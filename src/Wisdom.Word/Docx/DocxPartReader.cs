using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// DOCX 部件读取器：把 WordprocessingML 的 XML 元素（段落、表格、行内内容等）
    /// 解析为统一文档模型中的 <see cref="Block"/> 对象。
    /// Reads DOCX parts by parsing WordprocessingML XML elements (paragraphs, tables, inline content)
    /// into <see cref="Block"/> objects of the unified document model.
    /// </summary>
    /// <remarks>
    /// 该读取器绑定单个部件路径（正文、页眉或页脚），解析过程中通过
    /// <see cref="DocxPackage"/> 加载图片资源、通过 <see cref="DocxStyles"/> 还原样式。
    ///
    /// An instance is bound to one part path (body, header or footer); it loads images through
    /// <see cref="DocxPackage"/> and restores formatting through <see cref="DocxStyles"/> while parsing.
    /// </remarks>
    internal sealed class DocxPartReader
    {
        /// <summary>所属的 DOCX 包，用于按关系 ID 加载图片等资源。The owning DOCX package used to load images and other resources by relationship id.</summary>
        private readonly DocxPackage _package;
        /// <summary>文档样式表，用于将样式 ID 解析为段落/字符格式。The document style table that maps style ids to paragraph and character formats.</summary>
        private readonly DocxStyles _styles;
        /// <summary>当前解析的部件路径（如 word/document.xml），解析关系时以此为基准。The part path being parsed (e.g. word/document.xml); relationships are resolved relative to it.</summary>
        private readonly string _partPath;
        /// <summary>是否正在解析被删除的修订内容（w:del / w:moveFrom），此时 w:delText 计入正文。Whether deleted revision content (w:del / w:moveFrom) is being parsed; w:delText then contributes text.</summary>
        private bool _readingDeletedText;
        /// <summary>
        /// 本部件内 <c>w:bookmarkEnd</c> 的 id 到书签名的映射：<c>w:bookmarkEnd</c> 只带 <c>w:id</c>，
        /// 名字要靠先前 <c>w:bookmarkStart</c> 反查；书签 id 在 OOXML 中按部件独立编号，故映射随部件实例存在。
        /// Maps <c>w:bookmarkEnd</c> ids to bookmark names for this part: an end carries only <c>w:id</c>, so its
        /// name comes from the matching start. Bookmark ids restart per part, hence the map lives on the part instance.
        /// </summary>
        private readonly Dictionary<string, string> _bookmarkNamesById = new Dictionary<string, string>(StringComparer.Ordinal);

        /// <summary>
        /// 块级内容（表格嵌套、w:sdt 包装）的最大递归深度，超出后停止下探，防御畸形文档撑爆调用栈。
        /// The maximum recursion depth for block content (nested tables, w:sdt wrappers); deeper content is
        /// ignored so a malformed document cannot exhaust the call stack.
        /// </summary>
        private const int MaxBlockDepth = 64;

        /// <summary>
        /// 行内内容（域、超链接、文本框等容器）的最大递归深度，超出后停止下探。
        /// The maximum recursion depth for inline content (fields, hyperlinks, text boxes and similar wrappers).
        /// </summary>
        private const int MaxInlineDepth = 64;

        /// <summary>
        /// 单元格跨列数上限（OOXML 真实表格远小于此），防止恶意 gridSpan 让下游导出按列数分配超大数组。
        /// The cap on a cell's column span (real OOXML tables are far below it); it stops a hostile gridSpan from
        /// making downstream exports allocate arrays sized by the column count.
        /// </summary>
        private const int MaxColSpan = 1024;

        /// <summary>
        /// 行首/行尾网格缩进格数上限（OOXML 规定 gridBefore、gridAfter 取值 0-63）。
        /// The cap on a row's grid indent (OOXML restricts gridBefore and gridAfter to 0-63).
        /// </summary>
        private const int MaxGridOffset = 63;

        /// <summary>
        /// 创建绑定到指定包、样式表和部件路径的读取器实例。
        /// Creates a reader bound to the given package, style table and part path.
        /// </summary>
        /// <param name="package">DOCX 包对象。The DOCX package.</param>
        /// <param name="styles">文档样式表。The document style table.</param>
        /// <param name="partPath">当前部件的路径，用于解析图片等相对关系。The current part path used to resolve relative relationships such as images.</param>
        public DocxPartReader(DocxPackage package, DocxStyles styles, string partPath)
        {
            _package = package;
            _styles = styles;
            _partPath = partPath;
        }

        /// <summary>
        /// 读取容器元素下的全部块级内容并返回新列表。
        /// Reads all block-level content under a container element and returns a new list.
        /// </summary>
        /// <param name="container">包含 w:p / w:tbl 等子元素的容器（如 w:body、w:sdtContent）。A container holding w:p / w:tbl children (e.g. w:body, w:sdtContent).</param>
        /// <returns>解析得到的块对象列表。The parsed list of block objects.</returns>
        public List<Block> ReadBlocks(XElement container)
        {
            List<Block> blocks = new List<Block>();
            ReadBlocks(container, blocks, 0);
            return blocks;
        }

        /// <summary>
        /// 读取容器元素下的块级内容并追加到已有列表中；自动展开 w:sdt 与 w:customXml 包装。
        /// Reads block-level content under a container and appends it to an existing list;
        /// w:sdt and w:customXml wrappers are expanded automatically.
        /// </summary>
        /// <param name="container">待遍历的容器元素，为 null 时直接返回。The container to walk; returns immediately when null.</param>
        /// <param name="blocks">接收解析结果的块列表。The list that receives the parsed blocks.</param>
        public void ReadBlocks(XElement container, List<Block> blocks)
        {
            ReadBlocks(container, blocks, 0);
        }

        /// <summary>
        /// 按递归深度读取容器下的块级内容，超过 <see cref="MaxBlockDepth"/> 的嵌套不再解析。
        /// Reads block content under a container while tracking recursion depth; nesting deeper than
        /// <see cref="MaxBlockDepth"/> is not parsed.
        /// </summary>
        /// <param name="container">待遍历的容器元素。The container to walk.</param>
        /// <param name="blocks">接收解析结果的块列表。The list that receives the parsed blocks.</param>
        /// <param name="depth">当前块级嵌套深度。The current block nesting depth.</param>
        private void ReadBlocks(XElement container, List<Block> blocks, int depth)
        {
            if (container == null || depth > MaxBlockDepth) return;
            foreach (XElement child in container.Elements())
            {
                if (child.Name == Wml.W + "sdt")
                {
                    ReadBlocks(Wml.Child(child, "sdtContent"), blocks, depth + 1);
                    continue;
                }
                if (child.Name == Wml.W + "customXml")
                {
                    ReadBlocks(child, blocks, depth + 1);
                    continue;
                }
                Block block = ReadBlock(child, depth);
                if (block != null) blocks.Add(block);
            }
        }

        /// <summary>
        /// 按元素类型分派解析：w:p 解析为段落、w:tbl 解析为表格，其余返回 null。
        /// Dispatches parsing by element type: w:p becomes a paragraph, w:tbl a table, anything else returns null.
        /// </summary>
        /// <param name="element">待解析的块级元素。The block-level element to parse.</param>
        /// <param name="depth">当前块级嵌套深度。The current block nesting depth.</param>
        /// <returns>解析得到的块对象，或 null（表示不是可识别的块）。The parsed block, or null when the element is not a recognised block.</returns>
        private Block ReadBlock(XElement element, int depth)
        {
            if (element.Name == Wml.W + "p") return ReadParagraph(element);
            if (element.Name == Wml.W + "tbl") return ReadTable(element, depth + 1);
            return null;
        }

        /// <summary>
        /// 解析 w:p 段落元素：先应用段落样式与直接格式，再按样式基准格式解析行内内容。
        /// Parses a w:p paragraph: applies the paragraph style and direct formatting first, then parses
        /// inline content using the format inherited from the style.
        /// </summary>
        /// <param name="element">w:p 段落元素。The w:p paragraph element.</param>
        /// <returns>解析得到的 <see cref="Paragraph"/> 对象。The parsed <see cref="Paragraph"/> object.</returns>
        public Paragraph ReadParagraph(XElement element)
        {
            XElement properties = Wml.Child(element, "pPr");
            string styleId = null;
            if (properties != null)
            {
                XElement style = Wml.Child(properties, "pStyle");
                styleId = Wml.AttributeValue(style, "val");
            }
            if (string.IsNullOrEmpty(styleId)) styleId = _styles.DefaultParagraphStyleId;

            Paragraph paragraph = new Paragraph();
            _styles.ApplyParagraphStyle(paragraph, styleId);
            if (properties != null) _styles.ApplyParagraphProperties(paragraph, properties);
            paragraph.ListMarker = _styles.GetListMarker(properties);

            XElement markProperties = Wml.Child(properties, "rPr");
            XElement markInsert = Wml.Child(markProperties, "ins");
            XElement markDelete = Wml.Child(markProperties, "del");
            if (markInsert != null) paragraph.Revision = CreateRevision(markInsert, RevisionType.Insert);
            else if (markDelete != null) paragraph.Revision = CreateRevision(markDelete, RevisionType.Delete);

            TextFormat baseFormat = _styles.GetParagraphRunFormat(styleId);
            AppendInlines(element, paragraph, baseFormat, 0);
            return paragraph;
        }

        /// <summary>
        /// 解析 w:ins / w:del 等修订容器上的作者、时间与 ID。
        /// Parses the author, timestamp and id carried by a revision container such as w:ins or w:del.
        /// </summary>
        /// <param name="element">修订容器元素。The revision container element.</param>
        /// <param name="type">修订类型。The revision type.</param>
        /// <returns>修订元数据。The revision metadata.</returns>
        private static RevisionInfo CreateRevision(XElement element, RevisionType type)
        {
            RevisionInfo revision = new RevisionInfo(type, Wml.AttributeValue(element, "author"), null);
            revision.Id = Wml.AttributeValue(element, "id") ?? string.Empty;
            string rawDate = Wml.AttributeValue(element, "date");
            DateTime date;
            if (!string.IsNullOrEmpty(rawDate)
                && DateTime.TryParse(rawDate, CultureInfo.InvariantCulture,
                    DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out date))
            {
                revision.Date = date;
            }
            return revision;
        }

        /// <summary>
        /// 递归解析容器内的行内容元素（w:r、w:fldSimple、w:hyperlink 等）并追加到段落。
        /// w:del / w:moveFrom 表示被删除或移出的内容，按惯例跳过不读取。
        /// Recursively parses run-level children (w:r, w:fldSimple, w:hyperlink, ...) of a container and
        /// appends them to the paragraph. w:del / w:moveFrom mark deleted or moved-away content and is skipped.
        /// </summary>
        /// <param name="container">承载行内元素的容器。The container holding the inline elements.</param>
        /// <param name="paragraph">接收行内内容的目标段落。The target paragraph receiving the inline content.</param>
        /// <param name="baseFormat">段落样式提供的基准文字格式。The base text format supplied by the paragraph style.</param>
        /// <param name="depth">当前行内嵌套深度。The current inline nesting depth.</param>
        private void AppendInlines(XElement container, Paragraph paragraph, TextFormat baseFormat, int depth)
        {
            if (container == null || depth > MaxInlineDepth) return;
            foreach (XElement child in container.Elements())
            {
                string name = child.Name.LocalName;
                if (child.Name.Namespace != Wml.W && name != "AlternateContent" && name != "Choice" && name != "Fallback")
                {
                    continue;
                }
                switch (name)
                {
                    case "r":
                        ReadRun(child, paragraph, baseFormat);
                        break;
                    case "fldSimple":
                        ReadSimpleField(child, paragraph, baseFormat, depth);
                        break;
                    case "hyperlink":
                        {
                            string target = ResolveHyperlinkTarget(child);
                            int before = paragraph.Inlines.Count;
                            AppendInlines(child, paragraph, baseFormat, depth + 1);
                            if (!string.IsNullOrEmpty(target))
                            {
                                for (int i = before; i < paragraph.Inlines.Count; i++)
                                {
                                    TextRun linked = paragraph.Inlines[i] as TextRun;
                                    if (linked != null && string.IsNullOrEmpty(linked.Link)) linked.Link = target;
                                }
                            }
                        }
                        break;
                    case "ins":
                    case "moveTo":
                        {
                            RevisionInfo revision = CreateRevision(child, RevisionType.Insert);
                            int before = paragraph.Inlines.Count;
                            AppendInlines(child, paragraph, baseFormat, depth + 1);
                            for (int i = before; i < paragraph.Inlines.Count; i++)
                            {
                                paragraph.Inlines[i].Revision = revision.Clone();
                            }
                        }
                        break;
                    case "del":
                    case "moveFrom":
                        {
                            RevisionInfo revision = CreateRevision(child, RevisionType.Delete);
                            Paragraph deleted = new Paragraph();
                            bool saved = _readingDeletedText;
                            _readingDeletedText = true;
                            try
                            {
                                AppendInlines(child, deleted, baseFormat, depth + 1);
                            }
                            finally
                            {
                                _readingDeletedText = saved;
                            }
                            for (int i = 0; i < deleted.Inlines.Count; i++)
                            {
                                deleted.Inlines[i].Revision = revision.Clone();
                                paragraph.DeletedInlines.Add(deleted.Inlines[i]);
                            }
                        }
                        break;
                    case "smartTag":
                    case "customXml":
                    case "sdtContent":
                    case "sdt":
                    case "txbxContent":
                        AppendInlines(child, paragraph, baseFormat, depth + 1);
                        break;
                    case "bookmarkStart":
                        {
                            string bookmarkName = Wml.AttributeValue(child, "name");
                            string bookmarkId = Wml.AttributeValue(child, "id");
                            if (!string.IsNullOrEmpty(bookmarkName) && !bookmarkName.StartsWith("_GoBack", StringComparison.Ordinal))
                            {
                                paragraph.BookmarkName = bookmarkName;
                                paragraph.Inlines.Add(new BookmarkStart(bookmarkName));
                                // 记录 id 供配对的 w:bookmarkEnd 反查名字；id 缺失时只能靠名字配对。
                                // Record the id for the paired end; without an id the name alone pairs them.
                                if (!string.IsNullOrEmpty(bookmarkId)) _bookmarkNamesById[bookmarkId] = bookmarkName;
                            }
                        }
                        break;
                    case "bookmarkEnd":
                        {
                            string bookmarkId = Wml.AttributeValue(child, "id");
                            string pairedName;
                            if (string.IsNullOrEmpty(bookmarkId)
                                || !_bookmarkNamesById.TryGetValue(bookmarkId, out pairedName))
                            {
                                break;
                            }
                            _bookmarkNamesById.Remove(bookmarkId);
                            paragraph.Inlines.Add(new BookmarkEnd(pairedName));
                        }
                        break;
                }
            }
        }

        /// <summary>
        /// 解析超链接目标：优先取 w:anchor 表内锚点，其次按 r:id 读取外部关系目标；
        /// 表内锚点统一返回以 "#" 开头的形式，与统一模型的 <see cref="TextRun.Link"/> 约定一致。
        /// Resolves a hyperlink target: an in-document w:anchor wins, otherwise the target of the r:id
        /// relationship is used. In-document anchors are returned with a leading "#" to match the
        /// <see cref="TextRun.Link"/> convention of the unified model.
        /// </summary>
        /// <param name="hyperlink">w:hyperlink 元素。The w:hyperlink element.</param>
        /// <returns>超链接目标；无法解析时返回 null。The hyperlink target, or null when it cannot be resolved.</returns>
        private string ResolveHyperlinkTarget(XElement hyperlink)
        {
            string anchor = Wml.AttributeValue(hyperlink, "anchor");
            if (!string.IsNullOrEmpty(anchor)) return "#" + anchor;

            string relationshipId = Wml.AttributeValue(hyperlink, Wml.R + "id");
            if (string.IsNullOrEmpty(relationshipId)) return null;
            RelationshipInfo relationship = _package.GetRelationship(_partPath, relationshipId);
            if (relationship == null || string.IsNullOrEmpty(relationship.Target)) return null;
            if (relationship.External) return relationship.Target;

            string part = DocxPackage.ResolvePartPath(_partPath, relationship.Target);
            return string.IsNullOrEmpty(part) ? null : "#" + part;
        }

        /// <summary>
        /// 解析 w:fldSimple 简单域：提取域代码与缓存的显示文本，合并为一个 <see cref="FieldRun"/>。
        /// Parses a w:fldSimple field: extracts the field instruction and its cached display text and merges
        /// them into a single <see cref="FieldRun"/>.
        /// </summary>
        /// <param name="field">w:fldSimple 元素。The w:fldSimple element.</param>
        /// <param name="paragraph">接收域结果的目标段落。The target paragraph receiving the field result.</param>
        /// <param name="baseFormat">段落基准文字格式。The base text format of the paragraph.</param>
        /// <param name="depth">当前行内嵌套深度。The current inline nesting depth.</param>
        private void ReadSimpleField(XElement field, Paragraph paragraph, TextFormat baseFormat, int depth)
        {
            string instruction = Wml.AttributeValue(field, "instr") ?? string.Empty;
            Paragraph content = new Paragraph();
            AppendInlines(field, content, baseFormat, depth + 1);

            TextFormat format = baseFormat == null ? new TextFormat() : baseFormat.Clone();
            for (int i = 0; i < content.Inlines.Count; i++)
            {
                TextRun run = content.Inlines[i] as TextRun;
                if (run != null)
                {
                    format = run.Format == null ? format : run.Format.Clone();
                    break;
                }
            }
            paragraph.Inlines.Add(new FieldRun(instruction, content.GetText(), format));
        }

        /// <summary>
        /// 解析 w:r 运行元素：合并字符样式与直接格式（含中西文分别设置的字体），
        /// 并将 w:t 文本、制表符、换行、符号、图片等行内内容追加到段落。
        /// Parses a w:r run: merges the character style with direct formatting (including separate
        /// East-Asian and Latin fonts) and appends inline content such as w:t text, tabs, breaks, symbols
        /// and images to the paragraph.
        /// </summary>
        /// <param name="run">w:r 运行元素。The w:r run element.</param>
        /// <param name="paragraph">接收解析结果的目标段落。The target paragraph receiving the parsed content.</param>
        /// <param name="paragraphFormat">段落继承下来的基准文字格式。The base text format inherited from the paragraph.</param>
        private void ReadRun(XElement run, Paragraph paragraph, TextFormat paragraphFormat)
        {
            XElement properties = Wml.Child(run, "rPr");
            TextFormat format = paragraphFormat == null ? new TextFormat() : paragraphFormat.Clone();
            string runStyleId = null;
            if (properties != null)
            {
                XElement runStyle = Wml.Child(properties, "rStyle");
                runStyleId = Wml.AttributeValue(runStyle, "val");
            }
            if (!string.IsNullOrEmpty(runStyleId)) format = _styles.GetCharacterRunFormat(runStyleId, format);
            if (properties != null)
            {
                DocxStyles.ApplyRunProperties(format, properties);
                XElement fonts = Wml.Child(properties, "rFonts");
                if (fonts != null)
                {
                    string eastAsia = Wml.AttributeValue(fonts, "eastAsia");
                    string ascii = Wml.AttributeValue(fonts, "ascii") ?? Wml.AttributeValue(fonts, "hAnsi");
                    if (!string.IsNullOrEmpty(ascii))
                    {
                        format.AsciiFontFamily = ascii;
                        format.FontFamily = ascii;
                    }
                    if (!string.IsNullOrEmpty(eastAsia))
                    {
                        format.EastAsiaFontFamily = eastAsia;
                        bool hasCjk = HasCjk(CollectRunText(run));
                        if (hasCjk || string.IsNullOrEmpty(ascii)) format.FontFamily = eastAsia;
                    }
                }
            }

            StringBuilder text = new StringBuilder();
            foreach (XElement child in run.Elements())
            {
                if (child.Name == Wml.W + "rPr") continue;
                switch (child.Name.LocalName)
                {
                    case "t":
                        text.Append(child.Value);
                        break;
                    case "delText":
                        if (_readingDeletedText) text.Append(child.Value);
                        break;
                    case "instrText":
                        break;
                    case "tab":
                    case "ptab":
                        text.Append('\t');
                        break;
                    case "br":
                    case "cr":
                        text.Append(Wml.AttributeValue(child, "type") == "page" ? '\f' : '\n');
                        break;
                    case "noBreakHyphen":
                        text.Append('-');
                        break;
                    case "softHyphen":
                        break;
                    case "sym":
                        AppendSymbol(text, child);
                        break;
                    case "drawing":
                        FlushText(paragraph, text, format);
                        ReadDrawing(child, paragraph);
                        break;
                    case "pict":
                        FlushText(paragraph, text, format);
                        ReadVmlPicture(child, paragraph);
                        break;
                    case "AlternateContent":
                        FlushText(paragraph, text, format);
                        ReadAlternateContent(child, paragraph);
                        break;
                }
            }
            FlushText(paragraph, text, format);
        }

        /// <summary>
        /// 解析 mc:AlternateContent 备用内容：优先取 mc:Choice 分支，找不到时回退到 mc:Fallback，
        /// 再从中提取 w:drawing 或 v:pict 图形。
        /// Parses mc:AlternateContent markup: prefers the mc:Choice branch and falls back to mc:Fallback,
        /// then extracts the w:drawing or v:pict graphic from it.
        /// </summary>
        /// <param name="element">mc:AlternateContent 元素。The mc:AlternateContent element.</param>
        /// <param name="paragraph">接收图形结果的目标段落。The target paragraph receiving the graphic.</param>
        private void ReadAlternateContent(XElement element, Paragraph paragraph)
        {
            XElement choice = null;
            foreach (XElement child in element.Elements())
            {
                if (child.Name.LocalName == "Choice") { choice = child; break; }
                if (child.Name.LocalName == "Fallback") choice = child;
            }
            if (choice == null) return;
            XElement drawing = null;
            XElement pict = null;
            foreach (XElement descendant in choice.Descendants())
            {
                if (descendant.Name.LocalName == "drawing" && drawing == null) drawing = descendant;
                else if (descendant.Name.LocalName == "pict" && pict == null) pict = descendant;
            }
            if (drawing != null) ReadDrawing(drawing, paragraph);
            else if (pict != null) ReadVmlPicture(pict, paragraph);
        }

        /// <summary>
        /// 收集运行元素内所有 w:t / w:delText 子元素的纯文本拼接结果。
        /// Concatenates the plain text of every w:t / w:delText child of a run element.
        /// </summary>
        /// <param name="run">w:r 运行元素。The w:r run element.</param>
        /// <returns>运行的完整文本内容。The full text content of the run.</returns>
        private static string CollectRunText(XElement run)
        {
            StringBuilder builder = new StringBuilder();
            foreach (XElement child in run.Elements())
            {
                if (child.Name == Wml.W + "t" || child.Name == Wml.W + "delText") builder.Append(child.Value);
            }
            return builder.ToString();
        }

        /// <summary>
        /// 判断文本是否包含中日韩字符（CJK 统一表意文字、兼容表意文字或韩文音节区段），
        /// 用于决定运行字体回退到 eastAsia 字体还是西文字体。
        /// Detects whether the text contains CJK characters (Unified Ideographs, Compatibility Ideographs or
        /// Hangul Syllables) to decide whether the run font falls back to the eastAsia font or the Latin font.
        /// </summary>
        /// <param name="text">待检测的文本。The text to test.</param>
        /// <returns>包含任意 CJK 字符时返回 true。True when any CJK character is present.</returns>
        private static bool HasCjk(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if ((c >= 0x2E80 && c <= 0x9FFF) || (c >= 0xAC00 && c <= 0xD7AF) || (c >= 0xF900 && c <= 0xFAFF))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 将缓冲区中累积的文本作为 <see cref="TextRun"/> 刷入段落并清空缓冲区；缓冲区为空时不做任何事。
        /// Flushes the text accumulated in the buffer into the paragraph as a <see cref="TextRun"/> and clears
        /// the buffer; does nothing when the buffer is empty.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="text">待刷出的文本缓冲区。The text buffer to flush.</param>
        /// <param name="format">该段文本使用的文字格式。The text format applied to the flushed text.</param>
        private static void FlushText(Paragraph paragraph, StringBuilder text, TextFormat format)
        {
            if (text.Length == 0) return;
            string value = text.ToString();
            text.Length = 0;
            if (value.Length == 0) return;
            paragraph.Inlines.Add(new TextRun(value, format.Clone()));
        }

        /// <summary>
        /// 解析 w:sym 符号元素：将十六进制字符码转为 Unicode 文本追加到缓冲区，
        /// 并对私有区 0xF000 段的符号字体（Wingdings/Symbol 等）做偏移还原。
        /// Parses a w:sym symbol element: converts the hexadecimal character code into Unicode text and
        /// appends it to the buffer, undoing the 0xF000 private-area offset used by symbol fonts
        /// (Wingdings, Symbol, ...).
        /// </summary>
        /// <param name="text">文本缓冲区。The text buffer.</param>
        /// <param name="symbol">w:sym 元素。The w:sym element.</param>
        private static void AppendSymbol(StringBuilder text, XElement symbol)
        {
            string value = Wml.AttributeValue(symbol, "char");
            int code;
            if (!string.IsNullOrEmpty(value)
                && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
            {
                if (code >= 0xF000 && code <= 0xF0FF) code -= 0xF000;
                if (code > 0 && code <= 0x10FFFF)
                {
                    try
                    {
                        text.Append(char.ConvertFromUtf32(code));
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                    }
                }
            }
        }

        /// <summary>
        /// 解析 w:drawing  DrawingML 图形：通过 blip 的关系 ID 从包中加载图片，
        /// 并按 wp:extent 的 EMU 尺寸换算为磅值后加入段落。
        /// Parses a w:drawing DrawingML graphic: loads the image from the package using the blip relationship
        /// id and converts the wp:extent EMU size into points before adding it to the paragraph.
        /// </summary>
        /// <param name="drawing">w:drawing 元素。The w:drawing element.</param>
        /// <param name="paragraph">接收图片的目标段落。The target paragraph receiving the image.</param>
        private void ReadDrawing(XElement drawing, Paragraph paragraph)
        {
            XElement extent = null;
            XElement blip = null;
            foreach (XElement descendant in drawing.Descendants())
            {
                if (extent == null && descendant.Name == Wml.Wp + "extent") extent = descendant;
                if (blip == null && descendant.Name == Wml.A + "blip") blip = descendant;
            }
            if (blip == null) return;
            string relationshipId = Wml.AttributeValue(blip, Wml.R + "embed");
            if (string.IsNullOrEmpty(relationshipId)) relationshipId = Wml.AttributeValue(blip, Wml.R + "link");
            OfficeImage image = _package.LoadImage(_partPath, relationshipId);
            if (image == null) return;
            double width = image.DefaultWidthPoints;
            double height = image.DefaultHeightPoints;
            if (extent != null)
            {
                width = Wml.EmuToPoints(Wml.RawAttributeValue(extent, "cx"), width);
                height = Wml.EmuToPoints(Wml.RawAttributeValue(extent, "cy"), height);
            }
            if (width <= 0) width = image.DefaultWidthPoints;
            if (height <= 0) height = image.DefaultHeightPoints;
            paragraph.AddImage(image, width, height);
        }

        /// <summary>
        /// 解析 v:pict VML 旧式图片：通过 v:imagedata 的关系 ID 加载图片，
        /// 并从 v:shape 的 style 属性中解析 width/height 换算为磅值后加入段落。
        /// Parses a legacy v:pict VML image: loads the picture through the v:imagedata relationship id and
        /// reads width/height from the v:shape style attribute, converted to points, before adding it to the
        /// paragraph.
        /// </summary>
        /// <param name="pict">v:pict 元素。The v:pict element.</param>
        /// <param name="paragraph">接收图片的目标段落。The target paragraph receiving the image.</param>
        private void ReadVmlPicture(XElement pict, Paragraph paragraph)
        {
            XElement imageData = null;
            XElement shape = null;
            foreach (XElement descendant in pict.Descendants())
            {
                if (imageData == null && descendant.Name == Wml.V + "imagedata") imageData = descendant;
                if (shape == null && (descendant.Name == Wml.V + "shape" || descendant.Name == Wml.V + "rect"))
                {
                    shape = descendant;
                }
            }
            if (imageData == null) return;
            string relationshipId = Wml.AttributeValue(imageData, Wml.R + "id");
            OfficeImage image = _package.LoadImage(_partPath, relationshipId);
            if (image == null) return;
            double width = image.DefaultWidthPoints;
            double height = image.DefaultHeightPoints;
            if (shape != null)
            {
                string style = Wml.RawAttributeValue(shape, "style");
                if (!string.IsNullOrEmpty(style))
                {
                    string[] declarations = style.Split(';');
                    for (int i = 0; i < declarations.Length; i++)
                    {
                        int colon = declarations[i].IndexOf(':');
                        if (colon <= 0) continue;
                        string key = declarations[i].Substring(0, colon).Trim();
                        string value = declarations[i].Substring(colon + 1).Trim();
                        if (string.Equals(key, "width", StringComparison.OrdinalIgnoreCase))
                        {
                            width = Wml.MeasureToPoints(value, width);
                        }
                        else if (string.Equals(key, "height", StringComparison.OrdinalIgnoreCase))
                        {
                            height = Wml.MeasureToPoints(value, height);
                        }
                    }
                }
            }
            paragraph.AddImage(image, width, height);
        }

        /// <summary>
        /// 解析 w:tbl 表格元素：应用表格样式链与直接属性，读取网格列宽、表头行、行高、
        /// 单元格横向合并（gridSpan）、纵向合并（vMerge）与底色、边框、垂直对齐等。
        /// Parses a w:tbl table element: applies the table style chain and direct properties, then reads grid
        /// column widths, the header row, row heights, horizontal merges (gridSpan), vertical merges
        /// (vMerge), shading, borders and vertical alignment.
        /// </summary>
        /// <remarks>
        /// 纵向合并通过 <see cref="_previousRowCells"/> 记录上一行各列起点单元格，
        /// 遇到 continue 标记时将起点单元格的 RowSpan 加一，而不是生成新单元格。
        ///
        /// Vertical merges are handled by remembering the origin cell of each grid column in the previous row
        /// (<see cref="_previousRowCells"/>); a continue marker increases that cell's RowSpan instead of
        /// creating a new cell.
        /// </remarks>
        /// <param name="element">w:tbl 表格元素。The w:tbl table element.</param>
        /// <returns>解析得到的 <see cref="Table"/> 对象。The parsed <see cref="Table"/> object.</returns>
        public Table ReadTable(XElement element)
        {
            return ReadTable(element, 1);
        }

        /// <summary>
        /// 按嵌套深度解析表格：单元格内容以 depth+1 继续下探，并在返回前恢复外层表格的纵向合并状态，
        /// 使嵌套表格不会打断外层单元格的上延续行关系。
        /// Parses a table at the given nesting depth: cell content continues at depth+1, and the previous-row
        /// state is restored before returning so a nested table does not break the outer table's vertical merges.
        /// </summary>
        /// <param name="element">w:tbl 表格元素。The w:tbl table element.</param>
        /// <param name="depth">当前块级嵌套深度。The current block nesting depth.</param>
        /// <returns>解析得到的 <see cref="Table"/> 对象。The parsed <see cref="Table"/> object.</returns>
        private Table ReadTable(XElement element, int depth)
        {
            Dictionary<int, TableCell> savedRowCells = _previousRowCells;
            _previousRowCells = null;
            Table table = new Table();
            XElement tableProperties = Wml.Child(element, "tblPr");
            if (tableProperties != null)
            {
                string tableStyleId = null;
                XElement tableStyle = Wml.Child(tableProperties, "tblStyle");
                if (tableStyle != null) tableStyleId = Wml.AttributeValue(tableStyle, "val");
                table.Borders = TableBorders.None();
                if (!string.IsNullOrEmpty(tableStyleId))
                {
                    List<XElement> chain = _styles.GetTablePropertyChain(tableStyleId);
                    for (int i = 0; i < chain.Count; i++) ApplyTableProperties(table, chain[i]);
                }
                ApplyTableProperties(table, tableProperties);
            }
            else
            {
                table.Borders = TableBorders.None();
            }

            XElement grid = Wml.Child(element, "tblGrid");
            if (grid != null)
            {
                foreach (XElement column in Wml.Children(grid, "gridCol"))
                {
                    double width = Wml.TwipsToPoints(Wml.AttributeValue(column, "w"), 0);
                    if (width > 0) table.Columns.Add(new TableColumn(width));
                }
            }

            foreach (XElement rowElement in Wml.Children(element, "tr"))
            {
                TableRow row = table.AddRow();
                XElement rowProperties = Wml.Child(rowElement, "trPr");
                int gridBefore = 0;
                int gridAfter = 0;
                if (rowProperties != null)
                {
                    row.IsHeader = Wml.GetOnOff(rowProperties, "tblHeader", false);
                    XElement height = Wml.Child(rowProperties, "trHeight");
                    if (height != null)
                    {
                        row.Height = Wml.TwipsToPoints(Wml.AttributeValue(height, "val"), 0);
                    }
                    gridBefore = Math.Min(MaxGridOffset, Math.Max(0, Wml.GetInt(rowProperties, "gridBefore", 0)));
                    gridAfter = Math.Min(MaxGridOffset, Math.Max(0, Wml.GetInt(rowProperties, "gridAfter", 0)));
                }

                Dictionary<int, TableCell> previousCells = _previousRowCells;
                Dictionary<int, TableCell> currentCells = new Dictionary<int, TableCell>();
                int column = 0;
                for (int i = 0; i < gridBefore; i++)
                {
                    TableCell filler = new TableCell();
                    row.Cells.Add(filler);
                    currentCells[column++] = filler;
                }

                foreach (XElement cellElement in Wml.Children(rowElement, "tc"))
                {
                    TableCell cell = new TableCell();
                    XElement cellProperties = Wml.Child(cellElement, "tcPr");
                    bool verticalMergeContinue = false;
                    if (cellProperties != null)
                    {
                        XElement span = Wml.Child(cellProperties, "gridSpan");
                        if (span != null)
                        {
                            cell.ColSpan = Math.Min(MaxColSpan, Math.Max(1, Wml.ParseInt(Wml.AttributeValue(span, "val"), 1)));
                        }
                        XElement verticalMerge = Wml.Child(cellProperties, "vMerge");
                        if (verticalMerge != null)
                        {
                            string mergeValue = Wml.AttributeValue(verticalMerge, "val");
                            verticalMergeContinue = mergeValue == null || mergeValue == "continue";
                        }
                        XElement shading = Wml.Child(cellProperties, "shd");
                        if (shading != null)
                        {
                            string fill = Wml.AttributeValue(shading, "fill");
                            if (!string.IsNullOrEmpty(fill) && fill != "auto" && fill != "none")
                            {
                                OfficeColor color;
                                if (OfficeColor.TryParse(fill, out color)) cell.BackgroundColor = color;
                            }
                        }
                        cell.Borders = ReadCellBorders(Wml.Child(cellProperties, "tcBorders"));
                        XElement vertical = Wml.Child(cellProperties, "vAlign");
                        string verticalValue = Wml.AttributeValue(vertical, "val");
                        if (verticalValue == "center") cell.VerticalAlignment = VerticalAlignment.Middle;
                        else if (verticalValue == "bottom") cell.VerticalAlignment = VerticalAlignment.Bottom;
                    }

                    if (verticalMergeContinue && previousCells != null)
                    {
                        TableCell origin;
                        if (previousCells.TryGetValue(column, out origin))
                        {
                            origin.RowSpan = Math.Max(1, origin.RowSpan) + 1;
                            currentCells[column] = origin;
                            column += cell.ColSpan;
                            continue;
                        }
                    }

                    ReadBlocks(cellElement, cell.Blocks, depth);
                    if (cell.Blocks.Count == 0) cell.Blocks.Add(new Paragraph());
                    row.Cells.Add(cell);
                    currentCells[column] = cell;
                    column += Math.Max(1, cell.ColSpan);
                }

                for (int i = 0; i < gridAfter; i++)
                {
                    TableCell filler = new TableCell();
                    row.Cells.Add(filler);
                    currentCells[column++] = filler;
                }
                _previousRowCells = currentCells;
            }
            table.RepeatHeaderRow = table.Rows.Count > 0 && table.Rows[0].IsHeader;
            _previousRowCells = savedRowCells;
            return table;
        }

        /// <summary>表格解析期间记录的上一行「列号 -&gt; 单元格」映射，用于处理 vMerge 纵向合并。Mapping of "grid column -&gt; cell" for the previous row, kept while parsing tables to handle vMerge vertical merges.</summary>
        private Dictionary<int, TableCell> _previousRowCells;

        /// <summary>
        /// 解析 w:tcBorders 单元格边框元素，四个方向均缺失时返回 null。
        /// Parses a w:tcBorders cell-border element; returns null when all four sides are missing.
        /// </summary>
        /// <param name="element">w:tcBorders 元素，可为 null。The w:tcBorders element, may be null.</param>
        /// <returns>解析得到的边框定义，或 null。The parsed border definition, or null.</returns>
        private static TableCellBorders ReadCellBorders(XElement element)
        {
            if (element == null) return null;
            TableCellBorders borders = new TableCellBorders();
            bool any = false;
            XElement top = Wml.Child(element, "top");
            XElement bottom = Wml.Child(element, "bottom");
            XElement left = Wml.Child(element, "left") ?? Wml.Child(element, "start");
            XElement right = Wml.Child(element, "right") ?? Wml.Child(element, "end");
            if (top != null) { borders.Top = ReadBorder(top); any = true; }
            if (bottom != null) { borders.Bottom = ReadBorder(bottom); any = true; }
            if (left != null) { borders.Left = ReadBorder(left); any = true; }
            if (right != null) { borders.Right = ReadBorder(right); any = true; }
            return any ? borders : null;
        }

        /// <summary>
        /// 将一份表格属性元素（w:tblPr 或样式链中的属性）应用到表格：边框、单元格边距、对齐方式与自适应布局。
        /// Applies one table-property element (w:tblPr or an entry of the style chain) to the table: borders,
        /// cell margins, alignment and autofit layout.
        /// </summary>
        /// <param name="table">目标表格。The target table.</param>
        /// <param name="properties">表格属性元素，可为 null。The table properties element, may be null.</param>
        private void ApplyTableProperties(Table table, XElement properties)
        {
            if (properties == null) return;
            XElement borders = Wml.Child(properties, "tblBorders");
            if (borders != null)
            {
                XElement top = Wml.Child(borders, "top");
                XElement bottom = Wml.Child(borders, "bottom");
                XElement left = Wml.Child(borders, "left") ?? Wml.Child(borders, "start");
                XElement right = Wml.Child(borders, "right") ?? Wml.Child(borders, "end");
                XElement insideH = Wml.Child(borders, "insideH");
                XElement insideV = Wml.Child(borders, "insideV");
                if (top != null) table.Borders.Top = ReadBorder(top);
                if (bottom != null) table.Borders.Bottom = ReadBorder(bottom);
                if (left != null) table.Borders.Left = ReadBorder(left);
                if (right != null) table.Borders.Right = ReadBorder(right);
                if (insideH != null) table.Borders.InsideHorizontal = ReadBorder(insideH);
                if (insideV != null) table.Borders.InsideVertical = ReadBorder(insideV);
            }
            XElement margins = Wml.Child(properties, "tblCellMar");
            if (margins != null)
            {
                double left = ReadCellMargin(Wml.Child(margins, "left") ?? Wml.Child(margins, "start"));
                double right = ReadCellMargin(Wml.Child(margins, "right") ?? Wml.Child(margins, "end"));
                double padding = Math.Max(left, right);
                if (padding > 0) table.CellPadding = padding;
            }
            XElement alignment = Wml.Child(properties, "tblJc");
            string alignmentValue = Wml.AttributeValue(alignment, "val");
            if (alignmentValue == "center") table.Alignment = TableAlignment.Center;
            else if (alignmentValue == "right" || alignmentValue == "end") table.Alignment = TableAlignment.Right;
            XElement layout = Wml.Child(properties, "tblLayout");
            string layoutValue = Wml.AttributeValue(layout, "type");
            if (!string.IsNullOrEmpty(layoutValue)) table.Autofit = layoutValue == "autofit";
        }

        /// <summary>
        /// 读取单元格边距元素（w:tblCellMar 的子项），仅支持 dxa（缇）单位，其余单位返回 0。
        /// Reads a cell-margin element (a child of w:tblCellMar); only the dxa (twip) unit is supported, other
        /// units return 0.
        /// </summary>
        /// <param name="margin">边距元素，可为 null。The margin element, may be null.</param>
        /// <returns>换算为磅值的边距。The margin converted to points.</returns>
        private static double ReadCellMargin(XElement margin)
        {
            if (margin == null) return 0;
            string type = Wml.AttributeValue(margin, "type");
            string value = Wml.AttributeValue(margin, "w");
            if (type == "dxa" || type == null) return Wml.TwipsToPoints(value, 0);
            return 0;
        }

        /// <summary>
        /// 解析单个边框元素（w:top、w:bottom 等）：将 OOXML 边框值归类为无边框/双线/虚线/点线/单线，
        /// 并把以 1/8 磅为单位的厚度换算为磅值。
        /// Parses a single border element (w:top, w:bottom, ...): maps the OOXML border value onto
        /// none/double/dashed/dotted/single and converts the thickness, expressed in eighths of a point, to points.
        /// </summary>
        /// <param name="element">边框元素，为 null 时返回默认（无样式）边框。The border element; a null value yields the default (unstyled) border.</param>
        /// <returns>解析得到的 <see cref="BorderLine"/> 对象。The parsed <see cref="BorderLine"/> object.</returns>
        private static BorderLine ReadBorder(XElement element)
        {
            BorderLine border = new BorderLine();
            if (element == null) return border;
            string value = Wml.AttributeValue(element, "val");
            if (string.IsNullOrEmpty(value) || value == "none" || value == "nil")
            {
                border.Style = BorderStyle.None;
                return border;
            }
            if (value == "double" || value == "thinThickSmallGap" || value == "thickThinSmallGap")
            {
                border.Style = BorderStyle.Double;
            }
            else if (value.StartsWith("dash", StringComparison.Ordinal))
            {
                border.Style = BorderStyle.Dashed;
            }
            else if (value.StartsWith("dot", StringComparison.Ordinal))
            {
                border.Style = BorderStyle.Dotted;
            }
            else
            {
                border.Style = BorderStyle.Single;
            }
            int thickness = Wml.ParseInt(Wml.AttributeValue(element, "sz"), 4);
            border.Width = thickness > 0 ? thickness / 8.0 : 0.5;
            border.Color = Wml.ParseColor(Wml.AttributeValue(element, "color"), OfficeColor.Black);
            return border;
        }
    }
}
