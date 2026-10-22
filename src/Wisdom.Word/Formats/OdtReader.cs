using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 读取 OpenDocument 文本（.odt）：解析 content.xml 的正文、styles.xml 的自动样式/页面布局/主页面
    /// 以及 meta.xml 的元数据，重建段落、字符格式、页眉页脚、表格（含合并与底色）、书签、超链接、
    /// 页码域与内嵌图片。仅读取包内数据，不访问任何外部资源。
    /// Reads OpenDocument Text (.odt): parses the content.xml body, the automatic styles/page layouts/master
    /// pages of styles.xml and the meta.xml metadata, rebuilding paragraphs, character formatting, headers and
    /// footers, tables (including merges and shading), bookmarks, hyperlinks, page-number fields and embedded
    /// images. Only package data is read; no external resource is accessed.
    /// </summary>
    public static class OdtReader
    {
        /// <summary>解析时允许的最大 XML 节点数。The maximum number of XML nodes accepted while parsing.</summary>
        private const int MaxNodes = 400000;
        /// <summary>允许读取的内嵌图片数量上限。The maximum number of embedded images that may be read.</summary>
        private const int MaxImages = 512;
        /// <summary>单张内嵌图片的最大字节数（64 MB）。The maximum size of one embedded image (64 MB).</summary>
        private const int MaxImageBytes = 64 * 1024 * 1024;
        /// <summary>样式父链的最大解析深度。The maximum resolution depth of style parent chains.</summary>
        private const int MaxStyleDepth = 16;
        /// <summary>单个元素重复属性的最大展开次数。The maximum expansion of one element's repeat attribute.</summary>
        private const int MaxRepeat = 1024;
        /// <summary>表格列数的安全上限。The safety cap on the number of table columns.</summary>
        private const int MaxColumns = 4096;
        /// <summary>嵌套块的最大解析深度。The maximum parsing depth of nested blocks.</summary>
        private const int MaxBlockDepth = 32;

        private static readonly XNamespace OfficeNs = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        private static readonly XNamespace StyleNs = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        private static readonly XNamespace TextNs = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private static readonly XNamespace TableNs = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        private static readonly XNamespace DrawNs = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
        private static readonly XNamespace FoNs = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        private static readonly XNamespace SvgNs = "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
        private static readonly XNamespace XLinkNs = "http://www.w3.org/1999/xlink";
        private static readonly XNamespace MetaNs = "urn:oasis:names:tc:opendocument:xmlns:meta:1.0";
        private static readonly XNamespace DcNs = "http://purl.org/dc/elements/1.1/";

        /// <summary>从文件读取 ODT 文档。Loads an ODT document from a file.</summary>
        /// <param name="path">文件路径。The file path.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            using (FileStream stream = File.OpenRead(path))
            {
                return Load(stream);
            }
        }

        /// <summary>从流读取 ODT 文档；流保持打开。Loads an ODT document from a stream, leaving it open.</summary>
        /// <param name="stream">ODT 字节流。The stream of ODT bytes.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            if (stream.CanSeek)
            {
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read, true))
                {
                    return LoadPackage(archive);
                }
            }
            using (MemoryStream buffer = new MemoryStream())
            {
                byte[] chunk = new byte[65536];
                int read;
                while ((read = stream.Read(chunk, 0, chunk.Length)) > 0)
                {
                    buffer.Write(chunk, 0, read);
                }
                buffer.Position = 0;
                using (ZipArchive archive = new ZipArchive(buffer, ZipArchiveMode.Read, true))
                {
                    return LoadPackage(archive);
                }
            }
        }

        /// <summary>从字节数组读取 ODT 文档。Loads an ODT document from a byte array.</summary>
        /// <param name="data">ODT 字节。The ODT bytes.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            using (MemoryStream stream = new MemoryStream(data, false))
            {
                return Load(stream);
            }
        }

        /// <summary>读取整个 ODT 包并重建文档。Reads the whole ODT package and rebuilds the document.</summary>
        private static OfficeDocument LoadPackage(ZipArchive archive)
        {
            PartReader parts = new PartReader(archive);
            byte[] contentBytes = parts.Read("content.xml");
            if (contentBytes == null) throw new InvalidDataException("不是有效的 ODT 文档：缺少 content.xml。");

            XDocument content = Wml.LoadXml(contentBytes);
            if (content == null || content.Root == null) throw new InvalidDataException("不是有效的 ODT 文档：content.xml 为空。");

            XDocument stylesDocument = null;
            byte[] stylesBytes = parts.Read("styles.xml");
            if (stylesBytes != null)
            {
                try { stylesDocument = Wml.LoadXml(stylesBytes); }
                catch (XmlException) { stylesDocument = null; }
            }

            XDocument metaDocument = null;
            byte[] metaBytes = parts.Read("meta.xml");
            if (metaBytes != null)
            {
                try { metaDocument = Wml.LoadXml(metaBytes); }
                catch (XmlException) { metaDocument = null; }
            }

            StyleIndex styles = new StyleIndex();
            styles.Collect(content.Root);
            if (stylesDocument != null && stylesDocument.Root != null)
            {
                styles.Collect(stylesDocument.Root);
            }

            OfficeDocument document = new OfficeDocument();
            ParseMetadata(document.Metadata, metaDocument);

            ParseContext context = new ParseContext();
            context.Styles = styles;
            context.Parts = parts;

            XElement body = FindChild(content.Root, OfficeNs + "body");
            XElement text = body == null ? null : FindChild(body, OfficeNs + "text");
            if (text == null) throw new InvalidDataException("不是有效的 ODT 文档：缺少 office:text 正文。");

            BodyState state = new BodyState();
            state.Document = document;
            ParseBody(text, state, context, 0);
            if (document.Sections.Count == 0)
            {
                document.AddSection();
            }
            return document;
        }

        // ---------------------------------------------------------------- metadata

        /// <summary>解析 meta.xml 中的标题、作者、主题、关键词与时间。Parses title, author, subject, keywords and times from meta.xml.</summary>
        private static void ParseMetadata(DocumentMetadata metadata, XDocument meta)
        {
            if (metadata == null || meta == null || meta.Root == null) return;
            XElement officeMeta = FindChild(meta.Root, OfficeNs + "meta");
            if (officeMeta == null) return;
            StringBuilder keywords = new StringBuilder();
            foreach (XElement child in officeMeta.Elements())
            {
                if (child.Name == DcNs + "title") metadata.Title = child.Value;
                else if (child.Name == DcNs + "creator") metadata.Author = child.Value;
                else if (child.Name == DcNs + "subject") metadata.Subject = child.Value;
                else if (child.Name == MetaNs + "keyword")
                {
                    if (keywords.Length > 0) keywords.Append(", ");
                    keywords.Append(child.Value);
                }
                else if (child.Name == MetaNs + "generator")
                {
                    if (!string.IsNullOrEmpty(child.Value)) metadata.Creator = child.Value;
                }
                else if (child.Name == MetaNs + "creation-date")
                {
                    DateTime value;
                    if (TryParseDate(child.Value, out value)) metadata.Created = value;
                }
                else if (child.Name == DcNs + "date")
                {
                    DateTime value;
                    if (TryParseDate(child.Value, out value)) metadata.Modified = value;
                }
            }
            if (keywords.Length > 0) metadata.Keywords = keywords.ToString();
        }

        /// <summary>按 ISO 8601 解析日期时间。Parses a date-time value as ISO 8601.</summary>
        private static bool TryParseDate(string text, out DateTime value)
        {
            return DateTime.TryParse(text, CultureInfo.InvariantCulture,
                DateTimeStyles.RoundtripKind | DateTimeStyles.AllowWhiteSpaces, out value);
        }

        // ---------------------------------------------------------------- body

        /// <summary>正文解析状态：当前文档与当前节。Body parsing state: the document and the current section.</summary>
        private sealed class BodyState
        {
            /// <summary>正在填充的文档。The document being filled.</summary>
            public OfficeDocument Document;
            /// <summary>当前节，null 表示尚未创建。The current section; null until created.</summary>
            public Section Section;
            /// <summary>当前节使用的主页面名。The master-page name used by the current section.</summary>
            public string MasterPage;
        }

        /// <summary>解析正文级元素：段落、标题、表格、节、列表与分页。Parses body-level elements: paragraphs, headings, tables, sections, lists and page breaks.</summary>
        private static void ParseBody(XElement container, BodyState state, ParseContext context, int depth)
        {
            if (container == null || depth > MaxBlockDepth) return;
            foreach (XElement element in container.Elements())
            {
                context.CountNode();
                XName name = element.Name;
                if (name == TextNs + "p" || name == TextNs + "h")
                {
                    ParseBodyParagraph(element, state, context);
                }
                else if (name == TableNs + "table")
                {
                    EnsureSection(state, context).Blocks.Add(ParseTable(element, context, depth + 1));
                }
                else if (name == TextNs + "section")
                {
                    ParseBody(element, state, context, depth + 1);
                }
                else if (name == TextNs + "list")
                {
                    ParseBodyList(element, state, context, depth + 1);
                }
                else if (name == TextNs + "page-break" || name == TextNs + "soft-page-break")
                {
                    EnsureSection(state, context).Blocks.Add(new PageBreak());
                }
            }
        }

        /// <summary>解析列表项内的段落与表格。Parses the paragraphs and tables inside list items.</summary>
        private static void ParseBodyList(XElement list, BodyState state, ParseContext context, int depth)
        {
            if (depth > MaxBlockDepth) return;
            foreach (XElement item in list.Elements(TextNs + "list-item"))
            {
                foreach (XElement child in item.Elements())
                {
                    context.CountNode();
                    if (child.Name == TextNs + "p" || child.Name == TextNs + "h")
                    {
                        ParseBodyParagraph(child, state, context);
                    }
                    else if (child.Name == TableNs + "table")
                    {
                        EnsureSection(state, context).Blocks.Add(ParseTable(child, context, depth + 1));
                    }
                    else if (child.Name == TextNs + "list")
                    {
                        ParseBodyList(child, state, context, depth + 1);
                    }
                }
            }
        }

        /// <summary>解析一个正文段落；主页面变化时切换节，分页属性转换为分页块。Parses one body paragraph; a master-page change starts a new section and page-break properties become page-break blocks.</summary>
        private static void ParseBodyParagraph(XElement element, BodyState state, ParseContext context)
        {
            ParagraphInfo info = ParseParagraph(element, context);
            string master = info.MasterPage ?? string.Empty;
            bool switched = false;
            if (state.Section == null
                || (!string.IsNullOrEmpty(master) && !string.Equals(master, state.MasterPage ?? string.Empty, StringComparison.Ordinal)))
            {
                state.Section = CreateSection(state.Document, master, context);
                state.MasterPage = master;
                switched = true;
            }
            if (info.BreakBefore)
            {
                if (!switched) state.Section.Blocks.Add(new PageBreak());
                if (IsEmpty(info.Paragraph)) return;
            }
            state.Section.Blocks.Add(info.Paragraph);
        }

        /// <summary>判断段落是否没有任何内容。Checks whether a paragraph carries no content.</summary>
        private static bool IsEmpty(Paragraph paragraph)
        {
            return paragraph.Inlines.Count == 0
                && string.IsNullOrEmpty(paragraph.BookmarkName)
                && string.IsNullOrEmpty(paragraph.ListMarker);
        }

        /// <summary>取得当前节，必要时创建默认节。Gets the current section, creating a default one when needed.</summary>
        private static Section EnsureSection(BodyState state, ParseContext context)
        {
            if (state.Section == null)
            {
                state.Section = CreateSection(state.Document, null, context);
                state.MasterPage = string.Empty;
            }
            return state.Section;
        }

        /// <summary>按主页面创建节：复制页面布局并解析页眉页脚内容。Creates a section from a master page: copies the page layout and parses the header/footer content.</summary>
        private static Section CreateSection(OfficeDocument document, string masterName, ParseContext context)
        {
            MasterPageInfo master = null;
            if (!string.IsNullOrEmpty(masterName))
            {
                context.Styles.Masters.TryGetValue(masterName, out master);
            }
            if (master == null) return document.AddSection();

            Section section = document.AddSection(context.Styles.BuildPageSetup(master.LayoutName));
            if (master.Header != null) ParseBlocks(section.Header.Blocks, master.Header, context, 0);
            if (master.Footer != null) ParseBlocks(section.Footer.Blocks, master.Footer, context, 0);
            return section;
        }

        /// <summary>解析一组块级内容（用于页眉、页脚与表格单元格）。Parses a block list (used for headers, footers and table cells).</summary>
        private static void ParseBlocks(List<Block> target, XElement container, ParseContext context, int depth)
        {
            if (target == null || container == null || depth > MaxBlockDepth) return;
            foreach (XElement child in container.Elements())
            {
                context.CountNode();
                XName name = child.Name;
                if (name == TextNs + "p" || name == TextNs + "h")
                {
                    target.Add(ParseParagraph(child, context).Paragraph);
                }
                else if (name == TableNs + "table")
                {
                    target.Add(ParseTable(child, context, depth + 1));
                }
                else if (name == TextNs + "section" || name == TextNs + "list" || name == TextNs + "list-item")
                {
                    ParseBlocks(target, child, context, depth + 1);
                }
                else if (name == TextNs + "page-break" || name == TextNs + "soft-page-break")
                {
                    target.Add(new PageBreak());
                }
            }
        }

        /// <summary>解析段落的结果：段落本身、主页面与分页属性。The parse result of a paragraph: the paragraph, its master page and page-break attribute.</summary>
        private sealed class ParagraphInfo
        {
            /// <summary>解析出的段落。The parsed paragraph.</summary>
            public Paragraph Paragraph;
            /// <summary>段落样式引用的主页面名。The master-page name referenced by the paragraph style.</summary>
            public string MasterPage;
            /// <summary>段落前是否分页。Whether the paragraph starts on a new page.</summary>
            public bool BreakBefore;
        }

        /// <summary>解析一个段落或标题。Parses one paragraph or heading.</summary>
        private static ParagraphInfo ParseParagraph(XElement element, ParseContext context)
        {
            ParagraphInfo info = new ParagraphInfo();
            Paragraph paragraph = new Paragraph();
            paragraph.LineSpacing = 1;

            ResolvedStyle style = context.Styles.Resolve(Attr(element, TextNs + "style-name"));
            Dictionary<XName, string> properties = style == null ? null : style.Paragraph;
            info.MasterPage = style == null ? null : style.MasterPage;
            ApplyParagraphProperties(paragraph, properties);
            info.BreakBefore = IsPageBreakBefore(properties);

            if (element.Name == TextNs + "h")
            {
                int level = ParseInt(Attr(element, TextNs + "outline-level"), 1);
                if (level < 1) level = 1;
                if (level > 10) level = 10;
                paragraph.HeadingLevel = level;
            }

            TextFormat baseFormat = context.Styles.BuildTextFormat(style == null ? null : style.Text, null);
            ParseInlines(element, paragraph, baseFormat, null, context, 0);
            info.Paragraph = paragraph;
            return info;
        }

        /// <summary>把段落属性（对齐、间距、缩进、行距、孤行控制）应用到段落。Applies paragraph properties (alignment, spacing, indents, line height, keep-with-next) to a paragraph.</summary>
        private static void ApplyParagraphProperties(Paragraph paragraph, Dictionary<XName, string> properties)
        {
            if (paragraph == null || properties == null) return;
            paragraph.Alignment = ParseAlignment(Get(properties, FoNs + "text-align"), ParagraphAlignment.Left);
            paragraph.SpaceBefore = Length(properties, FoNs + "margin-top", 0);
            paragraph.SpaceAfter = Length(properties, FoNs + "margin-bottom", 0);
            paragraph.LeftIndent = Length(properties, FoNs + "margin-left", 0);
            paragraph.RightIndent = Length(properties, FoNs + "margin-right", 0);
            paragraph.FirstLineIndent = Length(properties, FoNs + "text-indent", 0);
            string lineHeight = Get(properties, FoNs + "line-height");
            if (!string.IsNullOrEmpty(lineHeight))
            {
                paragraph.LineSpacing = ParseLineSpacing(lineHeight, paragraph.LineSpacing);
            }
            paragraph.KeepWithNext = IsAlways(Get(properties, FoNs + "keep-with-next"));
        }

        /// <summary>判断段落属性中是否设置了段前分页。Checks whether the paragraph properties request a page break before the paragraph.</summary>
        private static bool IsPageBreakBefore(Dictionary<XName, string> properties)
        {
            string value = Get(properties, FoNs + "break-before");
            return value != null && string.Equals(value.Trim(), "page", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>解析对齐值。Parses an alignment value.</summary>
        private static ParagraphAlignment ParseAlignment(string text, ParagraphAlignment fallback)
        {
            if (string.IsNullOrEmpty(text)) return fallback;
            string value = text.Trim();
            if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) return ParagraphAlignment.Center;
            if (string.Equals(value, "end", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)) return ParagraphAlignment.Right;
            if (string.Equals(value, "justify", StringComparison.OrdinalIgnoreCase)) return ParagraphAlignment.Justify;
            return ParagraphAlignment.Left;
        }

        /// <summary>解析行高：百分比或倍数，带单位的磅值按 11 磅正文换算。Parses a line height: a percentage or multiple; a value with a unit is converted using an 11-point body size.</summary>
        private static double ParseLineSpacing(string text, double fallback)
        {
            if (string.IsNullOrEmpty(text)) return fallback;
            string value = text.Trim();
            if (value.EndsWith("%", StringComparison.Ordinal))
            {
                double percent;
                if (double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out percent) && percent > 0)
                {
                    return percent / 100.0;
                }
                return fallback;
            }
            double multiple;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out multiple) && multiple > 0)
            {
                return multiple > 10 ? multiple / 11.0 : multiple;
            }
            double points = WordFormatUtil.ParseLengthPoints(value, 0);
            return points > 0 ? points / 11.0 : fallback;
        }

        // ---------------------------------------------------------------- inlines

        /// <summary>解析段落内的行内内容。Parses the inline content of a paragraph.</summary>
        private static void ParseInlines(XElement container, Paragraph paragraph, TextFormat format, string link, ParseContext context, int depth)
        {
            if (container == null || depth > MaxBlockDepth) return;
            foreach (XNode node in container.Nodes())
            {
                XText textNode = node as XText;
                if (textNode != null)
                {
                    AppendText(paragraph, NormalizeText(textNode.Value), format, link);
                    continue;
                }
                XElement element = node as XElement;
                if (element == null) continue;
                context.CountNode();
                XName name = element.Name;

                if (name == TextNs + "span")
                {
                    ResolvedStyle style = context.Styles.Resolve(Attr(element, TextNs + "style-name"));
                    TextFormat spanFormat = style == null ? format : context.Styles.BuildTextFormat(style.Text, format);
                    ParseInlines(element, paragraph, spanFormat, link, context, depth + 1);
                    continue;
                }
                if (name == TextNs + "a")
                {
                    ParseInlines(element, paragraph, format, Attr(element, XLinkNs + "href"), context, depth + 1);
                    continue;
                }
                if (name == TextNs + "s")
                {
                    int count = ParseInt(Attr(element, TextNs + "c"), 1);
                    if (count < 1) count = 1;
                    if (count > 10000) count = 10000;
                    AppendText(paragraph, new string(' ', count), format, link);
                    continue;
                }
                if (name == TextNs + "tab")
                {
                    AppendText(paragraph, "\t", format, link);
                    continue;
                }
                if (name == TextNs + "line-break")
                {
                    AppendText(paragraph, "\n", format, link);
                    continue;
                }
                if (name == TextNs + "bookmark")
                {
                    if (string.IsNullOrEmpty(paragraph.BookmarkName))
                    {
                        paragraph.BookmarkName = Attr(element, TextNs + "name");
                    }
                    continue;
                }
                if (name == TextNs + "bookmark-start")
                {
                    // 起点标记既还原成行内书签范围，也沿用段落书签名以兼容旧读取语义。
                    // The start marker restores the inline range and keeps filling the paragraph bookmark name
                    // for compatibility with the previous reading behaviour.
                    string started = Attr(element, TextNs + "name");
                    if (!string.IsNullOrEmpty(started))
                    {
                        if (string.IsNullOrEmpty(paragraph.BookmarkName)) paragraph.BookmarkName = started;
                        paragraph.Inlines.Add(new BookmarkStart(started));
                    }
                    continue;
                }
                if (name == TextNs + "bookmark-end")
                {
                    string ended = Attr(element, TextNs + "name");
                    if (!string.IsNullOrEmpty(ended)) paragraph.Inlines.Add(new BookmarkEnd(ended));
                    continue;
                }
                if (name == TextNs + "note" || name == TextNs + "soft-page-break")
                {
                    continue;
                }
                if (name == DrawNs + "frame")
                {
                    ImageRun image = ParseImageFrame(element, context);
                    if (image != null) paragraph.Inlines.Add(image);
                    continue;
                }
                if (name == TextNs + "page-number")
                {
                    paragraph.Inlines.Add(new FieldRun("PAGE", "1", format.Clone()));
                    continue;
                }
                if (name == TextNs + "date" || name == TextNs + "time")
                {
                    string display = element.Value;
                    if (string.IsNullOrEmpty(display))
                    {
                        display = Attr(element, TextNs + "date-value") ?? Attr(element, TextNs + "time-value") ?? string.Empty;
                    }
                    paragraph.Inlines.Add(new FieldRun(name == TextNs + "date" ? "DATE" : "TIME", display, format.Clone()));
                    continue;
                }
                ParseInlines(element, paragraph, format, link, context, depth + 1);
            }
        }

        /// <summary>追加文本并与相邻同格式文本合并。Appends text, merging it with an adjacent run of the same format.</summary>
        private static void AppendText(Paragraph paragraph, string text, TextFormat format, string link)
        {
            if (string.IsNullOrEmpty(text)) return;
            if (paragraph.Inlines.Count > 0)
            {
                TextRun last = paragraph.Inlines[paragraph.Inlines.Count - 1] as TextRun;
                if (last != null && SameFormat(last.Format, format) && string.Equals(last.Link, link, StringComparison.Ordinal))
                {
                    last.Text += text;
                    return;
                }
            }
            TextRun run = new TextRun(text, format == null ? new TextFormat() : format.Clone());
            run.Link = link;
            paragraph.Inlines.Add(run);
        }

        /// <summary>判断两份文本格式是否完全相同。Checks whether two text formats are identical.</summary>
        private static bool SameFormat(TextFormat a, TextFormat b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null) return false;
            return string.Equals(a.GetAsciiFontFamily(), b.GetAsciiFontFamily(), StringComparison.Ordinal)
                && string.Equals(a.GetEastAsiaFontFamily(), b.GetEastAsiaFontFamily(), StringComparison.Ordinal)
                && Math.Abs(a.FontSize - b.FontSize) < 0.001
                && a.Bold == b.Bold
                && a.Italic == b.Italic
                && a.Underline == b.Underline
                && a.Strikeout == b.Strikeout
                && a.Superscript == b.Superscript
                && a.Subscript == b.Subscript
                && a.Color.A == b.Color.A
                && a.Color.R == b.Color.R
                && a.Color.G == b.Color.G
                && a.Color.B == b.Color.B;
        }

        /// <summary>把 XML 文本节点中的换行/制表符折叠为单个空格（ODF 的空白语义）。Collapses line breaks and tabs in an XML text node into single spaces (the ODF whitespace semantics).</summary>
        private static string NormalizeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            bool needsNormalization = false;
            for (int i = 0; i < text.Length; i++)
            {
                char value = text[i];
                if (value == '\r' || value == '\n' || value == '\t')
                {
                    needsNormalization = true;
                    break;
                }
            }
            if (!needsNormalization) return text;

            StringBuilder builder = new StringBuilder(text.Length);
            bool pendingSpace = false;
            for (int i = 0; i < text.Length; i++)
            {
                char value = text[i];
                if (value == '\r' || value == '\n' || value == '\t' || value == ' ')
                {
                    if (builder.Length > 0) pendingSpace = true;
                    continue;
                }
                if (pendingSpace)
                {
                    builder.Append(' ');
                    pendingSpace = false;
                }
                builder.Append(value);
            }
            return builder.ToString();
        }

        /// <summary>解析 draw:frame 中的内嵌图片（二进制数据或包内引用）。Parses the embedded image of a draw:frame (binary data or a package reference).</summary>
        private static ImageRun ParseImageFrame(XElement frame, ParseContext context)
        {
            double width = LengthAttr(frame, SvgNs + "width", 0);
            double height = LengthAttr(frame, SvgNs + "height", 0);
            foreach (XElement image in frame.Descendants(DrawNs + "image"))
            {
                byte[] data = null;
                foreach (XElement binary in image.Descendants(OfficeNs + "binary-data"))
                {
                    data = context.Parts.ReadBinaryImage(binary.Value);
                    if (data != null) break;
                }
                if (data == null)
                {
                    data = context.Parts.ReadImagePart(Attr(image, XLinkNs + "href"));
                }
                if (data == null) continue;
                return new ImageRun(new OfficeImage(data), width, height);
            }
            return null;
        }

        // ---------------------------------------------------------------- tables

        /// <summary>表格网格状态：已覆盖单元格与列数统计。Table grid state: covered cells and column statistics.</summary>
        private sealed class TableGridState
        {
            /// <summary>被跨行/跨列覆盖的网格槽位。The grid slots covered by row/column spans.</summary>
            public readonly HashSet<long> Covered = new HashSet<long>();
            /// <summary>当前行号。The current row index.</summary>
            public int RowIndex;
            /// <summary>已出现的最大列数。The largest column index seen so far.</summary>
            public int MaxColumns;
            /// <summary>从单元格样式读到的内边距。The padding read from cell styles.</summary>
            public double CellPadding;
        }

        /// <summary>解析一个表格。Parses one table.</summary>
        private static Table ParseTable(XElement element, ParseContext context, int depth)
        {
            Table table = new Table();
            table.Borders = TableBorders.None();

            ResolvedStyle tableStyle = context.Styles.Resolve(Attr(element, TableNs + "style-name"));
            if (tableStyle != null) ApplyTableProperties(table, tableStyle.Table);
            table.Alignment = ParseTableAlignment(Attr(element, TableNs + "align"), table.Alignment);

            List<double> widths = new List<double>();
            CollectColumns(element, widths, context, 0);

            TableGridState grid = new TableGridState();
            CollectRows(element, table, grid, context, false, depth);
            if (grid.CellPadding > 0) table.CellPadding = grid.CellPadding;

            int columnCount = Math.Max(widths.Count, grid.MaxColumns);
            if (columnCount <= 0) columnCount = 1;
            for (int i = 0; i < columnCount; i++)
            {
                table.Columns.Add(new TableColumn(i < widths.Count ? widths[i] : 0));
            }
            return table;
        }

        /// <summary>收集表格列定义（展开重复列与列组），读取列宽。Collects table column definitions (expanding repeats and groups) and reads column widths.</summary>
        private static void CollectColumns(XElement parent, List<double> widths, ParseContext context, int depth)
        {
            if (parent == null || depth > 8) return;
            foreach (XElement child in parent.Elements())
            {
                if (child.Name == TableNs + "table-column")
                {
                    int repeat = ParseRepeat(child);
                    double width = 0;
                    ResolvedStyle style = context.Styles.Resolve(Attr(child, TableNs + "style-name"));
                    if (style != null) width = Math.Max(0, Length(style.Column, StyleNs + "column-width", 0));
                    for (int i = 0; i < repeat && widths.Count < MaxColumns; i++)
                    {
                        widths.Add(width);
                    }
                    if (widths.Count >= MaxColumns) return;
                }
                else if (child.Name == TableNs + "table-column-group"
                    || child.Name == TableNs + "table-columns"
                    || child.Name == TableNs + "table-header-columns")
                {
                    CollectColumns(child, widths, context, depth + 1);
                }
            }
        }

        /// <summary>收集表格行（展开行组与表头行组）并逐行解析单元格。Collects table rows (expanding row groups and header-row groups) and parses cells row by row.</summary>
        private static void CollectRows(XElement parent, Table table, TableGridState grid, ParseContext context, bool header, int depth)
        {
            if (parent == null || depth > 8) return;
            foreach (XElement child in parent.Elements())
            {
                context.CountNode();
                if (child.Name == TableNs + "table-row")
                {
                    ParseRow(child, table, grid, context, header, depth);
                }
                else if (child.Name == TableNs + "table-header-rows")
                {
                    CollectRows(child, table, grid, context, true, depth + 1);
                }
                else if (child.Name == TableNs + "table-rows" || child.Name == TableNs + "table-row-group")
                {
                    CollectRows(child, table, grid, context, header, depth + 1);
                }
            }
        }

        /// <summary>解析一行：按网格位置还原跨行跨列单元格。Parses one row, rebuilding row/column spans from grid positions.</summary>
        private static void ParseRow(XElement rowElement, Table table, TableGridState grid, ParseContext context, bool header, int depth)
        {
            TableRow row = new TableRow();
            row.IsHeader = header;
            ResolvedStyle rowStyle = context.Styles.Resolve(Attr(rowElement, TableNs + "style-name"));
            if (rowStyle != null) row.Height = Math.Max(0, Length(rowStyle.Row, StyleNs + "row-height", 0));

            int column = 0;
            foreach (XElement cellElement in rowElement.Elements())
            {
                context.CountNode();
                if (cellElement.Name == TableNs + "table-cell")
                {
                    int repeat = ParseRepeat(cellElement);
                    for (int rep = 0; rep < repeat; rep++)
                    {
                        while (column < MaxColumns && grid.Covered.Contains(Key(grid.RowIndex, column))) column++;
                        if (column >= MaxColumns) break;
                        TableCell cell = ParseCell(cellElement, grid, context, depth + 1);
                        int colSpan = Clamp(ParseInt(Attr(cellElement, TableNs + "number-columns-spanned"), 1), 1, MaxColumns - column);
                        int rowSpan = Clamp(ParseInt(Attr(cellElement, TableNs + "number-rows-spanned"), 1), 1, 4096);
                        cell.ColSpan = colSpan;
                        cell.RowSpan = rowSpan;
                        for (int dr = 0; dr < rowSpan; dr++)
                        {
                            for (int dc = 0; dc < colSpan; dc++)
                            {
                                grid.Covered.Add(Key(grid.RowIndex + dr, column + dc));
                            }
                        }
                        row.Cells.Add(cell);
                        column += colSpan;
                        if (column > grid.MaxColumns) grid.MaxColumns = column;
                    }
                }
                else if (cellElement.Name == TableNs + "covered-table-cell")
                {
                    // 覆盖单元格没有内容，其网格位置已由起始单元格的跨行跨列登记，无需再次推进列号。
                    // A covered cell carries no content; its grid slots are already registered by the origin
                    // cell's spans, so the column cursor must not advance again.
                    continue;
                }
            }

            table.Rows.Add(row);
            grid.RowIndex++;
        }

        /// <summary>解析单元格内容与格式（底色、边框、垂直对齐、内边距）。Parses cell content and formatting (shading, borders, vertical alignment, padding).</summary>
        private static TableCell ParseCell(XElement cellElement, TableGridState grid, ParseContext context, int depth)
        {
            TableCell cell = new TableCell();
            ResolvedStyle style = context.Styles.Resolve(Attr(cellElement, TableNs + "style-name"));
            if (style != null)
            {
                Dictionary<XName, string> properties = style.Cell;
                OfficeColor background;
                if (TryParseColor(Get(properties, FoNs + "background-color"), out background))
                {
                    cell.BackgroundColor = background;
                }
                cell.Borders = ParseCellBorders(properties);
                double padding = Length(properties, FoNs + "padding", 0);
                if (padding > grid.CellPadding) grid.CellPadding = padding;
                string vertical = Get(properties, StyleNs + "vertical-align");
                if (vertical != null)
                {
                    if (string.Equals(vertical, "middle", StringComparison.OrdinalIgnoreCase)) cell.VerticalAlignment = VerticalAlignment.Middle;
                    else if (string.Equals(vertical, "bottom", StringComparison.OrdinalIgnoreCase)) cell.VerticalAlignment = VerticalAlignment.Bottom;
                }
            }
            ParseBlocks(cell.Blocks, cellElement, context, depth);
            return cell;
        }

        /// <summary>读取单元格四条边框；全部缺失时返回 null。Reads the four cell borders; returns null when every edge is absent.</summary>
        private static TableCellBorders ParseCellBorders(Dictionary<XName, string> properties)
        {
            if (properties == null) return null;
            BorderLine shorthand = ParseBorder(Get(properties, FoNs + "border"));
            BorderLine top = BorderFrom(properties, FoNs + "border-top", shorthand);
            BorderLine bottom = BorderFrom(properties, FoNs + "border-bottom", shorthand);
            BorderLine left = BorderFrom(properties, FoNs + "border-left", shorthand);
            BorderLine right = BorderFrom(properties, FoNs + "border-right", shorthand);
            if (top == null && bottom == null && left == null && right == null) return null;
            TableCellBorders borders = new TableCellBorders();
            borders.Top = top;
            borders.Bottom = bottom;
            borders.Left = left;
            borders.Right = right;
            return borders;
        }

        /// <summary>把表格样式属性应用到表格（对齐与外框）。Applies table style properties (alignment and outer borders) to a table.</summary>
        private static void ApplyTableProperties(Table table, Dictionary<XName, string> properties)
        {
            if (properties == null) return;
            table.Alignment = ParseTableAlignment(Get(properties, TableNs + "align"), table.Alignment);
            BorderLine shorthand = ParseBorder(Get(properties, FoNs + "border"));
            BorderLine top = BorderFrom(properties, FoNs + "border-top", shorthand);
            BorderLine bottom = BorderFrom(properties, FoNs + "border-bottom", shorthand);
            BorderLine left = BorderFrom(properties, FoNs + "border-left", shorthand);
            BorderLine right = BorderFrom(properties, FoNs + "border-right", shorthand);
            if (top != null) table.Borders.Top = top;
            if (bottom != null) table.Borders.Bottom = bottom;
            if (left != null) table.Borders.Left = left;
            if (right != null) table.Borders.Right = right;
        }

        /// <summary>解析表格对齐值。Parses a table alignment value.</summary>
        private static TableAlignment ParseTableAlignment(string text, TableAlignment fallback)
        {
            if (string.IsNullOrEmpty(text)) return fallback;
            string value = text.Trim();
            if (string.Equals(value, "center", StringComparison.OrdinalIgnoreCase)) return TableAlignment.Center;
            if (string.Equals(value, "right", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "end", StringComparison.OrdinalIgnoreCase)) return TableAlignment.Right;
            if (string.Equals(value, "left", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "start", StringComparison.OrdinalIgnoreCase)) return TableAlignment.Left;
            if (string.Equals(value, "margins", StringComparison.OrdinalIgnoreCase)) return fallback;
            return fallback;
        }

        /// <summary>解析一次元素重复次数并施加安全上限。Parses an element repeat count and applies the safety cap.</summary>
        private static int ParseRepeat(XElement element)
        {
            int repeat = ParseInt(Attr(element, TableNs + "number-columns-repeated"), 1);
            if (repeat < 1) repeat = 1;
            if (repeat > MaxRepeat) repeat = MaxRepeat;
            return repeat;
        }

        /// <summary>构造跨行跨列的网格键。Builds a grid key for a row/column pair.</summary>
        private static long Key(int row, int column)
        {
            return ((long)row << 32) | (uint)column;
        }

        // ---------------------------------------------------------------- styles

        /// <summary>样式定义：名称、族、父样式与各家族的属性集合。A style definition: name, family, parent and the property sets of each family.</summary>
        private sealed class StyleDef
        {
            /// <summary>样式名。The style name.</summary>
            public string Name;
            /// <summary>样式族（paragraph/text/table-cell 等）。The style family (paragraph/text/table-cell and so on).</summary>
            public string Family;
            /// <summary>父样式名。The parent style name.</summary>
            public string Parent;
            /// <summary>段落样式引用的主页面名。The master-page name referenced by a paragraph style.</summary>
            public string MasterPage;
            /// <summary>段落属性。The paragraph properties.</summary>
            public readonly Dictionary<XName, string> Paragraph = new Dictionary<XName, string>();
            /// <summary>文本属性。The text properties.</summary>
            public readonly Dictionary<XName, string> Text = new Dictionary<XName, string>();
            /// <summary>单元格属性。The table-cell properties.</summary>
            public readonly Dictionary<XName, string> Cell = new Dictionary<XName, string>();
            /// <summary>表行属性。The table-row properties.</summary>
            public readonly Dictionary<XName, string> Row = new Dictionary<XName, string>();
            /// <summary>表列属性。The table-column properties.</summary>
            public readonly Dictionary<XName, string> Column = new Dictionary<XName, string>();
            /// <summary>表格属性。The table properties.</summary>
            public readonly Dictionary<XName, string> Table = new Dictionary<XName, string>();
        }

        /// <summary>合并后的样式：已按父链叠加全部属性。A resolved style with all properties merged along the parent chain.</summary>
        private sealed class ResolvedStyle
        {
            /// <summary>合并后的段落属性。The merged paragraph properties.</summary>
            public readonly Dictionary<XName, string> Paragraph = new Dictionary<XName, string>();
            /// <summary>合并后的文本属性。The merged text properties.</summary>
            public readonly Dictionary<XName, string> Text = new Dictionary<XName, string>();
            /// <summary>合并后的单元格属性。The merged table-cell properties.</summary>
            public readonly Dictionary<XName, string> Cell = new Dictionary<XName, string>();
            /// <summary>合并后的表行属性。The merged table-row properties.</summary>
            public readonly Dictionary<XName, string> Row = new Dictionary<XName, string>();
            /// <summary>合并后的表列属性。The merged table-column properties.</summary>
            public readonly Dictionary<XName, string> Column = new Dictionary<XName, string>();
            /// <summary>合并后的表格属性。The merged table properties.</summary>
            public readonly Dictionary<XName, string> Table = new Dictionary<XName, string>();
            /// <summary>沿父链继承的主页面名。The master-page name inherited along the parent chain.</summary>
            public string MasterPage;
        }

        /// <summary>页面布局信息。Page-layout information.</summary>
        private sealed class LayoutInfo
        {
            /// <summary>页面布局元素。The page-layout element.</summary>
            public XElement Element;
        }

        /// <summary>主页面信息。Master-page information.</summary>
        private sealed class MasterPageInfo
        {
            /// <summary>页面布局名。The page-layout name.</summary>
            public string LayoutName;
            /// <summary>页眉元素。The header element.</summary>
            public XElement Header;
            /// <summary>页脚元素。The footer element.</summary>
            public XElement Footer;
        }

        /// <summary>ODF 样式索引：收集样式、页面布局与主页面，并按父链解析合并属性。The ODF style index: collects styles, page layouts and master pages, resolving merged properties along parent chains.</summary>
        private sealed class StyleIndex
        {
            private readonly Dictionary<string, StyleDef> _styles = new Dictionary<string, StyleDef>(StringComparer.Ordinal);
            private readonly Dictionary<string, StyleDef> _defaults = new Dictionary<string, StyleDef>(StringComparer.Ordinal);
            private readonly Dictionary<string, ResolvedStyle> _resolved = new Dictionary<string, ResolvedStyle>(StringComparer.Ordinal);
            private readonly Dictionary<string, LayoutInfo> _layouts = new Dictionary<string, LayoutInfo>(StringComparer.Ordinal);
            private readonly Dictionary<string, MasterPageInfo> _masters = new Dictionary<string, MasterPageInfo>(StringComparer.Ordinal);
            private TextFormat _defaultFormat;

            /// <summary>主页面名到主页面信息的映射。Master-page name to master-page information.</summary>
            public Dictionary<string, MasterPageInfo> Masters { get { return _masters; } }

            /// <summary>收集一个 XML 文档中的样式、页面布局与主页面。Collects the styles, page layouts and master pages of one XML document.</summary>
            public void Collect(XElement root)
            {
                if (root == null) return;
                foreach (XElement container in root.Elements())
                {
                    if (container.Name == OfficeNs + "styles" || container.Name == OfficeNs + "automatic-styles")
                    {
                        foreach (XElement child in container.Elements())
                        {
                            if (child.Name == StyleNs + "style") AddStyle(child);
                            else if (child.Name == StyleNs + "default-style") AddDefault(child);
                            else if (child.Name == StyleNs + "page-layout") AddLayout(child);
                            else if (child.Name == StyleNs + "master-page") AddMaster(child);
                        }
                    }
                    else if (container.Name == OfficeNs + "master-styles")
                    {
                        foreach (XElement child in container.Elements())
                        {
                            if (child.Name == StyleNs + "master-page") AddMaster(child);
                        }
                    }
                }
            }

            /// <summary>登记一个样式（同名样式以先出现者为准）。Registers a style (the first definition wins on duplicate names).</summary>
            private void AddStyle(XElement element)
            {
                string name = Attr(element, StyleNs + "name");
                if (string.IsNullOrEmpty(name) || _styles.ContainsKey(name)) return;
                StyleDef def = new StyleDef();
                def.Name = name;
                def.Family = Attr(element, StyleNs + "family") ?? string.Empty;
                def.Parent = Attr(element, StyleNs + "parent-style-name");
                def.MasterPage = Attr(element, StyleNs + "master-page-name");
                ReadProperties(element, def);
                _styles[name] = def;
            }

            /// <summary>登记一个默认样式。Registers a default style.</summary>
            private void AddDefault(XElement element)
            {
                string family = Attr(element, StyleNs + "family");
                if (string.IsNullOrEmpty(family) || _defaults.ContainsKey(family)) return;
                StyleDef def = new StyleDef();
                def.Name = "$" + family;
                def.Family = family;
                ReadProperties(element, def);
                _defaults[family] = def;
            }

            /// <summary>登记一个页面布局。Registers a page layout.</summary>
            private void AddLayout(XElement element)
            {
                string name = Attr(element, StyleNs + "name");
                if (string.IsNullOrEmpty(name) || _layouts.ContainsKey(name)) return;
                LayoutInfo info = new LayoutInfo();
                info.Element = element;
                _layouts[name] = info;
            }

            /// <summary>登记一个主页面。Registers a master page.</summary>
            private void AddMaster(XElement element)
            {
                string name = Attr(element, StyleNs + "name");
                if (string.IsNullOrEmpty(name) || _masters.ContainsKey(name)) return;
                MasterPageInfo info = new MasterPageInfo();
                info.LayoutName = Attr(element, StyleNs + "page-layout-name");
                info.Header = element.Element(StyleNs + "header");
                info.Footer = element.Element(StyleNs + "footer");
                _masters[name] = info;
            }

            /// <summary>把样式元素的直接属性子节点复制到对应属性集合。Copies the direct property children of a style element into the matching property sets.</summary>
            private static void ReadProperties(XElement element, StyleDef def)
            {
                foreach (XElement child in element.Elements())
                {
                    Dictionary<XName, string> target = null;
                    if (child.Name == StyleNs + "paragraph-properties") target = def.Paragraph;
                    else if (child.Name == StyleNs + "text-properties") target = def.Text;
                    else if (child.Name == StyleNs + "table-cell-properties") target = def.Cell;
                    else if (child.Name == StyleNs + "table-row-properties") target = def.Row;
                    else if (child.Name == StyleNs + "table-column-properties") target = def.Column;
                    else if (child.Name == StyleNs + "table-properties") target = def.Table;
                    if (target == null) continue;
                    foreach (XAttribute attribute in child.Attributes())
                    {
                        target[attribute.Name] = attribute.Value;
                    }
                }
            }

            /// <summary>解析样式，沿父链自底向上合并属性并缓存结果。Resolves a style, merging properties bottom-up along the parent chain and caching the result.</summary>
            public ResolvedStyle Resolve(string name)
            {
                if (string.IsNullOrEmpty(name)) return null;
                ResolvedStyle cached;
                if (_resolved.TryGetValue(name, out cached)) return cached;
                StyleDef def;
                if (!_styles.TryGetValue(name, out def)) return null;

                List<StyleDef> chain = new List<StyleDef>();
                HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
                StyleDef current = def;
                while (current != null && chain.Count < MaxStyleDepth && seen.Add(current.Name))
                {
                    chain.Add(current);
                    if (string.IsNullOrEmpty(current.Parent)) break;
                    StyleDef parent;
                    if (!_styles.TryGetValue(current.Parent, out parent)) break;
                    current = parent;
                }

                ResolvedStyle resolved = new ResolvedStyle();
                StyleDef defaultStyle;
                if (!string.IsNullOrEmpty(def.Family) && _defaults.TryGetValue(def.Family, out defaultStyle))
                {
                    MergeInto(resolved, defaultStyle);
                }
                for (int i = chain.Count - 1; i >= 0; i--)
                {
                    MergeInto(resolved, chain[i]);
                    if (!string.IsNullOrEmpty(chain[i].MasterPage)) resolved.MasterPage = chain[i].MasterPage;
                }
                _resolved[name] = resolved;
                return resolved;
            }

            /// <summary>把一份样式定义合并到解析结果上（子样式覆盖父样式）。Merges one style definition into the resolved result (child overrides parent).</summary>
            private static void MergeInto(ResolvedStyle target, StyleDef source)
            {
                CopyProperties(source.Paragraph, target.Paragraph);
                CopyProperties(source.Text, target.Text);
                CopyProperties(source.Cell, target.Cell);
                CopyProperties(source.Row, target.Row);
                CopyProperties(source.Column, target.Column);
                CopyProperties(source.Table, target.Table);
            }

            /// <summary>复制属性集合。Copies a property set.</summary>
            private static void CopyProperties(Dictionary<XName, string> source, Dictionary<XName, string> target)
            {
                foreach (KeyValuePair<XName, string> pair in source)
                {
                    target[pair.Key] = pair.Value;
                }
            }

            /// <summary>取得默认段落文本格式（缓存）。Gets the default paragraph text format (cached).</summary>
            private TextFormat DefaultFormat()
            {
                if (_defaultFormat == null)
                {
                    ResolvedStyle resolved = ResolveDefault("paragraph");
                    _defaultFormat = BuildFormatCore(resolved == null ? null : resolved.Text, new TextFormat());
                }
                return _defaultFormat;
            }

            /// <summary>解析某个族的默认样式。Resolves the default style of a family.</summary>
            private ResolvedStyle ResolveDefault(string family)
            {
                string key = "\u0001" + family;
                ResolvedStyle cached;
                if (_resolved.TryGetValue(key, out cached)) return cached;
                ResolvedStyle resolved = new ResolvedStyle();
                StyleDef def;
                if (_defaults.TryGetValue(family, out def)) MergeInto(resolved, def);
                _resolved[key] = resolved;
                return resolved;
            }

            /// <summary>基于基础格式叠加文本属性，生成新的文本格式。Builds a new text format by overlaying text properties on a base format.</summary>
            public TextFormat BuildTextFormat(Dictionary<XName, string> properties, TextFormat baseFormat)
            {
                return BuildFormatCore(properties, baseFormat ?? DefaultFormat());
            }

            /// <summary>生成页面设置（每次返回独立实例）。Builds a page setup (a fresh instance on every call).</summary>
            public PageSetup BuildPageSetup(string layoutName)
            {
                LayoutInfo info;
                if (string.IsNullOrEmpty(layoutName) || !_layouts.TryGetValue(layoutName, out info)) return new PageSetup();
                return ParsePageSetup(info.Element);
            }

            /// <summary>把文本属性叠加到基础格式上。Overlays text properties on a base format.</summary>
            private static TextFormat BuildFormatCore(Dictionary<XName, string> properties, TextFormat baseFormat)
            {
                TextFormat format = baseFormat == null ? new TextFormat() : baseFormat.Clone();
                if (properties == null) return format;

                string value;
                if (properties.TryGetValue(FoNs + "font-family", out value) && !string.IsNullOrEmpty(value))
                {
                    string family = Unquote(value);
                    format.FontFamily = family;
                    format.AsciiFontFamily = family;
                }
                if (properties.TryGetValue(StyleNs + "font-family-asian", out value) && !string.IsNullOrEmpty(value))
                {
                    format.EastAsiaFontFamily = Unquote(value);
                }
                if (properties.TryGetValue(FoNs + "font-size", out value))
                {
                    double size = WordFormatUtil.ParseLengthPoints(value, 0);
                    if (size > 0) format.FontSize = size;
                }
                if (properties.TryGetValue(FoNs + "font-weight", out value)) format.Bold = IsBold(value);
                if (properties.TryGetValue(FoNs + "font-style", out value)) format.Italic = IsItalic(value);
                if (properties.TryGetValue(StyleNs + "text-underline-style", out value)) format.Underline = !IsNone(value);
                if (properties.TryGetValue(StyleNs + "text-line-through-style", out value)) format.Strikeout = !IsNone(value);
                if (properties.TryGetValue(StyleNs + "text-position", out value))
                {
                    string position = value == null ? string.Empty : value.Trim().ToLowerInvariant();
                    format.Superscript = position.StartsWith("super", StringComparison.Ordinal);
                    format.Subscript = !format.Superscript && position.StartsWith("sub", StringComparison.Ordinal);
                }
                if (properties.TryGetValue(FoNs + "color", out value))
                {
                    OfficeColor color;
                    if (TryParseColor(value, out color)) format.Color = color;
                }
                return format;
            }
        }

        // ---------------------------------------------------------------- page setup

        /// <summary>从页面布局元素解析页面设置。Parses a page setup from a page-layout element.</summary>
        private static PageSetup ParsePageSetup(XElement layout)
        {
            PageSetup setup = new PageSetup();
            XElement properties = layout.Element(StyleNs + "page-layout-properties");
            if (properties != null)
            {
                setup.Width = LengthAttr(properties, FoNs + "page-width", setup.Width);
                setup.Height = LengthAttr(properties, FoNs + "page-height", setup.Height);
                setup.MarginTop = LengthAttr(properties, FoNs + "margin-top", setup.MarginTop);
                setup.MarginBottom = LengthAttr(properties, FoNs + "margin-bottom", setup.MarginBottom);
                setup.MarginLeft = LengthAttr(properties, FoNs + "margin-left", setup.MarginLeft);
                setup.MarginRight = LengthAttr(properties, FoNs + "margin-right", setup.MarginRight);
                string orientation = Attr(properties, StyleNs + "print-orientation");
                if (!string.IsNullOrEmpty(orientation))
                {
                    bool landscape = string.Equals(orientation.Trim(), "landscape", StringComparison.OrdinalIgnoreCase);
                    if (landscape != (setup.Width > setup.Height) && Math.Abs(setup.Width - setup.Height) > 0.01)
                    {
                        double width = setup.Width;
                        setup.Width = setup.Height;
                        setup.Height = width;
                    }
                }
            }
            setup.HeaderDistance = HeaderFooterDistance(layout, "header-style", setup.HeaderDistance);
            setup.FooterDistance = HeaderFooterDistance(layout, "footer-style", setup.FooterDistance);
            return setup;
        }

        /// <summary>读取页眉/页脚样式中的最小高度（作为页面设置中的距离）。Reads the minimum height of a header/footer style (used as the distance in the page setup).</summary>
        private static double HeaderFooterDistance(XElement layout, string styleName, double fallback)
        {
            XElement style = layout.Element(StyleNs + styleName);
            if (style == null) return fallback;
            XElement properties = style.Element(StyleNs + "header-footer-properties");
            if (properties == null) return fallback;
            return LengthAttr(properties, FoNs + "min-height", fallback);
        }

        // ---------------------------------------------------------------- formatting helpers

        /// <summary>解析边框文本，如 "0.5pt solid #000000"；空值返回 null，"none" 返回无边框线。Parses border text such as "0.5pt solid #000000"; empty returns null and "none" returns a borderless line.</summary>
        private static BorderLine ParseBorder(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
            string value = text.Trim();
            if (value.Length == 0) return null;
            string[] parts = value.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return null;
            if (string.Equals(parts[0], "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(parts[0], "hidden", StringComparison.OrdinalIgnoreCase))
            {
                return new BorderLine();
            }

            BorderLine line = new BorderLine(BorderStyle.Single, 0.5, OfficeColor.Black);
            for (int i = 0; i < parts.Length; i++)
            {
                string part = parts[i];
                if (part.Length == 0) continue;
                if (part[0] == '#')
                {
                    OfficeColor color;
                    if (TryParseColor(part, out color)) line.Color = color;
                    continue;
                }
                double width = WordFormatUtil.ParseLengthPoints(part, -1);
                if (width >= 0)
                {
                    line.Width = width;
                    continue;
                }
                if (string.Equals(part, "solid", StringComparison.OrdinalIgnoreCase)) line.Style = BorderStyle.Single;
                else if (string.Equals(part, "dashed", StringComparison.OrdinalIgnoreCase)) line.Style = BorderStyle.Dashed;
                else if (string.Equals(part, "dotted", StringComparison.OrdinalIgnoreCase)) line.Style = BorderStyle.Dotted;
                else if (string.Equals(part, "double", StringComparison.OrdinalIgnoreCase)) line.Style = BorderStyle.Double;
                else if (string.Equals(part, "none", StringComparison.OrdinalIgnoreCase)) line.Style = BorderStyle.None;
            }
            if (line.Width <= 0) line.Width = 0.5;
            return line;
        }

        /// <summary>读取一条边框属性；缺失时回退到简写边框。Reads one border property, falling back to the shorthand border when absent.</summary>
        private static BorderLine BorderFrom(Dictionary<XName, string> properties, XName name, BorderLine fallback)
        {
            string text = Get(properties, name);
            if (string.IsNullOrEmpty(text)) return fallback;
            BorderLine line = ParseBorder(text);
            return line ?? fallback;
        }

        /// <summary>解析 #RGB 或 #RRGGBB 颜色。Parses a #RGB or #RRGGBB colour.</summary>
        private static bool TryParseColor(string text, out OfficeColor color)
        {
            color = OfficeColor.Black;
            if (string.IsNullOrEmpty(text)) return false;
            string value = text.Trim();
            if (value.Length < 4 || value[0] != '#') return false;
            string hex = value.Substring(1);
            if (hex.Length == 3)
            {
                int red = ParseHexDigit(hex[0]);
                int green = ParseHexDigit(hex[1]);
                int blue = ParseHexDigit(hex[2]);
                if (red < 0 || green < 0 || blue < 0) return false;
                color = new OfficeColor((byte)(red * 17), (byte)(green * 17), (byte)(blue * 17));
                return true;
            }
            if (hex.Length == 6)
            {
                int red = ParseHexByte(hex.Substring(0, 2));
                int green = ParseHexByte(hex.Substring(2, 2));
                int blue = ParseHexByte(hex.Substring(4, 2));
                if (red < 0 || green < 0 || blue < 0) return false;
                color = new OfficeColor((byte)red, (byte)green, (byte)blue);
                return true;
            }
            return false;
        }

        /// <summary>解析一个十六进制数字。Parses one hexadecimal digit.</summary>
        private static int ParseHexDigit(char value)
        {
            if (value >= '0' && value <= '9') return value - '0';
            if (value >= 'a' && value <= 'f') return value - 'a' + 10;
            if (value >= 'A' && value <= 'F') return value - 'A' + 10;
            return -1;
        }

        /// <summary>解析两位十六进制字节。Parses a two-digit hexadecimal byte.</summary>
        private static int ParseHexByte(string text)
        {
            int high = ParseHexDigit(text[0]);
            int low = ParseHexDigit(text[1]);
            if (high < 0 || low < 0) return -1;
            return high * 16 + low;
        }

        /// <summary>判断字重是否为粗体（bold 或 600 及以上）。Checks whether a font weight is bold (bold or 600 and above).</summary>
        private static bool IsBold(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            string weight = value.Trim();
            if (string.Equals(weight, "bold", StringComparison.OrdinalIgnoreCase)
                || string.Equals(weight, "bolder", StringComparison.OrdinalIgnoreCase)) return true;
            int numeric;
            return int.TryParse(weight, NumberStyles.Integer, CultureInfo.InvariantCulture, out numeric) && numeric >= 600;
        }

        /// <summary>判断字形是否为倾斜。Checks whether a font style is italic or oblique.</summary>
        private static bool IsItalic(string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            string style = value.Trim();
            return string.Equals(style, "italic", StringComparison.OrdinalIgnoreCase)
                || string.Equals(style, "oblique", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>判断装饰值是否为 "none" 或空。Checks whether a decoration value is "none" or empty.</summary>
        private static bool IsNone(string value)
        {
            return string.IsNullOrEmpty(value) || string.Equals(value.Trim(), "none", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>判断值是否为 "always"。Checks whether a value is "always".</summary>
        private static bool IsAlways(string value)
        {
            return value != null && string.Equals(value.Trim(), "always", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>去掉字体名外围的单引号或双引号。Strips surrounding single or double quotes from a font-family name.</summary>
        private static string Unquote(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            string trimmed = value.Trim();
            if (trimmed.Length >= 2)
            {
                char first = trimmed[0];
                char last = trimmed[trimmed.Length - 1];
                if ((first == '\'' && last == '\'') || (first == '"' && last == '"'))
                {
                    return trimmed.Substring(1, trimmed.Length - 2);
                }
            }
            return trimmed;
        }

        /// <summary>读取属性值。Reads an attribute value.</summary>
        private static string Attr(XElement element, XName name)
        {
            if (element == null) return null;
            XAttribute attribute = element.Attribute(name);
            return attribute == null ? null : attribute.Value;
        }

        /// <summary>读取指定名字的第一个直接子元素。Reads the first direct child element with the given name.</summary>
        private static XElement FindChild(XElement element, XName name)
        {
            if (element == null) return null;
            foreach (XElement child in element.Elements())
            {
                if (child.Name == name) return child;
            }
            return null;
        }

        /// <summary>读取长度属性并换算为磅。Reads a length attribute and converts it to points.</summary>
        private static double LengthAttr(XElement element, XName name, double fallback)
        {
            return WordFormatUtil.ParseLengthPoints(Attr(element, name), fallback);
        }

        /// <summary>读取属性集合中的长度值并换算为磅。Reads a length value from a property set and converts it to points.</summary>
        private static double Length(Dictionary<XName, string> properties, XName name, double fallback)
        {
            string value = Get(properties, name);
            return WordFormatUtil.ParseLengthPoints(value, fallback);
        }

        /// <summary>读取属性集合中的字符串值。Reads a string value from a property set.</summary>
        private static string Get(Dictionary<XName, string> properties, XName name)
        {
            if (properties == null) return null;
            string value;
            return properties.TryGetValue(name, out value) ? value : null;
        }

        /// <summary>解析整数并回退。Parses an integer with a fallback.</summary>
        private static int ParseInt(string text, int fallback)
        {
            int value;
            if (!string.IsNullOrEmpty(text) && int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
            {
                return value;
            }
            return fallback;
        }

        /// <summary>把整数限制在给定区间。Clamps an integer to a range.</summary>
        private static int Clamp(int value, int minimum, int maximum)
        {
            if (maximum < minimum) maximum = minimum;
            if (value < minimum) return minimum;
            if (value > maximum) return maximum;
            return value;
        }

        // ---------------------------------------------------------------- parts

        /// <summary>解析上下文：样式索引、包读取器与节点计数。Parsing context: style index, package reader and node counter.</summary>
        private sealed class ParseContext
        {
            /// <summary>样式索引。The style index.</summary>
            public StyleIndex Styles;
            /// <summary>包内部件读取器。The package part reader.</summary>
            public PartReader Parts;
            private int _nodes;

            /// <summary>累计节点数并检查安全上限。Counts one node and enforces the safety cap.</summary>
            public void CountNode()
            {
                _nodes++;
                if (_nodes > MaxNodes) throw new InvalidDataException("ODT 文档结构超过安全上限。");
            }
        }

        /// <summary>包内部件读取器：按名字查找条目、读取 XML、限制图片大小与数量。Package part reader: looks entries up by name, reads XML and limits image size and count.</summary>
        private sealed class PartReader
        {
            private readonly ZipArchive _archive;
            private long _totalBytes;
            private int _imageCount;

            /// <summary>以压缩包创建读取器。Creates a reader over a ZIP archive.</summary>
            public PartReader(ZipArchive archive)
            {
                _archive = archive;
            }

            /// <summary>读取一个部件；缺失时返回 null。Reads a part, returning null when it is missing.</summary>
            public byte[] Read(string name)
            {
                if (string.IsNullOrEmpty(name)) return null;
                ZipArchiveEntry entry = FindEntry(_archive, name);
                return entry == null ? null : OfficePackageLimits.ReadEntryBytes(entry, ref _totalBytes);
            }

            /// <summary>读取一个 base64 内嵌图片，校验大小与数量上限。Reads one base64-embedded image, enforcing the size and count caps.</summary>
            public byte[] ReadBinaryImage(string base64)
            {
                if (string.IsNullOrEmpty(base64) || _imageCount >= MaxImages) return null;
                if (base64.Length / 4 * 3 > MaxImageBytes) throw new InvalidDataException("ODT 内嵌图片超过安全上限。");
                byte[] data;
                try { data = Convert.FromBase64String(base64); }
                catch (FormatException) { return null; }
                if (data == null || data.Length == 0) return null;
                if (data.Length > MaxImageBytes) throw new InvalidDataException("ODT 内嵌图片超过安全上限。");
                _imageCount++;
                return data;
            }

            /// <summary>读取包内图片部件；拒绝绝对路径、上级路径与超限大小。Reads an image part from the package, rejecting absolute paths, parent paths and oversized data.</summary>
            public byte[] ReadImagePart(string href)
            {
                if (string.IsNullOrEmpty(href) || _imageCount >= MaxImages) return null;
                string path = href.Trim();
                int fragment = path.IndexOf('#');
                if (fragment >= 0) path = path.Substring(0, fragment);
                path = path.Replace('\\', '/');
                while (path.StartsWith("./", StringComparison.Ordinal)) path = path.Substring(2);
                if (path.Length == 0 || path[0] == '/') return null;
                if (path.IndexOf("..", StringComparison.Ordinal) >= 0) return null;
                byte[] data = Read(path);
                if (data == null) return null;
                if (data.Length > MaxImageBytes) throw new InvalidDataException("ODT 内嵌图片超过安全上限。");
                _imageCount++;
                return data;
            }

            /// <summary>查找 ZIP 条目（精确、斜杠归一化、忽略大小写依次尝试）。Finds a ZIP entry (exact, slash-normalised and case-insensitive lookups in that order).</summary>
            private static ZipArchiveEntry FindEntry(ZipArchive archive, string name)
            {
                ZipArchiveEntry entry = archive.GetEntry(name);
                if (entry != null) return entry;
                string normalized = name.Replace('\\', '/');
                entry = archive.GetEntry(normalized);
                if (entry != null) return entry;
                foreach (ZipArchiveEntry candidate in archive.Entries)
                {
                    if (string.Equals(candidate.FullName.Replace('\\', '/'), normalized, StringComparison.OrdinalIgnoreCase))
                    {
                        return candidate;
                    }
                }
                return null;
            }
        }
    }
}
