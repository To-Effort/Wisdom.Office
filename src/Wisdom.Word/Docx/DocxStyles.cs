using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// styles.xml 中单个样式定义（w:style）的内存表示，保存样式标识、继承关系与各类属性元素。
    /// In-memory representation of a single style definition (w:style) from styles.xml, holding the style
    /// identity, its inheritance link and the various property elements.
    /// </summary>
    internal sealed class DocxStyleInfo
    {
        /// <summary>样式 ID（w:style/@w:styleId）。The style id (w:style/@w:styleId).</summary>
        public string Id;
        /// <summary>样式显示名称（w:name/@w:val），缺失时回退为样式 ID。The display name (w:name/@w:val); falls back to the style id when missing.</summary>
        public string Name;
        /// <summary>样式类型：paragraph、character、table 或 numbering。The style type: paragraph, character, table or numbering.</summary>
        public string Type;
        /// <summary>基于（继承）的父样式 ID，可为 null。The id of the base (inherited) style, may be null.</summary>
        public string BasedOn;
        /// <summary>是否为默认样式（w:style/@w:default="1"）。Whether the style is marked as default (w:style/@w:default="1").</summary>
        public bool IsDefault;
        /// <summary>字符属性元素 w:rPr，可为 null。The character-properties element w:rPr, may be null.</summary>
        public XElement RunProperties;
        /// <summary>段落属性元素 w:pPr，可为 null。The paragraph-properties element w:pPr, may be null.</summary>
        public XElement ParagraphProperties;
        /// <summary>表格属性元素 w:tblPr，可为 null。The table-properties element w:tblPr, may be null.</summary>
        public XElement TableProperties;
    }

    /// <summary>
    /// DOCX 样式表：解析 styles.xml 与 numbering.xml，提供样式链查找、
    /// 段落/字符格式还原以及编号（项目符号/数字）标记生成能力。
    /// DOCX style table: parses styles.xml and numbering.xml and provides style-chain lookup, paragraph and
    /// character format restoration, and generation of numbering markers (bullets and numbers).
    /// </summary>
    internal sealed class DocxStyles
    {
        /// <summary>编号定义的单个级别（abstractNum 下的一级 lvl）。A single level of a numbering definition (one lvl under an abstractNum).</summary>
        private sealed class NumberingLevel
        {
            /// <summary>起始序号（w:start），默认为 1。The starting number (w:start), 1 by default.</summary>
            public int Start = 1;
            /// <summary>编号格式（w:numFmt/@w:val）：decimal、bullet、lowerLetter、upperRoman 等。The numbering format (w:numFmt/@w:val): decimal, bullet, lowerLetter, upperRoman, ...</summary>
            public string Format;
            /// <summary>编号文本模板（w:lvlText/@w:val），如 "%1."。The numbering text template (w:lvlText/@w:val), e.g. "%1.".</summary>
            public string Text;
        }

        /// <summary>样式 ID 到样式定义的映射（区分大小写）。Case-sensitive map from style id to style definition.</summary>
        private readonly Dictionary<string, DocxStyleInfo> _styles =
            new Dictionary<string, DocxStyleInfo>(StringComparer.Ordinal);
        /// <summary>numId 到「级别 -&gt; 编号定义」的映射，来自 numbering.xml。Map from numId to "level -&gt; numbering definition" built from numbering.xml.</summary>
        private readonly Dictionary<int, Dictionary<int, NumberingLevel>> _numberedLists =
            new Dictionary<int, Dictionary<int, NumberingLevel>>();
        /// <summary>默认段落样式 ID（无样式段落回退使用），通常为 Normal 或显式标记 default 的段落样式。The default paragraph style id used as fallback, usually Normal or the style explicitly marked default.</summary>
        private string _defaultParagraphStyleId;

        /// <summary>
        /// 解析样式表与编号定义文档并构建查找字典；两者均可为 null（此时仅保留内置默认值）。
        /// Parses the style and numbering documents and builds the lookup dictionaries; both may be null, in
        /// which case only the built-in defaults remain.
        /// </summary>
        /// <param name="stylesXml">styles.xml 文档，可为 null。The styles.xml document, may be null.</param>
        /// <param name="numberingXml">numbering.xml 文档，可为 null。The numbering.xml document, may be null.</param>
        public DocxStyles(XDocument stylesXml, XDocument numberingXml)
        {
            DefaultTextFormat = new TextFormat("Calibri", 11);
            if (stylesXml != null && stylesXml.Root != null)
            {
                ReadDefaults(stylesXml.Root);
                ReadStyles(stylesXml.Root);
            }
            if (!_styles.ContainsKey("Normal"))
            {
                DocxStyleInfo normal = new DocxStyleInfo();
                normal.Id = "Normal";
                normal.Name = "Normal";
                normal.IsDefault = true;
                _styles["Normal"] = normal;
            }
            _defaultParagraphStyleId = FindDefaultParagraphStyle();
            ReadNumbering(numberingXml);
        }

        /// <summary>docDefaults 解析出的全局默认文字格式（未加载样式时为 Calibri 11 磅）。The global default text format parsed from docDefaults (Calibri 11pt when no styles are loaded).</summary>
        public TextFormat DefaultTextFormat { get; private set; }
        /// <summary>未显式指定样式段落所回退的默认段落样式 ID。The default paragraph style id used by paragraphs without an explicit style.</summary>
        public string DefaultParagraphStyleId { get { return _defaultParagraphStyleId; } }

        /// <summary>
        /// 读取 w:docDefaults/w:rPrDefault/w:rPr 全局默认字符属性并合入 <see cref="DefaultTextFormat"/>。
        /// Reads the global default character properties from w:docDefaults/w:rPrDefault/w:rPr and merges them
        /// into <see cref="DefaultTextFormat"/>.
        /// </summary>
        /// <param name="root">styles.xml 根元素。The root element of styles.xml.</param>
        private void ReadDefaults(XElement root)
        {
            XElement docDefaults = Wml.Child(root, "docDefaults");
            XElement runDefaults = Wml.Child(docDefaults, "rPrDefault");
            XElement runProperties = Wml.Child(runDefaults, "rPr");
            if (runProperties != null) ApplyRunProperties(DefaultTextFormat, runProperties);
        }

        /// <summary>
        /// 遍历 styles.xml 中所有 w:style 元素，按类型（段落/字符/表格/编号）建立样式字典。
        /// Walks every w:style element of styles.xml and fills the style dictionary, keyed by type
        /// (paragraph, character, table, numbering).
        /// </summary>
        /// <param name="root">styles.xml 根元素。The root element of styles.xml.</param>
        private void ReadStyles(XElement root)
        {
            foreach (XElement style in Wml.Children(root, "style"))
            {
                string type = Wml.AttributeValue(style, "type");
                if (type != "paragraph" && type != "character" && type != "table" && type != "numbering") continue;
                DocxStyleInfo info = new DocxStyleInfo();
                info.Id = Wml.AttributeValue(style, "styleId");
                if (string.IsNullOrEmpty(info.Id)) continue;
                info.Type = type;
                info.IsDefault = Wml.AttributeValue(style, "default") == "1"
                    || string.Equals(Wml.AttributeValue(style, "default"), "true", StringComparison.OrdinalIgnoreCase);
                XElement name = Wml.Child(style, "name");
                info.Name = name == null ? info.Id : Wml.AttributeValue(name, "val");
                XElement basedOn = Wml.Child(style, "basedOn");
                info.BasedOn = basedOn == null ? null : Wml.AttributeValue(basedOn, "val");
                info.RunProperties = Wml.Child(style, "rPr");
                info.ParagraphProperties = Wml.Child(style, "pPr");
                info.TableProperties = Wml.Child(style, "tblPr");
                _styles[info.Id] = info;
            }
        }

        /// <summary>
        /// 查找标记为 default 的段落样式 ID；不存在时回退为 "Normal"。
        /// Finds the id of the paragraph style marked as default; falls back to "Normal" when none exists.
        /// </summary>
        /// <returns>默认段落样式 ID。The default paragraph style id.</returns>
        private string FindDefaultParagraphStyle()
        {
            string result = "Normal";
            foreach (KeyValuePair<string, DocxStyleInfo> pair in _styles)
            {
                if (pair.Value.IsDefault && pair.Value.Type == "paragraph") result = pair.Key;
            }
            return result;
        }

        /// <summary>
        /// 按样式 ID 查找样式定义。
        /// Looks up a style definition by style id.
        /// </summary>
        /// <param name="styleId">样式 ID，可为空。The style id, may be null or empty.</param>
        /// <returns>找到的样式信息，未找到时返回 null。The style found, or null when there is no match.</returns>
        public DocxStyleInfo GetStyle(string styleId)
        {
            DocxStyleInfo style;
            if (!string.IsNullOrEmpty(styleId) && _styles.TryGetValue(styleId, out style)) return style;
            return null;
        }

        /// <summary>
        /// 沿 basedOn 继承链自根到叶收集样式定义（带环路保护，已反转为「父样式在前」顺序）。
        /// Collects the style definitions along the basedOn inheritance chain from root to leaf with cycle
        /// protection, reversed so that ancestors come first.
        /// </summary>
        /// <param name="styleId">起始样式 ID。The starting style id.</param>
        /// <returns>从最祖先样式到自身的样式链。The chain from the most ancestral style down to the style itself.</returns>
        public List<DocxStyleInfo> GetChain(string styleId)
        {
            List<DocxStyleInfo> chain = new List<DocxStyleInfo>();
            HashSet<string> visited = new HashSet<string>(StringComparer.Ordinal);
            string currentId = styleId;
            while (!string.IsNullOrEmpty(currentId) && visited.Add(currentId))
            {
                DocxStyleInfo style;
                if (!_styles.TryGetValue(currentId, out style)) break;
                chain.Add(style);
                currentId = style.BasedOn;
            }
            chain.Reverse();
            return chain;
        }

        /// <summary>
        /// 获取样式继承链上全部表格属性（w:tblPr），按「父样式在前」顺序返回。
        /// Returns every table property (w:tblPr) along the style inheritance chain, ancestors first.
        /// </summary>
        /// <param name="styleId">表格样式 ID。The table style id.</param>
        /// <returns>表格属性元素列表。The list of table property elements.</returns>
        public List<XElement> GetTablePropertyChain(string styleId)
        {
            List<XElement> result = new List<XElement>();
            List<DocxStyleInfo> chain = GetChain(styleId);
            for (int i = 0; i < chain.Count; i++)
            {
                if (chain[i].TableProperties != null) result.Add(chain[i].TableProperties);
            }
            return result;
        }

        /// <summary>
        /// 从默认格式出发，按样式链逐级合并段落样式的字符属性，得到段落基准文字格式。
        /// Starts from the default format and merges the character properties of each style in the chain to
        /// obtain the base run format of a paragraph style.
        /// </summary>
        /// <param name="styleId">段落样式 ID。The paragraph style id.</param>
        /// <returns>合并样式后的文字格式。The text format after the styles have been merged.</returns>
        public TextFormat GetParagraphRunFormat(string styleId)
        {
            TextFormat format = DefaultTextFormat.Clone();
            List<DocxStyleInfo> chain = GetChain(styleId);
            for (int i = 0; i < chain.Count; i++)
            {
                if (chain[i].RunProperties != null) ApplyRunProperties(format, chain[i].RunProperties);
            }
            return format;
        }

        /// <summary>
        /// 在基准格式上按字符样式链合并字符属性，得到应用字符样式后的文字格式。
        /// Merges the character-style chain on top of a base format and returns the resulting text format.
        /// </summary>
        /// <param name="styleId">字符样式 ID。The character style id.</param>
        /// <param name="baseFormat">基准格式，为 null 时从空白格式开始。The base format; a null value starts from an empty format.</param>
        /// <returns>合并字符样式后的文字格式。The text format after the character style has been merged.</returns>
        public TextFormat GetCharacterRunFormat(string styleId, TextFormat baseFormat)
        {
            TextFormat format = baseFormat == null ? new TextFormat() : baseFormat.Clone();
            List<DocxStyleInfo> chain = GetChain(styleId);
            for (int i = 0; i < chain.Count; i++)
            {
                if (chain[i].RunProperties != null) ApplyRunProperties(format, chain[i].RunProperties);
            }
            return format;
        }

        /// <summary>
        /// 沿样式链依次应用段落属性与样式名，并根据样式名/大纲级别推断标题级别。
        /// Applies paragraph properties and style names along the style chain and infers the heading level
        /// from the style name or the outline level.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="styleId">段落样式 ID。The paragraph style id.</param>
        public void ApplyParagraphStyle(Paragraph paragraph, string styleId)
        {
            List<DocxStyleInfo> chain = GetChain(styleId);
            for (int i = 0; i < chain.Count; i++)
            {
                DocxStyleInfo style = chain[i];
                paragraph.StyleName = string.IsNullOrEmpty(style.Name) ? style.Id : style.Name;
                if (style.ParagraphProperties != null) ApplyParagraphProperties(paragraph, style.ParagraphProperties);
            }
            paragraph.HeadingLevel = GetHeadingLevel(chain, paragraph.HeadingLevel);
        }

        /// <summary>
        /// 将一份 w:pPr 段落属性元素应用到段落：对齐、行距段距（缇换算为磅）、
        /// 缩进（含悬挂缩进）、保持与下段同页、段前分页与大纲级别。
        /// Applies one w:pPr paragraph-properties element to a paragraph: alignment, line and paragraph spacing
        /// (twips converted to points), indents (including hanging indent), keep-with-next, page-break-before
        /// and outline level.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="properties">w:pPr 元素，可为 null。The w:pPr element, may be null.</param>
        public void ApplyParagraphProperties(Paragraph paragraph, XElement properties)
        {
            if (properties == null) return;
            XElement alignment = Wml.Child(properties, "jc");
            string alignmentValue = Wml.AttributeValue(alignment, "val");
            if (!string.IsNullOrEmpty(alignmentValue))
            {
                switch (alignmentValue)
                {
                    case "center": paragraph.Alignment = ParagraphAlignment.Center; break;
                    case "right":
                    case "end": paragraph.Alignment = ParagraphAlignment.Right; break;
                    case "both":
                    case "distribute": paragraph.Alignment = ParagraphAlignment.Justify; break;
                    case "left":
                    case "start": paragraph.Alignment = ParagraphAlignment.Left; break;
                }
            }

            XElement spacing = Wml.Child(properties, "spacing");
            if (spacing != null)
            {
                string before = Wml.AttributeValue(spacing, "before");
                if (before != null) paragraph.SpaceBefore = Wml.TwipsToPoints(before, paragraph.SpaceBefore);
                string after = Wml.AttributeValue(spacing, "after");
                if (after != null) paragraph.SpaceAfter = Wml.TwipsToPoints(after, paragraph.SpaceAfter);
                string line = Wml.AttributeValue(spacing, "line");
                string rule = Wml.AttributeValue(spacing, "lineRule");
                if (line != null && (rule == null || rule == "auto"))
                {
                    double twips = Wml.ParseDouble(line, 0);
                    if (twips > 0) paragraph.LineSpacing = twips / 240.0;
                }
            }

            XElement indent = Wml.Child(properties, "ind");
            if (indent != null)
            {
                string left = Wml.AttributeValue(indent, "left") ?? Wml.AttributeValue(indent, "start");
                if (left != null) paragraph.LeftIndent = Wml.TwipsToPoints(left, paragraph.LeftIndent);
                string right = Wml.AttributeValue(indent, "right") ?? Wml.AttributeValue(indent, "end");
                if (right != null) paragraph.RightIndent = Wml.TwipsToPoints(right, paragraph.RightIndent);
                string firstLine = Wml.AttributeValue(indent, "firstLine");
                string hanging = Wml.AttributeValue(indent, "hanging");
                if (firstLine != null) paragraph.FirstLineIndent = Wml.TwipsToPoints(firstLine, paragraph.FirstLineIndent);
                else if (hanging != null) paragraph.FirstLineIndent = -Wml.TwipsToPoints(hanging, -paragraph.FirstLineIndent);
            }

            if (Wml.Child(properties, "keepNext") != null) paragraph.KeepWithNext = Wml.GetOnOff(properties, "keepNext", true);
            if (Wml.Child(properties, "pageBreakBefore") != null) paragraph.PageBreakBefore = Wml.GetOnOff(properties, "pageBreakBefore", true);

            XElement outline = Wml.Child(properties, "outlineLvl");
            if (outline != null)
            {
                int level = Wml.ParseInt(Wml.AttributeValue(outline, "val"), -1);
                if (level >= 0 && level < 9) paragraph.HeadingLevel = level + 1;
            }
        }

        /// <summary>
        /// 将一份 w:rPr 字符属性元素合入文字格式：中西文字体、字号（半磅换算）、
        /// 加粗/倾斜/删除线/下划线、颜色与上下标。
        /// Merges one w:rPr character-properties element into a text format: East-Asian and Latin fonts, size
        /// (half-point conversion), bold/italic/strike-through/underline, colour and super/subscript.
        /// </summary>
        /// <param name="format">接收属性的目标文字格式。The target text format that receives the properties.</param>
        /// <param name="properties">w:rPr 元素，可为 null。The w:rPr element, may be null.</param>
        public static void ApplyRunProperties(TextFormat format, XElement properties)
        {
            if (properties == null) return;
            XElement fonts = Wml.Child(properties, "rFonts");
            if (fonts != null)
            {
                string ascii = Wml.AttributeValue(fonts, "ascii");
                if (string.IsNullOrEmpty(ascii)) ascii = Wml.AttributeValue(fonts, "hAnsi");
                string eastAsia = Wml.AttributeValue(fonts, "eastAsia");
                if (!string.IsNullOrEmpty(ascii))
                {
                    format.AsciiFontFamily = ascii;
                    format.FontFamily = ascii;
                }
                if (!string.IsNullOrEmpty(eastAsia))
                {
                    format.EastAsiaFontFamily = eastAsia;
                    if (string.IsNullOrEmpty(ascii)) format.FontFamily = eastAsia;
                }
            }

            XElement size = Wml.Child(properties, "sz");
            if (size != null)
            {
                int halfPoints = Wml.ParseInt(Wml.AttributeValue(size, "val"), 0);
                if (halfPoints > 0) format.FontSize = halfPoints / 2.0;
            }

            XElement bold = Wml.Child(properties, "b");
            if (bold != null) format.Bold = Wml.AttributeValue(bold, "val") != "0" && Wml.AttributeValue(bold, "val") != "false";
            XElement italic = Wml.Child(properties, "i");
            if (italic != null) format.Italic = Wml.AttributeValue(italic, "val") != "0" && Wml.AttributeValue(italic, "val") != "false";
            XElement strike = Wml.Child(properties, "strike");
            if (strike != null) format.Strikeout = Wml.AttributeValue(strike, "val") != "0" && Wml.AttributeValue(strike, "val") != "false";

            XElement underline = Wml.Child(properties, "u");
            if (underline != null)
            {
                string value = Wml.AttributeValue(underline, "val");
                format.Underline = value != null && value != "none" && value != "0";
            }

            XElement color = Wml.Child(properties, "color");
            if (color != null)
            {
                string value = Wml.AttributeValue(color, "val");
                if (!string.IsNullOrEmpty(value) && value != "auto")
                {
                    OfficeColor parsed;
                    if (OfficeColor.TryParse(value, out parsed)) format.Color = parsed;
                }
            }

            XElement vertical = Wml.Child(properties, "vertAlign");
            if (vertical != null)
            {
                string value = Wml.AttributeValue(vertical, "val");
                format.Superscript = value == "superscript";
                format.Subscript = value == "subscript";
            }
        }

        /// <summary>
        /// 从段落属性 w:numPr 中提取编号引用（numId + ilvl）并生成对应的列表标记文本。
        /// Extracts the numbering reference (numId plus ilvl) from the paragraph properties w:numPr and builds
        /// the matching list marker text.
        /// </summary>
        /// <param name="paragraphProperties">段落属性元素 w:pPr，可为 null。The w:pPr element, may be null.</param>
        /// <returns>标记文本（如 "• " 或 "1. "），非编号段落返回 null。The marker text (e.g. "• " or "1. "); null for paragraphs without numbering.</returns>
        public string GetListMarker(XElement paragraphProperties)
        {
            XElement numPr = Wml.Child(paragraphProperties, "numPr");
            if (numPr == null) return null;
            int numId = Wml.GetInt(numPr, "numId", 0);
            int level = Wml.GetInt(numPr, "ilvl", 0);
            return GetMarker(numId, level);
        }

        /// <summary>
        /// 按编号定义生成指定级别的标记文本：项目符号、字母/罗马数字或十进制数字，
        /// 并替换 lvlText 模板中的 %n 占位符。
        /// Builds the marker text of a numbering level: bullet, letter, Roman or decimal number, replacing the
        /// %n placeholders of the lvlText template.
        /// </summary>
        /// <param name="numId">编号序列 ID（w:num/@w:numId）。The numbering id (w:num/@w:numId).</param>
        /// <param name="level">缩进级别（ilvl，从 0 开始）。The indentation level (ilvl, zero-based).</param>
        /// <returns>标记文本；无对应编号定义或格式为 none 时返回 null。The marker text; null when the numbering is unknown or its format is none.</returns>
        public string GetMarker(int numId, int level)
        {
            Dictionary<int, NumberingLevel> levels;
            if (!_numberedLists.TryGetValue(numId, out levels)) return null;
            NumberingLevel numbering;
            if (!levels.TryGetValue(level, out numbering))
            {
                if (!levels.TryGetValue(0, out numbering)) return null;
            }
            string marker = numbering.Text ?? string.Empty;
            string format = numbering.Format ?? "decimal";
            if (format == "bullet") return marker.Length == 0 ? "\u2022" : marker;
            if (format == "none") return null;

            string number;
            if (format == "lowerLetter") number = ToLetter(numbering.Start, false);
            else if (format == "upperLetter") number = ToLetter(numbering.Start, true);
            else if (format == "lowerRoman") number = ToRoman(numbering.Start, false);
            else if (format == "upperRoman") number = ToRoman(numbering.Start, true);
            else number = numbering.Start.ToString(CultureInfo.InvariantCulture);

            if (marker.IndexOf('%') >= 0)
            {
                int percent = marker.IndexOf('%');
                return marker.Substring(0, percent) + number + marker.Substring(Math.Min(marker.Length, percent + 2));
            }
            return number + ". ";
        }

        /// <summary>
        /// 将数值转换为单个英文字母序号（1-&gt;a/b...），用于 lowerLetter/upperLetter 编号格式。
        /// Converts a value into a single letter index (1 -&gt; a, b, ...) for the lowerLetter/upperLetter
        /// numbering formats.
        /// </summary>
        /// <param name="value">从 1 开始的序号。The one-based index.</param>
        /// <param name="upper">true 输出大写字母，false 输出小写。True for upper-case letters, false for lower-case.</param>
        /// <returns>字母文本。The letter text.</returns>
        private static string ToLetter(int value, bool upper)
        {
            char c = (char)('a' + Math.Max(0, Math.Min(25, value - 1)));
            return (upper ? char.ToUpperInvariant(c) : c).ToString();
        }

        /// <summary>
        /// 将数值转换为罗马数字文本（限 1-3999），用于 lowerRoman/upperRoman 编号格式。
        /// Converts a value into Roman numeral text (clamped to 1-3999) for the lowerRoman/upperRoman
        /// numbering formats.
        /// </summary>
        /// <param name="value">从 1 开始的序号，超出范围时截断。The one-based index, clamped when out of range.</param>
        /// <param name="upper">true 输出大写，false 输出小写。True for upper-case, false for lower-case.</param>
        /// <returns>罗马数字文本。The Roman numeral text.</returns>
        private static string ToRoman(int value, bool upper)
        {
            int number = Math.Max(1, Math.Min(3999, value));
            int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
            string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
            System.Text.StringBuilder builder = new System.Text.StringBuilder();
            for (int i = 0; i < values.Length; i++)
            {
                while (number >= values[i])
                {
                    builder.Append(symbols[i]);
                    number -= values[i];
                }
            }
            string result = builder.ToString();
            return upper ? result : result.ToLowerInvariant();
        }

        /// <summary>
        /// 从样式链推断标题级别：识别 "heading N" 样式名或 w:outlineLvl 大纲级别属性。
        /// Infers the heading level from the style chain by recognising "heading N" style names or the
        /// w:outlineLvl outline-level property.
        /// </summary>
        /// <param name="chain">样式继承链（父样式在前）。The style inheritance chain, ancestors first.</param>
        /// <param name="fallback">未识别出标题级别时的回退值。The value used when no heading level is recognised.</param>
        /// <returns>标题级别（1-9），未识别时返回 <paramref name="fallback"/>。The heading level (1-9), or <paramref name="fallback"/> when unknown.</returns>
        private int GetHeadingLevel(List<DocxStyleInfo> chain, int fallback)
        {
            int result = fallback;
            for (int i = 0; i < chain.Count; i++)
            {
                DocxStyleInfo style = chain[i];
                if (style.Name != null && style.Name.StartsWith("heading", StringComparison.OrdinalIgnoreCase))
                {
                    int level;
                    if (int.TryParse(style.Name.Substring(7).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out level))
                    {
                        result = level;
                    }
                }
                else if (style.Id != null && style.Id.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
                {
                    int level;
                    if (int.TryParse(style.Id.Substring(7), NumberStyles.Integer, CultureInfo.InvariantCulture, out level))
                    {
                        result = level;
                    }
                }
                if (style.ParagraphProperties != null)
                {
                    XElement outline = Wml.Child(style.ParagraphProperties, "outlineLvl");
                    if (outline != null)
                    {
                        int value = Wml.ParseInt(Wml.AttributeValue(outline, "val"), -1);
                        if (value >= 0 && value < 9) result = value + 1;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 解析 numbering.xml：先读取 abstractNum 各级编号定义，再按 num 建立 numId 到级别定义的映射。
        /// Parses numbering.xml: reads the level definitions of each abstractNum first, then maps numId to the
        /// level definitions through the num elements.
        /// </summary>
        /// <param name="numberingXml">numbering.xml 文档，可为 null。The numbering.xml document, may be null.</param>
        private void ReadNumbering(XDocument numberingXml)
        {
            if (numberingXml == null || numberingXml.Root == null) return;
            Dictionary<int, Dictionary<int, NumberingLevel>> abstracts =
                new Dictionary<int, Dictionary<int, NumberingLevel>>();
            foreach (XElement abstractNum in Wml.Children(numberingXml.Root, "abstractNum"))
            {
                int id = Wml.ParseInt(Wml.AttributeValue(abstractNum, "abstractNumId"), -1);
                if (id < 0) continue;
                Dictionary<int, NumberingLevel> levels = new Dictionary<int, NumberingLevel>();
                foreach (XElement lvl in Wml.Children(abstractNum, "lvl"))
                {
                    int level = Wml.ParseInt(Wml.AttributeValue(lvl, "ilvl"), 0);
                    NumberingLevel numbering = new NumberingLevel();
                    numbering.Start = Wml.GetInt(lvl, "start", 1);
                    XElement format = Wml.Child(lvl, "numFmt");
                    numbering.Format = format == null ? "decimal" : Wml.AttributeValue(format, "val");
                    XElement text = Wml.Child(lvl, "lvlText");
                    numbering.Text = text == null ? "%1." : Wml.AttributeValue(text, "val");
                    levels[level] = numbering;
                }
                abstracts[id] = levels;
            }

            foreach (XElement num in Wml.Children(numberingXml.Root, "num"))
            {
                int numId = Wml.ParseInt(Wml.AttributeValue(num, "numId"), -1);
                if (numId < 0) continue;
                int abstractId = Wml.GetInt(num, "abstractNumId", -1);
                Dictionary<int, NumberingLevel> levels;
                if (abstracts.TryGetValue(abstractId, out levels))
                {
                    _numberedLists[numId] = levels;
                }
            }
        }
    }
}
