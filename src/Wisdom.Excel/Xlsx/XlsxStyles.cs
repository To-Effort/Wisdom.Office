using System;
using System.Collections.Generic;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>Excel 内置数字格式（编号 0-49）与格式代码之间的双向映射。 Two-way mapping between Excel built-in number formats (ids 0-49) and their format codes.</summary>
    internal static class XlsxNumberFormats
    {
        /// <summary>按内置格式编号返回对应的格式代码。 Returns the format code for a built-in format id.</summary>
        /// <param name="id">内置格式编号（0-49）。 Built-in format id (0-49).</param>
        /// <returns>格式代码；编号未知时返回 null。 The format code, or null when the id is unknown.</returns>
        public static string GetBuiltIn(int id)
        {
            switch (id)
            {
                case 0: return "General";
                case 1: return "0";
                case 2: return "0.00";
                case 3: return "#,##0";
                case 4: return "#,##0.00";
                case 5: return "\"$\"#,##0_);(\"$\"#,##0)";
                case 6: return "\"$\"#,##0_);[Red](\"$\"#,##0)";
                case 7: return "\"$\"#,##0.00_);(\"$\"#,##0.00)";
                case 8: return "\"$\"#,##0.00_);[Red](\"$\"#,##0.00)";
                case 9: return "0%";
                case 10: return "0.00%";
                case 11: return "0.00E+00";
                case 12: return "# ?/?";
                case 13: return "# ??/??";
                case 14: return "mm-dd-yy";
                case 15: return "d-mmm-yy";
                case 16: return "d-mmm";
                case 17: return "mmm-yy";
                case 18: return "h:mm AM/PM";
                case 19: return "h:mm:ss AM/PM";
                case 20: return "h:mm";
                case 21: return "h:mm:ss";
                case 22: return "m/d/yy h:mm";
                case 37: return "#,##0 ;(#,##0)";
                case 38: return "#,##0 ;[Red](#,##0)";
                case 39: return "#,##0.00;(#,##0.00)";
                case 40: return "#,##0.00;[Red](#,##0.00)";
                case 41: return "_(* #,##0_);_(* (#,##0);_(* \"-\"_);_(@_)";
                case 42: return "_(\"$\"* #,##0_);_(\"$\"* (#,##0);_(\"$\"* \"-\"_);_(@_)";
                case 43: return "_(* #,##0.00_);_(* (#,##0.00);_(* \"-\"??_);_(@_)";
                case 44: return "_(\"$\"* #,##0.00_);_(\"$\"* (#,##0.00);_(\"$\"* \"-\"??_);_(@_)";
                case 45: return "mm:ss";
                case 46: return "[h]:mm:ss";
                case 47: return "mmss.0";
                case 48: return "##0.0E+0";
                case 49: return "@";
                default: return null;
            }
        }

        /// <summary>在内置格式中查找与格式代码完全一致的编号。 Finds the built-in id whose format code exactly matches the given code.</summary>
        /// <param name="format">格式代码文本。 Format code text.</param>
        /// <returns>匹配的内置编号；空或 General 返回 0，未找到返回 -1。 The matching built-in id; 0 for empty or General, -1 when no built-in format matches.</returns>
        public static int FindBuiltIn(string format)
        {
            if (string.IsNullOrEmpty(format) || format == "General") return 0;
            for (int id = 1; id <= 49; id++)
            {
                string candidate = GetBuiltIn(id);
                if (candidate != null && string.Equals(candidate, format, StringComparison.Ordinal))
                {
                    return id;
                }
            }
            return -1;
        }
    }

    /// <summary>
    /// 解析 xl/styles.xml：将 cellXfs 中的每个 xf 结合 fonts/fills/borders/numFmts
    /// 展开为统一的 CellStyle 列表，供读取单元格时按样式索引取用。
    /// Parses xl/styles.xml by expanding each cellXfs xf through fonts/fills/borders/numFmts
    /// into a flat list of CellStyle objects that can be looked up by style index while reading cells.
    /// </summary>
    internal sealed class XlsxStyles
    {
        /// <summary>BIFF/旧式索引颜色表（65 项，按 indexed 属性取值）。 BIFF/legacy indexed palette (65 entries) indexed by the indexed attribute.</summary>
        private static readonly byte[][] IndexedColors =
        {
            new byte[] { 0x00, 0x00, 0x00 }, new byte[] { 0xFF, 0xFF, 0xFF },
            new byte[] { 0xFF, 0x00, 0x00 }, new byte[] { 0x00, 0xFF, 0x00 },
            new byte[] { 0x00, 0x00, 0xFF }, new byte[] { 0xFF, 0xFF, 0x00 },
            new byte[] { 0xFF, 0x00, 0xFF }, new byte[] { 0x00, 0xFF, 0xFF },
            new byte[] { 0x00, 0x00, 0x00 }, new byte[] { 0xFF, 0xFF, 0xFF },
            new byte[] { 0xFF, 0x00, 0x00 }, new byte[] { 0x00, 0xFF, 0x00 },
            new byte[] { 0x00, 0x00, 0xFF }, new byte[] { 0xFF, 0xFF, 0x00 },
            new byte[] { 0xFF, 0x00, 0xFF }, new byte[] { 0x00, 0xFF, 0xFF },
            new byte[] { 0x80, 0x00, 0x00 }, new byte[] { 0x00, 0x80, 0x00 },
            new byte[] { 0x00, 0x00, 0x80 }, new byte[] { 0x80, 0x80, 0x00 },
            new byte[] { 0x80, 0x00, 0x80 }, new byte[] { 0x00, 0x80, 0x80 },
            new byte[] { 0xC0, 0xC0, 0xC0 }, new byte[] { 0x80, 0x80, 0x80 },
            new byte[] { 0x99, 0x99, 0xFF }, new byte[] { 0x99, 0x33, 0x66 },
            new byte[] { 0xFF, 0xFF, 0xCC }, new byte[] { 0xCC, 0xFF, 0xFF },
            new byte[] { 0x66, 0x00, 0x66 }, new byte[] { 0xFF, 0x80, 0x80 },
            new byte[] { 0x00, 0x66, 0xCC }, new byte[] { 0xCC, 0xCC, 0xFF },
            new byte[] { 0x00, 0x00, 0x80 }, new byte[] { 0xFF, 0x00, 0xFF },
            new byte[] { 0xFF, 0xFF, 0x00 }, new byte[] { 0x00, 0xFF, 0xFF },
            new byte[] { 0x80, 0x00, 0x80 }, new byte[] { 0x80, 0x00, 0x00 },
            new byte[] { 0x00, 0x80, 0x80 }, new byte[] { 0x00, 0x00, 0xFF },
            new byte[] { 0x00, 0xCC, 0xFF }, new byte[] { 0xCC, 0xFF, 0xFF },
            new byte[] { 0xCC, 0xFF, 0xCC }, new byte[] { 0xFF, 0xFF, 0x99 },
            new byte[] { 0x99, 0xCC, 0xFF }, new byte[] { 0xFF, 0x99, 0xCC },
            new byte[] { 0xCC, 0x99, 0xFF }, new byte[] { 0xFF, 0xCC, 0x99 },
            new byte[] { 0x33, 0x66, 0xFF }, new byte[] { 0x33, 0xCC, 0xCC },
            new byte[] { 0x99, 0xCC, 0x00 }, new byte[] { 0xFF, 0xCC, 0x00 },
            new byte[] { 0xFF, 0x99, 0x00 }, new byte[] { 0xFF, 0x66, 0x00 },
            new byte[] { 0x66, 0x66, 0x99 }, new byte[] { 0x96, 0x96, 0x96 },
            new byte[] { 0x00, 0x33, 0x66 }, new byte[] { 0x33, 0x99, 0x66 },
            new byte[] { 0x00, 0x33, 0x00 }, new byte[] { 0x33, 0x33, 0x00 },
            new byte[] { 0x99, 0x33, 0x33 }, new byte[] { 0x33, 0x33, 0x99 },
            new byte[] { 0x99, 0x99, 0x33 }
        };

        /// <summary>展开后的样式列表，列表下标即 cellXfs 的样式索引。 Expanded style list whose position matches the cellXfs style index.</summary>
        private readonly List<CellStyle> _styles = new List<CellStyle>();
        /// <summary>自定义数字格式表（numFmtId 到 formatCode 的映射）。 Custom number-format table mapping numFmtId to formatCode.</summary>
        private readonly Dictionary<int, string> _customFormats = new Dictionary<int, string>();
        /// <summary>与 _styles 同下标的日期格式预判定结果，避免逐格克隆样式再解析格式串。 Per-index precomputed date-format flags, so reading cells never clones a style just to parse its format string.</summary>
        private bool[] _dateFlags = Array.Empty<bool>();
        /// <summary>默认样式的日期格式预判定结果。 The precomputed date-format flag of the default style.</summary>
        private bool _defaultIsDate;

        /// <summary>默认样式（cellXfs 第一项的副本；无样式表时为新实例）。 Default style (a copy of the first cellXfs entry, or a fresh instance when there is no style table).</summary>
        public CellStyle DefaultStyle { get; private set; }

        /// <summary>从 styles.xml 文档解析样式表；文档为 null 时仅保留默认样式。 Parses the style table from a styles.xml document; when the document is null only the default style is kept.</summary>
        /// <param name="document">xl/styles.xml 对应的 XML 文档。 The XML document of xl/styles.xml.</param>
        public XlsxStyles(XDocument document)
        {
            DefaultStyle = new CellStyle();
            if (document == null || document.Root == null) return;

            XElement numberFormats = Xlsx.Child(document.Root, "numFmts");
            foreach (XElement format in Xlsx.Children(numberFormats, "numFmt"))
            {
                int id = Xlsx.IntAttribute(format, "numFmtId", -1);
                string code = Xlsx.Attribute(format, "formatCode");
                if (id >= 0 && code != null) _customFormats[id] = code;
            }

            List<XElement> fonts = new List<XElement>(Xlsx.Children(Xlsx.Child(document.Root, "fonts"), "font"));
            List<XElement> fills = new List<XElement>(Xlsx.Children(Xlsx.Child(document.Root, "fills"), "fill"));
            List<XElement> borders = new List<XElement>(Xlsx.Children(Xlsx.Child(document.Root, "borders"), "border"));
            XElement cellFormats = Xlsx.Child(document.Root, "cellXfs");

            foreach (XElement xf in Xlsx.Children(cellFormats, "xf"))
            {
                CellStyle style = new CellStyle();
                int fontId = Xlsx.IntAttribute(xf, "fontId", 0);
                int fillId = Xlsx.IntAttribute(xf, "fillId", 0);
                int borderId = Xlsx.IntAttribute(xf, "borderId", 0);
                int numFmtId = Xlsx.IntAttribute(xf, "numFmtId", 0);

                if (fontId >= 0 && fontId < fonts.Count) ApplyFont(style, fonts[fontId]);
                if (fillId >= 0 && fillId < fills.Count) ApplyFill(style, fills[fillId]);
                if (borderId >= 0 && borderId < borders.Count) ApplyBorder(style, borders[borderId]);
                style.NumberFormat = GetFormat(numFmtId);

                XElement alignment = Xlsx.Child(xf, "alignment");
                if (alignment != null)
                {
                    string horizontal = Xlsx.Attribute(alignment, "horizontal");
                    if (horizontal != null)
                    {
                        switch (horizontal)
                        {
                            case "center":
                            case "centerContinuous":
                                style.HorizontalAlignment = HorizontalAlignment.Center;
                                break;
                            case "right":
                                style.HorizontalAlignment = HorizontalAlignment.Right;
                                break;
                            case "justify":
                            case "distributed":
                                style.HorizontalAlignment = HorizontalAlignment.Justify;
                                break;
                            case "left":
                                style.HorizontalAlignment = HorizontalAlignment.Left;
                                break;
                        }
                    }
                    string vertical = Xlsx.Attribute(alignment, "vertical");
                    if (vertical == "top") style.VerticalAlignment = VerticalAlignment.Top;
                    else if (vertical == "center") style.VerticalAlignment = VerticalAlignment.Middle;
                    else if (vertical == "bottom") style.VerticalAlignment = VerticalAlignment.Bottom;
                    style.WrapText = Xlsx.ParseBool(Xlsx.Attribute(alignment, "wrapText"), false);
                }

                _styles.Add(style);
            }

            if (_styles.Count > 0) DefaultStyle = _styles[0].Clone();
            _defaultIsDate = DefaultStyle.IsDateFormat();
            _dateFlags = new bool[_styles.Count];
            for (int i = 0; i < _styles.Count; i++) _dateFlags[i] = _styles[i].IsDateFormat();
        }

        /// <summary>按 cellXfs 索引直接判定该样式是否日期格式，越界时按默认样式判定。 Decides by cellXfs index whether a style is a date format, falling back to the default style when the index is out of range.</summary>
        /// <param name="index">样式索引（单元格 s 属性值，可为负）。 Style index (the cell s attribute value; may be negative).</param>
        /// <returns>是日期/时间格式返回 true。 True when the format is a date/time format.</returns>
        public bool IsDateFormat(int index)
        {
            if (index >= 0 && index < _dateFlags.Length) return _dateFlags[index];
            return _defaultIsDate;
        }

        /// <summary>按 cellXfs 索引返回样式副本；索引越界时返回默认样式副本。 Returns a copy of the style at a cellXfs index, falling back to a copy of the default style when the index is out of range.</summary>
        /// <param name="index">样式索引（单元格 s 属性值）。 Style index (the cell s attribute value).</param>
        /// <returns>样式对象的独立副本。 An independent copy of the style object.</returns>
        public CellStyle GetStyle(int index)
        {
            if (index >= 0 && index < _styles.Count) return _styles[index].Clone();
            return DefaultStyle.Clone();
        }

        /// <summary>解析格式编号：优先自定义格式，其次内置格式，最终回退 General。 Resolves a format id: custom formats first, then built-in formats, finally falling back to General.</summary>
        /// <param name="numFmtId">数字格式编号。 Number-format id.</param>
        /// <returns>格式代码文本。 The format code text.</returns>
        public string GetFormat(int numFmtId)
        {
            string format;
            if (_customFormats.TryGetValue(numFmtId, out format)) return format;
            string builtIn = XlsxNumberFormats.GetBuiltIn(numFmtId);
            return builtIn ?? "General";
        }

        /// <summary>将 font 元素（名称、字号、加粗、斜体、下划线、颜色）应用到目标样式。 Applies a font element (name, size, bold, italic, underline, color) to the target style.</summary>
        /// <param name="style">目标样式。 Target style to write into.</param>
        /// <param name="font">fonts 集合中的单个 font 元素。 A single font element from the fonts collection.</param>
        private static void ApplyFont(CellStyle style, XElement font)
        {
            XElement name = Xlsx.Child(font, "name");
            if (name != null) style.FontFamily = Xlsx.Attribute(name, "val") ?? style.FontFamily;
            XElement size = Xlsx.Child(font, "sz");
            if (size != null) style.FontSize = Xlsx.DoubleAttribute(size, "val", style.FontSize);
            XElement bold = Xlsx.Child(font, "b");
            if (bold != null) style.Bold = Xlsx.ParseBool(Xlsx.Attribute(bold, "val"), true);
            XElement italic = Xlsx.Child(font, "i");
            if (italic != null) style.Italic = Xlsx.ParseBool(Xlsx.Attribute(italic, "val"), true);
            XElement underline = Xlsx.Child(font, "u");
            if (underline != null)
            {
                string value = Xlsx.Attribute(underline, "val");
                style.Underline = value == null || value != "none";
            }
            XElement color = Xlsx.Child(font, "color");
            if (color != null) style.FontColor = ReadColor(color, style.FontColor);
        }

        /// <summary>从 patternFill 的 fgColor 提取填充色；无填充、none 图案或索引 64/auto 时跳过。 Extracts the fill color from the patternFill fgColor, skipping absent fills, none patterns and indexed 64/auto colors.</summary>
        /// <param name="style">目标样式。 Target style to write into.</param>
        /// <param name="fill">fills 集合中的单个 fill 元素。 A single fill element from the fills collection.</param>
        private static void ApplyFill(CellStyle style, XElement fill)
        {
            XElement pattern = Xlsx.Child(fill, "patternFill");
            if (pattern == null) return;
            string patternType = Xlsx.Attribute(pattern, "patternType");
            if (string.IsNullOrEmpty(patternType) || patternType == "none") return;
            XElement foreground = Xlsx.Child(pattern, "fgColor");
            if (foreground == null) return;
            OfficeColor color = ReadColor(foreground, OfficeColor.White);
            if (Xlsx.Attribute(foreground, "indexed") == "64" || Xlsx.Attribute(foreground, "auto") == "1") return;
            style.FillColor = color;
        }

        /// <summary>读取 border 元素的左、右、上、下四条边线并写入样式。 Reads the left, right, top and bottom edges of a border element into the style.</summary>
        /// <param name="style">目标样式。 Target style to write into.</param>
        /// <param name="border">borders 集合中的单个 border 元素。 A single border element from the borders collection.</param>
        private static void ApplyBorder(CellStyle style, XElement border)
        {
            style.LeftBorder = ReadBorderLine(Xlsx.Child(border, "left"));
            style.RightBorder = ReadBorderLine(Xlsx.Child(border, "right"));
            style.TopBorder = ReadBorderLine(Xlsx.Child(border, "top"));
            style.BottomBorder = ReadBorderLine(Xlsx.Child(border, "bottom"));
        }

        /// <summary>将单边线元素映射为 BorderLine；省略线型统一近似为单线，粗细按样式名取近似宽度。 Maps one border-edge element to a BorderLine; omitted line styles are approximated as a single line and the weight is taken as an approximate width from the style name.</summary>
        /// <param name="element">left/right/top/bottom 边元素。 A left/right/top/bottom edge element.</param>
        /// <returns>边线描述；无有效线型时返回 null。 The border line description, or null when there is no valid line style.</returns>
        private static BorderLine ReadBorderLine(XElement element)
        {
            if (element == null) return null;
            string styleName = Xlsx.Attribute(element, "style");
            if (string.IsNullOrEmpty(styleName)) return null;
            BorderLine line = new BorderLine();
            line.Style = BorderStyle.Single;
            line.Width = 0.5;
            if (styleName == "dashed") line.Style = BorderStyle.Dashed;
            else if (styleName == "dotted" || styleName == "hair") line.Style = BorderStyle.Dotted;
            else if (styleName == "double") line.Style = BorderStyle.Double;
            else if (styleName == "medium") line.Width = 1.5;
            else if (styleName == "thick") line.Width = 2.5;
            else if (styleName == "thin") line.Width = 0.5;
            XElement color = Xlsx.Child(element, "color");
            if (color != null) line.Color = ReadColor(color, OfficeColor.Black);
            return line;
        }

        /// <summary>解析 color 元素：优先 rgb 十六进制，其次 indexed 索引色（64 黑、65 白），否则回退。 Parses a color element: rgb hexadecimal first, then indexed palette color (64 black, 65 white), otherwise the fallback.</summary>
        /// <param name="element">color 元素。 The color element.</param>
        /// <param name="fallback">无法解析时的回退颜色。 Fallback color used when parsing fails.</param>
        /// <returns>解析出的颜色。 The resolved color.</returns>
        private static OfficeColor ReadColor(XElement element, OfficeColor fallback)
        {
            string rgb = Xlsx.Attribute(element, "rgb");
            if (!string.IsNullOrEmpty(rgb))
            {
                OfficeColor parsed;
                if (Xlsx.TryParseColor(rgb, out parsed)) return parsed;
            }
            string indexed = Xlsx.Attribute(element, "indexed");
            int index;
            if (!string.IsNullOrEmpty(indexed) && int.TryParse(indexed, out index))
            {
                if (index >= 0 && index < IndexedColors.Length)
                {
                    byte[] value = IndexedColors[index];
                    return new OfficeColor(value[0], value[1], value[2]);
                }
                if (index == 64) return OfficeColor.Black;
                if (index == 65) return OfficeColor.White;
            }
            return fallback;
        }
    }
}
