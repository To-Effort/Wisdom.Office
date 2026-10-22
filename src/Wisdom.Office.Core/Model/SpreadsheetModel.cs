using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Wisdom.Office.Core
{
    /// <summary>单元格水平对齐方式。The horizontal alignment of a cell.</summary>
    public enum HorizontalAlignment
    {
        /// <summary>常规：按数据类型自动对齐。General: aligned automatically by data type.</summary>
        General = 0,
        /// <summary>左对齐。Left-aligned.</summary>
        Left = 1,
        /// <summary>居中对齐。Center-aligned.</summary>
        Center = 2,
        /// <summary>右对齐。Right-aligned.</summary>
        Right = 3,
        /// <summary>两端对齐。Justified.</summary>
        Justify = 4
    }

    /// <summary>单元格样式：字体、填充、对齐、边框与数字格式代码。A cell style: font, fill, alignment, borders and the number-format code.</summary>
    public sealed class CellStyle
    {
        /// <summary>字体名，默认 Calibri。The font family, Calibri by default.</summary>
        public string FontFamily { get; set; }
        /// <summary>字号，单位为磅。The font size, in points.</summary>
        public double FontSize { get; set; }
        /// <summary>是否加粗。Whether the text is bold.</summary>
        public bool Bold { get; set; }
        /// <summary>是否倾斜。Whether the text is italic.</summary>
        public bool Italic { get; set; }
        /// <summary>是否带下划线。Whether the text is underlined.</summary>
        public bool Underline { get; set; }
        /// <summary>字体颜色。The font color.</summary>
        public OfficeColor FontColor { get; set; }
        /// <summary>背景填充色，null 表示无填充。The background fill color; null means no fill.</summary>
        public OfficeColor? FillColor { get; set; }
        /// <summary>水平对齐方式。The horizontal alignment.</summary>
        public HorizontalAlignment HorizontalAlignment { get; set; }
        /// <summary>垂直对齐方式。The vertical alignment.</summary>
        public VerticalAlignment VerticalAlignment { get; set; }
        /// <summary>是否自动换行。Whether the text wraps automatically.</summary>
        public bool WrapText { get; set; }
        /// <summary>Excel 数字/日期格式代码，如 "General"、"0.00"、"yyyy-MM-dd"。The Excel number/date format code, such as "General", "0.00" or "yyyy-MM-dd".</summary>
        public string NumberFormat { get; set; }
        /// <summary>上边框线。The top border line.</summary>
        public BorderLine TopBorder { get; set; }
        /// <summary>下边框线。The bottom border line.</summary>
        public BorderLine BottomBorder { get; set; }
        /// <summary>左边框线。The left border line.</summary>
        public BorderLine LeftBorder { get; set; }
        /// <summary>右边框线。The right border line.</summary>
        public BorderLine RightBorder { get; set; }

        /// <summary>默认构造：Calibri 11 磅、黑色、底部对齐、General 格式。Default constructor: Calibri 11 points, black, bottom-aligned and the General format.</summary>
        public CellStyle()
        {
            FontFamily = "Calibri";
            FontSize = 11;
            FontColor = OfficeColor.Black;
            VerticalAlignment = VerticalAlignment.Bottom;
            NumberFormat = "General";
        }

        /// <summary>复制样式，并对四条边框线做深拷贝。Copies the style and deep-copies the four border lines.</summary>
        /// <returns>属性相同的新样式实例。A new style instance with the same properties.</returns>
        public CellStyle Clone()
        {
            CellStyle copy = (CellStyle)MemberwiseClone();
            if (TopBorder != null) copy.TopBorder = TopBorder.Clone();
            if (BottomBorder != null) copy.BottomBorder = BottomBorder.Clone();
            if (LeftBorder != null) copy.LeftBorder = LeftBorder.Clone();
            if (RightBorder != null) copy.RightBorder = RightBorder.Clone();
            return copy;
        }

        /// <summary>判断 <see cref="NumberFormat"/> 是否为日期/时间格式：跳过引号与方括号内容后扫描 yyyy/MM/dd/hh/ss/AM-PM 等令牌。Determines whether <see cref="NumberFormat"/> is a date/time format by scanning for tokens such as yyyy, MM, dd, hh, ss and AM-PM after skipping quoted and bracketed content.</summary>
        /// <returns>是日期/时间格式返回 true。True when the format is a date/time format.</returns>
        /// <remarks>用于决定 DateTime 值是套用格式代码显示，还是回退默认的 ISO 日期。"mm" 紧跟时分秒时按分钟处理，同样视为时间格式。The result decides whether a DateTime value is rendered with the format code or falls back to the default ISO date. A "mm" adjacent to hours or seconds is treated as minutes and still counts as a time format.</remarks>
        public bool IsDateFormat()
        {
            if (string.IsNullOrEmpty(NumberFormat))
            {
                return false;
            }

            string format = NumberFormat;
            if (format == "General" || format == "@")
            {
                return false;
            }

            bool inQuotes = false;
            bool inBrackets = false;
            for (int i = 0; i < format.Length; i++)
            {
                char c = format[i];
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    continue;
                }
                if (inQuotes)
                {
                    continue;
                }
                if (c == '[')
                {
                    inBrackets = true;
                    continue;
                }
                if (c == ']')
                {
                    inBrackets = false;
                    continue;
                }
                if (inBrackets)
                {
                    continue;
                }
                if (c == 'a' || c == 'A')
                {
                    int markerLength;
                    if (SpreadsheetFormat.IsAmPmMarker(format, i, out markerLength))
                    {
                        return true;
                    }
                    int run = i;
                    while (run < format.Length && (format[run] == 'a' || format[run] == 'A')) run++;
                    if (run - i >= 2) return true;
                    i = run - 1;
                    continue;
                }
                if (c == 'y' || c == 'Y' || c == 'd' || c == 'D' || c == 'm' || c == 'M'
                    || c == 'h' || c == 'H' || c == 's' || c == 'S')
                {
                    if (c == 'm' || c == 'M')
                    {
                        int next = i + 1;
                        while (next < format.Length && (format[next] == 'm' || format[next] == 'M'))
                        {
                            next++;
                        }
                        if (next < format.Length && (format[next] == 's' || format[next] == 'S'))
                        {
                            return true;
                        }
                    }
                    return true;
                }
            }
            return false;
        }
    }

    /// <summary>工作表单元格：值为任意对象，行列坐标均为 1 基。A worksheet cell: the value is any object and row/column coordinates are 1-based.</summary>
    public sealed class Cell
    {
        /// <summary>行号（从 1 开始）。The row number (starting from 1).</summary>
        public int Row { get; private set; }
        /// <summary>列号（从 1 开始）。The column number (starting from 1).</summary>
        public int Column { get; private set; }
        /// <summary>单元格值，支持数字、文本、日期、布尔等类型。The cell value; supports numbers, text, dates, booleans and other types.</summary>
        public object Value { get; set; }
        /// <summary>公式文本（不含前导 "="），null 表示无公式。The formula text (without a leading "="); null means no formula.</summary>
        public string Formula { get; set; }
        /// <summary>单元格样式，null 表示沿用默认样式。The cell style; null means the default style applies.</summary>
        public CellStyle Style { get; set; }
        /// <summary>超链接目标：外部 URL，或以 "#" 开头的表内位置（如 "#Sheet1!A1"）。The hyperlink target: an external URL, or an in-workbook location prefixed with "#" (such as "#Sheet1!A1").</summary>
        public string Hyperlink { get; set; }
        /// <summary>超链接的悬停提示文本，null 表示无提示。The hover tooltip of the hyperlink; null means no tooltip.</summary>
        public string HyperlinkTooltip { get; set; }

        /// <summary>创建位于指定行列的单元格。Creates a cell at the given row and column.</summary>
        /// <param name="row">行号，必须大于等于 1。The row number, must be at least 1.</param>
        /// <param name="column">列号，必须大于等于 1。The column number, must be at least 1.</param>
        /// <exception cref="ArgumentOutOfRangeException">行号或列号小于 1。The row or column is less than 1.</exception>
        public Cell(int row, int column)
        {
            if (row < 1) throw new ArgumentOutOfRangeException("row");
            if (column < 1) throw new ArgumentOutOfRangeException("column");
            Row = row;
            Column = column;
        }

        /// <summary>仅供模型内部在插入/删除行列时移动单元格位置，保持值与样式不变。Internal model helper that moves the cell when rows or columns are inserted or deleted, keeping its value and style.</summary>
        /// <param name="row">新行号。The new row number.</param>
        /// <param name="column">新列号。The new column number.</param>
        internal void MoveTo(int row, int column)
        {
            if (row < 1) throw new ArgumentOutOfRangeException("row");
            if (column < 1) throw new ArgumentOutOfRangeException("column");
            Row = row;
            Column = column;
        }

        /// <summary>确保单元格拥有样式实例：为空时按默认值创建。Ensures the cell has a style instance, creating a default one when null.</summary>
        /// <returns>单元格样式。The cell style.</returns>
        public CellStyle EnsureStyle()
        {
            if (Style == null)
            {
                Style = new CellStyle();
            }
            return Style;
        }

        /// <summary>按值类型与数字格式渲染展示文本：日期走 Excel 格式转换、布尔输出 TRUE/FALSE、数值按格式代码格式化，其余按不变区域转换。Renders the display text by value type and number format: dates go through Excel format conversion, booleans output TRUE/FALSE, numbers are formatted with the format code, and everything else uses the invariant culture.</summary>
        /// <returns>用于显示的字符串；值为 null 时返回空字符串。The string for display; an empty string when the value is null.</returns>
        public string DisplayText()
        {
            if (Value == null)
            {
                return string.Empty;
            }
            if (Value is DateTime)
            {
                DateTime date = (DateTime)Value;
                if (Style != null && Style.IsDateFormat())
                {
                    return SpreadsheetFormat.FormatDate(date, Style.NumberFormat);
                }
                return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            if (Value is bool)
            {
                return (bool)Value ? "TRUE" : "FALSE";
            }
            if (Value is double)
            {
                return SpreadsheetFormat.FormatNumber((double)Value, Style == null ? null : Style.NumberFormat, "G15");
            }
            if (Value is float)
            {
                return SpreadsheetFormat.FormatNumber(Convert.ToDouble(Value, CultureInfo.InvariantCulture),
                    Style == null ? null : Style.NumberFormat, "G7");
            }
            if (Value is decimal || Value is byte || Value is sbyte || Value is short || Value is ushort
                || Value is int || Value is uint || Value is long || Value is ulong)
            {
                return SpreadsheetFormat.FormatNumber((IFormattable)Value, Style == null ? null : Style.NumberFormat);
            }
            return Convert.ToString(Value, CultureInfo.InvariantCulture);
        }
    }

    /// <summary>内部格式化工具：把 Excel 数字/日期格式代码翻译为 .NET 格式并应用于显示文本。An internal formatting helper that translates Excel number/date format codes into .NET formats and applies them to display text.</summary>
    internal static class SpreadsheetFormat
    {
        /// <summary>按 Excel 数字格式代码格式化数值。Formats a number according to an Excel number-format code.</summary>
        /// <param name="value">待格式化的值。The value to format.</param>
        /// <param name="format">Excel 格式代码；空或 General/@ 视为无格式。The Excel format code; empty or General/@ is treated as no format.</param>
        /// <returns>格式化后的字符串；value 为 null 时返回空字符串。The formatted string; an empty string when value is null.</returns>
        public static string FormatNumber(IFormattable value, string format)
        {
            return FormatNumber(value, format, null);
        }

        /// <summary>按 Excel 数字格式代码格式化数值，格式缺失或非法时回退指定的一般格式。Formats a number with an Excel number-format code, falling back to the supplied general format when the code is missing or invalid.</summary>
        /// <param name="value">待格式化的值。The value to format.</param>
        /// <param name="format">Excel 数字格式代码。The Excel number-format code.</param>
        /// <param name="generalFormat">回退用的 .NET 一般格式（如 "G15"），null 表示默认 ToString。The fallback .NET general format (such as "G15"); null means the default ToString.</param>
        /// <returns>格式化后的字符串。The formatted string.</returns>
        public static string FormatNumber(IFormattable value, string format, string generalFormat)
        {
            if (value == null) return string.Empty;
            IFormattable target;
            string pattern = ToNumberPattern(SelectSection(format, value, out target));
            if (pattern.Length == 0)
            {
                return generalFormat == null
                    ? target.ToString(null, CultureInfo.InvariantCulture)
                    : target.ToString(generalFormat, CultureInfo.InvariantCulture);
            }
            try
            {
                return target.ToString(pattern, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                return generalFormat == null
                    ? target.ToString(null, CultureInfo.InvariantCulture)
                    : target.ToString(generalFormat, CultureInfo.InvariantCulture);
            }
        }

        /// <summary>按 Excel 日期格式代码格式化日期时间；格式非法时回退 "yyyy-MM-dd"。Formats a date and time with an Excel date-format code, falling back to "yyyy-MM-dd" when the code is invalid.</summary>
        /// <param name="value">日期时间值。The date and time value.</param>
        /// <param name="format">Excel 日期格式代码。The Excel date-format code.</param>
        /// <returns>格式化后的字符串。The formatted string.</returns>
        public static string FormatDate(DateTime value, string format)
        {
            string pattern = ToDatePattern(format);
            string single = FormatSingleToken(value, pattern);
            if (single != null) return single;
            try
            {
                return value.ToString(pattern, CultureInfo.InvariantCulture);
            }
            catch (FormatException)
            {
                return value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
        }

        // 单字符模式在 .NET 中是标准说明符（"d" 会展开为短日期模式），
        // 而 Excel 将其视为孤立字段（如"日"），需单独处理。
        /// <summary>处理长度为 1 的日期模式：按 Excel 语义输出单个字段值。Handles a length-1 date pattern by emitting the single field value with Excel semantics.</summary>
        /// <param name="value">日期时间值。The date and time value.</param>
        /// <param name="pattern">转换后的 .NET 日期模式。The converted .NET date pattern.</param>
        /// <returns>单字段的显示文本；非单字符模式返回 null。The single-field display text; null when the pattern is not a single character.</returns>
        private static string FormatSingleToken(DateTime value, string pattern)
        {
            if (pattern == null || pattern.Length != 1) return null;
            switch (pattern[0])
            {
                case 'd': return value.Day.ToString(CultureInfo.InvariantCulture);
                case 'M': return value.Month.ToString(CultureInfo.InvariantCulture);
                case 'y': return (value.Year % 100).ToString(CultureInfo.InvariantCulture);
                case 'H': return value.Hour.ToString(CultureInfo.InvariantCulture);
                case 'm': return value.Minute.ToString(CultureInfo.InvariantCulture);
                case 's': return value.Second.ToString(CultureInfo.InvariantCulture);
                default: return null;
            }
        }

        /// <summary>
        /// 按 Excel 分节语义选出本次取值应使用的格式节，并给出该节应格式化的数值。
        /// 财务报表普遍使用 "#,##0.00;(#,##0.00)"：负数须换用第二节并按绝对值渲染成括号，
        /// 若只取首节，(810,000.00) 会退化成 -810,000.00，与 Excel 显示不一致。
        /// Picks the format section Excel would use for this value and the value to render with it.
        /// Sections are positive;negative;zero;text: a negative number uses the second section with its
        /// absolute value (so "#,##0.00;(#,##0.00)" prints (810,000.00), not -810,000.00), and zero uses
        /// the third section when present, otherwise the positive section.
        /// </summary>
        /// <param name="format">完整 Excel 格式代码。The complete Excel format code.</param>
        /// <param name="value">待格式化的值。The value to format.</param>
        /// <param name="target">该节应格式化的值（负数取绝对值后返回）。The value to format with that section (negative numbers are returned as their absolute value).</param>
        /// <returns>应使用的格式节文本。The format section text to use.</returns>
        private static string SelectSection(string format, IFormattable value, out IFormattable target)
        {
            string[] sections = SplitSections(format);
            target = value;
            if (sections.Length == 1) return sections[0];

            bool numeric;
            bool negative = false;
            bool zero = false;
            IConvertible convertible = value as IConvertible;
            if (convertible == null) numeric = false;
            else
            {
                switch (convertible.GetTypeCode())
                {
                    case TypeCode.Double:
                    case TypeCode.Single:
                    case TypeCode.Decimal:
                    case TypeCode.SByte:
                    case TypeCode.Byte:
                    case TypeCode.Int16:
                    case TypeCode.UInt16:
                    case TypeCode.Int32:
                    case TypeCode.UInt32:
                    case TypeCode.Int64:
                    case TypeCode.UInt64:
                        decimal magnitude = convertible.ToDecimal(CultureInfo.InvariantCulture);
                        negative = magnitude < 0m;
                        zero = magnitude == 0m;
                        numeric = true;
                        break;
                    default:
                        numeric = false;
                        break;
                }
            }
            if (!numeric) return sections[0];
            if (negative)
            {
                target = Absolute(convertible);
                return sections[1];
            }
            if (zero && sections.Length >= 3) return sections[2];
            return sections[0];
        }

        /// <summary>取数值的绝对值并保留原始数值类型，避免整数与 decimal 转成 double 后丢精度。Takes the absolute value while preserving the original numeric type, avoiding the precision loss of casting integers or decimals to double.</summary>
        /// <param name="value">数值型可转换值。The numeric convertible value.</param>
        /// <returns>同类型的绝对值。The absolute value of the same type.</returns>
        private static IFormattable Absolute(IConvertible value)
        {
            switch (value.GetTypeCode())
            {
                case TypeCode.Double: return Math.Abs(value.ToDouble(CultureInfo.InvariantCulture));
                case TypeCode.Single: return Math.Abs(value.ToSingle(CultureInfo.InvariantCulture));
                case TypeCode.Decimal: return Math.Abs(value.ToDecimal(CultureInfo.InvariantCulture));
                case TypeCode.Int64:
                    {
                        long big = value.ToInt64(CultureInfo.InvariantCulture);
                        if (big == long.MinValue) return (double)long.MaxValue;
                        return big < 0 ? -big : big;
                    }
                default:
                    {
                        int small = value.ToInt32(CultureInfo.InvariantCulture);
                        return small < 0 ? -small : small;
                    }
            }
        }

        /// <summary>
        /// 按分号拆分 Excel 格式代码的各节。分号写在双引号里时只是字面文本，
        /// 例如会计格式 "¥"#,##0.00" ";"-"，不能据此切分。
        /// Splits a format code into its semicolon-separated sections; a semicolon inside double quotes
        /// is literal text (as in accounting formats) and does not split.
        /// </summary>
        /// <param name="format">完整格式代码。The complete format code.</param>
        /// <returns>至少一节；末尾空节会被保留（第三节为空表示零值显示为空）。At least one section; trailing empty sections are kept (an empty third section means zero values render as empty).</returns>
        private static string[] SplitSections(string format)
        {
            if (string.IsNullOrEmpty(format)) return new string[] { string.Empty };
            List<string> sections = new List<string>();
            StringBuilder current = new StringBuilder(format.Length);
            bool inQuote = false;
            for (int i = 0; i < format.Length; i++)
            {
                char c = format[i];
                if (c == '"')
                {
                    inQuote = !inQuote;
                    current.Append(c);
                    continue;
                }
                if (c == ';' && !inQuote)
                {
                    sections.Add(current.ToString());
                    current.Length = 0;
                    continue;
                }
                current.Append(c);
            }
            sections.Add(current.ToString());
            // 只有一节时保持单节语义；尾部空节（如 "0.00;"）表示负数沿用第一节。
            if (sections.Count == 2 && sections[1].Length == 0) return new string[] { sections[0] };
            string[] result = new string[sections.Count];
            sections.CopyTo(result);
            return result;
        }

        /// <summary>将 Excel 数字格式代码转换为 .NET 数字格式：剔除方括号条件、下划线/星号占位，转义反斜杠字符，"?" 替换为 "#"。Converts an Excel number-format code into a .NET number format: drops bracketed conditions and underscore/asterisk placeholders, escapes backslash characters, and replaces "?" with "#".</summary>
        /// <param name="format">Excel 数字格式代码（单个节）。The Excel number-format code (a single section).</param>
        /// <returns>可用的 .NET 格式；General、@ 或空输入返回空字符串。The usable .NET format; General, @ or empty input returns an empty string.</returns>
        private static string ToNumberPattern(string format)
        {
            if (string.IsNullOrEmpty(format)) return string.Empty;
            string section = FirstSection(format);
            if (section == "General" || section == "@") return string.Empty;

            StringBuilder builder = new StringBuilder(section.Length);
            for (int i = 0; i < section.Length; i++)
            {
                char c = section[i];
                if (c == '[')
                {
                    int close = section.IndexOf(']', i + 1);
                    if (close < 0) break;
                    i = close;
                    continue;
                }
                if ((c == '_' || c == '*') && i + 1 < section.Length)
                {
                    i++;
                    continue;
                }
                if (c == '\\' && i + 1 < section.Length)
                {
                    builder.Append('\\').Append(section[i + 1]);
                    i++;
                    continue;
                }
                if (c == '?')
                {
                    builder.Append('#');
                    continue;
                }
                // $ 在 .NET 自定义数值格式里是货币说明符（不变区域会输出 ¤），Excel 里只是字面货币符号。
                if (c == '$')
                {
                    builder.Append("\\$");
                    continue;
                }
                builder.Append(c);
            }
            return builder.ToString().Trim();
        }

        /// <summary>将 Excel 日期格式代码逐段翻译为 .NET 日期模式：处理引号字面量、方括号小时、AM/PM、年月日时分秒令牌，并区分 "m" 是月份还是分钟。Translates an Excel date-format code into a .NET date pattern segment by segment: handling quoted literals, bracketed hours, AM/PM and year/month/day/hour/minute/second tokens, and disambiguating whether "m" means month or minute.</summary>
        /// <param name="format">Excel 日期格式代码。The Excel date-format code.</param>
        /// <returns>.NET 日期模式；无法识别时回退 "yyyy-MM-dd"。The .NET date pattern; falls back to "yyyy-MM-dd" when unrecognizable.</returns>
        private static string ToDatePattern(string format)
        {
            if (string.IsNullOrEmpty(format)) return "yyyy-MM-dd";
            string section = FirstSection(format);
            if (section == "General" || section == "@") return "yyyy-MM-dd";

            bool[] literals = BuildLiteralMask(section);
            StringBuilder builder = new StringBuilder(section.Length + 4);
            for (int i = 0; i < section.Length; i++)
            {
                char c = section[i];
                if (c == '"')
                {
                    int end = section.IndexOf('"', i + 1);
                    if (end < 0) end = section.Length;
                    string text = section.Substring(i + 1, Math.Max(0, end - i - 1));
                    builder.Append('\'').Append(text.Replace("'", "''")).Append('\'');
                    i = end;
                    continue;
                }
                if (c == '[')
                {
                    int close = section.IndexOf(']', i + 1);
                    if (close < 0) { i = section.Length; continue; }
                    string inner = section.Substring(i + 1, close - i - 1);
                    if (inner.IndexOf('h') >= 0 || inner.IndexOf('H') >= 0)
                    {
                        builder.Append('h');
                    }
                    i = close;
                    continue;
                }
                if ((c == '_' || c == '*') && i + 1 < section.Length)
                {
                    if (c == '_') builder.Append(' ');
                    i++;
                    continue;
                }
                if (c == '\\' && i + 1 < section.Length)
                {
                    builder.Append('\\').Append(section[i + 1]);
                    i++;
                    continue;
                }
                if ((c == 'A' || c == 'a') && IsAmPmMarker(section, i, out int markerLength))
                {
                    builder.Append("tt");
                    i += markerLength - 1;
                    continue;
                }
                if (literals[i])
                {
                    builder.Append(c);
                    continue;
                }
                if (IsDateToken(c))
                {
                    int end = i;
                    while (end + 1 < section.Length && char.ToLowerInvariant(section[end + 1]) == char.ToLowerInvariant(c)
                        && !literals[end + 1])
                    {
                        end++;
                    }
                    int count = end - i + 1;
                    if (c == 'm' || c == 'M')
                    {
                        char token = IsMinuteMarker(section, i, literals) ? 'm' : 'M';
                        for (int k = 0; k < count; k++) builder.Append(token);
                    }
                    else if (c == 'a' || c == 'A')
                    {
                        builder.Append(count >= 4 ? "dddd" : "ddd");
                    }
                    else if (c == 'e' || c == 'E')
                    {
                        builder.Append("yyyy");
                    }
                    else if (c == 'h' || c == 'H')
                    {
                        builder.Append(count >= 2 ? "HH" : "H");
                    }
                    else if (c == 'y' || c == 'Y' || c == 'd' || c == 'D' || c == 's' || c == 'S')
                    {
                        char token = char.ToLowerInvariant(c);
                        for (int k = 0; k < count; k++) builder.Append(token);
                    }
                    i = end;
                    continue;
                }
                if (char.IsLetter(c))
                {
                    builder.Append('\\').Append(c);
                    continue;
                }
                builder.Append(c);
            }
            string pattern = builder.ToString();
            return pattern.Length == 0 ? "yyyy-MM-dd" : pattern;
        }

        /// <summary>取格式代码分号分隔的第一节（正数节）。Takes the first (positive) semicolon-separated section of a format code.</summary>
        /// <param name="format">完整格式代码。The complete format code.</param>
        /// <returns>第一节内容；无分号时即原串。The first section; the original string when there is no semicolon.</returns>
        private static string FirstSection(string format)
        {
            int semicolon = format.IndexOf(';');
            return semicolon >= 0 ? format.Substring(0, semicolon) : format;
        }

        /// <summary>构建字面量掩码：标记双引号内（含引号本身）的字符，避免其被当作日期令牌。Builds a literal mask marking characters inside double quotes (and the quotes themselves) so they are not treated as date tokens.</summary>
        /// <param name="text">格式节文本。The format section text.</param>
        /// <returns>与输入等长的布尔数组，true 表示该位置为字面量。A boolean array as long as the input; true means the position is a literal.</returns>
        private static bool[] BuildLiteralMask(string text)
        {
            bool[] mask = new bool[text.Length];
            bool inQuote = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '"')
                {
                    mask[i] = true;
                    inQuote = !inQuote;
                    continue;
                }
                mask[i] = inQuote;
            }
            return mask;
        }

        /// <summary>判断字符是否为 Excel 日期/时间令牌（y、m、d、h、s、e、g、a）。Reports whether a character is an Excel date/time token (y, m, d, h, s, e, g, a).</summary>
        /// <param name="c">待判断字符。The character to test.</param>
        /// <returns>是否日期令牌。Whether it is a date token.</returns>
        private static bool IsDateToken(char c)
        {
            char lower = char.ToLowerInvariant(c);
            return lower == 'y' || lower == 'm' || lower == 'd' || lower == 'h' || lower == 's'
                || lower == 'e' || lower == 'g' || lower == 'a';
        }

        /// <summary>判断 "m" 令牌是分钟还是月份：就近搜索非字面量上下文，先遇到 h/s 判为分钟，先遇到 y/d 判为月份。Decides whether an "m" token means minute or month by searching nearby non-literal context: an h/s encountered first means minute, a y/d encountered first means month.</summary>
        /// <param name="text">格式节文本。The format section text.</param>
        /// <param name="index">"m" 令牌起始位置。The start index of the "m" token.</param>
        /// <param name="literals">字面量掩码。The literal mask.</param>
        /// <returns>是分钟返回 true，否则为月份。True when it is a minute, otherwise a month.</returns>
        private static bool IsMinuteMarker(string text, int index, bool[] literals)
        {
            for (int i = index - 1; i >= 0; i--)
            {
                if (literals[i]) continue;
                char c = char.ToLowerInvariant(text[i]);
                if (c == 'h' || c == 's') return true;
                if (c == 'y' || c == 'd') return false;
                if (c == 'm' || c == 'g' || c == 'e') continue;
            }
            for (int i = index + 1; i < text.Length; i++)
            {
                if (literals[i]) continue;
                char c = char.ToLowerInvariant(text[i]);
                if (c == 'h' || c == 's') return true;
                if (c == 'y' || c == 'd') return false;
                if (c == 'm' || c == 'g' || c == 'e') continue;
            }
            return false;
        }

        /// <summary>判断指定位置是否为 "AM/PM" 或 "A/P" 上午下午标记。Reports whether the given position is an "AM/PM" or "A/P" day-period marker.</summary>
        /// <param name="text">格式节文本。The format section text.</param>
        /// <param name="index">待检查位置（应为 'a'/'A'）。The position to check (expected to be 'a'/'A').</param>
        /// <param name="length">匹配到的标记长度（5 或 3），未匹配为 0。The length of the matched marker (5 or 3), or 0 when unmatched.</param>
        /// <returns>是否为 AM/PM 标记。Whether it is an AM/PM marker.</returns>
        internal static bool IsAmPmMarker(string text, int index, out int length)
        {
            length = 0;
            if (Matches(text, index, "AM/PM"))
            {
                length = 5;
                return true;
            }
            if (Matches(text, index, "A/P"))
            {
                length = 3;
                return true;
            }
            return false;
        }

        /// <summary>忽略大小写地判断 text 从 index 起是否以 value 开头。Case-insensitively checks whether text starts with value at the given index.</summary>
        /// <param name="text">源文本。The source text.</param>
        /// <param name="index">起始位置。The starting index.</param>
        /// <param name="value">待匹配片段。The fragment to match.</param>
        /// <returns>是否匹配。Whether it matches.</returns>
        private static bool Matches(string text, int index, string value)
        {
            if (index + value.Length > text.Length) return false;
            return string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
        }
    }

    /// <summary>合并单元格区域：1 基起始行列加行列跨度。A merged cell range: a 1-based start row/column plus row/column spans.</summary>
    public struct MergedRange
    {
        /// <summary>起始行号（从 1 开始）。The starting row number (starting from 1).</summary>
        public int Row;
        /// <summary>起始列号（从 1 开始）。The starting column number (starting from 1).</summary>
        public int Column;
        /// <summary>合并的行数（至少为 1）。The number of merged rows (at least 1).</summary>
        public int RowSpan;
        /// <summary>合并的列数（至少为 1）。The number of merged columns (at least 1).</summary>
        public int ColumnSpan;

        /// <summary>创建合并区域，小于 1 的跨度自动钳制为 1。Creates a merged range, clamping spans below 1 to 1.</summary>
        /// <param name="row">起始行号。The starting row number.</param>
        /// <param name="column">起始列号。The starting column number.</param>
        /// <param name="rowSpan">合并行数。The number of merged rows.</param>
        /// <param name="columnSpan">合并列数。The number of merged columns.</param>
        public MergedRange(int row, int column, int rowSpan, int columnSpan)
        {
            Row = row;
            Column = column;
            RowSpan = rowSpan < 1 ? 1 : rowSpan;
            ColumnSpan = columnSpan < 1 ? 1 : columnSpan;
        }

        /// <summary>区域最后一行的行号。The row number of the last row in the range.</summary>
        public int LastRow { get { return Row + RowSpan - 1; } }
        /// <summary>区域最后一列的列号。The column number of the last column in the range.</summary>
        public int LastColumn { get { return Column + ColumnSpan - 1; } }

        /// <summary>判断指定坐标是否落在合并区域内。Reports whether the given coordinate falls inside the merged range.</summary>
        /// <param name="row">行号。The row number.</param>
        /// <param name="column">列号。The column number.</param>
        /// <returns>是否在区域内。Whether it is inside the range.</returns>
        public bool Contains(int row, int column)
        {
            return row >= Row && row <= LastRow && column >= Column && column <= LastColumn;
        }
    }

    /// <summary>数据验证的取值类型。The value type of a data-validation rule.</summary>
    public enum DataValidationType
    {
        /// <summary>不限制类型。No type restriction.</summary>
        None = 0,
        /// <summary>整数。Whole number.</summary>
        Whole = 1,
        /// <summary>小数。Decimal number.</summary>
        Decimal = 2,
        /// <summary>下拉列表。Drop-down list.</summary>
        List = 3,
        /// <summary>日期。Date.</summary>
        Date = 4,
        /// <summary>时间。Time.</summary>
        Time = 5,
        /// <summary>文本长度。Text length.</summary>
        TextLength = 6,
        /// <summary>自定义公式。Custom formula.</summary>
        Custom = 7
    }

    /// <summary>数据验证的比较运算符。The comparison operator of a data-validation rule.</summary>
    public enum DataValidationOperator
    {
        /// <summary>介于两者之间。Between the two formulas.</summary>
        Between = 0,
        /// <summary>不介于两者之间。Not between the two formulas.</summary>
        NotBetween = 1,
        /// <summary>等于。Equal to.</summary>
        Equal = 2,
        /// <summary>不等于。Not equal to.</summary>
        NotEqual = 3,
        /// <summary>小于。Less than.</summary>
        LessThan = 4,
        /// <summary>小于等于。Less than or equal to.</summary>
        LessThanOrEqual = 5,
        /// <summary>大于。Greater than.</summary>
        GreaterThan = 6,
        /// <summary>大于等于。Greater than or equal to.</summary>
        GreaterThanOrEqual = 7
    }

    /// <summary>数据验证失败时 Excel 的提示样式。The Excel alert style used when a data-validation rule fails.</summary>
    public enum DataValidationErrorStyle
    {
        /// <summary>阻止输入。Blocks the entry.</summary>
        Stop = 0,
        /// <summary>警告后可继续。Warns but allows the entry.</summary>
        Warning = 1,
        /// <summary>仅提示信息。Shows information only.</summary>
        Information = 2
    }

    /// <summary>一条数据验证规则：作用区域、类型、公式与提示文本。A data-validation rule: target range, type, formulas and message texts.</summary>
    public sealed class ExcelDataValidation
    {
        /// <summary>作用区域，A1 风格引用（如 "A2:A100"）。The target range as an A1-style reference (such as "A2:A100").</summary>
        public string Range { get; set; }
        /// <summary>验证类型。The validation type.</summary>
        public DataValidationType Type { get; set; }
        /// <summary>比较运算符，仅对数值/日期/文本长度类型有效。The comparison operator, used by number, date and text-length types.</summary>
        public DataValidationOperator Operator { get; set; }
        /// <summary>第一个公式：列表用逗号分隔的取值或区域引用，其他类型为比较值。The first formula: comma-separated values or a range for lists, the comparison value otherwise.</summary>
        public string Formula1 { get; set; }
        /// <summary>第二个公式，仅 Between/NotBetween 需要。The second formula, required only by Between and NotBetween.</summary>
        public string Formula2 { get; set; }
        /// <summary>是否允许空值，默认允许。Whether blank values are allowed; enabled by default.</summary>
        public bool AllowBlank { get; set; }
        /// <summary>是否显示下拉箭头；false 对应 OOXML 的 showDropDown="1"（隐藏箭头）。Whether the drop-down arrow is shown; false maps to OOXML showDropDown="1" (arrow suppressed).</summary>
        public bool ShowDropDown { get; set; }
        /// <summary>输入时是否显示提示。Whether the input prompt is shown.</summary>
        public bool ShowInputMessage { get; set; }
        /// <summary>输入不合法时是否显示错误提示。Whether an error alert is shown for invalid input.</summary>
        public bool ShowErrorMessage { get; set; }
        /// <summary>输入提示的标题。The title of the input prompt.</summary>
        public string PromptTitle { get; set; }
        /// <summary>输入提示的正文。The body of the input prompt.</summary>
        public string Prompt { get; set; }
        /// <summary>错误提示的标题。The title of the error alert.</summary>
        public string ErrorTitle { get; set; }
        /// <summary>错误提示的正文。The body of the error alert.</summary>
        public string ErrorMessage { get; set; }
        /// <summary>错误提示样式。The error alert style.</summary>
        public DataValidationErrorStyle ErrorStyle { get; set; }

        /// <summary>创建默认规则：允许空值、显示下拉箭头。Creates a default rule that allows blanks and shows the drop-down arrow.</summary>
        public ExcelDataValidation()
        {
            Range = "A1";
            Formula1 = string.Empty;
            Formula2 = string.Empty;
            AllowBlank = true;
            ShowDropDown = true;
            PromptTitle = string.Empty;
            Prompt = string.Empty;
            ErrorTitle = string.Empty;
            ErrorMessage = string.Empty;
        }
    }

    /// <summary>自动筛选中的一列条件：列偏移与允许的取值。One filter condition of an auto-filter: the column offset and the allowed values.</summary>
    public sealed class ExcelFilterColumn
    {
        /// <summary>列在筛选区域内的 0 基偏移（0 表示区域第一列）。The 0-based column offset inside the filter range (0 is the first column).</summary>
        public int ColumnId { get; set; }
        /// <summary>允许的取值文本；为空表示不限制。The allowed value texts; empty means no restriction.</summary>
        public List<string> Values { get; private set; }
        /// <summary>是否同时接受空单元格。Whether blank cells are accepted as well.</summary>
        public bool IncludeBlanks { get; set; }

        /// <summary>创建指定列偏移的筛选条件。Creates a filter condition for the given column offset.</summary>
        /// <param name="columnId">0 基列偏移。The 0-based column offset.</param>
        public ExcelFilterColumn(int columnId)
        {
            ColumnId = columnId;
            Values = new List<string>();
        }
    }

    /// <summary>自动筛选设置：作用区域与各列的取值条件。Auto-filter settings: the target range and per-column value conditions.</summary>
    public sealed class ExcelAutoFilter
    {
        /// <summary>筛选区域，A1 风格引用（如 "A1:D100"）。The filter range as an A1-style reference (such as "A1:D100").</summary>
        public string Range { get; set; }
        /// <summary>各列筛选条件。The per-column filter conditions.</summary>
        public List<ExcelFilterColumn> Columns { get; private set; }

        /// <summary>基于区域引用创建自动筛选。Creates an auto-filter for a range reference.</summary>
        /// <param name="range">A1 风格区域引用。The A1-style range reference.</param>
        public ExcelAutoFilter(string range)
        {
            Range = range;
            Columns = new List<ExcelFilterColumn>();
        }

        /// <summary>取得指定列偏移的筛选条件，不存在时创建。Gets the filter condition for a column offset, creating it when absent.</summary>
        /// <param name="columnId">0 基列偏移。The 0-based column offset.</param>
        /// <returns>筛选条件。The filter condition.</returns>
        public ExcelFilterColumn ForColumn(int columnId)
        {
            for (int i = 0; i < Columns.Count; i++)
            {
                if (Columns[i].ColumnId == columnId) return Columns[i];
            }
            ExcelFilterColumn column = new ExcelFilterColumn(columnId);
            Columns.Add(column);
            return column;
        }
    }

    /// <summary>工作簿定义名：全局或绑定到某张工作表的公式别名。A workbook defined name: a formula alias, either global or scoped to one worksheet.</summary>
    public sealed class ExcelDefinedName
    {
        /// <summary>定义名，如 "TaxRate"。The defined name, such as "TaxRate".</summary>
        public string Name { get; set; }
        /// <summary>引用或公式文本，如 "Sheet1!$A$1:$A$5"。The reference or formula text, such as "Sheet1!$A$1:$A$5".</summary>
        public string Reference { get; set; }
        /// <summary>局部作用域的工作表下标（从 0 开始）；-1 表示全局。The worksheet index (0-based) for a local scope; -1 means global.</summary>
        public int LocalSheetId { get; set; }

        /// <summary>创建全局定义名。Creates a global defined name.</summary>
        public ExcelDefinedName()
        {
            Name = string.Empty;
            Reference = string.Empty;
            LocalSheetId = -1;
        }

        /// <summary>创建指定名称与引用的定义名。Creates a defined name with the given name and reference.</summary>
        /// <param name="name">定义名。The defined name.</param>
        /// <param name="reference">引用或公式文本。The reference or formula text.</param>
        public ExcelDefinedName(string name, string reference)
        {
            Name = name;
            Reference = reference;
            LocalSheetId = -1;
        }
    }

    /// <summary>工作表模型：稀疏存储单元格，并提供行高列宽、合并、冻结、隐藏与打印区域等设置。A worksheet model: stores cells sparsely and offers row/column sizing, merging, freezing, hiding and print-area settings.</summary>
    public sealed class Worksheet
    {
        /// <summary>稀疏单元格字典，键由 <see cref="MakeKey"/> 生成并按行优先排序。The sparse cell dictionary whose keys are produced by <see cref="MakeKey"/> and ordered row-first.</summary>
        private readonly SortedDictionary<long, Cell> _cells = new SortedDictionary<long, Cell>();

        /// <summary>工作表名称。The worksheet name.</summary>
        public string Name { get; set; }
        /// <summary>合并单元格区域列表。The list of merged cell ranges.</summary>
        public List<MergedRange> MergedRanges { get; private set; }
        /// <summary>自定义列宽（键为 1 基列号，值为 Excel 字符单位）。Custom column widths (key is a 1-based column number, value is in Excel character units).</summary>
        public Dictionary<int, double> ColumnWidths { get; private set; }
        /// <summary>自定义行高（键为 1 基行号，值为磅）。Custom row heights (key is a 1-based row number, value is in points).</summary>
        public Dictionary<int, double> RowHeights { get; private set; }
        /// <summary>隐藏的行号集合。The set of hidden row numbers.</summary>
        public HashSet<int> HiddenRows { get; private set; }
        /// <summary>隐藏的列号集合。The set of hidden column numbers.</summary>
        public HashSet<int> HiddenColumns { get; private set; }
        /// <summary>工作表本身是否隐藏。Whether the worksheet itself is hidden.</summary>
        public bool IsHidden { get; set; }
        /// <summary>是否显示网格线。Whether gridlines are shown.</summary>
        public bool ShowGridLines { get; set; }
        /// <summary>零值是否显示。Whether zero values are shown.</summary>
        public bool ShowZeros { get; set; }
        /// <summary>顶部冻结的行数。The number of frozen rows at the top.</summary>
        public int FrozenRows { get; set; }
        /// <summary>左侧冻结的列数。The number of frozen columns on the left.</summary>
        public int FrozenColumns { get; set; }
        /// <summary>打印用页面设置（纸张与页边距，单位磅）。The page setup used for printing (paper size and margins, in points).</summary>
        public PageSetup PrintSetup { get; set; }
        /// <summary>打印时是否缩放内容至一页。Whether to scale content to one page when printing.</summary>
        public bool FitToPage { get; set; }
        /// <summary>打印时是否水平居中。Whether to center content horizontally when printing.</summary>
        public bool CenterHorizontally { get; set; }
        /// <summary>打印区域首行；0 表示未设置。The first row of the print area; 0 means unset.</summary>
        public int PrintAreaFirstRow { get; set; }
        /// <summary>打印区域末行。The last row of the print area.</summary>
        public int PrintAreaLastRow { get; set; }
        /// <summary>打印区域首列；0 表示未设置。The first column of the print area; 0 means unset.</summary>
        public int PrintAreaFirstColumn { get; set; }
        /// <summary>打印区域末列。The last column of the print area.</summary>
        public int PrintAreaLastColumn { get; set; }
        /// <summary>默认列宽，单位 Excel 字符（默认 8.43）。The default column width in Excel character units (8.43 by default).</summary>
        public double DefaultColumnWidth { get; set; }
        /// <summary>默认行高，单位磅（默认 15）。The default row height in points (15 by default).</summary>
        public double DefaultRowHeight { get; set; }
        /// <summary>表级默认单元格样式。The sheet-level default cell style.</summary>
        public CellStyle DefaultStyle { get; set; }
        /// <summary>嵌入本表的图表列表。The charts embedded in this worksheet.</summary>
        public List<ExcelChart> Charts { get; private set; }
        /// <summary>锚定到本表单元格的图片列表。The images anchored to cells of this worksheet.</summary>
        public List<ExcelImage> Images { get; private set; }
        /// <summary>本表的条件格式集合。The conditional-format sets of this worksheet.</summary>
        public List<ExcelConditionalFormat> ConditionalFormats { get; private set; }
        /// <summary>工作表标签颜色，null 表示默认颜色。The sheet tab color; null means the default color.</summary>
        public OfficeColor? TabColor { get; set; }
        /// <summary>自动筛选设置，null 表示未启用。The auto-filter settings; null means filtering is disabled.</summary>
        public ExcelAutoFilter AutoFilter { get; set; }
        /// <summary>本表的数据验证规则集合。The data-validation rules of this worksheet.</summary>
        public List<ExcelDataValidation> DataValidations { get; private set; }
        /// <summary>创建指定名称的工作表；名称为空时使用 "Sheet1"，默认显示网格线与零值。Creates a worksheet with the given name, falling back to "Sheet1" when empty; gridlines and zero values are shown by default.</summary>
        /// <param name="name">工作表名称。The worksheet name.</param>
        public Worksheet(string name)
        {
            Name = string.IsNullOrEmpty(name) ? "Sheet1" : name;
            MergedRanges = new List<MergedRange>();
            Charts = new List<ExcelChart>();
            Images = new List<ExcelImage>();
            ConditionalFormats = new List<ExcelConditionalFormat>();
            DataValidations = new List<ExcelDataValidation>();
            ColumnWidths = new Dictionary<int, double>();
            RowHeights = new Dictionary<int, double>();
            HiddenRows = new HashSet<int>();
            HiddenColumns = new HashSet<int>();
            ShowGridLines = true;
            ShowZeros = true;
            PrintSetup = new PageSetup();
            DefaultColumnWidth = 8.43;
            DefaultRowHeight = 15;
            DefaultStyle = new CellStyle();
        }

        /// <summary>将 1 基行/列号打包为排序键：行号左移 20 位后按列号取序。Packs 1-based row/column numbers into a sort key: the row is shifted left by 20 bits and ordered by column.</summary>
        /// <param name="row">行号。The row number.</param>
        /// <param name="column">列号。The column number.</param>
        /// <returns>单元格字典键。The cell dictionary key.</returns>
        internal static long MakeKey(int row, int column)
        {
            return ((long)row << 20) | (uint)column;
        }

        /// <summary>取得指定坐标的单元格；不存在时自动创建。Gets the cell at the given coordinates, creating it when absent.</summary>
        /// <param name="row">行号（从 1 开始）。The row number (starting from 1).</param>
        /// <param name="column">列号（从 1 开始）。The column number (starting from 1).</param>
        /// <returns>已有或新建的单元格。The existing or newly created cell.</returns>
        public Cell GetCell(int row, int column)
        {
            long key = MakeKey(row, column);
            Cell cell;
            if (!_cells.TryGetValue(key, out cell))
            {
                cell = new Cell(row, column);
                _cells.Add(key, cell);
            }
            return cell;
        }

        /// <summary>按 A1 引用（如 "B3"、"$B$3"、可带 "Sheet!" 前缀）取得单元格，不存在时创建。Gets a cell by A1 reference (such as "B3", "$B$3", optionally with a "Sheet!" prefix), creating it when absent.</summary>
        /// <param name="reference">单元格引用字符串。The cell reference string.</param>
        /// <returns>已有或新建的单元格。The existing or newly created cell.</returns>
        public Cell GetCell(string reference)
        {
            int row;
            int column;
            CellAddress.Parse(reference, out row, out column);
            return GetCell(row, column);
        }

        /// <summary>查找已存在的单元格，不创建新实例。Finds an existing cell without creating a new instance.</summary>
        /// <param name="row">行号（从 1 开始）。The row number (starting from 1).</param>
        /// <param name="column">列号（从 1 开始）。The column number (starting from 1).</param>
        /// <returns>找到的单元格；不存在返回 null。The found cell, or null when it does not exist.</returns>
        public Cell FindCell(int row, int column)
        {
            Cell cell;
            _cells.TryGetValue(MakeKey(row, column), out cell);
            return cell;
        }

        /// <summary>全部已存储单元格的枚举，按行优先顺序排列。An enumeration of all stored cells, ordered row-first.</summary>
        public IEnumerable<Cell> Cells
        {
            get { return _cells.Values; }
        }

        /// <summary>已用区域的最后一行：包含单元格与合并区域；被有意创建的空单元格同样计入。The last row of the used range, counting both cells and merged ranges; intentionally created empty cells are also included.</summary>
        public int LastRow
        {
            get
            {
                int last = 0;
                foreach (Cell cell in _cells.Values)
                {
                    if (cell.Row > last) last = cell.Row;
                    if (cell.Value != null || cell.Formula != null || cell.Style != null)
                    {
                        // 内容为空的单元格一旦被有意创建（已物化），仍计入已用区域。
                    }
                }
                for (int i = 0; i < MergedRanges.Count; i++)
                {
                    int end = MergedRanges[i].LastRow;
                    if (end > last) last = end;
                }
                return last;
            }
        }

        /// <summary>已用区域的最后一列：包含单元格与合并区域。The last column of the used range, counting both cells and merged ranges.</summary>
        public int LastColumn
        {
            get
            {
                int last = 0;
                foreach (Cell cell in _cells.Values)
                {
                    if (cell.Column > last) last = cell.Column;
                }
                for (int i = 0; i < MergedRanges.Count; i++)
                {
                    int end = MergedRanges[i].LastColumn;
                    if (end > last) last = end;
                }
                return last;
            }
        }

        /// <summary>若坐标位于某个合并区域内，返回该区域锚点（左上角）单元格；否则返回 null。If the coordinate lies in a merged range, returns that range's anchor (top-left) cell; otherwise returns null.</summary>
        /// <param name="row">行号。The row number.</param>
        /// <param name="column">列号。The column number.</param>
        /// <returns>合并区域锚点单元格或 null。The merged range's anchor cell or null.</returns>
        public Cell FindMergeAnchor(int row, int column)
        {
            for (int i = 0; i < MergedRanges.Count; i++)
            {
                MergedRange range = MergedRanges[i];
                if (range.Contains(row, column))
                {
                    return FindCell(range.Row, range.Column);
                }
            }
            return null;
        }

        /// <summary>按范围引用（如 "A1:B3"）添加一个合并区域。Adds a merged range from a range reference (such as "A1:B3").</summary>
        /// <param name="rangeReference">A1 风格范围引用。An A1-style range reference.</param>
        public void Merge(string rangeReference)
        {
            int row1, column1, row2, column2;
            CellAddress.ParseRange(rangeReference, out row1, out column1, out row2, out column2);
            MergedRanges.Add(new MergedRange(row1, column1, row2 - row1 + 1, column2 - column1 + 1));
        }

        /// <summary>解析范围引用并设置打印区域，自动规范化首末行列顺序。Parses a range reference and sets the print area, normalizing the first/last row and column order.</summary>
        /// <param name="reference">A1 风格范围引用。An A1-style range reference.</param>
        public void SetPrintArea(string reference)
        {
            int row1, column1, row2, column2;
            CellAddress.ParseRange(reference, out row1, out column1, out row2, out column2);
            PrintAreaFirstRow = Math.Min(row1, row2);
            PrintAreaLastRow = Math.Max(row1, row2);
            PrintAreaFirstColumn = Math.Min(column1, column2);
            PrintAreaLastColumn = Math.Max(column1, column2);
        }

        /// <summary>是否已设置打印区域（首行首列均大于 0）。Whether a print area is set (both the first row and first column are greater than 0).</summary>
        public bool HasPrintArea
        {
            get { return PrintAreaFirstRow > 0 && PrintAreaFirstColumn > 0; }
        }

        /// <summary>取得指定列宽（Excel 字符单位），未自定义时返回默认列宽。Gets the width of a column in Excel character units, returning the default width when it is not customized.</summary>
        /// <param name="column">列号（从 1 开始）。The column number (starting from 1).</param>
        /// <returns>列宽。The column width.</returns>
        public double GetColumnWidth(int column)
        {
            double width;
            if (ColumnWidths.TryGetValue(column, out width))
            {
                return width;
            }
            return DefaultColumnWidth;
        }

        /// <summary>将列宽换算为磅：Excel 按"字符"存储宽度，96 DPI 下每单位约 7 像素，另加 5 磅单元格边距。Converts a column width to points: Excel stores width in "characters", each unit being about 7 pixels at 96 DPI, plus a 5-point cell margin.</summary>
        /// <param name="column">列号（从 1 开始）。The column number (starting from 1).</param>
        /// <returns>以磅为单位的列宽。The column width in points.</returns>
        public double GetColumnWidthPoints(int column)
        {
            // Excel 以"字符"为单位存储列宽；96 DPI 下每单位约 7 像素。
            return GetColumnWidth(column) * 7.0 * 72.0 / 96.0 + 5.0;
        }

        /// <summary>取得指定行高（磅），未自定义时返回默认行高。Gets the height of a row in points, returning the default height when it is not customized.</summary>
        /// <param name="row">行号（从 1 开始）。The row number (starting from 1).</param>
        /// <returns>行高。The row height.</returns>
        public double GetRowHeight(int row)
        {
            double height;
            if (RowHeights.TryGetValue(row, out height))
            {
                return height;
            }
            return DefaultRowHeight;
        }

        /// <summary>设置指定列宽（Excel 字符单位）。Sets the width of a column in Excel character units.</summary>
        /// <param name="column">列号（从 1 开始）。The column number (starting from 1).</param>
        /// <param name="width">列宽。The column width.</param>
        public void SetColumnWidth(int column, double width)
        {
            ColumnWidths[column] = width;
        }

        /// <summary>设置指定行高（磅）。Sets the height of a row in points.</summary>
        /// <param name="row">行号（从 1 开始）。The row number (starting from 1).</param>
        /// <param name="height">行高。The row height.</param>
        public void SetRowHeight(int row, double height)
        {
            RowHeights[row] = height;
        }

        /// <summary>显示或隐藏指定行。Shows or hides a given row.</summary>
        /// <param name="row">行号（从 1 开始）。The row number (starting from 1).</param>
        /// <param name="hidden">true 隐藏，false 显示。True hides the row, false shows it.</param>
        public void SetRowHidden(int row, bool hidden)
        {
            if (hidden) HiddenRows.Add(row);
            else HiddenRows.Remove(row);
        }

        /// <summary>显示或隐藏指定列。Shows or hides a given column.</summary>
        /// <param name="column">列号（从 1 开始）。The column number (starting from 1).</param>
        /// <param name="hidden">true 隐藏，false 显示。True hides the column, false shows it.</param>
        public void SetColumnHidden(int column, bool hidden)
        {
            if (hidden) HiddenColumns.Add(column);
            else HiddenColumns.Remove(column);
        }

        /// <summary>移除指定坐标的单元格（若存在）。Removes the cell at the given coordinates, if present.</summary>
        /// <param name="row">行号（从 1 开始）。The row number (starting from 1).</param>
        /// <param name="column">列号（从 1 开始）。The column number (starting from 1).</param>
        public void RemoveCell(int row, int column)
        {
            _cells.Remove(MakeKey(row, column));
        }

        /// <summary>在指定行之前插入空行：单元格、行高、隐藏标记、合并区域、自动筛选、数据验证与打印区域随之下移或扩展。Inserts empty rows before the given row; cells, row heights, hidden flags, merged ranges, the auto-filter, data validations and the print area shift down or expand.</summary>
        /// <param name="rowIndex">插入位置的行号（从 1 开始）。The row number where rows are inserted (1-based).</param>
        /// <param name="count">插入的行数。The number of rows to insert.</param>
        /// <exception cref="ArgumentOutOfRangeException">行号小于 1，或插入后超过 Excel 最大行数。The row is below 1, or the insertion exceeds the Excel row limit.</exception>
        public void InsertRows(int rowIndex, int count)
        {
            if (rowIndex < 1 || rowIndex > CellAddress.MaxRowLiteral) throw new ArgumentOutOfRangeException("rowIndex");
            if (count <= 0) return;
            if ((long)GetLastLayoutRow() + count > CellAddress.MaxRowLiteral)
            {
                throw new ArgumentOutOfRangeException("count");
            }
            MoveRowCells(rowIndex, CellAddress.MaxRowLiteral, count);
            RebuildRowHeightEntries(rowIndex, 0, count);
            RebuildHiddenRows(rowIndex, 0, count);
            InsertMergedRows(rowIndex, count);
            ShiftFilterRows(rowIndex, count, true);
            ShiftValidationRows(rowIndex, count, true);
            InsertPrintAreaRows(rowIndex, count);
        }

        /// <summary>删除指定范围的整行：范围内的单元格与行设置被移除，下方内容上移。Deletes whole rows in the given range: cells and row settings inside the range are removed and content below moves up.</summary>
        /// <param name="rowIndex">起始行号（从 1 开始）。The first row number (1-based).</param>
        /// <param name="count">删除的行数。The number of rows to delete.</param>
        public void DeleteRows(int rowIndex, int count)
        {
            if (rowIndex < 1 || rowIndex > CellAddress.MaxRowLiteral) throw new ArgumentOutOfRangeException("rowIndex");
            if (count <= 0) return;
            int end = (int)Math.Min(CellAddress.MaxRowLiteral, (long)rowIndex + count - 1);
            DeleteRowCells(rowIndex, end);
            RebuildRowHeightEntries(rowIndex, end, -count);
            RebuildHiddenRows(rowIndex, end, -count);
            DeleteMergedRows(rowIndex, end, count);
            ShiftFilterRows(rowIndex, count, false);
            ShiftValidationRows(rowIndex, count, false);
            DeletePrintAreaRows(rowIndex, end, count);
        }

        /// <summary>在指定列之前插入空列：单元格、列宽、隐藏标记、合并区域、自动筛选、数据验证与打印区域随之右移或扩展。Inserts empty columns before the given column; cells, column widths, hidden flags, merged ranges, the auto-filter, data validations and the print area shift right or expand.</summary>
        /// <param name="columnIndex">插入位置的列号（从 1 开始）。The column number where columns are inserted (1-based).</param>
        /// <param name="count">插入的列数。The number of columns to insert.</param>
        /// <exception cref="ArgumentOutOfRangeException">列号小于 1，或插入后超过 Excel 最大列数。The column is below 1, or the insertion exceeds the Excel column limit.</exception>
        public void InsertColumns(int columnIndex, int count)
        {
            if (columnIndex < 1 || columnIndex > CellAddress.MaxColumnLiteral) throw new ArgumentOutOfRangeException("columnIndex");
            if (count <= 0) return;
            if ((long)GetLastLayoutColumn() + count > CellAddress.MaxColumnLiteral)
            {
                throw new ArgumentOutOfRangeException("count");
            }
            MoveColumnCells(columnIndex, CellAddress.MaxColumnLiteral, count);
            RebuildColumnWidthEntries(columnIndex, 0, count);
            RebuildHiddenColumns(columnIndex, 0, count);
            InsertMergedColumns(columnIndex, count);
            ShiftFilterColumns(columnIndex, count, true);
            ShiftValidationColumns(columnIndex, count, true);
            InsertPrintAreaColumns(columnIndex, count);
        }

        /// <summary>删除指定范围的整列：范围内的单元格与列设置被移除，右侧内容左移。Deletes whole columns in the given range: cells and column settings inside the range are removed and content on the right moves left.</summary>
        /// <param name="columnIndex">起始列号（从 1 开始）。The first column number (1-based).</param>
        /// <param name="count">删除的列数。The number of columns to delete.</param>
        public void DeleteColumns(int columnIndex, int count)
        {
            if (columnIndex < 1 || columnIndex > CellAddress.MaxColumnLiteral) throw new ArgumentOutOfRangeException("columnIndex");
            if (count <= 0) return;
            int end = (int)Math.Min(CellAddress.MaxColumnLiteral, (long)columnIndex + count - 1);
            DeleteColumnCells(columnIndex, end);
            RebuildColumnWidthEntries(columnIndex, end, -count);
            RebuildHiddenColumns(columnIndex, end, -count);
            DeleteMergedColumns(columnIndex, end, count);
            ShiftFilterColumns(columnIndex, count, false);
            ShiftValidationColumns(columnIndex, count, false);
            DeletePrintAreaColumns(columnIndex, end, count);
        }

        /// <summary>取得单元格、合并区域、行高与隐藏行覆盖的最大行号。Gets the largest row covered by cells, merged ranges, row heights and hidden rows.</summary>
        /// <returns>最大行号；空表返回 0。The largest row number, or 0 for an empty sheet.</returns>
        private int GetLastLayoutRow()
        {
            int last = LastRow;
            foreach (int row in RowHeights.Keys)
            {
                if (row > last) last = row;
            }
            foreach (int row in HiddenRows)
            {
                if (row > last) last = row;
            }
            return last;
        }

        /// <summary>取得单元格、合并区域、列宽与隐藏列覆盖的最大列号。Gets the largest column covered by cells, merged ranges, column widths and hidden columns.</summary>
        /// <returns>最大列号；空表返回 0。The largest column number, or 0 for an empty sheet.</returns>
        private int GetLastLayoutColumn()
        {
            int last = LastColumn;
            foreach (int column in ColumnWidths.Keys)
            {
                if (column > last) last = column;
            }
            foreach (int column in HiddenColumns)
            {
                if (column > last) last = column;
            }
            return last;
        }

        /// <summary>把指定行区间内的单元格平移 delta 行；区间末尾传 Excel 最大行号表示直到末尾。Shifts the cells inside a row interval by delta rows; the Excel row limit as the end means through the last row.</summary>
        /// <param name="rowIndex">起始行号。The first row number.</param>
        /// <param name="endRow">结束行号。The last row number.</param>
        /// <param name="delta">行偏移。The row delta.</param>
        private void MoveRowCells(int rowIndex, int endRow, int delta)
        {
            List<Cell> moving = new List<Cell>();
            foreach (Cell cell in _cells.Values)
            {
                if (cell.Row >= rowIndex && cell.Row <= endRow) moving.Add(cell);
            }
            for (int i = 0; i < moving.Count; i++)
            {
                _cells.Remove(MakeKey(moving[i].Row, moving[i].Column));
            }
            for (int i = 0; i < moving.Count; i++)
            {
                Cell cell = moving[i];
                cell.MoveTo(cell.Row + delta, cell.Column);
                _cells.Add(MakeKey(cell.Row, cell.Column), cell);
            }
        }

        /// <summary>删除行区间内的单元格并把下方单元格上移。Removes the cells in a row range and moves the rows below up.</summary>
        /// <param name="rowIndex">起始行号。The first row number.</param>
        /// <param name="endRow">结束行号。The last row number.</param>
        private void DeleteRowCells(int rowIndex, int endRow)
        {
            int count = endRow - rowIndex + 1;
            List<Cell> moving = new List<Cell>();
            List<long> removed = new List<long>();
            foreach (KeyValuePair<long, Cell> pair in _cells)
            {
                if (pair.Value.Row >= rowIndex && pair.Value.Row <= endRow)
                {
                    removed.Add(pair.Key);
                }
                else if (pair.Value.Row > endRow)
                {
                    moving.Add(pair.Value);
                }
            }
            for (int i = 0; i < removed.Count; i++) _cells.Remove(removed[i]);
            for (int i = 0; i < moving.Count; i++)
            {
                Cell cell = moving[i];
                _cells.Remove(MakeKey(cell.Row, cell.Column));
                cell.MoveTo(cell.Row - count, cell.Column);
                _cells.Add(MakeKey(cell.Row, cell.Column), cell);
            }
        }

        /// <summary>把指定列区间内的单元格平移 delta 列。Shifts the cells inside a column interval by delta columns.</summary>
        /// <param name="columnIndex">起始列号。The first column number.</param>
        /// <param name="endColumn">结束列号。The last column number.</param>
        /// <param name="delta">列偏移。The column delta.</param>
        private void MoveColumnCells(int columnIndex, int endColumn, int delta)
        {
            List<Cell> moving = new List<Cell>();
            foreach (Cell cell in _cells.Values)
            {
                if (cell.Column >= columnIndex && cell.Column <= endColumn) moving.Add(cell);
            }
            for (int i = 0; i < moving.Count; i++)
            {
                _cells.Remove(MakeKey(moving[i].Row, moving[i].Column));
            }
            for (int i = 0; i < moving.Count; i++)
            {
                Cell cell = moving[i];
                cell.MoveTo(cell.Row, cell.Column + delta);
                _cells.Add(MakeKey(cell.Row, cell.Column), cell);
            }
        }

        /// <summary>删除列区间内的单元格并把右侧单元格左移。Removes the cells in a column range and moves the columns on the right left.</summary>
        /// <param name="columnIndex">起始列号。The first column number.</param>
        /// <param name="endColumn">结束列号。The last column number.</param>
        private void DeleteColumnCells(int columnIndex, int endColumn)
        {
            int count = endColumn - columnIndex + 1;
            List<Cell> moving = new List<Cell>();
            List<long> removed = new List<long>();
            foreach (KeyValuePair<long, Cell> pair in _cells)
            {
                if (pair.Value.Column >= columnIndex && pair.Value.Column <= endColumn)
                {
                    removed.Add(pair.Key);
                }
                else if (pair.Value.Column > endColumn)
                {
                    moving.Add(pair.Value);
                }
            }
            for (int i = 0; i < removed.Count; i++) _cells.Remove(removed[i]);
            for (int i = 0; i < moving.Count; i++)
            {
                Cell cell = moving[i];
                _cells.Remove(MakeKey(cell.Row, cell.Column));
                cell.MoveTo(cell.Row, cell.Column - count);
                _cells.Add(MakeKey(cell.Row, cell.Column), cell);
            }
        }

        /// <summary>插入/删除行后重建行高表：移除 endRow 范围条目并把其后的行号加上 delta。Rebuilds the row-height table after a row edit: entries in the endRow range are dropped and later rows are shifted by delta.</summary>
        /// <param name="rowIndex">编辑起点行号。The first edited row number.</param>
        /// <param name="endRow">被删除的末行；插入时为 0。The last deleted row, or 0 for an insertion.</param>
        /// <param name="delta">行偏移（插入为正、删除为负）。The row delta (positive for insertion, negative for deletion).</param>
        private void RebuildRowHeightEntries(int rowIndex, int endRow, int delta)
        {
            if (RowHeights.Count == 0) return;
            Dictionary<int, double> shifted = new Dictionary<int, double>();
            foreach (KeyValuePair<int, double> pair in RowHeights)
            {
                if (endRow >= rowIndex && pair.Key >= rowIndex && pair.Key <= endRow) continue;
                int key = pair.Key >= rowIndex ? pair.Key + delta : pair.Key;
                shifted[key] = pair.Value;
            }
            RowHeights.Clear();
            foreach (KeyValuePair<int, double> pair in shifted) RowHeights[pair.Key] = pair.Value;
        }

        /// <summary>插入/删除行后重建隐藏行集合。Rebuilds the hidden-row set after a row edit.</summary>
        /// <param name="rowIndex">编辑起点行号。The first edited row number.</param>
        /// <param name="endRow">被删除的末行；插入时为 0。The last deleted row, or 0 for an insertion.</param>
        /// <param name="delta">行偏移。The row delta.</param>
        private void RebuildHiddenRows(int rowIndex, int endRow, int delta)
        {
            if (HiddenRows.Count == 0) return;
            HashSet<int> shifted = new HashSet<int>();
            foreach (int row in HiddenRows)
            {
                if (endRow >= rowIndex && row >= rowIndex && row <= endRow) continue;
                shifted.Add(row >= rowIndex ? row + delta : row);
            }
            HiddenRows.Clear();
            foreach (int row in shifted) HiddenRows.Add(row);
        }

        /// <summary>插入/删除列后重建列宽表。Rebuilds the column-width table after a column edit.</summary>
        /// <param name="columnIndex">编辑起点列号。The first edited column number.</param>
        /// <param name="endColumn">被删除的末列；插入时为 0。The last deleted column, or 0 for an insertion.</param>
        /// <param name="delta">列偏移。The column delta.</param>
        private void RebuildColumnWidthEntries(int columnIndex, int endColumn, int delta)
        {
            if (ColumnWidths.Count == 0) return;
            Dictionary<int, double> shifted = new Dictionary<int, double>();
            foreach (KeyValuePair<int, double> pair in ColumnWidths)
            {
                if (endColumn >= columnIndex && pair.Key >= columnIndex && pair.Key <= endColumn) continue;
                int key = pair.Key >= columnIndex ? pair.Key + delta : pair.Key;
                shifted[key] = pair.Value;
            }
            ColumnWidths.Clear();
            foreach (KeyValuePair<int, double> pair in shifted) ColumnWidths[pair.Key] = pair.Value;
        }

        /// <summary>插入/删除列后重建隐藏列集合。Rebuilds the hidden-column set after a column edit.</summary>
        /// <param name="columnIndex">编辑起点列号。The first edited column number.</param>
        /// <param name="endColumn">被删除的末列；插入时为 0。The last deleted column, or 0 for an insertion.</param>
        /// <param name="delta">列偏移。The column delta.</param>
        private void RebuildHiddenColumns(int columnIndex, int endColumn, int delta)
        {
            if (HiddenColumns.Count == 0) return;
            HashSet<int> shifted = new HashSet<int>();
            foreach (int column in HiddenColumns)
            {
                if (endColumn >= columnIndex && column >= columnIndex && column <= endColumn) continue;
                shifted.Add(column >= columnIndex ? column + delta : column);
            }
            HiddenColumns.Clear();
            foreach (int column in shifted) HiddenColumns.Add(column);
        }

        /// <summary>插入行后扩展或下移合并区域。Expands or shifts merged ranges after rows are inserted.</summary>
        /// <param name="rowIndex">插入位置行号。The insertion row number.</param>
        /// <param name="count">插入行数。The inserted row count.</param>
        private void InsertMergedRows(int rowIndex, int count)
        {
            for (int i = 0; i < MergedRanges.Count; i++)
            {
                MergedRange range = MergedRanges[i];
                if (range.Row >= rowIndex)
                {
                    range.Row += count;
                    MergedRanges[i] = range;
                }
                else if (rowIndex <= range.LastRow)
                {
                    range.RowSpan += count;
                    MergedRanges[i] = range;
                }
            }
        }

        /// <summary>删除行后裁剪、移除或上移合并区域。Clips, removes or shifts merged ranges after rows are deleted.</summary>
        /// <param name="rowIndex">删除起点行号。The first deleted row number.</param>
        /// <param name="endRow">删除终点行号。The last deleted row number.</param>
        /// <param name="count">删除行数。The deleted row count.</param>
        private void DeleteMergedRows(int rowIndex, int endRow, int count)
        {
            List<MergedRange> kept = new List<MergedRange>();
            for (int i = 0; i < MergedRanges.Count; i++)
            {
                MergedRange range = MergedRanges[i];
                if (range.LastRow < rowIndex) { kept.Add(range); continue; }
                if (range.Row > endRow)
                {
                    range.Row -= count;
                    kept.Add(range);
                    continue;
                }
                int start = Math.Min(range.Row, rowIndex);
                int last = range.LastRow > endRow ? range.LastRow - count : rowIndex - 1;
                if (last >= start) kept.Add(new MergedRange(start, range.Column, last - start + 1, range.ColumnSpan));
            }
            MergedRanges.Clear();
            MergedRanges.AddRange(kept);
        }

        /// <summary>插入列后扩展或右移合并区域。Expands or shifts merged ranges after columns are inserted.</summary>
        /// <param name="columnIndex">插入位置列号。The insertion column number.</param>
        /// <param name="count">插入列数。The inserted column count.</param>
        private void InsertMergedColumns(int columnIndex, int count)
        {
            for (int i = 0; i < MergedRanges.Count; i++)
            {
                MergedRange range = MergedRanges[i];
                if (range.Column >= columnIndex)
                {
                    range.Column += count;
                    MergedRanges[i] = range;
                }
                else if (columnIndex <= range.LastColumn)
                {
                    range.ColumnSpan += count;
                    MergedRanges[i] = range;
                }
            }
        }

        /// <summary>删除列后裁剪、移除或左移合并区域。Clips, removes or shifts merged ranges after columns are deleted.</summary>
        /// <param name="columnIndex">删除起点列号。The first deleted column number.</param>
        /// <param name="endColumn">删除终点列号。The last deleted column number.</param>
        /// <param name="count">删除列数。The deleted column count.</param>
        private void DeleteMergedColumns(int columnIndex, int endColumn, int count)
        {
            List<MergedRange> kept = new List<MergedRange>();
            for (int i = 0; i < MergedRanges.Count; i++)
            {
                MergedRange range = MergedRanges[i];
                if (range.LastColumn < columnIndex) { kept.Add(range); continue; }
                if (range.Column > endColumn)
                {
                    range.Column -= count;
                    kept.Add(range);
                    continue;
                }
                int start = Math.Min(range.Column, columnIndex);
                int last = range.LastColumn > endColumn ? range.LastColumn - count : columnIndex - 1;
                if (last >= start) kept.Add(new MergedRange(range.Row, start, range.RowSpan, last - start + 1));
            }
            MergedRanges.Clear();
            MergedRanges.AddRange(kept);
        }

        /// <summary>插入行后下移自动筛选区域。Shifts the auto-filter range after rows are inserted.</summary>
        /// <param name="rowIndex">插入位置行号。The insertion row number.</param>
        /// <param name="count">插入行数。The inserted row count.</param>
        /// <param name="insert">true 表示插入，false 表示删除。True for insertion, false for deletion.</param>
        private void ShiftFilterRows(int rowIndex, int count, bool insert)
        {
            if (AutoFilter == null) return;
            string reference = ShiftReferenceRows(AutoFilter.Range, rowIndex, count, insert);
            if (reference == null) AutoFilter = null;
            else AutoFilter.Range = reference;
        }

        /// <summary>插入/删除列后调整自动筛选区域及其列偏移。Adjusts the auto-filter range and its column offsets after a column edit.</summary>
        /// <param name="columnIndex">编辑起点列号。The first edited column number.</param>
        /// <param name="count">编辑列数。The edited column count.</param>
        /// <param name="insert">true 表示插入，false 表示删除。True for insertion, false for deletion.</param>
        private void ShiftFilterColumns(int columnIndex, int count, bool insert)
        {
            if (AutoFilter == null) return;
            string reference = ShiftReferenceColumns(AutoFilter.Range, columnIndex, count, insert);
            if (reference == null)
            {
                AutoFilter = null;
                return;
            }
            AutoFilter.Range = reference;
            List<ExcelFilterColumn> kept = new List<ExcelFilterColumn>();
            for (int i = 0; i < AutoFilter.Columns.Count; i++)
            {
                ExcelFilterColumn column = AutoFilter.Columns[i];
                int absolute = column.ColumnId + 1;
                if (!insert && absolute >= columnIndex && absolute < columnIndex + count) continue;
                if (absolute >= columnIndex) column.ColumnId = absolute + (insert ? count : -count) - 1;
                kept.Add(column);
            }
            AutoFilter.Columns.Clear();
            AutoFilter.Columns.AddRange(kept);
        }

        /// <summary>插入/删除行后调整数据验证区域；完全被删除的规则直接移除。Adjusts data-validation ranges after a row edit, removing rules that are deleted entirely.</summary>
        /// <param name="rowIndex">编辑起点行号。The first edited row number.</param>
        /// <param name="count">编辑行数。The edited row count.</param>
        /// <param name="insert">true 表示插入，false 表示删除。True for insertion, false for deletion.</param>
        private void ShiftValidationRows(int rowIndex, int count, bool insert)
        {
            for (int i = DataValidations.Count - 1; i >= 0; i--)
            {
                string reference = ShiftReferenceRows(DataValidations[i].Range, rowIndex, count, insert);
                if (reference == null) DataValidations.RemoveAt(i);
                else DataValidations[i].Range = reference;
            }
        }

        /// <summary>插入/删除列后调整数据验证区域；完全被删除的规则直接移除。Adjusts data-validation ranges after a column edit, removing rules that are deleted entirely.</summary>
        /// <param name="columnIndex">编辑起点列号。The first edited column number.</param>
        /// <param name="count">编辑列数。The edited column count.</param>
        /// <param name="insert">true 表示插入，false 表示删除。True for insertion, false for deletion.</param>
        private void ShiftValidationColumns(int columnIndex, int count, bool insert)
        {
            for (int i = DataValidations.Count - 1; i >= 0; i--)
            {
                string reference = ShiftReferenceColumns(DataValidations[i].Range, columnIndex, count, insert);
                if (reference == null) DataValidations.RemoveAt(i);
                else DataValidations[i].Range = reference;
            }
        }

        /// <summary>插入/删除行后调整打印区域。Adjusts the print area after a row edit.</summary>
        /// <param name="rowIndex">编辑起点行号。The first edited row number.</param>
        /// <param name="count">编辑行数。The edited row count.</param>
        private void InsertPrintAreaRows(int rowIndex, int count)
        {
            if (!HasPrintArea) return;
            if (PrintAreaFirstRow >= rowIndex)
            {
                PrintAreaFirstRow += count;
                PrintAreaLastRow += count;
            }
            else if (PrintAreaLastRow >= rowIndex)
            {
                PrintAreaLastRow += count;
            }
        }

        /// <summary>删除行后收缩或清除打印区域。Shrinks or clears the print area after rows are deleted.</summary>
        /// <param name="rowIndex">删除起点行号。The first deleted row number.</param>
        /// <param name="endRow">删除终点行号。The last deleted row number.</param>
        /// <param name="count">删除行数。The deleted row count.</param>
        private void DeletePrintAreaRows(int rowIndex, int endRow, int count)
        {
            if (!HasPrintArea) return;
            if (PrintAreaFirstRow > endRow)
            {
                PrintAreaFirstRow -= count;
                PrintAreaLastRow -= count;
            }
            else if (PrintAreaLastRow >= rowIndex)
            {
                PrintAreaLastRow -= count;
                if (PrintAreaLastRow < PrintAreaFirstRow)
                {
                    PrintAreaFirstRow = 0;
                    PrintAreaLastRow = 0;
                }
            }
        }

        /// <summary>插入列后调整打印区域。Adjusts the print area after a column edit.</summary>
        /// <param name="columnIndex">编辑起点列号。The first edited column number.</param>
        /// <param name="count">编辑列数。The edited column count.</param>
        private void InsertPrintAreaColumns(int columnIndex, int count)
        {
            if (!HasPrintArea) return;
            if (PrintAreaFirstColumn >= columnIndex)
            {
                PrintAreaFirstColumn += count;
                PrintAreaLastColumn += count;
            }
            else if (PrintAreaLastColumn >= columnIndex)
            {
                PrintAreaLastColumn += count;
            }
        }

        /// <summary>删除列后收缩或清除打印区域。Shrinks or clears the print area after columns are deleted.</summary>
        /// <param name="columnIndex">删除起点列号。The first deleted column number.</param>
        /// <param name="endColumn">删除终点列号。The last deleted column number.</param>
        /// <param name="count">删除列数。The deleted column count.</param>
        private void DeletePrintAreaColumns(int columnIndex, int endColumn, int count)
        {
            if (!HasPrintArea) return;
            if (PrintAreaFirstColumn > endColumn)
            {
                PrintAreaFirstColumn -= count;
                PrintAreaLastColumn -= count;
            }
            else if (PrintAreaLastColumn >= columnIndex)
            {
                PrintAreaLastColumn -= count;
                if (PrintAreaLastColumn < PrintAreaFirstColumn)
                {
                    PrintAreaFirstColumn = 0;
                    PrintAreaLastColumn = 0;
                }
            }
        }

        /// <summary>按插入/删除语义平移区域引用的行号；区域被完全删除时返回 null。Shifts the row numbers of a range reference for an insertion or deletion, returning null when the range disappears.</summary>
        /// <param name="reference">A1 风格区域引用。The A1-style range reference.</param>
        /// <param name="rowIndex">编辑起点行号。The first edited row number.</param>
        /// <param name="count">编辑行数。The edited row count.</param>
        /// <param name="insert">true 表示插入。True for insertion.</param>
        /// <returns>平移后的引用或 null。The shifted reference, or null.</returns>
        private static string ShiftReferenceRows(string reference, int rowIndex, int count, bool insert)
        {
            int row1, column1, row2, column2;
            try
            {
                CellAddress.ParseRange(reference, out row1, out column1, out row2, out column2);
            }
            catch (FormatException)
            {
                return reference;
            }
            if (row1 > row2) { int swap = row1; row1 = row2; row2 = swap; }
            if (column1 > column2) { int swap = column1; column1 = column2; column2 = swap; }
            if (insert)
            {
                if (row1 >= rowIndex)
                {
                    row1 += count;
                    row2 += count;
                }
                else if (row2 >= rowIndex)
                {
                    row2 += count;
                }
            }
            else
            {
                int endRow = rowIndex + count - 1;
                if (row1 > endRow)
                {
                    row1 -= count;
                    row2 -= count;
                }
                else if (row2 >= rowIndex)
                {
                    row2 -= count;
                    if (row2 < row1) return null;
                }
            }
            return CellAddress.Format(row1, column1) + ":" + CellAddress.Format(row2, column2);
        }

        /// <summary>按插入/删除语义平移区域引用的列号；区域被完全删除时返回 null。Shifts the column numbers of a range reference for an insertion or deletion, returning null when the range disappears.</summary>
        /// <param name="reference">A1 风格区域引用。The A1-style range reference.</param>
        /// <param name="columnIndex">编辑起点列号。The first edited column number.</param>
        /// <param name="count">编辑列数。The edited column count.</param>
        /// <param name="insert">true 表示插入。True for insertion.</param>
        /// <returns>平移后的引用或 null。The shifted reference, or null.</returns>
        private static string ShiftReferenceColumns(string reference, int columnIndex, int count, bool insert)
        {
            int row1, column1, row2, column2;
            try
            {
                CellAddress.ParseRange(reference, out row1, out column1, out row2, out column2);
            }
            catch (FormatException)
            {
                return reference;
            }
            if (row1 > row2) { int swap = row1; row1 = row2; row2 = swap; }
            if (column1 > column2) { int swap = column1; column1 = column2; column2 = swap; }
            if (insert)
            {
                if (column1 >= columnIndex)
                {
                    column1 += count;
                    column2 += count;
                }
                else if (column2 >= columnIndex)
                {
                    column2 += count;
                }
            }
            else
            {
                int endColumn = columnIndex + count - 1;
                if (column1 > endColumn)
                {
                    column1 -= count;
                    column2 -= count;
                }
                else if (column2 >= columnIndex)
                {
                    column2 -= count;
                    if (column2 < column1) return null;
                }
            }
            return CellAddress.Format(row1, column1) + ":" + CellAddress.Format(row2, column2);
        }
    }

    /// <summary>工作簿模型：包含多个工作表、文档元数据与全局默认单元格样式。A workbook model: it holds multiple worksheets, document metadata and a global default cell style.</summary>
    public sealed class OfficeWorkbook
    {
        /// <summary>工作表列表。The list of worksheets.</summary>
        public List<Worksheet> Worksheets { get; private set; }
        /// <summary>工作簿元数据。The workbook metadata.</summary>
        public DocumentMetadata Metadata { get; set; }
        /// <summary>工作簿级默认单元格样式。The workbook-level default cell style.</summary>
        public CellStyle DefaultStyle { get; set; }
        /// <summary>工作簿定义名集合（全局或绑定到工作表）。The workbook defined names (global or scoped to a worksheet).</summary>
        public List<ExcelDefinedName> DefinedNames { get; private set; }

        /// <summary>默认构造：空工作表列表、默认元数据与样式。Default constructor: an empty worksheet list with default metadata and style.</summary>
        public OfficeWorkbook()
        {
            Worksheets = new List<Worksheet>();
            Metadata = new DocumentMetadata();
            DefaultStyle = new CellStyle();
            DefinedNames = new List<ExcelDefinedName>();
        }

        /// <summary>添加工作表；与现有表重名（忽略大小写）时自动追加数字后缀保证唯一。Adds a worksheet, appending a numeric suffix to keep the name unique when it collides (case-insensitively) with an existing sheet.</summary>
        /// <param name="name">期望的工作表名称。The desired worksheet name.</param>
        /// <returns>新建的工作表。The newly created worksheet.</returns>
        public Worksheet AddWorksheet(string name)
        {
            string unique = name;
            int suffix = 2;
            while (FindWorksheet(unique) != null)
            {
                unique = name + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            }
            Worksheet sheet = new Worksheet(unique);
            Worksheets.Add(sheet);
            return sheet;
        }

        /// <summary>按名称忽略大小写查找工作表。Finds a worksheet by name, ignoring case.</summary>
        /// <param name="name">工作表名称。The worksheet name.</param>
        /// <returns>找到的工作表；不存在返回 null。The found worksheet, or null when it does not exist.</returns>
        public Worksheet FindWorksheet(string name)
        {
            for (int i = 0; i < Worksheets.Count; i++)
            {
                if (string.Equals(Worksheets[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return Worksheets[i];
                }
            }
            return null;
        }

        /// <summary>第一个工作表；尚不存在时自动创建 "Sheet1"。The first worksheet; "Sheet1" is created automatically when none exists.</summary>
        public Worksheet FirstWorksheet
        {
            get
            {
                if (Worksheets.Count == 0)
                {
                    AddWorksheet("Sheet1");
                }
                return Worksheets[0];
            }
        }

        /// <summary>用内置公式引擎重算所有公式单元格，把结果写回 <see cref="Cell.Value"/>（保留 <see cref="Cell.Formula"/>）。Recalculates every formula cell with the built-in engine and writes results into <see cref="Cell.Value"/> (keeping <see cref="Cell.Formula"/>).</summary>
        /// <returns>被计算的公式单元格数量。The number of formula cells computed.</returns>
        /// <remarks>默认不在读写路径自动触发，以保持既有输出字节稳定；需要缓存值时由调用方显式调用。
        /// It is not invoked automatically on the read/write path so existing output stays byte-stable; call it explicitly when cached values are wanted.</remarks>
        public int CalculateFormulas()
        {
            return Wisdom.Office.Core.Formulas.FormulaEngine.Calculate(this);
        }

        /// <summary>判断表名是否符合 Excel 规则：非空、不超过 31 字符、不含 [ ] : * ? / \、首尾无单引号。Checks whether a sheet name follows Excel rules: non-empty, at most 31 characters, without [ ] : * ? / \ and without leading or trailing apostrophes.</summary>
        /// <param name="name">待检查的表名。The sheet name to check.</param>
        /// <returns>合法返回 true。True when the name is valid.</returns>
        public static bool IsValidWorksheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            if (name.Length > 31) return false;
            if (name[0] == '\'' || name[name.Length - 1] == '\'') return false;
            for (int i = 0; i < name.Length; i++)
            {
                char c = name[i];
                if (c == '[' || c == ']' || c == ':' || c == '*' || c == '?' || c == '/' || c == '\\' || c < 32)
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>重命名工作表，校验名称合法且不与其他表重名（忽略大小写）。Renames a worksheet after validating that the name is legal and does not collide with another sheet (ignoring case).</summary>
        /// <param name="sheet">要重命名的工作表。The worksheet to rename.</param>
        /// <param name="newName">新名称。The new name.</param>
        /// <exception cref="ArgumentNullException">sheet 为 null。sheet is null.</exception>
        /// <exception cref="ArgumentException">名称非法或已被占用。The name is illegal or already in use.</exception>
        public void RenameWorksheet(Worksheet sheet, string newName)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            if (!IsValidWorksheetName(newName))
            {
                throw new ArgumentException("Invalid worksheet name.", "newName");
            }
            Worksheet existing = FindWorksheet(newName);
            if (existing != null && existing != sheet)
            {
                throw new ArgumentException("Worksheet name is already in use.", "newName");
            }
            sheet.Name = newName;
        }

        /// <summary>移除指定工作表；不存在时返回 false。Removes a worksheet, returning false when it is not part of the workbook.</summary>
        /// <param name="sheet">要移除的工作表。The worksheet to remove.</param>
        /// <returns>已移除返回 true。True when it was removed.</returns>
        public bool RemoveWorksheet(Worksheet sheet)
        {
            if (sheet == null) return false;
            return Worksheets.Remove(sheet);
        }

        /// <summary>按名称移除工作表；不存在时返回 false。Removes a worksheet by name, returning false when it does not exist.</summary>
        /// <param name="name">工作表名称。The worksheet name.</param>
        /// <returns>已移除返回 true。True when it was removed.</returns>
        public bool RemoveWorksheet(string name)
        {
            Worksheet sheet = FindWorksheet(name);
            return sheet != null && Worksheets.Remove(sheet);
        }

        /// <summary>把工作表移动到新的位置，越界的目标下标自动钳制到有效范围。Moves a worksheet to a new position, clamping an out-of-range target index.</summary>
        /// <param name="oldIndex">当前下标（从 0 开始）。The current index (0-based).</param>
        /// <param name="newIndex">目标下标（从 0 开始）。The target index (0-based).</param>
        /// <exception cref="ArgumentOutOfRangeException">当前下标越界。The current index is out of range.</exception>
        public void MoveWorksheet(int oldIndex, int newIndex)
        {
            if (oldIndex < 0 || oldIndex >= Worksheets.Count) throw new ArgumentOutOfRangeException("oldIndex");
            if (newIndex < 0) newIndex = 0;
            if (newIndex >= Worksheets.Count) newIndex = Worksheets.Count - 1;
            if (newIndex == oldIndex) return;
            Worksheet sheet = Worksheets[oldIndex];
            Worksheets.RemoveAt(oldIndex);
            Worksheets.Insert(newIndex, sheet);
        }

        /// <summary>复制工作表（单元格、样式、合并、行高列宽、筛选、验证与打印设置），副本追加到末尾；源表不存在时返回 null。Copies a worksheet (cells, styles, merges, sizing, filters, validations and print settings) and appends the copy; returns null when the source is missing.</summary>
        /// <param name="name">源工作表名称。The source worksheet name.</param>
        /// <param name="newName">副本名称；为空时使用"原名 副本"。The copy name; defaults to "original copy" when empty.</param>
        /// <returns>复制出的工作表或 null。The copied worksheet, or null.</returns>
        public Worksheet CopyWorksheet(string name, string newName = null)
        {
            Worksheet source = FindWorksheet(name);
            if (source == null) return null;
            Worksheet copy = AddWorksheet(string.IsNullOrEmpty(newName) ? source.Name + " 副本" : newName);
            copy.IsHidden = source.IsHidden;
            copy.ShowGridLines = source.ShowGridLines;
            copy.ShowZeros = source.ShowZeros;
            copy.FrozenRows = source.FrozenRows;
            copy.FrozenColumns = source.FrozenColumns;
            copy.FitToPage = source.FitToPage;
            copy.CenterHorizontally = source.CenterHorizontally;
            copy.PrintAreaFirstRow = source.PrintAreaFirstRow;
            copy.PrintAreaLastRow = source.PrintAreaLastRow;
            copy.PrintAreaFirstColumn = source.PrintAreaFirstColumn;
            copy.PrintAreaLastColumn = source.PrintAreaLastColumn;
            copy.DefaultColumnWidth = source.DefaultColumnWidth;
            copy.DefaultRowHeight = source.DefaultRowHeight;
            copy.DefaultStyle = source.DefaultStyle == null ? new CellStyle() : source.DefaultStyle.Clone();
            copy.TabColor = source.TabColor;
            PageSetup setup = source.PrintSetup ?? new PageSetup();
            copy.PrintSetup = new PageSetup
            {
                Width = setup.Width,
                Height = setup.Height,
                MarginLeft = setup.MarginLeft,
                MarginTop = setup.MarginTop,
                MarginRight = setup.MarginRight,
                MarginBottom = setup.MarginBottom,
                HeaderDistance = setup.HeaderDistance,
                FooterDistance = setup.FooterDistance
            };
            foreach (KeyValuePair<int, double> pair in source.ColumnWidths) copy.ColumnWidths[pair.Key] = pair.Value;
            foreach (KeyValuePair<int, double> pair in source.RowHeights) copy.RowHeights[pair.Key] = pair.Value;
            foreach (int row in source.HiddenRows) copy.HiddenRows.Add(row);
            foreach (int column in source.HiddenColumns) copy.HiddenColumns.Add(column);
            copy.MergedRanges.AddRange(source.MergedRanges);
            foreach (Cell cell in source.Cells)
            {
                Cell target = copy.GetCell(cell.Row, cell.Column);
                target.Value = cell.Value;
                target.Formula = cell.Formula;
                target.Hyperlink = cell.Hyperlink;
                target.HyperlinkTooltip = cell.HyperlinkTooltip;
                if (cell.Style != null) target.Style = cell.Style.Clone();
            }
            if (source.AutoFilter != null)
            {
                ExcelAutoFilter filter = new ExcelAutoFilter(source.AutoFilter.Range);
                for (int i = 0; i < source.AutoFilter.Columns.Count; i++)
                {
                    ExcelFilterColumn column = source.AutoFilter.Columns[i];
                    ExcelFilterColumn copyColumn = filter.ForColumn(column.ColumnId);
                    copyColumn.IncludeBlanks = column.IncludeBlanks;
                    copyColumn.Values.AddRange(column.Values);
                }
                copy.AutoFilter = filter;
            }
            for (int i = 0; i < source.DataValidations.Count; i++)
            {
                ExcelDataValidation validation = source.DataValidations[i];
                copy.DataValidations.Add(new ExcelDataValidation
                {
                    Range = validation.Range,
                    Type = validation.Type,
                    Operator = validation.Operator,
                    Formula1 = validation.Formula1,
                    Formula2 = validation.Formula2,
                    AllowBlank = validation.AllowBlank,
                    ShowDropDown = validation.ShowDropDown,
                    ShowInputMessage = validation.ShowInputMessage,
                    ShowErrorMessage = validation.ShowErrorMessage,
                    PromptTitle = validation.PromptTitle,
                    Prompt = validation.Prompt,
                    ErrorTitle = validation.ErrorTitle,
                    ErrorMessage = validation.ErrorMessage,
                    ErrorStyle = validation.ErrorStyle
                });
            }
            copy.Charts.AddRange(source.Charts);
            copy.Images.AddRange(source.Images);
            copy.ConditionalFormats.AddRange(source.ConditionalFormats);
            return copy;
        }
    }

    /// <summary>A1 风格单元格引用的解析与生成工具。Utilities for parsing and building A1-style cell references.</summary>
    public static class CellAddress
    {
        /// <summary>Excel 网格的最大列数（XFD）。The last column of the Excel grid (XFD).</summary>
        public const int MaxColumnLiteral = 16384;

        /// <summary>Excel 网格的最大行数。The last row of the Excel grid.</summary>
        public const int MaxRowLiteral = 1048576;

        /// <summary>解析单元格引用为 1 基行列坐标，支持 "$" 绝对引用与 "Sheet!B3" 前缀。Parses a cell reference into 1-based row/column coordinates, supporting "$" absolute references and a "Sheet!B3" prefix.</summary>
        /// <param name="reference">引用字符串，如 "B3"。The reference string, such as "B3".</param>
        /// <param name="row">输出行号（从 1 开始）。Outputs the row number (starting from 1).</param>
        /// <param name="column">输出列号（从 1 开始）。Outputs the column number (starting from 1).</param>
        /// <exception cref="FormatException">引用为空或格式非法。The reference is empty or malformed.</exception>
        public static void Parse(string reference, out int row, out int column)
        {
            if (string.IsNullOrEmpty(reference))
            {
                throw new FormatException("Cell reference is empty.");
            }

            string value = reference.Trim();
            int bang = value.LastIndexOf('!');
            if (bang >= 0)
            {
                value = value.Substring(bang + 1);
            }
            value = value.Replace("$", string.Empty);

            int index = 0;
            long col = 0;
            bool hasColumn = false;
            while (index < value.Length && char.IsLetter(value[index]))
            {
                char c = char.ToUpperInvariant(value[index]);
                col = col * 26 + (c - 'A' + 1);
                // 列字母串过长时数值会溢出 int：超过 Excel 的最大列数即视为非法引用，
                // 让调用方走回退分支，而不是静默绕回成另一个坐标。
                // A very long run of letters overflows int: anything past Excel's last column is rejected so
                // callers take their fallback path instead of silently wrapping onto another coordinate.
                if (col > MaxColumnLiteral) throw new FormatException("Invalid cell reference: " + reference);
                hasColumn = true;
                index++;
            }

            long rowNumber = 0;
            bool hasRow = false;
            while (index < value.Length && char.IsDigit(value[index]))
            {
                rowNumber = rowNumber * 10 + (value[index] - '0');
                if (rowNumber > MaxRowLiteral) throw new FormatException("Invalid cell reference: " + reference);
                hasRow = true;
                index++;
            }

            if (!hasColumn || !hasRow || col <= 0 || rowNumber <= 0)
            {
                throw new FormatException("Invalid cell reference: " + reference);
            }

            row = (int)rowNumber;
            column = (int)col;
        }

        /// <summary>解析范围引用（如 "A1:B5"）；单个引用时末坐标等于首坐标。Parses a range reference (such as "A1:B5"); for a single reference the end coordinates equal the start coordinates.</summary>
        /// <param name="reference">范围引用字符串。The range reference string.</param>
        /// <param name="row1">输出起始行号。Outputs the starting row number.</param>
        /// <param name="column1">输出起始列号。Outputs the starting column number.</param>
        /// <param name="row2">输出结束行号。Outputs the ending row number.</param>
        /// <param name="column2">输出结束列号。Outputs the ending column number.</param>
        /// <exception cref="FormatException">引用为空或格式非法。The reference is empty or malformed.</exception>
        public static void ParseRange(string reference, out int row1, out int column1, out int row2, out int column2)
        {
            if (string.IsNullOrEmpty(reference))
            {
                throw new FormatException("Range reference is empty.");
            }

            string[] parts = reference.Trim().Split(':');
            if (parts.Length == 1)
            {
                Parse(parts[0], out row1, out column1);
                row2 = row1;
                column2 = column1;
                return;
            }
            if (parts.Length != 2)
            {
                throw new FormatException("Invalid range reference: " + reference);
            }

            Parse(parts[0], out row1, out column1);
            Parse(parts[1], out row2, out column2);
        }

        /// <summary>由 1 基行列坐标生成引用字符串，如 (3, 2) → "B3"。Builds a reference string from 1-based coordinates, such as (3, 2) becoming "B3".</summary>
        /// <param name="row">行号。The row number.</param>
        /// <param name="column">列号。The column number.</param>
        /// <returns>A1 风格引用。An A1-style reference.</returns>
        public static string Format(int row, int column)
        {
            return ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>将 1 基列号转换为字母列名（1→A、26→Z、27→AA）。Converts a 1-based column number into its letter name (1 to A, 26 to Z, 27 to AA).</summary>
        /// <param name="column">列号，必须大于等于 1。The column number, must be at least 1.</param>
        /// <returns>字母列名。The letter column name.</returns>
        /// <exception cref="ArgumentOutOfRangeException">列号小于 1。The column number is less than 1.</exception>
        public static string ColumnName(int column)
        {
            if (column < 1)
            {
                throw new ArgumentOutOfRangeException("column");
            }
            StringBuilder builder = new StringBuilder();
            while (column > 0)
            {
                column--;
                builder.Insert(0, (char)('A' + (column % 26)));
                column /= 26;
            }
            return builder.ToString();
        }

        /// <summary>将字母列名转换为 1 基列号（A→1、Z→26、AA→27），忽略大小写。Converts a letter column name into a 1-based column number (A to 1, Z to 26, AA to 27), ignoring case.</summary>
        /// <param name="name">字母列名。The letter column name.</param>
        /// <returns>1 基列号。The 1-based column number.</returns>
        /// <exception cref="FormatException">名称为空或含非字母字符。The name is empty or contains non-letter characters.</exception>
        public static int ColumnIndex(string name)
        {
            int column = 0;
            if (string.IsNullOrEmpty(name))
            {
                throw new FormatException("Column name is empty.");
            }
            for (int i = 0; i < name.Length; i++)
            {
                char c = char.ToUpperInvariant(name[i]);
                if (c < 'A' || c > 'Z')
                {
                    throw new FormatException("Invalid column name: " + name);
                }
                column = column * 26 + (c - 'A' + 1);
            }
            return column;
        }
    }
}
