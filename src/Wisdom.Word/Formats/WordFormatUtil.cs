using System;
using System.Globalization;
using System.Text;

namespace Wisdom.Word
{
    /// <summary>Word 各格式读写器共享的小工具：数值格式化、单位换算、转义与链接安全校验。Small helpers shared by the Word format readers and writers: number formatting, unit conversion, escaping and hyperlink safety checks.</summary>
    internal static class WordFormatUtil
    {
        /// <summary>按不变文化输出数字，最多保留三位小数，去掉多余尾零。Formats a number with the invariant culture, keeping at most three decimals and trimming trailing zeros.</summary>
        /// <param name="value">数值。The value.</param>
        /// <returns>适合写入标记语言的数字文本。The numeric text suitable for markup.</returns>
        public static string Number(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return "0";
            return Math.Round(value, 3).ToString("0.###", CultureInfo.InvariantCulture);
        }

        /// <summary>将磅换算为缇（1 磅 = 20 缇），并限制在非负整数范围内。Converts points to twips (1 point = 20 twips), clamped to a non-negative integer.</summary>
        /// <param name="points">磅值。The value in points.</param>
        /// <returns>缇值。The value in twips.</returns>
        public static int PointsToTwips(double points)
        {
            double twips = Math.Round(points * 20.0);
            if (twips < 0) return 0;
            if (twips > int.MaxValue) return int.MaxValue;
            return (int)twips;
        }

        /// <summary>将 CSS/ODT 长度文本换算为磅；支持 pt、px、cm、mm、in、pc 与纯数字（按磅），无法识别时返回回退值。Converts a CSS/ODT length into points; pt, px, cm, mm, in, pc and bare numbers (points) are supported, and the fallback is returned when the text is not recognized.</summary>
        /// <param name="text">长度文本，如 "12pt"、"16px"。The length text, such as "12pt" or "16px".</param>
        /// <param name="fallback">解析失败时的回退值。The fallback when parsing fails.</param>
        /// <returns>磅值。The value in points.</returns>
        public static double ParseLengthPoints(string text, double fallback)
        {
            if (string.IsNullOrEmpty(text)) return fallback;
            string value = text.Trim();
            double number;
            int unitStart = 0;
            while (unitStart < value.Length
                && (char.IsDigit(value[unitStart]) || value[unitStart] == '.' || value[unitStart] == '-' || value[unitStart] == '+'))
            {
                unitStart++;
            }
            if (unitStart == 0
                || !double.TryParse(value.Substring(0, unitStart), NumberStyles.Float, CultureInfo.InvariantCulture, out number))
            {
                return fallback;
            }
            string unit = value.Substring(unitStart).Trim().ToLowerInvariant();
            switch (unit)
            {
                case "":
                case "pt": return number;
                case "px": return number * 72.0 / 96.0;
                case "pc": return number * 12.0;
                case "in": return number * 72.0;
                case "cm": return number * 72.0 / 2.54;
                case "mm": return number * 72.0 / 25.4;
                case "em": return number * 11.0;
                case "rem": return number * 11.0;
                default: return fallback;
            }
        }

        /// <summary>按 HTML 规则转义文本与属性值（&amp;、&lt;、&gt;、&quot;、&apos;）。Escapes text and attribute values with the HTML rules (&amp;, &lt;, &gt;, &quot;, &apos;).</summary>
        /// <param name="text">原始文本。The raw text.</param>
        /// <returns>转义后的文本。The escaped text.</returns>
        public static string EscapeHtml(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            StringBuilder builder = null;
            for (int i = 0; i < text.Length; i++)
            {
                char value = text[i];
                string replacement = null;
                switch (value)
                {
                    case '&': replacement = "&amp;"; break;
                    case '<': replacement = "&lt;"; break;
                    case '>': replacement = "&gt;"; break;
                    case '"': replacement = "&quot;"; break;
                    case '\'': replacement = "&apos;"; break;
                }
                if (replacement == null)
                {
                    if (builder != null) builder.Append(value);
                    continue;
                }
                if (builder == null)
                {
                    builder = new StringBuilder(text.Length + 16);
                    builder.Append(text, 0, i);
                }
                builder.Append(replacement);
            }
            return builder == null ? text : builder.ToString();
        }

        /// <summary>按 XML 规则转义文本与属性值（&amp;、&lt;、&gt;、&quot;），并丢弃非法控制字符。Escapes text and attribute values with XML rules (&amp;, &lt;, &gt;, &quot;) and drops illegal control characters.</summary>
        /// <param name="text">原始文本。The raw text.</param>
        /// <returns>转义后的文本。The escaped text.</returns>
        public static string EscapeXml(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            StringBuilder builder = new StringBuilder(text.Length + 16);
            for (int i = 0; i < text.Length; i++)
            {
                char value = text[i];
                switch (value)
                {
                    case '&': builder.Append("&amp;"); break;
                    case '<': builder.Append("&lt;"); break;
                    case '>': builder.Append("&gt;"); break;
                    case '"': builder.Append("&quot;"); break;
                    case '\r': break;
                    case '\n':
                    case '\t': builder.Append(value); break;
                    default:
                        if (value < 32 || (value >= 0x7F && value < 0xA0) || value == 0xFFFE || value == 0xFFFF)
                        {
                            break;
                        }
                        builder.Append(value);
                        break;
                }
            }
            return builder.ToString();
        }

        /// <summary>判断字符是否为需要东亚字体渲染的字符（CJK 统一表意文字、全角符号等）。Checks whether a character needs an East-Asian font (CJK unified ideographs, full-width punctuation and the like).</summary>
        /// <param name="value">待判断的字符。The character to inspect.</param>
        /// <returns>属于东亚文字范围时返回 true。True when the character is in an East-Asian range.</returns>
        public static bool IsEastAsian(char value)
        {
            return (value >= 0x2E80 && value <= 0x9FFF)
                || (value >= 0xAC00 && value <= 0xD7AF)
                || (value >= 0xF900 && value <= 0xFAFF)
                || (value >= 0xFE30 && value <= 0xFE4F)
                || (value >= 0xFF00 && value <= 0xFFEF);
        }

        /// <summary>判断字符串是否包含东亚字符。Checks whether a string contains any East-Asian character.</summary>
        /// <param name="text">待检查的文本。The text to inspect.</param>
        /// <returns>包含时返回 true。True when at least one East-Asian character is present.</returns>
        public static bool ContainsEastAsian(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            for (int i = 0; i < text.Length; i++)
            {
                if (IsEastAsian(text[i])) return true;
            }
            return false;
        }

        /// <summary>
        /// 校验外部超链接：拒绝超长目标、控制字符、javascript/data/vbscript 等脚本协议，以及指向本机磁盘或
        /// 网络共享的 file:/smb:/gopher:/ldap: 等协议；无协议的相对路径与常见安全协议（http/https/mailto/ftp/ftps/tel/sms）放行。
        /// Validates an external hyperlink: overlong targets, control characters, script schemes such as
        /// javascript, data and vbscript, and schemes that reach a local disk or a network share such as file:,
        /// smb:, gopher: or ldap: are rejected; relative paths without a scheme and common safe schemes
        /// (http/https/mailto/ftp/ftps/tel/sms) are accepted.
        /// </summary>
        /// <param name="target">外部链接目标。The external target.</param>
        /// <returns>可写出时返回 true。True when it can be written.</returns>
        /// <remarks>清单与 <c>XlsxHyperlinkParts.IsSafeTarget</c> 逐字一致，两个模块对同一份文档给出同样的链接取舍。
        /// The list is character-for-character the one in <c>XlsxHyperlinkParts.IsSafeTarget</c>, so both modules make
        /// the same call about the same document.</remarks>
        public static bool IsSafeHyperlinkTarget(string target)
        {
            if (string.IsNullOrEmpty(target) || target.Length > 2079) return false;
            for (int i = 0; i < target.Length; i++)
            {
                if (target[i] < 32) return false;
            }
            int colon = target.IndexOf(':');
            int slash = target.IndexOfAny(new char[] { '/', '\\', '?', '#' });
            if (colon < 0 || (slash >= 0 && colon > slash)) return true;
            if (colon == 0) return false;
            string scheme = target.Substring(0, colon).ToLowerInvariant();
            return scheme == "http" || scheme == "https" || scheme == "mailto" || scheme == "ftp"
                || scheme == "ftps" || scheme == "tel" || scheme == "sms";
        }

        /// <summary>把普通文本拆分为行，兼容 \r\n、\r 与 \n。Splits plain text into lines, accepting \r\n, \r and \n.</summary>
        /// <param name="text">原始文本，null 视为空字符串。The raw text; null is treated as empty.</param>
        /// <returns>行数组。The array of lines.</returns>
        public static string[] SplitLines(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        }
    }
}
