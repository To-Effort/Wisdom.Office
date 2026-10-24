using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// XML 文本防护。XML 1.0 不接受 0x00~0x08、0x0B、0x0C、0x0E~0x1F、未配对的代理项以及 0xFFFE/0xFFFF，
    /// System.Xml 在写出这些字符时直接抛 ArgumentException；一格来自数据库的脏字符就能让整份文档保存失败。
    /// 这里把非法字符丢掉，使脏数据最多损失一个字符，而不是让整个写出动作崩溃。
    /// Guard for XML text. XML 1.0 rejects 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F, unpaired surrogates and
    /// 0xFFFE/0xFFFF, and System.Xml throws ArgumentException while writing them, so a single dirty cell
    /// coming from a database can fail the whole save. The illegal characters are dropped here, so dirty
    /// data costs at most one character instead of crashing the writer.
    /// </summary>
    public static class OfficeXmlText
    {
        /// <summary>去掉文本里 XML 1.0 不接受的字符；文本干净时返回原字符串实例，不产生额外分配。Removes the characters XML 1.0 rejects; returns the same instance when the text is clean, so clean input allocates nothing.</summary>
        /// <param name="value">原始文本。The original text.</param>
        /// <returns>可安全写进 XML 的文本；入参为 null 时返回 null。Text that is safe to write into XML; null when the input is null.</returns>
        public static string Strip(string value)
        {
            if (string.IsNullOrEmpty(value)) return value;
            int length = value.Length;
            int first = -1;
            for (int i = 0; i < length; i++)
            {
                if (IsIllegalAt(value, i))
                {
                    first = i;
                    break;
                }
            }
            if (first < 0) return value;

            StringBuilder builder = new StringBuilder(length);
            builder.Append(value, 0, first);
            for (int i = first; i < length; i++)
            {
                char current = value[i];
                if (char.IsHighSurrogate(current))
                {
                    // 只有配上后续低代理项的高代理项才是合法字符，落单的必须丢弃。
                    // A high surrogate is legal only when a low surrogate follows it; a lone one must go.
                    if (i + 1 < length && char.IsLowSurrogate(value[i + 1]))
                    {
                        builder.Append(current);
                        builder.Append(value[i + 1]);
                        i++;
                    }
                    continue;
                }
                if (char.IsLowSurrogate(current)) continue;
                if (IsForbidden(current)) continue;
                builder.Append(current);
            }
            return builder.ToString();
        }

        /// <summary>就地把文档里所有文本节点与属性值中非法的 XML 字符清掉。文档本身不会被替换节点结构。Clears illegal XML characters from every text node and attribute value in place without restructuring the document.</summary>
        /// <param name="document">待清理的文档。The document to clean.</param>
        public static void Sanitize(XDocument document)
        {
            XElement root = document.Root;
            if (root == null) return;
            List<XText> dirty = null;
            List<string> cleaned = null;
            foreach (XNode node in root.DescendantNodesAndSelf())
            {
                XElement element = node as XElement;
                if (element != null)
                {
                    foreach (XAttribute attribute in element.Attributes())
                    {
                        string value = attribute.Value;
                        string stripped = Strip(value);
                        if (!ReferenceEquals(stripped, value)) attribute.Value = stripped;
                    }
                    continue;
                }
                XText text = node as XText;
                if (text == null) continue;
                string content = text.Value;
                string textStripped = Strip(content);
                if (ReferenceEquals(textStripped, content)) continue;
                // 文本节点的 Value 赋值可能移动节点，遍历结束后再统一改写。
                // Assigning XText.Value can move the node, so the rewrites happen after the walk.
                if (dirty == null)
                {
                    dirty = new List<XText>();
                    cleaned = new List<string>();
                }
                dirty.Add(text);
                cleaned.Add(textStripped);
            }
            if (dirty == null) return;
            for (int i = 0; i < dirty.Count; i++) dirty[i].Value = cleaned[i];
        }

        /// <summary>判断指定位置的字符是否无法写进 XML。Reports whether the character at the index cannot be written into XML.</summary>
        /// <param name="value">原始文本。The original text.</param>
        /// <param name="index">字符位置。The character index.</param>
        /// <returns>非法时返回 true。True when illegal.</returns>
        private static bool IsIllegalAt(string value, int index)
        {
            char current = value[index];
            if (char.IsHighSurrogate(current))
            {
                return index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1]);
            }
            // 紧跟在高代理项之后的低代理项属于同一对，不能当作落单代理项丢弃。
            // A low surrogate right behind a high one belongs to that pair and is not a lone surrogate.
            if (char.IsLowSurrogate(current)) return index == 0 || !char.IsHighSurrogate(value[index - 1]);
            return IsForbidden(current);
        }

        /// <summary>XML 1.0 的合法字符集为 #x9、#xA、#xD、[#x20-#xD7FF]、[#xE000-#xFFFD]。The legal XML 1.0 characters are #x9, #xA, #xD, [#x20-#xD7FF] and [#xE000-#xFFFD].</summary>
        /// <param name="value">待判定的基本字符。The basic character to test.</param>
        /// <returns>被规范排除时返回 true。True when the specification excludes it.</returns>
        private static bool IsForbidden(char value)
        {
            if (value < ' ') return value != '\t' && value != '\n' && value != '\r';
            return value == '\uFFFE' || value == '\uFFFF';
        }
    }
}
