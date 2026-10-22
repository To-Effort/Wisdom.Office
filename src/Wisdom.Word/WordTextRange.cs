using System;
using System.Collections.Generic;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 段落内连续文本运行的合并与跨运行替换工具：把相邻 <see cref="TextRun"/> 拼成一段虚拟文本，
    /// 使查找替换与邮件合并可以命中被 Word 拆分到多个运行中的标记。
    /// Utility for merging and editing consecutive text runs: adjacent <see cref="TextRun"/> instances are
    /// treated as one virtual string so find/replace and mail merge can hit markers that Word split across runs.
    /// </summary>
    internal static class WordTextRange
    {
        /// <summary>
        /// 把段落切分为若干连续文本组：图片、域等非文本行内元素会切断分组。
        /// Splits a paragraph into groups of consecutive text runs; non-text inlines such as images and fields
        /// terminate a group.
        /// </summary>
        /// <param name="paragraph">待切分的段落。The paragraph to split.</param>
        /// <returns>按出现顺序排列的文本运行组。The text-run groups in document order.</returns>
        public static List<List<TextRun>> GetGroups(Paragraph paragraph)
        {
            List<List<TextRun>> groups = new List<List<TextRun>>();
            List<TextRun> current = null;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run == null)
                {
                    current = null;
                    continue;
                }
                if (current == null)
                {
                    current = new List<TextRun>();
                    groups.Add(current);
                }
                current.Add(run);
            }
            return groups;
        }

        /// <summary>
        /// 把一组文本运行拼接为虚拟文本。
        /// Concatenates a group of text runs into one virtual string.
        /// </summary>
        /// <param name="runs">文本运行组。The text-run group.</param>
        /// <returns>拼接后的文本。The concatenated text.</returns>
        public static string GetText(List<TextRun> runs)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < runs.Count; i++)
            {
                builder.Append(runs[i].Text);
            }
            return builder.ToString();
        }

        /// <summary>
        /// 以新文本替换组内 [start, start+length) 区间：替换文字继承首个命中运行的格式，
        /// 其余被覆盖运行的文字被清空。
        /// Replaces the [start, start+length) range of the group. The replacement inherits the formatting of the
        /// first covered run, and the text of the remaining covered runs is cleared.
        /// </summary>
        /// <param name="runs">文本运行组。The text-run group.</param>
        /// <param name="start">替换区间起点。The start offset of the range.</param>
        /// <param name="length">替换区间长度，必须大于 0。The range length; must be greater than zero.</param>
        /// <param name="replacement">替换文本，null 视为空串。The replacement text; null is treated as empty.</param>
        /// <returns>区间有效并完成替换返回 true。True when the range was valid and replaced.</returns>
        public static bool ReplaceRange(List<TextRun> runs, int start, int length, string replacement)
        {
            if (runs == null || start < 0 || length <= 0) return false;
            if (replacement == null) replacement = string.Empty;

            int offset = 0;
            int first = -1;
            int last = -1;
            int startInFirst = 0;
            int endInLast = 0;
            for (int i = 0; i < runs.Count; i++)
            {
                int runLength = runs[i].Text == null ? 0 : runs[i].Text.Length;
                int runEnd = offset + runLength;
                if (first < 0 && start < runEnd)
                {
                    first = i;
                    startInFirst = start - offset;
                }
                int rangeEnd = start + length;
                if (first >= 0 && rangeEnd <= runEnd)
                {
                    last = i;
                    endInLast = rangeEnd - offset;
                    break;
                }
                offset = runEnd;
            }
            if (first < 0 || last < 0) return false;

            if (first == last)
            {
                string text = runs[first].Text ?? string.Empty;
                runs[first].Text = text.Substring(0, startInFirst) + replacement + text.Substring(endInLast);
                return true;
            }

            string firstText = runs[first].Text ?? string.Empty;
            runs[first].Text = firstText.Substring(0, startInFirst) + replacement;
            for (int i = first + 1; i < last; i++)
            {
                runs[i].Text = string.Empty;
            }
            string lastText = runs[last].Text ?? string.Empty;
            runs[last].Text = lastText.Substring(endInLast);
            return true;
        }

        /// <summary>
        /// 判断字符是否属于单词字符（字母、数字或下划线），用于全词匹配边界判断。
        /// Reports whether a character is a word character (letter, digit or underscore), used for whole-word
        /// boundary checks.
        /// </summary>
        /// <param name="value">待判断字符。The character to test.</param>
        /// <returns>是单词字符返回 true。True when the character is a word character.</returns>
        public static bool IsWordCharacter(char value)
        {
            return char.IsLetterOrDigit(value) || value == '_';
        }
    }
}
