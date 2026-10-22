using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 查找替换选项：大小写、全词匹配、正则表达式、仅替换首处以及搜索范围。
    /// Find/replace options: case sensitivity, whole-word matching, regular expressions, first-match-only and
    /// the search scope.
    /// </summary>
    public sealed class WordFindReplaceOptions
    {
        /// <summary>是否区分大小写，默认不区分。Whether matching is case-sensitive; false by default.</summary>
        public bool MatchCase { get; set; }
        /// <summary>是否只匹配完整单词，默认否。Whether only whole words match; false by default.</summary>
        public bool WholeWord { get; set; }
        /// <summary>是否把查找文本当作 .NET 正则表达式，默认否。Whether the search text is a .NET regular expression; false by default.</summary>
        public bool UseRegex { get; set; }
        /// <summary>是否只替换文档中的第一处匹配，默认否。Whether only the first match in the document is replaced; false by default.</summary>
        public bool FirstOnly { get; set; }
        /// <summary>是否搜索页眉和页脚，默认是。Whether headers and footers are searched; true by default.</summary>
        public bool IncludeHeadersAndFooters { get; set; } = true;
        /// <summary>是否搜索表格单元格，默认是。Whether table cells are searched; true by default.</summary>
        public bool IncludeTables { get; set; } = true;
    }

    /// <summary>
    /// 文档查找替换引擎：在正文、表格、页眉页脚中替换文本，支持跨运行命中、大小写/全词匹配与正则表达式。
    /// Document find/replace engine: replaces text in the body, tables, headers and footers, with support for
    /// matches split across runs, case/whole-word matching and regular expressions.
    /// </summary>
    public static class WordFindReplace
    {
        /// <summary>正则表达式的匹配超时，避免恶意模式造成灾难性回溯。Match timeout for regular expressions, guarding against catastrophic backtracking from malicious patterns.</summary>
        private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);

        /// <summary>
        /// 以默认选项替换文档中的全部匹配。
        /// Replaces every match of the search text in the document using the default options.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="searchText">查找文本（或正则表达式）。The search text (or regular expression).</param>
        /// <param name="replacement">替换文本，null 视为空串。The replacement text; null is treated as empty.</param>
        /// <returns>实际替换次数。The number of replacements made.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="searchText"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="searchText"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="searchText"/> 为空或正则表达式非法。Thrown when <paramref name="searchText"/> is empty or the regular expression is invalid.</exception>
        public static int Replace(OfficeDocument document, string searchText, string replacement)
        {
            return Replace(document, searchText, replacement, new WordFindReplaceOptions());
        }

        /// <summary>
        /// 使用指定选项替换文档中的匹配文本；替换文字继承每处首个命中运行的格式。
        /// Replaces matching text in the document with the given options; the replacement inherits the
        /// formatting of the first matching run at each occurrence.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="searchText">查找文本（或正则表达式）。The search text (or regular expression).</param>
        /// <param name="replacement">替换文本，null 视为空串。The replacement text; null is treated as empty.</param>
        /// <param name="options">替换选项，null 时使用默认值。The options; defaults are used when null.</param>
        /// <returns>实际替换次数。The number of replacements made.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="searchText"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="searchText"/> is null.</exception>
        /// <exception cref="ArgumentException"><paramref name="searchText"/> 为空或正则表达式非法。Thrown when <paramref name="searchText"/> is empty or the regular expression is invalid.</exception>
        public static int Replace(OfficeDocument document, string searchText, string replacement, WordFindReplaceOptions options)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (searchText == null) throw new ArgumentNullException("searchText");
            if (searchText.Length == 0) throw new ArgumentException("Search text cannot be empty.", "searchText");
            if (options == null) options = new WordFindReplaceOptions();
            if (replacement == null) replacement = string.Empty;

            Regex regex = options.UseRegex
                ? new Regex(searchText, options.MatchCase ? RegexOptions.None : RegexOptions.IgnoreCase, RegexTimeout)
                : null;

            int count = 0;
            bool firstDone = false;
            WordDocumentWalker.ForEachParagraph(
                document,
                options.IncludeHeadersAndFooters,
                options.IncludeTables,
                delegate (Paragraph paragraph)
                {
                    if (firstDone) return;
                    int before = count;
                    List<List<TextRun>> groups = WordTextRange.GetGroups(paragraph);
                    for (int g = 0; g < groups.Count && !firstDone; g++)
                    {
                        List<TextRun> runs = groups[g];
                        string text = WordTextRange.GetText(runs);
                        if (text.Length == 0) continue;

                        List<TextMatchRange> matches = FindMatches(text, searchText, regex, options);
                        if (matches.Count == 0) continue;

                        if (options.FirstOnly)
                        {
                            if (WordTextRange.ReplaceRange(runs, matches[0].Index, matches[0].Length, replacement))
                            {
                                count++;
                            }
                            firstDone = true;
                            continue;
                        }

                        for (int m = matches.Count - 1; m >= 0; m--)
                        {
                            if (WordTextRange.ReplaceRange(runs, matches[m].Index, matches[m].Length, replacement))
                            {
                                count++;
                            }
                        }
                    }
                    if (count != before) RemoveEmptyRuns(paragraph);
                });

            return count;
        }

        /// <summary>
        /// 单次替换匹配在合并文本中的位置。
        /// The location of one match within the merged text.
        /// </summary>
        private sealed class TextMatchRange
        {
            /// <summary>匹配起始偏移。The match start offset.</summary>
            public int Index;
            /// <summary>匹配长度。The match length.</summary>
            public int Length;
        }

        /// <summary>
        /// 在合并文本中查找全部匹配（可限定首处），并应用全词边界过滤。
        /// Finds all matches in the merged text (optionally only the first), applying whole-word boundary
        /// filtering.
        /// </summary>
        /// <param name="text">合并后的文本。The merged text.</param>
        /// <param name="searchText">查找文本。The search text.</param>
        /// <param name="regex">已编译的正则表达式，普通查找时为 null。The compiled regex, or null for plain search.</param>
        /// <param name="options">替换选项。The options.</param>
        /// <returns>按出现顺序排列的匹配集合。The matches in occurrence order.</returns>
        private static List<TextMatchRange> FindMatches(string text, string searchText, Regex regex, WordFindReplaceOptions options)
        {
            List<TextMatchRange> matches = new List<TextMatchRange>();
            if (regex != null)
            {
                foreach (Match match in regex.Matches(text))
                {
                    if (match.Length == 0) continue;
                    if (options.WholeWord && !IsWholeWord(text, match.Index, match.Length)) continue;
                    TextMatchRange range = new TextMatchRange();
                    range.Index = match.Index;
                    range.Length = match.Length;
                    matches.Add(range);
                    if (options.FirstOnly) break;
                }
                return matches;
            }

            StringComparison comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            int position = 0;
            while (position <= text.Length - searchText.Length)
            {
                int found = text.IndexOf(searchText, position, comparison);
                if (found < 0) break;
                if (!options.WholeWord || IsWholeWord(text, found, searchText.Length))
                {
                    TextMatchRange range = new TextMatchRange();
                    range.Index = found;
                    range.Length = searchText.Length;
                    matches.Add(range);
                    if (options.FirstOnly) break;
                    position = found + searchText.Length;
                }
                else
                {
                    position = found + 1;
                }
            }
            return matches;
        }

        /// <summary>
        /// 判断匹配区间两侧是否均为单词边界（非单词字符或文本首尾）。
        /// Reports whether both sides of the match are word boundaries (non-word characters or the text ends).
        /// </summary>
        /// <param name="text">合并文本。The merged text.</param>
        /// <param name="index">匹配起点。The match start.</param>
        /// <param name="length">匹配长度。The match length.</param>
        /// <returns>满足全词条件返回 true。True when the whole-word condition holds.</returns>
        private static bool IsWholeWord(string text, int index, int length)
        {
            if (index > 0 && WordTextRange.IsWordCharacter(text[index - 1])) return false;
            int end = index + length;
            if (end < text.Length && WordTextRange.IsWordCharacter(text[end])) return false;
            return true;
        }

        /// <summary>
        /// 删除被替换操作清空的文本运行，避免模型中残留空运行。
        /// Removes text runs emptied by the replacement, so no empty runs linger in the model.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        private static void RemoveEmptyRuns(Paragraph paragraph)
        {
            paragraph.Inlines.RemoveAll(delegate (Inline inline)
            {
                TextRun run = inline as TextRun;
                return run != null && string.IsNullOrEmpty(run.Text);
            });
        }
    }
}
