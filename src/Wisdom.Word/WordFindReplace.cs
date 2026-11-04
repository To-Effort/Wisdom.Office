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
        /// <exception cref="ArgumentException"><paramref name="searchText"/> 为空、正则表达式非法，或正则匹配超时；抛出超时文档未被改动。Thrown when <paramref name="searchText"/> is empty, the regular expression is invalid, or matching times out; on a timeout the document is untouched.</exception>
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
        /// <exception cref="ArgumentException"><paramref name="searchText"/> 为空、正则表达式非法，或正则匹配超时；抛出超时文档未被改动。Thrown when <paramref name="searchText"/> is empty, the regular expression is invalid, or matching times out; on a timeout the document is untouched.</exception>
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

            // 找匹配与改文本分成两趟：正则只在计划这一趟跑，所以超时抛出时文档一个字都没被改。
            // 原先到处是就地改写——第一段已经替换完、下一段才超时，调用方拿到的是改到一半的文档，
            // 既不知道改了几处，也没有回滚的余地（遍历器没有事务）。
            // Planning and applying are now two passes: the regex runs only while planning, so a timeout
            // escapes with not one character changed. The old loop rewrote in place, leaving callers with a
            // half-replaced document — no count, and nothing to roll back to, because the walker has no
            // transaction.
            List<ReplacementPlan> plan;
            try
            {
                plan = BuildPlan(document, searchText, regex, options);
            }
            catch (RegexMatchTimeoutException timeout)
            {
                throw new ArgumentException("正则匹配超时（" + RegexTimeout.TotalSeconds
                    + " 秒），文档未被改动。", timeout);
            }

            int count = 0;
            List<Paragraph> touched = new List<Paragraph>();
            for (int p = 0; p < plan.Count; p++)
            {
                ReplacementPlan entry = plan[p];
                // 同一组内从后往前改：改掉的长度不会移动前面匹配的下标。
                // Within one group the offsets run backwards, so a rewritten length never shifts the
                // indexes of the matches before it.
                for (int m = entry.Ranges.Count - 1; m >= 0; m--)
                {
                    if (WordTextRange.ReplaceRange(entry.Runs, entry.Ranges[m].Index, entry.Ranges[m].Length, replacement))
                    {
                        count++;
                        if (!touched.Contains(entry.Paragraph)) touched.Add(entry.Paragraph);
                    }
                }
            }

            for (int t = 0; t < touched.Count; t++) RemoveEmptyRuns(touched[t]);

            return count;
        }

        /// <summary>
        /// 遍历一遍文档，把每处替换记成计划；不改动任何文本，正则匹配只在这趟里发生。
        /// Walks the document and records every replacement as a plan without changing any text; this is the
        /// only pass where the regular expression runs.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="searchText">查找文本。The search text.</param>
        /// <param name="regex">已编译的正则表达式，普通查找时为 null。The compiled regex, or null for plain search.</param>
        /// <param name="options">替换选项。The options.</param>
        /// <returns>按遍历顺序排列的替换计划。The replacement plans in walk order.</returns>
        private static List<ReplacementPlan> BuildPlan(OfficeDocument document, string searchText, Regex regex, WordFindReplaceOptions options)
        {
            List<ReplacementPlan> plan = new List<ReplacementPlan>();
            bool firstDone = false;
            WordDocumentWalker.ForEachParagraph(
                document,
                options.IncludeHeadersAndFooters,
                options.IncludeTables,
                delegate (Paragraph paragraph)
                {
                    if (firstDone) return;
                    List<List<TextRun>> groups = WordTextRange.GetGroups(paragraph);
                    for (int g = 0; g < groups.Count && !firstDone; g++)
                    {
                        List<TextRun> runs = groups[g];
                        string text = WordTextRange.GetText(runs);
                        if (text.Length == 0) continue;

                        List<TextMatchRange> matches = FindMatches(text, searchText, regex, options);
                        if (matches.Count == 0) continue;

                        ReplacementPlan entry = new ReplacementPlan();
                        entry.Paragraph = paragraph;
                        entry.Runs = runs;
                        entry.Ranges = matches;
                        plan.Add(entry);
                        if (options.FirstOnly) firstDone = true;
                    }
                });
            return plan;
        }

        /// <summary>
        /// 一处待落地的替换：目标段落、该组运行列表与组内偏移。
        /// One planned replacement: the target paragraph, that group's runs and the offsets inside the group.
        /// </summary>
        private sealed class ReplacementPlan
        {
            /// <summary>被改动的段落，用于替换完成后清理空运行。The paragraph being changed, so emptied runs can be cleaned afterwards.</summary>
            public Paragraph Paragraph;
            /// <summary>命中所在的运行分组，偏移按这一组的合并文本计算。The run group the offsets belong to.</summary>
            public List<TextRun> Runs;
            /// <summary>组内按出现顺序排列的匹配。The matches inside this group, in occurrence order.</summary>
            public List<TextMatchRange> Ranges;
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
