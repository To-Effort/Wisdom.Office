using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 对统一文档模型执行规范化：段落整理、样式统一、伪标题升级、表格/图片清理与目录重建。
    /// Normalizes a unified document model: paragraph cleanup, style unification, fake heading
    /// promotion, table/image cleanup and table of contents regeneration.
    /// </summary>
    public static class WordNormalizer
    {
        /// <summary>变更类别：段落整理。Change category for paragraph cleanup.</summary>
        public const string CategoryParagraph = "段落整理";
        /// <summary>变更类别：样式整理。Change category for style unification.</summary>
        public const string CategoryStyle = "样式整理";
        /// <summary>变更类别：标题整理。Change category for heading cleanup.</summary>
        public const string CategoryHeading = "标题整理";
        /// <summary>变更类别：表格整理。Change category for table cleanup.</summary>
        public const string CategoryTable = "表格整理";
        /// <summary>变更类别：图片整理。Change category for image cleanup.</summary>
        public const string CategoryImage = "图片整理";
        /// <summary>变更类别：目录整理。Change category for table-of-contents rebuild.</summary>
        public const string CategoryToc = "目录整理";

        private static readonly Regex SentencePattern = new Regex(@"[^。！？!?；;\r\n]+[。！？!?；;]?", RegexOptions.Compiled);
        private static readonly Regex ChapterPattern = new Regex(@"^第[一二三四五六七八九十百千0-9]+[章节篇部分回]", RegexOptions.Compiled);

        /// <summary>
        /// 执行规范化管道：段落整理、样式统一、表格/图片清理、目录重建，并做整理后一致性校验。
        /// Runs the normalization pipeline: paragraph cleanup, style unification, table/image cleanup,
        /// TOC rebuild, and a post-normalization consistency check.
        /// </summary>
        /// <param name="document">待规范化的文档。The document to normalize.</param>
        /// <param name="options">规范化选项，为 null 时使用默认。The options; defaults are used when null.</param>
        /// <returns>规范化结果（变更、问题、前后指标）。The result with changes, issues and before/after metrics.</returns>
        public static WordNormalizeResult Normalize(OfficeDocument document, WordNormalizeOptions options)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (options == null) options = new WordNormalizeOptions();

            WordNormalizeResult result = new WordNormalizeResult();
            result.Document = document;
            result.Before = WordDocumentMetrics.Measure(document);

            int styledRuns = 0;
            int styledParagraphs = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                string location = "第" + (s + 1).ToString(CultureInfo.InvariantCulture) + "节";
                CleanBlockList(section.Blocks, options, result, location, true);
                PrepareHeadings(section.Blocks, options, result, location);
                UnifyEndings(section.Blocks, options, location);
                NormalizeStyles(section.Blocks, options, ref styledRuns, ref styledParagraphs, true);
                NormalizeImages(section, section.Blocks, options, result, location);
                NormalizeTables(section, options, result, location);
            }

            if (styledRuns > 0)
            {
                result.AddChange(CategoryStyle, "统一正文字体/字号，共 " + styledRuns.ToString(CultureInfo.InvariantCulture) + " 处文本", string.Empty);
            }
            if (styledParagraphs > 0)
            {
                result.AddChange(CategoryStyle, "统一行距/段距，共 " + styledParagraphs.ToString(CultureInfo.InvariantCulture) + " 个段落", string.Empty);
            }

            if (options.RegenerateTableOfContents)
            {
                BuildTableOfContents(document, options, result);
            }

            result.After = WordDocumentMetrics.Measure(document);
            result.Issues.AddRange(WordConsistencyChecker.Check(document, options));
            return result;
        }

        /// <summary>更新 PAGEREF 域的显示文本（用于分页后刷新页码）。Updates PAGEREF field display texts (used to refresh page numbers after pagination).</summary>
        /// <param name="document">目标文档。The document whose page numbers are refreshed.</param>
        /// <param name="pageNumbersByBookmark">书签名到页码的映射。Bookmark name to page number map.</param>
        /// <returns>被更新的域数量。The number of fields updated.</returns>
        public static int UpdateTocPageNumbers(OfficeDocument document, IDictionary<string, int> pageNumbersByBookmark)
        {
            if (document == null || pageNumbersByBookmark == null || pageNumbersByBookmark.Count == 0) return 0;
            int updated = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                updated += UpdatePageNumbers(document.Sections[s].Blocks, pageNumbersByBookmark);
            }
            return updated;
        }

        private static int UpdatePageNumbers(List<Block> blocks, IDictionary<string, int> pages)
        {
            int updated = 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    for (int r = 0; r < paragraph.Inlines.Count; r++)
                    {
                        FieldRun field = paragraph.Inlines[r] as FieldRun;
                        if (field == null) continue;
                        string bookmark = GetPageRefBookmark(field.Instruction);
                        if (bookmark == null) continue;
                        int page;
                        if (pages.TryGetValue(bookmark, out page))
                        {
                            field.DisplayText = page.ToString(CultureInfo.InvariantCulture);
                            updated++;
                        }
                    }
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                        {
                            updated += UpdatePageNumbers(table.Rows[r].Cells[c].Blocks, pages);
                        }
                    }
                }
            }
            return updated;
        }

        private static string GetPageRefBookmark(string instruction)
        {
            if (string.IsNullOrEmpty(instruction)) return null;
            string[] parts = instruction.Trim().Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2) return null;
            if (!parts[0].Equals("PAGEREF", StringComparison.OrdinalIgnoreCase)) return null;
            return parts[1];
        }

        // ---------- 段落整理 ----------

        private static void CleanBlockList(List<Block> blocks, WordNormalizeOptions options, WordNormalizeResult result, string location, bool isBody)
        {
            if (options.NormalizeWhitespace || options.UnifyEndingPunctuation || options.AddMissingEndingPunctuation)
            {
                for (int i = 0; i < blocks.Count; i++)
                {
                    Paragraph paragraph = blocks[i] as Paragraph;
                    if (paragraph != null)
                    {
                        NormalizeText(paragraph, options);
                        continue;
                    }
                    Table table = blocks[i] as Table;
                    if (table != null)
                    {
                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                            {
                                CleanBlockList(table.Rows[r].Cells[c].Blocks, options, result, location + " 表格", false);
                            }
                        }
                    }
                }
            }

            if (!isBody) return;
            if (options.RemoveEmptyParagraphs) RemoveEmptyParagraphs(blocks, result, location);
            if (options.MergeShortParagraphs) MergeShortParagraphs(blocks, options, result, location);
            if (options.SplitLongParagraphs) SplitLongParagraphs(blocks, options, result, location);
        }

        private static void NormalizeText(Paragraph paragraph, WordNormalizeOptions options)
        {
            if (!options.NormalizeWhitespace || IsTocParagraph(paragraph)) return;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run == null) continue;
                run.Text = CollapseWhitespace(run.Text);
            }
            RemoveEmptyRuns(paragraph);
        }

        private static void UnifyEndings(List<Block> blocks, WordNormalizeOptions options, string location)
        {
            if (!options.UnifyEndingPunctuation && !options.AddMissingEndingPunctuation) return;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph == null)
                {
                    Table table = blocks[i] as Table;
                    if (table == null) continue;
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                        {
                            UnifyEndings(table.Rows[r].Cells[c].Blocks, options, location);
                        }
                    }
                    continue;
                }
                if (IsTocParagraph(paragraph) || paragraph.HeadingLevel > 0) continue;
                if (!string.IsNullOrEmpty(paragraph.ListMarker) || WordDocumentMetrics.HasImage(paragraph)) continue;

                if (options.UnifyEndingPunctuation) UnifyEndingPunctuation(paragraph);
                if (!options.AddMissingEndingPunctuation) continue;

                string text = paragraph.GetText();
                if (!IsMostlyCjk(text) || text.Trim().Length < 10) continue;
                char last = GetLastVisibleChar(paragraph);
                if (last != '\0' && !IsSentenceEnder(last)) AppendCharacter(paragraph, '。');
            }
        }

        private static string CollapseWhitespace(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            if (normalized.IndexOf('\n') >= 0)
            {
                string[] lines = normalized.Split('\n');
                StringBuilder joined = new StringBuilder();
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (line.Length == 0) continue;
                    if (joined.Length == 0)
                    {
                        joined.Append(line);
                        continue;
                    }
                    char left = joined[joined.Length - 1];
                    char right = line[0];
                    if (!IsWideCharacter(left) && !IsWideCharacter(right)) joined.Append(' ');
                    joined.Append(line);
                }
                normalized = joined.ToString();
            }

            StringBuilder builder = new StringBuilder(normalized.Length);
            char previous = '\0';
            for (int i = 0; i < normalized.Length; i++)
            {
                char c = normalized[i];
                if (c == ' ' || c == '\t' || c == '\u3000')
                {
                    if (previous == ' ' || previous == '\t' || previous == '\u3000' || previous == '\0')
                    {
                        continue;
                    }
                    builder.Append(' ');
                    previous = ' ';
                    continue;
                }
                builder.Append(c);
                previous = c;
            }
            return builder.ToString().TrimEnd();
        }

        private static void RemoveEmptyRuns(Paragraph paragraph)
        {
            for (int i = paragraph.Inlines.Count - 1; i >= 0; i--)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null && run.Text.Length == 0) paragraph.Inlines.RemoveAt(i);
            }
        }

        private static void UnifyEndingPunctuation(Paragraph paragraph)
        {
            int runIndex;
            int charIndex;
            char last = GetLastVisibleChar(paragraph, out runIndex, out charIndex);
            if (last == '\0') return;

            char mapped = MapToFullWidth(last);
            if (mapped != last && IsMostlyCjk(paragraph.GetText()))
            {
                TextRun run = (TextRun)paragraph.Inlines[runIndex];
                run.Text = run.Text.Substring(0, charIndex) + mapped + run.Text.Substring(charIndex + 1);
                last = mapped;
            }

            // 折叠重复的段尾标点，例如“。。”或“！！！”
            while (true)
            {
                char previous;
                int previousRun;
                int previousChar;
                if (!TryGetPreviousVisibleChar(paragraph, runIndex, charIndex, out previous, out previousRun, out previousChar)) break;
                if (!IsDuplicateEnder(previous, last)) break;
                TextRun run = (TextRun)paragraph.Inlines[previousRun];
                run.Text = run.Text.Remove(previousChar, 1);
                runIndex = previousRun;
                charIndex = previousChar;
                char next = GetLastVisibleChar(paragraph, out runIndex, out charIndex);
                if (next == '\0') break;
                last = next;
            }
        }

        private static char MapToFullWidth(char c)
        {
            switch (c)
            {
                case '.': return '。';
                case ',': return '，';
                case ';': return '；';
                case ':': return '：';
                case '!': return '！';
                case '?': return '？';
                default: return c;
            }
        }

        private static bool IsDuplicateEnder(char previous, char current)
        {
            if (previous == current) return IsSentenceEnder(current);
            return (previous == '。' || previous == '．') && current == '。';
        }

        private static bool IsSentenceEnder(char c)
        {
            return c == '。' || c == '！' || c == '？' || c == '；' || c == '.' || c == '!' || c == '?' || c == ';';
        }

        private static char GetLastVisibleChar(Paragraph paragraph)
        {
            int runIndex;
            int charIndex;
            return GetLastVisibleChar(paragraph, out runIndex, out charIndex);
        }

        private static char GetLastVisibleChar(Paragraph paragraph, out int runIndex, out int charIndex)
        {
            for (int i = paragraph.Inlines.Count - 1; i >= 0; i--)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run == null) continue;
                for (int c = run.Text.Length - 1; c >= 0; c--)
                {
                    char value = run.Text[c];
                    if (value == ' ' || value == '\t' || value == '\n' || value == '\u3000') continue;
                    runIndex = i;
                    charIndex = c;
                    return value;
                }
            }
            runIndex = -1;
            charIndex = -1;
            return '\0';
        }

        private static bool TryGetPreviousVisibleChar(Paragraph paragraph, int runIndex, int charIndex, out char previous, out int previousRun, out int previousChar)
        {
            previous = '\0';
            previousRun = -1;
            previousChar = -1;
            for (int i = runIndex; i >= 0; i--)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run == null) continue;
                int start = i == runIndex ? charIndex - 1 : run.Text.Length - 1;
                for (int c = start; c >= 0; c--)
                {
                    char value = run.Text[c];
                    if (value == ' ' || value == '\t' || value == '\n' || value == '\u3000') continue;
                    previous = value;
                    previousRun = i;
                    previousChar = c;
                    return true;
                }
            }
            return false;
        }

        private static void AppendCharacter(Paragraph paragraph, char value)
        {
            for (int i = paragraph.Inlines.Count - 1; i >= 0; i--)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null)
                {
                    run.Text = run.Text.TrimEnd() + value;
                    return;
                }
            }
        }

        private static void RemoveEmptyParagraphs(List<Block> blocks, WordNormalizeResult result, string location)
        {
            for (int i = blocks.Count - 1; i >= 0; i--)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph == null) continue;
                string text = paragraph.GetText().Trim();
                if (text.Length > 0 || WordDocumentMetrics.HasImage(paragraph)) continue;
                if (paragraph.PageBreakBefore)
                {
                    blocks[i] = new PageBreak();
                    continue;
                }
                if (HasFields(paragraph)) continue;
                blocks.RemoveAt(i);
                result.AddChange(CategoryParagraph, "删除空段落", location + " 第" + (i + 1).ToString(CultureInfo.InvariantCulture) + "块");
            }
        }

        private static void MergeShortParagraphs(List<Block> blocks, WordNormalizeOptions options, WordNormalizeResult result, string location)
        {
            int merged = 0;
            for (int i = 0; i < blocks.Count - 1; i++)
            {
                Paragraph current = blocks[i] as Paragraph;
                if (current == null || !IsMergeCandidate(current, options)) continue;
                while (i < blocks.Count - 1)
                {
                    Paragraph next = blocks[i + 1] as Paragraph;
                    if (next == null || !IsMergeCandidate(next, options)) break;
                    if (next.Alignment != current.Alignment) break;
                    if (current.GetText().Length > options.ShortParagraphThreshold * 2) break;
                    char last = GetLastVisibleChar(current);
                    if (last != '\0' && IsSentenceEnder(last)) break;

                    string left = current.GetText();
                    string right = next.GetText();
                    if (NeedsSpace(left, right))
                    {
                        current.Inlines.Add(new TextRun(" ", GetFirstTextFormat(next)));
                    }
                    for (int r = 0; r < next.Inlines.Count; r++)
                    {
                        current.Inlines.Add(next.Inlines[r]);
                    }
                    blocks.RemoveAt(i + 1);
                    merged++;
                }
                if (merged > 0)
                {
                    result.AddChange(CategoryParagraph, "合并连续短段落", location + " 第" + (i + 1).ToString(CultureInfo.InvariantCulture) + "块");
                    merged = 0;
                }
            }
        }

        private static bool IsMergeCandidate(Paragraph paragraph, WordNormalizeOptions options)
        {
            if (paragraph.HeadingLevel > 0) return false;
            if (!string.IsNullOrEmpty(paragraph.ListMarker)) return false;
            if (paragraph.BookmarkName != null) return false;
            if (WordDocumentMetrics.HasImage(paragraph)) return false;
            if (HasFields(paragraph)) return false;
            if (paragraph.PageBreakBefore) return false;
            string text = paragraph.GetText().Trim();
            if (text.Length == 0) return false;
            return text.Length <= options.ShortParagraphThreshold;
        }

        private static void SplitLongParagraphs(List<Block> blocks, WordNormalizeOptions options, WordNormalizeResult result, string location)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph == null) continue;
                if (paragraph.HeadingLevel > 0 || !string.IsNullOrEmpty(paragraph.ListMarker)) continue;
                if (WordDocumentMetrics.HasImage(paragraph) || HasFields(paragraph)) continue;
                string text = paragraph.GetText();
                if (text.Length <= options.LongParagraphThreshold) continue;

                List<string> chunks = SplitSentences(text, options.LongParagraphThreshold);
                if (chunks.Count <= 1) continue;

                ReplaceParagraphText(paragraph, chunks[0]);
                for (int c = 1; c < chunks.Count; c++)
                {
                    Paragraph copy = paragraph.Clone();
                    ReplaceParagraphText(copy, chunks[c]);
                    blocks.Insert(i + c, copy);
                }
                i += chunks.Count - 1;
                result.AddChange(CategoryParagraph, "拆分过长段落为 " + chunks.Count.ToString(CultureInfo.InvariantCulture) + " 段", location + " 第" + (i + 1).ToString(CultureInfo.InvariantCulture) + "块");
            }
        }

        private static List<string> SplitSentences(string text, int threshold)
        {
            List<string> sentences = new List<string>();
            foreach (Match match in SentencePattern.Matches(text))
            {
                string sentence = match.Value.Trim();
                if (sentence.Length > 0) sentences.Add(sentence);
            }
            List<string> chunks = new List<string>();
            StringBuilder current = new StringBuilder();
            for (int i = 0; i < sentences.Count; i++)
            {
                if (current.Length > 0 && current.Length + sentences[i].Length > threshold)
                {
                    chunks.Add(current.ToString());
                    current.Length = 0;
                }
                current.Append(sentences[i]);
            }
            if (current.Length > 0) chunks.Add(current.ToString());
            if (chunks.Count == 0) chunks.Add(text);
            return chunks;
        }

        private static void ReplaceParagraphText(Paragraph paragraph, string text)
        {
            TextFormat format = GetFirstTextFormat(paragraph);
            paragraph.Inlines.Clear();
            paragraph.Inlines.Add(new TextRun(text, format));
        }

        private static TextFormat GetFirstTextFormat(Paragraph paragraph)
        {
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null) return run.Format;
            }
            return new TextFormat();
        }

        private static bool NeedsSpace(string left, string right)
        {
            if (left.Length == 0 || right.Length == 0) return false;
            return !IsWideCharacter(left[left.Length - 1]) && !IsWideCharacter(right[0]);
        }

        private static bool IsWideCharacter(char c)
        {
            return (c >= 0x2E80 && c <= 0x9FFF) || (c >= 0xF900 && c <= 0xFAFF)
                || (c >= 0xFE30 && c <= 0xFE4F) || (c >= 0xFF00 && c <= 0xFFEF)
                || (c >= 0x20000 && c <= 0x2FA1F);
        }

        private static bool IsMostlyCjk(string text)
        {
            int cjk = 0;
            int latin = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (char.IsDigit(c)) continue;
                if (IsWideCharacter(c) && !char.IsPunctuation(c) && !char.IsSymbol(c)) cjk++;
                else if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')) latin++;
            }
            return cjk > 0 && cjk >= latin;
        }

        private static bool HasFields(Paragraph paragraph)
        {
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                if (paragraph.Inlines[i] is FieldRun) return true;
            }
            return false;
        }

        // ---------- 标题整理 ----------

        private static void PrepareHeadings(List<Block> blocks, WordNormalizeOptions options, WordNormalizeResult result, string location)
        {
            if (options.PromoteFakeHeadings) PromoteFakeHeadings(blocks, options, result, location);
            if (options.FixHeadingLevelSkips) FixHeadingLevelSkips(blocks, result, location);
        }

        private static void PromoteFakeHeadings(List<Block> blocks, WordNormalizeOptions options, WordNormalizeResult result, string location)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph == null || paragraph.HeadingLevel > 0) continue;
                if (IsTocParagraph(paragraph)) continue;
                if (IsTableOfContentsTitle(paragraph, options)) continue;
                if (WordDocumentMetrics.HasImage(paragraph) || HasFields(paragraph)) continue;

                string text = paragraph.GetText().Trim();
                if (text.Length == 0 || text.Length > 80) continue;
                char last = GetLastVisibleChar(paragraph);
                if (IsSentenceEnder(last) && text.Length > 12) continue;

                double maxFontSize = 0;
                bool bold = false;
                for (int r = 0; r < paragraph.Inlines.Count; r++)
                {
                    TextRun run = paragraph.Inlines[r] as TextRun;
                    if (run == null) continue;
                    if (run.Format != null)
                    {
                        if (run.Format.FontSize > maxFontSize) maxFontSize = run.Format.FontSize;
                        if (run.Format.Bold) bold = true;
                    }
                }
                if (maxFontSize <= 0) maxFontSize = options.BodyFontSize;

                double ratio = maxFontSize / Math.Max(1, options.BodyFontSize);
                bool centered = paragraph.Alignment == ParagraphAlignment.Center;
                bool looksLikeHeading = ratio >= 1.15
                    || (bold && maxFontSize >= options.FakeHeadingMinFontSize)
                    || (bold && centered && ratio >= 1.08);
                if (!looksLikeHeading) continue;
                if (centered && !bold && ratio < 1.3 && text.Length > 30) continue;

                int level;
                if (ChapterPattern.IsMatch(text)) level = 1;
                else if (ratio >= 1.5) level = 1;
                else if (ratio >= 1.3) level = 2;
                else if (ratio >= 1.15) level = 3;
                else level = 4;

                paragraph.HeadingLevel = level;
                paragraph.KeepWithNext = true;
                paragraph.PageBreakBefore = false;
                paragraph.FirstLineIndent = 0;
                result.AddChange(CategoryHeading, "伪标题升级为 " + level.ToString(CultureInfo.InvariantCulture) + " 级标题: " + Shorten(text),
                    location + " 第" + (i + 1).ToString(CultureInfo.InvariantCulture) + "块");
            }
        }

        private static void FixHeadingLevelSkips(List<Block> blocks, WordNormalizeResult result, string location)
        {
            int previous = 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph == null || paragraph.HeadingLevel <= 0) continue;
                int level = paragraph.HeadingLevel;
                if (level < 1) level = 1;
                if (level > 9) level = 9;
                if (previous > 0 && level > previous + 1)
                {
                    result.AddChange(CategoryHeading, "标题层级跳跃 " + level.ToString(CultureInfo.InvariantCulture) + " -> " + (previous + 1).ToString(CultureInfo.InvariantCulture),
                        location + " 第" + (i + 1).ToString(CultureInfo.InvariantCulture) + "块");
                    level = previous + 1;
                }
                paragraph.HeadingLevel = level;
                previous = level;
            }
        }

        // ---------- 样式整理 ----------

        private static void NormalizeStyles(List<Block> blocks, WordNormalizeOptions options, ref int styledRuns, ref int styledParagraphs, bool isBody)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    NormalizeParagraphStyle(paragraph, options, ref styledRuns, ref styledParagraphs, isBody);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        NormalizeStyles(table.Rows[r].Cells[c].Blocks, options, ref styledRuns, ref styledParagraphs, false);
                    }
                }
            }
        }

        private static void NormalizeParagraphStyle(Paragraph paragraph, WordNormalizeOptions options, ref int styledRuns, ref int styledParagraphs, bool isBodyLevel)
        {
            if (!options.NormalizeStyles) return;
            if (IsTocParagraph(paragraph)) return;
            if (IsTableOfContentsTitle(paragraph, options)) return;

            bool isHeading = paragraph.HeadingLevel > 0;
            bool isList = !string.IsNullOrEmpty(paragraph.ListMarker);
            string eastAsia = isHeading ? options.HeadingEastAsiaFont : options.BodyEastAsiaFont;
            string ascii = isHeading ? options.HeadingAsciiFont : options.BodyAsciiFont;
            double fontSize = isHeading ? options.GetHeadingFontSize(paragraph.HeadingLevel) : options.BodyFontSize;

            if (isHeading)
            {
                paragraph.KeepWithNext = true;
                paragraph.FirstLineIndent = 0;
                paragraph.SpaceBefore = options.HeadingSpaceBefore;
                paragraph.SpaceAfter = options.HeadingSpaceAfter;
                paragraph.LineSpacing = options.HeadingLineSpacing;
                paragraph.StyleName = "Heading " + paragraph.HeadingLevel.ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                paragraph.SpaceBefore = options.BodySpaceBefore;
                paragraph.SpaceAfter = options.BodySpaceAfter;
                paragraph.LineSpacing = options.BodyLineSpacing;
                paragraph.StyleName = isList ? "List Paragraph" : "Normal";
                if (isBodyLevel && !isList && !WordDocumentMetrics.HasImage(paragraph)
                    && paragraph.Alignment != ParagraphAlignment.Center && paragraph.GetText().Trim().Length > 0)
                {
                    paragraph.FirstLineIndent = options.FirstLineIndentCharacters > 0
                        ? options.FirstLineIndentCharacters * options.BodyFontSize
                        : paragraph.FirstLineIndent;
                }
                else if (!isList)
                {
                    paragraph.FirstLineIndent = 0;
                }
            }
            styledParagraphs++;

            for (int r = 0; r < paragraph.Inlines.Count; r++)
            {
                TextRun run = paragraph.Inlines[r] as TextRun;
                if (run == null) continue;
                TextFormat format = run.Format;
                if (format == null)
                {
                    format = new TextFormat();
                    run.Format = format;
                }
                if (!options.ConvertManualFormatting) continue;
                format.AsciiFontFamily = ascii;
                format.EastAsiaFontFamily = eastAsia;
                format.FontFamily = eastAsia;
                format.FontSize = fontSize;
                if (isHeading) format.Bold = true;
                styledRuns++;
            }
        }

        // ---------- 表格整理 ----------

        private static void NormalizeTables(Section section, WordNormalizeOptions options, WordNormalizeResult result, string location)
        {
            if (!options.NormalizeTables) return;
            for (int i = 0; i < section.Blocks.Count; i++)
            {
                Table table = section.Blocks[i] as Table;
                if (table == null) continue;
                string tableLocation = location + " 表格" + (i + 1).ToString(CultureInfo.InvariantCulture);

                if (options.UnifyTableAlignment) table.Alignment = options.TableAlignment;
                if (options.AutofitTables) table.Autofit = true;
                if (options.RepeatTableHeaderRows && table.Rows.Count > 0)
                {
                    for (int r = 0; r < table.Rows.Count; r++) table.Rows[r].IsHeader = r == 0;
                    table.RepeatHeaderRow = true;
                }

                if (options.RemoveEmptyTableRows) RemoveEmptyTableRows(table, result, tableLocation);
                if (options.RemoveEmptyTableColumns) RemoveEmptyTableColumns(table, result, tableLocation);
                if (table.Rows.Count == 0)
                {
                    section.Blocks.RemoveAt(i);
                    i--;
                    result.AddChange(CategoryTable, "删除空表格", tableLocation);
                }
            }
        }

        private static void RemoveEmptyTableRows(Table table, WordNormalizeResult result, string tableLocation)
        {
            for (int r = table.Rows.Count - 1; r >= 0; r--)
            {
                TableRow row = table.Rows[r];
                bool empty = true;
                for (int c = 0; c < row.Cells.Count && empty; c++)
                {
                    empty = IsCellEmpty(row.Cells[c]);
                }
                if (!empty || row.Cells.Count == 0) continue;
                table.Rows.RemoveAt(r);
                result.AddChange(CategoryTable, "删除空行", tableLocation + " 第" + (r + 1).ToString(CultureInfo.InvariantCulture) + "行");
            }
        }

        private static void RemoveEmptyTableColumns(Table table, WordNormalizeResult result, string tableLocation)
        {
            int columns = int.MaxValue;
            for (int r = 0; r < table.Rows.Count; r++)
            {
                if (table.Rows[r].Cells.Count < columns) columns = table.Rows[r].Cells.Count;
            }
            if (columns == int.MaxValue || columns == 0) return;

            for (int c = columns - 1; c >= 0; c--)
            {
                bool empty = true;
                for (int r = 0; r < table.Rows.Count && empty; r++)
                {
                    TableCell cell = table.Rows[r].Cells[c];
                    empty = cell.ColSpan == 1 && IsCellEmpty(cell);
                }
                if (!empty) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    table.Rows[r].Cells.RemoveAt(c);
                }
                if (c < table.Columns.Count) table.Columns.RemoveAt(c);
                result.AddChange(CategoryTable, "删除空列", tableLocation + " 第" + CellAddress.ColumnName(c + 1) + "列");
            }
        }

        private static bool IsCellEmpty(TableCell cell)
        {
            for (int b = 0; b < cell.Blocks.Count; b++)
            {
                Paragraph paragraph = cell.Blocks[b] as Paragraph;
                if (paragraph != null)
                {
                    if (paragraph.GetText().Trim().Length > 0 || WordDocumentMetrics.HasImage(paragraph)) return false;
                    continue;
                }
                return false;
            }
            return true;
        }

        // ---------- 图片整理 ----------

        private static void NormalizeImages(Section section, List<Block> blocks, WordNormalizeOptions options, WordNormalizeResult result, string location)
        {
            if (!options.NormalizeImages) return;
            double contentWidth = Math.Max(36, section.PageSetup.Width - section.PageSetup.MarginLeft - section.PageSetup.MarginRight);
            double maxWidth = options.MaxImageWidth > 0 ? options.MaxImageWidth : contentWidth;

            for (int i = blocks.Count - 1; i >= 0; i--)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph == null) continue;
                bool removed = false;
                for (int r = paragraph.Inlines.Count - 1; r >= 0; r--)
                {
                    ImageRun image = paragraph.Inlines[r] as ImageRun;
                    if (image == null || image.Image == null) continue;

                    int pixelWidth = image.Image.PixelWidth;
                    int pixelHeight = image.Image.PixelHeight;
                    if (options.RemoveTinyImages && pixelWidth <= options.TinyImageMaxPixels && pixelHeight <= options.TinyImageMaxPixels)
                    {
                        paragraph.Inlines.RemoveAt(r);
                        result.AddChange(CategoryImage, "删除无意义小图(" + pixelWidth.ToString(CultureInfo.InvariantCulture) + "x" + pixelHeight.ToString(CultureInfo.InvariantCulture) + ")",
                            location + " 第" + (i + 1).ToString(CultureInfo.InvariantCulture) + "块");
                        removed = true;
                        continue;
                    }

                    double width = image.Width > 0 ? image.Width : image.Image.DefaultWidthPoints;
                    double height = image.Height > 0 ? image.Height : image.Image.DefaultHeightPoints;
                    if (width <= 0 && height <= 0)
                    {
                        width = pixelWidth;
                        height = pixelHeight;
                    }
                    if (width > maxWidth && width > 0)
                    {
                        double scale = maxWidth / width;
                        double newWidth = maxWidth;
                        double newHeight = height * scale;
                        image.Width = newWidth;
                        image.Height = newHeight;
                        result.AddChange(CategoryImage, "图片超宽，缩放为 " + newWidth.ToString("0.#", CultureInfo.InvariantCulture) + "x" + newHeight.ToString("0.#", CultureInfo.InvariantCulture) + "pt",
                            location + " 第" + (i + 1).ToString(CultureInfo.InvariantCulture) + "块");
                    }

                    ReportAbnormalImage(image, pixelWidth, pixelHeight, result, location + " 第" + (i + 1).ToString(CultureInfo.InvariantCulture) + "块");
                }

                if (options.CenterImages && !removed && IsImageOnly(paragraph))
                {
                    paragraph.Alignment = ParagraphAlignment.Center;
                    paragraph.FirstLineIndent = 0;
                }

                if (removed && paragraph.GetText().Trim().Length == 0 && !WordDocumentMetrics.HasImage(paragraph))
                {
                    blocks.RemoveAt(i);
                }
            }
        }

        private static void ReportAbnormalImage(ImageRun image, int pixelWidth, int pixelHeight, WordNormalizeResult result, string location)
        {
            if (pixelWidth <= 0 || pixelHeight <= 0)
            {
                result.AddIssue(WordIssueSeverity.Warning, "图片缺少有效的像素尺寸", location);
                return;
            }
            if (pixelWidth < 16 || pixelHeight < 16)
            {
                result.AddIssue(WordIssueSeverity.Info, "图片分辨率偏低 (" + pixelWidth.ToString(CultureInfo.InvariantCulture) + "x" + pixelHeight.ToString(CultureInfo.InvariantCulture) + " 像素)", location);
            }
            double aspect = Math.Max(pixelWidth, pixelHeight) / (double)Math.Max(1, Math.Min(pixelWidth, pixelHeight));
            if (aspect > 8)
            {
                result.AddIssue(WordIssueSeverity.Info, "图片宽高比例异常 (" + pixelWidth.ToString(CultureInfo.InvariantCulture) + "x" + pixelHeight.ToString(CultureInfo.InvariantCulture) + ")", location);
            }
        }

        private static bool IsImageOnly(Paragraph paragraph)
        {
            bool hasImage = false;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                if (paragraph.Inlines[i] is ImageRun)
                {
                    hasImage = true;
                    continue;
                }
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null && run.Text.Trim().Length > 0) return false;
            }
            return hasImage;
        }

        // ---------- 目录整理 ----------

        private static void BuildTableOfContents(OfficeDocument document, WordNormalizeOptions options, WordNormalizeResult result)
        {
            List<Paragraph> headings = CollectHeadings(document, options);
            if (headings.Count == 0)
            {
                result.AddIssue(WordIssueSeverity.Info, "文档中没有可识别的标题，跳过目录生成", string.Empty);
                return;
            }

            for (int i = 0; i < headings.Count; i++)
            {
                if (string.IsNullOrEmpty(headings[i].BookmarkName))
                {
                    headings[i].BookmarkName = "_Toc" + (9000000 + i).ToString(CultureInfo.InvariantCulture);
                }
            }

            List<Block> tocBlocks = CreateTableOfContentsBlocks(headings, options);
            int removed = RemoveExistingTableOfContents(document, out List<Block> insertTarget, out int insertIndex);
            if (insertTarget == null)
            {
                insertTarget = document.FirstSection.Blocks;
                insertIndex = 0;
            }
            insertTarget.InsertRange(insertIndex, tocBlocks);
            result.AddChange(CategoryToc, "重新生成目录（" + headings.Count.ToString(CultureInfo.InvariantCulture) + " 个标题，"
                + (removed > 0 ? "替换原有目录 " + removed.ToString(CultureInfo.InvariantCulture) + " 段" : "新建目录") + "）", string.Empty);
        }

        private static List<Paragraph> CollectHeadings(OfficeDocument document, WordNormalizeOptions options)
        {
            List<Paragraph> headings = new List<Paragraph>();
            for (int s = 0; s < document.Sections.Count; s++)
            {
                List<Block> blocks = document.Sections[s].Blocks;
                for (int i = 0; i < blocks.Count; i++)
                {
                    Paragraph paragraph = blocks[i] as Paragraph;
                    if (paragraph == null) continue;
                    if (paragraph.HeadingLevel <= 0 || paragraph.HeadingLevel > options.TableOfContentsMaxLevel) continue;
                    if (paragraph.GetText().Trim().Length == 0) continue;
                    if (IsTableOfContentsTitle(paragraph, options)) continue;
                    headings.Add(paragraph);
                }
            }
            return headings;
        }

        private static List<Block> CreateTableOfContentsBlocks(List<Paragraph> headings, WordNormalizeOptions options)
        {
            List<Block> blocks = new List<Block>();

            TextFormat titleFormat = new TextFormat();
            titleFormat.AsciiFontFamily = options.HeadingAsciiFont;
            titleFormat.EastAsiaFontFamily = options.HeadingEastAsiaFont;
            titleFormat.FontFamily = options.HeadingEastAsiaFont;
            titleFormat.FontSize = options.BodyFontSize + 4;
            titleFormat.Bold = true;
            Paragraph title = new Paragraph();
            title.AddText(options.TableOfContentsTitle, titleFormat);
            title.Alignment = ParagraphAlignment.Center;
            title.SpaceAfter = 12;
            title.LineSpacing = options.HeadingLineSpacing;
            title.KeepWithNext = true;
            blocks.Add(title);

            for (int i = 0; i < headings.Count; i++)
            {
                Paragraph heading = headings[i];
                TextFormat format = new TextFormat();
                format.AsciiFontFamily = options.BodyAsciiFont;
                format.EastAsiaFontFamily = options.BodyEastAsiaFont;
                format.FontFamily = options.BodyEastAsiaFont;
                format.FontSize = options.BodyFontSize;
                if (heading.HeadingLevel == 1) format.Bold = true;

                Paragraph entry = new Paragraph();
                entry.AddText(heading.GetText().Trim(), format);
                entry.AddText("\t", format);
                entry.Inlines.Add(new FieldRun("PAGEREF " + heading.BookmarkName + " \\h", "1", format));
                int level = Math.Max(1, heading.HeadingLevel);
                entry.LeftIndent = (level - 1) * 14;
                entry.FirstLineIndent = 0;
                entry.SpaceBefore = 0;
                entry.SpaceAfter = 0;
                entry.LineSpacing = Math.Min(options.BodyLineSpacing, 1.5);
                blocks.Add(entry);
            }

            if (options.AddPageBreakAfterTableOfContents)
            {
                blocks.Add(new PageBreak());
            }
            return blocks;
        }

        private static int RemoveExistingTableOfContents(OfficeDocument document, out List<Block> target, out int index)
        {
            target = null;
            index = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                List<Block> blocks = document.Sections[s].Blocks;
                for (int i = 0; i < blocks.Count; i++)
                {
                    if (!IsTableOfContentsStart(blocks[i])) continue;
                    int start = i;
                    int end = i + 1;
                    while (end < blocks.Count && IsTableOfContentsEntry(blocks[end])) end++;
                    if (end < blocks.Count && blocks[end] is PageBreak) end++;
                    int removed = end - start;
                    blocks.RemoveRange(start, end - start);
                    target = blocks;
                    index = start;
                    return removed;
                }
            }
            return 0;
        }

        private static bool IsTableOfContentsStart(Block block)
        {
            Paragraph paragraph = block as Paragraph;
            if (paragraph == null) return false;
            string text = paragraph.GetText().Trim();
            if (text == "目录" || string.Equals(text, "Contents", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "Table of Contents", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return IsTocParagraph(paragraph) && HasTocField(paragraph);
        }

        private static bool IsTableOfContentsEntry(Block block)
        {
            Paragraph paragraph = block as Paragraph;
            if (paragraph == null) return false;
            if (IsTocParagraph(paragraph)) return true;
            string text = paragraph.GetText().Trim();
            return text.Length == 0;
        }

        private static bool IsTableOfContentsTitle(Paragraph paragraph, WordNormalizeOptions options)
        {
            string text = paragraph.GetText().Trim();
            if (text.Length == 0) return false;
            string title = string.IsNullOrEmpty(options.TableOfContentsTitle) ? "目录" : options.TableOfContentsTitle;
            return string.Equals(text, title, StringComparison.Ordinal)
                || string.Equals(text, "目录", StringComparison.Ordinal)
                || string.Equals(text, "Contents", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "Table of Contents", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTocParagraph(Paragraph paragraph)
        {
            if (!string.IsNullOrEmpty(paragraph.StyleName)
                && paragraph.StyleName.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)) return true;
            return HasTocField(paragraph);
        }

        private static bool HasTocField(Paragraph paragraph)
        {
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field == null) continue;
                string instruction = field.Instruction ?? string.Empty;
                if (instruction.StartsWith("TOC", StringComparison.OrdinalIgnoreCase)
                    || instruction.StartsWith("PAGEREF", StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static string Shorten(string text)
        {
            if (text.Length <= 24) return text;
            return text.Substring(0, 24) + "…";
        }
    }
}