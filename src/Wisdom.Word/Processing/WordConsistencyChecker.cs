using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>校验文档格式一致性（通常在规范化之后执行）。Validates format consistency of a document (usually after normalization).</summary>
    public static class WordConsistencyChecker
    {
        /// <summary>
        /// 按规范化选项逐项扫描文档并汇总一致性问题。
        /// Scans the document against the normalization options and collects the consistency issues found.
        /// </summary>
        /// <param name="document">待校验文档。The document to validate.</param>
        /// <param name="options">规范化选项，为 null 时使用默认。The options to validate against; defaults are used when null.</param>
        /// <returns>问题清单，末尾附统计信息。The issue list, ending with a summary entry.</returns>
        public static List<WordIssue> Check(OfficeDocument document, WordNormalizeOptions options)
        {
            List<WordIssue> issues = new List<WordIssue>();
            if (document == null) return issues;
            if (options == null) options = new WordNormalizeOptions();

            int emptyParagraphs = 0;
            int fontMismatch = 0;
            int sizeMismatch = 0;
            int spacingMismatch = 0;
            int halfWidthEnding = 0;
            int controlCharacters = 0;
            int previousHeadingLevel = 0;
            int headingSkip = 0;
            int emptyTableRows = 0;
            int oversizedImages = 0;
            int tableCount = 0;
            int headingCount = 0;
            int tocEntries = 0;

            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                double contentWidth = Math.Max(36, section.PageSetup.Width - section.PageSetup.MarginLeft - section.PageSetup.MarginRight);
                CheckBlocks(section.Blocks, options, issues, contentWidth,
                    ref emptyParagraphs, ref fontMismatch, ref sizeMismatch, ref spacingMismatch,
                    ref halfWidthEnding, ref controlCharacters, ref previousHeadingLevel, ref headingSkip,
                    ref emptyTableRows, ref oversizedImages, ref tableCount, ref headingCount, ref tocEntries, true);
            }

            if (emptyParagraphs > 0) issues.Add(new WordIssue(WordIssueSeverity.Warning, "仍存在 " + emptyParagraphs.ToString(CultureInfo.InvariantCulture) + " 个空段落", string.Empty));
            if (fontMismatch > 0) issues.Add(new WordIssue(WordIssueSeverity.Warning, "正文存在 " + fontMismatch.ToString(CultureInfo.InvariantCulture) + " 处非规范字体", string.Empty));
            if (sizeMismatch > 0) issues.Add(new WordIssue(WordIssueSeverity.Warning, "正文存在 " + sizeMismatch.ToString(CultureInfo.InvariantCulture) + " 处非规范字号", string.Empty));
            if (spacingMismatch > 0) issues.Add(new WordIssue(WordIssueSeverity.Info, "存在 " + spacingMismatch.ToString(CultureInfo.InvariantCulture) + " 个段落行距/段距与规范不一致", string.Empty));
            if (halfWidthEnding > 0) issues.Add(new WordIssue(WordIssueSeverity.Info, "存在 " + halfWidthEnding.ToString(CultureInfo.InvariantCulture) + " 个段落以半角标点结尾", string.Empty));
            if (controlCharacters > 0) issues.Add(new WordIssue(WordIssueSeverity.Warning, "正文包含 " + controlCharacters.ToString(CultureInfo.InvariantCulture) + " 个不可见控制字符", string.Empty));
            if (headingSkip > 0) issues.Add(new WordIssue(WordIssueSeverity.Warning, "存在 " + headingSkip.ToString(CultureInfo.InvariantCulture) + " 处标题层级跳跃", string.Empty));
            if (emptyTableRows > 0) issues.Add(new WordIssue(WordIssueSeverity.Info, "表格中仍存在 " + emptyTableRows.ToString(CultureInfo.InvariantCulture) + " 个空行", string.Empty));
            if (oversizedImages > 0) issues.Add(new WordIssue(WordIssueSeverity.Warning, "存在 " + oversizedImages.ToString(CultureInfo.InvariantCulture) + " 张超出正文宽度的图片", string.Empty));

            issues.Add(new WordIssue(WordIssueSeverity.Info, "统计: 标题 " + headingCount.ToString(CultureInfo.InvariantCulture)
                + " 个, 表格 " + tableCount.ToString(CultureInfo.InvariantCulture)
                + " 个, 目录条目 " + tocEntries.ToString(CultureInfo.InvariantCulture) + " 条", string.Empty));
            if (options.RegenerateTableOfContents && headingCount == 0)
            {
                issues.Add(new WordIssue(WordIssueSeverity.Info, "未识别到任何标题，目录无法生成", string.Empty));
            }
            return issues;
        }

        private static void CheckBlocks(
            List<Block> blocks,
            WordNormalizeOptions options,
            List<WordIssue> issues,
            double contentWidth,
            ref int emptyParagraphs,
            ref int fontMismatch,
            ref int sizeMismatch,
            ref int spacingMismatch,
            ref int halfWidthEnding,
            ref int controlCharacters,
            ref int previousHeadingLevel,
            ref int headingSkip,
            ref int emptyTableRows,
            ref int oversizedImages,
            ref int tableCount,
            ref int headingCount,
            ref int tocEntries,
            bool isBody)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    CheckParagraph(paragraph, options, contentWidth,
                        ref emptyParagraphs, ref fontMismatch, ref sizeMismatch, ref spacingMismatch,
                        ref halfWidthEnding, ref controlCharacters, ref previousHeadingLevel, ref headingSkip,
                        ref oversizedImages, ref headingCount, ref tocEntries, isBody);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                tableCount++;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    TableRow row = table.Rows[r];
                    bool rowEmpty = row.Cells.Count > 0;
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        TableCell cell = row.Cells[c];
                        bool cellEmpty = true;
                        for (int b = 0; b < cell.Blocks.Count; b++)
                        {
                            Paragraph cellParagraph = cell.Blocks[b] as Paragraph;
                            if (cellParagraph != null && cellParagraph.GetText().Trim().Length > 0) cellEmpty = false;
                            if (cellParagraph != null && WordDocumentMetrics.HasImage(cellParagraph)) cellEmpty = false;
                        }
                        if (!cellEmpty) rowEmpty = false;
                        CheckBlocks(cell.Blocks, options, issues, contentWidth,
                            ref emptyParagraphs, ref fontMismatch, ref sizeMismatch, ref spacingMismatch,
                            ref halfWidthEnding, ref controlCharacters, ref previousHeadingLevel, ref headingSkip,
                            ref emptyTableRows, ref oversizedImages, ref tableCount, ref headingCount, ref tocEntries, false);
                    }
                    if (rowEmpty) emptyTableRows++;
                }
            }
        }

        private static void CheckParagraph(
            Paragraph paragraph,
            WordNormalizeOptions options,
            double contentWidth,
            ref int emptyParagraphs,
            ref int fontMismatch,
            ref int sizeMismatch,
            ref int spacingMismatch,
            ref int halfWidthEnding,
            ref int controlCharacters,
            ref int previousHeadingLevel,
            ref int headingSkip,
            ref int oversizedImages,
            ref int headingCount,
            ref int tocEntries,
            bool isBody)
        {
            bool hasImage = WordDocumentMetrics.HasImage(paragraph);
            bool isToc = HasPageRef(paragraph);
            if (isToc) tocEntries++;
            string text = paragraph.GetText();
            if (isBody && text.Trim().Length == 0 && !hasImage && !isToc) emptyParagraphs++;

            if (paragraph.HeadingLevel > 0)
            {
                headingCount++;
                if (previousHeadingLevel > 0 && paragraph.HeadingLevel > previousHeadingLevel + 1) headingSkip++;
                previousHeadingLevel = paragraph.HeadingLevel;
            }

            bool isHeading = paragraph.HeadingLevel > 0;
            if (!isHeading && !isToc && isBody)
            {
                if (Math.Abs(paragraph.LineSpacing - options.BodyLineSpacing) > 0.01
                    && !(paragraph.Alignment == ParagraphAlignment.Center && hasImage)) spacingMismatch++;
                for (int r = 0; r < paragraph.Inlines.Count; r++)
                {
                    TextRun run = paragraph.Inlines[r] as TextRun;
                    if (run == null || run.Format == null) continue;
                    string eastAsia = run.Format.GetEastAsiaFontFamily();
                    if (!string.IsNullOrEmpty(eastAsia) && !string.Equals(eastAsia, options.BodyEastAsiaFont, StringComparison.Ordinal)) fontMismatch++;
                    if (run.Format.FontSize > 0 && Math.Abs(run.Format.FontSize - options.BodyFontSize) > 0.01) sizeMismatch++;
                    controlCharacters += CountControlCharacters(run.Text);
                }
                if (IsMostlyCjk(text) && HalfWidthEnding(text)) halfWidthEnding++;
            }

            for (int r = 0; r < paragraph.Inlines.Count; r++)
            {
                ImageRun image = paragraph.Inlines[r] as ImageRun;
                if (image == null || image.Image == null) continue;
                double width = image.Width > 0 ? image.Width : image.Image.DefaultWidthPoints;
                if (width > contentWidth + 1) oversizedImages++;
            }
        }

        private static int CountControlCharacters(string text)
        {
            int count = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c < 32 && c != '\n' && c != '\t' && c != '\r') count++;
            }
            return count;
        }

        private static bool HasPageRef(Paragraph paragraph)
        {
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field == null) continue;
                if ((field.Instruction ?? string.Empty).StartsWith("PAGEREF", StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static bool IsMostlyCjk(string text)
        {
            int cjk = 0;
            int latin = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c >= 0x4E00 && c <= 0x9FFF) cjk++;
                else if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')) latin++;
            }
            return cjk > 0 && cjk >= latin;
        }

        private static bool HalfWidthEnding(string text)
        {
            string trimmed = text.TrimEnd();
            if (trimmed.Length == 0) return false;
            char last = trimmed[trimmed.Length - 1];
            return last == '.' || last == ',' || last == ';' || last == ':' || last == '!' || last == '?';
        }
    }
}