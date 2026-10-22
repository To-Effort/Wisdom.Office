using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>Word 整理/校验问题的严重级别。Severity level of a Word clean-up or validation issue.</summary>
    public enum WordIssueSeverity
    {
        /// <summary>提示级信息。Informational message.</summary>
        Info = 0,
        /// <summary>警告级问题。Warning-level problem.</summary>
        Warning = 1,
        /// <summary>错误级问题。Error-level problem.</summary>
        Error = 2
    }

    /// <summary>单条已应用到文档的规范化动作。A single normalization action that was applied to the document.</summary>
    public sealed class WordChange
    {
        /// <summary>变更类别。The change category name.</summary>
        public string Category { get; private set; }
        /// <summary>变更说明文本。The human-readable description of the change.</summary>
        public string Description { get; private set; }
        /// <summary>变更所在位置描述。The location description of the change.</summary>
        public string Location { get; private set; }

        /// <summary>
        /// 创建一条变更记录。
        /// Creates one change record.
        /// </summary>
        /// <param name="category">变更类别。The change category.</param>
        /// <param name="description">变更说明。The change description.</param>
        /// <param name="location">位置描述，可为空串。The location text; may be an empty string.</param>
        public WordChange(string category, string description, string location)
        {
            Category = category ?? string.Empty;
            Description = description ?? string.Empty;
            Location = location ?? string.Empty;
        }

        /// <summary>以 "[类别] 说明 (位置)" 格式输出该变更。Renders this change as "[category] description (location)" text.</summary>
        public override string ToString()
        {
            return "[" + Category + "] " + Description + (Location.Length == 0 ? string.Empty : " (" + Location + ")");
        }
    }

    /// <summary>规范化或校验过程中发现的单个问题。A problem detected while normalizing or validating the document.</summary>
    public sealed class WordIssue
    {
        /// <summary>问题严重级别。The issue severity.</summary>
        public WordIssueSeverity Severity { get; private set; }
        /// <summary>问题描述文本。The issue message text.</summary>
        public string Message { get; private set; }
        /// <summary>问题所在位置描述。The location description of the issue.</summary>
        public string Location { get; private set; }

        /// <summary>
        /// 创建一条问题记录。
        /// Creates one issue record.
        /// </summary>
        /// <param name="severity">问题严重级别。The issue severity.</param>
        /// <param name="message">问题描述。The issue message.</param>
        /// <param name="location">位置描述，可为空串。The location text; may be an empty string.</param>
        public WordIssue(WordIssueSeverity severity, string message, string location)
        {
            Severity = severity;
            Message = message ?? string.Empty;
            Location = location ?? string.Empty;
        }

        /// <summary>以 "级别: 消息 (位置)" 格式输出该问题。Renders this issue as "SEVERITY: message (location)" text.</summary>
        public override string ToString()
        {
            return Severity.ToString().ToUpperInvariant() + ": " + Message + (Location.Length == 0 ? string.Empty : " (" + Location + ")");
        }
    }

    /// <summary>文档统计指标快照，用于整理前后对比。A snapshot of document statistics used to compare the document before and after clean-up.</summary>
    public sealed class WordDocumentMetrics
    {
        /// <summary>节数量。Number of sections.</summary>
        public int SectionCount;
        /// <summary>段落总数。Total paragraph count.</summary>
        public int ParagraphCount;
        /// <summary>空段落数量。Count of empty paragraphs.</summary>
        public int EmptyParagraphCount;
        /// <summary>字符总数。Total character count.</summary>
        public int CharacterCount;
        /// <summary>表格总数。Total table count.</summary>
        public int TableCount;
        /// <summary>表格行总数。Total table row count.</summary>
        public int TableRowCount;
        /// <summary>图片总数。Total image count.</summary>
        public int ImageCount;
        /// <summary>各级标题数量，索引 1 至 9。Heading counts per level, indexed 1 through 9.</summary>
        public int[] HeadingCounts = new int[10];
        /// <summary>字体名到使用次数的映射。Font family name to usage count.</summary>
        public Dictionary<string, int> FontUsage = new Dictionary<string, int>();
        /// <summary>字号到使用次数的映射。Font size to usage count.</summary>
        public Dictionary<string, int> FontSizeUsage = new Dictionary<string, int>();

        /// <summary>所有级别标题的总数。The total number of headings across all levels.</summary>
        public int HeadingCount
        {
            get
            {
                int total = 0;
                for (int i = 1; i < HeadingCounts.Length; i++) total += HeadingCounts[i];
                return total;
            }
        }

        /// <summary>
        /// 统计文档各项指标。
        /// Measures the statistics of a document.
        /// </summary>
        /// <param name="document">待统计文档，为 null 时返回空指标。The document to measure; null yields an empty snapshot.</param>
        /// <returns>指标快照。The statistics snapshot.</returns>
        public static WordDocumentMetrics Measure(OfficeDocument document)
        {
            WordDocumentMetrics metrics = new WordDocumentMetrics();
            if (document == null) return metrics;
            metrics.SectionCount = document.Sections.Count;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                MeasureBlocks(document.Sections[s].Blocks, metrics);
                MeasureBlocks(document.Sections[s].Header.Blocks, metrics);
                MeasureBlocks(document.Sections[s].Footer.Blocks, metrics);
            }
            return metrics;
        }

        private static void MeasureBlocks(List<Block> blocks, WordDocumentMetrics metrics)
        {
            if (blocks == null) return;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    metrics.ParagraphCount++;
                    string text = paragraph.GetText();
                    metrics.CharacterCount += text.Length;
                    if (text.Trim().Length == 0 && !HasImage(paragraph)) metrics.EmptyParagraphCount++;
                    if (paragraph.HeadingLevel > 0 && paragraph.HeadingLevel < metrics.HeadingCounts.Length)
                    {
                        metrics.HeadingCounts[paragraph.HeadingLevel]++;
                    }
                    for (int r = 0; r < paragraph.Inlines.Count; r++)
                    {
                        TextRun run = paragraph.Inlines[r] as TextRun;
                        if (run != null)
                        {
                            RegisterFormat(metrics, run.Format);
                            continue;
                        }
                        if (paragraph.Inlines[r] is ImageRun) metrics.ImageCount++;
                    }
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    metrics.TableCount++;
                    metrics.TableRowCount += table.Rows.Count;
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                        {
                            MeasureBlocks(table.Rows[r].Cells[c].Blocks, metrics);
                        }
                    }
                }
            }
        }

        /// <summary>判断段落内是否包含图片运行。Determines whether the paragraph contains an image run.</summary>
        internal static bool HasImage(Paragraph paragraph)
        {
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                if (paragraph.Inlines[i] is ImageRun) return true;
            }
            return false;
        }

        private static void RegisterFormat(WordDocumentMetrics metrics, TextFormat format)
        {
            if (format == null) return;
            string eastAsia = format.GetEastAsiaFontFamily();
            string ascii = format.GetAsciiFontFamily();
            string font = string.IsNullOrEmpty(eastAsia) ? ascii : eastAsia;
            if (!string.IsNullOrEmpty(font))
            {
                int count;
                metrics.FontUsage.TryGetValue(font, out count);
                metrics.FontUsage[font] = count + 1;
            }
            if (format.FontSize > 0)
            {
                string size = format.FontSize.ToString("0.#", CultureInfo.InvariantCulture);
                int count;
                metrics.FontSizeUsage.TryGetValue(size, out count);
                metrics.FontSizeUsage[size] = count + 1;
            }
        }
    }

    /// <summary>规范化管道的执行结果：变更清单、问题清单、前后指标与差异文本。The normalization pipeline result: change list, issue list, before/after metrics and diff text.</summary>
    public sealed class WordNormalizeResult
    {
        /// <summary>被整理的文档对象。The document that was normalized.</summary>
        public OfficeDocument Document { get; set; }
        /// <summary>已应用的变更列表。The list of applied changes.</summary>
        public List<WordChange> Changes { get; private set; }
        /// <summary>发现的问题列表。The list of detected issues.</summary>
        public List<WordIssue> Issues { get; private set; }
        /// <summary>整理前的指标快照。Metrics measured before normalization.</summary>
        public WordDocumentMetrics Before { get; set; }
        /// <summary>整理后的指标快照。Metrics measured after normalization.</summary>
        public WordDocumentMetrics After { get; set; }
        /// <summary>整理前后的内容差异文本。The before/after content diff text.</summary>
        public string Diff { get; set; }

        /// <summary>初始化变更与问题列表。Initializes the change and issue lists.</summary>
        public WordNormalizeResult()
        {
            Changes = new List<WordChange>();
            Issues = new List<WordIssue>();
        }

        /// <summary>
        /// 追加一条变更记录。
        /// Appends one change record.
        /// </summary>
        /// <param name="category">变更类别。The change category.</param>
        /// <param name="description">变更说明。The change description.</param>
        /// <param name="location">位置描述。The location text.</param>
        public void AddChange(string category, string description, string location)
        {
            Changes.Add(new WordChange(category, description, location));
        }

        /// <summary>
        /// 追加一条问题记录。
        /// Appends one issue record.
        /// </summary>
        /// <param name="severity">问题严重级别。The issue severity.</param>
        /// <param name="message">问题描述。The issue message.</param>
        /// <param name="location">位置描述。The location text.</param>
        public void AddIssue(WordIssueSeverity severity, string message, string location)
        {
            Issues.Add(new WordIssue(severity, message, location));
        }

        /// <summary>
        /// 生成完整的中文整理报告文本。
        /// Builds the full Chinese clean-up report text.
        /// </summary>
        /// <returns>报告全文。The full report text.</returns>
        public string ToReport()
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("=== Word 整理报告 ===");
            builder.AppendLine("变更条目: " + Changes.Count.ToString(CultureInfo.InvariantCulture));
            Dictionary<string, int> categories = new Dictionary<string, int>();
            for (int i = 0; i < Changes.Count; i++)
            {
                int count;
                categories.TryGetValue(Changes[i].Category, out count);
                categories[Changes[i].Category] = count + 1;
            }
            foreach (KeyValuePair<string, int> pair in categories)
            {
                builder.AppendLine("  - " + pair.Key + ": " + pair.Value.ToString(CultureInfo.InvariantCulture));
            }
            builder.AppendLine();
            builder.AppendLine("--- 变更明细 ---");
            for (int i = 0; i < Changes.Count; i++)
            {
                builder.AppendLine("  " + Changes[i].ToString());
            }
            if (Issues.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("--- 问题清单 ---");
                for (int i = 0; i < Issues.Count; i++)
                {
                    builder.AppendLine("  " + Issues[i].ToString());
                }
            }
            builder.AppendLine();
            builder.AppendLine("--- 指标对比（整理前 -> 整理后）---");
            AppendMetrics(builder, Before, After);
            if (!string.IsNullOrEmpty(Diff))
            {
                builder.AppendLine();
                builder.AppendLine(Diff);
            }
            return builder.ToString();
        }

        private static void AppendMetrics(StringBuilder builder, WordDocumentMetrics before, WordDocumentMetrics after)
        {
            if (before == null || after == null) return;
            builder.AppendLine("  段落数: " + before.ParagraphCount.ToString(CultureInfo.InvariantCulture) + " -> " + after.ParagraphCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("  空段落: " + before.EmptyParagraphCount.ToString(CultureInfo.InvariantCulture) + " -> " + after.EmptyParagraphCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("  字符数: " + before.CharacterCount.ToString(CultureInfo.InvariantCulture) + " -> " + after.CharacterCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("  标题数: " + before.HeadingCount.ToString(CultureInfo.InvariantCulture) + " -> " + after.HeadingCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("  表格数: " + before.TableCount.ToString(CultureInfo.InvariantCulture) + " -> " + after.TableCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("  图片数: " + before.ImageCount.ToString(CultureInfo.InvariantCulture) + " -> " + after.ImageCount.ToString(CultureInfo.InvariantCulture));
            builder.Append("  字体使用: ");
            AppendDistribution(builder, before.FontUsage, after.FontUsage);
            builder.Append("  字号使用: ");
            AppendDistribution(builder, before.FontSizeUsage, after.FontSizeUsage);
        }

        private static void AppendDistribution(StringBuilder builder, Dictionary<string, int> before, Dictionary<string, int> after)
        {
            builder.AppendLine();
            builder.AppendLine("    整理前: " + FormatDistribution(before));
            builder.AppendLine("    整理后: " + FormatDistribution(after));
        }

        private static string FormatDistribution(Dictionary<string, int> values)
        {
            StringBuilder builder = new StringBuilder();
            foreach (KeyValuePair<string, int> pair in values)
            {
                if (builder.Length > 0) builder.Append(", ");
                builder.Append(pair.Key).Append('x').Append(pair.Value.ToString(CultureInfo.InvariantCulture));
            }
            return builder.Length == 0 ? "(无)" : builder.ToString();
        }
    }
}