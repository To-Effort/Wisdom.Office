using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>按行对比文档规范化前后的内容差异。Line-based comparison of a document before and after normalization.</summary>
    public sealed class WordDocumentDiff
    {
        private const int MaxListedLines = 200;

        /// <summary>删除行列表（最多列出前 200 条）。The removed-line list (up to the first 200 entries).</summary>
        public List<string> RemovedLines { get; private set; }
        /// <summary>新增行列表（最多列出前 200 条）。The added-line list (up to the first 200 entries).</summary>
        public List<string> AddedLines { get; private set; }
        /// <summary>删除的总行数。The total number of removed lines.</summary>
        public int RemovedCount { get; private set; }
        /// <summary>新增的总行数。The total number of added lines.</summary>
        public int AddedCount { get; private set; }
        /// <summary>未变化的总行数。The total number of unchanged lines.</summary>
        public int UnchangedCount { get; private set; }

        /// <summary>初始化空的删除/新增行列表。Initializes the empty removed and added line lists.</summary>
        public WordDocumentDiff()
        {
            RemovedLines = new List<string>();
            AddedLines = new List<string>();
        }

        /// <summary>是否存在删除行或新增行。Indicates whether any line was removed or added.</summary>
        public bool HasDifferences { get { return RemovedCount > 0 || AddedCount > 0; } }

        /// <summary>
        /// 按行统计两份文档的差异。
        /// Compares two documents line by line and counts the removals and additions.
        /// </summary>
        /// <param name="before">整理前的文档。The document before normalization.</param>
        /// <param name="after">整理后的文档。The document after normalization.</param>
        /// <returns>差异结果。The diff result.</returns>
        public static WordDocumentDiff Compare(OfficeDocument before, OfficeDocument after)
        {
            WordDocumentDiff diff = new WordDocumentDiff();
            List<string> beforeLines = CollectLines(before);
            List<string> afterLines = CollectLines(after);

            Dictionary<string, int> beforeCounts = CountLines(beforeLines);
            Dictionary<string, int> afterCounts = CountLines(afterLines);

            foreach (KeyValuePair<string, int> pair in beforeCounts)
            {
                int remaining;
                afterCounts.TryGetValue(pair.Key, out remaining);
                int removed = pair.Value - remaining;
                if (removed <= 0)
                {
                    diff.UnchangedCount += pair.Value;
                    continue;
                }
                diff.RemovedCount += removed;
                if (diff.RemovedLines.Count < MaxListedLines) diff.RemovedLines.Add(pair.Key);
            }

            foreach (KeyValuePair<string, int> pair in afterCounts)
            {
                int original;
                beforeCounts.TryGetValue(pair.Key, out original);
                int added = pair.Value - original;
                if (added <= 0) continue;
                diff.AddedCount += added;
                if (diff.AddedLines.Count < MaxListedLines) diff.AddedLines.Add(pair.Key);
            }

            return diff;
        }

        /// <summary>
        /// 生成中文差异对比文本。
        /// Builds the Chinese diff comparison text.
        /// </summary>
        /// <returns>对比全文。The full diff text.</returns>
        public string ToText()
        {
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("=== 整理前/后差异对比 ===");
            builder.AppendLine("未变化行: " + UnchangedCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("删除行: " + RemovedCount.ToString(CultureInfo.InvariantCulture));
            builder.AppendLine("新增行: " + AddedCount.ToString(CultureInfo.InvariantCulture));
            if (RemovedLines.Count > 0)
            {
                builder.AppendLine("--- 删除内容（前 200 条）---");
                for (int i = 0; i < RemovedLines.Count; i++)
                {
                    builder.AppendLine("  - " + Shorten(RemovedLines[i]));
                }
            }
            if (AddedLines.Count > 0)
            {
                builder.AppendLine("--- 新增内容（前 200 条）---");
                for (int i = 0; i < AddedLines.Count; i++)
                {
                    builder.AppendLine("  + " + Shorten(AddedLines[i]));
                }
            }
            return builder.ToString();
        }

        private static string Shorten(string line)
        {
            if (line.Length <= 60) return line;
            return line.Substring(0, 60) + "…";
        }

        private static Dictionary<string, int> CountLines(List<string> lines)
        {
            Dictionary<string, int> counts = new Dictionary<string, int>();
            for (int i = 0; i < lines.Count; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0) continue;
                int count;
                counts.TryGetValue(line, out count);
                counts[line] = count + 1;
            }
            return counts;
        }

        private static List<string> CollectLines(OfficeDocument document)
        {
            List<string> lines = new List<string>();
            if (document == null) return lines;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                CollectBlocks(document.Sections[s].Blocks, lines);
            }
            return lines;
        }

        private static void CollectBlocks(List<Block> blocks, List<string> lines)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    string text = paragraph.GetText();
                    if (text.Trim().Length > 0) lines.Add(text);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    StringBuilder row = new StringBuilder();
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        if (c > 0) row.Append('\t');
                        TableCell cell = table.Rows[r].Cells[c];
                        for (int b = 0; b < cell.Blocks.Count; b++)
                        {
                            Paragraph cellParagraph = cell.Blocks[b] as Paragraph;
                            if (cellParagraph != null) row.Append(cellParagraph.GetText());
                        }
                    }
                    if (row.Length > 0) lines.Add(row.ToString());
                }
            }
        }
    }

    /// <summary>文档深拷贝工具，用于规范化前保留原始内容。Deep-copy helper that snapshots the original document before normalization.</summary>
    internal static class DocumentCloner
    {
        /// <summary>
        /// 深拷贝整个文档模型。
        /// Creates a deep copy of the whole document model.
        /// </summary>
        /// <param name="source">源文档。The source document.</param>
        /// <returns>拷贝出的文档。The cloned document.</returns>
        public static OfficeDocument Clone(OfficeDocument source)
        {
            OfficeDocument copy = new OfficeDocument();
            copy.Metadata = CloneMetadata(source.Metadata);
            copy.DefaultTextFormat = source.DefaultTextFormat == null ? new TextFormat() : source.DefaultTextFormat.Clone();
            for (int s = 0; s < source.Sections.Count; s++)
            {
                Section section = copy.AddSection(source.Sections[s].PageSetup == null ? new PageSetup() : source.Sections[s].PageSetup.Clone());
                CloneBlocks(source.Sections[s].Blocks, section.Blocks);
                CloneHeaderFooter(source.Sections[s].Header, section.Header);
                CloneHeaderFooter(source.Sections[s].Footer, section.Footer);
            }
            return copy;
        }

        private static DocumentMetadata CloneMetadata(DocumentMetadata source)
        {
            DocumentMetadata copy = new DocumentMetadata();
            if (source == null) return copy;
            copy.Title = source.Title;
            copy.Author = source.Author;
            copy.Subject = source.Subject;
            copy.Keywords = source.Keywords;
            copy.Creator = source.Creator;
            copy.Created = source.Created;
            copy.Modified = source.Modified;
            if (source.CustomProperties != null)
            {
                foreach (KeyValuePair<string, object> pair in source.CustomProperties)
                {
                    copy.CustomProperties.Set(pair.Key, pair.Value);
                }
            }
            return copy;
        }

        private static void CloneHeaderFooter(HeaderFooter source, HeaderFooter copy)
        {
            if (source == null) return;
            CloneBlocks(source.Blocks, copy.Blocks);
        }

        internal static void CloneBlocks(List<Block> source, List<Block> copy)
        {
            for (int i = 0; i < source.Count; i++)
            {
                Paragraph paragraph = source[i] as Paragraph;
                if (paragraph != null)
                {
                    copy.Add(paragraph.Clone());
                    continue;
                }
                Table table = source[i] as Table;
                if (table != null)
                {
                    copy.Add(CloneTable(table));
                    continue;
                }
                if (source[i] is PageBreak) copy.Add(new PageBreak());
            }
        }

        internal static Table CloneTable(Table source)
        {
            Table copy = CreateTableShell(source);
            for (int r = 0; r < source.Rows.Count; r++)
            {
                copy.Rows.Add(CloneRow(source.Rows[r]));
            }
            return copy;
        }

        /// <summary>
        /// 创建只含表格格式（列、边框、内边距、对齐）的空表格，供区域展开逐项装箱。
        /// Creates an empty table carrying only the table formatting (columns, borders, padding,
        /// alignment), used by region expansion to stage cloned rows per data item.
        /// </summary>
        /// <param name="source">格式来源表格。The table supplying the formatting.</param>
        /// <returns>不含行的表格。A table without rows.</returns>
        internal static Table CreateTableShell(Table source)
        {
            Table copy = new Table();
            copy.CellPadding = source.CellPadding;
            copy.RepeatHeaderRow = source.RepeatHeaderRow;
            copy.Alignment = source.Alignment;
            copy.Autofit = source.Autofit;
            copy.Borders = CloneBorders(source.Borders);
            for (int c = 0; c < source.Columns.Count; c++) copy.Columns.Add(new TableColumn(source.Columns[c].Width));
            return copy;
        }

        /// <summary>
        /// 深拷贝表格行。
        /// Deep-copies a table row.
        /// </summary>
        /// <param name="source">源行。The source row.</param>
        /// <returns>行副本。The copied row.</returns>
        internal static TableRow CloneRow(TableRow source)
        {
            TableRow copy = new TableRow();
            copy.IsHeader = source.IsHeader;
            copy.Height = source.Height;
            for (int c = 0; c < source.Cells.Count; c++)
            {
                copy.Cells.Add(CloneCell(source.Cells[c]));
            }
            return copy;
        }

        /// <summary>
        /// 深拷贝表格单元格及其块内容。
        /// Deep-copies a table cell and its block content.
        /// </summary>
        /// <param name="source">源单元格。The source cell.</param>
        /// <returns>单元格副本。The copied cell.</returns>
        internal static TableCell CloneCell(TableCell source)
        {
            TableCell copy = new TableCell();
            copy.ColSpan = source.ColSpan;
            copy.RowSpan = source.RowSpan;
            copy.BackgroundColor = source.BackgroundColor;
            copy.VerticalAlignment = source.VerticalAlignment;
            copy.Borders = CloneCellBorders(source.Borders);
            CloneBlocks(source.Blocks, copy.Blocks);
            return copy;
        }

        /// <summary>
        /// 深拷贝单个块元素。
        /// Deep-copies a single block element.
        /// </summary>
        /// <param name="source">源块。The source block.</param>
        /// <returns>块副本。The copied block.</returns>
        internal static Block CloneBlock(Block source)
        {
            Paragraph paragraph = source as Paragraph;
            if (paragraph != null) return paragraph.Clone();
            Table table = source as Table;
            if (table != null) return CloneTable(table);
            return new PageBreak();
        }

        /// <summary>
        /// 深拷贝块列表。
        /// Deep-copies a block list.
        /// </summary>
        /// <param name="source">源块列表。The source block list.</param>
        /// <returns>块列表副本。The copied block list.</returns>
        internal static List<Block> CloneBlockList(List<Block> source)
        {
            List<Block> copy = new List<Block>(source.Count);
            for (int i = 0; i < source.Count; i++) copy.Add(CloneBlock(source[i]));
            return copy;
        }

        private static TableBorders CloneBorders(TableBorders source)
        {
            TableBorders copy = new TableBorders();
            if (source == null) return copy;
            copy.Top = source.Top == null ? null : source.Top.Clone();
            copy.Bottom = source.Bottom == null ? null : source.Bottom.Clone();
            copy.Left = source.Left == null ? null : source.Left.Clone();
            copy.Right = source.Right == null ? null : source.Right.Clone();
            copy.InsideHorizontal = source.InsideHorizontal == null ? null : source.InsideHorizontal.Clone();
            copy.InsideVertical = source.InsideVertical == null ? null : source.InsideVertical.Clone();
            return copy;
        }

        private static TableCellBorders CloneCellBorders(TableCellBorders source)
        {
            if (source == null) return null;
            TableCellBorders copy = new TableCellBorders();
            copy.Top = source.Top == null ? null : source.Top.Clone();
            copy.Bottom = source.Bottom == null ? null : source.Bottom.Clone();
            copy.Left = source.Left == null ? null : source.Left.Clone();
            copy.Right = source.Right == null ? null : source.Right.Clone();
            return copy;
        }
    }
}