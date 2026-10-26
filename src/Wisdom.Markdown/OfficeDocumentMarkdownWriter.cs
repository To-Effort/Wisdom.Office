using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Markdown
{
    /// <summary>
    /// 将统一文档模型反向生成 Markdown 文本的内部写入器，目标是与解析器「解析 → 生成」基本往返一致。
    /// 取舍：符号列表统一为 "-"、有序列表按原分隔符风格从 1 重新编号；下划线与颜色仅在启用原始 HTML 时以 u/span 表达；
    /// 表格的行列合并降级为占位空单元格；图片与 HTML 写入器保持一致，刻意采用 base64 data URI 以保证往返文本自包含。
    /// Internal writer that regenerates Markdown text from the unified document model, aiming for near round-trip parity with the parser.
    /// Trade-offs: bullet lists normalize to "-" and ordered lists renumber from 1 keeping the original delimiter style; underline and color
    /// use raw u/span only when raw HTML is enabled; cell spans degrade to empty filler cells; images deliberately use base64 data URIs,
    /// matching the HTML writer so the round-tripped text stays self-contained.
    /// </summary>
    internal static class OfficeDocumentMarkdownWriter
    {
        /// <summary>每层列表嵌套对应的左缩进磅数。The left indent, in points, that corresponds to one level of list nesting.</summary>
        private const double IndentPerLevel = 24.0;

        /// <summary>每层列表嵌套使用的空格数。The number of spaces used per level of list nesting.</summary>
        private const int SpacesPerLevel = 2;

        /// <summary>
        /// 管道表列数上限，防御上游模型中异常巨大的 ColSpan 让按列数分配的数组撑爆内存。
        /// The cap on pipe-table column count; it stops an absurd ColSpan in the source model from sizing
        /// the per-column allocation beyond memory.
        /// </summary>
        private const int MaxTableColumns = 4096;

        /// <summary>
        /// 渲染入口：标题转 ATX 井号、段落空行分隔、连续等宽段落转围栏代码块、列表转缩进列表、表格转 GFM 管道表。
        /// Entry point: headings become ATX hashes, paragraphs are separated by blank lines, consecutive monospaced paragraphs become fenced blocks,
        /// list paragraphs become indented lists and tables become GFM pipe tables.
        /// </summary>
        /// <param name="document">源文档；null 时返回空字符串。The source document; null returns an empty string.</param>
        /// <param name="options">渲染选项；null 时回退默认值。The render options; null falls back to defaults.</param>
        /// <returns>Markdown 文本。The Markdown text.</returns>
        internal static string Render(OfficeDocument document, MarkdownOptions options)
        {
            if (document == null)
            {
                return string.Empty;
            }
            if (options == null)
            {
                options = new MarkdownOptions();
            }

            List<Block> blocks = FlattenBlocks(document);
            StringBuilder builder = new StringBuilder();
            WriteTitleIfNeeded(builder, document, blocks, options);

            ListNumbering numbering = new ListNumbering();
            int index = 0;
            while (index < blocks.Count)
            {
                Paragraph paragraph = blocks[index] as Paragraph;
                if (paragraph == null)
                {
                    numbering.Reset();
                    Table table = blocks[index] as Table;
                    if (table != null)
                    {
                        WriteTable(builder, table, options);
                    }
                    else if (blocks[index] is PageBreak)
                    {
                        EnsureBlankLine(builder);
                        builder.Append("---\n");
                    }
                    index++;
                    continue;
                }

                if (IsCodeParagraph(paragraph, options))
                {
                    numbering.Reset();
                    int start = index;
                    while (index < blocks.Count)
                    {
                        Paragraph next = blocks[index] as Paragraph;
                        if (next == null || !IsCodeParagraph(next, options))
                        {
                            break;
                        }
                        index++;
                    }
                    EnsureBlankLine(builder);
                    WriteFencedCode(builder, blocks, start, index);
                    continue;
                }

                if (IsSeparator(paragraph))
                {
                    numbering.Reset();
                    EnsureBlankLine(builder);
                    builder.Append("---\n");
                    index++;
                    continue;
                }

                if (!string.IsNullOrEmpty(paragraph.ListMarker))
                {
                    WriteListItem(builder, paragraph, numbering, options);
                    index++;
                    continue;
                }

                numbering.Reset();
                WriteParagraph(builder, paragraph, options);
                index++;
            }

            string text = builder.ToString();
            if (options.TrimTrailingWhitespace)
            {
                text = TrimLines(text);
            }
            return text.Trim('\r', '\n');
        }

        /// <summary>把所有节的正文块按顺序摊平为一个列表。Flattens the body blocks of every section into one ordered list.</summary>
        /// <param name="document">源文档。The source document.</param>
        /// <returns>摊平后的块列表。The flattened block list.</returns>
        private static List<Block> FlattenBlocks(OfficeDocument document)
        {
            List<Block> blocks = new List<Block>();
            if (document.Sections == null)
            {
                return blocks;
            }
            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                if (section == null || section.Blocks == null)
                {
                    continue;
                }
                for (int i = 0; i < section.Blocks.Count; i++)
                {
                    if (section.Blocks[i] != null)
                    {
                        blocks.Add(section.Blocks[i]);
                    }
                }
            }
            return blocks;
        }

        /// <summary>若元数据标题存在且文档首个标题不是它，则在开头补输出一级标题。Emits a level-1 heading first when the metadata title exists and the document's first heading is not it.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="document">源文档。The source document.</param>
        /// <param name="blocks">摊平的块列表。The flattened block list.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteTitleIfNeeded(StringBuilder builder, OfficeDocument document, List<Block> blocks, MarkdownOptions options)
        {
            string title = document.Metadata != null ? document.Metadata.Title : null;
            if (string.IsNullOrWhiteSpace(title))
            {
                return;
            }
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null && paragraph.HeadingLevel > 0)
                {
                    if (paragraph.GetText().Trim() == title.Trim())
                    {
                        return;
                    }
                    break;
                }
            }
            builder.Append("# ").Append(EscapeText(title.Trim(), true, false, options)).Append('\n');
        }

        /// <summary>计算列表嵌套层级：每 24 磅左缩进一层，负值取 0、超过 8 层封顶。Computes the list nesting level: one level per 24 points of left indent, floored at 0 and capped at 8.</summary>
        /// <param name="paragraph">列表段落。The list paragraph.</param>
        /// <returns>层级（从 0 开始）。The zero-based level.</returns>
        private static int GetLevel(Paragraph paragraph)
        {
            int level = (int)Math.Floor(paragraph.LeftIndent / IndentPerLevel);
            if (level < 0)
            {
                level = 0;
            }
            return level > 8 ? 8 : level;
        }

        /// <summary>判断列表标记是否为有序：以数字加点/括号开头。Tests whether a list marker is ordered: digits followed by "." or ")".</summary>
        /// <param name="marker">列表标记文本。The list marker text.</param>
        /// <returns>是否有序。Whether it is ordered.</returns>
        private static bool IsOrderedMarker(string marker)
        {
            if (string.IsNullOrEmpty(marker))
            {
                return false;
            }
            int pos = 0;
            while (pos < marker.Length && marker[pos] >= '0' && marker[pos] <= '9')
            {
                pos++;
            }
            return pos > 0 && pos < marker.Length && (marker[pos] == '.' || marker[pos] == ')');
        }

        /// <summary>输出一个列表项：符号统一 "- "，有序按原分隔符从 1 重新编号；启用任务列表时识别行首复选框。Writes one list item: bullets normalize to "- ", ordered items renumber from 1 keeping the original delimiter; with task lists enabled a leading checkbox is recognized.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="paragraph">列表段落。The list paragraph.</param>
        /// <param name="numbering">有序编号状态。The ordered numbering state.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteListItem(StringBuilder builder, Paragraph paragraph, ListNumbering numbering, MarkdownOptions options)
        {
            int level = GetLevel(paragraph);
            bool ordered = IsOrderedMarker(paragraph.ListMarker);
            if (!numbering.Active)
            {
                EnsureBlankLine(builder);
                numbering.MarkActive();
            }
            string indent = new string(' ', level * SpacesPerLevel);
            string text = paragraph.GetText();
            bool isChecked;
            if (options.EnableTaskLists && TryGetTaskMarkerFromList(paragraph.ListMarker, out isChecked))
            {
                builder.Append(indent).Append("- ").Append(isChecked ? "[x] " : "[ ] ");
                WriteInlines(builder, paragraph.Inlines, options, new string(' ', indent.Length + 4), false, false);
                builder.Append('\n');
                return;
            }
            string rest;
            if (options.EnableTaskLists && TryGetTaskMarker(text, out rest, out isChecked))
            {
                // 任务项统一退回符号列表风格：CommonMark 任务列表只挂在无序项上。Task items fall back to bullet style: CommonMark attaches checkboxes only to unordered items.
                builder.Append(indent).Append("- ").Append(isChecked ? "[x] " : "[ ] ")
                       .Append(EscapeText(rest, false, false, options)).Append('\n');
                return;
            }
            string prefix = indent + (ordered ? numbering.Next(level, paragraph.ListMarker) : "- ");
            builder.Append(prefix);
            WriteInlines(builder, paragraph.Inlines, options, new string(' ', prefix.Length), false, false);
            builder.Append('\n');
        }

        /// <summary>
        /// 从列表符号里读取任务状态：解析器把 <c>☐</c> / <c>☑</c> 前缀写进 ListMarker。
        /// Reads the task state from the list marker: the parser stores the ☐ / ☑ prefix in ListMarker.
        /// </summary>
        /// <param name="marker">列表符号。The list marker.</param>
        /// <param name="isChecked">输出是否已勾选。Reports whether the item is checked.</param>
        /// <returns>是否为任务项。Whether the item is a task.</returns>
        private static bool TryGetTaskMarkerFromList(string marker, out bool isChecked)
        {
            isChecked = false;
            if (string.IsNullOrEmpty(marker))
            {
                return false;
            }
            if (marker[0] == '\u2611')
            {
                isChecked = true;
                return true;
            }
            return marker[0] == '\u2610';
        }

        /// <summary>识别任务列表标记 "[ ]"/"[x]"/"[X]" 并拆出正文与勾选状态。Recognizes the task markers "[ ]"/"[x]"/"[X]" and splits the body text and checked state.</summary>
        /// <param name="text">列表项全文。The full item text.</param>
        /// <param name="rest">去掉标记后的正文。The body without the marker.</param>
        /// <param name="isChecked">是否已勾选。Whether it is checked.</param>
        /// <returns>是否命中任务标记。Whether a task marker was found.</returns>
        private static bool TryGetTaskMarker(string text, out string rest, out bool isChecked)
        {
            rest = null;
            isChecked = false;
            if (text == null || text.Length < 3 || text[0] != '[')
            {
                return false;
            }
            if (text[1] != ' ' && text[1] != 'x' && text[1] != 'X')
            {
                return false;
            }
            if (text[2] != ']' || (text.Length > 3 && text[3] != ' '))
            {
                return false;
            }
            isChecked = text[1] == 'x' || text[1] == 'X';
            rest = text.Length > 4 ? text.Substring(4) : string.Empty;
            return true;
        }

        /// <summary>输出普通段落或 ATX 标题；超过 6 级的标题按 6 级处理，空段落跳过。Writes a plain paragraph or ATX heading; headings deeper than 6 clamp to 6 and empty paragraphs are skipped.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="paragraph">源段落。The source paragraph.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteParagraph(StringBuilder builder, Paragraph paragraph, MarkdownOptions options)
        {
            StringBuilder content = new StringBuilder();
            WriteInlines(content, paragraph.Inlines, options, string.Empty, false, paragraph.HeadingLevel > 0);
            string text = content.ToString().TrimEnd(' ', '\n');
            if (text.Trim(' ', '\n').Length == 0)
            {
                return;
            }
            EnsureBlankLine(builder);
            if (paragraph.HeadingLevel > 0)
            {
                int level = paragraph.HeadingLevel > 6 ? 6 : paragraph.HeadingLevel;
                builder.Append(new string('#', level)).Append(' ').Append(text.Trim(' ', '\n')).Append('\n');
                return;
            }
            builder.Append(text).Append('\n');
        }

        /// <summary>输出连续等宽段落组成的围栏代码块；围栏长度随内容中最长反引号串自适应加长以避免冲突。Writes consecutive monospaced paragraphs as a fenced code block; the fence grows past the longest backtick run in the content to avoid collisions.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="blocks">摊平的块列表。The flattened block list.</param>
        /// <param name="start">代码段起始下标。The inclusive start index.</param>
        /// <param name="end">代码段结束下标（不含）。The exclusive end index.</param>
        private static void WriteFencedCode(StringBuilder builder, List<Block> blocks, int start, int end)
        {
            StringBuilder content = new StringBuilder();
            int longestRun = 0;
            for (int i = start; i < end; i++)
            {
                if (i > start)
                {
                    content.Append('\n');
                }
                string line = GetRawText((Paragraph)blocks[i]).Replace('\r', ' ');
                content.Append(line);
                longestRun = Math.Max(longestRun, LongestBacktickRun(line));
            }
            int fenceLength = Math.Max(3, longestRun + 1);
            string fence = new string('`', fenceLength);
            builder.Append(fence).Append('\n').Append(content.ToString()).Append('\n').Append(fence).Append('\n');
        }

        /// <summary>统计文本中最长连续反引号串的长度。Counts the longest consecutive backtick run in the text.</summary>
        /// <param name="text">待统计文本。The text to scan.</param>
        /// <returns>最长串长度。The longest run length.</returns>
        private static int LongestBacktickRun(string text)
        {
            int longest = 0;
            int current = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '`')
                {
                    current++;
                    if (current > longest)
                    {
                        longest = current;
                    }
                }
                else
                {
                    current = 0;
                }
            }
            return longest;
        }

        /// <summary>输出表格：未启用时降级为逐行 "单元格 / 单元格" 段落；启用时输出 GFM 管道表，分隔行按列对齐写 :---、---:、:---:。Writes a table: degraded row paragraphs when tables are disabled; otherwise a GFM pipe table whose delimiter row carries per-column alignment markers.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="table">源表格。The source table.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteTable(StringBuilder builder, Table table, MarkdownOptions options)
        {
            if (table.Rows == null || table.Rows.Count == 0)
            {
                return;
            }
            EnsureBlankLine(builder);
            if (!options.EnableTables)
            {
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    TableRow row = table.Rows[r];
                    if (row == null || row.Cells == null)
                    {
                        continue;
                    }
                    StringBuilder line = new StringBuilder();
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        if (c > 0)
                        {
                            line.Append(" / ");
                        }
                        line.Append(EscapeText(GetCellText(row.Cells[c]), false, false, options));
                    }
                    builder.Append(line.ToString()).Append('\n');
                }
                return;
            }

            int columnCount = table.Columns != null ? table.Columns.Count : 0;
            for (int r = 0; r < table.Rows.Count; r++)
            {
                TableRow row = table.Rows[r];
                int sum = 0;
                if (row != null && row.Cells != null)
                {
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        TableCell cell = row.Cells[c];
                        sum += cell != null && cell.ColSpan > 1 ? cell.ColSpan : 1;
                    }
                }
                if (sum > columnCount)
                {
                    columnCount = sum;
                }
            }
            if (columnCount > MaxTableColumns) columnCount = MaxTableColumns;
            if (columnCount <= 0)
            {
                return;
            }

            ParagraphAlignment[] alignments = new ParagraphAlignment[columnCount];
            for (int r = 0; r < table.Rows.Count && r < 4; r++)
            {
                TableRow row = table.Rows[r];
                if (row == null || row.Cells == null)
                {
                    continue;
                }
                int slot = 0;
                for (int c = 0; c < row.Cells.Count && slot < columnCount; c++)
                {
                    TableCell cell = row.Cells[c];
                    if (cell == null)
                    {
                        slot++;
                        continue;
                    }
                    Paragraph first = cell.Blocks != null && cell.Blocks.Count > 0 ? cell.Blocks[0] as Paragraph : null;
                    if (first != null && alignments[slot] == ParagraphAlignment.Left && first.Alignment != ParagraphAlignment.Left)
                    {
                        alignments[slot] = first.Alignment;
                    }
                    slot += cell.ColSpan > 1 ? cell.ColSpan : 1;
                }
            }

            // GFM 管道表强制要求表头行：无论首行是否标记 IsHeader 都作为表头输出，保证结果可被再次解析。
            // GFM pipe tables require a header row: the first row is emitted as the header whether or not it is marked, so the result re-parses.
            WriteTableRow(builder, table.Rows[0], columnCount, options, true);
            WriteDelimiterRow(builder, columnCount, alignments);
            for (int r = 1; r < table.Rows.Count; r++)
            {
                WriteTableRow(builder, table.Rows[r], columnCount, options, false);
            }
        }

        /// <summary>输出一行管道表行；ColSpan 用占位空单元格补足，缺列的行尾部补空。Writes one pipe-table row; ColSpan is padded with filler empty cells and short rows are topped up.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="row">源行。The source row.</param>
        /// <param name="columnCount">列总数。The total column count.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <param name="suppressEmphasis">表头行跳过粗体包裹（GFM 表头默认加粗）。Whether to skip bold wrapping for the header row (GFM headers are bold by default).</param>
        private static void WriteTableRow(StringBuilder builder, TableRow row, int columnCount, MarkdownOptions options, bool suppressEmphasis)
        {
            builder.Append('|');
            int slot = 0;
            if (row != null && row.Cells != null)
            {
                for (int c = 0; c < row.Cells.Count && slot < columnCount; c++)
                {
                    TableCell cell = row.Cells[c];
                    builder.Append(' ');
                    if (cell != null)
                    {
                        builder.Append(GetCellMarkdown(cell, options, suppressEmphasis));
                    }
                    builder.Append(" |");
                    slot += cell != null && cell.ColSpan > 1 ? cell.ColSpan : 1;
                }
            }
            while (slot < columnCount)
            {
                builder.Append("  |");
                slot++;
            }
            builder.Append('\n');
        }

        /// <summary>生成一个单元格的单行 Markdown 内容：多段落与硬换行以 &lt;br&gt; 连接，管道符转义。Builds one cell's single-line Markdown: paragraphs and hard breaks joined with &lt;br&gt;, pipes escaped.</summary>
        /// <param name="cell">源单元格。The source cell.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <param name="suppressEmphasis">是否跳过粗体/斜体包裹。Whether to skip the bold/italic wrappers.</param>
        /// <returns>单元格内容文本。The cell content text.</returns>
        private static string GetCellMarkdown(TableCell cell, MarkdownOptions options, bool suppressEmphasis)
        {
            StringBuilder content = new StringBuilder();
            if (cell.Blocks == null)
            {
                return content.ToString();
            }
            bool first = true;
            for (int b = 0; b < cell.Blocks.Count; b++)
            {
                Paragraph paragraph = cell.Blocks[b] as Paragraph;
                if (paragraph == null)
                {
                    continue;
                }
                StringBuilder part = new StringBuilder();
                WriteInlines(part, paragraph.Inlines, options, string.Empty, true, suppressEmphasis);
                string text = part.ToString()
                    .Replace("  \n", "<br>")
                    .Replace("\n", "<br>")
                    .Trim(' ', '\t');
                if (text.Length == 0)
                {
                    continue;
                }
                if (!first)
                {
                    // 单元格内换行以 <br> 表达：管道表行必须保持单行。In-cell newlines become <br>: a pipe-table row must stay on one line.
                    content.Append("<br>");
                }
                content.Append(text);
                first = false;
            }
            return content.ToString();
        }

        /// <summary>输出表头分隔行：按列对齐写 :---、---:、:---: 或 ---（两端对齐退化为默认）。Writes the delimiter row: per-column :---, ---:, :---: or --- (justify degrades to default).</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="columnCount">列总数。The total column count.</param>
        /// <param name="alignments">各列对齐方式。The per-column alignments.</param>
        private static void WriteDelimiterRow(StringBuilder builder, int columnCount, ParagraphAlignment[] alignments)
        {
            builder.Append('|');
            for (int c = 0; c < columnCount; c++)
            {
                if (alignments[c] == ParagraphAlignment.Center)
                {
                    builder.Append(" :---: |");
                }
                else if (alignments[c] == ParagraphAlignment.Right)
                {
                    builder.Append(" ---: |");
                }
                else if (alignments[c] == ParagraphAlignment.Justify)
                {
                    builder.Append(" --- |");
                }
                else
                {
                    builder.Append(" :--- |");
                }
            }
            builder.Append('\n');
        }

        /// <summary>输出段落内全部行内元素：粗体转 **、斜体转 *、删除线转 ~~、等宽转反引号代码串；下划线与颜色仅在启用原始 HTML 时用 u/span 表达。Writes all inlines: bold to **, italic to *, strikeout to ~~, monospace to backtick spans; underline and color use raw u/span only when raw HTML is enabled.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="inlines">行内元素列表。The inline elements.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <param name="indent">硬换行续行的缩进串。The indent string for continuation lines after a hard break.</param>
        /// <param name="inTable">是否处于表格单元格（决定竖线转义）。Whether inside a table cell (decides pipe escaping).</param>
        /// <param name="suppressEmphasis">是否跳过粗体/斜体/删除线包裹（标题与表头本身就是强调样式）。Whether to skip the bold/italic/strike wrappers (headings and table headers are already emphasised).</param>
        private static void WriteInlines(StringBuilder builder, List<Inline> inlines, MarkdownOptions options, string indent, bool inTable, bool suppressEmphasis)
        {
            if (inlines == null)
            {
                return;
            }
            bool lineStart = builder.Length == 0;
            for (int i = 0; i < inlines.Count; i++)
            {
                Inline inline = inlines[i];
                TextRun text = inline as TextRun;
                if (text != null)
                {
                    lineStart = WriteTextRun(builder, text.Text, text.Format, options, lineStart, indent, inTable, text.Link, suppressEmphasis);
                    continue;
                }
                FieldRun field = inline as FieldRun;
                if (field != null)
                {
                    lineStart = WriteTextRun(builder, field.DisplayText, field.Format, options, lineStart, indent, inTable, null, suppressEmphasis);
                    continue;
                }
                ImageRun image = inline as ImageRun;
                if (image != null)
                {
                    WriteImage(builder, image);
                    lineStart = false;
                }
            }
        }

        /// <summary>输出单个文本运行；纯换行运行输出两个尾部空格加换行的 CommonMark 硬换行，续行按缩进对齐，返回输出后的行首状态。Writes one text run; a newline-only run emits the two-trailing-space CommonMark hard break with an indented continuation, returning the new line-start state.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="text">文本内容。The text content.</param>
        /// <param name="format">文本格式。The text format.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <param name="lineStart">当前是否处于行首。Whether output is currently at a line start.</param>
        /// <param name="indent">续行缩进。The continuation indent.</param>
        /// <param name="inTable">是否处于表格单元格。Whether inside a table cell.</param>
        /// <param name="link">超链接目标，null 或空表示普通文本。The hyperlink target; null or empty means plain text.</param>
        /// <param name="suppressEmphasis">是否跳过粗体/斜体包裹。Whether to skip the bold/italic wrappers.</param>
        /// <returns>输出后是否处于行首。Whether the output ends at a line start.</returns>
        private static bool WriteTextRun(StringBuilder builder, string text, TextFormat format, MarkdownOptions options, bool lineStart, string indent, bool inTable, string link, bool suppressEmphasis)
        {
            if (string.IsNullOrEmpty(text))
            {
                return lineStart;
            }
            TextFormat effective = format ?? new TextFormat();
            if (IsNewLineOnly(text))
            {
                // 硬换行：CommonMark 的两个尾部空格；恰在行首时退化为无操作。Hard break: the two CommonMark trailing spaces; a no-op when already at a line start.
                if (builder.Length > 0 && builder[builder.Length - 1] != '\n')
                {
                    builder.Append("  \n").Append(indent);
                }
                return true;
            }
            bool mono = IsMonospace(effective.GetAsciiFontFamily(), options);
            string body;
            if (mono)
            {
                body = WrapCodeSpan(text.Replace('\r', ' ').Replace('\n', ' '));
            }
            else
            {
                body = EscapeText(text, lineStart, inTable, options);
                if (options.HardLineBreaks)
                {
                    body = body.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "  \n" + indent);
                }
                else
                {
                    body = body.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ');
                }
            }
            if (body.Trim().Length > 0)
            {
                if (effective.Italic && !suppressEmphasis)
                {
                    body = "*" + body.Trim() + "*";
                }
                if (effective.Bold && !suppressEmphasis)
                {
                    body = "**" + body.Trim() + "**";
                }
                if (effective.Strikeout)
                {
                    body = "~~" + body.Trim() + "~~";
                }
                if (effective.Underline && options.EnableRawHtml)
                {
                    body = "<u>" + body + "</u>";
                }
                if (effective.Color != OfficeColor.Black && effective.Color.A != 0 && options.EnableRawHtml)
                {
                    body = "<span style=\"color:#" + effective.Color.ToHex() + "\">" + body + "</span>";
                }
            }
            if (body.Trim().Length > 0 && !string.IsNullOrEmpty(link))
            {
                body = "[" + body.Trim() + "](" + link.Replace(" ", "%20") + ")";
            }
            builder.Append(body);
            return false;
        }

        /// <summary>用自适应长度的反引号串包裹行内代码；内容首尾含反引号时按 CommonMark 补空格。Wraps inline code in a backtick string of adaptive length; pads with spaces when the content starts or ends with a backtick, per CommonMark.</summary>
        /// <param name="text">代码文本。The code text.</param>
        /// <returns>行内代码串。The inline code span.</returns>
        private static string WrapCodeSpan(string text)
        {
            string fence = new string('`', LongestBacktickRun(text) + 1);
            bool pad = text.StartsWith("`", StringComparison.Ordinal) || text.EndsWith("`", StringComparison.Ordinal);
            return pad ? fence + " " + text + " " + fence : fence + text + fence;
        }

        /// <summary>输出图片为 ![alt](data URI)；与 HTML 写入器一致采用 base64 内联，保证往返文本自包含。Writes the image as ![alt](data URI); like the HTML writer it inlines base64 so the round-tripped text stays self-contained.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="image">图片运行。The image run.</param>
        private static void WriteImage(StringBuilder builder, ImageRun image)
        {
            OfficeImage data = image.Image;
            if (data == null || data.Data == null || data.Data.Length == 0)
            {
                builder.Append("![image](image.png)");
                return;
            }
            string mime = data.Format == OfficeImageFormat.Jpeg ? "image/jpeg" : "image/png";
            builder.Append("![image](data:").Append(mime).Append(";base64,")
                   .Append(Convert.ToBase64String(data.Data)).Append(')');
        }

        /// <summary>判断字体名是否等宽：Consolas/Courier/Monospace 前缀匹配，或等于选项指定的代码字体。Tests whether a font family is monospaced: prefix match on Consolas/Courier/Monospace, or equal to the configured code font.</summary>
        /// <param name="fontFamily">字体名。The font family.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <returns>是否等宽。Whether it is monospaced.</returns>
        private static bool IsMonospace(string fontFamily, MarkdownOptions options)
        {
            if (string.IsNullOrEmpty(fontFamily))
            {
                return false;
            }
            string name = fontFamily.Trim();
            if (name.StartsWith("Consolas", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("Courier", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("Monospace", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return options != null && !string.IsNullOrEmpty(options.CodeFontFamily)
                && name.Equals(options.CodeFontFamily, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>判断代码段落：先认解析器打的结构标记，无标记的手写文档才退回等宽字体推断（空白的等宽运行也算，空围栏与代码块内的空行靠它保住）。Detects a code paragraph: the parser's structural marker wins, and only a hand-built document with no marker falls back to the monospaced-font test (an empty monospaced run counts too, since an empty fence and the blank lines inside a block depend on it).</summary>
        /// <param name="paragraph">待判断段落。The paragraph to test.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <returns>是否为代码段落。Whether it is a code paragraph.</returns>
        private static bool IsCodeParagraph(Paragraph paragraph, MarkdownOptions options)
        {
            if (paragraph.HeadingLevel > 0 || !string.IsNullOrEmpty(paragraph.ListMarker))
            {
                return false;
            }
            if (string.Equals(paragraph.StyleName, MarkdownParser.CodeStyleName, StringComparison.Ordinal))
            {
                return true;
            }
            if (string.Equals(paragraph.StyleName, MarkdownParser.InlineStyleName, StringComparison.Ordinal))
            {
                return false;
            }
            bool sawCodeRun = false;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null)
                {
                    string fontFamily = run.Format != null ? run.Format.GetAsciiFontFamily() : null;
                    if (!IsMonospace(fontFamily, options))
                    {
                        if (!string.IsNullOrEmpty(run.Text))
                        {
                            return false;
                        }
                        continue;
                    }
                    sawCodeRun = true;
                    continue;
                }
                if (!(paragraph.Inlines[i] is FieldRun))
                {
                    return false;
                }
            }
            return sawCodeRun;
        }

        /// <summary>判断分隔线段落：去空白后的正文恰为 em 划线或星号/下划线三连等常见分隔写法。Detects a separator paragraph: trimmed text that is exactly an em dash or the common triple star/dash/underscore forms.</summary>
        /// <param name="paragraph">待判断段落。The paragraph to test.</param>
        /// <returns>是否为分隔线。Whether it is a separator.</returns>
        private static bool IsSeparator(Paragraph paragraph)
        {
            if (!string.IsNullOrEmpty(paragraph.ListMarker) || paragraph.HeadingLevel > 0)
            {
                return false;
            }
            string text = paragraph.GetText().Trim();
            return text == "\u2014" || text == MarkdownParser.ThematicBreakText || text == "* * *"
                || text == "---" || text == "***" || text == "___";
        }

        /// <summary>提取段落原始文本：文本运行取正文、域取显示文本、图片忽略。Extracts the raw paragraph text: text runs contribute their body, fields their display text, images are skipped.</summary>
        /// <param name="paragraph">源段落。The source paragraph.</param>
        /// <returns>原始文本。The raw text.</returns>
        private static string GetRawText(Paragraph paragraph)
        {
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null)
                {
                    text.Append(run.Text);
                    continue;
                }
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field != null)
                {
                    text.Append(field.DisplayText);
                }
            }
            return text.ToString();
        }

        /// <summary>提取单元格纯文本（段落间以空格连接）。Extracts the cell's plain text (paragraphs joined by spaces).</summary>
        /// <param name="cell">源单元格。The source cell.</param>
        /// <returns>纯文本。The plain text.</returns>
        private static string GetCellText(TableCell cell)
        {
            StringBuilder text = new StringBuilder();
            if (cell == null || cell.Blocks == null)
            {
                return text.ToString();
            }
            for (int b = 0; b < cell.Blocks.Count; b++)
            {
                Paragraph paragraph = cell.Blocks[b] as Paragraph;
                if (paragraph == null)
                {
                    continue;
                }
                if (text.Length > 0)
                {
                    text.Append(' ');
                }
                text.Append(GetRawText(paragraph));
            }
            return text.ToString();
        }

        /// <summary>判断文本是否仅由换行符组成。Tests whether the text consists solely of line-break characters.</summary>
        /// <param name="text">待判断文本。The text to test.</param>
        /// <returns>是否为纯换行。Whether it is a pure line break.</returns>
        private static bool IsNewLineOnly(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] != '\n' && text[i] != '\r')
                {
                    return false;
                }
            }
            return text.Length > 0;
        }

        /// <summary>
        /// 对正文文本做最小必要的 Markdown 转义：反斜杠、反引号、星号、下划线、方括号一律转义；
        /// 尖括号仅在疑似标签/自动链接时转义；管道符仅在表格单元格内转义；行首再按标题/列表/引用语法补转义，避免正常文本被过度加噪。
        /// Applies minimal necessary Markdown escaping: backslash, backtick, asterisk, underscore and square brackets are always escaped;
        /// angle brackets only when they look like a tag or autolink; pipes only inside table cells; line starts gain extra escapes
        /// for heading/list/quote syntax so ordinary text stays readable.
        /// </summary>
        /// <param name="text">待转义文本。The text to escape.</param>
        /// <param name="lineStart">是否处于行首。Whether the text begins a line.</param>
        /// <param name="inTable">是否处于表格单元格。Whether the text is inside a table cell.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <returns>转义后的文本。The escaped text.</returns>
        private static string EscapeText(string text, bool lineStart, bool inTable, MarkdownOptions options)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            StringBuilder escaped = new StringBuilder(text.Length + 8);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                char next = i + 1 < text.Length ? text[i + 1] : '\0';
                if (i == 0 && lineStart && (c == '#' || c == '>' || c == '+' || c == '-'))
                {
                    escaped.Append('\\').Append(c);
                    continue;
                }
                if (lineStart && (c == '.' || c == ')') && i > 0 && i <= 10 && (next == ' ' || next == '\0') && AllDigitsBefore(text, i))
                {
                    escaped.Append('\\').Append(c);
                    continue;
                }
                switch (c)
                {
                    case '\\':
                        escaped.Append("\\\\");
                        break;
                    case '`':
                        escaped.Append("\\`");
                        break;
                    case '*':
                        escaped.Append("\\*");
                        break;
                    case '_':
                        escaped.Append("\\_");
                        break;
                    case '[':
                        escaped.Append("\\[");
                        break;
                    case ']':
                        escaped.Append("\\]");
                        break;
                    case '<':
                        if (IsAsciiLetter(next) || next == '!' || next == '/')
                        {
                            escaped.Append("\\<");
                        }
                        else
                        {
                            escaped.Append(c);
                        }
                        break;
                    case '|':
                        if (inTable)
                        {
                            escaped.Append("\\|");
                        }
                        else
                        {
                            escaped.Append(c);
                        }
                        break;
                    default:
                        escaped.Append(c);
                        break;
                }
            }
            return escaped.ToString();
        }

        /// <summary>判断字符串前 endExclusive 个字符是否全为数字。Tests whether the first endExclusive characters of the string are all digits.</summary>
        /// <param name="text">待判断文本。The text to test.</param>
        /// <param name="endExclusive">截止下标（不含）。The exclusive end index.</param>
        /// <returns>是否全为数字。Whether they are all digits.</returns>
        private static bool AllDigitsBefore(string text, int endExclusive)
        {
            for (int i = 0; i < endExclusive; i++)
            {
                if (text[i] < '0' || text[i] > '9')
                {
                    return false;
                }
            }
            return endExclusive > 0;
        }

        /// <summary>判断字符是否为 ASCII 字母。Tests whether the character is an ASCII letter.</summary>
        /// <param name="c">待判断字符。The character to test.</param>
        /// <returns>是否为字母。Whether it is a letter.</returns>
        private static bool IsAsciiLetter(char c)
        {
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
        }

        /// <summary>确保输出缓冲区当前块与下一块之间恰有一个空行；缓冲区为空时不输出。Ensures exactly one blank line between the current and next block; a no-op while the buffer is empty.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        private static void EnsureBlankLine(StringBuilder builder)
        {
            if (builder.Length == 0)
            {
                return;
            }
            int trailing = 0;
            for (int i = builder.Length - 1; i >= 0 && builder[i] == '\n'; i--)
            {
                trailing++;
            }
            if (trailing < 2)
            {
                builder.Append('\n', 2 - trailing);
            }
        }

        /// <summary>逐行去掉行尾空白，但保留恰好两个尾部空格（硬换行标记）、整段围栏代码内部不动，并去掉整体尾部空行。Trims trailing whitespace per line but keeps exactly two trailing spaces (the hard-break marker), leaves the inside of a fenced block untouched, then removes trailing blank lines.</summary>
        /// <param name="text">原始文本。The raw text.</param>
        /// <returns>整理后的文本。The tidied text.</returns>
        private static string TrimLines(string text)
        {
            string[] lines = text.Split('\n');
            StringBuilder trimmed = new StringBuilder(text.Length);
            bool inFence = false;
            char fenceChar = '\0';
            int fenceLen = 0;
            for (int i = 0; i < lines.Length; i++)
            {
                int run;
                int restIndex;
                char marker = ReadFenceMarker(lines[i], out run, out restIndex);
                // 只有同种字符且不短于起始围栏、其后只剩空白的行才能闭合，代码里的 ``` 因此不会被误判成结尾。
                // Only a run of the same character, at least as long as the opener and followed by whitespace alone,
                // closes the fence, so a ``` inside the code is never mistaken for the end.
                bool keep;
                if (inFence)
                {
                    keep = true;
                    if (marker == fenceChar && run >= fenceLen && lines[i].Substring(restIndex).Trim().Length == 0)
                    {
                        inFence = false;
                    }
                }
                else
                {
                    keep = marker != '\0' && lines[i].IndexOf(marker, restIndex) < 0;
                    if (keep)
                    {
                        fenceChar = marker;
                        fenceLen = run;
                        inFence = true;
                    }
                }
                // 围栏内部的行尾空格是代码内容，裁掉就不再是同一份代码；起止标记行同样原样保留。
                // Trailing spaces inside a fence are code content; trimming them changes the code. The marker lines
                // themselves are kept as written too.
                string line = lines[i];
                if (!keep)
                {
                    line = line.TrimEnd('\r', '\t');
                    int spaces = 0;
                    while (spaces < line.Length && line[line.Length - 1 - spaces] == ' ')
                    {
                        spaces++;
                    }
                    if (spaces >= 2)
                    {
                        // 保留恰好两个尾部空格以维持 CommonMark 硬换行。Keep exactly two trailing spaces so the CommonMark hard break survives.
                        line = line.Substring(0, line.Length - spaces) + "  ";
                    }
                    else
                    {
                        line = line.TrimEnd(' ');
                    }
                }
                if (i < lines.Length - 1)
                {
                    trimmed.Append(line).Append('\n');
                }
                else
                {
                    trimmed.Append(line);
                }
            }
            return trimmed.ToString();
        }

        /// <summary>
        /// 读取一行的围栏标记：≤3 格缩进、同种围栏字符连写 ≥3，返回该字符与连写长度及其后的下标。
        /// Reads a line's fence marker: at most three columns of indent and three or more identical fence characters,
        /// returning the character, the run length and the index just after it.
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <param name="run">输出的连写长度。The run length on output.</param>
        /// <param name="restIndex">输出标记之后的下标。The index just after the marker on output.</param>
        /// <returns>围栏字符，非标记行返回 0。The fence character, or 0 for a non-marker line.</returns>
        private static char ReadFenceMarker(string line, out int run, out int restIndex)
        {
            run = 0;
            restIndex = 0;
            int i = 0;
            while (i < line.Length && (line[i] == ' ' || line[i] == '\t') && i < 4)
            {
                i++;
            }
            if (i >= line.Length)
            {
                return '\0';
            }
            char marker = line[i];
            if (marker != '`' && marker != '~')
            {
                return '\0';
            }
            int start = i;
            while (i < line.Length && line[i] == marker)
            {
                i++;
            }
            run = i - start;
            restIndex = i;
            return run >= 3 ? marker : '\0';
        }

        /// <summary>
        /// 跟踪各列表层级的有序编号：进入新层级或列表类型切换时从 1 重新计数，列表被打断时整体清零。
        /// Tracks per-level ordered counters: numbering restarts at 1 when a level opens fresh or the list kind changes, and everything resets when the list run is interrupted.
        /// </summary>
        private sealed class ListNumbering
        {
            private readonly List<int> counters = new List<int>();
            private readonly List<bool> orderedFlags = new List<bool>();
            private readonly List<string> delimiters = new List<string>();

            /// <summary>当前是否正处于连续列表段内。Whether a contiguous list run is currently active.</summary>
            public bool Active { get; private set; }

            /// <summary>把当前状态标记为列表段进行中。Marks the current state as inside an ongoing list run.</summary>
            public void MarkActive()
            {
                Active = true;
            }

            /// <summary>取得该层级的下一个有序编号文本（含尾随空格）。Gets the next ordered item text for this level (with a trailing space).</summary>
            /// <param name="level">层级（从 0 开始）。The zero-based level.</param>
            /// <param name="marker">原列表标记，用于复用点号或括号分隔符。The original marker, reused to keep the dot or parenthesis delimiter.</param>
            /// <returns>形如 "1. " 的编号串。Numbering text such as "1. ".</returns>
            public string Next(int level, string marker)
            {
                while (counters.Count <= level)
                {
                    counters.Add(0);
                    orderedFlags.Add(false);
                    delimiters.Add(".");
                }
                while (counters.Count > level + 1)
                {
                    counters.RemoveAt(counters.Count - 1);
                    orderedFlags.RemoveAt(orderedFlags.Count - 1);
                    delimiters.RemoveAt(delimiters.Count - 1);
                }
                string delimiter = marker != null && marker.IndexOf(')') >= 0 ? ")" : ".";
                if (!orderedFlags[level])
                {
                    counters[level] = 0;
                    orderedFlags[level] = true;
                }
                else if (delimiters[level] != delimiter)
                {
                    counters[level] = 0;
                }
                delimiters[level] = delimiter;
                counters[level]++;
                return counters[level].ToString(CultureInfo.InvariantCulture) + delimiter + " ";
            }

            /// <summary>列表被打断时清空全部计数状态。Clears all counter state when the list run is interrupted.</summary>
            public void Reset()
            {
                Active = false;
                counters.Clear();
                orderedFlags.Clear();
                delimiters.Clear();
            }
        }
    }
}
