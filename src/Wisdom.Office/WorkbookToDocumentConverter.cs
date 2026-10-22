using System;
using System.Collections.Generic;
using Wisdom.Office.Core;

namespace Wisdom.Office
{
    /// <summary>
    /// 将 Excel 工作簿转换为可打印 Word/PDF 文档时的排版选项。
    /// 默认值适合常规打印：输出工作表名标题、跳过隐藏表、沿用打印设置并绘制网格线。
    /// Layout options for converting an Excel workbook into a printable Word/PDF document.
    /// The defaults suit normal printing: emit a sheet-name heading, skip hidden sheets, honor print setup, and draw grid lines.
    /// </summary>
    public sealed class WorkbookLayoutOptions
    {
        /// <summary>是否在每个工作表前输出一级标题（工作表名）。默认 true。Whether to emit a level-1 heading (the sheet name) before each worksheet. Default true.</summary>
        public bool IncludeSheetNames { get; set; }
        /// <summary>是否一并转换隐藏工作表。默认 false。Whether to also convert hidden worksheets. Default false.</summary>
        public bool IncludeHiddenWorksheets { get; set; }
        /// <summary>是否沿用工作表自身的打印设置（纸张、页边距等）。默认 true。Whether to honor each worksheet's own print setup (paper size, margins, etc.). Default true.</summary>
        public bool UsePrintSetup { get; set; }
        /// <summary>是否为表格绘制网格线边框。默认 true。Whether to draw grid-line borders for tables. Default true.</summary>
        public bool UseGridLines { get; set; }
        /// <summary>统一覆盖所有工作表的页面设置；null 表示不覆盖。Page setup that overrides all worksheets uniformly; null means no override.</summary>
        public PageSetup PageSetup { get; set; }
        /// <summary>正文字体族名，默认 "Arial"。Body font family, default "Arial".</summary>
        public string FontFamily { get; set; }
        /// <summary>正文字号（磅），默认 9。Body font size (points), default 9.</summary>
        public double FontSize { get; set; }
        /// <summary>网格线颜色，默认浅灰（#D0D0D0）。Grid line color, default light gray (#D0D0D0).</summary>
        public OfficeColor GridColor { get; set; }
        /// <summary>数值或日期文本放不下所在列宽时，是否按 Excel 规则输出 "####…" 溢出提示。默认 true。
        /// Whether to emit Excel's "####…" overflow indicator when a number or date is wider than its column. Default true.</summary>
        public bool EmulateNumberOverflow { get; set; }

        /// <summary>创建带默认排版设置的选项实例。Creates an options instance with the default layout settings.</summary>
        public WorkbookLayoutOptions()
        {
            IncludeSheetNames = true;
            IncludeHiddenWorksheets = false;
            UsePrintSetup = true;
            UseGridLines = true;
            FontFamily = "Arial";
            FontSize = 9;
            GridColor = new OfficeColor(0xD0, 0xD0, 0xD0);
            EmulateNumberOverflow = true;
        }
    }

    /// <summary>把工作表按可打印的 Word/PDF 表格布局转换为统一文档模型。Converts worksheets into the unified document model using a printable Word/PDF table layout.</summary>
    public static class WorkbookToDocumentConverter
    {
        /// <summary>工作表标题使用的深蓝色。Dark blue used for worksheet headings.</summary>
        private static readonly OfficeColor HeadingColor = new OfficeColor(0x1F, 0x3A, 0x5F);

        /// <summary>使用默认排版选项将工作簿转换为文档模型。Converts the workbook into a document model using the default layout options.</summary>
        /// <param name="workbook">要转换的工作簿。The workbook to convert.</param>
        /// <returns>可直接交给 PDF/DOCX 写出器的 <see cref="OfficeDocument"/>。An <see cref="OfficeDocument"/> ready for the PDF/DOCX writers.</returns>
        public static OfficeDocument Convert(OfficeWorkbook workbook)
        {
            return Convert(workbook, null);
        }

        /// <summary>
        /// 将工作簿按指定排版选项转换为文档模型：每个工作表生成独立节，
        /// 超出页宽的列自动分带换页，冻结行转为表头，合并单元格转为跨行/跨列表元。
        /// Converts the workbook into a document model using the given layout options: each worksheet becomes its own section,
        /// columns wider than the page are split into banded page breaks, frozen rows become header rows, and merged cells become row/column spans.
        /// </summary>
        /// <param name="workbook">要转换的工作簿。The workbook to convert.</param>
        /// <param name="options">排版选项；null 时使用 <see cref="WorkbookLayoutOptions"/> 的默认值。Layout options; the <see cref="WorkbookLayoutOptions"/> defaults are used when null.</param>
        /// <returns>文档模型；工作簿无可见内容时返回仅含空段落的文档。The document model; a document containing only empty paragraphs when the workbook has no visible content.</returns>
        /// <remarks>示例：XLSX 转 PDF 即先调用本方法得到文档，再交给 <c>PdfRenderer</c> 渲染。
        /// Example: converting XLSX to PDF calls this method first to get a document, then hands it to <c>PdfRenderer</c> for rendering.</remarks>
        public static OfficeDocument Convert(OfficeWorkbook workbook, WorkbookLayoutOptions options)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            if (options == null) options = new WorkbookLayoutOptions();

            OfficeDocument document = new OfficeDocument();
            CopyMetadata(workbook.Metadata, document.Metadata);
            document.DefaultTextFormat = new TextFormat(options.FontFamily, options.FontSize);

            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (sheet.IsHidden && !options.IncludeHiddenWorksheets) continue;
                PageSetup setup = ResolvePageSetup(sheet, options);
                Section section = document.AddSection(setup);

                if (options.IncludeSheetNames)
                {
                    Paragraph heading = new Paragraph();
                    heading.HeadingLevel = 1;
                    heading.SpaceAfter = 8;
                    TextFormat format = new TextFormat(options.FontFamily, 15);
                    format.Bold = true;
                    format.Color = HeadingColor;
                    heading.AddText(sheet.Name, format);
                    section.Blocks.Add(heading);
                }

                double contentWidth = Math.Max(36, setup.Width - setup.MarginLeft - setup.MarginRight);
                int firstRow = sheet.HasPrintArea ? sheet.PrintAreaFirstRow : 1;
                int lastRow = sheet.HasPrintArea ? Math.Min(sheet.PrintAreaLastRow, sheet.LastRow) : sheet.LastRow;
                int firstColumn = sheet.HasPrintArea ? sheet.PrintAreaFirstColumn : 1;
                int lastColumn = sheet.HasPrintArea ? Math.Min(sheet.PrintAreaLastColumn, sheet.LastColumn) : sheet.LastColumn;
                int[] columnContentRow = ComputeColumnContentRows(sheet, firstRow, lastRow, firstColumn, lastColumn);
                int contentLastRow = firstRow - 1;
                int contentLastColumn = firstColumn - 1;
                for (int column = firstColumn; column <= lastColumn; column++)
                {
                    if (columnContentRow[column] < firstRow) continue;
                    contentLastColumn = column;
                    if (columnContentRow[column] > contentLastRow) contentLastRow = columnContentRow[column];
                }
                if (contentLastRow >= firstRow) { lastRow = contentLastRow; lastColumn = contentLastColumn; }

                List<int[]> bands = ComputeColumnBands(sheet, contentWidth, firstColumn, lastColumn);
                List<Table> tables = new List<Table>();
                for (int band = 0; band < bands.Count; band++)
                {
                    int[] bandColumns = bands[band];
                    int bandLastRow = firstRow - 1;
                    for (int c = 0; c < bandColumns.Length; c++)
                    {
                        if (bandColumns[c] <= lastColumn && columnContentRow[bandColumns[c]] > bandLastRow)
                            bandLastRow = columnContentRow[bandColumns[c]];
                    }
                    if (bandLastRow < firstRow) continue;
                    Table table = BuildTable(sheet, workbook.DefaultStyle, options, bandColumns, contentWidth, firstRow, bandLastRow);
                    if (table != null && table.Rows.Count > 0) tables.Add(table);
                }
                for (int t = 0; t < tables.Count; t++)
                {
                    if (t > 0) section.Blocks.Add(new PageBreak());
                    section.Blocks.Add(tables[t]);
                }
                if (tables.Count == 0 && !options.IncludeSheetNames) section.Blocks.Add(new Paragraph(string.Empty));
            }

            if (document.Sections.Count == 0)
            {
                document.AddSection().AddParagraph(string.Empty);
            }
            return document;
        }

        /// <summary>
        /// 工作表列按页宽切分为若干"列带"。Excel 列宽面向屏幕显示，
        /// 宽表若强行压进一页会在数字或表头中间断开，因此宁可分带换页；
        /// 每个列带保留冻结的前导列，保证跨页时行仍可辨认。
        /// 本常量为"按宽度缩放"工作表可接受的最小压缩比例（0.7 即缩至 70%）。
        /// Worksheet columns are split into several "column bands" by page width. Excel column widths target on-screen display;
        /// forcing a wide sheet onto one page would break numbers or headers mid-value, so banded page breaks are preferred;
        /// each band keeps the frozen leading columns so rows stay identifiable across pages.
        /// This constant is the minimum acceptable shrink ratio for "fit to width" worksheets (0.7 means shrink to 70%).
        /// </summary>
        private const double FitToPageScale = 0.7;

        /// <summary>按内容宽度把列切分为若干"页宽列带"；一页能容纳时仅返回单个列带。Splits columns into "page-width column bands" by content width; returns a single band when everything fits on one page.</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="contentWidth">页面可打印内容宽度（磅）。The printable content width of the page (points).</param>
        /// <param name="firstColumn">起始列号（1 基）。The starting column number (1-based).</param>
        /// <param name="lastColumn">结束列号（1 基）。The ending column number (1-based).</param>
        /// <returns>列带列表，每个元素为该带上按顺序排列的列号。The list of column bands; each element holds the column numbers of that band in order.</returns>
        private static List<int[]> ComputeColumnBands(Worksheet sheet, double contentWidth, int firstColumn, int lastColumn)
        {
            List<int[]> bands = new List<int[]>();
            if (lastColumn < firstColumn) lastColumn = firstColumn;

            double[] widths = new double[lastColumn + 1];
            double[] layout = new double[lastColumn + 1];
            double total = 0;
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                widths[column] = sheet.GetColumnWidthPoints(column);
                // 单列宽到放不下时，等比压缩会让同带的姓名列按字符折断；
                // 因此分带只用页宽为这些列封顶，渲染时由缩放逻辑压缩超宽列本身。
                layout[column] = Math.Min(widths[column], contentWidth);
                total += widths[column];
            }

            int[] all = new int[lastColumn - firstColumn + 1];
            for (int column = firstColumn; column <= lastColumn; column++) all[column - firstColumn] = column;
            bands.Add(all);
            int dataColumns = lastColumn - firstColumn + 1;
            if (dataColumns < 2) return bands;
            if (total <= contentWidth) return bands;
            // 勾选"按宽度缩放分页"的工作表可接受适度压缩挤进一页；其余工作表则拆分到额外页。
            if (sheet.FitToPage && total * FitToPageScale <= contentWidth) return bands;

            int leading = sheet.FrozenColumns;
            if (leading <= 0) leading = InferLabelColumns(widths, firstColumn, lastColumn, contentWidth);
            if (leading > dataColumns - 2) leading = dataColumns - 2;

            bands.Clear();
            int next = firstColumn;
            bool first = true;
            while (next <= lastColumn)
            {
                List<int> columns = new List<int>();
                double used = 0;
                if (!first && leading > 0)
                {
                    for (int column = firstColumn; column < firstColumn + leading; column++)
                    {
                        columns.Add(column);
                        used += layout[column];
                    }
                }
                int added = 0;
                while (next <= lastColumn)
                {
                    if (added > 0 && used + layout[next] > contentWidth) break;
                    columns.Add(next);
                    used += layout[next];
                    added++;
                    next++;
                }
                bands.Add(columns.ToArray());
                first = false;
            }
            return bands;
        }

        /// <summary>
        /// 未冻结窗格时推断"标识列"：宽表分页后若不带姓名列，第二页起就只剩一串无法归属的日期数字。
        /// 因此把开头若干窄列（累计不超过页宽 30% 且不超过 200 磅，最多 3 列）在后续列带中重复输出。
        /// With no frozen panes, repeat the leading narrow "label" columns on later bands so a wide
        /// sheet's second page is not a run of orphaned numbers; capped by page width, 200pt and 3 columns.
        /// </summary>
        /// <param name="widths">列号到列宽（磅）的数组。Array mapping column numbers to widths (points).</param>
        /// <param name="firstColumn">起始列号（1 基）。The starting column number (1-based).</param>
        /// <param name="lastColumn">结束列号（1 基）。The ending column number (1-based).</param>
        /// <param name="contentWidth">页面内容宽度（磅）。The page content width (points).</param>
        /// <returns>应在后续列带中重复的前导列数，可能为 0。The number of leading columns to repeat on later bands; may be 0.</returns>
        private static int InferLabelColumns(double[] widths, int firstColumn, int lastColumn, double contentWidth)
        {
            double budget = Math.Min(200, contentWidth * 0.3);
            double used = 0;
            int count = 0;
            for (int column = firstColumn; column <= lastColumn - 2 && count < 3; column++)
            {
                used += widths[column];
                if (used > budget) break;
                count++;
            }
            return count;
        }

        /// <summary>
        /// 统计每列"最后一行有可见内容"的行号。Excel 的已用区域常被样式拉伸出成百个空行，
        /// 分带换页时每个列带只打印到自己真正的内容末尾，否则空行会独占整页，
        /// 转成 PDF 后就是空白页。
        /// Records the last row with visible content per column. Styled-but-empty trailing cells inflate
        /// the used range; trimming each column band to its own content prevents those rows from
        /// printing as blank PDF pages.
        /// </summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="firstRow">起始行号（1 基）。The starting row number (1-based).</param>
        /// <param name="lastRow">结束行号（1 基）。The ending row number (1-based).</param>
        /// <param name="firstColumn">起始列号（1 基）。The starting column number (1-based).</param>
        /// <param name="lastColumn">结束列号（1 基）。The ending column number (1-based).</param>
        /// <returns>列号到最后一个可见内容行号的数组；该列无内容时为 <c>firstRow - 1</c>。Array mapping column numbers to the last row with visible content; <c>firstRow - 1</c> when the column has no content.</returns>
        private static int[] ComputeColumnContentRows(Worksheet sheet, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            if (lastColumn < firstColumn) lastColumn = firstColumn;
            int[] contentRows = new int[lastColumn + 1];
            for (int column = firstColumn; column <= lastColumn; column++) contentRows[column] = firstRow - 1;
            if (lastRow < firstRow) return contentRows;

            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Row < firstRow || cell.Row > lastRow) continue;
                if (cell.Column < firstColumn || cell.Column > lastColumn) continue;
                if (!HasPrintedContent(cell, sheet.ShowZeros)) continue;
                if (cell.Row > contentRows[cell.Column]) contentRows[cell.Column] = cell.Row;
            }
            for (int i = 0; i < sheet.MergedRanges.Count; i++)
            {
                MergedRange range = sheet.MergedRanges[i];
                if (range.Row > lastRow || range.LastRow < firstRow) continue;
                if (range.Column > lastColumn || range.LastColumn < firstColumn) continue;
                if (!HasPrintedContent(sheet.FindCell(range.Row, range.Column), sheet.ShowZeros)) continue;
                int bottom = Math.Min(range.LastRow, lastRow);
                int left = Math.Max(range.Column, firstColumn);
                int right = Math.Min(range.LastColumn, lastColumn);
                for (int column = left; column <= right; column++)
                {
                    if (bottom > contentRows[column]) contentRows[column] = bottom;
                }
            }
            return contentRows;
        }

        /// <summary>单元格是否会在输出中留下可见文字（隐藏零值与纯空白按无内容处理）。Whether a cell leaves visible text in the output (hidden zeros and whitespace-only cells count as no content).</summary>
        /// <param name="cell">单元格，可为 null。The cell; may be null.</param>
        /// <param name="showZeros">工作表是否显示零值。Whether the worksheet displays zero values.</param>
        /// <returns>有可见内容时返回 true。True when there is visible content.</returns>
        private static bool HasPrintedContent(Cell cell, bool showZeros)
        {
            if (cell == null || cell.Value == null) return false;
            if (!showZeros && IsZeroValue(cell.Value)) return false;
            return cell.DisplayText().Trim().Length > 0;
        }

        /// <summary>
        /// 列带宽超出页宽时按"窄列保真、宽列让位"重新分配宽度。
        /// 整表等比缩放会把姓名、星期这类标识列压到按字符折断，报表反而不可读；
        /// 因此不高于均分宽度的列保持原始宽度，剩余空间按原宽比例分给超宽列。
        /// Resizes a band that is wider than the page by shrinking only its wide columns. Scaling every
        /// column proportionally wraps label columns such as name or weekday one character per line, so
        /// columns at or below the even share keep their natural width and the remaining budget is split
        /// among the wide columns in proportion to their original widths.
        /// </summary>
        /// <param name="table">待调整宽度的表格。The table whose column widths are adjusted.</param>
        /// <param name="natural">各列的原始宽度（磅）。The natural width of each column (points).</param>
        /// <param name="contentWidth">页面内容宽度（磅）。The page content width (points).</param>
        private static void FitBandWidth(Table table, double[] natural, double contentWidth)
        {
            double share = contentWidth / natural.Length;
            double fixedSum = 0;
            double wideSum = 0;
            for (int i = 0; i < natural.Length; i++)
            {
                if (natural[i] <= share) fixedSum += natural[i];
                else wideSum += natural[i];
            }
            double remaining = contentWidth - fixedSum;
            if (wideSum <= 0 || remaining <= 0) return;
            for (int i = 0; i < natural.Length; i++)
            {
                double width = natural[i] <= share ? natural[i] : natural[i] * remaining / wideSum;
                table.Columns[i].Width = Math.Max(12, width);
            }
        }

        /// <summary>构建某个列带在指定行区间内的表格，处理表头行、行高、合并单元格与边框。Builds the table for a column band within the given row range, handling header rows, row heights, merged cells and borders.</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="fallbackStyle">单元格未设置样式时使用的工作簿默认样式。The workbook default style used when a cell has no style.</param>
        /// <param name="options">排版选项。Layout options.</param>
        /// <param name="bandColumns">当前列带包含的列号。The column numbers included in the current band.</param>
        /// <param name="contentWidth">页面内容宽度（磅）。The page content width (points).</param>
        /// <param name="firstRow">起始行号（1 基）。The starting row number (1-based).</param>
        /// <param name="lastRow">结束行号（1 基）。The ending row number (1-based).</param>
        /// <returns>表格；区间无效时返回 null。The table; null when the range is invalid.</returns>
        private static Table BuildTable(Worksheet sheet, CellStyle fallbackStyle, WorkbookLayoutOptions options,
            int[] bandColumns, double contentWidth, int firstRow, int lastRow)
        {
            int lastColumn = sheet.LastColumn;
            if (lastRow < firstRow || bandColumns.Length < 1) return null;

            Table table = new Table();
            table.CellPadding = 3;
            table.Borders = options.UseGridLines ? CreateGrid(options.GridColor) : TableBorders.None();
            if (sheet.CenterHorizontally) table.Alignment = TableAlignment.Center;

            bool[] inBand = new bool[lastColumn + 2];
            double[] natural = new double[bandColumns.Length];
            double bandWidth = 0;
            for (int i = 0; i < bandColumns.Length; i++)
            {
                inBand[bandColumns[i]] = true;
                natural[i] = sheet.GetColumnWidthPoints(bandColumns[i]);
                bandWidth += natural[i];
                table.Columns.Add(new TableColumn(natural[i]));
            }
            if (bandWidth > contentWidth) FitBandWidth(table, natural, contentWidth);
            table.Autofit = bandWidth > contentWidth;

            for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++)
            {
                TableRow row = table.AddRow();
                if (sheet.FrozenRows > 0 && rowIndex - firstRow < sheet.FrozenRows) row.IsHeader = true;
                double rowHeight = sheet.GetRowHeight(rowIndex);
                if (rowHeight > 0) row.Height = rowHeight;

                for (int position = 0; position < bandColumns.Length; )
                {
                    int column = bandColumns[position];
                    MergedRange range;
                    if (!TryFindMerge(sheet, rowIndex, column, out range))
                    {
                        row.Cells.Add(CreateCell(sheet.FindCell(rowIndex, column), fallbackStyle, options, sheet.ShowZeros,
                            CapacityOf(options, sheet, column, column)));
                        position++;
                        continue;
                    }

                    int end = position;
                    while (end + 1 < bandColumns.Length && range.Contains(rowIndex, bandColumns[end + 1])) end++;
                    int span = end - position + 1;

                    if (range.Row == rowIndex && range.Column == column)
                    {
                        TableCell anchor = CreateCell(sheet.FindCell(range.Row, range.Column), fallbackStyle, options, sheet.ShowZeros,
                            CapacityOf(options, sheet, range.Column, range.LastColumn));
                        anchor.ColSpan = span;
                        anchor.RowSpan = Math.Min(range.RowSpan, lastRow - rowIndex + 1);
                        row.Cells.Add(anchor);
                    }
                    else if (range.Row == rowIndex || !inBand[range.Column])
                    {
                        // 真正的合并锚点位于其他列带，用空白单元格保持网格对齐。
                        TableCell filler = new TableCell(string.Empty);
                        filler.ColSpan = span;
                        row.Cells.Add(filler);
                    }
                    position = end + 1;
                }
            }
            return table;
        }

        /// <summary>把单个 Excel 单元格转换为文档表格单元格（文本、对齐、底色、边框及零值处理）。Converts a single Excel cell into a document table cell (text, alignment, background, borders, and zero-value handling).</summary>
        /// <param name="cell">源单元格，可为 null。The source cell; may be null.</param>
        /// <param name="fallbackStyle">缺省样式。The fallback style.</param>
        /// <param name="options">排版选项。Layout options.</param>
        /// <param name="showZeros">单元格值为 0 时是否仍显示。Whether to still display the value when the cell is 0.</param>
        /// <param name="capacity">所在列（含合并跨度）可容纳的字符数；不需要溢出模拟时为 double.MaxValue。
        /// Number of characters the column (including the merge span) can hold; double.MaxValue when overflow emulation is off.</param>
        /// <returns>文档表格单元格。The document table cell.</returns>
        private static TableCell CreateCell(Cell cell, CellStyle fallbackStyle, WorkbookLayoutOptions options, bool showZeros,
            double capacity)
        {
            TableCell result = new TableCell();
            CellStyle style = cell != null && cell.Style != null ? cell.Style : fallbackStyle;

            Paragraph paragraph = new Paragraph();
            paragraph.SpaceAfter = 0;
            paragraph.LineSpacing = 1.0;
            if (cell != null && cell.Value != null && (showZeros || !IsZeroValue(cell.Value)))
            {
                paragraph.AddText(FitText(cell, capacity), CreateTextFormat(style, options));
            }
            if (style != null)
            {
                paragraph.Alignment = ToParagraphAlignment(style.HorizontalAlignment);
                result.VerticalAlignment = ToVerticalAlignment(style.VerticalAlignment);
                if (style.FillColor.HasValue) result.BackgroundColor = style.FillColor.Value;
                TableCellBorders borders = CreateCellBorders(style);
                if (borders != null) result.Borders = borders;
            }
            result.Blocks.Add(paragraph);
            return result;
        }

        /// <summary>取合并跨度覆盖的各列宽之和作为可容纳字符数；关闭溢出模拟时返回 double.MaxValue。
        /// Sums the column widths covered by the merge span as the character capacity; returns double.MaxValue when emulation is off.</summary>
        /// <param name="options">排版选项。Layout options.</param>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="firstColumn">首列号（1 基）。The first column (1-based).</param>
        /// <param name="lastColumn">末列号（1 基）。The last column (1-based).</param>
        /// <returns>可容纳字符数。The character capacity.</returns>
        private static double CapacityOf(WorkbookLayoutOptions options, Worksheet sheet, int firstColumn, int lastColumn)
        {
            if (options == null || !options.EmulateNumberOverflow) return double.MaxValue;
            double capacity = 0;
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                capacity += sheet.GetColumnWidth(column);
            }
            return capacity;
        }

        /// <summary>按 Excel 规则给出单元格文本：数值或日期放不下所在列宽时输出与该列等宽的 "####…"。
        /// Renders the cell text the way Excel does: a number or date that does not fit its column becomes a "####…" run as wide as the column.</summary>
        /// <param name="cell">单元格。The cell.</param>
        /// <param name="capacity">可容纳字符数。The character capacity.</param>
        /// <returns>显示文本。The display text.</returns>
        private static string FitText(Cell cell, double capacity)
        {
            string text = cell.DisplayText();
            if (capacity == double.MaxValue || text.Length == 0) return text;
            if (!IsNumberOrDate(cell.Value)) return text;
            if (text.Length <= capacity) return text;
            int hashes = (int)Math.Floor(capacity);
            return hashes <= 0 ? string.Empty : new string('#', hashes);
        }

        /// <summary>判断值是否为会被 Excel 用 "####" 提示溢出的数值或日期类型。
        /// Determines whether the value is a number or date, the only kinds Excel flags with "####".</summary>
        /// <param name="value">单元格值。The cell value.</param>
        /// <returns>数值或日期时返回 true。True for numbers and dates.</returns>
        private static bool IsNumberOrDate(object value)
        {
            return value is double || value is float || value is decimal || value is DateTime
                || value is int || value is long || value is short || value is byte
                || value is uint || value is ulong || value is ushort || value is sbyte;
        }

        /// <summary>判断数值在各常见数值类型意义上是否为 0。Determines whether a numeric value is 0 across the common numeric types.</summary>
        /// <param name="value">单元格值。The cell value.</param>
        /// <returns>值为 0 时返回 true。True when the value is 0.</returns>
        private static bool IsZeroValue(object value)
        {
            if (value is double) return (double)value == 0;
            if (value is float) return (float)value == 0;
            if (value is decimal) return (decimal)value == 0;
            if (value is int) return (int)value == 0;
            if (value is long) return (long)value == 0;
            if (value is short) return (short)value == 0;
            if (value is byte) return (byte)value == 0;
            return false;
        }

        /// <summary>由单元格样式与排版选项生成文本格式（字体、字号、粗斜体、下划线、颜色）。Builds a text format from the cell style and layout options (font, size, bold/italic, underline, color).</summary>
        /// <param name="style">单元格样式，可为 null。The cell style; may be null.</param>
        /// <param name="options">排版选项，提供默认字体与字号。Layout options providing the default font and size.</param>
        /// <returns>文本格式。The text format.</returns>
        private static TextFormat CreateTextFormat(CellStyle style, WorkbookLayoutOptions options)
        {
            TextFormat format = new TextFormat(options.FontFamily, options.FontSize);
            if (style == null) return format;
            if (!string.IsNullOrEmpty(style.FontFamily)) format.FontFamily = style.FontFamily;
            if (style.FontSize > 0) format.FontSize = style.FontSize;
            format.Bold = style.Bold;
            format.Italic = style.Italic;
            format.Underline = style.Underline;
            format.Color = style.FontColor;
            return format;
        }

        /// <summary>转换单元格四边边框；四边均无边线时返回 null。Converts the four-edge borders of a cell; returns null when no edge has a border.</summary>
        /// <param name="style">单元格样式。The cell style.</param>
        /// <returns>边框集合或 null。The border set or null.</returns>
        private static TableCellBorders CreateCellBorders(CellStyle style)
        {
            if (!HasBorder(style.TopBorder) && !HasBorder(style.BottomBorder)
                && !HasBorder(style.LeftBorder) && !HasBorder(style.RightBorder))
            {
                return null;
            }
            TableCellBorders borders = new TableCellBorders();
            if (HasBorder(style.TopBorder)) borders.Top = style.TopBorder.Clone();
            if (HasBorder(style.BottomBorder)) borders.Bottom = style.BottomBorder.Clone();
            if (HasBorder(style.LeftBorder)) borders.Left = style.LeftBorder.Clone();
            if (HasBorder(style.RightBorder)) borders.Right = style.RightBorder.Clone();
            return borders;
        }

        /// <summary>判断边线是否实际存在（非 null 且样式不为 None）。Determines whether a border line actually exists (non-null and style not None).</summary>
        /// <param name="line">边线定义。The border line definition.</param>
        /// <returns>存在时返回 true。True when it exists.</returns>
        private static bool HasBorder(BorderLine line)
        {
            return line != null && line.Style != BorderStyle.None;
        }

        /// <summary>生成全网格表格边框（内外线均为 0.5 磅单实线）。Creates full-grid table borders (inner and outer lines are all 0.5pt single solid lines).</summary>
        /// <param name="color">网格线颜色。The grid line color.</param>
        /// <returns>表格边框集合。The table border set.</returns>
        private static TableBorders CreateGrid(OfficeColor color)
        {
            BorderLine line = new BorderLine(BorderStyle.Single, 0.5, color);
            TableBorders borders = new TableBorders();
            borders.Top = line;
            borders.Bottom = line;
            borders.Left = line;
            borders.Right = line;
            borders.InsideHorizontal = line;
            borders.InsideVertical = line;
            return borders;
        }

        /// <summary>查找覆盖指定单元格的合并区域。Finds the merged range covering the given cell.</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="row">行号（1 基）。The row number (1-based).</param>
        /// <param name="column">列号（1 基）。The column number (1-based).</param>
        /// <param name="result">找到的合并区域；未命中时为默认空区域。The merged range found; the default empty range when there is no hit.</param>
        /// <returns>单元格位于某个合并区域内时返回 true。True when the cell lies within a merged range.</returns>
        private static bool TryFindMerge(Worksheet sheet, int row, int column, out MergedRange result)
        {
            for (int i = 0; i < sheet.MergedRanges.Count; i++)
            {
                MergedRange range = sheet.MergedRanges[i];
                if (range.Contains(row, column))
                {
                    result = range;
                    return true;
                }
            }
            result = new MergedRange();
            return false;
        }

        /// <summary>解析页面设置，优先级：选项覆盖 &gt; 工作表打印设置 &gt; A4；尺寸无效时回退为 A4。Resolves the page setup, priority: options override &gt; worksheet print setup &gt; A4; falls back to A4 when the size is invalid.</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="options">排版选项。Layout options.</param>
        /// <returns>可直接使用的页面设置副本。A ready-to-use copy of the page setup.</returns>
        private static PageSetup ResolvePageSetup(Worksheet sheet, WorkbookLayoutOptions options)
        {
            PageSetup setup = null;
            if (options.PageSetup != null)
            {
                setup = options.PageSetup.Clone();
            }
            else if (options.UsePrintSetup && sheet.PrintSetup != null && sheet.PrintSetup.Width > 0 && sheet.PrintSetup.Height > 0)
            {
                setup = sheet.PrintSetup.Clone();
            }
            else
            {
                setup = PageSetup.A4();
            }
            if (setup.Width <= 0 || setup.Height <= 0)
            {
                setup.Width = 595.28;
                setup.Height = 841.89;
            }
            return setup;
        }

        /// <summary>将 Excel 水平对齐方式映射为段落水平对齐方式（未识别时按左对齐）。Maps the Excel horizontal alignment to a paragraph horizontal alignment (defaults to left when unrecognized).</summary>
        /// <param name="alignment">Excel 水平对齐方式。The Excel horizontal alignment.</param>
        /// <returns>对应的段落对齐方式。The corresponding paragraph alignment.</returns>
        private static ParagraphAlignment ToParagraphAlignment(HorizontalAlignment alignment)
        {
            switch (alignment)
            {
                case HorizontalAlignment.Center: return ParagraphAlignment.Center;
                case HorizontalAlignment.Right: return ParagraphAlignment.Right;
                case HorizontalAlignment.Justify: return ParagraphAlignment.Justify;
                default: return ParagraphAlignment.Left;
            }
        }

        /// <summary>将 Excel 垂直对齐方式映射为单元格垂直对齐方式（未识别时按底部对齐）。Maps the Excel vertical alignment to a cell vertical alignment (defaults to bottom when unrecognized).</summary>
        /// <param name="alignment">Excel 垂直对齐方式。The Excel vertical alignment.</param>
        /// <returns>对应的垂直对齐方式。The corresponding vertical alignment.</returns>
        private static VerticalAlignment ToVerticalAlignment(VerticalAlignment alignment)
        {
            switch (alignment)
            {
                case VerticalAlignment.Middle: return VerticalAlignment.Middle;
                case VerticalAlignment.Top: return VerticalAlignment.Top;
                default: return VerticalAlignment.Bottom;
            }
        }

        /// <summary>将来源元数据（标题、作者、主题、关键字、创建/修改时间等）复制到目标，null 字段按空值处理。Copies source metadata (title, author, subject, keywords, created/modified times, etc.) to the target, treating null fields as empty.</summary>
        /// <param name="source">来源元数据，可为 null。The source metadata; may be null.</param>
        /// <param name="target">目标元数据。The target metadata.</param>
        private static void CopyMetadata(DocumentMetadata source, DocumentMetadata target)
        {
            if (source == null) return;
            target.Title = source.Title ?? string.Empty;
            target.Author = source.Author ?? string.Empty;
            target.Subject = source.Subject ?? string.Empty;
            target.Keywords = source.Keywords ?? string.Empty;
            target.Creator = source.Creator ?? "Wisdom.Office";
            target.Created = source.Created;
            target.Modified = source.Modified;
        }
    }
}
