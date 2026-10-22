using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Excel;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Word;

namespace Wisdom.Office
{
    /// <summary>报表的视觉样式配置（字体、配色与斑马纹）；不修改时即为内置的深蓝灰商务风格。Visual style configuration for a report (fonts, colors and zebra stripes); the built-in dark blue-gray business theme when unmodified.</summary>
    public sealed class ReportStyle
    {
        /// <summary>全局字体族名，默认 "Arial"。Global font family, default "Arial".</summary>
        public string FontFamily { get; set; }
        /// <summary>正文字号（磅），默认 10。Body font size (points), default 10.</summary>
        public double BaseFontSize { get; set; }
        /// <summary>强调色，用于大标题与题注，默认深蓝 #1F3A5F。Accent color for titles and captions, default dark blue #1F3A5F.</summary>
        public OfficeColor AccentColor { get; set; }
        /// <summary>表头文字颜色，默认白色。Header text color, default white.</summary>
        public OfficeColor HeaderTextColor { get; set; }
        /// <summary>表头底色，默认深蓝 #1F3A5F。Header background color, default dark blue #1F3A5F.</summary>
        public OfficeColor HeaderFill { get; set; }
        /// <summary>隔行（斑马纹）行底色，默认浅蓝灰 #F2F5F9。Alternating (zebra) row background color, default light blue-gray #F2F5F9.</summary>
        public OfficeColor AlternateRowFill { get; set; }
        /// <summary>表格边框颜色，默认浅灰 #D0D0D0。Table border color, default light gray #D0D0D0.</summary>
        public OfficeColor BorderColor { get; set; }
        /// <summary>是否启用隔行变色。默认 true。Whether alternating row shading is enabled. Default true.</summary>
        public bool ZebraStripes { get; set; }

        /// <summary>创建使用默认商务配色的样式实例。Creates a style instance with the default business color scheme.</summary>
        public ReportStyle()
        {
            FontFamily = "Arial";
            BaseFontSize = 10;
            AccentColor = new OfficeColor(0x1F, 0x3A, 0x5F);
            HeaderTextColor = OfficeColor.White;
            HeaderFill = new OfficeColor(0x1F, 0x3A, 0x5F);
            AlternateRowFill = new OfficeColor(0xF2, 0xF5, 0xF9);
            BorderColor = new OfficeColor(0xD0, 0xD0, 0xD0);
            ZebraStripes = true;
        }
    }

    /// <summary>报表数据表的一列定义：表头文字、宽度、数字格式与对齐方式。A column definition in a report data table: header text, width, number format, and alignment.</summary>
    public sealed class ReportColumn
    {
        /// <summary>表头显示文字。Header display text.</summary>
        public string Header { get; set; }
        /// <summary>列宽（磅）；0 表示由导出器按内容自动决定。Column width (points); 0 lets the exporter auto-size it by content.</summary>
        public double Width { get; set; }
        /// <summary>数值的 .NET 格式字符串（如 "N2"、"0.00%"）；可为 null。.NET format string for numbers (e.g. "N2", "0.00%"); may be null.</summary>
        public string NumberFormat { get; set; }
        /// <summary>列内水平对齐方式，默认左对齐。Horizontal alignment within the column, default left.</summary>
        public HorizontalAlignment Alignment { get; set; }

        /// <summary>创建空列（表头为空、左对齐）。Creates an empty column (empty header, left-aligned).</summary>
        public ReportColumn()
        {
            Header = string.Empty;
            Alignment = HorizontalAlignment.Left;
        }

        /// <summary>创建带表头文字的列。Creates a column with the given header text.</summary>
        /// <param name="header">表头文字；null 视为空字符串。The header text; null is treated as an empty string.</param>
        public ReportColumn(string header)
            : this()
        {
            Header = header ?? string.Empty;
        }
    }

    /// <summary>报表中的一张数据表：若干列定义（<see cref="ReportColumn"/>）加若干行值，行值按列序位置对应。A data table in a report: several column definitions (<see cref="ReportColumn"/>) plus several row values, with each row value matched to its column by position.</summary>
    public sealed class ReportTable
    {
        /// <summary>表格题注，输出为表格上方的小标题；可为 null。Table caption, emitted as a small heading above the table; may be null.</summary>
        public string Caption { get; set; }
        /// <summary>表格跨页时是否重复表头行。默认 true。Whether to repeat the header row when the table spans pages. Default true.</summary>
        public bool RepeatHeaderRow { get; set; }
        /// <summary>列定义集合。The column definition collection.</summary>
        public List<ReportColumn> Columns { get; private set; }
        /// <summary>数据行集合，每行元素与 <see cref="Columns"/> 按位置对应。The data-row collection; each row's elements map to <see cref="Columns"/> by position.</summary>
        public List<object[]> Rows { get; private set; }

        /// <summary>创建无题注的空表。Creates an empty table with no caption.</summary>
        public ReportTable()
        {
            Columns = new List<ReportColumn>();
            Rows = new List<object[]>();
            RepeatHeaderRow = true;
        }

        /// <summary>创建带题注的空表。Creates an empty table with a caption.</summary>
        /// <param name="caption">表格题注。The table caption.</param>
        public ReportTable(string caption)
            : this()
        {
            Caption = caption;
        }

        /// <summary>添加一列并完整指定宽度、数字格式与对齐方式。Adds a column with fully specified width, number format, and alignment.</summary>
        /// <param name="header">表头文字。The header text.</param>
        /// <param name="width">列宽（磅），0 表示自动。Column width (points); 0 means automatic.</param>
        /// <param name="numberFormat">数字格式字符串，可为 null。The number-format string; may be null.</param>
        /// <param name="alignment">水平对齐方式。The horizontal alignment.</param>
        /// <returns>新添加的列，便于继续调整。The newly added column, for further adjustment.</returns>
        public ReportColumn AddColumn(string header, double width, string numberFormat, HorizontalAlignment alignment)
        {
            ReportColumn column = new ReportColumn(header);
            column.Width = width;
            column.NumberFormat = numberFormat;
            column.Alignment = alignment;
            Columns.Add(column);
            return column;
        }

        /// <summary>添加一列，使用默认自动宽度与左对齐。Adds a column with the default automatic width and left alignment.</summary>
        /// <param name="header">表头文字。The header text.</param>
        /// <returns>新添加的列。The newly added column.</returns>
        public ReportColumn AddColumn(string header)
        {
            return AddColumn(header, 0, null, HorizontalAlignment.Left);
        }

        /// <summary>追加一行数据；值按列顺序对应，超出列数的值被忽略，缺少的列填 null。Appends a data row; values map to columns in order, extra values are ignored, and missing columns are filled with null.</summary>
        /// <param name="values">各列的值（数字、日期、文本等，导出时自动格式化）。The values per column (numbers, dates, text, etc.; auto-formatted on export).</param>
        /// <returns>返回自身，便于链式调用 <c>AddRow(...).AddRow(...)</c>。Returns itself for chaining <c>AddRow(...).AddRow(...)</c>.</returns>
        public ReportTable AddRow(params object[] values)
        {
            object[] row = new object[Columns.Count];
            if (values != null)
            {
                for (int i = 0; i < row.Length && i < values.Length; i++)
                {
                    row[i] = values[i];
                }
            }
            Rows.Add(row);
            return this;
        }

        /// <summary>批量追加多行数据；null 行按当前列数补空。Appends multiple data rows in bulk; a null row is padded to the current column count.</summary>
        /// <param name="rows">行集合，可为 null（此时不做任何事）。The row collection; may be null (in which case nothing is done).</param>
        /// <returns>返回自身，便于链式调用。Returns itself for chaining.</returns>
        public ReportTable AddRows(IEnumerable<object[]> rows)
        {
            if (rows == null) return this;
            foreach (object[] row in rows)
            {
                Rows.Add(row ?? new object[Columns.Count]);
            }
            return this;
        }
    }

    /// <summary>
    /// 业务报表构建器：以"标题 + 若干数据表 + 备注"组织内容，一次构建即可导出为 PDF、Word 或 Excel。
    /// Business report builder: organizes content as "a title + several data tables + notes"; build once and export to PDF, Word, or Excel.
    /// </summary>
    /// <remarks>典型用法：
    /// <code>
    /// var report = new Report("月度销售汇总");
    /// report.Author = "财务部";
    /// var table = report.AddTable("各地区销售额");
    /// table.AddColumn("地区");
    /// table.AddColumn("金额", 90, "N2", HorizontalAlignment.Right);
    /// table.AddRow("华东", 123456.78);
    /// report.AddNote("数据来源：ERP 系统。");
    /// report.Export("月度销售.pdf");
    /// </code>
    /// Typical usage: construct a Report, add tables and notes, then export it.</remarks>
    public sealed class Report
    {
        /// <summary>报表主标题，输出为文档大标题并写入元数据。Main report title; emitted as the document title and written into metadata.</summary>
        public string Title { get; set; }
        /// <summary>副标题，显示在主标题下方；可为空。Subtitle shown below the main title; may be empty.</summary>
        public string Subtitle { get; set; }
        /// <summary>作者，写入文档元数据。Author; written into document metadata.</summary>
        public string Author { get; set; }
        /// <summary>页脚文字；为空时使用 "Wisdom.Office 报表"。Footer text; uses "Wisdom.Office 报表" when empty.</summary>
        public string FooterText { get; set; }
        /// <summary>导出 Excel 时的工作表名称，默认 "Report"（超过 31 字符会被截断）。Worksheet name when exporting to Excel, default "Report" (truncated past 31 characters).</summary>
        public string WorksheetName { get; set; }
        /// <summary>是否在正文顶部显示生成时间行。默认 true。Whether to show a generation-time line at the top of the body. Default true.</summary>
        public bool ShowGeneratedTime { get; set; }
        /// <summary>报表生成时间，默认构造时的当前时间。Report generation time, default the current time at construction.</summary>
        public DateTime GeneratedAt { get; set; }
        /// <summary>报表样式（字体与配色），默认使用内置商务样式。Report style (fonts and colors); the built-in business style by default.</summary>
        public ReportStyle Style { get; set; }
        /// <summary>报表包含的数据表集合，通过 <see cref="AddTable"/> 添加。The data-table collection of the report, added via <see cref="AddTable"/>.</summary>
        public List<ReportTable> Tables { get; private set; }
        /// <summary>文末备注条目集合，通过 <see cref="AddNote"/> 添加。The end-of-document notes collection, added via <see cref="AddNote"/>.</summary>
        public List<string> Notes { get; private set; }

        /// <summary>创建无标题报表，样式与各集合均已初始化。Creates an untitled report with the style and all collections initialized.</summary>
        public Report()
        {
            Title = string.Empty;
            Style = new ReportStyle();
            Tables = new List<ReportTable>();
            Notes = new List<string>();
            ShowGeneratedTime = true;
            GeneratedAt = DateTime.Now;
            WorksheetName = "Report";
        }

        /// <summary>创建带主标题的报表。Creates a report with the given main title.</summary>
        /// <param name="title">报表主标题。The main report title.</param>
        public Report(string title)
            : this()
        {
            Title = title;
        }

        /// <summary>添加一张带题注的数据表。Adds a data table with a caption.</summary>
        /// <param name="caption">表格题注。The table caption.</param>
        /// <returns>新添加的表，可继续 <c>AddColumn</c> / <c>AddRow</c>。The newly added table; continue with <c>AddColumn</c> / <c>AddRow</c>.</returns>
        public ReportTable AddTable(string caption)
        {
            ReportTable table = new ReportTable(caption);
            Tables.Add(table);
            return table;
        }

        /// <summary>追加一条文末备注（显示为灰色项目符号条目）；空白文本被忽略。Appends an end-of-document note (shown as a gray bulleted item); blank text is ignored.</summary>
        /// <param name="text">备注内容。The note content.</param>
        /// <returns>返回自身，便于链式调用。Returns itself for chaining.</returns>
        public Report AddNote(string text)
        {
            if (!string.IsNullOrEmpty(text)) Notes.Add(text);
            return this;
        }

        /// <summary>将报表渲染为统一文档模型：A4 版式，含页眉页脚、标题、数据表与备注。Renders the report into the unified document model: A4 layout with header/footer, title, data tables and notes.</summary>
        /// <returns>可直接导出 DOCX/PDF 的 <see cref="OfficeDocument"/>。An <see cref="OfficeDocument"/> ready to export as DOCX/PDF.</returns>
        public OfficeDocument ToDocument()
        {
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = Title ?? string.Empty;
            document.Metadata.Author = Author ?? string.Empty;
            document.DefaultTextFormat = new TextFormat(Style.FontFamily, Style.BaseFontSize);

            Section section = document.AddSection(PageSetup.A4());
            section.PageSetup.MarginLeft = 54;
            section.PageSetup.MarginRight = 54;
            section.PageSetup.MarginTop = 54;
            section.PageSetup.MarginBottom = 54;

            AddHeaderAndFooter(section);

            if (!string.IsNullOrEmpty(Title))
            {
                Paragraph title = new Paragraph();
                title.SpaceAfter = 4;
                TextFormat titleFormat = new TextFormat(Style.FontFamily, 20);
                titleFormat.Bold = true;
                titleFormat.Color = Style.AccentColor;
                title.AddText(Title, titleFormat);
                section.Blocks.Add(title);
            }
            if (!string.IsNullOrEmpty(Subtitle))
            {
                Paragraph subtitle = new Paragraph();
                subtitle.SpaceAfter = 4;
                TextFormat subtitleFormat = new TextFormat(Style.FontFamily, 11);
                subtitleFormat.Color = OfficeColor.Gray;
                subtitle.AddText(Subtitle, subtitleFormat);
                section.Blocks.Add(subtitle);
            }
            if (ShowGeneratedTime)
            {
                Paragraph generated = new Paragraph();
                generated.SpaceAfter = 12;
                TextFormat generatedFormat = new TextFormat(Style.FontFamily, 9);
                generatedFormat.Italic = true;
                generatedFormat.Color = OfficeColor.Gray;
                generated.AddText("生成时间：" + GeneratedAt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
                    generatedFormat);
                section.Blocks.Add(generated);
            }

            for (int i = 0; i < Tables.Count; i++)
            {
                AddTableToDocument(section, Tables[i]);
            }

            if (Notes.Count > 0)
            {
                Paragraph heading = new Paragraph();
                heading.SpaceBefore = 12;
                heading.SpaceAfter = 4;
                TextFormat headingFormat = new TextFormat(Style.FontFamily, 11);
                headingFormat.Bold = true;
                heading.AddText("备注", headingFormat);
                section.Blocks.Add(heading);

                for (int i = 0; i < Notes.Count; i++)
                {
                    Paragraph note = new Paragraph();
                    note.SpaceAfter = 2;
                    TextFormat noteFormat = new TextFormat(Style.FontFamily, 9);
                    noteFormat.Color = OfficeColor.Gray;
                    note.AddText("• " + Notes[i], noteFormat);
                    section.Blocks.Add(note);
                }
            }

            if (section.Blocks.Count == 0)
            {
                section.AddParagraph(string.Empty);
            }
            return document;
        }

        /// <summary>将报表渲染为 Excel 工作簿：标题合并居中、表头着色、隔行斑马纹并自动调整列宽。Renders the report as an Excel workbook: merged centered title, colored headers, alternating zebra rows, and auto-fitted column widths.</summary>
        /// <returns>单工作表的 <see cref="OfficeWorkbook"/>。A single-worksheet <see cref="OfficeWorkbook"/>.</returns>
        public OfficeWorkbook ToWorkbook()
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.Metadata.Title = Title ?? string.Empty;
            workbook.Metadata.Author = Author ?? string.Empty;

            string sheetName = string.IsNullOrEmpty(WorksheetName) ? "Report" : WorksheetName;
            if (sheetName.Length > 31) sheetName = sheetName.Substring(0, 31);
            Worksheet sheet = workbook.AddWorksheet(sheetName);
            sheet.ShowGridLines = false;
            sheet.FitToPage = true;

            int columnCount = 1;
            for (int i = 0; i < Tables.Count; i++)
            {
                if (Tables[i].Columns.Count > columnCount) columnCount = Tables[i].Columns.Count;
            }
            ApplyColumnWidths(sheet, columnCount);

            int row = 1;
            if (!string.IsNullOrEmpty(Title))
            {
                WriteMergedLine(sheet, row, columnCount, Title, CreateTitleStyle(16), 26);
                row++;
            }
            if (!string.IsNullOrEmpty(Subtitle))
            {
                CellStyle subtitleStyle = CreateTextStyle(10, OfficeColor.Gray, false);
                WriteMergedLine(sheet, row, columnCount, Subtitle, subtitleStyle, 18);
                row++;
            }
            if (ShowGeneratedTime)
            {
                CellStyle generatedStyle = CreateTextStyle(9, OfficeColor.Gray, false);
                generatedStyle.Italic = true;
                WriteMergedLine(sheet, row, columnCount,
                    "生成时间：" + GeneratedAt.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
                    generatedStyle, 16);
                row++;
            }
            row++;

            int firstHeaderRow = 0;
            for (int tableIndex = 0; tableIndex < Tables.Count; tableIndex++)
            {
                ReportTable table = Tables[tableIndex];
                if (!string.IsNullOrEmpty(table.Caption))
                {
                    CellStyle captionStyle = CreateTitleStyle(12);
                    WriteMergedLine(sheet, row, columnCount, table.Caption, captionStyle, 20);
                    row++;
                }

                CellStyle headerStyle = CreateHeaderStyle();
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    Cell cell = sheet.GetCell(row, c + 1);
                    cell.Value = table.Columns[c].Header ?? string.Empty;
                    cell.Style = headerStyle.Clone();
                    cell.Style.HorizontalAlignment = table.Columns[c].Alignment;
                }
                sheet.SetRowHeight(row, 20);
                if (firstHeaderRow == 0) firstHeaderRow = row;
                row++;

                for (int r = 0; r < table.Rows.Count; r++)
                {
                    object[] values = table.Rows[r];
                    for (int c = 0; c < table.Columns.Count; c++)
                    {
                        ReportColumn column = table.Columns[c];
                        Cell cell = sheet.GetCell(row, c + 1);
                        cell.Value = c < values.Length ? values[c] : null;
                        CellStyle style = CreateBodyStyle(column, r % 2 == 1 && Style.ZebraStripes);
                        cell.Style = style;
                    }
                    row++;
                }
                row++;
            }

            if (Notes.Count > 0)
            {
                CellStyle noteStyle = CreateTextStyle(9, OfficeColor.Gray, false);
                for (int i = 0; i < Notes.Count; i++)
                {
                    sheet.GetCell(row, 1).Value = "• " + Notes[i];
                    sheet.GetCell(row, 1).Style = noteStyle.Clone();
                    row++;
                }
            }

            if (firstHeaderRow > 0)
            {
                sheet.FrozenRows = firstHeaderRow;
            }
            return workbook;
        }

        /// <summary>将报表导出为 PDF 字节数组。Exports the report as a PDF byte array.</summary>
        /// <returns>PDF 文件字节。The PDF file bytes.</returns>
        public byte[] ToPdf()
        {
            return new PdfRenderer().RenderToBytes(ToDocument());
        }

        /// <summary>将报表导出为 DOCX（Word）字节数组。Exports the report as a DOCX (Word) byte array.</summary>
        /// <returns>DOCX 文件字节。The DOCX file bytes.</returns>
        public byte[] ToWord()
        {
            return DocxWriter.Write(ToDocument());
        }

        /// <summary>将报表导出为 XLSX（Excel）字节数组。Exports the report as an XLSX (Excel) byte array.</summary>
        /// <returns>XLSX 文件字节。The XLSX file bytes.</returns>
        public byte[] ToExcel()
        {
            return XlsxWriter.Write(ToWorkbook());
        }

        /// <summary>按目标路径的扩展名（.pdf / .docx / .xlsx）将报表导出到磁盘。Exports the report to disk based on the target path extension (.pdf / .docx / .xlsx).</summary>
        /// <param name="path">目标文件路径。The target file path.</param>
        /// <exception cref="NotSupportedException">扩展名不是受支持的导出格式时抛出。Thrown when the extension is not a supported export format.</exception>
        public void Export(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            switch (OfficeFileFormats.FromExtension(Path.GetExtension(path)))
            {
                case OfficeFileFormat.Pdf:
                    new PdfRenderer().Render(ToDocument(), path);
                    break;
                case OfficeFileFormat.Docx:
                    DocxWriter.Write(ToDocument(), path);
                    break;
                case OfficeFileFormat.Xlsx:
                    XlsxWriter.Write(ToWorkbook(), path);
                    break;
                default:
                    throw new NotSupportedException("报表仅支持导出为 .pdf、.docx 或 .xlsx。");
            }
        }

        /// <summary>为文档节添加页眉（右侧显示报表标题）与页脚（居中显示 <see cref="FooterText"/> 或默认文字）。Adds a header (report title on the right) and a footer (centered <see cref="FooterText"/> or default text) to the document section.</summary>
        /// <param name="section">目标节。The target section.</param>
        private void AddHeaderAndFooter(Section section)
        {
            TextFormat headerFormat = new TextFormat(Style.FontFamily, 8);
            headerFormat.Color = OfficeColor.Gray;
            if (!string.IsNullOrEmpty(Title))
            {
                Paragraph header = section.Header.AddParagraph(string.Empty);
                header.Alignment = ParagraphAlignment.Right;
                header.AddText(Title, headerFormat);
            }
            Paragraph footer = section.Footer.AddParagraph(string.Empty);
            footer.Alignment = ParagraphAlignment.Center;
            footer.AddText(string.IsNullOrEmpty(FooterText) ? "Wisdom.Office 报表" : FooterText, headerFormat);
        }

        /// <summary>把一张报表数据表（题注 + 着色表头 + 斑马纹数据行）转换为文档表格并追加到节末尾。Converts one report data table (caption + colored header + zebra-striped data rows) into a document table and appends it to the end of the section.</summary>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="table">报表数据表。The report data table.</param>
        private void AddTableToDocument(Section section, ReportTable table)
        {
            if (!string.IsNullOrEmpty(table.Caption))
            {
                Paragraph caption = new Paragraph();
                caption.SpaceBefore = 10;
                caption.SpaceAfter = 4;
                TextFormat captionFormat = new TextFormat(Style.FontFamily, 12);
                captionFormat.Bold = true;
                captionFormat.Color = Style.AccentColor;
                caption.AddText(table.Caption, captionFormat);
                section.Blocks.Add(caption);
            }

            Table target = new Table();
            target.RepeatHeaderRow = table.RepeatHeaderRow;
            target.CellPadding = 4;
            target.Borders = CreateBorders();

            double specified = 0;
            int specifiedCount = 0;
            for (int c = 0; c < table.Columns.Count; c++)
            {
                if (table.Columns[c].Width > 0)
                {
                    specified += table.Columns[c].Width;
                    specifiedCount++;
                }
            }
            double fallbackWidth = specifiedCount > 0 ? specified / specifiedCount : 90;
            for (int c = 0; c < table.Columns.Count; c++)
            {
                double width = table.Columns[c].Width > 0 ? table.Columns[c].Width : fallbackWidth;
                target.Columns.Add(new TableColumn(width));
            }

            TableRow headerRow = target.AddRow();
            headerRow.IsHeader = true;
            for (int c = 0; c < table.Columns.Count; c++)
            {
                TableCell cell = new TableCell();
                cell.BackgroundColor = Style.HeaderFill;
                Paragraph paragraph = cell.AddParagraph(string.Empty);
                paragraph.Alignment = ToParagraphAlignment(table.Columns[c].Alignment);
                paragraph.SpaceAfter = 0;
                TextFormat format = new TextFormat(Style.FontFamily, Style.BaseFontSize);
                format.Bold = true;
                format.Color = Style.HeaderTextColor;
                paragraph.AddText(table.Columns[c].Header ?? string.Empty, format);
                headerRow.Cells.Add(cell);
            }

            for (int r = 0; r < table.Rows.Count; r++)
            {
                object[] values = table.Rows[r];
                TableRow targetRow = target.AddRow();
                bool striped = Style.ZebraStripes && (r % 2 == 1);
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    ReportColumn column = table.Columns[c];
                    TableCell cell = new TableCell();
                    if (striped) cell.BackgroundColor = Style.AlternateRowFill;
                    Paragraph paragraph = cell.AddParagraph(string.Empty);
                    paragraph.Alignment = ToParagraphAlignment(column.Alignment);
                    paragraph.SpaceAfter = 0;
                    TextFormat format = new TextFormat(Style.FontFamily, Style.BaseFontSize);
                    string text = FormatValue(c < values.Length ? values[c] : null, column);
                    paragraph.AddText(text, format);
                    targetRow.Cells.Add(cell);
                }
            }

            section.Blocks.Add(target);
        }

        /// <summary>按各表同列的定义宽度与最长内容估算并设置 Excel 列宽（上限 40 个字符宽）。Estimates and sets Excel column widths from each column's defined width and longest content across all tables (capped at 40 characters wide).</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="columnCount">需要设置的最大列数。The maximum number of columns to set.</param>
        private void ApplyColumnWidths(Worksheet sheet, int columnCount)
        {
            for (int c = 0; c < columnCount; c++)
            {
                double width = 12;
                for (int t = 0; t < Tables.Count; t++)
                {
                    ReportTable table = Tables[t];
                    if (c >= table.Columns.Count) continue;
                    ReportColumn column = table.Columns[c];
                    if (column.Width > 0) width = Math.Max(width, column.Width / 7.0);
                    int longest = column.Header != null ? column.Header.Length : 0;
                    for (int r = 0; r < table.Rows.Count; r++)
                    {
                        object value = c < table.Rows[r].Length ? table.Rows[r][c] : null;
                        int length = FormatValue(value, column).Length;
                        if (length > longest) longest = length;
                    }
                    width = Math.Max(width, Math.Min(32, longest + 3));
                }
                sheet.SetColumnWidth(c + 1, Math.Min(40, width));
            }
        }

        /// <summary>写入一行横跨全部列的合并文字（用于标题、副标题、生成时间与表格题注）。Writes one merged row of text spanning all columns (used for the title, subtitle, generation time, and table captions).</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="row">行号（1 基）。The row number (1-based).</param>
        /// <param name="columnCount">合并的列数。The number of columns to merge.</param>
        /// <param name="text">显示文字。The display text.</param>
        /// <param name="style">单元格样式。The cell style.</param>
        /// <param name="height">行高。The row height.</param>
        private void WriteMergedLine(Worksheet sheet, int row, int columnCount, string text, CellStyle style, double height)
        {
            for (int c = 1; c <= columnCount; c++)
            {
                Cell cell = sheet.GetCell(row, c);
                cell.Style = style.Clone();
            }
            Cell anchor = sheet.GetCell(row, 1);
            anchor.Value = text ?? string.Empty;
            if (columnCount > 1)
            {
                sheet.Merge("A" + row.ToString(CultureInfo.InvariantCulture) + ":"
                    + CellAddress.ColumnName(columnCount) + row.ToString(CultureInfo.InvariantCulture));
            }
            sheet.SetRowHeight(row, height);
        }

        /// <summary>创建强调色、加粗且垂直居中的标题样式。Creates a title style in the accent color, bold and vertically centered.</summary>
        /// <param name="size">字号（磅）。The font size (points).</param>
        /// <returns>单元格样式。The cell style.</returns>
        private CellStyle CreateTitleStyle(double size)
        {
            CellStyle style = CreateTextStyle(size, Style.AccentColor, true);
            style.VerticalAlignment = VerticalAlignment.Middle;
            return style;
        }

        /// <summary>创建表头样式：深色底、浅色粗体字、自动换行并带网格边框。Creates the header style: dark background, light bold text, word wrap, and grid borders.</summary>
        /// <returns>单元格样式。The cell style.</returns>
        private CellStyle CreateHeaderStyle()
        {
            CellStyle style = new CellStyle();
            style.FontFamily = Style.FontFamily;
            style.FontSize = Style.BaseFontSize;
            style.Bold = true;
            style.FontColor = Style.HeaderTextColor;
            style.FillColor = Style.HeaderFill;
            style.VerticalAlignment = VerticalAlignment.Middle;
            style.WrapText = true;
            ApplyGridBorders(style);
            return style;
        }

        /// <summary>创建数据行样式：应用列对齐与数字格式，斑马纹行加底色，四边带网格线。Creates a data-row style: applies the column alignment and number format, shades zebra-striped rows, and adds grid lines on all four edges.</summary>
        /// <param name="column">所属列定义。The owning column definition.</param>
        /// <param name="striped">是否为需要加隔行底色的行。Whether this row needs alternating-row shading.</param>
        /// <returns>单元格样式。The cell style.</returns>
        private CellStyle CreateBodyStyle(ReportColumn column, bool striped)
        {
            CellStyle style = new CellStyle();
            style.FontFamily = Style.FontFamily;
            style.FontSize = Style.BaseFontSize;
            style.HorizontalAlignment = column.Alignment;
            style.VerticalAlignment = VerticalAlignment.Top;
            style.WrapText = true;
            if (!string.IsNullOrEmpty(column.NumberFormat)) style.NumberFormat = column.NumberFormat;
            if (striped) style.FillColor = Style.AlternateRowFill;
            style.TopBorder = CreateLine();
            style.BottomBorder = CreateLine();
            style.LeftBorder = CreateLine();
            style.RightBorder = CreateLine();
            return style;
        }

        /// <summary>创建无底色、垂直居中的纯文本样式。Creates a plain-text style with no background fill, vertically centered.</summary>
        /// <param name="size">字号（磅）。The font size (points).</param>
        /// <param name="color">文字颜色。The text color.</param>
        /// <param name="bold">是否加粗。Whether the text is bold.</param>
        /// <returns>单元格样式。The cell style.</returns>
        private CellStyle CreateTextStyle(double size, OfficeColor color, bool bold)
        {
            CellStyle style = new CellStyle();
            style.FontFamily = Style.FontFamily;
            style.FontSize = size;
            style.FontColor = color;
            style.Bold = bold;
            style.VerticalAlignment = VerticalAlignment.Middle;
            return style;
        }

        /// <summary>为样式四边应用默认网格边框线。Applies the default grid border lines to all four edges of the style.</summary>
        /// <param name="style">目标单元格样式。The target cell style.</param>
        private void ApplyGridBorders(CellStyle style)
        {
            style.TopBorder = CreateLine();
            style.BottomBorder = CreateLine();
            style.LeftBorder = CreateLine();
            style.RightBorder = CreateLine();
        }

        /// <summary>创建 0.5 磅单实线样式的报表边框线。Creates a report border line in the 0.5pt single solid style.</summary>
        /// <returns>边框线定义。The border line definition.</returns>
        private BorderLine CreateLine()
        {
            return new BorderLine(BorderStyle.Single, 0.5, Style.BorderColor);
        }

        /// <summary>创建文档表格的全网格边框（内外线统一为报表边框色）。Creates full-grid borders for a document table (inner and outer lines uniformly use the report border color).</summary>
        /// <returns>表格边框集合。The table border set.</returns>
        private TableBorders CreateBorders()
        {
            BorderLine line = CreateLine();
            TableBorders borders = new TableBorders();
            borders.Top = line;
            borders.Bottom = line;
            borders.Left = line;
            borders.Right = line;
            borders.InsideHorizontal = line;
            borders.InsideVertical = line;
            return borders;
        }

        /// <summary>将单元值转换为显示文本：优先列的数字格式，日期默认 yyyy-MM-dd，布尔输出 TRUE/FALSE。Converts a cell value to display text: the column's number format takes priority, dates default to yyyy-MM-dd, and booleans render as TRUE/FALSE.</summary>
        /// <param name="value">原始值，可为 null。The raw value; may be null.</param>
        /// <param name="column">所属列定义。The owning column definition.</param>
        /// <returns>显示文本；null 返回空字符串。The display text; null returns an empty string.</returns>
        private static string FormatValue(object value, ReportColumn column)
        {
            if (value == null) return string.Empty;
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            IFormattable formattable = value as IFormattable;
            if (formattable != null)
            {
                string format = !string.IsNullOrEmpty(column.NumberFormat)
                    ? column.NumberFormat
                    : (value is DateTime ? "yyyy-MM-dd" : null);
                return formattable.ToString(format, CultureInfo.InvariantCulture);
            }
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        /// <summary>将列的 Excel 水平对齐方式映射为段落对齐方式（未识别时按左对齐）。Maps the column's Excel horizontal alignment to a paragraph alignment (defaults to left when unrecognized).</summary>
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
    }
}
