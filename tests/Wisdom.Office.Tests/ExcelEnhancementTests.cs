using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Text;
using Wisdom.Excel;
using Wisdom.Office.Core;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>执行 Excel 增强能力测试：区域与行列编辑、集合映射、CSV、筛选/超链接/标签色/数据验证、流式查询。Runs the Excel enhancement tests: range and row/column editing, collection mapping, CSV, filter/hyperlink/tab color/data validation and streaming query.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelEnhancementTest(string outputDirectory)
        {
            RunExcelRangeEditingTest();
            RunExcelDecorationShiftTest(outputDirectory);
            RunExcelPrintTitleTest(outputDirectory);
            RunExcelTemplateDecorationTest(outputDirectory);
            RunExcelCollectionTest(outputDirectory);
            RunExcelCsvTest(outputDirectory);
            RunExcelSheetFeatureTest(outputDirectory);
            RunExcelQueryTest(outputDirectory);
            Console.WriteLine("Excel enhancement tests finished.");
        }

        /// <summary>区域 API 与行列插入/删除。Range API plus row/column insertion and deletion.</summary>
        private static void RunExcelRangeEditingTest()
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Data");
            sheet.Range("A1:C2").Values = new object[,]
            {
                { "Name", "Qty", "Price" },
                { "Apple", 3d, 1.5d }
            };
            Require(Equals(sheet.GetCell(1, 1).Value, "Name"), "range Values header");
            Require(Equals(sheet.GetCell(2, 2).Value, 3d), "range Values numeric");
            Require(sheet.Range("A1:C2").Address == "A1:C2", "range Address");

            ExcelRange header = sheet.Range("A1:C1");
            header.SetStyle(style => style.Bold = true);
            Require(header.FirstCell.Style != null && header.FirstCell.Style.Bold, "range SetStyle bold");
            Require(sheet.Range("A1:C2").Style.Bold, "range Style getter");

            object[,] snapshot = sheet.Range("A1:C2").Values;
            Require(Equals(snapshot[0, 1], "Qty") && Equals(snapshot[1, 1], 3d), "range Values getter");

            sheet.Range("A4:C4").Merge();
            Require(sheet.MergedRanges.Count == 1 && sheet.MergedRanges[0].Row == 4, "range Merge");
            sheet.Range("A4:C4").Merge();
            Require(sheet.MergedRanges.Count == 1, "range Merge de-duplicates");

            sheet.InsertRows(2, 1);
            Require(sheet.FindCell(2, 1) == null, "InsertRows leaves new row empty");
            Require(Equals(sheet.GetCell(3, 1).Value, "Apple"), "InsertRows shifts cells down");
            Require(sheet.MergedRanges[0].Row == 5, "InsertRows shifts merged range");
            sheet.DeleteRows(2, 1);
            Require(Equals(sheet.GetCell(2, 1).Value, "Apple"), "DeleteRows shifts cells up");
            Require(sheet.MergedRanges[0].Row == 4, "DeleteRows restores merged range");

            sheet.InsertColumns(2, 1);
            Require(Equals(sheet.GetCell(1, 3).Value, "Qty"), "InsertColumns shifts cells right");
            sheet.DeleteColumns(2, 1);
            Require(Equals(sheet.GetCell(1, 2).Value, "Qty"), "DeleteColumns shifts cells left");

            sheet.Range("A4:C4").Unmerge();
            Require(sheet.MergedRanges.Count == 0, "range Unmerge");

            sheet.Range("A6:B6").SetValue("fill");
            Require(Equals(sheet.GetCell(6, 2).Value, "fill"), "range SetValue fills");
            sheet.Range("A6:B6").Clear();
            Require(sheet.FindCell(6, 1) == null, "range Clear removes contentless cells");

            Worksheet copy = workbook.CopyWorksheet("Data", "Copy");
            Require(copy.Name == "Copy" && Equals(copy.GetCell(2, 1).Value, "Apple"), "CopyWorksheet copies cells");
            Require(!ReferenceEquals(copy.GetCell(2, 1), sheet.GetCell(2, 1)), "CopyWorksheet clones cells");
            workbook.MoveWorksheet(0, 1);
            Require(workbook.Worksheets[0].Name == "Copy", "MoveWorksheet moves the sheet");
            workbook.RenameWorksheet(sheet, "Renamed");
            Require(sheet.Name == "Renamed", "RenameWorksheet renames");
            Require(workbook.RemoveWorksheet(copy), "RemoveWorksheet by instance");
            Require(workbook.Worksheets.Count == 1, "RemoveWorksheet shrinks collection");
            Require(OfficeWorkbook.IsValidWorksheetName("Sheet1") && !OfficeWorkbook.IsValidWorksheetName("A/B")
                && !OfficeWorkbook.IsValidWorksheetName("'quoted"), "IsValidWorksheetName");

            sheet.GetCell(1, 4).Value = "季度销售额统计表";
            sheet.Range("A1:D3").AutoFitColumns(6, 40);
            Require(sheet.ColumnWidths.ContainsKey(4) && sheet.ColumnWidths[4] > 10 && sheet.ColumnWidths[4] <= 40,
                "AutoFitColumns measures full-width text");
            ExcelRange used = sheet.UsedRange();
            Require(used.Column1 == 1 && used.Column2 == 4, "UsedRange covers written cells");
        }

        /// <summary>
        /// 插入/删除行列时看板装饰（条件格式、图表、图片）随区域平移，整行 / 整列引用一并换算，并能在保存后原样读回。
        /// Dashboard decorations (conditional formats, charts, images) shift with row/column edits — whole-row and
        /// whole-column references included — and read back unchanged after a save.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelDecorationShiftTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Board");
            sheet.GetCell(5, 4).Value = 10d;
            sheet.GetCell(5, 5).Value = 12d;

            ExcelConditionalFormat bar = new ExcelConditionalFormat("D5:P5");
            bar.AddDataBar(OfficeColor.Blue);
            sheet.ConditionalFormats.Add(bar);
            ExcelConditionalFormat inside = new ExcelConditionalFormat("D7:P10");
            inside.AddColorScale(OfficeColor.Red, OfficeColor.Green);
            sheet.ConditionalFormats.Add(inside);

            ExcelChart chart = new ExcelChart { Type = ExcelChartType.Line, Row = 12, Column = 20 };
            ExcelChartSeries series = chart.AddSeries("Board!$D$4", "Board!$E$5:$P$5", "Board!$E$5:$P$5");
            sheet.Charts.Add(chart);

            ExcelImage icon = new ExcelImage(new OfficeImage(CreateTestPng(1, 1))) { Row = 14, Column = 3 };
            icon.WidthPoints = 12;
            icon.HeightPoints = 12;
            sheet.Images.Add(icon);

            sheet.InsertRows(3, 2);
            Require(bar.Range == "D7:P7", "InsertRows shifts the conditional-format range");
            Require(inside.Range == "D9:P12", "InsertRows shifts a range below the edit");
            Require(series.ValueRange == "Board!$E$7:$P$7", "InsertRows shifts a series range, keeping prefix and locks");
            Require(chart.Row == 14 && chart.Column == 20, "InsertRows shifts the chart anchor");
            Require(icon.Row == 16 && icon.Column == 3, "InsertRows shifts the image anchor");

            sheet.InsertRows(11, 1);
            Require(inside.Range == "D9:P13", "InsertRows inside a range extends it");
            Require(chart.Row == 15 && icon.Row == 17, "InsertRows below an anchor pushes it down");

            sheet.DeleteRows(11, 1);
            Require(inside.Range == "D9:P12", "DeleteRows inside a range clips it back");
            Require(chart.Row == 14 && icon.Row == 16, "DeleteRows restores anchors above the deleted row");

            sheet.DeleteRows(3, 2);
            Require(bar.Range == "D5:P5", "DeleteRows restores the conditional-format range");
            Require(inside.Range == "D7:P10", "DeleteRows lifts the whole range back");
            Require(series.ValueRange == "Board!$E$5:$P$5", "DeleteRows restores the series range");
            Require(chart.Row == 12 && icon.Row == 14, "DeleteRows restores the anchors");

            ExcelConditionalFormat doomed = new ExcelConditionalFormat("D30:D31");
            doomed.AddIconSet("3TrafficLights1", new double[] { 30, 70 });
            sheet.ConditionalFormats.Add(doomed);
            sheet.DeleteRows(30, 3);
            Require(!sheet.ConditionalFormats.Contains(doomed), "DeleteRows drops a swallowed conditional format");

            ExcelConditionalFormat multi = new ExcelConditionalFormat("A1:A5 C1:C5");
            multi.AddDataBar(OfficeColor.Green);
            sheet.ConditionalFormats.Add(multi);
            sheet.InsertRows(3, 2);
            Require(multi.Range == "A1:A7 C1:C7", "InsertRows shifts every sqref fragment");
            sheet.DeleteRows(3, 2);
            Require(multi.Range == "A1:A5 C1:C5", "DeleteRows shifts every sqref fragment back");

            ExcelConditionalFormat far = new ExcelConditionalFormat("D40:D41 F40:F41");
            far.AddDataBar(OfficeColor.Red);
            sheet.ConditionalFormats.Add(far);
            sheet.DeleteRows(40, 3);
            Require(!sheet.ConditionalFormats.Contains(far), "a rule whose fragments all vanish is dropped");
            Require(sheet.ConditionalFormats.Contains(multi), "an untouched multi-fragment rule survives");

            ExcelConditionalFormat wholeRows = new ExcelConditionalFormat("Board!$6:$8");
            wholeRows.AddDataBar(OfficeColor.Blue);
            sheet.ConditionalFormats.Add(wholeRows);
            ExcelConditionalFormat swallowed = new ExcelConditionalFormat("900:902");
            swallowed.AddDataBar(OfficeColor.Red);
            sheet.ConditionalFormats.Add(swallowed);
            sheet.InsertRows(3, 2);
            Require(wholeRows.Range == "Board!$8:$10", "InsertRows shifts a whole-row reference, keeping prefix and locks");
            sheet.DeleteRows(3, 2);
            Require(wholeRows.Range == "Board!$6:$8", "DeleteRows restores a whole-row reference");
            sheet.DeleteRows(900, 3);
            Require(!sheet.ConditionalFormats.Contains(swallowed), "DeleteRows drops a swallowed whole-row rule");

            ExcelConditionalFormat wholeColumns = new ExcelConditionalFormat("C:D");
            wholeColumns.AddDataBar(OfficeColor.Green);
            sheet.ConditionalFormats.Add(wholeColumns);
            ExcelConditionalFormat lockedColumns = new ExcelConditionalFormat("Board!$C:$D");
            lockedColumns.AddDataBar(OfficeColor.Blue);
            sheet.ConditionalFormats.Add(lockedColumns);
            sheet.InsertRows(5, 1);
            Require(wholeColumns.Range == "C:D", "a row edit leaves a whole-column reference alone");
            Require(lockedColumns.Range == "Board!$C:$D", "a row edit leaves a locked whole-column reference alone");
            sheet.DeleteRows(5, 1);

            sheet.InsertColumns(2, 1);
            Require(wholeColumns.Range == "D:E", "InsertColumns shifts a whole-column reference");
            Require(lockedColumns.Range == "Board!$D:$E", "InsertColumns shifts a locked whole-column reference");
            Require(wholeRows.Range == "Board!$6:$8", "a column edit leaves a whole-row reference alone");
            sheet.DeleteColumns(2, 1);
            Require(wholeColumns.Range == "C:D", "DeleteColumns restores the whole-column reference");
            Require(lockedColumns.Range == "Board!$C:$D", "DeleteColumns restores the locked whole-column reference");

            sheet.InsertColumns(2, 1);
            Require(bar.Range == "E5:Q5", "InsertColumns shifts the conditional-format columns");
            Require(series.ValueRange == "Board!$F$5:$Q$5", "InsertColumns shifts the series columns");
            Require(icon.Column == 4 && chart.Column == 21, "InsertColumns shifts the anchors");
            sheet.DeleteColumns(2, 1);
            Require(bar.Range == "D5:P5", "DeleteColumns restores the columns");
            Require(series.ValueRange == "Board!$E$5:$P$5", "DeleteColumns restores the series columns");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "excel-decoration-shift.xlsx"));
            XlsxWriter.Write(workbook, path);
            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Worksheet board = reloaded.FindWorksheet("Board");
            Require(board.Charts.Count == 1, "the shifted chart survives the round trip");
            Require(board.Images.Count == 1, "the shifted image survives the round trip");
            Require(board.Charts[0].Series[0].ValueRange.Replace("'", string.Empty) == "Board!$E$5:$P$5",
                "the reloaded series range points back at the shifted cells");
            Require(!string.IsNullOrEmpty(board.ConditionalFormats[0].Range), "the reloaded rule keeps its range");
        }

        /// <summary>
        /// 每页重复的标题行列写入 _xlnm.Print_Titles 并能在读取—保存往返后存活，插入删除时同步平移。
        /// Repeated title rows and columns land in _xlnm.Print_Titles, survive a load-save round trip and shift
        /// with row/column edits.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelPrintTitleTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Report");
            sheet.GetCell(1, 1).Value = "Month";
            sheet.GetCell(2, 1).Value = "Jan";
            sheet.SetPrintTitleRows("1:2");
            sheet.SetPrintTitleColumns("A:B");
            Require(sheet.HasPrintTitles, "print titles are set");

            Worksheet plain = workbook.AddWorksheet("Plain");
            plain.GetCell(1, 1).Value = 1d;

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "excel-print-titles.xlsx"));
            XlsxWriter.Write(workbook, path);
            string workbookXml;
            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                workbookXml = ReadZipEntryText(archive, "xl/workbook.xml");
            }
            RequireContains(workbookXml, "_xlnm.Print_Titles");
            RequireContains(workbookXml, "!$1:$2");
            RequireContains(workbookXml, "!$A:$B");
            Require(workbookXml.IndexOf("localSheetId=\"1\"", StringComparison.Ordinal) < 0,
                "a sheet without titles emits no defined name");

            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Worksheet report = reloaded.FindWorksheet("Report");
            Require(report.PrintTitleFirstRow == 1 && report.PrintTitleLastRow == 2, "title rows reload");
            Require(report.PrintTitleFirstColumn == 1 && report.PrintTitleLastColumn == 2, "title columns reload");
            Require(!reloaded.FindWorksheet("Plain").HasPrintTitles, "the untitled sheet stays untitled");

            report.InsertRows(1, 1);
            Require(report.PrintTitleFirstRow == 2 && report.PrintTitleLastRow == 3, "InsertRows shifts title rows");
            report.DeleteRows(1, 1);
            Require(report.PrintTitleFirstRow == 1 && report.PrintTitleLastRow == 2, "DeleteRows restores title rows");
            report.InsertColumns(1, 1);
            Require(report.PrintTitleFirstColumn == 2 && report.PrintTitleLastColumn == 3, "InsertColumns shifts title columns");

            report.SetPrintTitleRows("$3:$4");
            Require(report.PrintTitleFirstRow == 3 && report.PrintTitleLastRow == 4, "title rows accept locks and prefixes");
            report.SetPrintTitleRows(null);
            Require(report.PrintTitleFirstRow == 0, "clearing title rows");
            Require(report.PrintTitleFirstColumn == 2, "clearing rows leaves columns alone");

            bool rejected = false;
            try { report.SetPrintTitleRows("A:B"); }
            catch (FormatException) { rejected = true; }
            Require(rejected, "a column range is rejected as title rows");
        }

        /// <summary>
        /// 模板填充带走工作簿日期基准与表级装饰：循环展开后条件格式、筛选、数据验证、打印区域/重复标题、
        /// 隐藏行列与图表锚点都换算到输出行。
        /// Template fill carries the workbook's date epoch and the sheet-level decorations: once a loop expands,
        /// conditional formats, the filter, data validations, the print area and repeated titles, hidden rows and
        /// columns and the chart anchor all convert to output rows.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelTemplateDecorationTest(string outputDirectory)
        {
            OfficeWorkbook template = CreateDecorationTemplate();
            Require(template.Date1904, "the fixture is a 1904 workbook");

            OfficeWorkbook empty = XlsxTemplate.Fill(template, new TemplateContext());
            Worksheet plain = empty.Worksheets[0];
            Require(empty.Date1904, "fill carries date1904 over");
            Require(plain.LastRow == 2, "an empty loop leaves the header and the total row");
            Require(plain.ConditionalFormats.Count == 0,
                "a rule that only covered the vanished loop rows is dropped instead of pointing at nothing");
            Require(plain.DataValidations.Count == 0, "a validation over vanished rows is dropped");
            Require(plain.Charts.Count == 0, "a chart whose series lost every value row is dropped");
            Require(plain.HasPrintArea && plain.PrintAreaFirstRow == 1 && plain.PrintAreaLastRow == 47,
                "a print area past the used range extrapolates through an empty loop");

            OfficeWorkbook flat = CreateDecorationTemplate();
            Worksheet flatSheet = flat.Worksheets[0];
            flatSheet.GetCell(2, 1).Value = "行A";
            flatSheet.GetCell(4, 1).Value = "行B";
            Worksheet kept = XlsxTemplate.Fill(flat, new TemplateContext()).Worksheets[0];
            Require(kept.Charts.Count == 1 && ReferenceEquals(kept.Charts[0], flatSheet.Charts[0]),
                "when no row moves the chart is carried over as-is, literal caches intact");
            Require(kept.ConditionalFormats[0].Range == "A3:B3", "with no row movement the conditional range only loses its prefix");
            Require(kept.DataValidations[0].Range == "B3", "a degenerate range collapses to its single cell");
            Require(kept.Images.Count == 1 && kept.Images[0].Row == 5, "with no row movement the image stays put");

            TemplateContext context = new TemplateContext();
            List<object> items = new List<object>();
            for (int i = 0; i < 3; i++)
            {
                Dictionary<string, object> row = new Dictionary<string, object>();
                row["name"] = "行" + (i + 1);
                items.Add(row);
            }
            context.Set("items", items);
            OfficeWorkbook filled = XlsxTemplate.Fill(template, context);
            Worksheet sheet = filled.Worksheets[0];

            Require(sheet.LastRow == 5, "three items plus two fixed rows give five rows");
            Require(sheet.GetCell(2, 1).Value as string == "行1", "the loop starts right after the header");
            Require(sheet.GetCell(5, 1).Value as string == "合计", "the total row follows the expanded loop");
            Require(sheet.ConditionalFormats[0].Range == "A2:B4", "the conditional range follows the loop");
            Require(sheet.AutoFilter != null && sheet.AutoFilter.Range == "A1:C4", "the filter end snaps to the last loop row");
            Require(sheet.DataValidations[0].Range == "B2:B4", "the validation range follows the loop");
            Require(sheet.HasPrintArea && sheet.PrintAreaLastRow == 50,
                "a print area that runs past the used range is extrapolated, not dropped");
            Require(sheet.PrintTitleFirstRow == 1 && sheet.PrintTitleLastRow == 1
                && sheet.PrintTitleFirstColumn == 1 && sheet.PrintTitleLastColumn == 2,
                "repeated title rows and columns carry over");
            Require(sheet.HiddenColumns.Contains(3), "the hidden helper column stays hidden");
            Require(sheet.HiddenRows.Contains(5), "the hidden total row maps to the output row");
            Require(!sheet.ShowZeros && sheet.CenterHorizontally && sheet.TabColor != null,
                "view flags and the tab color carry over");
            Require(sheet.Charts.Count == 1 && sheet.Charts[0].Row == 6, "the chart anchor moves with the rows");
            Require(sheet.Charts[0].Series[0].CategoryRange == "A2:A4", "series ranges follow the loop");
            Require(sheet.Images.Count == 1 && sheet.Images[0].Row == 5, "the anchored image moves with the rows");
            Require(sheet.Images[0] != template.Worksheets[0].Images[0], "a moved image is a copy, the template stays intact");
            Require(template.Worksheets[0].ConditionalFormats[0].Range == "Report!$A$3:$B$3",
                "fill never rewrites the template's own decorations");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "excel-template-decorations.xlsx"));
            XlsxWriter.Write(filled, path);
            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Worksheet again = reloaded.Worksheets[0];
            Require(reloaded.Date1904, "the written file keeps the 1904 epoch");
            Require(again.GetCell(1, 1).Value is DateTime && (DateTime)again.GetCell(1, 1).Value == new DateTime(2024, 3, 5),
                "a 1904 date does not drift through fill and save");
            Require(again.ConditionalFormats.Count == 1 && again.ConditionalFormats[0].Range == "A2:B4",
                "the mapped conditional range reads back");
            Require(again.AutoFilter != null && again.AutoFilter.Range == "A1:C4", "the mapped filter reads back");
            Require(again.DataValidations.Count == 1 && again.DataValidations[0].Range == "B2:B4",
                "the mapped validation reads back");
            Require(again.HasPrintArea && again.PrintAreaLastRow == 50, "the mapped print area reads back");
            Require(again.PrintTitleFirstColumn == 1 && again.PrintTitleLastColumn == 2,
                "the repeated title columns read back");
            Require(again.Charts.Count == 1 && again.Charts[0].Row == 6, "the remapped chart anchor reads back");
            Require(again.Images.Count == 1 && again.Images[0].Row == 5, "the moved image reads back");
        }

        /// <summary>
        /// 构造带循环与各种表级装饰的模板工作簿。 Builds the template workbook that carries a loop and every sheet-level decoration.
        /// </summary>
        /// <returns>模板工作簿。 The template workbook.</returns>
        private static OfficeWorkbook CreateDecorationTemplate()
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.Date1904 = true;
            Worksheet sheet = workbook.AddWorksheet("Report");
            sheet.GetCell(1, 1).Value = new DateTime(2024, 3, 5);
            sheet.GetCell(1, 1).EnsureStyle().NumberFormat = "yyyy-mm-dd";
            sheet.GetCell(2, 1).Value = "{{#each items}}";
            sheet.GetCell(3, 1).Value = "{{name}}";
            sheet.GetCell(4, 1).Value = "{{/each}}";
            sheet.GetCell(5, 1).Value = "合计";
            sheet.SetColumnWidth(3, 12);
            sheet.SetColumnHidden(3, true);
            sheet.SetRowHidden(5, true);
            sheet.ShowZeros = false;
            sheet.CenterHorizontally = true;
            sheet.TabColor = new OfficeColor(0xFF, 0x00, 0x00);
            sheet.SetPrintArea("A1:C50");
            sheet.SetPrintTitleRows("1:1");
            sheet.SetPrintTitleColumns("A:B");

            ExcelConditionalFormat format = new ExcelConditionalFormat("Report!$A$3:$B$3");
            format.Rules.Add(new ExcelConditionalRule());
            sheet.ConditionalFormats.Add(format);

            ExcelDataValidation validation = new ExcelDataValidation();
            validation.Range = "B3:B3";
            validation.Type = DataValidationType.Whole;
            validation.Operator = DataValidationOperator.Between;
            validation.Formula1 = "1";
            validation.Formula2 = "100";
            sheet.DataValidations.Add(validation);

            sheet.AutoFilter = new ExcelAutoFilter("A1:C4");

            ExcelChart chart = new ExcelChart();
            chart.Row = 6;
            chart.Column = 5;
            chart.AddSeries("$A$1", "$A$3", "$A$3");
            sheet.Charts.Add(chart);

            string png = Path.Combine(Path.GetTempPath(), "office-template-dot.png");
            File.WriteAllBytes(png, CreateTestPng(4, 4));
            ExcelImage image = new ExcelImage(OfficeImage.FromFile(png));
            image.Row = 5;
            image.Column = 8;
            sheet.Images.Add(image);
            return workbook;
        }

        /// <summary>集合与 DataTable 映射往返。Collection and DataTable mapping round trip.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelCollectionTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            List<Person> people = new List<Person>
            {
                new Person { Name = "张三", Age = 30, Hired = new DateTime(2021, 3, 4) },
                new Person { Name = "李四", Age = 26, Hired = new DateTime(2023, 7, 1) }
            };
            ExcelCollection.Write(workbook, people, "People");
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "excel-collection.xlsx"));
            XlsxWriter.Write(workbook, path);

            OfficeWorkbook loaded = XlsxReader.Load(path);
            Worksheet sheet = loaded.Worksheets[0];
            Require(Equals(sheet.GetCell(1, 1).Value, "姓名"), "collection writes header");
            List<Person> read = ExcelCollection.ToList<Person>(sheet);
            Require(read.Count == 2 && read[0].Name == "张三" && read[0].Age == 30, "collection ToList");
            Require(read[1].Hired.Date == new DateTime(2023, 7, 1), "collection date round trip");

            DataTable table = new DataTable("T");
            table.Columns.Add("Code", typeof(string));
            table.Columns.Add("Amount", typeof(double));
            table.Rows.Add("A-1", 12.5d);
            table.Rows.Add("B-2", 99d);
            OfficeWorkbook tableWorkbook = new OfficeWorkbook();
            ExcelCollection.Write(tableWorkbook, table, "Table");
            string tablePath = PrepareOutputPath(Path.Combine(outputDirectory, "excel-datatable.xlsx"));
            XlsxWriter.Write(tableWorkbook, tablePath);
            DataTable back = ExcelCollection.ToDataTable(XlsxReader.Load(tablePath).Worksheets[0]);
            Require(back.Rows.Count == 2 && Equals(back.Rows[0]["Code"], "A-1"), "DataTable round trip");
            Require(Equals(back.Rows[1]["Amount"], 99d), "DataTable numeric round trip");
        }

        /// <summary>CSV 读写往返与转义。CSV read/write round trip and escaping.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelCsvTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Sheet1");
            sheet.GetCell(1, 1).Value = "名字";
            sheet.GetCell(1, 2).Value = "备注";
            sheet.GetCell(2, 1).Value = "张,三";
            sheet.GetCell(2, 2).Value = "第一行\n第二行 \"引号\"";
            sheet.GetCell(3, 1).Value = 42d;
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "excel-csv.csv"));
            ExcelCsv.Write(workbook, path);

            byte[] bytes = File.ReadAllBytes(path);
            Require(bytes.Length > 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF, "CSV writes UTF-8 BOM");
            string text = File.ReadAllText(path);
            RequireContains(text, "\"张,三\"");
            RequireContains(text, "\"第一行\n第二行 \"\"引号\"\"\"");

            OfficeWorkbook read = ExcelCsv.Read(path);
            Require(read.Worksheets.Count == 1, "CSV read creates one sheet");
            Worksheet loaded = read.Worksheets[0];
            Require(Equals(loaded.GetCell(2, 1).Value, "张,三"), "CSV quoted separator round trip");
            Require(Equals(loaded.GetCell(2, 2).Value, "第一行\n第二行 \"引号\""), "CSV quoted newline round trip");
            Require(Equals(loaded.GetCell(3, 1).Value, 42), "CSV numeric inference");
        }

        /// <summary>筛选、超链接、标签色、数据验证与定义名写出并读回。Filter, hyperlink, tab color, data validation and defined names survive a write/read cycle.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelSheetFeatureTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Features");
            sheet.Range("A1:B3").Values = new object[,]
            {
                { "Region", "Amount" },
                { "North", 10d },
                { "South", 20d }
            };
            sheet.AutoFilter = new ExcelAutoFilter("A1:B3");
            sheet.AutoFilter.ForColumn(0).Values.Add("North");
            sheet.TabColor = new OfficeColor(0x1F, 0x3A, 0x5F);

            ExcelDataValidation validation = new ExcelDataValidation();
            validation.Range = "B2:B3";
            validation.Type = DataValidationType.Decimal;
            validation.Operator = DataValidationOperator.GreaterThan;
            validation.Formula1 = "0";
            validation.ShowErrorMessage = true;
            validation.ErrorTitle = "Invalid";
            validation.ErrorMessage = "Amount must be positive";
            sheet.DataValidations.Add(validation);

            Cell link = sheet.GetCell(5, 1);
            link.Value = "Example";
            link.Hyperlink = "https://example.com/page";
            link.HyperlinkTooltip = "open example";
            Cell internalLink = sheet.GetCell(5, 2);
            internalLink.Value = "Top";
            internalLink.Hyperlink = "#Features!A1";
            Cell unsafeLink = sheet.GetCell(5, 3);
            unsafeLink.Value = "Unsafe";
            unsafeLink.Hyperlink = "javascript:alert(1)";
            ExcelImage icon = new ExcelImage(new OfficeImage(CreateTestPng(1, 1)));
            icon.Row = 2;
            icon.Column = 6;
            icon.WidthPoints = 12;
            icon.HeightPoints = 12;
            sheet.Images.Add(icon);
            workbook.DefinedNames.Add(new ExcelDefinedName("Amounts", "Features!$B$2:$B$3"));

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "excel-features.xlsx"));
            XlsxWriter.Write(workbook, path);

            string sheetXml;
            string workbookXml;
            string relationships;
            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                sheetXml = ReadZipEntryText(archive, "xl/worksheets/sheet1.xml");
                workbookXml = ReadZipEntryText(archive, "xl/workbook.xml");
                relationships = ReadZipEntryText(archive, "xl/worksheets/_rels/sheet1.xml.rels");
            }
            RequireContains(sheetXml, "<autoFilter");
            RequireContains(sheetXml, "<hyperlinks");
            RequireContains(sheetXml, "<dataValidations");
            RequireContains(sheetXml, "<tabColor");
            RequireContains(workbookXml, "<definedName");
            RequireContains(relationships, "https://example.com/page");
            RequireContains(relationships, "Id=\"rId1\"");
            RequireContains(relationships, "Id=\"rId2\"");
            Require(relationships.IndexOf("javascript", StringComparison.OrdinalIgnoreCase) < 0, "unsafe hyperlink scheme is rejected");

            OfficeWorkbook loaded = XlsxReader.Load(path);
            Worksheet read = loaded.Worksheets[0];
            Require(read.AutoFilter != null && read.AutoFilter.Range == "A1:B3", "auto filter round trip");
            Require(read.AutoFilter.ForColumn(0).Values.Count == 1
                && read.AutoFilter.ForColumn(0).Values[0] == "North", "auto filter values round trip");
            Require(read.TabColor.HasValue && read.TabColor.Value.ToHex() == "1F3A5F", "tab color round trip");
            Require(read.DataValidations.Count == 1 && read.DataValidations[0].Type == DataValidationType.Decimal
                && read.DataValidations[0].Formula1 == "0", "data validation round trip");
            Require(Equals(read.GetCell(5, 1).Hyperlink, "https://example.com/page"), "external hyperlink round trip");
            Require(Equals(read.GetCell(5, 1).HyperlinkTooltip, "open example"), "hyperlink tooltip round trip");
            Require(Equals(read.GetCell(5, 2).Hyperlink, "#Features!A1"), "internal hyperlink round trip");
            Require(read.GetCell(5, 3).Hyperlink == null, "unsafe hyperlink is dropped");
            Require(read.Images.Count == 1, "image survives along with hyperlinks");
            Require(loaded.DefinedNames.Count == 1 && loaded.DefinedNames[0].Name == "Amounts", "defined name round trip");
        }

        /// <summary>低内存流式查询：行/单元格/字典/对象四种形态及工作表选择。Low-memory streaming query: rows, cells, dictionaries, typed objects and sheet selection.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelQueryTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Query");
            sheet.Range("A1:C1").Values = new object[,] { { "Name", "Score", "Date" } };
            sheet.GetCell(2, 1).Value = "Alice";
            sheet.GetCell(2, 2).Value = 91.5d;
            sheet.GetCell(2, 3).Value = new DateTime(2026, 1, 2);
            sheet.GetCell(2, 3).EnsureStyle().NumberFormat = "yyyy-mm-dd";
            sheet.GetCell(2, 4).Formula = "B2*2";
            sheet.GetCell(3, 1).Value = "Bob";
            sheet.GetCell(3, 2).Value = 78d;
            sheet.GetCell(3, 3).Value = new DateTime(2025, 12, 31);
            sheet.GetCell(3, 3).EnsureStyle().NumberFormat = "yyyy-mm-dd";
            sheet.SetRowHidden(3, true);
            sheet.GetCell(4, 1).Value = "Cara";
            Worksheet noHeader = workbook.AddWorksheet("NoHeader");
            noHeader.GetCell(2, 1).Value = "Solo";

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "excel-query.xlsx"));
            XlsxWriter.Write(workbook, path);

            List<ExcelRowData> rows = new List<ExcelRowData>(ExcelQuery.ReadRows(path, "Query"));
            Require(rows.Count == 4, "query row count");
            Require(rows[1].RowIndex == 2 && Equals(rows[1][1], "Alice"), "query cell by column");
            Require(Equals(rows[1][2], 91.5d), "query numeric cell");
            Require(rows[1][3] is DateTime && ((DateTime)rows[1][3]).Date == new DateTime(2026, 1, 2), "query date cell");
            Require(rows[1].Cells.Count == 4 && rows[1].Cells[3].HasFormula, "query formula cell");
            Require(rows[2].Hidden, "query hidden row flag");
            Require(rows[3].Cells.Count == 1, "query sparse row");
            Require(rows[0].RowIndex == 1 && rows[0][4] == null, "query header row");

            List<ScoreRow> scores = new List<ScoreRow>(ExcelQuery.Query<ScoreRow>(path, "Query"));
            Require(scores.Count == 3 && scores[0].Name == "Alice" && Math.Abs(scores[0].Score - 91.5d) < 1e-9,
                "typed query maps header names");
            Require(scores[1].Name == "Bob" && scores[1].Date.Date == new DateTime(2025, 12, 31), "typed query converts dates");

            List<Dictionary<string, object>> dictionaries =
                new List<Dictionary<string, object>>(ExcelQuery.ReadDictionaries(path, "Query"));
            Require(dictionaries.Count == 3 && Equals(dictionaries[2]["Name"], "Cara"), "dictionary query");
            Require(dictionaries[0].ContainsKey("score"), "dictionary keys are case-insensitive");

            List<Dictionary<string, object>> fallback =
                new List<Dictionary<string, object>>(ExcelQuery.ReadDictionaries(path, "NoHeader"));
            Require(fallback.Count == 1 && Equals(fallback[0]["Column1"], "Solo"),
                "dictionary query falls back to ColumnN when the header row is missing");

            bool streamed = false;
            using (FileStream stream = File.OpenRead(path))
            {
                foreach (ExcelRowData row in ExcelQuery.ReadRows(stream, "Query"))
                {
                    if (row.RowIndex == 2) streamed = Equals(row[1], "Alice");
                }
            }
            Require(streamed, "query from stream");

            byte[] data = File.ReadAllBytes(path);
            int byteRows = 0;
            foreach (ExcelRowData row in ExcelQuery.ReadRows(data, "Query")) byteRows++;
            Require(byteRows == 4, "query from bytes");

            bool missing = false;
            try
            {
                foreach (ExcelRowData row in ExcelQuery.ReadRows(path, "Missing")) { }
            }
            catch (ArgumentException)
            {
                missing = true;
            }
            Require(missing, "query unknown sheet throws ArgumentException");
        }

        /// <summary>读取 zip 条目为文本。Reads a zip entry as text.</summary>
        /// <param name="archive">归档。The archive.</param>
        /// <param name="name">条目名。Entry name.</param>
        /// <returns>条目文本。The entry text.</returns>
        private static string ReadZipEntryText(ZipArchive archive, string name)
        {
            ZipArchiveEntry entry = archive.GetEntry(name);
            if (entry == null) return string.Empty;
            using (Stream stream = entry.Open())
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        /// <summary>集合映射测试用的人员类型。The person type used by collection-mapping tests.</summary>
        private sealed class Person
        {
            /// <summary>姓名。Name.</summary>
            [ExcelColumn("姓名", Order = 1)]
            public string Name { get; set; }
            /// <summary>年龄。Age.</summary>
            [ExcelColumn("年龄", Order = 2)]
            public int Age { get; set; }
            /// <summary>入职日期。Hire date.</summary>
            [ExcelColumn("入职日期", Order = 3, Format = "yyyy-mm-dd")]
            public DateTime Hired { get; set; }
        }

        /// <summary>流式查询测试用的成绩行。The score row used by streaming-query tests.</summary>
        private sealed class ScoreRow
        {
            /// <summary>姓名。Name.</summary>
            public string Name { get; set; }
            /// <summary>成绩。Score.</summary>
            public double Score { get; set; }
            /// <summary>日期。Date.</summary>
            public DateTime Date { get; set; }
        }
    }
}
