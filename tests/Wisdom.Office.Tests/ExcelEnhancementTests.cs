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
