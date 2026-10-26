using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Wisdom.Excel;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        private static int Main(string[] args)
        {
            string outputDirectory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "out"));
            Directory.CreateDirectory(outputDirectory);

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-basic.pdf"));
            RunPdfBasicTest(pdfPath);
            Console.WriteLine("PDF generated: " + pdfPath + " (" + new FileInfo(pdfPath).Length + " bytes)");

            RunPdfParserTest(pdfPath);
            RunPdfWatermarkTest(pdfPath, outputDirectory);
            RunPdfFormTest(pdfPath, outputDirectory);
            RunPdfEncryptionTest(pdfPath, outputDirectory);
            RunPdfHostileInputTest(outputDirectory);

            string fixtureDirectory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "fixtures"));
            string chromePdfPath = Path.Combine(fixtureDirectory, "chrome-print.pdf");
            if (File.Exists(chromePdfPath)) RunExternalPdfParserTest(chromePdfPath);

            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-basic.docx"));
            RunWordRoundTripTest(docxPath);
            ValidateOoxmlPackage(docxPath, "word/document.xml");
            Console.WriteLine("DOCX generated: " + docxPath + " (" + new FileInfo(docxPath).Length + " bytes)");

            string templatePath = PrepareOutputPath(Path.Combine(outputDirectory, "word-template.docx"));
            RunWordTemplateTest(templatePath, PrepareOutputPath(Path.Combine(outputDirectory, "word-template-filled.docx")),
                PrepareOutputPath(Path.Combine(outputDirectory, "word-template-empty.docx")));

            string legacyPath = Path.Combine(fixtureDirectory, "legacy-sample.doc");
            if (File.Exists(legacyPath)) RunLegacyDocTest(legacyPath);

            string xlsxPath = PrepareOutputPath(Path.Combine(outputDirectory, "excel-basic.xlsx"));
            RunExcelRoundTripTest(xlsxPath);
            ValidateOoxmlPackage(xlsxPath, "xl/workbook.xml");
            Console.WriteLine("XLSX generated: " + xlsxPath + " (" + new FileInfo(xlsxPath).Length + " bytes)");

            string xlsxTemplatePath = PrepareOutputPath(Path.Combine(outputDirectory, "excel-template.xlsx"));
            RunExcelTemplateTest(xlsxTemplatePath, PrepareOutputPath(Path.Combine(outputDirectory, "excel-template-filled.xlsx")),
                PrepareOutputPath(Path.Combine(outputDirectory, "excel-template-empty.xlsx")));

            string legacyXlsPath = Path.Combine(fixtureDirectory, "legacy-sample.xls");
            if (File.Exists(legacyXlsPath)) RunLegacyXlsTest(legacyXlsPath);

            string solverPath = Path.Combine(fixtureDirectory, "legacy-solver.xls");
            if (File.Exists(solverPath)) RunLegacyXlsSolverTest(solverPath,
                PrepareOutputPath(Path.Combine(outputDirectory, "excel-legacy-converted.xlsx")));

            RunCompoundFileWriterTest(outputDirectory);

            RunFacadeTest(docxPath, xlsxPath, pdfPath, legacyPath, legacyXlsPath, outputDirectory);
            RunReportTest(outputDirectory);
            RunWordNormalizeTest(outputDirectory);
            RunWordEnhancementTest(outputDirectory);
            RunWordFormatTest(outputDirectory);
            RunWordRevisionTest(outputDirectory);
            RunWordBookmarkTest(outputDirectory);
            RunWordLabelsAndPropertiesTest(outputDirectory);
            RunComplexReportTest(outputDirectory);
            RunComplexExcelTest(outputDirectory);
            RunAttendanceReportTest(outputDirectory);
            string realAttendancePath = Path.Combine(fixtureDirectory, "attendance-real.xlsx");
            if (File.Exists(realAttendancePath)) RunRealAttendanceTest(realAttendancePath, outputDirectory);
            RunFinancialStatementTest(outputDirectory);
            RunDashboardChartTest(outputDirectory);
            RunChartCatalogTest(outputDirectory);
            RunInspectionReportTest(outputDirectory);
            string pptxPath = PrepareOutputPath(Path.Combine(outputDirectory, "pptx-basic.pptx"));
            RunPptxTest(pptxPath, outputDirectory);
            RunPptAnimationTest(outputDirectory);
            RunPptThemeTest(outputDirectory);
            RunPptTemplateTest(outputDirectory);
            RunMarkdownTest();
            RunMarkdownLimitTest(outputDirectory);
            RunMarkdownFidelityTest();
            RunFormulaEngineTest();
            RunExcelEnhancementTest(outputDirectory);
            RunBatchConversionTest(outputDirectory);
            RunOutputHardeningTest(outputDirectory);
            RunExcelGridLimitTest(outputDirectory);
            RunWriteOutputTest(outputDirectory);
            RunResourceLimitTest(outputDirectory);
            RunSecurityTest(outputDirectory, pdfPath, docxPath, xlsxPath);

            Console.WriteLine("All tests finished.");
            return 0;
        }

        private static void RunPdfBasicTest(string path)
        {
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = "Wisdom 报表导出测试";
            document.Metadata.Author = "Wisdom.Office";
            document.Metadata.Subject = "PDF renderer smoke test";

            Section section = document.AddSection();
            Paragraph headerParagraph = section.Header.AddParagraph(string.Empty);
            headerParagraph.Alignment = ParagraphAlignment.Right;
            headerParagraph.AddText("Wisdom.Office 报表  |  第 {page} 页 / 共 {pages} 页", new TextFormat("Microsoft YaHei", 9) { Color = OfficeColor.Gray });

            Paragraph footer = section.Footer.AddParagraph(string.Empty);
            footer.Alignment = ParagraphAlignment.Center;
            footer.AddText("生成时间：2026-10-20    Wisdom.Office 自研文档引擎", new TextFormat("Microsoft YaHei", 9) { Color = OfficeColor.Gray });

            Paragraph title = new Paragraph();
            section.Blocks.Add(title);
            title.Alignment = ParagraphAlignment.Center;
            title.AddText("销售报表 / Sales Report", TextFormat.Heading(1));

            Paragraph subtitle = new Paragraph();
            section.Blocks.Add(subtitle);
            subtitle.Alignment = ParagraphAlignment.Center;
            subtitle.SpaceAfter = 18;
            subtitle.AddText("摘要：本季度销售保持稳健增长，海外市场表现突出。", new TextFormat("Microsoft YaHei", 11) { Color = OfficeColor.Gray });

            Paragraph mixed = new Paragraph();
            section.Blocks.Add(mixed);
            mixed.SpaceAfter = 10;
            mixed.AddText("样式测试：", new TextFormat("Microsoft YaHei", 11) { Bold = true });
            mixed.AddText("粗体 Bold", new TextFormat("Arial", 11) { Bold = true });
            mixed.AddText("，", new TextFormat("Microsoft YaHei", 11));
            mixed.AddText("斜体 Italic", new TextFormat("Arial", 11) { Italic = true });
            mixed.AddText("，", new TextFormat("Microsoft YaHei", 11));
            mixed.AddText("下划线 Underline", new TextFormat("Arial", 11) { Underline = true, Color = OfficeColor.Blue });
            mixed.AddText("，E = mc", new TextFormat("Arial", 11));
            mixed.AddText("2", new TextFormat("Arial", 11) { Superscript = true });
            mixed.AddText("，H", new TextFormat("Arial", 11));
            mixed.AddText("2", new TextFormat("Arial", 11) { Subscript = true });
            mixed.AddText("O。", new TextFormat("Arial", 11));

            Paragraph chinese = new Paragraph();
            section.Blocks.Add(chinese);
            chinese.SpaceAfter = 12;
            chinese.FirstLineIndent = 22;
            chinese.LineSpacing = 1.35;
            chinese.AddText("这是一段用于验证中文换行、标点处理与字体回退的长文本。PDF 引擎需要正确处理中日韩字符的全角宽度，"
                + "并在同一段落内混合 Latin 单词与数字 1234567890。当一行的剩余空间不足时，中文可以逐字换行，"
                + "而英文单词会保持完整，只有在单词本身超过行宽时才会被拆分。", new TextFormat("Microsoft YaHei", 11));

            Paragraph longWord = new Paragraph();
            section.Blocks.Add(longWord);
            longWord.SpaceAfter = 12;
            longWord.AddText("长单词测试：", new TextFormat("Microsoft YaHei", 11));
            longWord.AddText("Pneumonoultramicroscopicsilicovolcanoconiosis", new TextFormat("Arial", 11) { Color = OfficeColor.Red });

            Table table = new Table();
            table.Columns.Add(new TableColumn(50));
            table.Columns.Add(new TableColumn(190));
            table.Columns.Add(new TableColumn(80));
            table.Columns.Add(new TableColumn(120));
            TableRow header = table.AddRow("序号", "产品名称", "数量", "金额（元）");
            header.IsHeader = true;
            for (int i = 0; i < header.Cells.Count; i++)
            {
                header.Cells[i].BackgroundColor = new OfficeColor(0x1F, 0x3A, 0x5F);
                Paragraph cellParagraph = (Paragraph)header.Cells[i].Blocks[0];
                Paragraph styled = cellParagraph.Clone();
                styled.Alignment = i == 0 ? ParagraphAlignment.Center : ParagraphAlignment.Left;
                foreach (Inline inline in styled.Inlines)
                {
                    TextRun run = inline as TextRun;
                    if (run != null) run.Format = new TextFormat("Microsoft YaHei", 10) { Bold = true, Color = OfficeColor.White };
                }
                header.Cells[i].Blocks.Clear();
                header.Cells[i].Blocks.Add(styled);
            }

            string[] products = { "无线键盘", "机械鼠标", "27 英寸显示器", "USB-C 扩展坞", "降噪耳机", "笔记本电脑支架" };
            for (int i = 0; i < 32; i++)
            {
                TableRow row = table.AddRow(
                    (i + 1).ToString(),
                    products[i % products.Length] + " Pro",
                    (15 + i * 3).ToString(),
                    (1280.5 + i * 217.35).ToString("N2"));
                if (i % 2 == 1)
                {
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        row.Cells[c].BackgroundColor = new OfficeColor(0xF3, 0xF6, 0xFA);
                    }
                }
            }
            section.Blocks.Add(table);

            Paragraph afterTable = new Paragraph();
            section.Blocks.Add(afterTable);
            afterTable.SpaceBefore = 10;
            afterTable.Alignment = ParagraphAlignment.Right;
            afterTable.AddText("合计：￥ 128,640.00", new TextFormat("Microsoft YaHei", 11) { Bold = true, Color = new OfficeColor(0xC0, 0x39, 0x2B) });

            Paragraph imageParagraph = new Paragraph();
            section.Blocks.Add(imageParagraph);
            imageParagraph.SpaceBefore = 16;
            imageParagraph.Alignment = ParagraphAlignment.Center;
            imageParagraph.AddImage(new OfficeImage(CreateTestPng(320, 120)), 240, 90);

            Paragraph tail = new Paragraph();
            section.Blocks.Add(tail);
            tail.SpaceBefore = 8;
            tail.Alignment = ParagraphAlignment.Justify;
            tail.AddText("本段用于验证两端对齐（Justify）：The quick brown fox jumps over the lazy dog. 通过拉伸空格来对齐左右边界，"
                + "最后一行保持自然宽度。", new TextFormat("Microsoft YaHei", 11));

            PdfRenderer renderer = new PdfRenderer();
            renderer.Render(document, path);
        }

        private static void RunPdfParserTest(string path)
        {
            PdfDocument pdf = PdfDocument.Load(path);
            Console.WriteLine("[parser] pages=" + pdf.PageCount + ", title=" + pdf.Info.Title
                + ", author=" + pdf.Info.Author + ", created=" + (pdf.Info.CreationDate.HasValue ? pdf.Info.CreationDate.Value.ToString("yyyy-MM-dd") : "?"));
            Require(pdf.PageCount > 1, "expected a multi-page PDF");
            Require(pdf.Info.Title == "Wisdom 报表导出测试", "title mismatch: " + pdf.Info.Title);
            Require(pdf.Info.Author == "Wisdom.Office", "author mismatch: " + pdf.Info.Author);
            Require(pdf.Info.CreationDate.HasValue, "creation date missing");

            string text = pdf.ExtractText();
            RequireContains(text, "销售报表 / Sales Report");
            RequireContains(text, "摘要：本季度销售保持稳健增长，海外市场表现突出。");
            RequireContains(text, "Pneumonoultramicroscopicsilicovolcanoconiosis");
            RequireContains(text, "合计：￥ 128,640.00");
            RequireContains(text, "第 1 页 / 共 " + pdf.PageCount + " 页");
            RequireContains(text, "第 2 页 / 共 " + pdf.PageCount + " 页");
            RequireContains(text, "生成时间：2026-10-20");

            int totalTables = 0;
            for (int i = 0; i < pdf.PageCount; i++)
            {
                PdfPage page = pdf.Pages[i];
                Console.WriteLine("[parser] page " + (i + 1) + ": lines=" + page.GetTextLines().Count
                    + ", tables=" + page.GetTables().Count + ", size=" + page.Width.ToString("0.#") + "x" + page.Height.ToString("0.#"));
                totalTables += page.GetTables().Count;
            }
            Require(totalTables > 0, "no tables detected");

            PdfTable table = null;
            for (int i = 0; i < pdf.PageCount && table == null; i++)
            {
                if (pdf.Pages[i].GetTables().Count > 0) table = pdf.Pages[i].GetTables()[0];
            }
            Console.WriteLine("[parser] table " + table.RowCount + "x" + table.ColumnCount
                + " at " + table.X.ToString("0.#") + "," + table.Y.ToString("0.#"));
            Console.WriteLine("[parser] row 0: " + table.GetCell(0, 0) + " | " + table.GetCell(0, 1)
                + " | " + table.GetCell(0, 2) + " | " + table.GetCell(0, 3));
            Console.WriteLine("[parser] row 1: " + table.GetCell(1, 0) + " | " + table.GetCell(1, 1)
                + " | " + table.GetCell(1, 2) + " | " + table.GetCell(1, 3));
            Require(table.ColumnCount == 4, "expected 4 columns");
            Require(table.GetCell(0, 0) == "序号", "header column 0 mismatch: " + table.GetCell(0, 0));
            Require(table.GetCell(0, 1) == "产品名称", "header column 1 mismatch: " + table.GetCell(0, 1));
            Require(table.GetCell(1, 0) == "1", "row 1 column 0 mismatch: " + table.GetCell(1, 0));
            RequireContains(table.GetCell(1, 1), "无线键盘");

            PdfTable continuation = pdf.Pages[pdf.PageCount - 1].GetTables()[0];
            Require(continuation.GetCell(0, 0) == "序号", "continuation header not repeated: " + continuation.GetCell(0, 0));
            int lastRow = continuation.RowCount - 1;
            Console.WriteLine("[parser] continuation last row: " + continuation.GetCell(lastRow, 0) + " | "
                + continuation.GetCell(lastRow, 1) + " | " + continuation.GetCell(lastRow, 2) + " | " + continuation.GetCell(lastRow, 3));
            Require(continuation.GetCell(lastRow, 0) == "32", "last row index mismatch: " + continuation.GetCell(lastRow, 0));
            RequireContains(continuation.GetCell(lastRow, 3), "8,018.35");

            OfficeDocument converted = pdf.ToOfficeDocument();
            Require(converted.Sections.Count == pdf.PageCount, "converted section count mismatch");
            Require(converted.Sections[0].Blocks.Count > 0, "converted document has no content");

            Console.WriteLine("[parser] extracted " + text.Length + " characters, " + totalTables + " table(s)");
        }

        private static void RunExternalPdfParserTest(string path)
        {
            PdfDocument pdf = PdfDocument.Load(path);
            Console.WriteLine("[external] pages=" + pdf.PageCount + ", producer=" + pdf.Info.Producer);
            string text = pdf.ExtractText();
            RequireContains(text, "Quarterly Operations Review");
            RequireContains(text, "APAC / 亚太区");
            RequireContains(text, "728,940.25");

            PdfTable table = null;
            for (int i = 0; i < pdf.PageCount && table == null; i++)
            {
                for (int t = 0; t < pdf.Pages[i].GetTables().Count; t++)
                {
                    PdfTable candidate = pdf.Pages[i].GetTables()[t];
                    if (candidate.RowCount >= 5) { table = candidate; break; }
                }
            }
            Require(table != null, "external PDF table not detected");
            Console.WriteLine("[external] table " + table.RowCount + "x" + table.ColumnCount + ": "
                + table.GetCell(0, 0) + " | " + table.GetCell(0, 1) + " | " + table.GetCell(0, 2));
            Console.WriteLine("[external] row 3: " + table.GetCell(3, 0) + " | " + table.GetCell(3, 1) + " | " + table.GetCell(3, 2));
            Require(table.ColumnCount == 3, "external table column count mismatch: " + table.ColumnCount);
            Require(table.GetCell(1, 0) == "North America", "external cell mismatch: " + table.GetCell(1, 0));
            RequireContains(table.GetCell(3, 2), "728,940.25");
        }

        private static void RunWordRoundTripTest(string path)
        {
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = "Wisdom Word 读写测试";
            document.Metadata.Author = "Wisdom.Office";
            document.Metadata.Subject = "DOCX round trip";
            document.Metadata.Keywords = "word, docx, wisom";
            document.Metadata.Created = new DateTime(2026, 10, 20, 8, 30, 0, DateTimeKind.Utc);
            document.Metadata.Modified = new DateTime(2026, 10, 20, 9, 45, 0, DateTimeKind.Utc);

            Section section = document.AddSection();
            section.PageSetup.MarginLeft = 60;
            section.PageSetup.MarginRight = 60;
            section.PageSetup.MarginTop = 70;
            section.PageSetup.MarginBottom = 70;
            section.PageSetup.HeaderDistance = 30;
            section.PageSetup.FooterDistance = 30;

            Paragraph header = section.Header.AddParagraph("页眉：Wisdom.Office 文档引擎");
            header.Alignment = ParagraphAlignment.Right;
            header.AddText(string.Empty, new TextFormat("Microsoft YaHei", 9) { Color = OfficeColor.Gray });
            Paragraph footer = section.Footer.AddParagraph("页脚：内部资料");
            footer.Alignment = ParagraphAlignment.Center;

            Paragraph heading = new Paragraph();
            heading.HeadingLevel = 1;
            heading.SpaceAfter = 14;
            heading.AddText("第一章 概述", new TextFormat("Microsoft YaHei", 20) { Bold = true, Color = new OfficeColor(0x1F, 0x3A, 0x5F) });
            section.Blocks.Add(heading);

            Paragraph mixed = new Paragraph();
            mixed.SpaceAfter = 10;
            mixed.AddText("中文正文测试，", new TextFormat("Microsoft YaHei", 11));
            mixed.AddText("加粗 Bold", new TextFormat("Microsoft YaHei", 11) { Bold = true });
            mixed.AddText("，", new TextFormat("Microsoft YaHei", 11));
            mixed.AddText("斜体 Italic", new TextFormat("Microsoft YaHei", 11) { Italic = true, Color = OfficeColor.Blue });
            mixed.AddText("，下划线内容。", new TextFormat("Microsoft YaHei", 11) { Underline = true });
            section.Blocks.Add(mixed);

            Paragraph multiline = new Paragraph();
            multiline.SpaceAfter = 8;
            multiline.AddText("第一行\n第二行\t制表位后的文字", new TextFormat("Microsoft YaHei", 10));
            section.Blocks.Add(multiline);

            Paragraph imageParagraph = new Paragraph();
            imageParagraph.Alignment = ParagraphAlignment.Center;
            imageParagraph.SpaceAfter = 12;
            imageParagraph.AddImage(new OfficeImage(CreateTestPng(200, 80)), 200, 80);
            section.Blocks.Add(imageParagraph);

            Table table = new Table();
            table.Columns.Add(new TableColumn(60));
            table.Columns.Add(new TableColumn(170));
            table.Columns.Add(new TableColumn(110));
            TableRow tableHeader = table.AddRow("序号", "项目", "金额");
            tableHeader.IsHeader = true;
            string[] headerValues = { "序号", "项目", "金额" };
            for (int i = 0; i < tableHeader.Cells.Count; i++)
            {
                tableHeader.Cells[i].BackgroundColor = new OfficeColor(0x1F, 0x3A, 0x5F);
                tableHeader.Cells[i].Blocks.Clear();
                Paragraph cellParagraph = new Paragraph();
                cellParagraph.AddText(headerValues[i], new TextFormat("Microsoft YaHei", 10) { Bold = true, Color = OfficeColor.White });
                tableHeader.Cells[i].Blocks.Add(cellParagraph);
            }
            TableRow mergedRow = table.AddRow("合并行", "跨两列");
            mergedRow.Cells[0].RowSpan = 2;
            mergedRow.Cells[1].ColSpan = 2;
            TableRow mergeContinue = table.AddRow("Y");
            TableRow lastRow = table.AddRow("1", "测试项目 A", "10.00");
            section.Blocks.Add(table);

            Paragraph closing = new Paragraph();
            closing.SpaceBefore = 12;
            closing.AddText("以上为全部数据。", new TextFormat("Microsoft YaHei", 10) { Italic = true });
            section.Blocks.Add(closing);

            PageSetup landscape = new PageSetup();
            landscape.Width = 792;
            landscape.Height = 612;
            landscape.MarginLeft = 50;
            landscape.MarginRight = 50;
            landscape.MarginTop = 50;
            landscape.MarginBottom = 50;
            Section appendix = document.AddSection(landscape);
            Paragraph appendixHeading = new Paragraph();
            appendixHeading.HeadingLevel = 1;
            appendixHeading.AddText("附录 A：横向页面", new TextFormat("Microsoft YaHei", 18) { Bold = true });
            appendix.Blocks.Add(appendixHeading);
            appendix.Blocks.Add(new PageBreak());
            appendix.AddParagraph("横向页第二页内容。");

            DocxWriter.Write(document, path);
            Require(new FileInfo(path).Length > 2000, "docx output too small");

            OfficeDocument loaded = DocxReader.Load(path);
            Require(loaded.Sections.Count == 2, "docx section count mismatch: " + loaded.Sections.Count);
            Require(loaded.Metadata.Title == "Wisdom Word 读写测试", "docx title mismatch: " + loaded.Metadata.Title);
            Require(loaded.Metadata.Author == "Wisdom.Office", "docx author mismatch: " + loaded.Metadata.Author);
            Require(loaded.Metadata.Keywords == "word, docx, wisom", "docx keywords mismatch: " + loaded.Metadata.Keywords);
            Require(loaded.Metadata.Created.Year == 2026 && loaded.Metadata.Created.Month == 10
                && loaded.Metadata.Created.Day == 20 && loaded.Metadata.Created.Hour == 8,
                "docx created timestamp mismatch: " + loaded.Metadata.Created.ToString("o"));

            string headerText = GetBlocksText(loaded.Sections[0].Header.Blocks);
            RequireContains(headerText, "页眉：Wisdom.Office 文档引擎");
            RequireContains(GetBlocksText(loaded.Sections[0].Footer.Blocks), "页脚：内部资料");
            RequireContains(GetBlocksText(loaded.Sections[1].Header.Blocks), "页眉：Wisdom.Office 文档引擎");

            Require(Math.Abs(loaded.Sections[0].PageSetup.Width - 595.28) < 1.0,
                "A4 width mismatch: " + loaded.Sections[0].PageSetup.Width.ToString("0.##"));
            Require(loaded.Sections[1].PageSetup.Width > loaded.Sections[1].PageSetup.Height,
                "landscape page setup not preserved");
            Require(Math.Abs(loaded.Sections[1].PageSetup.Width - 792) < 1.0, "landscape width mismatch");

            Paragraph loadedHeading = FindParagraph(loaded.Sections[0].Blocks, "第一章 概述");
            Require(loadedHeading != null, "heading paragraph missing");
            Require(loadedHeading.HeadingLevel == 1, "heading level mismatch: " + loadedHeading.HeadingLevel);

            Paragraph loadedMixed = FindParagraph(loaded.Sections[0].Blocks, "加粗 Bold");
            Require(loadedMixed != null, "mixed paragraph missing");
            bool hasBold = false;
            bool hasItalic = false;
            bool hasUnderline = false;
            for (int i = 0; i < loadedMixed.Inlines.Count; i++)
            {
                TextRun run = loadedMixed.Inlines[i] as TextRun;
                if (run == null) continue;
                if (run.Text.IndexOf("加粗", StringComparison.Ordinal) >= 0 && run.Format.Bold) hasBold = true;
                if (run.Text.IndexOf("斜体", StringComparison.Ordinal) >= 0 && run.Format.Italic) hasItalic = true;
                if (run.Text.IndexOf("下划线", StringComparison.Ordinal) >= 0 && run.Format.Underline) hasUnderline = true;
            }
            Require(hasBold && hasItalic && hasUnderline, "run formatting not preserved");

            Paragraph loadedMultiline = FindParagraph(loaded.Sections[0].Blocks, "制表位");
            Require(loadedMultiline != null && loadedMultiline.GetText().Contains("\n") && loadedMultiline.GetText().Contains("\t"),
                "line break or tab not preserved");

            bool hasImage = false;
            foreach (Block block in loaded.Sections[0].Blocks)
            {
                Paragraph paragraph = block as Paragraph;
                if (paragraph == null) continue;
                for (int i = 0; i < paragraph.Inlines.Count; i++)
                {
                    ImageRun image = paragraph.Inlines[i] as ImageRun;
                    if (image != null && image.Image.PixelWidth == 200 && image.Image.PixelHeight == 80
                        && Math.Abs(image.Width - 200) < 1)
                    {
                        hasImage = true;
                    }
                }
            }
            Require(hasImage, "image extent not preserved");

            Table loadedTable = FindFirstTable(loaded.Sections[0].Blocks);
            Require(loadedTable != null, "table missing after read");
            Require(loadedTable.Rows.Count == 4, "table row count mismatch: " + loadedTable.Rows.Count);
            Require(loadedTable.Rows[0].IsHeader, "table header flag missing");
            Require(loadedTable.Rows[0].Cells[0].BackgroundColor.HasValue, "table header shading missing");
            Require(loadedTable.Rows[1].Cells[0].RowSpan == 2, "vertical merge not preserved: " + loadedTable.Rows[1].Cells[0].RowSpan);
            Require(loadedTable.Rows[1].Cells[1].ColSpan == 2, "horizontal merge not preserved: " + loadedTable.Rows[1].Cells[1].ColSpan);
            RequireContains(GetBlocksText(loadedTable.Rows[1].Cells[0].Blocks), "合并行");
            RequireContains(GetBlocksText(loadedTable.Rows[2].Cells[0].Blocks), "Y");
            Require(loadedTable.Borders.Top.Style == BorderStyle.Single, "table border style missing");

            bool hasPageBreak = false;
            for (int i = 0; i < loaded.Sections[1].Blocks.Count; i++)
            {
                if (loaded.Sections[1].Blocks[i] is PageBreak) hasPageBreak = true;
            }
            Require(hasPageBreak, "explicit page break not preserved");

            string roundTripPath = PrepareOutputPath(Path.Combine(Path.GetDirectoryName(path), "word-roundtrip.docx"));
            DocxWriter.Write(loaded, roundTripPath);
            OfficeDocument reloaded = DocxReader.Load(roundTripPath);
            Require(reloaded.Sections.Count == 2, "round-trip section count mismatch");
            Require(reloaded.Metadata.Title == "Wisdom Word 读写测试", "round-trip title mismatch");
            Table roundTripTable = FindFirstTable(reloaded.Sections[0].Blocks);
            Require(roundTripTable != null && roundTripTable.Rows.Count == 4, "round-trip table mismatch");
            RequireContains(GetBlocksText(reloaded.Sections[0].Blocks), "以上为全部数据。");
            RequireContains(GetBlocksText(reloaded.Sections[1].Header.Blocks), "页眉");
            Console.WriteLine("[docx] sections=" + loaded.Sections.Count + ", table=" + loadedTable.Rows.Count + "x"
                + loadedTable.Columns.Count + ", header='" + headerText + "', roundtrip=" + new FileInfo(roundTripPath).Length + " bytes");
        }

        private static void RunWordTemplateTest(string templatePath, string filledPath, string emptyPath)
        {
            OfficeDocument template = new OfficeDocument();
            Section section = template.AddSection();
            section.Header.AddParagraph("{{reportTitle}}");

            Paragraph greeting = new Paragraph();
            greeting.SpaceAfter = 10;
            TextFormat greetingFormat = new TextFormat("Microsoft YaHei", 12) { Bold = true };
            greeting.AddText("客户：", greetingFormat);
            greeting.AddText("{{cus", greetingFormat);
            greeting.AddText("tomerName}}", greetingFormat);
            section.Blocks.Add(greeting);

            Table table = new Table();
            table.Columns.Add(new TableColumn(200));
            table.Columns.Add(new TableColumn(120));
            TableRow tableHeader = table.AddRow("名称", "金额");
            tableHeader.IsHeader = true;
            TableRow loopRow = table.AddRow(string.Empty, string.Empty);
            loopRow.Cells[0].Blocks.Clear();
            Paragraph nameCell = new Paragraph();
            nameCell.AddText("{{#each it", new TextFormat("Microsoft YaHei", 10));
            nameCell.AddText("ems}}{{name}}", new TextFormat("Microsoft YaHei", 10));
            loopRow.Cells[0].Blocks.Add(nameCell);
            loopRow.Cells[1].Blocks.Clear();
            Paragraph amountCell = new Paragraph();
            amountCell.AddText("{{amount:0.00}}{{/each}}", new TextFormat("Microsoft YaHei", 10) { Bold = true });
            loopRow.Cells[1].Blocks.Add(amountCell);
            section.Blocks.Add(table);

            section.AddParagraph("备注：");
            section.AddParagraph("{{#each notes}}");
            section.AddParagraph("- {{text}}");
            section.AddParagraph("{{/each}}");
            section.AddParagraph("共 {{items.count}} 项");

            DocxWriter.Write(template, templatePath);

            TemplateContext context = new TemplateContext();
            context.Set("reportTitle", "2026 年第三季度报告");
            context.Set("customerName", "华东供应链有限公司");
            List<Dictionary<string, object>> items = new List<Dictionary<string, object>>();
            items.Add(new Dictionary<string, object> { { "name", "物料 A" }, { "amount", 1250.5m } });
            items.Add(new Dictionary<string, object> { { "name", "物料 B" }, { "amount", 880m } });
            items.Add(new Dictionary<string, object> { { "name", "物料 C" }, { "amount", 12.345m } });
            context.Set("items", items);
            List<Dictionary<string, object>> notes = new List<Dictionary<string, object>>();
            notes.Add(new Dictionary<string, object> { { "text", "第一条备注" } });
            notes.Add(new Dictionary<string, object> { { "text", "第二条备注" } });
            notes.Add(new Dictionary<string, object> { { "text", "第三条备注" } });
            context.Set("notes", notes);

            DocxTemplate.Fill(templatePath, filledPath, context);
            OfficeDocument filled = DocxReader.Load(filledPath);
            string filledText = filled.GetPlainText();
            RequireContains(GetBlocksText(filled.Sections[0].Header.Blocks), "2026 年第三季度报告");
            RequireContains(filledText, "客户：华东供应链有限公司");
            RequireContains(filledText, "物料 A");
            RequireContains(filledText, "1250.50");
            RequireContains(filledText, "12.35");
            RequireContains(filledText, "第一条备注");
            RequireContains(filledText, "第三条备注");
            RequireContains(filledText, "共 3 项");
            Require(filledText.IndexOf("{{", StringComparison.Ordinal) < 0, "unresolved template tokens remain");

            Paragraph filledGreeting = FindParagraph(filled.Sections[0].Blocks, "客户：");
            Require(filledGreeting != null, "filled greeting missing");
            bool tokenFormatting = false;
            for (int i = 0; i < filledGreeting.Inlines.Count; i++)
            {
                TextRun run = filledGreeting.Inlines[i] as TextRun;
                if (run != null && run.Text.IndexOf("华东供应链", StringComparison.Ordinal) >= 0 && run.Format.Bold) tokenFormatting = true;
            }
            Require(tokenFormatting, "template replacement lost run formatting");

            Table filledTable = FindFirstTable(filled.Sections[0].Blocks);
            Require(filledTable != null, "filled table missing");
            Require(filledTable.Rows.Count == 4, "template row loop count mismatch: " + filledTable.Rows.Count);
            RequireContains(GetBlocksText(filledTable.Rows[1].Cells[0].Blocks), "物料 A");
            RequireContains(GetBlocksText(filledTable.Rows[3].Cells[1].Blocks), "12.35");

            int noteCount = 0;
            for (int i = 0; i < filled.Sections[0].Blocks.Count; i++)
            {
                Paragraph paragraph = filled.Sections[0].Blocks[i] as Paragraph;
                if (paragraph != null && paragraph.GetText().StartsWith("- ", StringComparison.Ordinal)) noteCount++;
            }
            Require(noteCount == 3, "paragraph loop count mismatch: " + noteCount);

            TemplateContext empty = new TemplateContext();
            empty.Set("reportTitle", "空数据报告");
            empty.Set("customerName", "无");
            empty.Set("items", new List<object>());
            empty.Set("notes", new List<object>());
            DocxTemplate.Fill(templatePath, emptyPath, empty);
            OfficeDocument emptyDocument = DocxReader.Load(emptyPath);
            string emptyText = emptyDocument.GetPlainText();
            Require(emptyText.IndexOf("{{", StringComparison.Ordinal) < 0, "empty template left markers");
            Require(emptyText.IndexOf("- ", StringComparison.Ordinal) < 0, "empty paragraph loop was not removed");
            Table emptyTable = FindFirstTable(emptyDocument.Sections[0].Blocks);
            Require(emptyTable != null && emptyTable.Rows.Count == 1, "empty row loop was not removed");
            Console.WriteLine("[docx-template] filled=" + new FileInfo(filledPath).Length + " bytes, rows="
                + filledTable.Rows.Count + ", notes=" + noteCount + ", empty rows=" + (emptyTable == null ? -1 : emptyTable.Rows.Count));
        }

        private static void RunLegacyDocTest(string path)
        {
            OfficeDocument document = DocReader.Load(path);
            string text = document.GetPlainText();
            Require(document.Sections.Count > 0, "legacy doc has no sections");
            RequireContains(text, "单位");
            RequireContains(text, "电子邮件");
            RequireContains(text, "地址（英文）");
            Require(document.Metadata.Title == "名片地址", "legacy doc title mismatch: " + document.Metadata.Title);
            RequireContains(document.Metadata.Author, "微软");
            int tables = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                for (int b = 0; b < document.Sections[s].Blocks.Count; b++)
                {
                    if (document.Sections[s].Blocks[b] is Table) tables++;
                }
            }
            Require(tables > 0, "legacy doc table missing");
            Console.WriteLine("[doc] title='" + document.Metadata.Title + "', author='" + document.Metadata.Author
                + "', chars=" + text.Length + ", tables=" + tables);
        }

        private static void RunExcelRoundTripTest(string path)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.Metadata.Title = "Wisdom Excel 读写测试";
            workbook.Metadata.Author = "Wisdom.Office";
            workbook.Metadata.Subject = "XLSX round trip";
            workbook.Metadata.Keywords = "excel, xlsx, wisom";
            workbook.Metadata.Created = new DateTime(2026, 10, 20, 8, 30, 0, DateTimeKind.Utc);
            workbook.Metadata.Modified = new DateTime(2026, 10, 20, 9, 45, 0, DateTimeKind.Utc);

            Worksheet sheet = workbook.AddWorksheet("销售数据");
            sheet.ShowGridLines = false;
            sheet.FrozenRows = 1;
            sheet.FrozenColumns = 1;
            sheet.FitToPage = true;
            sheet.SetColumnWidth(1, 18.5);
            sheet.SetColumnWidth(2, 12);
            sheet.SetColumnWidth(3, 14);
            sheet.SetColumnWidth(4, 16);
            sheet.SetRowHeight(1, 26);
            sheet.PrintSetup.MarginLeft = 54;
            sheet.PrintSetup.MarginTop = 54;

            CellStyle headerStyle = new CellStyle();
            headerStyle.FontFamily = "Microsoft YaHei";
            headerStyle.FontSize = 11;
            headerStyle.Bold = true;
            headerStyle.FontColor = OfficeColor.White;
            headerStyle.FillColor = new OfficeColor(0x1F, 0x3A, 0x5F);
            headerStyle.HorizontalAlignment = HorizontalAlignment.Center;
            headerStyle.VerticalAlignment = VerticalAlignment.Middle;
            headerStyle.WrapText = true;
            headerStyle.BottomBorder = new BorderLine(BorderStyle.Single, 0.5, OfficeColor.Black);

            string[] headers = { "产品", "数量", "单价", "日期" };
            for (int c = 0; c < headers.Length; c++)
            {
                Cell cell = sheet.GetCell(1, c + 1);
                cell.Value = headers[c];
                cell.Style = headerStyle.Clone();
            }

            CellStyle moneyStyle = new CellStyle();
            moneyStyle.NumberFormat = "#,##0.00";
            moneyStyle.TopBorder = BorderLine.Thin();

            CellStyle dateStyle = new CellStyle();
            dateStyle.NumberFormat = "yyyy-mm-dd";
            dateStyle.HorizontalAlignment = HorizontalAlignment.Center;

            string[] products = { "无线键盘", "机械鼠标", "27 英寸显示器" };
            double[] quantities = { 12, 30, 5 };
            double[] prices = { 299.5, 89.9, 1899 };
            DateTime[] dates = { new DateTime(2026, 10, 1), new DateTime(2026, 10, 2), new DateTime(2026, 10, 3) };
            for (int i = 0; i < products.Length; i++)
            {
                int row = i + 2;
                sheet.GetCell(row, 1).Value = products[i];
                sheet.GetCell(row, 2).Value = quantities[i];
                Cell price = sheet.GetCell(row, 3);
                price.Value = prices[i];
                price.Style = moneyStyle.Clone();
                Cell date = sheet.GetCell(row, 4);
                date.Value = dates[i];
                date.Style = dateStyle.Clone();
            }

            sheet.Merge("A5:B5");
            Cell totalLabel = sheet.GetCell(5, 1);
            totalLabel.Value = "合计";
            totalLabel.Style = new CellStyle();
            totalLabel.Style.Bold = true;
            totalLabel.Style.HorizontalAlignment = HorizontalAlignment.Right;
            Cell totalValue = sheet.GetCell(5, 3);
            totalValue.Formula = "SUM(C2:C4)";
            totalValue.Value = 2288.4;
            totalValue.Style = moneyStyle.Clone();
            totalValue.Style.Bold = true;

            sheet.GetCell(6, 1).Value = "已审核";
            sheet.GetCell(6, 2).Value = true;

            sheet.Merge("A7:D7");
            Cell note = sheet.GetCell(7, 1);
            note.Value = "备注：本季度价格含税。";
            CellStyle noteStyle = new CellStyle();
            noteStyle.Italic = true;
            noteStyle.FillColor = OfficeColor.LightGray;
            note.Style = noteStyle;

            Worksheet second = workbook.AddWorksheet("说明");
            second.GetCell(1, 1).Value = "本工作簿由 Wisdom.Excel 生成。";
            second.GetCell(2, 1).Value = 42.5;
            second.SetColumnWidth(1, 40);

            XlsxWriter.Write(workbook, path);
            Require(new FileInfo(path).Length > 2000, "xlsx output too small");

            OfficeWorkbook loaded = XlsxReader.Load(path);
            Require(loaded.Worksheets.Count == 2, "xlsx worksheet count mismatch: " + loaded.Worksheets.Count);
            Require(loaded.Worksheets[0].Name == "销售数据", "xlsx sheet name mismatch");
            Require(loaded.Worksheets[1].Name == "说明", "xlsx second sheet name mismatch");
            Require(loaded.Metadata.Title == "Wisdom Excel 读写测试", "xlsx title mismatch: " + loaded.Metadata.Title);
            Require(loaded.Metadata.Author == "Wisdom.Office", "xlsx author mismatch: " + loaded.Metadata.Author);
            Require(loaded.Metadata.Keywords == "excel, xlsx, wisom", "xlsx keywords mismatch");
            Require(loaded.Metadata.Created.Year == 2026 && loaded.Metadata.Created.Month == 10
                && loaded.Metadata.Created.Day == 20 && loaded.Metadata.Created.Hour == 8,
                "xlsx created timestamp mismatch: " + loaded.Metadata.Created.ToString("o"));

            Worksheet loadedSheet = loaded.Worksheets[0];
            Require(!loadedSheet.ShowGridLines, "xlsx gridline flag not preserved");
            Require(loadedSheet.FrozenRows == 1 && loadedSheet.FrozenColumns == 1, "xlsx freeze panes not preserved");
            Require(loadedSheet.FitToPage, "xlsx fit-to-page not preserved");
            Require(Math.Abs(loadedSheet.GetColumnWidth(1) - 18.5) < 0.01, "xlsx column width mismatch: " + loadedSheet.GetColumnWidth(1));
            Require(Math.Abs(loadedSheet.GetRowHeight(1) - 26) < 0.01, "xlsx row height mismatch");
            Require(Math.Abs(loadedSheet.PrintSetup.MarginLeft - 54) < 0.5, "xlsx margin mismatch");

            Cell loadedHeader = loadedSheet.GetCell(1, 1);
            Require(loadedHeader.Value as string == "产品", "xlsx header value mismatch");
            Require(loadedHeader.Style != null && loadedHeader.Style.Bold, "xlsx header bold missing");
            Require(loadedHeader.Style.FontFamily == "Microsoft YaHei", "xlsx header font mismatch: " + loadedHeader.Style.FontFamily);
            RequireColor(loadedHeader.Style.FontColor, OfficeColor.White, "xlsx header font color");
            Require(loadedHeader.Style.FillColor.HasValue, "xlsx header fill missing");
            RequireColor(loadedHeader.Style.FillColor.Value, new OfficeColor(0x1F, 0x3A, 0x5F), "xlsx header fill");
            Require(loadedHeader.Style.HorizontalAlignment == HorizontalAlignment.Center, "xlsx header alignment mismatch");
            Require(loadedHeader.Style.WrapText, "xlsx header wrap flag missing");
            Require(loadedHeader.Style.BottomBorder != null && loadedHeader.Style.BottomBorder.Style == BorderStyle.Single,
                "xlsx header border missing");

            Require(Convert.ToDouble(loadedSheet.GetCell(2, 2).Value) == 12, "xlsx number mismatch");
            Require(Math.Abs(Convert.ToDouble(loadedSheet.GetCell(2, 3).Value) - 299.5) < 0.0001, "xlsx decimal mismatch");
            Require(loadedSheet.GetCell(2, 3).Style.NumberFormat == "#,##0.00", "xlsx number format mismatch: "
                + loadedSheet.GetCell(2, 3).Style.NumberFormat);
            Require(loadedSheet.GetCell(2, 3).Style.TopBorder != null
                && loadedSheet.GetCell(2, 3).Style.TopBorder.Style == BorderStyle.Single, "xlsx cell border missing");

            object dateValue = loadedSheet.GetCell(2, 4).Value;
            Require(dateValue is DateTime, "xlsx date cell did not round-trip: " + (dateValue == null ? "null" : dateValue.GetType().Name));
            Require(((DateTime)dateValue).Date == new DateTime(2026, 10, 1), "xlsx date mismatch");
            Require(loadedSheet.GetCell(2, 4).Style.IsDateFormat(), "xlsx date format not preserved");

            Cell loadedFormula = loadedSheet.GetCell(5, 3);
            Require(loadedFormula.Formula == "SUM(C2:C4)", "xlsx formula mismatch: " + loadedFormula.Formula);
            Require(Math.Abs(Convert.ToDouble(loadedFormula.Value) - 2288.4) < 0.0001, "xlsx formula cached value mismatch");
            Require(loadedSheet.GetCell(6, 2).Value is bool && (bool)loadedSheet.GetCell(6, 2).Value, "xlsx boolean mismatch");

            Require(HasMerge(loadedSheet, 5, 1, 1, 2), "xlsx merged range A5:B5 missing");
            Require(HasMerge(loadedSheet, 7, 1, 1, 4), "xlsx merged range A7:D7 missing");
            Require(loadedSheet.GetCell(7, 1).Style != null && loadedSheet.GetCell(7, 1).Style.Italic, "xlsx italic style missing");
            Require(loadedSheet.GetCell(7, 1).Style.FillColor.HasValue, "xlsx note fill missing");
            RequireColor(loadedSheet.GetCell(7, 1).Style.FillColor.Value, OfficeColor.LightGray, "xlsx note fill");

            string text = XlsxReader.ExtractText(path);
            RequireContains(text, "无线键盘");
            RequireContains(text, "备注：本季度价格含税。");
            Require(text.IndexOf('\t') >= 0, "xlsx extracted text is not tab separated");

            // 稀疏文本导出只遍历已有单元格：空行不出行，行内每空一列补一个制表符（与旧的整网格扫描一致）。
            // Sparse text export walks only the cells that exist: an empty row emits no line and every skipped
            // column inside a row emits one tab, matching the old full-grid sweep.
            OfficeWorkbook sparseWorkbook = new OfficeWorkbook();
            Worksheet sparse = sparseWorkbook.AddWorksheet("稀疏");
            sparse.GetCell(1, 1).Value = "甲";
            sparse.GetCell(1, 3).Value = "丙";
            sparse.GetCell(3, 2).Value = "乙";
            sparse.GetCell(5, 1).Value = "丁";
            StringBuilder sparseText = new StringBuilder();
            XlsxReader.AppendSheetText(sparseText, sparse);
            string sparseLayout = sparseText.ToString().Replace("\r\n", "\n");
            Require(sparseLayout == "甲\t\t丙\n\t乙\n丁\n",
                "sparse text layout mismatch: " + sparseLayout.Replace('\t', '>').Replace('\n', '|'));

            string roundTripPath = PrepareOutputPath(Path.Combine(Path.GetDirectoryName(path), "excel-roundtrip.xlsx"));
            XlsxWriter.Write(loaded, roundTripPath);
            OfficeWorkbook reloaded = XlsxReader.Load(roundTripPath);
            Require(reloaded.Worksheets.Count == 2, "xlsx round-trip sheet count mismatch");
            Require(HasMerge(reloaded.Worksheets[0], 5, 1, 1, 2), "xlsx round-trip merge missing");
            Require(reloaded.Worksheets[0].GetCell(5, 3).Formula == "SUM(C2:C4)", "xlsx round-trip formula missing");
            Require(reloaded.Worksheets[0].GetCell(2, 4).Value is DateTime, "xlsx round-trip date missing");

            Console.WriteLine("[xlsx] sheets=" + loaded.Worksheets.Count + ", merges=" + loadedSheet.MergedRanges.Count
                + ", freeze=" + loadedSheet.FrozenRows + "/" + loadedSheet.FrozenColumns
                + ", roundtrip=" + new FileInfo(roundTripPath).Length + " bytes");
        }

        private static void RunExcelTemplateTest(string templatePath, string filledPath, string emptyPath)
        {
            OfficeWorkbook template = new OfficeWorkbook();
            template.Metadata.Title = "Excel 模板测试";

            Worksheet sheet = template.AddWorksheet("报表");
            sheet.GetCell(1, 1).Value = "{{reportTitle}}";
            sheet.Merge("A1:B1");
            sheet.GetCell(2, 1).Value = "客户：";
            sheet.GetCell(2, 2).Value = "{{customerName}}";
            sheet.GetCell(3, 1).Value = "{{date:yyyy-mm-dd}}";
            sheet.GetCell(4, 1).Value = "共 {{items.count}} 项，合计 {{total:0.00}} 元";
            sheet.SetColumnWidth(1, 20);
            sheet.SetRowHeight(5, 22);

            CellStyle headerStyle = new CellStyle();
            headerStyle.Bold = true;
            headerStyle.FillColor = new OfficeColor(0xDD, 0xE7, 0xF3);
            string[] headers = { "名称", "数量", "单价", "序号" };
            for (int c = 0; c < headers.Length; c++)
            {
                Cell cell = sheet.GetCell(5, c + 1);
                cell.Value = headers[c];
                cell.Style = headerStyle.Clone();
            }

            sheet.GetCell(6, 1).Value = "{{#each items}}{{name}}";
            sheet.GetCell(6, 2).Value = "{{qty}}";
            sheet.GetCell(6, 3).Value = "{{price:0.00}}";
            sheet.GetCell(6, 4).Value = "{{@number}}";
            sheet.GetCell(6, 5).Value = "{{/each}}";
            sheet.GetCell(7, 1).Value = "备注：{{note}}";

            Worksheet notes = template.AddWorksheet("备注");
            notes.GetCell(1, 1).Value = "备注明细";
            notes.GetCell(2, 1).Value = "{{#each notes}}";
            notes.GetCell(3, 1).Value = "- {{text}}";
            notes.GetCell(4, 1).Value = "{{/each}}";
            notes.GetCell(5, 1).Value = "共 {{notes.count}} 条";
            notes.Merge("A3:B3");
            notes.SetColumnWidth(1, 30);

            XlsxWriter.Write(template, templatePath);

            TemplateContext context = new TemplateContext();
            context.Set("reportTitle", "2026 年第三季度销售报表");
            context.Set("customerName", "华东供应链有限公司");
            context.Set("date", new DateTime(2026, 10, 20));
            context.Set("total", 5000.5m);
            context.Set("note", "价格含税");
            List<Dictionary<string, object>> items = new List<Dictionary<string, object>>();
            items.Add(new Dictionary<string, object> { { "name", "物料 A" }, { "qty", 2 }, { "price", 1250.5m } });
            items.Add(new Dictionary<string, object> { { "name", "物料 B" }, { "qty", 1 }, { "price", 880m } });
            items.Add(new Dictionary<string, object> { { "name", "物料 C" }, { "qty", 4 }, { "price", 37.75m } });
            context.Set("items", items);
            List<Dictionary<string, object>> noteItems = new List<Dictionary<string, object>>();
            noteItems.Add(new Dictionary<string, object> { { "text", "第一条备注" } });
            noteItems.Add(new Dictionary<string, object> { { "text", "第二条备注" } });
            context.Set("notes", noteItems);

            XlsxTemplate.Fill(templatePath, filledPath, context);
            OfficeWorkbook filled = XlsxReader.Load(filledPath);
            Worksheet report = filled.Worksheets[0];
            Require(report.GetCell(1, 1).Value as string == "2026 年第三季度销售报表", "xlsx template title mismatch");
            Require(HasMerge(report, 1, 1, 1, 2), "xlsx template merge A1:B1 not preserved");
            Require(report.GetCell(2, 2).Value as string == "华东供应链有限公司", "xlsx template customer mismatch");
            Require(report.GetCell(3, 1).Value is DateTime, "xlsx template date type mismatch");
            Require(((DateTime)report.GetCell(3, 1).Value).Date == new DateTime(2026, 10, 20), "xlsx template date mismatch");
            Require(report.GetCell(3, 1).Style != null && report.GetCell(3, 1).Style.IsDateFormat(), "xlsx template date format missing");
            Require(report.GetCell(4, 1).Value as string == "共 3 项，合计 5000.50 元",
                "xlsx template mixed token mismatch: " + report.GetCell(4, 1).Value);
            Require(report.GetCell(5, 1).Value as string == "名称", "xlsx template header missing");
            Require(Math.Abs(report.GetRowHeight(5) - 22) < 0.01, "xlsx template row height lost");
            Require(report.GetCell(5, 1).Style != null && report.GetCell(5, 1).Style.Bold, "xlsx template header style lost");

            Require(report.GetCell(6, 1).Value as string == "物料 A", "xlsx template loop row 1 mismatch");
            Require(Convert.ToDouble(report.GetCell(6, 2).Value) == 2, "xlsx template numeric token lost type");
            Require(Math.Abs(Convert.ToDouble(report.GetCell(6, 3).Value) - 1250.5) < 0.0001, "xlsx template decimal mismatch");
            Require(report.GetCell(6, 3).Style.NumberFormat == "0.00", "xlsx template price format mismatch: "
                + report.GetCell(6, 3).Style.NumberFormat);
            Require(Convert.ToDouble(report.GetCell(6, 4).Value) == 1, "xlsx template @number mismatch");
            Require(report.GetCell(7, 1).Value as string == "物料 B", "xlsx template loop row 2 mismatch");
            Require(report.GetCell(8, 1).Value as string == "物料 C", "xlsx template loop row 3 mismatch");
            Require(Convert.ToDouble(report.GetCell(8, 4).Value) == 3, "xlsx template @number last mismatch");
            Require(report.GetCell(9, 1).Value as string == "备注：价格含税", "xlsx template trailing row mismatch: "
                + report.GetCell(9, 1).Value);

            Worksheet notesSheet = filled.Worksheets[1];
            Require(notesSheet.GetCell(2, 1).Value as string == "- 第一条备注", "xlsx template multi-row loop 1 mismatch");
            Require(notesSheet.GetCell(3, 1).Value as string == "- 第二条备注", "xlsx template multi-row loop 2 mismatch");
            Require(notesSheet.GetCell(4, 1).Value as string == "共 2 条", "xlsx template count mismatch: "
                + notesSheet.GetCell(4, 1).Value);
            Require(HasMerge(notesSheet, 2, 1, 1, 2), "xlsx template loop merge 1 not replicated");
            Require(HasMerge(notesSheet, 3, 1, 1, 2), "xlsx template loop merge 2 not replicated");

            string filledText = XlsxReader.ExtractText(filledPath);
            Require(filledText.IndexOf("{{", StringComparison.Ordinal) < 0, "xlsx template tokens remain: " + filledText);

            TemplateContext empty = new TemplateContext();
            empty.Set("reportTitle", "空数据报表");
            empty.Set("customerName", "无");
            empty.Set("date", new DateTime(2026, 10, 20));
            empty.Set("total", 0m);
            empty.Set("note", string.Empty);
            empty.Set("items", new List<object>());
            empty.Set("notes", new List<object>());
            XlsxTemplate.Fill(templatePath, emptyPath, empty);
            OfficeWorkbook emptyBook = XlsxReader.Load(emptyPath);
            Worksheet emptyReport = emptyBook.Worksheets[0];
            Require(emptyReport.GetCell(4, 1).Value as string == "共 0 项，合计 0.00 元", "xlsx empty total mismatch: "
                + emptyReport.GetCell(4, 1).Value);
            Require(emptyReport.GetCell(6, 1).Value as string == "备注：", "xlsx empty row loop not removed: "
                + emptyReport.GetCell(6, 1).Value);
            Worksheet emptyNotes = emptyBook.Worksheets[1];
            Require(emptyNotes.GetCell(2, 1).Value as string == "共 0 条", "xlsx empty multi-row loop not removed: "
                + emptyNotes.GetCell(2, 1).Value);
            Require(emptyNotes.MergedRanges.Count == 0, "xlsx empty loop left merged ranges");
            string emptyText = XlsxReader.ExtractText(emptyPath);
            Require(emptyText.IndexOf("{{", StringComparison.Ordinal) < 0, "xlsx empty template tokens remain");
            Require(emptyText.IndexOf("物料", StringComparison.Ordinal) < 0, "xlsx empty template kept loop data");

            Console.WriteLine("[xlsx-template] filled=" + new FileInfo(filledPath).Length + " bytes, rows="
                + report.LastRow + ", notes rows=" + notesSheet.LastRow + ", empty rows=" + emptyReport.LastRow);
        }

        private static void RunLegacyXlsTest(string path)
        {
            OfficeWorkbook workbook = XlsReader.Load(path);
            Require(workbook.Worksheets.Count == 1, "legacy xls worksheet count mismatch");
            Worksheet sheet = workbook.Worksheets[0];
            Require(sheet.LastRow == 1, "legacy xls row count mismatch: " + sheet.LastRow);
            Require(sheet.LastColumn == 12, "legacy xls column count mismatch: " + sheet.LastColumn);
            Require(Convert.ToString(sheet.GetCell(1, 1).Value).Trim() == "单位", "legacy xls A1 mismatch");
            Require(Convert.ToString(sheet.GetCell(1, 7).Value).Trim() == "电子邮件", "legacy xls G1 mismatch: "
                + sheet.GetCell(1, 7).Value);
            Require(Convert.ToString(sheet.GetCell(1, 8).Value).Trim() == "邮编", "legacy xls H1 mismatch: "
                + sheet.GetCell(1, 8).Value);
            Require(Convert.ToString(sheet.GetCell(1, 12).Value).Trim() == "地址（英文）", "legacy xls L1 mismatch: "
                + sheet.GetCell(1, 12).Value);

            string text = XlsReader.ExtractText(path);
            Require(text.Length > 0, "legacy xls extracted no text");
            RequireContains(text, "单位");
            RequireContains(text, "电子邮件");
            RequireContains(text, "地址（英文）");
            Require(workbook.Metadata.Title == "名片地址", "legacy xls title mismatch: " + workbook.Metadata.Title);
            RequireContains(workbook.Metadata.Author, "微软");
            Console.WriteLine("[xls] sheets=1, columns=" + sheet.LastColumn + ", title='" + workbook.Metadata.Title
                + "', author='" + workbook.Metadata.Author + "'");
        }

        private static void RunLegacyXlsSolverTest(string path, string convertedPath)
        {
            OfficeWorkbook workbook = XlsReader.Load(path);
            Require(workbook.Worksheets.Count == 6, "solver xls worksheet count mismatch: " + workbook.Worksheets.Count);
            Require(workbook.Metadata.Title == "Solver Samples", "solver xls title mismatch: " + workbook.Metadata.Title);
            Require(workbook.Worksheets[0].Name == "Product Mix", "solver sheet 1 name mismatch: " + workbook.Worksheets[0].Name);
            Require(workbook.Worksheets[5].Name == "Engineering Design", "solver sheet 6 name mismatch");

            Worksheet productMix = workbook.Worksheets[0];
            Require(productMix.LastRow == 18 && productMix.LastColumn == 8,
                "solver sheet 1 dimensions mismatch: " + productMix.LastRow + "x" + productMix.LastColumn);
            RequireContains(productMix.GetCell(1, 1).DisplayText(), "Product mix problem");
            Require(Convert.ToString(productMix.GetCell(11, 1).Value).Trim() == "Chassis", "solver cell A11 mismatch: "
                + productMix.GetCell(11, 1).Value);
            Require(Math.Abs(Convert.ToDouble(productMix.GetCell(11, 2).Value) - 450) < 0.0001, "solver inventory mismatch");
            Require(Math.Abs(Convert.ToDouble(productMix.GetCell(11, 3).Value) - 200) < 0.0001, "solver usage mismatch");
            Require(Math.Abs(Convert.ToDouble(productMix.GetCell(9, 4).Value) - 100) < 0.0001, "solver build quantity mismatch");
            Require(Math.Abs(productMix.GetRowHeight(1) - 13.5) < 0.01, "solver row height mismatch: " + productMix.GetRowHeight(1));
            Require(Math.Abs(productMix.DefaultRowHeight - 10.5) < 0.01, "solver default row height mismatch");

            Worksheet income = workbook.Worksheets[3];
            Require(income.Name == "Maximizing Income", "solver sheet 4 name mismatch");
            Require(Math.Abs(Convert.ToDouble(income.GetCell(11, 2).Value) - 400000) < 0.0001, "solver cash mismatch");
            Require(Math.Abs(Convert.ToDouble(income.GetCell(11, 8).Value) - 125400) < 0.0001, "solver ending cash mismatch");

            Worksheet portfolio = workbook.Worksheets[4];
            Require(Math.Abs(Convert.ToDouble(portfolio.GetCell(10, 2).Value) - 0.8) < 0.0001, "solver beta mismatch");

            Require(workbook.Worksheets[2].ShowGridLines, "solver staff sheet gridlines missing");
            Require(!productMix.ShowGridLines, "solver product mix gridlines state mismatch");

            string text = XlsReader.ExtractText(path);
            RequireContains(text, "[Product Mix]");
            RequireContains(text, "[Shipping Routes]");
            RequireContains(text, "TV set");

            XlsxWriter.Write(workbook, convertedPath);
            OfficeWorkbook converted = XlsxReader.Load(convertedPath);
            Require(converted.Worksheets.Count == 6, "converted xlsx worksheet count mismatch");
            Require(Math.Abs(Convert.ToDouble(converted.Worksheets[0].GetCell(11, 2).Value) - 450) < 0.0001,
                "converted xlsx inventory mismatch");
            Require(converted.Worksheets[4].GetCell(10, 2).Value != null, "converted xlsx beta missing");

            Console.WriteLine("[xls-solver] sheets=" + workbook.Worksheets.Count + ", rows=" + productMix.LastRow
                + "x" + productMix.LastColumn + ", defaultH=" + productMix.DefaultRowHeight
                + ", converted=" + new FileInfo(convertedPath).Length + " bytes");
        }

        private static void RunFacadeTest(string docxPath, string xlsxPath, string pdfPath, string docPath,
            string xlsPath, string outputDirectory)
        {
            Require(OfficeFileFormats.Detect(docxPath) == OfficeFileFormat.Docx, "docx detection failed");
            Require(OfficeFileFormats.Detect(xlsxPath) == OfficeFileFormat.Xlsx, "xlsx detection failed");
            Require(OfficeFileFormats.Detect(pdfPath) == OfficeFileFormat.Pdf, "pdf detection failed");
            Require(OfficeFileFormats.Detect(docPath) == OfficeFileFormat.Doc, "doc detection failed");
            Require(OfficeFileFormats.Detect(xlsPath) == OfficeFileFormat.Xls, "xls detection failed");
            Require(OfficeFileFormats.Detect(new byte[] { 1, 2, 3, 4 }) == OfficeFileFormat.Unknown,
                "unknown data detection failed");

            OfficeFile wordFile = OfficeFile.Load(docxPath);
            Require(wordFile.Format == OfficeFileFormat.Docx, "docx facade format mismatch");
            Require(wordFile.Document != null && wordFile.Workbook == null, "docx facade content mismatch");
            RequireContains(wordFile.ExtractText(), "以上为全部数据。");

            string wordPdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "facade-word.pdf"));
            OfficeConverter.Convert(docxPath, wordPdfPath);
            PdfDocument wordPdf = PdfDocument.Load(wordPdfPath);
            Require(wordPdf.PageCount >= 2, "word to pdf page count mismatch: " + wordPdf.PageCount);
            RequireContains(wordPdf.ExtractText(), "以上为全部数据。");

            OfficeFile pdfFile = OfficeFile.Load(wordPdfPath);
            Require(pdfFile.Format == OfficeFileFormat.Pdf && pdfFile.Pdf != null, "pdf facade format mismatch");
            string pdfDocxPath = PrepareOutputPath(Path.Combine(outputDirectory, "facade-pdf.docx"));
            OfficeConverter.Convert(wordPdfPath, pdfDocxPath);
            ValidateOoxmlPackage(pdfDocxPath, "word/document.xml");
            OfficeFile pdfAsWord = OfficeFile.Load(pdfDocxPath);
            Require(pdfAsWord.Format == OfficeFileFormat.Docx, "pdf to docx format mismatch");
            RequireContains(pdfAsWord.ExtractText(), "以上为全部数据。");

            OfficeFile excelFile = OfficeFile.Load(xlsxPath);
            Require(excelFile.Workbook != null && excelFile.Workbook.Worksheets.Count == 2, "xlsx facade workbook mismatch");
            string excelPdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "facade-excel.pdf"));
            OfficeConverter.Convert(xlsxPath, excelPdfPath);
            PdfDocument excelPdf = PdfDocument.Load(excelPdfPath);
            Require(excelPdf.PageCount >= 2, "excel to pdf page count mismatch: " + excelPdf.PageCount);
            string excelPdfText = excelPdf.ExtractText();
            RequireContains(excelPdfText, "销售数据");
            RequireContains(excelPdfText, "无线键盘");
            RequireContains(excelPdfText, "2026-10-01");
            RequireContains(excelPdfText, "2,288.40");

            OfficeDocument layout = WorkbookToDocumentConverter.Convert(excelFile.Workbook);
            Require(layout.Sections.Count == 2, "workbook document section count mismatch");
            Table layoutTable = FindFirstTable(layout.Sections[0].Blocks);
            Require(layoutTable != null && layoutTable.Columns.Count == 4, "workbook document table mismatch");
            Require(layoutTable.Rows.Count == 7, "workbook document row count mismatch: " + layoutTable.Rows.Count);
            bool mergeSpan = false;
            for (int i = 0; i < layoutTable.Rows[4].Cells.Count; i++)
            {
                if (layoutTable.Rows[4].Cells[i].ColSpan == 2) mergeSpan = true;
            }
            Require(mergeSpan, "workbook document merge span missing");

            string excelDocxPath = PrepareOutputPath(Path.Combine(outputDirectory, "facade-excel.docx"));
            OfficeConverter.Convert(xlsxPath, excelDocxPath);
            ValidateOoxmlPackage(excelDocxPath, "word/document.xml");
            OfficeDocument excelDoc = DocxReader.Load(excelDocxPath);
            RequireContains(excelDoc.GetPlainText(), "无线键盘");
            RequireContains(excelDoc.GetPlainText(), "2026-10-01");

            string docxFromDocPath = PrepareOutputPath(Path.Combine(outputDirectory, "facade-legacy.docx"));
            OfficeConverter.Convert(docPath, docxFromDocPath);
            ValidateOoxmlPackage(docxFromDocPath, "word/document.xml");
            OfficeFile convertedDoc = OfficeFile.Load(docxFromDocPath);
            Require(convertedDoc.Format == OfficeFileFormat.Docx, "doc to docx format mismatch");
            Require(convertedDoc.Metadata.Title == "名片地址", "doc to docx metadata title mismatch");
            RequireContains(convertedDoc.ExtractText(), "电子邮件");

            string xlsxFromXlsPath = PrepareOutputPath(Path.Combine(outputDirectory, "facade-legacy.xlsx"));
            OfficeConverter.Convert(xlsPath, xlsxFromXlsPath);
            ValidateOoxmlPackage(xlsxFromXlsPath, "xl/workbook.xml");
            OfficeFile convertedXls = OfficeFile.Load(xlsxFromXlsPath);
            Require(convertedXls.Format == OfficeFileFormat.Xlsx && convertedXls.Workbook.Worksheets.Count == 1,
                "xls to xlsx conversion mismatch");
            Require(Convert.ToString(convertedXls.Workbook.Worksheets[0].GetCell(1, 1).Value).Trim() == "单位",
                "xls to xlsx content mismatch");

            string reSavedPath = PrepareOutputPath(Path.Combine(outputDirectory, "facade-resaved.xlsx"));
            excelFile.Save(reSavedPath, OfficeFileFormat.Xlsx);
            ValidateOoxmlPackage(reSavedPath, "xl/workbook.xml");
            Require(new FileInfo(reSavedPath).Length > 2000, "facade save produced a small file");
            Require(OfficeFile.Load(reSavedPath).Workbook.Worksheets.Count == 2, "facade save round-trip mismatch");

            Console.WriteLine("[facade] word-pdf pages=" + wordPdf.PageCount + ", excel-pdf pages=" + excelPdf.PageCount
                + ", pdf-docx='" + pdfAsWord.Metadata.Title + "', legacy xls->xlsx sheets="
                + convertedXls.Workbook.Worksheets.Count);
        }

        private static void RunReportTest(string outputDirectory)
        {
            Report report = new Report("季度销售报表");
            report.Subtitle = "2026 年第三季度";
            report.Author = "Wisdom.Office";
            report.FooterText = "Wisdom.Office 自动生成";
            report.GeneratedAt = new DateTime(2026, 10, 20, 9, 30, 0);

            ReportTable table = report.AddTable("产品销售明细");
            table.AddColumn("产品", 120, null, HorizontalAlignment.Left);
            table.AddColumn("数量", 60, "0", HorizontalAlignment.Right);
            table.AddColumn("金额", 100, "#,##0.00", HorizontalAlignment.Right);
            table.AddColumn("日期", 90, "yyyy-MM-dd", HorizontalAlignment.Center);
            table.AddRow("无线键盘", 12, 3594.0, new DateTime(2026, 7, 5));
            table.AddRow("机械鼠标", 30, 2697.0, new DateTime(2026, 7, 6));
            table.AddRow("27 英寸显示器", 5, 9495.0, new DateTime(2026, 7, 7));
            report.AddNote("金额单位：人民币元。");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "report-facade.pdf"));
            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "report-facade.docx"));
            string xlsxPath = PrepareOutputPath(Path.Combine(outputDirectory, "report-facade.xlsx"));
            report.Export(pdfPath);
            report.Export(docxPath);
            report.Export(xlsxPath);
            ValidateOoxmlPackage(docxPath, "word/document.xml");
            ValidateOoxmlPackage(xlsxPath, "xl/workbook.xml");

            PdfDocument pdf = PdfDocument.Load(pdfPath);
            Require(pdf.PageCount >= 1, "report pdf has no pages");
            string pdfText = pdf.ExtractText();
            RequireContains(pdfText, "季度销售报表");
            RequireContains(pdfText, "产品销售明细");
            RequireContains(pdfText, "无线键盘");
            RequireContains(pdfText, "3,594.00");
            RequireContains(pdfText, "2026-07-05");
            RequireContains(pdfText, "金额单位：人民币元。");

            OfficeDocument doc = DocxReader.Load(docxPath);
            string docText = doc.GetPlainText();
            RequireContains(docText, "季度销售报表");
            RequireContains(docText, "产品销售明细");
            RequireContains(docText, "27 英寸显示器");
            RequireContains(docText, "金额单位：人民币元。");
            Require(FindFirstTable(doc.Sections[0].Blocks) != null, "report docx table missing");

            OfficeWorkbook workbook = XlsxReader.Load(xlsxPath);
            Worksheet sheet = workbook.Worksheets[0];
            Require(sheet.GetCell(1, 1).Value as string == "季度销售报表", "report xlsx title mismatch");
            Require(HasMerge(sheet, 1, 1, 1, 4), "report xlsx title merge missing");
            Require(sheet.GetCell(5, 1).Value as string == "产品销售明细", "report xlsx caption mismatch");
            Require(sheet.GetCell(6, 1).Value as string == "产品", "report xlsx header mismatch");
            Require(Convert.ToDouble(sheet.GetCell(7, 2).Value) == 12, "report xlsx quantity mismatch");
            Require(Math.Abs(Convert.ToDouble(sheet.GetCell(7, 3).Value) - 3594) < 0.0001, "report xlsx amount mismatch");
            Require(sheet.GetCell(7, 4).Value is DateTime, "report xlsx date type mismatch");
            Require(sheet.GetCell(7, 3).Style.NumberFormat == "#,##0.00", "report xlsx number format mismatch");
            Require(sheet.GetCell(6, 1).Style.FillColor.HasValue, "report xlsx header fill missing");
            Require(sheet.FrozenRows == 6, "report xlsx frozen header mismatch: " + sheet.FrozenRows);

            Console.WriteLine("[report] pdf pages=" + pdf.PageCount + ", docx=" + new FileInfo(docxPath).Length
                + " bytes, xlsx rows=" + sheet.LastRow + ", title='" + workbook.Metadata.Title + "'");
        }

        private static bool HasMerge(Worksheet sheet, int row, int column, int rowSpan, int columnSpan)
        {
            for (int i = 0; i < sheet.MergedRanges.Count; i++)
            {
                MergedRange range = sheet.MergedRanges[i];
                if (range.Row == row && range.Column == column && range.RowSpan == rowSpan && range.ColumnSpan == columnSpan)
                {
                    return true;
                }
            }
            return false;
        }

        private static void RequireColor(OfficeColor actual, OfficeColor expected, string message)
        {
            Require(actual == expected, "xlsx color mismatch (" + message + "): " + actual + " != " + expected);
        }

        private static Paragraph FindParagraph(List<Block> blocks, string text)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null && paragraph.GetText().IndexOf(text, StringComparison.Ordinal) >= 0) return paragraph;
            }
            return null;
        }

        private static Table FindFirstTable(List<Block> blocks)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Table table = blocks[i] as Table;
                if (table != null) return table;
            }
            return null;
        }

        private static List<Table> FindTables(List<Block> blocks)
        {
            List<Table> tables = new List<Table>();
            for (int i = 0; i < blocks.Count; i++)
            {
                Table table = blocks[i] as Table;
                if (table != null) tables.Add(table);
            }
            return tables;
        }

        private static bool ContainsCellText(Table table, string text)
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                {
                    if (GetBlocksText(table.Rows[r].Cells[c].Blocks).IndexOf(text, StringComparison.Ordinal) >= 0) return true;
                }
            }
            return false;
        }

        private static int CountOccurrences(string text, string value)
        {
            int count = 0;
            int index = 0;
            while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
            {
                count++;
                index += value.Length;
            }
            return count;
        }

        private static string ReadPackageEntryText(string path, string entryName)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = archive.GetEntry(entryName);
                Require(entry != null, "OOXML part missing: " + entryName);
                using (StreamReader reader = new StreamReader(entry.Open(), Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        private static string GetBlocksText(List<Block> blocks)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    builder.Append(paragraph.GetText());
                    builder.Append(' ');
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        builder.Append(GetBlocksText(table.Rows[r].Cells[c].Blocks));
                        builder.Append(' ');
                    }
                }
            }
            return builder.ToString();
        }

        private static void ValidateOoxmlPackage(string path, string mainPart)
        {
            Require(File.Exists(path), "OOXML package missing: " + path);
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                Dictionary<string, ZipArchiveEntry> entries =
                    new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < archive.Entries.Count; i++)
                {
                    ZipArchiveEntry entry = archive.Entries[i];
                    entries[NormalizePackagePath(entry.FullName)] = entry;
                }

                Require(entries.ContainsKey("[Content_Types].xml"), "OOXML content types missing: " + path);
                Require(entries.ContainsKey("_rels/.rels"), "OOXML package relationships missing: " + path);
                Require(entries.ContainsKey(mainPart), "OOXML main part missing: " + mainPart);

                foreach (KeyValuePair<string, ZipArchiveEntry> pair in entries)
                {
                    if (pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        || pair.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                    {
                        LoadPackageXml(pair.Value, path);
                    }
                }

                XDocument contentTypes = LoadPackageXml(entries["[Content_Types].xml"], path);
                XNamespace contentTypeNs = "http://schemas.openxmlformats.org/package/2006/content-types";
                Require(contentTypes.Root != null && contentTypes.Root.Name == contentTypeNs + "Types",
                    "OOXML content types root mismatch: " + path);
                bool hasRelationshipsDefault = false;
                bool hasXmlDefault = false;
                foreach (XElement element in contentTypes.Root.Elements())
                {
                    if (element.Name == contentTypeNs + "Default")
                    {
                        string extension = (string)element.Attribute("Extension");
                        if (string.Equals(extension, "rels", StringComparison.OrdinalIgnoreCase)) hasRelationshipsDefault = true;
                        if (string.Equals(extension, "xml", StringComparison.OrdinalIgnoreCase)) hasXmlDefault = true;
                    }
                    else if (element.Name == contentTypeNs + "Override")
                    {
                        string partName = (string)element.Attribute("PartName");
                        Require(!string.IsNullOrEmpty(partName), "OOXML content type override has no part name: " + path);
                        Require(entries.ContainsKey(NormalizePackagePath(partName)),
                            "OOXML content type points to missing part: " + partName);
                    }
                }
                Require(hasRelationshipsDefault, "OOXML relationships content type missing: " + path);
                Require(hasXmlDefault, "OOXML xml content type missing: " + path);

                XNamespace relationshipNs = "http://schemas.openxmlformats.org/package/2006/relationships";
                foreach (KeyValuePair<string, ZipArchiveEntry> pair in entries)
                {
                    if (!pair.Key.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                    XDocument relationships = LoadPackageXml(pair.Value, path);
                    Require(relationships.Root != null && relationships.Root.Name == relationshipNs + "Relationships",
                        "OOXML relationships root mismatch: " + pair.Key);
                    string sourcePart = GetRelationshipSourcePart(pair.Key);
                    HashSet<string> ids = new HashSet<string>(StringComparer.Ordinal);
                    foreach (XElement relationship in relationships.Root.Elements(relationshipNs + "Relationship"))
                    {
                        string id = (string)relationship.Attribute("Id");
                        Require(!string.IsNullOrEmpty(id), "OOXML relationship without id: " + pair.Key);
                        Require(ids.Add(id), "OOXML duplicate relationship id: " + id + " in " + pair.Key);

                        string targetMode = (string)relationship.Attribute("TargetMode");
                        if (string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase)) continue;
                        string target = (string)relationship.Attribute("Target");
                        Require(!string.IsNullOrEmpty(target), "OOXML relationship without target: " + id);
                        string resolved = ResolvePackageTarget(sourcePart, target);
                        Require(entries.ContainsKey(resolved),
                            "OOXML relationship " + id + " points to missing part: " + resolved);
                    }
                }
            }
        }

        private static XDocument LoadPackageXml(ZipArchiveEntry entry, string packagePath)
        {
            try
            {
                using (Stream entryStream = entry.Open())
                {
                    return XDocument.Load(entryStream);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Test failed: invalid XML part '" + entry.FullName
                    + "' in " + packagePath + ": " + ex.Message, ex);
            }
        }

        private static string GetRelationshipSourcePart(string relationshipsPath)
        {
            string normalized = NormalizePackagePath(relationshipsPath);
            int marker = normalized.LastIndexOf("_rels/", StringComparison.OrdinalIgnoreCase);
            string prefix = marker < 0 ? string.Empty : normalized.Substring(0, marker);
            string fileName = marker < 0 ? normalized : normalized.Substring(marker + 6);
            if (fileName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                fileName = fileName.Substring(0, fileName.Length - 5);
            }
            return prefix + fileName;
        }

        private static string ResolvePackageTarget(string sourcePart, string target)
        {
            if (target == null) return string.Empty;
            target = target.Replace('\\', '/');
            int fragment = target.IndexOf('#');
            if (fragment >= 0) target = target.Substring(0, fragment);
            target = Uri.UnescapeDataString(target);

            string baseDirectory = string.Empty;
            int slash = sourcePart.LastIndexOf('/');
            if (slash >= 0) baseDirectory = sourcePart.Substring(0, slash + 1);
            string combined = target.StartsWith("/", StringComparison.Ordinal)
                ? target.Substring(1)
                : baseDirectory + target;

            List<string> segments = new List<string>();
            string[] parts = combined.Split('/');
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i].Length == 0 || parts[i] == ".") continue;
                if (parts[i] == "..")
                {
                    if (segments.Count > 0) segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                segments.Add(parts[i]);
            }
            return string.Join("/", segments.ToArray());
        }

        private static string NormalizePackagePath(string path)
        {
            string normalized = (path ?? string.Empty).Replace('\\', '/');
            while (normalized.StartsWith("/", StringComparison.Ordinal))
            {
                normalized = normalized.Substring(1);
            }
            return normalized;
        }

        private static void RunWordNormalizeTest(string outputDirectory)
{
    OfficeDocument document = new OfficeDocument();
    Section section = document.AddSection();

    Paragraph title = new Paragraph();
    section.Blocks.Add(title);
    title.Alignment = ParagraphAlignment.Center;
    title.AddText("某某公司年度报告", new TextFormat("微软雅黑", 18) { Bold = true });

    section.Blocks.Add(new Paragraph("   "));

    Paragraph shortFirst = new Paragraph();
    section.Blocks.Add(shortFirst);
    shortFirst.AddText("本文介绍了", new TextFormat("微软雅黑", 11));

    Paragraph shortSecond = new Paragraph();
    section.Blocks.Add(shortSecond);
    shortSecond.AddText("公司的发展历史", new TextFormat("微软雅黑", 11));

    Paragraph halfWidth = new Paragraph();
    section.Blocks.Add(halfWidth);
    halfWidth.AddText("这一段使用半角句号结尾.", new TextFormat("微软雅黑", 11));

    Paragraph spaces = new Paragraph();
    section.Blocks.Add(spaces);
    spaces.AddText("这里有   多个    空格", new TextFormat("微软雅黑", 11));
    spaces.AddLineBreak();
    spaces.AddText("以及多余换行。", new TextFormat("微软雅黑", 11));

    StringBuilder longText = new StringBuilder();
    for (int i = 1; i <= 12; i++)
    {
        longText.Append("第").Append(i).Append("句用于验证超长段落拆分逻辑，内容需要足够长。");
    }
    Paragraph longParagraph = new Paragraph();
    section.Blocks.Add(longParagraph);
    longParagraph.AddText(longText.ToString(), new TextFormat("微软雅黑", 11));

    Paragraph fakeHeading2 = new Paragraph();
    section.Blocks.Add(fakeHeading2);
    fakeHeading2.AddText("二、业务概览", new TextFormat("微软雅黑", 15) { Bold = true });

    Paragraph skipHeading = new Paragraph();
    section.Blocks.Add(skipHeading);
    skipHeading.HeadingLevel = 4;
    skipHeading.AddText("细节小节", new TextFormat("微软雅黑", 10));

    Paragraph tinyImage = new Paragraph();
    section.Blocks.Add(tinyImage);
    tinyImage.AddImage(new OfficeImage(CreateTestPng(1, 1)), 10, 10);

    Paragraph bigImage = new Paragraph();
    section.Blocks.Add(bigImage);
    bigImage.AddImage(new OfficeImage(CreateTestPng(1200, 400)), 0, 0);

    Table table = new Table();
    table.Columns.Add(new TableColumn(80));
    table.Columns.Add(new TableColumn(80));
    table.Columns.Add(new TableColumn(80));
    table.Columns.Add(new TableColumn(80));
    table.AddRow("项目", "数值", "备注", string.Empty);
    table.AddRow("A", "1", string.Empty, string.Empty);
    table.AddRow(string.Empty, string.Empty, string.Empty, string.Empty);
    section.Blocks.Add(table);

    WordNormalizeOptions options = new WordNormalizeOptions();
    options.AddMissingEndingPunctuation = true;
    WordNormalizeResult result = WordNormalizer.Normalize(document, options);

    Require(result.Changes.Count > 0, "word normalize recorded no changes");
    Require(result.Before.EmptyParagraphCount >= 1, "word normalize before metrics missed empty paragraph");
    Require(result.After.EmptyParagraphCount < result.Before.EmptyParagraphCount, "word normalize left empty paragraphs");

    Paragraph tocTitle = document.FirstSection.Blocks[0] as Paragraph;
    Require(tocTitle != null && tocTitle.GetText() == "目录", "word normalize toc title missing");
    Require(FindParagraph(document, delegate (Paragraph p) { return HasFieldInstruction(p, "PAGEREF"); }) != null, "word normalize toc entries missing");
    Require(ContainsPageBreak(document.FirstSection.Blocks), "word normalize toc page break missing");

    Paragraph promoted = FindParagraph(document, delegate (Paragraph p) { return p.GetText().Trim() == "某某公司年度报告"; });
    Require(promoted != null && promoted.HeadingLevel == 1, "word normalize title not promoted to level 1");
    Paragraph level2 = FindParagraph(document, delegate (Paragraph p) { return p.GetText().Trim() == "二、业务概览"; });
    Require(level2 != null && level2.HeadingLevel == 2, "word normalize fake heading level mismatch: " + (level2 == null ? "null" : level2.HeadingLevel.ToString()));
    Paragraph fixedLevel = FindParagraph(document, delegate (Paragraph p) { return p.GetText().Trim() == "细节小节"; });
    Require(fixedLevel != null && fixedLevel.HeadingLevel == 3, "word normalize heading skip not fixed");

    Paragraph merged = FindParagraph(document, delegate (Paragraph p) { return p.GetText().IndexOf("本文介绍了公司的发展历史", StringComparison.Ordinal) >= 0; });
    Require(merged != null, "word normalize short paragraphs not merged");
    Require(merged.GetText().EndsWith("。", StringComparison.Ordinal), "word normalize merged paragraph punctuation missing: " + merged.GetText());

    Paragraph halfWidthParagraph = FindParagraph(document, delegate (Paragraph p) { return p.GetText().IndexOf("半角句号结尾", StringComparison.Ordinal) >= 0; });
    Require(halfWidthParagraph != null && halfWidthParagraph.GetText().EndsWith("。", StringComparison.Ordinal), "word normalize half-width ending not converted");

    Paragraph spaceParagraph = FindParagraph(document, delegate (Paragraph p) { return p.GetText().IndexOf("多个", StringComparison.Ordinal) >= 0 && p.GetText().IndexOf("空格", StringComparison.Ordinal) >= 0; });
    Require(spaceParagraph != null, "word normalize space paragraph missing");
    Require(spaceParagraph.GetText().IndexOf("  ", StringComparison.Ordinal) < 0, "word normalize extra spaces not collapsed");
    Require(spaceParagraph.GetText().IndexOf('\n') < 0, "word normalize extra line break not removed");

    Paragraph firstChunk = FindParagraph(document, delegate (Paragraph p) { return p.GetText().IndexOf("第1句", StringComparison.Ordinal) >= 0; });
    Paragraph lastChunk = FindParagraph(document, delegate (Paragraph p) { return p.GetText().IndexOf("第12句", StringComparison.Ordinal) >= 0; });
    Require(firstChunk != null && lastChunk != null && !object.ReferenceEquals(firstChunk, lastChunk), "word normalize long paragraph not split");
    Require(firstChunk.GetText().Length <= options.LongParagraphThreshold + 40, "word normalize split chunk too long");

    Table normalizedTable = null;
    for (int i = 0; i < document.FirstSection.Blocks.Count; i++)
    {
        Table candidate = document.FirstSection.Blocks[i] as Table;
        if (candidate != null && candidate.Rows.Count > 0 && candidate.Rows[0].Cells.Count > 0
            && candidate.Rows[0].Cells[0].Blocks.Count > 0)
        {
            Paragraph cellParagraph = candidate.Rows[0].Cells[0].Blocks[0] as Paragraph;
            if (cellParagraph != null && cellParagraph.GetText() == "项目") normalizedTable = candidate;
        }
    }
    Require(normalizedTable != null, "word normalize table missing");
    Require(normalizedTable.Rows.Count == 2, "word normalize empty table row not removed: " + normalizedTable.Rows.Count);
    Require(normalizedTable.Rows[0].Cells.Count == 3, "word normalize empty table column not removed: " + normalizedTable.Rows[0].Cells.Count);
    Require(normalizedTable.Alignment == TableAlignment.Center, "word normalize table alignment mismatch");
    Require(normalizedTable.Autofit, "word normalize table autofit not enabled");
    Require(normalizedTable.RepeatHeaderRow && normalizedTable.Rows[0].IsHeader, "word normalize table header repeat not enabled");

    Require(!ContainsTinyImage(document), "word normalize 1x1 image not removed");
    Require(result.Changes.FindAll(delegate (WordChange c) { return c.Category == WordNormalizer.CategoryImage; }).Count > 0, "word normalize image changes missing");

    int styledRuns = 0;
    for (int s = 0; s < document.Sections.Count; s++)
    {
        CheckBodyFonts(document.Sections[s].Blocks, options, ref styledRuns);
    }
    Require(styledRuns > 0, "word normalize body font check found no runs");

    string sampleBookmark = null;
    for (int i = 0; i < document.FirstSection.Blocks.Count; i++)
    {
        Paragraph paragraph = document.FirstSection.Blocks[i] as Paragraph;
        if (paragraph == null) continue;
        if (paragraph.GetText().Trim() == "某某公司年度报告" && !string.IsNullOrEmpty(paragraph.BookmarkName))
        {
            sampleBookmark = paragraph.BookmarkName;
        }
    }
    Require(sampleBookmark != null, "word normalize heading bookmark missing");
    Dictionary<string, int> pages = new Dictionary<string, int>();
    pages[sampleBookmark] = 3;
    int updatedFields = WordNormalizer.UpdateTocPageNumbers(document, pages);
    Require(updatedFields == 1, "word normalize toc page number update failed: " + updatedFields);
    Paragraph tocEntry = FindParagraph(document, delegate (Paragraph p) { return HasFieldDisplay(p, sampleBookmark, "3"); });
    Require(tocEntry != null, "word normalize toc display number not updated");

    string normalizedPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-normalized.docx"));
    DocxWriter.Write(document, normalizedPath);
    ValidateOoxmlPackage(normalizedPath, "word/document.xml");

    OfficeDocument reloaded = DocxReader.Load(normalizedPath);
    Paragraph reloadedTocTitle = reloaded.FirstSection.Blocks[0] as Paragraph;
    Require(reloadedTocTitle != null && reloadedTocTitle.GetText() == "目录", "word normalize reloaded toc title mismatch");
    Paragraph reloadedHeading = FindParagraph(reloaded, delegate (Paragraph p) { return p.GetText().Trim() == "某某公司年度报告"; });
    Require(reloadedHeading != null && reloadedHeading.HeadingLevel == 1, "word normalize reloaded heading mismatch");
    Require(reloadedHeading.BookmarkName == sampleBookmark, "word normalize reloaded bookmark mismatch");
    Require(FindParagraph(reloaded, delegate (Paragraph p) { return HasFieldDisplay(p, sampleBookmark, "3"); }) != null, "word normalize reloaded toc field mismatch");

    string processedPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-processed.docx"));
    WordProcessResult processResult = WordProcessor.Process(normalizedPath, processedPath, new WordNormalizeOptions());
    Require(File.Exists(processedPath), "word processor output missing");
    Require(File.Exists(processResult.BackupPath), "word processor backup missing");
    Require(File.Exists(processResult.ReportPath), "word processor report missing");
    RequireContains(processResult.Report, "Word 整理报告");
    RequireContains(processResult.Report, "差异对比");
    Require(processResult.Diff != null && processResult.Diff.UnchangedCount > 0, "word processor diff missing");
    WordProcessor.Rollback(processResult.BackupPath, processedPath);
    Require(new FileInfo(processedPath).Length == new FileInfo(normalizedPath).Length, "word processor rollback size mismatch");

    Console.WriteLine("[word-normalize] headings=" + CountHeadings(document) + ", changes=" + result.Changes.Count
        + ", issues=" + result.Issues.Count + ", table rows=" + normalizedTable.Rows.Count
        + ", toc entries=" + CountTocEntries(document) + ", report=" + new FileInfo(processResult.ReportPath).Length + " bytes");
    if (processResult.Normalize.Changes.Count > 0)
    {
        Console.WriteLine("[word-normalize] second pass change: " + processResult.Normalize.Changes[0].ToString());
    }
}

        private static OfficeDocument BuildComplexReportDocument()
        {
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = "2026 年度经营分析报告";
            document.Metadata.Author = "Wisdom.Office";
            document.Metadata.Subject = "Complex report stress test";
            document.DefaultTextFormat = new TextFormat("宋体", 11);

            Section section = document.AddSection();
            Paragraph header = section.Header.AddParagraph(string.Empty);
            header.Alignment = ParagraphAlignment.Right;
            header.AddText("2026 年度经营分析报告  |  第 {page} 页 / 共 {pages} 页",
                new TextFormat("Microsoft YaHei", 9) { Color = OfficeColor.Gray });
            Paragraph footer = section.Footer.AddParagraph(string.Empty);
            footer.Alignment = ParagraphAlignment.Center;
            footer.AddText("Wisdom.Office 复杂报表压力测试", new TextFormat("Microsoft YaHei", 9) { Color = OfficeColor.Gray });

            Paragraph title = new Paragraph();
            title.HeadingLevel = 1;
            title.Alignment = ParagraphAlignment.Center;
            title.SpaceAfter = 18;
            title.AddText("2026 年度经营分析报告",
                new TextFormat("Microsoft YaHei", 20) { Bold = true, Color = new OfficeColor(0x1F, 0x3A, 0x5F) });
            section.Blocks.Add(title);

            section.Blocks.Add(new Paragraph("   "));
            Paragraph introFirst = new Paragraph();
            introFirst.AddText("本报告基于销售、供应链与财务系统的明细数据汇总。", new TextFormat("宋体", 11));
            section.Blocks.Add(introFirst);
            Paragraph introSecond = new Paragraph();
            introSecond.AddText("报告覆盖 2026 年度四个季度的经营情况。", new TextFormat("宋体", 11));
            section.Blocks.Add(introSecond);

            StringBuilder longParagraphText = new StringBuilder();
            for (int k = 1; k <= 16; k++)
            {
                longParagraphText.Append("第").Append(k)
                    .Append("项经营指标保持稳定，同比变化处于合理区间，需要结合行业景气度持续跟踪。");
            }
            Paragraph longParagraph = new Paragraph();
            longParagraph.FirstLineIndent = 22;
            longParagraph.LineSpacing = 1.3;
            longParagraph.AddText(longParagraphText.ToString(), new TextFormat("宋体", 11));
            section.Blocks.Add(longParagraph);

            string[] chapters =
            {
                "市场概览", "销售分析", "产品结构", "渠道表现", "客户分析", "供应链",
                "成本与费用", "财务摘要", "风险提示", "人力资源", "信息化建设", "下一年度规划"
            };

            for (int i = 0; i < chapters.Length; i++)
            {
                if (i > 0) section.Blocks.Add(new PageBreak());

                Paragraph heading = new Paragraph();
                heading.HeadingLevel = 1;
                heading.SpaceBefore = 12;
                heading.SpaceAfter = 6;
                heading.AddText("第" + (i + 1) + "章 " + chapters[i],
                    new TextFormat("Microsoft YaHei", 16) { Bold = true, Color = new OfficeColor(0x1F, 0x3A, 0x5F) });
                section.Blocks.Add(heading);

                for (int p = 1; p <= 24; p++)
                {
                    Paragraph paragraph = new Paragraph();
                    paragraph.FirstLineIndent = 22;
                    paragraph.LineSpacing = 1.25;
                    paragraph.SpaceAfter = 4;
                    if (p % 7 == 0)
                    {
                        StringBuilder builder = new StringBuilder();
                        for (int k = 1; k <= 16; k++)
                        {
                            builder.Append("第").Append(k).Append("项指标环比变化处于合理区间，");
                        }
                        paragraph.AddText(builder.ToString() + "需要持续跟踪。", new TextFormat("宋体", 11));
                    }
                    else
                    {
                        paragraph.AddText("第" + (i + 1) + "章第" + p + "段：本期数据来自业务系统导出，"
                            + "经过清洗、去重与口径校验后形成；" + chapters[i] + "相关指标详见下表与趋势图。",
                            new TextFormat("宋体", 11));
                    }
                    section.Blocks.Add(paragraph);
                }

                Table table = new Table();
                table.Columns.Add(new TableColumn(90));
                table.Columns.Add(new TableColumn(80));
                table.Columns.Add(new TableColumn(80));
                table.Columns.Add(new TableColumn(60));
                table.Columns.Add(new TableColumn(130));
                if (i % 4 == 0)
                {
                    TableRow captionRow = table.AddRow();
                    TableCell captionCell = new TableCell(chapters[i] + " 关键指标");
                    captionCell.ColSpan = 5;
                    captionCell.BackgroundColor = new OfficeColor(0xDD, 0xE7, 0xF3);
                    captionRow.Cells.Add(captionCell);
                }
                TableRow tableHeader = table.AddRow("指标", "本期", "上期", "同比", "说明");
                tableHeader.IsHeader = true;
                for (int c = 0; c < tableHeader.Cells.Count; c++)
                {
                    TableCell cell = tableHeader.Cells[c];
                    cell.BackgroundColor = new OfficeColor(0x1F, 0x3A, 0x5F);
                    cell.VerticalAlignment = VerticalAlignment.Middle;
                    Paragraph cellParagraph = (Paragraph)cell.Blocks[0];
                    cellParagraph.Alignment = ParagraphAlignment.Center;
                    for (int k = 0; k < cellParagraph.Inlines.Count; k++)
                    {
                        TextRun run = cellParagraph.Inlines[k] as TextRun;
                        if (run != null)
                        {
                            run.Format = new TextFormat("Microsoft YaHei", 9) { Bold = true, Color = OfficeColor.White };
                        }
                    }
                }
                for (int r = 1; r <= 13; r++)
                {
                    TableRow row = table.AddRow("指标 " + (i + 1) + "-" + r,
                        (1200 + i * 37 + r * 11).ToString(),
                        (1100 + i * 31 + r * 9).ToString(),
                        (r % 3 == 0 ? "-" : "+") + (r * 2 + i) + "%",
                        chapters[i] + " 明细说明");
                    if (r % 2 == 0)
                    {
                        for (int c = 0; c < row.Cells.Count; c++)
                        {
                            row.Cells[c].BackgroundColor = new OfficeColor(0xF2, 0xF5, 0xFA);
                        }
                    }
                }
                section.Blocks.Add(table);

                Paragraph figure = new Paragraph();
                figure.Alignment = ParagraphAlignment.Center;
                figure.SpaceBefore = 6;
                figure.AddImage(new OfficeImage(CreateTestPng(320, 120)), 240, 90);
                section.Blocks.Add(figure);
                Paragraph caption = new Paragraph();
                caption.Alignment = ParagraphAlignment.Center;
                caption.AddText("图 " + (i + 1) + " " + chapters[i] + "趋势", new TextFormat("Microsoft YaHei", 9) { Color = OfficeColor.Gray });
                section.Blocks.Add(caption);

                if (i == 4)
                {
                    Paragraph tinyImage = new Paragraph();
                    tinyImage.AddImage(new OfficeImage(CreateTestPng(1, 1)), 8, 8);
                    section.Blocks.Add(tinyImage);
                }
            }

            Paragraph appendix = new Paragraph();
            appendix.SpaceBefore = 12;
            appendix.SpaceAfter = 6;
            appendix.AddText("附录：指标定义", new TextFormat("Microsoft YaHei", 16) { Bold = true });
            section.Blocks.Add(appendix);
            section.Blocks.Add(new Paragraph("指标口径说明：所有金额单位为人民币元，同比为与上年同期比较。"));

            return document;
        }

        private static void RunComplexReportTest(string outputDirectory)
        {
            OfficeDocument document = BuildComplexReportDocument();

            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "complex-report.docx"));
            Stopwatch watch = Stopwatch.StartNew();
            DocxWriter.Write(document, docxPath);
            watch.Stop();
            long writeMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(docxPath, "word/document.xml");

            watch.Restart();
            OfficeDocument loaded = DocxReader.Load(docxPath);
            watch.Stop();
            long readMs = watch.ElapsedMilliseconds;
            Require(CountTables(loaded) >= 12, "complex docx table count mismatch: " + CountTables(loaded));
            Require(CountImages(loaded) >= 12, "complex docx image count mismatch: " + CountImages(loaded));
            Require(CountParagraphs(loaded) > 250, "complex docx paragraph count mismatch: " + CountParagraphs(loaded));

            WordNormalizeOptions options = new WordNormalizeOptions();
            watch.Restart();
            WordNormalizeResult result = WordNormalizer.Normalize(loaded, options);
            watch.Stop();
            long normalizeMs = watch.ElapsedMilliseconds;
            Require(result.Changes.Count > 0, "complex report normalize recorded no changes");
            Require(CountHeadings(loaded) >= 13, "complex report headings not promoted: " + CountHeadings(loaded));
            Require(CountTocEntries(loaded) >= 13, "complex report toc entries missing: " + CountTocEntries(loaded));
            Require(CountImages(loaded) >= 12, "complex report images lost during normalize: " + CountImages(loaded));
            Require(!ContainsTinyImage(loaded), "complex report tiny image not removed");

            string normalizedPath = PrepareOutputPath(Path.Combine(outputDirectory, "complex-report-normalized.docx"));
            watch.Restart();
            DocxWriter.Write(loaded, normalizedPath);
            watch.Stop();
            long normalizedWriteMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(normalizedPath, "word/document.xml");
            Require(new FileInfo(normalizedPath).Length > 20000, "complex normalized docx is unexpectedly small");

            OfficeDocument reloaded = DocxReader.Load(normalizedPath);
            Require(CountHeadings(reloaded) >= 13, "complex reloaded headings mismatch: " + CountHeadings(reloaded));
            Require(CountTables(reloaded) >= 12, "complex reloaded tables mismatch: " + CountTables(reloaded));
            Require(CountImages(reloaded) >= 12, "complex reloaded images mismatch: " + CountImages(reloaded));

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "complex-report.pdf"));
            watch.Restart();
            new PdfRenderer().Render(reloaded, pdfPath);
            watch.Stop();
            long renderMs = watch.ElapsedMilliseconds;
            Require(new FileInfo(pdfPath).Length > 100000, "complex report pdf is unexpectedly small");

            watch.Restart();
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            string text = pdf.ExtractText();
            int pdfTables = 0;
            for (int i = 0; i < pdf.PageCount; i++)
            {
                pdfTables += pdf.Pages[i].GetTables().Count;
            }
            watch.Stop();
            long pdfReadMs = watch.ElapsedMilliseconds;
            Require(pdf.PageCount > 12, "complex report pdf pages mismatch: " + pdf.PageCount);
            RequireContains(text, "第1章 市场概览");
            RequireContains(text, "附录：指标定义");
            RequireContains(text, "第 1 页 / 共 " + pdf.PageCount + " 页");
            Require(pdfTables > 0, "complex report pdf tables not detected");

            Require(writeMs < 10000, "complex docx write too slow: " + writeMs + " ms");
            Require(readMs < 10000, "complex docx read too slow: " + readMs + " ms");
            Require(normalizeMs < 10000, "complex docx normalize too slow: " + normalizeMs + " ms");
            Require(renderMs < 30000, "complex pdf render too slow: " + renderMs + " ms");
            Require(pdfReadMs < 15000, "complex pdf read too slow: " + pdfReadMs + " ms");

            Console.WriteLine("[complex-word] paragraphs=" + CountParagraphs(reloaded) + ", headings=" + CountHeadings(reloaded)
                + ", tables=" + CountTables(reloaded) + ", images=" + CountImages(reloaded)
                + ", docx=" + new FileInfo(normalizedPath).Length + " bytes, write=" + writeMs + " ms, read=" + readMs
                + " ms, normalize=" + normalizeMs + " ms, normalized-write=" + normalizedWriteMs
                + " ms, pdf=" + pdf.PageCount + " pages/" + new FileInfo(pdfPath).Length + " bytes, render=" + renderMs
                + " ms, extract=" + pdfReadMs + " ms, pdf tables=" + pdfTables);
        }

        private static OfficeWorkbook BuildComplexWorkbook(int dataRows)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.Metadata.Title = "复杂销售明细压力测试";
            workbook.Metadata.Author = "Wisdom.Office";

            Worksheet detail = workbook.AddWorksheet("销售明细");
            detail.FrozenRows = 1;
            detail.FrozenColumns = 2;
            detail.ShowGridLines = false;
            detail.FitToPage = true;
            detail.PrintSetup.MarginLeft = 36;
            detail.PrintSetup.MarginRight = 36;
            detail.SetRowHeight(1, 26);
            detail.SetRowHidden(10, true);
            detail.SetColumnHidden(9, true);
            for (int c = 1; c <= 9; c++)
            {
                detail.SetColumnWidth(c, c == 1 ? 16 : c == 3 ? 20 : c == 8 ? 24 : 12);
            }

            CellStyle headerStyle = new CellStyle();
            headerStyle.FontFamily = "Microsoft YaHei";
            headerStyle.FontSize = 10;
            headerStyle.Bold = true;
            headerStyle.FontColor = OfficeColor.White;
            headerStyle.FillColor = new OfficeColor(0x1F, 0x3A, 0x5F);
            headerStyle.HorizontalAlignment = HorizontalAlignment.Center;
            headerStyle.VerticalAlignment = VerticalAlignment.Middle;
            headerStyle.WrapText = true;
            headerStyle.BottomBorder = BorderLine.Thin();

            CellStyle textStyle = new CellStyle();
            textStyle.FontFamily = "Microsoft YaHei";
            textStyle.FontSize = 10;
            CellStyle textAltStyle = textStyle.Clone();
            textAltStyle.FillColor = new OfficeColor(0xF2, 0xF5, 0xFA);
            CellStyle integerStyle = textStyle.Clone();
            integerStyle.NumberFormat = "#,##0";
            integerStyle.HorizontalAlignment = HorizontalAlignment.Right;
            CellStyle integerAltStyle = integerStyle.Clone();
            integerAltStyle.FillColor = new OfficeColor(0xF2, 0xF5, 0xFA);
            CellStyle moneyStyle = textStyle.Clone();
            moneyStyle.NumberFormat = "#,##0.00";
            moneyStyle.HorizontalAlignment = HorizontalAlignment.Right;
            CellStyle moneyAltStyle = moneyStyle.Clone();
            moneyAltStyle.FillColor = new OfficeColor(0xF2, 0xF5, 0xFA);
            CellStyle dateStyle = textStyle.Clone();
            dateStyle.NumberFormat = "yyyy-mm-dd";
            dateStyle.HorizontalAlignment = HorizontalAlignment.Center;
            CellStyle dateAltStyle = dateStyle.Clone();
            dateAltStyle.FillColor = new OfficeColor(0xF2, 0xF5, 0xFA);
            CellStyle boolStyle = textStyle.Clone();
            boolStyle.HorizontalAlignment = HorizontalAlignment.Center;
            CellStyle boolAltStyle = boolStyle.Clone();
            boolAltStyle.FillColor = new OfficeColor(0xF2, 0xF5, 0xFA);
            CellStyle noteStyle = textStyle.Clone();
            noteStyle.Italic = true;
            noteStyle.FontColor = OfficeColor.Gray;

            string[] headers = { "订单号", "日期", "客户", "数量", "单价", "金额", "加急", "备注", "毛利" };
            for (int c = 0; c < headers.Length; c++)
            {
                Cell cell = detail.GetCell(1, c + 1);
                cell.Value = headers[c];
                cell.Style = headerStyle;
            }

            DateTime baseDate = new DateTime(2025, 1, 1);
            string[] customers = { "华东供应链", "北方电子", "华南商贸", "西部智能", "海外事业部" };
            double totalAmount = 0;
            for (int i = 0; i < dataRows; i++)
            {
                int row = i + 2;
                int qty = (i % 40) + 1;
                double price = Math.Round(12.5 + (i % 900) * 1.35, 2);
                double amount = Math.Round(qty * price, 2);
                totalAmount += amount;
                bool alternate = i % 2 == 1;

                Cell order = detail.GetCell(row, 1);
                order.Value = "SO-" + (20260000 + i);
                order.Style = alternate ? textAltStyle : textStyle;

                Cell date = detail.GetCell(row, 2);
                date.Value = baseDate.AddDays(i % 730);
                date.Style = alternate ? dateAltStyle : dateStyle;

                Cell customer = detail.GetCell(row, 3);
                customer.Value = customers[i % customers.Length];
                customer.Style = alternate ? textAltStyle : textStyle;

                Cell quantity = detail.GetCell(row, 4);
                quantity.Value = qty;
                quantity.Style = alternate ? integerAltStyle : integerStyle;

                Cell unitPrice = detail.GetCell(row, 5);
                unitPrice.Value = price;
                unitPrice.Style = alternate ? moneyAltStyle : moneyStyle;

                Cell amountCell = detail.GetCell(row, 6);
                amountCell.Formula = "D" + row + "*E" + row;
                amountCell.Value = amount;
                amountCell.Style = alternate ? moneyAltStyle : moneyStyle;

                Cell urgent = detail.GetCell(row, 7);
                urgent.Value = i % 5 == 0;
                urgent.Style = alternate ? boolAltStyle : boolStyle;

                if (i % 11 == 0) detail.GetCell(row, 8).Value = "重点客户，优先发货";
                else if (i % 7 == 0) detail.GetCell(row, 8).Value = "促销订单";

                Cell margin = detail.GetCell(row, 9);
                margin.Formula = "F" + row + "*0.35";
                margin.Value = Math.Round(amount * 0.35, 2);
                margin.Style = alternate ? moneyAltStyle : moneyStyle;
            }

            detail.GetCell(dataRows + 3, 1).Value = "数据来源：ERP 系统导出";
            detail.GetCell(dataRows + 3, 1).Style = noteStyle;

            Worksheet summary = workbook.AddWorksheet("汇总");
            summary.SetColumnWidth(1, 18);
            summary.SetColumnWidth(2, 18);
            summary.SetColumnWidth(3, 16);
            summary.SetColumnWidth(4, 24);
            summary.SetRowHeight(1, 30);

            CellStyle summaryTitle = new CellStyle();
            summaryTitle.FontFamily = "Microsoft YaHei";
            summaryTitle.FontSize = 16;
            summaryTitle.Bold = true;
            summaryTitle.FontColor = OfficeColor.White;
            summaryTitle.FillColor = new OfficeColor(0x1F, 0x3A, 0x5F);
            summaryTitle.HorizontalAlignment = HorizontalAlignment.Center;

            summary.GetCell(1, 1).Value = "销售汇总看板";
            summary.GetCell(1, 1).Style = summaryTitle;
            summary.Merge("A1:D1");

            string[] summaryHeaders = { "指标", "数值", "单位", "说明" };
            for (int c = 0; c < summaryHeaders.Length; c++)
            {
                Cell cell = summary.GetCell(3, c + 1);
                cell.Value = summaryHeaders[c];
                cell.Style = headerStyle;
            }

            summary.GetCell(4, 1).Value = "订单行数";
            summary.GetCell(4, 2).Value = dataRows;
            summary.GetCell(4, 2).Style = integerStyle;
            summary.GetCell(4, 3).Value = "行";
            summary.GetCell(4, 4).Value = "覆盖 2025-2026 年度";
            summary.GetCell(5, 1).Value = "总金额";
            summary.GetCell(5, 2).Formula = "SUM(销售明细!F2:F" + (dataRows + 1) + ")";
            summary.GetCell(5, 2).Value = Math.Round(totalAmount, 2);
            summary.GetCell(5, 2).Style = moneyStyle;
            summary.GetCell(5, 3).Value = "元";
            summary.GetCell(6, 1).Value = "平均单价";
            summary.GetCell(6, 2).Value = Math.Round(totalAmount / Math.Max(1, dataRows), 2);
            summary.GetCell(6, 2).Style = moneyStyle;
            summary.GetCell(6, 3).Value = "元/行";
            summary.GetCell(7, 1).Value = "统计口径：仅统计已审核订单。";
            summary.Merge("A7:D7");

            Worksheet internalSheet = workbook.AddWorksheet("内部数据");
            internalSheet.IsHidden = true;
            internalSheet.GetCell(1, 1).Value = "机密-内部数据-2026";
            internalSheet.GetCell(2, 1).Value = "仅限管理层查看";
            internalSheet.GetCell(3, 1).Value = 42;

            return workbook;
        }

        private static void RunComplexExcelTest(string outputDirectory)
        {
            const int dataRows = 5000;
            OfficeWorkbook workbook = BuildComplexWorkbook(dataRows);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "complex-excel.xlsx"));
            Stopwatch watch = Stopwatch.StartNew();
            XlsxWriter.Write(workbook, path);
            watch.Stop();
            long writeMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(path, "xl/workbook.xml");

            watch.Restart();
            OfficeWorkbook loaded = XlsxReader.Load(path);
            watch.Stop();
            long readMs = watch.ElapsedMilliseconds;
            Require(loaded.Worksheets.Count == 3, "complex xlsx sheet count mismatch");
            Require(loaded.Worksheets[2].IsHidden, "complex xlsx hidden sheet flag lost");

            Worksheet detail = loaded.Worksheets[0];
            Require(detail.FrozenRows == 1 && detail.FrozenColumns == 2, "complex xlsx freeze panes lost");
            Require(detail.HiddenRows.Contains(10), "complex xlsx hidden row lost");
            Require(detail.HiddenColumns.Contains(9), "complex xlsx hidden column lost");
            Require(Convert.ToString(detail.GetCell(2, 1).Value) == "SO-20260000", "complex xlsx first order mismatch");
            Require(detail.GetCell(2, 2).Value is DateTime, "complex xlsx date cell missing");
            Require(Convert.ToInt32(detail.GetCell(2, 4).Value) == 1, "complex xlsx quantity mismatch");
            Require(Math.Abs(Convert.ToDouble(detail.GetCell(2, 5).Value) - 12.5) < 0.0001, "complex xlsx price mismatch");
            Require(detail.GetCell(2, 6).Formula == "D2*E2", "complex xlsx formula mismatch: " + detail.GetCell(2, 6).Formula);
            Require(Math.Abs(Convert.ToDouble(detail.GetCell(2, 6).Value) - 12.5) < 0.0001, "complex xlsx cached formula value mismatch");
            Require(detail.GetCell(2, 7).Value is bool && (bool)detail.GetCell(2, 7).Value, "complex xlsx boolean mismatch");
            int lastDataRow = dataRows + 1;
            Require(detail.GetCell(lastDataRow, 6).Formula == "D" + lastDataRow + "*E" + lastDataRow,
                "complex xlsx last-row formula mismatch");
            Require(Convert.ToDouble(detail.GetCell(lastDataRow, 6).Value) > 0, "complex xlsx last-row value missing");
            Require(Convert.ToString(detail.GetCell(lastDataRow + 2, 1).Value) == "数据来源：ERP 系统导出",
                "complex xlsx trailing note missing");

            watch.Restart();
            OfficeFile excelFile = OfficeFile.Load(path);
            string text = excelFile.ExtractText();
            watch.Stop();
            long extractMs = watch.ElapsedMilliseconds;
            RequireContains(text, "SO-20260000");
            RequireContains(text, "数据来源：ERP 系统导出");
            Require(text.IndexOf("机密-内部数据-2026", StringComparison.Ordinal) < 0, "hidden sheet text leaked into extraction");

            watch.Restart();
            OfficeDocument layout = WorkbookToDocumentConverter.Convert(loaded);
            watch.Stop();
            long layoutMs = watch.ElapsedMilliseconds;
            Require(layout.Sections.Count == 2, "complex xlsx visible section count mismatch: " + layout.Sections.Count);
            Table detailTable = FindFirstTable(layout.Sections[0].Blocks);
            Require(detailTable != null && detailTable.Rows.Count > dataRows,
                "complex xlsx layout table rows mismatch: " + (detailTable == null ? "null" : detailTable.Rows.Count.ToString()));

            WorkbookLayoutOptions withHidden = new WorkbookLayoutOptions();
            withHidden.IncludeHiddenWorksheets = true;
            OfficeDocument hiddenLayout = WorkbookToDocumentConverter.Convert(loaded, withHidden);
            Require(hiddenLayout.Sections.Count == 3, "complex xlsx hidden-inclusive section count mismatch");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "complex-excel.pdf"));
            watch.Restart();
            OfficeConverter.WorkbookToPdf(loaded, pdfPath);
            watch.Stop();
            long pdfMs = watch.ElapsedMilliseconds;
            Require(new FileInfo(pdfPath).Length > 10000, "complex xlsx pdf is unexpectedly small");

            watch.Restart();
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            string pdfText = pdf.ExtractText();
            watch.Stop();
            long pdfReadMs = watch.ElapsedMilliseconds;
            Require(pdf.PageCount > 10, "complex xlsx pdf page count mismatch: " + pdf.PageCount);
            string compactPdfText = pdfText.Replace(" ", string.Empty).Replace("\r", string.Empty)
                .Replace("\n", string.Empty).Replace("\t", string.Empty);
            Require(compactPdfText.IndexOf("SO-202600", StringComparison.Ordinal) >= 0,
                "complex xlsx pdf text missing first order row");
            int orderHits = 0;
            int searchIndex = 0;
            while ((searchIndex = pdfText.IndexOf("SO-2026", searchIndex, StringComparison.Ordinal)) >= 0)
            {
                orderHits++;
                searchIndex += 7;
            }
            Require(orderHits > 100, "complex xlsx pdf text missing order rows: " + orderHits);
            RequireContains(pdfText, "销售明细");
            Require(pdfText.IndexOf("机密-内部数据-2026", StringComparison.Ordinal) < 0, "hidden sheet text leaked into pdf");

            Require(writeMs < 15000, "complex xlsx write too slow: " + writeMs + " ms");
            Require(readMs < 15000, "complex xlsx read too slow: " + readMs + " ms");
            Require(extractMs < 15000, "complex xlsx extract too slow: " + extractMs + " ms");
            Require(layoutMs < 15000, "complex xlsx layout too slow: " + layoutMs + " ms");
            Require(pdfMs < 30000, "complex xlsx pdf too slow: " + pdfMs + " ms");
            Require(pdfReadMs < 15000, "complex xlsx pdf read too slow: " + pdfReadMs + " ms");

            Console.WriteLine("[complex-excel] rows=" + dataRows + ", cells=" + (dataRows * 9) + ", file="
                + new FileInfo(path).Length + " bytes, write=" + writeMs + " ms, read=" + readMs
                + " ms, extract=" + extractMs + " ms, layout=" + layoutMs + " ms, pdf=" + pdf.PageCount
                + " pages/" + new FileInfo(pdfPath).Length + " bytes, render=" + pdfMs + " ms, pdf read=" + pdfReadMs + " ms");
        }

        private static CellStyle BorderedStyle(string fontFamily, double fontSize, bool bold, OfficeColor? fill, HorizontalAlignment alignment)
        {
            CellStyle style = new CellStyle();
            style.FontFamily = fontFamily;
            style.FontSize = fontSize;
            style.Bold = bold;
            style.FillColor = fill;
            style.HorizontalAlignment = alignment;
            style.VerticalAlignment = VerticalAlignment.Middle;
            style.TopBorder = BorderLine.Thin();
            style.BottomBorder = BorderLine.Thin();
            style.LeftBorder = BorderLine.Thin();
            style.RightBorder = BorderLine.Thin();
            return style;
        }

        /// <summary>Builds a 12-sheet monthly attendance workbook shaped like a typical HR attendance form.</summary>
        private static OfficeWorkbook BuildAttendanceWorkbook(int year, out int[,] attendDays, out int[,] leaveHours)
        {
            string[] names = { "张三", "李四", "王二", "赵五" };
            string[] shifts = { "上午", "下午", "晚上", "请假" };
            string[] weekdayNames = { "日", "一", "二", "三", "四", "五", "六" };

            const int FirstDayColumn = 4;
            const int AttendColumn = 35;
            const int LeaveColumn = 36;
            const int SummaryFirstColumn = 38;
            const int LastColumn = 41;

            attendDays = new int[12, names.Length];
            leaveHours = new int[12, names.Length];

            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.Metadata.Title = "员工考勤表";
            workbook.Metadata.Author = "Wisdom.Office";
            workbook.Metadata.Subject = "月度考勤报表";

            OfficeColor titleFill = new OfficeColor(0xC5, 0xD5, 0xE3);
            OfficeColor headerFill = new OfficeColor(0xDD, 0xEB, 0xF7);
            OfficeColor weekendFill = new OfficeColor(0xFF, 0xD9, 0x66);
            OfficeColor labelFill = new OfficeColor(0xFC, 0xE4, 0xD6);
            OfficeColor leaveLabelFill = new OfficeColor(0xA9, 0xDF, 0xDF);
            OfficeColor selectorFill = new OfficeColor(0xFF, 0xF2, 0xCC);
            OfficeColor totalFill = new OfficeColor(0xFD, 0xE9, 0xD9);
            OfficeColor holidayFont = new OfficeColor(0xC0, 0x00, 0x00);

            for (int month = 1; month <= 12; month++)
            {
                Worksheet sheet = workbook.AddWorksheet(month + "月");
                sheet.ShowGridLines = true;
                sheet.FrozenRows = 5;
                sheet.FrozenColumns = 3;
                sheet.FitToPage = true;
                sheet.PrintSetup.IsLandscape = true;
                sheet.PrintSetup.MarginLeft = 28;
                sheet.PrintSetup.MarginRight = 28;
                sheet.PrintSetup.MarginTop = 32;
                sheet.PrintSetup.MarginBottom = 32;
                sheet.SetColumnWidth(1, 5);
                sheet.SetColumnWidth(2, 9);
                sheet.SetColumnWidth(3, 6.5);
                for (int d = 1; d <= 31; d++) sheet.SetColumnWidth(FirstDayColumn + d - 1, 3.6);
                sheet.SetColumnWidth(AttendColumn, 9);
                sheet.SetColumnWidth(LeaveColumn, 9);
                sheet.SetColumnWidth(SummaryFirstColumn, 5);
                sheet.SetColumnWidth(SummaryFirstColumn + 1, 9);
                sheet.SetColumnWidth(SummaryFirstColumn + 2, 9);
                sheet.SetColumnWidth(SummaryFirstColumn + 3, 9);
                sheet.SetRowHeight(1, 26);
                sheet.SetRowHeight(2, 22);
                sheet.SetRowHeight(4, 18);
                sheet.SetRowHeight(5, 15);

                int daysInMonth = DateTime.DaysInMonth(year, month);
                string lastColumn = CellAddress.ColumnName(LastColumn);

                CellStyle titleStyle = BorderedStyle("宋体", 18, true, titleFill, HorizontalAlignment.Center);
                CellStyle subTitleStyle = BorderedStyle("宋体", 14, true, titleFill, HorizontalAlignment.Center);
                titleStyle.BottomBorder = null;
                subTitleStyle.TopBorder = null;

                Cell title = sheet.GetCell(1, 1);
                title.Value = "员工考勤表";
                title.Style = titleStyle;
                for (int c = 2; c <= LastColumn; c++) sheet.GetCell(1, c).Style = titleStyle;
                sheet.Merge("A1:" + lastColumn + "1");

                Cell subTitle = sheet.GetCell(2, 1);
                subTitle.Value = year + "年" + month + "月考勤表";
                subTitle.Style = subTitleStyle;
                for (int c = 2; c <= LastColumn; c++) sheet.GetCell(2, c).Style = subTitleStyle;
                sheet.Merge("A2:" + lastColumn + "2");

                CellStyle selectorLabel = BorderedStyle("宋体", 10, true, null, HorizontalAlignment.Center);
                CellStyle selectorValue = BorderedStyle("宋体", 10, true, selectorFill, HorizontalAlignment.Center);
                sheet.GetCell(3, 1).Value = "选择时间";
                sheet.GetCell(3, 1).Style = selectorLabel;
                sheet.GetCell(3, 2).Style = selectorLabel;
                sheet.Merge("A3:B3");
                sheet.GetCell(3, 3).Value = year;
                sheet.GetCell(3, 3).Style = selectorValue;
                sheet.GetCell(3, 4).Value = "年";
                sheet.GetCell(3, 4).Style = selectorLabel;
                sheet.GetCell(3, 5).Value = month;
                sheet.GetCell(3, 5).Style = selectorValue;
                sheet.GetCell(3, 6).Value = "月";
                sheet.GetCell(3, 6).Style = selectorLabel;

                CellStyle dayHeader = BorderedStyle("宋体", 9, true, headerFill, HorizontalAlignment.Center);
                CellStyle weekendHeader = BorderedStyle("宋体", 9, true, weekendFill, HorizontalAlignment.Center);
                CellStyle weekHeader = BorderedStyle("宋体", 8, false, headerFill, HorizontalAlignment.Center);
                CellStyle weekEndHeader = BorderedStyle("宋体", 8, false, weekendFill, HorizontalAlignment.Center);
                CellStyle corner = BorderedStyle("宋体", 9, true, headerFill, HorizontalAlignment.Center);
                corner.WrapText = true;

                for (int r = 4; r <= 5; r++)
                {
                    sheet.GetCell(r, 1).Style = corner;
                    sheet.GetCell(r, 2).Style = corner;
                    sheet.GetCell(r, 3).Style = corner;
                    sheet.GetCell(r, AttendColumn).Style = corner;
                    sheet.GetCell(r, LeaveColumn).Style = corner;
                }
                sheet.GetCell(4, 1).Value = "序号";
                sheet.GetCell(4, 2).Value = "姓名";
                sheet.GetCell(4, 3).Value = "班次";
                sheet.GetCell(4, AttendColumn).Value = "出勤天数（天）";
                sheet.GetCell(4, LeaveColumn).Value = "请假工时（小时）";
                sheet.Merge("A4:A5");
                sheet.Merge("B4:B5");
                sheet.Merge("C4:C5");
                sheet.Merge(CellAddress.ColumnName(AttendColumn) + "4:" + CellAddress.ColumnName(AttendColumn) + "5");
                sheet.Merge(CellAddress.ColumnName(LeaveColumn) + "4:" + CellAddress.ColumnName(LeaveColumn) + "5");

                for (int d = 1; d <= 31; d++)
                {
                    int column = FirstDayColumn + d - 1;
                    bool inMonth = d <= daysInMonth;
                    DayOfWeek dayOfWeek = inMonth ? new DateTime(year, month, d).DayOfWeek : DayOfWeek.Monday;
                    bool isWeekend = !inMonth || dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday;
                    Cell dayCell = sheet.GetCell(4, column);
                    Cell weekCell = sheet.GetCell(5, column);
                    dayCell.Style = isWeekend ? weekendHeader : dayHeader;
                    weekCell.Style = isWeekend ? weekEndHeader : weekHeader;
                    if (inMonth)
                    {
                        dayCell.Value = d;
                        weekCell.Value = weekdayNames[(int)dayOfWeek];
                    }
                }

                CellStyle nameStyle = BorderedStyle("宋体", 10, false, labelFill, HorizontalAlignment.Center);
                CellStyle shiftStyle = BorderedStyle("宋体", 9, false, labelFill, HorizontalAlignment.Center);
                CellStyle leaveShiftStyle = BorderedStyle("宋体", 9, true, leaveLabelFill, HorizontalAlignment.Center);
                CellStyle markStyle = BorderedStyle("宋体", 9, false, null, HorizontalAlignment.Center);
                CellStyle holidayStyle = BorderedStyle("宋体", 9, true, null, HorizontalAlignment.Center);
                holidayStyle.FontColor = holidayFont;
                CellStyle hoursStyle = BorderedStyle("宋体", 9, false, null, HorizontalAlignment.Center);
                hoursStyle.NumberFormat = "0";
                CellStyle totalStyle = BorderedStyle("宋体", 11, true, totalFill, HorizontalAlignment.Center);

                string firstDay = CellAddress.ColumnName(FirstDayColumn);
                string lastDay = CellAddress.ColumnName(FirstDayColumn + daysInMonth - 1);
                int monthPresentDays = 0;
                int monthVacationHours = 0;

                for (int e = 0; e < names.Length; e++)
                {
                    int baseRow = 6 + e * 4;
                    for (int r = baseRow; r < baseRow + 4; r++)
                    {
                        sheet.GetCell(r, 1).Style = nameStyle;
                        sheet.GetCell(r, 2).Style = nameStyle;
                        sheet.GetCell(r, AttendColumn).Style = totalStyle;
                        sheet.GetCell(r, LeaveColumn).Style = totalStyle;
                        for (int d = 1; d <= 31; d++)
                        {
                            sheet.GetCell(r, FirstDayColumn + d - 1).Style = markStyle;
                        }
                    }
                    sheet.GetCell(baseRow, 1).Value = e + 1;
                    sheet.GetCell(baseRow, 2).Value = names[e];
                    sheet.Merge("A" + baseRow + ":A" + (baseRow + 3));
                    sheet.Merge("B" + baseRow + ":B" + (baseRow + 3));
                    sheet.Merge(CellAddress.ColumnName(AttendColumn) + baseRow + ":" + CellAddress.ColumnName(AttendColumn) + (baseRow + 3));
                    sheet.Merge(CellAddress.ColumnName(LeaveColumn) + baseRow + ":" + CellAddress.ColumnName(LeaveColumn) + (baseRow + 3));

                    for (int s = 0; s < shifts.Length; s++)
                    {
                        Cell shiftCell = sheet.GetCell(baseRow + s, 3);
                        shiftCell.Value = shifts[s];
                        shiftCell.Style = s == 3 ? leaveShiftStyle : shiftStyle;
                    }

                    int presentDays = 0;
                    int vacationHours = 0;
                    for (int d = 1; d <= daysInMonth; d++)
                    {
                        int column = FirstDayColumn + d - 1;
                        DayOfWeek dayOfWeek = new DateTime(year, month, d).DayOfWeek;
                        if (dayOfWeek == DayOfWeek.Saturday || dayOfWeek == DayOfWeek.Sunday) continue;

                        bool[] onLeave =
                        {
                            (d + e) % 11 == 0,
                            (d * (e + 3)) % 13 == 0,
                            (d + e * 2) % 17 == 0
                        };
                        bool fullDay = !onLeave[0] && !onLeave[1] && !onLeave[2];
                        if (fullDay) presentDays++;

                        for (int s = 0; s < 3; s++)
                        {
                            Cell cell = sheet.GetCell(baseRow + s, column);
                            if (onLeave[s])
                            {
                                cell.Value = "假";
                                cell.Style = holidayStyle;
                                Cell hoursCell = sheet.GetCell(baseRow + 3, column);
                                hoursCell.Value = 4;
                                hoursCell.Style = hoursStyle;
                                vacationHours += 4;
                            }
                            else
                            {
                                cell.Value = "√";
                            }
                        }
                    }

                    attendDays[month - 1, e] = presentDays;
                    leaveHours[month - 1, e] = vacationHours;
                    monthPresentDays += presentDays;
                    monthVacationHours += vacationHours;

                    Cell attendCell = sheet.GetCell(baseRow, AttendColumn);
                    attendCell.Formula = "COUNTIFS(" + firstDay + baseRow + ":" + lastDay + baseRow + ",\"√\","
                        + firstDay + (baseRow + 1) + ":" + lastDay + (baseRow + 1) + ",\"√\","
                        + firstDay + (baseRow + 2) + ":" + lastDay + (baseRow + 2) + ",\"√\")";
                    attendCell.Value = presentDays;

                    Cell leaveCell = sheet.GetCell(baseRow, LeaveColumn);
                    leaveCell.Formula = "SUM(" + firstDay + (baseRow + 3) + ":" + lastDay + (baseRow + 3) + ")";
                    leaveCell.Value = vacationHours;
                }

                int noteRow = 6 + names.Length * 4;
                CellStyle noteStyle = BorderedStyle("宋体", 9, false, null, HorizontalAlignment.Left);
                noteStyle.TopBorder = null;
                noteStyle.BottomBorder = null;
                noteStyle.LeftBorder = null;
                noteStyle.RightBorder = null;
                Cell note = sheet.GetCell(noteRow, 1);
                note.Value = "说明：√ 表示出勤，假 表示请假（半天 4 小时），空白表示休息日。";
                note.Style = noteStyle;
                sheet.Merge("A" + noteRow + ":" + CellAddress.ColumnName(LeaveColumn) + noteRow);

                CellStyle summaryTitle = BorderedStyle("宋体", 11, true, titleFill, HorizontalAlignment.Center);
                CellStyle summaryHeader = BorderedStyle("宋体", 9, true, headerFill, HorizontalAlignment.Center);
                summaryHeader.WrapText = true;
                CellStyle summaryText = BorderedStyle("宋体", 9, false, null, HorizontalAlignment.Center);
                CellStyle summaryTotal = BorderedStyle("宋体", 9, true, totalFill, HorizontalAlignment.Center);

                sheet.SetRowHeight(4, 18);
                Cell summaryCell = sheet.GetCell(4, SummaryFirstColumn);                summaryCell.Value = year + "年" + month + "考勤汇总";
                for (int c = SummaryFirstColumn; c <= LastColumn; c++) sheet.GetCell(4, c).Style = summaryTitle;
                sheet.Merge(CellAddress.ColumnName(SummaryFirstColumn) + "4:" + CellAddress.ColumnName(LastColumn) + "4");

                string[] summaryHeaders = { "序号", "姓名", "出勤天数", "请假工时" };
                for (int c = 0; c < summaryHeaders.Length; c++)
                {
                    Cell cell = sheet.GetCell(5, SummaryFirstColumn + c);
                    cell.Value = summaryHeaders[c];
                    cell.Style = summaryHeader;
                }

                for (int e = 0; e < names.Length; e++)
                {
                    int row = 6 + e;
                    int sourceRow = 6 + e * 4;
                    sheet.GetCell(row, SummaryFirstColumn).Value = e + 1;
                    sheet.GetCell(row, SummaryFirstColumn).Style = summaryText;
                    sheet.GetCell(row, SummaryFirstColumn + 1).Value = names[e];
                    sheet.GetCell(row, SummaryFirstColumn + 1).Style = summaryText;
                    Cell rowAttend = sheet.GetCell(row, SummaryFirstColumn + 2);
                    rowAttend.Formula = CellAddress.ColumnName(AttendColumn) + sourceRow;
                    rowAttend.Value = attendDays[month - 1, e];
                    rowAttend.Style = summaryText;
                    Cell rowLeave = sheet.GetCell(row, SummaryFirstColumn + 3);
                    rowLeave.Formula = CellAddress.ColumnName(LeaveColumn) + sourceRow;
                    rowLeave.Value = leaveHours[month - 1, e];
                    rowLeave.Style = summaryText;
                }

                int totalRow = 6 + names.Length;
                sheet.GetCell(totalRow, SummaryFirstColumn).Value = "合计";
                sheet.GetCell(totalRow, SummaryFirstColumn).Style = summaryTotal;
                sheet.GetCell(totalRow, SummaryFirstColumn + 1).Style = summaryTotal;
                sheet.Merge(CellAddress.ColumnName(SummaryFirstColumn) + totalRow + ":"
                    + CellAddress.ColumnName(SummaryFirstColumn + 1) + totalRow);
                Cell totalAttend = sheet.GetCell(totalRow, SummaryFirstColumn + 2);
                totalAttend.Formula = "SUM(" + CellAddress.ColumnName(SummaryFirstColumn + 2) + "6:"
                    + CellAddress.ColumnName(SummaryFirstColumn + 2) + (totalRow - 1) + ")";
                totalAttend.Value = monthPresentDays;
                totalAttend.Style = summaryTotal;
                Cell totalLeave = sheet.GetCell(totalRow, SummaryFirstColumn + 3);
                totalLeave.Formula = "SUM(" + CellAddress.ColumnName(SummaryFirstColumn + 3) + "6:"
                    + CellAddress.ColumnName(SummaryFirstColumn + 3) + (totalRow - 1) + ")";
                totalLeave.Value = monthVacationHours;
                totalLeave.Style = summaryTotal;
            }

            return workbook;
        }

        private static void RunAttendanceReportTest(string outputDirectory)
        {
            const int Year = 2026;
            int[,] attendDays;
            int[,] leaveHours;
            OfficeWorkbook workbook = BuildAttendanceWorkbook(Year, out attendDays, out leaveHours);
            int totalAttendDays = 0;
            for (int i = 0; i < attendDays.GetLength(1); i++) totalAttendDays += attendDays[0, i];

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "attendance-report.xlsx"));
            Stopwatch watch = Stopwatch.StartNew();
            XlsxWriter.Write(workbook, path);
            watch.Stop();
            long writeMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(path, "xl/workbook.xml");

            watch.Restart();
            OfficeWorkbook loaded = XlsxReader.Load(path);
            watch.Stop();
            long readMs = watch.ElapsedMilliseconds;
            Require(loaded.Worksheets.Count == 12, "attendance sheet count mismatch: " + loaded.Worksheets.Count);
            Require(loaded.Worksheets[0].Name == "1月" && loaded.Worksheets[11].Name == "12月", "attendance sheet names mismatch");

            Worksheet january = loaded.Worksheets[0];
            Require(january.FrozenRows == 5 && january.FrozenColumns == 3, "attendance freeze panes lost");
            Require(Math.Abs(january.GetColumnWidth(10) - 3.6) < 0.01, "attendance day column width lost");
            Require(january.MergedRanges.Count == 27, "attendance merged range count mismatch: " + january.MergedRanges.Count);
            Require(Convert.ToString(january.GetCell(1, 1).Value) == "员工考勤表", "attendance title mismatch");
            Require(Convert.ToString(january.GetCell(2, 1).Value) == "2026年1月考勤表", "attendance subtitle mismatch");
            Require(Convert.ToString(january.GetCell(3, 3).Value) == "2026", "attendance selector year mismatch");
            Require(Convert.ToString(january.GetCell(4, 4).Value) == "1", "attendance first day mismatch");
            Require(Convert.ToString(january.GetCell(4, 34).Value) == "31", "attendance last day mismatch");
            Require(Convert.ToString(january.GetCell(5, 4).Value) == "四", "attendance weekday mismatch");
            Require(Convert.ToString(january.GetCell(6, 2).Value) == "张三", "attendance employee name mismatch");
            Require(Convert.ToString(january.GetCell(6, 3).Value) == "上午", "attendance shift label mismatch");
            Require(Convert.ToString(january.GetCell(9, 3).Value) == "请假", "attendance leave shift label mismatch");

            Cell januaryAttend = january.GetCell(6, 35);
            Require(januaryAttend.Formula != null && januaryAttend.Formula.StartsWith("COUNTIFS(", StringComparison.Ordinal),
                "attendance COUNTIFS formula lost: " + januaryAttend.Formula);
            Require(Convert.ToInt32(januaryAttend.Value) == attendDays[0, 0],
                "attendance cached days mismatch: " + januaryAttend.Value + " != " + attendDays[0, 0]);
            Cell januaryLeave = january.GetCell(6, 36);
            Require(januaryLeave.Formula == "SUM(D9:AH9)", "attendance SUM formula mismatch: " + januaryLeave.Formula);
            Require(Convert.ToInt32(januaryLeave.Value) == leaveHours[0, 0],
                "attendance cached hours mismatch: " + januaryLeave.Value + " != " + leaveHours[0, 0]);
            Require(Convert.ToString(january.GetCell(10, 38).Value) == "合计", "attendance summary total label mismatch");
            Require(Convert.ToInt32(january.GetCell(10, 40).Value) == totalAttendDays,
                "attendance summary total value mismatch");
            Worksheet december = loaded.Worksheets[11];
            Require(Convert.ToInt32(december.GetCell(6, 35).Value) == attendDays[11, 0],
                "december cached days mismatch: " + december.GetCell(6, 35).Value + " != " + attendDays[11, 0]);
            Require(Convert.ToString(december.GetCell(2, 1).Value) == "2026年12月考勤表", "december subtitle mismatch");

            Cell weekendHeader = january.GetCell(4, 7);
            Require(weekendHeader.Style != null && weekendHeader.Style.FillColor.HasValue, "attendance weekend fill lost");
            RequireColor(weekendHeader.Style.FillColor.Value, new OfficeColor(0xFF, 0xD9, 0x66), "attendance weekend fill");
            Cell shiftLabel = january.GetCell(6, 3);
            Require(shiftLabel.Style != null && shiftLabel.Style.TopBorder != null
                && shiftLabel.Style.TopBorder.Style == BorderStyle.Single, "attendance grid border lost");
            Require(january.GetCell(4, 35).Style != null && january.GetCell(4, 35).Style.WrapText, "attendance header wrap lost");

            int markedCells = 0;
            for (int r = 6; r <= 21; r++)
            {
                for (int c = 4; c <= 34; c++)
                {
                    if (january.FindCell(r, c) != null && january.FindCell(r, c).Value != null) markedCells++;
                }
            }
            Require(markedCells > 60, "attendance mark cells missing: " + markedCells);

            watch.Restart();
            OfficeDocument layout = WorkbookToDocumentConverter.Convert(loaded);
            watch.Stop();
            long layoutMs = watch.ElapsedMilliseconds;
            Require(layout.Sections.Count == 12, "attendance layout section count mismatch: " + layout.Sections.Count);
            List<Table> januaryTables = FindTables(layout.Sections[0].Blocks);
            Require(januaryTables.Count == 2, "attendance wide sheet band count mismatch: " + januaryTables.Count);
            Table firstBand = januaryTables[0];
            Table secondBand = januaryTables[1];
            Require(firstBand.Rows.Count == 22, "attendance layout row count mismatch: " + firstBand.Rows.Count);
            Require(firstBand.Columns.Count < 41, "attendance wide sheet was not split across pages");
            Require(firstBand.Columns.Count + secondBand.Columns.Count == 44,
                "attendance bands must cover all columns and repeat the 3 frozen ones: "
                + firstBand.Columns.Count + "+" + secondBand.Columns.Count);
            Require(!firstBand.Autofit, "attendance table should use fixed layout widths");

            PageSetup januarySetup = layout.Sections[0].PageSetup;
            double pageContent = januarySetup.Width - januarySetup.MarginLeft - januarySetup.MarginRight;
            Require(januarySetup.Width > januarySetup.Height, "attendance page should be landscape");
            for (int t = 0; t < januaryTables.Count; t++)
            {
                double bandWidth = 0;
                for (int c = 0; c < januaryTables[t].Columns.Count; c++) bandWidth += januaryTables[t].Columns[c].Width;
                Require(bandWidth <= pageContent + 0.01,
                    "attendance band " + t + " exceeds printable width: " + bandWidth + " > " + pageContent);
            }
            Require(ContainsCellText(firstBand, "张三"), "attendance first band lost employee name");
            Require(ContainsCellText(secondBand, "张三"), "attendance second band lost repeated name column");
            Require(ContainsCellText(secondBand, "出勤天数（天）"), "attendance second band lost attendance header");
            Require(ContainsCellText(firstBand, "27"), "attendance first band lost day 27");
            Require(ContainsCellText(secondBand, "31"), "attendance second band lost end-of-month days");
            Require(!ContainsCellText(firstBand, "2026年1考勤汇总"), "attendance summary block duplicated into first band");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "attendance-report.pdf"));
            watch.Restart();
            OfficeConverter.WorkbookToPdf(loaded, pdfPath);
            watch.Stop();
            long pdfMs = watch.ElapsedMilliseconds;
            Require(new FileInfo(pdfPath).Length > 10000, "attendance pdf is unexpectedly small");

            watch.Restart();
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            string pdfText = pdf.ExtractText();
            watch.Stop();
            long pdfReadMs = watch.ElapsedMilliseconds;
            Require(pdf.PageCount >= 24, "attendance pdf page count mismatch: " + pdf.PageCount);
            string compact = pdfText.Replace(" ", string.Empty).Replace("\r", string.Empty)
                .Replace("\n", string.Empty).Replace("\t", string.Empty);
            RequireContains(compact, "员工考勤表");
            RequireContains(compact, "2026年1月考勤表");
            Require(CountOccurrences(compact, "2026") >= 36,
                "attendance year cell broken across lines: " + CountOccurrences(compact, "2026"));
            Require(CountOccurrences(compact, "张三") >= 24,
                "attendance employee name missing from bands: " + CountOccurrences(compact, "张三"));
            RequireContains(compact, "出勤天数");
            RequireContains(compact, "请假工时");
            PdfPage bandPage = pdf.Pages[1];
            Require(bandPage.Width > bandPage.Height, "attendance pdf page should be landscape");
            IReadOnlyList<PdfTable> firstPageTables = pdf.Pages[0].GetTables();
            IReadOnlyList<PdfTable> bandTables = bandPage.GetTables();
            Require(firstPageTables.Count == 1 && bandTables.Count == 1,
                "attendance pdf page should hold one column band: " + firstPageTables.Count
                + "/" + bandTables.Count);
            Require(firstPageTables[0].ColumnCount == firstBand.Columns.Count,
                "attendance pdf first band geometry mismatch: " + firstPageTables[0].ColumnCount
                + " != " + firstBand.Columns.Count);
            Require(bandTables[0].ColumnCount == secondBand.Columns.Count,
                "attendance pdf second band geometry mismatch: " + bandTables[0].ColumnCount
                + " != " + secondBand.Columns.Count);
            bool repeatedIdentity = false;
            for (int r = 0; r < bandTables[0].RowCount && !repeatedIdentity; r++)
            {
                repeatedIdentity = bandTables[0].GetCell(r, 0) == "序号";
            }
            Require(repeatedIdentity, "attendance pdf second band lost repeated identity column");
            Require(compact.IndexOf("√", StringComparison.Ordinal) >= 0, "attendance marks missing from pdf text");
            int titleHits = 0;
            int cursor = 0;
            while ((cursor = compact.IndexOf("考勤表", cursor, StringComparison.Ordinal)) >= 0)
            {
                titleHits++;
                cursor += 3;
            }
            Require(titleHits >= 24, "attendance pdf missing monthly titles: " + titleHits);

            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "attendance-report.docx"));
            watch.Restart();
            OfficeConverter.Convert(path, docxPath);
            watch.Stop();
            long docxMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(docxPath, "word/document.xml");
            OfficeDocument docx = DocxReader.Load(docxPath);
            Require(docx.Sections.Count == 12, "attendance docx section count mismatch: " + docx.Sections.Count);
            List<Table> docxTables = FindTables(docx.Sections[0].Blocks);
            Require(docxTables.Count == 2, "attendance docx band count mismatch: " + docxTables.Count);
            int docxTableTotal = 0;
            for (int s = 0; s < docx.Sections.Count; s++) docxTableTotal += FindTables(docx.Sections[s].Blocks).Count;
            Require(docxTableTotal == 24, "attendance docx total band count mismatch: " + docxTableTotal);
            Require(docx.Sections[0].PageSetup.Width > docx.Sections[0].PageSetup.Height,
                "attendance docx page should be landscape");
            Require(docxTables[0].Columns.Count == firstBand.Columns.Count,
                "attendance docx first band columns mismatch: " + docxTables[0].Columns.Count
                + " != " + firstBand.Columns.Count);
            Require(ContainsCellText(docxTables[0], "张三"), "attendance docx first band lost employee name");
            Require(ContainsCellText(docxTables[1], "张三"), "attendance docx second band lost repeated name column");
            Require(ContainsCellText(docxTables[1], "请假工时"), "attendance docx second band lost leave header");
            Require(ContainsCellText(docxTables[1], "出勤天数（天）"), "attendance docx second band lost attendance header");
            string documentXml = ReadPackageEntryText(docxPath, "word/document.xml");
            Require(CountOccurrences(documentXml, "autofit") == 0,
                "attendance docx tables should not auto-fit: " + CountOccurrences(documentXml, "autofit"));
            Require(CountOccurrences(documentXml, "fixed") == docxTableTotal,
                "attendance docx tables should use fixed layout: " + CountOccurrences(documentXml, "fixed"));
            string docxText = GetBlocksText(docx.Sections[0].Blocks);
            RequireContains(docxText.Replace(" ", string.Empty), "请假工时");

            Console.WriteLine("[attendance] sheets=12, merged/sheet=" + january.MergedRanges.Count
                + ", markCells=" + markedCells + ", file=" + new FileInfo(path).Length + " bytes, write=" + writeMs
                + " ms, read=" + readMs + " ms, layout=" + layoutMs + " ms, pdf=" + pdf.PageCount + " pages/"
                + new FileInfo(pdfPath).Length + " bytes, render=" + pdfMs + " ms, pdf read=" + pdfReadMs
                + " ms, docx=" + new FileInfo(docxPath).Length + " bytes/" + docxMs + " ms");
        }

        private static bool HasExactCellText(Table table, string text)
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                {
                    if (GetBlocksText(table.Rows[r].Cells[c].Blocks).Trim() == text) return true;
                }
            }
            return false;
        }

        private static void RunRealAttendanceTest(string path, string outputDirectory)
        {
            Stopwatch watch = new Stopwatch();
            watch.Restart();
            OfficeWorkbook workbook = XlsxReader.Load(path);
            watch.Stop();
            long readMs = watch.ElapsedMilliseconds;

            Require(workbook.Worksheets.Count == 5, "real workbook sheet count mismatch: " + workbook.Worksheets.Count);
            Worksheet reference = workbook.Worksheets[0];
            Require(reference.Name == "数据引用区域" && reference.IsHidden, "real workbook hidden sheet lost");
            Require(reference.LastRow == 1556, "real workbook reference sheet truncated: " + reference.LastRow);

            Worksheet september = workbook.Worksheets[1];
            Require(september.Name == "9月考勤", "real workbook sheet 1 name mismatch: " + september.Name);
            Require(september.MergedRanges.Count == 16, "real workbook merge count mismatch: " + september.MergedRanges.Count);
            Require(september.HasPrintArea, "print area not parsed");
            Require(september.PrintAreaFirstRow == 2 && september.PrintAreaLastRow == 16,
                "print area mismatch: " + september.PrintAreaFirstRow + "-" + september.PrintAreaLastRow);
            Require(september.PrintAreaFirstColumn == 1 && september.PrintAreaLastColumn == 36,
                "print area columns mismatch: " + september.PrintAreaFirstColumn + "-" + september.PrintAreaLastColumn);
            Require(!september.ShowZeros, "sheetView showZeros not parsed");
            Require(september.CenterHorizontally, "printOptions horizontalCentered not parsed");
            Require(september.PrintSetup.IsLandscape, "real workbook landscape lost");
            Require(workbook.Worksheets[2].PrintSetup.IsLandscape == false, "10月 should stay portrait");

            Cell firstDay = september.FindCell(5, 3);
            Require(firstDay != null && firstDay.DisplayText() == "1",
                "excel single-field date format not applied: " + (firstDay == null ? "null" : firstDay.DisplayText()));
            Cell lastDay = september.FindCell(5, september.PrintAreaFirstColumn + 31);
            Require(lastDay != null && lastDay.DisplayText() == "30",
                "day header mismatch: " + (lastDay == null ? "null" : lastDay.DisplayText()));
            Cell weekday = workbook.Worksheets[3].FindCell(4, 3);
            Require(weekday != null && weekday.DisplayText() != "aaa",
                "weekday number format leaked as literal: " + (weekday == null ? "null" : weekday.DisplayText()));
            Cell zeroCell = september.FindCell(7, 35);
            Require(zeroCell != null && zeroCell.Value is double && (double)zeroCell.Value == 0,
                "expected a zero summary cell for the showZeros check");

            watch.Restart();
            OfficeDocument layout = WorkbookToDocumentConverter.Convert(workbook);
            watch.Stop();
            long layoutMs = watch.ElapsedMilliseconds;
            Require(layout.Sections.Count == 4, "real workbook layout sections mismatch: " + layout.Sections.Count);

            List<Table> bands = FindTables(layout.Sections[0].Blocks);
            Require(bands.Count >= 2, "real wide sheet should split into column bands: " + bands.Count);
            int bandColumns = 0;
            for (int t = 0; t < bands.Count; t++)
            {
                // 每个列带只打印到自己真正的内容末尾，尾部空行不再独占整页。
                Require(bands[t].Rows.Count > 0 && bands[t].Rows.Count <= september.PrintAreaLastRow - september.PrintAreaFirstRow + 1,
                    "band rows should be trimmed to their own content in band " + t + ": " + bands[t].Rows.Count);
                Require(bands[t].Alignment == TableAlignment.Center, "real sheet should center each band");
                Require(!HasExactCellText(bands[t], "0"), "zero values should be hidden in band " + t);
                bandColumns += bands[t].Columns.Count;
            }
            Require(bands[0].Columns.Count < september.PrintAreaLastColumn,
                "a landscape A4 page cannot hold every print-area column: " + bands[0].Columns.Count);
            Require(bandColumns > september.PrintAreaLastColumn,
                "later bands must repeat the identity columns: " + bandColumns);

            List<Table> october = FindTables(layout.Sections[1].Blocks);
            Require(october.Count >= 2, "10月 should split into column bands: " + october.Count);
            for (int t = 0; t < october.Count; t++)
            {
                Require(october[t].Rows.Count > 0 && october[t].Rows.Count <= workbook.Worksheets[2].LastRow,
                    "sheet without a print area should still trim band " + t + ": " + october[t].Rows.Count);
            }

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "real-attendance.pdf"));
            watch.Restart();
            OfficeConverter.WorkbookToPdf(workbook, pdfPath);
            watch.Stop();
            long pdfMs = watch.ElapsedMilliseconds;
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            string pdfText = pdf.ExtractText();
            string compact = pdfText.Replace(" ", string.Empty).Replace("\r", string.Empty)
                .Replace("\n", string.Empty).Replace("\t", string.Empty);
            RequireContains(compact, "9月考勤表");
            // Narrow label columns wrap, so assert on the fragment that survives a wrap.
            RequireContains(compact, "个人签");
            RequireContains(compact, "星期");
            RequireContains(compact, "姚蕾");
            Require(compact.IndexOf("√", StringComparison.Ordinal) >= 0, "attendance marks missing from pdf text");
            int blankPages = 0;
            for (int p = 0; p < pdf.PageCount; p++)
            {
                IReadOnlyList<PdfTextLine> lines = pdf.Pages[p].GetTextLines();
                bool filled = false;
                for (int l = 0; l < lines.Count && !filled; l++)
                {
                    filled = lines[l].Text != null && lines[l].Text.Trim().Length > 0;
                }
                if (!filled) blankPages++;
            }
            Require(blankPages == 0, "real workbook export has blank pages: " + blankPages + "/" + pdf.PageCount);

            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "real-attendance.docx"));
            watch.Restart();
            OfficeConverter.Convert(path, docxPath);
            watch.Stop();
            long docxMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(docxPath, "word/document.xml");
            OfficeDocument docx = DocxReader.Load(docxPath);
            Require(docx.Sections.Count == 4, "real docx section count mismatch: " + docx.Sections.Count);
            string docxText = GetBlocksText(docx.Sections[0].Blocks).Replace(" ", string.Empty);
            RequireContains(docxText, "姚蕾");
            Require(docx.Sections[0].PageSetup.IsLandscape, "real docx should keep the landscape setup");
            int docxBands = FindTables(docx.Sections[0].Blocks).Count;
            Require(docxBands == bands.Count, "docx band count mismatch: " + docxBands + " != " + bands.Count);
            List<Table> docxTables = FindTables(docx.Sections[0].Blocks);
            for (int t = 1; t < docxTables.Count; t++)
            {
                Require(docxTables[t].Columns[0].Width == docxTables[0].Columns[0].Width,
                    "docx band " + t + " should repeat the identity columns");
                Require(HasExactCellText(docxTables[t], "姚蕾"),
                    "docx band " + t + " should carry the name column so rows stay identifiable");
            }

            Console.WriteLine("[real-attendance] sheets=" + workbook.Worksheets.Count + ", read=" + readMs
                + " ms, layout=" + layoutMs + " ms, bands=" + bands.Count + ", pdf=" + pdf.PageCount + " pages/"
                + new FileInfo(pdfPath).Length + " bytes/" + pdfMs + " ms, docx="
                + new FileInfo(docxPath).Length + " bytes/" + docxMs + " ms");
        }

        /// <summary>
        /// 财务报表族的功能测试：资产负债表（双栏镜像）、利润表、财务收支公布表（多级合并表头）、
        /// 成本预算核算表（差异率百分号、人民币符号）与月度分析看板（超宽列带），
        /// 覆盖负数括号、百分号缩放、货币符号字面量、日期中文格式与合并跨页等真实报表写法。
        /// Functional test for the financial-statement family: two-panel balance sheet, income
        /// statement, income-and-expense disclosure with multi-level merged headers, cost budget sheet
        /// (percent variance, yuan sign) and a wide monthly dashboard; covers parenthesised negatives,
        /// percent scaling, literal currency symbols, Chinese date formats and merge handling across bands.
        /// </summary>
        /// <param name="outputDirectory">输出目录。</param>
        private static void RunFinancialStatementTest(string outputDirectory)
        {
            Stopwatch watch = new Stopwatch();
            OfficeWorkbook workbook = BuildFinancialWorkbook();

            string xlsxPath = PrepareOutputPath(Path.Combine(outputDirectory, "financial-statements.xlsx"));
            watch.Restart();
            XlsxWriter.Write(workbook, xlsxPath);
            watch.Stop();
            long writeMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(xlsxPath, "xl/workbook.xml");

            watch.Restart();
            OfficeWorkbook loaded = XlsxReader.Load(xlsxPath);
            watch.Stop();
            long readMs = watch.ElapsedMilliseconds;
            Require(loaded.Worksheets.Count == 6, "financial sheet count mismatch: " + loaded.Worksheets.Count);
            Require(loaded.Worksheets[5].IsHidden, "financial hidden dictionary sheet lost");

            Worksheet balance = loaded.Worksheets[0];
            Require(balance.PrintSetup.IsLandscape, "balance sheet should stay landscape");
            Require(balance.HasPrintArea, "balance sheet print area lost");
            Require(balance.PrintAreaFirstRow == 1 && balance.PrintAreaLastRow == 9
                && balance.PrintAreaFirstColumn == 1 && balance.PrintAreaLastColumn == 8,
                "balance sheet print area mismatch: " + balance.PrintAreaFirstRow + "-" + balance.PrintAreaLastRow
                + "," + balance.PrintAreaFirstColumn + "-" + balance.PrintAreaLastColumn);
            Require(HasMerge(balance, 1, 1, 1, 8), "balance sheet merged title lost");
            Require(HasMerge(balance, 3, 1, 1, 4), "balance sheet merged asset header lost");
            Require(balance.GetCell(6, 3).Style.NumberFormat == "#,##0.00;(#,##0.00);0.00",
                "balance sheet money format lost: " + balance.GetCell(6, 3).Style.NumberFormat);
            Require(balance.GetCell(4, 1).Style.BottomBorder != null, "balance sheet header border lost");

            // 财务格式的实际渲染结果，逐项与 Excel 显示对齐。
            Require(balance.GetCell(6, 3).DisplayText() == "(810,000.00)",
                "parenthesised negative mismatch: " + balance.GetCell(6, 3).DisplayText());
            Require(balance.GetCell(5, 3).DisplayText() == "1,280,500.00",
                "thousands format mismatch: " + balance.GetCell(5, 3).DisplayText());
            Require(balance.GetCell(7, 3).DisplayText() == "0.00",
                "zero should use the third format section: " + balance.GetCell(7, 3).DisplayText());
            Require(balance.GetCell(2, 5).DisplayText() == "2026年9月30日",
                "chinese date format mismatch: " + balance.GetCell(2, 5).DisplayText());

            Worksheet income = loaded.Worksheets[1];
            Require(income.GetCell(5, 3).DisplayText() == "(156,800.00)",
                "income statement negative mismatch: " + income.GetCell(5, 3).DisplayText());
            Require(income.GetCell(14, 3).DisplayText() == "0.4200",
                "earnings per share precision mismatch: " + income.GetCell(14, 3).DisplayText());

            Worksheet disclosure = loaded.Worksheets[2];
            Require(HasMerge(disclosure, 3, 2, 1, 3), "disclosure revenue group header merge lost");
            Require(HasMerge(disclosure, 3, 5, 1, 3), "disclosure expense group header merge lost");
            Require(disclosure.GetCell(9, 5).DisplayText() == "(660,000.00)",
                "disclosure total negative mismatch: " + disclosure.GetCell(9, 5).DisplayText());

            Worksheet cost = loaded.Worksheets[3];
            Require(cost.GetCell(4, 5).DisplayText() == "12.34%",
                "percent should scale by 100: " + cost.GetCell(4, 5).DisplayText());
            Require(cost.GetCell(4, 2).DisplayText() == "¥860,000.00",
                "yuan literal mismatch: " + cost.GetCell(4, 2).DisplayText());
            Require(cost.GetCell(5, 4).DisplayText() == "¥(51,230.25)",
                "yuan with parenthesised negative mismatch: " + cost.GetCell(5, 4).DisplayText());
            Require(cost.ShowZeros == false, "cost sheet showZeros=false lost");
            Require(cost.CenterHorizontally, "cost sheet horizontal centering lost");

            Worksheet dashboard = loaded.Worksheets[4];
            Require(dashboard.LastColumn >= 26, "dashboard should stay wide: " + dashboard.LastColumn);

            watch.Restart();
            OfficeDocument layout = WorkbookToDocumentConverter.Convert(loaded);
            watch.Stop();
            long layoutMs = watch.ElapsedMilliseconds;
            Require(layout.Sections.Count == 5, "financial visible section count mismatch: " + layout.Sections.Count);

            List<Table> balanceTables = FindTables(layout.Sections[0].Blocks);
            Require(balanceTables.Count == 1,
                "eight-column balance sheet should fit one landscape page: " + balanceTables.Count);
            Require(balanceTables[0].Rows.Count == 9,
                "balance sheet rows mismatch: " + balanceTables[0].Rows.Count);
            TableCell titleCell = balanceTables[0].Rows[0].Cells[0];
            Require(titleCell.ColSpan == 8, "balance sheet title should span all eight columns: " + titleCell.ColSpan);
            Require(balanceTables[0].Rows[2].Cells[0].ColSpan == 4,
                "asset group header should span four columns: " + balanceTables[0].Rows[2].Cells[0].ColSpan);
            Require(HasExactCellText(balanceTables[0], "(810,000.00)"), "parenthesised negative lost in layout");

            List<Table> costTables = FindTables(layout.Sections[3].Blocks);
            Require(costTables.Count == 1, "cost sheet should not band: " + costTables.Count);
            for (int t = 0; t < costTables.Count; t++)
            {
                // 责任部门列只有被隐藏的零值，不应把打印区域撑到第 9 行之后。
                Require(costTables[t].Rows.Count == 8,
                    "cost sheet rows should stop at the last printed value: " + costTables[t].Rows.Count);
            }

            List<Table> dashboardTables = FindTables(layout.Sections[4].Blocks);
            Require(dashboardTables.Count >= 2, "wide dashboard should split into column bands: " + dashboardTables.Count);
            Require(dashboardTables[0].Columns.Count < dashboard.LastColumn,
                "first band should be page-width only: " + dashboardTables[0].Columns.Count);
            Require(dashboardTables[1].Columns[0].Width == dashboardTables[0].Columns[0].Width,
                "later bands must repeat the label column");
            Require(HasExactCellText(dashboardTables[1], "指标1"),
                "repeated label column should carry the indicator name");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "financial-statements.pdf"));
            watch.Restart();
            OfficeConverter.WorkbookToPdf(loaded, pdfPath);
            watch.Stop();
            long pdfMs = watch.ElapsedMilliseconds;
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            string compact = pdf.ExtractText().Replace(" ", string.Empty).Replace("\r", string.Empty)
                .Replace("\n", string.Empty).Replace("\t", string.Empty);
            RequireContains(compact, "资产负债表");
            RequireContains(compact, "利润表");
            RequireContains(compact, "财务收支公布表");
            RequireContains(compact, "成本预算核算表");
            RequireContains(compact, "(810,000.00)");
            RequireContains(compact, "12.34%");
            RequireContains(compact, "2026年9月30日");
            Require(compact.IndexOf("机密字典数据", StringComparison.Ordinal) < 0, "hidden sheet leaked into pdf");
            int blankPages = 0;
            for (int p = 0; p < pdf.PageCount; p++)
            {
                IReadOnlyList<PdfTextLine> lines = pdf.Pages[p].GetTextLines();
                bool filled = false;
                for (int l = 0; l < lines.Count && !filled; l++)
                {
                    filled = lines[l].Text != null && lines[l].Text.Trim().Length > 0;
                }
                if (!filled) blankPages++;
            }
            Require(blankPages == 0, "financial pdf has blank pages: " + blankPages + "/" + pdf.PageCount);

            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "financial-statements.docx"));
            watch.Restart();
            OfficeConverter.Convert(xlsxPath, docxPath);
            watch.Stop();
            long docxMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(docxPath, "word/document.xml");
            OfficeDocument docx = DocxReader.Load(docxPath);
            Require(docx.Sections.Count == 5, "financial docx section count mismatch: " + docx.Sections.Count);
            Require(docx.Sections[0].PageSetup.IsLandscape, "financial docx should keep landscape balance sheet");
            List<Table> docxBalance = FindTables(docx.Sections[0].Blocks);
            Require(docxBalance.Count >= 1, "financial docx balance sheet table missing");
            Require(docxBalance[0].Rows[0].Cells[0].ColSpan == 8,
                "docx merged title span lost: " + docxBalance[0].Rows[0].Cells[0].ColSpan);
            Require(HasExactCellText(docxBalance[0], "(810,000.00)"), "docx parenthesised negative lost");
            Require(HasExactCellText(docxBalance[0], "1,280,500.00"), "docx thousands value lost");

            Require(writeMs < 5000, "financial write too slow: " + writeMs);
            Require(readMs < 5000, "financial read too slow: " + readMs);
            Require(layoutMs < 5000, "financial layout too slow: " + layoutMs);
            Require(pdfMs < 15000, "financial pdf too slow: " + pdfMs);

            Console.WriteLine("[financial] sheets=" + loaded.Worksheets.Count + ", write=" + writeMs
                + " ms, read=" + readMs + " ms, layout=" + layoutMs + " ms, bands=" + dashboardTables.Count
                + ", pdf=" + pdf.PageCount + " pages/" + new FileInfo(pdfPath).Length + " bytes/" + pdfMs
                + " ms, docx=" + new FileInfo(docxPath).Length + " bytes/" + docxMs + " ms");
        }

        private static OfficeWorkbook BuildFinancialWorkbook()
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.Metadata.Title = "2026年第三季度财务报表";
            workbook.Metadata.Author = "Wisdom.Office";

            CellStyle base10 = new CellStyle();
            base10.FontFamily = "Microsoft YaHei";
            base10.FontSize = 10;
            CellStyle title = base10.Clone();
            title.FontSize = 16;
            title.Bold = true;
            title.HorizontalAlignment = HorizontalAlignment.Center;
            CellStyle meta = base10.Clone();
            meta.HorizontalAlignment = HorizontalAlignment.Right;
            CellStyle groupHeader = base10.Clone();
            groupHeader.Bold = true;
            groupHeader.HorizontalAlignment = HorizontalAlignment.Center;
            groupHeader.FillColor = new OfficeColor(0xEE, 0xF2, 0xF8);
            groupHeader.TopBorder = BorderLine.Thin();
            groupHeader.BottomBorder = BorderLine.Thin();
            groupHeader.LeftBorder = BorderLine.Thin();
            groupHeader.RightBorder = BorderLine.Thin();
            CellStyle header = base10.Clone();
            header.Bold = true;
            header.HorizontalAlignment = HorizontalAlignment.Center;
            header.FillColor = new OfficeColor(0xD8, 0xE2, 0xEF);
            header.TopBorder = BorderLine.Thin();
            header.BottomBorder = BorderLine.Thin();
            header.LeftBorder = BorderLine.Thin();
            header.RightBorder = BorderLine.Thin();
            CellStyle label = base10.Clone();
            label.LeftBorder = BorderLine.Thin();
            label.RightBorder = BorderLine.Thin();
            label.BottomBorder = BorderLine.Thin();
            CellStyle labelBold = label.Clone();
            labelBold.Bold = true;
            labelBold.BottomBorder = new BorderLine(BorderStyle.Double, 1.5, OfficeColor.Black);
            CellStyle index = base10.Clone();
            index.HorizontalAlignment = HorizontalAlignment.Center;
            index.LeftBorder = BorderLine.Thin();
            index.RightBorder = BorderLine.Thin();
            index.BottomBorder = BorderLine.Thin();
            CellStyle money = base10.Clone();
            money.NumberFormat = "#,##0.00;(#,##0.00);0.00";
            money.HorizontalAlignment = HorizontalAlignment.Right;
            money.LeftBorder = BorderLine.Thin();
            money.RightBorder = BorderLine.Thin();
            money.BottomBorder = BorderLine.Thin();
            CellStyle moneyBold = money.Clone();
            moneyBold.Bold = true;
            moneyBold.BottomBorder = new BorderLine(BorderStyle.Double, 1.5, OfficeColor.Black);
            CellStyle yuan = money.Clone();
            yuan.NumberFormat = "\"¥\"#,##0.00;\"¥\"(#,##0.00)";
            CellStyle percent = money.Clone();
            percent.NumberFormat = "0.00%";
            CellStyle rate4 = money.Clone();
            rate4.NumberFormat = "0.0000";
            CellStyle dateStyle = base10.Clone();
            dateStyle.NumberFormat = "yyyy年m月d日";
            dateStyle.HorizontalAlignment = HorizontalAlignment.Left;

            Worksheet balance = workbook.AddWorksheet("资产负债表");
            balance.PrintSetup = PageSetup.A4();
            balance.PrintSetup.IsLandscape = true;
            balance.PrintSetup.MarginLeft = 28;
            balance.PrintSetup.MarginRight = 28;
            balance.ShowGridLines = false;
            balance.CenterHorizontally = true;
            balance.SetColumnWidth(1, 22);
            balance.SetColumnWidth(2, 5);
            balance.SetColumnWidth(3, 14);
            balance.SetColumnWidth(4, 14);
            balance.SetColumnWidth(5, 24);
            balance.SetColumnWidth(6, 5);
            balance.SetColumnWidth(7, 14);
            balance.SetColumnWidth(8, 14);
            balance.SetRowHeight(1, 30);
            Put(balance, 1, 1, "资产负债表", title);
            balance.Merge("A1:H1");
            Put(balance, 2, 1, "编制单位：某某建设装饰有限公司", meta);
            balance.Merge("A2:D2");
            Put(balance, 2, 5, new DateTime(2026, 9, 30), dateStyle);
            Put(balance, 2, 8, "会资01表", meta);
            Put(balance, 3, 1, "资产", groupHeader);
            balance.Merge("A3:D3");
            Put(balance, 3, 5, "负债和所有者权益", groupHeader);
            balance.Merge("E3:H3");
            Put(balance, 4, 1, "资产", header);
            Put(balance, 4, 2, "行次", header);
            Put(balance, 4, 3, "期末余额", header);
            Put(balance, 4, 4, "年初余额", header);
            Put(balance, 4, 5, "负债和所有者权益", header);
            Put(balance, 4, 6, "行次", header);
            Put(balance, 4, 7, "期末余额", header);
            Put(balance, 4, 8, "年初余额", header);
            string[,] left =
            {
                { "货币资金", "1", "1,280,500.00" },
                { "应收账款", "2", "(810,000.00)" },
                { "存货", "3", "0" },
                { "固定资产原值", "4", "3,450,000.00" }
            };
            string[,] right =
            {
                { "短期借款", "31", "500,000.00" },
                { "应付账款", "32", "(810,000.00)" },
                { "应交税费", "33", "24,530.00" },
                { "实收资本", "34", "2,000,000.00" }
            };
            for (int r = 0; r < 4; r++)
            {
                int row = 5 + r;
                Put(balance, row, 1, left[r, 0], label);
                Put(balance, row, 2, r + 1, index);
                Put(balance, row, 3, r == 1 ? -810000d : (r == 2 ? 0d : new double[] { 1280500, 0, 0, 3450000 }[r]), money);
                Put(balance, row, 4, new double[] { 1280500, 960000, 275000, 3450000 }[r], money);
                Put(balance, row, 5, right[r, 0], label);
                Put(balance, row, 6, 31 + r, index);
                Put(balance, row, 7, r == 1 ? -810000d : new double[] { 500000, 0, 24530, 2000000 }[r], money);
                Put(balance, row, 8, new double[] { 500000, 620000, 18900, 2000000 }[r], money);
            }
            Put(balance, 9, 1, "资产总计", labelBold);
            Put(balance, 9, 2, 5, index);
            Put(balance, 9, 3, 3920500d - 810000d, moneyBold);
            Put(balance, 9, 4, 5690500d, moneyBold);
            Put(balance, 9, 5, "负债和所有者权益总计", labelBold);
            Put(balance, 9, 6, 35, index);
            Put(balance, 9, 7, 3920500d - 810000d, moneyBold);
            Put(balance, 9, 8, 5690500d, moneyBold);
            Put(balance, 10, 1, "制表：姚蕾", label);
            Put(balance, 10, 5, "审核：李震", label);
            balance.SetPrintArea("A1:H9");

            Worksheet income = workbook.AddWorksheet("利润表");
            income.PrintSetup = PageSetup.A4();
            income.SetColumnWidth(1, 30);
            income.SetColumnWidth(2, 6);
            income.SetColumnWidth(3, 16);
            income.SetColumnWidth(4, 16);
            Put(income, 1, 1, "利润表", title);
            income.Merge("A1:D1");
            Put(income, 2, 1, "2026年9月　单位：元", meta);
            income.Merge("A2:D2");
            Put(income, 3, 1, "项目", header);
            Put(income, 3, 2, "行次", header);
            Put(income, 3, 3, "本期金额", header);
            Put(income, 3, 4, "本年累计金额", header);
            string[] incomeItems = { "一、营业收入", "减：营业成本", "税金及附加", "销售费用", "管理费用", "二、营业利润", "营业外收入", "三、利润总额", "减：所得税费用", "四、净利润" };
            double[] monthly = { 2450000, 1620000, 15300, 88000, 156800, 570099.5, 12000, 582099.5, 145525, 436574.5 };
            for (int r = 0; r < incomeItems.Length; r++)
            {
                int row = 4 + r;
                Put(income, row, 1, incomeItems[r], r >= 5 ? labelBold : label);
                Put(income, row, 2, r + 1, index);
                Put(income, row, 3, r == 1 ? -156800d : monthly[r], r == 5 || r == 9 ? moneyBold : money);
                Put(income, row, 4, monthly[r] * 9, r == 5 || r == 9 ? moneyBold : money);
            }
            Put(income, 14, 1, "基本每股收益", label);
            Put(income, 14, 2, 11, index);
            Put(income, 14, 3, 0.42d, rate4);
            Put(income, 14, 4, 3.78d, rate4);

            Worksheet disclosure = workbook.AddWorksheet("财务收支公布表");
            disclosure.PrintSetup = PageSetup.A4();
            disclosure.SetColumnWidth(1, 14);
            for (int c = 2; c <= 7; c++) disclosure.SetColumnWidth(c, 13);
            Put(disclosure, 1, 1, "财务收支公布表", title);
            disclosure.Merge("A1:G1");
            Put(disclosure, 2, 1, "某某小区 2026年第三季度", meta);
            disclosure.Merge("A2:G2");
            Put(disclosure, 3, 1, "项目", header);
            Put(disclosure, 3, 2, "收入", groupHeader);
            disclosure.Merge("B3:D3");
            Put(disclosure, 3, 5, "支出", groupHeader);
            disclosure.Merge("E3:G3");
            Put(disclosure, 4, 1, "类别", header);
            string[] sub = { "本月", "本季", "累计", "本月", "本季", "累计" };
            for (int c = 0; c < 6; c++) Put(disclosure, 4, 2 + c, sub[c], header);
            string[] groups = { "物业费", "停车费", "广告收入", "维修支出" };
            for (int r = 0; r < groups.Length; r++)
            {
                int row = 5 + r;
                Put(disclosure, row, 1, groups[r], label);
                for (int c = 2; c <= 7; c++)
                {
                    Put(disclosure, row, c, c <= 4 ? 120000d + r * 1000 : -(48000d + r * 2500), money);
                }
            }
            Put(disclosure, 9, 1, "合计", labelBold);
            for (int c = 2; c <= 7; c++) Put(disclosure, 9, c, c <= 4 ? 540000d : -660000d, moneyBold);

            Worksheet cost = workbook.AddWorksheet("成本预算核算表");
            cost.PrintSetup = PageSetup.A4();
            cost.PrintSetup.IsLandscape = true;
            cost.ShowZeros = false;
            cost.CenterHorizontally = true;
            cost.SetColumnWidth(1, 20);
            for (int c = 2; c <= 5; c++) cost.SetColumnWidth(c, 16);
            cost.SetColumnWidth(6, 14);
            Put(cost, 1, 1, "成本预算核算表", title);
            cost.Merge("A1:F1");
            Put(cost, 2, 1, "项目名称：精装A座", meta);
            Put(cost, 3, 1, "成本项", header);
            Put(cost, 3, 2, "预算金额", header);
            Put(cost, 3, 3, "实际金额", header);
            Put(cost, 3, 4, "差异额", header);
            Put(cost, 3, 5, "差异率", header);
            Put(cost, 3, 6, "责任部门", header);
            string[] costItems = { "材料费", "人工费", "机械租赁" };
            double[] budget = { 860000, 912500.5, 145000 };
            for (int r = 0; r < costItems.Length; r++)
            {
                int row = 4 + r;
                Put(cost, row, 1, costItems[r], label);
                Put(cost, row, 2, budget[r], yuan);
                Put(cost, row, 3, budget[r] * 1.05, yuan);
                Put(cost, row, 4, r == 1 ? -51230.25 : budget[r] * 0.05, yuan);
                Put(cost, row, 5, 0.1234, percent);
                Put(cost, row, 6, 0, money);
            }
            Put(cost, 8, 1, "说明：差异率＝差异额／预算金额", label);
            Put(cost, 8, 5, 0.0321, percent);

            Worksheet dashboard = workbook.AddWorksheet("月度分析看板");
            dashboard.PrintSetup = PageSetup.A4();
            dashboard.PrintSetup.IsLandscape = true;
            dashboard.FrozenRows = 2;
            dashboard.SetColumnWidth(1, 18);
            dashboard.SetColumnWidth(2, 10);
            for (int c = 3; c <= 28; c++) dashboard.SetColumnWidth(c, 9);
            Put(dashboard, 1, 1, "月度经营分析看板", title);
            dashboard.Merge("A1:AB1");
            Put(dashboard, 2, 1, "指标", header);
            Put(dashboard, 2, 2, "单位", header);
            for (int c = 3; c <= 28; c++) Put(dashboard, 2, c, c - 2 + "月", header);
            for (int r = 0; r < 8; r++)
            {
                int row = 3 + r;
                Put(dashboard, row, 1, "指标" + (r + 1), label);
                Put(dashboard, row, 2, "元", index);
                for (int c = 3; c <= 28; c++) Put(dashboard, row, c, r * 1000 + c * 13.5, money);
            }
            Put(dashboard, 11, 1, "数据来源：ERP 结转", label);

            Worksheet dictionary = workbook.AddWorksheet("数据字典");
            dictionary.IsHidden = true;
            Put(dictionary, 1, 1, "机密字典数据", label);
            Put(dictionary, 2, 1, "科目编码表", label);

            return workbook;
        }

        /// <summary>覆盖 20+ 图表类型与组合图：逐类构建、写出、读回，断言类型与区域引用不丢失，
        /// 组合图产出多个图表组且次坐标轴系列带覆盖标记。
        /// Covers the 20+ chart catalog and combo charts: build each type, write, reload and assert the type
        /// and range references survive; a combo yields multiple groups and its secondary series carries the override flag.</summary>
        /// <param name="outputDirectory">输出目录。Output directory.</param>
        private static void RunChartCatalogTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("图");
            CellStyle plain = new CellStyle();
            string[] cats = { "Q1", "Q2", "Q3", "Q4" };
            for (int i = 0; i < 4; i++) Put(sheet, 1, 2 + i, cats[i], plain);
            for (int i = 0; i < 4; i++) Put(sheet, 2, 2 + i, (i + 1) * 10.0, plain);
            for (int i = 0; i < 4; i++) Put(sheet, 3, 2 + i, (i + 1) * 4.0, plain);
            Put(sheet, 1, 1, "类目", plain);
            Put(sheet, 2, 1, "系列A", plain);
            Put(sheet, 3, 1, "系列B", plain);

            // 九种"复用类目 + 数值轴"的新类型 + 散点/气泡/雷达 + 一张柱线组合次轴图。
            // Nine axis-reuse types plus scatter/bubble/radar and one column-line combo on a secondary axis.
            ExcelChartType[] axisTypes =
            {
                ExcelChartType.StackedColumn, ExcelChartType.PercentStackedColumn,
                ExcelChartType.StackedBar, ExcelChartType.PercentStackedBar,
                ExcelChartType.LineWithMarkers, ExcelChartType.StackedLine, ExcelChartType.PercentStackedLine,
                ExcelChartType.StackedArea, ExcelChartType.PercentStackedArea
            };
            int col = 6;
            foreach (ExcelChartType type in axisTypes)
            {
                ExcelChart chart = new ExcelChart();
                chart.Type = type;
                chart.Column = col;
                col += 3;
                chart.AddSeries("图!$A$2", "图!$B$1:$E$1", "图!$B$2:$E$2");
                sheet.Charts.Add(chart);
            }

            ExcelChart scatter = new ExcelChart();
            scatter.Type = ExcelChartType.ScatterMarkers;
            scatter.Column = col; col += 3;
            scatter.AddScatterSeries("图!$A$2", "图!$B$1:$E$1", "图!$B$2:$E$2", null);
            sheet.Charts.Add(scatter);

            ExcelChart bubble = new ExcelChart();
            bubble.Type = ExcelChartType.Bubble;
            bubble.Column = col; col += 3;
            bubble.AddScatterSeries("图!$A$2", "图!$B$1:$E$1", "图!$B$2:$E$2", "图!$B$3:$E$3");
            sheet.Charts.Add(bubble);

            ExcelChart radar = new ExcelChart();
            radar.Type = ExcelChartType.RadarFilled;
            radar.Column = col; col += 3;
            radar.AddSeries("图!$A$2", "图!$B$1:$E$1", "图!$B$2:$E$2");
            sheet.Charts.Add(radar);

            ExcelChart combo = new ExcelChart();
            combo.Type = ExcelChartType.Column;
            combo.Title = "收入与增长率组合";
            combo.Column = col;
            combo.AddSeries("图!$A$2", "图!$B$1:$E$1", "图!$B$2:$E$2");
            combo.AddComboSeries("图!$A$3", "图!$B$1:$E$1", "图!$B$3:$E$3", ExcelChartType.Line, true);
            sheet.Charts.Add(combo);

            int expected = axisTypes.Length + 4;
            Require(sheet.Charts.Count == expected, "chart catalog count mismatch: " + sheet.Charts.Count);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "chart-catalog.xlsx"));
            XlsxWriter.Write(workbook, path);
            ValidateOoxmlPackage(path, "xl/workbook.xml");

            // 断言新类型的关键 OOXML 元素确实写出。
            // Assert the key OOXML elements for the new types were actually written.
            string[] chartXml = new string[expected];
            for (int i = 0; i < expected; i++) chartXml[i] = ReadPackageEntryText(path, "xl/charts/chart" + (i + 1) + ".xml");
            RequireContains(FirstContaining(chartXml, "percentStacked"), "<c:barChart>");
            RequireContains(FirstContaining(chartXml, "percentStacked"), "<c:barDir val=\"col\"");
            RequireContains(FirstContaining(chartXml, "scatterChart"), "<c:scatterStyle val=\"marker\"");
            RequireContains(FirstContaining(chartXml, "scatterChart"), "<c:xVal>");
            RequireContains(FirstContaining(chartXml, "bubbleChart"), "<c:bubbleSize>");
            RequireContains(FirstContaining(chartXml, "radarChart"), "<c:radarStyle val=\"filled\"");
            string comboXml = FirstContaining(chartXml, "333333333");
            Require(comboXml.Contains("<c:barChart>") && comboXml.Contains("<c:lineChart>"),
                "combo chart must contain both a bar and a line group");
            Require(comboXml.Contains("<c:valAx") && CountOccurrences(comboXml, "<c:valAx") == 2,
                "combo chart must carry a primary and a secondary value axis");

            OfficeWorkbook loaded = XlsxReader.Load(path);
            Worksheet back = loaded.Worksheets[0];
            Require(back.Charts.Count == expected, "reloaded chart count mismatch: " + back.Charts.Count);

            for (int i = 0; i < axisTypes.Length; i++)
            {
                Require(back.Charts[i].Type == axisTypes[i],
                    "axis type " + i + " round-trip mismatch: " + back.Charts[i].Type);
                Require(back.Charts[i].Series[0].ValueRange == "图!$B$2:$E$2",
                    "axis type " + i + " value range lost");
            }
            Require(back.Charts[axisTypes.Length].Type == ExcelChartType.ScatterMarkers, "scatter type lost");
            Require(back.Charts[axisTypes.Length].Series[0].CategoryRange == "图!$B$1:$E$1", "scatter x range lost");
            Require(back.Charts[axisTypes.Length + 1].Type == ExcelChartType.Bubble, "bubble type lost");
            Require(back.Charts[axisTypes.Length + 1].Series[0].SizeRange == "图!$B$3:$E$3", "bubble size range lost");
            Require(back.Charts[axisTypes.Length + 2].Type == ExcelChartType.RadarFilled, "radar type lost");

            ExcelChart reloadedCombo = back.Charts[expected - 1];
            Require(reloadedCombo.Type == ExcelChartType.Column, "combo base type mismatch: " + reloadedCombo.Type);
            Require(reloadedCombo.Series.Count == 2, "combo series count mismatch: " + reloadedCombo.Series.Count);
            Require(!reloadedCombo.Series[0].SecondaryAxis && reloadedCombo.Series[0].ChartType == null,
                "combo primary series must not override");
            Require(reloadedCombo.Series[1].SecondaryAxis, "combo secondary flag lost");
            Require(reloadedCombo.Series[1].ChartType == ExcelChartType.Line, "combo line override lost");
            Require(reloadedCombo.Title == "收入与增长率组合", "combo title lost");

            // 再次写出应保持稳定，读回的图表能原样再序列化。
            // Re-writing must stay stable; the reloaded charts must serialize again in place.
            string resavePath = PrepareOutputPath(Path.Combine(outputDirectory, "chart-catalog-resaved.xlsx"));
            XlsxWriter.Write(loaded, resavePath);
            ValidateOoxmlPackage(resavePath, "xl/workbook.xml");

            Console.WriteLine("[chart-catalog] types=" + expected + ", combo groups=2, reloaded=" + back.Charts.Count);
        }

        /// <summary>返回第一个包含指定片段的图表 XML；都不包含时返回空串。 Returns the first chart XML containing a fragment, or an empty string when none do.</summary>
        /// <param name="parts">图表 XML 数组。 The chart XML array.</param>
        /// <param name="needle">要查找的片段。 The fragment to find.</param>
        /// <returns>匹配的 XML 文本。 The matched XML text.</returns>
        private static string FirstContaining(string[] parts, string needle)
        {
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i] != null && parts[i].Contains(needle)) return parts[i];
            }
            return string.Empty;
        }

        /// <summary>企业级批量与异步转换门面端到端测试：有界并行的整批转换、目录扫描转换、单项异步封装、
        /// 进度回调与单项失败隔离（坏文件不拖垮整批）。
        /// End-to-end test for the enterprise batch and async façade: bounded-parallel batch runs, directory scans,
        /// single-file async wrappers, progress reporting and per-item failure isolation (a bad file must not abort the batch).</summary>
        /// <param name="outputDirectory">输出目录。Output directory.</param>
        private static void RunBatchConversionTest(string outputDirectory)
        {
            string root = PrepareOutputPath(Path.Combine(outputDirectory, "batch-src"));
            if (Directory.Exists(root)) Directory.Delete(root, true);
            Directory.CreateDirectory(root);

            const int count = 6;
            string[] sources = new string[count];
            CellStyle plain = new CellStyle();
            for (int i = 0; i < count; i++)
            {
                OfficeWorkbook workbook = new OfficeWorkbook();
                Worksheet sheet = workbook.AddWorksheet("S" + i);
                Put(sheet, 1, 1, "批次" + i, plain);
                Put(sheet, 1, 2, (i + 1) * 100.0, plain);
                sources[i] = Path.Combine(root, "book" + i + ".xlsx");
                XlsxWriter.Write(workbook, sources[i]);
            }

            // 故意写一个坏 xlsx，验证单项失败被捕获而不抛出。
            // Deliberately write a bogus xlsx to prove a single failure is captured rather than thrown.
            string broken = Path.Combine(root, "broken.xlsx");
            File.WriteAllText(broken, "not a real xlsx package");

            // 目录扫描先跑：此刻 root 仅含 6 个有效 xlsx 与 1 个坏 xlsx，共 7 个受支持输入。
            // Run the directory scan first: at this point root holds exactly 6 valid xlsx plus 1 bogus one = 7 supported inputs.
            string targetDir = PrepareOutputPath(Path.Combine(outputDirectory, "batch-out"));
            if (Directory.Exists(targetDir)) Directory.Delete(targetDir, true);
            BatchConversionReport dir = OfficeBatchConverter.ConvertDirectory(root, targetDir, ".docx", 4);
            Require(dir.Total == count + 1, "directory total mismatch: " + dir.Total);
            Require(dir.Succeeded == count, "directory succeeded mismatch: " + dir.Succeeded);
            Require(dir.Failed == 1, "directory failed mismatch: " + dir.Failed);

            string pdfDir = PrepareOutputPath(Path.Combine(outputDirectory, "batch-pdf"));
            if (Directory.Exists(pdfDir)) Directory.Delete(pdfDir, true);
            Directory.CreateDirectory(pdfDir);

            List<BatchConversionItem> items = new List<BatchConversionItem>();
            List<BatchConversionItem> valid = new List<BatchConversionItem>();
            for (int i = 0; i < count; i++)
            {
                BatchConversionItem item = new BatchConversionItem(sources[i], Path.Combine(pdfDir, "book" + i + ".pdf"), null);
                items.Add(item);
                valid.Add(item);
            }
            items.Add(new BatchConversionItem(broken, Path.Combine(pdfDir, "broken.pdf"), null));

            int reported = 0;
            Progress<BatchConversionOutcome> progress = new Progress<BatchConversionOutcome>(
                delegate(BatchConversionOutcome outcome) { Interlocked.Increment(ref reported); });
            BatchConversionReport report = OfficeBatchConverter.Convert(items, 4, progress);
            Require(report.Total == count + 1, "batch total mismatch: " + report.Total);
            Require(report.Succeeded == count, "batch succeeded mismatch: " + report.Succeeded);
            Require(report.Failed == 1, "batch failed mismatch: " + report.Failed);
            Require(report.ElapsedMilliseconds >= 0, "elapsed must be non-negative");

            bool failureCarriesError = false;
            for (int i = 0; i < report.Outcomes.Count; i++)
            {
                if (!report.Outcomes[i].Success) failureCarriesError = report.Outcomes[i].Error != null;
            }
            Require(failureCarriesError, "a failed job must carry an exception");
            for (int i = 0; i < count; i++) Require(File.Exists(items[i].TargetPath), "missing pdf " + i);
            for (int i = 0; i < count; i++)
            {
                PdfDocument pdf = PdfDocument.Load(items[i].TargetPath);
                Require(pdf.PageCount >= 1, "pdf page count " + i);
            }

            // 单项异步封装：await 风格完成且产物是合法 OOXML。
            // Single-file async wrapper: completes await-style and the product is valid OOXML.
            string asyncDocx = Path.Combine(pdfDir, "book0-async.docx");
            OfficeBatchConverter.ConvertFileAsync(sources[0], asyncDocx, null, CancellationToken.None)
                .GetAwaiter().GetResult();
            Require(File.Exists(asyncDocx), "async single-file output missing");
            ValidateOoxmlPackage(asyncDocx, "word/document.xml");

            // 异步整批：返回逐项报告。
            // Async batch: yields a per-item report.
            BatchConversionReport asyncReport = OfficeBatchConverter.ConvertAsync(valid, 3, null, CancellationToken.None)
                .GetAwaiter().GetResult();
            Require(asyncReport.Total == count, "async batch total mismatch: " + asyncReport.Total);
            Require(asyncReport.Succeeded == count, "async batch succeeded mismatch: " + asyncReport.Succeeded);

            // 取消令牌：用一个已取消的令牌，应立即抛出且不影响成功用例。
            // Cancellation token: an already-cancelled token must throw without disturbing the success cases.
            CancellationTokenSource cts = new CancellationTokenSource();
            cts.Cancel();
            bool cancelled = false;
            try
            {
                OfficeBatchConverter.Convert(valid, 2, null, cts.Token);
            }
            catch (OperationCanceledException)
            {
                cancelled = true;
            }
            Require(cancelled, "a cancelled batch must throw OperationCanceledException");
            cts.Dispose();

            Console.WriteLine("[batch] jobs=" + report.Total + " ok=" + report.Succeeded + " fail=" + report.Failed
                + " dir=" + dir.Total + " async-ok=" + asyncReport.Succeeded + " progress=" + reported);
        }

        /// <summary>看板原生特性端到端测试：原生图表（柱形/环形/折线）、条件格式（色阶/图标集/数据条/高亮）、
        /// 单元格图片和列宽不足时的 "########" 溢出提示。校验 OOXML 部件结构、读回链路与 Word/PDF 转换。
        /// End-to-end test for native dashboard features: charts (column, doughnut, line), conditional
        /// formatting (color scale, icon set, data bar, highlight), anchored cell images and the
        /// "########" overflow hint. Verifies the OOXML parts, the read-back path and the Word/PDF conversions.
        /// </summary>
        /// <param name="outputDirectory">输出目录。Output directory.</param>
        private static void RunDashboardChartTest(string outputDirectory)
        {
            Stopwatch watch = new Stopwatch();
            OfficeWorkbook workbook = BuildDashboardWorkbook();

            string xlsxPath = PrepareOutputPath(Path.Combine(outputDirectory, "dashboard-native.xlsx"));
            watch.Restart();
            XlsxWriter.Write(workbook, xlsxPath);
            watch.Stop();
            long writeMs = watch.ElapsedMilliseconds;
            ValidateOoxmlPackage(xlsxPath, "xl/workbook.xml");

            string sheetXml = ReadPackageEntryText(xlsxPath, "xl/worksheets/sheet1.xml");
            RequireContains(sheetXml, "sqref=\"B6:M6\"");
            RequireContains(sheetXml, "sqref=\"B8:M8\"");
            RequireContains(sheetXml, "sqref=\"B9:M9\"");
            RequireContains(sheetXml, "<cfRule type=\"colorScale\" priority=\"1\">");
            RequireContains(sheetXml, "<cfRule type=\"iconSet\" priority=\"2\">");
            RequireContains(sheetXml, "<cfRule type=\"dataBar\" priority=\"3\">");
            RequireContains(sheetXml, "<cfRule type=\"cellIs\" operator=\"lessThan\" dxfId=\"0\" priority=\"4\">");
            RequireContains(sheetXml, "<drawing r:id=\"rId1\"");
            Require(sheetXml.IndexOf("conditionalFormatting", StringComparison.Ordinal)
                < sheetXml.IndexOf("<drawing", StringComparison.Ordinal),
                "conditionalFormatting must precede the drawing element");

            string styles = ReadPackageEntryText(xlsxPath, "xl/styles.xml");
            RequireContains(styles, "<dxfs count=\"1\">");
            RequireContains(styles, "FFC00000");

            string columnChart = ReadPackageEntryText(xlsxPath, "xl/charts/chart1.xml");
            RequireContains(columnChart, "<c:barChart>");
            RequireContains(columnChart, "<c:barDir val=\"col\"");
            RequireContains(columnChart, "<c:grouping val=\"clustered\"");
            RequireContains(columnChart, "看板!$B$5:$M$5");
            RequireContains(columnChart, "看板!$B$6:$M$6");
            RequireContains(columnChart, "<c:strRef>");
            RequireContains(columnChart, "<c:numCache>");
            RequireContains(columnChart, "<c:ptCount val=\"12\"");
            RequireContains(columnChart, "<c:title>");
            RequireContains(columnChart, "月度收入与成本对比");
            RequireContains(columnChart, "<c:legendPos val=\"b\"");
            Require(CountOccurrences(columnChart, "<c:catAx") == 1,
                "each axis must appear once, inside plotArea only");
            Require(CountOccurrences(columnChart, "<c:valAx") == 1,
                "each axis must appear once, inside plotArea only");
            RequireContains(columnChart, "<c:tx><c:strRef><c:f>看板!$A$6</c:f>");

            string doughnut = ReadPackageEntryText(xlsxPath, "xl/charts/chart2.xml");
            RequireContains(doughnut, "<c:doughnutChart>");
            RequireContains(doughnut, "<c:holeSize val=\"60\"");
            RequireContains(doughnut, "<c:dPt>");
            RequireContains(doughnut, "FF70A0");

            string line = ReadPackageEntryText(xlsxPath, "xl/charts/chart3.xml");
            RequireContains(line, "<c:lineChart>");
            RequireContains(line, "<c:smooth val=\"1\"");
            RequireContains(line, "看板!$B$9:$M$9");

            string drawing = ReadPackageEntryText(xlsxPath, "xl/drawings/drawing1.xml");
            RequireContains(drawing, "<xdr:wsDr");
            int anchors = drawing.Split(new[] { "<xdr:oneCellAnchor>" }, StringSplitOptions.None).Length - 1;
            Require(anchors == 4, "one sheet should carry three chart frames plus one picture: " + anchors);
            RequireContains(drawing, "<xdr:graphicFrame");
            RequireContains(drawing, "<xdr:pic>");
            RequireContains(drawing, "r:embed=\"rId4\"");

            string drawingRels = ReadPackageEntryText(xlsxPath, "xl/drawings/_rels/drawing1.xml.rels");
            RequireContains(drawingRels, "../charts/chart1.xml");
            RequireContains(drawingRels, "../charts/chart3.xml");
            RequireContains(drawingRels, "../media/image1.png");

            string sheetRels = ReadPackageEntryText(xlsxPath, "xl/worksheets/_rels/sheet1.xml.rels");
            RequireContains(sheetRels, "../drawings/drawing1.xml");

            string contentTypes = ReadPackageEntryText(xlsxPath, "[Content_Types].xml");
            RequireContains(contentTypes, "PartName=\"/xl/charts/chart1.xml\"");
            RequireContains(contentTypes, "ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"");
            RequireContains(contentTypes, "PartName=\"/xl/drawings/drawing1.xml\"");
            RequireContains(contentTypes, "ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"");
            RequireContains(contentTypes, "Extension=\"png\"");

            long mediaBytes = PackageEntryLength(xlsxPath, "xl/media/image1.png");
            Require(mediaBytes > 0, "anchored image part is empty");

            watch.Restart();
            OfficeWorkbook loaded = XlsxReader.Load(xlsxPath);
            watch.Stop();
            long readMs = watch.ElapsedMilliseconds;
            Require(loaded.Worksheets.Count == 1, "dashboard sheet count mismatch: " + loaded.Worksheets.Count);
            Worksheet back = loaded.Worksheets[0];
            Require(back.GetCell(6, 2).DisplayText() == "1,180.00",
                "month value lost: " + back.GetCell(6, 2).DisplayText());
            Require(back.GetColumnWidth(14) == 4, "narrow total column lost: " + back.GetColumnWidth(14));
            Require(HasMerge(back, 1, 1, 1, 14), "dashboard merged title lost");

            Require(back.Charts.Count == 3, "reloaded chart count mismatch: " + back.Charts.Count);
            Require(back.Images.Count == 1, "reloaded image count mismatch: " + back.Images.Count);
            Require(back.ConditionalFormats.Count == 3, "reloaded format count mismatch: " + back.ConditionalFormats.Count);
            Require(back.Charts[0].Type == ExcelChartType.Column,
                "first chart type mismatch: " + back.Charts[0].Type);
            Require(back.Charts[1].Type == ExcelChartType.Doughnut,
                "second chart type mismatch: " + back.Charts[1].Type);
            Require(back.Charts[2].Type == ExcelChartType.Line,
                "third chart type mismatch: " + back.Charts[2].Type);
            Require(back.Charts[0].Series[0].ValueRange == "看板!$B$6:$M$6",
                "chart value range lost: " + back.Charts[0].Series[0].ValueRange);
            Require(back.Charts[0].Title == "月度收入与成本对比", "chart title lost: " + back.Charts[0].Title);
            Require(Math.Abs(back.Charts[0].WidthPoints - 420) < 1,
                "chart width lost: " + back.Charts[0].WidthPoints);
            Require(back.Charts[1].Series[0].Categories.Count == 4,
                "doughnut literal categories lost: " + back.Charts[1].Series[0].Categories.Count);
            Require(back.Charts[1].Series[0].Values.Count == 4 && back.Charts[1].Series[0].Values[0] == 620,
                "doughnut literal values lost");
            Require(back.Charts[1].Series[0].PointColors.Count == 4,
                "doughnut point colors lost: " + back.Charts[1].Series[0].PointColors.Count);
            Require(back.Charts[2].Series[0].Smooth, "line smooth flag lost");
            Require(back.Images[0].Image.PixelWidth == 1, "image bytes lost: " + back.Images[0].Image.PixelWidth);
            Require(back.Images[0].Row == 2 && back.Images[0].Column == 14,
                "image anchor lost: " + back.Images[0].Row + "," + back.Images[0].Column);
            Require(back.ConditionalFormats[0].Rules[0].Kind == ExcelConditionalKind.ColorScale
                && back.ConditionalFormats[0].Rules[0].Colors.Count == 3, "color scale lost");
            Require(back.ConditionalFormats[2].Rules[0].Kind == ExcelConditionalKind.DataBar, "data bar lost");
            Require(back.ConditionalFormats[2].Rules[1].Kind == ExcelConditionalKind.CellIs
                && back.ConditionalFormats[2].Rules[1].Format != null
                && back.ConditionalFormats[2].Rules[1].Format.Bold
                && back.ConditionalFormats[2].Rules[1].Format.FontColor == new OfficeColor(0xC0, 0x00, 0x00),
                "cellIs dxf style lost");

            string resavePath = PrepareOutputPath(Path.Combine(outputDirectory, "dashboard-native-resaved.xlsx"));
            XlsxWriter.Write(loaded, resavePath);
            ValidateOoxmlPackage(resavePath, "xl/workbook.xml");
            RequireContains(ReadPackageEntryText(resavePath, "xl/charts/chart1.xml"), "看板!$B$6:$M$6");
            RequireContains(ReadPackageEntryText(resavePath, "xl/worksheets/sheet1.xml"), "<cfRule type=\"dataBar\"");
            RequireContains(ReadPackageEntryText(resavePath, "xl/drawings/drawing1.xml"), "r:embed=\"rId4\"");
            Require(PackageEntryLength(resavePath, "xl/media/image1.png") == mediaBytes,
                "image bytes changed on resave");

            OfficeDocument layout = WorkbookToDocumentConverter.Convert(loaded);
            List<Table> tables = FindTables(layout.Sections[0].Blocks);
            Require(tables.Count >= 1, "dashboard table missing: " + tables.Count);
            bool overflow = false;
            for (int t = 0; t < tables.Count; t++)
            {
                overflow = overflow || HasExactCellText(tables[t], "####");
            }
            Require(overflow, "the four-character total column should render as #### like Excel");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "dashboard-native.pdf"));
            watch.Restart();
            OfficeConverter.WorkbookToPdf(workbook, pdfPath);
            watch.Stop();
            long pdfMs = watch.ElapsedMilliseconds;
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            string compact = pdf.ExtractText().Replace(" ", string.Empty).Replace("\r", string.Empty)
                .Replace("\n", string.Empty).Replace("\t", string.Empty);
            RequireContains(compact, "月度财务收入成本费用分析报表");
            RequireContains(compact, "####");

            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "dashboard-native.docx"));
            OfficeConverter.Convert(xlsxPath, docxPath);
            ValidateOoxmlPackage(docxPath, "word/document.xml");
            RequireContains(ReadPackageEntryText(docxPath, "word/document.xml"), "####");

            Require(writeMs < 5000, "dashboard write too slow: " + writeMs);
            Require(readMs < 5000, "dashboard read too slow: " + readMs);
            Require(pdfMs < 15000, "dashboard pdf too slow: " + pdfMs);

            Console.WriteLine("[dashboard] anchors=" + anchors + ", media=" + mediaBytes
                + " bytes, rules=" + CountOccurrences(sheetXml, "<cfRule ")
                + ", write=" + writeMs + " ms, read=" + readMs + " ms, pdf=" + pdf.PageCount
                + " pages/" + pdfMs + " ms, reloaded charts=" + back.Charts.Count
                + ", reloaded formats=" + back.ConditionalFormats.Count);
        }

        /// <summary>构造带 KPI 卡片、12 个月明细、三张原生图表、四类条件格式与单元格图标的看板工作簿。
        /// Builds a dashboard workbook with KPI cards, a twelve-month table, three native charts,
        /// four conditional-format rules and an anchored cell image.</summary>
        /// <returns>看板工作簿。The dashboard workbook.</returns>
        private static OfficeWorkbook BuildDashboardWorkbook()
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.Metadata.Title = "月度财务收入成本费用分析报表";
            workbook.Metadata.Author = "Wisdom.Office";

            CellStyle base11 = new CellStyle();
            base11.FontFamily = "Microsoft YaHei";
            base11.FontSize = 11;
            CellStyle title = base11.Clone();
            title.FontSize = 18;
            title.Bold = true;
            title.HorizontalAlignment = HorizontalAlignment.Center;
            CellStyle cardLabel = base11.Clone();
            cardLabel.FontSize = 9;
            cardLabel.FontColor = OfficeColor.Gray;
            cardLabel.HorizontalAlignment = HorizontalAlignment.Center;
            cardLabel.FillColor = new OfficeColor(0xF3, 0xF8, 0xF3);
            cardLabel.TopBorder = BorderLine.Thin();
            cardLabel.BottomBorder = BorderLine.Thin();
            cardLabel.LeftBorder = BorderLine.Thin();
            cardLabel.RightBorder = BorderLine.Thin();
            CellStyle cardValue = cardLabel.Clone();
            cardValue.FontSize = 16;
            cardValue.Bold = true;
            cardValue.FontColor = new OfficeColor(0x1F, 0x70, 0x40);
            cardValue.NumberFormat = "#,##0.00";
            CellStyle header = base11.Clone();
            header.Bold = true;
            header.FontColor = OfficeColor.White;
            header.FillColor = new OfficeColor(0x00, 0xB0, 0x50);
            header.HorizontalAlignment = HorizontalAlignment.Center;
            header.TopBorder = BorderLine.Thin();
            header.BottomBorder = BorderLine.Thin();
            header.LeftBorder = BorderLine.Thin();
            header.RightBorder = BorderLine.Thin();
            CellStyle label = base11.Clone();
            label.LeftBorder = BorderLine.Thin();
            label.RightBorder = BorderLine.Thin();
            label.BottomBorder = BorderLine.Thin();
            CellStyle money = label.Clone();
            money.NumberFormat = "#,##0.00";
            money.HorizontalAlignment = HorizontalAlignment.Right;
            CellStyle percent = money.Clone();
            percent.NumberFormat = "0.00%";
            CellStyle total = money.Clone();
            total.Bold = true;
            total.FillColor = new OfficeColor(0xEE, 0xF2, 0xF8);

            Worksheet sheet = workbook.AddWorksheet("看板");
            sheet.PrintSetup = PageSetup.A4();
            sheet.PrintSetup.IsLandscape = true;
            sheet.ShowGridLines = false;
            sheet.CenterHorizontally = true;
            sheet.SetColumnWidth(1, 18);
            for (int c = 2; c <= 13; c++) sheet.SetColumnWidth(c, 8);
            sheet.SetColumnWidth(14, 4);
            sheet.SetRowHeight(1, 30);
            sheet.SetRowHeight(3, 26);

            Put(sheet, 1, 1, "月度财务收入成本费用分析报表", title);
            sheet.Merge("A1:N1");

            string[] cards = { "本月收入", "本月成本", "本月费用", "本月利润" };
            double[] cardValues = { 9440.00, 5210.75, 1480.40, 2748.85 };
            for (int i = 0; i < cards.Length; i++)
            {
                int first = 1 + i * 3;
                Put(sheet, 2, first, cards[i], cardLabel);
                sheet.Merge(CellAddress.Format(2, first) + ":" + CellAddress.Format(2, first + 2));
                Put(sheet, 3, first, cardValues[i], cardValue);
                sheet.Merge(CellAddress.Format(3, first) + ":" + CellAddress.Format(3, first + 2));
            }
            Put(sheet, 2, 13, "利润率", cardLabel);
            Put(sheet, 3, 13, 0.2912, percent);

            Put(sheet, 5, 1, "项目", header);
            for (int c = 0; c < 12; c++) Put(sheet, 5, 2 + c, (c + 1) + "月", header);
            Put(sheet, 5, 14, "合计", header);

            string[] rows = { "收入", "成本", "费用", "利润" };
            double[] revenue = { 1180, 960, 1040, 870, 720, 655, 690, 745, 810, 890, 940, 990 };
            double[] rates = { 1, 0.56, 0.16, 0.28 };
            for (int r = 0; r < rows.Length; r++)
            {
                int row = 6 + r;
                Put(sheet, row, 1, rows[r], label);
                double sum = 0;
                for (int c = 0; c < 12; c++)
                {
                    double value = Math.Round(revenue[c] * rates[r], 2);
                    Put(sheet, row, 2 + c, r == 3 && c == 4 ? -620.45 : value, money);
                    sum += value;
                }
                Put(sheet, row, 14, Math.Round(sum, 2), total);
            }

            ExcelChart column = new ExcelChart();
            column.Type = ExcelChartType.Column;
            column.Title = "月度收入与成本对比";
            column.Row = 12;
            column.Column = 1;
            column.WidthPoints = 420;
            column.HeightPoints = 240;
            column.ValueAxisFormat = "#,##0";
            ExcelChartSeries income = column.AddSeries("看板!$A$6", "看板!$B$5:$M$5", "看板!$B$6:$M$6");
            income.FillColor = new OfficeColor(0x00, 0xB0, 0x50);
            income.ShowDataLabels = true;
            income.DataLabelFormat = "#,##0";
            ExcelChartSeries cost = column.AddSeries("看板!$A$7", "看板!$B$5:$M$5", "看板!$B$7:$M$7");
            cost.FillColor = new OfficeColor(0xF0, 0x8C, 0x00);
            sheet.Charts.Add(column);

            ExcelChart doughnut = new ExcelChart();
            doughnut.Type = ExcelChartType.Doughnut;
            doughnut.Title = "费用构成";
            doughnut.Row = 12;
            doughnut.Column = 8;
            doughnut.WidthPoints = 300;
            doughnut.HeightPoints = 240;
            doughnut.HoleSize = 60;
            ExcelChartSeries parts = doughnut.AddSeries("费用构成",
                new List<string> { "人工", "材料", "机械", "管理" },
                new List<double> { 620, 430, 240, 190 });
            parts.PointColors.Add(new OfficeColor(0x00, 0xB0, 0x50));
            parts.PointColors.Add(new OfficeColor(0xFF, 0x70, 0xA0));
            parts.PointColors.Add(new OfficeColor(0x40, 0x80, 0xE0));
            parts.PointColors.Add(new OfficeColor(0xF0, 0xC0, 0x40));
            sheet.Charts.Add(doughnut);

            ExcelChart trend = new ExcelChart();
            trend.Type = ExcelChartType.Line;
            trend.Title = "利润趋势";
            trend.Row = 26;
            trend.Column = 1;
            trend.WidthPoints = 420;
            trend.HeightPoints = 220;
            trend.ShowLegend = false;
            ExcelChartSeries profit = trend.AddSeries("看板!$A$9", "看板!$B$5:$M$5", "看板!$B$9:$M$9");
            profit.Smooth = true;
            profit.FillColor = new OfficeColor(0x40, 0x80, 0xE0);
            sheet.Charts.Add(trend);

            ExcelConditionalFormat scale = new ExcelConditionalFormat("B6:M6");
            scale.AddColorScale(new OfficeColor(0xF8, 0x69, 0x6A), new OfficeColor(0xFF, 0xEB, 0x84),
                new OfficeColor(0x63, 0xC3, 0x84));
            sheet.ConditionalFormats.Add(scale);

            ExcelConditionalFormat arrows = new ExcelConditionalFormat("B8:M8");
            arrows.AddIconSet("3Arrows", new List<double> { 33, 67 });
            sheet.ConditionalFormats.Add(arrows);

            ExcelConditionalFormat profitBar = new ExcelConditionalFormat("B9:M9");
            profitBar.AddDataBar(new OfficeColor(0x63, 0xC3, 0x84));
            CellStyle loss = new CellStyle();
            loss.Bold = true;
            loss.FontColor = new OfficeColor(0xC0, 0x00, 0x00);
            loss.FillColor = new OfficeColor(0xFD, 0xE7, 0xE7);
            profitBar.AddCellIs("lessThan", "0", loss);
            sheet.ConditionalFormats.Add(profitBar);

            ExcelImage icon = new ExcelImage(new OfficeImage(CreateTestPng(1, 1)));
            icon.Row = 2;
            icon.Column = 14;
            icon.WidthPoints = 24;
            icon.HeightPoints = 24;
            sheet.Images.Add(icon);

            return workbook;
        }

        /// <summary>读取部件字节长度；部件缺失时返回 0。Reads a part's byte length, returning 0 when the part is missing.</summary>
        /// <param name="path">包路径。The package path.</param>
        /// <param name="entryName">部件名。The entry name.</param>
        /// <returns>字节长度。The length in bytes.</returns>
        private static long PackageEntryLength(string path, string entryName)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = archive.GetEntry(entryName);
                Require(entry != null, "OOXML part missing: " + entryName);
                return entry.Length;
            }
        }

        private static void Put(Worksheet sheet, int row, int column, object value, CellStyle style)
        {
            Cell cell = sheet.GetCell(row, column);
            cell.Value = value;
            if (style != null) cell.Style = style;
        }

        /// <summary>
        /// 检验／审核报告类报表测试：页眉文件编号、页脚页码域、多级合并表头、跨页重复表头、签署栏。
        /// Functional test for inspection/audit reports: a header with the document number, footer page
        /// fields, merged multi-level table headers, a header row repeated across pages and a signature block.
        /// </summary>
        /// <param name="outputDirectory">输出目录。Output directory.</param>
        private static void RunInspectionReportTest(string outputDirectory)
        {
            Stopwatch timer = Stopwatch.StartNew();
            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "inspection-report.docx"));
            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "inspection-report.pdf"));

            OfficeDocument document = BuildInspectionDocument();
            long build = timer.ElapsedMilliseconds;

            DocxWriter.Write(document, docxPath);
            ValidateOoxmlPackage(docxPath, "word/document.xml");
            long write = timer.ElapsedMilliseconds;

            OfficeDocument parsed = DocxReader.Load(docxPath);
            Section restored = parsed.Sections[0];
            Require(restored.Header.IsEnabled, "inspection header should survive the docx round-trip");
            RequireContains(GetBlocksText(restored.Header.Blocks), "QF-QC-018");
            Require(restored.Footer.IsEnabled, "inspection footer should survive the docx round-trip");
            Paragraph footer = restored.Footer.Blocks[0] as Paragraph;
            Require(footer != null, "inspection footer paragraph lost in round-trip");
            Require(HasFieldInstruction(footer, "PAGE"), "footer PAGE field lost: " + footer.GetText());
            Require(HasFieldInstruction(footer, "NUMPAGES"), "footer NUMPAGES field lost: " + footer.GetText());

            List<Table> roundTrip = FindTables(restored.Blocks);
            Require(roundTrip.Count == 3, "inspection docx should round-trip three tables: " + roundTrip.Count);
            Table back = roundTrip[1];
            Require(back.Rows[0].Cells[0].RowSpan == 2, "vertical merge on 序号 lost: " + back.Rows[0].Cells[0].RowSpan);
            Require(back.Rows[0].Cells[4].ColSpan == 2, "horizontal merge on 实测结果 lost: " + back.Rows[0].Cells[4].ColSpan);
            Require(back.Rows[0].IsHeader, "inspection item table header row flag lost");
            Require(roundTrip[2].Rows.Count == 2 && roundTrip[2].Columns.Count == 3,
                "signature block lost: " + roundTrip[2].Rows.Count + "x" + roundTrip[2].Columns.Count);
            Require(ContainsCellText(roundTrip[0], "精装A座3号楼"), "client name lost in the commission grid");
            long read = timer.ElapsedMilliseconds;

            OfficeConverter.ToPdf(document, pdfPath);
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            Require(pdf.PageCount >= 2, "a 30-item inspection table should overflow to a second page: " + pdf.PageCount);
            int blankPages = 0;
            int repeatedHeaders = 0;
            int pageNumbers = 0;
            int itemHeaders = 0;
            for (int p = 0; p < pdf.PageCount; p++)
            {
                IReadOnlyList<PdfTextLine> lines = pdf.Pages[p].GetTextLines();
                StringBuilder text = new StringBuilder();
                for (int l = 0; l < lines.Count; l++) text.Append(lines[l].Text);
                string pageText = text.ToString();
                if (pageText.Trim().Length == 0) blankPages++;
                if (pageText.Contains("QF-QC-018")) repeatedHeaders++;
                if (pageText.Contains("技术要求")) itemHeaders++;
                if (pageText.Contains("第 " + (p + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " 页")) pageNumbers++;
            }
            Require(blankPages == 0, "inspection pdf has blank pages: " + blankPages);
            Require(repeatedHeaders == pdf.PageCount,
                "文件编号 header must repeat on every page: " + repeatedHeaders + "/" + pdf.PageCount);
            Require(pageNumbers == pdf.PageCount,
                "footer page number must be substituted on every page: " + pageNumbers + "/" + pdf.PageCount);
            Require(itemHeaders == pdf.PageCount,
                "the merged item header must repeat on every page: " + itemHeaders + "/" + pdf.PageCount);

            string all = string.Empty;
            for (int p = 0; p < pdf.PageCount; p++)
            {
                IReadOnlyList<PdfTextLine> lines = pdf.Pages[p].GetTextLines();
                for (int l = 0; l < lines.Count; l++) all += lines[l].Text + "\n";
            }
            Require(!all.Contains("{page}") && !all.Contains("{pages}"), "page tokens must not survive into pdf");
            RequireContains(all, "成品检验报告");
            RequireContains(all, "不合格");
            RequireContains(all, "耐洗刷性");
            RequireContains(all, "编制：");
            long pdfMs = timer.ElapsedMilliseconds;

            Console.WriteLine("[inspection] build=" + build + " ms, docx=" + write + " ms/"
                + new FileInfo(docxPath).Length + " bytes, read=" + (read - write) + " ms, pdf="
                + (pdfMs - read) + " ms/" + pdf.PageCount + " pages/" + new FileInfo(pdfPath).Length + " bytes");
        }

        /// <summary>
        /// 构造一张真实检验报告版式：A4 纵向、页眉文件编号、页脚页码域、四列委托信息网格、
        /// 六列检验项目表（多级合并表头 + 18 行数据以便跨页）与三列签署栏。
        /// Builds a real inspection-report layout: portrait A4, a header carrying the document number,
        /// footer page fields, a four-column commission grid, a six-column item table with merged
        /// multi-level headers and 18 rows so it overflows one page, and a three-column signature block.
        /// </summary>
        /// <returns>待导出的文档模型。The document model to export.</returns>
        private static OfficeDocument BuildInspectionDocument()
        {
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = "成品检验报告";
            document.Metadata.Author = "Wisdom.Office";

            PageSetup setup = PageSetup.A4();
            setup.MarginLeft = 56;
            setup.MarginRight = 56;
            setup.MarginBottom = 64;
            Section section = document.AddSection(setup);

            TextFormat meta = InspectionFormat(9, false);
            Paragraph header = section.Header.AddParagraph(string.Empty);
            header.Alignment = ParagraphAlignment.Right;
            header.AddText("文件编号：QF-QC-018　版本：A/2", meta);

            Paragraph footer = section.Footer.AddParagraph(string.Empty);
            footer.Alignment = ParagraphAlignment.Center;
            footer.AddText("第 ", meta);
            footer.Inlines.Add(new FieldRun("PAGE", "{page}", meta));
            footer.AddText(" 页　共 ", meta);
            footer.Inlines.Add(new FieldRun("NUMPAGES", "{pages}", meta));
            footer.AddText(" 页", meta);

            Paragraph title = section.AddParagraph(string.Empty);
            title.Alignment = ParagraphAlignment.Center;
            title.AddText("成品检验报告", InspectionFormat(18, true));
            title.HeadingLevel = 1;
            Paragraph subtitle = section.AddParagraph(string.Empty);
            subtitle.Alignment = ParagraphAlignment.Center;
            subtitle.AddText("报告编号：JC-2026-0917　检验类别：出厂检验", InspectionFormat(9, false));

            OfficeColor label = new OfficeColor(238, 242, 245);
            Table commission = new Table();
            commission.Autofit = false;
            commission.RepeatHeaderRow = false;
            int[] commissionWidths = { 78, 146, 78, 142 };
            for (int c = 0; c < commissionWidths.Length; c++) commission.Columns.Add(new TableColumn(commissionWidths[c]));
            string[,] pairs =
            {
                { "产品名称", "内墙乳胶漆", "规格型号", "TH-200" },
                { "生产批号", "P20260912", "检验数量", "1 200 件" },
                { "委托单位", "某某建设装饰有限公司", "工程部位", "精装A座3号楼" },
                { "检验日期", "2026-09-17", "检验依据", "GB/T 9756-2019" }
            };
            for (int r = 0; r < pairs.GetLength(0); r++)
            {
                TableRow row = commission.AddRow();
                for (int c = 0; c < 4; c++)
                {
                    OfficeColor? fill = c % 2 == 0 ? new OfficeColor?(label) : null;
                    AddInspectionCell(row, pairs[r, c], 1, 1, fill);
                }
            }
            section.Blocks.Add(commission);

            Table items = new Table();
            items.Autofit = false;
            items.RepeatHeaderRow = true;
            int[] itemWidths = { 28, 84, 104, 64, 58, 58, 46 };
            for (int c = 0; c < itemWidths.Length; c++) items.Columns.Add(new TableColumn(itemWidths[c]));

            TableRow top = items.AddRow();
            top.IsHeader = true;
            top.Height = 22;
            AddInspectionCell(top, "序号", 1, 2, null);
            AddInspectionCell(top, "检验项目", 1, 2, null);
            AddInspectionCell(top, "技术要求", 1, 2, null);
            AddInspectionCell(top, "检验方法", 1, 2, null);
            AddInspectionCell(top, "实测结果", 2, 1, null);
            AddInspectionCell(top, "判定", 1, 2, null);
            TableRow sub = items.AddRow();
            sub.IsHeader = true;
            AddInspectionCell(sub, "单件一", 1, 1, null);
            AddInspectionCell(sub, "单件二", 1, 1, null);

            string[] names = { "容器中状态", "施工性", "干燥时间", "涂膜外观", "耐碱性", "耐洗刷性", "对比率", "附着强度" };
            for (int r = 0; r < 30; r++)
            {
                TableRow row = items.AddRow();
                row.Height = 24;
                string group = names[r % names.Length];
                AddInspectionCell(row, (r + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), 1, 1, null);
                AddInspectionCell(row, group + (r >= names.Length ? "（复测）" : string.Empty), 1, 1, null);
                AddInspectionCell(row, "标准值 ≤0.5g/L", 1, 1, null);
                AddInspectionCell(row, "GB/T 9756", 1, 1, null);
                AddInspectionCell(row, "0." + (r + 12).ToString(System.Globalization.CultureInfo.InvariantCulture), 1, 1, null);
                AddInspectionCell(row, "0." + (r + 21).ToString(System.Globalization.CultureInfo.InvariantCulture), 1, 1, null);
                AddInspectionCell(row, r == 5 ? "不合格" : "合格", 1, 1, r == 5 ? new OfficeColor?(new OfficeColor(255, 230, 230)) : null);
            }
            section.Blocks.Add(items);

            Paragraph verdict = section.AddParagraph(string.Empty);
            verdict.SpaceBefore = 10;
            verdict.AddText("检验结论：", InspectionFormat(11, true));
            verdict.AddText("本批产品按 GB/T 9756-2019 检验，所检项目中耐洗刷性不合格，其余合格，判本批不合格，退回返工后复验。", InspectionFormat(11, false));
            Paragraph note = section.AddParagraph(string.Empty);
            note.AddText("备注：抽样方案见 QB-QC-003，本报告不经批准无效。", InspectionFormat(9, false));

            Table signature = new Table();
            signature.Autofit = false;
            signature.RepeatHeaderRow = false;
            signature.Borders = TableBorders.None();
            for (int c = 0; c < 3; c++) signature.Columns.Add(new TableColumn(148));
            TableRow roles = signature.AddRow();
            roles.Height = 30;
            AddInspectionCell(roles, "编制：姚蕾", 1, 1, null);
            AddInspectionCell(roles, "审核：李震", 1, 1, null);
            AddInspectionCell(roles, "批准：王平", 1, 1, null);
            TableRow dates = signature.AddRow();
            AddInspectionCell(dates, "日期：2026-09-17", 1, 1, null);
            AddInspectionCell(dates, "日期：2026-09-17", 1, 1, null);
            AddInspectionCell(dates, "日期：2026-09-18", 1, 1, null);
            section.Blocks.Add(signature);

            return document;
        }

        /// <summary>检验报告使用的文本格式：宋体、指定字号与加粗。</summary>
        /// <param name="size">字号（磅）。Font size in points.</param>
        /// <param name="bold">是否加粗。Whether the text is bold.</param>
        /// <returns>文本格式实例。A fresh text format.</returns>
        private static TextFormat InspectionFormat(double size, bool bold)
        {
            TextFormat format = new TextFormat("SimSun", size);
            format.EastAsiaFontFamily = "宋体";
            format.Bold = bold;
            return format;
        }

        /// <summary>向表格行追加一个垂直居中的单元格，可选加粗标签底色。</summary>
        /// <param name="row">目标行。The row to append to.</param>
        /// <param name="text">单元格文本。Cell text.</param>
        /// <param name="colSpan">横向跨度。Horizontal span.</param>
        /// <param name="rowSpan">纵向跨度。Vertical span.</param>
        /// <param name="background">背景色，null 表示无填充。Background fill, null for none.</param>
        /// <returns>新建的单元格。The created cell.</returns>
        private static TableCell AddInspectionCell(TableRow row, string text, int colSpan, int rowSpan, OfficeColor? background)
        {
            TableCell cell = new TableCell();
            cell.ColSpan = colSpan;
            cell.RowSpan = rowSpan;
            cell.VerticalAlignment = VerticalAlignment.Middle;
            cell.BackgroundColor = background;
            Paragraph paragraph = new Paragraph();
            paragraph.AddText(text, InspectionFormat(9.5f, background != null));
            cell.Blocks.Add(paragraph);
            row.Cells.Add(cell);
            return cell;
        }

        private static void RunSecurityTest(string outputDirectory, string seedPdfPath, string seedDocxPath, string seedXlsxPath)
        {
            string securityDirectory = Path.Combine(outputDirectory, "security");
            Directory.CreateDirectory(securityDirectory);
            Stopwatch watch = new Stopwatch();

            string xxeXml = "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<!DOCTYPE document [<!ENTITY xxe SYSTEM \"file:///C:/Windows/win.ini\">]>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:body><w:p><w:r><w:t>&xxe;</w:t></w:r></w:p></w:body></w:document>";

            Dictionary<string, byte[]> docxEntries = new Dictionary<string, byte[]>();
            docxEntries["word/document.xml"] = Encoding.UTF8.GetBytes(xxeXml);
            byte[] xxeDocx = CreateZipBytes(docxEntries);
            watch.Restart();
            bool xxeDocxBlocked = ThrowsInvalidData(delegate { DocxReader.Load(xxeDocx); });
            watch.Stop();
            Require(xxeDocxBlocked, "docx external entity payload was not rejected");
            Require(watch.ElapsedMilliseconds < 5000, "docx xxe rejection too slow");

            Dictionary<string, byte[]> xlsxEntries = new Dictionary<string, byte[]>();
            xlsxEntries["xl/workbook.xml"] = Encoding.UTF8.GetBytes(xxeXml);
            byte[] xxeXlsx = CreateZipBytes(xlsxEntries);
            bool xxeXlsxBlocked = ThrowsInvalidData(delegate { XlsxReader.Load(xxeXlsx); });
            Require(xxeXlsxBlocked, "xlsx external entity payload was not rejected");

            long originalEntryLimit = OfficePackageLimits.MaxEntryBytes;
            long bombMs = 0;
            try
            {
                OfficePackageLimits.MaxEntryBytes = 1024 * 1024;
                Dictionary<string, byte[]> bombEntries = new Dictionary<string, byte[]>();
                bombEntries["word/document.xml"] = new byte[8 * 1024 * 1024];
                byte[] bomb = CreateZipBytes(bombEntries);
                watch.Restart();
                bool bombBlocked = ThrowsInvalidData(delegate { DocxReader.Load(bomb); });
                watch.Stop();
                bombMs = watch.ElapsedMilliseconds;
                Require(bombBlocked, "zip bomb payload was not rejected");
                Require(bombMs < 2000, "zip bomb rejection too slow: " + bombMs + " ms");
            }
            finally
            {
                OfficePackageLimits.MaxEntryBytes = originalEntryLimit;
            }

            // 深层嵌套表格会让块级解析递归下探，畸形文档可以借此把调用栈打满导致进程崩溃（不可捕获）。
            // 读取器按 MaxBlockDepth 停止下探，这里用 2000 层嵌套表验证它只是少解析几层，而不是崩溃。
            // Deeply nested tables drive the block walk deeper and deeper; a malformed document can use them to
            // exhaust the call stack (an uncatchable crash). The reader stops at its depth cap instead, so here a
            // 2000-level document must still load with only the first few levels parsed.
            StringBuilder nestedBuilder = new StringBuilder();
            nestedBuilder.Append("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>");
            for (int level = 0; level < 2000; level++) nestedBuilder.Append("<w:tbl><w:tr><w:tc>");
            for (int level = 0; level < 2000; level++) nestedBuilder.Append("</w:tc></w:tr></w:tbl>");
            nestedBuilder.Append("</w:body></w:document>");
            Dictionary<string, byte[]> nestedEntries = new Dictionary<string, byte[]>();
            nestedEntries["word/document.xml"] = Encoding.UTF8.GetBytes(nestedBuilder.ToString());
            OfficeDocument nestedDocument = DocxReader.Load(CreateZipBytes(nestedEntries));
            int nestedDepth = 0;
            for (int s = 0; s < nestedDocument.Sections.Count; s++)
            {
                int sectionDepth = MeasureTableDepth(nestedDocument.Sections[s].Blocks);
                if (sectionDepth > nestedDepth) nestedDepth = sectionDepth;
            }
            Require(nestedDepth > 0 && nestedDepth <= 65,
                "nested table depth is not bounded: " + nestedDepth);

            // 工作表部件走限量流式读取：造一个大表，把上限压到只有它会越界，
            // 证明单件与累计上限在流式路径上同样生效，且默认上限下仍能正常读回。
            // Worksheet parts are read through a quota-bounded stream: build one sheet large enough to be the
            // only part over the tightened limits, proving both the per-entry and cumulative guards fire while
            // streaming, and that the default limits still load it fine.
            long originalTotalLimit = OfficePackageLimits.MaxTotalBytes;
            byte[] streamingWorkbook = BuildStreamingProbePackage();
            watch.Restart();
            Require(XlsxReader.Load(streamingWorkbook).Worksheets.Count == 1,
                "streaming probe workbook failed to load with the default limits");
            watch.Stop();
            long streamMs = watch.ElapsedMilliseconds;
            try
            {
                OfficePackageLimits.MaxEntryBytes = 65536;
                Require(ThrowsInvalidData(delegate { XlsxReader.Load(streamingWorkbook); }),
                    "streamed worksheet ignored the per-entry decompression limit");
            }
            finally
            {
                OfficePackageLimits.MaxEntryBytes = originalEntryLimit;
            }
            try
            {
                OfficePackageLimits.MaxTotalBytes = 65536;
                Require(ThrowsInvalidData(delegate { XlsxReader.Load(streamingWorkbook); }),
                    "streamed worksheet ignored the total decompression limit");
            }
            finally
            {
                OfficePackageLimits.MaxTotalBytes = originalTotalLimit;
            }

            try
            {
                OfficePackageLimits.MaxTotalBytes = 1024 * 1024;
                Dictionary<string, byte[]> multiPart = new Dictionary<string, byte[]>();
                multiPart["word/document.xml"] = new byte[700 * 1024];
                multiPart["word/styles.xml"] = new byte[700 * 1024];
                byte[] package = CreateZipBytes(multiPart);
                bool totalBlocked = false;
                using (MemoryStream memory = new MemoryStream(package))
                using (ZipArchive archive = new ZipArchive(memory, ZipArchiveMode.Read))
                {
                    long totalBytes = 0;
                    try
                    {
                        OfficePackageLimits.ReadEntryBytes(archive.GetEntry("word/document.xml"), ref totalBytes);
                        OfficePackageLimits.ReadEntryBytes(archive.GetEntry("word/styles.xml"), ref totalBytes);
                    }
                    catch (InvalidDataException)
                    {
                        totalBlocked = true;
                    }
                }
                Require(totalBlocked, "package total decompression limit was not enforced");
            }
            finally
            {
                OfficePackageLimits.MaxTotalBytes = originalTotalLimit;
            }

            byte[] compressed = Zlib.Compress(new byte[4 * 1024 * 1024]);
            bool zlibBlocked = false;
            try
            {
                Zlib.Decompress(compressed, 0, compressed.Length, 1024 * 1024);
            }
            catch (DecompressionLimitException)
            {
                zlibBlocked = true;
            }
            Require(zlibBlocked, "zlib decompression limit was not enforced");

            Dictionary<string, byte[]> traversalEntries = new Dictionary<string, byte[]>();
            traversalEntries["word/document.xml"] = Encoding.UTF8.GetBytes(
                "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
                + "<w:body><w:p><w:r><w:t>客户：{{name}}</w:t></w:r></w:p></w:body></w:document>");
            traversalEntries["../evil.txt"] = Encoding.UTF8.GetBytes("path traversal");
            string traversalTemplate = PrepareOutputPath(Path.Combine(securityDirectory, "traversal-template.docx"));
            File.WriteAllBytes(traversalTemplate, CreateZipBytes(traversalEntries));
            string traversalOutput = PrepareOutputPath(Path.Combine(securityDirectory, "traversal-filled.docx"));
            TemplateContext traversalContext = new TemplateContext();
            traversalContext.Set("name", "安全测试");
            bool traversalBlocked = ThrowsInvalidData(delegate { DocxTemplate.Fill(traversalTemplate, traversalOutput, traversalContext); });
            Require(traversalBlocked, "path traversal package was not rejected");
            Require(!File.Exists(traversalOutput), "rejected traversal package still produced an output file");
            Require(!File.Exists(Path.Combine(securityDirectory, "evil.txt")), "path traversal wrote evil.txt beside the output");
            Require(!File.Exists(Path.Combine(outputDirectory, "evil.txt")), "path traversal escaped into the output directory");

            Random random = new Random(1234);
            Exception unexpected = null;
            int accepted = 0;
            watch.Restart();
            for (int i = 0; i < 120; i++)
            {
                byte[] data = new byte[16 + random.Next(48 * 1024)];
                random.NextBytes(data);
                if (i % 4 == 0)
                {
                    byte[] headerBytes = Encoding.ASCII.GetBytes("%PDF-1.7\n");
                    Buffer.BlockCopy(headerBytes, 0, data, 0, headerBytes.Length);
                }
                try
                {
                    PdfDocument.Load(data);
                    accepted++;
                }
                catch (InvalidDataException)
                {
                }
                catch (NotSupportedException)
                {
                }
                catch (DecompressionLimitException)
                {
                }
                catch (Exception ex)
                {
                    unexpected = ex;
                    break;
                }
            }
            watch.Stop();
            long fuzzMs = watch.ElapsedMilliseconds;
            Require(unexpected == null, "pdf fuzz raised an unexpected exception: " + unexpected);
            Require(fuzzMs < 10000, "pdf fuzz too slow: " + fuzzMs + " ms");

            StringBuilder bombPdf = new StringBuilder();
            bombPdf.Append("%PDF-1.7\n");
            bombPdf.Append("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
            bombPdf.Append("2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n");
            int xrefOffset = Encoding.ASCII.GetByteCount(bombPdf.ToString());
            bombPdf.Append("xref\n0 2000000000\n");
            bombPdf.Append("0000000000 65535 f \n");
            bombPdf.Append("trailer\n<< /Size 2000000000 /Root 1 0 R >>\n");
            bombPdf.Append("startxref\n").Append(xrefOffset).Append("\n%%EOF\n");
            byte[] xrefBomb = Encoding.ASCII.GetBytes(bombPdf.ToString());
            watch.Restart();
            bool xrefHandled = false;
            try
            {
                PdfDocument xrefDocument = PdfDocument.Load(xrefBomb);
                Require(xrefDocument.PageCount == 0, "crafted xref pdf produced unexpected pages");
                xrefHandled = true;
            }
            catch (InvalidDataException)
            {
                xrefHandled = true;
            }
            watch.Stop();
            long xrefMs = watch.ElapsedMilliseconds;
            Require(xrefHandled, "crafted xref pdf was not handled");
            Require(xrefMs < 5000, "crafted xref pdf processed too slowly: " + xrefMs + " ms");

            int mutatedAccepted = 0;
            Exception mutatedPdfFailure = null;
            watch.Restart();
            if (File.Exists(seedPdfPath))
            {
                byte[] seedPdf = File.ReadAllBytes(seedPdfPath);
                for (int i = 0; i < 60 && mutatedPdfFailure == null; i++)
                {
                    byte[] mutated = MutateBytes(random, seedPdf, 1 + random.Next(12));
                    try
                    {
                        PdfDocument mutatedDocument = PdfDocument.Load(mutated);
                        mutatedDocument.ExtractText();
                        mutatedAccepted++;
                    }
                    catch (InvalidDataException)
                    {
                    }
                    catch (NotSupportedException)
                    {
                    }
                    catch (DecompressionLimitException)
                    {
                    }
                    catch (Exception ex)
                    {
                        mutatedPdfFailure = ex;
                    }
                }
            }
            watch.Stop();
            long mutatedPdfMs = watch.ElapsedMilliseconds;
            Require(mutatedPdfFailure == null, "mutated pdf raised an unexpected exception: " + mutatedPdfFailure);
            Require(mutatedPdfMs < 10000, "mutated pdf fuzzing too slow: " + mutatedPdfMs + " ms");

            // 空名称（光杆斜杠）属畸形词元，必须按已声明的 InvalidDataException 上报；
            // PdfName 构造函数的 ArgumentException 只留给写侧调用方，不能漏过交叉表与内容流的窄 catch。
            // A bare slash is a malformed token and must surface as the documented InvalidDataException:
            // PdfName's ArgumentException belongs to the writing side and must not slip past the narrow
            // recovery catches in the cross-reference walk or the content interpreter.
            byte[] bareNameTrailer = Encoding.ASCII.GetBytes(
                "%PDF-1.7\n1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n"
                + "2 0 obj\n<< /Type /Pages /Kids [] /Count 0 >>\nendobj\n"
                + "xref\n0 3\n0000000000 65535 f \ntrailer\n<< /Size 3 /Root 1 0 R / >>\nstartxref\n15\n%%EOF\n");
            Exception bareNameTrailerFailure = null;
            try
            {
                PdfDocument.Load(bareNameTrailer);
            }
            catch (InvalidDataException)
            {
            }
            catch (Exception ex)
            {
                bareNameTrailerFailure = ex;
            }
            Require(bareNameTrailerFailure == null, "trailer slash name leaked " + bareNameTrailerFailure);

            if (File.Exists(seedPdfPath))
            {
                byte[] slashSeed = File.ReadAllBytes(seedPdfPath);
                string slashBefore = PdfDocument.Load(slashSeed).ExtractText();
                PdfFile slashFile = PdfFile.Load(slashSeed);
                slashFile.PrependPageContent(0, Encoding.ASCII.GetBytes("/ "));
                Exception slashContentFailure = null;
                string slashAfter = null;
                try
                {
                    slashAfter = PdfDocument.Load(slashFile.SaveToBytes()).ExtractText();
                }
                catch (Exception ex)
                {
                    slashContentFailure = ex;
                }
                Require(slashContentFailure == null, "content stream slash name leaked " + slashContentFailure);
                Require(slashAfter == slashBefore, "content stream slash name changed the extracted text");
            }

            Exception mutatedDocxFailure = null;
            int mutatedDocxAccepted = 0;
            if (File.Exists(seedDocxPath))
            {
                byte[] seedDocx = File.ReadAllBytes(seedDocxPath);
                for (int i = 0; i < 40 && mutatedDocxFailure == null; i++)
                {
                    byte[] mutated = MutateBytes(random, seedDocx, 1 + random.Next(6));
                    try
                    {
                        DocxReader.Load(mutated);
                        mutatedDocxAccepted++;
                    }
                    catch (InvalidDataException)
                    {
                    }
                    catch (NotSupportedException)
                    {
                    }
                    catch (DecompressionLimitException)
                    {
                    }
                    catch (Exception ex)
                    {
                        mutatedDocxFailure = ex;
                    }
                }
            }
            Require(mutatedDocxFailure == null, "mutated docx raised an unexpected exception: " + mutatedDocxFailure);

            Exception mutatedXlsxFailure = null;
            int mutatedXlsxAccepted = 0;
            if (File.Exists(seedXlsxPath))
            {
                byte[] seedXlsx = File.ReadAllBytes(seedXlsxPath);
                for (int i = 0; i < 40 && mutatedXlsxFailure == null; i++)
                {
                    byte[] mutated = MutateBytes(random, seedXlsx, 1 + random.Next(6));
                    try
                    {
                        XlsxReader.Load(mutated);
                        mutatedXlsxAccepted++;
                    }
                    catch (InvalidDataException)
                    {
                    }
                    catch (NotSupportedException)
                    {
                    }
                    catch (DecompressionLimitException)
                    {
                    }
                    catch (Exception ex)
                    {
                        mutatedXlsxFailure = ex;
                    }
                }
            }
            Require(mutatedXlsxFailure == null, "mutated xlsx raised an unexpected exception: " + mutatedXlsxFailure);

            // Markdown 行内配对：旧实现从每个未闭合的 [ 、( 、< 起向后重扫剩余文本，18 万字符的畸形行内串会退化到
            // O(N²) 而停顿数十秒；改为按文本一次性建配对表后只允许线性耗时。
            // Markdown inline pairing: the old scan restarted from every unmatched '[', '(' or '<' and re-walked the
            // remaining text, so a 180k-character hostile inline degraded to O(N²) and stalled for tens of seconds.
            // Building one pairing table per text must keep the parse linear.
            string hostileInline = new string('[', 60000) + new string('<', 60000) + new string('(', 60000);
            Stopwatch inlineWatch = Stopwatch.StartNew();
            OfficeDocument hostileMarkdown = Wisdom.Markdown.Markdown.ToDocument(hostileInline);
            inlineWatch.Stop();
            long inlineMs = inlineWatch.ElapsedMilliseconds;
            string inlineText = hostileMarkdown.GetPlainText();
            Require(CountChar(inlineText, '[') == 60000 && CountChar(inlineText, '<') == 60000
                && CountChar(inlineText, '(') == 60000,
                "hostile inline text lost or invented brackets: " + inlineText.Length);
            Require(inlineMs < 3000, "markdown inline parse is not linear: " + inlineMs + " ms");

            Console.WriteLine("[security] xxe-docx=blocked, xxe-xlsx=blocked, zip-bomb=blocked(" + bombMs
                + " ms), total-limit=blocked, zlib-limit=blocked, traversal=rejected, nested-tables="
                + nestedDepth + "/2000 levels, stream-limits=blocked(" + streamMs + " ms), markdown-inline="
                + inlineMs + " ms/180k, fuzz=" + accepted
                + "/120 accepted in " + fuzzMs + " ms, xref=" + xrefMs + " ms, mutated pdf=" + mutatedAccepted
                + "/60 accepted in " + mutatedPdfMs + " ms, mutated docx=" + mutatedDocxAccepted
                + "/40, mutated xlsx=" + mutatedXlsxAccepted + "/40");
        }

        private static byte[] MutateBytes(Random random, byte[] source, int mutations)
        {
            byte[] data = (byte[])source.Clone();
            for (int i = 0; i < mutations; i++)
            {
                int position = random.Next(data.Length);
                data[position] = (byte)random.Next(256);
            }
            return data;
        }

        /// <summary>统计字符在文本中出现的次数。Counts how many times a character occurs in the text.</summary>
        /// <param name="text">待统计文本。The text to scan.</param>
        /// <param name="target">目标字符。The character to count.</param>
        /// <returns>出现次数。The occurrence count.</returns>
        private static int CountChar(string text, char target)
        {
            int count = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == target) count++;
            }
            return count;
        }

        /// <summary>
        /// 递归统计块列表中的最大表格嵌套层数（单元格内容继续下探）。
        /// Measures, recursively, the deepest table nesting in a block list; cell content continues the walk.
        /// </summary>
        /// <param name="blocks">待统计的块列表。The blocks to measure.</param>
        /// <returns>最大表格嵌套层数，无表格时为 0。The deepest table level, 0 when no table is present.</returns>
        private static int MeasureTableDepth(List<Block> blocks)
        {
            int max = 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Table table = blocks[i] as Table;
                if (table == null) continue;
                int childMax = 0;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        int depth = MeasureTableDepth(table.Rows[r].Cells[c].Blocks);
                        if (depth > childMax) childMax = depth;
                    }
                }
                if (childMax + 1 > max) max = childMax + 1;
            }
            return max;
        }

        /// <summary>
        /// 构造一个只有工作表部件会超过 64 KB 的小 xlsx，用于校验流式读取路径上的解压上限。
        /// Builds a small xlsx whose worksheet part is the only one exceeding 64 KB, used to verify the
        /// decompression limits on the streaming read path.
        /// </summary>
        /// <returns>xlsx 包字节。The xlsx package bytes.</returns>
        private static byte[] BuildStreamingProbePackage()
        {
            OfficeWorkbook probe = new OfficeWorkbook();
            Worksheet sheet = probe.AddWorksheet("流式上限");
            for (int row = 1; row <= 2000; row++)
            {
                for (int column = 1; column <= 4; column++)
                {
                    sheet.GetCell(row, column).Value = (double)(row * column);
                }
            }
            return XlsxWriter.Write(probe);
        }

        private static bool ThrowsInvalidData(Action action)
        {
            try
            {
                action();
            }
            catch (InvalidDataException)
            {
                return true;
            }
            return false;
        }

        private static byte[] CreateZipBytes(Dictionary<string, byte[]> entries)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                using (ZipArchive archive = new ZipArchive(memory, ZipArchiveMode.Create, true))
                {
                    foreach (KeyValuePair<string, byte[]> pair in entries)
                    {
                        ZipArchiveEntry entry = archive.CreateEntry(pair.Key, CompressionLevel.Optimal);
                        using (Stream stream = entry.Open())
                        {
                            stream.Write(pair.Value, 0, pair.Value.Length);
                        }
                    }
                }
                return memory.ToArray();
            }
        }

        private static int CountParagraphs(OfficeDocument document)
        {
            int count = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                count += CountParagraphs(document.Sections[s].Blocks);
            }
            return count;
        }

        private static int CountParagraphs(List<Block> blocks)
        {
            int count = 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    count++;
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        count += CountParagraphs(table.Rows[r].Cells[c].Blocks);
                    }
                }
            }
            return count;
        }

        private static int CountTables(OfficeDocument document)
        {
            int count = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                count += CountTablesInBlocks(document.Sections[s].Blocks);
            }
            return count;
        }

        private static int CountTablesInBlocks(List<Block> blocks)
        {
            int count = 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Table table = blocks[i] as Table;
                if (table == null) continue;
                count++;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        count += CountTablesInBlocks(table.Rows[r].Cells[c].Blocks);
                    }
                }
            }
            return count;
        }

        private static int CountImages(OfficeDocument document)
        {
            int count = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                count += CountImagesInBlocks(document.Sections[s].Blocks);
            }
            return count;
        }

        private static int CountImagesInBlocks(List<Block> blocks)
        {
            int count = 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    for (int k = 0; k < paragraph.Inlines.Count; k++)
                    {
                        if (paragraph.Inlines[k] is ImageRun) count++;
                    }
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        count += CountImagesInBlocks(table.Rows[r].Cells[c].Blocks);
                    }
                }
            }
            return count;
        }

private static int CountHeadings(OfficeDocument document)
{
    int count = 0;
    for (int s = 0; s < document.Sections.Count; s++)
    {
        for (int i = 0; i < document.Sections[s].Blocks.Count; i++)
        {
            Paragraph paragraph = document.Sections[s].Blocks[i] as Paragraph;
            if (paragraph != null && paragraph.HeadingLevel > 0 && paragraph.GetText().Trim() != "目录") count++;
        }
    }
    return count;
}

private static int CountTocEntries(OfficeDocument document)
{
    int count = 0;
    for (int s = 0; s < document.Sections.Count; s++)
    {
        for (int i = 0; i < document.Sections[s].Blocks.Count; i++)
        {
            Paragraph paragraph = document.Sections[s].Blocks[i] as Paragraph;
            if (paragraph != null && HasFieldInstruction(paragraph, "PAGEREF")) count++;
        }
    }
    return count;
}

private static Paragraph FindParagraph(OfficeDocument document, Predicate<Paragraph> match)
{
    for (int s = 0; s < document.Sections.Count; s++)
    {
        Paragraph found = FindParagraph(document.Sections[s].Blocks, match);
        if (found != null) return found;
    }
    return null;
}

private static Paragraph FindParagraph(List<Block> blocks, Predicate<Paragraph> match)
{
    for (int i = 0; i < blocks.Count; i++)
    {
        Paragraph paragraph = blocks[i] as Paragraph;
        if (paragraph != null && match(paragraph)) return paragraph;
        Table table = blocks[i] as Table;
        if (table == null) continue;
        for (int r = 0; r < table.Rows.Count; r++)
        {
            for (int c = 0; c < table.Rows[r].Cells.Count; c++)
            {
                Paragraph found = FindParagraph(table.Rows[r].Cells[c].Blocks, match);
                if (found != null) return found;
            }
        }
    }
    return null;
}

private static bool HasFieldInstruction(Paragraph paragraph, string prefix)
{
    for (int i = 0; i < paragraph.Inlines.Count; i++)
    {
        FieldRun field = paragraph.Inlines[i] as FieldRun;
        if (field != null && (field.Instruction ?? string.Empty).StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return true;
    }
    return false;
}

private static bool HasFieldDisplay(Paragraph paragraph, string bookmark, string display)
{
    for (int i = 0; i < paragraph.Inlines.Count; i++)
    {
        FieldRun field = paragraph.Inlines[i] as FieldRun;
        if (field == null) continue;
        string instruction = field.Instruction ?? string.Empty;
        if (instruction.IndexOf(bookmark, StringComparison.Ordinal) >= 0 && field.DisplayText == display) return true;
    }
    return false;
}

private static bool ContainsPageBreak(List<Block> blocks)
{
    for (int i = 0; i < blocks.Count; i++)
    {
        if (blocks[i] is PageBreak) return true;
    }
    return false;
}

private static bool ContainsTinyImage(OfficeDocument document)
{
    for (int s = 0; s < document.Sections.Count; s++)
    {
        if (ContainsTinyImage(document.Sections[s].Blocks)) return true;
    }
    return false;
}

private static bool ContainsTinyImage(List<Block> blocks)
{
    for (int i = 0; i < blocks.Count; i++)
    {
        Paragraph paragraph = blocks[i] as Paragraph;
        if (paragraph != null)
        {
            for (int r = 0; r < paragraph.Inlines.Count; r++)
            {
                ImageRun image = paragraph.Inlines[r] as ImageRun;
                if (image == null || image.Image == null) continue;
                if (image.Image.PixelWidth <= 1 && image.Image.PixelHeight <= 1) return true;
            }
            continue;
        }
        Table table = blocks[i] as Table;
        if (table == null) continue;
        for (int r = 0; r < table.Rows.Count; r++)
        {
            for (int c = 0; c < table.Rows[r].Cells.Count; c++)
            {
                if (ContainsTinyImage(table.Rows[r].Cells[c].Blocks)) return true;
            }
        }
    }
    return false;
}

private static void CheckBodyFonts(List<Block> blocks, WordNormalizeOptions options, ref int checkedRuns)
{
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
                    CheckBodyFonts(table.Rows[r].Cells[c].Blocks, options, ref checkedRuns);
                }
            }
            continue;
        }
        if (paragraph.HeadingLevel > 0 || HasFieldInstruction(paragraph, "PAGEREF")) continue;
        if (paragraph.GetText().Trim() == "目录") continue;
        for (int r = 0; r < paragraph.Inlines.Count; r++)
        {
            TextRun run = paragraph.Inlines[r] as TextRun;
            if (run == null || run.Format == null) continue;
            Require(run.Format.GetEastAsiaFontFamily() == options.BodyEastAsiaFont, "word normalize body eastAsia font mismatch: " + run.Format.GetEastAsiaFontFamily());
            Require(run.Format.GetAsciiFontFamily() == options.BodyAsciiFont, "word normalize body ascii font mismatch: " + run.Format.GetAsciiFontFamily());
            Require(Math.Abs(run.Format.FontSize - options.BodyFontSize) < 0.01, "word normalize body font size mismatch: " + run.Format.FontSize);
            checkedRuns++;
        }
    }
}

private static string PrepareOutputPath(string path)
        {
            try
            {
                using (new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                }
                return path;
            }
            catch (FileNotFoundException)
            {
                return path;
            }
            catch (DirectoryNotFoundException)
            {
                return path;
            }
            catch (IOException)
            {
                string directory = Path.GetDirectoryName(path);
                string name = Path.GetFileNameWithoutExtension(path);
                string extension = Path.GetExtension(path);
                return Path.Combine(directory, name + "-" + DateTime.Now.ToString("HHmmssfff") + extension);
            }
            catch (UnauthorizedAccessException)
            {
                string directory = Path.GetDirectoryName(path);
                string name = Path.GetFileNameWithoutExtension(path);
                string extension = Path.GetExtension(path);
                return Path.Combine(directory, name + "-" + DateTime.Now.ToString("HHmmssfff") + extension);
            }
        }

        private static void RequireContains(string text, string expected)
        {
            Require(text.IndexOf(expected, StringComparison.Ordinal) >= 0, "text not found: " + expected);
        }

        private static void Require(bool condition, string message)
        {
            if (!condition) throw new InvalidOperationException("Test failed: " + message);
        }

        private static byte[] CreateTestPng(int width, int height)
        {
            byte[] raw = new byte[height * (width * 4 + 1)];
            int stride = width * 4 + 1;
            for (int y = 0; y < height; y++)
            {
                int row = y * stride;
                raw[row] = 0;
                for (int x = 0; x < width; x++)
                {
                    int offset = row + 1 + x * 4;
                    byte alpha = (byte)(40 + 215 * (height - y) / height);
                    raw[offset] = (byte)(40 + 180 * x / width);
                    raw[offset + 1] = (byte)(90 + 120 * y / height);
                    raw[offset + 2] = (byte)(200 - 120 * x / width);
                    raw[offset + 3] = alpha;
                }
            }

            using (MemoryStream stream = new MemoryStream())
            {
                byte[] signature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
                stream.Write(signature, 0, signature.Length);

                using (MemoryStream header = new MemoryStream())
                {
                    WriteBigEndian(header, width);
                    WriteBigEndian(header, height);
                    header.WriteByte(8);
                    header.WriteByte(6);
                    header.WriteByte(0);
                    header.WriteByte(0);
                    header.WriteByte(0);
                    WriteChunk(stream, "IHDR", header.ToArray());
                }
                WriteChunk(stream, "IDAT", Zlib.Compress(raw));
                WriteChunk(stream, "IEND", new byte[0]);
                return stream.ToArray();
            }
        }

        private static void WriteChunk(Stream stream, string type, byte[] data)
        {
            byte[] typeBytes = Encoding.ASCII.GetBytes(type);
            WriteBigEndian(stream, data.Length);
            stream.Write(typeBytes, 0, typeBytes.Length);
            stream.Write(data, 0, data.Length);
            uint crc = ComputeCrc(typeBytes, data);
            stream.WriteByte((byte)(crc >> 24));
            stream.WriteByte((byte)(crc >> 16));
            stream.WriteByte((byte)(crc >> 8));
            stream.WriteByte((byte)crc);
        }

        private static void WriteBigEndian(Stream stream, int value)
        {
            stream.WriteByte((byte)(value >> 24));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)value);
        }

        private static uint ComputeCrc(byte[] first, byte[] second)
        {
            uint crc = 0xFFFFFFFF;
            crc = UpdateCrc(crc, first);
            crc = UpdateCrc(crc, second);
            return crc ^ 0xFFFFFFFF;
        }

        private static uint UpdateCrc(uint crc, byte[] data)
        {
            for (int i = 0; i < data.Length; i++)
            {
                crc ^= data[i];
                for (int bit = 0; bit < 8; bit++)
                {
                    if ((crc & 1) != 0) crc = (crc >> 1) ^ 0xEDB88320;
                    else crc >>= 1;
                }
            }
            return crc;
        }
    }
}
