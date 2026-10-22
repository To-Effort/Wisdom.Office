using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>执行自定义属性、标签纸与信封的边界测试。Runs edge-case tests for custom properties, label sheets and envelopes.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWordLabelsAndPropertiesEdgeTest(string outputDirectory)
        {
            RunCustomPropertyEdgeTest(outputDirectory);
            RunLabelEdgeTest(outputDirectory);
            RunLabelFormTest(outputDirectory);
            RunRetentionLabelSheetTest(outputDirectory);
            RunChemicalSafetyLabelTest(outputDirectory);
            RunEnvelopeEdgeTest(outputDirectory);
            Console.WriteLine("Word label/property edge case tests finished.");
        }

        /// <summary>验证特殊文本、高精度时间、浮点极值、大小写覆盖、DOTX 往返与非法输入。Verifies special text, high-precision timestamps, floating-point extremes, case-insensitive overwrite, DOTX round trips and invalid input.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunCustomPropertyEdgeTest(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-custom-properties-edge.docx"));
            OfficeDocument document = new OfficeDocument();
            document.AddSection().AddParagraph("边界测试正文。");
            CustomPropertyCollection properties = document.Metadata.CustomProperties;

            string special = "<标签 & \"引号\" '单引号'>\n第二行";
            properties.Set("Special", special);
            properties.Set("Version", 1);
            properties.Set("version", 2);
            Require(properties.Count == 2, "case-insensitive names should share one slot: " + properties.Count);
            Require(properties.GetInt32("VERSION") == 2, "case-insensitive lookup failed");

            DateTime precise = new DateTime(2026, 10, 20, 8, 30, 0, DateTimeKind.Utc).AddTicks(1234567);
            properties.Set("Precise", precise);
            double[] doubles = { Math.PI, double.Epsilon, double.MaxValue, -double.MaxValue, 0.1, 1.0 / 3.0 };
            for (int i = 0; i < doubles.Length; i++)
            {
                properties.Set("Double" + i, doubles[i]);
            }
            properties.Set("Negative", -1234567890123L);

            DocxWriter.Write(document, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            RequireContains(ReadPackageEntryText(path, "docProps/custom.xml"),
                "<vt:filetime>2026-10-20T08:30:00.1234567Z</vt:filetime>");

            OfficeDocument loaded = DocxReader.Load(path);
            Require(loaded.Metadata.CustomProperties.GetString("Special") == special,
                "special text custom property mismatch");
            Require(loaded.Metadata.CustomProperties.GetDateTime("Precise") == precise,
                "sub-second precision lost: " + loaded.Metadata.CustomProperties.GetDateTime("Precise").ToString("O"));
            for (int i = 0; i < doubles.Length; i++)
            {
                double actual = loaded.Metadata.CustomProperties.GetDouble("Double" + i);
                Require(BitConverter.DoubleToInt64Bits(actual) == BitConverter.DoubleToInt64Bits(doubles[i]),
                    "double custom property " + i + " not bit-identical: " + actual.ToString("R"));
            }
            Require(loaded.Metadata.CustomProperties.GetInt64("Negative") == -1234567890123L,
                "negative long custom property mismatch");

            string dotxPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-custom-properties-edge.dotx"));
            DocxWriter.WriteTemplate(document, dotxPath);
            OfficeDocument dotx = DocxReader.Load(dotxPath);
            Require(dotx.Metadata.CustomProperties.GetString("Special") == special,
                "DOTX custom property lost");
            Require(dotx.Metadata.CustomProperties.GetInt32("version") == 2,
                "DOTX case-insensitive custom property lookup failed");

            RequireThrows<ArgumentException>(() => properties.Set("Bad", double.NaN),
                "NaN double must be rejected");
            RequireThrows<ArgumentException>(() => properties.Set("Bad", double.PositiveInfinity),
                "infinite double must be rejected");
            RequireThrows<ArgumentException>(() => properties.Set("bad\u0001name", "x"),
                "control character in name must be rejected");
            RequireThrows<ArgumentException>(() => properties.Set("Bad", "x\u0001"),
                "control character in text value must be rejected");
            RequireThrows<ArgumentException>(() => properties.Set("bad\ud800", "x"),
                "unpaired surrogate in name must be rejected");
            RequireThrows<ArgumentException>(() => properties.Set(new string('x', 256), "x"),
                "over-long name must be rejected");
            RequireThrows<ArgumentException>(() => properties.Set("Bad", new object()),
                "unsupported type must be rejected");

            int before = properties.Count;
            properties["Special"] = null;
            Require(properties.Count == before - 1 && !properties.Contains("Special"),
                "null indexer should remove the property");
        }

        /// <summary>验证垂直间隙并入行高、分页边界、空白标签、单列版式、数据行模板与非法版式。Verifies vertical-gap row heights, pagination edges, blank labels, single-column layouts, DataRow templates and invalid layouts.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunLabelEdgeTest(string outputDirectory)
        {
            WordLabelLayout gapLayout = WordLabelLayout.FromMillimeters(
                "gapped", 210, 297, 3, 7, 63.5, 38.1, 15.1, 7, 2.5, 2.5);
            OfficeDocument gapSheet = WordLabels.Create(gapLayout, new[] { new[] { "间隙标签" } });
            Table gapTable = gapSheet.Sections[0].Blocks[0] as Table;
            Require(gapTable != null, "gapped label table missing");
            Require(gapTable.Rows.Count == 7, "vertical gap must be folded into rows, got " + gapTable.Rows.Count);
            double labelHeight = 38.1 * 72 / 25.4;
            double gapHeight = 2.5 * 72 / 25.4;
            RequireNear(gapTable.Rows[0].Height, labelHeight + gapHeight, 0.01,
                "first label row should include the vertical gap");
            RequireNear(gapTable.Rows[6].Height, labelHeight, 0.01,
                "last label row should not include the vertical gap");

            string gapPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-labels-gap.docx"));
            DocxWriter.Write(gapSheet, gapPath);
            ValidateOoxmlPackage(gapPath, "word/document.xml");
            OfficeDocument gapRead = DocxReader.Load(gapPath);
            Table gapReadTable = gapRead.Sections[0].Blocks[0] as Table;
            Require(gapReadTable != null && gapReadTable.Rows.Count == 7,
                "gapped label rows lost in DOCX round trip");

            OfficeDocument exact = WordLabels.Create(WordLabelLayout.AveryL7160(), CreateLabels(21));
            Require(exact.Sections[0].Blocks.Count == 1, "21 labels should fit on one page");
            OfficeDocument twoPages = WordLabels.Create(WordLabelLayout.AveryL7160(), CreateLabels(42));
            Require(twoPages.Sections[0].Blocks.Count == 3, "42 labels should produce two pages");
            Require(((Table)twoPages.Sections[0].Blocks[2]).Rows.Count == 7,
                "second label page should keep the full grid");

            OfficeDocument blank = WordLabels.Create(WordLabelLayout.AveryL7160(), new string[0][]);
            Require(blank.Sections[0].Blocks.Count == 1, "blank label sheet should have one page");
            Table blankTable = blank.Sections[0].Blocks[0] as Table;
            Require(blankTable != null && blankTable.Rows.Count == 7, "blank label grid mismatch");
            Require(TextOf(blankTable.Rows[0].Cells[0]).Length == 0, "blank label should be empty");

            WordLabelLayout singleLayout = WordLabelLayout.FromMillimeters(
                "single", 100, 100, 1, 2, 90, 40, 5, 5, 0, 0);
            OfficeDocument single = WordLabels.Create(singleLayout,
                new[] { new[] { "A" }, new[] { "B" } });
            Table singleTable = single.Sections[0].Blocks[0] as Table;
            Require(singleTable != null && singleTable.Columns.Count == 1 && singleTable.Rows.Count == 2,
                "single-column label layout mismatch");

            WordLabelLayout tooWide = WordLabelLayout.FromMillimeters(
                "wide", 100, 100, 3, 1, 60, 20, 5, 5, 0, 0);
            RequireThrows<ArgumentException>(
                () => { WordLabels.Create(tooWide, new string[0][]); },
                "too-wide label grid must be rejected");
            WordLabelLayout tooTall = WordLabelLayout.FromMillimeters(
                "tall", 100, 100, 1, 3, 60, 40, 5, 5, 0, 0);
            RequireThrows<ArgumentException>(
                () => { WordLabels.Create(tooTall, new string[0][]); },
                "too-tall label grid must be rejected");

            DataTable data = new DataTable();
            data.Columns.Add("Name", typeof(string));
            data.Columns.Add("Amount", typeof(decimal));
            data.Rows.Add("张三", 1234.5m);
            List<DataRow> rows = new List<DataRow>();
            rows.Add(data.Rows[0]);
            OfficeDocument fromRows = WordLabels.CreateFromRecords(
                WordLabelLayout.AveryL7160(), rows, "{{Name}}\n{{Amount:N2}}");
            Table rowTable = fromRows.Sections[0].Blocks[0] as Table;
            Require(TextOf(rowTable.Rows[0].Cells[0])
                == "张三" + Environment.NewLine + "1,234.50", "DataRow label formatting mismatch");

            OfficeDocument keep = WordLabels.CreateFromRecords(
                WordLabelLayout.AveryL7160(),
                new[] { new { Name = "王五" } },
                "{{Missing}}",
                new WordLabelOptions { LeaveUnresolvedTokens = true });
            Require(TextOf(((Table)keep.Sections[0].Blocks[0]).Rows[0].Cells[0]) == "{{Missing}}",
                "unresolved label token should be preserved when requested");
            OfficeDocument drop = WordLabels.CreateFromRecords(
                WordLabelLayout.AveryL7160(),
                new[] { new { Name = "王五" } },
                "{{Missing}}");
            Require(TextOf(((Table)drop.Sections[0].Blocks[0]).Rows[0].Cells[0]).Length == 0,
                "unresolved label token should be removed by default");
        }

        /// <summary>
        /// 验证表单标签纸：2×5 带边框版式、每张标签三行“标签文字 + 填写线”、嵌套表格底对齐与下边框、
        /// 间隙列无边框、DOCX 往返与非法输入。
        /// Verifies form label sheets: a bordered 2×5 layout, three "label text + rule" rows per label,
        /// bottom-aligned nested tables with a bottom border, borderless gap columns, DOCX round trips and
        /// invalid input.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunLabelFormTest(string outputDirectory)
        {
            WordLabelLayout layout = WordLabelLayout.FromMillimeters(
                "学生信息标签 A4 2×5", 210, 297, 2, 5, 82, 48, 25, 22, 4, 4);
            WordLabelOptions options = new WordLabelOptions { ShowCellBorders = true };
            string[] fields = { "姓名：", "年级：", "学号：" };
            OfficeDocument form = WordLabels.CreateForm(layout, fields, options);

            Table table = form.Sections[0].Blocks[0] as Table;
            Require(table != null, "form label table missing");
            Require(table.Columns.Count == 3, "form label table should have 2 label columns and 1 gap, got "
                + table.Columns.Count);
            Require(table.Rows.Count == 5, "form label table should have 5 rows, got " + table.Rows.Count);
            RequireNear(table.Rows[0].Height,
                layout.LabelHeight + layout.VerticalGap, 0.01,
                "form label row height should include the vertical gap");
            RequireNear(table.Rows[4].Height, layout.LabelHeight, 0.01,
                "last form label row should not include the vertical gap");
            RequireNear(table.CellPadding, 0, 0.01,
                "form label page table should keep zero inset so the label box spans the cell");

            TableRow firstRow = table.Rows[0];
            Require(firstRow.Cells.Count == 3, "form label row should have 3 cells, got " + firstRow.Cells.Count);
            TableCell labelCell = firstRow.Cells[0];
            Require(labelCell.Borders == null,
                "form label box should be drawn by the nested table, not the outer cell");
            Require(firstRow.Cells[1].Borders == null,
                "gap cells between label columns must stay borderless");

            Require(labelCell.Blocks.Count == 2, "each label should hold one nested field table plus a tail paragraph, got "
                + labelCell.Blocks.Count);
            Table lines = labelCell.Blocks[0] as Table;
            Require(lines != null, "nested field table missing");
            Require(labelCell.Blocks[1] is Paragraph
                && ((Paragraph)labelCell.Blocks[1]).GetText().Length == 0,
                "label cell should end with an empty tail paragraph");
            double padding = layout.CellPadding;
            double expectedRowHeight = Math.Max(8, options.FontSize * 1.4) + options.FieldSpacing;
            double expectedFillerHeight = layout.LabelHeight - fields.Length * expectedRowHeight;
            Require(lines.Rows.Count == fields.Length + 1,
                "nested table should hold one row per field plus a filler row, got " + lines.Rows.Count);
            RequireNear(lines.CellPadding, padding, 0.01,
                "nested field table should use the layout padding");
            Require(lines.Borders != null && lines.Borders.Top != null && lines.Borders.Bottom != null
                && lines.Borders.Left != null && lines.Borders.Right != null,
                "nested field table should carry the four label box borders");
            RequireNear(lines.Borders.Top.Width, options.BorderWidth, 0.001,
                "form label box border width mismatch");
            Require(lines.Borders.InsideHorizontal.Style == BorderStyle.None
                && lines.Borders.InsideVertical.Style == BorderStyle.None,
                "form label box should not draw internal table lines");
            Require(lines.Columns.Count == 3,
                "nested field table should have text, rule and right inset columns, got " + lines.Columns.Count);
            RequireNear(lines.Columns[0].Width + lines.Columns[1].Width + lines.Columns[2].Width,
                layout.LabelWidth, 0.01,
                "nested field table should span the label box width");
            RequireNear(lines.Columns[2].Width, Math.Max(padding, 6), 0.01,
                "the right inset column should stay at or above the six-point minimum");
            for (int i = 0; i < fields.Length; i++)
            {
                TableRow fieldRow = lines.Rows[i];
                Require(fieldRow.Cells.Count == 3, "form field " + i + " should have text, rule and inset cells");
                RequireNear(fieldRow.Height, expectedRowHeight, 0.01,
                    "form field " + i + " row height mismatch");
                Require(TextOf(fieldRow.Cells[0]) == fields[i],
                    "form field " + i + " label text mismatch: " + TextOf(fieldRow.Cells[0]));

                TableCell labelPart = fieldRow.Cells[0];
                TableCell rulePart = fieldRow.Cells[1];
                Require(labelPart.VerticalAlignment == VerticalAlignment.Bottom
                    && rulePart.VerticalAlignment == VerticalAlignment.Bottom,
                    "form field " + i + " cells should be bottom aligned");
                Require(rulePart.Borders != null && rulePart.Borders.Bottom != null
                    && rulePart.Borders.Top == null && rulePart.Borders.Left == null
                    && rulePart.Borders.Right == null,
                    "form field " + i + " rule cell should only carry a bottom border");
                RequireNear(rulePart.Borders.Bottom.Width, options.BorderWidth, 0.001,
                    "form field " + i + " rule width mismatch");
                Require(fieldRow.Cells[2].Borders == null,
                    "form field " + i + " inset cell should stay borderless");
            }
            TableRow fillerRow = lines.Rows[fields.Length];
            RequireNear(fillerRow.Height, expectedFillerHeight, 0.01,
                "form label filler row should complete the label height");
            Require(fillerRow.Cells.Count == 3, "form label filler row should keep three cells");
            Require(fillerRow.Cells[0].Borders == null && fillerRow.Cells[1].Borders == null
                && fillerRow.Cells[2].Borders == null,
                "form label filler cells should not draw borders");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-form-labels.docx"));
            DocxWriter.Write(form, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            string documentXml = ReadPackageEntryText(path, "word/document.xml");
            RequireContains(documentXml, "姓名：");
            RequireContains(documentXml, "<w:vAlign w:val=\"bottom\"");
            OfficeDocument loaded = DocxReader.Load(path);
            Table loadedTable = loaded.Sections[0].Blocks[0] as Table;
            Require(loadedTable != null && loadedTable.Rows.Count == 5
                && loadedTable.Rows[0].Cells.Count == 3,
                "form label grid lost in DOCX round trip");
            TableCell loadedCell = loadedTable.Rows[0].Cells[0];
            Require(loadedCell.Borders == null,
                "form label box should stay on the nested table through DOCX round trip");
            Require(loadedCell.Blocks.Count == 2,
                "nested field table and tail paragraph lost in DOCX round trip");
            Table loadedLines = loadedCell.Blocks[0] as Table;
            Require(loadedLines != null && loadedLines.Borders != null
                && loadedLines.Borders.Top != null && loadedLines.Borders.Bottom != null
                && loadedLines.Borders.Left != null && loadedLines.Borders.Right != null,
                "form label box border lost in DOCX round trip");
            Require(loadedLines.Rows.Count == fields.Length + 1,
                "nested field rows or filler lost in DOCX round trip");
            for (int i = 0; i < fields.Length; i++)
            {
                TableRow loadedRow = loadedLines.Rows[i];
                Require(loadedRow.Cells.Count == 3,
                    "nested form field " + i + " lost in DOCX round trip");
                Require(TextOf(loadedRow.Cells[0]) == fields[i],
                    "nested form field " + i + " text lost in DOCX round trip");
                Require(loadedRow.Cells[1].Borders != null
                    && loadedRow.Cells[1].Borders.Bottom != null,
                    "nested form field " + i + " rule lost in DOCX round trip");
            }

            OfficeDocument plain = WordLabels.CreateForm(layout, new[] { "姓名：", " " });
            Table plainTable = plain.Sections[0].Blocks[0] as Table;
            Require(plainTable.Rows[0].Cells[0].Borders == null,
                "form labels should stay borderless by default");
            Table plainLines = plainTable.Rows[0].Cells[0].Blocks[0] as Table;
            Require(plainLines != null && plainLines.Rows.Count == 2,
                "blank field labels should be dropped, leaving one field row plus the filler");
            Require(plainLines.Borders != null && plainLines.Borders.Top.Style == BorderStyle.None,
                "form labels should not draw a label box by default");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-form-labels.pdf"));
            OfficeConverter.ToPdf(form, pdfPath);
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            Require(pdf.PageCount == 1, "form labels should fit on one A4 page, got " + pdf.PageCount);
            System.Text.StringBuilder pdfText = new System.Text.StringBuilder();
            IReadOnlyList<PdfTextLine> pdfLines = pdf.Pages[0].GetTextLines();
            List<double> nameXs = new List<double>();
            List<double> nameYs = new List<double>();
            for (int i = 0; i < pdfLines.Count; i++)
            {
                pdfText.Append(pdfLines[i].Text);
                for (int f = 0; f < pdfLines[i].Fragments.Count; f++)
                {
                    PdfTextFragment fragment = pdfLines[i].Fragments[f];
                    if (!fragment.Text.Contains("姓名：")) continue;
                    nameXs.Add(Math.Round(fragment.X, 1));
                    nameYs.Add(Math.Round(pdfLines[i].Y, 1));
                }
            }
            Require(pdfText.ToString().Contains("姓名：") && pdfText.ToString().Contains("年级：")
                && pdfText.ToString().Contains("学号："),
                "form label text missing from the rendered PDF");
            Require(nameXs.Count == layout.LabelsPerPage,
                "rendered PDF should contain " + layout.LabelsPerPage + " 姓名 rows, got " + nameXs.Count);
            Require(new HashSet<double>(nameXs).Count == layout.Columns,
                "rendered PDF should arrange 姓名 rows in " + layout.Columns + " columns");
            Require(new HashSet<double>(nameYs).Count == layout.Rows,
                "rendered PDF should arrange 姓名 rows in " + layout.Rows + " rows");

            RequireThrows<ArgumentNullException>(() => WordLabels.CreateForm(null, fields),
                "null layout must be rejected");
            RequireThrows<ArgumentNullException>(() => WordLabels.CreateForm(layout, null),
                "null field labels must be rejected");
            RequireThrows<ArgumentException>(() => WordLabels.CreateForm(layout, new[] { " ", "" }),
                "blank field labels must be rejected");
        }

        /// <summary>
        /// 验证留样标签纸：A4 4×10、每张四条文字、外部框线与三条行分隔线、DOCX 往返与 PDF 网格。
        /// Verifies the retention label sheet: A4 4×10, four text rows per label, an outer box with three
        /// row separators, DOCX round trips and the rendered PDF grid.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunRetentionLabelSheetTest(string outputDirectory)
        {
            WordLabelLayout layout = WordLabelLayout.FromMillimeters(
                "留样标签 A4 4×10", 210, 297, 4, 10, 45, 25, 12, 10.5, 3, 2);
            WordLabelOptions options = new WordLabelOptions
            {
                FontFamily = "Arial",
                EastAsiaFontFamily = "SimSun",
                FontSize = 9,
                ShowCellBorders = true,
                ShowRowSeparators = true,
                BorderWidth = 0.75
            };
            string[] template =
            {
                "201  年  月  日  点  分",
                "食品名：          （    克）",
                "留样人：",
                "审核人："
            };
            List<string[]> contents = new List<string[]>();
            for (int i = 0; i < layout.LabelsPerPage; i++) contents.Add(template);
            OfficeDocument sheet = WordLabels.Create(layout, contents, options);

            Section section = sheet.Sections[0];
            RequireNear(section.PageSetup.MarginBottom,
                layout.CreatePageSetup().MarginBottom - 18, 0.01,
                "row separator labels should reserve space for Word's trailing cell paragraphs");
            Table table = section.Blocks[0] as Table;
            Require(table != null, "retention label table missing");
            Require(table.Columns.Count == 7,
                "retention sheet should have four label columns and three gaps, got " + table.Columns.Count);
            Require(table.Rows.Count == 10, "retention sheet should have ten rows, got " + table.Rows.Count);

            TableCell firstCell = table.Rows[0].Cells[0];
            Require(firstCell.Borders == null,
                "row separator labels should draw the box on the nested table, not the outer cell");
            Require(firstCell.Blocks.Count == 2,
                "each retention label should hold one nested table plus a tail paragraph, got "
                + firstCell.Blocks.Count);
            Table body = firstCell.Blocks[0] as Table;
            Require(body != null, "retention label nested table missing");
            Require(body.Columns.Count == 1, "retention label body should be one column");
            Require(body.Rows.Count == template.Length,
                "retention label body should keep four rows, got " + body.Rows.Count);
            RequireNear(body.Rows[0].Height, layout.LabelHeight / template.Length, 0.01,
                "retention label rows should divide the label height equally");
            Require(body.Borders != null && body.Borders.Top != null
                && body.Borders.Top.Style == BorderStyle.Single
                && body.Borders.Bottom.Style == BorderStyle.Single
                && body.Borders.Left.Style == BorderStyle.Single
                && body.Borders.Right.Style == BorderStyle.Single,
                "retention label should draw a complete outer box");
            Require(body.Borders.InsideHorizontal != null
                && body.Borders.InsideHorizontal.Style == BorderStyle.Single,
                "retention label should draw row separators");
            Require(body.Borders.InsideVertical != null
                && body.Borders.InsideVertical.Style == BorderStyle.None,
                "retention label should not draw a vertical separator");
            RequireNear(body.Borders.Top.Width, options.BorderWidth, 0.001,
                "retention label border width mismatch");
            Require(TextOf(body.Rows[1].Cells[0]).StartsWith("食品名："),
                "retention label second row text mismatch");
            Require(TextOf(body.Rows[2].Cells[0]) == "留样人：",
                "retention label third row text mismatch");
            Require(TextOf(body.Rows[3].Cells[0]) == "审核人：",
                "retention label fourth row text mismatch");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-retention-labels.docx"));
            DocxWriter.Write(sheet, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            string documentXml = ReadPackageEntryText(path, "word/document.xml");
            RequireContains(documentXml, "留样人：");
            RequireContains(documentXml, "<w:insideH w:val=\"single\"");
            OfficeDocument loaded = DocxReader.Load(path);
            Table loadedTable = loaded.Sections[0].Blocks[0] as Table;
            Require(loadedTable != null && loadedTable.Rows.Count == 10,
                "retention label grid lost in DOCX round trip");
            TableCell loadedCell = loadedTable.Rows[0].Cells[0];
            Require(loadedCell.Blocks.Count == 2,
                "retention nested table or tail paragraph lost in DOCX round trip");
            Table loadedBody = loadedCell.Blocks[0] as Table;
            Require(loadedBody != null && loadedBody.Rows.Count == template.Length
                && loadedBody.Rows[0].Cells.Count == 1,
                "retention nested rows lost in DOCX round trip");
            Require(loadedBody.Borders != null && loadedBody.Borders.InsideHorizontal != null
                && loadedBody.Borders.InsideHorizontal.Style == BorderStyle.Single,
                "retention row separators lost in DOCX round trip");
            Require(TextOf(loadedBody.Rows[2].Cells[0]) == "留样人：",
                "retention label text lost in DOCX round trip");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-retention-labels.pdf"));
            OfficeConverter.ToPdf(sheet, pdfPath);
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            Require(pdf.PageCount == 1,
                "retention labels should fit on one A4 page, got " + pdf.PageCount);
            IReadOnlyList<PdfTextLine> pdfLines = pdf.Pages[0].GetTextLines();
            List<double> labelXs = new List<double>();
            List<double> labelYs = new List<double>();
            for (int i = 0; i < pdfLines.Count; i++)
            {
                for (int f = 0; f < pdfLines[i].Fragments.Count; f++)
                {
                    if (!pdfLines[i].Fragments[f].Text.Contains("留样人：")) continue;
                    labelXs.Add(Math.Round(pdfLines[i].Fragments[f].X, 1));
                    labelYs.Add(Math.Round(pdfLines[i].Y, 1));
                }
            }
            Require(labelXs.Count == layout.LabelsPerPage,
                "rendered PDF should contain " + layout.LabelsPerPage + " 留样人 rows, got " + labelXs.Count);
            Require(new HashSet<double>(labelXs).Count == layout.Columns,
                "rendered PDF should arrange retention labels in " + layout.Columns + " columns");
            Require(new HashSet<double>(labelYs).Count == layout.Rows,
                "rendered PDF should arrange retention labels in " + layout.Rows + " rows");

            PdfFile file = PdfFile.Load(pdfPath);
            List<double[]> segments = ExtractHorizontalPdfSegments(file.GetPageContentBytes(0));
            Require(segments.Count >= layout.LabelsPerPage * template.Length * 2,
                "rendered PDF should draw every retention label edge, got " + segments.Count);
            HashSet<double> lineYs = new HashSet<double>();
            HashSet<string> lineSpans = new HashSet<string>();
            for (int i = 0; i < segments.Count; i++)
            {
                lineYs.Add(Math.Round(segments[i][1], 1));
                lineSpans.Add(Math.Round(segments[i][0], 1).ToString() + ":"
                    + Math.Round(segments[i][2], 1).ToString());
            }
            Require(lineYs.Count == layout.Rows * (template.Length + 1),
                "rendered PDF should draw five distinct lines per retention label row, got " + lineYs.Count);
            Require(lineSpans.Count == layout.Columns,
                "rendered PDF should draw retention lines at " + layout.Columns + " horizontal spans, got "
                + lineSpans.Count);
        }

        /// <summary>
        /// 验证化学品安全标签：A5 横向、三列合并单元格、GHS 图形、多段落文字、底部双行联系表、
        /// DOCX 往返与 PDF 渲染。
        /// Verifies the chemical safety label: A5 landscape, three columns with merged cells, the GHS
        /// pictogram, multi-paragraph text, a two-row contact table at the bottom, DOCX round trips and
        /// PDF rendering.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunChemicalSafetyLabelTest(string outputDirectory)
        {
            OfficeDocument document = CreateChemicalSafetyDocument();
            Section section = document.Sections[0];
            Require(section.PageSetup.Width > section.PageSetup.Height,
                "chemical safety label should be landscape");
            RequireNear(section.PageSetup.Width, 210.0 * 72 / 25.4, 0.01,
                "chemical safety label width mismatch");

            Table table = section.Blocks[0] as Table;
            Require(table != null && table.Rows.Count == 6,
                "chemical safety table should have six rows, got "
                + (table == null ? 0 : table.Rows.Count));
            Require(table.Columns.Count == 3, "chemical safety table should have three columns");
            Require(ContainsText(table.Rows[0].Cells[0], "纯度：") && ContainsText(table.Rows[0].Cells[0], "≥30%"),
                "purity cell text missing");
            Require(ContainsText(table.Rows[0].Cells[1], "hydrochloric acid")
                && ContainsText(table.Rows[0].Cells[1], "盐酸")
                && ContainsText(table.Rows[0].Cells[1], "HCl"),
                "chemical title text missing");
            Require(ContainsText(table.Rows[1].Cells[0], "酸性腐蚀，对人体有害"),
                "danger statement missing");
            Require(ContainsText(table.Rows[2].Cells[1], "安全措施：")
                && ContainsText(table.Rows[2].Cells[1], "皮肤接触")
                && ContainsText(table.Rows[2].Cells[1], "灭火方法："),
                "safety guidance text missing");
            Require(ContainsText(table.Rows[4].Cells[0], "生产企业：南宁化工股份有限公司")
                && ContainsText(table.Rows[4].Cells[0], "0771-4818653"),
                "manufacturer contact row missing");
            Require(ContainsText(table.Rows[4].Cells[1], "UN NO.:1789")
                && ContainsText(table.Rows[5].Cells[1], "CN NO.:81013"),
                "UN/CN numbers missing");

            TableCell purityCell = table.Rows[0].Cells[0];
            Require(purityCell.RowSpan == 2, "purity cell should span two rows");
            TableCell pictogramCell = table.Rows[0].Cells[2];
            Require(pictogramCell.RowSpan == 4, "pictogram cell should span four rows");
            Require(document.Sections[0].Blocks.Count == 1, "chemical safety table should be the only body block");
            Require(table.Rows[1].Cells.Count == 1,
                "danger row should only hold the center cell, got " + table.Rows[1].Cells.Count);
            Require(table.Rows[2].Cells[0].RowSpan == 2, "storage cell should span two rows");
            Require(table.Rows[4].Cells[0].ColSpan == 2,
                "manufacturer row should merge the left and center columns");
            Require(table.Rows[5].Cells[0].ColSpan == 2,
                "emergency row should merge the left and center columns");
            Table contact = table.Rows[4].Cells[0].Blocks[0] as Table;
            Require(contact != null && contact.Columns.Count == 2 && contact.Rows.Count == 1,
                "manufacturer row should hold a two-column nested contact table");
            Paragraph imageParagraph = pictogramCell.Blocks[0] as Paragraph;
            Require(imageParagraph != null && imageParagraph.Inlines.Count == 1
                && imageParagraph.Inlines[0] is ImageRun,
                "pictogram image missing from the safety label");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-chemical-safety.docx"));
            DocxWriter.Write(document, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            string documentXml = ReadPackageEntryText(path, "word/document.xml");
            RequireContains(documentXml, "盐酸");
            RequireContains(documentXml, "<w:vMerge w:val=\"restart\"");
            RequireContains(documentXml, "<w:gridSpan w:val=\"2\"");
            OfficeDocument loaded = DocxReader.Load(path);
            Table loadedTable = loaded.Sections[0].Blocks[0] as Table;
            Require(loadedTable != null && loadedTable.Rows.Count == 6,
                "chemical safety table lost in DOCX round trip");
            Require(loadedTable.Rows[0].Cells[0].RowSpan == 2
                && loadedTable.Rows[0].Cells[2].RowSpan == 4,
                "vertical merges lost in DOCX round trip");
            Require(loadedTable.Rows[2].Cells[0].RowSpan == 2,
                "storage row span lost in DOCX round trip");
            Require(loadedTable.Rows[4].Cells[0].ColSpan == 2
                && loadedTable.Rows[5].Cells[0].ColSpan == 2,
                "bottom row merges lost in DOCX round trip");
            Require(loaded.GetPlainText().Contains("氢氧化钠"),
                "safety text lost in DOCX round trip");
            Table loadedContact = loadedTable.Rows[4].Cells[0].Blocks[0] as Table;
            Require(loadedContact != null && loadedContact.Columns.Count == 2,
                "nested contact table lost in DOCX round trip");
            bool imageReloaded = false;
            for (int b = 0; b < loadedTable.Rows[0].Cells[2].Blocks.Count; b++)
            {
                Paragraph paragraph = loadedTable.Rows[0].Cells[2].Blocks[b] as Paragraph;
                if (paragraph == null) continue;
                for (int i = 0; i < paragraph.Inlines.Count; i++)
                {
                    if (paragraph.Inlines[i] is ImageRun) imageReloaded = true;
                }
            }
            Require(imageReloaded, "pictogram image lost in DOCX round trip");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-chemical-safety.pdf"));
            OfficeConverter.ToPdf(document, pdfPath);
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            Require(pdf.PageCount == 1,
                "chemical safety label should fit on one page, got " + pdf.PageCount);
            System.Text.StringBuilder pdfText = new System.Text.StringBuilder();
            IReadOnlyList<PdfTextLine> pdfLines = pdf.Pages[0].GetTextLines();
            for (int i = 0; i < pdfLines.Count; i++) pdfText.Append(pdfLines[i].Text);
            Require(pdfText.ToString().Contains("盐酸")
                && pdfText.ToString().Contains("氢氧化钠")
                && pdfText.ToString().Contains("UN NO.:1789"),
                "chemical safety text missing from the rendered PDF");
            for (int i = 0; i < pdfLines.Count; i++)
            {
                string lineText = pdfLines[i].Text.TrimStart();
                if (lineText.Length == 0) continue;
                Require(lineText[0] != '。' && lineText[0] != '，' && lineText[0] != '、'
                    && lineText[0] != '；' && lineText[0] != '：',
                    "rendered PDF line should not start with closing punctuation: " + lineText);
            }
            PdfFile file = PdfFile.Load(pdfPath);
            double pageWidth;
            double pageHeight;
            file.GetPageSize(0, out pageWidth, out pageHeight);
            RequireNear(pageWidth, 210.0 * 72 / 25.4, 0.5,
                "rendered PDF should keep the A5 landscape width");
            Require(pageHeight < pageWidth, "rendered PDF should keep the landscape orientation");
            string content = System.Text.Encoding.ASCII.GetString(file.GetPageContentBytes(0));
            Require(content.Contains(" Do"),
                "rendered PDF should place the GHS pictogram image");
        }

        /// <summary>
        /// 构建化学品安全标签示例：A5 横向、三列六行、纵向/横向合并、GHS 图形与底部联系表。
        /// Builds the chemical safety label sample: A5 landscape, three columns by six rows, vertical and
        /// horizontal merges, a GHS pictogram and a contact table at the bottom.
        /// </summary>
        /// <returns>化学品安全标签文档。The chemical safety label document.</returns>
        private static OfficeDocument CreateChemicalSafetyDocument()
        {
            PageSetup setup = new PageSetup();
            setup.Width = 210.0 * 72 / 25.4;
            setup.Height = 148.0 * 72 / 25.4;
            setup.MarginLeft = 6.0 * 72 / 25.4;
            setup.MarginRight = setup.MarginLeft;
            setup.MarginTop = 6.0 * 72 / 25.4;
            setup.MarginBottom = setup.MarginTop;
            setup.HeaderDistance = 0;
            setup.FooterDistance = 0;

            OfficeDocument document = new OfficeDocument();
            document.DefaultTextFormat = SafetyFormat(8.5, false);
            Section section = document.AddSection(setup);

            Table table = new Table();
            table.Borders = SafetyBorders();
            table.CellPadding = 3;
            table.Autofit = false;
            table.Alignment = TableAlignment.Left;
            table.RepeatHeaderRow = false;
            table.Columns.Add(new TableColumn(38.0 * 72 / 25.4));
            table.Columns.Add(new TableColumn(120.0 * 72 / 25.4));
            table.Columns.Add(new TableColumn(40.0 * 72 / 25.4));
            section.Blocks.Add(table);

            TableRow titleRow = table.AddRow();
            titleRow.Height = 26.0 * 72 / 25.4;
            TableCell purity = new TableCell();
            purity.RowSpan = 2;
            purity.VerticalAlignment = VerticalAlignment.Middle;
            AddCenteredParagraph(purity, "纯度：", SafetyFormat(8.5, false), 0, 0);
            AddCenteredParagraph(purity, "≥30%", SafetyFormat(11, true), 1, 0);
            titleRow.Cells.Add(purity);

            TableCell title = new TableCell();
            title.VerticalAlignment = VerticalAlignment.Middle;
            AddCenteredParagraph(title, "hydrochloric acid", SafetyFormat(10.5, true), 0, 0);
            AddCenteredParagraph(title, "盐酸", SafetyFormat(16, true), 0, 0);
            AddCenteredParagraph(title, "HCl", SafetyFormat(10, true), 0, 0);
            titleRow.Cells.Add(title);

            TableCell pictogram = new TableCell();
            pictogram.RowSpan = 4;
            pictogram.VerticalAlignment = VerticalAlignment.Middle;
            Paragraph imageParagraph = new Paragraph();
            imageParagraph.Alignment = ParagraphAlignment.Center;
            imageParagraph.AddImage(new OfficeImage(CreateHazardPng(320)), 34.0 * 72 / 25.4, 34.0 * 72 / 25.4);
            pictogram.Blocks.Add(imageParagraph);
            titleRow.Cells.Add(pictogram);

            TableRow dangerRow = table.AddRow();
            dangerRow.Height = 14.0 * 72 / 25.4;
            TableCell danger = new TableCell();
            danger.VerticalAlignment = VerticalAlignment.Middle;
            AddCenteredParagraph(danger, "危  险", SafetyFormat(12, true), 0, 1);
            AddCenteredParagraph(danger, "酸性腐蚀，对人体有害", SafetyFormat(9, true), 0, 0);
            dangerRow.Cells.Add(danger);

            TableRow safetyRow = table.AddRow();
            safetyRow.Height = 55.0 * 72 / 25.4;
            TableCell storage = new TableCell();
            storage.RowSpan = 2;
            storage.VerticalAlignment = VerticalAlignment.Middle;
            AddCenteredParagraph(storage, "最大储存量：", SafetyFormat(8.5, false), 0, 0);
            AddCenteredParagraph(storage, "60 吨", SafetyFormat(11, true), 1, 0);
            safetyRow.Cells.Add(storage);

            TableCell safety = new TableCell();
            safety.VerticalAlignment = VerticalAlignment.Top;
            AddLeftParagraph(safety, "安全措施：", SafetyFormat(9, true), 0, 1);
            AddSafetyBullet(safety,
                "◆ 皮肤接触：立即脱去污染的衣着，用大量流动清水冲洗至少 15 分钟，再用 0.5%碳酸氢钠溶液清洗。就医。");
            AddSafetyBullet(safety,
                "◆ 眼睛接触：立即提起眼睑，用大量流动清水或生理盐水彻底冲洗至少 15 分钟，再用 0.5%碳酸氢钠溶液清洗。就医。");
            AddSafetyBullet(safety,
                "◆ 吸入：保持呼吸道通畅。如呼吸困难，给输氧。如呼吸停止，立即进行人工呼吸。就医。");
            AddSafetyBullet(safety,
                "◆ 泄漏应急处理：泄漏后，穿防酸碱工作服、戴防护眼镜。少量泄漏：用砂土、干燥石灰或苏打灰混合或大量水冲洗，冲洗水放入废水系统处理。大量泄漏：构筑围堤或挖坑收容；用泵转移废液并回收处置。");
            AddLeftParagraph(safety, "灭火方法：", SafetyFormat(9, true), 3, 0);
            AddLeftParagraph(safety,
                "用碱性物质如碳酸氢钠、碳酸钠、消石灰、氢氧化钠等中和。", SafetyFormat(8.5, false), 0, 0);
            ((Paragraph)safety.Blocks[safety.Blocks.Count - 1]).LeftIndent = 11;
            safetyRow.Cells.Add(safety);

            TableRow noticeRow = table.AddRow();
            noticeRow.Height = 8.0 * 72 / 25.4;
            TableCell notice = new TableCell();
            notice.VerticalAlignment = VerticalAlignment.Middle;
            AddCenteredParagraph(notice, "请参阅安全技术说明书", SafetyFormat(8.5, false), 0, 0);
            noticeRow.Cells.Add(notice);

            TableRow manufacturerRow = table.AddRow();
            manufacturerRow.Height = 8.0 * 72 / 25.4;
            TableCell manufacturer = new TableCell();
            manufacturer.ColSpan = 2;
            manufacturer.VerticalAlignment = VerticalAlignment.Middle;
            manufacturer.Blocks.Add(CreateContactRow(
                "生产企业：南宁化工股份有限公司", "应急电话：0771-4818653"));
            AddTailParagraph(manufacturer);
            manufacturerRow.Cells.Add(manufacturer);
            TableCell unNumber = new TableCell();
            unNumber.VerticalAlignment = VerticalAlignment.Middle;
            AddCenteredParagraph(unNumber, "UN NO.:1789", SafetyFormat(8, false), 0, 0);
            manufacturerRow.Cells.Add(unNumber);

            TableRow emergencyRow = table.AddRow();
            emergencyRow.Height = 8.0 * 72 / 25.4;
            TableCell emergency = new TableCell();
            emergency.ColSpan = 2;
            emergency.VerticalAlignment = VerticalAlignment.Middle;
            emergency.Blocks.Add(CreateContactRow(
                "应急电话：火警：3668119", "值班：3668181　化学值班：3668311"));
            AddTailParagraph(emergency);
            emergencyRow.Cells.Add(emergency);
            TableCell cnNumber = new TableCell();
            cnNumber.VerticalAlignment = VerticalAlignment.Middle;
            AddCenteredParagraph(cnNumber, "CN NO.:81013", SafetyFormat(8, false), 0, 0);
            emergencyRow.Cells.Add(cnNumber);

            return document;
        }

        /// <summary>创建化学品安全标签的黑色细边框集合。Creates the thin black border set of the chemical safety label.</summary>
        /// <returns>六向单实线边框。Six single solid border lines.</returns>
        private static TableBorders SafetyBorders()
        {
            BorderLine line = new BorderLine(BorderStyle.Single, 1.0, OfficeColor.Black);
            TableBorders borders = new TableBorders();
            borders.Top = line;
            borders.Bottom = line;
            borders.Left = line;
            borders.Right = line;
            borders.InsideHorizontal = line;
            borders.InsideVertical = line;
            return borders;
        }

        /// <summary>创建标签文字格式：Arial 西文、宋体中文、黑色。Creates a label text format: Arial Latin, SimSun East Asian, black.</summary>
        /// <param name="fontSize">字号（磅）。The font size in points.</param>
        /// <param name="bold">是否加粗。Whether the text is bold.</param>
        /// <returns>文字格式。The text format.</returns>
        private static TextFormat SafetyFormat(double fontSize, bool bold)
        {
            TextFormat format = new TextFormat("Arial", fontSize);
            format.EastAsiaFontFamily = "SimSun";
            format.Bold = bold;
            return format;
        }

        /// <summary>向单元格追加居中段落。Appends a centered paragraph to a cell.</summary>
        /// <param name="cell">目标单元格。The target cell.</param>
        /// <param name="text">段落文本。The paragraph text.</param>
        /// <param name="format">文字格式。The text format.</param>
        /// <param name="spaceBefore">段前间距。The space before the paragraph.</param>
        /// <param name="spaceAfter">段后间距。The space after the paragraph.</param>
        private static void AddCenteredParagraph(
            TableCell cell, string text, TextFormat format, double spaceBefore, double spaceAfter)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.Alignment = ParagraphAlignment.Center;
            paragraph.SpaceBefore = spaceBefore;
            paragraph.SpaceAfter = spaceAfter;
            paragraph.AddText(text, format);
            cell.Blocks.Add(paragraph);
        }

        /// <summary>向单元格追加左对齐段落。Appends a left-aligned paragraph to a cell.</summary>
        /// <param name="cell">目标单元格。The target cell.</param>
        /// <param name="text">段落文本。The paragraph text.</param>
        /// <param name="format">文字格式。The text format.</param>
        /// <param name="spaceBefore">段前间距。The space before the paragraph.</param>
        /// <param name="spaceAfter">段后间距。The space after the paragraph.</param>
        private static void AddLeftParagraph(
            TableCell cell, string text, TextFormat format, double spaceBefore, double spaceAfter)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.Alignment = ParagraphAlignment.Left;
            paragraph.SpaceBefore = spaceBefore;
            paragraph.SpaceAfter = spaceAfter;
            paragraph.LineSpacing = 1.1;
            paragraph.AddText(text, format);
            cell.Blocks.Add(paragraph);
        }

        /// <summary>追加带悬挂缩进的安全措施条目。Appends a safety bullet with a hanging indent.</summary>
        /// <param name="cell">目标单元格。The target cell.</param>
        /// <param name="text">条目文本。The bullet text.</param>
        private static void AddSafetyBullet(TableCell cell, string text)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.Alignment = ParagraphAlignment.Left;
            paragraph.LineSpacing = 1.1;
            paragraph.SpaceAfter = 2;
            paragraph.LeftIndent = 11;
            paragraph.FirstLineIndent = -11;
            paragraph.AddText(text, SafetyFormat(8.5, false));
            cell.Blocks.Add(paragraph);
        }

        /// <summary>创建底部联系信息的无边框两列表格。Creates the borderless two-column contact table used at the bottom.</summary>
        /// <param name="left">左列文本。The left column text.</param>
        /// <param name="right">右列文本。The right column text.</param>
        /// <returns>联系信息表格。The contact table.</returns>
        private static Table CreateContactRow(string left, string right)
        {
            Table table = new Table();
            table.Borders = TableBorders.None();
            table.CellPadding = 0;
            table.Autofit = false;
            table.Alignment = TableAlignment.Left;
            table.RepeatHeaderRow = false;
            table.Columns.Add(new TableColumn(92.0 * 72 / 25.4));
            table.Columns.Add(new TableColumn(62.0 * 72 / 25.4));
            TableRow row = table.AddRow();
            TableCell leftCell = new TableCell();
            leftCell.VerticalAlignment = VerticalAlignment.Middle;
            AddLeftParagraph(leftCell, left, SafetyFormat(8.5, false), 0, 0);
            row.Cells.Add(leftCell);
            TableCell rightCell = new TableCell();
            rightCell.VerticalAlignment = VerticalAlignment.Middle;
            Paragraph paragraph = new Paragraph();
            paragraph.Alignment = ParagraphAlignment.Right;
            paragraph.AddText(right, SafetyFormat(8.5, false));
            rightCell.Blocks.Add(paragraph);
            row.Cells.Add(rightCell);
            return table;
        }

        /// <summary>在嵌套表格后补一个极小空段落，避免 Word 自动补段撑高行。Adds a minimal tail paragraph after a nested table so Word's own trailing paragraph cannot grow the row.</summary>
        /// <param name="cell">目标单元格。The target cell.</param>
        private static void AddTailParagraph(TableCell cell)
        {
            Paragraph tail = new Paragraph();
            tail.LineSpacing = 0.1;
            tail.SpaceBefore = 0;
            tail.SpaceAfter = 0;
            cell.Blocks.Add(tail);
        }

        /// <summary>判断单元格文本（含嵌套表格）是否包含指定片段。Checks whether a cell's text, including nested tables, contains the given fragment.</summary>
        /// <param name="cell">目标单元格。The target cell.</param>
        /// <param name="fragment">待查找文本。The text to find.</param>
        /// <returns>包含时返回 true。True when the fragment is present.</returns>
        private static bool ContainsText(TableCell cell, string fragment)
        {
            return BlocksContainText(cell.Blocks, fragment);
        }

        /// <summary>递归判断块列表中的段落或嵌套表格是否包含指定片段。Recursively checks paragraphs and nested tables in a block list for the given fragment.</summary>
        /// <param name="blocks">块列表。The block list.</param>
        /// <param name="fragment">待查找文本。The text to find.</param>
        /// <returns>包含时返回 true。True when the fragment is present.</returns>
        private static bool BlocksContainText(List<Block> blocks, string fragment)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null
                    && paragraph.GetText().IndexOf(fragment, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
                Table table = blocks[i] as Table;
                if (table != null && TableContainsText(table, fragment)) return true;
            }
            return false;
        }

        /// <summary>递归判断表格是否包含指定片段。Recursively checks whether a table contains the given fragment.</summary>
        /// <param name="table">表格。The table.</param>
        /// <param name="fragment">待查找文本。The text to find.</param>
        /// <returns>包含时返回 true。True when the fragment is present.</returns>
        private static bool TableContainsText(Table table, string fragment)
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                {
                    if (BlocksContainText(table.Rows[r].Cells[c].Blocks, fragment)) return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 从 PDF 内容流中提取水平线段，返回起点与终点坐标。
        /// Extracts horizontal line segments from a PDF content stream, returning their endpoint coordinates.
        /// </summary>
        /// <param name="content">解码后的内容流。The decoded content stream.</param>
        /// <returns>每条水平线段的 x1、y1、x2、y2。The x1, y1, x2 and y2 of each horizontal segment.</returns>
        private static List<double[]> ExtractHorizontalPdfSegments(byte[] content)
        {
            List<double[]> segments = new List<double[]>();
            if (content == null) return segments;
            string text = System.Text.Encoding.ASCII.GetString(content);
            MatchCollection matches = Regex.Matches(text,
                @"(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+m\s+(-?\d+(?:\.\d+)?)\s+(-?\d+(?:\.\d+)?)\s+l\s+S");
            for (int i = 0; i < matches.Count; i++)
            {
                Match match = matches[i];
                double x1 = double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
                double y1 = double.Parse(match.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
                double x2 = double.Parse(match.Groups[3].Value, System.Globalization.CultureInfo.InvariantCulture);
                double y2 = double.Parse(match.Groups[4].Value, System.Globalization.CultureInfo.InvariantCulture);
                if (Math.Abs(y1 - y2) > 0.01) continue;
                segments.Add(new[] { x1, y1, x2, y2 });
            }
            return segments;
        }

        /// <summary>
        /// 生成黑白的 GHS 腐蚀图形 PNG：菱形边框、倾斜的试管、滴落的液体、手与台面轮廓。
        /// Generates a black-and-white GHS corrosion pictogram PNG: diamond frame, tilted test tube,
        /// dripping liquid, hand and surface silhouettes.
        /// </summary>
        /// <param name="size">图像边长（像素）。The image side length in pixels.</param>
        /// <returns>PNG 字节。The PNG bytes.</returns>
        private static byte[] CreateHazardPng(int size)
        {
            int stride = size * 4 + 1;
            byte[] raw = new byte[stride * size];
            for (int y = 0; y < size; y++)
            {
                int row = y * stride;
                raw[row] = 0;
                for (int x = 0; x < size; x++)
                {
                    int offset = row + 1 + x * 4;
                    byte value = HazardInk(x, y, size) ? (byte)0 : (byte)255;
                    raw[offset] = value;
                    raw[offset + 1] = value;
                    raw[offset + 2] = value;
                    raw[offset + 3] = 255;
                }
            }

            using (MemoryStream stream = new MemoryStream())
            {
                byte[] signature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
                stream.Write(signature, 0, signature.Length);
                using (MemoryStream header = new MemoryStream())
                {
                    WriteBigEndian(header, size);
                    WriteBigEndian(header, size);
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

        /// <summary>判断像素是否属于 GHS 图形墨色区域。Determines whether a pixel belongs to the GHS pictogram ink.</summary>
        /// <param name="x">像素 X。The pixel X.</param>
        /// <param name="y">像素 Y。The pixel Y.</param>
        /// <param name="size">图像边长。The image side length.</param>
        /// <returns>属于墨色返回 true。True when the pixel is ink.</returns>
        private static bool HazardInk(int x, int y, int size)
        {
            double cx = size / 2.0;
            double cy = size / 2.0;
            double distance = Math.Abs(x - cx) + Math.Abs(y - cy);
            double outer = size * 0.46;
            double inner = outer - size * 0.035;
            if (distance <= outer && distance >= inner) return true;
            if (distance > inner) return false;

            double u = x / (double)size;
            double v = y / (double)size;

            if (v >= 0.64 && v <= 0.68 && u >= 0.30 && u <= 0.70) return true;
            if (u >= 0.43 && u <= 0.64 && v >= 0.52 && v <= 0.64) return true;
            if (v >= 0.40 && v <= 0.54)
            {
                if ((u >= 0.45 && u <= 0.48) || (u >= 0.49 && u <= 0.52)
                    || (u >= 0.53 && u <= 0.56) || (u >= 0.57 && u <= 0.60)) return true;
            }
            if (u >= 0.39 && u <= 0.44 && v >= 0.47 && v <= 0.54) return true;
            if (SegmentInk(x, y, size, 0.28, 0.30, 0.43, 0.46, 0.030)) return true;
            if (SegmentInk(x, y, size, 0.42, 0.45, 0.47, 0.52, 0.010)) return true;
            if (CircleInk(u, v, 0.47, 0.55, 0.022)
                || CircleInk(u, v, 0.52, 0.58, 0.013)
                || CircleInk(u, v, 0.60, 0.58, 0.018)) return true;
            return false;
        }

        /// <summary>判断点是否落在一条带粗细的线段上。Checks whether a point lies on a thick line segment.</summary>
        /// <param name="x">像素 X。The pixel X.</param>
        /// <param name="y">像素 Y。The pixel Y.</param>
        /// <param name="size">图像边长。The image side length.</param>
        /// <param name="x1">起点 U。The start U.</param>
        /// <param name="y1">起点 V。The start V.</param>
        /// <param name="x2">终点 U。The end U.</param>
        /// <param name="y2">终点 V。The end V.</param>
        /// <param name="thickness">线宽（相对边长）。The thickness relative to the side length.</param>
        /// <returns>命中线段返回 true。True when the point lies on the segment.</returns>
        private static bool SegmentInk(int x, int y, int size, double x1, double y1, double x2, double y2, double thickness)
        {
            double u = x / (double)size;
            double v = y / (double)size;
            double dx = x2 - x1;
            double dy = y2 - y1;
            double length = dx * dx + dy * dy;
            if (length <= 0) return false;
            double t = ((u - x1) * dx + (v - y1) * dy) / length;
            if (t < 0) t = 0;
            else if (t > 1) t = 1;
            double px = x1 + dx * t - u;
            double py = y1 + dy * t - v;
            return px * px + py * py <= thickness * thickness;
        }

        /// <summary>判断相对坐标是否落在圆内。Checks whether relative coordinates lie inside a circle.</summary>
        /// <param name="u">相对 X。The relative X.</param>
        /// <param name="v">相对 Y。The relative Y.</param>
        /// <param name="cx">圆心 X。The center X.</param>
        /// <param name="cy">圆心 Y。The center Y.</param>
        /// <param name="radius">半径。The radius.</param>
        /// <returns>落在圆内返回 true。True when inside the circle.</returns>
        private static bool CircleInk(double u, double v, double cx, double cy, double radius)
        {
            double dx = u - cx;
            double dy = v - cy;
            return dx * dx + dy * dy <= radius * radius;
        }

        /// <summary>验证 C5/中式 5 号尺寸、仅收件人/仅寄件人、空地址、长寄件人、批量信封与非法版式。Verifies C5/Chinese No. 5 sizes, recipient-only and sender-only envelopes, empty addresses, long senders, batch envelopes and invalid layouts.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunEnvelopeEdgeTest(string outputDirectory)
        {
            WordEnvelopeLayout c5 = WordEnvelopeLayout.C5();
            OfficeDocument c5Document = WordEnvelopes.Create(new WordEnvelopeOptions
            {
                Layout = c5,
                Recipient = "张三"
            });
            RequireNear(c5Document.Sections[0].PageSetup.Width, c5.PageWidth, 0.01, "C5 width mismatch");
            RequireNear(c5Document.Sections[0].PageSetup.Height, c5.PageHeight, 0.01, "C5 height mismatch");
            Require(c5Document.Sections[0].PageSetup.Width > c5Document.Sections[0].PageSetup.Height,
                "C5 envelope should be landscape");

            WordEnvelopeLayout chinese5 = WordEnvelopeLayout.Chinese5();
            OfficeDocument chineseDocument = WordEnvelopes.Create(new WordEnvelopeOptions
            {
                Layout = chinese5,
                Recipient = "张三"
            });
            RequireNear(chineseDocument.Sections[0].PageSetup.Width, chinese5.PageWidth, 0.01,
                "Chinese No. 5 width mismatch");
            RequireNear(chineseDocument.Sections[0].PageSetup.Height, chinese5.PageHeight, 0.01,
                "Chinese No. 5 height mismatch");

            WordEnvelopeOptions recipientOnly = new WordEnvelopeOptions
            {
                Recipient = "张三\n上海市浦东新区世纪大道 100 号"
            };
            OfficeDocument recipientDocument = WordEnvelopes.Create(recipientOnly);
            Require(recipientDocument.Sections[0].Blocks.Count == 2, "recipient-only envelope block count mismatch");
            Paragraph firstRecipient = recipientDocument.Sections[0].Blocks[0] as Paragraph;
            Require(firstRecipient != null, "recipient-only paragraph missing");
            RequireNear(firstRecipient.SpaceBefore,
                recipientOnly.Layout.RecipientTop - recipientOnly.Layout.MarginTop, 0.01,
                "recipient-only vertical position mismatch");
            RequireNear(firstRecipient.LeftIndent,
                recipientOnly.Layout.RecipientLeft - recipientOnly.Layout.MarginLeft, 0.01,
                "recipient-only horizontal position mismatch");

            OfficeDocument empty = WordEnvelopes.Create(new WordEnvelopeOptions());
            Require(empty.Sections[0].Blocks.Count == 1 && empty.Sections[0].Blocks[0] is Paragraph,
                "empty envelope should contain a single empty paragraph");

            List<string> lines = new List<string>();
            for (int i = 0; i < 20; i++) lines.Add("寄件人第 " + (i + 1) + " 行");
            OfficeDocument longSender = WordEnvelopes.Create(new WordEnvelopeOptions
            {
                Sender = string.Join("\n", lines.ToArray()),
                Recipient = "张三"
            });
            Require(longSender.Sections[0].Blocks.Count == 21, "long sender block count mismatch");
            Paragraph pushedRecipient = longSender.Sections[0].Blocks[20] as Paragraph;
            Require(pushedRecipient != null && pushedRecipient.SpaceBefore == 0,
                "long sender should push the recipient instead of overlapping");

            WordEnvelopeLayout invalidLayout = WordEnvelopeLayout.Dl();
            invalidLayout.RecipientTop = invalidLayout.PageHeight + 1;
            RequireThrows<ArgumentException>(
                () => { WordEnvelopes.Create(new WordEnvelopeOptions { Layout = invalidLayout, Recipient = "张三" }); },
                "recipient outside the envelope must be rejected");
            WordEnvelopeLayout invalidSender = WordEnvelopeLayout.Dl();
            invalidSender.SenderLeft = invalidSender.PageWidth + 1;
            RequireThrows<ArgumentException>(
                () => { WordEnvelopes.Create(new WordEnvelopeOptions { Layout = invalidSender, Sender = "寄件人" }); },
                "sender outside the envelope must be rejected");

            OfficeDocument emptyBatch = WordEnvelopes.CreateMany(new object[0], "{{Name}}");
            Require(emptyBatch.Sections.Count == 1 && emptyBatch.Sections[0].Blocks.Count == 1,
                "empty envelope batch should keep one empty envelope");

            OfficeDocument batch = WordEnvelopes.CreateMany(
                new[] { new { Name = "甲" }, new { Name = "乙" } },
                "{{Name}}\n地址",
                new WordEnvelopeOptions { Sender = "寄件人" });
            Require(batch.Sections[0].Blocks.Count == 7, "batch envelope block count mismatch");
            int breaks = 0;
            for (int i = 0; i < batch.Sections[0].Blocks.Count; i++)
            {
                if (batch.Sections[0].Blocks[i] is PageBreak) breaks++;
            }
            Require(breaks == 1, "batch envelopes should insert one page break, got " + breaks);

            string batchPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-envelopes-edge.docx"));
            DocxWriter.Write(batch, batchPath);
            ValidateOoxmlPackage(batchPath, "word/document.xml");
            OfficeDocument batchRead = DocxReader.Load(batchPath);
            int readBreaks = 0;
            for (int i = 0; i < batchRead.Sections[0].Blocks.Count; i++)
            {
                if (batchRead.Sections[0].Blocks[i] is PageBreak) readBreaks++;
            }
            Require(readBreaks == 1 && batchRead.GetPlainText().Contains("乙"),
                "batch envelope structure lost in DOCX round trip");
        }

        /// <summary>生成指定数量的标签内容。Creates the requested number of label contents.</summary>
        /// <param name="count">标签数量。The number of labels.</param>
        /// <returns>标签内容列表。The label contents.</returns>
        private static List<string[]> CreateLabels(int count)
        {
            List<string[]> labels = new List<string[]>();
            for (int i = 0; i < count; i++)
            {
                labels.Add(new[] { "标签 " + (i + 1), "地址 " + (i + 1) });
            }
            return labels;
        }

        /// <summary>断言指定操作抛出指定异常。Requires the given action to throw the expected exception.</summary>
        /// <typeparam name="TException">期望的异常类型。The expected exception type.</typeparam>
        /// <param name="action">待执行的操作。The action to execute.</param>
        /// <param name="message">失败信息。The failure message.</param>
        private static void RequireThrows<TException>(Action action, string message) where TException : Exception
        {
            try
            {
                action();
            }
            catch (TException)
            {
                return;
            }
            catch (Exception exception)
            {
                throw new InvalidOperationException(
                    message + ": got " + exception.GetType().Name, exception);
            }
            throw new InvalidOperationException(message + ": no exception was thrown");
        }
    }
}
