using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Wisdom.Office.Core;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>执行自定义属性、标签纸与信封测试。Runs the custom-property, label-sheet and envelope tests.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWordLabelsAndPropertiesTest(string outputDirectory)
        {
            RunCustomPropertiesTest(outputDirectory);
            RunLabelSheetTest(outputDirectory);
            RunEnvelopeTest(outputDirectory);
            RunWordLabelsAndPropertiesEdgeTest(outputDirectory);
            Console.WriteLine("Word custom properties, labels and envelopes tests finished.");
        }

        /// <summary>验证带类型的自定义属性写出、包结构、读回、克隆与再次往返。Verifies typed custom-property writing, package structure, reading, cloning and a second round trip.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunCustomPropertiesTest(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-custom-properties.docx"));
            OfficeDocument document = new OfficeDocument();
            document.AddSection().AddParagraph("自定义属性测试正文。");
            CustomPropertyCollection properties = document.Metadata.CustomProperties;
            properties.Set("项目编号", "QF-2026-018");
            properties.Set("版本", 3);
            properties.Set("文件大小", 4294967296L);
            properties.Set("已审核", true);
            properties.Set("抽检比例", 0.15);
            properties.Set("金额", 1234.56m);
            DateTime archived = new DateTime(2026, 10, 20, 8, 30, 0, DateTimeKind.Utc);
            properties.Set("归档时间", archived);
            Require(properties.Count == 7, "custom property count mismatch: " + properties.Count);

            DocxWriter.Write(document, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            Require(PackageContainsEntry(path, "docProps/custom.xml"), "custom.xml part is missing");
            string customXml = ReadPackageEntryText(path, "docProps/custom.xml");
            RequireContains(customXml, "custom-properties");
            RequireContains(customXml, "<vt:lpwstr>QF-2026-018</vt:lpwstr>");
            RequireContains(customXml, "<vt:i4>3</vt:i4>");
            RequireContains(customXml, "<vt:i8>4294967296</vt:i8>");
            RequireContains(customXml, "<vt:bool>true</vt:bool>");
            RequireContains(customXml, "<vt:r8>0.15</vt:r8>");
            RequireContains(customXml, "<vt:decimal>1234.56</vt:decimal>");
            RequireContains(customXml, "<vt:filetime>2026-10-20T08:30:00Z</vt:filetime>");

            string contentTypes = ReadPackageEntryText(path, "[Content_Types].xml");
            RequireContains(contentTypes, "custom-properties+xml");
            string rootRelationships = ReadPackageEntryText(path, "_rels/.rels");
            RequireContains(rootRelationships, "relationships/custom-properties");
            RequireContains(rootRelationships, "docProps/custom.xml");

            OfficeDocument loaded = DocxReader.Load(path);
            CustomPropertyCollection read = loaded.Metadata.CustomProperties;
            Require(read.Count == 7, "custom property count not restored: " + read.Count);
            Require(read.GetString("项目编号") == "QF-2026-018", "text custom property mismatch");
            Require(read.GetInt32("版本") == 3, "int custom property mismatch");
            Require(read.GetInt64("文件大小") == 4294967296L, "long custom property mismatch");
            Require(read.GetBoolean("已审核"), "bool custom property mismatch");
            Require(Math.Abs(read.GetDouble("抽检比例") - 0.15) < 0.000001, "double custom property mismatch");
            object amount;
            Require(read.TryGetValue("金额", out amount) && amount is decimal && (decimal)amount == 1234.56m,
                "decimal custom property mismatch");
            Require(read.GetDateTime("归档时间") == archived, "DateTime custom property mismatch");
            Require(read.Names[0] == "项目编号" && read.Names[6] == "归档时间", "custom property order not preserved");

            OfficeDocument cloned = WordDocumentOperations.Clone(loaded);
            cloned.Metadata.CustomProperties.Set("版本", 4);
            Require(loaded.Metadata.CustomProperties.GetInt32("版本") == 3, "clone shares custom properties with the source");
            Require(cloned.Metadata.CustomProperties.GetInt32("版本") == 4, "clone custom property was not updated");

            string secondPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-custom-properties-roundtrip.docx"));
            DocxWriter.Write(cloned, secondPath);
            OfficeDocument second = DocxReader.Load(secondPath);
            Require(second.Metadata.CustomProperties.GetInt32("版本") == 4, "custom property second round trip failed");
            Require(second.Metadata.CustomProperties.GetString("项目编号") == "QF-2026-018", "text property second round trip failed");

            string emptyPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-custom-properties-empty.docx"));
            DocxWriter.Write(new OfficeDocument(), emptyPath);
            Require(!PackageContainsEntry(emptyPath, "docProps/custom.xml"), "empty custom-property part should not be written");
        }

        /// <summary>验证标签纸版式、页容量、分页、表格几何与记录模板。Verifies label layout, per-page capacity, pagination, table geometry and record templates.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunLabelSheetTest(string outputDirectory)
        {
            WordLabelLayout layout = WordLabelLayout.AveryL7160();
            Require(layout.Columns == 3 && layout.Rows == 7, "L7160 grid mismatch");
            Require(layout.LabelsPerPage == 21, "L7160 labels per page mismatch");
            RequireNear(layout.PageWidth, 595.28, 0.02, "L7160 page width mismatch");
            RequireNear(layout.LabelWidth, 180, 0.01, "L7160 label width mismatch");
            PageSetup setup = layout.CreatePageSetup();
            RequireNear(setup.MarginLeft, layout.LeftMargin, 0.01, "label left margin mismatch");
            RequireNear(setup.MarginTop, layout.TopMargin, 0.01, "label top margin mismatch");

            List<string[]> labels = new List<string[]>();
            for (int i = 0; i < 22; i++)
            {
                labels.Add(new[] { "标签 " + (i + 1), "上海市浦东新区示例路 " + (i + 1) + " 号" });
            }

            OfficeDocument sheet = WordLabels.Create(layout, labels);
            Require(sheet.Sections.Count == 1, "label sheet section count mismatch");
            List<Block> blocks = sheet.Sections[0].Blocks;
            Require(blocks.Count == 3, "label sheet should contain two tables and one page break, got " + blocks.Count);
            Table first = blocks[0] as Table;
            Table second = blocks[2] as Table;
            Require(first != null && second != null, "label pages should be tables");
            Require(blocks[1] is PageBreak, "label pages should be separated by a page break");
            Require(first.Columns.Count == 5, "label table should have 3 label columns and 2 gaps, got " + first.Columns.Count);
            Require(first.Rows.Count == 7, "first label page should have 7 rows, got " + first.Rows.Count);
            Require(second.Rows.Count == 7, "second label page should have 7 rows, got " + second.Rows.Count);
            Require(first.Borders.Top.Style == BorderStyle.None, "label table should be borderless");
            Require(first.Rows[0].Cells.Count == 5, "label row should include gap cells");
            Require(TextOf(first.Rows[0].Cells[0]) == "标签 1" + Environment.NewLine + "上海市浦东新区示例路 1 号",
                "first label text mismatch: " + TextOf(first.Rows[0].Cells[0]));
            Require(TextOf(first.Rows[6].Cells[4]) == "标签 21" + Environment.NewLine + "上海市浦东新区示例路 21 号",
                "last label on first page mismatch");
            Require(TextOf(second.Rows[0].Cells[0]) == "标签 22" + Environment.NewLine + "上海市浦东新区示例路 22 号",
                "first label on second page mismatch");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-labels.docx"));
            DocxWriter.Write(sheet, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            string xml = ReadPackageEntryText(path, "word/document.xml");
            RequireContains(xml, "标签 22");
            RequireContains(xml, "w:type=\"page\"");

            OfficeDocument readBack = DocxReader.Load(path);
            Require(readBack.GetPlainText().Contains("标签 22"), "label text missing after DOCX round trip");
            List<Block> readBlocks = readBack.Sections[0].Blocks;
            // 正文以表格结尾时写出器会补一个空段落，这是 Word 的兼容性要求。When the body ends with a table the
            // writer appends an empty paragraph, which Word requires for compatibility.
            Require(readBlocks.Count == 4, "label page structure not preserved: " + readBlocks.Count);
            Require(readBlocks[0] is Table && readBlocks[2] is Table, "label tables missing after DOCX round trip");
            Require(readBlocks[1] is PageBreak, "label page break missing after DOCX round trip");
            Paragraph trailing = readBlocks[3] as Paragraph;
            Require(trailing != null && trailing.GetText().Length == 0,
                "label sheet should end with the compatibility paragraph");

            OfficeDocument fromRecords = WordLabels.CreateFromRecords(
                WordLabelLayout.Avery5163(),
                new[]
                {
                    new { Name = "张三", Address = "上海市浦东新区世纪大道 100 号", Zip = "200120" },
                    new { Name = "李四", Address = "北京市朝阳区建国路 8 号", Zip = "100022" }
                },
                "{{Name}}\n{{Address}}\n{{Zip}}");
            Table recordTable = fromRecords.Sections[0].Blocks[0] as Table;
            Require(recordTable != null, "record label table missing");
            Require(TextOf(recordTable.Rows[0].Cells[0]) == "张三" + Environment.NewLine + "上海市浦东新区世纪大道 100 号" + Environment.NewLine + "200120",
                "record label text mismatch: " + TextOf(recordTable.Rows[0].Cells[0]));
            Require(TextOf(recordTable.Rows[0].Cells[2]) == "李四" + Environment.NewLine + "北京市朝阳区建国路 8 号" + Environment.NewLine + "100022",
                "second record label text mismatch: " + TextOf(recordTable.Rows[0].Cells[2]));

            string recordPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-labels-records.docx"));
            DocxWriter.Write(fromRecords, recordPath);
            ValidateOoxmlPackage(recordPath, "word/document.xml");
        }

        /// <summary>验证信封页面尺寸、地址定位、模板批量生成与 DOCX 往返。Verifies envelope page size, address placement, template batches and DOCX round trips.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunEnvelopeTest(string outputDirectory)
        {
            WordEnvelopeLayout layout = WordEnvelopeLayout.Dl();
            RequireNear(layout.PageWidth, 623.62, 0.02, "DL envelope width mismatch");
            RequireNear(layout.PageHeight, 311.81, 0.02, "DL envelope height mismatch");

            WordEnvelopeOptions options = new WordEnvelopeOptions
            {
                Layout = layout,
                Sender = "智慧办公\n上海市浦东新区示例路 1 号\n200120",
                Recipient = "张三\n上海市浦东新区世纪大道 100 号\n200120"
            };
            OfficeDocument envelope = WordEnvelopes.Create(options);
            Section section = envelope.Sections[0];
            RequireNear(section.PageSetup.Width, layout.PageWidth, 0.01, "envelope page width not applied");
            RequireNear(section.PageSetup.Height, layout.PageHeight, 0.01, "envelope page height not applied");
            Require(section.Blocks.Count == 6, "envelope paragraph count mismatch: " + section.Blocks.Count);

            Paragraph firstSender = section.Blocks[0] as Paragraph;
            Paragraph firstRecipient = section.Blocks[3] as Paragraph;
            Require(firstSender != null && firstRecipient != null, "envelope address blocks missing");
            RequireNear(firstSender.SpaceBefore, layout.SenderTop - layout.MarginTop, 0.01, "sender top position mismatch");
            RequireNear(firstSender.LeftIndent, layout.SenderLeft - layout.MarginLeft, 0.01, "sender left position mismatch");
            Require(firstRecipient.SpaceBefore > 0, "recipient top position should use space-before");
            RequireNear(firstRecipient.LeftIndent, layout.RecipientLeft - layout.MarginLeft, 0.01, "recipient left position mismatch");
            Require(firstRecipient.GetText() == "张三", "recipient first line mismatch");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-envelope.docx"));
            DocxWriter.Write(envelope, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            OfficeDocument readBack = DocxReader.Load(path);
            string text = readBack.GetPlainText();
            RequireContains(text, "智慧办公");
            RequireContains(text, "张三");
            // pgSz 以整数 twips 存储，往返存在不足 0.05 磅的量化误差。pgSz is stored in whole twips, so the round
            // trip carries a quantization error below 0.05 points.
            RequireNear(readBack.Sections[0].PageSetup.Width, layout.PageWidth, 0.1, "envelope width lost in round trip");

            OfficeDocument many = WordEnvelopes.CreateMany(
                new[]
                {
                    new { Name = "张三", Address = "上海市浦东新区世纪大道 100 号" },
                    new { Name = "李四", Address = "北京市朝阳区建国路 8 号" }
                },
                "{{Name}}\n{{Address}}",
                options,
                "智慧办公");
            Require(many.Sections[0].Blocks.Count > 4, "batch envelopes missing content");
            int pageBreaks = 0;
            int recipients = 0;
            for (int i = 0; i < many.Sections[0].Blocks.Count; i++)
            {
                if (many.Sections[0].Blocks[i] is PageBreak) pageBreaks++;
                Paragraph paragraph = many.Sections[0].Blocks[i] as Paragraph;
                if (paragraph != null && paragraph.GetText() == "张三") recipients++;
            }
            Require(pageBreaks == 1, "batch envelopes should insert one page break, got " + pageBreaks);
            Require(recipients == 1, "batch envelope recipient block missing");
            Require(many.GetPlainText().Contains("李四"), "second batch recipient missing");

            string manyPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-envelopes-batch.docx"));
            DocxWriter.Write(many, manyPath);
            ValidateOoxmlPackage(manyPath, "word/document.xml");
        }

        /// <summary>读取单元格内全部段落的文本（段落之间以换行连接）。Reads all paragraph text in a cell, joining paragraphs with line breaks.</summary>
        /// <param name="cell">表格单元格。The table cell.</param>
        /// <returns>单元格文本。The cell text.</returns>
        private static string TextOf(TableCell cell)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < cell.Blocks.Count; i++)
            {
                Paragraph paragraph = cell.Blocks[i] as Paragraph;
                if (paragraph == null) continue;
                if (builder.Length > 0) builder.Append(Environment.NewLine);
                builder.Append(paragraph.GetText());
            }
            return builder.ToString();
        }

        /// <summary>判断 ZIP 包是否包含指定条目。Determines whether the ZIP package contains the given entry.</summary>
        /// <param name="path">包文件路径。The package path.</param>
        /// <param name="entryName">条目名。The entry name.</param>
        /// <returns>存在时返回 true。True when present.</returns>
        private static bool PackageContainsEntry(string path, string entryName)
        {
            using (FileStream stream = File.OpenRead(path))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                return archive.GetEntry(entryName) != null;
            }
        }

        /// <summary>断言两个数值在容差内相等。Requires two numbers to be equal within a tolerance.</summary>
        /// <param name="actual">实际值。The actual value.</param>
        /// <param name="expected">期望值。The expected value.</param>
        /// <param name="tolerance">允许误差。The allowed difference.</param>
        /// <param name="message">失败信息。The failure message.</param>
        private static void RequireNear(double actual, double expected, double tolerance, string message)
        {
            Require(Math.Abs(actual - expected) <= tolerance, message + ": expected " + expected + ", got " + actual);
        }
    }
}
