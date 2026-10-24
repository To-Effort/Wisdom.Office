using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Wisdom.Excel;
using Wisdom.Office.Core;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>
        /// 输出侧与模板侧的策略性加固回归：CSV 公式注入开关、危险链接协议（file: 与 smb:）拒写、
        /// 模板填充剥离宏与嵌入对象。
        /// Regression tests for the output- and template-side policy hardening: the CSV formula-injection
        /// switch, refusal to write dangerous link schemes (file: and smb:), and the template fill stripping
        /// macro and embedded-object parts.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunOutputHardeningTest(string outputDirectory)
        {
            RunCsvFormulaEscapeTest(outputDirectory);
            RunDangerousLinkSchemeTest(outputDirectory);
            RunTemplateCarryoverStripTest(outputDirectory);
            RunTemplateCarryoverEvasionTest(outputDirectory);
            Console.WriteLine("Output hardening tests finished.");
        }

        /// <summary>
        /// CSV 公式注入：默认逐字写出（保持往返一致），显式打开开关后文本字段前置单引号，数值不变。
        /// CSV formula injection: the default writes fields verbatim (keeping the round trip exact); with the
        /// switch on, text fields gain a leading apostrophe while numbers stay untouched.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunCsvFormulaEscapeTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Csv");
            sheet.GetCell(1, 1).Value = "=1+1";
            sheet.GetCell(1, 2).Value = "+8613800000000";
            sheet.GetCell(1, 3).Value = "-45";
            sheet.GetCell(1, 4).Value = "@SUM(A1)";
            sheet.GetCell(1, 5).Value = 45d;
            sheet.GetCell(1, 6).Value = "普通文本";

            string plainPath = PrepareOutputPath(Path.Combine(outputDirectory, "csv-formula-default.csv"));
            ExcelCsv.Write(workbook, plainPath);
            string plain = FirstLine(plainPath);
            Require(plain == "=1+1,+8613800000000,-45,@SUM(A1),45,普通文本",
                "CSV default export must keep text verbatim: " + plain);

            ExcelCsvOptions guarded = new ExcelCsvOptions();
            guarded.EscapeFormulaLikeText = true;
            string guardedPath = PrepareOutputPath(Path.Combine(outputDirectory, "csv-formula-guarded.csv"));
            ExcelCsv.Write(workbook, guardedPath, guarded);
            string guardedLine = FirstLine(guardedPath);
            Require(guardedLine == "'=1+1,'+8613800000000,'-45,'@SUM(A1),45,普通文本",
                "CSV guarded export must neutralize text and keep numbers: " + guardedLine);

            OfficeWorkbook reloaded = ExcelCsv.Read(guardedPath);
            Worksheet read = reloaded.Worksheets[0];
            Require(Equals(read.GetCell(1, 1).Value, "'=1+1"), "guarded apostrophe is kept as literal text");
            Require(Equals(read.GetCell(1, 6).Value, "普通文本"), "safe text is not escaped");

            Console.WriteLine("[hardening] csv formula switch verified");
        }

        /// <summary>只取文本文件的第一行（自动去 BOM）。Returns the first line of a text file, BOM removed.</summary>
        /// <param name="path">文件路径。The file path.</param>
        /// <returns>第一行内容。The first line.</returns>
        private static string FirstLine(string path)
        {
            string text = File.ReadAllText(path);
            if (text.Length > 0 && text[0] == '\uFEFF') text = text.Substring(1);
            int breakIndex = text.IndexOf("\r\n", StringComparison.Ordinal);
            if (breakIndex < 0) breakIndex = text.IndexOf('\n');
            return breakIndex < 0 ? text : text.Substring(0, breakIndex);
        }

        /// <summary>
        /// file: 与 smb: 在 Word 与 Excel 的写出侧一律拒写：链接文字保留，关系不落地。
        /// file: and smb: are refused on the write path of both Word and Excel: the link text stays, the
        /// relationship is never emitted.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunDangerousLinkSchemeTest(string outputDirectory)
        {
            WordDocumentBuilder builder = new WordDocumentBuilder();
            builder.Font.FontFamily = "Microsoft YaHei";
            builder.InsertParagraph("链接：");
            builder.InsertHyperlink("官网", "https://example.com/report");
            builder.InsertParagraph("危险链接：");
            builder.InsertHyperlink("本机", "file:///C:/Windows/win.ini");
            builder.InsertHyperlink("共享", "smb://server/share/doc.docx");
            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "hardening-links.docx"));
            DocxWriter.Write(builder.Document, docxPath);

            string docxRelationships;
            using (ZipArchive archive = ZipFile.OpenRead(docxPath))
            {
                docxRelationships = ReadZipEntryText(archive, "word/_rels/document.xml.rels");
            }
            Require(docxRelationships.IndexOf("file:", StringComparison.OrdinalIgnoreCase) < 0,
                "word must not write a file: relationship");
            Require(docxRelationships.IndexOf("smb:", StringComparison.OrdinalIgnoreCase) < 0,
                "word must not write a smb: relationship");
            RequireContains(docxRelationships, "https://example.com/report");
            string wordText = DocxReader.Load(docxPath).GetPlainText();
            RequireContains(wordText, "本机");
            RequireContains(wordText, "共享");

            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Links");
            sheet.GetCell(1, 1).Value = "本机";
            sheet.GetCell(1, 1).Hyperlink = "file:///C:/Windows/win.ini";
            sheet.GetCell(2, 1).Value = "共享";
            sheet.GetCell(2, 1).Hyperlink = "smb://server/share/doc.docx";
            sheet.GetCell(3, 1).Value = "官网";
            sheet.GetCell(3, 1).Hyperlink = "https://example.com/report";
            string xlsxPath = PrepareOutputPath(Path.Combine(outputDirectory, "hardening-links.xlsx"));
            XlsxWriter.Write(workbook, xlsxPath);

            string sheetRelationships;
            using (ZipArchive archive = ZipFile.OpenRead(xlsxPath))
            {
                sheetRelationships = ReadZipEntryText(archive, "xl/worksheets/_rels/sheet1.xml.rels");
            }
            Require(sheetRelationships.IndexOf("file:", StringComparison.OrdinalIgnoreCase) < 0,
                "excel must not write a file: hyperlink");
            Require(sheetRelationships.IndexOf("smb:", StringComparison.OrdinalIgnoreCase) < 0,
                "excel must not write a smb: hyperlink");
            RequireContains(sheetRelationships, "https://example.com/report");
            Worksheet read = XlsxReader.Load(xlsxPath).Worksheets[0];
            Require(read.GetCell(1, 1).Hyperlink == null, "file: hyperlink is dropped on read back");
            Require(read.GetCell(2, 1).Hyperlink == null, "smb: hyperlink is dropped on read back");
            Require(Equals(read.GetCell(3, 1).Hyperlink, "https://example.com/report"), "safe hyperlink survives");

            Console.WriteLine("[hardening] file: and smb: links refused");
        }

        /// <summary>
        /// 模板填充剥离宏与嵌入对象：vbaProject.bin、word/embeddings/* 不再随结果带走，指向它们的关系统通删除，
        /// [Content_Types].xml 的 Override 一并清理，包结构保持自洽，超链接文字与占位符替换不受影响。
        /// The template fill strips macro and embedded-object parts: vbaProject.bin and word/embeddings/* no
        /// longer travel with the result, the relationships pointing at them go, the matching
        /// [Content_Types].xml overrides are cleaned too, the package stays self-consistent, and neither the
        /// hyperlink text nor placeholder replacement is affected.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunTemplateCarryoverStripTest(string outputDirectory)
        {
            const string documentXml =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:body>"
                + "<w:p><w:r><w:t>客户：{{customerName}}</w:t></w:r></w:p>"
                + "<w:p><w:hyperlink r:id=\"rIdSafe\"><w:r><w:t>官网</w:t></w:r></w:hyperlink>"
                + "<w:hyperlink r:id=\"rIdFile\"><w:r><w:t>本机</w:t></w:r></w:hyperlink>"
                + "<w:hyperlink r:id=\"rIdSmb\"><w:r><w:t>共享</w:t></w:r></w:hyperlink></w:p>"
                + "<w:p><w:r><w:object r:id=\"rIdEmbed\"><w:t>OLE</w:t></w:object></w:r></w:p>"
                + "</w:body></w:document>";

            const string relationships =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rIdSafe\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\""
                + " Target=\"https://example.com/report\" TargetMode=\"External\"/>"
                + "<Relationship Id=\"rIdFile\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\""
                + " Target=\"file:///C:/Windows/win.ini\" TargetMode=\"External\"/>"
                + "<Relationship Id=\"rIdSmb\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\""
                + " Target=\"smb://server/share/doc.docx\" TargetMode=\"External\"/>"
                + "<Relationship Id=\"rIdEmbed\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject\""
                + " Target=\"embeddings/Microsoft_Excel_Worksheet1.xlsx\"/>"
                + "<Relationship Id=\"rIdVba\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\""
                + " Target=\"vbaProject.bin\"/>"
                + "<Relationship Id=\"rIdImage\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\""
                + " Target=\"media/image1.png\"/>"
                + "</Relationships>";

            const string contentTypes =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Default Extension=\"xlsx\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\"/>"
                + "<Default Extension=\"png\" ContentType=\"image/png\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>"
                + "</Types>";

            const string packageRelationships =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\""
                + " Target=\"word/document.xml\"/>"
                + "</Relationships>";

            Dictionary<string, byte[]> entries = new Dictionary<string, byte[]>();
            entries["[Content_Types].xml"] = Encoding.UTF8.GetBytes(contentTypes);
            entries["_rels/.rels"] = Encoding.UTF8.GetBytes(packageRelationships);
            entries["word/document.xml"] = Encoding.UTF8.GetBytes(documentXml);
            entries["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(relationships);
            entries["word/vbaProject.bin"] = Encoding.UTF8.GetBytes("VBProject placeholder");
            entries["word/embeddings/Microsoft_Excel_Worksheet1.xlsx"] = Encoding.UTF8.GetBytes("OLE payload");
            entries["word/media/image1.png"] = CreateTestPng(1, 1);
            byte[] templateBytes = CreateZipBytes(entries);

            TemplateContext context = new TemplateContext();
            context.Set("customerName", "华东供应链有限公司");
            byte[] filledBytes = DocxTemplate.Fill(templateBytes, context);
            string filledPath = PrepareOutputPath(Path.Combine(outputDirectory, "hardening-template-filled.docx"));
            File.WriteAllBytes(filledPath, filledBytes);

            Dictionary<string, string> parts = ReadPackageBytes(filledBytes);
            Require(!parts.ContainsKey("word/vbaProject.bin"),
                "the macro project must not be carried over; entries: " + string.Join(" | ", parts.Keys));
            Require(!parts.ContainsKey("word/embeddings/Microsoft_Excel_Worksheet1.xlsx"),
                "the embedded OLE object must not be carried over");
            Require(parts.ContainsKey("word/media/image1.png"), "an ordinary image part stays");
            Require(parts.ContainsKey("word/document.xml"), "the main document part stays");

            string document = parts["word/document.xml"];
            RequireContains(document, "华东供应链有限公司");
            Require(document.IndexOf("rIdEmbed", StringComparison.Ordinal) < 0, "the object reference must be gone");
            Require(document.IndexOf("<w:object", StringComparison.Ordinal) < 0, "the embedded object element must be gone");
            Require(document.IndexOf("rIdFile", StringComparison.Ordinal) < 0, "the file: hyperlink reference must be gone");
            Require(document.IndexOf("rIdSmb", StringComparison.Ordinal) < 0, "the smb: hyperlink reference must be gone");
            Require(document.IndexOf("rIdSafe", StringComparison.Ordinal) >= 0, "the safe hyperlink keeps its reference");

            string remainingRelationships = parts["word/_rels/document.xml.rels"];
            Require(remainingRelationships.IndexOf("file:", StringComparison.OrdinalIgnoreCase) < 0,
                "the file: relationship must be removed");
            Require(remainingRelationships.IndexOf("smb:", StringComparison.OrdinalIgnoreCase) < 0,
                "the smb: relationship must be removed");
            Require(remainingRelationships.IndexOf("vbaProject") < 0, "the macro relationship must be removed");
            Require(remainingRelationships.IndexOf("embeddings") < 0, "the embedded object relationship must be removed");
            RequireContains(remainingRelationships, "https://example.com/report");
            RequireContains(remainingRelationships, "media/image1.png");

            Require(parts["[Content_Types].xml"].IndexOf("vbaProject", StringComparison.Ordinal) < 0,
                "the macro content type override must be removed");
            RequireContains(parts["[Content_Types].xml"], "/word/document.xml");

            // 包结构必须在剥离后仍然自洽：每一个关系目标与每一个 Override 都指向真实存在的部件。
            // The package must stay self-consistent after the strip: every relationship target and every
            // override still resolves to a part that exists.
            ValidateOoxmlPackage(filledPath, "word/document.xml");
            string plainText = DocxReader.Load(filledPath).GetPlainText();
            RequireContains(plainText, "官网");
            RequireContains(plainText, "本机");
            RequireContains(plainText, "共享");
            RequireContains(plainText, "客户：华东供应链有限公司");

            Console.WriteLine("[hardening] template carryover stripped");
        }

        /// <summary>
        /// 剥离链的四种绕过：../ 相对目标、规范化后落在根级的 embeddings/、不带占位符的脚注部件里的引用、
        /// 以及被删部件自己的关系文件；同时验证 context 为 null 时悬空 r:id 依然被清理。
        /// Four ways the strip used to be bypassed: a ../ relative target, an embeddings/ folder that lands at
        /// the package root after normalisation, a reference inside footnotes.xml which carries no placeholders,
        /// and the relationships file of a dropped part; the dangling ids are also expected to be cleaned when
        /// the context is null.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunTemplateCarryoverEvasionTest(string outputDirectory)
        {
            const string mainDocument =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:body>"
                + "<w:p><w:r><w:t>客户：{{customerName}}</w:t></w:r></w:p>"
                + "<w:p><w:r><w:control r:id=\"rIdAx\"/></w:r></w:p>"
                + "<w:p><w:r><w:object r:id=\"rIdRoot\"><w:t>OLE</w:t></w:object></w:r></w:p>"
                + "<w:p><w:r><w:t>脚注</w:t></w:r></w:p>"
                + "</w:body></w:document>";

            const string footnotes =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<w:footnotes xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<w:footnote w:id=\"1\">"
                + "<w:p><w:r><w:object r:id=\"rIdFnOle\"><w:t>FN</w:t></w:object></w:r></w:p>"
                + "<w:p><w:hyperlink r:id=\"rIdFnSmb\"><w:r><w:t>共享文件</w:t></w:r></w:hyperlink></w:p>"
                + "</w:footnote></w:footnotes>";

            const string mainRelationships =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rIdAx\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/control\""
                + " Target=\"activeX/activeX1.xml\"/>"
                + "<Relationship Id=\"rIdRoot\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package\""
                + " Target=\"../embeddings/Workbook.bin\"/>"
                + "<Relationship Id=\"rIdVba\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\""
                + " Target=\"vbaProject.bin\"/>"
                + "<Relationship Id=\"rIdFootnotes\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\""
                + " Target=\"footnotes.xml\"/>"
                + "</Relationships>";

            const string footnoteRelationships =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rIdFnOle\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package\""
                + " Target=\"../embeddings/FootnoteOle.bin\"/>"
                + "<Relationship Id=\"rIdFnSmb\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\""
                + " Target=\"smb://server/share/doc.docx\" TargetMode=\"External\"/>"
                + "</Relationships>";

            const string activeXRelationships =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rIdAxBin\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject\""
                + " Target=\"activeX1.bin\"/>"
                + "</Relationships>";

            const string contentTypes =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Default Extension=\"bin\" ContentType=\"application/octet-stream\"/>"
                + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                + "<Override PartName=\"/word/footnotes.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml\"/>"
                + "<Override PartName=\"/word/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>"
                + "<Override PartName=\"/word/activeX/activeX1.xml\" ContentType=\"application/vnd.ms-office.activeX+xml\"/>"
                + "</Types>";

            const string packageRelationships =
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\""
                + " Target=\"word/document.xml\"/>"
                + "</Relationships>";

            Dictionary<string, byte[]> entries = new Dictionary<string, byte[]>();
            entries["[Content_Types].xml"] = Encoding.UTF8.GetBytes(contentTypes);
            entries["_rels/.rels"] = Encoding.UTF8.GetBytes(packageRelationships);
            entries["word/document.xml"] = Encoding.UTF8.GetBytes(mainDocument);
            entries["word/_rels/document.xml.rels"] = Encoding.UTF8.GetBytes(mainRelationships);
            entries["word/footnotes.xml"] = Encoding.UTF8.GetBytes(footnotes);
            entries["word/_rels/footnotes.xml.rels"] = Encoding.UTF8.GetBytes(footnoteRelationships);
            entries["word/activeX/activeX1.xml"] = Encoding.UTF8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"utf-8\"?><ax:activeX xmlns:ax=\"urn:schemas-microsoft-com:office:activeX\"/>");
            entries["word/activeX/_rels/activeX1.xml.rels"] = Encoding.UTF8.GetBytes(activeXRelationships);
            entries["word/activeX/activeX1.bin"] = Encoding.UTF8.GetBytes("activeX control binary");
            entries["word/vbaProject.bin"] = Encoding.UTF8.GetBytes("VBProject placeholder");
            entries["embeddings/Workbook.bin"] = Encoding.UTF8.GetBytes("root level OLE payload");
            entries["embeddings/FootnoteOle.bin"] = Encoding.UTF8.GetBytes("footnote OLE payload");
            byte[] templateBytes = CreateZipBytes(entries);

            TemplateContext context = new TemplateContext();
            context.Set("customerName", "华东供应链有限公司");
            string filledPath = PrepareOutputPath(Path.Combine(outputDirectory, "hardening-template-evasion.docx"));
            File.WriteAllBytes(filledPath, DocxTemplate.Fill(templateBytes, context));

            Dictionary<string, string> parts = ReadPackageParts(filledPath);
            Require(!parts.ContainsKey("word/vbaProject.bin"), "the macro project must be stripped");
            Require(!parts.ContainsKey("word/activeX/activeX1.xml"), "the activeX part must be stripped");
            Require(!parts.ContainsKey("word/activeX/_rels/activeX1.xml.rels"),
                "the relationships file of a stripped part must go with it");
            Require(!parts.ContainsKey("word/activeX/activeX1.bin"),
                "a part reachable only from a stripped part must go as well");
            Require(!parts.ContainsKey("embeddings/Workbook.bin"),
                "a ../ target that lands at the package root must still be stripped");
            Require(!parts.ContainsKey("embeddings/FootnoteOle.bin"),
                "an embedded object referenced from footnotes.xml must be stripped");
            Require(parts.ContainsKey("word/footnotes.xml"), "footnotes.xml itself stays");

            Require(parts["word/document.xml"].IndexOf("rIdAx", StringComparison.Ordinal) < 0, "the control id is gone");
            Require(parts["word/document.xml"].IndexOf("rIdRoot", StringComparison.Ordinal) < 0, "the object id is gone");
            Require(parts["word/document.xml"].IndexOf("<w:object", StringComparison.Ordinal) < 0, "the object element is gone");
            Require(parts["word/document.xml"].IndexOf("<w:control", StringComparison.Ordinal) < 0, "the control element is gone");
            RequireContains(parts["word/document.xml"], "华东供应链有限公司");

            // 脚注部件里没有占位符，但它同样是引用者：失效的关系必须清掉，链接文字要留下。
            // footnotes.xml holds no placeholders yet it is a referrer too: the dead id has to go and the
            // link text has to stay.
            Require(parts["word/footnotes.xml"].IndexOf("rIdFnOle", StringComparison.Ordinal) < 0,
                "the footnotes object id is gone");
            Require(parts["word/footnotes.xml"].IndexOf("rIdFnSmb", StringComparison.Ordinal) < 0,
                "the footnotes smb: id is gone");
            Require(parts["word/footnotes.xml"].IndexOf("<w:object", StringComparison.Ordinal) < 0,
                "the footnotes object element is gone");
            RequireContains(parts["word/footnotes.xml"], "共享文件");
            Require(parts["word/_rels/footnotes.xml.rels"].IndexOf("smb:") < 0, "the smb: relationship is pruned");
            Require(parts["word/_rels/footnotes.xml.rels"].IndexOf("embeddings") < 0, "the object relationship is pruned");
            Require(parts["[Content_Types].xml"].IndexOf("vbaProject") < 0, "the macro override is gone");
            Require(parts["[Content_Types].xml"].IndexOf("activeX") < 0, "the activeX override is gone");
            RequireContains(parts["[Content_Types].xml"], "/word/footnotes.xml");
            ValidateOoxmlPackage(filledPath, "word/document.xml");

            // 上下文为 null 时不展开占位符，但剥离与引用清理照旧。
            // A null context expands no placeholders, yet the strip and the reference cleanup still happen.
            Dictionary<string, string> emptyContextParts =
                ReadPackageBytes(DocxTemplate.Fill(templateBytes, null));
            Require(!emptyContextParts.ContainsKey("word/vbaProject.bin"), "the strip works without a context");
            Require(emptyContextParts["word/document.xml"].IndexOf("rIdAx", StringComparison.Ordinal) < 0,
                "dangling ids are cleaned without a context");
            RequireContains(emptyContextParts["word/document.xml"], "{{customerName}}");

            Console.WriteLine("[hardening] template strip evasions closed");
        }

        /// <summary>读出包里每个条目的文本（二进制条目记为空串）。Reads every entry of a package as text, binary entries map to an empty string.</summary>
        /// <param name="path">包路径。The package path.</param>
        /// <returns>条目名到文本。The map from entry name to text.</returns>
        private static Dictionary<string, string> ReadPackageParts(string path)
        {
            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                return ReadPackageEntries(archive);
            }
        }

        /// <summary>把包字节展开为条目名到文本的映射。Expands package bytes into the map from entry name to text.</summary>
        /// <param name="data">包字节。The package bytes.</param>
        /// <returns>条目名到文本。The map from entry name to text.</returns>
        private static Dictionary<string, string> ReadPackageBytes(byte[] data)
        {
            using (MemoryStream stream = new MemoryStream(data, false))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                return ReadPackageEntries(archive);
            }
        }

        /// <summary>遍历包条目并取出文本。Walks the package entries and reads their text.</summary>
        /// <param name="archive">已打开的包。The open package.</param>
        /// <returns>条目名到文本。The map from entry name to text.</returns>
        private static Dictionary<string, string> ReadPackageEntries(ZipArchive archive)
        {
            Dictionary<string, string> parts = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    || entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    parts[entry.FullName] = ReadZipEntryText(archive, entry.FullName);
                }
                else
                {
                    parts[entry.FullName] = string.Empty;
                }
            }
            return parts;
        }
    }
}
