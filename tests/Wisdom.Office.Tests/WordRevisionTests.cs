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
        /// <summary>执行 Word 修订、SmartArt 与数字签名检查测试。Runs the Word revisions, SmartArt and digital-signature inspection tests.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWordRevisionTest(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-revisions.docx"));
            RunWordTrackedChangeTest(path);
            string secondPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-revisions-roundtrip.docx"));
            RunWordTrackedChangeRoundTripTest(path, secondPath);
            RunWordAcceptRejectTest(path);
            RunWordCloneRevisionTest(path);
            RunSmartArtInspectionTest(outputDirectory);
            RunDigitalSignatureInspectionTest(outputDirectory);
            Console.WriteLine("Word revision and package inspection tests finished.");
        }

        /// <summary>验证插入、删除与段落标记修订的 DOCX 写出与读取。Verifies writing and reading tracked insertions, deletions and paragraph-mark revisions.</summary>
        /// <param name="path">输出 .docx 路径。The output .docx path.</param>
        private static void RunWordTrackedChangeTest(string path)
        {
            OfficeDocument document = BuildRevisionDocument();
            DocxWriter.Write(document, path);
            ValidateOoxmlPackage(path, "word/document.xml");

            string xml = ReadPackageEntryText(path, "word/document.xml");
            RequireContains(xml, "<w:ins ");
            RequireContains(xml, "<w:del ");
            RequireContains(xml, "<w:delText");
            RequireContains(xml, "w:author=\"张三\"");
            RequireContains(xml, "w:author=\"李四\"");

            OfficeDocument loaded = DocxReader.Load(path);
            Paragraph first = FindWordParagraph(loaded, "可见开头");
            Require(first != null, "revision fixture paragraph is readable");
            RequireContains(first.GetText(), "新增内容");
            Require(first.GetText().IndexOf("已删除内容", StringComparison.Ordinal) < 0,
                "deleted content is hidden from the visible text");
            Require(first.DeletedInlines.Count == 1, "deleted content is kept separately");
            RequireContains(first.GetDeletedText(), "已删除内容");
            Require(first.DeletedInlines[0].Revision != null
                && first.DeletedInlines[0].Revision.Type == RevisionType.Delete,
                "deleted content carries a delete revision");
            Require(first.DeletedInlines[0].Revision.Author == "李四", "deleted revision author round trips");

            TextRun inserted = null;
            for (int i = 0; i < first.Inlines.Count; i++)
            {
                TextRun candidate = first.Inlines[i] as TextRun;
                if (candidate != null && candidate.Text == "新增内容") { inserted = candidate; break; }
            }
            Require(inserted != null && inserted.Revision != null, "inserted run carries a revision");
            Require(inserted.Revision.Type == RevisionType.Insert, "inserted revision type round trips");
            Require(inserted.Revision.Author == "张三", "inserted revision author round trips");
            Require(inserted.Revision.Date.HasValue
                && inserted.Revision.Date.Value.ToUniversalTime() == new DateTime(2026, 10, 20, 8, 30, 0, DateTimeKind.Utc),
                "inserted revision timestamp round trips");

            Paragraph second = FindWordParagraph(loaded, "整段插入");
            Require(second != null && second.Revision != null
                && second.Revision.Type == RevisionType.Insert,
                "paragraph mark insertion round trips");
            Paragraph third = FindWordParagraph(loaded, "段落标记删除");
            Require(third != null && third.Revision != null
                && third.Revision.Type == RevisionType.Delete,
                "paragraph mark deletion round trips");
        }

        /// <summary>验证修订文档再次写出后仍然保留修订结构与作者信息。Verifies a revision document keeps its revision structure and authors after another write.</summary>
        /// <param name="sourcePath">首轮写出的 .docx。The first-round .docx.</param>
        /// <param name="path">第二轮写出的 .docx。The second-round .docx.</param>
        private static void RunWordTrackedChangeRoundTripTest(string sourcePath, string path)
        {
            OfficeDocument loaded = DocxReader.Load(sourcePath);
            DocxWriter.Write(loaded, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            string xml = ReadPackageEntryText(path, "word/document.xml");
            RequireContains(xml, "<w:ins ");
            RequireContains(xml, "<w:del ");
            RequireContains(xml, "<w:delText");

            OfficeDocument reloaded = DocxReader.Load(path);
            Paragraph first = FindWordParagraph(reloaded, "可见开头");
            Require(first != null && first.DeletedInlines.Count == 1, "deleted content survives a second round trip");
            RequireContains(first.GetDeletedText(), "已删除内容");
            RequireContains(first.GetText(), "新增内容");
        }

        /// <summary>验证接受/拒绝全部修订的语义与修订计数。Verifies the accept/reject semantics and revision counting.</summary>
        /// <param name="path">修订测试文档。The revision fixture document.</param>
        private static void RunWordAcceptRejectTest(string path)
        {
            OfficeDocument loaded = DocxReader.Load(path);
            Require(WordRevisions.Count(loaded) == 4, "revision count includes runs, deleted content and paragraph marks, got " + WordRevisions.Count(loaded));

            OfficeDocument accepted = DocxReader.Load(path);
            int acceptedCount = WordRevisions.AcceptAll(accepted);
            Require(acceptedCount == 4, "accept handles every revision, got " + acceptedCount);
            Require(WordRevisions.Count(accepted) == 0, "accept clears every revision marker");
            Paragraph acceptedFirst = FindWordParagraph(accepted, "可见开头");
            Require(acceptedFirst != null, "accepted paragraph survives");
            RequireContains(acceptedFirst.GetText(), "新增内容");
            Require(acceptedFirst.GetText().IndexOf("已删除内容", StringComparison.Ordinal) < 0,
                "accept drops deleted content");
            Require(acceptedFirst.DeletedInlines.Count == 0, "accept clears the deleted-content list");

            OfficeDocument rejected = DocxReader.Load(path);
            int rejectedCount = WordRevisions.RejectAll(rejected);
            Require(rejectedCount == 4, "reject handles every revision, got " + rejectedCount);
            Require(WordRevisions.Count(rejected) == 0, "reject clears every revision marker");
            Paragraph rejectedFirst = FindWordParagraph(rejected, "可见开头");
            Require(rejectedFirst != null, "rejected paragraph survives");
            RequireContains(rejectedFirst.GetText(), "已删除内容");
            Require(rejectedFirst.GetText().IndexOf("新增内容", StringComparison.Ordinal) < 0,
                "reject removes inserted content");
            Require(rejectedFirst.DeletedInlines.Count == 0, "reject restores and clears deleted content");
        }

        /// <summary>验证段落深拷贝保留修订与删除内容。Verifies a paragraph deep copy keeps revisions and deleted content.</summary>
        /// <param name="path">修订测试文档。The revision fixture document.</param>
        private static void RunWordCloneRevisionTest(string path)
        {
            OfficeDocument loaded = DocxReader.Load(path);
            Paragraph source = FindWordParagraph(loaded, "可见开头");
            Require(source != null, "clone source paragraph exists");
            Paragraph clone = source.Clone();
            Require(clone.GetText() == source.GetText(), "clone keeps visible text");
            Require(clone.DeletedInlines.Count == source.DeletedInlines.Count, "clone keeps deleted content");
            Require(clone.DeletedInlines.Count > 0
                && clone.DeletedInlines[0].Revision != null
                && clone.DeletedInlines[0].Revision.Author == "李四",
                "clone keeps deleted revision metadata");
            Require(clone.DeletedInlines[0].Revision != source.DeletedInlines[0].Revision,
                "clone deep-copies revision metadata");
        }

        /// <summary>构造包含插入、删除与段落标记修订的测试文档。Builds the fixture document with insertions, deletions and paragraph-mark revisions.</summary>
        /// <returns>文档模型。The document model.</returns>
        private static OfficeDocument BuildRevisionDocument()
        {
            TextFormat format = new TextFormat("Microsoft YaHei", 11);
            DateTime insertDate = new DateTime(2026, 10, 20, 8, 30, 0, DateTimeKind.Utc);
            DateTime deleteDate = new DateTime(2026, 10, 20, 9, 0, 0, DateTimeKind.Utc);

            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();

            Paragraph first = new Paragraph();
            first.AddText("可见开头 ", format);
            TextRun inserted = new TextRun("新增内容", format);
            inserted.Revision = new RevisionInfo(RevisionType.Insert, "张三", insertDate);
            first.Inlines.Add(inserted);
            first.AddText(" 可见结尾", format);
            TextRun deleted = new TextRun("已删除内容", format);
            deleted.Revision = new RevisionInfo(RevisionType.Delete, "李四", deleteDate);
            first.DeletedInlines.Add(deleted);
            section.Blocks.Add(first);

            Paragraph second = new Paragraph("整段插入");
            second.Revision = new RevisionInfo(RevisionType.Insert, "张三", insertDate);
            section.Blocks.Add(second);

            Paragraph third = new Paragraph("段落标记删除");
            third.Revision = new RevisionInfo(RevisionType.Delete, "李四", deleteDate);
            section.Blocks.Add(third);
            return document;
        }

        /// <summary>验证 SmartArt 数据部件解析：节点数量与文本提取，并能区分无 SmartArt 的普通文档。Verifies SmartArt data-part parsing (node count and text extraction) and that an ordinary document has no diagrams.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunSmartArtInspectionTest(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-smartart.docx"));
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteTestZipEntry(archive, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                    + "</Types>");
                WriteTestZipEntry(archive, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                    + "</Relationships>");
                WriteTestZipEntry(archive, "word/document.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                    + " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                    + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\""
                    + " xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">"
                    + "<w:body><w:p><w:r><w:drawing><a:graphic><a:graphicData"
                    + " uri=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">"
                    + "<dgm:relIds r:dm=\"rId2\" r:lo=\"rId3\" r:qs=\"rId4\" r:cs=\"rId5\"/>"
                    + "</a:graphicData></a:graphic></w:drawing></w:r></w:p></w:body></w:document>");
                WriteTestZipEntry(archive, "word/_rels/document.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData\" Target=\"diagrams/data1.xml\"/>"
                    + "</Relationships>");
                WriteTestZipEntry(archive, "word/diagrams/data1.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\""
                    + " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                    + "<dgm:ptLst>"
                    + "<dgm:pt modelId=\"1\" type=\"doc\"><dgm:t><a:bodyPr/><a:p><a:r><a:t>年度总结</a:t></a:r></a:p></dgm:t></dgm:pt>"
                    + "<dgm:pt modelId=\"2\" type=\"node\"><dgm:t><a:bodyPr/><a:p><a:r><a:t>第一季度</a:t></a:r></a:p></dgm:t></dgm:pt>"
                    + "<dgm:pt modelId=\"3\" type=\"node\"><dgm:t><a:bodyPr/><a:p><a:r><a:t>第二季度</a:t></a:r></a:p></dgm:t></dgm:pt>"
                    + "</dgm:ptLst></dgm:dataModel>");
            }

            IList<WordSmartArt> diagrams = WordSmartArtReader.Read(path);
            Require(diagrams.Count == 1, "smart art inspection finds one diagram, got " + diagrams.Count);
            Require(diagrams[0].PartPath == "word/diagrams/data1.xml", "smart art data part resolves");
            Require(diagrams[0].SourcePart == "word/document.xml", "smart art source part is recorded");
            Require(diagrams[0].NodeCount == 3, "smart art node count matches, got " + diagrams[0].NodeCount);
            RequireContains(diagrams[0].Text, "年度总结");
            RequireContains(diagrams[0].Text, "第一季度");
            RequireContains(diagrams[0].Text, "第二季度");

            IList<WordSmartArt> none = WordSmartArtReader.Read(CreateSmartArtFreeDocument(outputDirectory));
            Require(none.Count == 0, "smart art reader reports no diagrams for an ordinary document");
        }

        /// <summary>创建一个不含 SmartArt 的普通文档并返回路径。Creates an ordinary document without SmartArt and returns its path.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        /// <returns>普通 .docx 路径。The ordinary .docx path.</returns>
        private static string CreateSmartArtFreeDocument(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-smartart-empty.docx"));
            OfficeDocument document = new OfficeDocument();
            document.AddSection().AddParagraph("普通段落");
            DocxWriter.Write(document, path);
            return path;
        }

        /// <summary>验证数字签名元数据读取：签名存在性、算法、引用部件与无证书时的容错。Verifies digital-signature metadata reading: presence, algorithms, referenced parts and tolerance of a missing certificate.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunDigitalSignatureInspectionTest(string outputDirectory)
        {
            string signedPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-signed.docx"));
            const string signatureXml =
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<ds:Signature xmlns:ds=\"http://www.w3.org/2000/09/xmldsig#\">"
                + "<ds:SignedInfo>"
                + "<ds:CanonicalizationMethod Algorithm=\"http://www.w3.org/TR/2001/REC-xml-c14n-20010315\"/>"
                + "<ds:SignatureMethod Algorithm=\"http://www.w3.org/2001/04/xmldsig-more#rsa-sha256\"/>"
                + "<ds:Reference URI=\"\"><ds:DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\"/><ds:DigestValue>AAAA</ds:DigestValue></ds:Reference>"
                + "<ds:Reference URI=\"http://schemas.openxmlformats.org/package/2006/relationships\"><ds:DigestMethod Algorithm=\"http://www.w3.org/2001/04/xmlenc#sha256\"/><ds:DigestValue>BBBB</ds:DigestValue></ds:Reference>"
                + "</ds:SignedInfo>"
                + "<ds:SignatureValue>CCCC</ds:SignatureValue>"
                + "<ds:KeyInfo/>"
                + "</ds:Signature>";
            using (FileStream stream = new FileStream(signedPath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteTestZipEntry(archive, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "</Types>");
                WriteTestZipEntry(archive, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
                WriteTestZipEntry(archive, "word/document.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>");
                WriteTestZipEntry(archive, "_xmlsignatures/origin.sigs", string.Empty);
                WriteTestZipEntry(archive, "_xmlsignatures/_rels/origin.sigs.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature\" Target=\"sig1.xml\"/>"
                    + "</Relationships>");
                WriteTestZipEntry(archive, "_xmlsignatures/sig1.xml", signatureXml);
            }

            WordDigitalSignatureReport report = WordDigitalSignatureReader.Read(signedPath);
            Require(report.IsSigned, "signature origin part is detected");
            Require(report.Signatures.Count == 1, "one signature is parsed, got " + report.Signatures.Count);
            WordDigitalSignature signature = report.Signatures[0];
            Require(signature.PartPath == "_xmlsignatures/sig1.xml", "signature part path is recorded");
            RequireContains(signature.SignatureMethod, "rsa-sha256");
            RequireContains(signature.DigestMethod, "sha256");
            RequireContains(signature.CanonicalizationMethod, "xml-c14n");
            Require(signature.SignedParts.Count == 2, "both signed references are captured");
            Require(signature.SignedParts[0] == string.Empty, "package reference URI is kept");
            RequireContains(signature.SignedParts[1], "relationships");
            Require(signature.HasSignatureValue, "signature value presence is reported");
            Require(signature.SignerSubject == null, "missing certificate yields a null subject");

            string unsignedPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-unsigned.docx"));
            using (FileStream stream = new FileStream(unsignedPath, FileMode.Create, FileAccess.Write, FileShare.None))
            using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create))
            {
                WriteTestZipEntry(archive, "word/document.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body/></w:document>");
            }
            WordDigitalSignatureReport unsigned = WordDigitalSignatureReader.Read(unsignedPath);
            Require(!unsigned.IsSigned && unsigned.Signatures.Count == 0,
                "an unsigned package reports no signatures");
        }

        /// <summary>以 UTF-8 文本向测试 ZIP 包写入一个条目。Writes one UTF-8 text entry into a test ZIP package.</summary>
        /// <param name="archive">目标 ZIP 包。The target ZIP package.</param>
        /// <param name="entryName">条目路径。The entry path.</param>
        /// <param name="text">条目文本。The entry text.</param>
        private static void WriteTestZipEntry(ZipArchive archive, string entryName, string text)
        {
            ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using (Stream stream = entry.Open())
            {
                byte[] data = Encoding.UTF8.GetBytes(text);
                stream.Write(data, 0, data.Length);
            }
        }
    }
}
