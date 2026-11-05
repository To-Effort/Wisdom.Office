using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Wisdom.Excel;
using Wisdom.Office.Core;
using Wisdom.Ppt;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 4 的 Word 条目：异常契约违规与写出器不判空。
    /// Phase-four Word entries of remediation design episode two: exception-contract violations and
    /// writers that dereference optional model state.
    /// </summary>
    /// <remarks>
    /// 本文件的条目都以「调用方按声明的异常类型兜错」为前提，所以断言钉的是异常类型而不是返回值。
    /// Every entry here assumes callers catch the documented exception type, so the assertions pin exception
    /// types rather than return values.
    /// </remarks>
    internal static partial class Program
    {
        /// <summary>Word Phase 4 入口。Entry point for the Word half of phase four.</summary>
        private static void RunRemediationWordPhaseFourTest()
        {
            RunOdtContentXmlContractTest();
            RunRegexTimeoutAtomicReplaceTest();
            RunOversizedOdtImageIsSkippedTest();
            RunRtfNonBmpCodeUnitTest();
            RunWritersTolerateNullModelFieldsTest();
            RunWritersNeverTruncateTheOriginalTest();
            RunCorruptHeaderPartIsReportedTest();

            Console.WriteLine("[remediation] word phase-four contracts ok");
        }

        /// <summary>
        /// WE2：<c>content.xml</c> 不是合法 XML 时，<c>OdtReader.Load</c> 与门面必须报
        /// <see cref="InvalidDataException"/>，不得让 <see cref="XmlException"/> 穿出契约，也不得丢掉内层异常。
        /// WE2: when <c>content.xml</c> is not well-formed XML, <c>OdtReader.Load</c> and the façade must report
        /// <see cref="InvalidDataException"/> instead of leaking <see cref="XmlException"/> past the contract,
        /// and the inner exception must survive.
        /// </summary>
        private static void RunOdtContentXmlContractTest()
        {
            string validContent = ValidOdtContent("Quiet");
            string brokenContent = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\">"
                + "<office:body><office:text><<<</office:text></office:body></office:document-content>";

            // 控制一：正文合法、styles.xml 坏掉。样式部件一直是按支降级读入的，
            // 这一支必须照常出文，否则「只有主部件致命」这条界就钉不住。
            // Control one: a valid body with a broken styles.xml. The styles part has always degraded,
            // so this branch must still yield text — otherwise the "only the main part is fatal" boundary
            // is not pinned at all.
            OfficeDocument degraded = OdtReader.Load(BuildOdtPackage(validContent, "<<< not xml >>>"));
            Require(degraded.GetPlainText().Contains("Quiet"),
                "a broken styles.xml stopped the document body, text reads '" + degraded.GetPlainText() + "'");

            // 控制二：缺 content.xml 走的就是同一条契约（本来就抛 InvalidDataException），
            // 用它确认这一族确实被当成「坏文件」而不是「读不出来」。
            // Control two: a missing body already honors this contract, confirming the family is treated
            // as a corrupt file rather than an unreadable one.
            RequireThrows<InvalidDataException>(() => OdtReader.Load(BuildOdtPackage(null, null)),
                "a package without content.xml must report InvalidDataException");

            RequireThrows<InvalidDataException>(() => OdtReader.Load(BuildOdtPackage(brokenContent, null)),
                "a malformed content.xml must reach the caller as InvalidDataException");

            // 门面这一路是条目点名的漏网路径：调用方按 InvalidDataException 兜错，
            // XmlException 一旦穿出就直接落到业务上。
            // The façade route is the leak the entry names: callers catch InvalidDataException, so an
            // escaping XmlException lands on the business code untouched.
            RequireThrows<InvalidDataException>(() => WordFormatConverter.Load(BuildOdtPackage(brokenContent, null)),
                "the façade must not leak the parser's exception type for a malformed content.xml");

            // 包装不能只换个类型就把现场也换掉：内层异常是排查「哪个部件的哪一列」的唯一线索。
            // Rewrapping must not replace the scene with the type: the inner exception is the only lead
            // for which part failed at which line.
            try
            {
                OdtReader.Load(BuildOdtPackage(brokenContent, null));
                Require(false, "a malformed content.xml loaded without any exception");
            }
            catch (InvalidDataException wrapped)
            {
                Require(wrapped.InnerException is XmlException,
                    "the wrapped InvalidDataException dropped the inner XmlException, inner is "
                    + (wrapped.InnerException == null ? "null" : wrapped.InnerException.GetType().Name));
            }
        }

        /// <summary>
        /// WE6：正则匹配超时必须转成方法已声明的 <see cref="ArgumentException"/>，抛出时文档一个字都不能被改动。
        /// WE6: a regular-expression match timeout must surface as the declared
        /// <see cref="ArgumentException"/>, and the document must be untouched when it escapes.
        /// </summary>
        private static void RunRegexTimeoutAtomicReplaceTest()
        {
            string bomb = new string('a', 40);

            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();
            Paragraph hit = new Paragraph("ab");
            Paragraph timedOut = new Paragraph(bomb);
            section.Blocks.Add(hit);
            section.Blocks.Add(timedOut);

            WordFindReplaceOptions options = new WordFindReplaceOptions();
            options.UseRegex = true;

            ArgumentException declared = null;
            try
            {
                WordFindReplace.Replace(document, "(a+)+b", "X", options);
                Require(false, "a catastrophic regex pattern finished without any exception");
            }
            catch (ArgumentException expected)
            {
                declared = expected;
            }
            catch (Exception escaped)
            {
                Require(false, "a catastrophic regex escaped as '" + escaped.GetType().Name
                    + "'; Replace declares only ArgumentNullException and ArgumentException");
            }

            Require(declared.InnerException is RegexMatchTimeoutException,
                "the ArgumentException dropped the inner timeout, inner is "
                + (declared.InnerException == null ? "null" : declared.InnerException.GetType().Name));

            // 半替换是本条的另一半：段落一在段落二超时之前就已经改完了。
            // The half-replaced document is the other half of this entry: the first paragraph had already
            // been rewritten before the second one timed out.
            Require(hit.GetText() == "ab",
                "the timeout left the first paragraph replaced, it reads '" + hit.GetText() + "'");
            Require(timedOut.GetText() == bomb, "the paragraph that timed out was modified");

            // 控制一：同一份文档换一个不回溯的模式，替换照常生效——超时没把文档留在坏状态里。
            // Control one: a non-backtracking pattern over the same document still replaces, so the
            // timeout left nothing damaged.
            int count = WordFindReplace.Replace(document, "b", "X", options);
            Require(count == 1 && hit.GetText() == "aX",
                "a safe regex after the timeout replaced " + count + " match(es), the first paragraph "
                + "reads '" + hit.GetText() + "'");

            // 控制二：FirstOnly 走正则这一路仍然只动文档里的第一处。
            // Control two: FirstOnly over the regex route still replaces exactly one occurrence.
            OfficeDocument firstOnly = new OfficeDocument();
            Section scope = firstOnly.AddSection();
            Paragraph one = new Paragraph("ab");
            Paragraph two = new Paragraph("cb");
            scope.Blocks.Add(one);
            scope.Blocks.Add(two);
            WordFindReplaceOptions once = new WordFindReplaceOptions();
            once.UseRegex = true;
            once.FirstOnly = true;
            count = WordFindReplace.Replace(firstOnly, "b", "X", once);
            Require(count == 1 && one.GetText() == "aX" && two.GetText() == "cb",
                "FirstOnly under the regex route replaced " + count + " match(es): '" + one.GetText()
                + "' / '" + two.GetText() + "'");
        }

        /// <summary>
        /// WE9：一张超限的内嵌图片不该让整篇 ODT 读失败——正文必须照常读回，丢的只是那张图。
        /// HTML 读端一直是「丢图不丢文档」，ODT 过去对单张超限图片抛 <see cref="InvalidDataException"/>，
        /// 把一篇结构完全合法的扫描稿连正文一起废掉。
        /// WE9: one oversized embedded picture must not fail the whole ODT read — the body must still
        /// load and only the picture is dropped. The HTML reader has always followed "drop the picture,
        /// keep the document"; ODT used to throw <see cref="InvalidDataException"/> for a single
        /// oversized image and throw away a structurally valid scanned document along with it.
        /// </summary>
        private static void RunOversizedOdtImageIsSkippedTest()
        {
            // 控制：同一份包换成一张正常大小的 GIF，图片必须照常读回，否则红的会是夹具而不是本条。
            // Control: the same package with a normal-sized GIF must keep its picture, otherwise the
            // fixture rather than the entry under test is what fails.
            OfficeDocument small = OdtReader.Load(
                BuildOdtPackageWithImage("正文可用", "Pictures/small.gif", CreateTestGif(8, 6)));
            Require(ContainsText(small, "正文可用"), "the control ODT lost its body text");
            Require(FindFirstImage(small) != null, "the control ODT lost its normal-sized picture");

            OfficeDocument document = OdtReader.Load(
                BuildOdtPackageWithImage("正文可用", "Pictures/big.gif", null));
            Require(ContainsText(document, "正文可用"),
                "an oversized embedded picture took the whole document down with it");
            Require(FindFirstImage(document) == null,
                "an embedded picture past the 64 MB cap was kept in the model");
        }

        /// <summary>
        /// WE11：RTF 的 <c>\uN</c> 只接受有符号 16 位。增补平面字符（emoji）在 .NET 里是一对代理项，
        /// 码元 55357 直接写成正数就越界了；规范写法是减去 65536 的负值，读端本来也按「负值加 65536」还原，
        /// 所以这条只改输出字节，往返行为不变。
        /// WE11: RTF's <c>\uN</c> takes a signed 16-bit value. A supplementary-plane character is a surrogate
        /// pair in .NET, and writing code unit 55357 as a positive number is out of range; the spec form is
        /// that value minus 65536. The reader already adds 65536 back for negatives, so this changes only the
        /// emitted bytes and the round trip stays the same.
        /// </summary>
        private static void RunRtfNonBmpCodeUnitTest()
        {
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection(PageSetup.A4());
            section.Blocks.Add(new Paragraph("表情 \U0001F600 与中文"));

            byte[] rtf = WordFormatConverter.ToArray(document, WordDocumentFormat.Rtf);
            string emitted = Encoding.ASCII.GetString(rtf);

            Require(emitted.IndexOf("\\u55357", StringComparison.Ordinal) < 0,
                "a surrogate code unit was written as the positive 55357, outside RTF's signed 16-bit range");
            Require(emitted.IndexOf("\\u-10179", StringComparison.Ordinal) >= 0,
                "the high surrogate did not come out as the signed value -10179");
            Require(emitted.IndexOf("\\u-8704", StringComparison.Ordinal) >= 0,
                "the low surrogate did not come out as the signed value -8704");

            OfficeDocument back = WordFormatConverter.Load(rtf, WordDocumentFormat.Rtf);
            Require(ContainsText(back, "表情 \U0001F600 与中文"),
                "the emoji did not survive the RTF round trip: " + emitted);
        }

        /// <summary>
        /// WE10：模型上的 `TextRun.Text`、`TextRun.Format` 与 `Table.Borders` 都是公开可写属性，
        /// 宿主把它们置空之后，四个写出器必须给出同一结果——照常出文，而不是 DOCX/RTF 崩、HTML/ODT 正常。
        /// WE10: `TextRun.Text`, `TextRun.Format` and `Table.Borders` all expose public setters. Once a host
        /// nulls them, all four writers must agree: produce a document, instead of DOCX/RTF throwing while
        /// HTML/ODT keep working.
        /// </summary>
        private static void RunWritersTolerateNullModelFieldsTest()
        {
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection(PageSetup.A4());

            Paragraph paragraph = new Paragraph();
            TextRun emptyText = new TextRun("x");
            emptyText.Text = null;
            paragraph.Inlines.Add(emptyText);
            TextRun noFormat = new TextRun("只留文本");
            noFormat.Format = null;
            paragraph.Inlines.Add(noFormat);
            section.Blocks.Add(paragraph);

            Table table = new Table();
            table.Columns.Add(new TableColumn(120));
            table.AddRow("无边框表格");
            table.Borders = null;
            section.Blocks.Add(table);

            WordDocumentFormat[] formats = new WordDocumentFormat[]
            {
                WordDocumentFormat.Docx, WordDocumentFormat.Rtf, WordDocumentFormat.Html, WordDocumentFormat.Odt
            };
            for (int i = 0; i < formats.Length; i++)
            {
                byte[] bytes = null;
                try
                {
                    bytes = WordFormatConverter.ToArray(document, formats[i]);
                }
                catch (Exception exception)
                {
                    Require(false, "the " + formats[i] + " writer failed on a null model field: "
                        + exception.GetType().Name + ": " + exception.Message);
                }
                Require(bytes != null && bytes.Length > 0,
                    "the " + formats[i] + " writer produced no output for a null model field");
            }
        }

        /// <summary>
        /// E7：目标文件必须写完整再替换。过去 `DocxWriter` / `XlsxWriter` / `PptxWriter` 都直接用
        /// `FileMode.Create` 打开目标，`Write(null, path)` 在参数校验失败之前就把用户原文件截成了 0 字节。
        /// 现在失败被限制在同目录临时文件上：目标保持原样，也不留 `.tmp` 残骸。
        /// E7: the target must be complete before it replaces anything. The writers used to open the target
        /// with FileMode.Create, so Write(null, path) truncated the user's file to zero bytes before the
        /// argument check even ran; now a failure stays in a sibling temporary file, the target keeps its
        /// bytes and no .tmp litter is left behind.
        /// </summary>
        private static void RunWritersNeverTruncateTheOriginalTest()
        {
            string directory = Path.Combine(Path.GetTempPath(), "wisdom-atomic-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                RequireFailedWriteKeepsTarget(directory, "docx", delegate(string path) { DocxWriter.Write(null, path); });
                RequireFailedWriteKeepsTarget(directory, "xlsx", delegate(string path) { XlsxWriter.Write(null, path); });
                RequireFailedWriteKeepsTarget(directory, "pptx", delegate(string path) { PptxWriter.Write(null, path); });
            }
            finally
            {
                try { Directory.Delete(directory, true); }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }

        /// <summary>断言「失败的写入既没有改动目标文件，也没有留下临时文件」。Requires that a failed write neither touched the target nor left a temporary file behind.</summary>
        /// <param name="directory">夹具目录。The fixture directory.</param>
        /// <param name="extension">目标扩展名。The target extension.</param>
        /// <param name="write">触发失败的写入委托。The write delegate that must fail.</param>
        private static void RequireFailedWriteKeepsTarget(string directory, string extension, Action<string> write)
        {
            string path = Path.Combine(directory, "target." + extension);
            byte[] sentinel = Encoding.UTF8.GetBytes("ORIGINAL-" + extension);
            File.WriteAllBytes(path, sentinel);

            bool rejected = false;
            try
            {
                write(path);
            }
            catch (ArgumentNullException)
            {
                rejected = true;
            }
            Require(rejected, "the " + extension + " writer accepted a null model instead of rejecting it");

            byte[] after = File.ReadAllBytes(path);
            Require(BytesEqual(after, sentinel),
                "a failed " + extension + " write left the original file with " + after.Length
                + " byte(s) instead of the " + sentinel.Length + " it had");

            string[] leftovers = Directory.GetFiles(directory, "." + Path.GetFileName(path) + ".*.tmp");
            Require(leftovers.Length == 0,
                "a failed " + extension + " write left " + leftovers.Length + " temporary file(s) behind");
        }

        /// <summary>
        /// WE7：`word/header1.xml` 存在但内容不是合法 XML 时，读端过去把「解析失败」与「没有页眉」都表达成
        /// `null`，页眉被静默丢弃且调用方无从分辨。现在这类降级记进 `OfficeDocument.Warnings`：
        /// 文档照常加载、正文可读，那条告警点名是哪个部件。
        /// WE7: when `word/header1.xml` exists but is not well-formed XML, the reader used to express both
        /// "parse failed" and "no header" as `null`, dropping the header silently with no way for the caller
        /// to tell. Such downgrades now land in `OfficeDocument.Warnings`: the document still loads, the body
        /// is readable, and the notice names the part that was dropped.
        /// </summary>
        private static void RunCorruptHeaderPartIsReportedTest()
        {
            OfficeDocument source = new OfficeDocument();
            Section section = source.AddSection(PageSetup.A4());
            section.Header.Blocks.Add(new Paragraph("页眉内容"));
            section.Blocks.Add(new Paragraph("正文内容"));

            byte[] clean = WordFormatConverter.ToArray(source, WordDocumentFormat.Docx);
            byte[] corrupted = RewriteZipEntry(clean, "word/header1.xml",
                Encoding.UTF8.GetBytes("<w:hdr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><<<<"));

            OfficeDocument loaded = WordFormatConverter.Load(corrupted, WordDocumentFormat.Docx);
            Require(ContainsText(loaded, "正文内容"),
                "a corrupt header must not take the body down with it");
            Require(loaded.Warnings.Count == 1,
                "expected exactly one downgrade notice for the corrupt header, got " + loaded.Warnings.Count);
            Require(loaded.Warnings[0].IndexOf("header1.xml", StringComparison.OrdinalIgnoreCase) >= 0,
                "the notice does not name the part that was dropped: " + loaded.Warnings[0]);

            // 控制：未损坏的同一份文档不得产生任何告警——「缺失」与「坏了」必须真的分开。
            // Control: the same document with an intact header must produce no notice at all, so "missing"
            // and "broken" really are distinguishable.
            OfficeDocument ok = WordFormatConverter.Load(clean, WordDocumentFormat.Docx);
            Require(ok.Warnings.Count == 0, "a clean document reported " + ok.Warnings.Count + " downgrade notice(s)");
        }

        /// <summary>在内存里重打包，替换指定条目的内容；其余条目原样复制。 Repackages in memory, replacing the named entry and copying every other entry as-is.</summary>
        /// <param name="package">原包字节。The original package bytes.</param>
        /// <param name="entryName">要替换的条目名。The entry to replace.</param>
        /// <param name="content">新的条目内容。The new entry content.</param>
        /// <returns>重打包后的字节。The repackaged bytes.</returns>
        private static byte[] RewriteZipEntry(byte[] package, string entryName, byte[] content)
        {
            using (MemoryStream buffer = new MemoryStream())
            {
                using (ZipArchive source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read))
                using (ZipArchive target = new ZipArchive(buffer, ZipArchiveMode.Create, true))
                {
                    foreach (ZipArchiveEntry entry in source.Entries)
                    {
                        ZipArchiveEntry copy = target.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                        using (Stream output = copy.Open())
                        {
                            if (string.Equals(entry.FullName, entryName, StringComparison.OrdinalIgnoreCase))
                            {
                                output.Write(content, 0, content.Length);
                            }
                            else
                            {
                                using (Stream input = entry.Open())
                                {
                                    input.CopyTo(output);
                                }
                            }
                        }
                    }
                }
                return buffer.ToArray();
            }
        }

        /// <summary>正文加一个引用图片的 <c>draw:frame</c>。 A body paragraph followed by a <c>draw:frame</c> that references a picture.</summary>
        /// <param name="text">正文文本。The body text.</param>
        /// <param name="href">图片在包内的路径。The in-package path of the picture.</param>
        /// <returns>content.xml 文本。The content.xml text.</returns>
        private static string OdtContentWithImage(string text, string href)
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<office:document-content"
                + " xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\""
                + " xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\""
                + " xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\""
                + " xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\""
                + " xmlns:xlink=\"http://www.w3.org/1999/xlink\">"
                + "<office:body><office:text><text:p>" + text
                + "<draw:frame svg:width=\"2cm\" svg:height=\"2cm\">"
                + "<draw:image xlink:href=\"" + href + "\" xlink:type=\"simple\""
                + " xlink:show=\"embed\" xlink:actuate=\"onLoad\"/>"
                + "</draw:frame></text:p></office:text></office:body></office:document-content>";
        }

        /// <summary>
        /// 拼一个带图片部件的 ODT 包；<paramref name="image"/> 为 null 时写入一张刚超过 64 MB 上限的图，
        /// 内容按块流式写入，夹具自身只占一个 64 KB 缓冲。
        /// Assembles an ODT package with a picture part; a null <paramref name="image"/> writes one just past
        /// the 64 MB cap, streamed in blocks so the fixture itself holds a single 64 KB buffer.
        /// </summary>
        /// <param name="text">正文文本。The body text.</param>
        /// <param name="href">图片在包内的路径。The in-package path of the picture.</param>
        /// <param name="image">图片字节，null 表示超限图。The picture bytes; null means the oversized one.</param>
        /// <returns>ODT 字节。The ODT bytes.</returns>
        private static byte[] BuildOdtPackageWithImage(string text, string href, byte[] image)
        {
            using (MemoryStream buffer = new MemoryStream())
            {
                using (ZipArchive archive = new ZipArchive(buffer, ZipArchiveMode.Create, true))
                {
                    WriteOdtTextEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text");
                    WriteOdtTextEntry(archive, "META-INF/manifest.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                        + "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">"
                        + "<manifest:file-entry manifest:full-path=\"/\" "
                        + "manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>"
                        + "</manifest:manifest>");
                    WriteOdtTextEntry(archive, "content.xml", OdtContentWithImage(text, href));

                    ZipArchiveEntry picture = archive.CreateEntry(href, CompressionLevel.Fastest);
                    using (Stream target = picture.Open())
                    {
                        if (image != null)
                        {
                            target.Write(image, 0, image.Length);
                        }
                        else
                        {
                            byte[] chunk = new byte[65536];
                            long remaining = 64L * 1024 * 1024 + 64;
                            while (remaining > 0)
                            {
                                int count = (int)Math.Min(chunk.Length, remaining);
                                target.Write(chunk, 0, count);
                                remaining -= count;
                            }
                        }
                    }
                }
                return buffer.ToArray();
            }
        }

        /// <summary>往 ODT 包里写一个 UTF-8 文本条目。 Writes one UTF-8 text entry into an ODT package.</summary>
        /// <param name="archive">目标压缩包。The target archive.</param>
        /// <param name="name">条目名。The entry name.</param>
        /// <param name="value">文本内容。The text content.</param>
        private static void WriteOdtTextEntry(ZipArchive archive, string name, string value)
        {
            ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream target = entry.Open())
            using (StreamWriter writer = new StreamWriter(target, new UTF8Encoding(false)))
            {
                writer.Write(value);
            }
        }

        /// <summary>判断文档的段落文本里是否含指定串（含表格单元格内的段落）。 Reports whether any paragraph text contains the given string, including paragraphs inside table cells.</summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="needle">待查找的文本。The text to look for.</param>
        /// <returns>命中返回 true。True when found.</returns>
        private static bool ContainsText(OfficeDocument document, string needle)
        {
            for (int s = 0; s < document.Sections.Count; s++)
            {
                foreach (Block block in document.Sections[s].Blocks)
                {
                    if (ContainsText(block, needle)) return true;
                }
            }
            return false;
        }

        /// <summary>在单个块里查找文本，表格递归进单元格。 Looks for the text inside one block, recursing into table cells.</summary>
        /// <param name="block">待查块。The block to search.</param>
        /// <param name="needle">待查找的文本。The text to look for.</param>
        /// <returns>命中返回 true。True when found.</returns>
        private static bool ContainsText(Block block, string needle)
        {
            Paragraph paragraph = block as Paragraph;
            if (paragraph != null)
            {
                return paragraph.GetText().IndexOf(needle, StringComparison.Ordinal) >= 0;
            }
            Table table = block as Table;
            if (table == null) return false;
            foreach (TableRow row in table.Rows)
            {
                foreach (TableCell cell in row.Cells)
                {
                    foreach (Block inner in cell.Blocks)
                    {
                        if (ContainsText(inner, needle)) return true;
                    }
                }
            }
            return false;
        }

        /// <summary>一份能读通的 ODT 正文，只带一个段落。A readable ODT body carrying a single paragraph.</summary>
        private static string ValidOdtContent(string text)
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<office:document-content"
                + " xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\""
                + " xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\">"
                + "<office:body><office:text><text:p>" + text + "</text:p></office:text></office:body>"
                + "</office:document-content>";
        }

        /// <summary>
        /// 按字节拼一个 ODT 包：mimetype、manifest 与正文；<c>content</c> 为 null 即不放 <c>content.xml</c>。
        /// Assembles an ODT package from bytes: mimetype, manifest and the body; a null <c>content</c> means
        /// no <c>content.xml</c> at all.
        /// </summary>
        private static byte[] BuildOdtPackage(string content, string styles)
        {
            List<KeyValuePair<string, string>> entries = new List<KeyValuePair<string, string>>();
            entries.Add(new KeyValuePair<string, string>("mimetype", "application/vnd.oasis.opendocument.text"));
            entries.Add(new KeyValuePair<string, string>("META-INF/manifest.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">"
                + "<manifest:file-entry manifest:full-path=\"/\" "
                + "manifest:media-type=\"application/vnd.oasis.opendocument.text\"/>"
                + "</manifest:manifest>"));
            if (content != null) entries.Add(new KeyValuePair<string, string>("content.xml", content));
            if (styles != null) entries.Add(new KeyValuePair<string, string>("styles.xml", styles));

            using (MemoryStream buffer = new MemoryStream())
            {
                using (ZipArchive archive = new ZipArchive(buffer, ZipArchiveMode.Create, true))
                {
                    foreach (KeyValuePair<string, string> pair in entries)
                    {
                        ZipArchiveEntry entry = archive.CreateEntry(pair.Key, CompressionLevel.Optimal);
                        using (Stream target = entry.Open())
                        using (StreamWriter writer = new StreamWriter(target, new UTF8Encoding(false)))
                        {
                            writer.Write(pair.Value);
                        }
                    }
                }
                return buffer.ToArray();
            }
        }
    }
}
