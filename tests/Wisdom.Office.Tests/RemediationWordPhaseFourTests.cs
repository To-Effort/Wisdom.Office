using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Wisdom.Office.Core;
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
