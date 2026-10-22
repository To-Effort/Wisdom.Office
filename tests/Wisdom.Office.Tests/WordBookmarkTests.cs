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
        /// <summary>执行书签 API 的名称校验、范围往返、文本替换、光标定位与页眉页脚测试。Runs the bookmark tests covering name rules, range round trips, text replacement, cursor navigation and header/footer scoping.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWordBookmarkTest(string outputDirectory)
        {
            RunBookmarkNameRules();
            RunBookmarkRangeRoundTrip(outputDirectory);
            RunBookmarkCrossParagraph(outputDirectory);
            RunBookmarkTextReplaceAndRemove(outputDirectory);
            RunBookmarkCursor(outputDirectory);
            RunBookmarkHeaderFooterAndGoBack(outputDirectory);
            RunBookmarkFormatMatrix(outputDirectory);
            Console.WriteLine("[bookmarks] names, ranges, cross-paragraph, text, cursor, header/footer, _GoBack, format matrix ok");
        }

        /// <summary>验证范围书签在 RTF/ODT/HTML 中的往返表现。Verifies how a range bookmark survives RTF, ODT and HTML round trips.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunBookmarkFormatMatrix(string outputDirectory)
        {
            WordDocumentBuilder builder = new WordDocumentBuilder();
            builder.Write("前 ");
            builder.StartBookmark("Mid");
            builder.Write("中段");
            builder.EndBookmark("Mid");
            builder.Write(" 后");
            OfficeDocument source = builder.Document;

            WordDocumentFormat[] formats = new WordDocumentFormat[]
            {
                WordDocumentFormat.Rtf,
                WordDocumentFormat.Odt,
                WordDocumentFormat.Html
            };
            for (int i = 0; i < formats.Length; i++)
            {
                WordDocumentFormat format = formats[i];
                string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-range."
                    + WordDocumentFormats.GetExtension(format)));
                WordFormatConverter.Save(source, path, format);
                OfficeDocument loaded = WordFormatConverter.Load(path);
                Require(loaded.GetPlainText().Contains("前 中段 后"), format + " keeps the paragraph text");
                Bookmark bookmark = loaded.Bookmarks["Mid"];
                Require(bookmark != null, format + " keeps the bookmark name");
                if (format == WordDocumentFormat.Html)
                {
                    // HTML 只有点锚：名字与位置还在，范围退化为整段。HTML has point anchors only: the name and
                    // position survive but the range degrades to the whole paragraph.
                    Require(bookmark.IsWholeParagraph, "HTML degrades the range to a whole-paragraph bookmark");
                    Require(bookmark.Text.Contains("中段"), "HTML bookmark text still covers the range");
                    continue;
                }
                Require(bookmark.Text == "中段", format + " keeps the range text, got [" + bookmark.Text + "]");
                Require(!bookmark.IsWholeParagraph, format + " keeps the range shape");

                string again = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-range-re."
                    + WordDocumentFormats.GetExtension(format)));
                WordFormatConverter.Save(loaded, again, format);
                OfficeDocument third = WordFormatConverter.Load(again);
                Require(third.Bookmarks.Count == 1, format + " resave keeps one bookmark, got " + third.Bookmarks.Count);
                Require(third.Bookmarks["Mid"].Text == "中段", format + " resave keeps the range text");
            }
        }

        /// <summary>校验书签名规则与非法名拒绝。Verifies the bookmark name rules and rejection of illegal names.</summary>
        private static void RunBookmarkNameRules()
        {
            Require(BookmarkNames.IsValid("Intro"), "plain name accepted");
            Require(BookmarkNames.IsValid("_Tail"), "underscore-prefixed name accepted");
            Require(BookmarkNames.IsValid("姓名"), "CJK name accepted");
            Require(BookmarkNames.IsValid(new string('a', 40)), "40-character name accepted");
            Require(!BookmarkNames.IsValid(null), "null name rejected");
            Require(!BookmarkNames.IsValid(string.Empty), "empty name rejected");
            Require(!BookmarkNames.IsValid(new string('a', 41)), "41-character name rejected");
            Require(!BookmarkNames.IsValid("1Intro"), "digit-prefixed name rejected");
            Require(!BookmarkNames.IsValid("my-bookmark"), "hyphen rejected");
            Require(!BookmarkNames.IsValid("my bookmark"), "space rejected");

            Require(BookmarkNames.Validate("  Intro  ") == "Intro", "validate trims the name");
            try
            {
                BookmarkNames.Validate("1bad");
                Require(false, "validate should reject an illegal name");
            }
            catch (ArgumentException)
            {
            }

            WordDocumentBuilder builder = new WordDocumentBuilder();
            try
            {
                builder.StartBookmark("9lives");
                Require(false, "StartBookmark should reject an illegal name");
            }
            catch (ArgumentException)
            {
            }
            try
            {
                builder.EndBookmark("9lives");
                Require(false, "EndBookmark should reject an illegal name");
            }
            catch (ArgumentException)
            {
            }
        }

        /// <summary>验证段落内范围书签按 w:id 配对往返。Verifies that an in-paragraph range pairs by w:id across a round trip.</summary>
        private static void RunBookmarkRangeRoundTrip(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-range.docx"));
            WordDocumentBuilder builder = new WordDocumentBuilder();
            builder.Write("前 ");
            builder.StartBookmark("Target");
            builder.Write("重点内容");
            builder.EndBookmark("Target");
            builder.Write(" 后");

            Require(builder.Document.Bookmarks.Count == 1, "builder range bookmark counted before save");
            Require(builder.Document.Bookmarks["Target"].Text == "重点内容",
                "builder range text read before save, got " + builder.Document.Bookmarks["Target"].Text);

            DocxWriter.Write(builder.Document, path);
            ValidateOoxmlPackage(path, "word/document.xml");

            string part = ReadPackagePartText(path, "word/document.xml");
            XDocument xml = XDocument.Parse(part);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            List<XElement> starts = new List<XElement>();
            foreach (XElement element in xml.Descendants(w + "bookmarkStart")) starts.Add(element);
            List<XElement> ends = new List<XElement>();
            foreach (XElement element in xml.Descendants(w + "bookmarkEnd")) ends.Add(element);
            Require(starts.Count == 1, "one w:bookmarkStart written, got " + starts.Count);
            Require(ends.Count == 1, "one w:bookmarkEnd written, got " + ends.Count);
            Require((string)starts[0].Attribute(w + "name") == "Target", "bookmark name written");
            Require((string)ends[0].Attribute(w + "id") == (string)starts[0].Attribute(w + "id"),
                "w:bookmarkEnd reuses the w:bookmarkStart id");

            OfficeDocument loaded = DocxReader.Load(path);
            Require(loaded.GetPlainText().Contains("前 重点内容 后"), "range round trip keeps paragraph text");
            Require(loaded.Bookmarks.Count == 1, "one bookmark after reload, got " + loaded.Bookmarks.Count);
            Bookmark target = loaded.Bookmarks["Target"];
            Require(target != null, "bookmark reachable by name after reload");
            Require(!target.IsWholeParagraph, "reloaded bookmark is a range");
            Require(target.IsValid, "reloaded range is paired");
            Require(target.Text == "重点内容", "reloaded range text, got " + target.Text);
            Require(loaded.Sections[0].Blocks[0] is Paragraph
                    && ((Paragraph)loaded.Sections[0].Blocks[0]).BookmarkName == "Target",
                "legacy whole-paragraph bookmark name still filled for compatibility");

            // 再存一次：编号可以变，配对必须仍成立。Resaving may renumber but pairing must survive.
            string again = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-range-again.docx"));
            DocxWriter.Write(loaded, again);
            OfficeDocument third = DocxReader.Load(again);
            Require(third.Bookmarks.Count == 1, "resave keeps one bookmark");
            Require(third.Bookmarks["Target"].Text == "重点内容",
                "resave keeps the range text, got " + third.Bookmarks["Target"].Text);

            // 嵌套书签：外层含内层。Nested bookmarks: the outer one wraps the inner one.
            WordDocumentBuilder nested = new WordDocumentBuilder();
            nested.StartBookmark("Outer");
            nested.Write("外");
            nested.StartBookmark("Inner");
            nested.Write("内");
            nested.EndBookmark("Inner");
            nested.Write("末");
            nested.EndBookmark("Outer");
            Require(nested.Document.Bookmarks.Count == 2, "nested bookmarks both counted");
            Require(nested.Document.Bookmarks["Inner"].Text == "内", "inner range text");
            Require(nested.Document.Bookmarks["Outer"].Text == "外内末", "outer range text");
            Require(nested.Document.Bookmarks[0].Name == "Outer", "indexer follows document order");
        }

        /// <summary>验证跨段落范围书签。Verifies a bookmark range spanning paragraphs.</summary>
        private static void RunBookmarkCrossParagraph(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-span.docx"));
            WordDocumentBuilder builder = new WordDocumentBuilder();
            builder.Write("第一段");
            builder.StartBookmark("Span");
            builder.InsertParagraph("第二段");
            builder.Write("结尾");
            builder.EndBookmark("Span");
            builder.InsertParagraph("第三段");

            DocxWriter.Write(builder.Document, path);
            OfficeDocument loaded = DocxReader.Load(path);
            Bookmark span = loaded.Bookmarks["Span"];
            Require(span != null, "cross-paragraph bookmark survives the round trip");
            Require(!span.IsWholeParagraph, "cross-paragraph bookmark stays a range");
            Require(span.IsValid, "cross-paragraph range is paired");
            Require(span.Text == "\n第二段结尾", "cross-paragraph text joined by newline, got [" + span.Text + "]");
            Require(loaded.GetPlainText().Contains("第三段"), "trailing paragraph unaffected");

            string part = ReadPackagePartText(path, "word/document.xml");
            XDocument xml = XDocument.Parse(part);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            List<XElement> paragraphs = new List<XElement>();
            foreach (XElement element in xml.Descendants(w + "p")) paragraphs.Add(element);
            XElement rangeStart = null;
            XElement rangeEnd = null;
            foreach (XElement element in xml.Descendants(w + "bookmarkStart"))
            {
                if ((string)element.Attribute(w + "name") == "Span") rangeStart = element;
            }
            Require(rangeStart != null, "range start written for the cross-paragraph bookmark");
            string startId = (string)rangeStart.Attribute(w + "id");
            foreach (XElement element in xml.Descendants(w + "bookmarkEnd"))
            {
                if ((string)element.Attribute(w + "id") == startId) rangeEnd = element;
            }
            Require(rangeEnd != null, "range end written with the matching id");
            Require(paragraphs.IndexOf(rangeStart.Parent) == 0, "range starts inside the first paragraph");
            Require(paragraphs.IndexOf(rangeEnd.Parent) == 1, "range ends inside the second paragraph");
            Require(paragraphs.Count >= 3, "paragraphs written, got " + paragraphs.Count);
        }

        /// <summary>验证书签文本替换与移除。Verifies replacing and removing bookmark content.</summary>
        private static void RunBookmarkTextReplaceAndRemove(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-text.docx"));
            WordDocumentBuilder builder = new WordDocumentBuilder();
            builder.Write("价格 ");
            builder.StartBookmark("Price");
            builder.Write("100 元");
            builder.EndBookmark("Price");
            builder.Write(" 含税");
            builder.InsertParagraph("整段书签");
            builder.InsertBookmark("Chapter");
            DocxWriter.Write(builder.Document, path);

            OfficeDocument loaded = DocxReader.Load(path);
            Require(loaded.Bookmarks.Count == 2, "range plus whole-paragraph bookmark, got " + loaded.Bookmarks.Count);

            Bookmark price = loaded.Bookmarks["Price"];
            price.Text = "120 元";
            Require(price.Text == "120 元", "range text read back after assignment");
            Require(loaded.Bookmarks.Count == 1 + 1, "replacement keeps both bookmarks");
            Require(loaded.GetPlainText().Contains("价格 120 元 含税"),
                "replacement rewrites the paragraph, got " + loaded.GetPlainText());

            Bookmark chapter = loaded.Bookmarks["Chapter"];
            Require(chapter.IsWholeParagraph, "InsertBookmark produces a whole-paragraph bookmark");
            Require(chapter.Text == "整段书签", "whole-paragraph text, got " + chapter.Text);
            chapter.Text = "改写整段";
            Require(chapter.Paragraph.GetText() == "改写整段", "whole-paragraph assignment replaces the paragraph");

            Require(loaded.Bookmarks.Contains("Price"), "Contains finds the range");
            Require(loaded.Bookmarks.Remove("Price"), "Remove reports success");
            Require(!loaded.Bookmarks.Contains("Price"), "range gone after Remove");
            Require(loaded.GetPlainText().Contains("120 元"), "removal keeps the content");
            loaded.Bookmarks.Clear();
            Require(loaded.Bookmarks.Count == 0, "Clear drops every bookmark");
            Require(loaded.GetPlainText().Contains("改写整段"), "Clear keeps the content");

            string saved = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-cleared.docx"));
            DocxWriter.Write(loaded, saved);
            OfficeDocument reloaded = DocxReader.Load(saved);
            Require(reloaded.Bookmarks.Count == 0, "cleared document writes no bookmark, got " + reloaded.Bookmarks.Count);
            Require(!ReadPackagePartText(saved, "word/document.xml").Contains("bookmarkStart"),
                "no w:bookmarkStart left in the cleared package");

            // 未配对的起点：范围存在但不可替换。An unpaired start: the bookmark exists but cannot be replaced.
            OfficeDocument orphan = new OfficeDocument();
            List<Block> blocks = orphan.FirstSection.Blocks;
            Paragraph orphanParagraph = blocks.Count > 0 ? blocks[0] as Paragraph : null;
            if (orphanParagraph == null)
            {
                orphanParagraph = new Paragraph();
                blocks.Add(orphanParagraph);
            }
            orphanParagraph.AddText("正文");
            orphanParagraph.Inlines.Add(new BookmarkStart("Loose"));
            Require(orphan.Bookmarks.Count == 1, "unpaired start counted once");
            Bookmark loose = orphan.Bookmarks["Loose"];
            Require(!loose.IsValid, "unpaired range reported invalid");
            Require(loose.Text == string.Empty, "unpaired range has no text");
            try
            {
                loose.Text = "x";
                Require(false, "unpaired range should refuse replacement");
            }
            catch (InvalidOperationException)
            {
            }

            try
            {
                orphan.Bookmarks.Add("Ghost", new BookmarkStart("Ghost"), new BookmarkEnd("Ghost"));
                Require(false, "Add should reject markers outside the tree");
            }
            catch (ArgumentException)
            {
            }
        }

        /// <summary>验证 MoveToBookmark 的行内定位。Verifies inline positioning through MoveToBookmark.</summary>
        private static void RunBookmarkCursor(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-cursor.docx"));
            WordDocumentBuilder builder = new WordDocumentBuilder();
            builder.Write("前 ");
            builder.StartBookmark("Note");
            builder.Write("备注");
            builder.EndBookmark("Note");
            builder.Write(" 后");
            DocxWriter.Write(builder.Document, path);

            OfficeDocument loaded = DocxReader.Load(path);
            WordDocumentBuilder after = new WordDocumentBuilder(loaded);
            after.MoveToBookmark("Note", false, true);
            after.Write("！");
            Require(loaded.GetPlainText().Contains("备注！ 后"),
                "cursor lands after the end marker, got " + loaded.GetPlainText());

            OfficeDocument before = DocxReader.Load(path);
            WordDocumentBuilder head = new WordDocumentBuilder(before);
            head.MoveToBookmark("Note", true, false);
            head.Write("*");
            Require(before.GetPlainText().Contains("前 *备注"),
                "cursor lands before the start marker, got " + before.GetPlainText());

            OfficeDocument missing = DocxReader.Load(path);
            WordDocumentBuilder noop = new WordDocumentBuilder(missing);
            int blocks = missing.Sections[0].Blocks.Count;
            noop.MoveToBookmark("NoSuchBookmark", true, true);
            noop.Write("尾部");
            Require(missing.Sections[0].Blocks.Count == blocks + 1,
                "unknown bookmark appends a trailing paragraph like any plain write");
            Paragraph last = missing.Sections[0].Blocks[blocks] as Paragraph;
            Require(last != null && last.GetText() == "尾部",
                "unknown bookmark keeps the cursor at the document end");
        }

        /// <summary>验证页眉页脚书签与 _GoBack 过滤。Verifies header/footer bookmarks and the _GoBack filter.</summary>
        private static void RunBookmarkHeaderFooterAndGoBack(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-header.docx"));
            OfficeDocument document = new OfficeDocument();
            WordDocumentBuilder builder = new WordDocumentBuilder(document);
            builder.Write("正文");
            builder.InsertBookmark("Body");

            Paragraph headerParagraph = new Paragraph();
            headerParagraph.AddText("页眉");
            headerParagraph.BookmarkName = "Heading";
            document.FirstSection.Header.Blocks.Add(headerParagraph);

            Paragraph footerParagraph = new Paragraph();
            footerParagraph.AddText("页脚 ");
            footerParagraph.Inlines.Add(new BookmarkStart("Foot"));
            footerParagraph.AddText("第 1 页");
            footerParagraph.Inlines.Add(new BookmarkEnd("Foot"));
            document.FirstSection.Footer.Blocks.Add(footerParagraph);

            DocxWriter.Write(document, path);
            OfficeDocument loaded = DocxReader.Load(path);
            Require(loaded.Bookmarks["Body"] != null, "body bookmark found");
            Require(loaded.Bookmarks["Heading"] != null, "header bookmark found");
            Require(loaded.Bookmarks["Foot"] != null, "footer range bookmark found");
            Require(loaded.Bookmarks["Foot"].Text == "第 1 页",
                "footer range text, got " + loaded.Bookmarks["Foot"].Text);

            // _GoBack 是 Word 的内部书签，不应出现在集合里。_GoBack is Word's internal bookmark and must stay out.
            string withGoBack = PrepareOutputPath(Path.Combine(outputDirectory, "word-bookmark-goback.docx"));
            File.Copy(path, withGoBack, true);
            InjectGoBackBookmark(withGoBack, "word/document.xml");
            OfficeDocument patched = DocxReader.Load(withGoBack);
            Require(!patched.Bookmarks.Contains("_GoBack"), "_GoBack ignored");
            Require(patched.Bookmarks["Body"] != null, "_GoBack does not disturb real bookmarks");
            Require(patched.Bookmarks.Count == loaded.Bookmarks.Count,
                "bookmark count unchanged by _GoBack");
        }

        /// <summary>读取包内部件的文本内容。Reads the text of a package part.</summary>
        /// <param name="path">包路径。The package path.</param>
        /// <param name="part">部件名。The part name.</param>
        /// <returns>部件文本。The part text.</returns>
        private static string ReadPackagePartText(string path, string part)
        {
            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                ZipArchiveEntry entry = archive.GetEntry(part);
                Require(entry != null, "package part missing: " + part);
                using (StreamReader reader = new StreamReader(entry.Open(), Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>向部件正文末尾注入 Word 内部的 _GoBack 书签。Injects Word's internal _GoBack bookmark at the end of a part body.</summary>
        /// <param name="path">包路径。The package path.</param>
        /// <param name="part">部件名。The part name.</param>
        private static void InjectGoBackBookmark(string path, string part)
        {
            string xml = ReadPackagePartText(path, part);
            string markers = "<w:bookmarkStart w:id=\"99\" w:name=\"_GoBack\"/>"
                + "<w:bookmarkEnd w:id=\"99\"/>";
            int insertAt = xml.LastIndexOf("</w:p>", StringComparison.Ordinal);
            Require(insertAt >= 0, "paragraph end found for _GoBack injection");
            string patched = xml.Substring(0, insertAt) + markers + xml.Substring(insertAt);
            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update))
            {
                ZipArchiveEntry entry = archive.GetEntry(part);
                Require(entry != null, "package part missing: " + part);
                entry.Delete();
                ZipArchiveEntry replacement = archive.CreateEntry(part);
                using (StreamWriter writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false)))
                {
                    writer.Write(patched);
                }
            }
        }
    }
}
