using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Markdown;
using Wisdom.Pdf;
using Wisdom.Ppt;

namespace Wisdom.Office.Tests
{
    /// <summary>PPTX 读写与 Markdown 解析/渲染的回归测试。Regression tests for PPTX round-trips and Markdown parse/render.</summary>
    internal static partial class Program
    {
        private static void RunPptxTest(string pptxPath, string outputDirectory)
        {
            OfficePresentation presentation = OfficePresentation.Widescreen();
            presentation.Metadata.Title = "PPT 回归测试";
            presentation.Metadata.Author = "Wisdom.Office";

            Slide cover = presentation.AddSlide(SlideLayoutKind.Title);
            TextBoxShape coverTitle = (TextBoxShape)cover.Shapes[0];
            coverTitle.Paragraphs[0].AddText("统一模型演示");
            cover.NotesText = "备注：控制节奏 30 秒。";

            Slide body = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
            TextBoxShape bodyTitle = null;
            TextBoxShape content = null;
            for (int i = 0; i < body.Shapes.Count; i++)
            {
                TextBoxShape box = body.Shapes[i] as TextBoxShape;
                if (box == null) continue;
                if (box.Placeholder == PlaceholderKind.Title) bodyTitle = box;
                if (box.Placeholder == PlaceholderKind.Content) content = box;
            }
            Require(bodyTitle != null && content != null, "title/content placeholders missing");
            bodyTitle.Paragraphs[0].AddText("三区明细");
            content.Paragraphs[0].AddText("华东：588 万元");
            content.Paragraphs[0].ListMarker = "•";
            Paragraph second = content.AddParagraph(string.Empty);
            second.AddText("华南：384 万元");
            second.ListMarker = "•";
            body.Transition = SlideTransitionKind.Fade;
            body.TransitionDuration = 0.6;

            TableShape tableShape = body.AddTable(80, 330, 800, 120);
            tableShape.HeaderColor = new OfficeColor(0x1F, 0x3A, 0x5F);
            tableShape.Table.AddRow("区域", "收入", "占比");
            tableShape.Table.AddRow("华东", "588", "46%");
            tableShape.Table.AddRow("华南", "384", "30%");

            Slide graphics = presentation.AddSlide(SlideLayoutKind.Blank);
            AutoShape banner = graphics.AddShape(AutoShapeKind.RoundedRectangle, 96, 60, 768, 80, "达成率 96%");
            banner.FillColor = new OfficeColor(0x00, 0xB0, 0x50);
            graphics.AddImage(new OfficeImage(CreateTestPng(8, 6)), 96, 180, 120, 90);

            PptxWriter.Write(presentation, pptxPath);
            ValidateOoxmlPackage(pptxPath, "ppt/presentation.xml");

            OfficeFileFormat detected = OfficeFileFormats.Detect(pptxPath);
            Require(detected == OfficeFileFormat.Pptx, "format detection failed: " + detected);

            OfficePresentation loaded = PptxReader.Load(pptxPath);
            Require(loaded.Slides.Count == 3, "slide count mismatch: " + loaded.Slides.Count);
            Require(Math.Abs(loaded.SlideWidth - 960.0) < 0.5, "slide width mismatch: " + loaded.SlideWidth);
            Require(Math.Abs(loaded.SlideHeight - 540.0) < 0.5, "slide height mismatch: " + loaded.SlideHeight);
            Require(loaded.Metadata.Title == "PPT 回归测试", "metadata title mismatch: " + loaded.Metadata.Title);
            Require(loaded.Slides[0].GetTitle() == "统一模型演示", "cover title mismatch: " + loaded.Slides[0].GetTitle());
            Require(loaded.Slides[0].Layout == SlideLayoutKind.Title, "cover layout mismatch: " + loaded.Slides[0].Layout);
            Require(loaded.Slides[0].NotesText != null && loaded.Slides[0].NotesText.Contains("备注"),
                "notes missing: " + loaded.Slides[0].NotesText);
            Require(loaded.Slides[1].Transition == SlideTransitionKind.Fade, "transition mismatch: " + loaded.Slides[1].Transition);
            Require(Math.Abs(loaded.Slides[1].TransitionDuration - 0.6) < 0.001,
                "transition duration mismatch: " + loaded.Slides[1].TransitionDuration);

            int bulletRuns = CountBulletParagraphs(loaded.Slides[1]);
            Require(bulletRuns >= 2, "bulleted paragraphs lost: " + bulletRuns);

            TableShape loadedTable = FindTable(loaded.Slides[1]);
            Require(loadedTable != null, "table shape lost after reload");
            Require(loadedTable.Table.Rows.Count == 3, "table rows mismatch: " + loadedTable.Table.Rows.Count);
            Require(loadedTable.Table.Rows[1].Cells[0].Blocks.Count > 0
                && GetCellText(loadedTable.Table.Rows[1].Cells[0]) == "华东", "table cell text mismatch");

            PictureShape loadedPicture = FindPicture(loaded.Slides[2]);
            Require(loadedPicture != null, "picture shape lost after reload");
            Require(loadedPicture.Image != null && loadedPicture.Image.Data.Length > 0, "picture bytes lost");

            AutoShape loadedBanner = FindAutoShape(loaded.Slides[2]);
            Require(loadedBanner != null, "auto shape lost after reload");
            Require(loadedBanner.FillColor != null && loadedBanner.FillColor.Value.ToHex() == "00B050",
                "shape fill mismatch: " + (loadedBanner.FillColor == null ? "null" : loadedBanner.FillColor.Value.ToHex()));

            OfficeFile file = OfficeFile.Load(pptxPath);
            Require(file.Format == OfficeFileFormat.Pptx, "facade format mismatch: " + file.Format);
            Require(file.ExtractText().IndexOf("统一模型演示", StringComparison.Ordinal) >= 0, "facade text missing title");
            OfficeDocument asWord = file.AsDocument();
            Require(CountBlocks(asWord) > 0, "presentation-to-document conversion produced nothing");

            string pdfPath = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-convert.pdf"));
            OfficeConverter.PresentationToPdf(loaded, pdfPath);
            PdfDocument pdf = PdfDocument.Load(pdfPath);
            Require(pdf.PageCount == loaded.Slides.Count, "expected exactly one PDF page per slide, got " + pdf.PageCount);

            // 超页内容回归：单页放一整段超长文本，导出 PDF 仍须恰为一页（不自动续页，与 PowerPoint 导出裁切一致）。
            // Overflow regression: one slide packed with a long paragraph must still export to exactly one PDF page (no auto page break, same as PowerPoint clipping).
            OfficePresentation tall = new OfficePresentation();
            Slide overflow = tall.AddSlide(SlideLayoutKind.Blank);
            TextBoxShape longBox = new TextBoxShape(20, 20, 900, 500);
            for (int line = 0; line < 120; line++)
            {
                longBox.AddParagraph("超长内容第 " + line + " 行：幻灯片转 PDF 必须严格一页幻对应一页，超出部分留在本页裁切。");
            }
            overflow.Shapes.Add(longBox);
            string tallPath = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-overflow.pdf"));
            OfficeConverter.PresentationToPdf(tall, tallPath);
            PdfDocument overflowPdf = PdfDocument.Load(tallPath);
            Require(overflowPdf.PageCount == 1, "overflow slide must stay on exactly one page, got " + overflowPdf.PageCount);

            // 固定版式保真回归：验证 PPT→PDF 按绝对坐标渲染后，项目符号、形状内文本位置、圆角几何都落到正确页面上，
            // 而不是退回流式排版把全部文本堆在页面左侧。
            // Fixed-layout fidelity: after absolute-coordinate rendering, bullets, in-shape text positions and rounded-rect
            // geometry must land correctly instead of collapsing back to flow layout (all text stacked at the left margin).
            string bodyText = pdf.Pages[1].ExtractText();
            Require(bodyText.IndexOf('•') >= 0, "list markers lost in fixed-layout PDF: " + bodyText);

            bool bannerPlaced = false;
            bool bannerGeometry = false;
            byte[] graphicsContent = pdf.Pages[2].GetContentBytes();
            string graphicsAscii = Encoding.GetEncoding("ISO-8859-1").GetString(graphicsContent);
            foreach (PdfTextLine line in pdf.Pages[2].GetTextLines())
            {
                if (line.Text == null || line.Text.IndexOf("达成率", StringComparison.Ordinal) < 0) continue;
                // 文本应落在横幅矩形内（Left≈96、Top≈60、H≈80），而非页边距处。
                // Text should sit inside the banner rect (Left~96, Top~60, H~80), not at the page margin.
                bannerPlaced = line.X >= 80.0 && line.Y >= 50.0 && line.Y <= 150.0;
            }
            bannerGeometry = graphicsAscii.IndexOf(" c", StringComparison.Ordinal) >= 0;
            Require(bannerPlaced, "banner text not placed at the shape position in fixed-layout PDF");
            Require(bannerGeometry, "rounded-rect shape geometry (bezier path) missing from fixed-layout PDF");

            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "ppt-convert.docx"));
            OfficeConverter.PresentationToWord(loaded, docxPath);
            OfficeDocument word = OfficeFile.Load(docxPath).AsDocument();
            Require(CountBlocks(word) > 0, "presentation-to-word produced nothing");

            string fromWordPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-to-pptx.pptx"));
            OfficeDocument source = BuildHeadingDocument();
            OfficeConverter.DocumentToPptx(source, fromWordPath);
            OfficePresentation split = PptxReader.Load(fromWordPath);
            Require(split.Slides.Count == CountHeadings(source),
                "expected one slide per heading, got " + split.Slides.Count);

            using (MemoryStream stream = new MemoryStream())
            {
                PptxWriter.Write(loaded, stream);
                Require(stream.Length > 1024, "stream write too small: " + stream.Length);
                OfficePresentation again = PptxReader.Load(stream.ToArray());
                Require(again.Slides.Count == loaded.Slides.Count, "byte round-trip slide mismatch");
            }

            Console.WriteLine("[pptx] slides=" + loaded.Slides.Count + ", size=" + (int)loaded.SlideWidth + "x"
                + (int)loaded.SlideHeight + ", bytes=" + new FileInfo(pptxPath).Length
                + ", pdf=" + pdf.PageCount + " pages, docx=" + new FileInfo(docxPath).Length
                + " bytes, split=" + split.Slides.Count + " slides");
        }

        private static void RunMarkdownTest()
        {
            string source = "---\ntitle: 回归样例\nauthor: Wisdom.Office\n---\n\n"
                + "# 一级标题\n\n"
                + "正文含 **粗体**、*斜体*、~~删除~~、`代码` 与 [示例链接](https://example.com/doc)。\n\n"
                + "- [x] 已完成项\n- [ ] 待办项\n\n"
                + "1. 有序一\n2. 有序二\n\n"
                + "| 区域 | 收入 |\n|:-----|-----:|\n| 华东 | 588 |\n| 华南 | 384 |\n\n"
                + "> 引用行\n\n"
                + "```csharp\nint answer = 42;\n```\n\n"
                + "不安全链接：<script>alert(1)</script> 与 [点击](javascript:alert(1))。\n";

            OfficeDocument document = Wisdom.Markdown.Markdown.ToDocument(source);
            Require(document.Metadata.Title == "回归样例", "front-matter title mismatch: " + document.Metadata.Title);
            Require(document.Metadata.Author == "Wisdom.Office", "front-matter author mismatch: " + document.Metadata.Author);

            int headings = CountHeadings(document);
            Require(headings >= 1, "no headings parsed");
            int tables = CountTables(document);
            Require(tables == 1, "expected one table, got " + tables);
            Require(CountListItems(document) >= 4, "list items lost: " + CountListItems(document));

            string link = FindFirstLink(document);
            Require(link == "https://example.com/doc", "link target lost: " + (link == null ? "null" : link));

            string html = Wisdom.Markdown.Markdown.ToHtml(source);
            RequireContains(html, "<a href=\"https://example.com/doc\">");
            RequireContains(html, "<table>");
            RequireContains(html, "<input type=\"checkbox\"");
            RequireContains(html, "checked=\"checked\"");
            RequireContains(html, "<pre><code>");
            Require(html.IndexOf("href=\"javascript:", StringComparison.Ordinal) < 0, "unsafe javascript: link was emitted");
            Require(html.IndexOf("<script>", StringComparison.Ordinal) < 0, "raw script markup was not escaped");
            RequireContains(html, "&lt;script&gt;");

            string markdown = Wisdom.Markdown.Markdown.ToMarkdown(document);
            RequireContains(markdown, "[示例链接](https://example.com/doc)");
            RequireContains(markdown, "| 区域 | 收入 |");
            RequireContains(markdown, "- [x] 已完成项");
            OfficeDocument reparsed = Wisdom.Markdown.Markdown.ToDocument(markdown);
            // 回写会把 Metadata.Title 补成一行一级标题，因此标题数 +1。The writer promotes Metadata.Title to a heading, so the count grows by one.
            Require(CountHeadings(reparsed) == headings + 1,
                "round-trip heading count changed: " + CountHeadings(reparsed) + " vs " + headings);
            Require(CountTables(reparsed) == tables, "round-trip table count changed");
            Require(FindFirstLink(reparsed) == link, "round-trip link target changed");

            MarkdownOptions noTables = new MarkdownOptions();
            noTables.EnableTables = false;
            noTables.EnableTaskLists = false;
            OfficeDocument plain = Wisdom.Markdown.Markdown.ToDocument(source, noTables);
            Require(CountTables(plain) == 0, "table parsed while disabled");
            Require(CountListItems(plain) >= 4, "list markers lost with task lists disabled");

            // 分割线与五层嵌套列表的往返保真：模型把分割线承载为全角星号文本，两个写器必须按同一常量识别；
            // 层级换算用 24 磅，解析器也要按 24 磅产出缩进，否则第 3 层以后会被压平。
            // Round-trip fidelity for thematic breaks and deep lists: the model carries a break as full-width asterisk
            // text that both writers must match with the same constant, and the parser has to emit the 24-point per-level
            // indent the writers divide by, otherwise level three and deeper collapse.
            string nesting = "- 一层\n  - 二层\n    - 三层\n      - 四层\n        - 五层\n\n***\n\n正文\n";
            OfficeDocument nestedDoc = Wisdom.Markdown.Markdown.ToDocument(nesting);
            string nestedHtml = Wisdom.Markdown.Markdown.ToHtml(nesting);
            Require(CountOccurrences(nestedHtml, "<ul>") == 5, "nested list levels lost: " + CountOccurrences(nestedHtml, "<ul>"));
            Require(nestedHtml.IndexOf("<hr />", StringComparison.Ordinal) >= 0, "thematic break not rendered as <hr />");
            string nestedMarkdown = Wisdom.Markdown.Markdown.ToMarkdown(nestedDoc);
            RequireContains(nestedMarkdown, "        - 五层");
            RequireContains(nestedMarkdown, "---");
            Require(nestedMarkdown.IndexOf('＊') < 0, "thematic break leaked as literal asterisk text");

            // 加固项：深层嵌套不撑爆栈、默认不读本地磁盘、协议相对链接不进 href。
            // Hardening: deep nesting must not overflow the stack, local disk stays unread by default,
            // and a protocol-relative target must not reach href.
            StringBuilder deepBuilder = new StringBuilder();
            for (int i = 0; i < 2000; i++)
            {
                deepBuilder.Append('[');
            }
            deepBuilder.Append('x');
            for (int i = 0; i < 2000; i++)
            {
                deepBuilder.Append("](u)");
            }
            string deepLinks = deepBuilder.ToString();
            Require(CountBlocks(Wisdom.Markdown.Markdown.ToDocument(deepLinks)) >= 1, "deep link nesting produced no blocks");
            string deepHtml = Wisdom.Markdown.Markdown.ToHtml(deepLinks);
            Require(CountOccurrences(deepHtml, "<a href=") <= 64, "deep link nesting exceeded the depth cap: " + CountOccurrences(deepHtml, "<a href="));

            StringBuilder emphasisBuilder = new StringBuilder();
            for (int i = 0; i < 200; i++)
            {
                emphasisBuilder.Append(i % 2 == 0 ? "**" : "__");
            }
            emphasisBuilder.Append('x');
            for (int i = 0; i < 200; i++)
            {
                emphasisBuilder.Append(i % 2 == 0 ? "__" : "**");
            }
            Require(CountBlocks(Wisdom.Markdown.Markdown.ToDocument(emphasisBuilder.ToString())) >= 1, "deep emphasis nesting produced no blocks");

            Require(string.IsNullOrEmpty(new MarkdownOptions().ImageBaseDirectory),
                "local image loading should default to disabled");
            OfficeDocument absoluteImage = Wisdom.Markdown.Markdown.ToDocument("![外部](C:/Windows/win.ini)");
            Require(CountImages(absoluteImage) == 0, "an absolute image path was read from disk");

            MarkdownOptions sandboxed = new MarkdownOptions();
            sandboxed.ImageBaseDirectory = Path.Combine(Path.GetTempPath(), "wisdom-md-sandbox");
            OfficeDocument escaped = Wisdom.Markdown.Markdown.ToDocument("![逃逸](../../Windows/win.ini)", sandboxed);
            Require(CountImages(escaped) == 0, "image path escaped the base directory");

            string protocolRelative = Wisdom.Markdown.Markdown.ToHtml("[外链](//evil.example/a) 与 [相对](/docs/a)");
            Require(protocolRelative.IndexOf("href=\"//", StringComparison.Ordinal) < 0,
                "a protocol-relative link was emitted into href");
            RequireContains(protocolRelative, "href=\"/docs/a\"");

            Console.WriteLine("[markdown] headings=" + headings + ", tables=" + tables
                + ", listItems=" + CountListItems(document) + ", html=" + html.Length
                + " chars, roundTrip=" + markdown.Length
                + " chars, nestedLevels=" + CountOccurrences(nestedHtml, "<ul>"));
        }

        private static int CountBlocks(OfficeDocument document)
        {
            int total = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                total += document.Sections[s].Blocks.Count;
            }
            return total;
        }

        private static int CountListItems(OfficeDocument document)
        {
            int total = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                for (int i = 0; i < document.Sections[s].Blocks.Count; i++)
                {
                    Paragraph paragraph = document.Sections[s].Blocks[i] as Paragraph;
                    if (paragraph != null && !string.IsNullOrEmpty(paragraph.ListMarker)) total++;
                }
            }
            return total;
        }

        private static string FindFirstLink(OfficeDocument document)
        {
            for (int s = 0; s < document.Sections.Count; s++)
            {
                for (int i = 0; i < document.Sections[s].Blocks.Count; i++)
                {
                    Paragraph paragraph = document.Sections[s].Blocks[i] as Paragraph;
                    if (paragraph == null) continue;
                    for (int k = 0; k < paragraph.Inlines.Count; k++)
                    {
                        TextRun run = paragraph.Inlines[k] as TextRun;
                        if (run != null && !string.IsNullOrEmpty(run.Link)) return run.Link;
                    }
                }
            }
            return null;
        }

        private static OfficeDocument BuildHeadingDocument()
        {
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = "由文档拆分";
            Section section = document.AddSection();
            for (int i = 1; i <= 3; i++)
            {
                Paragraph heading = section.AddParagraph("第 " + i + " 章");
                heading.HeadingLevel = 1;
                section.AddParagraph("正文段落 " + i + "。");
            }
            return document;
        }

        private static int CountBulletParagraphs(Slide slide)
        {
            int total = 0;
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                TextBoxShape box = slide.Shapes[i] as TextBoxShape;
                if (box == null) continue;
                for (int p = 0; p < box.Paragraphs.Count; p++)
                {
                    if (!string.IsNullOrEmpty(box.Paragraphs[p].ListMarker)) total++;
                }
            }
            return total;
        }

        private static TableShape FindTable(Slide slide)
        {
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                TableShape table = slide.Shapes[i] as TableShape;
                if (table != null) return table;
            }
            return null;
        }

        private static PictureShape FindPicture(Slide slide)
        {
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                PictureShape picture = slide.Shapes[i] as PictureShape;
                if (picture != null) return picture;
            }
            return null;
        }

        private static AutoShape FindAutoShape(Slide slide)
        {
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                AutoShape shape = slide.Shapes[i] as AutoShape;
                if (shape != null) return shape;
            }
            return null;
        }

        private static string GetCellText(TableCell cell)
        {
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < cell.Blocks.Count; i++)
            {
                Paragraph paragraph = cell.Blocks[i] as Paragraph;
                if (paragraph != null) text.Append(paragraph.GetText());
            }
            return text.ToString();
        }
    }
}
