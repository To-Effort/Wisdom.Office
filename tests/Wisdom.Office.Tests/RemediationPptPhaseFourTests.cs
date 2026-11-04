using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Wisdom.Office.Core;
using Wisdom.Ppt;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 4 的 PPT 条目：产出损坏文件与静默丢失的形状。
    /// Phase-four PPT entries of remediation design episode two: corrupt output and silently dropped shapes.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>PPT Phase 4 入口。Entry point for the PPT half of phase four.</summary>
        private static void RunRemediationPptPhaseFourTest()
        {
            RunStyledPlaceholderSurvivesTest();
            RunPartialRunFormatStaysInheritedTest();

            Console.WriteLine("[remediation] ppt phase-four shapes and media ok");
        }

        /// <summary>
        /// PT6：库自己补出来的空占位符，被调用者加过底色或描边之后必须写得出门；
        /// 谁都没动过的那一份才允许按「PowerPoint 的空提示」省掉。
        /// PT6: a placeholder this library generated must survive the save once the caller has given it a
        /// fill or an outline; only the one nobody touched may still be dropped as PowerPoint's empty prompt.
        /// </summary>
        private static void RunStyledPlaceholderSurvivesTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();

            Slide slide = deck.AddSlide(SlideLayoutKind.TitleAndContent);
            ((TextBoxShape)slide.Shapes[0]).Paragraphs[0].AddText("标题");
            TextBoxShape block = (TextBoxShape)slide.Shapes[1];
            block.FillColor = new OfficeColor(0xFF, 0xE8, 0xD9);
            block.Outline = new BorderLine(BorderStyle.Single, 1.5, new OfficeColor(0xC0, 0x39, 0x2B));
            double width = block.Width;
            double height = block.Height;

            // 控制：另一页只填标题，版式自带的内容占位符没人动过，仍然应当被省掉——
            // 「程序生成的演示文稿不该留一个等用户点进去的空框」这条承诺不能因为本次改动丢掉。
            // Control: on the next slide only the title is filled, so the untouched layout placeholder must
            // still be dropped — the promise that generated decks carry no "click to add" boxes cannot be
            // lost with this change.
            Slide plain = deck.AddSlide(SlideLayoutKind.TitleAndContent);
            ((TextBoxShape)plain.Shapes[0]).Paragraphs[0].AddText("只有标题");
            Require(plain.Shapes.Count == 2, "the plain slide must start with two generated placeholders");

            OfficePresentation loaded;
            using (MemoryStream buffer = new MemoryStream())
            {
                PptxWriter.Write(deck, buffer);
                loaded = PptxReader.Load(buffer.ToArray());
            }

            Slide first = loaded.Slides[0];
            Require(first.Shapes.Count == 2,
                "the styled placeholder vanished on save: the slide carries " + first.Shapes.Count
                + " shape(s) instead of 2");
            TextBoxShape back = (TextBoxShape)first.Shapes[1];
            RequireHex(back.FillColor, "FFE8D9", "placeholder fill");
            Require(back.Outline != null && back.Outline.Style == BorderStyle.Single,
                "the placeholder outline was not written back");
            Require(Math.Abs(back.Width - width) < 0.5 && Math.Abs(back.Height - height) < 0.5,
                "the placeholder geometry changed: " + back.Width + " x " + back.Height);

            Slide second = loaded.Slides[1];
            Require(second.Shapes.Count == 1,
                "an untouched generated placeholder was written out: the plain slide carries "
                + second.Shapes.Count + " shape(s)");
        }

        /// <summary>
        /// PT5：只显式设了一处属性的 run，不该被写进模型预置的 11 磅与 Arial。
        /// 母版里 18pt 的正文套上一个只加粗的词，写出后被钉成 11pt，就是这条预置值在作乱。
        /// PT5: a run that sets one property explicitly must not leave the file carrying the model's preset
        /// 11 pt and Arial. Bold one word inside an 18pt body placeholder and that preset shrinks it to 11pt.
        /// </summary>
        private static void RunPartialRunFormatStaysInheritedTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            TextBoxShape box = slide.AddTextBox(40, 40, 700, 120, "起始");
            Paragraph paragraph = box.Paragraphs[0];
            paragraph.Inlines.Clear();

            TextRun boldOnly = new TextRun("只加粗");
            boldOnly.Format.Bold = true;
            TextRun sizeOnly = new TextRun("只设字号");
            sizeOnly.Format.FontSize = 22.0;
            TextRun fontOnly = new TextRun("只设字体");
            fontOnly.Format.FontFamily = "宋体";
            paragraph.Inlines.Add(boldOnly);
            paragraph.Inlines.Add(sizeOnly);
            paragraph.Inlines.Add(fontOnly);

            byte[] written = WriteToBuffer(deck);
            XElement[] runProps = ReadRunProps(written, "ppt/slides/slide1.xml");
            Require(runProps.Length == 3, "expected three runs in the written slide, got " + runProps.Length);

            Require((string)runProps[0].Attribute("b") == "1", "the bold-only run lost its b attribute");
            Require((string)runProps[0].Attribute("sz") == null,
                "a run that never set a size was pinned to the model default: sz="
                + (string)runProps[0].Attribute("sz"));
            Require(runProps[0].Element(A + "latin") == null,
                "a run that never set a font was pinned to an explicit a:latin");
            Require(runProps[0].Element(A + "ea") == null,
                "a run that never set a font was pinned to an explicit a:ea");
            Require(runProps[0].Element(A + "solidFill") == null,
                "a run that never set a colour was pinned to an explicit a:solidFill");

            Require((string)runProps[1].Attribute("sz") == "2200",
                "the explicit 22pt size was not written: sz=" + (string)runProps[1].Attribute("sz"));
            Require(runProps[1].Element(A + "latin") == null,
                "the size-only run invented an explicit font");
            XElement latin = runProps[2].Element(A + "latin");
            Require(latin != null && (string)latin.Attribute("typeface") == "宋体",
                "the explicit typeface was not written on the third run");
            Require((string)runProps[2].Attribute("sz") == null,
                "the font-only run invented an explicit size");

            // 往返一轮再写：读回来的格式不该把没读到的属性当成「显式设过」。
            // After a round trip the re-written file must stay lean: the reader must not report the
            // properties it never found in the source as explicitly set.
            OfficePresentation loaded = PptxReader.Load(written);
            TextBoxShape loadedBox = (TextBoxShape)loaded.Slides[0].Shapes[0];
            Paragraph loadedParagraph = loadedBox.Paragraphs[0];
            Require(loadedParagraph.Inlines.Count == 3,
                "the three runs did not read back, got " + loadedParagraph.Inlines.Count);
            TextRun loadedBold = (TextRun)loadedParagraph.Inlines[0];
            TextRun loadedSize = (TextRun)loadedParagraph.Inlines[1];
            TextRun loadedFont = (TextRun)loadedParagraph.Inlines[2];
            Require(loadedBold.Format.Bold, "bold did not survive the round trip");
            Require(Math.Abs(loadedSize.Format.FontSize - 22.0) < 0.0001,
                "size did not survive the round trip: " + loadedSize.Format.FontSize);
            Require(loadedFont.Format.FontFamily == "宋体",
                "typeface did not survive the round trip: " + loadedFont.Format.FontFamily);

            byte[] rewritten = WriteToBuffer(loaded);
            XElement[] again = ReadRunProps(rewritten, "ppt/slides/slide1.xml");
            Require(again.Length == 3, "the rewritten deck lost runs, got " + again.Length);
            Require((string)again[0].Attribute("sz") == null && again[0].Element(A + "latin") == null,
                "the round-tripped bold-only run was pinned to defaults on the second write");
        }

        /// <summary>DrawingML 命名空间，测试里手写一份，避免依赖写端的私有常量。 The DrawingML namespace, spelled out here so the test does not lean on a writer private.</summary>
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        /// <summary>把演示文稿写到内存并返回字节。 Writes the presentation to memory and returns the bytes.</summary>
        /// <param name="presentation">待写出的演示文稿。 The presentation to write.</param>
        /// <returns>pptx 字节。 The pptx bytes.</returns>
        private static byte[] WriteToBuffer(OfficePresentation presentation)
        {
            using (MemoryStream buffer = new MemoryStream())
            {
                PptxWriter.Write(presentation, buffer);
                return buffer.ToArray();
            }
        }

        /// <summary>
        /// 取出幻灯片部件里每个 <c>a:r</c> 的 <c>a:rPr</c>，按文档顺序返回。
        /// Collects the <c>a:rPr</c> of every <c>a:r</c> in a slide part, in document order.
        /// </summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径，如 ppt/slides/slide1.xml。The part path, such as ppt/slides/slide1.xml.</param>
        /// <returns>运行属性元素列表。The run property elements.</returns>
        private static XElement[] ReadRunProps(byte[] archive, string entryName)
        {
            using (ZipArchive zip = new ZipArchive(new MemoryStream(archive), ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = zip.GetEntry(entryName);
                Require(entry != null, "slide part missing: " + entryName);
                using (Stream stream = entry.Open())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    XDocument doc = XDocument.Parse(reader.ReadToEnd());
                    List<XElement> found = new List<XElement>();
                    foreach (XElement run in doc.Descendants(A + "r"))
                    {
                        XElement props = run.Element(A + "rPr");
                        if (props != null) found.Add(props);
                    }
                    return found.ToArray();
                }
            }
        }
    }
}
