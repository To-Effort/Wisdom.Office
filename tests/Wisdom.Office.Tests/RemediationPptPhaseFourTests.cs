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
            RunUndeclaredTableHeaderStaysPlainTest();
            RunZeroHeightSeparatorSurvivesTest();
            RunTransitionDirectionSpeedClickTest();

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

        /// <summary>
        /// PT3 结构档：没声明表头的表格，写端不该替它立 firstRow 与 bandRow、把首行刷成表头色加粗白字，
        /// 更不该给每个格子写 <c>a:noFill</c> 把 <c>tableStyleId</c> 的话头堵死。
        /// PT3 (structure tier): a table whose header was never declared must not leave the file with
        /// firstRow/bandRow switched on, a painted bold-white first row, and an a:noFill in every cell that
        /// silences the tableStyleId PowerPoint would otherwise apply.
        /// </summary>
        private static void RunUndeclaredTableHeaderStaysPlainTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);

            TableShape plain = slide.AddTable(60, 60, 700, 160);
            plain.Table.AddRow("华东", "588", "46%");
            plain.Table.AddRow("华南", "384", "30%");

            TableShape declared = slide.AddTable(60, 300, 700, 160);
            declared.HeaderColor = new OfficeColor(0x1F, 0x3A, 0x5F);
            declared.Table.AddRow("区域", "收入", "占比");
            declared.Table.AddRow("华北", "300", "24%");

            byte[] written = WriteToBuffer(deck);
            XElement[] tables = ReadTables(written, "ppt/slides/slide1.xml");
            Require(tables.Length == 2, "expected two tables on the slide, got " + tables.Length);

            XElement plainPr = tables[0].Element(A + "tblPr");
            Require(plainPr != null, "the undeclared table lost its a:tblPr");
            Require((string)plainPr.Attribute("firstRow") == null,
                "firstRow was switched on for a table that never declared a header row");
            Require((string)plainPr.Attribute("bandRow") == null,
                "bandRow was switched on although the model carries no banding to express");
            XElement plainCell = FirstCell(FirstRow(tables[0]));
            XElement plainTcPr = plainCell == null ? null : plainCell.Element(A + "tcPr");
            Require(plainTcPr != null, "the undeclared table's first cell lost its a:tcPr");
            Require(plainTcPr.Element(A + "solidFill") == null,
                "the first row was painted a header fill nobody asked for");
            Require(plainTcPr.Element(A + "noFill") == null,
                "a:noFill was written for a cell that set no background, muting the table style");
            XElement plainRun = FirstRunProps(plainCell);
            Require(plainRun != null && (string)plainRun.Attribute("b") == null,
                "the undeclared first row's text was pinned bold as a header");

            XElement declaredPr = tables[1].Element(A + "tblPr");
            Require(declaredPr != null && (string)declaredPr.Attribute("firstRow") == "1",
                "HeaderColor was set but a:tblPr carries no firstRow, so the table style cannot paint it");
            XElement headerCell = FirstCell(FirstRow(tables[1]));
            XElement headerTcPr = headerCell == null ? null : headerCell.Element(A + "tcPr");
            XElement headerFill = headerTcPr == null ? null : headerTcPr.Element(A + "solidFill");
            XElement headerClr = headerFill == null ? null : headerFill.Element(A + "srgbClr");
            Require(headerClr != null && (string)headerClr.Attribute("val") == "1F3A5F",
                "the declared header colour was not written");
            XElement headerRun = FirstRunProps(headerCell);
            Require(headerRun != null && (string)headerRun.Attribute("b") == "1",
                "the declared header row lost its bold emphasis");

            // 往返：a:tblPr@firstRow 是模型里 IsHeader 唯一的回来路，读不回来等于二次写出把表头整段抹平。
            // Round trip: a:tblPr@firstRow is the only route back to TableRow.IsHeader; skipping it would
            // flatten a declared header on the second save.
            OfficePresentation loaded = PptxReader.Load(written);
            List<TableShape> back = FindTables(loaded.Slides[0]);
            Require(back.Count == 2, "both tables must read back, got " + back.Count);
            Require(back[1].Table.Rows[0].IsHeader,
                "the declared header row did not survive the round trip");
            Require(!back[0].Table.Rows[0].IsHeader,
                "the undeclared table's first row was read back as a header");
        }

        /// <summary>
        /// PT7：零宽度或零高度是真实几何，不是「读不到」。水平分隔线的写法就是一条
        /// <c>&lt;a:ext cx="700pt" cy="0"/&gt;</c>，旧实现把 <c>cy="0"</c> 判成缺失几何，
        /// 回落到按幻灯片比例推算的默认大框，一根细线于是变成铺满半屏的矩形。
        /// PT7: a zero width or height is real geometry, not missing data. A horizontal separator is
        /// written as <c>&lt;a:ext cx="700pt" cy="0"/&gt;</c>; the old implementation read <c>cy="0"</c>
        /// as absent geometry, fell back to the proportional default box, and turned a hairline into a
        /// rectangle covering half the slide.
        /// </summary>
        private static void RunZeroHeightSeparatorSurvivesTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            AutoShape rule = slide.AddShape(AutoShapeKind.Line, 60, 120, 700, 0, null);
            rule.Outline = new BorderLine(BorderStyle.Single, 1.0, OfficeColor.Black);

            // 写出端必须照原样落 cx/cy，符号与量级都不能自己改。
            // The writer must emit cx/cy exactly as asked; neither sign nor magnitude is its business.
            byte[] written = WriteToBuffer(deck);
            XElement ext = FirstShapeExtent(written, "ppt/slides/slide1.xml");
            Require(ext != null && (string)ext.Attribute("cy") == "0",
                "the writer padded a zero-height shape to cy="
                + (ext == null ? "<missing a:ext>" : (string)ext.Attribute("cy")));

            OfficePresentation loaded = PptxReader.Load(written);
            AutoShape back = null;
            for (int i = 0; i < loaded.Slides[0].Shapes.Count; i++)
            {
                AutoShape candidate = loaded.Slides[0].Shapes[i] as AutoShape;
                if (candidate != null) back = candidate;
            }
            Require(back != null, "the separator line did not read back");
            Require(Math.Abs(back.Height) < 0.5,
                "a zero-height separator was inflated to " + back.Height + " pt on the way back in");
            Require(Math.Abs(back.Width - 700) < 0.5,
                "the separator lost its width on the way back in: " + back.Width);
        }

        /// <summary>
        /// 取幻灯片里第一个形状自身的几何 <c>a:ext</c>（其 <c>a:xfrm</c> 挂在 <c>spPr</c> 下），
        /// 跳过 <c>p:grpSpPr</c> 上那圈整页尺寸的组变换。
        /// Takes the <c>a:ext</c> of the first shape's own geometry — the <c>a:xfrm</c> under <c>spPr</c> —
        /// skipping the slide-sized group transform on <c>p:grpSpPr</c>.
        /// </summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径。The part path.</param>
        /// <returns>几何元素或 null。The geometry element, or null.</returns>
        private static XElement FirstShapeExtent(byte[] archive, string entryName)
        {
            foreach (XElement xfrm in ReadSlideXml(archive, entryName).Descendants(A + "xfrm"))
            {
                if (xfrm.Parent == null || xfrm.Parent.Name.LocalName != "spPr") continue;
                XElement ext = xfrm.Element(A + "ext");
                if (ext != null && ext.Attribute("cx") != null && ext.Attribute("cy") != null) return ext;
            }
            return null;
        }

        /// <summary>
        /// PT9：切换效果的方向、速度与「是否需要单击」过去全被写死成 <c>dir="l"</c>、<c>spd="med"</c>、
        /// <c>advClick="1"</c>，于是「向上推入」读回来变成向左推入、设成自动换片的稿子又变回必须点击。
        /// 现在三个字段都是模型的一部分，写得出也读得回；没设过的那些不写，输出与旧稿一致。
        /// PT9: direction, speed and "advance on click" used to be hardcoded to dir="l", spd="med" and
        /// advClick="1", so an "up" push came back as a left push and an auto-advancing deck needed a click
        /// again. All three are model state now, written and read symmetrically; the ones never set stay
        /// unwritten, so existing decks produce the same bytes as before.
        /// </summary>
        private static void RunTransitionDirectionSpeedClickTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();

            Slide declared = deck.AddSlide(SlideLayoutKind.Blank);
            declared.Transition = SlideTransitionKind.Push;
            declared.TransitionDirection = SlideTransitionDirection.Up;
            declared.TransitionSpeed = SlideTransitionSpeed.Fast;
            declared.TransitionDuration = 1.5;
            declared.AdvanceAfterSeconds = 2.5;
            declared.AdvanceOnClick = false;

            // 控制：默认字段一个都没设的那一页，写出的形状必须与旧实现一致（dir="l"、无 spd、无 advClick）。
            // Control: the slide that sets nothing must keep the old shape — dir="l", no spd, no advClick.
            Slide plain = deck.AddSlide(SlideLayoutKind.Blank);
            plain.Transition = SlideTransitionKind.Push;

            byte[] written = WriteToBuffer(deck);

            XElement transition = FirstTransition(written, "ppt/slides/slide1.xml");
            Require(transition != null, "the slide part carries no p:transition");
            XElement push = transition.Element(P + "push");
            Require(push != null, "the push effect element is missing");
            Require((string)push.Attribute("dir") == "u",
                "an upward push was written as dir=" + (string)push.Attribute("dir"));
            Require((string)transition.Attribute("spd") == "fast",
                "the fast speed was not written: spd=" + (string)transition.Attribute("spd"));
            Require((string)transition.Attribute("advClick") == "0",
                "an auto-advancing slide still demanded a click: advClick="
                + (string)transition.Attribute("advClick"));
            Require((string)transition.Attribute("advTm") == "2500",
                "the auto-advance time did not survive: advTm=" + (string)transition.Attribute("advTm"));

            XElement plainTransition = FirstTransition(written, "ppt/slides/slide2.xml");
            Require(plainTransition != null, "the control slide lost its transition");
            XElement plainPush = plainTransition.Element(P + "push");
            Require(plainPush != null && (string)plainPush.Attribute("dir") == "l",
                "the default push direction changed: dir="
                + (plainPush == null ? "<no push>" : (string)plainPush.Attribute("dir")));
            Require((string)plainTransition.Attribute("spd") == null,
                "a slide that never set a speed still wrote spd");
            Require((string)plainTransition.Attribute("advClick") == null,
                "a slide that never set advClick still wrote it");

            OfficePresentation loaded = PptxReader.Load(written);
            Slide back = loaded.Slides[0];
            Require(back.Transition == SlideTransitionKind.Push, "the transition kind did not read back");
            Require(back.TransitionDirection == SlideTransitionDirection.Up,
                "the direction did not read back: " + back.TransitionDirection);
            Require(back.TransitionSpeed == SlideTransitionSpeed.Fast,
                "the speed did not read back: " + back.TransitionSpeed);
            Require(!back.AdvanceOnClick, "the slide started demanding a click again");
            Require(Math.Abs(back.AdvanceAfterSeconds - 2.5) < 0.001,
                "the auto-advance time did not read back: " + back.AdvanceAfterSeconds);
            Require(loaded.Slides[1].AdvanceOnClick, "the default slide lost its click-to-advance");
        }

        /// <summary>取幻灯片部件里第一个 <c>p:transition</c>（带时长时它住在 mc:Choice 里）。 Takes the first <c>p:transition</c> in a slide part, which lives inside mc:Choice when a duration is present.</summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径。The part path.</param>
        /// <returns>transition 元素，缺失时 null。The transition element, null when absent.</returns>
        private static XElement FirstTransition(byte[] archive, string entryName)
        {
            foreach (XElement transition in ReadSlideXml(archive, entryName).Descendants(P + "transition"))
            {
                return transition;
            }
            return null;
        }

        /// <summary>PresentationML 命名空间，测试里手写一份。 The PresentationML namespace, spelled out here.</summary>
        private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

        /// <summary>DrawingML 命名空间，测试里手写一份，避免依赖写端的私有常量。 The DrawingML namespace, spelled out here so the test does not lean on a writer private.</summary>
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        /// <summary>取幻灯片部件里的每张 <c>a:tbl</c>，按文档顺序返回。 Collects every <c>a:tbl</c> in a slide part, in document order.</summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径。The part path.</param>
        /// <returns>表格元素列表。The table elements.</returns>
        private static XElement[] ReadTables(byte[] archive, string entryName)
        {
            List<XElement> found = new List<XElement>();
            foreach (XElement tbl in ReadSlideXml(archive, entryName).Descendants(A + "tbl")) found.Add(tbl);
            return found.ToArray();
        }

        /// <summary>读出幻灯片部件的 XML 根。 Reads a slide part into an XML document.</summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径。The part path.</param>
        /// <returns>解析好的文档。The parsed document.</returns>
        private static XDocument ReadSlideXml(byte[] archive, string entryName)
        {
            using (ZipArchive zip = new ZipArchive(new MemoryStream(archive), ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = zip.GetEntry(entryName);
                Require(entry != null, "slide part missing: " + entryName);
                using (Stream stream = entry.Open())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return XDocument.Parse(reader.ReadToEnd());
                }
            }
        }

        /// <summary>取表格的首行 <c>a:tr</c>。 Takes a table's first <c>a:tr</c>.</summary>
        /// <param name="tbl">a:tbl 元素。The a:tbl element.</param>
        /// <returns>首行元素，空表为 null。 The first row, or null for an empty table.</returns>
        private static XElement FirstRow(XElement tbl)
        {
            foreach (XElement tr in tbl.Elements(A + "tr")) return tr;
            return null;
        }

        /// <summary>取行里的第一个单元格 <c>a:tc</c>。 Takes a row's first <c>a:tc</c>.</summary>
        /// <param name="row">a:tr 元素，可为 null。The a:tr element, may be null.</param>
        /// <returns>第一个单元格元素，取不到为 null。 The first cell, or null.</returns>
        private static XElement FirstCell(XElement row)
        {
            if (row == null) return null;
            foreach (XElement tc in row.Elements(A + "tc")) return tc;
            return null;
        }

        /// <summary>取单元格里第一个文本运行的 <c>a:rPr</c>。 Takes the <c>a:rPr</c> of a cell's first text run.</summary>
        /// <param name="cell">a:tc 元素，可为 null。The a:tc element, may be null.</param>
        /// <returns>运行属性元素，无运行为 null。 The run properties, or null when the cell has no run.</returns>
        private static XElement FirstRunProps(XElement cell)
        {
            if (cell == null) return null;
            foreach (XElement run in cell.Descendants(A + "r"))
            {
                XElement props = run.Element(A + "rPr");
                if (props != null) return props;
            }
            return null;
        }

        /// <summary>按 z 序列出这一页的表格形状。 Lists the slide's table shapes in z-order.</summary>
        /// <param name="slide">读回的幻灯片。The reloaded slide.</param>
        /// <returns>表格形状列表。The table shapes.</returns>
        private static List<TableShape> FindTables(Slide slide)
        {
            List<TableShape> found = new List<TableShape>();
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                TableShape table = slide.Shapes[i] as TableShape;
                if (table != null) found.Add(table);
            }
            return found;
        }

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
