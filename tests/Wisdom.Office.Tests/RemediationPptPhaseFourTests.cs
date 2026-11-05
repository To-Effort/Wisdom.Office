using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
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
            RunUnsetShapeFillStaysInheritedTest();
            RunUnmarkedParagraphDoesNotSuppressBulletsTest();
            RunTextBoxWrapNoneSurvivesTest();
            RunDoubleBorderIsCompoundLineTest();
            RunTableBorderDirectionsLandTest();
            RunCellBorderNoneMeansNoLineTest();
            RunTableCellBordersComeBackTest();
            RunZeroHeightSeparatorSurvivesTest();
            RunTransitionDirectionSpeedClickTest();
            RunSectionHeaderLayoutSurvivesTest();
            RunAppPropertiesSchemaOrderTest();
            RunTimingTreeShapeTest();
            RunNonFirstShapeAnimationDelayTest();
            RunBaselineAndSpacingRoundTripTest();
            RunVerticalMergeGridTest();

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
        /// PT3 结构档第二批：形状没设填充、没设描边时，写端不得替它写 <c>a:noFill</c> 与一条空 <c>a:ln</c>。
        /// 本库自造的母版与版式压根没给占位符定义填充或线条，省略这两项与写 <c>a:noFill</c> 在自己产出的
        /// 稿子上看起来一样；从真实文件读回的形状就不同了——读端把主题 <c>fillRef</c>、渐变、图案填充统一
        /// 读成 null，写端却一律钉成 <c>a:noFill</c>，一次 Load+Save 就把借来的主题填充与线条抹个干净。
        /// PT3 structure tier, second batch: a shape that sets neither fill nor outline must not leave the
        /// file with <c>a:noFill</c> and an empty <c>a:ln</c>. The masters and layouts this library generates
        /// define no fill or line for their placeholders, so omitting both looks the same as writing
        /// <c>a:noFill</c> in its own decks; shapes read back from a real file are another matter — the reader
        /// resolves theme <c>fillRef</c>, gradient and pattern fills to null, and pinning every one of them to
        /// <c>a:noFill</c> wipes the borrowed fill and line out on the first Load+Save.
        /// </summary>
        private static void RunUnsetShapeFillStaysInheritedTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);

            slide.AddShape(AutoShapeKind.Rectangle, 40, 40, 240, 100, "没设填充也没设线");
            AutoShape filled = slide.AddShape(AutoShapeKind.Rectangle, 320, 40, 240, 100, "只设填充");
            filled.FillColor = new OfficeColor(0x00, 0xB0, 0x50);
            AutoShape unlined = slide.AddShape(AutoShapeKind.Ellipse, 620, 40, 160, 160, "显式无线");
            unlined.Outline = new BorderLine(BorderStyle.None, 0, OfficeColor.Black);
            slide.AddTextBox(40, 220, 520, 80, "文本框同样不写");

            byte[] written = WriteToBuffer(deck);
            XElement[] props = ReadShapeProperties(written, "ppt/slides/slide1.xml");
            Require(props.Length == 4, "expected four p:sp on the slide, got " + props.Length);

            Require(props[0].Element(A + "noFill") == null,
                "a shape that never set a fill was pinned with a:noFill, silencing the inherited one");
            Require(props[0].Element(A + "solidFill") == null,
                "a shape that never set a fill was given a solidFill anyway");
            Require(props[0].Element(A + "ln") == null,
                "a shape that never set an outline still got an a:ln element");

            XElement filledFill = props[1].Element(A + "solidFill");
            XElement filledClr = filledFill == null ? null : filledFill.Element(A + "srgbClr");
            Require(filledClr != null && (string)filledClr.Attribute("val") == "00B050",
                "the shape with an explicit FillColor lost it — the new rule must not swallow real intent");

            XElement unlinedLn = props[2].Element(A + "ln");
            Require(unlinedLn != null && unlinedLn.Element(A + "noFill") != null,
                "an outline explicitly set to BorderStyle.None must still reach the file as a:ln with noFill");

            Require(props[3].Element(A + "noFill") == null && props[3].Element(A + "ln") == null,
                "the text box got the same fabricated noFill and empty a:ln");

            // 往返：省下来的属性读回仍是「没设过」，二次写出不得把它们补成 noFill，也不得吞掉显式那两份。
            // Round trip: an omitted property reads back as unset, so a second write must neither invent a
            // noFill for it nor swallow the two shapes that did state something.
            OfficePresentation loaded = PptxReader.Load(written);
            byte[] again = WriteToBuffer(loaded);
            XElement[] second = ReadShapeProperties(again, "ppt/slides/slide1.xml");
            Require(second.Length == 4, "expected four p:sp after the round trip, got " + second.Length);
            Require(second[0].Element(A + "noFill") == null && second[0].Element(A + "ln") == null,
                "the second write re-invented a:noFill and an empty a:ln for the untouched shape");
            Require(second[1].Element(A + "solidFill") != null,
                "the explicitly filled shape came back without a solidFill");
        }

        /// <summary>
        /// PT3 结构档第三批：段落没给列表标记时，写端不得替它写 <c>a:buNone</c>，也不该为塞这句谎话而
        /// 恒产出一个 <c>a:pPr</c>。<c>buNone</c> 在 PresentationML 里是显式的「本段不要项目符号」，会把母版
        /// 与版式本来要给的符号压掉。本库自造的母版没定义任何符号，自家稿子上看不出差别；但第三方稿子
        /// 的正文段落走一遍 Load+Save，母版项目符号就被逐段取消了——模型里从没说过「不要符号」。
        /// PT3 structure tier, third batch: a paragraph with no list marker must not leave the file carrying
        /// <c>a:buNone</c>, and the writer must not fabricate an <c>a:pPr</c> just to house that lie.
        /// <c>buNone</c> is PresentationML for "this paragraph takes no bullet", which suppresses whatever the
        /// master or layout would supply. The masters this library generates define no bullet at all, so its
        /// own decks look identical either way; but body paragraphs read back from a third-party file lose the
        /// master's bullets one by one on the first Load+Save — the model never said "no bullet".
        /// </summary>
        private static void RunUnmarkedParagraphDoesNotSuppressBulletsTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            TextBoxShape body = slide.AddTextBox(40, 40, 700, 260, "带标记的一段");
            body.Paragraphs[0].ListMarker = "•";
            Paragraph bare = body.AddParagraph(string.Empty);
            bare.AddText("既没标记也没别的属性");
            Paragraph aligned = body.AddParagraph(string.Empty);
            aligned.AddText("只设了居中");
            aligned.Alignment = ParagraphAlignment.Center;

            byte[] written = WriteToBuffer(deck);
            XElement[] paragraphs = ReadParagraphElements(written, "ppt/slides/slide1.xml");
            Require(paragraphs.Length == 3, "expected three a:p on the slide, got " + paragraphs.Length);

            XElement markedPr = paragraphs[0].Element(A + "pPr");
            Require(markedPr != null && markedPr.Element(A + "buChar") != null,
                "the control failed: a declared marker must still reach the file as buChar");

            Require(paragraphs[1].Element(A + "pPr") == null,
                "a paragraph that set nothing still got an a:pPr, housing a fabricated a:buNone");

            XElement alignedPr = paragraphs[2].Element(A + "pPr");
            Require(alignedPr != null && (string)alignedPr.Attribute("algn") == "ctr",
                "the paragraph that set alignment lost it — the new rule must not swallow real intent");
            Require(alignedPr != null && alignedPr.Element(A + "buNone") == null
                && alignedPr.Element(A + "buChar") == null && alignedPr.Element(A + "buAutoNum") == null,
                "an unmarked paragraph was told to suppress the inherited bullet");

            // 往返：标记仍要读得回来，省下的那两处不能在第二次写出时补回去。
            // Round trip: the declared marker still has to come back, and the two omissions must not be
            // patched in by the second write.
            OfficePresentation loaded = PptxReader.Load(written);
            Require(ReadFirstMarker(loaded) == "•",
                "the declared marker did not survive the round trip");
            byte[] again = WriteToBuffer(loaded);
            XElement[] second = ReadParagraphElements(again, "ppt/slides/slide1.xml");
            Require(second.Length == 3, "expected three a:p after the round trip, got " + second.Length);
            Require(second[1].Element(A + "pPr") == null,
                "the second write invented an a:pPr for the paragraph that sets nothing");
            Require(second[2].Element(A + "pPr") != null
                && second[2].Element(A + "pPr").Element(A + "buNone") == null,
                "the second write suppressed the inherited bullet again");
        }

        /// <summary>
        /// PT3 读端补齐（wrap 一档）：文本框「不换行」是 PowerPoint 里单行标题与竖排标签的常态写法
        /// <c>&lt;a:bodyPr wrap="none"/&gt;</c>。写端过去无论模型说没说都钉一个 <c>wrap="square"</c>，
        /// 读端从不看这个属性，于是载入别人的演示稿再存盘一次，那行标题就开始在框内折行——
        /// 版式被静默改掉，形状数、文本一个字都没少，肉眼只能在放映时才发现。
        /// PT3 read parity (the wrap case): "do not wrap" is how PowerPoint spells a single-line title or a
        /// vertical label — <c>&lt;a:bodyPr wrap="none"/&gt;</c>. The writer pinned <c>wrap="square"</c>
        /// whatever the model said and the reader never looked at the attribute, so loading someone else's
        /// deck and saving it once made that title start folding inside the box. Nothing is lost textually;
        /// only the layout silently changes, which is exactly what a round trip must not do.
        /// </summary>
        private static void RunTextBoxWrapNoneSurvivesTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            slide.AddTextBox(40, 40, 380, 40, "一行到底、不许折行的长标题，超出右边界也照样继续排下去");
            slide.AddTextBox(40, 120, 380, 200, "这段照常按框宽折行，作为对照组");

            // 夹具：把第一个文本框写成 wrap="none"，第二个留作 square 对照。
            // Fixture: spell the first box as wrap="none" and keep the second one as the square control.
            byte[] written = WriteToBuffer(deck);
            string slideXml = SlideEntryText(written, "ppt/slides/slide1.xml");
            int first = slideXml.IndexOf("wrap=\"square\"", StringComparison.Ordinal);
            Require(first > 0, "the fixture found no wrap attribute to rewrite");
            string patched = slideXml.Remove(first, "wrap=\"square\"".Length)
                .Insert(first, "wrap=\"none\"");
            Require(!patched.Contains("wrap=\"emergency\""), "the fixture edit ran away");
            byte[] source = RewriteZipEntryText(written, "ppt/slides/slide1.xml", patched);

            XElement[] before = ReadBodyPrs(source, "ppt/slides/slide1.xml");
            Require(before.Length == 2, "expected two a:bodyPr in the fixture, got " + before.Length);
            Require((string)before[0].Attribute("wrap") == "none",
                "the fixture did not carry wrap=\"none\": " + (string)before[0].Attribute("wrap"));
            Require((string)before[1].Attribute("wrap") == "square",
                "the fixture's control box lost its wrap=\"square\"");

            OfficePresentation loaded = PptxReader.Load(source);
            Require(loaded.Slides[0].Shapes.Count == 2,
                "the round trip changed the shape count: " + loaded.Slides[0].Shapes.Count);
            byte[] again = WriteToBuffer(loaded);
            XElement[] after = ReadBodyPrs(again, "ppt/slides/slide1.xml");
            Require(after.Length == 2, "expected two a:bodyPr after the round trip, got " + after.Length);
            Require((string)after[0].Attribute("wrap") == "none",
                "wrap=\"none\" was flattened to " + (string)after[0].Attribute("wrap")
                + ": the box that must not fold started folding");
            Require((string)after[1].Attribute("wrap") == "square",
                "the control box changed too: " + (string)after[1].Attribute("wrap"));

            // 读端必须把这件事交到模型上，否则第二次写出只能靠猜。
            // The reader has to hand this to the model, otherwise the second write can only guess.
            TextBoxShape foldedBack = loaded.Slides[0].Shapes[0] as TextBoxShape;
            Require(foldedBack != null && foldedBack.WrapText == false,
                "the reader left wrap=\"none\" unset in the model: "
                + (foldedBack == null ? "<not a text box>" : foldedBack.WrapText.ToString()));

            // 本库自己的写出：显式不折行才写 none，显式折行与「没说过」都落 square（PowerPoint 的原生默认）。
            // Our own write: only an explicit no-fold becomes none; an explicit fold and an unset box both
            // land on square, PowerPoint's native default.
            OfficePresentation declared = OfficePresentation.Widescreen();
            Slide page = declared.AddSlide(SlideLayoutKind.Blank);
            TextBoxShape single = page.AddTextBox(40, 40, 380, 40, "不许折行的长标题");
            single.WrapText = false;
            TextBoxShape folding = page.AddTextBox(40, 120, 380, 100, "显式要折行");
            folding.WrapText = true;
            TextBoxShape quiet = page.AddTextBox(40, 260, 380, 100, "没提过折行");
            byte[] own = WriteToBuffer(declared);
            XElement[] ownPrs = ReadBodyPrs(own, "ppt/slides/slide1.xml");
            Require(ownPrs.Length == 3, "expected three a:bodyPr, got " + ownPrs.Length);
            Require((string)ownPrs[0].Attribute("wrap") == "none",
                "the declared no-fold box wrote " + (string)ownPrs[0].Attribute("wrap"));
            Require((string)ownPrs[1].Attribute("wrap") == "square",
                "the declared fold box wrote " + (string)ownPrs[1].Attribute("wrap"));
            Require((string)ownPrs[2].Attribute("wrap") == "square",
                "the box that never spoke changed its written default: " + (string)ownPrs[2].Attribute("wrap"));

            OfficePresentation ownBack = PptxReader.Load(own);
            TextBoxShape[] boxes = new TextBoxShape[3];
            for (int i = 0; i < 3; i++) boxes[i] = ownBack.Slides[0].Shapes[i] as TextBoxShape;
            Require(boxes[0] != null && boxes[1] != null && boxes[2] != null,
                "the three boxes did not come back as text boxes");
            Require(boxes[0].WrapText == false,
                "the no-fold setting did not read back: " + boxes[0].WrapText);
            // square 就是默认值的另一种写法，文件里没有任何一句「要折行」值得当成声明：读回 null。
            // square is only the default spelled out; nothing in the file declares "fold", so it reads back null.
            Require(boxes[1].WrapText == null,
                "the fold box came back with a declaration it never made: " + boxes[1].WrapText);
            Require(boxes[2].WrapText == null,
                "the box that never spoke gained a setting: " + boxes[2].WrapText);
            byte[] third = WriteToBuffer(ownBack);
            XElement[] thirdPrs = ReadBodyPrs(third, "ppt/slides/slide1.xml");
            Require(thirdPrs.Length == 3
                && (string)thirdPrs[0].Attribute("wrap") == "none"
                && (string)thirdPrs[1].Attribute("wrap") == "square"
                && (string)thirdPrs[2].Attribute("wrap") == "square",
                "the second write did not reproduce the first one's wrapping");
        }

        /// <summary>
        /// 双线在 DrawingML 里是**复合线**，写法是 <c>a:ln@cmpd="dbl"</c>（同一个写出器给表格边框就写了
        /// <c>cmpd="sng"</c>），而 <c>a:prstDash</c> 的取值集合里没有 "dbl" 这一档。旧实现把 Double 塞进
        /// prstDash：文件里留下一个架构不接受的枚举值，读端 <c>MapDashStyle</c> 也认不出来，一律回落
        /// Single——双线边框存一次盘就变成单线。同批还有点线：<c>BorderStyle.Dotted</c> 被映射成
        /// <c>sysDash</c>（长划图案），画出来是虚线而不是点线。
        /// Double is a *compound* line in DrawingML and spells itself <c>a:ln@cmpd="dbl"</c> — the very same
        /// writer emits <c>cmpd="sng"</c> for table edges — while <c>a:prstDash</c> has no "dbl" in its value
        /// set. The old code stuffed Double into prstDash, which leaves a value the schema does not accept and
        /// which <c>MapDashStyle</c> cannot read back, so a double border silently became single after one
        /// save. The dotted case ships in the same batch: <c>BorderStyle.Dotted</c> mapped to <c>sysDash</c>,
        /// a long-dash pattern, so a dotted line drew as a dashed one.
        /// </summary>
        private static void RunDoubleBorderIsCompoundLineTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            AutoShape doubled = slide.AddShape(AutoShapeKind.Rectangle, 40, 40, 320, 120, "双线边框");
            doubled.Outline = new BorderLine(BorderStyle.Double, 2.0, OfficeColor.Black);
            AutoShape dotted = slide.AddShape(AutoShapeKind.Rectangle, 40, 200, 320, 120, "点线边框");
            dotted.Outline = new BorderLine(BorderStyle.Dotted, 1.0, OfficeColor.Black);
            TableShape table = slide.AddTable(420, 40, 480, 160);
            table.Table.Borders = new TableBorders();
            table.Table.Borders.InsideHorizontal = new BorderLine(BorderStyle.Double, 1.5, OfficeColor.Black);
            table.Table.AddRow("甲", "乙");
            table.Table.AddRow("丙", "丁");

            byte[] written = WriteToBuffer(deck);
            XDocument doc = ReadSlideXml(written, "ppt/slides/slide1.xml");

            // prstDash 的取值集合里没有双线：这个值出现在文件里就是非法枚举。
            // prstDash has no double value; seeing it in the file is an out-of-enumeration value.
            foreach (XElement dash in doc.Descendants(A + "prstDash"))
            {
                string value = (string)dash.Attribute("val");
                Require(value != "dbl",
                    "a:prstDash/@val was written as \"dbl\", which is a compound-line value, not a dash pattern");
            }

            XElement[] spPrs = ReadShapeProperties(written, "ppt/slides/slide1.xml");
            Require(spPrs.Length == 2, "expected two p:spPr, got " + spPrs.Length);
            XElement doubleLn = spPrs[0].Element(A + "ln");
            Require(doubleLn != null && (string)doubleLn.Attribute("cmpd") == "dbl",
                "a double outline must be a compound line: a:ln@cmpd="
                + (doubleLn == null ? "<no a:ln>" : (string)doubleLn.Attribute("cmpd") ?? "<absent>"));
            XElement dotLn = spPrs[1].Element(A + "ln");
            XElement dotDash = dotLn == null ? null : dotLn.Element(A + "prstDash");
            string dotValue = dotDash == null ? null : (string)dotDash.Attribute("val");
            Require(dotValue == "sysDot" || dotValue == "dot" || dotValue == "sqDot" || dotValue == "roundDot",
                "BorderStyle.Dotted wrote prstDash=" + (dotValue == null ? "<absent>" : dotValue)
                + ", which is a dash pattern rather than dots");

            bool tableEdgeIsDouble = false;
            foreach (XElement edge in doc.Descendants(A + "lnL").Concat(doc.Descendants(A + "lnR"))
                .Concat(doc.Descendants(A + "lnT")).Concat(doc.Descendants(A + "lnB")))
            {
                if ((string)edge.Attribute("cmpd") == "dbl") tableEdgeIsDouble = true;
            }
            Require(tableEdgeIsDouble, "the table's double inside border reached no a:lnX@cmpd=\"dbl\"");

            OfficePresentation loaded = PptxReader.Load(written);
            Require(loaded.Slides[0].Shapes.Count == 3,
                "expected three shapes back, got " + loaded.Slides[0].Shapes.Count);
            BorderLine doubleBack = loaded.Slides[0].Shapes[0].Outline;
            Require(doubleBack != null && doubleBack.Style == BorderStyle.Double,
                "the double outline read back as " + (doubleBack == null ? "<null>" : doubleBack.Style.ToString()));
            BorderLine dotBack = loaded.Slides[0].Shapes[1].Outline;
            Require(dotBack != null && dotBack.Style == BorderStyle.Dotted,
                "the dotted outline read back as " + (dotBack == null ? "<null>" : dotBack.Style.ToString()));
        }

        /// <summary>
        /// PT12：<c>Table.Borders</c> 有六个方向，写出器只认其中一个——<c>InsideHorizontal</c> 被当作
        /// 四条边共用的默认线，<c>Top</c>/<c>Bottom</c>/<c>Left</c>/<c>Right</c>/<c>InsideVertical</c>
        /// 五份声明在文件里查无此线。外圈加粗、内部变淡是演示稿表格最常见的写法，存一次盘全部塌回一种线。
        /// PT12: <c>Table.Borders</c> carries six directions and the writer reads exactly one of them —
        /// <c>InsideHorizontal</c> becomes the shared default for every edge, so the other five declarations
        /// leave no trace in the file. A bold ring with lighter inner rules, the most common deck-table
        /// idiom, collapsed into a single line style on save.
        /// </summary>
        private static void RunTableBorderDirectionsLandTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            TableShape frame = slide.AddTable(60, 60, 700, 160);
            Table model = frame.Table;
            model.Borders.Top = new BorderLine(BorderStyle.Single, 3.0, new OfficeColor(0xFF, 0x00, 0x00));
            model.Borders.Bottom = new BorderLine(BorderStyle.Single, 3.0, new OfficeColor(0x00, 0x00, 0xFF));
            model.Borders.Left = new BorderLine(BorderStyle.Single, 3.0, new OfficeColor(0x00, 0xFF, 0x00));
            model.Borders.Right = new BorderLine(BorderStyle.Single, 3.0, new OfficeColor(0xFF, 0x80, 0x00));
            model.Borders.InsideHorizontal = new BorderLine(BorderStyle.Single, 1.0, new OfficeColor(0x80, 0x80, 0x80));
            model.Borders.InsideVertical = new BorderLine(BorderStyle.Single, 1.0, new OfficeColor(0x80, 0x00, 0x80));
            model.AddRow("甲", "乙", "丙");
            model.AddRow("丁", "戊", "己");

            byte[] written = WriteToBuffer(deck);
            XElement[] tables = ReadTables(written, "ppt/slides/slide1.xml");
            Require(tables.Length == 1, "expected one a:tbl, got " + tables.Length);
            XElement[] trs = tables[0].Elements(A + "tr").ToArray();
            Require(trs.Length == 2, "expected two a:tr, got " + trs.Length);

            for (int r = 0; r < trs.Length; r++)
            {
                XElement[] tcs = trs[r].Elements(A + "tc").ToArray();
                Require(tcs.Length == 3, "row " + r + " carries " + tcs.Length + " cells, expected 3");
                for (int c = 0; c < tcs.Length; c++)
                {
                    XElement tcPr = tcs[c].Element(A + "tcPr");
                    Require(tcPr != null, "row " + r + " cell " + c + " was written without a:tcPr");
                    string where = "row " + r + " cell " + c + " ";
                    Require(EdgeHex(tcPr, "lnT") == (r == 0 ? "FF0000" : "808080"),
                        where + "lnT was " + EdgeHex(tcPr, "lnT") + ", expected " + (r == 0 ? "FF0000" : "808080"));
                    Require(EdgeHex(tcPr, "lnB") == (r == trs.Length - 1 ? "0000FF" : "808080"),
                        where + "lnB was " + EdgeHex(tcPr, "lnB"));
                    Require(EdgeHex(tcPr, "lnL") == (c == 0 ? "00FF00" : "800080"),
                        where + "lnL was " + EdgeHex(tcPr, "lnL"));
                    Require(EdgeHex(tcPr, "lnR") == (c == tcs.Length - 1 ? "FF8000" : "800080"),
                        where + "lnR was " + EdgeHex(tcPr, "lnR"));
                    Require(EdgeWidth(tcPr, "lnT") == (r == 0 ? "38100" : "12700"),
                        where + "lnT width was " + EdgeWidth(tcPr, "lnT"));
                    Require(EdgeWidth(tcPr, "lnL") == (c == 0 ? "38100" : "12700"),
                        where + "lnL width was " + EdgeWidth(tcPr, "lnL"));
                }
            }
        }

        /// <summary>取单元格某条边线的颜色十六进制，缺什么就报什么形状。 Reads one cell edge's colour hex, naming whatever is missing.</summary>
        /// <param name="tcPr">a:tcPr 元素。 The a:tcPr element.</param>
        /// <param name="tag">边线标签（lnL/lnR/lnT/lnB）。 The edge tag.</param>
        /// <returns>十六进制或说明性占位。 The hex value or a descriptive placeholder.</returns>
        private static string EdgeHex(XElement tcPr, string tag)
        {
            XElement edge = tcPr == null ? null : tcPr.Element(A + tag);
            if (edge == null) return "<no edge>";
            XElement srgb = edge.Element(A + "solidFill");
            srgb = srgb == null ? null : srgb.Element(A + "srgbClr");
            if (srgb != null) return (string)srgb.Attribute("val") ?? "<no val>";
            return edge.Element(A + "noFill") != null ? "<noFill>" : "<no color>";
        }

        /// <summary>取单元格某条边线的 <c>w</c> 属性（EMU 文本）。 Reads one cell edge's <c>w</c> attribute as EMU text.</summary>
        /// <param name="tcPr">a:tcPr 元素。 The a:tcPr element.</param>
        /// <param name="tag">边线标签。 The edge tag.</param>
        /// <returns>EMU 文本或占位说明。 The EMU text or a placeholder.</returns>
        private static string EdgeWidth(XElement tcPr, string tag)
        {
            XElement edge = tcPr == null ? null : tcPr.Element(A + tag);
            string width = edge == null ? null : (string)edge.Attribute("w");
            return width ?? "<absent>";
        }

        /// <summary>
        /// PT13：<c>BorderStyle.None</c> 是「这条边不画」，同库的 DOCX 写出器就照这个意思写
        /// <c>w:val="nil"</c>。演示稿这边却照旧补一条 <c>a:solidFill</c>：于是
        /// <c>Table.Borders = TableBorders.None()</c> 存出门是一张满黑格的表，与声明正相反。
        /// PT13: <c>BorderStyle.None</c> means "draw no line here", and this library's own DOCX
        /// writer spells it <c>w:val="nil"</c>. The PPTX cell path still emitted an
        /// <c>a:solidFill</c> anyway, so <c>Table.Borders = TableBorders.None()</c> saved out as a
        /// fully ruled black grid — the opposite of what was declared.
        /// </summary>
        private static void RunCellBorderNoneMeansNoLineTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            TableShape frame = slide.AddTable(60, 60, 700, 160);
            Table model = frame.Table;
            model.Borders = TableBorders.None();
            model.Borders.Top = BorderLine.Thin();
            model.AddRow("甲", "乙");
            model.AddRow("丙", "丁");

            byte[] written = WriteToBuffer(deck);
            XElement[] tables = ReadTables(written, "ppt/slides/slide1.xml");
            Require(tables.Length == 1, "expected one a:tbl, got " + tables.Length);
            XElement[] trs = tables[0].Elements(A + "tr").ToArray();
            Require(trs.Length == 2, "expected two a:tr, got " + trs.Length);

            for (int r = 0; r < trs.Length; r++)
            {
                XElement[] tcs = trs[r].Elements(A + "tc").ToArray();
                for (int c = 0; c < tcs.Length; c++)
                {
                    XElement tcPr = tcs[c].Element(A + "tcPr");
                    string where = "row " + r + " cell " + c + " ";
                    Require(EdgeHex(tcPr, "lnT") == (r == 0 ? "000000" : "<noFill>"),
                        where + "lnT was " + EdgeHex(tcPr, "lnT") + (r == 0 ? ", the declared thin ring line is gone" : ", a borderless edge still drew a line"));
                    Require(EdgeHex(tcPr, "lnB") == "<noFill>", where + "lnB was " + EdgeHex(tcPr, "lnB"));
                    Require(EdgeHex(tcPr, "lnL") == "<noFill>", where + "lnL was " + EdgeHex(tcPr, "lnL"));
                    Require(EdgeHex(tcPr, "lnR") == "<noFill>", where + "lnR was " + EdgeHex(tcPr, "lnR"));
                }
            }
        }

        /// <summary>
        /// PT14：写出器逐格写出的 <c>a:lnL</c>/<c>a:lnR</c>/<c>a:lnT</c>/<c>a:lnB</c> 读端从不解析，
        /// 于是别人的表格载进来边框全丢，本库自己写的逐格边框也活不过第二次保存——模型里
        /// <c>TableCellBorders</c> 四个字段本来就在，只是没人往里放东西。
        /// PT14: the writer's per-cell <c>a:lnL</c>/<c>a:lnR</c>/<c>a:lnT</c>/<c>a:lnB</c> were never parsed
        /// on the way in, so a foreign table lost every border on load and this library's own per-cell
        /// borders did not survive a second save — <c>TableCellBorders</c> had the four fields all along,
        /// nothing ever filled them.
        /// </summary>
        private static void RunTableCellBordersComeBackTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            TableShape ring = slide.AddTable(60, 60, 700, 160);
            ring.Table.Borders.Top = new BorderLine(BorderStyle.Single, 3.0, new OfficeColor(0xFF, 0x00, 0x00));
            ring.Table.Borders.Bottom = new BorderLine(BorderStyle.Single, 3.0, new OfficeColor(0x00, 0x00, 0xFF));
            ring.Table.Borders.Left = new BorderLine(BorderStyle.Single, 3.0, new OfficeColor(0x00, 0xFF, 0x00));
            ring.Table.Borders.Right = new BorderLine(BorderStyle.Single, 3.0, new OfficeColor(0xFF, 0x80, 0x00));
            ring.Table.Borders.InsideHorizontal = new BorderLine(BorderStyle.Single, 1.0, new OfficeColor(0x80, 0x80, 0x80));
            ring.Table.Borders.InsideVertical = new BorderLine(BorderStyle.Double, 1.5, new OfficeColor(0x80, 0x00, 0x80));
            ring.Table.AddRow("甲", "乙", "丙");
            ring.Table.AddRow("丁", "戊", "己");
            TableShape bare = slide.AddTable(60, 260, 700, 120);
            bare.Table.Borders = TableBorders.None();
            bare.Table.Borders.Top = BorderLine.Thin();
            bare.Table.AddRow("庚", "辛");
            bare.Table.AddRow("壬", "癸");

            byte[] written = WriteToBuffer(deck);
            OfficePresentation loaded = PptxReader.Load(written);
            List<TableShape> tables = FindTables(loaded.Slides[0]);
            Require(tables.Count == 2, "expected two table shapes back, got " + tables.Count);

            TableCell corner = tables[0].Table.Rows[0].Cells[0];
            TableCellBorders cornerEdges = corner.Borders;
            Require(cornerEdges != null,
                "the loaded table kept no cell borders: row 0 cell 0 has Borders="
                + (cornerEdges == null ? "<null>" : "<set>"));
            Require(HexOf(cornerEdges == null ? null : cornerEdges.Top) == "FF0000",
                "row 0 cell 0 top edge read back as " + HexOf(cornerEdges == null ? null : cornerEdges.Top)
                + ", expected the declared FF0000 ring line");
            Require(HexOf(cornerEdges == null ? null : cornerEdges.Bottom) == "808080",
                "row 0 cell 0 bottom edge read back as " + HexOf(cornerEdges == null ? null : cornerEdges.Bottom)
                + ", expected the inner rule 808080");
            Require(HexOf(cornerEdges == null ? null : cornerEdges.Left) == "00FF00",
                "row 0 cell 0 left edge read back as " + HexOf(cornerEdges == null ? null : cornerEdges.Left));
            Require(HexOf(cornerEdges == null ? null : cornerEdges.Right) == "800080",
                "row 0 cell 0 right edge read back as " + HexOf(cornerEdges == null ? null : cornerEdges.Right));
            BorderLine innerVertical = cornerEdges == null ? null : cornerEdges.Right;
            Require(innerVertical != null && innerVertical.Style == BorderStyle.Double,
                "the double inner rule read back as "
                + (innerVertical == null ? "<null>" : innerVertical.Style.ToString()));
            TableCell tail = tables[0].Table.Rows[1].Cells[2];
            Require(HexOf(tail.Borders == null ? null : tail.Borders.Bottom) == "0000FF",
                "the last row's bottom edge read back as " + HexOf(tail.Borders == null ? null : tail.Borders.Bottom)
                + ", expected the declared 0000FF ring line");
            Require(HexOf(tail.Borders == null ? null : tail.Borders.Right) == "FF8000",
                "the last column's right edge read back as " + HexOf(tail.Borders == null ? null : tail.Borders.Right));

            TableCell bareCorner = tables[1].Table.Rows[0].Cells[0];
            BorderLine bareBottom = bareCorner.Borders == null ? null : bareCorner.Borders.Bottom;
            Require(bareBottom != null && bareBottom.Style == BorderStyle.None,
                "a cell edge written as a:noFill read back as "
                + (bareBottom == null ? "<null>" : bareBottom.Style.ToString()));
            BorderLine bareTop = bareCorner.Borders == null ? null : bareCorner.Borders.Top;
            Require(bareTop != null && bareTop.Style == BorderStyle.Single,
                "the one declared line of the borderless table read back as "
                + (bareTop == null ? "<null>" : bareTop.Style.ToString()));

            // 第二次保存不得改动任何一条边线：读回来的值必须正好是文件里那一份。
            // The second save must not move a single edge: what came back has to be exactly what the file said.
            byte[] again = WriteToBuffer(loaded);
            XElement[] before = ReadTables(written, "ppt/slides/slide1.xml");
            XElement[] after = ReadTables(again, "ppt/slides/slide1.xml");
            Require(before.Length == after.Length, "the re-save changed the table count from "
                + before.Length + " to " + after.Length);
            string firstEdges = EdgeSignatures(before);
            string secondEdges = EdgeSignatures(after);
            Require(firstEdges == secondEdges, "per-cell edges changed on re-save:\n  first " + firstEdges
                + "\n  again " + secondEdges);
        }

        /// <summary>边框线的颜色十六进制，线不存在就说是 null。 A border line's colour hex, or null.</summary>
        /// <param name="line">边框线。 The line.</param>
        /// <returns>十六进制或占位。 The hex text or a placeholder.</returns>
        private static string HexOf(BorderLine line)
        {
            return line == null ? "<null>" : line.Color.ToHex();
        }

        /// <summary>把所有表格每格四条边线拼成可比较的签名串。 Builds a comparable signature of every cell edge of every table.</summary>
        /// <param name="tables">a:tbl 元素列表。 The a:tbl elements.</param>
        /// <returns>签名串。 The signature text.</returns>
        private static string EdgeSignatures(XElement[] tables)
        {
            StringBuilder text = new StringBuilder();
            for (int t = 0; t < tables.Length; t++)
            {
                XElement[] trs = tables[t].Elements(A + "tr").ToArray();
                for (int r = 0; r < trs.Length; r++)
                {
                    XElement[] tcs = trs[r].Elements(A + "tc").ToArray();
                    for (int c = 0; c < tcs.Length; c++)
                    {
                        XElement tcPr = tcs[c].Element(A + "tcPr");
                        foreach (string tag in new string[] { "lnL", "lnR", "lnT", "lnB" })
                        {
                            text.Append(t).Append('/').Append(r).Append('/').Append(c).Append('/').Append(tag)
                                .Append('=').Append(EdgeHex(tcPr, tag)).Append(',').Append(EdgeWidth(tcPr, tag))
                                .Append(';');
                        }
                    }
                }
            }
            return text.ToString();
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

        /// <summary>
        /// PT8：章节过渡页只有一个标题占位符，与「仅标题页」形状完全一样，形状推断永远认不出来；
        /// 版式部件自己声明的 <c>p:sldLayout@type="secHead"</c> 是唯一能把它带回来的证据。
        /// 同时守住反方向：手工加表格的那一页仍按形状推断成 TitleAndTable（既有行为不变），
        /// 说明这次只是让版式声明为章节页这一条生效，没有整体压过推断。
        /// PT8: a section header carries a single title placeholder and is shape-for-shape identical to a
        /// title-only slide, so inference can never recognise it; the layout part's own
        /// <c>p:sldLayout@type="secHead"</c> is the only evidence that can bring it back. The opposite
        /// direction is guarded too: a slide with a hand-added table still infers TitleAndTable (existing
        /// behaviour unchanged), proving the declared kind is consulted for the section header only and does
        /// not overrule inference generally.
        /// </summary>
        private static void RunSectionHeaderLayoutSurvivesTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();

            Slide section = deck.AddSlide(SlideLayoutKind.SectionHeader);
            ((TextBoxShape)section.Shapes[0]).Paragraphs[0].AddText("第二部分");

            Slide titleOnly = deck.AddSlide(SlideLayoutKind.Title);
            ((TextBoxShape)titleOnly.Shapes[0]).Paragraphs[0].AddText("仅标题");

            Slide mixed = deck.AddSlide(SlideLayoutKind.TitleAndContent);
            ((TextBoxShape)mixed.Shapes[0]).Paragraphs[0].AddText("混合");
            mixed.AddTable(60, 200, 600, 100).Table.AddRow("A", "B");

            byte[] written = WriteToBuffer(deck);

            // 版式部件必须自己说清楚：有一个 type="secHead"，名字仍叫「节标题」。
            // The layout part has to say so itself: one carries type="secHead" and keeps the 节标题 name.
            int sectionLayouts = 0;
            using (ZipArchive zip = new ZipArchive(new MemoryStream(written), ZipArchiveMode.Read))
            {
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (!entry.FullName.StartsWith("ppt/slideLayouts/slideLayout", StringComparison.Ordinal)
                        || !entry.FullName.EndsWith(".xml", StringComparison.Ordinal)) continue;
                    using (Stream stream = entry.Open())
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        XElement root = XDocument.Parse(reader.ReadToEnd()).Root;
                        if (root == null) continue;
                        if ((string)root.Attribute("type") != "secHead") continue;
                        sectionLayouts++;
                        XElement cSld = root.Element(P + "cSld");
                        Require(cSld != null && (string)cSld.Attribute("name") == "节标题",
                            "the section layout lost its name: "
                            + (cSld == null ? "<no cSld>" : (string)cSld.Attribute("name")));
                    }
                }
            }
            Require(sectionLayouts == 1, "expected exactly one secHead layout part, got " + sectionLayouts);

            OfficePresentation loaded = PptxReader.Load(written);
            Require(loaded.Slides[0].Layout == SlideLayoutKind.SectionHeader,
                "the section header slide came back as " + loaded.Slides[0].Layout);
            Require(loaded.Slides[1].Layout == SlideLayoutKind.Title,
                "the title-only slide came back as " + loaded.Slides[1].Layout);
            Require(loaded.Slides[2].Layout == SlideLayoutKind.TitleAndTable,
                "a hand-added table no longer infers TitleAndTable: " + loaded.Slides[2].Layout);
        }

        /// <summary>
        /// PT10：`docProps/app.xml` 的元素必须按 ISO/IEC 29500 的 `CT_Properties` 序列排（Company → Slides →
        /// Application → AppVersion），而不是 PowerPoint 实测写出的「Application 提前」次序；后者不满足架构
        /// 定义的 sequence，严格校验会报错，而这个部件只承载统计信息，按架构写不掉任何能力。
        /// PT10: the elements of `docProps/app.xml` have to follow the `CT_Properties` sequence in ISO/IEC
        /// 29500 (Company → Slides → Application → AppVersion) rather than the "Application first" order
        /// PowerPoint actually emits; the latter violates the schema's sequence and fails strict validation,
        /// and this part only carries statistics, so following the schema costs no capability.
        /// </summary>
        private static void RunAppPropertiesSchemaOrderTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            deck.AddSlide(SlideLayoutKind.Blank);
            byte[] written = WriteToBuffer(deck);

            XElement root = ReadSlideXml(written, "docProps/app.xml").Root;
            Require(root != null, "docProps/app.xml is missing");
            List<string> order = new List<string>();
            foreach (XElement child in root.Elements()) order.Add(child.Name.LocalName);

            string[] expected = { "Company", "Slides", "Application", "AppVersion" };
            Require(order.Count == expected.Length,
                "unexpected element count in docProps/app.xml: " + string.Join(", ", order.ToArray()));
            for (int i = 0; i < expected.Length; i++)
            {
                Require(order[i] == expected[i],
                    "docProps/app.xml is not in CT_Properties order (position " + i + "): "
                    + string.Join(", ", order.ToArray()));
            }
        }

        /// <summary>
        /// PT2：`p:timing` 必须长成 <c>tnLst/par/cTn(nodeType="tmRoot")/childTnLst/seq</c> 的形状，
        /// id 全局唯一，`dur`/`delay` 用规范词法 `indefinite`，类别用 `nodeType` 而不是 `type`，
        /// `p:set` 内容按 <c>cBhvr + p:to/p:strVal</c> 给出。旧实现是一个 `p:par` 里并列两个 `p:cTn`、
        /// `seq` 与根 `cTn` 平级、两个节点同用 `id="1"`、并把 `a:clrMapOvr` 塞进 `p:setTo`——PowerPoint 打开即走修复。
        /// PT2: `p:timing` has to take the <c>tnLst/par/cTn(nodeType="tmRoot")/childTnLst/seq</c> shape, with
        /// tree-wide unique ids, the schema spellings `indefinite` for dur/delay, `nodeType` instead of `type`,
        /// and a `p:set` built from <c>cBhvr + p:to/p:strVal</c>. The old tree put two `p:cTn` inside one
        /// `p:par`, kept `seq` beside the root `cTn`, gave two nodes `id="1"`, and tucked an `a:clrMapOvr`
        /// into `p:setTo` — any of which sends PowerPoint into its repair path.
        /// </summary>
        private static void RunTimingTreeShapeTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            TextBoxShape box = slide.AddTextBox(40, 40, 400, 80, "淡入");
            box.Animation = SlideAnimationKind.Fade;
            box.AnimationDurationSeconds = 0.5;
            box.AnimationDelaySeconds = 0.25;

            byte[] written = WriteToBuffer(deck);
            XElement timing = ReadSlideXml(written, "ppt/slides/slide1.xml").Root.Element(P + "timing");
            Require(timing != null, "the animated slide carries no p:timing");
            Require(timing.Elements().Count() == 1 && timing.Elements().First().Name == P + "tnLst",
                "p:timing must hold exactly one p:tnLst");

            XElement rootCtn = timing.Element(P + "tnLst").Element(P + "par");
            rootCtn = rootCtn == null ? null : rootCtn.Element(P + "cTn");
            Require(rootCtn != null, "p:tnLst must hold a par wrapping one cTn");
            Require((string)rootCtn.Attribute("nodeType") == "tmRoot",
                "the root cTn must carry nodeType=\"tmRoot\": " + (string)rootCtn.Attribute("nodeType"));
            Require((string)rootCtn.Attribute("dur") == "indefinite",
                "the root dur must be indefinite: " + (string)rootCtn.Attribute("dur"));
            Require(rootCtn.Elements(P + "cTn").Count() == 0,
                "two cTn siblings inside one par is exactly what CT_TLTimeNodeParallel forbids");

            XElement seq = rootCtn.Element(P + "childTnLst") == null
                ? null : rootCtn.Element(P + "childTnLst").Element(P + "seq");
            Require(seq != null, "the sequence must live inside the root cTn's childTnLst");
            XElement seqCtn = seq.Element(P + "cTn");
            Require(seqCtn != null && (string)seqCtn.Attribute("nodeType") == "mainSeq",
                "the sequence cTn must carry nodeType=\"mainSeq\"");

            // 全树翻一遍：不许出现 indef；不许把类别写进 type；每个 cTn 的 id 必须唯一。
            // Sweep the whole tree: no "indef", no kind smuggled into type, and every cTn id unique.
            HashSet<string> ids = new HashSet<string>();
            foreach (XElement element in timing.DescendantsAndSelf())
            {
                foreach (XAttribute attribute in element.Attributes())
                {
                    Require(attribute.Value != "indef",
                        "the schema spelling is 'indefinite', not 'indef' (on " + attribute.Name + ")");
                }
                if (element.Name != P + "cTn") continue;
                string id = (string)element.Attribute("id");
                Require(!string.IsNullOrEmpty(id) && ids.Add(id),
                    "duplicate or missing cTn id: " + id);
            }

            XElement effect = null;
            foreach (XElement candidate in timing.Descendants(P + "cTn"))
            {
                if ((string)candidate.Attribute("presetClass") == "entr") effect = candidate;
            }
            Require(effect != null, "no entrance behaviour was written");
            Require((string)effect.Attribute("nodeType") == "clickEffect",
                "the effect node type is not written to nodeType: " + (string)effect.Attribute("nodeType"));

            XElement set = effect.Element(P + "childTnLst") == null ? null : effect.Element(P + "childTnLst").Element(P + "set");
            Require(set != null, "the entrance behaviour lost its visibility p:set");
            Require(set.Descendants(A + "clrMapOvr").Count() == 0,
                "a:clrMapOvr inside p:setTo is not part of the p:set content model");
            XElement toValue = set.Element(P + "to") == null ? null : set.Element(P + "to").Element(P + "strVal");
            Require(toValue != null && (string)toValue.Attribute("val") == "visible",
                "p:set must end with p:to/p:strVal val=\"visible\"");
            Require(set.Descendants(P + "attrName").Count() == 1,
                "p:set must declare exactly one p:attrName");

            // 往返：新形状仍要能读回类别、时长与延迟。
            // Round trip: the new shape still reads back its kind, duration and delay.
            OfficePresentation loaded = PptxReader.Load(written);
            SlideShape back = loaded.Slides[0].Shapes[0];
            Require(back.Animation == SlideAnimationKind.Fade,
                "the animation kind did not read back: " + back.Animation);
            Require(Math.Abs(back.AnimationDurationSeconds - 0.5) < 0.001,
                "the animation duration did not read back: " + back.AnimationDurationSeconds);
            Require(Math.Abs(back.AnimationDelaySeconds - 0.25) < 0.001,
                "the animation delay did not read back: " + back.AnimationDelaySeconds);
        }

        /// <summary>
        /// PT2 的连带口径：每一形状自己的 <c>AnimationDelaySeconds</c> 都要写进它那个延迟节点的数值
        /// <c>delay</c>，读回也要逐形状对上；旧写法只让第一个（自动放映那条）带数值，其余一律
        /// <c>indefinite</c>，回读成 0。触发节点仍是单击语义，不许把 <c>indefinite</c> 当延迟读出来。
        /// The follow-on rule from PT2: every shape's own <c>AnimationDelaySeconds</c> must land as a numeric
        /// <c>delay</c> on that shape's delay node and read back per shape; the old writer gave only the first
        /// (auto-advance) shape a number and left the rest <c>indefinite</c>, which read back as 0. Trigger
        /// nodes keep their click semantics and their <c>indefinite</c> must never be read as a delay.
        /// </summary>
        private static void RunNonFirstShapeAnimationDelayTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            double[] delays = new double[] { 1.2, 0.8, 0.4, 0.0 };
            string[] expected = new string[] { "1200", "800", "400", "0" };
            for (int i = 0; i < delays.Length; i++)
            {
                TextBoxShape box = slide.AddTextBox(40, 40 + i * 90, 320, 70, "形状 " + (i + 1).ToString());
                box.Animation = SlideAnimationKind.Fade;
                box.AnimationDelaySeconds = delays[i];
            }

            byte[] written = WriteToBuffer(deck);
            XElement timing = ReadSlideXml(written, "ppt/slides/slide1.xml").Root.Element(P + "timing");
            Require(timing != null, "the animated slide carries no p:timing");

            List<string> values = timing.Descendants(P + "cond")
                .Select(c => (string)c.Attribute("delay"))
                .Where(v => v != null)
                .ToList();
            for (int i = 0; i < expected.Length; i++)
            {
                if (expected[i] == "0") continue;
                Require(values.Contains(expected[i]),
                    "the delay of " + expected[i] + " ms never reached the file, wrote: "
                    + string.Join(", ", values.ToArray()));
            }

            OfficePresentation loaded = PptxReader.Load(written);
            Require(loaded.Slides.Count == 1 && loaded.Slides[0].Shapes.Count == 4,
                "the round trip lost shapes: " + loaded.Slides[0].Shapes.Count);
            for (int i = 0; i < delays.Length; i++)
            {
                SlideShape back = loaded.Slides[0].Shapes[i];
                Require(Math.Abs(back.AnimationDelaySeconds - delays[i]) < 0.001,
                    "shape " + i + " delay did not read back: expected " + delays[i]
                    + ", got " + back.AnimationDelaySeconds);
            }
        }

        /// <summary>
        /// PT3 读端补齐（上下标与段落间距）：`Superscript`/`Subscript` 读端一直按 `a:rPr/@baseline` 还原，
        /// 写端却从不输出；`lnSpc`/`spcBef`/`spcAft` 反过来——写端一直输出、读端从不读。两个方向都补齐后，
        /// 「存进去再读回来」才真的等于存进去的东西。
        /// PT3 read-side parity (super/subscript and paragraph spacing): the reader has always restored
        /// `Superscript`/`Subscript` from `a:rPr/@baseline` while the writer never emitted it, and
        /// `lnSpc`/`spcBef`/`spcAft` went the other way — always written, never read. With both directions
        /// filled in, "save then load" finally equals what was saved.
        /// </summary>
        private static void RunBaselineAndSpacingRoundTripTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            TextBoxShape box = slide.AddTextBox(40, 40, 520, 140, "基线");
            Paragraph paragraph = box.Paragraphs[0];
            paragraph.Inlines.Clear();

            TextRun superscript = new TextRun("上标");
            superscript.Format.Superscript = true;
            TextRun subscript = new TextRun("下标");
            subscript.Format.Subscript = true;
            paragraph.Inlines.Add(superscript);
            paragraph.Inlines.Add(subscript);
            paragraph.LineSpacing = 1.5;
            paragraph.SpaceBefore = 6.0;
            paragraph.SpaceAfter = 9.0;

            byte[] written = WriteToBuffer(deck);
            OfficePresentation loaded = PptxReader.Load(written);
            Paragraph back = ((TextBoxShape)loaded.Slides[0].Shapes[0]).Paragraphs[0];

            Require(back.Inlines.Count == 2, "the two runs did not read back: " + back.Inlines.Count);
            Require(((TextRun)back.Inlines[0]).Format.Superscript,
                "superscript was lost: the writer never emits a:rPr/@baseline");
            Require(((TextRun)back.Inlines[1]).Format.Subscript,
                "subscript was lost: the writer never emits a:rPr/@baseline");
            Require(Math.Abs(back.LineSpacing - 1.5) < 0.0001,
                "line spacing did not read back: " + back.LineSpacing);
            Require(Math.Abs(back.SpaceBefore - 6.0) < 0.0001,
                "space before did not read back: " + back.SpaceBefore);
            Require(Math.Abs(back.SpaceAfter - 9.0) < 0.0001,
                "space after did not read back: " + back.SpaceAfter);
        }

        /// <summary>
        /// PT3 读端补齐（纵向合并）：写端过去只写 <c>gridSpan</c>，`RowSpan` 被完全丢掉，续接格也不占位，
        /// 于是第二行的真实单元格顶到第一列——数据整体左移一列。现在首格写 <c>rowSpan</c>，后续行按网格
        /// 补 <c>vMerge="1"</c>（块首列）与 <c>hMerge="1"</c>（同块其余列）。
        /// PT3 read-side parity (vertical merge): the writer used to emit only <c>gridSpan</c>, dropping
        /// `RowSpan` entirely and leaving no continuation cells, so the second row's real cells landed on the
        /// first column — the data shifted one column left. The origin cell now carries <c>rowSpan</c> and the
        /// following rows get <c>vMerge="1"</c> (block start) and <c>hMerge="1"</c> (the rest of the block).
        /// </summary>
        private static void RunVerticalMergeGridTest()
        {
            OfficePresentation deck = OfficePresentation.Widescreen();
            Slide slide = deck.AddSlide(SlideLayoutKind.Blank);
            TableShape shape = slide.AddTable(40, 40, 600, 200);
            Table table = shape.Table;

            TableRow header = table.AddRow("区域", "收入");
            header.IsHeader = true;
            TableRow first = table.AddRow("华东");
            first.Cells[0].RowSpan = 2;
            first.Cells.Add(new TableCell());
            first.Cells[1].Blocks.Add(new Paragraph("588"));
            TableRow second = table.AddRow();                       // 只有一格：另一格被上面的纵向合并占住
            second.Cells.Add(new TableCell());
            second.Cells[0].Blocks.Add(new Paragraph("384"));

            byte[] written = WriteToBuffer(deck);
            XElement[] tables = ReadTables(written, "ppt/slides/slide1.xml");
            Require(tables.Length == 1, "expected one table, got " + tables.Length);
            XElement[] trs = tables[0].Elements(A + "tr").ToArray();
            Require(trs.Length == 3, "expected three rows, got " + trs.Length);

            XElement origin = FirstCell(trs[1]);
            Require(origin != null && (string)origin.Attribute("rowSpan") == "2",
                "the origin cell did not carry rowSpan=\"2\"");
            Require((string)origin.Attribute("gridSpan") == null,
                "a one-column cell must not carry gridSpan");

            // 第二数据行：第一列是被占住的续接格，真实内容落在第二列。
            // The second data row: the first column is the continuation cell and the real content sits in
            // the second column.
            XElement[] secondRowCells = trs[2].Elements(A + "tc").ToArray();
            Require(secondRowCells.Length == 2,
                "the continuation row should still carry two cells, got " + secondRowCells.Length);
            Require((string)secondRowCells[0].Attribute("vMerge") == "1",
                "the continuation cell is missing vMerge=\"1\"");
            string[] texts = new string[secondRowCells.Length];
            for (int i = 0; i < secondRowCells.Length; i++)
            {
                texts[i] = secondRowCells[i].Value;
            }
            Require(texts[1].IndexOf("384", StringComparison.Ordinal) >= 0,
                "the second row's data slid into the merged column: " + string.Join(" | ", texts));
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

        /// <summary>取幻灯片部件里的每个 <c>a:p</c>，按文档顺序返回。 Collects every <c>a:p</c> in a slide part, in document order.</summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径。The part path.</param>
        /// <returns>段落元素列表。The paragraph elements.</returns>
        private static XElement[] ReadParagraphElements(byte[] archive, string entryName)
        {
            List<XElement> found = new List<XElement>();
            foreach (XElement p in ReadSlideXml(archive, entryName).Descendants(A + "p")) found.Add(p);
            return found.ToArray();
        }

        /// <summary>取第一页第一个文本框首段的项目符号标记。 Reads the list marker of the first paragraph of the first text box on slide one.</summary>
        /// <param name="presentation">回读的演示文稿。The reloaded presentation.</param>
        /// <returns>标记文本，没有则 null。The marker text, or null when absent.</returns>
        private static string ReadFirstMarker(OfficePresentation presentation)
        {
            foreach (SlideShape shape in presentation.Slides[0].Shapes)
            {
                TextBoxShape box = shape as TextBoxShape;
                if (box == null) continue;
                if (box.Paragraphs.Count == 0) continue;
                return box.Paragraphs[0].ListMarker;
            }
            return null;
        }

        /// <summary>取幻灯片部件里每个 <c>p:sp</c> 的 <c>p:spPr</c>，按文档顺序返回。 Collects the <c>p:spPr</c> of every <c>p:sp</c> in a slide part, in document order.</summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径。The part path.</param>
        /// <returns>形状属性元素列表。The shape-property elements.</returns>
        private static XElement[] ReadShapeProperties(byte[] archive, string entryName)
        {
            List<XElement> found = new List<XElement>();
            foreach (XElement sp in ReadSlideXml(archive, entryName).Descendants(P + "sp"))
            {
                XElement spPr = sp.Element(P + "spPr");
                if (spPr != null) found.Add(spPr);
            }
            return found.ToArray();
        }

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

        /// <summary>按原文读出幻灯片部件的文本，用于构造夹具。 Reads a slide part verbatim so a fixture can edit it as text.</summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径。The part path.</param>
        /// <returns>部件文本。The part text.</returns>
        private static string SlideEntryText(byte[] archive, string entryName)
        {
            using (ZipArchive zip = new ZipArchive(new MemoryStream(archive), ZipArchiveMode.Read))
            {
                ZipArchiveEntry entry = zip.GetEntry(entryName);
                Require(entry != null, "slide part missing: " + entryName);
                using (Stream stream = entry.Open())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>在内存里重打包，替换指定条目的文本，其余条目原样复制。 Repackages in memory, replacing one entry's text and copying the rest as-is.</summary>
        /// <param name="package">原包字节。The original package bytes.</param>
        /// <param name="entryName">要替换的条目名。The entry to replace.</param>
        /// <param name="content">新的条目文本。The new entry text.</param>
        /// <returns>重打包后的字节。The repackaged bytes.</returns>
        private static byte[] RewriteZipEntryText(byte[] package, string entryName, string content)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(content);
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
                                output.Write(bytes, 0, bytes.Length);
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

        /// <summary>取幻灯片部件里每个 <c>a:bodyPr</c>，按文档顺序返回。 Collects every <c>a:bodyPr</c> in a slide part, in document order.</summary>
        /// <param name="archive">pptx 字节。The pptx bytes.</param>
        /// <param name="entryName">部件路径。The part path.</param>
        /// <returns>bodyPr 元素列表。The bodyPr elements.</returns>
        private static XElement[] ReadBodyPrs(byte[] archive, string entryName)
        {
            List<XElement> found = new List<XElement>();
            foreach (XElement bodyPr in ReadSlideXml(archive, entryName).Descendants(A + "bodyPr")) found.Add(bodyPr);
            return found.ToArray();
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
