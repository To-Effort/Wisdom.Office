using System;
using System.IO;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Ppt;

// PPTX 读写演示：
// 1. 用统一演示文稿模型构建 4 页幻灯片（标题页、要点页、表格页、图形+图片页）
// 2. PptxWriter 写出 .pptx，再用 PptxReader 回读校验页数与标题
// 3. OfficeFile 统一入口按文件头识别 PPTX 并提取纯文本
// 4. 演示文稿转 PDF / Word；Word 文档按一级标题反向拆分为幻灯片
internal static class Program
{
    private static string OutDir
    {
        get
        {
            string dir = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "out");
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    private static void Main()
    {
        Console.WriteLine("=== Wisdom.Sample.Ppt：PPTX 读写演示 ===");

        // 1. 构建演示文稿
        OfficePresentation presentation = OfficePresentation.Widescreen();
        presentation.Metadata.Title = "季度经营回顾";
        presentation.Metadata.Author = "Wisdom.Office";

        Slide cover = presentation.AddSlide(SlideLayoutKind.Title);
        SetText(Title(cover), "2026 年第三季度经营回顾");
        SetText(Content(cover), "财务分析组 · 10 月");
        cover.NotesText = "开场 2 分钟，说明本次回顾的三个议题。";

        Slide bullets = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
        SetText(Title(bullets), "本季度要点");
        Bullet(Content(bullets), "收入 1,280 万元，环比增长 12%");
        Bullet(Content(bullets), "成本控制在预算内，毛利率 41.6%");
        Bullet(Content(bullets), "华东区贡献了 46% 的增量收入");
        bullets.Transition = SlideTransitionKind.Fade;
        bullets.TransitionDuration = 0.7;

        Slide agenda = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
        SetText(Title(agenda), "区域明细");
        TableShape table = agenda.AddTable(96, 170, 768, 180);
        table.HeaderColor = new OfficeColor(0x1F, 0x3A, 0x5F);
        table.Table.AddRow("区域", "收入（万元）", "占比");
        table.Table.AddRow("华东", "588", "46%");
        table.Table.AddRow("华南", "384", "30%");
        table.Table.AddRow("华北", "308", "24%");

        Slide shapes = presentation.AddSlide(SlideLayoutKind.Blank);
        AutoShape banner = shapes.AddShape(AutoShapeKind.RoundedRectangle, 96, 60, 768, 80, "达成率 96%");
        banner.FillColor = new OfficeColor(0x00, 0xB0, 0x50);
        banner.CornerRadius = 0.25;
        banner.Paragraphs[0].Alignment = ParagraphAlignment.Center;
        shapes.AddImage(new OfficeImage(SolidPng()), 96, 180, 120, 120);
        shapes.AddTextBox(240, 205, 500, 60, "绿色圆角矩形为达标线以上的达成率");

        string pptxPath = Path.Combine(OutDir, "sample-deck.pptx");
        PptxWriter.Write(presentation, pptxPath);
        Console.WriteLine("[写出] PPTX: " + new FileInfo(pptxPath).Length + " 字节，页数=" + presentation.Slides.Count);

        // 2. 回读校验
        OfficePresentation loaded = PptxReader.Load(pptxPath);
        Console.WriteLine("[回读] 页数=" + loaded.Slides.Count
            + "，尺寸=" + (int)loaded.SlideWidth + "x" + (int)loaded.SlideHeight);
        for (int i = 0; i < loaded.Slides.Count; i++)
        {
            Slide slide = loaded.Slides[i];
            Console.WriteLine("  第 " + (i + 1) + " 页 形状=" + slide.Shapes.Count
                + " 版式=" + slide.Layout + " 标题=\"" + slide.GetTitle() + "\"");
        }

        // 3. 统一入口
        OfficeFile file = OfficeFile.Load(pptxPath);
        Console.WriteLine("[识别] " + OfficeFileFormats.GetDisplayName(file.Format)
            + "，元数据标题=\"" + file.Metadata.Title + "\"");
        Console.WriteLine("[纯文本] 长度=" + file.ExtractText().Length);

        // 4. 跨格式输出
        string pdfPath = Path.Combine(OutDir, "sample-deck.pdf");
        OfficeConverter.PresentationToPdf(loaded, pdfPath);
        Console.WriteLine("[转换] PDF: " + new FileInfo(pdfPath).Length + " 字节");

        string docxPath = Path.Combine(OutDir, "sample-deck.docx");
        OfficeConverter.PresentationToWord(loaded, docxPath);
        Console.WriteLine("[转换] DOCX: " + new FileInfo(docxPath).Length + " 字节");

        // Word 文档按一级标题反向拆分为幻灯片
        OfficeDocument document = new OfficeDocument();
        document.Metadata.Title = "由文档生成的演示文稿";
        Section section = document.AddSection();
        Heading(section, "第一章 市场概况");
        section.AddParagraph("整体需求平稳，头部客户占比上升。");
        Heading(section, "第二章 产品路线");
        section.AddParagraph("下半年聚焦报表引擎与格式兼容。");
        string fromDocPath = Path.Combine(OutDir, "sample-from-word.pptx");
        OfficeConverter.DocumentToPptx(document, fromDocPath);
        OfficePresentation split = PptxReader.Load(fromDocPath);
        Console.WriteLine("[文档转 PPTX] 页数=" + split.Slides.Count
            + "，标题=\"" + split.Slides[0].GetTitle() + "\"");

        // 5. 入场动画、自动换页、主题深改与 .potx 模板
        OfficePresentation branded = OfficePresentation.Widescreen();
        branded.Theme = new PresentationTheme
        {
            Name = "品牌主题",
            Text = new OfficeColor(0x1A, 0x1A, 0x2E),
            Background = new OfficeColor(0xFA, 0xFA, 0xF5),
            Hyperlink = new OfficeColor(0xC0, 0x39, 0x2B),
            FontMajor = "微软雅黑",
            FontMinor = "Consolas"
        };
        Slide animated = branded.AddSlide(SlideLayoutKind.Blank);
        TextBoxShape headline = animated.AddTextBox(96, 80, 768, 70, "季度收入 1,280 万元");
        headline.Animation = SlideAnimationKind.FlyIn;
        headline.AnimationDurationSeconds = 0.8;
        headline.AnimationDelaySeconds = 0.5;
        TextBoxShape subhead = animated.AddTextBox(96, 160, 768, 50, "环比增长 12%");
        subhead.Animation = SlideAnimationKind.Wipe;
        animated.AdvanceAfterSeconds = 3.0;

        string potxPath = Path.Combine(OutDir, "sample-brand.potx");
        PptxWriter.WriteTemplate(branded, potxPath);
        OfficeFile templateFile = OfficeFile.Load(potxPath);
        OfficePresentation rebranded = PptxReader.Load(potxPath);
        Slide back = rebranded.Slides[0];
        TextBoxShape firstBox = (TextBoxShape)back.Shapes[0];
        Console.WriteLine("[模板] " + OfficeFileFormats.GetDisplayName(templateFile.Format)
            + " " + new FileInfo(potxPath).Length + " 字节，动画=" + firstBox.Animation
            + "/" + firstBox.AnimationDurationSeconds + "s+" + firstBox.AnimationDelaySeconds + "s"
            + "，第二形状=" + ((TextBoxShape)back.Shapes[1]).Animation
            + "，自动换页=" + back.AdvanceAfterSeconds + "s"
            + "，主题=" + rebranded.Theme.Name + "/" + rebranded.Theme.FontMajor
            + "/" + rebranded.Theme.FontMinor
            + "/#" + rebranded.Theme.Text.Value.ToHex());

        Console.WriteLine("PPTX 演示完成，输出目录: " + Path.GetFullPath(OutDir));
    }

    private static void SetText(TextBoxShape box, string text)
    {
        Paragraph paragraph = box.Paragraphs.Count > 0 ? box.Paragraphs[0] : box.AddParagraph(string.Empty);
        paragraph.AddText(text);
    }

    private static void Bullet(TextBoxShape box, string text)
    {
        Paragraph paragraph = box.Paragraphs.Count > 1 || box.Paragraphs[0].Inlines.Count > 0
            ? box.AddParagraph(string.Empty)
            : box.Paragraphs[0];
        paragraph.AddText(text);
        paragraph.ListMarker = "•";
        paragraph.LeftIndent = 18;
    }

    private static void Heading(Section section, string text)
    {
        Paragraph paragraph = section.AddParagraph(text);
        paragraph.HeadingLevel = 1;
    }

    private static TextBoxShape Title(Slide slide)
    {
        return Find(slide, PlaceholderKind.Title);
    }

    private static TextBoxShape Content(Slide slide)
    {
        TextBoxShape box = Find(slide, PlaceholderKind.Content);
        if (box == null) box = Find(slide, PlaceholderKind.Body);
        return box ?? slide.AddTextBox(96, 170, 768, 260, string.Empty);
    }

    private static TextBoxShape Find(Slide slide, PlaceholderKind kind)
    {
        for (int i = 0; i < slide.Shapes.Count; i++)
        {
            TextBoxShape box = slide.Shapes[i] as TextBoxShape;
            if (box != null && box.Placeholder == kind) return box;
        }
        return null;
    }

    private static byte[] SolidPng()
    {
        // 1x1 纯色 PNG，仅用于演示图片部件的嵌入与回读。
        return Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==");
    }
}
