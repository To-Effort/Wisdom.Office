using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Wisdom.Office.Core;
using Wisdom.Ppt;

// PPT 侧复杂效果覆盖面探测 + 写出/读取性能基准。
// Probes the complex-effect coverage of Wisdom.Ppt and benchmarks write/read throughput.
internal static class PptMicro
{
    private static readonly SlideTransitionKind[] Transitions =
    {
        SlideTransitionKind.None, SlideTransitionKind.Fade, SlideTransitionKind.Push,
        SlideTransitionKind.Wipe, SlideTransitionKind.Split
    };

    private static readonly AutoShapeKind[] ShapeKinds =
    {
        AutoShapeKind.Rectangle, AutoShapeKind.RoundedRectangle, AutoShapeKind.Ellipse,
        AutoShapeKind.Triangle, AutoShapeKind.Diamond, AutoShapeKind.Pentagon,
        AutoShapeKind.Hexagon, AutoShapeKind.RightArrow, AutoShapeKind.Chevron, AutoShapeKind.Line
    };

    private static readonly SlideLayoutKind[] LayoutCycle =
    {
        SlideLayoutKind.Title, SlideLayoutKind.TitleAndContent, SlideLayoutKind.TwoContent,
        SlideLayoutKind.SectionHeader, SlideLayoutKind.TitleAndTable, SlideLayoutKind.TitleAndPicture,
        SlideLayoutKind.Blank
    };

    private static OfficeImage MakeImage(int seed)
    {
        // 1x1 PNG，尾部附加一个可区分实例的注释字节，用于测试媒体去重与多图写出。
        // 1x1 PNG with per-instance trailing bytes to exercise media de-duplication and multi-image writes.
        byte[] base64 = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");
        byte[] data = new byte[base64.Length + 4];
        Array.Copy(base64, data, base64.Length);
        data[base64.Length] = (byte)seed;
        data[base64.Length + 1] = (byte)(seed * 7);
        return new OfficeImage(data);
    }

    private static TextBoxShape FindContent(Slide slide)
    {
        foreach (SlideShape shape in slide.Shapes)
        {
            TextBoxShape box = shape as TextBoxShape;
            if (box != null && (box.Placeholder == PlaceholderKind.Content
                || box.Placeholder == PlaceholderKind.Body))
            {
                return box;
            }
        }
        return slide.AddTextBox(96, 170, 768, 260, string.Empty);
    }

    /// <summary>构建一叠尽量吃满全部可表达效果的复杂幻灯片。Builds a deck that saturates every expressible effect.</summary>
    private static OfficePresentation BuildDeck(int slideCount, OfficeImage[] media, string tag)
    {
        OfficePresentation presentation = OfficePresentation.Widescreen();
        presentation.Metadata.Title = "复杂效果压测 " + tag;
        presentation.Metadata.Author = "PptMicro";

        for (int i = 0; i < slideCount; i++)
        {
            Slide slide = presentation.AddSlide(LayoutCycle[i % LayoutCycle.Length]);
            slide.Transition = Transitions[i % Transitions.Length];
            slide.TransitionDuration = 0.3 + (i % Transitions.Length) * 0.4;
            slide.NotesText = "备注 " + i.ToString(CultureInfo.InvariantCulture) + "：本页覆盖切换、备注与富文本。";
            if (i % 3 == 0)
            {
                slide.BackgroundColor = new OfficeColor((byte)(0xF0 + i % 8), 0xF7, 0xFF);
            }

            TextBoxShape title = null;
            foreach (SlideShape shape in slide.Shapes)
            {
                TextBoxShape box = shape as TextBoxShape;
                if (box != null && box.Placeholder == PlaceholderKind.Title) { title = box; break; }
            }
            if (title == null)
            {
                title = slide.AddTextBox(60, 30, 840, 90, string.Empty);
                title.Placeholder = PlaceholderKind.Title;
            }
            title.Paragraphs[0].AddText("第 " + (i + 1).ToString(CultureInfo.InvariantCulture) + " 页 · 复杂效果");
            title.Paragraphs[0].Alignment = ParagraphAlignment.Left;

            TextBoxShape body = FindContent(slide);
            body.Paragraphs[0].Inlines.Clear();
            TextFormat strong = new TextFormat("微软雅黑", 22.0) { Bold = true, Color = new OfficeColor(0x1F, 0x3A, 0x5F) };
            TextFormat emphasis = new TextFormat("Georgia", 18.0) { Italic = true, Underline = true };
            TextFormat plain = new TextFormat("Calibri", 18.0);
            body.Paragraphs[0].AddText("收入 1,280 万元", strong);
            body.Paragraphs[0].AddText("（环比 +12%）", emphasis);
            body.Paragraphs[0].ListMarker = "•";
            body.Paragraphs[0].LeftIndent = 18;
            for (int b = 1; b <= 5; b++)
            {
                Paragraph p = body.AddParagraph(string.Empty);
                p.AddText("要点 " + b.ToString(CultureInfo.InvariantCulture) + "：", plain);
                p.AddText("毛利率 41.6%，成本控制在预算内。", b % 2 == 0 ? emphasis : strong);
                p.ListMarker = "•";
                p.LeftIndent = 18;
            }
            body.ShrinkToFit = true;
            body.VerticalAlignment = SlideTextAnchor.Top;

            for (int s = 0; s < 10; s++)
            {
                AutoShape shape = slide.AddShape(ShapeKinds[s], 60 + s * 86, 400, 76, 56,
                    "形状" + s.ToString(CultureInfo.InvariantCulture));
                shape.FillColor = presentation.AccentColors[s % presentation.AccentColors.Count];
                shape.Outline = new BorderLine(BorderStyle.Single, 1.5, new OfficeColor(0x1F, 0x3A, 0x5F));
                shape.CornerRadius = 0.25;
                shape.Paragraphs[0].Alignment = ParagraphAlignment.Center;
            }

            PictureShape picture = slide.AddImage(media[i % media.Length], 700, 300, 120, 90);
            picture.AlternativeText = "替代文字 " + i.ToString(CultureInfo.InvariantCulture);

            TableShape table = slide.AddTable(60, 470, 600, 60);
            table.HeaderColor = new OfficeColor(0x1F, 0x3A, 0x5F);
            table.Table.AddRow("区域", "收入", "占比", "环比");
            for (int r = 0; r < 5; r++)
            {
                table.Table.AddRow("区域" + r, (588 + r).ToString(CultureInfo.InvariantCulture),
                    (40 + r) + "%", "+" + r + "%");
            }
        }
        return presentation;
    }

    /// <summary>效果矩阵探测：写出后逐项读回比对，并列出模型无法表达的复杂效果。
    /// Effect-matrix probe: writes a saturating deck, reads it back item by item, and lists effects the model cannot express.</summary>
    private static void EffectsProbe()
    {
        Console.WriteLine("=== 复杂效果支持矩阵（写出→读回） Complex-effect matrix (write → read-back) ===");
        OfficeImage[] media = { MakeImage(1), MakeImage(2), MakeImage(3) };
        OfficePresentation deck = BuildDeck(7, media, "probe");
        string path = Path.Combine(AppContext.BaseDirectory, "probe-complex.pptx");
        byte[] bytes = PptxWriter.Write(deck);
        File.WriteAllBytes(path, bytes);
        OfficePresentation back = PptxReader.Load(bytes);

        Check("页数", deck.Slides.Count, back.Slides.Count);
        Slide src = deck.Slides[1];
        Slide dst = back.Slides[1];

        TextBoxShape srcTitle = (TextBoxShape)src.Shapes[0];
        TextBoxShape dstTitle = FindByKind(dst, PlaceholderKind.Title);
        Check("标题文本", srcTitle.GetText(), dstTitle != null ? dstTitle.GetText() : "<null>");

        TextBoxShape srcBody = FindContent(src);
        TextBoxShape dstBody = FindByKind(dst, PlaceholderKind.Content);
        Check("正文多运行文本", srcBody.GetText(), dstBody != null ? dstBody.GetText() : "<null>");
        Check("项目符号", srcBody.Paragraphs[0].ListMarker, dstBody != null ? dstBody.Paragraphs[0].ListMarker : "<null>");
        Check("左缩进", srcBody.Paragraphs[0].LeftIndent, dstBody != null ? dstBody.Paragraphs[0].LeftIndent : -1);

        TextRun srcRun = (TextRun)srcBody.Paragraphs[0].Inlines[0];
        TextRun dstRun = dstBody != null && dstBody.Paragraphs[0].Inlines.Count > 0
            ? dstBody.Paragraphs[0].Inlines[0] as TextRun : null;
        Check("粗体", srcRun.Format.Bold, dstRun != null && dstRun.Format.Bold);
        Check("字号", srcRun.Format.FontSize, dstRun != null ? dstRun.Format.FontSize : 0);
        Check("字体", srcRun.Format.FontFamily, dstRun != null ? dstRun.Format.FontFamily : "<null>");
        Check("颜色", srcRun.Format.Color.ToHex(), dstRun != null ? dstRun.Format.Color.ToHex() : "<null>");
        Check("下划线", srcBody.Paragraphs[1].Inlines.Count > 1
            ? ((TextRun)srcBody.Paragraphs[1].Inlines[1]).Format.Underline : false,
            dstBody != null && dstBody.Paragraphs.Count > 1 && dstBody.Paragraphs[1].Inlines.Count > 1
            ? ((TextRun)dstBody.Paragraphs[1].Inlines[1]).Format.Underline : false);

        AutoShape srcShape = null;
        foreach (SlideShape shape in src.Shapes)
        {
            AutoShape a = shape as AutoShape;
            if (a != null && a.Kind == AutoShapeKind.RoundedRectangle) { srcShape = a; break; }
        }
        AutoShape dstShape = null;
        foreach (SlideShape shape in dst.Shapes)
        {
            AutoShape a = shape as AutoShape;
            if (a != null && a.Kind == AutoShapeKind.RoundedRectangle) { dstShape = a; break; }
        }
        Check("自选图形几何 Left", srcShape.Left, dstShape != null ? dstShape.Left : -1);
        Check("填充色", srcShape.FillColor.Value.ToHex(), dstShape != null && dstShape.FillColor.HasValue ? dstShape.FillColor.Value.ToHex() : "<null>");
        Check("描边", srcShape.Outline != null, dstShape != null && dstShape.Outline != null);
        Check("圆角半径", srcShape.CornerRadius, dstShape != null ? dstShape.CornerRadius : -1);

        PictureShape srcPic = (PictureShape)src.Shapes.Find(shape => shape is PictureShape);
        PictureShape dstPic = (PictureShape)dst.Shapes.Find(shape => shape is PictureShape);
        Check("图片字节往返", srcPic.Image.Data.Length, dstPic != null ? dstPic.Image.Data.Length : -1);
        Check("替代文字", srcPic.AlternativeText, dstPic != null ? dstPic.AlternativeText : "<null>");

        TableShape srcTable = (TableShape)src.Shapes.Find(shape => shape is TableShape);
        TableShape dstTable = (TableShape)dst.Shapes.Find(shape => shape is TableShape);
        Check("表格行数", srcTable.Table.Rows.Count, dstTable != null ? dstTable.Table.Rows.Count : -1);
        Check("表格单元文本", CellText(srcTable.Table.Rows[1].Cells[0]),
            dstTable != null ? CellText(dstTable.Table.Rows[1].Cells[0]) : "<null>");

        Slide srcFade = deck.Slides[1];
        Slide dstFade = back.Slides[1];
        Check("切换类型", srcFade.Transition.ToString(), dstFade.Transition.ToString());
        Check("切换时长", srcFade.TransitionDuration, dstFade.TransitionDuration);
        Slide srcBg = deck.Slides[0];
        Slide dstBg = back.Slides[0];
        Check("背景色读回", srcBg.BackgroundColor.HasValue ? srcBg.BackgroundColor.Value.ToHex() : "<null>",
            dstBg.BackgroundColor.HasValue ? dstBg.BackgroundColor.Value.ToHex() : "<null>");
        Check("备注", src.NotesText, dst.NotesText);
        Check("ShrinkToFit", srcBody.ShrinkToFit, dstBody != null && dstBody.ShrinkToFit);

        Console.WriteLine();
        Console.WriteLine("模型未提供、因此无法写出也无法读回的复杂效果（编译期即缺失）：");
        Console.WriteLine("  - 渐变填充（线性/径向/路径）gradient fill");
        Console.WriteLine("  - 形状阴影 / 辉光 / 柔化边缘 / 三维旋转 shadow, glow, soft edges, 3-D");
        Console.WriteLine("  - 动画序列（进入/强调/退出/路径）animation sequences");
        Console.WriteLine("  - 图表 / SmartArt / 嵌入 OLE 对象 charts, SmartArt, embedded OLE");
        Console.WriteLine("  - 连接线与自由曲线（custGeom）connectors and freeform geometry");
        Console.WriteLine("  - 外部模板 .potx 与第三方母版继承 external .potx templates");
        Console.WriteLine("  - 旧版 .ppt（97-2003 二进制）legacy .ppt");
        Console.WriteLine();
    }

    private static string CellText(TableCell cell)
    {
        if (cell.Blocks.Count == 0) return string.Empty;
        Paragraph paragraph = cell.Blocks[0] as Paragraph;
        return paragraph != null ? paragraph.GetText() : "<block>";
    }

    private static TextBoxShape FindByKind(Slide slide, PlaceholderKind kind)
    {
        foreach (SlideShape shape in slide.Shapes)
        {
            TextBoxShape box = shape as TextBoxShape;
            if (box != null && box.Placeholder == kind) return box;
        }
        return null;
    }

    private static void Check(string label, object expected, object actual)
    {
        bool ok = Equals(expected, actual);
        Console.WriteLine((ok ? "  [OK]      " : "  [MISMATCH]") + " " + label +
            "  写出=" + Convert.ToString(expected, CultureInfo.InvariantCulture) +
            "  读回=" + Convert.ToString(actual, CultureInfo.InvariantCulture));
    }

    /// <summary>先写临时文件再原子替换：目标被 PowerPoint 等进程锁住时落到带时间戳的旁路文件并告警，
    /// 保证压测不中断、产物不半写。 Writes to a temp file then replaces the target atomically; when the
    /// target is locked by PowerPoint or another process, falls back to a timestamped side file with a
    /// warning so the benchmark never aborts mid-write.</summary>
    private static void SafeWrite(OfficePresentation deck, string path)
    {
        string temp = path + ".tmp";
        PptxWriter.Write(deck, temp);
        try
        {
            File.Move(temp, path, true);
        }
        catch (IOException)
        {
            string side = path.Substring(0, path.Length - 5) + "-" + DateTime.Now.ToString("HHmmss") + ".pptx";
            File.Move(temp, side, true);
            Console.WriteLine("  [WARN] " + Path.GetFileName(path) + " 被占用，本次产物写到 " + Path.GetFileName(side) + "，请关闭 PowerPoint 中的旧文件");
        }
    }

    private static void Run(string label, int slides, int iterations)
    {
        OfficeImage[] media = { MakeImage(1), MakeImage(2), MakeImage(3), MakeImage(4) };
        string path = Path.Combine(AppContext.BaseDirectory, "micro-" + slides + ".pptx");
        Action warm = () => SafeWrite(BuildDeck(slides, media, "warm"), path);
        warm();

        double[] samples = new double[iterations];
        long allocatedBefore = GC.GetTotalAllocatedBytes(true);
        for (int i = 0; i < iterations; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Stopwatch sw = Stopwatch.StartNew();
            SafeWrite(BuildDeck(slides, media, "t" + i), path);
            sw.Stop();
            samples[i] = sw.Elapsed.TotalMilliseconds;
        }
        long allocated = (GC.GetTotalAllocatedBytes(true) - allocatedBefore) / iterations;
        Array.Sort(samples);
        Console.WriteLine(label +
            " 写出(含建模型) min=" + F(samples[0]) +
            " med=" + F(samples[iterations / 2]) +
            " p95=" + F(samples[Math.Min(iterations - 1, (int)Math.Ceiling(iterations * 0.95) - 1)]) +
            " alloc=" + F(allocated / 1048576.0) + " MB" +
            " file=" + new FileInfo(path).Length);

        OfficePresentation built = BuildDeck(slides, media, "frozen");
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        Stopwatch modelSw = Stopwatch.StartNew();
        OfficePresentation again = BuildDeck(slides, media, "frozen2");
        modelSw.Stop();
        again = null;
        Console.WriteLine("    仅构建模型 time=" + F(modelSw.Elapsed.TotalMilliseconds) + " ms");

        int writeIterations = Math.Max(1, iterations);
        double[] writeSamples = new double[writeIterations];
        long writeBefore = GC.GetTotalAllocatedBytes(true);
        for (int i = 0; i < writeIterations; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Stopwatch sw = Stopwatch.StartNew();
            using (MemoryStream stream = new MemoryStream())
            {
                PptxWriter.Write(built, stream);
            }
            sw.Stop();
            writeSamples[i] = sw.Elapsed.TotalMilliseconds;
        }
        long writeAlloc = (GC.GetTotalAllocatedBytes(true) - writeBefore) / writeIterations;
        Array.Sort(writeSamples);
        Console.WriteLine("    仅写出(内存流) min=" + F(writeSamples[0]) +
            " ms alloc=" + F(writeAlloc / 1048576.0) + " MB");

        byte[] package = PptxWriter.Write(built);
        int readIterations = Math.Max(1, Math.Min(iterations, 3));
        double[] readSamples = new double[readIterations];
        long readBefore = GC.GetTotalAllocatedBytes(true);
        OfficePresentation read = null;
        for (int i = 0; i < readIterations; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Stopwatch sw = Stopwatch.StartNew();
            read = PptxReader.Load(package);
            sw.Stop();
            readSamples[i] = sw.Elapsed.TotalMilliseconds;
        }
        long readAlloc = (GC.GetTotalAllocatedBytes(true) - readBefore) / readIterations;
        Array.Sort(readSamples);
        Console.WriteLine("    读取 min=" + F(readSamples[0]) +
            " med=" + F(readSamples[readIterations / 2]) +
            " ms alloc=" + F(readAlloc / 1048576.0) + " MB" +
            " 页数=" + read.Slides.Count);
        read = null;
    }

    private static string F(double value)
    {
        return value.ToString("F1", CultureInfo.InvariantCulture);
    }

    private static int Main(string[] args)
    {
        EffectsProbe();
        Run("30 页 × 全效果  ", 30, 5);
        Run("120 页 × 全效果 ", 120, 3);
        Run("400 页 × 全效果 ", 400, 1);
        return 0;
    }
}
