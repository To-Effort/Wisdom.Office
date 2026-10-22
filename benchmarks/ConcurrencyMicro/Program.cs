// 六模块高并发压测：并行写、并行读、CPU 密集转换与门面加载，统计吞吐/延迟分位/GC/错误，并做产物回读校验。
// High-concurrency stress harness for the six modules: parallel write, parallel read, CPU-bound
// conversion and facade loading, reporting throughput, latency percentiles, GC pressure and errors,
// with a read-back integrity check for every write scenario.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Wisdom.Excel;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Ppt;
using Wisdom.Word;

internal static class Program
{
    private static string OutDir;

    private static int Main(string[] args)
    {
        // 每次运行用独立输出目录：上一轮产物若被 Office 打开锁定，重跑不会崩在写文件上。
        // Each run uses a fresh output directory so a previous artifact left open in Office cannot crash the writes.
        string root = Path.Combine(AppContext.BaseDirectory, "cc-out");
        Directory.CreateDirectory(root);
        foreach (string old in Directory.GetDirectories(root))
        {
            try { Directory.Delete(old, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
        }
        OutDir = Path.Combine(root, "run-" + DateTime.Now.ToString("yyyyMMdd-HHmmss"));
        Directory.CreateDirectory(OutDir);
        bool minThreads = args != null && args.Length > 0 && args[0] == "mt";
        if (minThreads) ThreadPool.SetMinThreads(64, 64);
        Console.WriteLine("CPU cores: " + Environment.ProcessorCount
            + "  gcServer=" + (GCSettings.IsServerGC ? "on" : "off")
            + "  minThreads=" + (minThreads ? "64" : "default"));
        int[] dops = new int[] { 1, 4, 8, Math.Max(16, Environment.ProcessorCount) };

        string wordPath = PrepareWord();
        string xlsxPath = PrepareExcel();
        string pptxPath = PreparePpt();
        string pdfPath = PreparePdf(wordPath);
        string mdPath = PrepareMarkdown();
        Console.WriteLine("fixtures ready: word/excel/ppt/pdf/markdown");

        OfficeDocument wordDoc = DocxReader.Load(wordPath);
        OfficeWorkbook workbook = XlsxReader.Load(xlsxPath);
        OfficePresentation presentation = PptxReader.Load(pptxPath);
        string markdown = File.ReadAllText(mdPath);

        List<Result> results = new List<Result>();
        foreach (int dop in dops)
        {
            results.Add(Run("word   write", dop, 4, i => DocxWriter.Write(wordDoc, Path.Combine(OutDir, "w-" + i + ".docx")),
                () => Require(DocxReader.Load(Path.Combine(OutDir, "w-0.docx")).Sections.Count >= 1)));
            results.Add(Run("word   read ", dop, 8, i => { OfficeDocument d = DocxReader.Load(wordPath); Require(d.Sections.Count >= 1); }, null));

            results.Add(Run("excel  write", dop, 4, i => XlsxWriter.Write(workbook, Path.Combine(OutDir, "x-" + i + ".xlsx")),
                () => Require(XlsxReader.Load(Path.Combine(OutDir, "x-0.xlsx")).Worksheets.Count == workbook.Worksheets.Count)));
            results.Add(Run("excel  read ", dop, 8, i => { OfficeWorkbook w = XlsxReader.Load(xlsxPath); Require(w.Worksheets.Count >= 1); }, null));

            results.Add(Run("ppt    write", dop, 4, i => PptxWriter.Write(presentation, Path.Combine(OutDir, "p-" + i + ".pptx")),
                () => Require(PptxReader.Load(Path.Combine(OutDir, "p-0.pptx")).Slides.Count == presentation.Slides.Count)));
            results.Add(Run("ppt    read ", dop, 8, i => { OfficePresentation p = PptxReader.Load(pptxPath); Require(p.Slides.Count >= 1); }, null));

            results.Add(Run("pdf    read ", dop, 8, i => { PdfDocument d = PdfDocument.Load(pdfPath); Require(d.Pages.Count >= 1); }, null));
            results.Add(Run("pdf  convert", dop, 3, i => { byte[] bytes = OfficeConverter.ToPdf(wordDoc); Require(bytes.Length > 1000); }, null));

            results.Add(Run("mark   html ", dop, 30, i => { string html = Wisdom.Markdown.Markdown.ToHtml(markdown); Require(html.Length > 1000); }, null));
            results.Add(Run("mark roundtrp", dop, 15, i =>
            {
                OfficeDocument doc = Wisdom.Markdown.Markdown.ToDocument(markdown);
                string back = Wisdom.Markdown.Markdown.ToMarkdown(doc);
                Require(back.Length > 1000);
            }, null));

            results.Add(Run("facade load ", dop, 8, i => { OfficeFile f = OfficeFile.Load(xlsxPath); Require(f != null); }, null));
            Console.WriteLine("--- dop=" + dop + " done");
        }

        Console.WriteLine();
        Console.WriteLine("scenario      dop  elapsed_ms   ops/s      p50ms     p95ms   allocMB  gen0/gen1/gen2  err");
        foreach (Result r in results)
        {
            Console.WriteLine(r.Label + "  " + Pad(r.Dop.ToString(), 3) + Pad(r.ElapsedMs.ToString("0"), 10)
                + Pad((r.OpsPerSec).ToString("0.0"), 9) + Pad(r.P50.ToString("0.0"), 9) + Pad(r.P95.ToString("0.0"), 9)
                + Pad(r.AllocMb.ToString("0.0"), 9) + Pad(r.Gen0 + "/" + r.Gen1 + "/" + r.Gen2, 14)
                + (r.Errors == 0 ? "" : "  ERR=" + r.Errors + " " + r.FirstError));
        }
        return results.Any(r => r.Errors > 0) ? 1 : 0;
    }

    private sealed class Result
    {
        public string Label;
        public int Dop;
        public double ElapsedMs;
        public double OpsPerSec;
        public double P50;
        public double P95;
        public double AllocMb;
        public int Gen0;
        public int Gen1;
        public int Gen2;
        public int Errors;
        public string FirstError;
    }

    /// <summary>以 dop 个任务并发执行 op，统计墙钟吞吐、单次延迟分位、线程侧分配与 GC 次数。 Runs op dop-way concurrently and aggregates throughput, latency percentiles, allocation and GC counts.</summary>
    private static Result Run(string label, int dop, int opsPerThread, Action<int> op, Action verify)
    {
        // 预热一次，排除 JIT 与字体索引等一次性成本对并发数据的影响。
        // One warm-up pass so JIT and first-use indexes do not pollute the concurrent numbers.
        op(999);

        Result r = new Result();
        r.Label = label;
        r.Dop = dop;
        double[][] latencies = new double[dop][];
        // GetTotalAllocatedBytes 是进程级累计值，只能在主线程整体取差值，按线程求和会重复计数。
        // GetTotalAllocatedBytes is process-wide: take one delta on the main thread; summing per-thread deltas double counts.
        long allocBefore = GC.GetTotalAllocatedBytes(true);
        int g0 = GC.CollectionCount(0);
        int g1 = GC.CollectionCount(1);
        int g2 = GC.CollectionCount(2);
        Stopwatch sw = Stopwatch.StartNew();
        Task[] tasks = new Task[dop];
        for (int t = 0; t < dop; t++)
        {
            int threadIndex = t;
            tasks[t] = Task.Run(() =>
            {
                double[] times = new double[opsPerThread];
                for (int i = 0; i < opsPerThread; i++)
                {
                    Stopwatch one = Stopwatch.StartNew();
                    try { op(threadIndex); }
                    catch (Exception ex) { Interlocked.Increment(ref r.Errors); Interlocked.CompareExchange(ref r.FirstError, ex.GetType().Name + ": " + ex.Message, null); }
                    one.Stop();
                    times[i] = one.Elapsed.TotalMilliseconds;
                }
                latencies[threadIndex] = times;
            });
        }
        Task.WaitAll(tasks);
        sw.Stop();
        r.AllocMb = (GC.GetTotalAllocatedBytes(true) - allocBefore) / 1048576.0;
        r.ElapsedMs = sw.Elapsed.TotalMilliseconds;
        int totalOps = dop * opsPerThread;
        r.OpsPerSec = totalOps / (r.ElapsedMs / 1000.0);
        r.Gen0 = GC.CollectionCount(0) - g0;
        r.Gen1 = GC.CollectionCount(1) - g1;
        r.Gen2 = GC.CollectionCount(2) - g2;
        List<double> all = new List<double>();
        foreach (double[] arr in latencies) if (arr != null) all.AddRange(arr);
        all.Sort();
        r.P50 = all[(int)(all.Count * 0.50)];
        r.P95 = all[Math.Min(all.Count - 1, (int)(all.Count * 0.95))];
        if (verify != null && r.Errors == 0) verify();
        return r;
    }

    private static void Require(bool condition)
    {
        if (!condition) throw new InvalidOperationException("integrity check failed");
    }

    private static string Pad(string value, int width)
    {
        StringBuilder builder = new StringBuilder(value);
        while (builder.Length < width) builder.Append(' ');
        return builder.ToString();
    }

    // ---- 固定测试文档 fixtures ----

    private static string PrepareWord()
    {
        OfficeDocument document = new OfficeDocument();
        document.Metadata.Title = "并发压测文档";
        Section section = document.AddSection();
        for (int i = 0; i < 60; i++)
        {
            Paragraph p = section.AddParagraph("第 " + i + " 段：并发读写一致性验证文本，包含中英文 mixed content 与数字 1234567890。");
            if (i % 5 == 0)
            {
                TextRun run = p.Inlines[0] as TextRun;
                if (run != null) run.Format.Bold = true;
            }
        }
        for (int t = 0; t < 2; t++)
        {
            Table table = new Table();
            for (int r = 0; r < 20; r++) table.AddRow("r" + r, (r * 3.5).ToString("0.0"));
            section.Blocks.Add(table);
        }
        string path = Path.Combine(OutDir, "fixture.docx");
        DocxWriter.Write(document, path);
        return path;
    }

    private static string PrepareExcel()
    {
        OfficeWorkbook workbook = new OfficeWorkbook();
        for (int s = 0; s < 3; s++)
        {
            Worksheet sheet = workbook.AddWorksheet("表" + s);
            sheet.SetRowHeight(1, 24);
            for (int r = 1; r <= 1500; r++)
            {
                sheet.GetCell(r, 1).Value = "行" + r + "-s" + s;
                sheet.GetCell(r, 2).Value = r * 1.5;
                sheet.GetCell(r, 3).Value = (r % 7 == 0);
                sheet.GetCell(r, 4).Value = "文本列 " + r;
            }
        }
        string path = Path.Combine(OutDir, "fixture.xlsx");
        XlsxWriter.Write(workbook, path);
        return path;
    }

    private static string PreparePpt()
    {
        OfficePresentation presentation = OfficePresentation.Widescreen();
        for (int i = 0; i < 24; i++)
        {
            Slide slide = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
            TextBoxShape title = slide.Shapes.OfType<TextBoxShape>().FirstOrDefault(b => b.Placeholder == PlaceholderKind.Title);
            if (title != null) title.Paragraphs[0].AddText("并发页 " + i);
            TextBoxShape body = slide.AddTextBox(80, 160, 800, 300, string.Empty);
            for (int b = 0; b < 6; b++) body.AddParagraph("要点 " + b + "：并发吞吐验证内容。");
            slide.AddShape(AutoShapeKind.RoundedRectangle, 100 + i * 3, 420, 90, 50, "形状");
            slide.NotesText = "备注 " + i;
        }
        string path = Path.Combine(OutDir, "fixture.pptx");
        PptxWriter.Write(presentation, path);
        return path;
    }

    private static string PreparePdf(string wordPath)
    {
        OfficeDocument document = DocxReader.Load(wordPath);
        string path = Path.Combine(OutDir, "fixture.pdf");
        OfficeConverter.ToPdf(document, path);
        return path;
    }

    private static string PrepareMarkdown()
    {
        StringBuilder builder = new StringBuilder();
        builder.AppendLine("# 并发压测标题");
        for (int i = 0; i < 200; i++)
        {
            builder.AppendLine("## 小节 " + i);
            builder.AppendLine("这是第 " + i + " 段正文，包含 **粗体**、`代码` 与 [链接](https://example.com/" + i + ")。");
            builder.AppendLine("- 列表项 A");
            builder.AppendLine("- 列表项 B");
            builder.AppendLine();
        }
        string path = Path.Combine(OutDir, "fixture.md");
        File.WriteAllText(path, builder.ToString(), new UTF8Encoding(false));
        return path;
    }
}
