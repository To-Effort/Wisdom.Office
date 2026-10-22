using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Excel;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Word;

// Wisdom.* 逐库性能基准：冷启动、写入/读取/转换吞吐、延迟分布（min/median/P95）与托管分配。
// 用法：dotnet run -c Release --project benchmarks/Wisdom.Benchmarks
internal static class Program
{
    private sealed class Outcome
    {
        public double Units;
        public string Unit;
        public string Detail;
    }

    private sealed class Row
    {
        public string Library;
        public string Scenario;
        public int Iterations;
        public double MinMs;
        public double MedianMs;
        public double P95Ms;
        public long AllocatedPerOp;
        public long PeakHeap;
        public long PeakRss;
        public string Rate;
        public string Detail;
    }

    private static readonly List<Row> Rows = new List<Row>();
    private static readonly string Machine = Environment.ProcessorCount + " 核 / " + RuntimeInformation();

    private static string OutDir
    {
        get
        {
            // 按进程号分目录：并发或残留的基准轮次不会争抢同一个输出文件。
            string dir = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "out", "pid-" + Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture));
            Directory.CreateDirectory(dir);
            return dir;
        }
    }

    private static string FixtureDir
    {
        get { return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "Wisdom.Office.Tests", "fixtures")); }
    }

    private static string GeneratedDir
    {
        get { return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", "tests", "Wisdom.Office.Tests", "out")); }
    }

    private static void Main()
    {
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        Stopwatch warm = Stopwatch.StartNew();
        Console.WriteLine("=== Wisdom.* 性能基准 ===");
        Console.WriteLine("机器：" + Machine);
        Console.WriteLine("模式：" + (IsDebugRun() ? "Debug（数字不可用于结论）" : "Release"));
        Console.WriteLine();

        RunColdStart();
        RunCore();
        RunExcel();
        RunWord();
        RunPdf();
        RunOffice();

        warm.Stop();
        PrintTable();
        Console.WriteLine();
        Console.WriteLine("总耗时 " + Math.Round(warm.Elapsed.TotalSeconds, 1) + " 秒，输出目录：" + Path.GetFullPath(OutDir));
    }

    private static bool IsDebugRun()
    {
#if DEBUG
        return true;
#else
        return false;
#endif
    }

    private static string RuntimeInformation()
    {
        return System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription;
    }

    // ---------------- 冷启动：首次调用（含类型加载、JIT、字体扫描） ----------------

    private static void RunColdStart()
    {
        string xlsx = Path.Combine(OutDir, "bench-cold.xlsx");
        string docx = Path.Combine(OutDir, "bench-cold.docx");
        string pdf = Path.Combine(OutDir, "bench-cold.pdf");

        OfficeWorkbook tiny = BuildWorkbook(20, 5);
        OfficeDocument tinyDoc = BuildDocument(20, 2);

        Console.WriteLine("[冷启动] 首次 XlsxWriter.Write      " + One(() => XlsxWriter.Write(tiny, xlsx)) + " ms");
        Console.WriteLine("[冷启动] 首次 XlsxReader.Load       " + One(() => XlsxReader.Load(xlsx)) + " ms");
        Console.WriteLine("[冷启动] 首次 DocxWriter.Write      " + One(() => DocxWriter.Write(tinyDoc, docx)) + " ms");
        Console.WriteLine("[冷启动] 首次 DocxReader.Load       " + One(() => DocxReader.Load(docx)) + " ms");
        Console.WriteLine("[冷启动] 首次 PdfRenderer.Render    " + One(() => new PdfRenderer().Render(tinyDoc, pdf)) + " ms（含字体加载与子集化）");
        Console.WriteLine("[冷启动] 首次 PdfDocument.Load      " + One(() => PdfDocument.Load(pdf)) + " ms");
        Console.WriteLine("[冷启动] 首次 OfficeFile.Load       " + One(() => OfficeFile.Load(xlsx)) + " ms");
        Console.WriteLine();
    }

    private static double One(Action action)
    {
        Stopwatch sw = Stopwatch.StartNew();
        action();
        sw.Stop();
        return Math.Round(sw.Elapsed.TotalMilliseconds, 1);
    }

    // ---------------- Wisdom.Office.Core：统一模型与格式化（纯内存） ----------------

    private static void RunCore()
    {
        Measure("Core", "构建 5 万单元格工作簿（样式+数字格式）", 5, () =>
        {
            OfficeWorkbook workbook = BuildWorkbook(2500, 20);
            return new Outcome { Units = 2500 * 20, Unit = "cells", Detail = "含 3 种共享样式" };
        });

        Measure("Core", "数字/日期 DisplayText 格式化", 5, () =>
        {
            Worksheet sheet = BuildWorkbook(1000, 6).Worksheets[0];
            int calls = 0;
            for (int r = 2; r <= 1000; r++)
                for (int c = 1; c <= 6; c++) { sheet.GetCell(r, c).DisplayText(); calls++; }
            return new Outcome { Units = calls, Unit = "formats" };
        });

        Measure("Core", "构建统一文档模型（段落+表格）", 5, () =>
        {
            OfficeDocument document = BuildDocument(8000, 400);
            return new Outcome { Units = 8000 + 400, Unit = "blocks" };
        });

        Measure("Core", "工作簿内存驻留（1 万行 × 20 列）", 3, () =>
        {
            OfficeWorkbook workbook = BuildWorkbook(10000, 20);
            long heap = GC.GetTotalMemory(false);
            GC.KeepAlive(workbook);
            return new Outcome { Units = 10000 * 20, Unit = "cells", Detail = "托管堆 " + Mb(heap) };
        });
    }

    // ---------------- Wisdom.Excel：XLSX 写入 / 读取 / 旧版 XLS ----------------

    private static void RunExcel()
    {
        string small = Path.Combine(OutDir, "bench-excel-1k.xlsx");
        string medium = Path.Combine(OutDir, "bench-excel-10k.xlsx");
        string large = Path.Combine(OutDir, "bench-excel-50k.xlsx");

        Measure("Excel", "XlsxWriter.Write 1 千行 × 20 列", 5, () =>
        {
            OfficeWorkbook workbook = BuildWorkbook(1000, 20);
            XlsxWriter.Write(workbook, small);
            return new Outcome { Units = 1000, Unit = "rows", Detail = FileSize(small) };
        });

        Measure("Excel", "XlsxWriter.Write 1 万行 × 20 列", 3, () =>
        {
            OfficeWorkbook workbook = BuildWorkbook(10000, 20);
            XlsxWriter.Write(workbook, medium);
            return new Outcome { Units = 10000, Unit = "rows", Detail = FileSize(medium) };
        });

        Measure("Excel", "XlsxWriter.Write 5 万行 × 20 列（100 万单元格）", 1, () =>
        {
            OfficeWorkbook workbook = BuildWorkbook(50000, 20);
            XlsxWriter.Write(workbook, large);
            return new Outcome { Units = 50000, Unit = "rows", Detail = FileSize(large) };
        });

        Measure("Excel", "XlsxReader.Load 5 万行（100 万单元格）", 1, () =>
        {
            OfficeWorkbook loaded = XlsxReader.Load(large);
            int cells = loaded.Worksheets[0].LastRow * 20;
            return new Outcome { Units = cells, Unit = "cells", Detail = "回读 " + loaded.Worksheets[0].LastRow + " 行" };
        });

        Measure("Excel", "XlsxReader.Load 1 万行", 3, () =>
        {
            OfficeWorkbook loaded = XlsxReader.Load(medium);
            return new Outcome { Units = 10000 * 20, Unit = "cells", Detail = "回读 " + loaded.Worksheets[0].LastRow + " 行" };
        });

        string attendance = Path.Combine(FixtureDir, "attendance-real.xlsx");
        if (File.Exists(attendance))
        {
            Measure("Excel", "XlsxReader.Load 真实考勤表（5 表/合并/格式）", 3, () =>
            {
                OfficeWorkbook loaded = XlsxReader.Load(attendance);
                int sheets = loaded.Worksheets.Count;
                return new Outcome { Units = sheets, Unit = "sheets", Detail = new FileInfo(attendance).Length + " 字节" };
            });
        }

        string legacyXls = Path.Combine(FixtureDir, "legacy-solver.xls");
        if (File.Exists(legacyXls))
        {
            Measure("Excel", "XlsReader.Load 旧版 BIFF8（6 表）", 3, () =>
            {
                OfficeWorkbook loaded = XlsReader.Load(legacyXls);
                return new Outcome { Units = loaded.Worksheets.Count, Unit = "sheets", Detail = "BIFF8 OLE2 解析" };
            });
        }
    }

    // ---------------- Wisdom.Word：DOCX 写入 / 读取 / 整理 / 旧版 DOC ----------------

    private static void RunWord()
    {
        string big = Path.Combine(OutDir, "bench-word-20k.docx");

        Measure("Word", "DocxWriter.Write 2 万段落 + 200 表格", 3, () =>
        {
            OfficeDocument document = BuildDocument(20000, 200);
            DocxWriter.Write(document, big);
            return new Outcome { Units = 20000 + 200, Unit = "blocks", Detail = FileSize(big) };
        });

        Measure("Word", "DocxReader.Load 2 万段落文档", 3, () =>
        {
            OfficeDocument loaded = DocxReader.Load(big);
            int blocks = CountBlocks(loaded);
            return new Outcome { Units = blocks, Unit = "blocks", Detail = "回读 " + loaded.Sections.Count + " 节" };
        });

        Measure("Word", "WordNormalizer.Normalize 2 万段落", 3, () =>
        {
            OfficeDocument loaded = DocxReader.Load(big);
            WordNormalizeResult result = WordNormalizer.Normalize(loaded, new WordNormalizeOptions());
            return new Outcome { Units = CountBlocks(loaded), Unit = "blocks", Detail = "变更 " + result.Changes.Count + " 处" };
        });

        string complex = Path.Combine(GeneratedDir, "complex-report.docx");
        if (File.Exists(complex))
        {
            Measure("Word", "DocxReader.Load 含 12 张 PNG 的真实文档", 3, () =>
            {
                OfficeDocument loaded = DocxReader.Load(complex);
                return new Outcome { Units = new FileInfo(complex).Length, Unit = "bytes", Detail = FileSize(complex) };
            });
        }

        string legacyDoc = Path.Combine(FixtureDir, "legacy-sample.doc");
        if (File.Exists(legacyDoc))
        {
            Measure("Word", "DocReader.Load 旧版 .doc（Word 97-2003）", 3, () =>
            {
                OfficeDocument loaded = DocReader.Load(legacyDoc);
                return new Outcome { Units = CountBlocks(loaded), Unit = "blocks", Detail = "二进制 FIB 解析" };
            });
        }

        Measure("Word", "WordMailMerge 扁平合并（2 万段 × 3 标记）", 3, () =>
        {
            OfficeDocument document = BuildMailMergeDocument(20000);
            TemplateContext context = new TemplateContext();
            context.Set("name", "客户名称");
            context.Set("amount", 1234.5m);
            context.Set("code", "A-001");
            int merged = WordMailMerge.Execute(document, context);
            return new Outcome { Units = merged, Unit = "merges", Detail = "字典上下文" };
        });

        Measure("Word", "WordMailMerge.ExecuteWithRegions 嵌套区域（500 单 × 3 行）", 2, () =>
        {
            OfficeDocument document = BuildNestedRegionDocument();
            TemplateContext context = new TemplateContext();
            List<object> orders = new List<object>();
            for (int i = 0; i < 500; i++)
            {
                Dictionary<string, object> order = new Dictionary<string, object>();
                order["orderNo"] = "A-" + i.ToString(CultureInfo.InvariantCulture);
                order["customer"] = "客户 " + i.ToString(CultureInfo.InvariantCulture);
                List<object> lines = new List<object>();
                for (int l = 0; l < 3; l++)
                {
                    Dictionary<string, object> line = new Dictionary<string, object>();
                    line["sku"] = "SKU-" + l.ToString(CultureInfo.InvariantCulture);
                    line["qty"] = l + 1;
                    lines.Add(line);
                }
                order["lines"] = lines;
                orders.Add(order);
            }
            context.Set("orders", orders);
            int merged = WordMailMerge.ExecuteWithRegions(document, context);
            return new Outcome { Units = merged, Unit = "merges", Detail = "父区域 + 子区域" };
        });

        Measure("Word", "WordFindReplace 2 万段落（6 万处命中）", 3, () =>
        {
            OfficeDocument document = BuildDocument(20000, 50);
            int count = WordFindReplace.Replace(document, "正文内容", "body text");
            return new Outcome { Units = count, Unit = "replacements", Detail = "跨 Run 匹配" };
        });

        string formatSource = Path.Combine(OutDir, "bench-word-format-source.docx");
        OfficeDocument formatDocument = BuildDocument(2000, 20);
        DocxWriter.Write(formatDocument, formatSource);
        Measure("Word", "WordFormatConverter DOCX → ODT（2 千段往返）", 3, () =>
        {
            byte[] odt = WordFormatConverter.ToArray(formatDocument, WordDocumentFormat.Odt);
            OfficeDocument loaded = WordFormatConverter.Load(odt, WordDocumentFormat.Odt);
            return new Outcome { Units = CountBlocks(loaded), Unit = "blocks", Detail = odt.Length + " 字节" };
        });
        Measure("Word", "WordFormatConverter DOCX → RTF（2 千段往返）", 3, () =>
        {
            byte[] rtf = WordFormatConverter.ToArray(formatDocument, WordDocumentFormat.Rtf);
            OfficeDocument loaded = WordFormatConverter.Load(rtf, WordDocumentFormat.Rtf);
            return new Outcome { Units = CountBlocks(loaded), Unit = "blocks", Detail = rtf.Length + " 字节" };
        });
        Measure("Word", "WordFormatConverter DOCX → HTML（2 千段往返）", 3, () =>
        {
            byte[] html = WordFormatConverter.ToArray(formatDocument, WordDocumentFormat.Html);
            OfficeDocument loaded = WordFormatConverter.Load(html, WordDocumentFormat.Html);
            return new Outcome { Units = CountBlocks(loaded), Unit = "blocks", Detail = html.Length + " 字节" };
        });
    }

    /// <summary>构建邮件合并基准文档：每段包含三个 {{token}}。Builds the mail-merge benchmark document: three {{token}} markers per paragraph.</summary>
    private static OfficeDocument BuildMailMergeDocument(int paragraphs)
    {
        OfficeDocument document = new OfficeDocument();
        Section section = document.AddSection();
        for (int i = 0; i < paragraphs; i++)
        {
            Paragraph paragraph = new Paragraph("{{code}} ");
            paragraph.AddText("{{name}}");
            paragraph.AddText(" 金额 {{amount:N2}} 元");
            section.Blocks.Add(paragraph);
        }
        return document;
    }

    /// <summary>构建嵌套区域基准表格：orders 区域包含 lines 子区域。Builds the nested-region benchmark table: an orders region containing a lines child region.</summary>
    private static OfficeDocument BuildNestedRegionDocument()
    {
        OfficeDocument document = new OfficeDocument();
        Section section = document.AddSection();
        Table table = new Table();
        table.Columns.Add(new TableColumn(120));
        table.Columns.Add(new TableColumn(120));
        table.Rows.Add(CreateBenchmarkMarkerRow("MERGEFIELD TableStart:orders", 2));
        TableRow orderRow = table.AddRow();
        orderRow.Cells.Add(new TableCell("{{orderNo}}"));
        orderRow.Cells.Add(new TableCell("{{customer}}"));
        table.Rows.Add(CreateBenchmarkMarkerRow("MERGEFIELD TableStart:lines", 2));
        TableRow lineRow = table.AddRow();
        lineRow.Cells.Add(new TableCell("{{sku}}"));
        lineRow.Cells.Add(new TableCell("{{qty}}"));
        table.Rows.Add(CreateBenchmarkMarkerRow("MERGEFIELD TableEnd:lines", 2));
        table.Rows.Add(CreateBenchmarkMarkerRow("MERGEFIELD TableEnd:orders", 2));
        section.Blocks.Add(table);
        return document;
    }

    /// <summary>创建只含一个区域域标记的两列表行。Creates a two-cell table row holding a single region field marker.</summary>
    private static TableRow CreateBenchmarkMarkerRow(string instruction, int cellCount)
    {
        TableRow row = new TableRow();
        for (int i = 0; i < cellCount; i++)
        {
            TableCell cell = new TableCell();
            if (i == 0)
            {
                Paragraph paragraph = new Paragraph();
                paragraph.Inlines.Add(new FieldRun(instruction, string.Empty, new TextFormat("Arial", 9)));
                cell.Blocks.Add(paragraph);
            }
            row.Cells.Add(cell);
        }
        return row;
    }

    // ---------------- Wisdom.Pdf：生成 / 解析 / 表格识别 ----------------

    private static void RunPdf()
    {
        string pdf50 = Path.Combine(OutDir, "bench-pdf-50.doc.pdf");
        string pdf200 = Path.Combine(OutDir, "bench-pdf-200.pdf");

        Measure("Pdf", "PdfRenderer.Render 约 50 页", 3, () =>
        {
            OfficeDocument document = BuildDocument(4000, 60);
            new PdfRenderer().Render(document, pdf50);
            int pages = PdfDocument.Load(pdf50).PageCount;
            return new Outcome { Units = pages, Unit = "pages", Detail = FileSize(pdf50) };
        });

        Measure("Pdf", "PdfRenderer.Render 大文档（8 千段落 + 400 表格）", 1, () =>
        {
            OfficeDocument document = BuildDocument(8000, 400);
            new PdfRenderer().Render(document, pdf200);
            int pages = PdfDocument.Load(pdf200).PageCount;
            return new Outcome { Units = pages, Unit = "pages", Detail = FileSize(pdf200) };
        });

        Measure("Pdf", "PdfDocument.Load 200 页", 3, () =>
        {
            PdfDocument pdf = PdfDocument.Load(pdf200);
            return new Outcome { Units = pdf.PageCount, Unit = "pages", Detail = "xref + 页树解析" };
        });

        Measure("Pdf", "ExtractText 200 页", 1, () =>
        {
            PdfDocument pdf = PdfDocument.Load(pdf200);
            string text = pdf.ExtractText();
            return new Outcome { Units = pdf.PageCount, Unit = "pages", Detail = "提取 " + text.Length + " 字符" };
        });

        Measure("Pdf", "GetTables 表格识别 200 页", 1, () =>
        {
            PdfDocument pdf = PdfDocument.Load(pdf200);
            int tables = 0;
            foreach (PdfPage page in pdf.Pages) tables += page.GetTables().Count;
            return new Outcome { Units = pdf.PageCount, Unit = "pages", Detail = "识别 " + tables + " 个表格" };
        });
    }

    // ---------------- Wisdom.Office：版面换算 / 报表 / 端到端转换 ----------------

    private static void RunOffice()
    {
        string reportXlsx = Path.Combine(OutDir, "bench-report.xlsx");
        string reportDocx = Path.Combine(OutDir, "bench-report.docx");
        string reportPdf = Path.Combine(OutDir, "bench-report.pdf");
        string medium = Path.Combine(OutDir, "bench-excel-10k.xlsx");
        string layoutPdf = Path.Combine(OutDir, "bench-layout-10k.pdf");
        string layoutDocx = Path.Combine(OutDir, "bench-layout-10k.docx");
        string pdfToDocx = Path.Combine(OutDir, "bench-pdf2docx.docx");

        Measure("Office", "Report.Export -> XLSX", 5, () =>
        {
            BuildReport().Export(reportXlsx);
            return new Outcome { Units = new FileInfo(reportXlsx).Length, Unit = "bytes", Detail = FileSize(reportXlsx) };
        });

        Measure("Office", "Report.Export -> DOCX", 5, () =>
        {
            BuildReport().Export(reportDocx);
            return new Outcome { Units = new FileInfo(reportDocx).Length, Unit = "bytes", Detail = FileSize(reportDocx) };
        });

        Measure("Office", "Report.Export -> PDF", 3, () =>
        {
            BuildReport().Export(reportPdf);
            return new Outcome { Units = new FileInfo(reportPdf).Length, Unit = "bytes", Detail = FileSize(reportPdf) };
        });

        Measure("Office", "WorkbookToDocumentConverter 版面分带（1 万行）", 3, () =>
        {
            OfficeWorkbook workbook = XlsxReader.Load(medium);
            OfficeDocument layout = WorkbookToDocumentConverter.Convert(workbook);
            return new Outcome { Units = 10000 * 20, Unit = "cells", Detail = "版面块 " + CountBlocks(layout) + " 个 / " + layout.Sections.Count + " 节" };
        });

        Measure("Office", "OfficeConverter.Convert XLSX -> PDF（1 万行）", 1, () =>
        {
            OfficeConverter.Convert(medium, layoutPdf);
            int pages = PdfDocument.Load(layoutPdf).PageCount;
            return new Outcome { Units = pages, Unit = "pages", Detail = FileSize(layoutPdf) };
        });

        Measure("Office", "OfficeConverter.Convert XLSX -> DOCX（1 万行）", 1, () =>
        {
            OfficeConverter.Convert(medium, layoutDocx);
            return new Outcome { Units = new FileInfo(layoutDocx).Length, Unit = "bytes", Detail = FileSize(layoutDocx) };
        });

        Measure("Office", "OfficeConverter.Convert PDF -> DOCX", 1, () =>
        {
            OfficeConverter.Convert(reportPdf, pdfToDocx);
            return new Outcome { Units = new FileInfo(pdfToDocx).Length, Unit = "bytes", Detail = FileSize(pdfToDocx) };
        });

        Measure("Office", "OfficeFile.Load + ExtractText（DOCX）", 3, () =>
        {
            OfficeFile file = OfficeFile.Load(reportDocx);
            string text = file.ExtractText();
            return new Outcome { Units = text.Length, Unit = "chars", Detail = "格式 " + file.Format };
        });

        Measure("Office", "OfficeFileFormats.Detect 魔数识别", 5, () =>
        {
            int calls = 0;
            for (int i = 0; i < 20000; i++) { OfficeFileFormats.Detect(reportXlsx); calls++; }
            return new Outcome { Units = calls, Unit = "detects" };
        });

        string complexExcel = Path.Combine(GeneratedDir, "complex-excel.xlsx");
        if (File.Exists(complexExcel))
        {
            Measure("Office", "端到端：真实 5 千行考勤数据 XLSX -> PDF", 1, () =>
            {
                string target = Path.Combine(OutDir, "bench-e2e-real.pdf");
                OfficeConverter.Convert(complexExcel, target);
                int pages = PdfDocument.Load(target).PageCount;
                return new Outcome { Units = pages, Unit = "pages", Detail = FileSize(target) };
            });
        }
    }

    // ---------------- 数据构造 ----------------

    private static OfficeWorkbook BuildWorkbook(int rows, int columns)
    {
        OfficeWorkbook workbook = new OfficeWorkbook();
        workbook.Metadata.Title = "压测数据";
        Worksheet sheet = workbook.AddWorksheet("明细");

        CellStyle header = new CellStyle
        {
            Bold = true,
            HorizontalAlignment = HorizontalAlignment.Center,
            FillColor = new OfficeColor((byte)217, (byte)226, (byte)243)
        };
        CellStyle money = new CellStyle { NumberFormat = "#,##0.00" };
        CellStyle date = new CellStyle { NumberFormat = "yyyy-MM-dd" };
        DateTime day = new DateTime(2026, 1, 1);

        for (int c = 1; c <= columns; c++)
        {
            Cell cell = sheet.GetCell(1, c);
            cell.Value = "列" + c.ToString(CultureInfo.InvariantCulture);
            cell.Style = header;
        }

        for (int r = 2; r <= rows; r++)
        {
            for (int c = 1; c <= columns; c++)
            {
                Cell cell = sheet.GetCell(r, c);
                if (c == 1) cell.Value = "行" + r.ToString(CultureInfo.InvariantCulture);
                else if (c % 7 == 0) { cell.Value = day.AddDays(r % 300); cell.Style = date; }
                else if (c % 3 == 0) { cell.Value = r * 1.37; cell.Style = money; }
                else cell.Value = r * c;
            }
        }

        sheet.FrozenRows = 1;
        sheet.SetColumnWidth(1, 20);
        return workbook;
    }

    private static OfficeDocument BuildDocument(int paragraphs, int tables)
    {
        OfficeDocument document = new OfficeDocument();
        document.Metadata.Title = "压测文档";
        Section section = document.AddSection(PageSetup.A4());
        TextFormat small = new TextFormat("Arial", 9);

        for (int i = 0; i < paragraphs; i++)
        {
            Paragraph paragraph = new Paragraph("段落 " + i.ToString(CultureInfo.InvariantCulture) + "：正文内容用于测量排版吞吐。");
            paragraph.AddText("（附注 " + i.ToString(CultureInfo.InvariantCulture) + "）", small);
            section.Blocks.Add(paragraph);
        }

        for (int t = 0; t < tables; t++)
        {
            Table table = new Table();
            table.Columns.Add(new TableColumn(160));
            table.Columns.Add(new TableColumn(100));
            table.AddRow("项目", "金额");
            for (int r = 0; r < 8; r++) table.AddRow("条目 " + r.ToString(CultureInfo.InvariantCulture), (r * 1234.5).ToString("#,##0.00", CultureInfo.InvariantCulture));
            section.Blocks.Add(table);
        }

        return document;
    }

    private static Report BuildReport()
    {
        Report report = new Report("性能基准报表");
        report.Subtitle = "Wisdom.Office";
        report.Author = "Wisdom";
        ReportTable table = report.AddTable("明细");
        table.AddColumn("项目", 120, null, HorizontalAlignment.Left);
        table.AddColumn("数量", 60, "0", HorizontalAlignment.Right);
        table.AddColumn("金额", 100, "#,##0.00", HorizontalAlignment.Right);
        table.AddColumn("日期", 90, "yyyy-MM-dd", HorizontalAlignment.Center);
        for (int i = 0; i < 1500; i++)
            table.AddRow("条目 " + i.ToString(CultureInfo.InvariantCulture), i, i * 12.5, new DateTime(2026, 1, 1).AddDays(i % 300));
        report.AddNote("金额单位：人民币元。");
        return report;
    }

    private static int CountBlocks(OfficeDocument document)
    {
        int count = 0;
        foreach (Section section in document.Sections) count += section.Blocks.Count;
        return count;
    }

    private static string FileSize(string path)
    {
        return new FileInfo(path).Length.ToString("N0", CultureInfo.InvariantCulture) + " 字节";
    }

    private static string Mb(long bytes)
    {
        return Math.Round(bytes / 1048576.0, 1) + " MB";
    }

    // ---------------- 测量与输出 ----------------

    private static void Measure(string library, string scenario, int iterations, Func<Outcome> action)
    {
        Outcome outcome = action();   // 预热：触发类型加载与 JIT，不计入统计

        double[] samples = new double[iterations];
        long peakHeap = 0;
        long peakWorkingSet = 0;
        Process process = Process.GetCurrentProcess();
        long allocatedBefore = GC.GetTotalAllocatedBytes(true);
        for (int i = 0; i < iterations; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Stopwatch sw = Stopwatch.StartNew();
            outcome = action() ?? outcome;
            sw.Stop();
            samples[i] = sw.Elapsed.TotalMilliseconds;
            long heap = GC.GetTotalMemory(false);
            if (heap > peakHeap) peakHeap = heap;
            long workingSet = process.WorkingSet64;
            if (workingSet > peakWorkingSet) peakWorkingSet = workingSet;
        }
        long allocated = (GC.GetTotalAllocatedBytes(true) - allocatedBefore) / Math.Max(1, iterations);

        Array.Sort(samples);
        double min = samples[0];
        double median = samples[samples.Length / 2];
        double p95 = samples[Math.Min(samples.Length - 1, (int)Math.Ceiling(samples.Length * 0.95) - 1)];

        string rate = Rate(outcome, median);
        Rows.Add(new Row
        {
            Library = library,
            Scenario = scenario,
            Iterations = iterations,
            MinMs = min,
            MedianMs = median,
            P95Ms = p95,
            AllocatedPerOp = allocated,
            PeakHeap = peakHeap,
            PeakRss = peakWorkingSet,
            Rate = rate,
            Detail = outcome == null || outcome.Detail == null ? string.Empty : outcome.Detail
        });

        Console.WriteLine("  [" + library + "] " + scenario + " -> 中位 " + Math.Round(median, 1) + " ms"
            + (string.IsNullOrEmpty(rate) ? string.Empty : "，" + rate)
            + "，分配 " + Mb(allocated) + "，峰值 RSS " + Mb(peakWorkingSet));
    }

    private static string Rate(Outcome outcome, double medianMs)
    {
        if (outcome == null || outcome.Units <= 0 || medianMs <= 0) return string.Empty;
        double perSecond = outcome.Units / (medianMs / 1000.0);
        string value;
        if (perSecond >= 1000000) value = Math.Round(perSecond / 1000000.0, 2) + "M";
        else if (perSecond >= 1000) value = Math.Round(perSecond / 1000.0, 1) + "k";
        else value = Math.Round(perSecond, 1).ToString(CultureInfo.InvariantCulture);
        return value + " " + outcome.Unit + "/s";
    }

    private static void PrintTable()
    {
        const int Width = 130;
        Console.WriteLine();
        Console.WriteLine("================================ 汇总表 ================================");
        Console.WriteLine(Pad("库", 8) + Pad("场景", 44) + Right("次数", 6) + Right("最小ms", 10)
            + Right("中位ms", 10) + Right("P95ms", 10) + Right("吞吐", 18) + Right("分配", 10) + Right("峰值RSS", 12));
        Console.WriteLine(new string('-', Width));
        foreach (Row row in Rows)
        {
            Console.WriteLine(Pad(row.Library, 8) + Pad(Trim(row.Scenario, 42), 44) + Right(row.Iterations.ToString(), 6)
                + Right(Math.Round(row.MinMs, 1).ToString(CultureInfo.InvariantCulture), 10)
                + Right(Math.Round(row.MedianMs, 1).ToString(CultureInfo.InvariantCulture), 10)
                + Right(Math.Round(row.P95Ms, 1).ToString(CultureInfo.InvariantCulture), 10)
                + Right(row.Rate, 18) + Right(Mb(row.AllocatedPerOp), 10) + Right(Mb(row.PeakRss), 12));
            if (!string.IsNullOrEmpty(row.Detail)) Console.WriteLine("         └─ " + row.Detail);
        }
        Console.WriteLine(new string('=', Width));
        Console.WriteLine("整轮进程峰值 RSS：" + Mb(Process.GetCurrentProcess().PeakWorkingSet64));
    }

    private static string Trim(string value, int max)
    {
        return value.Length <= max ? value : value.Substring(0, max - 1) + "…";
    }

    private static string Pad(string value, int width)
    {
        value = value ?? string.Empty;
        int visual = 0;
        foreach (char ch in value) visual += ch > 0x2E00 ? 2 : 1;
        int padding = Math.Max(0, width - visual);
        return value + new string(' ', padding);
    }

    private static string Right(string value, int width)
    {
        value = value ?? string.Empty;
        int visual = 0;
        foreach (char ch in value) visual += ch > 0x2E00 ? 2 : 1;
        int padding = Math.Max(1, width - visual);
        return new string(' ', padding) + value;
    }
}
