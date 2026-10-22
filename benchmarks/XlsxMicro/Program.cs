using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using Wisdom.Excel;
using Wisdom.Office.Core;

internal static class XlsxMicro
{
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

    private static void Run(string label, int rows, int iterations)
    {
        string path = Path.Combine(AppContext.BaseDirectory, "micro-" + rows + ".xlsx");
        Action warm = () => XlsxWriter.Write(BuildWorkbook(rows, 20), path);
        warm();

        double[] samples = new double[iterations];
        long allocatedBefore = GC.GetTotalAllocatedBytes(true);
        for (int i = 0; i < iterations; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Stopwatch sw = Stopwatch.StartNew();
            XlsxWriter.Write(BuildWorkbook(rows, 20), path);
            sw.Stop();
            samples[i] = sw.Elapsed.TotalMilliseconds;
        }
        long allocated = (GC.GetTotalAllocatedBytes(true) - allocatedBefore) / iterations;
        Array.Sort(samples);
        Console.WriteLine(label +
            " min=" + samples[0].ToString("F1", CultureInfo.InvariantCulture) +
            " med=" + samples[iterations / 2].ToString("F1", CultureInfo.InvariantCulture) +
            " p95=" + samples[Math.Min(iterations - 1, (int)Math.Ceiling(iterations * 0.95) - 1)].ToString("F1", CultureInfo.InvariantCulture) +
            " alloc=" + (allocated / 1048576.0).ToString("F1", CultureInfo.InvariantCulture) + " MB" +
            " file=" + new FileInfo(path).Length);

        // 模型构建与纯写出拆分：写出到内存流，排除文件 IO 与模型构造。
        // Splits model construction from pure serialization into a memory stream.
        OfficeWorkbook built = BuildWorkbook(rows, 20);
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        Stopwatch modelSw = Stopwatch.StartNew();
        long modelBefore = GC.GetTotalAllocatedBytes(true);
        OfficeWorkbook again = BuildWorkbook(rows, 20);
        modelSw.Stop();
        long modelAlloc = GC.GetTotalAllocatedBytes(true) - modelBefore;
        again = null;
        Console.WriteLine("    仅构建模型 time=" + modelSw.Elapsed.TotalMilliseconds.ToString("F1", CultureInfo.InvariantCulture) +
            " ms alloc=" + (modelAlloc / 1048576.0).ToString("F1", CultureInfo.InvariantCulture) + " MB");

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
                XlsxWriter.Write(built, stream);
            }
            sw.Stop();
            writeSamples[i] = sw.Elapsed.TotalMilliseconds;
        }
        long writeAlloc = (GC.GetTotalAllocatedBytes(true) - writeBefore) / writeIterations;
        Array.Sort(writeSamples);
        Console.WriteLine("    仅写出(内存流) min=" + writeSamples[0].ToString("F1", CultureInfo.InvariantCulture) +
            " ms alloc=" + (writeAlloc / 1048576.0).ToString("F1", CultureInfo.InvariantCulture) + " MB");

        // 读回：同一文件的 XlsxReader.Load 与 ExtractText。
        // Read-back: XlsxReader.Load and ExtractText over the same file.
        byte[] data = File.ReadAllBytes(path);
        double[] readSamples = new double[writeIterations];
        long readBefore = GC.GetTotalAllocatedBytes(true);
        int cellCount = 0;
        for (int i = 0; i < writeIterations; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Stopwatch sw = Stopwatch.StartNew();
            OfficeWorkbook loaded = XlsxReader.Load(data);
            sw.Stop();
            readSamples[i] = sw.Elapsed.TotalMilliseconds;
            cellCount = System.Linq.Enumerable.Count(loaded.Worksheets[0].Cells);
        }
        long readAlloc = (GC.GetTotalAllocatedBytes(true) - readBefore) / writeIterations;
        Array.Sort(readSamples);
        Console.WriteLine("    仅读回 min=" + readSamples[0].ToString("F1", CultureInfo.InvariantCulture) +
            " ms alloc=" + (readAlloc / 1048576.0).ToString("F1", CultureInfo.InvariantCulture) +
            " MB cells=" + cellCount);

        Stopwatch textSw = Stopwatch.StartNew();
        string text = XlsxReader.ExtractText(path);
        textSw.Stop();
        Console.WriteLine("    ExtractText=" + textSw.Elapsed.TotalMilliseconds.ToString("F1", CultureInfo.InvariantCulture) +
            " ms chars=" + text.Length);
    }

    private static int Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "address") { return AddressCheck(); }
        if (args.Length > 0 && args[0] == "read")
        {
            Run("1 万行 × 20 列  ", 10000, 3);
            XmlTreeProbe();
            return 0;
        }
        PhaseProbe();
        Run("1 千行 × 20 列  ", 1000, 5);
        Run("1 万行 × 20 列  ", 10000, 3);
        Run("5 万行 × 20 列  ", 50000, 1);
        return 0;
    }

    /// <summary>对比整棵 XDocument 与 XmlReader 单遍扫描同一工作表部件的开销。 Compares a whole XDocument tree against a single XmlReader pass over the same worksheet part.</summary>
    private static void XmlTreeProbe()
    {
        string path = Path.Combine(AppContext.BaseDirectory, "micro-10000.xlsx");
        byte[] data = File.ReadAllBytes(path);

        using (MemoryStream first = new MemoryStream(data, false))
        using (System.IO.Compression.ZipArchive zip = new System.IO.Compression.ZipArchive(first, System.IO.Compression.ZipArchiveMode.Read, false))
        {
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml");
            long before = GC.GetTotalAllocatedBytes(true);
            Stopwatch sw = Stopwatch.StartNew();
            int elements = 0;
            using (System.IO.Stream stream = entry.Open())
            {
                System.Xml.Linq.XDocument doc = System.Xml.Linq.XDocument.Load(stream);
                foreach (System.Xml.Linq.XElement ignored in doc.Descendants()) elements++;
            }
            sw.Stop();
            Console.WriteLine("XDocument 树 min=" + elements + " 元素 time=" + sw.Elapsed.TotalMilliseconds.ToString("F1", CultureInfo.InvariantCulture) +
                " ms alloc=" + ((GC.GetTotalAllocatedBytes(true) - before) / 1048576.0).ToString("F1", CultureInfo.InvariantCulture) + " MB");

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            before = GC.GetTotalAllocatedBytes(true);
            sw = Stopwatch.StartNew();
            int nodes = 0;
            using (System.IO.Stream stream = entry.Open())
            {
                using (System.Xml.XmlReader reader = System.Xml.XmlReader.Create(stream))
                {
                    while (reader.Read()) if (reader.NodeType == System.Xml.XmlNodeType.Element) nodes++;
                }
            }
            sw.Stop();
            Console.WriteLine("XmlReader 扫描 min=" + nodes + " 元素 time=" + sw.Elapsed.TotalMilliseconds.ToString("F1", CultureInfo.InvariantCulture) +
                " ms alloc=" + ((GC.GetTotalAllocatedBytes(true) - before) / 1048576.0).ToString("F1", CultureInfo.InvariantCulture) + " MB");
        }
    }

    /// <summary>用反射把内部地址拼装与 CellAddress.Format 做穷举比对。 Brute-forces the internal address builder against CellAddress.Format through reflection.</summary>
    private static int AddressCheck()
    {
        Type stateType = typeof(XlsxWriter).GetNestedType("WriterState", System.Reflection.BindingFlags.NonPublic);
        object state = Activator.CreateInstance(stateType,
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic,
            null, new object[] { new OfficeWorkbook() }, CultureInfo.InvariantCulture);
        System.Reflection.MethodInfo format = stateType.GetMethod("CellAddressText",
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);

        int[] rows = { 1, 2, 3, 9, 10, 11, 25, 26, 27, 52, 53, 99, 100, 101, 999, 1000, 1001, 1048575, 1048576 };
        int checkedCount = 0;
        int bad = 0;
        for (int column = 1; column <= 17000; column++)
        {
            for (int i = 0; i < rows.Length; i++)
            {
                string expected = CellAddress.Format(rows[i], column);
                string actual = (string)format.Invoke(state, new object[] { rows[i], column });
                checkedCount++;
                if (expected != actual)
                {
                    bad++;
                    if (bad < 10) Console.WriteLine("MISMATCH r=" + rows[i] + " c=" + column + " expected=" + expected + " actual=" + actual);
                }
            }
        }
        Console.WriteLine("address check: compared=" + checkedCount + " mismatches=" + bad);
        return bad == 0 ? 0 : 1;
    }

    /// <summary>分离测量各阶段的裸开销，用于定位剩余分配与耗时来源。 Isolates raw per-phase costs to locate remaining allocations and time.</summary>
    private static void PhaseProbe()
    {
        const int N = 1000000;
        double[] numbers = new double[N];
        for (int i = 0; i < N; i++) numbers[i] = (i + 2) * 1.37;

        long before = GC.GetTotalAllocatedBytes(true);
        Stopwatch sw = Stopwatch.StartNew();
        double sink = 0;
        for (int i = 0; i < N; i++) sink += double.Parse(((double)i).ToString("G17", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        sw.Stop();
        Report("double.ToString(G17) x1M", sw.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(true) - before, sink);

        before = GC.GetTotalAllocatedBytes(true);
        sw = Stopwatch.StartNew();
        sink = 0;
        for (int i = 0; i < N; i++) sink += i.ToString(CultureInfo.InvariantCulture).Length;
        sw.Stop();
        Report("int.ToString x1M", sw.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(true) - before, sink);

        before = GC.GetTotalAllocatedBytes(true);
        sw = Stopwatch.StartNew();
        sink = 0;
        for (int i = 0; i < N; i++) sink += CellAddress.Format(i + 1, (i % 20) + 1).Length;
        sw.Stop();
        Report("CellAddress.Format x1M", sw.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(true) - before, sink);

        before = GC.GetTotalAllocatedBytes(true);
        sw = Stopwatch.StartNew();
        sink = 0;
        using (MemoryStream stream = new MemoryStream())
        {
            using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(stream, new System.Xml.XmlWriterSettings { Indent = false, Encoding = new System.Text.UTF8Encoding(false), CloseOutput = false }))
            {
                writer.WriteStartDocument(true);
                writer.WriteStartElement("root");
                for (int i = 0; i < N; i++)
                {
                    writer.WriteStartElement("c");
                    writer.WriteAttributeString("r", "A" + i.ToString(CultureInfo.InvariantCulture));
                    writer.WriteElementString("v", ((double)i).ToString("G17", CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
            sink = stream.Length;
        }
        sw.Stop();
        Report("XmlWriter 1M <c r><v>G17</v></c>", sw.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(true) - before, sink);
    }

    private static void Report(string label, double ms, long bytes, double sink)
    {
        Console.WriteLine("  [probe] " + label + " = " + ms.ToString("F1", CultureInfo.InvariantCulture) +
            " ms, " + (bytes / 1048576.0).ToString("F1", CultureInfo.InvariantCulture) + " MB, check=" + sink.ToString("F0", CultureInfo.InvariantCulture));
    }
}
