using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using Wisdom.Excel;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;

internal static class Program
{
    private static void Main(string[] args)
    {
        int rows = args.Length > 0 ? int.Parse(args[0], CultureInfo.InvariantCulture) : 10000;
        string variant = args.Length > 1 ? args[1] : "full";
        int rounds = args.Length > 2 ? int.Parse(args[2], CultureInfo.InvariantCulture) : 3;

        if (variant == "two") { rows = 10000; }

        OfficeWorkbook workbook = BuildWorkbook(rows, 20);
        long beforeConvert = GC.GetTotalAllocatedBytes(true);
        Stopwatch convertWatch = Stopwatch.StartNew();
        OfficeDocument document = WorkbookToDocumentConverter.Convert(workbook);
        convertWatch.Stop();
        long convertAlloc = GC.GetTotalAllocatedBytes(true) - beforeConvert;
        Console.WriteLine($"convert: {convertWatch.Elapsed.TotalMilliseconds:F0} ms, alloc {convertAlloc / 1048576.0:F0} MB");

        switch (variant)
        {
            case "noborder":
                RemoveBorders(document);
                break;
            case "notext":
                ClearText(document);
                break;
            case "textonly":
                RemoveBorders(document);
                break;
        }

        // Warm the static font face cache with a tiny document.
        PdfRenderer warm = new PdfRenderer();
        warm.RenderToBytes(WorkbookToDocumentConverter.Convert(BuildWorkbook(20, 5)));

        double[] times = new double[rounds];
        long[] allocs = new long[rounds];
        int pages = 0;
        long bytes = 0;
        for (int i = 0; i < rounds; i++)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            long before = GC.GetTotalAllocatedBytes(true);
            Stopwatch watch = Stopwatch.StartNew();
            byte[] pdf = new PdfRenderer().RenderToBytes(document);
            watch.Stop();
            allocs[i] = GC.GetTotalAllocatedBytes(true) - before;
            times[i] = watch.Elapsed.TotalMilliseconds;
            bytes = pdf.Length;
            if (i == 0) pages = PdfDocument.Load(pdf).PageCount;
            File.WriteAllBytes(Path.Combine(Path.GetTempPath(), "pdfcell.pdf"), pdf);
        }

        Array.Sort(times);
        Array.Sort(allocs);
        double median = times[times.Length / 2];
        Console.WriteLine($"{variant} {rows}x20: pages={pages} bytes={bytes} ms(min/med)={times[0]:F0}/{median:F0} alloc med={allocs[allocs.Length / 2] / 1048576.0:F0} MB peak={Environment.WorkingSet / 1048576.0:F0} MB");
        Console.WriteLine($"  cells/s={rows * 20.0 / median * 1000:F0}  ms/cell={median / (rows * 20) * 1000:F3}");
    }

    private static void RemoveBorders(OfficeDocument document)
    {
        foreach (Section section in document.Sections)
        {
            foreach (Table table in section.Blocks.OfType<Table>())
            {
                table.Borders = TableBorders.None();
                foreach (TableRow row in table.Rows)
                {
                    foreach (TableCell cell in row.Cells) cell.Borders = null;
                }
            }
        }
    }

    private static void ClearText(OfficeDocument document)
    {
        foreach (Section section in document.Sections)
        {
            foreach (Table table in section.Blocks.OfType<Table>())
            {
                foreach (TableRow row in table.Rows)
                {
                    foreach (TableCell cell in row.Cells)
                    {
                        foreach (Paragraph paragraph in cell.Blocks.OfType<Paragraph>()) paragraph.Inlines.Clear();
                    }
                }
            }
        }
    }

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
}
