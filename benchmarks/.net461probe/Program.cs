using System;
using System.IO;
using System.Linq;
using Wisdom.Excel;
using Wisdom.Office.Core;

internal static class Program
{
    private static int Main()
    {
        string path = Path.Combine(Path.GetTempPath(), "net461-probe.xlsx");
        try
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.Metadata.Title = "probe";
            Worksheet sheet = workbook.AddWorksheet("S1");
            CellStyle style = new CellStyle { Bold = true, NumberFormat = "#,##0.00" };
            for (int r = 1; r <= 50; r++)
            {
                Cell a = sheet.GetCell(r, 1);
                a.Value = "行" + r;
                Cell b = sheet.GetCell(r, 2);
                b.Value = r * 1.5;
                b.Style = style.Clone();
            }
            sheet.SetColumnWidth(1, 18.5);
            sheet.FrozenRows = 1;
            byte[] bytes = XlsxWriter.Write(workbook);
            File.WriteAllBytes(path, bytes);
            Console.WriteLine("written " + new FileInfo(path).Length);

            OfficeWorkbook read = XlsxReader.Load(path);
            Console.WriteLine("sheets=" + read.Worksheets.Count);
            Worksheet ws = read.Worksheets[0];
            int lastRow = ws.LastRow;
            int lastCol = ws.LastColumn;
            int cellCount = System.Linq.Enumerable.Count(ws.Cells);
            Console.WriteLine("lastRow=" + lastRow + " lastCol=" + lastCol + " cellCount=" + cellCount);
            Cell probeA = ws.FindCell(1, 1);
            Cell probeB = ws.FindCell(50, 2);
            Console.WriteLine("A1=" + (probeA == null ? "<null>" : probeA.DisplayText()));
            Console.WriteLine("B50=" + (probeB == null ? "<null>" : probeB.DisplayText()));
            Console.WriteLine("bold=" + (probeB == null ? "<null>" : probeB.Style.Bold.ToString()));
            Console.WriteLine("colwidths=" + string.Join(",", System.Linq.Enumerable.Select(ws.ColumnWidths, p => p.Key + ":" + p.Value)));
            Console.WriteLine("frozen=" + ws.FrozenRows);
            Walker.Run(path);
            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine("TYPE: " + ex.GetType().FullName);
            Console.WriteLine("MSG: " + SafeMessage(ex));
            Console.WriteLine("STACK: " + SafeTrace(ex));
            if (ex.InnerException != null)
            {
                Console.WriteLine("INNER TYPE: " + ex.InnerException.GetType().FullName);
                Console.WriteLine("INNER MSG: " + SafeMessage(ex.InnerException));
                Console.WriteLine("INNER STACK: " + SafeTrace(ex.InnerException));
            }
            return 1;
        }
    }

    private static string SafeMessage(Exception ex)
    {
        try { return ex.Message; } catch (Exception e) { return "<message failed: " + e.GetType().Name + ">"; }
    }

    private static string SafeTrace(Exception ex)
    {
        try { return ex.StackTrace; } catch (Exception e) { return "<trace failed: " + e.GetType().Name + ">"; }
    }
}
