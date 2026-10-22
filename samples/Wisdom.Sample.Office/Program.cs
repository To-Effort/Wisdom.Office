using System;
using System.IO;
using Wisdom.Office;
using Wisdom.Office.Core;

// Wisdom.Office 统一门面演示：
// 1. Report 报表一套数据导出 PDF / DOCX / XLSX
// 2. OfficeFile 统一读取：格式检测、元数据、文本提取、另存
// 3. OfficeConverter 格式互转：DOCX->PDF、XLSX->PDF、PDF->DOCX
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
        Console.WriteLine("=== Wisdom.Sample.Office：统一门面演示 ===");

        // 1. 报表导出三种格式
        Report report = new Report("季度销售报表");
        report.Subtitle = "2026 年第三季度";
        report.Author = "Wisdom.Office";
        report.FooterText = "Wisdom.Office 自动生成";
        report.GeneratedAt = new DateTime(2026, 10, 20, 9, 30, 0);

        ReportTable table = report.AddTable("产品销售明细");
        table.AddColumn("产品", 120, null, HorizontalAlignment.Left);
        table.AddColumn("数量", 60, "0", HorizontalAlignment.Right);
        table.AddColumn("金额", 100, "#,##0.00", HorizontalAlignment.Right);
        table.AddColumn("日期", 90, "yyyy-MM-dd", HorizontalAlignment.Center);

        table.AddRow("无线键盘", 12, 3594.0, new DateTime(2026, 7, 5));
        table.AddRow("机械鼠标", 30, 2697.0, new DateTime(2026, 7, 6));
        table.AddRow("27 英寸显示器", 5, 9495.0, new DateTime(2026, 7, 7));
        report.AddNote("金额单位：人民币元。");

        string reportPdf = Path.Combine(OutDir, "report.pdf");
        string reportDocx = Path.Combine(OutDir, "report.docx");
        string reportXlsx = Path.Combine(OutDir, "report.xlsx");
        report.Export(reportPdf);
        report.Export(reportDocx);
        report.Export(reportXlsx);
        Console.WriteLine("[报表] pdf=" + new FileInfo(reportPdf).Length
            + "，docx=" + new FileInfo(reportDocx).Length
            + "，xlsx=" + new FileInfo(reportXlsx).Length + " 字节");

        // 2. 统一读取与格式检测
        OfficeFile file = OfficeFile.Load(reportXlsx);
        Console.WriteLine("[读取] 格式=" + file.Format
            + "，标题=" + file.Metadata.Title
            + "，文本预览=" + Truncate(file.ExtractText(), 40));

        string copyPath = Path.Combine(OutDir, "report-copy.xlsx");
        file.Save(copyPath);
        Console.WriteLine("[另存] " + new FileInfo(copyPath).Length + " 字节，检测=" + OfficeFileFormats.Detect(copyPath));

        // 3. 格式互转
        string docxPdf = Path.Combine(OutDir, "report.docx.pdf");
        OfficeConverter.Convert(reportDocx, docxPdf);
        Console.WriteLine("[互转] DOCX->PDF: " + new FileInfo(docxPdf).Length + " 字节");

        string xlsxPdf = Path.Combine(OutDir, "report.xlsx.pdf");
        OfficeConverter.Convert(reportXlsx, xlsxPdf, new WorkbookLayoutOptions
        {
            IncludeSheetNames = true,
            UseGridLines = true,
            FontFamily = "Microsoft YaHei",
            FontSize = 9
        });
        Console.WriteLine("[互转] XLSX->PDF: " + new FileInfo(xlsxPdf).Length + " 字节");

        string pdfDocx = Path.Combine(OutDir, "report.pdf.docx");
        OfficeConverter.Convert(reportPdf, pdfDocx);
        Console.WriteLine("[互转] PDF->DOCX: " + new FileInfo(pdfDocx).Length + " 字节");

        Console.WriteLine("Office 门面演示完成，输出目录: " + Path.GetFullPath(OutDir));
    }

    private static string Truncate(string text, int max)
    {
        text = (text ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
        return text.Length <= max ? text : text.Substring(0, max) + "...";
    }
}
