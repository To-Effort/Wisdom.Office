using System;
using System.IO;
using Wisdom.Excel;
using Wisdom.Office;
using Wisdom.Office.Core;

// Excel（XLSX）读写演示：
// 1. 构建工作簿（样式、数字格式、合并、冻结、列宽）-> XlsxWriter 写出 .xlsx
// 2. XlsxReader 回读校验
// 3. OfficeConverter 将 XLSX 直接转成 PDF
// 4. 看板元素：原生图表、条件格式数据条、锚定单元格图片（写出后回读校验）
internal static class Program
{
    // 一张 1x1 像素的 PNG 常量字节，用于演示锚定到单元格的图标。
    private static readonly byte[] OnePixelPng =
    {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4, 0x89,
        0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4,
        0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82
    };
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
        Console.WriteLine("=== Wisdom.Sample.Excel：XLSX 读写演示 ===");

        // 1. 写入 XLSX
        OfficeWorkbook workbook = new OfficeWorkbook();
        workbook.Metadata.Title = "销售数据";

        Worksheet sheet = workbook.AddWorksheet("明细");
        CellStyle header = new CellStyle { Bold = true, HorizontalAlignment = HorizontalAlignment.Center, FillColor = new OfficeColor((byte)217, (byte)226, (byte)243) };

        Cell h1 = sheet.GetCell(1, 1); h1.Value = "产品"; h1.Style = header;
        Cell h2 = sheet.GetCell(1, 2); h2.Value = "数量"; h2.Style = header;
        Cell h3 = sheet.GetCell(1, 3); h3.Value = "金额"; h3.Style = header;

        sheet.GetCell(2, 1).Value = "无线键盘";
        sheet.GetCell(2, 2).Value = 12;
        Cell amount = sheet.GetCell(2, 3);
        amount.Value = 3594.0;
        amount.Style = new CellStyle { NumberFormat = "#,##0.00", Bold = true };

        sheet.GetCell(3, 1).Value = "27 英寸显示器";
        sheet.GetCell(3, 2).Value = 5;
        Cell amount2 = sheet.GetCell(3, 3);
        amount2.Value = 9495.0;
        amount2.Style = new CellStyle { NumberFormat = "#,##0.00", Bold = true };

        Cell dateCell = sheet.GetCell(4, 1);
        dateCell.Value = new DateTime(2026, 7, 5);
        dateCell.Style = new CellStyle { NumberFormat = "yyyy-MM-dd" };

        sheet.Merge("A6:C6");
        sheet.GetCell(6, 1).Value = "备注：本季度价格含税。";
        sheet.FrozenRows = 1;
        sheet.SetColumnWidth(1, 24);

        string xlsxPath = Path.Combine(OutDir, "sample-excel.xlsx");
        XlsxWriter.Write(workbook, xlsxPath);
        Console.WriteLine("[写入] XLSX: " + new FileInfo(xlsxPath).Length + " 字节");

        // 2. 回读校验
        OfficeWorkbook loaded = XlsxReader.Load(xlsxPath);
        Worksheet loadedSheet = loaded.Worksheets[0];
        Console.WriteLine("[回读] 工作表=" + loadedSheet.Name
            + "，B2=" + loadedSheet.GetCell(2, 2).Value
            + "，C2=" + loadedSheet.GetCell(2, 3).Value
            + "，合并区域数=" + loadedSheet.MergedRanges.Count
            + "，冻结行=" + loadedSheet.FrozenRows);

        // 3. XLSX -> PDF
        string pdfPath = Path.Combine(OutDir, "sample-excel.pdf");
        OfficeConverter.Convert(xlsxPath, pdfPath, new WorkbookLayoutOptions
        {
            IncludeSheetNames = true,
            UseGridLines = true,
            FontFamily = "Microsoft YaHei",
            FontSize = 9
        });
        Console.WriteLine("[转换] PDF: " + new FileInfo(pdfPath).Length + " 字节");

        // 4. 看板元素：原生图表 + 条件格式数据条 + 锚定单元格图片
        OfficeWorkbook dashboard = new OfficeWorkbook();
        Worksheet board = dashboard.AddWorksheet("看板");
        board.SetColumnWidth(1, 16);
        for (int column = 2; column <= 7; column++) board.SetColumnWidth(column, 10);

        CellStyle boardHeader = new CellStyle { Bold = true, HorizontalAlignment = HorizontalAlignment.Center, FillColor = new OfficeColor((byte)0x00, (byte)0xB0, (byte)0x50), FontColor = OfficeColor.White };
        Cell region = board.GetCell(1, 1); region.Value = "区域"; region.Style = boardHeader;
        string[] months = { "1月", "2月", "3月", "4月", "5月", "6月" };
        for (int i = 0; i < months.Length; i++)
        {
            Cell cell = board.GetCell(1, i + 2);
            cell.Value = months[i];
            cell.Style = boardHeader;
        }
        board.GetCell(2, 1).Value = "销售额";
        double[] sales = { 820, 960, 735, 1040, 1180, 990 };
        for (int i = 0; i < sales.Length; i++)
        {
            Cell cell = board.GetCell(2, i + 2);
            cell.Value = sales[i];
            cell.Style = new CellStyle { NumberFormat = "#,##0" };
        }
        // 合计列故意只给 4 个字符宽度，回读/转 Word、PDF 时会显示为 #### 溢出提示。
        board.SetColumnWidth(8, 4);
        Cell total = board.GetCell(2, 8);
        total.Value = 5725.0;
        total.Style = new CellStyle { NumberFormat = "#,##0.00", Bold = true };

        ExcelChart salesChart = new ExcelChart();
        salesChart.Type = ExcelChartType.Column;
        salesChart.Title = "上半年销售额";
        salesChart.Row = 4;
        salesChart.Column = 1;
        salesChart.WidthPoints = 360;
        salesChart.HeightPoints = 200;
        salesChart.ValueAxisFormat = "#,##0";
        ExcelChartSeries salesSeries = salesChart.AddSeries("看板!$A$2", "看板!$B$1:$G$1", "看板!$B$2:$G$2");
        salesSeries.FillColor = new OfficeColor((byte)0x00, (byte)0xB0, (byte)0x50);
        salesSeries.ShowDataLabels = true;
        board.Charts.Add(salesChart);

        ExcelConditionalFormat salesBar = new ExcelConditionalFormat("B2:G2");
        salesBar.AddDataBar(new OfficeColor((byte)0x63, (byte)0xC3, (byte)0x84));
        board.ConditionalFormats.Add(salesBar);

        ExcelImage flag = new ExcelImage(new OfficeImage(OnePixelPng));
        flag.Row = 2;
        flag.Column = 9;
        flag.WidthPoints = 24;
        flag.HeightPoints = 24;
        board.Images.Add(flag);

        string dashboardPath = Path.Combine(OutDir, "sample-dashboard.xlsx");
        XlsxWriter.Write(dashboard, dashboardPath);
        OfficeWorkbook reloaded = XlsxReader.Load(dashboardPath);
        Worksheet reloadedBoard = reloaded.Worksheets[0];
        Console.WriteLine("[看板] XLSX: " + new FileInfo(dashboardPath).Length + " 字节"
            + "，回读图表=" + reloadedBoard.Charts.Count
            + "，条件格式=" + reloadedBoard.ConditionalFormats.Count
            + "，图片=" + reloadedBoard.Images.Count
            + "，图表首系列=" + reloadedBoard.Charts[0].Series[0].ValueRange);

        Console.WriteLine("Excel 演示完成，输出目录: " + Path.GetFullPath(OutDir));
    }
}
