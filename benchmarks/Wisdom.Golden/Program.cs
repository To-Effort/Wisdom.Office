using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using Wisdom.Excel;
using Wisdom.Office.Core;

// 逐字节对照工具：用当前写出器生成覆盖各分支的 .xlsx，并把每个 OOXML 部件原文导出到目录。
// 修改 XlsxWriter 前后各跑一次，用 diff -r 比对部件字节，证明输出未变。
internal static class Program
{
    private static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("usage: Wisdom.Golden <output-directory>");
            return 1;
        }

        string directory = Path.GetFullPath(args[0]);
        if (Directory.Exists(directory)) Directory.Delete(directory, true);
        Directory.CreateDirectory(directory);

        byte[] package = XlsxWriter.Write(BuildMaximalWorkbook());
        using (MemoryStream input = new MemoryStream(package))
        using (ZipArchive archive = new ZipArchive(input, ZipArchiveMode.Read))
        {
            for (int i = 0; i < archive.Entries.Count; i++)
            {
                ZipArchiveEntry entry = archive.Entries[i];
                if (string.IsNullOrEmpty(entry.Name) && entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
                string target = Path.Combine(directory, entry.FullName.Replace('/', '_').Replace(":", "-") + ".bin");
                using (Stream source = entry.Open())
                using (FileStream output = new FileStream(target, FileMode.Create, FileAccess.Write))
                {
                    source.CopyTo(output);
                }
            }
            Console.WriteLine("parts=" + archive.Entries.Count + " package=" + package.Length + " bytes -> " + directory);
        }
        return 0;
    }

    private static OfficeWorkbook BuildMaximalWorkbook()
    {
        OfficeWorkbook workbook = new OfficeWorkbook();
        workbook.Metadata.Title = "黄金对照 & <tag>";
        workbook.Metadata.Author = "第一行\n第二行\t制表 & <转义>";
        workbook.Metadata.Subject = "XlsxWriter 输出基线";
        workbook.Metadata.Keywords = "a;b,c";
        workbook.Metadata.Created = new DateTime(2026, 3, 4, 5, 6, 7);
        workbook.Metadata.Modified = new DateTime(2026, 3, 5, 8, 9, 10);

        CellStyle bold = new CellStyle
        {
            Bold = true,
            Italic = true,
            Underline = true,
            FontFamily = "Microsoft YaHei",
            FontSize = 12,
            FontColor = new OfficeColor((byte)0xC0, (byte)0x00, (byte)0x00),
            FillColor = new OfficeColor((byte)0xFF, (byte)0xF2, (byte)0xCC),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Middle,
            WrapText = true,
            NumberFormat = "#,##0.00",
            TopBorder = BorderLine.Thin(),
            BottomBorder = new BorderLine(BorderStyle.Double, 1.5, new OfficeColor((byte)0x80, (byte)0x80, (byte)0x80)),
            LeftBorder = new BorderLine(BorderStyle.Dotted, 0.5, OfficeColor.Black),
            RightBorder = new BorderLine(BorderStyle.None, 0.5, OfficeColor.Black)
        };
        CellStyle date = new CellStyle { NumberFormat = "yyyy-MM-dd", Italic = true };
        CellStyle percent = new CellStyle { NumberFormat = "0.00%", HorizontalAlignment = HorizontalAlignment.Right };
        CellStyle textStyle = new CellStyle { FontSize = 9, TopBorder = new BorderLine(BorderStyle.Dashed, 2.0, OfficeColor.Black) };
        CellStyle onlyStyle = new CellStyle { Bold = true, FillColor = new OfficeColor((byte)10, (byte)20, (byte)30) };

        Worksheet sheet = workbook.AddWorksheet("明细 [非法]*?字符:名字");
        Worksheet second = workbook.AddWorksheet("明细 [非法]*?字符:名字");

        sheet.GetCell(1, 1).Value = "文本 & <标签> \"引号\" '单引'";
        sheet.GetCell(1, 1).Style = bold;
        sheet.GetCell(1, 2).Value = "换行\n与\t制表\r回车";
        sheet.GetCell(1, 3).Value = "Ünïcödé 中文 日本語 한국";
        sheet.GetCell(1, 4).Value = "   前后空格   ";
        sheet.GetCell(1, 5).Value = string.Empty;

        sheet.GetCell(2, 1).Value = 1234.5;
        sheet.GetCell(2, 1).Style = bold;
        sheet.GetCell(2, 2).Value = 0.2574;
        sheet.GetCell(2, 2).Style = percent;
        sheet.GetCell(2, 3).Value = new DateTime(2026, 10, 20, 0, 0, 0);
        sheet.GetCell(2, 3).Style = date;
        sheet.GetCell(2, 4).Value = true;
        sheet.GetCell(2, 5).Value = false;
        sheet.GetCell(2, 6).Value = -0.000001;
        sheet.GetCell(2, 7).Value = 1.0 / 0.0;
        sheet.GetCell(2, 8).Value = long.MaxValue;
        sheet.GetCell(2, 9).Value = 12345678901234567890m;

        sheet.GetCell(3, 1).Formula = "=SUM(A2:A3)";
        sheet.GetCell(3, 2).Formula = "B2*2";
        sheet.GetCell(3, 2).Value = 2469.0;
        sheet.GetCell(3, 3).Formula = "=A1&\"x\"";
        sheet.GetCell(3, 3).Value = "文本 & <标签> \"引号\" '单引'x";

        sheet.GetCell(4, 1).Style = onlyStyle;
        sheet.GetCell(4, 2).Value = "带样式无字体";
        sheet.GetCell(4, 2).Style = textStyle;

        sheet.Merge("A6:C6");
        sheet.GetCell(6, 1).Value = "合并起点";
        sheet.Merge("D6:F7");
        sheet.GetCell(6, 4).Value = "跨行跨列合并";

        sheet.FrozenRows = 1;
        sheet.FrozenColumns = 2;
        sheet.SetColumnWidth(1, 24.5);
        sheet.SetColumnWidth(2, 8);
        sheet.SetRowHeight(6, 30);
        sheet.SetRowHeight(7, 0.5);
        sheet.SetRowHidden(9, true);
        sheet.SetColumnHidden(11, true);
        sheet.ShowZeros = false;
        sheet.DefaultRowHeight = 13.25;
        sheet.DefaultColumnWidth = 9.5;

        for (int c = 1; c <= 14; c++)
        {
            sheet.GetCell(5, c).Value = c % 2 == 0 ? (object)(c * 1000.25) : ("列" + c.ToString());
            sheet.GetCell(5, c).Style = c % 3 == 0 ? bold : percent;
        }

        ExcelConditionalFormat bars = new ExcelConditionalFormat("B6:F6");
        bars.AddDataBar(new OfficeColor((byte)0x63, (byte)0x8E, (byte)0xC5));
        bars.AddCellIs("lessThan", "0", new CellStyle
        {
            FontColor = new OfficeColor((byte)0xC0, (byte)0, (byte)0),
            Bold = true
        });
        sheet.ConditionalFormats.Add(bars);

        ExcelConditionalFormat scales = new ExcelConditionalFormat("A5:F5");
        scales.AddColorScale(
            new OfficeColor((byte)0xF8, (byte)0x69, (byte)0x5A),
            new OfficeColor((byte)0xFF, (byte)0xFF, (byte)0xFF),
            new OfficeColor((byte)0x63, (byte)0xC5, (byte)0x8E));
        sheet.ConditionalFormats.Add(scales);

        second.GetCell(1, 1).Value = "第二个工作表";
        PageSetup setup = new PageSetup();
        setup.MarginLeft = 20;
        setup.MarginRight = 25;
        setup.MarginTop = 30;
        setup.MarginBottom = 35;
        setup.HeaderDistance = 0.6;
        setup.FooterDistance = 0.4;
        setup.Width = 792;
        setup.Height = 612;
        second.PrintSetup = setup;
        second.FitToPage = true;
        second.ShowGridLines = true;

        Worksheet third = workbook.AddWorksheet("Sheet3");
        third.IsHidden = true;
        third.GetCell(1, 1).Value = "隐藏表内容";

        return workbook;
    }
}
