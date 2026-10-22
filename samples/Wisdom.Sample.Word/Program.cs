using System;
using System.IO;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Word;

// Word（DOCX）读写演示：
// 1. 用统一模型构建文档 -> DocxWriter 写出 .docx
// 2. DocxReader 回读校验段落、表格
// 3. 用 OfficeFile 统一入口检测格式并提取纯文本
// 4. 通过 OfficeConverter 将 DOCX 转成 PDF
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
        Console.WriteLine("=== Wisdom.Sample.Word：DOCX 读写演示 ===");

        // 1. 构建统一文档模型并写入 DOCX
        OfficeDocument document = new OfficeDocument();
        document.Metadata.Title = "Word 读写示例";
        document.Metadata.Author = "Wisdom.Office";

        Section section = document.AddSection(PageSetup.A4());
        section.Blocks.Add(new Paragraph("项目验收报告").AddText(" —— 由 Wisdom.Word 生成", new TextFormat("Arial", 10)));
        section.Blocks.Add(new Paragraph("本报告演示段落、文本样式与表格的写入与回读。"));

        Table table = new Table();
        table.Columns.Add(new TableColumn(160));
        table.Columns.Add(new TableColumn(100));
        table.AddRow("项目", "金额");
        table.AddRow("软件许可", "12,000.00");
        table.AddRow("实施服务", "10,888.40");
        section.Blocks.Add(table);

        string docxPath = Path.Combine(OutDir, "sample-word.docx");
        DocxWriter.Write(document, docxPath);
        Console.WriteLine("[写入] DOCX: " + new FileInfo(docxPath).Length + " 字节");

        // 2. 回读校验
        OfficeDocument loaded = DocxReader.Load(docxPath);
        int paragraphs = 0, tables = 0;
        foreach (Section s in loaded.Sections)
        {
            foreach (Block block in s.Blocks)
            {
                if (block is Paragraph) paragraphs++;
                if (block is Table) tables++;
            }
        }
        Console.WriteLine("[回读] 标题=" + loaded.Metadata.Title + "，段落数=" + paragraphs + "，表格数=" + tables);

        // 3. 统一入口：格式检测 + 纯文本提取
        OfficeFile file = OfficeFile.Load(docxPath);
        Console.WriteLine("[检测] 格式=" + file.Format + "，文本预览=" + Truncate(file.ExtractText(), 40));

        // 4. DOCX -> PDF
        string pdfPath = Path.Combine(OutDir, "sample-word.pdf");
        OfficeConverter.Convert(docxPath, pdfPath);
        Console.WriteLine("[转换] PDF: " + new FileInfo(pdfPath).Length + " 字节");

        Console.WriteLine("Word 演示完成，输出目录: " + Path.GetFullPath(OutDir));
    }

    private static string Truncate(string text, int max)
    {
        text = (text ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Trim();
        return text.Length <= max ? text : text.Substring(0, max) + "...";
    }
}
