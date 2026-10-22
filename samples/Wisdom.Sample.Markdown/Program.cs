using System;
using System.IO;
using System.Text;
using Wisdom.Markdown;
using Wisdom.Office;
using Wisdom.Office.Core;

// Markdown 演示：
// 1. Markdown 源文本 -> 统一文档模型（标题、列表、表格、代码块、引用、粗斜体、链接）
// 2. 文档模型 -> HTML 片段
// 3. 文档模型 -> Markdown 文本（往返）
// 4. .md 文件经 OfficeFile 统一入口转出 PDF 与 DOCX
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

    private const string Source = @"---
title: 报表引擎使用说明
author: Wisdom.Office
---

# 报表引擎使用说明

** Wisdom.Excel ** 与 **Wisdom.Pdf** 负责把数据表排成可打印的版式。

## 支持的语法

- 粗体、*斜体*、~~删除线~~、`行内代码`
- [指向文档的链接](https://example.com/doc)
  - 两级缩进的列表项

1. 有序列表第一项
2. 有序列表第二项

- [x] 已完成：图表与条件格式
- [ ] 待办：数据透视

> 引用块：导出前请先确认纸张尺寸。
> 第二行引用。

| 区域 | 收入 | 占比 |
|:-----|-----:|:----:|
| 华东 | 588 | 46% |
| 华南 | 384 | 30% |

```csharp
OfficeFile file = OfficeFile.Load(""report.xlsx"");
file.Save(""report.pdf"", OfficeFileFormat.Pdf);
```

---

结尾段落里列宽不足会输出 `########`。
";

    private static void Main()
    {
        Console.WriteLine("=== Wisdom.Sample.Markdown：Markdown 解析与互转演示 ===");

        // 1. Markdown -> 文档模型
        MarkdownOptions options = new MarkdownOptions();
        options.EnableTables = true;
        options.EnableTaskLists = true;
        OfficeDocument document = Wisdom.Markdown.Markdown.ToDocument(Source, options);
        int paragraphs = 0;
        int tables = 0;
        int headings = 0;
        for (int s = 0; s < document.Sections.Count; s++)
        {
            for (int i = 0; i < document.Sections[s].Blocks.Count; i++)
            {
                Paragraph paragraph = document.Sections[s].Blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    paragraphs++;
                    if (paragraph.HeadingLevel > 0) headings++;
                    continue;
                }
                if (document.Sections[s].Blocks[i] is Table) tables++;
            }
        }
        Console.WriteLine("[解析] 块=" + paragraphs + " 段落/标题，标题=" + headings
            + "，表格=" + tables + " 张，元数据标题=\"" + document.Metadata.Title + "\"");

        // 2. 文档模型 -> HTML
        string html = Wisdom.Markdown.Markdown.ToHtml(Source, options);
        string htmlPath = Path.Combine(OutDir, "sample-markdown.html");
        File.WriteAllText(htmlPath, WrapHtml(html), new UTF8Encoding(false));
        Console.WriteLine("[HTML] 长度=" + html.Length + "，文件=" + new FileInfo(htmlPath).Length + " 字节");
        Console.WriteLine("  片段: " + FirstLines(html, 3));

        // 3. 文档模型 -> Markdown（往返）
        string back = Wisdom.Markdown.Markdown.ToMarkdown(document);
        string mdPath = Path.Combine(OutDir, "sample-roundtrip.md");
        File.WriteAllText(mdPath, back, new UTF8Encoding(false));
        OfficeDocument reparsed = Wisdom.Markdown.Markdown.ToDocument(back, options);
        Console.WriteLine("[往返] Markdown 长度=" + back.Length
            + "，再解析标题=\"" + reparsed.Metadata.Title + "\"");

        // 4. 经统一入口把 .md 转为 PDF 与 DOCX
        string sourceMdPath = Path.Combine(OutDir, "sample-source.md");
        File.WriteAllText(sourceMdPath, Source, new UTF8Encoding(false));
        OfficeFile file = OfficeFile.Load(sourceMdPath);
        Console.WriteLine("[识别] " + OfficeFileFormats.GetDisplayName(file.Format)
            + "，纯文本长度=" + file.ExtractText().Length);

        string pdfPath = Path.Combine(OutDir, "sample-markdown.pdf");
        file.Save(pdfPath);
        Console.WriteLine("[转换] PDF: " + new FileInfo(pdfPath).Length + " 字节");

        string docxPath = Path.Combine(OutDir, "sample-markdown.docx");
        OfficeConverter.Convert(file, docxPath);
        Console.WriteLine("[转换] DOCX: " + new FileInfo(docxPath).Length + " 字节");

        Console.WriteLine("Markdown 演示完成，输出目录: " + Path.GetFullPath(OutDir));
    }

    private static string FirstLines(string text, int count)
    {
        string[] lines = text.Replace("\r\n", "\n").Split('\n');
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < lines.Length && i < count; i++)
        {
            if (i > 0) builder.Append(' ');
            builder.Append(lines[i].Trim());
        }
        return builder.ToString();
    }

    private static string WrapHtml(string body)
    {
        return "<!DOCTYPE html>\n<html lang=\"zh-CN\">\n<head>\n<meta charset=\"utf-8\" />\n"
            + "<title>报表引擎使用说明</title>\n</head>\n<body>\n" + body + "</body>\n</html>\n";
    }
}
