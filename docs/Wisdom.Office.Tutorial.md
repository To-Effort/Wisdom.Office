# Wisdom.Office 使用教程

门面包的实战教程：从一份数据出发产出 PDF / Word / Excel 三种报表，把已有 `.xlsx` 排成 PDF，做各格式互转，并在不落盘的服务场景里直接拿字节。

阅读前请先安装 `Wisdom.Office`（安装方式见[包说明](../src/Wisdom.Office/README.md#安装)）。本教程示例统一假定 .NET 8 控制台工程，需要的命名空间在每个步骤顶部标出。

## 你将完成

```text
out/
  sales.pdf      由 Report 生成，含标题、表头底纹、斑马纹、备注
  sales.docx     同一份数据的 Word 版
  sales.xlsx     同一份数据的 Excel 版（带冻结窗格）
  detail.pdf     已有工作簿按排版选项转出的 PDF
  legacy.docx    旧版 .doc 转换结果
  summary.docx   PDF 反向转成可编辑 Word
  deck.pptx      用演示文稿模型搭出的幻灯片
  deck.pdf       同一份幻灯片按"每页一节"导出的 PDF
  article.pptx   由 article.md 直接生成的幻灯片
  article.md     任意文档回写的 Markdown
```

正文里的 `t.md` 是 Markdown 源文件，内容在第 8 节开头给出；其余输入都来自前面几步的产物。

## 1. 准备数据

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using Wisdom.Office;
using Wisdom.Office.Core;

record Sale(string Product, int Qty, decimal Amount, DateTime Date);

List<Sale> sales = new List<Sale>
{
    new Sale("无线键盘", 12, 3594.00m, new DateTime(2026, 7, 5)),
    new Sale("机械鼠标", 30, 2697.00m, new DateTime(2026, 7, 6)),
    new Sale("27 英寸显示器", 5, 9495.00m, new DateTime(2026, 7, 7)),
};

Directory.CreateDirectory("out");
```

## 2. 用 `Report` 一次产出三种格式

`Report` 是构建在统一模型上的报表 DSL：描述数据与列格式，导出时由对应引擎处理标题、表头、斑马纹、备注与冻结窗格。

```csharp
Report report = new Report("季度销售报表");
report.Subtitle = "2026 年第三季度";
report.Author = "示例公司";
report.FooterText = "由 Wisdom.Office 自动生成";
report.GeneratedAt = new DateTime(2026, 10, 21, 9, 30, 0);
report.ShowGeneratedTime = true;

ReportTable table = report.AddTable("产品销售明细");
table.AddColumn("产品", 120, null, HorizontalAlignment.Left);
table.AddColumn("数量", 60, "0", HorizontalAlignment.Right);
table.AddColumn("金额", 100, "#,##0.00", HorizontalAlignment.Right);
table.AddColumn("日期", 90, "yyyy-MM-dd", HorizontalAlignment.Center);

foreach (Sale sale in sales)
{
    table.AddRow(sale.Product, sale.Qty, sale.Amount, sale.Date);
}

report.AddNote("金额单位：人民币元，均为含税价。");
```

需要改外观时，在导出之前配置 `report.Style`（斑马纹、表头底色、字体都会同时作用于三种格式）：

```csharp
report.Style.FontFamily = "Microsoft YaHei";
report.Style.BaseFontSize = 9;
report.Style.HeaderFill = OfficeColor.FromHtml("#1F4E79");
report.Style.HeaderTextColor = OfficeColor.White;
report.Style.AlternateRowFill = OfficeColor.FromHtml("#F2F7FC");
report.Style.ZebraStripes = true;

report.Export("out/sales.pdf");
report.Export("out/sales.docx");
report.Export("out/sales.xlsx");
```

`Export` 按扩展名分派；想要中间模型继续手工加工，用 `ToDocument()`（Word/PDF 侧）或 `ToWorkbook()`（Excel 侧）：

```csharp
OfficeDocument document = report.ToDocument();
document.Sections[0].Blocks.Add(new PageBreak());     // 备注另起一页

OfficeWorkbook workbook = report.ToWorkbook();
workbook.Worksheets[0].SetColumnWidth(1, 28);         // 报表默认列宽可再调整
```

## 3. 读任意文件：不用先知道格式

`OfficeFile.Load` 按文件内容签名判断格式，扩展名错了也能读。

```csharp
OfficeFile file = OfficeFile.Load("out/sales.xlsx");

Console.WriteLine(file.Format);                       // Xlsx
Console.WriteLine(file.Metadata.Title);               // 季度销售报表
Console.WriteLine(file.Document == null);             // True —— Excel 还没转换成文档模型

string text = file.ExtractText();
Console.WriteLine(text.Substring(0, Math.Min(60, text.Length)));  // [Report]（工作表名 + 制表符分隔的行列）

file.Save("out/sales-copy.xlsx");                     // 同格式回存
byte[] pdfBytes = file.ToArray(OfficeFileFormat.Pdf); // 跨格式取字节
```

实测：`ExtractText()` 188 字符、`sales-copy.xlsx` 4,666 字节、`ToArray(Pdf)` 64,833 字节、`ToArray(Pptx)` 8,223 字节、`ToArray(Markdown)` 465 字节。

按格式取底层模型：

| 属性 / 方法 | 适用格式 | 得到 |
| --- | --- | --- |
| `file.Workbook` | Xlsx、Xls | `OfficeWorkbook` |
| `file.Document` | Docx、Doc、Pdf、Markdown | `OfficeDocument`（PDF 与 Markdown 在加载时已经解析成文档模型） |
| `file.Pdf` | Pdf | `PdfDocument` |
| `file.Presentation` | Pptx | `OfficePresentation` |
| `file.AsDocument()` | Xlsx、Xls、Pptx | 按默认排版映射出的 `OfficeDocument`（结果会缓存） |
| `file.AsDocument(options)` | Xlsx、Xls | 按指定排版选项映射出的 `OfficeDocument`（传 options 时不缓存） |
| `file.AsPresentation()` | Docx、Doc、Xlsx、Xls、Pdf、Markdown | 按标题层级拆页得到的 `OfficePresentation`（结果会缓存） |

四个属性是"原始模型"，只反映加载时拿到的内容：Excel 文件刚加载时 `Document` 与 `Presentation` 都是 `null`，调用过 `AsDocument()`（或任何触发排版的方法如 `ToArray(Pdf)`）之后 `Document` 才不为 null。跨格式取值一律走 `AsDocument()` / `AsPresentation()`。

只要判断格式、不需要内容时用 `OfficeFileFormats.Detect(path)`，它不会构造完整模型。

## 4. Excel 转 PDF 并控制排版

工作簿本身没有"页"的概念，`WorkbookLayoutOptions` 决定工作表如何映射为文档：是否输出表名、是否用网格线、是否沿用打印设置与打印区域、列宽不足时是否模拟 `####`。

```csharp
OfficeConverter.Convert("out/sales.xlsx", "out/detail.pdf", new WorkbookLayoutOptions
{
    IncludeSheetNames = true,        // 每个工作表前加一个表名标题
    UsePrintSetup = true,            // 沿用纸张、方向、页边距与打印区域
    UseGridLines = true,             // 用网格线代替逐格边框
    EmulateNumberOverflow = true,    // 列宽不够时按 Excel 规则输出 ####
    FontFamily = "Microsoft YaHei",
    FontSize = 9,
    GridColor = OfficeColor.LightGray
});
```

横向大表的推荐组合：在源工作簿上把该表标成横向并框定打印区域，转换时沿用这些设置。

```csharp
using Wisdom.Excel;

OfficeWorkbook data = XlsxReader.Load("out/sales.xlsx");

Worksheet sheet = data.Worksheets[0];
sheet.PrintSetup.IsLandscape = true;
sheet.PrintAreaFirstRow = 1;
sheet.PrintAreaLastRow = 60;
sheet.PrintAreaFirstColumn = 1;
sheet.PrintAreaLastColumn = 16;

OfficeConverter.WorkbookToPdf(data, "out/wide.pdf", new WorkbookLayoutOptions
{
    UsePrintSetup = true,
    IncludeSheetNames = true,
    IncludeHiddenWorksheets = false
});
```

不想经过文件时，`WorkbookToDocumentConverter.Convert(workbook, options)` 直接给出映射后的 `OfficeDocument`，可继续手工调整再交给 `OfficeConverter.ToPdf`。

## 5. 格式互转矩阵

下面的文件名是示意，换成你自己的输入输出即可。

```csharp
// DOCX / DOC -> PDF
OfficeConverter.Convert("draft.docx", "draft.pdf");

// DOC -> DOCX（旧版只读，先转成 OOXML 再处理）
OfficeConverter.Convert("legacy.doc", "out/legacy.docx");

// XLS -> XLSX
OfficeConverter.Convert("legacy.xls", "legacy.xlsx");

// PDF -> DOCX（解析文本与表格后按统一模型重新排版）
OfficeConverter.Convert("scan-output.pdf", "out/summary.docx");

// 内存中转换：入参是源字节 + 目标格式
byte[] docx = OfficeConverter.Convert(File.ReadAllBytes("draft.docx"), OfficeFileFormat.Docx);
```

实测的转换矩阵（同一份 `sales.*` 产物 + 一份 Markdown、一份 PPTX，目标由扩展名决定；格子里是输出字节数）：

| 源 \ 目标 | PDF | DOCX | XLSX | PPTX | Markdown |
| --- | --- | --- | --- | --- | --- |
| Markdown | 56,889 | 3,964 | 不支持 | 7,714 | 206 |
| DOCX | 45,450 | 5,205 | 不支持 | 8,064 | 404 |
| PDF | 51,036 | 4,016 | 不支持 | 7,964 | 441 |
| XLSX | 64,833 | 4,388 | 4,666 | 8,223 | 465 |
| PPTX | 50,558 | 3,843 | 不支持 | 10,861 | 161 |

只有"非电子表格 → XLSX"会抛 `NotSupportedException`（消息：`Word/PDF 文档不能直接保存为 Excel 工作簿。`），其余 24 条路径全部通过。`.doc`、`.xls` 只读，但作为源可用。同一份内容两次跑，DOCX / PPTX / PDF 的字节数会差 1～2 字节——包里记了生成时间戳，所以回归断言要比内容与页/行/列数量，别比哈希（见文末常见问题）。

格式判断的三个入口也都实测过：

```csharp
Console.WriteLine(OfficeFileFormats.Detect("t.md"));                      // Unknown —— Markdown 没有内容签名
Console.WriteLine(OfficeFile.Load("t.md").Format);                        // Markdown —— 按扩展名兜底
Console.WriteLine(OfficeFileFormats.GetExtension(OfficeFileFormat.Pptx)); // .pptx
Console.WriteLine(OfficeFileFormats.GetDisplayName(OfficeFileFormat.Markdown)); // Markdown 文档

// FromExtension 覆盖全部受支持扩展名
// .md/.markdown/.mkd -> Markdown，.pptx -> Pptx，.doc -> Doc，.xls -> Xls
```

从字节或流加载时没有扩展名可兜底，Markdown 必须显式指定格式：

```csharp
byte[] md = File.ReadAllBytes("t.md");
Console.WriteLine(OfficeFile.Load(md, OfficeFileFormat.Markdown).Format);  // Markdown
OfficeFile.Load(md);          // 抛 NotSupportedException: 无法识别文件格式。支持 PDF、DOCX、DOC、XLSX、XLS、PPTX 和 Markdown。
using (Stream s = File.OpenRead("t.md"))
{
    OfficeFile.Load(s);       // 同样抛 NotSupportedException —— Load(Stream) 没有格式重载
}
```

## 6. 服务场景：全程字节，不碰文件系统

不需要排版选项时，字节进字节出只需一行；需要 `WorkbookLayoutOptions` 时先按字节读成 `OfficeWorkbook`，再走 `WorkbookToPdf` 的字节重载：

```csharp
using Wisdom.Excel;
using Wisdom.Office;
using Wisdom.Office.Core;

// 单向、无需排版选项
public static byte[] ToPdfBytes(byte[] sourceDocx)
{
    return OfficeConverter.Convert(sourceDocx, OfficeFileFormat.Pdf);
}

// Excel -> PDF 且要控制排版
public static byte[] BuildPdf(byte[] sourceXlsx)
{
    OfficeWorkbook workbook = XlsxReader.Load(sourceXlsx);
    return OfficeConverter.WorkbookToPdf(workbook, new WorkbookLayoutOptions
    {
        IncludeSheetNames = true,
        UseGridLines = true
    });
}
```

`OfficeFile.Load` 有 `(string)`、`(Stream)`、`(byte[])`、`(byte[], OfficeFileFormat)` 四个入口；`OfficeConverter.Convert` 的字节重载只接受 `(byte[], OfficeFileFormat)`，返回目标格式字节，适合直接写进 HTTP 响应。已知源格式时可显式传入 `OfficeFile.Load(bytes, format)` 跳过内容嗅探。

## 7. PPTX：模型直接进门面

演示文稿模型定义在 `Wisdom.Office.Core`，写出与回读由 `Wisdom.Ppt` 完成，门面包把它们串起来。几何量全部是磅，原点为幻灯片左上角。

```csharp
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Ppt;

OfficePresentation presentation = OfficePresentation.Widescreen();   // 960 x 540
presentation.Metadata.Title = "季度经营回顾";

Slide cover = presentation.AddSlide(SlideLayoutKind.Title);
((TextBoxShape)cover.Shapes[0]).Paragraphs[0].AddText("2026 年第三季度");
cover.NotesText = "开场 2 分钟。";
cover.Transition = SlideTransitionKind.Fade;
cover.TransitionDuration = 1.5;

Slide body = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
((TextBoxShape)body.Shapes[0]).Paragraphs[0].AddText("业务进展");
TextBoxShape content = (TextBoxShape)body.Shapes[1];
content.Paragraphs.Clear();
content.AddParagraph("完成报表引擎");
content.AddParagraph("支持 Markdown 解析");
TableShape table = body.AddTable(480, 90, 300, 90);
table.Table.AddRow("模块", "状态");
table.Table.AddRow("PPT", "已完成");

OfficeConverter.ToPptx(presentation, "out/deck.pptx");              // 10,843 字节
```

读回来即可拿到模型，`OfficeFile` 对 PPTX 只填 `Presentation`，其余属性为 `null`：

```csharp
OfficeFile deck = OfficeFile.Load("out/deck.pptx");
Console.WriteLine(deck.Format);                                     // Pptx
Console.WriteLine(deck.Presentation.Slides.Count);                  // 2
Console.WriteLine(deck.Presentation.Slides[0].GetTitle());           // 2026 年第三季度
Console.WriteLine(deck.ExtractText().Split('\n')[0].Trim());         // 第 1 页：2026 年第三季度（全文 71 字符）
Console.WriteLine(deck.Presentation.Slides[0].Transition);           // Fade，时长 1.5
Console.WriteLine(deck.Presentation.Slides[0].NotesText);            // 开场 2 分钟。
Console.WriteLine(deck.Presentation.Slides[1].Shapes.Count);         // 3（标题 + 正文 + 表格）
Console.WriteLine(deck.Presentation.Slides[1].Layout);               // TitleAndTable —— 版式按形状推断，写入的是 TitleAndContent
```

导出其他格式走"每页幻灯片一节"，`AsDocument()` 的结果可以直接继续手工排版：

```csharp
Console.WriteLine(deck.AsDocument().Sections.Count);                 // 2 —— 一页一节
Console.WriteLine(deck.ToArray(OfficeFileFormat.Pdf).Length);        // 50,558
Console.WriteLine(deck.ToArray(OfficeFileFormat.Markdown).Length);   // 161

OfficeConverter.PresentationToPdf(presentation, "out/deck.pdf");      // 48,259 字节
OfficeConverter.PresentationToWord(presentation, "out/deck.docx");    // 3,768 字节
```

同一个模型直接导出 48,259 字节，从 `.pptx` 回读后再导出 50,558 字节。差异来自 `Table` 的列定义：写出时不落库，回读时按每行单元格数补齐（2x0 → 2x2），排版阶段据此才画出完整表格。

不想经过文件时，用内存模型建入口对象（初始格式即 Pptx）：

```csharp
OfficeFile fromModel = OfficeFile.FromPresentation(presentation);
Console.WriteLine(fromModel.Format);                                  // Pptx
fromModel.Save("out/deck-b.pptx");                                     // 10,843 字节
fromModel.Save("out/deck-c.pdf", OfficeFileFormat.Pdf);                // 48,259 字节
fromModel.Save("out/deck-c.docx", OfficeFileFormat.Docx);              // 3,768 字节
fromModel.Save("out/deck-c.md", OfficeFileFormat.Markdown);            // 161 字节
```

反向也行：任何文档按标题层级拆成幻灯片，`HeadingLevel` 越大越靠后。

```csharp
OfficeDocument article = Wisdom.Markdown.Markdown.ToDocument(File.ReadAllText("t.md"));
OfficeConverter.DocumentToPptx(article, "out/article.pptx");            // 7,714 字节
Console.WriteLine(PptxReader.Load("out/article.pptx").Slides.Count);   // 1 —— 只有一处一级标题

OfficeFile docx = OfficeFile.Load("out/sales.docx");
Console.WriteLine(docx.AsPresentation().Slides.Count);                 // 1 —— 报表只有一级标题
```

## 8. Markdown：没有签名的第五种格式

```csharp
using Wisdom.Markdown;

string markdown = "# 季度经营回顾\n\n营收增长 **18%**，成本持平。详见[官网](https://example.com)。\n\n"
    + "- 完成报表引擎\n- [x] 支持 Markdown 解析\n\n| 模块 | 状态 |\n| --- | --- |\n| PPT | 已完成 |\n";

OfficeDocument document = Markdown.ToDocument(markdown);
Console.WriteLine(document.Sections[0].Blocks.Count);                    // 5
Console.WriteLine(document.Sections[0].Blocks.OfType<Paragraph>().First().HeadingLevel); // 1
Console.WriteLine(document.Sections[0].Blocks.Count(b => b is Table));   // 1

string html = Markdown.ToHtml(markdown);                                 // 518 字符
Console.WriteLine(html.Contains("<a href=\"https://example.com\">"));    // True
Console.WriteLine(html.Contains("<input type=\"checkbox\""));            // True —— GFM 任务列表
```

链接与任务状态都落在统一模型上，不需要额外类型：链接是 `TextRun.Link`，任务项把 `☐` / `☑` 前缀写进 `Paragraph.ListMarker`（实测本例为 `☑ •`），加粗仍是 `TextFormat.Bold`。

模型回写 Markdown 会保留链接语法与 `[x]`，列表符号统一成 `-`，表格分隔行统一成 `:---`（实测 130 字符）：

```csharp
string back = Markdown.ToMarkdown(document);
Console.WriteLine(back.Contains("[官网](https://example.com)"));   // True
Console.WriteLine(back.Contains("- [x] 支持 Markdown 解析"));       // True
Console.WriteLine(back.Contains("| :--- | :--- |"));               // True
```

Markdown 没有内容签名，所以只有走文件路径才能自动识别：

```csharp
File.WriteAllText("t.md", markdown);
OfficeFile md = OfficeFile.Load("t.md");
Console.WriteLine(md.Format);                                       // Markdown（Detect 返回 Unknown，靠扩展名兜底）

md.Save("out/article.pdf");                                         // 56,889 字节
md.Save("out/article.docx");                                        // 3,964 字节
Console.WriteLine(OfficeConverter.ToMarkdown(docx.AsDocument()).Length); // 268 —— 任意文档 -> Markdown
OfficeConverter.ToMarkdown(document, "out/article.md");             // 206 字节，UTF-8 无 BOM
```

`链接`在 DOCX / PDF 里只剩文字：写出时不生成 `w:hyperlink` 关系，PDF 也不写链接注释。接着上面两个产物验证一下：

```csharp
using System.Linq;

OfficeFile linkDocx = OfficeFile.Load("out/article.docx");
int linkRuns = linkDocx.Document.Sections
    .SelectMany(s => s.Blocks.OfType<Paragraph>())
    .SelectMany(p => p.Inlines.OfType<TextRun>())
    .Count(t => t.Link != null);
Console.WriteLine(linkRuns);                                      // 0 —— 链接目标丢了
Console.WriteLine(linkDocx.ExtractText().Contains("官网"));         // True —— 链接文字还在

OfficeFile linkPdf = OfficeFile.Load("out/article.pdf");
Console.WriteLine(linkPdf.ExtractText().Contains("example.com"));  // False
Console.WriteLine(linkPdf.ExtractText().Contains("☑"));            // False —— 任务符号取不到
Console.WriteLine(linkPdf.ExtractText().Contains("支持 Markdown 解析")); // True —— 条目文字完整
```

需要可点击链接请交付 Markdown 或 HTML。

## 9. 验证产物

```csharp
OfficeFile pdf = OfficeFile.Load("out/sales.pdf");
string text = pdf.ExtractText();
Console.WriteLine(text.Contains("季度销售报表"));      // 标题已排入
Console.WriteLine(text.Contains("27 英寸显示器"));      // 数据行已排入
```

回读 Excel 时注意行号。`Report.ToWorkbook()` 在数据表之前会先写标题、副标题、生成时间、表格标题各占一行，然后才是表头行，所以本例中表头落在第 6 行、第一条数据在第 7 行：

```csharp
OfficeFile roundTrip = OfficeFile.Load("out/sales.xlsx");
Worksheet back = roundTrip.Workbook.Worksheets[0];

Console.WriteLine(back.GetCell(1, 1).Value);   // 季度销售报表（标题行）
Console.WriteLine(back.GetCell(6, 1).Value);   // 产品（表头行）
Console.WriteLine(back.GetCell(7, 1).Value);   // 无线键盘（第一条数据）
Console.WriteLine(back.GetCell(7, 3).Value);   // 3594 —— decimal 也按数值写入
```

不必硬记行号时，遍历 `sheet.Cells` 找目标文本更稳：

```csharp
foreach (Cell cell in back.Cells)
{
    if (cell.Value is string s && s == "无线键盘")
    {
        Console.WriteLine($"row {cell.Row}, col {cell.Column}");   // row 7, col 1
    }
}
```

五种格式各留一条最省事的断言，回归时一眼能看出是哪条链路断了：

```csharp
OfficeFile deckCheck = OfficeFile.Load("out/deck.pptx");
Console.WriteLine(deckCheck.Presentation.Slides.Count == 2);             // True
Console.WriteLine(deckCheck.Presentation.Slides[1].Shapes.Count == 3);    // True（标题 + 正文 + 表格）

OfficeFile mdCheck = OfficeFile.Load("out/article.md");
Console.WriteLine(mdCheck.Format);                                        // Markdown
Console.WriteLine(mdCheck.ExtractText().Contains("季度经营回顾"));          // True
```

`out/article.md` 那份产物只需检查"块数 > 0 且文本含关键词"，不必比较字节数——Markdown 与 PDF 的输出字节会随生成时间、表格列补齐等情况变化（实测同一份内容 161～465 字符之间浮动）。

库自身的测试套件用同一套断言覆盖了财务报表、看板图表、检验报告、考勤明细四类真实版式；把自己产物也纳入同样检查（标题存在、数据行存在、页数合理、无空页），是回归时最省事的护栏。

## 常见问题

**转换后中文变方块或空白。** PDF 生成需要从系统字体目录嵌入 TrueType 字体。Windows 默认可用；精简容器里请安装 Noto 或文泉驿，或用 `Wisdom.Pdf` 的 `FontResolver` 显式注册字体文件。

**PDF 只有一页或表格被切断。** 检查是否把整表放进了不允许拆分的位置：`Table.RepeatHeaderRow = true` 让表头在每页重复，`TableRow.Height` 与列宽决定单页能容纳多少行；Excel 侧则先确认打印区域与方向设置。

**Excel 数字在 Word/PDF 里变成 `####`。** 这是 `EmulateNumberOverflow = true` 时按 Excel 规则的模拟输出。要么加宽源工作表的列，要么把该选项设为 `false` 让排版直接输出文本。

**隐藏的工作表被输出了，或被漏掉了。** 两层默认值不同，容易搞混：`XlsxReader.Load` 读工作簿时**默认包含**隐藏表/行/列（`ExcelReadOptions` 三个开关默认 `true`）；而转换层的 `WorkbookLayoutOptions.IncludeHiddenWorksheets` **默认 `false`**，所以默认不会把隐藏表排进 PDF/DOCX。要连隐藏表一起输出就显式置 `true`；想在读取阶段就丢掉隐藏行列，用 `ExcelReadOptions { IncludeHiddenRows = false, IncludeHiddenColumns = false }`。

**想要图表出现在 PDF/Word 里。** 图表按 OOXML 保存，只有 Excel/WPS 会渲染它；导出 Word/PDF 时输出占位文本。需要图形请在源工作簿里改用 `ExcelImage` 锚定图片。

**幻灯片的版式回读后变了。** 版式是按形状推断的，不是存储的：实测 `SectionHeader` 回读成 `Title`、`TitleAndTable` 回读成 `TitleAndContent`，而带真实表格形状的 `TitleAndContent` 页会被推断成 `TitleAndTable`。形状本身（位置、尺寸、文本、填充）都原样回来，需要精确版式就 `AddSlide()` 拿到空白页自己摆形状。

**Markdown 从字节或流加载抛 `NotSupportedException`。** 它没有内容签名，`Detect` 只能返回 `Unknown`；走文件路径会按 `.md` / `.markdown` / `.mkd` 兜底，从内存加载必须显式 `OfficeFile.Load(bytes, OfficeFileFormat.Markdown)`（`Load(Stream)` 没有格式重载，先读成字节数组）。

**Word / PDF 里的链接点不动。** 模型用 `TextRun.Link` 承载链接，Markdown 与 HTML 输出会写成 `<a href>`，DOCX 写出与 PDF 排版只保留链接文字、丢弃目标地址。需要可点击链接就交付 Markdown / HTML，或在 DOCX 里自行补 `w:hyperlink`。

**同一份内容两次导出字节数不一样。** 两个已知来源：包里的生成时间戳，以及 `Table` 的列定义——写出时不入库，回读时按每行单元格数补齐，补齐后的表格在排版阶段会多画单元格（实测同一份幻灯片 48,259 → 50,558 字节）。回归断言请比较文本内容与页/行/列数量，不要比较文件哈希。
