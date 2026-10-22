# Wisdom.Office

`Wisdom.Office` 是整套库的**统一门面包**：自动检测格式、读取任意受支持的 Office 文件、在 PDF / Word / Excel / PPTX / Markdown 之间互转，并提供开箱即用的报表导出 `Report`。引用这一个包即可拿到 Core、Pdf、Word、Excel、Ppt、Markdown 六个包的全部能力。

- 包名：`Wisdom.Office`
- 版本：`1.0.0`
- 依赖：`Wisdom.Office.Core`、`Wisdom.Pdf`、`Wisdom.Word`、`Wisdom.Excel`、`Wisdom.Ppt`、`Wisdom.Markdown`（同版本）
- 第三方依赖：无
- 命名空间：`Wisdom.Office`

本文是能力与 API 参考；分步实战教程见源码仓库的 `docs/Wisdom.Office.Tutorial.md`。

## 支持的框架

包内同时提供六份资产，NuGet 按目标框架自动选择，无需为 .NET 9 / 10 单独打包：

| 目标框架 | 包内资产 | 典型运行时 |
| --- | --- | --- |
| `net461` | `lib/net461` | .NET Framework 4.6.1 / 4.7.2 / 4.8.x（Windows Desktop） |
| `netstandard2.0` | `lib/netstandard2.0` | .NET Core 2.0-3.1、.NET 5/6、Mono、Unity |
| `net7.0` | `lib/net7.0` | .NET 7 |
| `net8.0` | `lib/net8.0` | .NET 8 |
| `net9.0` | `lib/net9.0` | .NET 9 |
| `net10.0` | `lib/net10.0` | .NET 10 / 11 |

.NET Framework 4.6.0 及以下不支持。六份资产由同一份源码编译，源码中没有任何 `#if NET*` 条件编译，功能完全一致。

## 安装

```bash
# 从 nuget.org
dotnet add package Wisdom.Office

# 指定版本
dotnet add package Wisdom.Office --version 1.0.0
```

`.NET Framework` 旧式 `packages.config` 项目：

```powershell
Install-Package Wisdom.Office -Version 1.0.0
```

## 快速开始

### 统一读取与保存

```csharp
using Wisdom.Office;

OfficeFile file = OfficeFile.Load("input.xlsx");

Console.WriteLine(file.Format);          // Xlsx
Console.WriteLine(file.Metadata.Title);  // 文档属性
Console.WriteLine(file.ExtractText());   // 纯文本摘要

file.Save("copy.xlsx");
```

`OfficeFile` 同时暴露底层模型：`Document`（`OfficeDocument`）、`Workbook`（`OfficeWorkbook`）、`Presentation`（`OfficePresentation`）、`Pdf`（`PdfDocument`）；`AsDocument()` 把 Excel 按排版选项、把 PPTX 按"每页一节"转成文档模型，`AsPresentation()` 反向把 Word/Excel/PDF 按标题层级拆成幻灯片。

```csharp
OfficeFile file = OfficeFile.Load("deck.pptx");
Console.WriteLine(file.Presentation.Slides.Count);      // 页数
Console.WriteLine(file.Presentation.Slides[0].GetTitle()); // 首页标题
Console.WriteLine(file.ExtractText());                  // "第 N 页：" + 形状文本
```

### 格式检测

```csharp
OfficeFileFormat format = OfficeFileFormats.Detect("input.bin");
// Pdf / Docx / Doc / Xlsx / Xls / Pptx / Unknown —— 基于文件内容签名，不依赖扩展名
```

Markdown 是纯文本、没有内容签名：`Detect` 对它返回 `Unknown`，`OfficeFile.Load(path)` 会在识别失败后按扩展名（`.md` / `.markdown` / `.mkd`）兜底，实测 `.md` 文件读回 `Format=Markdown`。从字节或流加载时没有扩展名可用，必须显式指定：`OfficeFile.Load(bytes, OfficeFileFormat.Markdown)`；`Load(byte[])` 与 `Load(Stream)` 对纯文本会抛 `NotSupportedException`（`Load(Stream)` 没有格式重载，请先读成字节数组）。

### 格式互转

```csharp
OfficeConverter.Convert("report.docx", "report.pdf");        // DOCX -> PDF
OfficeConverter.Convert("legacy.doc", "legacy.docx");        // DOC   -> DOCX
OfficeConverter.Convert("legacy.xls", "legacy.xlsx");        // XLS   -> XLSX
OfficeConverter.Convert("notes.md", "notes.pptx");           // Markdown -> PPTX
OfficeConverter.Convert("deck.pptx", "deck.docx");           // PPTX -> DOCX
OfficeConverter.Convert("data.xlsx", "data.pdf", new WorkbookLayoutOptions
{
    IncludeSheetNames = true,
    UseGridLines = true,
    FontFamily = "Microsoft YaHei",
    FontSize = 9
});

byte[] pdf = OfficeConverter.Convert(File.ReadAllBytes("report.docx"), OfficeFileFormat.Pdf);
```

实测的转换矩阵（`OfficeConverter.Convert(源, 目标)`，目标由扩展名决定）：

| 源 \ 目标 | PDF | DOCX | XLSX | PPTX | Markdown |
| --- | --- | --- | --- | --- | --- |
| Markdown | 通过 | 通过 | 不支持 | 通过 | 通过 |
| DOCX / DOC | 通过 | 通过 | 不支持 | 通过 | 通过 |
| PDF | 通过 | 通过 | 不支持 | 通过 | 通过 |
| XLSX / XLS | 通过 | 通过 | 通过 | 通过 | 通过 |
| PPTX | 通过 | 通过 | 不支持 | 通过 | 通过 |

只有"非电子表格 → XLSX"这一类会抛 `NotSupportedException`（消息：`Word/PDF 文档不能直接保存为 Excel 工作簿。`）；其余路径都先汇入统一模型再排版，所以 Word、PDF、PPTX、Markdown 之间可以任意互转。`.doc`、`.xls` 只读，但作为源可用。

### Markdown 读写

```csharp
using Wisdom.Markdown;

OfficeDocument document = Markdown.ToDocument(markdownText);   // CommonMark + GFM 表格/任务列表
string html = Markdown.ToHtml(markdownText);                   // HTML 片段，危险协议链接降级为纯文本
string back = Markdown.ToMarkdown(document);                   // 文档模型回写 Markdown
```

```csharp
OfficeFile md = OfficeFile.Load("article.md");
md.Save("article.pdf");                                        // Markdown -> PDF
md.Save("article.docx");                                       // Markdown -> DOCX
string markdown = OfficeConverter.ToMarkdown(md.AsDocument()); // 任意文档 -> Markdown
```

语法覆盖、`MarkdownOptions` 默认值与往返细节见 [src/Wisdom.Markdown/README.md](src/Wisdom.Markdown/README.md)。

### PPTX 读写

```csharp
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Ppt;

OfficePresentation presentation = OfficePresentation.Widescreen();   // 960 x 540 磅
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

PptxWriter.Write(presentation, "deck.pptx");
PptxWriter.WriteTemplate(presentation, "deck.potx");   // 同一模型直接存为 PowerPoint 模板

OfficePresentation loaded = PptxReader.Load("deck.pptx");
Console.WriteLine(loaded.Slides.Count);            // 2
Console.WriteLine(loaded.Slides[0].GetTitle());    // 2026 年第三季度
Console.WriteLine(loaded.Slides[0].TransitionDuration); // 1.5
```

实测保真：页序、页大小、标题/正文/备注、切换类型与时长、入场动画（类别 / 延迟 / 时长与 `advTm` 自动换页）、主题深改字段（`PresentationTheme` 的六个色板阶段、主/次字体与主题名）、自选图形（类型、填充色、几何，`CornerRadius` 默认 0.15）、图片（字节、alt 文本、尺寸）、表格（含单元格文本）与页面背景色都能写出并读回；`.potx` 模板可检测、可加载，按扩展名保存即在 `.pptx` / `.potx` 间切换。两点差异：表头占位形状回读成 `TextBoxShape`，`Table` 的列定义在写出时不入库、回读时按单元格数补齐；版式是"按形状推断"的，`SectionHeader` 回读成 `Title`、`TitleAndTable` 回读成 `TitleAndContent`，而带真实表格形状的 `TitleAndContent` 页会被推断成 `TitleAndTable`。导出 PDF 时按"每页幻灯片一节"排版，Word 同理。细节见 [src/Wisdom.Ppt/README.md](src/Wisdom.Ppt/README.md)。

### 报表导出

`Report` 用一套数据同时产出 PDF、DOCX、XLSX，自动处理标题、表头、斑马纹、边框、备注和冻结窗格。

```csharp
using System;
using Wisdom.Office;
using Wisdom.Office.Core;

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

report.AddNote("金额单位：人民币元。");

report.Export("report.pdf");   // 按扩展名分派
report.Export("report.docx");
report.Export("report.xlsx");

byte[] pdfBytes = report.ToPdf();
OfficeDocument doc = report.ToDocument();      // 可继续手工排版
OfficeWorkbook wb = report.ToWorkbook();
```

外观由 `ReportStyle` 控制：`FontFamily`、`BaseFontSize`、`AccentColor`、`HeaderFill`、`HeaderTextColor`、`AlternateRowFill`、`BorderColor`、`ZebraStripes`。

### Excel 转文档的排版选项

```csharp
WorkbookLayoutOptions options = new WorkbookLayoutOptions
{
    IncludeSheetNames = true,        // 每个工作表前输出表名标题
    IncludeHiddenWorksheets = false, // 隐藏表不参与排版
    UsePrintSetup = true,            // 沿用工作表的纸张、方向、页边距与打印区域
    UseGridLines = true,             // 用网格线代替逐格边框
    EmulateNumberOverflow = true,    // 列宽不足时按 Excel 规则输出 ####
    FontFamily = "Microsoft YaHei",
    FontSize = 9,
    GridColor = OfficeColor.LightGray,
    PageSetup = PageSetup.A4()
};
OfficeDocument document = WorkbookToDocumentConverter.Convert(workbook, options);
```

## 类型一览

| 类型 | 用途 |
| --- | --- |
| `OfficeFile` | 任意受支持文件的统一加载 / 检测 / 保存入口 |
| `OfficeFileFormat`、`OfficeFileFormats` | 格式枚举与内容签名检测 |
| `OfficeConverter` | 文件、流、字节三个层面的格式互转，含 `ToPptx` / `DocumentToPptx` / `PresentationToPdf` / `PresentationToWord` / `ToMarkdown` |
| `WorkbookToDocumentConverter` | 工作簿 → 文档模型的排版映射 |
| `WorkbookLayoutOptions` | 上述映射的排版开关 |
| `PresentationToDocumentConverter` | 演示文稿 → 文档模型（每页一节，供 PDF / Word 排版） |
| `DocumentToPresentationConverter` | 文档模型 → 演示文稿（按标题层级拆页） |
| `Markdown`（`Wisdom.Markdown`） | Markdown ↔ 文档模型 ↔ HTML |
| `PptxReader`、`PptxWriter`（`Wisdom.Ppt`） | `.pptx` 读写 |
| `Report`、`ReportTable`、`ReportColumn`、`ReportStyle` | 报表 DSL |

## 已知限制

- PDF 保存是按统一模型重新排版，不保留原始 PDF 的字节与完整视觉布局。
- 图表、形状等 Excel 装饰对象不会转入 PDF/Word，只在原 `.xlsx` 中保留；需要在 PDF/Word 里呈现请改用 `ExcelImage` 图片锚定。
- 非电子表格内容（Word、PDF、PPTX、Markdown）不能存为 XLSX；`.doc`、`.xls` 只读，`Report` 也只产出 PDF / DOCX / XLSX 三种报表文件。
- PPTX → PDF / Word 走的是"每页一节"的线性排版，不还原幻灯片的绝对坐标、母版与动画；幻灯片里的自选图形与图片按形状文本降级输出。
- Markdown 没有内容签名，只能靠扩展名兜底；`ToMarkdown` 会丢弃字体、字号与引用块语义，细节见 `Wisdom.Markdown` 的说明文档。
- 依赖扩展名的场景请以 `OfficeFileFormats.Detect` 为准，`OfficeFile.Load` 已经内置内容检测。
