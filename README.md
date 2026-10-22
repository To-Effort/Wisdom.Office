# Wisdom.Office

纯 C# 自研的 Office 文档读写与转换库，覆盖 **PDF、Word、Excel、PowerPoint、Markdown**，不依赖 Microsoft Office、WPS 或任何第三方 NuGet 组件。

- 目标框架：`net461` + `netstandard2.0` + `net7.0` + `net8.0` + `net9.0` + `net10.0`
- 零第三方依赖：ZIP、XML、OLE2 复合文档、PDF、OOXML、BIFF8、PresentationML、Markdown 解析器均自行实现
- 统一文档模型：Word、PowerPoint 与 Markdown 可汇入同一套 `OfficeDocument` 模型，统一导出 PDF / DOCX / PPTX
- 面向场景：报表导出、文档互转、模板填充、PDF 文本/表格提取、批量出幻灯片

## 格式支持

| 格式 | 读取 | 写入 | 说明 |
| --- | --- | --- | --- |
| PDF 1.7 | 支持 | 支持 | 解析文本、表格、元数据；生成由统一模型排版 |
| DOCX | 支持 | 支持 | 段落、表格、合并单元格、页眉页脚、图片、元数据 |
| DOC | 支持 | 不支持 | Word 97-2003 复合文档，读取正文与表格 |
| XLSX | 支持 | 支持 | 多工作表、样式、数字/日期格式、公式、合并、冻结窗格、打印设置、原生图表、条件格式、单元格图片 |
| XLS | 支持 | 不支持 | BIFF8（Excel 97-2003），读取值、样式、合并、窗格和元数据 |
| PPTX | 支持 | 支持 | 版式占位符、文本与富文本、表格、图片、自选图形、备注、切换动画；读回时按形状推断版式 |
| Markdown | 支持 | 支持 | CommonMark 主体 + GFM 表格/任务列表，可渲染 HTML 片段；无 Office 语义的字节级保真，按语法往返 |

### 文件检测

`OfficeFileFormats.Detect` 基于文件内容签名判断格式，不依赖扩展名：

```csharp
using Wisdom.Office;

OfficeFileFormat format = OfficeFileFormats.Detect("input.bin");
// Pdf / Docx / Doc / Xlsx / Xls / Pptx / Unknown
```

Markdown 是纯文本、没有文件头签名，只能由 `OfficeFile.Load(path)` 在识别失败后按扩展名（`.md` / `.markdown` / `.mkd`）兜底；从流或字节数组加载时要把 `OfficeFileFormat.Markdown` 显式传给 `Load(byte[], OfficeFileFormat)`。

## 项目结构

```text
src/
  Wisdom.Office.Core/   统一文档模型、工作簿模型、演示模型、OLE2、模板引擎、基础工具
  Wisdom.Pdf/           PDF 1.7 生成、解析、文本与表格提取、字体嵌入
  Wisdom.Word/          DOCX/DOTX/RTF/HTML/ODT/TXT 读写、DOC 读取、模板填充、标签纸/信封、修订跟踪、SmartArt/数字签名检查、文档整理与校验
  Wisdom.Excel/         XLSX 读写、模板填充、图表/条件格式/图片、XLS 读取
  Wisdom.Ppt/           PPTX 读写（PresentationML）、幻灯片版式与形状
  Wisdom.Markdown/      Markdown 解析、HTML 渲染与文档模型回写
  Wisdom.Office/        统一门面、格式互转、报表现成导出
tests/
  Wisdom.Office.Tests/  端到端测试、真实 DOC/XLS 样本、OOXML 包结构校验
samples/                四类真实场景的验证程序
docs/                   七个包各一份分步使用教程 + 一篇四场景功能导读
```

每个 `src/<包名>/README.md` 是该包的独立说明文档，打包时自动作为对应 NuGet 包内的 README（缺失时才回退到本文件）。

## 构建与测试

```bash
dotnet build Wisdom.Office.sln -c Release

# net8.0
dotnet run --project tests/Wisdom.Office.Tests/Wisdom.Office.Tests.csproj -f net8.0

# .NET Framework 4.6.1（Windows，直接跑生成的 exe）
dotnet build tests/Wisdom.Office.Tests/Wisdom.Office.Tests.csproj -c Release -f net461
"tests/Wisdom.Office.Tests/bin/Release/net461/Wisdom.Office.Tests.exe"
```

测试输出位于 `tests/Wisdom.Office.Tests/out/`。测试覆盖 PDF 生成与回读、DOCX/XLSX 往返、PPTX 往返（版式占位符、备注、切换、表格与图片）、Markdown 解析与渲染（front-matter、GFM 表格、任务列表、HTML 出口）、模板填充、真实 `.doc`/`.xls` 样本解析、格式转换、报表导出（财务报表、看板图表、检验报告、考勤明细等真实版式），以及 OOXML 内容类型、关系目标和 XML 合法性校验。同一套断言在 `net8.0` 与 `net461` 上各跑一遍，用于保证跨框架行为一致。

## 各库说明文档

每个包都有独立的说明文档，同时也是发布到 nuget.org 后包内的 README：

| 包 | 说明文档 | 内容 |
| --- | --- | --- |
| `Wisdom.Office` | [src/Wisdom.Office/README.md](src/Wisdom.Office/README.md) | 统一门面、格式检测与互转、`Report` 报表导出、`WorkbookLayoutOptions` |
| `Wisdom.Office.Core` | [src/Wisdom.Office.Core/README.md](src/Wisdom.Office.Core/README.md) | 统一文档模型、工作簿模型、模板引擎、OLE2 与 ZIP 基础设施 |
| `Wisdom.Pdf` | [src/Wisdom.Pdf/README.md](src/Wisdom.Pdf/README.md) | PDF 生成与排版、字体嵌入与子集化、解析与文本/表格提取、读取上限 |
| `Wisdom.Word` | [src/Wisdom.Word/README.md](src/Wisdom.Word/README.md) | DOCX/DOTX/RTF/HTML/ODT/TXT 读写、`.doc` 读取、模板填充、邮件合并区域、自定义属性、标签纸 / 信封、修订跟踪、SmartArt / 数字签名检查、文档整理 / 校验 / 差异 / 统计 |
| `Wisdom.Excel` | [src/Wisdom.Excel/README.md](src/Wisdom.Excel/README.md) | XLSX 读写、样式与数字格式、图表 / 条件格式 / 图片、模板填充、`.xls` 读取 |
| `Wisdom.Ppt` | [src/Wisdom.Ppt/README.md](src/Wisdom.Ppt/README.md) | 演示模型与版式占位符、PPTX 读写、表格 / 图片 / 自选图形 / 备注 / 切换动画 |
| `Wisdom.Markdown` | [src/Wisdom.Markdown/README.md](src/Wisdom.Markdown/README.md) | Markdown 解析、语法覆盖、HTML 渲染与安全策略、文档模型回写 |

## 使用教程

说明文档讲"有什么"，教程讲"怎么做"——每个包一份分步教程：输入文件由教程自身生成，代码片段拼起来就是可运行的控制台程序，文中输出全部是实测结果。

| 包 | 教程 | 你会做出什么 |
| --- | --- | --- |
| `Wisdom.Office` | [docs/Wisdom.Office.Tutorial.md](docs/Wisdom.Office.Tutorial.md) | 一份报表导出 DOCX / XLSX / PDF，再读回、互转、落盘 |
| `Wisdom.Office.Core` | [docs/Wisdom.Office.Core.Tutorial.md](docs/Wisdom.Office.Core.Tutorial.md) | 手工搭带页眉页脚与合并单元的文档模型、模板上下文、直接用 OLE2 / ZIP / 压缩工具 |
| `Wisdom.Pdf` | [docs/Wisdom.Pdf.Tutorial.md](docs/Wisdom.Pdf.Tutorial.md) | 长表格分页 PDF、字体解析与子集化、解析回读与提取表格 |
| `Wisdom.Word` | [docs/Wisdom.Word.Tutorial.md](docs/Wisdom.Word.Tutorial.md) | 脏草稿批量整理成规范文档、模板填充与邮件合并、全格式互转、修订接受/拒绝、SmartArt / 数字签名检查、自定义属性与标签纸 / 信封、一致性校验与两版差异 |
| `Wisdom.Excel` | [docs/Wisdom.Excel.Tutorial.md](docs/Wisdom.Excel.Tutorial.md) | 带样式与数字格式的台账、图表 / 条件格式 / 图片、模板填充、读真实报表与 `.xls` |
| `Wisdom.Ppt` | [docs/Wisdom.Ppt.Tutorial.md](docs/Wisdom.Ppt.Tutorial.md) | 版式占位符搭出的汇报稿、表格 / 图片 / 自选图形 / 备注 / 切换，读回与导出 PDF / Word |
| `Wisdom.Markdown` | [docs/Wisdom.Markdown.Tutorial.md](docs/Wisdom.Markdown.Tutorial.md) | 一份 Markdown 解析成文档模型、渲染 HTML 片段、导出 PDF / DOCX / PPTX 与回写 Markdown |

想先看"我的需求该从哪个入口进来"，读[《零依赖 C# Office 库到底怎么用——四个场景讲透 Wisdom.Office》](docs/Wisdom.Office.功能导读.md)：按报表导出、文档互转、模板填充、通用文档模型四个场景组织，含选型规则、默认值清单、安全上限与十一条实测出来的行为。

## 安装（NuGet）

> 七个包的版本号不再强制齐步：`Wisdom.Office`、`Wisdom.Office.Core`、`Wisdom.Pdf`、`Wisdom.Word`、`Wisdom.Excel` 当前为 `1.0.1`；`Wisdom.Ppt` 与 `Wisdom.Markdown` 自 `1.0.0` 起源码未变，仍按 `1.0.0` 发布，不产出内容重复的包。门面包会把这六个版本原样递传下来，所以 `dotnet add package Wisdom.Office` 一条命令即可。引用门面包 `Wisdom.Office` 即可获得全部能力；只需要单一格式时可只装 `Wisdom.Word`、`Wisdom.Excel`、`Wisdom.Ppt`、`Wisdom.Pdf`、`Wisdom.Markdown` 或 `Wisdom.Office.Core`。发布步骤见下一节。

```bash
dotnet add package Wisdom.Office

# 或指定版本
dotnet add package Wisdom.Office --version 1.0.1
```

## 发布流程（nuget.org Trusted Publishing）

发布由 `.github/workflows/publish-nupkg.yml` 完成，走 nuget.org 的可信发布（OIDC 短期凭据），仓库内不保存长期 API Key。

1. 在 nuget.org 打开头像菜单的 **Trusted Publishing**，按下表新建策略（大小写不敏感，必须逐字一致）：

   | 表单项 | 取值 |
   | --- | --- |
   | Policy Name | `wisdom-office-github-actions` |
   | Package Owner | `YF.f` |
   | CI/CD Provider | `GitHub Actions` |
   | Repository Owner | `To-Effort` |
   | Repository | `Wisdom.Office` |
   | Workflow File | `publish-nupkg.yml` |
   | Environment | 留空 |

   `Workflow File` 只填文件名，不带 `.github/workflows/` 路径；策略作用于该 Owner 名下的全部包，若只想覆盖部分包，可在 Policy Scopes 里用 glob 限定 `Wisdom.*`。`Environment` 留空表示不校验 GitHub 环境；若填了名字（如 `production`），必须在 workflow 的 job 上声明同名 `environment:`，否则令牌校验失败。

2. 把 `Directory.Build.props` 里的 `<Version>` 改成新版本号，提交并推送到 `master`。

3. 打标签触发发布，标签必须与版本号一致（流水线会校验，不一致直接失败）：

   ```bash
   git tag v1.0.1
   git push origin v1.0.1
   ```

4. 在 Actions 的 `Publish NuGet packages` 流水线查看结果；包在 nuget.org 索引生效需要几分钟。

私有仓库首次发布的策略会先进入 7 天待激活状态，期间成功发布一次即转为永久生效。

临时手工发布（仅在 CI 不可用时）：创建一个仅对 `Wisdom.*` 前缀生效、带过期时间的 scoped API Key，然后执行：

```powershell
dotnet pack Wisdom.Office.sln -c Release -o ./nupkg
dotnet nuget push .\nupkg\*.nupkg --api-key $env:NUGET_API_KEY --source https://api.nuget.org/v3/index.json --skip-duplicate
```

### 支持的运行环境

每个包同时提供六份资产，安装后 NuGet 按目标框架精确匹配、未命中时向前回退：

| 运行时 | 采用的包内资产 | 支持情况 |
| --- | --- | --- |
| .NET Framework 4.6.1 / 4.6.2 / 4.7.2 / 4.8.x | `lib/net461` | 支持 |
| .NET Core 2.0 - 3.1 | `lib/netstandard2.0` | 支持 |
| .NET 5 / 6 | `lib/netstandard2.0` | 支持 |
| .NET 7 | `lib/net7.0` | 支持 |
| .NET 8 | `lib/net8.0` | 支持 |
| .NET 9 | `lib/net9.0` | 支持 |
| .NET 10 / 11 | `lib/net10.0` | 支持 |
| Mono / Unity（兼容 netstandard2.0 的运行时） | `lib/netstandard2.0` | 支持 |
| .NET Framework 4.6.0 及以下 | 无 | 不支持 |

- 六份资产由同一份源码编译，源码中没有任何 `#if NET*` 条件编译，因此较低或较高的框架版本都不会丢功能；新增的 .NET 版本会自动选中 `net10.0`（或向前回退到 `netstandard2.0`），不必重新打包。
- `net461` 资产引用框架自带的 `System.IO.Compression` 与 `System.IO.Compression.FileSystem`，无需额外 NuGet 包。
- 已实测的消费端组合：.NET 7、.NET 8、.NET 9、.NET 10 与 .NET Framework 4.6.1 / 4.7.1，各自从 NuGet 包引用后跑通「写 XLSX → 转 PDF / DOCX → 回读」全流程；本仓库的端到端测试套件同时在 `net8.0` 与 `net461` 上全绿。

- 零第三方依赖，不依赖 Microsoft Office、WPS，Windows / Linux / macOS 均可运行。
- PDF 生成需要从系统字体目录嵌入 TrueType/TTC 字体；无桌面字体的精简容器环境请预先安装字体（如 Noto、文泉驿）。

## 快速开始

安装 `Wisdom.Office` 包后即可使用（下面各示例的 `using` 已标注所需命名空间）。

### 统一读取、检测与保存

```csharp
using Wisdom.Office;

OfficeFile file = OfficeFile.Load("input.xlsx");

Console.WriteLine(file.Format);          // Xlsx
Console.WriteLine(file.Metadata.Title);  // 文档属性
Console.WriteLine(file.ExtractText());   // 纯文本摘要

file.Save("copy.xlsx");
```

`OfficeFile` 同时暴露底层模型：

- `Document`：Word/PDF/Markdown 对应的 `OfficeDocument`
- `Workbook`：Excel 对应的 `OfficeWorkbook`
- `Presentation`：PPTX 对应的 `OfficePresentation`
- `Pdf`：PDF 解析结果 `PdfDocument`

`AsDocument()` 与 `AsPresentation()` 负责跨模型：前者把 Excel 按排版选项、把 PPTX 按"每页一节"转成文档模型，后者把 Word/Excel/PDF 按标题层级反向拆成幻灯片。

### 格式互转

```csharp
using Wisdom.Office;

// DOCX -> PDF
OfficeConverter.Convert("report.docx", "report.pdf");

// XLSX -> PDF，并控制工作表到文档的排版
OfficeConverter.Convert("data.xlsx", "data.pdf", new WorkbookLayoutOptions
{
    IncludeSheetNames = true,
    UseGridLines = true,
    FontFamily = "Microsoft YaHei",
    FontSize = 9
});

// DOC -> DOCX，XLS -> XLSX
OfficeConverter.Convert("legacy.doc", "legacy.docx");
OfficeConverter.Convert("legacy.xls", "legacy.xlsx");

// Markdown -> PPTX / PDF / HTML 之外的 Office 格式
OfficeConverter.Convert("notes.md", "notes.pptx");

// PPTX -> PDF / DOCX / Markdown
OfficeConverter.Convert("deck.pptx", "deck.docx");
```

实测的转换矩阵（`OfficeConverter.Convert(源, 目标)`，目标格式由扩展名决定）：

| 源 \ 目标 | PDF | DOCX | XLSX | PPTX | Markdown |
| --- | --- | --- | --- | --- | --- |
| Markdown | 通过 | 通过 | 不支持 | 通过 | 通过 |
| DOCX / DOC | 通过 | 通过 | 不支持 | 通过 | 通过 |
| PDF | 通过 | 通过 | 不支持 | 通过 | 通过 |
| XLSX / XLS | 通过 | 通过 | 通过 | 通过 | 通过 |
| PPTX | 通过 | 通过 | 不支持 | 通过 | 通过 |

只有"非电子表格 → XLSX"会抛 `NotSupportedException`（消息：`Word/PDF 文档不能直接保存为 Excel 工作簿。`），其余组合都先汇入统一模型再排版。旧版 `.doc`、`.xls` 仅支持读取，但作为源可用。

### 报表导出

`Report` 用一套数据同时生成 PDF、DOCX、XLSX，自动处理标题、表头、斑马纹、边框、备注和冻结窗格。

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
table.AddRow("27 英寸显示器", 5, 9495.0, new DateTime(2026, 7, 7));

report.AddNote("金额单位：人民币元。");

report.Export("report.pdf");
report.Export("report.docx");
report.Export("report.xlsx");

// 也可获取字节数组
byte[] pdf = report.ToPdf();
byte[] docx = report.ToWord();
byte[] xlsx = report.ToExcel();
```

### Excel 读写

```csharp
using Wisdom.Excel;
using Wisdom.Office.Core;

OfficeWorkbook workbook = new OfficeWorkbook();
workbook.Metadata.Title = "销售数据";

Worksheet sheet = workbook.AddWorksheet("明细");
sheet.GetCell(1, 1).Value = "产品";
sheet.GetCell(1, 2).Value = "数量";
sheet.GetCell(1, 3).Value = "金额";

sheet.GetCell(2, 1).Value = "无线键盘";
sheet.GetCell(2, 2).Value = 12;

Cell amount = sheet.GetCell(2, 3);
amount.Value = 299.5;
amount.Style = new CellStyle
{
    NumberFormat = "#,##0.00",
    Bold = true
};

sheet.Merge("A4:C4");
sheet.GetCell(4, 1).Value = "备注：本季度价格含税。";
sheet.FrozenRows = 1;
sheet.SetColumnWidth(1, 24);

XlsxWriter.Write(workbook, "sales.xlsx");

OfficeWorkbook loaded = XlsxReader.Load("sales.xlsx");
Console.WriteLine(loaded.Worksheets[0].GetCell(2, 2).Value);
```

`Cell.Value` 支持 `string`、`double`、`bool`、`DateTime`；`Cell.Formula` 保存公式，`Value` 可保存缓存结果。

### Excel 图表、条件格式与图片

图表、条件格式和锚定图片随 `XlsxWriter` 一起写出，`XlsxReader` 回读时同样还原，因此可以「加载 → 修改 → 保存」而不丢装饰：

```csharp
using Wisdom.Excel;
using Wisdom.Office.Core;

OfficeWorkbook workbook = new OfficeWorkbook();
Worksheet sheet = workbook.AddWorksheet("看板");
for (int c = 1; c <= 7; c++) sheet.SetColumnWidth(c, 10);
for (int i = 0; i < 6; i++) sheet.GetCell(1, i + 2).Value = (i + 1) + "月";
sheet.GetCell(2, 1).Value = "销售额";
for (int i = 0; i < 6; i++) sheet.GetCell(2, i + 2).Value = 820.0 + i * 100;

// 原生图表：柱状/条形/堆积/折线/面积/饼/环形/散点，数据源用区域引用或字面量集合
ExcelChart chart = new ExcelChart
{
    Type = ExcelChartType.Column,
    Title = "上半年销售额",
    Row = 4,
    Column = 1,
    WidthPoints = 360,
    HeightPoints = 200,
    ValueAxisFormat = "#,##0"
};
ExcelChartSeries series = chart.AddSeries("看板!$A$2", "看板!$B$1:$G$1", "看板!$B$2:$G$2");
series.FillColor = OfficeColor.FromHtml("#00B050");
series.ShowDataLabels = true;
sheet.Charts.Add(chart);

// 条件格式：数据条、色阶、图标集、按数值高亮（dxf 样式真实落盘）
ExcelConditionalFormat format = new ExcelConditionalFormat("B2:G2");
format.AddDataBar(new OfficeColor(0x63, 0xC3, 0x84));
format.AddColorScale(new OfficeColor(0xF8, 0x69, 0x6A), new OfficeColor(0x63, 0xC3, 0x84));
sheet.ConditionalFormats.Add(format);

// 锚定到单元格的图片（PNG / JPEG 自动识别）
ExcelImage image = new ExcelImage(OfficeImage.FromFile("flag.png"))
{
    Row = 2,
    Column = 9,
    WidthPoints = 24,
    HeightPoints = 24
};
sheet.Images.Add(image);

XlsxWriter.Write(workbook, "dashboard.xlsx");

OfficeWorkbook loaded = XlsxReader.Load("dashboard.xlsx");
Console.WriteLine(loaded.Worksheets[0].Charts.Count);          // 1
Console.WriteLine(loaded.Worksheets[0].Charts[0].Series[0].ValueRange); // 看板!$B$2:$G$2
```

图表本身按 OOXML 保存，由 Excel/WPS 渲染；导出 Word/PDF 时图表位置输出占位文本，需要图片版请改用 `ExcelImage`。列宽不足以容纳格式化后的数字时，Word/PDF 排版会按 Excel 的规则输出 `####`，可用 `WorkbookLayoutOptions.EmulateNumberOverflow = false` 关闭。

### PowerPoint 读写

幻灯片模型同样以磅为单位（`Widescreen()` 为 960 x 540），`AddSlide(版式)` 按版式生成占位符，`AddSlide()` 不带参数得到完全自由定位的空白页。

```csharp
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Ppt;

OfficePresentation presentation = OfficePresentation.Widescreen();
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

PptxWriter.Write(presentation, "review.pptx");

OfficePresentation loaded = PptxReader.Load("review.pptx");
Console.WriteLine(loaded.Slides.Count);                  // 2
Console.WriteLine(loaded.Slides[0].GetTitle());          // 2026 年第三季度
Console.WriteLine(loaded.Slides[0].TransitionDuration);  // 1.5
Console.WriteLine(loaded.Slides[1].Layout);              // TitleAndTable（按形状推断）

// 经门面导出
OfficeConverter.Convert("review.pptx", "review.pdf");
OfficeConverter.Convert("review.pptx", "review.docx");
```

`Slide.AddShape` / `AddImage` / `AddTable` 分别产出 `AutoShape`（十种预设图形、`CornerRadius` 默认 0.15）、`PictureShape`（PNG / JPEG，带 `AlternativeText`）与 `TableShape`；页背景色、演讲者备注与五种切换效果（含时长）写出后都能原样读回。版式是读回时按形状推断的，`SectionHeader` 会归成 `Title`，带真实表格形状的页面会归成 `TitleAndTable`。

### Markdown 读写

```csharp
using System.IO;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Markdown;

string text = File.ReadAllText("article.md", Encoding.UTF8);

OfficeDocument document = Markdown.ToDocument(text);      // CommonMark 主体 + GFM 表格/任务列表
string html = Markdown.ToHtml(text);                      // HTML 片段，图片转 data URI，危险协议链接降级
string markdown = Markdown.ToMarkdown(document);          // 文档模型回写 Markdown

OfficeFile file = OfficeFile.Load("article.md");          // 按扩展名识别为 Markdown
file.Save("article.pdf");                                 // Markdown -> PDF
file.Save("article.pptx");                                // Markdown -> PPTX（按标题拆页）
```

front-matter 的 `title` / `author` / `date` 会写进 `document.Metadata`；`MarkdownOptions` 控制表格、任务列表、硬换行、原始 HTML、等宽字体与基准字号（默认 `EnableTables` / `EnableTaskLists` / `TrimTrailingWhitespace` 为真，`HardLineBreaks` / `EnableRawHtml` 为假，等宽字体 `Consolas`，正文 11 磅）。HTML 出口只接受相对地址、`#锚点` 与 `http` / `https` / `mailto` / `ftp` / `ftps` / `tel`，`javascript:` 一类降级为不可点击的蓝色下划线文本。引用式链接、脚注与列表项内的缩进续段不解析，按字面文本保留。

### 模板填充

模板占位符格式为 `{{key}}`、`{{key:format}}`，支持跨 Run 占位符、点路径、列表循环和空列表清理。

Word 段落/表格行循环：

```text
{{customerName}}，您好：

{{#each items}}
{{name}}    数量：{{qty}}    金额：{{amount:0.00}}
{{/each}}
```

Excel 单元格循环（循环可跨一行或多行）：

```text
A6: {{#each items}}{{name}}
B6: {{qty}}
C6: {{price:0.00}}
D6: {{@number}}
E6: {{/each}}
```

```csharp
using Wisdom.Excel;
using Wisdom.Office.Core;
using Wisdom.Word;

TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["items"] = new[]
{
    new { name = "无线键盘", qty = 2, price = 199.0 },
    new { name = "机械鼠标", qty = 5, price = 89.9 }
};

XlsxTemplate.Fill("template.xlsx", "filled.xlsx", context);
DocxTemplate.Fill("template.docx", "filled.docx", context);
```

模板选项可控制未解析占位符、XML 转义和格式化区域：

```csharp
TemplateOptions options = new TemplateOptions
{
    LeaveUnresolvedTokens = false,
    HtmlEncodeValues = true,
    Culture = System.Globalization.CultureInfo.InvariantCulture
};
```

### PDF 解析

```csharp
using Wisdom.Pdf;

PdfDocument pdf = PdfDocument.Load("input.pdf");
Console.WriteLine(pdf.PageCount);

foreach (PdfPage page in pdf.Pages)
{
    Console.WriteLine(page.ExtractText());
    foreach (PdfTable table in page.GetTables())
    {
        Console.WriteLine($"表格 {table.RowCount} x {table.ColumnCount}");
    }
}

OfficeDocument document = pdf.ToOfficeDocument();
```

### 直接用统一模型生成 PDF/Word

```csharp
using Wisdom.Office;
using Wisdom.Office.Core;

OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "示例文档";

Section section = document.AddSection(PageSetup.A4());
section.Blocks.Add(new Paragraph("标题").AddText(" - 副标题", new TextFormat("Arial", 10)));

Table table = new Table();
table.Columns.Add(new TableColumn(160));
table.Columns.Add(new TableColumn(100));
table.AddRow("项目", "金额");
TableRow row = table.AddRow("合计", "2,288.40");
row.IsHeader = false;
section.Blocks.Add(table);

OfficeConverter.ToPdf(document, "document.pdf");
OfficeConverter.ToWord(document, "document.docx");
```

## 设计说明

- **统一模型**：Word/PDF/Markdown 共用 `OfficeDocument`（节、段落、表格、图片、页眉页脚），Excel 使用独立 `OfficeWorkbook`，PowerPoint 使用独立 `OfficePresentation`，分别经 `WorkbookToDocumentConverter`、`PresentationToDocumentConverter` 映射后统一渲染；`DocumentToPresentationConverter` 反向按标题层级拆页。
- **纯托管实现**：`ZipArchive` 处理 OOXML 包，LINQ to XML 处理部件，`CompoundFile` 读取 OLE2；PDF 生成器自行输出对象、字体与内容流；PresentationML 与 Markdown 解析器同样手写，不依赖 Markdig。
- **字体**：PDF 生成会从系统字体目录嵌入 TrueType/TTC 字体，无需安装 Office；Linux/macOS 下可回退到 Noto、Liberation、文泉驿等已注册字体，也可通过 `FontResolver` 注册自定义字体。
- **报表层**：`Report` 是构建在统一模型之上的轻量 DSL，不是独立排版引擎。

## 已知限制

- 旧版 `.doc`、`.xls` 仅支持读取，不支持写回；`.doc` 的复杂域、批注、修订不保证还原。
- `.xls` 仅支持 BIFF8（Excel 97-2003），不支持 BIFF5/7、加密工作簿；公式只保留缓存值，不还原公式文本。
- PDF 不支持加密文档；PDF 保存是依据统一模型重新排版，不保留原始 PDF 的字节和完整视觉结构。
- PDF 文本提取依赖文本操作符，扫描件需要额外 OCR；表格识别为启发式，复杂版式可能需人工校正。
- PDF 字体嵌入支持 TrueType `glyf` 轮廓，不支持 CFF/OpenType（`OTTO`）字体嵌入；PNG 不支持隔行扫描。
- Word/PDF/PPTX/Markdown 不能直接输出为 Excel；`WorkbookLayoutOptions` 不支持的复杂对象（如图表、形状）不会转入 PDF/Word。
- PPTX 不读写旧版 `.ppt`，不处理图表、SmartArt、嵌入 OLE 与动画序列；版式与主题由本库生成一套固定的 `slideMaster` / `slideLayout` / `theme`，不解析外部模板 `.potx`。导出 PDF / Word 时按"每页一节"线性排版，不还原绝对坐标。
- Markdown 没有文件头签名，只能靠扩展名兜底；不是完整 CommonMark / GFM 实现（引用式链接、脚注、HTML 块内 Markdown、列表项内缩进续段按字面文本保留），`ToMarkdown` 会丢弃字体、字号与引用块语义。
- 生成文件会通过 ZIP 包、内容类型、关系和 XML 合法性校验；正式上线前仍建议用目标 Office 版本抽查复杂排版效果。

## 1.0 版本说明

相对 0.1.0 预发布版本：

- **框架覆盖扩展为六份资产**：`net461`、`netstandard2.0`、`net7.0`、`net8.0`、`net9.0`、`net10.0`。`.NET Framework 4.6.1 及以上`拥有专属资产而不再复用 `netstandard2.0`。
- **跨框架行为统一（缺陷修复）**：`net461` 的 `ZipArchive` 延迟解析中央目录，损坏的 OOXML 条目名会在无关调用点抛 `ArgumentException`，绕过读取器的 `catch (IOException)`。新增 `PackageArchive.OpenForRead` 立即物化中央目录并把这类异常统一收敛为 `InvalidDataException`，已应用于 XLSX/DOCX 读取、DOCX 模板填充与格式检测四处入口；端到端测试套件同时新增 `net461` 目标并验证通过。
- **Excel 装饰对象**：原生图表（柱/条/折/面积/饼/环/堆积）、条件格式（数据条、色阶、图标集、数值比较高亮，含 `dxf` 差量样式）、锚定图片，写出与回读双向支持。
- **Excel → PDF/Word 排版保真**：数字格式分节（正;负;零;文本）、`ShowZeros`、打印区域、按分带裁剪与标识列重复、列宽不足时按 Excel 规则输出 `####`（可用 `WorkbookLayoutOptions.EmulateNumberOverflow` 关闭）。
- **新增两个包**：`Wisdom.Ppt` 负责 PresentationML 读写（七种版式与占位符、富文本段落、表格含行列合并、图片与替代文字、十种自选图形、背景色、备注与五种切换效果），`Wisdom.Markdown` 负责 Markdown ↔ 统一文档模型 ↔ HTML（CommonMark 主体 + GFM 管道表格与任务列表、front-matter、等宽代码块、危险协议链接降级）。两者与门面包一同汇入 `OfficeDocument`，因此 Markdown / PPTX 与 PDF / DOCX / XLSX 之间可以互转。
- **真实报表回归测试**：新增财务报表、数据看板、检验/审核报告、考勤明细等版式的端到端断言，修复跨页合并表头重复、行跨度原子拆分与空页问题。
- **文档**：七个包各附一份说明文档（同时作为包内 README），本文件更新为 1.0 的框架矩阵与安装说明。
- **许可证**：MIT。

## 许可证

本项目采用 [MIT](LICENSE) 许可，见 LICENSE 文件。
