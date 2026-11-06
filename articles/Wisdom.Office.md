# 一个 NuGet 包，把 Word / Excel / PPT / PDF / Markdown 全部打通

> 「纯 .NET 操作 Office」系列 · 门面包篇

## 这套库是怎么组织的

前面几篇分别讲了 Excel、Word、PDF 侧的读写。它们共享一个数据模型（`OfficeDocument`、`OfficeWorkbook`、`OfficePresentation`，都在 `Wisdom.Office.Core` 里），但各自成包，谁也不依赖谁。

问题是：真实项目里最常提的需求不是"读一个 docx"，而是**"把这个 Word 转成 PDF"、"把 Excel 排进报告"、"这份 Markdown 要出三个格式"**。跨格式就需要有人把所有模型串起来——这就是 `Wisdom.Office`：

```bash
dotnet add package Wisdom.Office
```

引用这一个包，就同时拿到 `Wisdom.Office.Core`、`Wisdom.Pdf`、`Wisdom.Word`、`Wisdom.Excel`、`Wisdom.Ppt`、`Wisdom.Markdown` 六个包的全部能力。零第三方依赖，不要求安装 Microsoft Office 或 WPS，支持 `net461` / `netstandard2.0` / `net7.0` / `net8.0` / `net9.0` / `net10.0`。

## 一、统一入口：OfficeFile

不必关心文件是什么格式：

```csharp
using Wisdom.Office;

OfficeFile file = OfficeFile.Load("input.xlsx");

Console.WriteLine(file.Format);          // Xlsx
Console.WriteLine(file.Metadata.Title);  // 文档属性
Console.WriteLine(file.ExtractText());   // 纯文本摘要

file.Save("copy.xlsx");
```

`Save` 按目标扩展名分派，所以同一个 `OfficeFile` 可以：

```csharp
OfficeFile file = OfficeFile.Load("report.md");
file.Save("report.pdf");     // Markdown → PDF
file.Save("report.docx");    // Markdown → DOCX
```

底层模型也都暴露着：`Document`（`OfficeDocument`）、`Workbook`（`OfficeWorkbook`）、`Presentation`（`OfficePresentation`）、`Pdf`（`PdfDocument`）。

还有两个方向转换方法，是整套互转的关键：

```csharp
OfficeDocument asWord = file.AsDocument();              // Excel 按排版选项、PPTX 按"每页一节"转成文档模型
OfficePresentation asSlides = file.AsPresentation();    // Word/Excel/PDF 按标题层级拆成幻灯片
```

PPTX 侧特别实用：

```csharp
OfficeFile file = OfficeFile.Load("deck.pptx");
Console.WriteLine(file.Presentation.Slides.Count);           // 页数
Console.WriteLine(file.Presentation.Slides[0].GetTitle());   // 首页标题
Console.WriteLine(file.ExtractText());                       // "第 N 页：" + 形状文本
```

## 二、格式检测不靠扩展名

```csharp
OfficeFileFormat format = OfficeFileFormats.Detect("input.bin");
// Pdf / Docx / Doc / Xlsx / Xls / Pptx / Unknown —— 基于文件内容签名
```

用户上传一个 `.docx` 其实是 zip 里的别的玩意，或者扩展名被改来改去，这种情况 `Detect` 比 `Path.GetExtension` 靠谱得多。`OfficeFile.Load(path)` 内部已经带内容检测。

**一个必须知道的例外**：Markdown 是纯文本，没有内容签名。`Detect` 对它返回 `Unknown`，`OfficeFile.Load(path)` 在识别失败后按扩展名（`.md` / `.markdown` / `.mkd`）兜底。也就是说：

```csharp
OfficeFile.Load(bytes, OfficeFileFormat.Markdown);   // 从字节加载必须显式指定
```

`Load(byte[])` 与 `Load(Stream)` 对纯文本会抛 `NotSupportedException`；`Load(Stream)` 没有格式重载，纯文本请先读成字节数组。

## 三、转换矩阵：哪些路走得通

```csharp
OfficeConverter.Convert("report.docx", "report.pdf");        // DOCX → PDF
OfficeConverter.Convert("legacy.doc", "legacy.docx");        // DOC  → DOCX
OfficeConverter.Convert("legacy.xls", "legacy.xlsx");        // XLS  → XLSX
OfficeConverter.Convert("notes.md", "notes.pptx");           // Markdown → PPTX
OfficeConverter.Convert("deck.pptx", "deck.docx");           // PPTX → DOCX

byte[] pdf = OfficeConverter.Convert(File.ReadAllBytes("report.docx"), OfficeFileFormat.Pdf);
```

实测矩阵（目标格式由扩展名决定）：

| 源 \ 目标 | PDF | DOCX | XLSX | PPTX | Markdown |
| --- | --- | --- | --- | --- | --- |
| Markdown | 通过 | 通过 | 不支持 | 通过 | 通过 |
| DOCX / DOC | 通过 | 通过 | 不支持 | 通过 | 通过 |
| PDF | 通过 | 通过 | 不支持 | 通过 | 通过 |
| XLSX / XLS | 通过 | 通过 | 通过 | 通过 | 通过 |
| PPTX | 通过 | 通过 | 不支持 | 通过 | 通过 |

只有"非电子表格 → XLSX"会抛 `NotSupportedException`（消息：`Word/PDF 文档不能直接保存为 Excel 工作簿。`）。其余路径都先汇入统一模型再排版，所以 Word、PDF、PPTX、Markdown 之间可以任意互转。`.doc`、`.xls` 只读，但作为源可用。

显式入口也有，不想绕 `OfficeFile` 时直接调：

```csharp
OfficeConverter.PresentationToPdf(presentation, "review.pdf");
OfficeConverter.PresentationToWord(presentation, "review.docx");
OfficeConverter.DocumentToPptx(document, "from-word.pptx");
string md = OfficeConverter.ToMarkdown(document);
```

## 四、Excel 转文档时的排版开关

Excel 是网格，Word/PDF 是流式版面，两者之间必须有一个"排版映射层"，`WorkbookLayoutOptions` 就是这个层的旋钮：

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

`EmulateNumberOverflow` 值得一提：Excel 里列宽不够会显示 `####`，这是渲染层行为不是数据。要不要在导出的 PDF/Word 里模拟这个行为，交给你决定；`Cell.DisplayText()` 本身始终给完整文本。

一个需要提前知道的取舍：**图表、形状等 Excel 装饰对象不会转入 PDF/Word**，只在原 `.xlsx` 里保留。要在 PDF 里出现图形，请在 Excel 里另存一份图片版并用 `ExcelImage` 锚定。PPTX → PDF/Word 同理，走"每页幻灯片一节"的线性排版，不还原绝对坐标、母版与动画，自选图形与图片按形状文本降级输出。

## 五、Report：一套数据出三种格式

这是门面包独有的高层 API，也是我个人觉得最能省时间的部分。做过的都知道：一个报表要出 PDF、Word、Excel，通常得写三套排版代码，然后维护三份对齐的样式。`Report` 把它压成一次声明：

```csharp
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

report.Export("report.pdf");    // 按扩展名分派
report.Export("report.docx");
report.Export("report.xlsx");

byte[] pdfBytes = report.ToPdf();
OfficeDocument doc = report.ToDocument();     // 想继续手工排版就拿模型
OfficeWorkbook wb = report.ToWorkbook();
```

标题、表头、斑马纹、边框、备注、冻结窗格全自动处理。外观由 `ReportStyle` 控制：`FontFamily`、`BaseFontSize`、`AccentColor`、`HeaderFill`、`HeaderTextColor`、`AlternateRowFill`、`BorderColor`、`ZebraStripes`。

列定义里的第三个参数是 Excel 数字格式串，PDF 和 Word 走同一套格式化实现，所以不会出现"PDF 里日期变成序列号"这种经典事故。

需要更多控制就落到模型层：`report.ToDocument()` / `ToWorkbook()` 返回的就是标准 `OfficeDocument` / `OfficeWorkbook`，可以继续加段落、合并单元格、图表，再交给对应 writer。

## 六、其它两个包的能力也一起进来了

引用门面包之后，`Wisdom.Word`、`Wisdom.Excel`、`Wisdom.Pdf`、`Wisdom.Ppt`、`Wisdom.Markdown` 的类型都可直接用：

```csharp
using Wisdom.Markdown;

OfficeDocument document = Markdown.ToDocument(markdownText);   // CommonMark + GFM 表格 / 任务列表
string html = Markdown.ToHtml(markdownText);                   // HTML 片段
string back = Markdown.ToMarkdown(document);                   // 模型回写 Markdown
```

```csharp
using Wisdom.Ppt;

OfficePresentation presentation = OfficePresentation.Widescreen();
Slide cover = presentation.AddSlide(SlideLayoutKind.Title);
((TextBoxShape)cover.Shapes[0]).Paragraphs[0].AddText("2026 年第三季度");
cover.NotesText = "开场 2 分钟。";
cover.Transition = SlideTransitionKind.Fade;
cover.TransitionDuration = 1.5;

PptxWriter.Write(presentation, "deck.pptx");
PptxWriter.WriteTemplate(presentation, "deck.potx");   // 同一模型直接存为 PowerPoint 模板
```

## 该引哪个包

| 需求 | 引用 |
| --- | --- |
| 只读写 Excel | `Wisdom.Excel` |
| 只读写 Word | `Wisdom.Word` |
| 只生成 / 解析 / 编辑 PDF | `Wisdom.Pdf` |
| 只读写 PPTX | `Wisdom.Ppt` |
| 只要 Markdown ↔ 模型 ↔ HTML | `Wisdom.Markdown` |
| 只要模型对象或模板引擎 | `Wisdom.Office.Core` |
| **任意跨格式转换、报表导出** | **`Wisdom.Office`** |

一个原则：**只碰单一格式就别引门面包**。`Wisdom.Word` 不依赖 `Wisdom.Pdf`，`Wisdom.Ppt` 不依赖 `Wisdom.Pdf` 与 `Wisdom.Word`——这是刻意的分包，让桌面应用和小程序端不必为用不到的引擎付体积。需要"Word 转 PDF"这类跨引擎动作时，门面包是唯一正确选择。

## 已知限制

- PDF 保存是按统一模型重新排版，不保留原始 PDF 的字节与完整视觉布局（要保原版面请用 `Wisdom.Pdf` 的 `PdfFile` 做对象级操作）。
- 图表、形状等 Excel 装饰对象不进入 PDF/Word。
- 非电子表格内容不能存为 XLSX；`.doc` / `.xls` 只读；`Report` 只产 PDF / DOCX / XLSX 三种。
- PPTX → PDF / Word 是线性排版，幻灯片里的自选图形与图片按形状文本降级输出。
- Markdown 无内容签名，只能靠扩展名兜底；`ToMarkdown` 会丢弃字体、字号与引用块语义。

## 小结

`Wisdom.Office` 的价值不在"多一个 API"，而在把**六种格式收敛到一个模型**：所有互转都走"汇入统一模型 → 重新排版"这一条路，因此新增一种格式不会带来 N×N 的组合爆炸，转换矩阵里除了"文档类 → 电子表格"这一类语义上讲不通的方向，其余全通。

配合 `Report` 这个高层 DSL，常见后端需求——"业务方要 PDF，财务要 Excel，存档要 Word"——大概能压缩到十几行代码。
