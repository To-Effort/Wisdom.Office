# 为什么 Office 库需要一个"统一文档模型"：Wisdom.Office.Core 拆解

> 「纯 .NET 操作 Office」系列 · 内核篇

## 先回答一个设计问题

假设你要写一套能同时处理 Word、Excel、PPT、PDF、Markdown 的库。最省事的做法是每个格式一套 API：`WordDoc`、`ExcelBook`、`PptDeck`……然后"Word 转 PPT"就得写 N×N 个转换器。

`Wisdom` 这套库选了另一条路：先定义**三种通用模型**——文档（`OfficeDocument`）、工作簿（`OfficeWorkbook`）、演示文稿（`OfficePresentation`）——所有格式的读写器只负责"把自家文件映射到模型 / 从模型落盘"，跨格式转换统统走"汇入模型 → 重新排版"。

承载这些模型和基础设施的就是 `Wisdom.Office.Core`：

```bash
dotnet add package Wisdom.Office.Core
```

它是另外六个包（Pdf / Word / Excel / Ppt / Markdown / Office 门面）的**唯一依赖，自身零依赖**。只想要模型对象或模板引擎的场合，可以单独引用它。支持 `net461` / `netstandard2.0` / `net7.0` / `net8.0` / `net9.0` / `net10.0`，源码无任何 `#if NET*` 条件编译，六份资产功能完全一致。

## 一、文档模型：Word 和 PDF 共用一套

层级是 `OfficeDocument → Section → Block`，`Block` 只有三种：`Paragraph`、`Table`、`PageBreak`。

```csharp
using Wisdom.Office.Core;

OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "示例文档";
document.Metadata.Author = "Wisdom";

Section section = document.AddSection(PageSetup.A4());

Paragraph title = new Paragraph("季度总结").AddText("　2026 Q3", new TextFormat("Arial", 10));
title.HeadingLevel = 1;
title.Alignment = ParagraphAlignment.Center;
title.KeepWithNext = true;
section.Blocks.Add(title);

Table table = new Table();
table.Columns.Add(new TableColumn(160));
table.Columns.Add(new TableColumn(100));
table.AddRow("项目", "金额");
table.RepeatHeaderRow = true;
section.Blocks.Add(table);
```

关键类型一览：

| 类型 | 关键成员 |
| --- | --- |
| `OfficeDocument` | `Metadata`、`Sections`、`DefaultTextFormat`、`AddSection()`、`FirstSection` |
| `Section` | `PageSetup`、`Blocks`、`Header`、`Footer`、`AddParagraph()` |
| `PageSetup` | `Width/Height`、四边 `Margin*`、`HeaderDistance`、`FooterDistance`、`IsLandscape`、`A4()`、`Letter()` |
| `HeaderFooter` | `Blocks`、`IsEnabled`、`AddParagraph()` |
| `Paragraph` | `Inlines`、`Alignment`、`SpaceBefore/After`、`LineSpacing`、缩进、`HeadingLevel`、`KeepWithNext`、`PageBreakBefore`、`ListMarker`、`AddText()`、`AddImage()`、`GetText()` |
| `TextRun` / `FieldRun` / `ImageRun` | 文本、超链接 `Link`、域（`PAGE`、`NUMPAGES`）、行内图片 |
| `TextFormat` | `FontFamily`、`AsciiFontFamily`、`EastAsiaFontFamily`、`FontSize`、`Bold/Italic/Underline/Strikeout`、`Color`、上下标、`Heading(level)` |
| `Table` / `TableRow` / `TableCell` | 列宽、行 `IsHeader`/`Height`、单元格 `ColSpan`/`RowSpan`/`BackgroundColor`/`VerticalAlignment`/`Borders`、`CellPadding`、`Autofit` |
| `OfficeColor` | `FromHtml("#00B050")`、`FromArgb`、`ToHex`、`TryParse`、预置色 |
| `OfficeImage` | `FromFile()`、`Data`、`Format`、`PixelWidth/Height`、`DefaultWidthPoints` |

### 同一段模型，Word 里是"域"，PDF 里是真实页码

这是统一模型最漂亮的一处设计。页脚这样写一次：

```csharp
HeaderFooter footer = section.Footer;
Paragraph p = footer.AddParagraph(string.Empty);
TextFormat small = new TextFormat("SimSun", 8);
p.AddText("第 ", small);
p.Inlines.Add(new FieldRun("PAGE", "{page}", small));
p.AddText(" 页　共 ", small);
p.Inlines.Add(new FieldRun("NUMPAGES", "{pages}", small));
p.AddText(" 页", small);
```

交给 Word 写出器，产出真实的 `fldSimple` 域，Word 打开时自己算页码；交给 PDF 渲染器，`{page}` / `{pages}` 在排版时被替换成实际页码。**模型不关心谁会消费它**，两个占位串（域代码 + 替换标记）就把两种渲染时机都覆盖了。

### 中西文字体是两个字段

`TextFormat` 把西文族名和中文族名分开（`AsciiFontFamily` / `EastAsiaFontFamily`），对应 OOXML 里的 `w:rFonts` 的 `ascii` / `eastAsia`。这是中文排版的刚需：正文西文 Times New Roman + 中文宋体这种混排，只给一个 `FontFamily` 是写不出来的。

### 模型层不做分页

`{page}` 的实际替换发生在 PDF 渲染阶段，模型本身不计算分页。同理 `TextRun.Link` 目前只被 Markdown / HTML 写出消费，DOCX 写出时不生成 `w:hyperlink` 关系、PDF 也不写链接注释——需要真正的 Word 超链接请使用 `Wisdom.Word` 的 `WordDocumentBuilder.InsertHyperlink`。

## 二、工作簿模型：Excel

```csharp
OfficeWorkbook workbook = new OfficeWorkbook();
workbook.Metadata.Title = "销售数据";

Worksheet sheet = workbook.AddWorksheet("明细");
sheet.GetCell(1, 1).Value = "产品";
sheet.GetCell(2, 2).Value = 12;                 // string / double / bool / DateTime
sheet.GetCell(2, 3).Style = new CellStyle { NumberFormat = "#,##0.00", Bold = true };
sheet.Merge("A4:C4");
sheet.SetColumnWidth(1, 24);
sheet.FrozenRows = 1;
sheet.SetPrintArea("A1:D20");
```

`Worksheet` 不只装单元格，还承载装饰集合 `Charts`、`Images`、`ConditionalFormats`，以及表格级属性 `MergedRanges`、`ColumnWidths`、`RowHeights`、`HiddenRows/Columns`、`IsHidden`、`ShowGridLines`、`ShowZeros`、`FrozenRows/Columns`、`PrintSetup`、`FitToPage`、`DefaultStyle`。单元格地址工具是 `CellAddress` 与 `SpreadsheetRange`。

要注意分工：**模型定义在本包，落盘与回读由 `Wisdom.Excel` 实现**。在这里你能构造一个 `ExcelChart`，但它不会自己变成文件。

## 三、演示文稿模型：PowerPoint

层级 `OfficePresentation → Slide → SlideShape`，形状派生 `TextBoxShape`、`AutoShape`、`PictureShape`、`TableShape`。

所有几何量以**磅**为单位，原点在幻灯片左上角——不是 EMU，不是像素。写出时由 `Wisdom.Ppt` 按 1 磅 = 12700 EMU 换算。

```csharp
OfficePresentation presentation = OfficePresentation.Widescreen();   // 960 x 540
presentation.Metadata.Title = "季度经营回顾";

Slide cover = presentation.AddSlide(SlideLayoutKind.Title);          // 版式自动生成占位符
((TextBoxShape)cover.Shapes[0]).Paragraphs[0].AddText("2026 年第三季度");
cover.NotesText = "开场 2 分钟。";
cover.Transition = SlideTransitionKind.Fade;
cover.TransitionDuration = 1.5;

Slide free = presentation.AddSlide();                                // 等价于 Blank，全部自己定位
AutoShape box = free.AddShape(AutoShapeKind.RoundedRectangle, 60, 60, 200, 80, "圆角矩形");
box.FillColor = OfficeColor.FromHtml("#00B050");
free.AddImage(OfficeImage.FromFile("flag.png"), 320, 60, 120, 120);
TableShape table = free.AddTable(60, 260, 300, 90);
table.Table.AddRow("模块", "状态");                                   // 复用文档模型的 Table
free.BackgroundColor = OfficeColor.FromHtml("#F2F7FF");
```

版式占位符按幻灯片尺寸**等比推算**：左右边距为宽度的 8%，标题占位符位于高度的 6%、占 20%（`SectionHeader` 例外，下移到 34%、占 26%），内容区从高度的 30% 起。七种版式的实测产出：

| `SlideLayoutKind` | 占位符 |
| --- | --- |
| `Title` | `Title#0` |
| `TitleAndContent` | `Title#0` + `Content#1` |
| `SectionHeader` | 仅 `Title#0` |
| `TwoContent` | `Title#0` + 左右两个 `Content` |
| `TitleAndPicture` | `Title#0` + `Picture#1` |
| `TitleAndTable` | `Title#0` + `Table#1` |
| `Blank` | 无形状 |

几个容易踩的小约定：

- `AddTextBox()` / `AddShape()` / `AddImage()` / `AddTable()` 最后一个字符串参数是**形状文本**，不是形状名称。
- 版式占位符名称由模型自动分配中文（`标题`、`内容占位符 1`），手工添加的形状 `Name` 是空串。
- `Slide.Layout` **不写进文件**，读取时按页内形状组合推断。
- `Table` 的 `Columns` 只是列占位，实测写出时不落库、读回时按每行单元格数补齐。

`TableShape` 内部直接复用文档模型的 `Table`——这就是统一模型的收益：Word 侧学会的表格语义（行列、`ColSpan`/`RowSpan`、`IsHeader`），在 PPT 里完全一样。

## 四、TemplateEngine：只做标量替换

`{{key}}`、`{{key:format}}`，支持嵌套对象点路径（不支持 `[i]` 下标）：

```csharp
TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["order"] = new { id = "SO-1024", total = 398.0 };

string text = TemplateEngine.Replace("{{customerName}} 的订单 {{order.id}} 合计 {{order.total:0.00}} 元", context);
// 示例客户 的订单 SO-1024 合计 398.00 元
```

**`Replace()` 只做标量替换**：遇到 `{{#each list}} … {{/each}}` 时标记原样保留（实测返回 `{{#each items}}{{name}}{{/each}}`），真正的循环展开由 `Wisdom.Word.DocxTemplate` / `Wisdom.Excel.XlsxTemplate` 完成——因为只有它们知道"一行表格行"和"一个段落"该怎么复制。

本包提供的是它们复用的零件：

- 变量注入：`{{@index}}`（从 0）、`{{@number}}`（从 1）、`{{@first}}`、`{{@last}}`、`{{.}}`（当前项本身）
- 公开工具方法：`ContainsToken`、`TryGetLoopName`、`IsLoopEnd`、`RemoveLoopMarkers`、`ResolveList`、`CreateRowContext`、`Merge`、`ConvertToText`

想给自己的文件格式（比如 HTML 模板）写填充逻辑，这几个方法就是全部原料。一个约束：`{{#each}}` 只匹配整格内容形如 `{{#each 名称}}` 的单元格或段落，**标记必须独占一行 / 一格**，和正文写在同一格不会被识别。

`TemplateOptions`：`LeaveUnresolvedTokens`（默认 `true`，未解析占位符原样保留）、`HtmlEncodeValues`、`Culture`。

## 五、基础设施：这些是"自研"两个字的具体含义

零第三方依赖不是口号，它意味着下面这些轮子都得自己造：

| 类型 | 用途 |
| --- | --- |
| `CompoundFile` | OLE2 复合文档（`.doc` / `.xls` 容器）读取：`HasStream`、流枚举 |
| `PackageArchive` | ZIP 包安全打开 |
| `OfficePackageLimits` | 包解压上限：`MaxEntryBytes`（默认 256 MB）、`MaxTotalBytes`（默认 1 GB） |
| `Zlib` / `Adler32` | 自研 DEFLATE 解压，供 PDF 与 XLSX 使用 |
| `LittleEndianReader` / `BigEndianReader` / `ByteUtils` | 二进制读取 |
| `UnitConverter` | 磅 / 像素 / EMU / OOXML 列宽行高之间的换算 |
| `DateUtils`、`SpreadsheetFormat` | 序列日期、数字格式解析 |

`PackageArchive` 值得单独说一句。`.NET Framework 4.6.1` 上的 `ZipArchive` 会延迟解析中央目录，损坏的条目名可能在**完全不同的调用点**抛出 `ArgumentException`——排查起来毫无线索。`PackageArchive.OpenForRead` 立即物化目录，并把这类异常统一收敛为 `InvalidDataException`，让六个目标框架的容错行为一致。

`OfficePackageLimits` 则是针对"用户上传的包解压炸弹"的兜底：单个条目和总量都有上限，超限直接中止而不是把内存吃光。自研解压（`Zlib`）能设这种闸，也是不用第三方压缩库换来的好处之一。

## 已知限制

- **本包只定义模型与基础设施，不产生任何文件**。读写请用 `Wisdom.Word` / `Wisdom.Excel` / `Wisdom.Pdf` / `Wisdom.Ppt` / `Wisdom.Markdown`。
- `Cell.Value` 只接受 `string`、`double`、`bool`、`DateTime`，其它类型需先自行转换。
- 模型层不做分页计算。
- `TextRun.Link` 只被 Markdown / HTML 写出消费。
- `TemplateEngine.Replace()` 不展开 `{{#each}}`。
- `Slide.Layout` 与 `Table.Columns` 不写进 PPTX。

## 小结

`Wisdom.Office.Core` 存在的意义，是把"Office 文档"这件事拆成两个正交问题：**内容长什么样**（模型）和**字节怎么排**（读写器）。三种模型——文档、工作簿、演示文稿——覆盖了办公场景的三种基本介质，剩下六个包各自解决一个文件格式族。

收益在后面几篇文章里都能看到：Word 的表格能直接放进 PPT 的 `TableShape`；Markdown 解析的结果不需要任何新类型；Excel 转 PDF 只是"工作簿 → 文档模型 → 渲染"两步；`Report` 能用一次声明同时产出三种格式。

单独引用它的场合其实不少：做 DSL 生成中间模型、给自己的模板格式写填充器、或者只想要 `UnitConverter` 这类换算工具而不引入任何读写引擎。
