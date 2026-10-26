# Wisdom.Office.Core

`Wisdom.Office.Core` 是整套库的**共享内核**：统一文档模型、工作簿模型、演示文稿模型、模板引擎、OLE2 复合文档读取、ZIP 安全读取和二进制/压缩基础工具。`Wisdom.Pdf`、`Wisdom.Word`、`Wisdom.Excel`、`Wisdom.Ppt`、`Wisdom.Markdown`、`Wisdom.Office` 全部依赖它；只想要模型对象或模板引擎的场合也可以单独引用。

- 包名：`Wisdom.Office.Core`
- 版本：`1.0.2`
- 依赖：无（其余六个包的共同底座）
- 命名空间：`Wisdom.Office.Core`

本文是能力与 API 参考；分步实战教程见源码仓库的 `docs/Wisdom.Office.Core.Tutorial.md`。

## 支持的框架

| 目标框架 | 包内资产 | 典型运行时 |
| --- | --- | --- |
| `net461` | `lib/net461` | .NET Framework 4.6.1 / 4.7.2 / 4.8.x |
| `netstandard2.0` | `lib/netstandard2.0` | .NET Core 2.0-3.1、.NET 5/6、Mono、Unity |
| `net7.0` | `lib/net7.0` | .NET 7 |
| `net8.0` | `lib/net8.0` | .NET 8 |
| `net9.0` | `lib/net9.0` | .NET 9 |
| `net10.0` | `lib/net10.0` | .NET 10 / 11 |

源码无任何 `#if NET*` 条件编译，六份资产功能一致。`net461` 资产需要运行时自带 `System.IO.Compression`（.NET Framework 4.5+ 已内置），无需额外 NuGet 包。

## 安装

```bash
dotnet add package Wisdom.Office.Core --version 1.0.2
```

## 文档模型（Word / PDF 共用）

层级为 `OfficeDocument` → `Section` → `Block`，`Block` 有 `Paragraph`、`Table`、`PageBreak` 三种。

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
TableRow total = table.AddRow("合计", "2,288.40");
total.IsHeader = false;
table.RepeatHeaderRow = true;
section.Blocks.Add(table);
```

常用类型：

| 类型 | 关键成员 |
| --- | --- |
| `OfficeDocument` | `Metadata`、`Sections`、`DefaultTextFormat`、`AddSection()`、`FirstSection` |
| `Section` | `PageSetup`、`Blocks`、`Header`、`Footer`、`AddParagraph()` |
| `PageSetup` | `Width/Height`、四边 `Margin*`、`HeaderDistance`、`FooterDistance`、`IsLandscape`、`A4()`、`Letter()` |
| `HeaderFooter` | `Blocks`、`IsEnabled`、`AddParagraph()` |
| `Paragraph` | `Inlines`、`Alignment`、`SpaceBefore/After`、`LineSpacing`、缩进、`HeadingLevel`、`KeepWithNext`、`PageBreakBefore`、`ListMarker`、`AddText()`、`AddImage()`、`AddLineBreak()`、`GetText()` |
| `TextRun` / `FieldRun` / `ImageRun` | 文本、超链接 `Link`、域（`PAGE`、`NUMPAGES` 等）、行内图片 |
| `TextFormat` | `FontFamily`、`AsciiFontFamily`、`EastAsiaFontFamily`、`FontSize`、`Bold/Italic/Underline/Strikeout`、`Color`、上下标、`Heading(level)` |
| `Table` / `TableRow` / `TableCell` | 列宽、行 `IsHeader`/`Height`、单元格 `ColSpan`/`RowSpan`/`BackgroundColor`/`VerticalAlignment`/`Borders`、`CellPadding`、`Autofit`、`Alignment` |
| `TableBorders` / `TableCellBorders` / `BorderLine` | `Thin()`、`None()`、`BorderStyle` 枚举 |
| `OfficeColor` | `FromHtml("#00B050")`、`FromArgb`、`ToHex`、`TryParse`、预置色 |
| `OfficeImage` | `FromFile()`、`Data`、`Format`、`PixelWidth/Height`、`DefaultWidthPoints` |

页码域的写法（Word 输出真实 `fldSimple`，PDF 输出实际页码）：

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

### Markdown 在这套模型里长什么样

`Wisdom.Markdown` 解析出的结构不引入新类型，全部落在上面的文档模型上；反过来，任何文档也能回写成 Markdown：

| Markdown 语法 | 模型字段（实测） |
| --- | --- |
| `# 标题` | `Paragraph.HeadingLevel = 1`（`######` 到 6） |
| `[官网](https://example.com)` | 该 `TextRun.Link = "https://example.com"`，加粗仍是 `TextFormat.Bold` |
| `- 项` | `Paragraph.ListMarker = "•"` |
| `- [x] 项` / `- [ ] 项` | `ListMarker = "☑ •"` / `"☐ •"` |
| GFM 表格（含 `| --- | --- |`） | `Table`，分隔行那一 `TableRow.IsHeader = true` |

```csharp
// 本包只提供字段载体；解析与回写在 Wisdom.Markdown（Markdown.ToDocument / ToMarkdown / ToHtml）。
Paragraph heading = new Paragraph("季度经营回顾") { HeadingLevel = 1 };  // "# 季度经营回顾"
Paragraph task = new Paragraph { ListMarker = "☑ •" };                   // "- [x] 支持 Markdown 解析"
task.AddText("支持 Markdown 解析");
TextRun link = new TextRun("官网") { Link = "https://example.com" };     // "[官网](https://example.com)"
```

`Link` 只有 Markdown / HTML 写出会消费：DOCX 写出时不生成 `w:hyperlink` 关系，PDF 也不写链接注释，两头都只保留文字。`HeadingLevel` 还是文档转演示文稿的分页依据——每遇一级标题起新的一页。

## 工作簿模型（Excel）

```csharp
OfficeWorkbook workbook = new OfficeWorkbook();
workbook.Metadata.Title = "销售数据";

Worksheet sheet = workbook.AddWorksheet("明细");
sheet.GetCell(1, 1).Value = "产品";
sheet.GetCell(2, 2).Value = 12;                 // Cell.Value 支持 string / double / bool / DateTime
sheet.GetCell(2, 3).Style = new CellStyle { NumberFormat = "#,##0.00", Bold = true };
sheet.Merge("A4:C4");
sheet.SetColumnWidth(1, 24);
sheet.FrozenRows = 1;
sheet.SetPrintArea("A1:D20");
```

`Worksheet` 还承载装饰对象：`Charts`、`Images`、`ConditionalFormats`；表格级属性含 `MergedRanges`、`ColumnWidths`、`RowHeights`、`HiddenRows/Columns`、`IsHidden`、`ShowGridLines`、`ShowZeros`、`FrozenRows/Columns`、`PrintSetup`、`FitToPage`、`DefaultStyle`。单元格地址工具为 `CellAddress` 与 `SpreadsheetRange`。

图表 / 条件格式 / 图片的完整用法见 `Wisdom.Excel` 的说明文档——模型定义在本包，落盘与回读由 XLSX 实现。

## 演示文稿模型（PowerPoint）

层级为 `OfficePresentation` → `Slide` → `SlideShape`，形状派生出 `TextBoxShape`、`AutoShape`、`PictureShape`、`TableShape`。所有几何量以 **磅** 为单位，原点为幻灯片左上角。

```csharp
using Wisdom.Office.Core;

OfficePresentation presentation = OfficePresentation.Widescreen();   // 960 x 540
presentation.Metadata.Title = "季度经营回顾";

Slide cover = presentation.AddSlide(SlideLayoutKind.Title);          // 版式自动生成占位符
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

Slide free = presentation.AddSlide();                                // 等价于 Blank，全部自己定位
AutoShape box = free.AddShape(AutoShapeKind.RoundedRectangle, 60, 60, 200, 80, "圆角矩形");
box.FillColor = OfficeColor.FromHtml("#00B050");
free.AddImage(OfficeImage.FromFile("flag.png"), 320, 60, 120, 120);
TableShape table = free.AddTable(60, 260, 300, 90);
table.Table.AddRow("模块", "状态");
free.BackgroundColor = OfficeColor.FromHtml("#F2F7FF");
```

`AddSlide()` 不带参数得到空白页；带 `SlideLayoutKind` 时按幻灯片尺寸等比推算占位符：左右边距为宽度的 8%，标题占位符位于高度的 6%、占 20%（`SectionHeader` 例外，下移到 34%、占 26%），内容区从高度的 30% 起。七种版式生成的占位符实测如下：

| `SlideLayoutKind` | 占位符（名称为模型自动分配的中文） |
| --- | --- |
| `Title` | `Title#0` |
| `TitleAndContent` | `Title#0` + `Content#1` |
| `SectionHeader` | 仅 `Title#0` |
| `TwoContent` | `Title#0` + `Content#1` + `Content#2`（左右分栏） |
| `TitleAndPicture` | `Title#0` + `Picture#1` |
| `TitleAndTable` | `Title#0` + `Table#1` |
| `Blank` | 无形状 |

`AddTextBox()` / `AddShape()` / `AddImage()` / `AddTable()` 的最后一个字符串参数是形状文本，不是形状名称；版式占位符的名称由模型自动分配（`标题`、`内容占位符 1` 等中文），手工添加的形状 `Name` 是空串。版式（`Slide.Layout`）不写进文件，读取时按页内形状组合推断。

| 类型 | 关键成员 |
| --- | --- |
| `OfficePresentation` | `Metadata`、`SlideWidth/SlideHeight`、`Slides`、`DefaultTextFormat`（Calibri 18）、`DefaultTitleFormat`（Calibri 40 粗体）、`AccentColors`、`Theme`（`PresentationTheme`，null 为默认主题）、`Widescreen()`、`Standard()`（720 x 540）、`AddSlide()` |
| `Slide` | `Layout`、`BackgroundColor`、`NotesText`、`Transition`、`TransitionDuration`、`AdvanceAfterSeconds`（`advTm` 自动换页）、`Shapes`、`AddTextBox()`、`AddShape()`、`AddImage()`、`AddTable()`、`GetTitle()` |
| `SlideShape` | `Id`、`Name`、`Left/Top/Width/Height`、`Placeholder`、`PlaceholderIndex`、`FillColor`、`Outline`、`Animation`（`SlideAnimationKind`）、`AnimationDelaySeconds`、`AnimationDurationSeconds` |
| `TextBoxShape` | `Paragraphs`、`VerticalAlignment`（`SlideTextAnchor`）、四边 `Margin*`、`ShrinkToFit`、`AddParagraph()`、`GetText()` |
| `AutoShape` | `Kind`（`AutoShapeKind`：矩形 / 圆角矩形 / 椭圆 / 三角 / 菱形 / 五边 / 六边 / 右箭头 / 燕尾 / 直线）、`CornerRadius`（默认 0.15） |
| `PictureShape` | `Image`（`OfficeImage`）、`AlternativeText` |
| `TableShape` | `Table`（复用文档模型的 `Table`）、`HeaderColor` |
| `SlideLayoutKind` / `PlaceholderKind` / `SlideTransitionKind` / `SlideAnimationKind` | 版式、占位符类别（含 `Date` / `Footer` / `SlideNumber` / `Chart`）、切换效果（`None` / `Fade` / `Push` / `Wipe` / `Split`）、入场动画（`None` / `Fade` / `FlyIn` / `Wipe` / `Zoom` / `Split`） |
| `PresentationTheme` | 主题深改：`Name`、`Text` / `Background`（dk1 / lt1）、`TextSecondary` / `BackgroundSecondary`（dk2 / lt2）、`Hyperlink` / `FollowedHyperlink`、`FontMajor` / `FontMinor`，null 字段保持内置默认 |

模型本身不产生文件；读写由 `Wisdom.Ppt` 的 `PptxReader` / `PptxWriter` 负责，导出 PDF / Word 由 `Wisdom.Office` 的 `PresentationToDocumentConverter` 按"每页一节"排版。`Table` 的 `Columns` 只是列占位，实测写出时不落库、读回时按每行单元格数补齐。

## 模板引擎

占位符语法 `{{key}}`、`{{key:format}}`，支持嵌套对象的点路径（不支持 `[i]` 下标）。`Replace()` **只做标量替换**：遇到 `{{#each list}} … {{/each}}` 时标记原样保留（实测返回 `{{#each items}}{{name}}{{/each}}`），展开由 `Wisdom.Word.DocxTemplate` / `Wisdom.Excel.XlsxTemplate` 完成，本包提供它们复用的循环工具与变量注入：`{{@index}}`（从 0）、`{{@number}}`（从 1）、`{{@first}}`、`{{@last}}`、`{{.}}`（当前项本身）。

```csharp
TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["order"] = new { id = "SO-1024", total = 398.0 };
context["items"] = new[]
{
    new { name = "无线键盘", qty = 2, price = 199.0 },
    new { name = "机械鼠标", qty = 5, price = 89.9 }
};

// 点路径取嵌套成员；{{key:format}} 用 IFormattable 格式串
string text = TemplateEngine.Replace("{{customerName}} 的订单 {{order.id}} 合计 {{order.total:0.00}} 元", context);
// 示例客户 的订单 SO-1024 合计 398.00 元
```

`TemplateOptions`：`LeaveUnresolvedTokens`（默认 `true`，未解析的占位符原样保留；置 `false` 才替换成空串）、`HtmlEncodeValues`、`Culture`。`TemplateEngine` 另外公开 `ContainsToken`、`TryGetLoopName`、`IsLoopEnd`、`RemoveLoopMarkers`、`ResolveList`、`CreateRowContext`、`Merge`、`ConvertToText`，供自定义填充逻辑复用——这几个方法就是循环展开的全部零件，`{{#each}}` 只匹配整格内容形如 `{{#each 名称}}` 的单元格或段落，标记必须独占一行 / 一格，和正文写在同一格不会被识别。

## 基础设施

| 类型 | 用途 |
| --- | --- |
| `CompoundFile` | OLE2 复合文档（`.doc`/`.xls` 容器）读取：`HasStream`、流枚举 |
| `PackageArchive` | ZIP 包安全打开：立即物化中央目录，把 .NET Framework 上延迟抛出的 `ArgumentException` 统一转为 `InvalidDataException` |
| `OfficePackageLimits` | 包解压上限：`MaxEntryBytes`（默认 256 MB）、`MaxTotalBytes`（默认 1 GB）、`ReadEntryBytes`、`ReadAllBytes` / `CopyBounded`（按上限分块复制，不凭部件声明的长度预分配内存） |
| `OfficeXmlText` | 写出前的 XML 文本防护：`Strip(string)` 丢掉 XML 1.0 不接受的字符（`0x00`–`0x08`、未配对代理项等），`Sanitize(XDocument)` 就地清理整棵树；文本干净时返回原实例，不额外分配 |
| `Zlib` / `Adler32` | 自研 DEFLATE 解压，供 PDF 与 XLSX 使用，不依赖第三方压缩库 |
| `LittleEndianReader` / `BigEndianReader` / `ByteUtils` | 二进制读取工具 |
| `UnitConverter` | 磅 / 像素 / EMU / OOXML 列宽行高之间的换算 |
| `DateUtils`、`SpreadsheetFormat` | 序列日期、数字格式解析 |

## 已知限制

- 本包只定义模型与基础设施，不产生任何文件；读写请使用 `Wisdom.Word` / `Wisdom.Excel` / `Wisdom.Pdf` / `Wisdom.Ppt` / `Wisdom.Markdown`。
- `Cell.Value` 只接受 `string`、`double`、`bool`、`DateTime`，其它类型需先自行转换。
- 模型层不做分页计算；`{page}` / `{pages}` 的实际替换发生在 PDF 渲染阶段。
- `TextRun.Link` 只被 Markdown / HTML 写出消费，DOCX 与 PDF 只保留链接文字，不生成超链接关系或注释。
- `TemplateEngine.Replace()` 不展开 `{{#each}}`，循环行 / 列的复制由 `DocxTemplate` / `XlsxTemplate` 负责。
- 演示文稿的 `Slide.Layout` 与 `Table.Columns` 不写进 PPTX：前者读取时按形状推断，后者读回时按每行单元格数补齐。
