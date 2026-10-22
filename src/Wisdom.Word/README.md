# Wisdom.Word

`Wisdom.Word` 负责 **Word 侧的全部读写**：DOCX/DOTX/RTF/HTML/ODT/纯文本的读取与生成、DOCX 生成与模板填充、旧版 `.doc` 读取、流式文档构建、邮件合并、查找替换、文档克隆与合并、自定义文档属性、标签纸与信封生成，以及建立在统一文档模型之上的修订跟踪、SmartArt/数字签名检查、文档整理、一致性校验、差异比对与字数统计引擎。

- 包名：`Wisdom.Word`
- 版本：`1.0.1`
- 依赖：`Wisdom.Office.Core`（唯一依赖）
- 第三方依赖：无；不要求安装 Microsoft Office 或 WPS
- 命名空间：`Wisdom.Word`

本文是能力与 API 参考；分步实战教程见源码仓库的 `docs/Wisdom.Word.Tutorial.md`。

## 支持的框架

| 目标框架 | 包内资产 | 典型运行时 |
| --- | --- | --- |
| `net461` | `lib/net461` | .NET Framework 4.6.1 / 4.7.2 / 4.8.x |
| `netstandard2.0` | `lib/netstandard2.0` | .NET Core 2.0-3.1、.NET 5/6、Mono、Unity |
| `net7.0` | `lib/net7.0` | .NET 7 |
| `net8.0` | `lib/net8.0` | .NET 8 |
| `net9.0` | `lib/net9.0` | .NET 9 |
| `net10.0` | `lib/net10.0` | .NET 10 / 11 |

DOCX 由 `System.IO.Compression.ZipArchive` + LINQ to XML 直接处理，六份资产由同一份源码编译，功能一致。

## 安装

```bash
dotnet add package Wisdom.Word --version 1.0.1
```

需要把 Word 转成 PDF 时请改引 `Wisdom.Office` 门面包（本包不依赖 `Wisdom.Pdf`）。

## 读取 DOCX

```csharp
using Wisdom.Word;
using Wisdom.Office.Core;

OfficeDocument document = DocxReader.Load("report.docx");   // 亦支持 Stream / byte[]

Console.WriteLine(document.Metadata.Title);
foreach (Section section in document.Sections)
{
    foreach (Block block in section.Blocks)
    {
        if (block is Paragraph paragraph) Console.WriteLine(paragraph.GetText());
        if (block is Table table) Console.WriteLine($"{table.Rows.Count} 行 x {table.Columns.Count} 列");
    }
}

string text = DocxReader.ExtractText("report.docx");   // 只要纯文本时用这个
```

读取覆盖段落与行内格式、表格（含合并单元格 `ColSpan`/`RowSpan`）、页眉页脚、图片、超链接（外部 URL 与 `#书签` 锚点）、书签、`HeadingLevel`、对齐与缩进、文档元数据。

## 生成 DOCX

```csharp
using Wisdom.Word;
using Wisdom.Office.Core;

OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "季度总结";

Section section = document.AddSection(PageSetup.A4());

// 页眉：右对齐的文件编号
section.Header.AddParagraph("文件编号：QF-QC-018").Alignment = ParagraphAlignment.Right;

// 页脚：真实页码域，Word 打开时自动计算
Paragraph footer = section.Footer.AddParagraph(string.Empty);
TextFormat small = new TextFormat("SimSun", 8);
footer.AddText("第 ", small);
footer.Inlines.Add(new FieldRun("PAGE", "{page}", small));
footer.AddText(" 页　共 ", small);
footer.Inlines.Add(new FieldRun("NUMPAGES", "{pages}", small));
footer.AddText(" 页", small);

// 正文
Paragraph title = new Paragraph("成品检验报告");
title.HeadingLevel = 1;
title.Alignment = ParagraphAlignment.Center;
section.Blocks.Add(title);

// 表格：跨列跨行 + 表头底纹 + 重复表头
Table table = new Table();
table.Columns.Add(new TableColumn(28));
table.Columns.Add(new TableColumn(84));
table.Columns.Add(new TableColumn(104));
table.RepeatHeaderRow = true;

TableRow header = table.AddRow();
header.IsHeader = true;
AddSpanCell(header, "序号", 1, 2, OfficeColor.LightGray);
AddSpanCell(header, "检验项目", 1, 2, OfficeColor.LightGray);
AddSpanCell(header, "技术要求", 1, 2, OfficeColor.LightGray);
AddSpanCell(header, "检验方法", 1, 2, OfficeColor.LightGray);
AddSpanCell(header, "实测结果", 2, 1, OfficeColor.LightGray);   // 跨两列
AddSpanCell(header, "判定", 1, 2, OfficeColor.LightGray);

TableRow sub = table.AddRow();                                   // 表头第二行：两个单件列
sub.Cells.Add(new TableCell("单件一"));
sub.Cells.Add(new TableCell("单件二"));

foreach (string[] item in items)
{
    TableRow row = table.AddRow();
    row.Height = 24;
    foreach (string value in item) row.Cells.Add(new TableCell(value));
}
section.Blocks.Add(table);

DocxWriter.Write(document, "inspection.docx");     // 亦支持 Stream，或 byte[] Write(document)
```

合并单元格辅助写法：

```csharp
static void AddSpanCell(TableRow row, string text, int colSpan, int rowSpan, OfficeColor background)
{
    TableCell cell = new TableCell();
    cell.ColSpan = colSpan;
    cell.RowSpan = rowSpan;
    cell.VerticalAlignment = VerticalAlignment.Middle;
    cell.BackgroundColor = background;
    cell.Blocks.Add(new Paragraph(text));
    row.Cells.Add(cell);
}
```

生成的包通过内容类型、关系目标与 XML 合法性校验，可被 Word 2010+、WPS 直接打开。

## 流式构建：WordDocumentBuilder

`WordDocumentBuilder` 参照 Aspose.Words 的 `DocumentBuilder`：维护写入游标与当前字符/段落格式，按顺序追加内容，适合"拼接报告"这类过程式生成。

```csharp
using Wisdom.Word;
using Wisdom.Office.Core;

WordDocumentBuilder builder = new WordDocumentBuilder();
builder.Font.FontFamily = "Microsoft YaHei";
builder.Font.EastAsiaFontFamily = "Microsoft YaHei";

builder.HeadingLevel = 1;
builder.WriteLine("2026 年度报告");
builder.HeadingLevel = 0;

builder.InsertParagraph("正文第一段。");
builder.InsertBookmark("Intro");
builder.InsertHyperlink("访问官网", "https://example.com/report");   // 外部 URL
builder.InsertHyperlink("回到引言", "#Intro");                        // 文档内锚点

builder.StartTable();
builder.InsertCell("项目");
builder.InsertCell("金额");
builder.EndRow();
builder.InsertCell("合计");
builder.InsertCell("100");
builder.EndTable();

builder.InsertImage(new OfficeImage(File.ReadAllBytes("logo.png")), 120, 40);
builder.InsertTableOfContents(3);      // TOC 域，Word 打开后按 F9 更新
builder.InsertField("PAGE", "1");      // 页码域

DocxWriter.Write(builder.Document, "report.docx");
```

可设置 `Alignment`、`SpaceBefore/SpaceAfter`、`LineSpacing`、各向缩进、`HeadingLevel`、`KeepWithNext`、`PageBreakBefore`；`Write` 追加文字，`InsertXxx` 系列插入段落、图片、书签、超链接、域、目录与表格。外部超链接会生成 `TargetMode="External"` 关系，锚点写成 `w:anchor`，书签写成 `w:bookmarkStart/End`。

## 书签

两种形态：覆盖整段的书签（沿用 `Paragraph.BookmarkName`，API 兼容旧版本）和任意位置的范围书签（`BookmarkStart`/`BookmarkEnd` 两个行内标记，可落在段落中间、跨段落、甚至跨表格单元格）。`document.Bookmarks` 每次访问即时配对，读改存之间不需要刷新。

```csharp
using Wisdom.Word;
using Wisdom.Office.Core;

WordDocumentBuilder builder = new WordDocumentBuilder();
builder.InsertParagraph("引言");
builder.InsertBookmark("Intro");                 // 整段书签

builder.InsertParagraph();
builder.Write("合计 ");
builder.StartBookmark("Total");                  // 范围书签：只框住"128 万元"
builder.Write("128 万元");
builder.EndBookmark("Total");
builder.Write("（含税）");

OfficeDocument document = builder.Document;
document.Bookmarks["Total"].Text = "130 万元";   // 就地替换范围内正文，书签保留
Bookmark intro = document.Bookmarks["Intro"];
Console.WriteLine(intro.IsWholeParagraph);       // True
Console.WriteLine(intro.Text);                   // 引言

// 按名称或文档顺序遍历，也可以移动光标就地续写
foreach (Bookmark bookmark in document.Bookmarks)
{
    Console.WriteLine(bookmark.Name + " -> " + bookmark.Text);
}
builder.MoveToBookmark("Total", false, true);    // 定位到终点之后
builder.Write("，已复核");
document.Bookmarks.Remove("Total");              // 只删标记，不删正文

string safe = BookmarkNames.Validate(" Total ");     // 去空白并校验：1-40 字符、字母或下划线开头、允许中文
DocxWriter.Write(document, "report.docx");
```

DOCX 里 `w:bookmarkStart` 带 `w:id` 与 `w:name`，`w:bookmarkEnd` 只带 `w:id`；读取端按 id 反查名称配对，写出端复用起点的 id，因此起点未落笔时不会留下无主 `w:bookmarkEnd`。id 在每个部件内唯一（正文、页眉、页脚各自计数），Word 内部书签 `_GoBack` 会被忽略。RTF（`\bkmkstart`/`\bkmkend`）与 ODT（`text:bookmark-start`/`text:bookmark-end`）同样保留范围；HTML 只有点锚，范围读取后退化为整段书签。

## 查找替换

```csharp
// 最简单：替换全部（不区分大小写），跨 Run 命中，替换文字继承首个命中 Run 的格式
int count = WordFindReplace.Replace(document, "旧公司名", "新公司名");

WordFindReplaceOptions options = new WordFindReplaceOptions
{
    MatchCase = false,               // 区分大小写
    WholeWord = true,                // 全词匹配
    FirstOnly = false,               // 只替换第一处
    IncludeHeadersAndFooters = true, // 搜索页眉页脚
    IncludeTables = true             // 搜索表格单元格
};
count = WordFindReplace.Replace(document, "公司", "集团", options);

// 正则：内置 2 秒匹配超时，避免灾难性回溯
WordFindReplaceOptions regex = new WordFindReplaceOptions { UseRegex = true };
count = WordFindReplace.Replace(document, @"\d{4}年", "2026年", regex);
```

## 邮件合并

`WordMailMerge` 在统一模型上填充 Word 的 `MERGEFIELD` 域与 `{{token}}` / `{{token:format}}` 标记；数据可来自字符串字典、`DataRow` 或 `TemplateContext`，支持点路径、格式串、页眉页脚与图片值。`ExecuteWithRegions` 另外支持 `TableStart` / `TableEnd` 区域与任意深度嵌套。

模板里可以直接用 Word 的"插入域 → MergeField"，也可以写双花括号标记：

```text
MERGEFIELD customer.name
MERGEFIELD Balance \# "#,##0.00"
{{amount:N2}}
```

```csharp
using Wisdom.Word;
using Wisdom.Office.Core;

TemplateContext context = new TemplateContext();
context["customer"] = new { name = "示例客户" };            // 点路径 customer.name
context["amount"] = 1234.5m;
context["Balance"] = 9876.5m;
context["Logo"] = new OfficeImage(File.ReadAllBytes("logo.png"));

int merged = WordMailMerge.Execute(document, context, new WordMailMergeOptions
{
    LeaveUnresolvedFields = false,       // 未解析域/标记置空；默认 true 原样保留
    Culture = CultureInfo.InvariantCulture,
    IncludeHeadersAndFooters = true
});

WordMailMerge.Execute(document, dataRow);                     // DataRow：列名为字段名，DBNull 视为空
WordMailMerge.Execute(document, new Dictionary<string, object> { ["title"] = "季度报告" });
```

合并后的 `MERGEFIELD` 会转成普通文本或图文，避免 Word 更新域时把结果清空。图片值只接受 `OfficeImage` 或图片 `byte[]`，不接受文件路径；图片请放在 `MERGEFIELD` 位置，文本标记中间不能嵌图片。

### 区域与嵌套区域

在表格行或普通段落中放一对区域标记，区域内容会按集合逐项重复：

```text
MERGEFIELD TableStart:orders                  ← 表格首行第一格
{{orderNo}}        {{customer}}
MERGEFIELD TableStart:lines                   ← 子区域起点
{{sku}}            {{qty}}
MERGEFIELD TableEnd:lines
MERGEFIELD TableEnd:orders                    ← 表格末行
```

```csharp
TemplateContext root = new TemplateContext();
root["title"] = "年度报价";
root["orders"] = new[]
{
    new
    {
        orderNo = "A-001",
        customer = "海森公司",
        lines = new[] { new { sku = "键盘", qty = 2 }, new { sku = "鼠标", qty = 4 } }
    }
};

int merged = WordMailMerge.ExecuteWithRegions(document, root);
```

规则：

- 区域名对应数据上下文中的集合键；`DataTable` / `DataSet` 用表名作为区域名，表间 `DataRelation` 自动成为嵌套区域。
- 表格区域从 `TableStart` 行到 `TableEnd` 行整体作为模板，逐项克隆；仅包含标记、没有其它内容的首尾行会被移除，`IsHeader` 行只输出一次。
- 段落区域从起始标记段落到结束标记段落整体重复；区域可放在表格单元格内。
- 嵌套区域的字段先按子行解析，再回退到父行/全局上下文；子行同名键覆盖父行。
- 集合为空或未提供时整个区域被移除；区域嵌套深度上限为 32，防止自引用数据无限展开。
- 同一区域名为同层标记，不支持同名区域自我嵌套；嵌套请使用不同的区域名。

```csharp
// DataTable：表名 = 区域名；DataRelation 的子表自动成为嵌套区域
WordMailMerge.ExecuteWithRegions(document, ordersTable);
WordMailMerge.ExecuteWithRegions(document, dataSet);      // 每个表按表名注册为顶层区域

// 字典：集合值对应区域，标量作为全局值
WordMailMerge.ExecuteWithRegions(document, new Dictionary<string, object>
{
    ["title"] = "季度报告",
    ["orders"] = orderList
});
```

## 自定义文档属性（文档标签）

Word 的“自定义属性”（文件 → 信息 → 属性 → 高级属性 → 自定义）保存在 `docProps/custom.xml`。`DocumentMetadata.CustomProperties` 提供强类型读写：名称不区分大小写，枚举时保持插入顺序。

```csharp
OfficeDocument document = new OfficeDocument();
document.AddSection().AddParagraph("正文");

CustomPropertyCollection properties = document.Metadata.CustomProperties;
properties.Set("项目编号", "QF-2026-018");
properties.Set("版本", 3);
properties.Set("文件大小", 4294967296L);
properties.Set("已审核", true);
properties.Set("抽检比例", 0.15);
properties.Set("金额", 1234.56m);
properties.Set("归档时间", DateTime.UtcNow);

DocxWriter.Write(document, "tagged.docx");

OfficeDocument loaded = DocxReader.Load("tagged.docx");
Console.WriteLine(loaded.Metadata.CustomProperties.GetString("项目编号"));
Console.WriteLine(loaded.Metadata.CustomProperties.GetInt32("版本"));

object value;
if (loaded.Metadata.CustomProperties.TryGetValue("已审核", out value))
{
    Console.WriteLine(value is bool && (bool)value);       // true
}
```

| CLR 类型 | OOXML vt 类型 |
| --- | --- |
| `string` | `vt:lpwstr` |
| `bool` | `vt:bool` |
| `int` | `vt:i4` |
| `long` | `vt:i8` |
| `double` | `vt:r8` |
| `decimal` | `vt:decimal` |
| `DateTime` | `vt:filetime`（按 UTC 写出） |

读取时兼容 `lpstr` / `bstr` / `ui4` / `r4` 等常见变体；无法识别的 vt 类型按文本保留。`WordDocumentOperations.Clone` 会深拷贝自定义属性，修改副本不影响原件。

## 标签纸与信封

### 标签纸

`WordLabels` 把多行文本或数据记录排成标准标签纸，内置 Avery 5160（US Letter 3×10）、Avery 5163（US Letter 2×5）与 Avery L7160（A4 3×7），也可用 `WordLabelLayout.FromMillimeters` 自定义。

```csharp
using Wisdom.Word;

// 已排好的多行文本：每个 string[] 是一张标签
OfficeDocument sheet = WordLabels.Create(
    WordLabelLayout.AveryL7160(),
    new[]
    {
        new[] { "张三", "上海市浦东新区世纪大道 100 号", "200120" },
        new[] { "李四", "北京市朝阳区建国路 8 号", "100022" }
    },
    new WordLabelOptions { FontSize = 10.5, Bold = false });

DocxWriter.Write(sheet, "labels.docx");

// 数据记录 + 多行模板：按行解析 {{字段}}，超出每页容量自动分页
OfficeDocument fromData = WordLabels.CreateFromRecords(
    WordLabelLayout.Avery5163(),
    records,                                   // 任意对象、字典、DataRow 或 DataTable 行
    "{{Name}}\n{{Address}}\n{{Zip}}");

// 空白标签纸，用于手工套打
OfficeDocument blank = WordLabels.CreateBlank(WordLabelLayout.Avery5160());
```

生成结果为每页一张无边框表格：标签列之间留出版式间隙，页与页之间插入分页符。标签行高按版式设为“最小值”，内容超出单张高度时行高会扩展，后续行随之下移。

自定义版式（毫米）：

```csharp
WordLabelLayout custom = WordLabelLayout.FromMillimeters(
    name: "63.5×38.1 三联",
    pageWidthMm: 210, pageHeightMm: 297,
    columns: 3, rows: 7,
    labelWidthMm: 63.5, labelHeightMm: 38.1,
    topMarginMm: 15.1, leftMarginMm: 7,
    horizontalGapMm: 2.5, verticalGapMm: 0);
```

### 表单标签纸

学生信息标签这类需要手写的标签用 `CreateForm`：每张标签自动输出“字段名 + 填写横线”，横线延伸到标签框右侧并留出 6 磅空白，可为每张标签加细边框。

```csharp
WordLabelLayout studentLayout = WordLabelLayout.FromMillimeters(
    name: "学生信息标签 A4 2×5",
    pageWidthMm: 210, pageHeightMm: 297,
    columns: 2, rows: 5,
    labelWidthMm: 82, labelHeightMm: 48,
    topMarginMm: 25, leftMarginMm: 22,
    horizontalGapMm: 4, verticalGapMm: 4);

OfficeDocument formLabels = WordLabels.CreateForm(
    studentLayout,
    new[] { "姓名：", "年级：", "学号：" },
    new WordLabelOptions
    {
        FontSize = 10.5,
        ShowCellBorders = true,     // 标签外框；邮件标签保持默认 false
        FieldSpacing = 6,           // 相邻填写线之间的距离
        BorderWidth = 0.5
    });
DocxWriter.Write(formLabels, "student-labels.docx");
```

填写行由一张三列嵌套表格实现：左列放字段名，中列是只有下边框的空白单元格，右列留 6 磅空白让横线不贴住边框；三列都底对齐，横线紧贴文字下方。表格外框就是标签边框，表尾用空白行补足标签高度，因此标签之间的垂直间隙留在框外，框与框不会粘连。`FieldSpacing` 控制行距，`BorderColor` / `BorderWidth` 控制外框与填写线的颜色和粗细，版式的 `CellPadding` 控制内容与标签边缘的距离。

### 分格标签纸

留样标签、样品签这类“每行文字之间画一条横线”的标签，把 `ShowRowSeparators` 打开即可：每张标签的正文放进一列多行的嵌套表格，各行等高、行间画满分隔线；同时打开 `ShowCellBorders` 时，外框与内线组成一张完整的表格。同一页内所有标签取该页最大文本行数作为行数，格子上下对齐，空白行也保留边框。

```csharp
WordLabelLayout retentionLayout = WordLabelLayout.FromMillimeters(
    name: "留样标签 A4 4×10",
    pageWidthMm: 210, pageHeightMm: 297,
    columns: 4, rows: 10,
    labelWidthMm: 45, labelHeightMm: 25,
    topMarginMm: 12, leftMarginMm: 10.5,
    horizontalGapMm: 3, verticalGapMm: 2);

string[] retentionLines =
{
    "201  年  月  日  点  分",
    "食品名：          （    克）",
    "留样人：",
    "审核人："
};
List<string[]> retentionLabels = new List<string[]>();
for (int i = 0; i < 40; i++) retentionLabels.Add(retentionLines);

OfficeDocument retentionSheet = WordLabels.Create(
    retentionLayout,
    retentionLabels,
    new WordLabelOptions
    {
        FontFamily = "Arial",          // 西文用 Arial，空格更窄，长行不会被挤断
        EastAsiaFontFamily = "SimSun", // 中文用宋体
        FontSize = 9,
        ShowCellBorders = true,      // 标签外框
        ShowRowSeparators = true,    // 四行之间的横线
        BorderWidth = 0.75
    });
DocxWriter.Write(retentionSheet, "retention-labels.docx");
```

分格标签的内部只有横线、没有竖线；行高按标签高度均分，文本顶部对齐。每个标签单元格在嵌套表格后保留一个极小空段落，避免 Word 自动补段把行高撑开，因此页面的下边距会预留 18 磅。

### 表格版式标签

化学品安全标签这类带合并单元格、警示图形和底部联系信息的版式，直接用 `Table` 排版：`RowSpan` / `ColSpan` 做合并，`VerticalAlignment` 控制内容垂直居中，单元格内可放段落、图片或嵌套表格。下面是一张 A5 横向标签的骨架：

```csharp
PageSetup page = new PageSetup();
page.Width = 210.0 * 72 / 25.4;              // A5 横向
page.Height = 148.0 * 72 / 25.4;
page.MarginLeft = page.MarginRight = 6.0 * 72 / 25.4;
page.MarginTop = page.MarginBottom = 6.0 * 72 / 25.4;

OfficeDocument label = new OfficeDocument();
Section section = label.AddSection(page);

BorderLine line = new BorderLine(BorderStyle.Single, 1.0, OfficeColor.Black);
TableBorders box = new TableBorders
{
    Top = line, Bottom = line, Left = line, Right = line,
    InsideHorizontal = line, InsideVertical = line
};
Table table = new Table { Borders = box, CellPadding = 3, Autofit = false };
table.Columns.Add(new TableColumn(38.0 * 72 / 25.4));
table.Columns.Add(new TableColumn(120.0 * 72 / 25.4));
table.Columns.Add(new TableColumn(40.0 * 72 / 25.4));
section.Blocks.Add(table);

TableRow header = table.AddRow();
header.Height = 26.0 * 72 / 25.4;

TableCell purity = new TableCell { RowSpan = 2, VerticalAlignment = VerticalAlignment.Middle };
Paragraph purityText = new Paragraph { Alignment = ParagraphAlignment.Center };
purityText.AddText("纯度：");
purityText.AddLineBreak();
purityText.AddText("≥30%", new TextFormat("Arial", 11) { Bold = true });
purity.Blocks.Add(purityText);
header.Cells.Add(purity);

TableCell title = new TableCell { VerticalAlignment = VerticalAlignment.Middle };
title.AddParagraph("盐酸").Alignment = ParagraphAlignment.Center;
header.Cells.Add(title);

TableCell pictogram = new TableCell { RowSpan = 4, VerticalAlignment = VerticalAlignment.Middle };
Paragraph imageParagraph = new Paragraph { Alignment = ParagraphAlignment.Center };
imageParagraph.AddImage(new OfficeImage(pngBytes), 34.0 * 72 / 25.4, 34.0 * 72 / 25.4);
pictogram.Blocks.Add(imageParagraph);
header.Cells.Add(pictogram);

// 危险行：纯度与图形仍在跨行，这里只放中间一格
TableRow danger = table.AddRow();
danger.Height = 14.0 * 72 / 25.4;
danger.Cells.Add(new TableCell("危  险"));

// 底部一行：左两列合并，内部再放无边框两列联系表；右列放 UN/CN 编号
TableRow bottom = table.AddRow();
bottom.Height = 8.0 * 72 / 25.4;
TableCell contact = new TableCell { ColSpan = 2, VerticalAlignment = VerticalAlignment.Middle };
Table contactTable = new Table
{
    Borders = TableBorders.None(),
    CellPadding = 0,
    Autofit = false,
    RepeatHeaderRow = false
};
contactTable.Columns.Add(new TableColumn(92.0 * 72 / 25.4));
contactTable.Columns.Add(new TableColumn(62.0 * 72 / 25.4));
TableRow contactRow = contactTable.AddRow();
contactRow.Cells.Add(new TableCell("生产企业：南宁化工股份有限公司"));
TableCell phone = new TableCell();
phone.AddParagraph("应急电话：0771-4818653").Alignment = ParagraphAlignment.Right;
contactRow.Cells.Add(phone);
contact.Blocks.Add(contactTable);
contact.Blocks.Add(new Paragraph { LineSpacing = 0.1 });
bottom.Cells.Add(contact);
bottom.Cells.Add(new TableCell("UN NO.:1789"));
```

要点：

- 被 `RowSpan` 覆盖的后续行不能再放该列单元格，否则 Word 中会错位；被 `ColSpan` 覆盖的列同理。写出 DOCX 时生成 `w:vMerge` / `w:gridSpan`，回读后 `RowSpan` / `ColSpan` 保持不变。
- 图片按磅指定显示尺寸，段落居中即可；转成 PDF 时作为嵌入对象绘制。
- 单元格内的嵌套表格用于底部左右对齐的联系信息，表格后补一个 `LineSpacing = 0.1` 的空段落，防止 Word 的结尾段落撑高行。

### 信封

`WordEnvelopes` 按版式把寄件人/收件人地址放到信封上，内置 DL（220×110 mm）、C5（229×162 mm）与中式 5 号（220×110 mm），也可自定义尺寸。

```csharp
OfficeDocument envelope = WordEnvelopes.Create(new WordEnvelopeOptions
{
    Layout = WordEnvelopeLayout.Dl(),
    Sender = "智慧办公\n上海市浦东新区示例路 1 号\n200120",
    Recipient = "张三\n上海市浦东新区世纪大道 100 号\n200120",
    FontSize = 12,
    SenderFontSize = 9
});
DocxWriter.Write(envelope, "envelope.docx");

// 批量：每条记录一页，收件人模板按记录替换；senderTemplate 为固定寄件人文本
OfficeDocument batch = WordEnvelopes.CreateMany(
    records,
    "{{Name}}\n{{Address}}\n{{Zip}}",
    new WordEnvelopeOptions { Layout = WordEnvelopeLayout.Chinese5() },
    "智慧办公");
```

地址用段前间距控制纵向位置、左缩进控制横向位置，输出为普通段落，不需要文本框；Word/WPS 与 PDF 导出都能直接渲染。版式中比页边距更靠边的位置会被贴齐到页边距。

## 克隆、追加与合并

```csharp
OfficeDocument copy = WordDocumentOperations.Clone(document);   // 深拷贝，改副本不影响原件

int sections = WordDocumentOperations.Append(target, source);  // 追加源文档全部节，返回节数

OfficeDocument merged = WordDocumentOperations.Merge(report1, report2, report3);
```

追加与合并按"节"保留页面设置、页眉页脚和节分隔，所有内容都是深拷贝；`Merge` 使用第一个非空文档的元数据，null 文档被跳过。

## 读取旧版 .doc

```csharp
OfficeDocument legacy = DocReader.Load("legacy.doc");   // Word 97-2003 复合文档
string plain = DocReader.ExtractText("legacy.doc");

DocxWriter.Write(legacy, "legacy.docx");                // 常用：.doc 转 .docx
```

`.doc` 仅支持读取，不支持写回；复杂域、批注、修订不保证还原。

## 修订跟踪、SmartArt 与数字签名

### 读取与写回修订

DOCX/DOTX 的插入、删除修订会进入统一模型：行内元素通过 `Inline.Revision` 标记，段落标记修订存放在 `Paragraph.Revision`，被删除的正文保存到 `Paragraph.DeletedInlines`（不参与 `GetText()` 与版面）。`DocxWriter` 会重新写出 `w:ins` / `w:del` / `w:delText` 与段落标记修订。

```csharp
using Wisdom.Word;
using Wisdom.Office.Core;

OfficeDocument document = DocxReader.Load("review.docx");

foreach (Section section in document.Sections)
    foreach (Block block in section.Blocks)
    {
        Paragraph paragraph = block as Paragraph;
        if (paragraph == null) continue;
        foreach (Inline inline in paragraph.Inlines)
        {
            RevisionInfo revision = inline.Revision;
            if (revision != null)
                Console.WriteLine($"{revision.Type} {revision.Author} {revision.Date}");
        }
        if (paragraph.DeletedInlines.Count > 0)
            Console.WriteLine("已删除：" + paragraph.GetDeletedText());
    }

int changes = WordRevisions.Count(document);   // 行内修订 + 删除内容 + 段落标记
WordRevisions.AcceptAll(document);             // 保留插入、丢弃删除、清除标记
// WordRevisions.RejectAll(document);          // 移除插入、恢复删除、清除标记
```

主动写入修订：

```csharp
TextRun inserted = new TextRun("新增内容", format);
inserted.Revision = new RevisionInfo(RevisionType.Insert, "张三", DateTime.UtcNow);
paragraph.Inlines.Add(inserted);

paragraph.DeletedInlines.Add(new TextRun("已删除内容", format)
{
    Revision = new RevisionInfo(RevisionType.Delete, "李四", DateTime.UtcNow)
});
```

接受/拒绝按"可见结果"处理：拒绝修订时被删除内容会追加到段落末尾（模型不记录原始字符偏移），段落标记修订直接清除、不拆分或合并相邻段落。修订信息只在 DOCX/DOTX 中读写；RTF、HTML、ODT 与纯文本转换会输出最终可见内容。

### 读取 SmartArt

`WordSmartArtReader` 从正文、页眉与页脚中定位 `dgm:relIds` 引用，解析数据部件并提取每个节点的文本：

```csharp
IList<WordSmartArt> diagrams = WordSmartArtReader.Read("report.docx");
foreach (WordSmartArt diagram in diagrams)
{
    Console.WriteLine($"{diagram.NodeCount} 个节点：{diagram.Text}");
    Console.WriteLine(diagram.PartPath);    // 如 word/diagrams/data1.xml
}
```

SmartArt 的布局、样式与配色分散在多个包级部件中，统一模型不承载这些结构；检查器用于内容提取与存在性判断，重新生成的包不会保留原 SmartArt 图形。

### 检查数字签名

`WordDigitalSignatureReader` 检测 `_xmlsignatures/` 下的签名部件，读取签名/摘要/规范化算法、被引用的部件与证书主体：

```csharp
WordDigitalSignatureReport report = WordDigitalSignatureReader.Read("signed.docx");
if (report.IsSigned)
{
    foreach (WordDigitalSignature signature in report.Signatures)
    {
        Console.WriteLine(signature.SignatureMethod);
        Console.WriteLine(signature.SignerSubject);
        foreach (string part in signature.SignedParts) Console.WriteLine(part);
    }
}
```

该 API 只做结构检查，不验证签名的密码学有效性；任何读取后重新写出的包都会使原签名失效，写出器不会复制签名部件。

## 多格式读写与转换

`WordFormatConverter` 以统一文档模型为中转，把 Word 生态的常见格式统一成一套 API：

| 格式 | 读取 | 写入 | 说明 |
| --- | :---: | :---: | --- |
| `.docx` | ✓ | ✓ | OOXML 主格式 |
| `.dotx` | ✓ | ✓ | 模板，Content-Type 为 `wordprocessingml.template.main+xml` |
| `.doc` | ✓ | ✗ | Word 97-2003 二进制，只读 |
| `.rtf` | ✓ | ✓ | 多节、页眉页脚、表格合并、图片、书签与超链接 |
| `.html` | ✓ | ✓ | 自包含单文件，CSS 与 Base64 图片内嵌，不引用外部资源 |
| `.odt` | ✓ | ✓ | ODF 1.3，含 `content.xml` / `styles.xml` / `meta.xml` / manifest |
| `.txt` | ✓ | ✓ | UTF-8，仅保留段落文本 |

7 种输入 × 6 种输出 = **42 种转换组合**；扩展名与文件头都会参与识别，`.docx` 与 `.dotx` 通过 Content-Type 区分。

```csharp
using Wisdom.Word;

// 按扩展名自动识别输入、按目标扩展名选择输出
WordFormatConverter.Convert("legacy.doc", "report.odt");
WordFormatConverter.Convert("report.odt", "report.rtf");
WordFormatConverter.Convert("report.rtf", "report.html");
WordFormatConverter.Convert("report.html", "report.docx");
WordFormatConverter.Convert("report.docx", "template.dotx");

// 显式格式 + 流/字节数组
OfficeDocument document = WordFormatConverter.Load("report.rtf");
byte[] odt = WordFormatConverter.ToArray(document, WordDocumentFormat.Odt);
WordFormatConverter.Save(document, stream, WordDocumentFormat.Html);

// 格式嗅探
WordDocumentFormat format = WordFormatConverter.DetectFormat("report.odt");   // Odt
```

### 在格式之间保留什么

| 能力 | DOCX / DOTX | RTF | HTML | ODT |
| --- | :---: | :---: | :---: | :---: |
| 段落 / 字符格式 | ✓ | ✓ | ✓ | ✓ |
| 多节页面设置 | ✓ | ✓ | ✓ | ✓ |
| 页眉 / 页脚 | ✓ | ✓ | ✓ | ✓ |
| 表格与合并单元格 | ✓ | ✓ | ✓ | ✓ |
| 图片（PNG/JPEG） | ✓ | ✓ | ✓ | ✓ |
| 书签 / 超链接 | ✓ | ✓ 范围 | ✓ 整段 | ✓ 范围 |
| 分页符 / PAGE 域 | ✓ | ✓ | CSS 分页 | ✓ |
| 文档元数据 | ✓ | ✓ | `<title>`/`<meta>` | ✓ |

跨格式转换是"结构等价"而非像素级还原：样式表、主题、复杂域、修订与数字签名不进入统一模型，转换后不会保留；`.doc` 与 `.odt` 的读取侧重正文、表格与常用格式。

## 模板填充

占位符 `{{key}}`、`{{key:format}}`；支持跨 Run 的占位符（Word 把一个词拆进多个 Run 也能识别）、点路径、段落与表格行循环。

模板内容：

```text
{{customerName}}，您好：

{{#each items}}
{{name}}    数量：{{qty}}    金额：{{amount:0.00}}
{{/each}}
```

```csharp
using Wisdom.Word;
using Wisdom.Office.Core;

TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["items"] = new[]
{
    new { name = "无线键盘", qty = 2, amount = 398.0 },
    new { name = "机械鼠标", qty = 5, amount = 449.5 }
};

DocxTemplate.Fill("template.docx", "filled.docx", context);
byte[] result = DocxTemplate.Fill(File.ReadAllBytes("template.docx"), context);

DocxTemplate.Fill(input, output, context, new TemplateOptions
{
    LeaveUnresolvedTokens = false,   // 未解析的占位符置空
    HtmlEncodeValues = true,         // 值做 XML/HTML 转义
    Culture = CultureInfo.InvariantCulture
});
```

`{{#each}}` 在 Word 中按段落或表格行整体循环；空列表时整个循环块被移除。

## 文档整理引擎

`WordNormalizer` 在统一模型上做批量规范化——适合把多人协作、格式混乱的文档拉回统一标准。

```csharp
using Wisdom.Word;

// 默认构造已按中文公文常用值打开绝大多数开关，下面的显式赋值是为了改目标值或关掉不需要的步骤
WordNormalizeOptions options = new WordNormalizeOptions
{
    RemoveEmptyParagraphs = true,
    NormalizeWhitespace = true,
    BodyEastAsiaFont = "宋体",
    BodyAsciiFont = "Times New Roman",
    BodyFontSize = 12,
    BodyLineSpacing = 1.5,             // 行距是倍数，不是磅值
    FirstLineIndentCharacters = 2,
    PromoteFakeHeadings = true,          // 把"加粗大字"提升为真标题
    FixHeadingLevelSkips = true,         // 修正 1 级直接跳 3 级
    RegenerateTableOfContents = false,   // 默认 true：会插入目录并加分页符
    MergeShortParagraphs = false,        // 默认 true：短于 ShortParagraphThreshold 的段落会被并到一起
    NormalizeTables = true,
    TableAlignment = TableAlignment.Center,
    RepeatTableHeaderRows = true,
    NormalizeImages = true,
    MaxImageWidth = 420
};

WordNormalizeResult result = WordNormalizer.Normalize(document, options);
OfficeDocument cleaned = result.Document;
Console.WriteLine(result.ToReport());     // 变更列表 + 问题列表 + 前后统计
Console.WriteLine(result.Changes.Count);  // 逐项变更，含 Category / Description / Location
Console.WriteLine(result.Issues.Count);   // WordIssueSeverity：Info / Warning / Error
```

`WordProcessor` 是"读文件 → 备份 → 整理 → 写出 → 出报告"的一键流水线：

```csharp
WordProcessResult r = WordProcessor.Process("draft.docx", "final.docx", options);
Console.WriteLine(r.Report);
Console.WriteLine(r.BackupPath);                       // draft.backup-<时间戳>.docx，与原文件同目录
Console.WriteLine(r.ReportPath);                       // 固定为 outputPath + ".report.txt"
WordProcessor.Rollback(r.BackupPath, "draft.docx");    // 用备份覆盖目标文件（通常覆盖回原文件）
```

## 校验、差异与统计

```csharp
// 一致性校验：空段落、字体/字号混用、行距段距不一致、半角句末标点、控制字符、
// 标题跳级、表格空行、图片超出正文宽度等
List<WordIssue> issues = WordConsistencyChecker.Check(document, options);

// 两版本文本差异
WordDocumentDiff diff = WordDocumentDiff.Compare(before, after);
if (diff.HasDifferences) Console.WriteLine(diff.ToText());
// AddedLines / RemovedLines / AddedCount / RemovedCount / UnchangedCount

// 文档度量：段落数、字数、表格数、各级标题数、字体使用分布
WordDocumentMetrics metrics = WordDocumentMetrics.Measure(document);
Console.WriteLine($"{metrics.CharacterCount} 字，{metrics.TableCount} 表");
foreach (KeyValuePair<string, int> f in metrics.FontUsage) Console.WriteLine(f.Key + " " + f.Value);
```

`WordNormalizer.Normalize` 就地修改传入的文档模型（`result.Document` 就是它本身）。想在不影响原模型的前提下试算，直接从文件重新加载一份，或先用 `WordProcessor.CreateBackup` 备份文件。

## 类型一览

| 类型 | 用途 |
| --- | --- |
| `DocxReader` | DOCX → `OfficeDocument`，`ExtractText` |
| `DocxWriter` | `OfficeDocument` → DOCX（文件 / 流 / 字节） |
| `WordDocumentBuilder` | 流式文档构建（Aspose DocumentBuilder 风格） |
| `DocxTemplate` | DOCX 模板填充（包级，含 `{{#each}}` 循环） |
| `WordMailMerge` / `WordMailMergeOptions` | 模型级邮件合并（MERGEFIELD、区域/嵌套区域与双花括号） |
| `CustomPropertyCollection` | 带类型的自定义文档属性（DOCX/DOTX `docProps/custom.xml`） |
| `WordLabelLayout` / `WordLabelOptions` / `WordLabels` | 标签纸版式、Avery 5160/5163/L7160 与自定义毫米版式、记录模板、带填写线的表单标签、带行分隔线的分格标签 |
| `WordEnvelopeLayout` / `WordEnvelopeOptions` / `WordEnvelopes` | 信封版式（DL/C5/中式 5 号）与单个/批量生成 |
| `WordFindReplace` / `WordFindReplaceOptions` | 跨 Run 查找替换、全词与正则 |
| `WordDocumentOperations` | 深拷贝、追加与合并 |
| `DocReader` | 旧版 `.doc` 读取 |
| `WordFormatConverter` / `WordDocumentFormat` / `WordDocumentFormats` | 7 种格式的加载、保存、嗅探与互转 |
| `DocxReader` / `DocxWriter` | DOCX / DOTX 读写（`DocxWriter.WriteTemplate` 输出模板） |
| `RtfReader` / `RtfWriter` | RTF 读写 |
| `HtmlReader` / `HtmlWriter` | 自包含 HTML 读写 |
| `OdtReader` / `OdtWriter` | OpenDocument 文本读写 |
| `WordNormalizer` / `WordNormalizeOptions` / `WordNormalizeResult` | 批量整理 |
| `WordProcessor` / `WordProcessResult` | 文件级整理流水线与备份回滚 |
| `WordConsistencyChecker` / `WordIssue` / `WordIssueSeverity` | 规范校验 |
| `WordDocumentDiff` / `WordChange` | 版本差异 |
| `WordDocumentMetrics` | 结构与用字统计 |
| `RevisionInfo` / `RevisionType` | 修订元数据（作者、时间、类型） |
| `WordRevisions` | 修订计数、接受全部、拒绝全部 |
| `WordSmartArtReader` / `WordSmartArt` | SmartArt 节点文本与数据部件检查 |
| `WordDigitalSignatureReader` / `WordDigitalSignatureReport` / `WordDigitalSignature` | 数字签名结构检查（不验证密码学有效性） |

## 已知限制

- 不支持写入 `.doc`；`.doc` 只做读取。
- 跨格式转换覆盖正文结构（段落、格式、表格、图片、书签、链接、页眉页脚、分页），不保留样式表、主题、批注、修订、SmartArt 与数字签名等高级包结构。
- DOCX/DOTX 保留行内插入/删除与段落标记修订；不支持移动修订（`w:moveFrom`/`w:moveTo` 读取为删除/插入）、表格行/单元格/section 修订与修订批注气泡。接受/拒绝修订不重建原始字符位置，也不拆分或合并相邻段落。
- SmartArt 仅提供内容提取（节点文本），重新生成的包不会保留 SmartArt 图形结构。
- 数字签名仅做结构与算法信息检查，不验证密码学有效性；写出器不复制签名部件，任何重新生成包的操作都会使原签名失效。
- 不处理批注、嵌入对象与 VBA；读取时忽略。
- 分页、页码等运行时排版由 Word/WPS 决定；本包输出的是 `PAGE`/`NUMPAGES` 域与分页属性，不做像素级分页模拟。
- 图片仅支持 PNG / JPEG。
- 邮件合并区域不支持同名区域自我嵌套（嵌套区域请使用不同名称），也不支持跨节区域。
- 自定义文档属性只在 DOCX/DOTX 包中往返；转成 RTF/HTML/ODT/TXT 或重写旧版 `.doc` 时不会保留。
- 标签纸的行高是“最小值”（atLeast），单张内容过多时行高会扩展并可能把后续标签挤到下一页；打印前建议按实际打印机校准页边距。
- 信封仅放置文字地址，不绘制回邮窗口、条码或邮票框；版式坐标按页面计算，打印机不可打印边距可能造成偏移。
- 查找替换只处理普通文本运行，不搜索域显示文本或图片；正则匹配内置 2 秒超时。
- 目录（TOC）写入的是域，目录条目由 Word/WPS 打开后更新生成，本包不预渲染目录内容。
