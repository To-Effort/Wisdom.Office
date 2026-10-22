# Wisdom.Excel

`Wisdom.Excel` 负责 **Excel 侧的全部读写**：XLSX 读取与生成（样式、数字格式、公式、合并、冻结窗格、打印设置、原生图表、条件格式、单元格图片、自动筛选、数据验证、超链接与标签色）、集合 / `DataTable` 映射、CSV 导入导出、低内存流式查询、区域与行列编辑、工作表管理、XLSX 模板填充，以及旧版 `.xls`（BIFF8，Excel 97-2003）读取。

- 包名：`Wisdom.Excel`
- 版本：`1.0.1`
- 依赖：`Wisdom.Office.Core`（唯一依赖）
- 第三方依赖：无；不要求安装 Microsoft Office 或 WPS
- 命名空间：`Wisdom.Excel`（数据模型位于 `Wisdom.Office.Core`）

本文是能力与 API 参考；分步实战教程见源码仓库的 `docs/Wisdom.Excel.Tutorial.md`。

## 支持的框架

| 目标框架 | 包内资产 | 典型运行时 |
| --- | --- | --- |
| `net461` | `lib/net461` | .NET Framework 4.6.1 / 4.7.2 / 4.8.x |
| `netstandard2.0` | `lib/netstandard2.0` | .NET Core 2.0-3.1、.NET 5/6、Mono、Unity |
| `net7.0` | `lib/net7.0` | .NET 7 |
| `net8.0` | `lib/net8.0` | .NET 8 |
| `net9.0` | `lib/net9.0` | .NET 9 |
| `net10.0` | `lib/net10.0` | .NET 10 / 11 |

XLSX 包由 `System.IO.Compression.ZipArchive` + LINQ to XML 直接处理；`.xls` 的 OLE2 容器与 BIFF8 记录由本库自行解析。

> 关于 .NET Framework：`net461` 上的 `ZipArchive` 会延迟解析中央目录，损坏的条目名可能在完全不同的调用点抛出 `ArgumentException`。本包统一改用 `PackageArchive.OpenForRead` 立即物化目录并把这类异常收敛为 `InvalidDataException`，因此六个框架下的容错行为一致（1.0 修复项）。

## 安装

```bash
dotnet add package Wisdom.Excel --version 1.0.1
```

需要把 Excel 转成 PDF / DOCX 时请改引 `Wisdom.Office` 门面包。

## 读取 XLSX

```csharp
using Wisdom.Excel;
using Wisdom.Office.Core;

OfficeWorkbook workbook = XlsxReader.Load("sales.xlsx");   // 亦支持 Stream / byte[]

foreach (Worksheet sheet in workbook.Worksheets)
{
    Console.WriteLine(sheet.Name);
    foreach (Cell cell in sheet.Cells)                      // 只遍历实际存在的单元格
    {
        if (cell.Value != null) Console.WriteLine($"{cell.Row}:{cell.Column} = {cell.DisplayText()}");
    }
    Console.WriteLine($"范围 {sheet.LastRow} x {sheet.LastColumn}，合并 {sheet.MergedRanges.Count} 处");
}

string text = XlsxReader.ExtractText("sales.xlsx");
```

隐藏内容的取舍由 `ExcelReadOptions` 控制：

```csharp
OfficeWorkbook all = XlsxReader.Load("sales.xlsx", new ExcelReadOptions
{
    IncludeHiddenWorksheets = true,
    IncludeHiddenRows = true,
    IncludeHiddenColumns = true
});
```

读取还原：值与缓存公式结果、`CellStyle`（字体、颜色、填充、对齐、边框、换行、数字格式）、合并区域、列宽行高、冻结窗格、打印区域与纸张方向、隐藏行列、以及图表 / 条件格式 / 锚定图片。

## 写入 XLSX

```csharp
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
    Bold = true,
    FillColor = OfficeColor.FromHtml("#F2F2F2"),
    HorizontalAlignment = HorizontalAlignment.Right
};

sheet.GetCell(3, 3).Formula = "SUM(C2:C2)";    // 公式文本；Value 可另存缓存结果
sheet.Merge("A4:C4");
sheet.GetCell(4, 1).Value = "备注：本季度价格含税。";
sheet.FrozenRows = 1;
sheet.SetColumnWidth(1, 24);
sheet.SetRowHeight(4, 30);
sheet.SetColumnHidden(2, true);
sheet.ShowGridLines = false;
sheet.ShowZeros = false;                        // 关闭 0 值显示
sheet.SetPrintArea("A1:C20");
sheet.PrintSetup = PageSetup.A4();
sheet.PrintSetup.IsLandscape = true;            // 由 Width/Height 关系决定方向
sheet.FitToPage = true;
sheet.CenterHorizontally = true;

XlsxWriter.Write(workbook, "sales.xlsx");       // 亦支持 Stream，或 byte[] Write(workbook)
```

`Cell.Value` 支持 `string`、`double`、`bool`、`DateTime`（`decimal` 也能写，落盘时转为 `double`）。数字格式遵循 Excel 的分节语法 `正数;负数;零;文本`，`DisplayText()` 与导出的 PDF/Word 使用同一套格式化实现；列宽放不下数值或日期时输出 `####` 是**转换层**的行为，由 `Wisdom.Office` 的 `WorkbookLayoutOptions.EmulateNumberOverflow`（默认 `true`）控制，`DisplayText()` 本身始终给出完整文本。

## 集合映射与 DataTable

`ExcelCollection` 提供 MiniExcel / EPPlus 风格的集合映射：对象集合直接写成表，读回时按表头名称还原对象。

```csharp
public sealed class Order
{
    [ExcelColumn("订单号", Order = 1)] public string Code { get; set; }
    [ExcelColumn("数量", Order = 2)] public int Quantity { get; set; }
    [ExcelColumn("金额", Order = 3, Format = "#,##0.00")] public decimal Amount { get; set; }
    [ExcelColumn("下单时间", Order = 4, Format = "yyyy-mm-dd")] public DateTime Created { get; set; }
    [ExcelColumn(Ignore = true)] public string InternalNote { get; set; }
}

OfficeWorkbook workbook = new OfficeWorkbook();
ExcelCollection.Write(workbook, orders, "订单");          // 表头 + 数据 + 数字格式

List<Order> read = ExcelCollection.ToList<Order>(workbook.Worksheets[0]);
DataTable table = ExcelCollection.ToDataTable(workbook.Worksheets[0]);

ExcelCollection.Write(workbook, dataTable, "导入");       // DataTable → 工作表
DataTable back = ExcelCollection.ToDataTable(workbook.Worksheets[0]);
```

`ExcelColumnAttribute` 的 `Order` 从 1 开始，未指定时排在最后；`Format` 写入数字/日期格式，日期值未指定格式时自动补 `yyyy-mm-dd hh:mm:ss`。读取时表头名不区分大小写，找不到的成员保留默认值；`ExcelCollectionOptions` 控制起始行、起始列、首行是否为表头。

## 低内存流式查询与 CSV

数据量达到百万行时不必构建整个工作簿模型：`ExcelQuery` 直接流式解析工作表部件，逐行返回，只有枚举期间持有文件句柄。

```csharp
// 逐行读取（数字 → double、布尔 → bool、日期 → DateTime）
foreach (ExcelRowData row in ExcelQuery.ReadRows("big.xlsx", "明细"))
{
    object name = row[1];                       // 按列号取值，不存在返回 null
    if (row.Hidden) continue;                   // 跳过隐藏行
}

// 首行表头按名称映射到对象，返回 IEnumerable<T>，边读边反序列化
foreach (OrderRow item in ExcelQuery.Query<OrderRow>("big.xlsx", "明细"))
{
    Process(item);
}

// 每行一个字典（键不区分大小写，重复表头自动加序号）
foreach (Dictionary<string, object> row in ExcelQuery.ReadDictionaries("big.xlsx", "明细"))
{
    Console.WriteLine(row["订单号"]);
}
```

`Query<T>` / `ReadDictionaries` 同样接受 `Stream`、`byte[]` 与 `ExcelCollectionOptions`；不可定位的流会被复制到内存。共享字符串表仍会整体加载（与 MiniExcel 语义一致），样式与日期系统按工作簿设置解释。

CSV 读写覆盖 RFC 4180 的引号、逗号与字段内换行，并带行列数与单字段长度上限：

```csharp
ExcelCsv.Write(workbook, "export.csv");                       // 默认 UTF-8 BOM、逗号
OfficeWorkbook csv = ExcelCsv.Read("import.csv");             // 数值/布尔自动转换
ExcelCsv.Write(sheet, "export-tab.csv", new ExcelCsvOptions { Separator = '\t', WriteBom = false });
```

## 区域 API 与行列编辑

`ExcelRange` 把 A1 区域当作一个单元格块来读写，偏移量从 1 开始：

```csharp
ExcelRange header = sheet.Range("A1:C1");
header.Values = new object[,] { { "产品", "数量", "金额" } };  // 二维数组，0 基
header.SetStyle(style => { style.Bold = true; style.FillColor = OfficeColor.LightGray; });

sheet.Range("A2:C2").CopyTo(sheet.Range("A3:C3"));
sheet.Range("A5:C5").Merge();
sheet.Range("A1:C20").AutoFitColumns(minWidth: 8, maxWidth: 40);   // 全角字符按 2 字符计

sheet.InsertRows(2, 1);        // 在第 2 行前插入 1 行
sheet.DeleteRows(2, 1);
sheet.InsertColumns(2, 1);     // 在第 2 列前插入 1 列
sheet.DeleteColumns(2, 1);
```

行列编辑会同步平移单元格、行高列宽、隐藏标记、合并区域、自动筛选、数据验证与打印区域；`ExcelRange.Clear()` 保留样式只清内容，`ClearAll()` 连同样式与区域内的合并一起清除。

## 工作表管理与高级表设置

```csharp
Worksheet copy = workbook.CopyWorksheet("明细", "明细副本");   // 深拷贝单元格与样式
workbook.MoveWorksheet(0, 1);
workbook.RenameWorksheet(copy, "归档");                        // 校验 31 字符与非法字符
workbook.RemoveWorksheet("归档");
bool valid = OfficeWorkbook.IsValidWorksheetName("报表 2026");

sheet.TabColor = OfficeColor.FromHtml("#1F3A5F");             // 表标签颜色
sheet.AutoFilter = new ExcelAutoFilter("A1:D100");
sheet.AutoFilter.ForColumn(0).Values.Add("华东");

sheet.GetCell(10, 1).Value = "官网";
sheet.GetCell(10, 1).Hyperlink = "https://example.com";        // 外链，写出为 r:id 关系
sheet.GetCell(10, 1).HyperlinkTooltip = "打开官网";
sheet.GetCell(10, 2).Hyperlink = "#明细!A1";                   // 内链，写出为 location

sheet.DataValidations.Add(new ExcelDataValidation
{
    Range = "B2:B100",
    Type = DataValidationType.List,
    Formula1 = "待处理,处理中,已完成",                          // 列表字面量会自动加引号
    ShowErrorMessage = true,
    ErrorMessage = "请从下拉列表中选择"
});

workbook.DefinedNames.Add(new ExcelDefinedName("Amounts", "明细!$C$2:$C$100"));
```

超链接只接受 `http/https/mailto/ftp/ftps/file/tel/smb` 与相对路径，`javascript:`、`data:` 等方案在写出时直接丢弃；筛选、数据验证、标签色、超链接与定义名均可被 `XlsxReader` 回读。

## 图表、条件格式与图片

三类装饰对象随 `XlsxWriter` 写出，`XlsxReader` 回读时同样还原，因此可以「加载 → 修改 → 保存」而不丢装饰。

```csharp
Worksheet sheet = workbook.AddWorksheet("看板");
for (int c = 1; c <= 7; c++) sheet.SetColumnWidth(c, 10);
for (int i = 0; i < 6; i++) sheet.GetCell(1, i + 2).Value = (i + 1) + "月";
sheet.GetCell(2, 1).Value = "销售额";
for (int i = 0; i < 6; i++) sheet.GetCell(2, i + 2).Value = 820.0 + i * 100;

// 原生图表：柱状 / 条形 / 折线 / 面积 / 饼 / 环形 / 堆积柱
ExcelChart chart = new ExcelChart
{
    Type = ExcelChartType.Column,
    Title = "上半年销售额",
    Row = 4,                    // 锚定单元格
    Column = 1,
    WidthPoints = 360,
    HeightPoints = 200,
    ShowLegend = true,
    LegendPosition = "b",
    ValueAxisFormat = "#,##0"
};
// 数据源用区域引用，图表随单元格数据自动更新
ExcelChartSeries series = chart.AddSeries("看板!$A$2", "看板!$B$1:$G$1", "看板!$B$2:$G$2");
series.FillColor = OfficeColor.FromHtml("#00B050");
series.ShowDataLabels = true;
series.DataLabelFormat = "#,##0";
sheet.Charts.Add(chart);

// 也可用字面量集合建图（不依赖工作表区域）
ExcelChart pie = new ExcelChart { Type = ExcelChartType.Pie, Row = 4, Column = 8 };
pie.AddSeries("占比", new[] { "甲", "乙", "丙" }, new double[] { 45, 35, 20 });
sheet.Charts.Add(pie);

// 条件格式：数据条、双色/三色色阶、图标集、数值比较高亮（dxf 样式真实落盘）
ExcelConditionalFormat format = new ExcelConditionalFormat("B2:G2");
format.AddDataBar(new OfficeColor(0x63, 0xC3, 0x84));
format.AddColorScale(new OfficeColor(0xF8, 0x69, 0x6A), new OfficeColor(0x63, 0xC3, 0x84));
format.AddIconSet("3TrafficLights1", new double[] { 33, 67 });
format.AddCellIs("greaterThan", "1000", new CellStyle { FontColor = OfficeColor.Red, Bold = true });
sheet.ConditionalFormats.Add(format);

// 锚定到单元格的图片（PNG / JPEG 自动识别尺寸）
sheet.Images.Add(new ExcelImage(OfficeImage.FromFile("logo.png"))
{
    Row = 2,
    Column = 9,
    WidthPoints = 24,
    HeightPoints = 24
});

XlsxWriter.Write(workbook, "dashboard.xlsx");

OfficeWorkbook loaded = XlsxReader.Load("dashboard.xlsx");
Console.WriteLine(loaded.Worksheets[0].Charts.Count);                   // 2
Console.WriteLine(loaded.Worksheets[0].Charts[0].Series[0].ValueRange);  // 看板!$B$2:$G$2
```

图表按 OOXML 保存，由 Excel / WPS 渲染。导出 Word/PDF 时图表位置输出占位文本；需要在 PDF/Word 里出现图形，请改存一份图片版并用 `ExcelImage` 锚定。

## 模板填充

单元格里的占位符为 `{{key}}`、`{{key:format}}`；循环标记 `{{#each list}}` 与 `{{/each}}` 可以跨一行或多行。

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

TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["items"] = new[]
{
    new { name = "无线键盘", qty = 2, price = 199.0 },
    new { name = "机械鼠标", qty = 5, price = 89.9 }
};

XlsxTemplate.Fill("template.xlsx", "filled.xlsx", context);

// 也可在内存模型上直接填充：返回填充后的新工作簿，传入的模板工作簿不被修改
OfficeWorkbook wb = XlsxReader.Load("template.xlsx");
OfficeWorkbook filled = XlsxTemplate.Fill(wb, context, new TemplateOptions
{
    LeaveUnresolvedTokens = false,
    HtmlEncodeValues = false,
    Culture = CultureInfo.InvariantCulture
});
```

循环按行展开时保留模板行的样式、行高与合并；空列表会清掉整个循环块。

## 读取旧版 .xls

```csharp
OfficeWorkbook legacy = XlsReader.Load("legacy.xls");   // BIFF8，Excel 97-2003
string text = XlsReader.ExtractText("legacy.xls");
```

值、样式、合并区域、窗格与元数据均可读取；公式只保留缓存值，不还原公式文本。不支持 BIFF5/7 与加密工作簿；`.xls` 不支持写回。

## 类型一览

| 类型 | 用途 |
| --- | --- |
| `XlsxReader` | XLSX → `OfficeWorkbook`，`ExtractText` |
| `XlsxWriter` | `OfficeWorkbook` → XLSX（文件 / 流 / 字节） |
| `XlsxTemplate` | XLSX 模板填充（文件或内存模型） |
| `XlsReader` | 旧版 `.xls`（BIFF8）读取 |
| `ExcelCollection`、`ExcelColumnAttribute`、`ExcelCollectionOptions` | 对象集合 / `DataTable` ↔ 工作表映射 |
| `ExcelQuery`、`ExcelRowData`、`ExcelCellData` | 低内存流式查询（行 / 字典 / 对象） |
| `ExcelCsv`、`ExcelCsvOptions` | CSV 导入导出 |
| `ExcelRange`、`ExcelRangeExtensions` | A1 区域读写、合并、复制、自动列宽、`UsedRange` |
| `ExcelAutoFilter`、`ExcelFilterColumn` | 自动筛选条件 |
| `ExcelDataValidation`、`DataValidationType` | 数据验证规则 |
| `ExcelDefinedName` | 工作簿定义名 |
| `ExcelReadOptions` | 隐藏表 / 行 / 列的取舍 |
| `OfficeWorkbook`、`Worksheet`、`Cell`、`CellStyle` | 数据模型（定义在 `Wisdom.Office.Core`） |
| `ExcelChart`、`ExcelChartSeries`、`ExcelChartType` | 图表模型 |
| `ExcelConditionalFormat`、`ExcelConditionalRule`、`ExcelConditionalKind` | 条件格式模型 |
| `ExcelImage` | 锚定图片模型 |

## 已知限制

- 不支持写入 `.xls`；`.xlsx` 中的宏（`xlsm`）不保留 VBA 内容。
- 加密工作簿、外部工作簿链接、数据透视表、切片器、形状/批注不解析；单元格超链接与定义名支持读写。
- 图表按 OOXML 落盘，本包不做位图渲染；PDF/Word 导出只输出占位文本。
- 公式文本可以写入，但本库不做计算引擎，值需由调用方提供或由 Excel 打开时重算。
- CSV 只表示单个工作表；流式查询不构建模型，但共享字符串表仍需整体加载，且未提供异步枚举（`netstandard2.0` 无 `IAsyncEnumerable`）。
