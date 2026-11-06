# 不装 Office 也能操作 Excel：Wisdom.Excel 使用指南

> 「纯 .NET 操作 Office」系列 · Excel 篇

## 写在前面

服务端生成 Excel 报表，大概是 .NET 项目里最常见的轮子需求之一。可选方案无非三类：`Microsoft.Office.Interop`（要求机器上装 Office，服务端基本不能用）、EPPlus/ClosedXML（功能强但许可证和体积让人犹豫）、MiniExcel（轻量但只做数据搬运）。

`Wisdom.Excel` 走的是第四条路：**零第三方依赖，自己解析 OOXML**。全文只有一个内部依赖 `Wisdom.Office.Core`（数据模型），不要求安装 Microsoft Office 或 WPS。NuGet 安装：

```bash
dotnet add package Wisdom.Excel
```

支持 `net461` / `netstandard2.0` / `net7.0` / `net8.0` / `net9.0` / `net10.0` 六份目标框架，由同一份源码编译（源码里没有任何 `#if NET*` 条件编译），.NET Framework 4.6.1 以上的老项目和 Unity 都能用。

这篇覆盖：读取、生成、对象映射、百万行流式查询、模板填充，以及图表 / 条件格式这些"装饰型"能力。

## 一、数据模型：三个类走天下

所有操作都建立在 `OfficeWorkbook → Worksheet → Cell` 这条链上，模型定义在 `Wisdom.Office.Core`：

```csharp
using Wisdom.Excel;
using Wisdom.Office.Core;

OfficeWorkbook workbook = XlsxReader.Load("sales.xlsx");   // 也支持 Stream / byte[]

foreach (Worksheet sheet in workbook.Worksheets)
{
    Console.WriteLine(sheet.Name);
    foreach (Cell cell in sheet.Cells)                     // 只遍历真实存在的单元格
    {
        if (cell.Value != null)
            Console.WriteLine($"{cell.Row}:{cell.Column} = {cell.DisplayText()}");
    }
    Console.WriteLine($"范围 {sheet.LastRow} x {sheet.LastColumn}，合并 {sheet.MergedRanges.Count} 处");
}
```

`Cell.Value` 接受 `string` / `double` / `bool` / `DateTime`（`decimal` 也能写，落盘时转 `double`）。只想抽文本的话，一行就够：

```csharp
string text = XlsxReader.ExtractText("sales.xlsx");
```

隐藏表 / 隐藏行列默认不参与读取，需要时显式打开：

```csharp
OfficeWorkbook all = XlsxReader.Load("sales.xlsx", new ExcelReadOptions
{
    IncludeHiddenWorksheets = true,
    IncludeHiddenRows = true,
    IncludeHiddenColumns = true
});
```

读回来的东西比想象中多：值与缓存公式结果、字体 / 颜色 / 填充 / 对齐 / 边框 / 数字格式、合并区域、列宽行高、冻结窗格、打印区域与纸张方向，甚至图表、条件格式和锚定图片。这意味着「加载 → 改几个格子 → 保存」不会把原文件的装饰弄丢。

## 二、生成：样式、公式、打印设置

```csharp
OfficeWorkbook workbook = new OfficeWorkbook();
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

sheet.GetCell(3, 3).Formula = "SUM(C2:C2)";   // 公式文本
sheet.Merge("A4:C4");
sheet.FrozenRows = 1;                          // 冻结表头
sheet.SetColumnWidth(1, 24);
sheet.ShowGridLines = false;

sheet.SetPrintArea("A1:C20");
sheet.PrintSetup = PageSetup.A4();
sheet.PrintSetup.IsLandscape = true;
sheet.FitToPage = true;
sheet.CenterHorizontally = true;

XlsxWriter.Write(workbook, "sales.xlsx");      // 亦可写 Stream 或返回 byte[]
```

数字格式沿用 Excel 的分节语法 `正数;负数;零;文本`，`DisplayText()` 和导出 PDF / Word 时用的是**同一套格式化实现**——这一点在跨格式报表里很重要，不会出现"Excel 里是日期，PDF 里成了序列号"。

`####`（列宽不足时 Excel 的显示行为）在导出 PDF / Word 时可以选择是否模拟，由 `Wisdom.Office` 包的 `WorkbookLayoutOptions.EmulateNumberOverflow` 控制，而 `DisplayText()` 永远给完整文本。

## 三、对象集合直接收发

`ExcelCollection` 提供 MiniExcel 风格的映射，注解控制列名、顺序与数字格式：

```csharp
public sealed class Order
{
    [ExcelColumn("订单号", Order = 1)] public string Code { get; set; }
    [ExcelColumn("数量", Order = 2)] public int Quantity { get; set; }
    [ExcelColumn("金额", Order = 3, Format = "#,##0.00")] public decimal Amount { get; set; }
    [ExcelColumn("下单时间", Order = 4, Format = "yyyy-mm-dd")] public DateTime Created { get; set; }
    [ExcelColumn(Ignore = true)] public string InternalNote { get; set; }
}

ExcelCollection.Write(workbook, orders, "订单");
List<Order> read = ExcelCollection.ToList<Order>(workbook.Worksheets[0]);
DataTable table = ExcelCollection.ToDataTable(workbook.Worksheets[0]);
```

读取时表头名不区分大小写，Excel 里缺的列保留属性默认值；日期没指定格式时自动补 `yyyy-mm-dd hh:mm:ss`。`DataTable` 与工作表可以双向。

## 四、百万行：ExcelQuery 流式读

`XlsxReader.Load` 会构建整个工作簿模型，遇到几十上百万行的文件就很吃力。`ExcelQuery` 绕过模型直接流式解析工作表部件，逐行返回：

```csharp
foreach (ExcelRowData row in ExcelQuery.ReadRows("big.xlsx", "明细"))
{
    object name = row[1];          // 按列号取值，不存在返回 null
    if (row.Hidden) continue;
}

foreach (OrderRow item in ExcelQuery.Query<OrderRow>("big.xlsx", "明细"))
    Process(item);                 // 首行表头映射到对象，边读边反序列化

foreach (Dictionary<string, object> row in ExcelQuery.ReadDictionaries("big.xlsx", "明细"))
    Console.WriteLine(row["订单号"]);
```

语义与 MiniExcel 一致：共享字符串表仍会整体加载，枚举期间持有文件句柄，样式与日期系统按工作簿设置解释。目前未提供异步枚举（`netstandard2.0` 没有 `IAsyncEnumerable`）。

## 五、区域与行列编辑

`ExcelRange` 把 A1 区域当作一个块来操作：

```csharp
ExcelRange header = sheet.Range("A1:C1");
header.Values = new object[,] { { "产品", "数量", "金额" } };
header.SetStyle(s => { s.Bold = true; s.FillColor = OfficeColor.LightGray; });

sheet.Range("A2:C2").CopyTo(sheet.Range("A3:C3"));
sheet.Range("A5:C5").Merge();
sheet.Range("A1:C20").AutoFitColumns(minWidth: 8, maxWidth: 40);   // 全角字符按 2 字符计

sheet.InsertRows(2, 1);
sheet.DeleteColumns(2, 1);
```

插入 / 删除行列时，单元格、行高列宽、隐藏标记、合并区域、自动筛选、数据验证、打印区域、每页重复的标题行列，以及条件格式区域、图表锚点与系列取数区域、图片锚点都会同步平移——这是手工写偏移最容易漏的地方；整行 / 整列引用（`2:8`、`A:A`，带 `$` 也一样）按同一规则走，行编辑不碰列引用，列编辑不碰行引用。公式文本里的引用是个例外：库不重写公式，插入行后 `=SUM(A1:A5)` 仍写着 `A1:A5`，得自行改或者改完 `FormulaEngine.Calculate` 复核。`Clear()` 只清内容保留样式，`ClearAll()` 连同样式与合并一起清除。

## 六、表级功能

```csharp
Worksheet copy = workbook.CopyWorksheet("明细", "明细副本");   // 深拷贝内容与样式
workbook.MoveWorksheet(0, 1);
workbook.RenameWorksheet(copy, "归档");                        // 校验 31 字符与非法字符
bool valid = OfficeWorkbook.IsValidWorksheetName("报表 2026");

sheet.TabColor = OfficeColor.FromHtml("#1F3A5F");              // 标签色
sheet.AutoFilter = new ExcelAutoFilter("A1:D100");
sheet.AutoFilter.ForColumn(0).Values.Add("华东");

sheet.GetCell(10, 1).Hyperlink = "https://example.com";        // 外链
sheet.GetCell(10, 2).Hyperlink = "#明细!A1";                   // 文档内跳转

sheet.DataValidations.Add(new ExcelDataValidation
{
    Range = "B2:B100",
    Type = DataValidationType.List,
    Formula1 = "待处理,处理中,已完成",
    ShowErrorMessage = true,
    ErrorMessage = "请从下拉列表中选择"
});

workbook.DefinedNames.Add(new ExcelDefinedName("Amounts", "明细!$C$2:$C$100"));
```

超链接只放行 `http/https/mailto/ftp/ftps/file/tel/smb` 与相对路径，`javascript:`、`data:` 这类危险协议在写出时直接丢弃——把它当作导出组件时，这是一条省心的人文安全兜底。筛选、数据验证、标签色、超链接与定义名都能被读回。

## 七、原生图表、条件格式与图片

```csharp
ExcelChart chart = new ExcelChart
{
    Type = ExcelChartType.Column,      // 柱/条/折线/面积/饼/环形/堆积柱
    Title = "上半年销售额",
    Row = 4, Column = 1,
    WidthPoints = 360, HeightPoints = 200,
    ShowLegend = true, LegendPosition = "b",
    ValueAxisFormat = "#,##0"
};
ExcelChartSeries series = chart.AddSeries("看板!$A$2", "看板!$B$1:$G$1", "看板!$B$2:$G$2");
series.FillColor = OfficeColor.FromHtml("#00B050");
series.ShowDataLabels = true;
sheet.Charts.Add(chart);

ExcelConditionalFormat format = new ExcelConditionalFormat("B2:G2");
format.AddDataBar(new OfficeColor(0x63, 0xC3, 0x84));                  // 数据条
format.AddColorScale(new OfficeColor(0xF8, 0x69, 0x6A), new OfficeColor(0x63, 0xC3, 0x84));  // 色阶
format.AddIconSet("3TrafficLights1", new double[] { 33, 67 });         // 图标集
format.AddCellIs("greaterThan", "1000", new CellStyle { FontColor = OfficeColor.Red, Bold = true });
sheet.ConditionalFormats.Add(format);

sheet.Images.Add(new ExcelImage(OfficeImage.FromFile("logo.png"))
{
    Row = 2, Column = 9, WidthPoints = 24, HeightPoints = 24
});
```

图表数据源用区域引用，改动单元格后 Excel 里图表自动跟着变；不想依赖区域也可以直接喂字面量数组。条件格式的 `dxf` 样式是真实落盘的，不是截图糊上去的。

一点预期管理：图表按 OOXML 保存、由 Excel / WPS 渲染，本库不做位图渲染。所以导出 PDF / Word 时图表位置输出的是占位文本——要在 PDF 里看到图，请另存一份图片版并用 `ExcelImage` 锚定。

## 八、模板填充

业务方给一张排好版的 `.xlsx` 模板，程序只管往里灌数据，这是国内项目的高频需求。占位符 `{{key}}` / `{{key:format}}`，循环 `{{#each list}}` … `{{/each}}`：

```text
A6: {{#each items}}{{name}}    B6: {{qty}}    C6: {{price:0.00}}    D6: {{@number}}    E6: {{/each}}
```

```csharp
TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["items"] = new[]
{
    new { name = "无线键盘", qty = 2, price = 199.0 },
    new { name = "机械鼠标", qty = 5, price = 89.9 }
};

XlsxTemplate.Fill("template.xlsx", "filled.xlsx", context);
```

循环按行展开时保留模板行的样式、行高与合并；空列表会把整个循环块清掉，而不是留下一行孤零零的标记。覆盖循环行的条件格式、筛选、数据验证、打印区域与重复标题、隐藏行列、图表/图片锚点也按同一行映射换算到输出行，工作簿的 1904 日期基准与定义名同样带走；只有公式与定义名里的引用文本不跟着重写。也可在内存模型上填充（`XlsxTemplate.Fill(workbook, context)` 返回新工作簿，不动原模板）。

## 九、CSV 与旧版 .xls

```csharp
ExcelCsv.Write(workbook, "export.csv");              // 默认 UTF-8 BOM、逗号
OfficeWorkbook csv = ExcelCsv.Read("import.csv");    // 数值 / 布尔自动识别
ExcelCsv.Write(sheet, "export-tab.csv", new ExcelCsvOptions { Separator = '\t', WriteBom = false });
```

CSV 覆盖 RFC 4180 的引号、逗号与字段内换行，并带行列数与单字段长度上限（防 zip 炸弹式的畸形输入）。

老古董文件也能读：

```csharp
OfficeWorkbook legacy = XlsReader.Load("legacy.xls");   // BIFF8，Excel 97-2003
```

`XlsReader` 自己解析 OLE2 容器与 BIFF8 记录，值、样式、合并区域、窗格与元数据都能读；公式只保留缓存值，不支持 BIFF5/7 和加密工作簿，`.xls` 也不支持写回。

## 已知边界

- 不支持写 `.xls`；`.xlsx` 里的宏（`xlsm`）不保留 VBA 内容。
- 加密工作簿、外部工作簿链接、数据透视表、切片器、形状 / 批注不解析。
- 条件格式只识别 `dataBar` / `colorScale` / `iconSet` / `cellIs` 四类规则；"包含文本"、"Top 10"、公式型 `expression` 这些类型读取时直接跳过，重写或模板填充后不再出现。
- **公式可以重算**：写出前调用 `FormulaEngine.Calculate`（`Wisdom.Office.Core.Formulas`）即可整簿或单表重算并把结果写回缓存值，引擎覆盖 137 个函数；不实现数组公式（CSE）、外部工作簿引用与迭代计算。
- 一个已知坑位：`net461` 上 `ZipArchive` 延迟解析中央目录，损坏的条目名会在完全无关的调用点抛 `ArgumentException`。本包统一走 `PackageArchive.OpenForRead`，立即物化目录并把这类异常收敛成 `InvalidDataException`，六个框架行为一致。

## 小结

`Wisdom.Excel` 的定位很清晰：把 XLSX 当作一个可以直接读写的 ZIP + XML 包来处理，覆盖到「企业报表」这个场景需要的绝大部分——样式、数字格式、公式、合并、冻结、打印、图表、条件格式、模板填充、流式读取、CSV、老 `.xls`，代价是明确放弃宏与透视表这些重活；公式求值则交给同系列的 `Wisdom.Office.Core.Formulas`。

如果你的项目只需要数据搬运，MiniExcel 更轻；如果你需要 PDF / Word / PPT 之间互转，去看同系列的 `Wisdom.Office` 门面包。只操作 Excel 且在意依赖纯净度的话，这个包值得一试。
