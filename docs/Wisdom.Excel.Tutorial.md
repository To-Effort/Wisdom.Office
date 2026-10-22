# Wisdom.Excel 使用教程

Excel 侧的实战教程：从零写出带样式、数字格式、合并、冻结与打印设置的台账，读回来核对；再挂原生图表、条件格式与锚定图片，做模板填充，最后读真实报表与旧版 `.xls`。安装方式见[包说明](../src/Wisdom.Excel/README.md#安装)。

数据模型（`OfficeWorkbook` / `Worksheet` / `Cell` / `CellStyle` / `ExcelChart` …）都在 `Wisdom.Office.Core`，读写器在本包。文中所有输出都是实测结果。

## 你将完成

```text
out/
  ledger.xlsx           第 1 步写出的台账（样式、数字格式、合并、冻结、打印设置）
  dashboard.xlsx        第 3 步的看板（图表 + 条件格式 + 锚定图片）
  order-template.xlsx   第 4 步生成的模板
  order.xlsx            模板填充结果
  types.xlsx            第 7 步的值类型实测表
```

## 1. 写一张台账

`GetCell(row, column)` 取单元格，行列都从 1 开始；单元格不存在时自动创建。样式挂在 `Cell.Style` 上，`NumberFormat` 用 Excel 的分节语法。

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Wisdom.Excel;
using Wisdom.Office.Core;

OfficeWorkbook workbook = new OfficeWorkbook();
workbook.Metadata.Title = "设备采购台账";
workbook.Metadata.Author = "报表组";

Worksheet sheet = workbook.AddWorksheet("台账");

// 标题行：合并 A1:E1
sheet.GetCell(1, 1).Value = "设备采购台账";
sheet.Merge("A1:E1");
sheet.GetCell(1, 1).Style = new CellStyle
{
    FontFamily = "微软雅黑",
    FontSize = 16,
    Bold = true,
    HorizontalAlignment = HorizontalAlignment.Center,
    VerticalAlignment = VerticalAlignment.Middle
};
sheet.SetRowHeight(1, 28);

string[] headers = { "品名", "数量", "单价", "金额", "采购日期" };
CellStyle headerStyle = new CellStyle
{
    Bold = true,
    FillColor = OfficeColor.FromHtml("#DDEBF7"),
    HorizontalAlignment = HorizontalAlignment.Center,
    TopBorder = new BorderLine(BorderStyle.Single, 0.5, OfficeColor.Gray),
    BottomBorder = new BorderLine(BorderStyle.Single, 0.5, OfficeColor.Gray)
};
for (int c = 0; c < headers.Length; c++)
{
    Cell head = sheet.GetCell(2, c + 1);
    head.Value = headers[c];
    head.Style = headerStyle;
}

object[,] rows =
{
    { "无线键盘", 12.0, 199.0, 2388.0, new DateTime(2026, 1, 9) },
    { "机械鼠标", 5.0, 89.9, 449.5, new DateTime(2026, 2, 18) },
    { "显示器支架", 0.0, 260.0, -1.0, new DateTime(2026, 3, 2) },
    { "清洁服务", 1.0, 420.5, 0.0, new DateTime(2026, 4, 25) }
};
CellStyle moneyStyle = new CellStyle { NumberFormat = "#,##0.00", HorizontalAlignment = HorizontalAlignment.Right };
CellStyle qtyStyle = new CellStyle { NumberFormat = "0;-0;\"—\"", HorizontalAlignment = HorizontalAlignment.Center };
CellStyle dateStyle = new CellStyle { NumberFormat = "yyyy-MM-dd", HorizontalAlignment = HorizontalAlignment.Center };

for (int r = 0; r < rows.GetLength(0); r++)
{
    int row = r + 3;
    sheet.GetCell(row, 1).Value = (string)rows[r, 0];
    sheet.GetCell(row, 2).Value = (double)rows[r, 1];
    sheet.GetCell(row, 2).Style = qtyStyle;
    sheet.GetCell(row, 3).Value = (double)rows[r, 2];
    sheet.GetCell(row, 3).Style = moneyStyle;
    sheet.GetCell(row, 4).Value = (double)rows[r, 3];
    sheet.GetCell(row, 4).Style = moneyStyle;
    sheet.GetCell(row, 5).Value = (DateTime)rows[r, 4];
    sheet.GetCell(row, 5).Style = dateStyle;
}

// 合计：公式文本 + 缓存值（本库不计算，值要自己给）
Cell total = sheet.GetCell(7, 4);
total.Value = 2836.5;
total.Formula = "SUM(D3:D6)";
total.Style = new CellStyle
{
    NumberFormat = "#,##0.00",
    Bold = true,
    TopBorder = new BorderLine(BorderStyle.Double, 1.0, OfficeColor.Black)
};
sheet.GetCell(7, 1).Value = "合计";
sheet.GetCell(7, 1).Style = new CellStyle { Bold = true };

sheet.GetCell(8, 1).Value = "本表含税";
sheet.GetCell(8, 1).Style = new CellStyle { Italic = true, FontColor = OfficeColor.Gray };

sheet.SetColumnWidth(1, 22);
sheet.SetColumnWidth(2, 10);
sheet.SetColumnWidth(3, 12);
sheet.SetColumnWidth(4, 14);
sheet.SetColumnWidth(5, 14);
sheet.FrozenRows = 2;                     // 冻结前两行（表头）
sheet.FrozenColumns = 1;                  // 冻结首列（品名）
sheet.ShowGridLines = false;
sheet.SetColumnHidden(3, true);           // 单价列先藏起来，第 2 步演示取舍
sheet.SetPrintArea("A1:E8");
sheet.PrintSetup = PageSetup.A4();
sheet.PrintSetup.IsLandscape = true;      // 方向由 Width/Height 关系决定，setter 会交换宽高
sheet.FitToPage = true;
sheet.CenterHorizontally = true;

Directory.CreateDirectory("out");
XlsxWriter.Write(workbook, "out/ledger.xlsx");     // 亦支持 Stream，或 byte[] Write(workbook)
```

实测：

```text
file=4582 bytes=4582
sheets=1 name=台账
lastRow=8 lastColumn=5 merged=1
hasPrintArea=True landscape=True
```

列宽单位是 Excel 的"字符宽"，`GetColumnWidthPoints` 才换算成磅；`LastRow` / `LastColumn` 统计已用区域，包含合并区域与有意创建的空单元格。

## 2. 读回来核对

```csharp
OfficeWorkbook workbook = XlsxReader.Load("out/ledger.xlsx");   // 亦支持 Stream / byte[]
Worksheet sheet = workbook.FirstWorksheet;

Console.WriteLine("title=" + workbook.Metadata.Title + " author=" + workbook.Metadata.Author);
Console.WriteLine("range=" + sheet.LastRow + "x" + sheet.LastColumn + " merged=" + sheet.MergedRanges.Count
                  + " frozen=" + sheet.FrozenRows + "/" + sheet.FrozenColumns
                  + " gridLines=" + sheet.ShowGridLines + " hiddenCols=[" + string.Join(",", sheet.HiddenColumns) + "]");
Console.WriteLine("A1=" + sheet.GetCell("A1").DisplayText()
                  + " 样式=" + sheet.GetCell("A1").Style.FontFamily + "/" + sheet.GetCell("A1").Style.FontSize);

for (int row = 2; row <= 8; row++)
{
    List<string> line = new List<string>();
    for (int c = 1; c <= 5; c++) line.Add(sheet.GetCell(row, c).DisplayText());
    Console.WriteLine(row + ": " + string.Join(" | ", line));
}

Console.WriteLine("D7 formula=" + sheet.GetCell(7, 4).Formula);
```

```text
title=设备采购台账 author=报表组
range=8x5 merged=1 frozen=2/1 gridLines=False hiddenCols=[3]
A1=设备采购台账 样式=微软雅黑/16
2: 品名 | 数量 | 单价 | 金额 | 采购日期
3: 无线键盘 | 12 | 199.00 | 2,388.00 | 2026-01-09
4: 机械鼠标 | 5 | 89.90 | 449.50 | 2026-02-18
5: 显示器支架 | — | 260.00 | -1.00 | 2026-03-02
6: 清洁服务 | 1 | 420.50 | 0.00 | 2026-04-25
7: 合计 |  |  | 2,836.50 |
8: 本表含税 |  |  |  |
D7 formula=SUM(D3:D6)
```

样式、数字格式、合并、冻结、隐藏列、打印区域都按原样读回。第 5 行的数量 `0` 被 `"0;-0;\"—\""` 的第三段渲染成 `—`，第 6 行的金额 `0` 用的是 `#,##0.00`，所以仍是 `0.00`——分节语法 `正数;负数;零;文本` 就是在这里生效的。

`GetCell` 找不到时会新建单元格，只想探测存在性要用 `FindCell`（不存在返回 `null`）；`FindMergeAnchor` 用来判断某个格子是否落在合并区域内：

```csharp
Console.WriteLine("anchor(1,3)=" + (sheet.FindMergeAnchor(1, 3) == null
    ? "null" : CellAddress.Format(sheet.FindMergeAnchor(1, 3).Row, sheet.FindMergeAnchor(1, 3).Column)));
Console.WriteLine("find(9,9)=" + (sheet.FindCell(9, 9) == null ? "null（未创建）" : "created"));
Console.WriteLine("width B=" + sheet.GetColumnWidth(2) + " points=" + sheet.GetColumnWidthPoints(2) + " height1=" + sheet.GetRowHeight(1));
```

```text
anchor(1,3)=A1
find(9,9)=null（未创建）
width B=10 points=57.5 height1=28
```

隐藏内容默认全部导入；不想要时按维度关掉：

```csharp
OfficeWorkbook lean = XlsxReader.Load("out/ledger.xlsx", new ExcelReadOptions { IncludeHiddenColumns = false });
Worksheet leanSheet = lean.FirstWorksheet;
Console.WriteLine("without-hidden-cols: lastColumn=" + leanSheet.LastColumn
                  + " C3=[" + (leanSheet.FindCell(3, 3) == null ? "已剔除" : "仍在") + "]");
Console.WriteLine("extract-text-lines=" + XlsxReader.ExtractText("out/ledger.xlsx").Split('\n').Length);
```

```text
without-hidden-cols: lastColumn=5 C3=[已剔除]
extract-text-lines=10
```

被剔除的只是隐藏列里的**单元格**，`LastColumn` 仍是 5，因为 E 列本身有数据。只要纯文本时用 `XlsxReader.ExtractText(path)`。

## 3. 图表、条件格式与锚定图片

三类装饰对象挂在 `Worksheet` 的 `Charts` / `ConditionalFormats` / `Images` 上，写出后回读同样还原，所以可以「加载 → 改 → 保存」而不丢装饰。

```csharp
OfficeWorkbook workbook = new OfficeWorkbook();
Worksheet sheet = workbook.AddWorksheet("看板");
string[] months = { "1月", "2月", "3月", "4月", "5月", "6月" };
double[] sales = { 820, 910, 760, 1020, 1180, 990 };
sheet.GetCell(1, 1).Value = "月份";
sheet.GetCell(2, 1).Value = "销售额";
for (int i = 0; i < months.Length; i++)
{
    sheet.GetCell(1, i + 2).Value = months[i];
    sheet.GetCell(2, i + 2).Value = sales[i];
    sheet.SetColumnWidth(i + 2, 10);
}

ExcelChart chart = new ExcelChart
{
    Type = ExcelChartType.Column,
    Title = "上半年销售额",
    Row = 4,                       // 锚定到 D4（行列从 1 开始）
    Column = 1,
    WidthPoints = 360,
    HeightPoints = 200,
    ShowLegend = true,
    LegendPosition = "b",
    ValueAxisFormat = "#,##0"
};
// 数据源用区域引用：图表随单元格数据自动更新
ExcelChartSeries series = chart.AddSeries("看板!$A$2", "看板!$B$1:$G$1", "看板!$B$2:$G$2");
series.FillColor = OfficeColor.FromHtml("#00B050");
series.ShowDataLabels = true;
series.DataLabelFormat = "#,##0";
sheet.Charts.Add(chart);

// 也可用字面量集合建图（不依赖工作表区域）；饼图/环形图按 PointColors 逐点着色
ExcelChart pie = new ExcelChart { Type = ExcelChartType.Pie, Row = 4, Column = 8, Title = "占比" };
ExcelChartSeries literal = pie.AddSeries("占比", new[] { "线上", "门店", "渠道" }, new double[] { 45, 35, 20 });
literal.PointColors.Add(OfficeColor.FromHtml("#4472C4"));
literal.PointColors.Add(OfficeColor.FromHtml("#ED7D31"));
literal.PointColors.Add(OfficeColor.FromHtml("#A5A5A5"));
sheet.Charts.Add(pie);

ExcelConditionalFormat format = new ExcelConditionalFormat("B2:G2");
format.AddDataBar(new OfficeColor(0x63, 0xC3, 0x84));
format.AddColorScale(new OfficeColor(0xF8, 0x69, 0x6A), new OfficeColor(0xFF, 0xF3, 0xB5), new OfficeColor(0x63, 0xC3, 0x84));
format.AddIconSet("3TrafficLights1", new double[] { 33, 67 });
format.AddCellIs("greaterThan", "1000", new CellStyle { FontColor = OfficeColor.Red, Bold = true });
sheet.ConditionalFormats.Add(format);

byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
sheet.Images.Add(new ExcelImage(new OfficeImage(png))
{
    Row = 1,
    Column = 9,
    WidthPoints = 48,
    HeightPoints = 48
});

XlsxWriter.Write(workbook, "out/dashboard.xlsx");

OfficeWorkbook back = XlsxReader.Load("out/dashboard.xlsx");
Worksheet bs = back.FirstWorksheet;
Console.WriteLine("charts=" + bs.Charts.Count + " types=[" + string.Join(",", bs.Charts.Select(x => x.Type)) + "]");
Console.WriteLine("series range=" + bs.Charts[0].Series[0].ValueRange + " literal=" + string.Join(",", bs.Charts[1].Series[0].Values));
Console.WriteLine("cf=" + bs.ConditionalFormats.Count + " rules=" + bs.ConditionalFormats[0].Rules.Count
                  + " kinds=[" + string.Join(",", bs.ConditionalFormats[0].Rules.Select(r => r.Kind)) + "]");
Console.WriteLine("images=" + bs.Images.Count + " anchor=" + bs.Images[0].Row + "," + bs.Images[0].Column
                  + " size=" + bs.Images[0].WidthPoints + "x" + bs.Images[0].HeightPoints);
```

```text
bytes=7343
charts=2 types=[Column,Pie]
series range=看板!$B$2:$G$2 literal=45,35,20
cf=1 rules=4 kinds=[DataBar,ColorScale,IconSet,CellIs]
images=1 anchor=1,9 size=48x48
```

`ExcelChartType` 支持 `Column` / `Bar` / `Line` / `Pie` / `Doughnut` / `Area` / `StackedColumn`；区域引用在读写两端只记录字符串，取值发生在写出时。`CellIs` 的差量样式（`dxf`）会真实落盘，`AddCellIs("between", …)` 时记得再给 `SecondFormula`。图片格式与像素尺寸由 `OfficeImage` 按文件头魔数嗅探，`ExcelImage` 的默认尺寸取图片按 96 DPI 换算的磅值。

## 4. 模板填充

单元格里的占位符是 `{{key}}` / `{{key:format}}`；行循环用 `{{#each 列表}}` 与 `{{/each}}`，两个标记可以分别放在循环行的首列和末列，也可以跨多行。

先生成模板：

```csharp
OfficeWorkbook template = new OfficeWorkbook();
Worksheet sheet = template.AddWorksheet("订单");
sheet.GetCell(1, 1).Value = "客户：{{customerName}}";
sheet.GetCell(1, 4).Value = "签订日期：{{signDate:yyyy-MM-dd}}";
sheet.GetCell(3, 1).Value = "品名";
sheet.GetCell(3, 2).Value = "数量";
sheet.GetCell(3, 3).Value = "金额";
sheet.GetCell(3, 4).Value = "序号";
for (int c = 1; c <= 4; c++) sheet.GetCell(3, c).Style = new CellStyle { Bold = true };

sheet.GetCell(4, 1).Value = "{{#each items}}{{name}}";
sheet.GetCell(4, 2).Value = "{{qty}}";
sheet.GetCell(4, 3).Value = "{{amount:0.00}}";
sheet.GetCell(4, 4).Value = "{{@number}}";
sheet.GetCell(4, 5).Value = "{{/each}}";
sheet.GetCell(4, 3).Style = new CellStyle { NumberFormat = "#,##0.00", HorizontalAlignment = HorizontalAlignment.Right };
sheet.SetColumnWidth(1, 18);

sheet.GetCell(6, 1).Value = "合计";
sheet.GetCell(6, 3).Value = "{{total:0.00}}";
XlsxWriter.Write(template, "out/order-template.xlsx");
```

再填充：

```csharp
TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["signDate"] = new DateTime(2026, 10, 21);
context["total"] = 1268.0;
context["items"] = new[]
{
    new { name = "无线键盘", qty = 2, amount = 398.0 },
    new { name = "机械鼠标", qty = 5, amount = 449.5 },
    new { name = "清洁服务", qty = 1, amount = 420.5 }
};

XlsxTemplate.Fill("out/order-template.xlsx", "out/order.xlsx", context, new TemplateOptions
{
    LeaveUnresolvedTokens = false,
    Culture = CultureInfo.InvariantCulture
});

OfficeWorkbook filled = XlsxReader.Load("out/order.xlsx");
Worksheet fs = filled.FirstWorksheet;
Console.WriteLine("A1=" + fs.GetCell("A1").DisplayText() + " D1=" + fs.GetCell("D1").DisplayText());
for (int row = 3; row <= 8; row++)
{
    List<string> line = new List<string>();
    for (int c = 1; c <= 4; c++) line.Add(fs.GetCell(row, c).DisplayText());
    Console.WriteLine(row + ": " + string.Join(" | ", line));
}
Console.WriteLine("lastRow=" + fs.LastRow);
```

```text
A1=客户：示例客户 D1=签订日期：2026-10-21
3: 品名 | 数量 | 金额 | 序号
4: 无线键盘 | 2 | 398.00 | 1
5: 机械鼠标 | 5 | 449.50 | 2
6: 清洁服务 | 1 | 420.50 | 3
7:  |  |  |
8: 合计 |  | 1268.00 |
lastRow=8
```

循环行按条目数展开成 3 行，模板行的样式、行高、列宽与合并区域都跟着复制，`合计` 那行整体下移。同一份模板可以在内存里反复填充——`XlsxTemplate.Fill(OfficeWorkbook, …)` 返回**新**工作簿，传入的模板不被修改：

```csharp
OfficeWorkbook inMemory = XlsxTemplate.Fill(XlsxReader.Load("out/order-template.xlsx"), context);
Console.WriteLine("in-memory template untouched=" + XlsxReader.Load("out/order-template.xlsx").FirstWorksheet.GetCell("A1").DisplayText()
                  + " result=" + inMemory.FirstWorksheet.GetCell("A1").DisplayText());
```

```text
in-memory template untouched=客户：{{customerName}} result=客户：示例客户
```

列表为空时整个循环块被删掉，后面的行上移：

```csharp
TemplateContext empty = new TemplateContext();
empty["customerName"] = "零订单客户";
empty["items"] = new object[0];
OfficeWorkbook none = XlsxTemplate.Fill(XlsxReader.Load("out/order-template.xlsx"), empty);
Console.WriteLine("empty-list lastRow=" + none.FirstWorksheet.LastRow + " A1=" + none.FirstWorksheet.GetCell("A1").DisplayText());
Console.WriteLine("empty A5=[" + none.FirstWorksheet.FindCell(5, 1).DisplayText() + "]");
```

```text
empty-list lastRow=5 A1=客户：零订单客户
empty A5=[合计]
```

循环变量 `@index`（从 0）、`@number`（从 1）、`@first`、`@last` 与 `.`（当前项本身）都可用；嵌套属性用点路径 `{{order.id}}`，不支持 `[i]` 下标。

## 5. 读一份真实报表

`tests/Wisdom.Office.Tests/fixtures/attendance-real.xlsx` 是仓库里的真实考勤台账，拿它验证读取覆盖面：

```csharp
OfficeWorkbook workbook = XlsxReader.Load("attendance-real.xlsx");
Console.WriteLine("sheets=" + workbook.Worksheets.Count);
foreach (Worksheet sheet in workbook.Worksheets)
{
    Console.WriteLine("  " + sheet.Name + " " + sheet.LastRow + "x" + sheet.LastColumn
                      + " merged=" + sheet.MergedRanges.Count + " charts=" + sheet.Charts.Count
                      + " cf=" + sheet.ConditionalFormats.Count + " images=" + sheet.Images.Count
                      + " hidden=" + sheet.IsHidden);
}

Worksheet sept = workbook.FindWorksheet("9月考勤");
for (int row = 1; row <= 4; row++)
{
    List<string> line = new List<string>();
    for (int c = 1; c <= 8; c++) line.Add(sept.GetCell(row, c).DisplayText());
    Console.WriteLine("  " + row + ": " + string.Join(" | ", line));
}

OfficeWorkbook visibleOnly = XlsxReader.Load("attendance-real.xlsx", new ExcelReadOptions { IncludeHiddenWorksheets = false });
Console.WriteLine("visible-only sheets=" + visibleOnly.Worksheets.Count
                  + " names=[" + string.Join(",", visibleOnly.Worksheets.Select(s => s.Name)) + "]");
```

```text
sheets=5
  数据引用区域 1556x15 merged=1 charts=0 cf=0 images=0 hidden=True
  9月考勤 28x37 merged=16 charts=0 cf=0 images=0 hidden=False
  10月考勤 24x37 merged=16 charts=0 cf=0 images=0 hidden=False
  11月考勤 24x37 merged=15 charts=0 cf=1 images=0 hidden=False
  Sheet1 0x0 merged=0 charts=0 cf=0 images=0 hidden=False
  1: 请选择日期 |  |  |  |  |  | 9 |
  2: 9月考勤表 |  |  |  |  |  |
  3:  | √ | 出勤 |  | ● | 休假 |  | ○
  4: 姓 名 | 星期 | 一 | 二 | 三 | 四 | 五 | 六
visible-only sheets=4 names=[9月考勤,10月考勤,11月考勤,Sheet1]
```

注意三个层次：第一个工作表是隐藏的取数区（1556 行），关掉 `IncludeHiddenWorksheets` 才能只拿到业务表；`11月考勤` 带 1 组条件格式，`10月考勤` 有 16 处合并——这些装饰都会被还原；`Sheet1` 是空表，`LastRow` 为 0。多级表头（第 3～4 行）在这里是**两个普通行**，合并区域靠 `MergedRanges` 描述，需要自己按 `FindMergeAnchor` 归位。

## 6. 读旧版 .xls

BIFF8（Excel 97-2003）只读，用 `XlsReader`；API 与 `XlsxReader` 一致。

```csharp
OfficeWorkbook workbook = XlsReader.Load("legacy-sample.xls");
Worksheet sheet = workbook.FirstWorksheet;
Console.WriteLine("sheets=" + workbook.Worksheets.Count);
Console.WriteLine("range=" + sheet.LastRow + "x" + sheet.LastColumn + " merged=" + sheet.MergedRanges.Count);
Console.WriteLine("1: " + string.Join(" | ", Enumerable.Range(1, 6).Select(c => sheet.GetCell(1, c).DisplayText())));
Console.WriteLine("text-lines=" + XlsReader.ExtractText("legacy-sample.xls").Split('\n').Length);
```

```text
sheets=1
range=1x12 merged=0
1: 单位 |  姓名 |  职称 |  地址 |  电话 |  传真
text-lines=3
```

再用一个带公式的样例（`legacy-solver.xls`）看公式的处理方式：

```csharp
OfficeWorkbook solver = XlsReader.Load("legacy-solver.xls");
Worksheet ss = solver.FirstWorksheet;
Console.WriteLine("solver sheets=" + solver.Worksheets.Count + " range=" + ss.LastRow + "x" + ss.LastColumn);
for (int row = 9; row <= 12; row++)
{
    Console.WriteLine("solver " + row + ": " + string.Join(" | ", Enumerable.Range(1, 6)
        .Select(c => ss.GetCell(row, c).Value == null ? "" : ss.GetCell(row, c).DisplayText() + "(" + ss.GetCell(row, c).Value.GetType().Name + ")")));
}
Console.WriteLine("solver formula cells=" + ss.Cells.Count(x => !string.IsNullOrEmpty(x.Formula)));
```

```text
solver sheets=6 range=18x8
solver 9:  |  | Number to Build->(String) | 100(Double) | 100(Double) | 100(Double)
solver 10: Part Name(String) | Inventory(String) | No. Used(String) |  |  |
solver 11: Chassis(String) | 450(Double) | 200(Double) | 1(Double) | 1(Double) | 0(Double)
solver 12: Picture Tube(String) | 250(Double) | 100(Double) | 1(Double) | 0(Double) | 0(Double)
solver formula cells=0
```

`formula cells=0` 说明 `.xls` 只保留公式的**缓存结果**（那些 `Double`），不还原公式文本；要转成 XLSX 就先读成模型再 `XlsxWriter.Write`。不支持 BIFF5/7 与加密工作簿，也不支持写回 `.xls`。

## 7. 值类型、数字格式与容错

```csharp
OfficeWorkbook wb = new OfficeWorkbook();
Worksheet s = wb.AddWorksheet("类型");
s.GetCell(1, 1).Value = 12.50m;                                   // decimal
s.GetCell(1, 2).Value = true;                                     // bool
s.GetCell(1, 3).Value = new DateTime(2026, 10, 21, 18, 30, 0);
s.GetCell(1, 3).Style = new CellStyle { NumberFormat = "yyyy-MM-dd HH:mm" };
s.GetCell(1, 4).Value = "  前后空格  ";
s.GetCell(1, 5).Value = 0.0;
s.GetCell(1, 5).Style = new CellStyle { NumberFormat = "0.00;[Red]-0.00;零" };
s.GetCell(1, 6).Value = -3.5;
s.GetCell(1, 6).Style = new CellStyle { NumberFormat = "0.00;[Red]-0.00;零" };
s.GetCell(1, 7).Value = 1234567890.0;
s.GetCell(1, 7).Style = new CellStyle { NumberFormat = "0.00" };
s.SetColumnWidth(7, 4);
XlsxWriter.Write(wb, "out/types.xlsx");

OfficeWorkbook back = XlsxReader.Load("out/types.xlsx");
Worksheet bs = back.FirstWorksheet;
for (int c = 1; c <= 7; c++)
{
    Cell cell = bs.GetCell(1, c);
    Console.WriteLine("  col" + c + " 值类型=" + (cell.Value == null ? "null" : cell.Value.GetType().Name)
                      + " 显示=[" + cell.DisplayText() + "]");
}
```

```text
  col1 值类型=Double 显示=[12.5]
  col2 值类型=Boolean 显示=[TRUE]
  col3 值类型=DateTime 显示=[2026-10-21 18:30]
  col4 值类型=String 显示=[  前后空格  ]
  col5 值类型=Double 显示=[零]
  col6 值类型=Double 显示=[-3.50]
  col7 值类型=Double 显示=[1234567890.00]
```

要点：`decimal` 落盘转成 `double`（`12.50m` 读回是 `12.5`）；布尔值以 XLSX 的布尔类型存储，读回是 `Boolean`；带 `NumberFormat` 的日期读回是 `DateTime`，日期序列号只在未标日期格式时才表现为数字；文本首尾空格原样保留。`[Red]` 这类颜色指令只影响显示颜色，`DisplayText()` 只给数字部分。第 7 列宽度只有 4 个字符，`DisplayText()` 仍返回完整数字——把它渲染成 `####` 是 `Wisdom.Office` 转换层的行为，由 `WorkbookLayoutOptions.EmulateNumberOverflow`（默认 `true`）控制。

坏文件的异常类型在六个框架下统一：

```csharp
byte[] good = XlsxWriter.Write(wb);
byte[] truncated = new byte[good.Length / 2];
Array.Copy(good, truncated, truncated.Length);
try { XlsxReader.Load(truncated); } catch (Exception ex) { Console.WriteLine("truncated=" + ex.GetType().Name); }
try { XlsxReader.Load(System.Text.Encoding.UTF8.GetBytes("这不是 xlsx 文件")); }
catch (Exception ex) { Console.WriteLine("not-zip=" + ex.GetType().Name); }
```

```text
truncated=InvalidDataException
not-zip=InvalidDataException
```

## 常见问题

**读出来 `LastRow` / `LastColumn` 比预期大。** 已用区域包含合并区域和"被有意创建的空单元格"——`GetCell` 一旦调用就会创建并计入。只探测存在性请用 `FindCell`。

**公式算不出值。** 本库没有计算引擎，`Cell.Formula` 只是文本；写出时要同时给 `Cell.Value` 作为缓存结果，否则 Excel/WPS 打开后才会重算，而 `DisplayText()` 拿到的是空值。

**图表导到 PDF / Word 里是占位文字。** 图表按 OOXML 落盘，由 Excel / WPS 渲染，本库不做位图渲染；要在 PDF/Word 中出现图形，请另存一份图片版并用 `ExcelImage` 锚定。

**隐藏的工作表/行/列进来了。** `ExcelReadOptions` 三个开关默认全开，按需关闭；判断工作表是否隐藏还可以直接看 `Worksheet.IsHidden`。

**中文列宽换算困惑。** `SetColumnWidth` 用的是 Excel 的"字符宽"单位，磅值请走 `GetColumnWidthPoints`；反过来从磅设置字符宽会带来取整误差。

**`.xls` 想写回。** 不支持——只能读。要产出 XLSX 就读成 `OfficeWorkbook` 再用 `XlsxWriter.Write`。
