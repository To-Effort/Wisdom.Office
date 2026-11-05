# Wisdom.Pdf 使用教程

PDF 引擎的实战教程：从模型排出一份多页报表并核对分页，注册与解析字体，反向解析既有 PDF 的文本与表格，把 PDF 转成可编辑 Word，在既有文档上做对象级编辑与水印，填写并扁平化 AcroForm 表单，以及在服务里只输出字节。

安装方式见[包说明](../src/Wisdom.Pdf/README.md#安装)。本教程示例是 .NET 8 控制台工程；落盘验证需要 `Wisdom.Word` 时已单独标注。

## 你将完成

```text
out/
  report.pdf              手工模型排版的三页报表（页码已替换、跨页表头重复）
  report.json.txt         从 report.pdf 反解出来的文本行与坐标
  tables.txt              从 report.pdf 识别出的表格
  editable.docx           PDF 反向转成的 Word（需要 Wisdom.Word）
  report-stamped.pdf      对象级编辑后的文档：首页转横向 + 平铺前景水印 + 图像印章
  form-blank.pdf          对象级拼出的四字段空白表单
  form-filled.pdf         填好值并生成外观的表单（中文已子集嵌入）
  form-flat.pdf           扁平化后的定稿，控件与 AcroForm 都已移除
```

## 1. 生成第一份 PDF

模型只描述"内容 + 版式意图"，分页由渲染器计算。

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Wisdom.Pdf;
using Wisdom.Office.Core;

OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "月度巡检汇总";
document.Metadata.Creator = "设施部";

PageSetup setup = PageSetup.A4();
setup.MarginLeft = setup.MarginRight = 60;
setup.MarginTop = setup.MarginBottom = 60;

Section section = document.AddSection(setup);
```

```csharp
Directory.CreateDirectory("out");

TextFormat titleFormat = new TextFormat("SimHei", 16);
titleFormat.EastAsiaFontFamily = "黑体";
titleFormat.Bold = true;

Paragraph title = new Paragraph();
title.Alignment = ParagraphAlignment.Center;
title.HeadingLevel = 1;
title.KeepWithNext = true;
title.SpaceAfter = 4;
title.AddText("月度巡检汇总", titleFormat);
section.Blocks.Add(title);

TextFormat body = new TextFormat("SimSun", 10);
TextFormat bodyBold = new TextFormat("SimSun", 10) { Bold = true };

Paragraph subtitle = new Paragraph();
subtitle.Alignment = ParagraphAlignment.Center;
subtitle.SpaceAfter = 12;
subtitle.AddText("2026 年 10 月　", body);
subtitle.AddText("共 3 个区域", bodyBold);
section.Blocks.Add(subtitle);
```

`TextFormat` 的构造函数只有 `TextFormat()` 与 `TextFormat(string fontFamily, double fontSize)` 两种，其余属性（`Bold`、`Italic`、`Color`、`EastAsiaFontFamily` 等）用对象初始化器或赋值设置；格式挂在行内 `TextRun` 上，段落本身不带格式。

## 2. 表格与跨页

表格行数超出单页时会自动拆分；把表头行标成 `IsHeader` 并开启 `RepeatHeaderRow`，每页顶部都会重复这些行。`TableRow.Height` 用来控制每页能装下多少行。

```csharp
Table table = new Table();
table.CellPadding = 4;
table.RepeatHeaderRow = true;
table.Columns.Add(new TableColumn(34));
table.Columns.Add(new TableColumn(150));
table.Columns.Add(new TableColumn(86));
table.Columns.Add(new TableColumn(70));

TableRow header = table.AddRow();
header.IsHeader = true;
AddCell(header, "序号", HorizontalAlignment.Center);
AddCell(header, "巡检项", HorizontalAlignment.Left);
AddCell(header, "结果", HorizontalAlignment.Center);
AddCell(header, "用时(min)", HorizontalAlignment.Right);

for (int i = 0; i < 52; i++)
{
    TableRow row = table.AddRow();
    row.Height = 20;
    AddCell(row, (i + 1).ToString(), HorizontalAlignment.Center);
    AddCell(row, "配电柜温度与接线端子检查 " + (i + 1), HorizontalAlignment.Left);
    AddCell(row, i % 11 == 5 ? "异常" : "正常", HorizontalAlignment.Center);
    AddCell(row, (6 + i % 5).ToString(), HorizontalAlignment.Right);
}

section.Blocks.Add(table);
```

```csharp
static void AddCell(TableRow row, string text, HorizontalAlignment alignment)
{
    TableCell cell = new TableCell();
    cell.VerticalAlignment = VerticalAlignment.Middle;

    Paragraph paragraph = new Paragraph();
    paragraph.Alignment = alignment == HorizontalAlignment.Right
        ? ParagraphAlignment.Right
        : alignment == HorizontalAlignment.Center ? ParagraphAlignment.Center : ParagraphAlignment.Left;
    paragraph.AddText(text, new TextFormat("SimSun", 9.5f));
    cell.Blocks.Add(paragraph);

    row.Cells.Add(cell);
}
```

`HorizontalAlignment` 是 Excel 侧的枚举，文档模型里的段落对齐用 `ParagraphAlignment`；上面做了映射，直接写文档模型时可以只用 `ParagraphAlignment`。

页脚页码用 `FieldRun`，渲染时替换实际页码：

```csharp
TextFormat small = new TextFormat("SimSun", 8);
Paragraph footer = section.Footer.AddParagraph(string.Empty);
footer.Alignment = ParagraphAlignment.Center;
footer.AddText("第 ", small);
footer.Inlines.Add(new FieldRun("PAGE", "{page}", small));
footer.AddText(" / ", small);
footer.Inlines.Add(new FieldRun("NUMPAGES", "{pages}", small));
footer.AddText(" 页", small);
```

## 3. 渲染输出

```csharp
PdfRenderer renderer = new PdfRenderer();
renderer.Render(document, "out/report.pdf");                          // 文件

using (FileStream stream = File.Create("out/report2.pdf"))
{
    renderer.Render(document, stream);                                 // 流
}

byte[] bytes = renderer.RenderToBytes(document);                       // 字节，适合直接写 HTTP 响应
Console.WriteLine(bytes.Length);
```

## 4. 字体

渲染器默认扫描系统字体目录（Windows Fonts、`%LOCALAPPDATA%\Microsoft\Windows\Fonts`、`/usr/share/fonts` 等，缺失目录跳过），按 `TextFormat.FontFamily` / `EastAsiaFontFamily` 请求族名；族名解析不到时，按字符码位回退到能覆盖它的已注册字体（Noto、Liberation、文泉驿等常见命名）。

显式注册自有字体，避免依赖机器上的字体：

```csharp
FontResolver resolver = new FontResolver();

resolver.RegisterFont("MySans", false, false, @"C:\fonts\MySans-Regular.ttf");
resolver.RegisterFont("MySans", true, false, @"C:\fonts\MySans-Bold.ttf");
resolver.RegisterFont("MySerif", false, false, File.ReadAllBytes("MySerif.ttf"), faceIndex: 0);

PdfRenderer renderer = new PdfRenderer(resolver);
renderer.Render(document, "out/report-fonts.pdf");
```

查询解析结果，便于排查"字没生效"：

```csharp
TrueTypeFace regular = resolver.Resolve("MySans", false, false);
TrueTypeFace cjk = resolver.ResolveFallback(0x4E2D, false, false);
Console.WriteLine(FontResolver.IsCjkCodePoint(0x4E2D));                // True
Console.WriteLine(FontResolver.NormalizeFamily("  Microsoft  YaHei "));  // microsoftyahei
```

`NormalizeFamily` 去掉空格并转小写，用于族名比对；`Resolve` 按族名精确命中，命中不到时 `ResolveFallback` 按码位找能覆盖该字符的字体，两者都返回 `TrueTypeFace`（找不到返回 `null`）。

字体只写出文档里实际用到的字形（子集化），所以体积跟用到的字符数相关，而不是字体文件大小。

嵌入子集的字体描述符按 ISO 32000-1 表 223 写标志位：`Symbolic` 恒置（自定义编码的子集不提供标准拉丁字符集），斜体字面再并上 `Italic` 位。规范里**没有 Bold 位**，粗体只能由 `StemV`（粗体 120、常规 80）传达，因此想让粗体真正变粗，请保证解析命中的是 Bold 字面（用 `RegisterFont(family, bold: true, …)` 显式注册），而不是指望查看器合成描边。

## 5. 解析既有 PDF

```csharp
PdfDocument pdf = PdfDocument.Load("out/report.pdf");       // 也支持 Stream / byte[]

Console.WriteLine(pdf.PageCount);
Console.WriteLine(pdf.Info.Title);
Console.WriteLine(pdf.Info.Creator);
Console.WriteLine(pdf.Info.CreationDate);

foreach (PdfPage page in pdf.Pages)
{
    Console.WriteLine($"p{page.Number} {page.Width:F0}x{page.Height:F0} rot={page.Rotation}");
    foreach (PdfTextLine line in page.GetTextLines())
    {
        Console.WriteLine($"  y={line.Y:F1} x={line.X:F1} [{line.Text}]");
        PdfTextFragment first = line.Fragments.FirstOrDefault();
        if (first != null)
        {
            Console.WriteLine($"    size={first.FontSize:F1} w={first.Width:F1} base={first.Baseline:F1}");
        }
    }
}
```

同一行内的多个片段（`Fragments`）保留各自的坐标与字号，做表格还原或字段定位时比整行文本更可靠。

## 6. 表格识别

```csharp
foreach (PdfPage page in pdf.Pages)
{
    foreach (PdfTable t in page.GetTables())
    {
        Console.WriteLine($"table {t.RowCount}x{t.ColumnCount} at ({t.X:F0},{t.Y:F0}) {t.Width:F0}x{t.Height:F0}");
        Console.WriteLine("  header: " + t.GetCell(0, 0) + " | " + t.GetCell(0, 1));
        Console.WriteLine(t.ToText("\t", "\n"));
    }
}
```

识别基于文本位置与线条，属启发式：合并单元格、无边框表格、跨页表格需要按坐标自行校正。

## 7. 转成可编辑 Word

```csharp
using Wisdom.Word;

OfficeDocument fromPdf = pdf.ToOfficeDocument();
DocxWriter.Write(fromPdf, "out/editable.docx");
```

结果是"内容 + 大致结构"的 Word 文档，不追求与原 PDF 逐像素一致；需要精排版仍应在源数据侧控制模型。

## 8. 对象级编辑与水印

前面几章走"模型 → 渲染"的路径，产物是按模型重排的。`PdfFile` 直接在既有 PDF 的对象层工作：解析交叉表、按对象号取代替、改页面、加水印，再整量或增量写出，原页面的版式与字体资源都不动。

```csharp
using System.Text;
using Wisdom.Pdf;

PdfFile file = PdfFile.Load("out/report.pdf");
file.SetPageRotation(0, 90);                                        // 首页转横向阅读
file.AppendPageContent(0, Encoding.ASCII.GetBytes("q 0.9 G 20 20 m 200 20 l S Q\n"));  // 追加一段画线
byte[] content = file.GetPageContentBytes(0);                       // 原内容 + 新片段，已解码
```

对象级改写还能直接动对象：`GetObject<T>(number)`、`AddObject`、`SetObject`、`RemoveObject`、`MarkModified`
拿到的是 PDF 原始字典，`Trailer` 与 `Catalog` 也可读写；页面内容用 `SetPageContent` 整段替换，
`PrependPageContent` 只在原内容前面插一段。

`SetPageRotation` 收任意角度，写出前一律规整到 0/90/180/270（就近取 90 的倍数，负角先加满 360）：
`/Rotate` 在规范里只允许这四个值，传 45 得到的是 90 而不是 45，想要 45° 的斜跨效果请走水印的
`WatermarkOptions.Angle`，那是内容流里的 `cm` 旋转，与页面 `/Rotate` 是两回事。

水印就在这一层工作：默认斜跨 45°、浅灰、30% 不透明度、背景层，中文自动回退到中文字体。

```csharp
file.AddWatermark("机密文件 CONFIDENTIAL");

WatermarkOptions tile = WatermarkOptions.ForText("内部资料 INTERNAL USE");
tile.Layer = WatermarkLayer.Foreground;             // 压在正文之上
tile.Placement = WatermarkPlacement.Tile;           // Diagonal/Center/Tile/四角/Free
tile.FontSize = 28;
tile.Bold = true;
tile.Angle = 30;                                    // 逆时针角度，斜跨与平铺默认 45
tile.Opacity = 0.15;
file.AddWatermark(tile, 0);                         // 只加给第 1 页
file.Save("out/report-watermarked.pdf");
```

页面方向与图像印章放在一起演示时顺序很重要：先改 `/Rotate` 再加水印，写出端才会按旋转后的阅读方向做补偿。

```csharp
PdfFile second = PdfFile.Load("out/report-watermarked.pdf");
second.SetPageRotation(0, 90);

WatermarkOptions stamp = WatermarkOptions.ForImage(File.ReadAllBytes("seal.png"));
stamp.Placement = WatermarkPlacement.BottomRight;
stamp.ImageWidth = 120;                 // 只给宽度时，高度按图片原始比例推算
stamp.Layer = WatermarkLayer.Foreground;
stamp.Opacity = 0.6;
second.AddWatermark(stamp);
second.SaveIncremental("out/report-stamped.pdf");   // 原字节保留，只追加改动对象
```

「只追加」有一个前提：这份 PDF 的交叉表必须是完好的。如果 `startxref` 或交叉表已经坏到要由全文扫描救回，
内存里的条目就不再对应文件里任何一个交叉表区段，追加段也就没有 `/Prev` 可指，写出去只会少掉未改动的页树与
内容流。这种情况下 `SaveIncremental` / `SaveIncrementalToBytes` 退化为全量重写：文件照样能开、内容照旧不缺，
但原字节不再逐字保留，挂在原字节上的数字签名也随之作废。

其余页面操作：`PageCount`、`GetPageDictionary`、`GetPageResources`、`GetPageSize`、`GetPageRotation`、
`SetPageCropBox`、`RemovePage`、`ReorderPages` 处理单篇文档，`Merge`、`AppendPagesFrom`、`InsertPagesFrom`、
`ExtractPages`、`Split` 跨文档搬页。

```csharp
PdfFile stamped = PdfFile.Load("out/report-stamped.pdf");
IReadOnlyList<PdfFile> parts = stamped.Split(1);            // 每页一份
Console.WriteLine(parts.Count);                             // 页数即份数

PdfFile merged = PdfFile.Merge(parts[0], stamped);          // 首页 + 整篇
Console.WriteLine(merged.PageCount);
```

水印的实现是每个水印一个 Form XObject：所有页共用同一份字体与图像资源（只嵌入一次），每页多出一个与
MediaBox 等大的 XObject 和一段 `q … /Wm Do … Q` 内容流；背景层插到 `/Contents` 前面，前景层追加到后面。
不透明度小于 1 会把版本升到 1.4 并写 `/ExtGState` 与透明 `/Group`；带 `/StructParents` 的标签 PDF 会把水印标为
`/Artifact`，因此不会打断屏幕阅读器的朗读顺序。

注意 `AddWatermark("文字", 0)` 的第二个参数是 `WatermarkLayer`（整数 `0` 会隐式转成背景层），
要给指定页加水印必须走 `AddWatermark(options, 0)`。

## 9. 表单填写与扁平化

AcroForm 同样在对象层工作：读出字段树、写入值、自己生成外观流，最后可以把表单扁平化成固定版式。
本章从上一节的报表开始，先用对象级 API 拼一份四字段表单——只是为了有个可演示的输入，日常场景是直接
`PdfFile.Load("问卷.pdf")` 打开一份现成的表单。

```csharp
using System;
using System.Collections.Generic;
using Wisdom.Pdf;

PdfFile blank = PdfFile.Load("out/report-watermarked.pdf");
PdfReference page = blank.GetPageReference(0);
PdfDictionary pageDict = blank.GetPageDictionary(0);
PdfArray fields = new PdfArray();
PdfArray annots = new PdfArray();

fields.Add(AddField(blank, page, "applicant.name", "Tx", 0, new double[] { 72, 700, 292, 716 }, annots));
fields.Add(AddField(blank, page, "applicant.address", "Tx", 0x800, new double[] { 72, 660, 292, 694 }, annots)); // 0x800 = 多行

PdfReference city = AddField(blank, page, "applicant.city", "Ch", 0x20000, new double[] { 320, 700, 440, 716 }, annots);
PdfDictionary cityDict = blank.GetObject<PdfDictionary>(city.Number);
cityDict.Set("Opt", new PdfArray()
    .Add(new PdfArray().AddName("beijing").Add(PdfString.Unicode("北京 Beijing")))
    .Add(new PdfArray().AddName("shenzhen").Add(PdfString.Unicode("深圳 Shenzhen"))));
blank.MarkModified(city.Number);
fields.Add(city);

// 复选框的勾选态名称取自 /AP /N 的键，所以要显式给出开与关两个状态
PdfReference agree = AddField(blank, page, "applicant.agree", "Btn", 0x40000, new double[] { 72, 636, 84, 648 }, annots);
PdfDictionary box = new PdfDictionary();
box.SetName("Type", "XObject");
box.SetName("Subtype", "Form");
box.Set("BBox", new PdfArray().AddNumber(0).AddNumber(0).AddNumber(12).AddNumber(12));
byte[] nothing = Encoding.ASCII.GetBytes("q\nQ\n");
PdfDictionary states = new PdfDictionary();
states.Set("Off", blank.AddObjectReference(new PdfStream(box.Clone(), nothing, false)));
states.Set("Yes", blank.AddObjectReference(new PdfStream(box.Clone(), nothing, false)));
PdfDictionary agreeDict = blank.GetObject<PdfDictionary>(agree.Number);
agreeDict.Set("AP", new PdfDictionary().Set("N", states));
agreeDict.SetName("AS", "Off");
blank.MarkModified(agree.Number);
fields.Add(agree);

PdfDictionary acroForm = new PdfDictionary();
acroForm.Set("Fields", fields);
acroForm.SetString("DA", "/Helv 11 Tf 0 g");            // 默认外观串：字号 0 表示自动缩放
acroForm.SetBoolean("NeedAppearances", true);           // 请阅读器自行排版新值
pageDict.Set("Annots", annots);
blank.MarkModified(page.Number);
blank.Catalog.Set("AcroForm", blank.AddObjectReference(acroForm));
blank.MarkModified(blank.CatalogNumber);
blank.Save("out/form-blank.pdf");
```

`using System.Text;` 沿用第 8 章的引入。`PdfDictionary.Set` 返回字典自身，所以上面可以写成
`new PdfDictionary().Set("N", states)`。

`AddField` 就是把"字段"和"控件"合成一个字典——终端字段通常自己就是 `/Type /Annot /Subtype /Widget`：

```csharp
static PdfReference AddField(PdfFile file, PdfReference page, string name, string fieldType,
    int flags, double[] rect, PdfArray annots)
{
    PdfDictionary field = new PdfDictionary();
    field.SetName("Type", "Annot");
    field.SetName("Subtype", "Widget");
    field.SetName("FT", fieldType);                 // Tx / Ch / Btn / Sig
    field.SetString("T", name);
    if (flags != 0) field.SetNumber("Ff", flags);
    field.Set("Rect", new PdfArray().AddNumber(rect[0]).AddNumber(rect[1]).AddNumber(rect[2]).AddNumber(rect[3]));
    field.Set("P", page);
    field.SetString("DA", "/Helv 0 Tf 0 g");
    PdfReference reference = file.AddObjectReference(field);
    annots.Add(reference);
    return reference;
}
```

读字段树：`GetForm()` 返回整棵树，`Fields` 是文档顺序的全部节点（含中间层父域），`GetField("父.子")`
按点号名查找（大小写不敏感）。

```csharp
PdfFile file = PdfFile.Load("out/form-blank.pdf");
PdfForm form = file.GetForm();
foreach (PdfFormField field in form.Fields)
    Console.WriteLine(field.FullyQualifiedName + " [" + field.Kind + "] page=" + field.PageIndex
        + " rect=" + (field.Rect.HasValue ? Math.Round(field.Rect.Value.Width) + "x" + Math.Round(field.Rect.Value.Height) : "-"));
```

填写：`FillForm` 一次写多个值并统一生成外观（字体子集在整批里只嵌一次），返回实际写入的字段数；
未知字段名直接跳过，所以可以放心传一份完整的数据表。

```csharp
Dictionary<string, string> values = new Dictionary<string, string>();
values["applicant.name"] = "王小明 Wang Xiaoming";
values["applicant.address"] = "上海市浦东新区世纪大道 100 号\n100 Century Avenue, Pudong";
values["applicant.city"] = "shenzhen";              // 命中 /Opt 导出值时写成名字
values["applicant.agree"] = "true";                 // true/状态名 勾选，false/Off/"" 取消
int filled = file.FillForm(values);
file.Save("out/form-filled.pdf");
Console.WriteLine(filled);                          // 4
```

各类型的取值口径：文本域原样写入（含换行、中文）；组合框/列表框给导出值，可编辑组合框也接受自由文本；
复选框的 `true` 会翻译成该控件自己的勾选态名（本例 `/Yes`）；单选组可以给状态名、选项名或 1 起的序号，
值记在父域 `/V` 上，各选项只同步自己的 `/AS`。签名域在任何填写入口都抛 `InvalidOperationException`——它的 `/V`
是一份 `/Sig` 字典，配合域上的 `/ByteRange` 与 `/Contents` 标定签名区间，写进字符串就是把签好的文档改坏；
批量填写的映射里点了签名域同样抛，抛出时先前各域已写进内存但尚未保存。名单从 `form.FillableFields` 取，
父域、只读域、普通按钮与签名域都不在里面，照着它填保证写得进去。

外观的字号与配色取自 `/DA`，顺序是本域 → 父域 → AcroForm：Acrobat 生成的表单常只在 AcroForm 上写一份文档级默认外观
（如 `/Helv 14 Tf 0 0 1 rg`），字段本域不再重复，这一份同样参与生成外观。`/DA` 里字号写 0 表示由阅读器按控件矩形自动定，
本库按同一口径从默认字号往下缩到放得下为止。配色认 `rg`、灰度 `g` 与 CMYK 的 `k`／`c`：印刷向的稿子常写
`/Helv 12 Tf 0 0.75 0.75 0.25 k`，本库按无 ICC 配置的减法混色换算成 RGB 再写进外观；分隔色（`sc`／`scn`）与自定义
色彩空间认不出来，此时按黑色起排。

只想写值、把排版交给阅读器时用 `file.FillForm(values, false)`——此时 `/NeedAppearances` 会被置为 `true`，
避免"值已改、外观还是旧的"被阅读器显示成旧值；随后调用 `GenerateFormAppearances()` 会生成外观并把它复位。

扁平化：先补一次外观，再把每个控件的 `/AP /N` 按控件矩形画进页面内容，控件注释与字段节点随之移除。

```csharp
PdfFile flat = PdfFile.Load("out/form-filled.pdf");
int flattened = flat.FlattenForm();
flat.Save("out/form-flat.pdf");

PdfFile after = PdfFile.Load("out/form-flat.pdf");
Console.WriteLine(flattened);                       // 4
Console.WriteLine(after.HasForm);                   // False——已经是普通文档
Console.WriteLine(PdfDocument.Load("out/form-flat.pdf").Pages[0].ExtractText().Contains("Wang Xiaoming"));  // True
```

画不出外观的字段不会被移除：密码域、没有 `/AP` 的字段、控件没挂进 `/Annots` 的字段，以及签名域与普通按钮
都保持交互态，扁平化绝不丢值。`FlattenForm()` 的返回值只统计真正落地的字段数。

混合表单（AcroForm + `/XFA`）里阅读器优先按 XFA 模板渲染，填进 AcroForm 的值可能看不见；
`file.RemoveFormXfa()` 整体剥离 XFA，让 AcroForm 成为唯一内容（本包不解析 XFA 模板本身）。

## 10. 解析资源上限

解析不受信任的 PDF 前，先把上限调到与业务相称的规模。超限会中止而不是无界占用内存：解压超限抛 `DecompressionLimitException`，结构非法抛 `InvalidDataException`。

```csharp
PdfReadLimits.MaxDecodedStreamBytes = 64L * 1024 * 1024;   // 默认 256 MB
PdfReadLimits.MaxXrefEntries = 200000;                     // 默认 2000000
PdfReadLimits.MaxObjectStreamObjects = 100000;             // 默认 1000000
PdfReadLimits.MaxPageTreeDepth = 32;                       // 默认 64
PdfReadLimits.MaxObjectNumber = 5000000;                   // 默认 100000000
PdfReadLimits.MaxObjectDepth = 16;                         // 默认 64，数组/字典嵌套层数
PdfReadLimits.MaxGlyphMapEntries = 50000;                  // 默认 300000，单个字体的 CMap 与 /W 宽度表条目

try
{
    PdfDocument risky = PdfDocument.Load("upload.pdf");
    Console.WriteLine(risky.PageCount);
}
catch (Wisdom.Office.Core.DecompressionLimitException ex)
{
    Console.WriteLine("解压超限: " + ex.Message);
}
catch (InvalidDataException ex)
{
    Console.WriteLine("结构非法: " + ex.Message);
}
```

`MaxObjectNumber` 同时约束对象号计数器与交叉表的 `/Size`——谎报 `2147483647` 的文件不再能把编辑后的输出搅坏；`MaxObjectDepth` 挡住畸形词元用超深嵌套打满调用栈；`MaxGlyphMapEntries` 应对的是"每条跨度都合法、加起来把映射表堆到内存耗尽"的 CMap 流，超出部分直接丢弃。流内容按 `OfficePackageLimits.CopyBounded` 分块复制，**不按对象声明的 `/Length` 预分配缓冲区**，所以声明长度造假换不来内存占用。把上限压到不合理的值，同一份正常文件也会被拦：

```csharp
// 上一节生成的 probe.pdf（20 583 字节，1 页）
Console.WriteLine("正常解析页数=" + PdfDocument.Load("out/probe.pdf").Pages.Count);

int saved = PdfReadLimits.MaxObjectNumber;
PdfReadLimits.MaxObjectNumber = 3;
try
{
    PdfDocument.Load("out/probe.pdf");
    Console.WriteLine("未拦截");
}
catch (Exception ex)
{
    Console.WriteLine(ex.GetType().Name + ": " + ex.Message);
}
finally
{
    PdfReadLimits.MaxObjectNumber = saved;
}
```

```text
正常解析页数=1
InvalidDataException: The PDF document catalog is invalid.
```

这些字段是**进程级静态值**，多线程共享，改完务必复原；库内部用它做只读判断，不会自动按文件规模调整。

不是所有结构损坏都走异常：页面树节点声明 `/Type /Pages` 却取不到 `/Kids` 时按空容器处理，这一支贡献零页；对象流表头读到非法的「号 偏移」配对就停止，已经命中的偏移照常可用。字体侧同理，`maxp` 声明零字形的 TrueType 在构造期即被拒绝，嵌入与子集化按既定路径降级，不会把整次保存炸掉。

资源上限不在这类静默降级之列。一页的正文可能由多个流引用、表单可以被反复绘制，单看每一份都合法，合起来却是几十倍的体积，因此解码量越界只在这一刻显出形状：`ExtractText()` / `GetContentBytes()` 会在**单条流解码后超限、整页累计解码量超限、或各段正文合计超限**时抛 `DecompressionLimitException`，内存分配失败也照原样上抛，不会再伪装成「这一页没有文字」。批量索引不可信文件时，请按页包住调用并捕获这一条——跳过该页继续下一本，比让空文本静默进库好。

## 11. 验证产物

```csharp
PdfDocument check = PdfDocument.Load("out/report.pdf");
string text = check.ExtractText();

Console.WriteLine(check.PageCount >= 2);                       // 52 行表格已跨页
Console.WriteLine(text.Contains("月度巡检汇总"));                 // 标题已排入
Console.WriteLine(text.Contains("{page}") == false);           // 页码占位符已替换
Console.WriteLine(text.Contains("配电柜温度与接线端子检查 52"));   // 末行数据存在
```

按行定位比整串匹配更稳（同一行的片段可能被拆进不同 `PdfTextLine`）：

```csharp
int pagesWithHeader = check.Pages.Count(page =>
    page.GetTextLines().Any(l => l.Text.Contains("用时(min)")));
Console.WriteLine($"{pagesWithHeader}/{check.PageCount}");
```

本例的实际输出：`pages=2`、`hasLastRow=True`、`pagesWithHeader=2/2`（表头在每页重复）、页脚每页都能匹配到 `第 N /`。注意 `PdfTextLine.Text` 会把同一行的多个单元格用空格连起来，例如 `序号 巡检项 结果 用时(min)`，所以按整段原文匹配容易失败，建议按 `Contains` 或逐片段坐标判断。

## 常见问题

**生成的 PDF 是空白或方块。** 系统里没有可用 TrueType 字体时无法嵌入字形。容器镜像里请安装 `fonts-noto-cjk` 或文泉驿，或用 `FontResolver.RegisterFont` 指向随应用发布的字体文件。

**要求支持 CFF/OpenType。** 目前嵌入只支持 TrueType `glyf` 轮廓；`.otf` 里若是 CFF 轮廓不能作为嵌入源（部分以 `.otf` 后缀发布的 TrueType 字体可正常使用）。

**加密 PDF 读不了。** 口令与权限加密走标准安全处理器：`PdfFile.Load(path, "口令")` 可开 RC4 40/128、AES-128、AES-256（R5 与 R6）的文档，口令错时抛 `PdfPasswordException`；写出时 `PdfEncryptionAlgorithm.Aes256` 生成 R6，四种算法都已与第三方阅读器双向验证。有些工具把 `/P` 写成 `4294967292` 这样的无符号值，读取端会按低 32 位还原，不影响开箱。只支持标准处理器，公钥加密与 Adobe 自定义安全处理器不在范围内。

**填了值但阅读器里看不到。** 三种原因：文档是混合表单（带 `/XFA`），阅读器按 XFA 模板渲染，`RemoveFormXfa()` 剥离即可；只调 `FillForm(values, false)` 且阅读器不重建外观，本库已把 `/NeedAppearances` 置为 `true`，个别阅读器仍会忽略，此时改用 `FillForm(values)` 或显式 `GenerateFormAppearances()`；字段是密码域，外观按规范留空，值只在 `/V` 里。

**填进去的值生成的外观跟表单其余部分不像同一份稿子（字号偏小、颜色发黑）。** 已修。这类表单的 `/DA` 只写在 AcroForm 上作全文档默认，字段本域没有；读端过去沿父域走到顶就停，把这份默认读成空，于是生成外观时字号退回自动、颜色一律黑。现在按本域 → 父域 → AcroForm 取，`field.DefaultAppearance` 也一并返回继承来的那一份。

**单选组填了值但没有任何一项被选中。** 传入的必须是某一项的勾选态名（`/AP /N` 里的键）、选项部分名，或 1 起的序号；本库不会猜中文标签，遍历 `Children` 取 `OnState`（为空时退回 `Name`）就能列出合法取值。

**别人的表单打开再存一次，选项名对不上了。** 已修。名称在文件里是字节：写出端把非 ASCII 名称按 UTF-8 展开成 `#HH`，读回端过去却一个字节还原一个字符，「中」回来成一串拉丁字母，载进来再存盘就变形一次；星外字符（成对代理）更是每个半边退化成 `?`，直接在文件里丢掉。现在两端成对——读回按 UTF-8 还原，写出的代理对按整个码位展开；把高位写成单字节的历史名称仍按逐字节读，不会多出替换符。同类口径也用于字体名与任何 `/Name` 键值。

**批量填表单抛出 "Signature fields cannot be filled."。** 映射里点了签名域的名字，这是有意的拒绝：签名域的 `/V` 是一份 `/Sig` 字典，域上还有 `/ByteRange` 与 `/Contents` 标定签名在文件里的字节区间，写进一个字符串等于把签好的文档改坏。本库只读不填签名域，`SetFormFieldValue`、`FillForm` 与 `field.Value` 三条入口同一口径。名单改用 `form.FillableFields`，它不列签名域与普通按钮。

**加水印后正文被盖住了。** 默认水印在背景层（`WatermarkLayer.Background`，排在原内容之前）；显式设成 `Foreground` 才会压在正文上，此时把 `Opacity` 调低一些。

**水印文字提取出来是散的。** 旋转文字按流向归并，但若用 `GetTables()` 看旋转页，表格识别仍以轴对齐坐标为准；正文提取不受影响。

**别人的 PDF 在阅读器里显示正常，抽出来却是空的。** 先看那份稿子的正文流怎么声明过滤器：规范允许把 `/Filter` 缩写成 `/F`、把 `/DecodeParms` 缩写成 `/DP`，本库现已两个都认（此前只认 `/DP`，`<< /F /Fl >>` 的压缩正文被当成未过滤原字节，整页文字静默为空）。`/F` 取不出过滤器名时按无过滤器处理，富文本流与加密转换里那个写作 `/F` 的整数标志不会被误认。仍读不出来的多是图像型页面，见下一条。

**扫描件没有文本。** 提取依赖文本操作符，图像型页面需要外部 OCR。

**页码显示成 `{page}`。** 说明这段文本走的是非渲染路径（例如直接从模型取 `DisplayText`）；只有 `PdfRenderer` 渲染时才把 `{page}` / `{pages}` 换成实际数字，DOCX 侧则会输出真实的 `PAGE` / `NUMPAGES` 域。

**给第三方 PDF 逐页登记资源，另一页的同名条目被顶掉了。** 已修。这类 PDF 常把资源字典挂在 `/Pages` 上由全篇继承，那份实例同时是每个兄弟页的资源字典；`SetPageResource` 先前把它钉在本页上就地落笔，等于一次登记替全篇生效——后一页的 `/Logo` 覆盖前一页的，第三页什么都没登记也跟着多出这个分组。现在缺资源的页先按页复制一份（连同 `/Font`、`/XObject` 这些分组一起脱开）再写，继承来的条目照常可读，各页同名资源互不干扰，多页水印的资源名也不再被串页逼出 `_1` 后缀。
