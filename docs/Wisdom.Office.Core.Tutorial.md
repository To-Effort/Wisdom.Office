# Wisdom.Office.Core 使用教程

内核包的实战教程：手工搭出一份带页眉页脚、合并单元格与页码域的文档模型，用工作簿模型描述一张带格式的表，用演示文稿模型搭出幻灯片，看 Markdown 落到模型的哪些成员上，绑定模板上下文，并直接复用底层的 OLE2 / ZIP / 压缩工具。安装方式见[包说明](../src/Wisdom.Office.Core/README.md#安装)。

## 你将完成

```text
out/
  model.docx     由手工模型生成的 Word 文件（页眉、页码域、跨页表格）
  model.pdf      同一份模型的 PDF 版（页码已替换成实际页码）
  model.xlsx     工作簿模型生成的 Excel 文件
  deck.pptx      演示文稿模型写出的幻灯片（3 页）
  deck.pdf       同一份演示文稿按"每页一节"导出的 PDF
  deck.docx      同一份演示文稿导出的 Word
  template.docx  带 {{key}} 与 {{#each}} 标记的 Word 模板
  filled.docx    用模板上下文填充后的 Word
  sample.png     任意一张 PNG，供图片示例使用（本文用它测得 240x120 / 592 字节）
```

内核不含读写器，所以本教程在"看效果"的那几步借助 `Wisdom.Office`（`OfficeConverter`）、`Wisdom.Excel`（`XlsxWriter`、`XlsxTemplate`）、`Wisdom.Ppt`（`PptxWriter`、`PptxReader`）、`Wisdom.Word`（`DocxTemplate`）与 `Wisdom.Markdown` 落盘；模型代码本身只依赖 `Wisdom.Office.Core`。

## 1. 建立文档骨架

`OfficeDocument` → `Section` → `Block`；`Block` 有 `Paragraph`、`Table`、`PageBreak` 三种。

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using Wisdom.Office;
using Wisdom.Office.Core;

OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "成品检验报告";
document.Metadata.Author = "质检部";
document.Metadata.Subject = "乳胶漆批次检验";
document.Metadata.Keywords = "检验,报告,批次";

PageSetup setup = PageSetup.A4();
setup.MarginLeft = setup.MarginRight = 56;
setup.MarginTop = setup.MarginBottom = 64;
setup.HeaderDistance = 32;
setup.FooterDistance = 32;

Section section = document.AddSection(setup);
```

`PageSetup.IsLandscape` 赋值时会自动交换宽高，不需要手工改 `Width`/`Height`：

```csharp
PageSetup wide = PageSetup.A4();
wide.IsLandscape = true;      // 595.28x841.89 -> 841.89x595.28
```

## 2. 页眉与页码域

页眉页脚各自是一个 `HeaderFooter`，里面同样放 `Block`。`FieldRun` 写 DOCX 时输出真实域代码，写 PDF 时把 `DisplayText` 里的 `{page}` / `{pages}` 替换成实际页码——同一份模型两种输出都对。

```csharp
TextFormat meta = new TextFormat("SimSun", 8);
meta.EastAsiaFontFamily = "宋体";

Paragraph header = section.Header.AddParagraph("文件编号：QF-QC-018　版本：A/2");
header.Alignment = ParagraphAlignment.Right;

Paragraph footer = section.Footer.AddParagraph(string.Empty);
footer.Alignment = ParagraphAlignment.Center;
footer.AddText("第 ", meta);
footer.Inlines.Add(new FieldRun("PAGE", "{page}", meta));
footer.AddText(" 页　共 ", meta);
footer.Inlines.Add(new FieldRun("NUMPAGES", "{pages}", meta));
footer.AddText(" 页", meta);
```

## 3. 段落与文本格式

`TextFormat` 的西文族名与中文族名分开，中英混排时按字符分别取字体。

```csharp
Paragraph title = new Paragraph("成品检验报告");
title.HeadingLevel = 1;
title.Alignment = ParagraphAlignment.Center;
title.SpaceAfter = 6;
title.KeepWithNext = true;
section.Blocks.Add(title);

Paragraph body = new Paragraph();
body.AddText("委托方：", InspectionFormat(10, true));
body.AddText("示例建材有限公司　", InspectionFormat(10, false));
body.AddText("依据标准：", InspectionFormat(10, true));
body.AddText("GB/T 9756-2018", InspectionFormat(10, false));
body.FirstLineIndent = 21;            // 首行缩进两个中文字（磅）
body.LineSpacing = 18;
section.Blocks.Add(body);
```

辅助方法（后面几节继续用）：

```csharp
static TextFormat InspectionFormat(double size, bool bold)
{
    TextFormat format = new TextFormat("SimSun", size);
    format.EastAsiaFontFamily = "宋体";
    format.Bold = bold;
    return format;
}
```

## 4. 合并单元格表格

两行表头 + 跨列跨行：`TableCell.ColSpan` / `RowSpan` 描述跨度，`TableRow.IsHeader` 标记表头行（跨页时会在每页顶部重复），`TableRow.Height` 给数据行留手写空间。

```csharp
Table table = new Table();
table.CellPadding = 3;
table.RepeatHeaderRow = true;
foreach (double width in new double[] { 28, 84, 104, 64, 58, 58, 46 })
{
    table.Columns.Add(new TableColumn(width));
}

OfficeColor label = OfficeColor.FromHtml("#EEF2F5");

TableRow header = table.AddRow();               // AddRow 已把行加入表格
header.IsHeader = true;
AddSpanCell(header, "序号", 1, 2, label);
AddSpanCell(header, "检验项目", 1, 2, label);
AddSpanCell(header, "技术要求", 1, 2, label);
AddSpanCell(header, "检验方法", 1, 2, label);
AddSpanCell(header, "实测结果", 2, 1, label);   // 跨两列，纵向不合并
AddSpanCell(header, "判定", 1, 2, label);

TableRow sub = table.AddRow();                  // 表头第二行：只补被跨列覆盖的那两列
sub.IsHeader = true;
AddSpanCell(sub, "单件一", 1, 1, label);
AddSpanCell(sub, "单件二", 1, 1, label);

string[] items =
{
    "容器中状态", "细度", "施工性", "干燥时间", "遮盖力", "耐洗刷性",
    "附着强度", "耐候性", "耐碱性", "耐水性"
};

for (int i = 0; i < 30; i++)
{
    TableRow row = table.AddRow();
    row.Height = 24;                            // 留出填写空间，同时保证表格跨页
    AddSpanCell(row, (i + 1).ToString(), 1, 1, null);
    AddSpanCell(row, items[i % items.Length], 1, 1, null);
    AddSpanCell(row, "符合 GB/T 9756 优等品", 1, 1, null);
    AddSpanCell(row, "GB/T 9756", 1, 1, null);
    AddSpanCell(row, (86 + i % 10).ToString(), 1, 1, null);
    AddSpanCell(row, (88 + i % 8).ToString(), 1, 1, null);
    AddSpanCell(row, i == 5 ? "不合格" : "合格", 1, 1,
        i == 5 ? OfficeColor.FromHtml("#FFE6E6") : null);
}

section.Blocks.Add(table);
```

单元格辅助方法——三种情况（普通格、跨列格、跨行格）都走同一个入口：

```csharp
static TableCell AddSpanCell(TableRow row, string text, int colSpan, int rowSpan, OfficeColor? background)
{
    TableCell cell = new TableCell();
    cell.ColSpan = colSpan;
    cell.RowSpan = rowSpan;
    cell.VerticalAlignment = VerticalAlignment.Middle;
    cell.BackgroundColor = background;

    Paragraph paragraph = new Paragraph();
    paragraph.Alignment = ParagraphAlignment.Center;
    paragraph.AddText(text, InspectionFormat(9.5f, background != null));
    cell.Blocks.Add(paragraph);

    row.Cells.Add(cell);
    return cell;
}
```

**关键约束**：`RowSpan = 2` 的单元格，其覆盖到的下一行必须不再为该列放单元格。上面表头第二行只有"单件一 / 单件二"两格，正是因为序号、检验项目、技术要求、检验方法、判定五列都被第一行的跨行格占了位；写错会导致 DOCX 合并结构与 Word 显示不一致。

## 5. 图片、分页与第二个节

```csharp
Paragraph picture = new Paragraph();
picture.Alignment = ParagraphAlignment.Center;
picture.AddImage(OfficeImage.FromFile("sample.png"), 120, 60);
section.Blocks.Add(picture);

section.Blocks.Add(new PageBreak());

Section appendix = document.AddSection(PageSetup.A4());
appendix.Header.AddParagraph("附件").Alignment = ParagraphAlignment.Center;
appendix.AddParagraph("附录 A：检测点位示意");
```

`AddParagraph` 是 `Section` 上的便捷方法，等价于"新建段落并追加到 `Blocks`"；需要设属性时直接用 `section.Blocks.Add(...)`。

## 6. 落盘验证

```csharp
using System.Linq;

Directory.CreateDirectory("out");
OfficeConverter.ToWord(document, "out/model.docx");
OfficeConverter.ToPdf(document, "out/model.pdf");

OfficeFile docx = OfficeFile.Load("out/model.docx");
string headerText = string.Join(" ", docx.Document.FirstSection.Header.Blocks
    .OfType<Paragraph>().Select(p => p.GetText()));
Console.WriteLine(headerText.Contains("QF-QC-018"));    // True：页眉已回读

OfficeFile pdf = OfficeFile.Load("out/model.pdf");
string pdfText = pdf.ExtractText();
Console.WriteLine(pdf.Pdf.PageCount >= 2);              // True：30 行表格已跨页（本例 4 页）
Console.WriteLine(pdfText.Contains("{page}") == false); // True：占位符已替换成实际页码
```

注意 `PageCount` 在 PDF 模型上：`OfficeFile` 暴露的是 `file.Pdf`（`PdfDocument`），`file.Pdf.PageCount` 才是页数；`file.ExtractText()` 则已代理到对应格式。

实测：`model.docx` 7,647 字节、`model.pdf` 61,856 字节、`PageCount=4`（30 行表格 + 附录页），页眉回读含 `QF-QC-018`，文本层里取不到 `{page}` 且出现"第 1 页"。

## 7. 工作簿模型

```csharp
OfficeWorkbook workbook = new OfficeWorkbook();
workbook.Metadata.Title = "库存台账";

Worksheet sheet = workbook.AddWorksheet("台账");
sheet.GetCell(1, 1).Value = "物料";
sheet.GetCell(1, 2).Value = "数量";
sheet.GetCell(1, 3).Value = "金额";
sheet.GetCell(1, 4).Value = "入库日";

sheet.GetCell(2, 1).Value = "主板 M-200";
sheet.GetCell(2, 2).Value = 40;
sheet.GetCell(2, 3).Value = 12800.50;
sheet.GetCell(2, 4).Value = new DateTime(2026, 9, 1);

sheet.GetCell(2, 3).Style = new CellStyle
{
    NumberFormat = "#,##0.00",
    HorizontalAlignment = HorizontalAlignment.Right,
    FillColor = OfficeColor.FromHtml("#F2F7FC"),
    Bold = true
};
sheet.GetCell(2, 4).Style = new CellStyle { NumberFormat = "yyyy-MM-dd" };

sheet.Merge("A5:D5");
sheet.GetCell(5, 1).Value = "备注：金额为含税成本。";
sheet.FrozenRows = 1;
sheet.SetColumnWidth(1, 22);
sheet.SetRowHeight(5, 28);
sheet.ShowGridLines = false;
sheet.SetPrintArea("A1:D30");
```

`Cell.Value` 支持 `string`、`double`、`bool`、`DateTime`；`DisplayText()` 按 `CellStyle.NumberFormat` 给出与 Excel 一致的显示串（含 `正;负;零;文本` 分节）。

```csharp
Cell amount = sheet.GetCell(2, 3);
Console.WriteLine(amount.DisplayText());                 // 12,800.50

Console.WriteLine(sheet.LastRow + " x " + sheet.LastColumn);
Console.WriteLine(sheet.FindMergeAnchor(5, 2).Value);     // 备注所在合并区的锚点单元格
```

写盘由 `Wisdom.Excel` 负责（模型自己不含 IO）：

```csharp
using Wisdom.Excel;

XlsxWriter.Write(workbook, "out/model.xlsx");              // 4,064 字节
```

## 8. 演示文稿模型（PowerPoint）

层级是 `OfficePresentation` → `Slide` → `SlideShape`，形状派生出 `TextBoxShape`、`AutoShape`、`PictureShape`、`TableShape`。几何量一律是**磅**，原点在幻灯片左上角；本包只定义模型，写出与回读由 `Wisdom.Ppt` 完成。

```csharp
using System;
using System.Linq;
using Wisdom.Office.Core;
using Wisdom.Ppt;

OfficePresentation presentation = OfficePresentation.Widescreen();   // 960 x 540
Console.WriteLine(presentation.SlideWidth + "x" + presentation.SlideHeight);   // 960x540
Console.WriteLine(OfficePresentation.Standard().SlideWidth);                    // 720（4:3）
Console.WriteLine(presentation.DefaultTextFormat.FontSize);                     // 18（Calibri）
Console.WriteLine(presentation.DefaultTitleFormat.FontSize + "/" + presentation.DefaultTitleFormat.Bold); // 40/True
Console.WriteLine(presentation.AccentColors.Count);                             // 6 个强调色
```

`AddSlide(SlideLayoutKind)` 按幻灯片尺寸等比推算占位符：左右边距 8%、标题上边距 6% 高 20%、内容区从 30% 起到下边距。实测七种版式：

| `SlideLayoutKind` | 形状（`Name` + `PlaceholderIndex`） | `Placeholder` | 几何（磅） |
| --- | --- | --- | --- |
| `Title` | `标题#0` | `Title` | 左 76.8 上 32.4 宽 806.4 高 108 |
| `TitleAndContent` | `标题#0` + `内容占位符#1` | `Title` + `Content` | 内容 上 162 高 313.2 |
| `SectionHeader` | `标题#0` | `Title` | 上 183.6 高 140.4（下移居中） |
| `TwoContent` | `标题#0` + `内容占位符 1#1` + `内容占位符 2#2` | `Title` + `Content`×2 | 左右各宽 394.2，第二栏左 489 |
| `TitleAndPicture` | `标题#0` + `内容占位符#1` | `Title` + `Picture` | 内容同 `TitleAndContent` |
| `TitleAndTable` | `标题#0` + `内容占位符#1` | `Title` + `Table` | 内容同 `TitleAndContent` |
| `Blank` | 无形状 | —— | —— |

占位符的名字是模型自动分配的中文（`标题`、`内容占位符`），类别看 `Placeholder` 枚举，配对序号看 `PlaceholderIndex`（非占位符形状为 `-1`，`Placeholder` 为 `None`）。

```csharp
Slide cover = presentation.AddSlide(SlideLayoutKind.Title);
Console.WriteLine(cover.Shapes[0].Name + "#" + cover.Shapes[0].PlaceholderIndex);   // 标题#0
((TextBoxShape)cover.Shapes[0]).Paragraphs[0].AddText("2026 年第三季度");
cover.NotesText = "开场 2 分钟。";
cover.Transition = SlideTransitionKind.Fade;
cover.TransitionDuration = 1.5;

Slide body = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
((TextBoxShape)body.Shapes[0]).Paragraphs[0].AddText("业务进展");
TextBoxShape content = (TextBoxShape)body.Shapes[1];
content.Paragraphs.Clear();                       // 版式自带一个空段落，先清掉
content.AddParagraph("完成报表引擎");
content.AddParagraph("支持 Markdown 解析");
Console.WriteLine(content.GetText());             // 完成报表引擎\n支持 Markdown 解析
```

`AddSlide()` 不带参数等价于 `Blank`，全部自己定位。文本框、自选图形、图片、表格四个入口都在 `Slide` 上：

```csharp
Slide free = presentation.AddSlide();
Console.WriteLine(free.Layout);                                                    // Blank
TextBoxShape note = free.AddTextBox(60, 180, 200, 60, "小注");
note.VerticalAlignment = SlideTextAnchor.Middle;
note.MarginLeft = note.MarginRight = note.MarginTop = note.MarginBottom = 4;

AutoShape box = free.AddShape(AutoShapeKind.RoundedRectangle, 60, 60, 200, 80, "圆角矩形");
box.FillColor = OfficeColor.FromHtml("#00B050");
Console.WriteLine(box.CornerRadius);                       // 0.15 —— 圆角比例，按短边归一
Console.WriteLine(box.FillColor.ToHex());                  // 00B050

OfficeImage logo = OfficeImage.FromFile("sample.png");     // 本例 240x120 PNG，592 字节
Console.WriteLine(logo.PixelWidth + "x" + logo.PixelHeight + " " + logo.Format);   // 240x120 Png
free.AddImage(logo, 320, 60, 120, 120);

TableShape table = free.AddTable(60, 260, 300, 90);
table.Table.AddRow("模块", "状态");                // 复用文档模型的 Table
table.Table.AddRow("PPT", "已完成");
free.BackgroundColor = OfficeColor.FromHtml("#F2F7FF");

Console.WriteLine(string.Join(",", free.Shapes.Select(s => s.GetType().Name)));
// TextBoxShape,AutoShape,PictureShape,TableShape —— 按添加顺序
Console.WriteLine(note.GetText() + " / " + box.GetText());   // 小注 / 圆角矩形
```

三个入口的最后一个参数都是**形状文本**，不是形状名称：手工添加的形状 `Name` 保持空串，只有版式占位符才带自动分配的中文名。

写出、回读，再用导出器把每页排成一节：

```csharp
PptxWriter.Write(presentation, "out/deck.pptx");            // 3 页 = 13,214 字节

OfficePresentation back = PptxReader.Load("out/deck.pptx");
Console.WriteLine(back.Slides.Count);                        // 3
Console.WriteLine(back.Slides[0].GetTitle());                 // 2026 年第三季度
Console.WriteLine(back.Slides[0].Transition + "/" + back.Slides[0].TransitionDuration);  // Fade/1.5
Console.WriteLine(back.Slides[0].Layout);                     // Title
Console.WriteLine(back.Slides[2].Shapes.Count);               // 4
Console.WriteLine(back.Slides[2].BackgroundColor.ToHex());    // F2F7FF
```

```csharp
AutoShape backBox = back.Slides[2].Shapes.OfType<AutoShape>().First();
Console.WriteLine(backBox.Kind + " " + backBox.Left + " " + backBox.CornerRadius + " " + backBox.FillColor.ToHex());
// RoundedRectangle 60 0.15 00B050

PictureShape backPicture = back.Slides[2].Shapes.OfType<PictureShape>().First();
Console.WriteLine(backPicture.Image.Data.Length);             // 592 —— 图片字节原样回来
TableShape backTable = back.Slides[2].Shapes.OfType<TableShape>().First();
Console.WriteLine(backTable.Table.Rows.Count + "x" + backTable.Table.Columns.Count);   // 2x2
```

`2x2` 是回读补齐的结果：`Table.Columns` 只是列占位，写出时不入库，读回时按每行单元格数补齐。`PlaceholderIndex`、文本、几何、填充色、切换与备注都能写出读回；版式是"按形状推断"的，`SectionHeader` 回读成 `Title`、带真实表格形状的 `TitleAndContent` 页回读成 `TitleAndTable`。

导出 PDF / Word 由 `Wisdom.Office` 的 `PresentationToDocumentConverter` 按"每页一节"排版，模型层不含 IO：

```csharp
using Wisdom.Office;

Console.WriteLine(PresentationToDocumentConverter.Convert(presentation).Sections.Count); // 3
OfficeConverter.PresentationToPdf(presentation, "out/deck.pdf");     // 51,628 字节
OfficeConverter.PresentationToWord(presentation, "out/deck.docx");   // 4,417 字节
```

排版是线性的：不还原绝对坐标、母版与动画，自选图形与图片按形状文本降级输出。

## 9. Markdown 在这套模型里长什么样

`Wisdom.Markdown` 解析出来的东西就是本包的文档模型，没有第二套类型——理解映射关系才能手工改模板。

```csharp
using Wisdom.Markdown;

string markdown = "# 标题一\n\n段落带[链接](https://example.com/a)与 **加粗**。\n\n"
    + "- 普通项\n- [x] 已完成项\n\n## 标题二\n\n| A | B |\n| --- | --- |\n| 1 | 2 |\n";

OfficeDocument doc = Markdown.ToDocument(markdown);
Console.WriteLine(doc.Sections[0].Blocks.Count);                     // 6
Console.WriteLine(string.Join(",", doc.Sections[0].Blocks.OfType<Paragraph>()
    .Where(p => p.HeadingLevel > 0).Select(p => p.HeadingLevel + ":" + p.GetText())));
// 1:标题一,2:标题二
```

| Markdown 语法 | 落在模型的哪个成员 |
| --- | --- |
| `#` / `##` | `Paragraph.HeadingLevel`（1..6），`TextFormat.Heading(level)` 提供配套字号 |
| `[文字](url)` | `TextRun.Link` = url，文字仍是普通 `Text` |
| `**粗体**` | `TextRun.Format.Bold` |
| `- 项` | `Paragraph.ListMarker` = `•` |
| `- [x] 项` | `Paragraph.ListMarker` = `☑ •`（未完成是 `☐ •`） |
| GFM 表格 | `Table`，首行 `TableRow.IsHeader = true`，`Columns` 按列数补齐（实测 2x2） |

```csharp
TextRun link = doc.Sections[0].Blocks.OfType<Paragraph>()
    .SelectMany(p => p.Inlines).OfType<TextRun>().First(t => t.Link != null);
Console.WriteLine(link.Text + " -> " + link.Link);            // 链接 -> https://example.com/a
Console.WriteLine(string.Join(",", doc.Sections[0].Blocks.OfType<Paragraph>()
    .Where(p => p.ListMarker != null).Select(p => "[" + p.ListMarker + "]")));
// [•],[☑ •]
```

`Link` 只有 Markdown 与 HTML 输出会消费：`Markdown.ToHtml` 写成 `<a href="…">`（实测含 `<input type="checkbox"` 任务框），DOCX 写出与 PDF 排版只留链接文字、丢弃地址。

`HeadingLevel` 同时是拆页依据。`DocumentToPresentationConverter` 按标题层级切幻灯片（在 `Wisdom.Office` 里）：

```csharp
using Wisdom.Office;

OfficePresentation split = DocumentToPresentationConverter.Convert(doc);
Console.WriteLine(split.Slides.Count);                       // 2 —— 两个一级/二级标题段
Console.WriteLine(split.Slides[0].Layout);                    // TitleAndContent
Console.WriteLine(split.Slides[0].GetTitle());                // 标题一
Console.WriteLine(split.Slides[1].Shapes.Count);              // 3 —— 标题 + 正文 + GFM 表格变 TableShape
Console.WriteLine(DocumentToPresentationConverter
    .Convert(Markdown.ToDocument("只有正文，没有标题。")).Slides.Count);   // 1 —— 无标题时整篇一页
```

## 10. 模板上下文

```csharp
TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["order"] = new { id = "SO-1024", total = 398.0 };
context["items"] = new[]
{
    new { name = "无线键盘", qty = 2, price = 199.0 },
    new { name = "机械鼠标", qty = 5, price = 89.9 }
};

Console.WriteLine(TemplateEngine.Replace(
    "{{customerName}} 的订单 {{order.id}} 合计 {{order.total:0.00}} 元", context));
// 示例客户 的订单 SO-1024 合计 398.00 元

TemplateOptions options = new TemplateOptions
{
    LeaveUnresolvedTokens = false,
    HtmlEncodeValues = true,
    Culture = System.Globalization.CultureInfo.InvariantCulture
};

Console.WriteLine(TemplateEngine.Replace("{{missing}}", context, options)); // 空串（显式清掉）
Console.WriteLine(TemplateEngine.Replace("{{missing}}", context));           // {{missing}} —— 默认保留
Console.WriteLine(TemplateEngine.ContainsToken("{{a}}"));                   // True

Console.WriteLine(TemplateEngine.ContainsLoopStart("{{#each items}}"));      // True
if (TemplateEngine.TryGetLoopName("{{#each items}}", out string loopName))
{
    Console.WriteLine(loopName);                                            // items
}
Console.WriteLine(TemplateEngine.IsLoopEnd("{{/each}}"));                    // True
Console.WriteLine(TemplateEngine.RemoveLoopMarkers("{{#each items}}x{{/each}}"));  // x
```

未解析的占位符行为由 `LeaveUnresolvedTokens` 决定：默认 `true` 原样保留，`false` 替换为空串。值一律经过 XML 转义后再写入包，`HtmlEncodeValues` 控制的是输出内容本身是否再做一次转义（用于把值嵌进富文本的场景）。

**`Replace` 只做标量替换，不展开循环。** 把 `{{#each}}` 交给它是原样返回——这两个方法分别用于"文本模板"和"包模板"，走的是不同代码路径：

```csharp
Console.WriteLine(TemplateEngine.Replace("{{#each items}}{{name}}{{/each}}", context));
// {{#each items}}{{name}}{{/each}} —— 标记完整保留，交给填充器处理
```

循环标记由 `Wisdom.Word` 的 `DocxTemplate` 与 `Wisdom.Excel` 的 `XlsxTemplate` 消费：它们扫描包里的标记行/单元格，用 `ResolveList` 解析数据源、`CreateRowContext` 为每一项建上下文（`{{@index}}` 从 0、`{{@number}}` 从 1、`{{@first}}`、`{{@last}}`、`{{.}}` 当前项本身），再克隆承载元素。所以模板里 **`{{#each}}` 与 `{{/each}}` 各自独占一行**，中间那行才是重复的行。

```csharp
using Wisdom.Excel;
using Wisdom.Word;

OfficeWorkbook tpl = new OfficeWorkbook();
Worksheet tsheet = tpl.AddWorksheet("清单");
tsheet.GetCell(1, 1).Value = "合计：{{order.total:0.00}}";
tsheet.GetCell(2, 1).Value = "{{#each items}}";
tsheet.GetCell(3, 1).Value = "{{@number}}";
tsheet.GetCell(3, 2).Value = "{{name}}";
tsheet.GetCell(3, 3).Value = "{{qty}}";
tsheet.GetCell(4, 1).Value = "{{/each}}";

OfficeWorkbook filled = XlsxTemplate.Fill(tpl, context);
Worksheet fs = filled.Worksheets[0];
Console.WriteLine(fs.LastRow);                       // 3 —— 两行标记被吃掉，数据行复制成 2 行
Console.WriteLine(fs.GetCell(1, 1).DisplayText());   // 合计：398.00
Console.WriteLine(fs.GetCell(2, 2).DisplayText() + " " + fs.GetCell(2, 3).DisplayText());   // 无线键盘 2
```

Word 侧同一套语法，模板用模型搭出来即可（先 `OfficeConverter.ToWord` 存成 `.docx` 就能拿去当模板）：

```csharp
OfficeDocument docTpl = new OfficeDocument();
Section dsec = docTpl.AddSection(PageSetup.A4());
dsec.AddParagraph("客户：{{customerName}}　订单：{{order.id}}");
Table dtable = new Table();
dtable.Columns.Add(new TableColumn(120));
dtable.Columns.Add(new TableColumn(60));
dtable.AddRow("品名", "数量");
dtable.AddRow("{{#each items}}", string.Empty);
dtable.AddRow("{{name}}", "{{qty}}");
dtable.AddRow("{{/each}}", string.Empty);
dsec.Blocks.Add(dtable);

byte[] templateBytes = OfficeConverter.ToWord(docTpl);          // 3,693 字节
byte[] output = DocxTemplate.Fill(templateBytes, context);      // 3,694 字节

OfficeDocument result = OfficeFile.Load(output).Document;
Console.WriteLine(result.Sections[0].Blocks.Count);             // 3
Console.WriteLine(result.Sections[0].Blocks.OfType<Paragraph>().First().GetText());
// 客户：示例客户　订单：SO-1024
Table filledTable = result.Sections[0].Blocks.OfType<Table>().First();
Console.WriteLine(filledTable.Rows.Count);                       // 3 —— 表头 + 两条明细
Console.WriteLine(filledTable.Rows[1].Cells[0].Blocks.OfType<Paragraph>().First().GetText());   // 无线键盘
```

## 11. 直接用底层工具

读取 OLE2 复合文档（`.doc` / `.xls` 的容器）：

```csharp
byte[] legacyBytes = File.ReadAllBytes("legacy.xls");

Console.WriteLine(CompoundFile.IsCompoundFile(legacyBytes));   // True —— 先看签名再构造

CompoundFile compound = new CompoundFile(legacyBytes);
Console.WriteLine(compound.HasStream("Workbook"));             // True —— 据此区分 XLS 与 DOC
Console.WriteLine(string.Join(",", compound.GetStreamNames()));
// Workbook,SummaryInformation,DocumentSummaryInformation

byte[] workbookStream = compound.GetStream("Workbook");        // 取流内容，实测 4,096 字节；不存在时 TryGetStream 返回 null
```

`.doc` 的容器结构不同，用它来区分两类旧格式（`HasStream("Workbook")` 为 `False`）：

```csharp
CompoundFile doc = new CompoundFile(File.ReadAllBytes("legacy.doc"));
Console.WriteLine(doc.HasStream("Workbook"));                    // False
Console.WriteLine(string.Join(",", doc.GetStreamNames()));
// Data,1Table,CompObj,WordDocument,SummaryInformation,DocumentSummaryInformation
```

安全地打开 OOXML 包（注意入参必须是 zip 容器，`.xls`/`.doc` 这类 OLE2 文件要用上面的 `CompoundFile`）并限制解压规模：

```csharp
byte[] packageBytes = File.ReadAllBytes("out/model.xlsx");

using (System.IO.Compression.ZipArchive archive =
    PackageArchive.OpenForRead(new MemoryStream(packageBytes), false))
{
    long totalBytes = 0;
    foreach (System.IO.Compression.ZipArchiveEntry entry in archive.Entries)
    {
        byte[] part = OfficePackageLimits.ReadEntryBytes(entry, ref totalBytes);
        Console.WriteLine($"{entry.FullName} {part.Length}");
    }
}
```

`PackageArchive.OpenForRead` 会立即物化中央目录，把 .NET Framework 上延迟抛出的 `ArgumentException` 统一成 `InvalidDataException`，因此损坏包在六个目标框架上的表现一致（实测传入 4 个随机字节：`InvalidDataException: Central Directory corrupt.`）。上限可以按需收紧：

```csharp
OfficePackageLimits.MaxEntryBytes = 32L * 1024 * 1024;
OfficePackageLimits.MaxTotalBytes = 128L * 1024 * 1024;
```

默认值是 `MaxEntryBytes = 268,435,456`（256 MB）、`MaxTotalBytes = 1,073,741,824`（1 GB）；上面那份 `model.xlsx` 实测 9 个部件、合计 6,918 字节：

```text
[Content_Types].xml=1061  _rels/.rels=589  docProps/core.xml=630  docProps/app.xml=305
xl/workbook.xml=616  xl/_rels/workbook.xml.rels=567  xl/worksheets/sheet1.xml=1182
xl/styles.xml=1516  xl/sharedStrings.xml=452
```

其它可复用的工具：

```csharp
byte[] packed = Zlib.Compress(legacyBytes);                   // 10,752 -> 1,435 字节
byte[] restored = Zlib.Decompress(packed);                    // 自行实现的 inflate，无需第三方压缩库
// restored.Length == legacyBytes.Length，Adler32.Compute(restored, 0, restored.Length) = 3BE8A341

Console.WriteLine(UnitConverter.PointsToTwips(72));            // 1440   （Word 表宽/边距用 twips）
Console.WriteLine(UnitConverter.PointsToEmu(72));              // 914400 （Drawing 锚定用 EMU）
Console.WriteLine(UnitConverter.PointsToPixels(72));           // 96     （CSS px）

CellAddress.Parse("C7", out int row, out int column);           // 行 7、列 3；CellAddress.Format(7, 3) = "C7"
SpreadsheetRange.TryParse("台账!$B$2:$D$9",
    out int firstRow, out int lastRow, out int firstColumn, out int lastColumn);   // 2,9,2,4
```

`LittleEndianReader` / `BigEndianReader` 用于 BIFF 与 PDF 的二进制读取，`Adler32` 是 zlib 校验和，`DateUtils` 负责 Excel 序列日期与 `DateTime` 互转（含 1904 日期系统）。

## 常见问题

**段落不进 `Blocks` 就不显示。** `AddParagraph` 只在 `Section` / `HeaderFooter` 上有；对已经拿到手的 `Paragraph`，要显式 `section.Blocks.Add(p)`。

**表格列宽加起来超过页面宽度。** 排版按可用宽度等比压缩各列；跨页时表头重复由前导的 `IsHeader` 行决定，请保证表头行连续排在最前。

**`RowSpan` / `ColSpan` 设了但 Word 里没有合并。** 合并行必须真的少放被覆盖的单元格：`RowSpan = 2` 的第二行不应再为该列补单元格（本教程第 4 节的表头第二行只放"单件一 / 单件二"就是这个原因）。

**生成的 PDF 字体不对。** 模型里的 `TextFormat.FontFamily` 只声明族名，实际字形由 `Wisdom.Pdf` 的 `FontResolver` 在系统字体目录里解析；需要指定文件时参考 PDF 教程的字体注册步骤。

**`{{#each}}` 写了没反应。** 两个原因：`TemplateEngine.Replace` 只做标量替换、不展开循环（本教程第 10 节实测原样返回），循环要交给 `DocxTemplate` / `XlsxTemplate`；而这两个填充器要求 `{{#each items}}` 与 `{{/each}}` 各自独占一行（或独占一个只放标记的单元格），中间那一行才是被复制的数据行。

**形状没有名字。** 只有版式占位符会拿到模型自动分配的中文名（`标题`、`内容占位符`）；`AddTextBox` / `AddShape` / `AddImage` / `AddTable` 添加的形状 `Name` 是空串，最后一个字符串参数是**形状文本**。识别形状请用 `GetType()` 与 `Placeholder`，不要依赖 `Name`。

**幻灯片版式回读后变了 / 表格列数变了。** 版式在包里不单独存储，读回时按形状推断（`SectionHeader` → `Title`，带表格形状的 `TitleAndContent` → `TitleAndTable`）；`Table.Columns` 只是列占位，写出时不入库，回读时按每行单元格数补齐（2x0 → 2x2）。需要稳定的排版就别依赖版式，直接 `AddSlide()` 摆形状。

**导出 PDF / Word 后幻灯片版式全变了。** 那是"每页一节"的线性排版（`PresentationToDocumentConverter`），不还原绝对坐标、母版与动画；自选图形与图片按形状文本降级输出。需要还原视觉请在应用端直接交付 `.pptx`。
