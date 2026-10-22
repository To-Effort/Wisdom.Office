# Wisdom.Ppt 使用教程

演示文稿引擎的实战教程：从统一模型逐页搭一份季度回顾，替换版式生成的占位符，排一张表格页，画自选图形并嵌图片，回读校验版式与切换时长，再经 `Wisdom.Office` 门面转成 PDF / Word，把一份 Word 文档反向拆成幻灯片，最后挂入场动画、自动换页并用品牌主题存成 `.potx` 模板。

安装方式见[包说明](../src/Wisdom.Ppt/README.md#安装)。本教程示例是 .NET 8 控制台工程，文中所有输出与字节数都是这份工程实测打印的；第 6、7 两步与第 12 步的门面识别用到 `Wisdom.Office` 门面包（`OfficeFile` / `OfficeConverter`），其余只用 `Wisdom.Ppt` + `Wisdom.Office.Core`。

## 你将完成

```text
out/
  logo.png            教程自己写出的 1 x 1 PNG，70 字节
  review.pptx         五页回顾（标题页 / 要点页 / 双栏页 / 表格页 / 图形页），18027 字节
  review.pdf          review.pptx 渲染出的 PDF，87714 字节 / 5 页
  review.docx         review.pptx 转出的 Word，4761 字节 / 15 块
  facade.pdf          门面 Save 出的同一份 PDF，87714 字节
  roadmap.pptx        由 Word 文档反向拆出的 3 页，9544 字节
  plain.pptx          无一级标题的文档 → 1 页
  long.pptx           25 段正文 + 1 个一级标题 → 3 页
  review-stream.pptx  流写出，17991 字节
  motion.pptx         五个入场动画 + 3 秒自动换页，7716 字节
  brand.potx          品牌主题写出的 PowerPoint 模板，7010 字节
  brand-again.potx    模板回读后再写一次，6783 字节
```

工程入口先建目录，后面全部用相对路径：

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Ppt;

Directory.CreateDirectory("out");
```

## 1. 第一份演示文稿

`OfficePresentation.Widescreen()` 给出 16:9 的 960 x 540 磅画布。`AddSlide(SlideLayoutKind)` 按版式生成占位符形状，每个占位符里预置一个空段落，取出后直接往 `Paragraphs[0]` 里加文本。

```csharp
OfficePresentation presentation = OfficePresentation.Widescreen();
presentation.Metadata.Title = "季度经营回顾";
presentation.Metadata.Author = "Wisdom.Office";

Slide cover = presentation.AddSlide(SlideLayoutKind.Title);
TextBoxShape coverTitle = (TextBoxShape)cover.Shapes[0];       // Title 版式只有标题占位符
coverTitle.Paragraphs[0].AddText("2026 年第三季度经营回顾");
cover.NotesText = "开场 2 分钟，说明本次回顾的三个议题。";
```

要点页要按占位符**类别**找标题与正文，不要按下标猜：`TitleAndContent` 生成的是"标题 + 内容"两个占位符，但类别才是模型的契约。

```csharp
Slide points = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
TextBoxShape pointsTitle = null;
TextBoxShape body = null;
foreach (SlideShape shape in points.Shapes)
{
    TextBoxShape box = shape as TextBoxShape;
    if (box != null && box.Placeholder == PlaceholderKind.Title) pointsTitle = box;
    if (box != null && box.Placeholder == PlaceholderKind.Content) body = box;
}
pointsTitle.Paragraphs[0].AddText("本季度要点");
```

段落本身不带字符格式，格式挂在 `TextRun` 上；一段里可以放多个 `TextRun`，所以"同一句只有后半段加粗变色"就是加两个运行：

```csharp
body.Paragraphs[0].AddText("收入 1,280 万元");
body.Paragraphs[0].AddText("，环比增长 12%", new TextFormat
{
    FontSize = 20,
    Bold = true,
    Color = new OfficeColor(0x00, 0xB0, 0x50)
});
body.Paragraphs[0].ListMarker = "•";
body.Paragraphs[0].LeftIndent = 18;

string[] extra = { "成本控制在预算内，毛利率 41.6%", "华东区贡献了 46% 的增量收入" };
foreach (string item in extra)
{
    Paragraph paragraph = body.AddParagraph(string.Empty);
    paragraph.AddText(item);
    paragraph.ListMarker = "•";
    paragraph.LeftIndent = 18;
}
points.Transition = SlideTransitionKind.Fade;
points.TransitionDuration = 0.7;                                // 单位：秒

PptxWriter.Write(presentation, "out/review.pptx");
Console.WriteLine("[1] 页数=" + presentation.Slides.Count
    + "，文件=" + new FileInfo("out/review.pptx").Length + " 字节");
```

实测输出：`[1] 页数=2，文件=10667 字节`。此时只有两页——`Write` 可以反复调用，模型改到哪一页就写到哪一页，不必一次成型。

## 2. 双栏版式与自己的文本框

`TwoContent` 生成标题 + 两个内容占位符（左右各一栏），按类别取出来填即可；占位符以外的位置用 `AddTextBox` 自己摆。

```csharp
Slide compare = presentation.AddSlide(SlideLayoutKind.TwoContent);
TextBoxShape compareTitle = null;
List<TextBoxShape> columns = new List<TextBoxShape>();
foreach (SlideShape shape in compare.Shapes)
{
    TextBoxShape box = shape as TextBoxShape;
    if (box == null) continue;
    if (box.Placeholder == PlaceholderKind.Title) compareTitle = box;
    if (box.Placeholder == PlaceholderKind.Content) columns.Add(box);
}
compareTitle.Paragraphs[0].AddText("计划与实际");
columns[0].Paragraphs[0].AddText("计划：1,150 万元");
columns[1].Paragraphs[0].AddText("实际：1,280 万元");

TextBoxShape note = compare.AddTextBox(96, 460, 768, 40, "两栏占位符由版式自动生成，坐标单位是磅");
note.VerticalAlignment = SlideTextAnchor.Middle;
note.ShrinkToFit = true;
((TextRun)note.Paragraphs[0].Inlines[0]).Format.FontSize = 11;
Console.WriteLine("[2] 双栏占位符=" + columns.Count + " 个，形状总数=" + compare.Shapes.Count);
```

实测输出：`[2] 双栏占位符=2 个，形状总数=4`。`AddTextBox(left, top, width, height, text)` 的四个几何量都是**磅**，原点为该页左上角；写出时按 1 磅 = 12700 EMU 换算成 `a:off` / `a:ext`，所以 960 x 540 的画布上 `top + height` 别超过 540。

## 3. 表格页：占位符是空框，要换成表格

`TitleAndTable` 生成的第二个占位符是**文本框**（`Placeholder = PlaceholderKind.Table`），里面只有一个空段落——模型里没有"表格占位符"这种形状。直接 `AddTable` 会让这一页多出一个叠在表格上的空框，正确做法是借用它的几何位置，再把空框移出形状列表：

```csharp
Slide detail = presentation.AddSlide(SlideLayoutKind.TitleAndTable);
TextBoxShape detailTitle = null;
TextBoxShape tableHolder = null;
foreach (SlideShape shape in detail.Shapes)
{
    TextBoxShape box = shape as TextBoxShape;
    if (box == null) continue;
    if (box.Placeholder == PlaceholderKind.Title) detailTitle = box;
    if (box.Placeholder == PlaceholderKind.Table) tableHolder = box;
}
detailTitle.Paragraphs[0].AddText("区域明细");

double holderLeft = tableHolder.Left, holderTop = tableHolder.Top;
double holderWidth = tableHolder.Width, holderHeight = tableHolder.Height;
detail.Shapes.Remove(tableHolder);                              // Shapes 是可变的 List

TableShape tableShape = detail.AddTable(holderLeft, holderTop, holderWidth, holderHeight);
tableShape.HeaderColor = new OfficeColor(0x1F, 0x3A, 0x5F);
tableShape.Table.RepeatHeaderRow = true;
tableShape.Table.Columns.Add(new TableColumn(280));
tableShape.Table.Columns.Add(new TableColumn(280));
tableShape.Table.Columns.Add(new TableColumn(208));
tableShape.Table.AddRow("区域", "收入（万元）", "占比");
tableShape.Table.AddRow("华东", "588", "46%");
tableShape.Table.AddRow("华南", "384", "30%");
tableShape.Table.AddRow("华北", "308", "24%");
Console.WriteLine("[3] 占位符几何=" + (int)holderLeft + "," + (int)holderTop + " "
    + (int)holderWidth + "x" + (int)holderHeight
    + "，表格=" + tableShape.Table.Rows.Count + " 行 x "
    + tableShape.Table.Rows[0].Cells.Count + " 列，本页形状数=" + detail.Shapes.Count);
```

实测输出：`[3] 占位符几何=76,162 806x313，表格=4 行 x 3 列，本页形状数=2`。三点要注意：

- 占位符坐标由画布等比推出：960 x 540 上是 `left=76`（左右各留画布宽的 8%，即 76.8 取整）、`top=162`（高 x 0.30）、`806 x 313`（宽 = 画布宽 − 2 x 8%，高 = 高 x 0.58）。换 4:3 画布数值整体变小，别写死。
- `Table.Columns` 只用来登记**列宽定义**，`AddRow(params string[])` 不会往里加东西；数列数用 `Rows[0].Cells.Count`。
- `TableShape.Table` 就是 Word 侧的 `Table`，合并单元格、单元格内多段落等语义完全复用（`TableCell.Blocks`）。

## 4. 自选图形与图片

`Blank` 版式不生成任何占位符，适合完全自己排的页。

```csharp
Slide graphics = presentation.AddSlide(SlideLayoutKind.Blank);
AutoShape banner = graphics.AddShape(AutoShapeKind.RoundedRectangle, 96, 60, 768, 80, "达成率 96%");
banner.FillColor = new OfficeColor(0x00, 0xB0, 0x50);
banner.Outline = new BorderLine(BorderStyle.Single, 1.5, new OfficeColor(0x1F, 0x3A, 0x5F));
banner.CornerRadius = 0.25;                                     // 比例 0-0.5，仅圆角矩形生效
banner.Paragraphs[0].Alignment = ParagraphAlignment.Center;

AutoShape flag = graphics.AddShape(AutoShapeKind.Chevron, 96, 380, 300, 70, "下一阶段");
flag.FillColor = new OfficeColor(0x44, 0x72, 0xC4);

string pngPath = "out/logo.png";
File.WriteAllBytes(pngPath, Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg=="));
PictureShape picture = graphics.AddImage(OfficeImage.FromFile(pngPath), 700, 380, 120, 90);
picture.AlternativeText = "公司标识";
Console.WriteLine("[4] 形状=" + graphics.Shapes.Count + "，图片字节=" + picture.Image.Data.Length);

PptxWriter.Write(presentation, "out/review.pptx");
Console.WriteLine("[4] 写出=" + new FileInfo("out/review.pptx").Length + " 字节，页数=" + presentation.Slides.Count);
```

实测输出：`[4] 形状=3，图片字节=70` 与 `[4] 写出=18027 字节，页数=5`。

`BorderLine` 的构造顺序是 `(BorderStyle, 粗细磅, OfficeColor)`；`CornerRadius` 写出为 `roundRect` 的 `adj val="25000"`（0.25 x 100000）。`AddImage` 只接受 PNG / JPEG 字节（`OfficeImage.FromFile` 与 `new OfficeImage(bytes)` 两条路都行），`AlternativeText` 写进图片的 `p:cNvPr@descr`，读回时仍在 `PictureShape.AlternativeText` 上。

## 5. 回读校验

写出与读回共用同一套模型，"写了什么就该读回什么"可以逐页核对。

```csharp
OfficePresentation loaded = PptxReader.Load("out/review.pptx");   // 也支持 Stream / byte[]
Console.WriteLine("[5] 回读页数=" + loaded.Slides.Count + "，尺寸="
    + (int)loaded.SlideWidth + "x" + (int)loaded.SlideHeight);
for (int i = 0; i < loaded.Slides.Count; i++)
{
    Slide slide = loaded.Slides[i];
    Console.WriteLine("    第 " + (i + 1) + " 页 版式=" + slide.Layout + " 形状=" + slide.Shapes.Count
        + " 切换=" + slide.Transition + "/" + slide.TransitionDuration
        + " 标题=\"" + slide.GetTitle() + "\"");
}
```

实测输出：

```text
[5] 回读页数=5，尺寸=960x540
    第 1 页 版式=Title 形状=1 切换=None/0 标题="2026 年第三季度经营回顾"
    第 2 页 版式=TitleAndContent 形状=2 切换=Fade/0.7 标题="本季度要点"
    第 3 页 版式=TwoContent 形状=4 切换=None/0 标题="计划与实际"
    第 4 页 版式=TitleAndTable 形状=2 切换=None/0 标题="区域明细"
    第 5 页 版式=Blank 形状=3 切换=None/0 标题=""
```

`Slide.Layout` 是**推断值**：pptx 里存的是版式关系与占位符，不是统一模型的 `SlideLayoutKind`，读取器按"有无 title 占位符 / 有无图片或表格占位符 / 内容占位符几个"回推。第 4 页的表格是不带占位符的 `p:graphicFrame`，仍能判回 `TitleAndTable`，因为形状类型（`TableShape`、`PictureShape`）也算证据。

再核三件容易漏的东西：表格内容、图形填充、备注与切换时长。

```csharp
TableShape loadedTable = null;
AutoShape loadedBanner = null;
foreach (SlideShape shape in loaded.Slides[3].Shapes)
{
    TableShape t = shape as TableShape;
    if (t != null) loadedTable = t;
}
foreach (SlideShape shape in loaded.Slides[4].Shapes)
{
    AutoShape s = shape as AutoShape;
    if (s != null && s.Kind == AutoShapeKind.RoundedRectangle) loadedBanner = s;
}
Console.WriteLine("[5] 表格回读=" + loadedTable.Table.Rows.Count + " 行 x "
    + loadedTable.Table.Rows[0].Cells.Count + " 列，单元格(1,0)=\""
    + CellText(loadedTable.Table, 1, 0) + "\"，图形填充=" + loadedBanner.FillColor.Value.ToHex());

OfficePresentation fromBytes = PptxReader.Load(File.ReadAllBytes("out/review.pptx"));
Console.WriteLine("[5] 字节流回读页数=" + fromBytes.Slides.Count
    + "，备注=\"" + fromBytes.Slides[0].NotesText + "\"");
Console.WriteLine("[5] 纯文本提取长度=" + PptxReader.ExtractText("out/review.pptx").Length);
```

实测输出：

```text
[5] 表格回读=4 行 x 3 列，单元格(1,0)="华东"，图形填充=00B050
[5] 字节流回读页数=5，备注="开场 2 分钟，说明本次回顾的三个议题。"
[5] 纯文本提取长度=243
```

`SlideShape.FillColor` 是 `OfficeColor?`——null 表示"没设填充，继承版式与主题"，取色前要判 null，或像上面那样确认形状种类后再用 `.Value`。

```csharp
static string CellText(Table table, int row, int column)
{
    Paragraph paragraph = table.Rows[row].Cells[column].Blocks[0] as Paragraph;
    return paragraph == null ? string.Empty : paragraph.GetText();
}
```

## 6. 经 `Wisdom.Office` 门面识别与取文本

不想在业务代码里区分 Word / Excel / PPT 时用统一入口：`OfficeFile.Load` 按文件头嗅探，`.pptx` 识别为 `Pptx`。

```csharp
OfficeFile file = OfficeFile.Load("out/review.pptx");
Console.WriteLine("[6] 识别=" + OfficeFileFormats.GetDisplayName(file.Format)
    + "，元数据标题=\"" + file.Metadata.Title + "\"，纯文本长度=" + file.ExtractText().Length);
Console.WriteLine("    " + file.ExtractText().Replace("\n", " / "));

OfficePresentation viaFacade = file.AsPresentation();
Console.WriteLine("[6] AsPresentation 页数=" + viaFacade.Slides.Count
    + "，AsDocument 节数=" + file.AsDocument().Sections.Count);

file.Save("out/facade.pdf");
Console.WriteLine("[6] 门面 Save 出 PDF=" + new FileInfo("out/facade.pdf").Length + " 字节");
```

实测输出（门面文本按行折行展示）：

```text
[6] 识别=PowerPoint 演示文稿，元数据标题="季度经营回顾"，纯文本长度=256
    第 1 页：2026 年第三季度经营回顾 /  / 第 2 页：本季度要点 / 收入 1,280 万元，环比增长 12%
    / 成本控制在预算内，毛利率 41.6% / 华东区贡献了 46% 的增量收入 /  / 第 3 页：计划与实际
    / 计划：1,150 万元 / 实际：1,280 万元 / 两栏占位符由版式自动生成，坐标单位是磅 /  / 第 4 页：区域明细
    / 区域	收入（万元）	占比 / 华东	588	46% / 华南	384	30% / 华北	308	24% /  / 第 5 页
    / 达成率 96% / 下一阶段 / [图片] 公司标识 /
[6] AsPresentation 页数=5，AsDocument 节数=5
[6] 门面 Save 出 PDF=87714 字节
```

两条提取路径格式不同，别按某一条的字符数写断言：`PptxReader.ExtractText` 实测 243 字符，每页只输出标题行 + 各形状文本，表格行内用制表符分隔，备注以 `[备注]` 前缀跟在对应页后，图片不输出；门面 `ExtractText` 实测 256 字符，每页前面加 `第 N 页：`，图片输出成 `[图片] 替代文字`，但不带备注。`AsDocument()` 把每页降格成一个 Word 节（5 页 → 5 节），标题降为普通段落，适合喂给只做文本处理的下游。

## 7. 转 PDF 与 Word

转换入口在 `OfficeConverter`，需要 `Wisdom.Office` 门面包（`Wisdom.Ppt` 本身不依赖 `Wisdom.Pdf` / `Wisdom.Word`）。

```csharp
OfficeConverter.PresentationToPdf(loaded, "out/review.pdf");
OfficeConverter.PresentationToWord(loaded, "out/review.docx");

PdfDocument pdf = PdfDocument.Load("out/review.pdf");
OfficeDocument word = OfficeFile.Load("out/review.docx").AsDocument();
Console.WriteLine("[7] PDF=" + new FileInfo("out/review.pdf").Length + " 字节/" + pdf.PageCount + " 页，DOCX="
    + new FileInfo("out/review.docx").Length + " 字节/" + CountBlocks(word) + " 块");
```

实测输出：`[7] PDF=87714 字节/5 页，DOCX=4761 字节/15 块`。PDF 一页对应一张幻灯片，按画布比例排版，形状按 z 序（即 `Shapes` 顺序）叠放；Word 侧每页一节，15 块是这五页的段落与表格总数（`Section.Blocks` 里段落、表格都算一块）。

```csharp
static int CountBlocks(OfficeDocument document)
{
    int total = 0;
    for (int s = 0; s < document.Sections.Count; s++)
    {
        total += document.Sections[s].Blocks.Count;
    }
    return total;
}
```

## 8. 反向：Word 文档拆成幻灯片

`DocumentToPptx` 按一级、二级标题切页，标题进标题占位符，其余内容进正文。

```csharp
OfficeDocument source = new OfficeDocument();
source.Metadata.Title = "产品路线图";
Section section = source.AddSection(PageSetup.A4());
string[] chapters = { "上半年：报表引擎", "下半年：格式兼容", "明年：协作与批注" };
foreach (string chapter in chapters)
{
    Paragraph heading = source.Sections[0].AddParagraph(chapter);
    heading.HeadingLevel = 1;
    source.Sections[0].AddParagraph("要点与验收标准见详细方案。");
}

OfficeConverter.DocumentToPptx(source, "out/roadmap.pptx");
OfficePresentation roadmap = PptxReader.Load("out/roadmap.pptx");
Console.WriteLine("[8] 文档块=" + CountBlocks(source) + "，拆出页数=" + roadmap.Slides.Count
    + "，首页标题=\"" + roadmap.Slides[0].GetTitle() + "\"，文件="
    + new FileInfo("out/roadmap.pptx").Length + " 字节");
```

实测输出：`[8] 文档块=6，拆出页数=3，首页标题="上半年：报表引擎"，文件=9544 字节`。六个文档块（3 标题 + 3 正文）压成三页，说明拆分依据是**标题层级**而不是块数。

## 9. 流、字节与 4:3 画布

服务端不落盘时，三种写出出口等价：

```csharp
using (Stream stream = File.Create("out/review-stream.pptx"))
{
    PptxWriter.Write(loaded, stream);
}
byte[] deckBytes = PptxWriter.Write(loaded);

OfficePresentation standard = OfficePresentation.Standard();      // 4:3，720 x 540
standard.AddSlide(SlideLayoutKind.Title);
Console.WriteLine("[9] 流写出=" + new FileInfo("out/review-stream.pptx").Length
    + " 字节，字节写出=" + deckBytes.Length + " 字节，4:3 画布="
    + (int)standard.SlideWidth + "x" + (int)standard.SlideHeight);
```

实测输出：`[9] 流写出=17991 字节，字节写出=17991 字节，4:3 画布=720x540`。

同一份模型两次写出（流与字节）字节数相同，说明包内容不含时间戳之类的随机量。但它比第 4 步的 18027 字节少 36 字节：`loaded` 是读回来的模型，读取器没解析的属性不会重新写出，所以**别拿字节数当等价性断言**，要断言就按第 5 步逐页比内容。

## 10. 拆分规则的两条边界

正文段数有上限（每页 12 段），且没有一级/二级标题的文档不会凭空造封面页。

```csharp
OfficeDocument plain = new OfficeDocument();
plain.Metadata.Title = "无标题文档";
plain.AddSection(PageSetup.A4());
plain.Sections[0].AddParagraph("第一段正文。");
plain.Sections[0].AddParagraph("第二段正文。");
plain.Sections[0].AddParagraph("第三段正文。");
OfficeConverter.DocumentToPptx(plain, "out/plain.pptx");
OfficePresentation plainDeck = PptxReader.Load("out/plain.pptx");
Console.WriteLine("[10] 无一级标题：文档块=" + CountBlocks(plain) + "，拆出页数=" + plainDeck.Slides.Count);
```

实测输出：`[10] 无一级标题：文档块=3，拆出页数=1`，且这一页是 `TitleAndContent` 版式、标题占位符里是一个空格 `" "`（模型需要一个标题文本，转换器对无标题文档用空格占位）。

段落超上限时按同一标题续页：

```csharp
OfficeDocument longDoc = new OfficeDocument();
longDoc.Metadata.Title = "长文档";
longDoc.AddSection(PageSetup.A4());
Paragraph firstHeading = longDoc.Sections[0].AddParagraph("唯一的一级标题");
firstHeading.HeadingLevel = 1;
for (int i = 1; i <= 25; i++)
{
    longDoc.Sections[0].AddParagraph("正文第 " + i + " 段。");
}
OfficeConverter.DocumentToPptx(longDoc, "out/long.pptx");
OfficePresentation longDeck = PptxReader.Load("out/long.pptx");
Console.WriteLine("[10] 单页装不下时：文档块=" + CountBlocks(longDoc) + "，拆出页数=" + longDeck.Slides.Count);
```

实测输出：`[10] 单页装不下时：文档块=26，拆出页数=3`，三页的标题依次是 `唯一的一级标题`、`唯一的一级标题（续）`、`唯一的一级标题（续）`，正文段数依次 12 / 12 / 1（第 1–12 段、第 13–24 段、第 25 段）。续页标题只带一个"（续）"，不会叠成"（续）（续）"。

## 11. 入场动画与自动换页

`SlideShape` 上的三个属性 `Animation` / `AnimationDurationSeconds` / `AnimationDelaySeconds` 描述一个入场动画，写出器据此生成 PowerPoint 原生的 `p:timing` 主序列；`Slide.AdvanceAfterSeconds` 是 `p:transition/@advTm`，到点自动换页。

```csharp
OfficePresentation deck = OfficePresentation.Widescreen();
Slide motion = deck.AddSlide(SlideLayoutKind.Blank);
string[] labels = { "飞入标题", "擦除要点", "缩放要点", "淡入要点", "拆分要点" };
SlideAnimationKind[] kinds =
{
    SlideAnimationKind.FlyIn, SlideAnimationKind.Wipe, SlideAnimationKind.Zoom,
    SlideAnimationKind.Fade, SlideAnimationKind.Split
};
for (int i = 0; i < kinds.Length; i++)
{
    TextBoxShape box = motion.AddTextBox(96, 60 + i * 80, 768, 60, labels[i]);
    box.Animation = kinds[i];
    if (i == 0) { box.AnimationDurationSeconds = 0.8; box.AnimationDelaySeconds = 0.5; }
}
TextBoxShape still = motion.AddTextBox(96, 470, 768, 40, "这一行没有动画");
motion.AdvanceAfterSeconds = 3.0;
motion.Transition = SlideTransitionKind.Push;
motion.TransitionDuration = 0.5;
PptxWriter.Write(deck, "out/motion.pptx");

OfficePresentation motionBack = PptxReader.Load("out/motion.pptx");
Slide motionAgain = motionBack.Slides[0];
Console.WriteLine("[11] 写出=" + new FileInfo("out/motion.pptx").Length + " 字节，形状数="
    + motionAgain.Shapes.Count + "，自动换页=" + motionAgain.AdvanceAfterSeconds + "s");
for (int i = 0; i < motionAgain.Shapes.Count; i++)
{
    SlideShape shape = motionAgain.Shapes[i];
    Console.WriteLine("     " + shape.Name + " 动画=" + shape.Animation
        + " 时长=" + shape.AnimationDurationSeconds + "s 延迟=" + shape.AnimationDelaySeconds + "s");
}
```

实测输出：

```text
[11] 写出=7716 字节，形状数=6，自动换页=3s
     Shape 动画=FlyIn 时长=0.8s 延迟=0.5s
     Shape 动画=Wipe 时长=0.5s 延迟=0s
     Shape 动画=Zoom 时长=1s 延迟=0s
     Shape 动画=Fade 时长=0.5s 延迟=0s
     Shape 动画=Split 时长=0.5s 延迟=0s
     Shape 动画=None 时长=0s 延迟=0s
```

四条规则：

- **`AnimationDurationSeconds = 0` 是"取类别默认"，不是"没有动画"。** 默认时长 FlyIn / Zoom 各 1 秒，Fade / Wipe / Split 各 0.5 秒，实测第 3 个形状没设时长也回读出 `1s`。开关只有 `Animation = SlideAnimationKind.None`。
- **延迟只对"自动播放的那一个"有意义。** 写出器把第一个动画的形状写成 `afterEvent`（配合 `AdvanceAfterSeconds`），其余形状仍是单击触发，所以第 2–5 个形状即便设了延迟也会写成 `indef`，回读成 0（实测两个形状都设 `1.2s`，回读 `1.2 / 0`）。
- **动画目标是 `cNvPr id`，形状名字回读是 `Shape`。** `p:timing` 用 `p:spid/@val` 指形状，模型侧不保存这个 id 与 `Name` 的关系，按 `Shapes` 顺序对位即可。
- **只有入场动画。** 强调、退出、路径与媒体动画写出时不生成，读取时按 `presetClass="entr"` 过滤，第三方 PPTX 里的其他动画节点直接跳过。

## 12. 主题深改与 `.potx` 模板

`OfficePresentation.Theme`（`PresentationTheme`）覆盖 `theme1.xml` 的六个色板阶段与主/次字体，null 字段保持内置默认；强调色仍走 `AccentColors`。母版与版式的文字、背景色通过 `p:clrMap`（`bg1→lt1`、`tx1→dk1`）自动跟随，不需要逐个形状改填充。

```csharp
OfficePresentation branded = OfficePresentation.Widescreen();
branded.Theme = new PresentationTheme
{
    Name = "品牌主题",
    Text = new OfficeColor(0x1A, 0x1A, 0x2E),          // dk1
    Background = new OfficeColor(0xFA, 0xFA, 0xF5),    // lt1
    TextSecondary = new OfficeColor(0x2F, 0x4D, 0x6B), // dk2
    Hyperlink = new OfficeColor(0xC0, 0x39, 0x2B),
    FontMajor = "微软雅黑",                             // majorFont/latin
    FontMinor = "Consolas"                             // minorFont/latin
};
Slide brandSlide = branded.AddSlide(SlideLayoutKind.TitleAndContent);
brandSlide.AddTextBox(96, 460, 768, 40, "主题色经 clrMap 影响母版与版式占位符");
PptxWriter.WriteTemplate(branded, "out/brand.potx");   // 文件 / 流 / 字节三个重载都有

OfficePresentation brandedBack = PptxReader.Load("out/brand.potx");
PresentationTheme themeBack = brandedBack.Theme;
Console.WriteLine("[12] 主题名=\"" + themeBack.Name + "\"，dk1=#" + themeBack.Text.Value.ToHex()
    + "，lt1=#" + themeBack.Background.Value.ToHex() + "，dk2=#" + themeBack.TextSecondary.Value.ToHex()
    + "，超链接=#" + themeBack.Hyperlink.Value.ToHex()
    + "，主字体=" + themeBack.FontMajor + "，次字体=" + themeBack.FontMinor);

string againPath = "out/brand-again.potx";
PptxWriter.WriteTemplate(brandedBack, againPath);
Console.WriteLine("[12] 首次写出=" + new FileInfo("out/brand.potx").Length
    + " 字节，回读再写=" + new FileInfo(againPath).Length + " 字节");

OfficeFile templateFile = OfficeFile.Load("out/brand.potx");
Console.WriteLine("[12] 门面识别=" + OfficeFileFormats.GetDisplayName(templateFile.Format)
    + "，AsPresentation 页数=" + templateFile.AsPresentation().Slides.Count);
```

实测输出：

```text
[12] 主题名="品牌主题"，dk1=#1A1A2E，lt1=#FAFAF5，dk2=#2F4D6B，超链接=#C0392B，主字体=微软雅黑，次字体=Consolas
[12] 首次写出=7010 字节，回读再写=6783 字节
[12] 门面识别=PowerPoint 模板，AsPresentation 页数=1
```

三点要注意：

- **`.potx` 与 `.pptx` 只差一个内容类型。** 包结构、关系类型（仍是 `officeDocument`）完全相同，`[Content_Types].xml` 里 `/ppt/presentation.xml` 的 `Override` 换成 `…presentationml.template.main+xml`；`OfficeFileFormats.Detect` 就是读这一处来区分，所以 `WriteTemplate` 出来的文件能被识别成 `Potx`，按扩展名 `Save("deck.pptx")` 又变回普通演示文稿。
- **读回来的 `Theme` 每个字段都非 null。** 读取器把 `clrScheme` / `fontScheme` 里读到的值原样写进模型，包括与内置默认相同的那些，因此 `Theme.Text` 可以直接 `.Value`；没设过主题的普通 PPTX 也读得出默认值（实测 `名=Office，dk1=#000000，lt1=#FFFFFF，主字体=Calibri，次字体=Calibri`）。
- **回读再写会掉 227 字节**（7010 → 6783），原因与第 9 步相同：读取器不解析的属性不会重新写出，模板的字节数不能当等价性断言。

## 常见问题

**`Table.Columns.Count` 是 0。** 正常。`Columns` 只在需要显式指定列宽时填（`new TableColumn(磅)`），`AddRow(params string[])` 只加单元格。不填列宽时写出器按总宽与列数均分。

**表格页上多出一个空框。** `TitleAndTable` / `TitleAndPicture` 生成的第二个占位符是文本框，不是表格或图片形状。用它的 `Left/Top/Width/Height` 定位 `AddTable` / `AddImage`，再 `slide.Shapes.Remove(占位符)`，见第 3 步。

**读回来的 `Layout` 跟写进去的不一样。** 版式是推断值。带占位符的形状按 `ph type` 判，`obj`（模型里的 `Content` / `Table` 占位符）都归到内容类；表格与图片即使没有占位符，也按 `TableShape` / `PictureShape` 计入证据。

**切换动画没生效或时长丢了。** `TransitionDuration` 单位是秒，写出为 `p14:dur`（毫秒），包在 `mc:AlternateContent` 的 `mc:Choice` 里，`mc:Fallback` 只留 `spd="med"`——不支持 p14 扩展的老版本仍有切换，只是按 `spd` 走。`SlideTransitionKind.None` 不写效果节点，但 `AdvanceAfterSeconds > 0` 时仍会写一个只带 `advTm` 的 `p:transition`，自动换页不依赖切换效果。

**形状动画没出现。** 检查三件事：`Animation` 是不是 `None`、这页或前一页是否需要先手动触发（动画只在放映模式播放，编辑视图看不到）、以及是不是只有第一个形状带上了延迟（其余形状的延迟写出为 `indef`，见第 11 步）。动画节点写在 `p:timing` 里，与 `p:transition` 是两个部件。

**形状颜色读回来是 null。** `FillColor` 是 `OfficeColor?`，null 表示继承版式/主题色，不是丢失。

**图片读回来字节不对 / 打不开。** 只支持 PNG 与 JPEG；`.emf`、`.svg`、`.gif` 不解析。教程里那颗 1 x 1 的 PNG（70 字节）就是最省事的自造图片。

**要不要装 Office。** 不需要。`.pptx` 与 `.potx` 由 `ZipArchive` + LINQ to XML 直接读写；旧版二进制 `.ppt`、图表、SmartArt、OLE 对象都不在支持范围，母版只暴露主题色板与主/次字体这一层深改（第 12 步），动画只支持入场类别（第 11 步），详见包说明的「已知限制」。
