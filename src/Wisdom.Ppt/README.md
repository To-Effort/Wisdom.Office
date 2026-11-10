# Wisdom.Ppt

`Wisdom.Ppt` 负责 **PowerPoint 侧的全部读写**：用统一演示文稿模型 `OfficePresentation` 构建幻灯片并写出 PPTX，或把 `.pptx` 读回同一模型。全部由 `System.IO.Compression.ZipArchive` + LINQ to XML 直接处理 PresentationML，不依赖 Office、WPS 或任何第三方库。

- 包名：`Wisdom.Ppt`
- 版本：`1.0.3`
- 依赖：`Wisdom.Office.Core`（唯一依赖）
- 第三方依赖：无；不要求安装 Microsoft Office 或 WPS
- 命名空间：`Wisdom.Ppt`（读写器）、`Wisdom.Office.Core`（`OfficePresentation`、`Slide`、各形状类型）

本文是能力与 API 参考；分步实战教程见源码仓库的 `docs/Wisdom.Ppt.Tutorial.md`。

## 支持的框架

| 目标框架 | 包内资产 | 典型运行时 |
| --- | --- | --- |
| `net461` | `lib/net461` | .NET Framework 4.6.1 / 4.7.2 / 4.8.x |
| `netstandard2.0` | `lib/netstandard2.0` | .NET Core 2.0-3.1、.NET 5/6、Mono、Unity |
| `net7.0` | `lib/net7.0` | .NET 7 |
| `net8.0` | `lib/net8.0` | .NET 8 |
| `net9.0` | `lib/net9.0` | .NET 9 |
| `net10.0` | `lib/net10.0` | .NET 10 / 11 |

六份资产由同一份源码编译，功能一致。

## 安装

```bash
dotnet add package Wisdom.Ppt --version 1.0.3
```

需要把演示文稿转成 PDF / Word 时请改引 `Wisdom.Office` 门面包（本包不依赖 `Wisdom.Pdf` 与 `Wisdom.Word`）。

## 单位与坐标系

所有几何量以 **磅（point）** 为单位，写出时按 1 磅 = 12700 EMU 换算成 PresentationML 的 `a:off` / `a:ext`。坐标原点为幻灯片左上角。

| 尺寸 | 属性 | 用途 |
| --- | --- | --- |
| 16:9 宽屏 | `OfficePresentation.Widescreen()`，960 x 540 | 默认，现代投影仪 |
| 4:3 标准 | `OfficePresentation.Standard()`，720 x 540 | 老式投影与打印 |

## 构建演示文稿

`AddSlide(SlideLayoutKind)` 会按版式生成占位符形状，每个占位符内预置一个空段落，取出后直接 `Paragraphs[0].AddText(...)` 填内容。

```csharp
using Wisdom.Ppt;
using Wisdom.Office.Core;

OfficePresentation presentation = OfficePresentation.Widescreen();
presentation.Metadata.Title = "季度经营回顾";

Slide cover = presentation.AddSlide(SlideLayoutKind.Title);
TextBoxShape coverTitle = (TextBoxShape)cover.Shapes[0];   // Title 版式只有标题占位符
coverTitle.Paragraphs[0].AddText("2026 年第三季度经营回顾");
cover.NotesText = "开场 2 分钟，说明本次回顾的三个议题。";

Slide bullets = presentation.AddSlide(SlideLayoutKind.TitleAndContent);
TextBoxShape title = (TextBoxShape)bullets.Shapes[0];
title.Paragraphs[0].AddText("本季度要点");

TextBoxShape body = FindContent(bullets);                  // 见下方辅助方法
body.Paragraphs[0].AddText("收入 1,280 万元，环比增长 12%");
body.Paragraphs[0].ListMarker = "•";
body.Paragraphs[0].LeftIndent = 18;
foreach (string item in new[] { "成本控制在预算内，毛利率 41.6%", "华东区贡献了 46% 的增量收入" })
{
    Paragraph paragraph = body.AddParagraph(string.Empty);
    paragraph.AddText(item);
    paragraph.ListMarker = "•";
    paragraph.LeftIndent = 18;
}
bullets.Transition = SlideTransitionKind.Fade;
bullets.TransitionDuration = 0.7;

PptxWriter.Write(presentation, "review.pptx");             // 亦支持 Stream，或 byte[] Write(presentation)
```

```csharp
static TextBoxShape FindContent(Slide slide)
{
    foreach (SlideShape shape in slide.Shapes)
    {
        TextBoxShape box = shape as TextBoxShape;
        if (box != null && (box.Placeholder == PlaceholderKind.Content
            || box.Placeholder == PlaceholderKind.Body))
        {
            return box;
        }
    }
    return slide.AddTextBox(96, 170, 768, 260, string.Empty);
}
```

生成的包通过内容类型、关系目标与 XML 合法性校验，可被 PowerPoint 2010+、WPS 演示直接打开。

## 版式与占位符

`AddSlide(SlideLayoutKind)` 按当前幻灯片尺寸等比推算占位符位置：

| `SlideLayoutKind` | 生成的占位符 |
| --- | --- |
| `Title` | `Title` |
| `TitleAndContent` | `Title` + `Content` |
| `SectionHeader` | `Title`（垂直居中、位置偏低） |
| `TwoContent` | `Title` + 左右两个 `Content` |
| `TitleAndPicture` | `Title` + `Picture` |
| `TitleAndTable` | `Title` + `Table` |
| `Blank` | 无 |

`AddSlide()` 不带参数等价于 `Blank`，全部形状自己定位。占位符的字体不在模型里写死，由写出器套用 `DefaultTitleFormat` / `DefaultTextFormat`。

空占位符的省略只认来源：`AddSlide(SlideLayoutKind)` 自己补出来的那一份，如果到写出时仍然没有任何可见效果（文本为空、没有底色、没有描边、没有入场动画），保存时整体不写出，免得 PowerPoint 在页面上留一个"单击此处添加"的空框。一旦你给它刷了底色或描边，它就照常写出。从文件读回的占位符、以及 `AddTextBox` 一类自己摆的形状都不受这条规则影响——哪怕里面没字，也是页面上的一个形状，必须保留。来源标记是 `SlideShape.IsGeneratedPlaceholder`（只能由本库置位）。

## 文本、图形、图片与表格

| 形状 | 构造方式 | 说明 |
| --- | --- | --- |
| `TextBoxShape` | `slide.AddTextBox(left, top, width, height, text)` | `Paragraphs` 列表、`VerticalAlignment`、四边内边距、`ShrinkToFit` |
| `AutoShape` | `slide.AddShape(AutoShapeKind, ...)` | `FillColor`、`Outline`、`CornerRadius` |
| `PictureShape` | `slide.AddImage(new OfficeImage(bytes), ...)` | 仅 PNG / JPEG；`AlternativeText` 作为替代文字；`FillColor`、`Outline` 与 `AutoShape` 同一口径（设过才写、读得回） |
| `TableShape` | `slide.AddTable(left, top, width, height)` | 内部就是一个 `Table`，复用 Word 侧的行列与合并单元格语义 |

```csharp
Slide shapes = presentation.AddSlide(SlideLayoutKind.Blank);

AutoShape banner = shapes.AddShape(AutoShapeKind.RoundedRectangle, 96, 60, 768, 80, "达成率 96%");
banner.FillColor = new OfficeColor(0x00, 0xB0, 0x50);
banner.Outline = new BorderLine(BorderStyle.Single, 1, new OfficeColor(0x1F, 0x3A, 0x5F));
banner.CornerRadius = 0.25;
banner.Paragraphs[0].Alignment = ParagraphAlignment.Center;

shapes.AddImage(OfficeImage.FromFile("chart.png"), 96, 180, 320, 200);
shapes.AddTextBox(440, 205, 424, 120, "绿色圆角矩形为达标线以上的达成率");

TableShape table = shapes.AddTable(96, 330, 768, 150);
table.HeaderColor = new OfficeColor(0x1F, 0x3A, 0x5F);
table.Table.AddRow("区域", "收入（万元）", "占比");
table.Table.AddRow("华东", "588", "46%");
table.Table.AddRow("华南", "384", "30%");
```

`AutoShapeKind` 支持 `Rectangle`、`RoundedRectangle`、`Ellipse`、`Triangle`、`Diamond`、`Pentagon`、`Hexagon`、`RightArrow`、`Chevron`、`Line`，逐个映射到 PresentationML 的 `prstGeom`。
宽或高为 `0` 的形状是合法几何，照原样写出也照原样读回：水平分隔线就是一条 `Height = 0` 的 `Line`，读回时不会因为「`cy` 是 0」被当成几何缺失而换成按幻灯片比例推算的默认大框（只有负值才判为坏几何）。
`SlideTransitionKind` 支持 `None`、`Fade`、`Push`、`Wipe`、`Split`，配合 `TransitionDuration`（秒）。
方向、速度与「是否需要单击」也各有字段：`TransitionDirection`（`SlideTransitionDirection` 的 `Left`/`Right`/`Up`/`Down`，对应 `dir` 的 `l`/`r`/`u`/`d`）、`TransitionSpeed`（`SlideTransitionSpeed` 的 `Slow`/`Medium`/`Fast`，对应 `spd` 的 `slow`/`med`/`fast`）与 `AdvanceOnClick`（默认 `true`，写 `advClick="1"`；置 `false` 表示自动换片、不接受点击）。三者都是可空/带默认的：没设过就不写，写出形状与旧稿一致，读回时也保持「未设置」而不是猜一个值——推入的原生默认方向仍然是向左；配合 `AdvanceAfterSeconds` 设成自动换片时，`advClick="0"` 现在能原值往返，不会再被还原成必须点击。

文本运行只写你**显式设过**的字符属性。`new TextFormat()` 预置的 Arial、11 磅、黑色是模型内置默认值，不算意图：一个只设了 `Bold = true` 的 run 写出的是 `<a:rPr b="1"/>`，字号、字体、颜色都留给版式与母版的继承链，不会被钉成 11 磅的 Calibri。`TextFormat` 上的 `HasExplicitFontFamily`、`HasExplicitFontSize`、`HasExplicitColor` 三个只读标记就是这条判定，赋过值才点亮；`new TextFormat("Calibri", 18)` 这类显式构造则三项都算设过。整段没有任何格式的文本沿用演示文稿的默认字体（`DefaultTextFormat` / `DefaultTitleFormat`，未显式设置时取主题主/次字体）。

段落属性走的是同一条规矩：`a:pPr` 里只装模型说过的话，一段什么都没设时整个 `a:pPr` 省略不写。项目符号最要紧——`a:buNone` 不是「没有符号」的留空，而是一条**关掉本层符号**的显式指令，从前无条件跟在每段后面，等于把加载进来的第三方母版与版式里定义好的符号逐段压制掉。现在只有 `ListMarker` 非空才写 `buChar`（数字开头写 `buAutoNum`），没标记就不写符号子元素，让继承链自己说话；本库自建的母版与版式从不定义符号，自家产物因此字节不变。代价是模型里「显式关掉符号」与「没想过这件事」仍是同一个值（读取端把 `buNone` 读回 `null`），载入带 `buNone` 的外来稿再存盘时那份显式关闭不会被保留。

文本框的折行走的是同一句规矩：`a:bodyPr@wrap` 里只有 `none` 是一句真话——PowerPoint 用它写单行标题与竖排标签，文字超出右边界照样画出去，不再在框内折行。`TextBoxShape.WrapText` 就是这个开关：置 `false` 写 `wrap="none"`；没设置或置 `true` 都写 `wrap="square"`，那正是 PowerPoint 自己不写这个属性时的默认，所以自家产物一个字节都不变。读端把 `none` 读回 `false`，`square`、`emergency` 与属性缺失一律读回 `null`——文件里没有任何一句「要折行」值得当成声明，把默认值读成一条声明正是 `TextFormat` 那档教训的复发。从前写端无论模型说没说都钉一个 `wrap="square"`，读端从不看这个属性，载入别人的演示稿存一次盘，那行不许折行的标题就开始在框里折行：文本、形状数一个字都没少，只有版式静默变了。

表格的结构开关同样只写模型里说过的话。`a:tblPr@firstRow` 只在**声明过**表头时写——判据是某一行 `TableRow.IsHeader`，或给过 `TableShape.HeaderColor`；两者都没有时首行按普通数据行走，不会凭空加粗、变白。读回时 `firstRow` 被还原成首行的 `IsHeader`，所以声明过的表头能活过第二次保存（模型里没有第二个字段承载这件事）。隔行条纹（`bandRow`）在统一模型里没有对应字段，写出时不再恒带；单元格没设底色时也不补 `a:noFill`——那等于替调用者宣布「这格不许有底色」，会把 `tableStyleId` 本来要给的外观一并抹掉。列宽之外的整表样式（`tableStyleId`）固定为内置的一套，不开放选择。

线型分两处编码。`BorderLine.Style` 取 `None`、`Single`、`Dashed`、`Dotted`、`Double`：虚线与点线是**划线图案**，写进 `a:ln` 下的 `a:prstDash`（`dash` 与 `sysDot`）；双线不是图案，是一种**复合线**，写在 `a:ln@cmpd="dbl"` 上——`ST_PresetLineDashVal` 的取值集合里根本没有 `dbl` 这一档。从前写出器把 Double 塞进 `prstDash`，文件里落下一个架构不接受的枚举值，读端又认不出来，双线边框存一次盘就静默降成单线；点线当时映射成 `sysDash`（长划图案），画出来是虚线而不是点。现在形状描边与表格边线两条路径都按 `cmpd` 写双线，读端优先看 `cmpd="dbl"`，同时继续认旧文件里 `prstDash val="dbl"` 这份历史写法，早前版本存出的双线不会读成实线。单线与虚线的字节不变（`cmpd` 缺省即 `sng`，形状描边非双线时不写这个属性）。

表格边框认六个方向。`Table.Borders` 的 `Top`/`Bottom`/`Left`/`Right` 管外圈那四条边，`InsideHorizontal`/`InsideVertical` 管内部的两条分隔线；写出器按单元格落在网格的哪一格来决定每条边取哪一份声明，跨多行多列的合并块按它自己的首末行列判定，`TableCell.Borders` 再逐边覆盖（四边的 `TableCellBorders` 优先于表格方向）。从前这四条边一律取 `InsideHorizontal`，另外五份声明连一个字节都写不出门——外圈加粗、内部变淡这种最常见的写法，存一次盘就塌回一种线。`TableBorders` 的六个方向默认都是 0.5 磅黑色细实线，没动过 `Borders` 的稿子字节不变。`BorderStyle.None` 是「这条边不画」，写出 `<a:lnT><a:noFill/></a:lnT>`（同库 DOCX 侧写 `w:val="nil"`，同一个意思）；从前这里不看线型，照旧补一条 `a:solidFill`，于是 `Table.Borders = TableBorders.None()` 存出门是一张满黑格的表，声明被反着执行。读端也解析这四条边线，逐格落进 `TableCell.Borders`：四条里一条都没声明的单元格保持 `null`，让表级六向继续做主。`a:noFill` 认成 `BorderStyle.None`（「这条边不画」），不是「没说过」。要说边界：`.pptx` 没有表级边框元素，六向声明本来就摊在每格的边线上，所以 `Table.Borders` 不会从文件里回来——载入别人的表格后权威在逐格那一份，整表改边框得先清掉每格的 `Borders`。划线图案的细档（`dashDot`、`lgDash` 这类）在统一模型里归成 `Dashed`，`a:lnX@algn` 与逐格的 `marT`/`marB`/`marR` 内边距不往返。

图片的底色与描边走与 `p:sp` 同一条规矩：**设过才写**。写出器过去给每张 `p:pic` 恒补一条 `<a:ln><a:noFill/></a:ln>`，等于替调用者宣布「这张图不要边框」，而模型里从没说过这句话；`BorderStyle.None` 是调用者自己说的，照旧写出。另一半在读端：`ReadPicture` 过去既不调 `ReadFill` 也不调 `ReadOutline`（`p:sp` 那条路这两行都在），写端留在盘上的 `a:solidFill` 与 `a:ln` 载回来就没了——自家稿子 Load+Save 一趟把明写过的两条声明双双抹掉，第三方稿的图片底色与边框第一次读就丢。现在两条路对齐：文件里没写时读回 `null`（不是「不要线」），二次写出既不补 `a:ln` 也不吞声明，图形页两次写出 `slide5.xml` 字节一致。产物字节的变化有两处：从没设过描边的图片少掉那条恒补的 `a:ln`（这一变化本身就是修复）；设过底色或描边的图片，其声明现在能活过第二次保存。

## 入场动画、自动换页与主题深改

每个 `SlideShape` 都可以挂一个入场动画；写出器据此生成 `p:timing` 主序列（PowerPoint 原生动画节点），读回时按 `presetClass="entr"` 的行为节点还原类别、延迟与时长：

```csharp
Slide animated = presentation.AddSlide(SlideLayoutKind.Blank);
TextBoxShape headline = animated.AddTextBox(96, 80, 768, 70, "季度收入 1,280 万元");
headline.Animation = SlideAnimationKind.FlyIn;         // None / Fade / FlyIn / Wipe / Zoom / Split
headline.AnimationDurationSeconds = 0.8;               // 0 表示按类别取默认时长
headline.AnimationDelaySeconds = 0.5;                  // 写出后自动触发时的延迟

TextBoxShape subhead = animated.AddTextBox(96, 160, 768, 50, "环比增长 12%");
subhead.Animation = SlideAnimationKind.Wipe;

animated.AdvanceAfterSeconds = 3.0;                    // p:transition 的 advTm，3 秒后自动换页
```

`PresentationTheme` 覆盖 `theme1.xml` 的六个色板阶段与主/次字体（null 字段保持内置默认；强调色仍走 `AccentColors`），母版与版式经 `clrMap` 自动跟随：

```csharp
presentation.Theme = new PresentationTheme
{
    Name = "品牌主题",
    Text = new OfficeColor(0x1A, 0x1A, 0x2E),          // dk1
    Background = new OfficeColor(0xFA, 0xFA, 0xF5),    // lt1
    TextSecondary = new OfficeColor(0x2F, 0x4D, 0x6B), // dk2
    Hyperlink = new OfficeColor(0xC0, 0x39, 0x2B),
    FontMajor = "微软雅黑",                             // majorFont/latin
    FontMinor = "Consolas"                             // minorFont/latin
};
```

模板格式 `.potx` 与 `.pptx` 包结构完全一致，仅 `/ppt/presentation.xml` 的内容类型不同：

```csharp
PptxWriter.WriteTemplate(presentation, "brand.potx");  // 文件 / 流 / 字节重载均有
OfficeFile file = OfficeFile.Load("brand.potx");       // 识别为 OfficeFileFormat.Potx
file.Save("brand.pptx");                               // 按扩展名转回普通演示文稿
```

## 读取 PPTX

```csharp
OfficePresentation loaded = PptxReader.Load("review.pptx");   // 亦支持 Stream / byte[]

Console.WriteLine($"{loaded.Slides.Count} 页，{loaded.SlideWidth} x {loaded.SlideHeight}");
foreach (Slide slide in loaded.Slides)
{
    Console.WriteLine(slide.GetTitle());                       // 标题占位符文本
    foreach (SlideShape shape in slide.Shapes)
    {
        TextBoxShape box = shape as TextBoxShape;
        if (box != null) Console.WriteLine(box.GetText());
        TableShape table = shape as TableShape;
        if (table != null) Console.WriteLine($"{table.Table.Rows.Count} 行");
    }
}

string plain = PptxReader.ExtractText("review.pptx");          // 只要纯文本时用这个
```

读取覆盖页数与幻灯片尺寸、版式类别、占位符类型、形状几何与填充、段落与行内格式（粗体、斜体、下划线、删除线、上下标、字号、字体、颜色）、项目符号、表格（含合并单元格）、图片字节与替代文字、图片自身的底色与描边（`p:spPr` 上的 `a:solidFill` 与 `a:ln`）、演讲者备注、切换效果、入场动画（类别 / 延迟 / 时长）、`advTm` 自动换页、主题深改字段（`PresentationTheme`，`.potx` 同样可加载）以及 `core.xml` 里的文档元数据。

`.pptx` 里没有可直接对照的"版式类别"字段，读回时分两层：先看幻灯片所用版式部件自己声明的 `p:sldLayout@type`（`secHead`，或 `cSld@name` 为 `Section Header` / `节标题`），命中即是 `SectionHeader`——章节过渡页只有一个标题占位符，与「仅标题页」形状完全一样，形状推断永远分不开，这是唯一能带回它的证据；其余一律按形状推断（`PptxReader.InferLayoutKind`）：先看图片形状或 `pic` 占位符，再看表格形状与 `tbl` / `chart` 占位符，都没有时按标题与内容占位符的个数归入 `TitleAndContent` / `TwoContent` / `Title` / `Blank`。因此往返后版式名仍可能被"修正"，实测两条：`AddSlide(TitleAndTable)` 生成的表格占位符按 `obj` 类型写出，回读变成 `TitleAndContent`；反过来，`AddSlide(TitleAndContent)` 上手工 `AddTable(...)` 的真实表格形状，回读会被推断成 `TitleAndTable`。占位符几何（实测 `Left` 60 原样读回）、图片字节与 alt 文本、备注、切换类型与时长（`Push` / 2.25 秒）、背景色（`#F2F7FF`）与自选图形（`RoundedRectangle` 填充 `#00B050`、`CornerRadius` 默认 0.15）都按原值读回；表格的列定义 `Table.Columns` 写出时不入库（实测写出为 2 行 0 列），回读按每行单元格数补齐（2 行 2 列）。

## 经 `Wisdom.Office` 门面使用

```csharp
using Wisdom.Office;
using Wisdom.Office.Core;

OfficeFile file = OfficeFile.Load("review.pptx");              // 按文件头识别为 Pptx
Console.WriteLine(OfficeFileFormats.GetDisplayName(file.Format));   // PowerPoint 演示文稿
Console.WriteLine(file.ExtractText());                               // 第 N 页：标题 + 各形状文本

OfficeDocument asWord = file.AsDocument();                     // 每页转成一节，标题降级为段落
OfficePresentation backToSlides = file.AsPresentation();       // Word 文档按一级标题反向拆分

file.Save("review.pdf");                                       // 需要 Wisdom.Office 门面包
```

`OfficeConverter` 还提供了显式转换入口：

```csharp
OfficeConverter.PresentationToPdf(presentation, "review.pdf");
OfficeConverter.PresentationToWord(presentation, "review.docx");
OfficeConverter.DocumentToPptx(document, "from-word.pptx");    // byte[] 重载亦有
```

## 类型一览

| 类型 | 用途 |
| --- | --- |
| `PptxWriter` | `OfficePresentation` → PPTX / POTX（文件 / 流 / 字节，模板走 `WriteTemplate`） |
| `PptxReader` | PPTX / POTX → `OfficePresentation`，`ExtractText` |
| `OfficePresentation` | 演示文稿模型：尺寸、元数据、默认字体、强调色、`Theme` |
| `PresentationTheme` | 主题深改：六个色板阶段、主/次字体、主题名 |
| `Slide` | 单页：形状列表、版式、背景色、备注、切换效果、`AdvanceAfterSeconds` |
| `SlideShape` / `TextBoxShape` / `AutoShape` / `PictureShape` / `TableShape` | 形状族（`SlideShape` 含入场动画三属性） |
| `SlideLayoutKind` / `PlaceholderKind` / `AutoShapeKind` / `SlideTextAnchor` / `SlideTransitionKind` / `SlideAnimationKind` | 枚举 |

以上模型类型都在 `Wisdom.Office.Core` 命名空间，读写器在 `Wisdom.Ppt`。

## 已知限制

- 不支持写回旧版 `.ppt`（PowerPoint 97-2003 二进制），也不读取 `.ppt`。
- 不处理图表、SmartArt 与嵌入 OLE 对象；读取时忽略这些部件，写出时不生成。动画只支持入场类别（`p:timing` 主序列），强调、退出、路径与媒体动画不建模；读取时只识别入场节点，其余 timing 节点跳过。
- 版式与主题底稿由本库生成一套固定的 `slideMaster` / `slideLayout` / `theme`，深改经 `PresentationTheme` 显式覆盖六个色板阶段与主/次字体；`.potx` 模板可读写，但不继承第三方母版的精确占位符坐标（占位符位置仍按幻灯片尺寸等比推算）。
- 文本换行、自动字号、形状阴影与 3D 效果由 PowerPoint / WPS 运行时决定，本包只输出对应属性，不做排版模拟。
- 列表符号只有 `Paragraph.ListMarker` 一个字段：没有标记就不写任何 bullet 子元素，符号由母版与版式继承决定。因此「显式关掉上层符号」（`a:buNone`）与「没设置过符号」在模型里是同一个值，读入带 `buNone` 的第三方演示稿再存盘时，那份显式关闭不会被保留。隔行条纹（`bandRow`）同理没有对应字段。
- 折行只表达「不许折行」这一档：`TextBoxShape.WrapText` 置 `false` 才写 `wrap="none"`，`true` 与 `null` 在文件里是同一个样子（都写 `square`），回读只会得到 `null`；`wrap="emergency"`（紧急收缩）本库不建模，读来当作没说过。
- 表格边框逐格往返，表级六向不回：写出时按 `Table.Borders` 的六个方向给每个单元格写四边 `a:lnL`/`lnR`/`lnT`/`lnB`（贴外圈的那条边取 `Top`/`Bottom`/`Left`/`Right`，内部取两条分隔线；含双线的 `cmpd="dbl"`、点线的 `sysDot` 与「无边框」的 `a:noFill` 编码），读回时这四条边逐格落进 `TableCell.Borders`。`Table.Borders` 的六向声明不从文件还原——`.pptx` 本来就没有表级边框元素，声明摊在每格的边线上；载入别人的表格后要整表改边框，得先把每格的 `Borders` 清掉。划线图案的细档（`dashDot`、`lgDash` 等）归成 `Dashed`，`a:lnX@algn` 与逐格 `marT`/`marB`/`marR` 内边距不往返。
- 图片仅支持 PNG / JPEG。
- 文本里 XML 1.0 不接受的字符（控制字符、未配对代理项）在写出前由 `OfficeXmlText` 剔除，避免一个脏字符让整个演示文稿保存失败；因此这类字符不会往返保留。
