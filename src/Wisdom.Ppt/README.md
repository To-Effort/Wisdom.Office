# Wisdom.Ppt

`Wisdom.Ppt` 负责 **PowerPoint 侧的全部读写**：用统一演示文稿模型 `OfficePresentation` 构建幻灯片并写出 PPTX，或把 `.pptx` 读回同一模型。全部由 `System.IO.Compression.ZipArchive` + LINQ to XML 直接处理 PresentationML，不依赖 Office、WPS 或任何第三方库。

- 包名：`Wisdom.Ppt`
- 版本：`1.0.2`
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
dotnet add package Wisdom.Ppt --version 1.0.2
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

## 文本、图形、图片与表格

| 形状 | 构造方式 | 说明 |
| --- | --- | --- |
| `TextBoxShape` | `slide.AddTextBox(left, top, width, height, text)` | `Paragraphs` 列表、`VerticalAlignment`、四边内边距、`ShrinkToFit` |
| `AutoShape` | `slide.AddShape(AutoShapeKind, ...)` | `FillColor`、`Outline`、`CornerRadius` |
| `PictureShape` | `slide.AddImage(new OfficeImage(bytes), ...)` | 仅 PNG / JPEG；`AlternativeText` 作为替代文字 |
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
`SlideTransitionKind` 支持 `None`、`Fade`、`Push`、`Wipe`、`Split`，配合 `TransitionDuration`（秒）。

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

读取覆盖页数与幻灯片尺寸、版式类别、占位符类型、形状几何与填充、段落与行内格式（粗体、斜体、下划线、删除线、上下标、字号、字体、颜色）、项目符号、表格（含合并单元格）、图片字节与替代文字、演讲者备注、切换效果、入场动画（类别 / 延迟 / 时长）、`advTm` 自动换页、主题深改字段（`PresentationTheme`，`.potx` 同样可加载）以及 `core.xml` 里的文档元数据。

`.pptx` 里没有可直接对照的"版式类别"字段，读回时按形状推断（`PptxReader.InferLayoutKind`）：先看图片形状或 `pic` 占位符，再看表格形状与 `tbl` / `chart` 占位符，都没有时按标题与内容占位符的个数归入 `TitleAndContent` / `TwoContent` / `Title` / `Blank`。因此往返后版式名可能被"修正"，实测两条：`AddSlide(TitleAndTable)` 生成的表格占位符按 `obj` 类型写出，回读变成 `TitleAndContent`；反过来，`AddSlide(TitleAndContent)` 上手工 `AddTable(...)` 的真实表格形状，回读会被推断成 `TitleAndTable`。占位符几何（实测 `Left` 60 原样读回）、图片字节与 alt 文本、备注、切换类型与时长（`Push` / 2.25 秒）、背景色（`#F2F7FF`）与自选图形（`RoundedRectangle` 填充 `#00B050`、`CornerRadius` 默认 0.15）都按原值读回；表格的列定义 `Table.Columns` 写出时不入库（实测写出为 2 行 0 列），回读按每行单元格数补齐（2 行 2 列）。

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
- 图片仅支持 PNG / JPEG。
- 文本里 XML 1.0 不接受的字符（控制字符、未配对代理项）在写出前由 `OfficeXmlText` 剔除，避免一个脏字符让整个演示文稿保存失败；因此这类字符不会往返保留。
