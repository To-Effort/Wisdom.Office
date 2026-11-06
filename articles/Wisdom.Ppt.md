# 不装 Office 也能生成 PPT：Wisdom.Ppt 使用指南

> 「纯 .NET 操作 Office」系列 · PowerPoint 篇

## 写在前面

服务端生成 PowerPoint，比 Excel / Word 更让人头疼。`Interop` 要求机器上装 Office，服务端基本不能用；`Open XML SDK` 虽然纯净，但要手写冗长的 PresentationML，还得时刻记住坐标是 EMU、位移写在 `a:off`、尺寸写在 `a:ext`；商用库则绕不开授权。

`Wisdom.Ppt` 的定位很明确：**用模型描述幻灯片，写出器负责 PresentationML**。你用 `OfficePresentation` 摆好版式、形状、文本、动画，`PptxWriter` 把它落成 `.pptx`；反过来 `PptxReader` 把 `.pptx` 读回同一模型。全程由 `System.IO.Compression.ZipArchive` + LINQ to XML 直接处理，不依赖 Office、WPS 或任何第三方库，唯一依赖是自己家的数据模型包 `Wisdom.Office.Core`。

```bash
dotnet add package Wisdom.Ppt
```

支持 `net461` / `netstandard2.0` / `net7.0` / `net8.0` / `net9.0` / `net10.0` 六份目标框架，同一份源码编译，老 .NET Framework 项目和 Unity 都能引。

## 一、单位与坐标系

这是最容易搞错的地方，先讲清楚。所有几何量以**磅（point）**为单位，写出时按 **1 磅 = 12700 EMU** 换算成 PresentationML 的 `a:off` / `a:ext`，坐标**原点为幻灯片左上角**，X 向右、Y 向下。你不需要碰 EMU，但需要知道画布多大。

| 尺寸 | 入口 | 画布（磅） | 用途 |
| --- | --- | --- | --- |
| 16:9 宽屏 | `OfficePresentation.Widescreen()` | 960 x 540 | 默认，现代投影仪 |
| 4:3 标准 | `OfficePresentation.Standard()` | 720 x 540 | 老式投影与打印 |

摆一个 `left=96, top=60, width=768, height=80` 的文本框，就是距左上角 96 磅、宽 768 磅——在 960 宽的画布上左右各留 96 磅边距。两套尺寸下占位符按各自宽度等比推算，同一份代码换 `Standard()` 会自动收缩。

## 二、构建演示文稿

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
((TextBoxShape)bullets.Shapes[0]).Paragraphs[0].AddText("本季度要点");

TextBoxShape body = FindContent(bullets);
body.Paragraphs[0].AddText("收入 1,280 万元，环比增长 12%");
body.Paragraphs[0].ListMarker = "•";          // 项目符号
body.Paragraphs[0].LeftIndent = 18;           // 缩进（磅）
foreach (string item in new[] { "成本控制在预算内，毛利率 41.6%", "华东区贡献了 46% 的增量收入" })
{
    Paragraph paragraph = body.AddParagraph(string.Empty);
    paragraph.AddText(item);
    paragraph.ListMarker = "•";
    paragraph.LeftIndent = 18;
}
bullets.Transition = SlideTransitionKind.Fade;   // 切换效果
bullets.TransitionDuration = 0.7;                // 秒

PptxWriter.Write(presentation, "review.pptx");   // 亦支持 Stream，或 byte[] Write(presentation)
```

`SlideTransitionKind` 支持 `None`、`Fade`、`Push`、`Wipe`、`Split`。生成的包通过内容类型、关系目标与 XML 合法性校验，可被 PowerPoint 2010+、WPS 演示直接打开。

## 三、七种版式与占位符

`AddSlide` 不带参数等价于 `Blank`，全部形状自己定位。占位符的字体不在模型里写死，由写出器套用 `DefaultTitleFormat` / `DefaultTextFormat`。

| `SlideLayoutKind` | 生成的占位符 |
| --- | --- |
| `Title` | `Title` |
| `TitleAndContent` | `Title` + `Content` |
| `SectionHeader` | `Title`（垂直居中、位置偏低） |
| `TwoContent` | `Title` + 左右两个 `Content` |
| `TitleAndPicture` | `Title` + `Picture` |
| `TitleAndTable` | `Title` + `Table` |
| `Blank` | 无 |

## 四、文本、图形、图片与表格

四种形状对应四种构造方式：

| 形状 | 构造 | 说明 |
| --- | --- | --- |
| `TextBoxShape` | `slide.AddTextBox(left, top, width, height, text)` | `Paragraphs` 列表、`VerticalAlignment`、四边内边距、`ShrinkToFit` |
| `AutoShape` | `slide.AddShape(AutoShapeKind, ...)` | `FillColor`、`Outline`、`CornerRadius` |
| `PictureShape` | `slide.AddImage(new OfficeImage(bytes), ...)` | 仅 PNG / JPEG；`AlternativeText` 作替代文字 |
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

`AutoShapeKind` 支持 `Rectangle`、`RoundedRectangle`、`Ellipse`、`Triangle`、`Diamond`、`Pentagon`、`Hexagon`、`RightArrow`、`Chevron`、`Line` 十种，逐个映射到 PresentationML 的 `prstGeom`。

## 五、入场动画与自动换页

每个 `SlideShape` 都能挂一个入场动画；写出器据此生成 `p:timing` 主序列（PowerPoint 原生动画节点）。三个属性即够：

```csharp
Slide animated = presentation.AddSlide(SlideLayoutKind.Blank);
TextBoxShape headline = animated.AddTextBox(96, 80, 768, 70, "季度收入 1,280 万元");
headline.Animation = SlideAnimationKind.FlyIn;    // None / Fade / FlyIn / Wipe / Zoom / Split
headline.AnimationDurationSeconds = 0.8;          // 0 表示按类别取默认时长
headline.AnimationDelaySeconds = 0.5;             // 自动触发时的延迟

TextBoxShape subhead = animated.AddTextBox(96, 160, 768, 50, "环比增长 12%");
subhead.Animation = SlideAnimationKind.Wipe;

animated.AdvanceAfterSeconds = 3.0;               // p:transition 的 advTm，3 秒后自动换页
```

注意：只有**入场类别**被建模，强调、退出、路径与媒体动画不处理。

## 六、主题深改与 .potx 模板

`PresentationTheme` 覆盖 `theme1.xml` 的六个色板阶段与主/次字体，null 字段保持内置默认（强调色仍走 `AccentColors`），母版与版式经 `clrMap` 自动跟随：

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
```

## 七、读取 PPTX

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
        TableShape tbl = shape as TableShape;
        if (tbl != null) Console.WriteLine($"{tbl.Table.Rows.Count} 行");
    }
}

string plain = PptxReader.ExtractText("review.pptx");          // 只要纯文本时用这个
```

读取覆盖页数与幻灯片尺寸、版式类别、占位符类型、形状几何与填充、段落与行内格式（粗体、斜体、下划线、删除线、上下标、字号、字体、颜色）、项目符号、表格（含合并单元格）、图片字节与替代文字、演讲者备注、切换效果、入场动画、`advTm` 自动换页、`PresentationTheme` 深改字段（`.potx` 同样可加载）以及 `core.xml` 里的文档元数据。

## 八、往返保真：诚实说明差异

`.pptx` 里没有可直接对照的"版式类别"字段，读回时版式是**按形状推断**的（`PptxReader.InferLayoutKind`）：先看图片形状或 `pic` 占位符，再看表格形状与 `tbl` / `chart` 占位符，都没有时按标题与内容占位符的个数归入 `TitleAndContent` / `TwoContent` / `Title` / `Blank`。因此往返后版式名可能被"修正"，这里不粉饰，实测几条差异讲清楚：

- `AddSlide(TitleAndTable)` 生成的表格占位符按 `obj` 类型写出，回读变成 `TitleAndContent`；反过来，`AddSlide(TitleAndContent)` 上手工 `AddTable(...)` 得到的真实表格形状，回读会被推断成 `TitleAndTable`。
- 表头占位形状回读成 `TextBoxShape`。
- 表格的列定义 `Table.Columns` 写出时不入库（实测写出为 2 行 0 列），回读按每行单元格数补齐（2 行 2 列）。

好消息是几何与内容层面基本按原值读回：占位符几何（实测 `Left` 60 原样读回）、图片字节与 alt 文本、备注、切换类型与时长（`Push` / 2.25 秒）、背景色（`#F2F7FF`）与自选图形（`RoundedRectangle` 填充 `#00B050`、`CornerRadius` 默认 0.15）都能回读。一句话：**内容和坐标可靠，版式名以推断为准**。

## 九、经 Wisdom.Office 门面使用

`Wisdom.Ppt` 只做 PPTX 读写；要转 PDF / Word，请改引 `Wisdom.Office` 门面包（本包不依赖 `Wisdom.Pdf` 与 `Wisdom.Word`）。

```csharp
using Wisdom.Office;
using Wisdom.Office.Core;

OfficeFile file = OfficeFile.Load("review.pptx");     // 按文件头识别为 Pptx
OfficeDocument asWord = file.AsDocument();            // 每页转成一节，标题降级为段落
file.Save("review.pdf");                              // 需要门面包

// 显式转换入口
OfficeConverter.PresentationToPdf(presentation, "review.pdf");
OfficeConverter.DocumentToPptx(document, "from-word.pptx");   // byte[] 重载亦有
```

`file.AsPresentation()` 还能把 Word 文档按一级标题反向拆成幻灯片。

## 已知边界

- 不支持读写旧版 `.ppt`（PowerPoint 97-2003 二进制）。
- 不处理图表、SmartArt 与嵌入 OLE 对象：读取时忽略这些部件，写出时不生成。
- 动画只支持入场类别（`p:timing` 主序列）；强调、退出、路径与媒体动画不建模，读取时只识别入场节点。
- 母版 / 版式 / 主题底稿由本库固定生成一套 `slideMaster` / `slideLayout` / `theme`，深改经 `PresentationTheme` 显式覆盖；`.potx` 可读写，但**不继承第三方母版的精确占位符坐标**，占位符位置仍按幻灯片尺寸等比推算。
- 文本换行、自动字号、形状阴影与 3D 效果由 PowerPoint / WPS 运行时决定，本包只输出对应属性，不做排版模拟。
- 图片仅支持 PNG / JPEG。

## 小结

`Wisdom.Ppt` 的思路和系列里其它包一致：把 PresentationML 当作一个可以直接读写的 ZIP + XML 包来处理，你只面对 `OfficePresentation` 模型。它覆盖了服务端生成 PPT 的核心场景——版式、四种形状、十种自选图形、项目符号、备注、切换、入场动画、自动换页、主题深改与 `.potx` 模板，代价是明确放弃旧 `.ppt`、图表 / SmartArt / OLE、非入场动画和排版模拟。

如果目标是"用代码描述一份结构清晰的演示文稿并写出 PPTX"，而不是"100% 还原设计师在 PowerPoint 里排好的复杂母版"，这个包够用，且没有许可证和体积负担。需要跨 PDF / Word 互转的，再看同系列的 `Wisdom.Office` 门面包。
