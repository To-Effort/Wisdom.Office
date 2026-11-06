# Markdown 进、Word 出：Wisdom.Markdown 与统一文档模型

> 「纯 .NET 操作 Office」系列 · Markdown 篇

## 写在前面

这个包在系列里的角色和其他几个不一样。`Wisdom.Word` / `Wisdom.Excel` / `Wisdom.Ppt` 都是"读写某种 Office 文件"，`Wisdom.Markdown` 补的是**一个纯文本入口**：业务方用 Markdown 写文档，程序负责把它变成 PDF / Word / PPT；反方向也成立，任何文档都能回写成 Markdown。

实现上是纯 C# 手写解析器，不依赖 Markdig、CommonMark.NET 这类外部解析器，唯一依赖是共享模型包 `Wisdom.Office.Core`。这个包定义了整套库共用的文档模型：`OfficeDocument → Section → Block`，`Block` 只有 `Paragraph`、`Table`、`PageBreak` 三种，行内内容是 `TextRun` / `FieldRun` / `ImageRun`。Markdown 侧用得到的只有标题级别、列表标记、行内格式、图片和表格。

```bash
dotnet add package Wisdom.Markdown
```

支持 `net461` 到 `net10.0` 六份目标框架，同一份源码编译。命名空间 `Wisdom.Markdown`，对外只有两个类型：门面 `Markdown` 和选项 `MarkdownOptions`，解析器与写出器都是内部类。要转 PDF / Word 的话，实际项目里一般直接引 `Wisdom.Office` 门面包。

## 一、核心设计：解析结果不引入任何新类型

这是本库最值得讲的架构决策。很多库会定义一套 Markdown 专属 AST（`Document` / `Heading` / `ListItem` 之类），转 Word 时再单独写一遍映射。这里没有——解析产物**直接就是共享的文档模型 `OfficeDocument`**，Markdown 语法逐条落在已有字段上：

| Markdown 语法 | 落到文档模型的形式（实测） |
| --- | --- |
| `# 标题` | `Paragraph.HeadingLevel`（超过 6 级按 6 级） |
| `[官网](url)` | `TextRun.Link` 保存目标地址，同时置为蓝色下划线 |
| `- 项` | `Paragraph.ListMarker = "•"` |
| `- [x] 项` / `- [ ] 项` | `ListMarker = "☑ •"` / `"☐ •"`，勾选状态记在标记前缀里 |
| GFM 表格 | `Table` + 分隔行那一 `TableRow.IsHeader = true`，三种对齐落在单元格段落的 `Alignment` |
| 列表嵌套 | `LeftIndent`，第 N 层为 `12 + (N-1) * 24` 磅 |

好处是解析完就已经在流水线上了：`Markdown.ToDocument` 出来的对象，`Wisdom.Word` / `Wisdom.Pdf` 认得，`AsPresentation()` 也认得（`HeadingLevel` 正是文档拆幻灯片分页的依据）。反过来也一样，`Markdown.ToMarkdown` 的输入不必来自 Markdown。代价是模型里没有的语义只能近似表达或者丢弃，下面「已知边界」逐条列了。

字段是双向可读写的，手工构造模型再吐成 Markdown 也成立：

```csharp
// 本包只提供字段载体；解析与回写在 Wisdom.Markdown。
Paragraph heading = new Paragraph("季度经营回顾") { HeadingLevel = 1 };  // "# 季度经营回顾"
Paragraph task = new Paragraph { ListMarker = "☑ •" };                   // "- [x] 支持 Markdown 解析"
task.AddText("支持 Markdown 解析");
TextRun link = new TextRun("官网") { Link = "https://example.com" };     // "[官网](https://example.com)"
```

层级换算 `12 + (N-1) * 24` 磅这条公式是双向一致的：写器按"每 24 磅算一层"还原，所以解析器给出的缩进正好落进这个换算，五层嵌套可以原样往返。

## 二、语法覆盖

| 类别 | 支持写法 | 说明 |
| --- | --- | --- |
| 标题 | ATX `#`~`######`、Setext（`===` / `---`） | — |
| 强调 | `**粗**`、`*斜*`、`***粗斜***`、`~~删除线~~`、`` `行内代码` `` | 对应 `TextRun` 的 `Bold` / `Italic` / `Strikeout` / 等宽 `TextFormat` |
| 链接 | `[文字](URL "title")`、相对路径、`<https://…>`、`<mailto:…>` | title 会被识别，但模型无对应字段，落地时丢弃 |
| 图片 | `![alt](本地路径)` | 本地 PNG / JPEG 直接嵌入，`Width`/`Height` 留 0 由导出器按像素与 DPI 回填自然尺寸；网络地址与缺失文件降级为 alt 文本 |
| 列表 | `-` `*` `+`、`1.` / `1)`、缩进嵌套、任务列表 | 见上表 |
| 代码块 | 围栏 ```` ``` ```` 与 `~~~`、4 空格缩进块 | 落成连续等宽段落，字号比正文小 0.5 磅 |
| 引用 | `>` 单层、`>>` 多层 | 斜体段落 + 递进左缩进（单层 36 磅，每加一层 18 磅） |
| 表格 | GFM 管道表，`:---` / `---:` / `:---:` | 灰底加粗表头 |
| 其他 | 分割线 `---` / `***`、反斜杠转义、单元格内 `\|` | 分割线落成居中写 `＊ ＊ ＊` 的段落 |
| Front-matter | 文件开头 `---` 包裹的 YAML 块 | 仅取 `title` / `author` / `date` 写入 `document.Metadata` |

几个容易踩的点：图片只支持 PNG / JPEG，网络地址和不存在的本地文件不报错，降级成 alt 文本；引用和分割线是模型里"没有对应块类型"的两处，只能靠段落属性近似；front-matter 是白名单，那三个键之外的字段直接忽略；转义只认反斜杠，单元格里要输出字面 `|` 得写 `\|`。

## 三、三个方向的 API

**Markdown → 文档模型**：

```csharp
using Wisdom.Markdown;
using Wisdom.Office.Core;

OfficeDocument document = Markdown.ToDocument(markdownText);   // 默认选项
OfficeDocument fromFile = Markdown.LoadFile("readme.md");      // UTF-8 读取并自动去 BOM

foreach (Section section in document.Sections)
    foreach (Block block in section.Blocks)
    {
        if (block is Paragraph p && p.HeadingLevel > 0)
            Console.WriteLine($"{p.HeadingLevel} 级标题：{p.GetText()}");
        if (block is Table t) Console.WriteLine($"{t.Rows.Count} 行 x {t.Rows[0].Cells.Count} 列");
    }

Console.WriteLine(document.Metadata.Title);   // 来自 front-matter
```

**Markdown → HTML**：`string html = Markdown.ToHtml(markdownText);`，细节见第五节。

**文档模型 → Markdown**：模型侧照常构造，最后一步 `Markdown.ToMarkdown(document)`。

```csharp
OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "发布说明";
Section section = document.AddSection();

Paragraph heading = section.AddParagraph("新版本要点");
heading.HeadingLevel = 2;

Paragraph item = section.AddParagraph("支持 Excel 原生图表");
item.ListMarker = "-";

Table table = new Table();
table.AddRow("模块", "状态");
table.AddRow("Word", "已完成");
section.Blocks.Add(table);

string markdown = Markdown.ToMarkdown(document);
```

## 四、MarkdownOptions

```csharp
MarkdownOptions options = new MarkdownOptions
{
    HardLineBreaks = false,           // 默认 false：单个换行不断行，行尾两空格才硬换行
    EnableTables = true,              // 默认 true：解析 GFM 管道表格
    EnableTaskLists = true,           // 默认 true：解析 - [ ] / - [x]
    EnableRawHtml = false,            // 默认 false：原始 HTML 标签按纯文本保留
    CodeFontFamily = "Consolas",      // 默认 Consolas：代码块与行内代码的等宽字体
    BodyFontFamily = null,            // 默认 null：沿用文档默认字体
    BaseFontSize = 11.0,              // 默认 11 磅：正文字号基准
    TrimTrailingWhitespace = true     // 默认 true：去行尾空白（硬换行的两空格除外）
};
```

`MarkdownOptions` 同时影响解析和渲染：README 里给出的用法是 `Markdown.ToDocument(markdownText, options)` 与 `Markdown.ToHtml(markdownText, options)`，回写侧的下划线与颜色是否用 `<u>` / `<span>` 表达也由这里的 `EnableRawHtml` 决定。

中文项目里最常动的是字体和换行。`BodyFontFamily` 设成 `"Microsoft YaHei"` 之后导出的 Word / PDF 会整体跟着换，代码块仍走 `CodeFontFamily`，字号在 `BaseFontSize` 基础上小 0.5 磅。`HardLineBreaks` 与 `TrimTrailingWhitespace` 是一对：默认单个换行不断行、只有行尾两空格才算硬换行，而去行尾空白时这两空格是例外。实际项目里 `options` 常常是静态共享的，改之前用 `options.Clone()` 复制一份，别污染共用实例。

## 五、HTML 渲染与安全处理

输出的是 HTML **片段**，不含 `<html>` / `<head>` 外壳。段落转 `<p>`、标题转 `<h1>`~`<h6>`、列表转嵌套 `<ul>` / `<ol>`、任务项转 `<input type="checkbox" disabled>`、连续等宽段落转 `<pre><code>`、表格转 `<table>`、分割线转 `<hr />`。图片一律转成 `data:` URI 内联，产物是个能离线打开的单文件。

安全这块的处理值得单独说，因为 Markdown 渲染器最常见的事故就是链接协议放行过宽：

- `<a href>` 只接受相对地址、`#锚点` 与 `http` / `https` / `mailto` / `ftp` / `ftps` / `tel`；
- 其余协议（`javascript:`、`data:` 等）**不进 `href`**，渲染成蓝 + 下划线的纯文本 `<span>`——看着像链接但点不动；
- 文本与属性值统一做 `&`、`<`、`>`、`"` 转义，`EnableRawHtml` 也管不到这一步；
- 行首就是 `<` 的整行按原始 HTML 行透传，该行不再做行内解析（`<b>粗</b> 与 [链接](https://example.com)` 里的链接不会转成 `<a>`）；`<` 出现在行中则只转义标签本身，其余照常解析。

`EnableRawHtml` 的作用范围比名字听起来窄：HTML 渲染阶段所有文本一律转义，它只影响回写 Markdown 时用 `<u>` / `<span>` 表达下划线与颜色。`<b>`、`<br>` 这类行内标签在 HTML 出口永远是纯文本。

要完整页面时自己包外壳，标题从 front-matter 来：

```csharp
string page = "<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">"
    + "<title>" + document.Metadata.Title + "</title></head><body>"
    + html + "</body></html>";
```

## 六、回写规则

能还原的：粗体、斜体、删除线、行内代码、链接 `[文字](URL)`、图片、表格与三种对齐。符号层面有几条明确约定——无序列表统一输出 `- `（不管原文用 `-` 还是 `*`），有序列表按原分隔符（`.` 或 `)`）从 1 重新编号，也就是说原文从 `5.` 开始的续编号会归位成 `1.`，任务项按 `ListMarker` 里的 `☑` / `☐` 还原成 `- [x]` / `- [ ]`，代码块用围栏输出。`Metadata.Title` 若文档首个标题不是它，则在开头补一个 `# 标题`。标题行与表头行不再重复包 `**`，因为这两种语境本身就带强调样式。

代码块的判定方式值得留意：写器按"整段等宽"识别代码块，模型里并没有 `CodeBlock` 这个块类型。这是共用一套字段的代价之一——正文里如果整段都被显式设成等宽字体，回写时它会被当成代码块。

不回写的：引用的 `>`（只剩 `*斜体*`）、front-matter（不重新生成 `---` 块）、链接 title、图片的本地路径与 alt 文本（模型不带 alt，统一写成 `![image](data:image/png;base64,…)`）、字体名与字号（Markdown 里没有对应表达）。

## 七、经 Wisdom.Office 门面使用

```csharp
using Wisdom.Office;

OfficeFile file = OfficeFile.Load("report.md");              // .md / .markdown / .mkd
Console.WriteLine(OfficeFileFormats.GetDisplayName(file.Format));  // Markdown 文档
Console.WriteLine(file.ExtractText());

file.Save("report.pdf");                                     // Markdown → PDF
OfficeConverter.Convert(file, "report.docx");                // Markdown → DOCX
string markdown = OfficeConverter.ToMarkdown(file.AsDocument());  // 任意文档 → Markdown
```

有一点要提前知道：**Markdown 是纯文本，没有内容签名**，`OfficeFileFormats.Detect` 对它返回 `Unknown`。`OfficeFile.Load(path)` 是在识别失败后按扩展名兜底的，所以从路径加载没问题；从字节或流加载时没有扩展名可用，必须显式指定 `OfficeFile.Load(bytes, OfficeFileFormat.Markdown)`，`Load(byte[])` 与 `Load(Stream)` 对纯文本会抛 `NotSupportedException`。

按 `Wisdom.Office` 实测的转换矩阵，Markdown 作为源可以到 PDF、DOCX、PPTX、Markdown 四种目标；`ExtractText()` 拿到的是不带标记的纯文本，适合做搜索索引。

## 已知边界

- 不是完整的 CommonMark / GFM 实现。不支持引用式链接 `[text][ref]` 与其定义行、脚注 `[^1]`、HTML 块内再套 Markdown、列表项内以缩进续写的多段落（这些续行退化成不带标记的普通段落）。以上都不抛异常，只按字面文本保留。
- 表格单元格内的块级内容（列表、代码块）不解析，只按行内文本处理；`Table.Columns` 在 Markdown 表里只是 `Width = 0` 的列占位，行列数要看 `Rows` 与 `Cells`。
- 代码围栏的语言标记（```` ```csharp ````）在文档模型里没有对应字段，解析时丢弃，回写只输出无语言的围栏，HTML 也不加 `class`。
- `TextRun.Link` 目前只在 Markdown 侧填充。DOCX、PPTX 里的超链接读入模型时只保留文字与格式，URL 不入库，所以**Word 文档转 Markdown 后链接会退化成普通文本**。
- 分割线、引用这类模型里没有对应块类型的语法靠段落属性近似表达：HTML / Markdown 出口能识别回 `<hr />` 和 `---`，Word / PDF 里分割线就是排成居中的 `＊ ＊ ＊` 这几个字符。
- 由 Markdown 解析出的文档再转 Word / PDF 时排版按默认值；回写出的 data URI 图片再解析时同样降级为 alt 文本，字节不再入库——**别用 md → md 往返来存图**。

## 小结

`Wisdom.Markdown` 做的事可以压缩成一句：把 Markdown 翻译进 `OfficeDocument`，再从 `OfficeDocument` 翻译回 Markdown 或 HTML，中间不产生任何私有类型。语法覆盖按常见写作习惯做了取舍，模型表达不了的语义明确降级或丢弃，不静默造假数据。

如果你的需求是"业务方交一份 Markdown，程序出 PDF/Word"，或者"给现有文档加一个 Markdown 导出格式"，这个包配合 `Wisdom.Office` 门面基本够用，而且不涉及任何第三方解析器和它的许可证。
