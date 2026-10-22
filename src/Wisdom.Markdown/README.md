# Wisdom.Markdown

`Wisdom.Markdown` 负责 **Markdown 侧的全部读写**：把 Markdown 文本解析为统一文档模型 `OfficeDocument`，把文档模型回写为 Markdown，并在两者之间提供 HTML 渲染。解析结果可以直接交给 `Wisdom.Word` / `Wisdom.Pdf` 生成 DOCX 或 PDF。

- 包名：`Wisdom.Markdown`
- 版本：`1.0.0`
- 依赖：`Wisdom.Office.Core`（唯一依赖）
- 第三方依赖：无；不依赖 Markdig、CommonMark.NET 等外部解析器
- 命名空间：`Wisdom.Markdown`

本文是能力与 API 参考；分步实战教程见源码仓库的 `docs/Wisdom.Markdown.Tutorial.md`。

## 支持的框架

| 目标框架 | 包内资产 | 典型运行时 |
| --- | --- | --- |
| `net461` | `lib/net461` | .NET Framework 4.6.1 / 4.7.2 / 4.8.x |
| `netstandard2.0` | `lib/netstandard2.0` | .NET Core 2.0-3.1、.NET 5/6、Mono、Unity |
| `net7.0` | `lib/net7.0` | .NET 7 |
| `net8.0` | `lib/net8.0` | .NET 8 |
| `net9.0` | `lib/net9.0` | .NET 9 |
| `net10.0` | `lib/net10.0` | .NET 10 / 11 |

纯 C# 手写解析器，六份资产由同一份源码编译，功能一致。

## 安装

```bash
dotnet add package Wisdom.Markdown --version 1.0.0
```

需要把 Markdown 转成 PDF / Word 时请改引 `Wisdom.Office` 门面包。

## 语法覆盖

| 类别 | 支持写法 | 落到文档模型的形式 |
| --- | --- | --- |
| 标题 | ATX `#` 到 `######`、Setext（下划线 `===` / `---`） | `Paragraph.HeadingLevel`（超过 6 级按 6 级） |
| 强调 | `**粗体**`、`*斜体*`、`***粗斜体***`、`~~删除线~~`、`` `行内代码` `` | `TextRun` 的 `Bold` / `Italic` / `Strikeout` / 等宽 `TextFormat` |
| 链接 | `[文字](URL "title")`、`[文字](相对路径)`、`<https://…>`、`<mailto:…>` 自动链接 | `TextRun.Link` 保存目标地址，同时置为蓝色下划线；title 字符串被识别但模型无对应字段，落地时丢弃 |
| 图片 | `![alt](本地路径)` | `ImageRun`；本地 PNG / JPEG 直接嵌入（`Width`/`Height` 留 0，由导出器按图片像素与 DPI 回填自然尺寸），网络地址与缺失文件降级为 alt 文本 |
| 列表 | `-` `*` `+` 无序、`1.` / `1)` 有序、缩进嵌套、`- [ ]` / `- [x]` 任务列表 | `Paragraph.ListMarker` + `LeftIndent`（第 N 层为 `12 + (N-1) * 24` 磅，写器按同一换算还原层级）；任务项的勾选状态记在标记前缀 `☐` / `☑` 里 |
| 代码块 | 围栏 ```` ``` ```` 与 `~~~`（可带语言信息）、4 空格缩进块 | 连续等宽字体段落（字号比正文小 0.5 磅），写出时按"整段等宽"识别为代码块 |
| 引用 | `>` 单层与 `>>` 多层 | 斜体段落 + 递进左缩进（单层 36 磅、每加一层 18 磅）；模型没有引用块类型，回写 Markdown 与渲染 HTML 时只剩斜体 |
| 表格 | GFM 管道表格，含 `:---`、`---:`、`:---:` 三种对齐 | `Table` + `TableRow.IsHeader` + 灰底加粗表头；三种对齐落在每个单元格段落的 `Alignment` 上，`Table.Columns` 只按列数补 `Width = 0` 的占位 |
| 其他 | 分割线 `---` / `***`、反斜杠转义、`\|` 单元格内转义 | 分割线落成居中、写着 `＊ ＊ ＊` 的段落（段落模型无下边框属性）；HTML 与 Markdown 出口按同一常量识别回 `<hr />` 与 `---`，Word / PDF 仍排这几个字符 |
| Front-matter | 文件开头 `---` 包裹的 YAML 块 | 仅识别 `title` / `author` / `date`，写入 `document.Metadata` |

## Markdown → 文档模型

```csharp
using Wisdom.Markdown;
using Wisdom.Office.Core;

OfficeDocument document = Markdown.ToDocument(markdownText);       // 默认选项
OfficeDocument fromFile = Markdown.LoadFile("readme.md");          // UTF-8 读取并自动去 BOM

foreach (Section section in document.Sections)
{
    foreach (Block block in section.Blocks)
    {
        if (block is Paragraph paragraph && paragraph.HeadingLevel > 0)
        {
            Console.WriteLine($"{paragraph.HeadingLevel} 级标题：{paragraph.GetText()}");
        }
        if (block is Table table) Console.WriteLine($"{table.Rows.Count} 行 x {table.Rows[0].Cells.Count} 列");
    }
}

Console.WriteLine(document.Metadata.Title);                         // 来自 front-matter
```

## 渲染选项

```csharp
using Wisdom.Markdown;

MarkdownOptions options = new MarkdownOptions
{
    HardLineBreaks = false,            // 默认 false：单个换行不断行，行尾两空格才硬换行
    EnableTables = true,               // 默认 true：解析 GFM 管道表格
    EnableTaskLists = true,            // 默认 true：解析 - [ ] / - [x]
    EnableRawHtml = false,             // 默认 false：false 时原始 HTML 标签按纯文本保留
    CodeFontFamily = "Consolas",       // 默认 Consolas：代码块与行内代码的等宽字体
    BodyFontFamily = null,             // 默认 null：沿用文档默认字体
    BaseFontSize = 11.0,               // 默认 11 磅：正文字号基准
    TrimTrailingWhitespace = true      // 默认 true：去除行尾空白（硬换行的两空格除外）
};

OfficeDocument document = Markdown.ToDocument(markdownText, options);
string html = Markdown.ToHtml(markdownText, options);
```

`options.Clone()` 可复制一份再改，避免影响共用实例。

## Markdown → HTML

```csharp
string html = Markdown.ToHtml(File.ReadAllText("article.md", Encoding.UTF8));
```

输出的是 HTML **片段**（不含 `<html>` / `<head>` 外壳）：段落转 `<p>`、标题转 `<h1>`~`<h6>`、列表转嵌套 `<ul>` / `<ol>`、任务项转 `<input type="checkbox" disabled>`、连续等宽段落转 `<pre><code>`、表格转 `<table>`、分割线转 `<hr />`、链接转 `<a href="…">`。图片一律转成 `data:` URI 内联，产物可离线单文件打开。

`<a href>` 只接受相对地址、`#锚点` 与 `http` / `https` / `mailto` / `ftp` / `ftps` / `tel`；其余协议（`javascript:`、`data:` 等）不进 `href`，渲染成蓝 + 下划线的纯文本 `<span>`——看着像链接但点不动。文本与属性值统一做 `&`、`<`、`>`、`"` 转义，`EnableRawHtml` 也管不到这一步。行首就是 `<` 的整行按原始 HTML 行透传，该行不再做行内解析（`<b>粗</b> 与 [链接](https://example.com)` 里的链接不会转成 `<a>`）；`<` 出现在行中则只转义标签本身，其余照常解析。

需要完整页面时自行包裹：

```csharp
string page = "<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">"
    + "<title>" + document.Metadata.Title + "</title></head><body>"
    + html + "</body></html>";
```

## 文档模型 → Markdown

```csharp
using Wisdom.Markdown;
using Wisdom.Office.Core;

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
table.AddRow("PPT", "已完成");
section.Blocks.Add(table);

string markdown = Markdown.ToMarkdown(document);
```

回写规则：无序符号统一为 `- `，有序列表按原分隔符（`.` 或 `)`）从 1 重新编号；缩进每 24 磅算一层（解析器给出的 `12 + (N-1) * 24` 磅正好落进这个换算，五层嵌套可原样还原）；任务项按 `ListMarker` 里的 `☑` / `☐` 还原成 `- [x]` / `- [ ]`；代码块用围栏输出（围栏不保留语言标记，模型里没有这个字段）；`Metadata.Title` 若文档首个标题不是它，则在开头补一个 `# 标题`；粗体、斜体、删除线、行内代码、链接（`[文字](URL)`）、图片、表格与三种对齐都能还原；标题与表头行不再重复包 `**`（这两种语境本身就是强调样式），下划线与字体颜色只在 `EnableRawHtml = true` 时以 `<u>` / `<span>` 形式保留；字体名与字号在 Markdown 里没有对应表达，会被丢弃。

不回写的还有：引用的 `>`（只剩 `*斜体*`）、front-matter（不重新生成 `---` 块）、链接 title、图片的本地路径与 alt 文本（模型不带 alt，统一写成 `![image](data:image/png;base64,…)`）。

## 经 `Wisdom.Office` 门面使用

```csharp
using Wisdom.Office;
using Wisdom.Office.Core;

OfficeFile file = OfficeFile.Load("report.md");                    // .md / .markdown / .mkd
Console.WriteLine(OfficeFileFormats.GetDisplayName(file.Format));   // Markdown 文档
Console.WriteLine(file.ExtractText());

file.Save("report.pdf");                                           // Markdown → PDF
OfficeConverter.Convert(file, "report.docx");                       // Markdown → DOCX
string markdown = OfficeConverter.ToMarkdown(file.AsDocument());    // 任意文档 → Markdown
```

## 类型一览

| 类型 | 用途 |
| --- | --- |
| `Markdown` | 门面：`ToDocument` / `LoadFile` / `ToHtml` / `ToMarkdown` |
| `MarkdownOptions` | 解析与渲染开关、字体与字号 |

内部类（`MarkdownParser`、`MarkdownHtmlWriter`、`OfficeDocumentMarkdownWriter`）不对外暴露，统一走 `Markdown` 门面。

## 已知限制

- 不是完整 CommonMark / GFM 实现：不支持引用式链接 `[text][ref]` 与其定义行、脚注 `[^1]`、HTML 块内再套 Markdown、列表项内以缩进续写的多段落内容（这些续行会退化成不带标记的普通段落）。以上都不抛异常，只按字面文本保留。
- Front-matter 只取 `title` / `author` / `date` 三个字段，其余键忽略；回写 Markdown 时不会重新生成 front-matter 块。
- 网络图片与不存在的本地图片降级为 alt 文本；图片仅支持 PNG / JPEG。回写出的 data URI 图片再解析时同样降级为 alt 文本，字节不再入库，别用 md → md 往返存图。
- 表格单元格内的块级内容（列表、代码块）不解析，只按行内文本处理；`Table.Columns` 在 Markdown 表里只是 `Width = 0` 的列占位，行列数看 `Rows` 与 `Cells`。
- 代码围栏的语言标记（```` ```csharp ````）在文档模型里没有对应字段，解析时丢弃，回写只输出无语言的围栏；HTML 也不加 `class`。
- 行内 HTML 标签（`<b>`、`<br>` 等）始终按纯文本保留，`EnableRawHtml` 只影响回写 Markdown 时用 `<u>` / `<span>` 表达下划线与颜色，HTML 渲染阶段所有文本一律转义。
- `TextRun.Link` 目前只在 Markdown 侧填充：DOCX、PPTX 里的超链接读入模型时只保留文字与格式，URL 不入库，因此 Word 文档转 Markdown 后链接会退化为普通文本。
- 字体、字号、缩进磅值等文档模型专有格式在 Markdown 中无对应语法，`ToMarkdown` 会丢失；由 Markdown 解析出的文档再转 Word / PDF 时排版按默认值。
- 分割线、引用这类"模型里没有对应块类型"的语法靠段落属性近似表达：分割线是居中的 `＊ ＊ ＊` 段落（HTML / Markdown 出口识别回 `<hr />` / `---`，Word / PDF 排成这行字符），引用是斜体 + 缩进段落（回写只剩斜体）。

分步实战与全部实测数字见 `docs/Wisdom.Markdown.Tutorial.md`。
