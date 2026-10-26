# Wisdom.Markdown 使用教程

Markdown 引擎的实战教程：把一份带 front-matter 的发布说明解析成统一文档模型，逐类看标题、强调、链接、列表、任务项、代码块、引用与表格在模型里落成什么形状，渲染成可离线打开的 HTML 片段，再用 `MarkdownOptions` 逐个开关看效果，最后把文档模型回写成 Markdown、经 `Wisdom.Office` 门面转 PDF / Word，实测哪些写法会降级，并逐组核对与 CommonMark 对齐的行内与块级细节。

安装方式见[包说明](../src/Wisdom.Markdown/README.md#安装)。本教程示例是 .NET 8 控制台工程，文中所有输出、字符数与字节数都是这份工程实测打印的；只有第 9 步用到 `Wisdom.Office` 门面包（`OfficeFile` / `OfficeConverter`），其余只用 `Wisdom.Markdown` + `Wisdom.Office.Core`。

## 你将完成

```text
out/
  logo.png          教程自己写出的 1 x 1 PNG，70 字节
  article.md        教程的 Markdown 源文，供门面按扩展名识别
  article.pdf       Markdown 直接导出的 PDF，121895 字节 / 1 页
  article.docx      Markdown 直接导出的 Word，5004 字节
  article-back.md   文档模型回写的 Markdown，804 字节（回写文本是 576 字符）
```

工程入口先建目录，后面全部用相对路径；那颗 1 x 1 的 PNG 是第 1 步源码里 `![公司标识](out/logo.png)` 要读的文件：

```csharp
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Markdown;

Directory.CreateDirectory("out");
File.WriteAllBytes("out/logo.png", Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
```

## 1. 一份样例源文与解析入口

后面每一步都复用同一份源文，把 CommonMark 主体语法与 GFM 扩展（表格、任务列表）都覆盖进去。逐行数组比逐字字符串好核对，"行尾两空格"这种看不见的语法用普通字符串写出来更直观：

```csharp
string source = string.Join("\n", new string[]
{
    "---",
    "title: 报表引擎发布说明",
    "author: YF",
    "date: 2026-10-20",
    "---",
    "",
    "# 报表引擎发布说明",
    "",
    "本段落含 **粗体**、*斜体*、***粗斜体***、~~删除线~~、`行内代码` 与 [官方文档](https://example.com/docs \"悬浮提示\")。",
    "",
    "## 新版要点",
    "",
    "- Excel 原生图表",
    "- 条件格式",
    "  - 数据条",
    "  - 色阶",
    "",
    "1. 解析 Markdown",
    "2. 导出 PDF",
    "",
    "### 任务清单",
    "",
    "- [x] 图表部件写入",
    "- [ ] 图标集对齐值",
    "",
    "```csharp",
    "var doc = Markdown.ToDocument(text);",
    "Markdown.ToHtml(text);",
    "```",
    "",
    "> 第一层引用。",
    ">> 第二层引用。",
    "",
    "| 模块 | 状态 | 进度 |",
    "|:-----|:----:|-----:|",
    "| Word | 已完成 | 100% |",
    "| PPT | 进行中 | 60% |",
    "",
    "![公司标识](out/logo.png)",
    "",
    "---",
    "",
    "结尾段落，含裸标签 <script>alert(1)</script> 与危险链接 [点我](javascript:alert(1))。",
    ""
});

OfficeDocument document = Markdown.ToDocument(source);
Console.WriteLine("[1] 块=" + CountBlocks(document) + " 节=" + document.Sections.Count
    + " 元数据标题=\"" + document.Metadata.Title + "\" 作者=\"" + document.Metadata.Author
    + "\" 创建=" + document.Metadata.Created.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
foreach (Paragraph heading in Headings(document))
{
    Console.WriteLine("    " + heading.HeadingLevel + " 级标题：" + heading.GetText());
}
```

实测输出：

```text
[1] 块=20 节=1 元数据标题="报表引擎发布说明" 作者="YF" 创建=2026-10-20
    1 级标题：报表引擎发布说明
    2 级标题：新版要点
    3 级标题：任务清单
```

四个入口按需选：`Markdown.ToDocument(text)` 与 `ToDocument(text, options)` 吃字符串，`Markdown.LoadFile(path)` 与 `LoadFile(path, options)` 按 UTF-8 读文件并自动去掉起始 BOM，`Markdown.ToHtml(...)` 是"解析 + 渲染"的一次性便捷调用（第 6 步），`Markdown.ToMarkdown(document)` 是反向出口（第 8 步）。整个解析器单遍扫描、不用带嵌套量词的正则，行内分隔符没闭合时按字面文本回退，所以再长的输入也不会指数爆炸。

20 个块是这份源文的"段落 + 表格"总数，`document.Sections` 只有一个节——Markdown 没有分节概念，整篇落进同一节。`Metadata.Created` 来自 front-matter 的 `date: 2026-10-20`；front-matter 只认 `title` / `author` / `date` 三个键，其余键忽略，块本身也不进正文。若没有 front-matter，首个一级标题（ATX 或 Setext 都算）会被抄进 `Metadata.Title`。

教程里反复用到的统计函数放在工程末尾：

```csharp
static int CountBlocks(OfficeDocument doc)
{
    int total = 0;
    foreach (Section section in doc.Sections) total += section.Blocks.Count;
    return total;
}

static List<Paragraph> Paragraphs(OfficeDocument doc)
{
    List<Paragraph> list = new List<Paragraph>();
    foreach (Section section in doc.Sections)
    {
        foreach (Block block in section.Blocks)
        {
            Paragraph paragraph = block as Paragraph;
            if (paragraph != null) list.Add(paragraph);
        }
    }
    return list;
}

static List<Paragraph> Headings(OfficeDocument doc)
{
    List<Paragraph> list = new List<Paragraph>();
    foreach (Paragraph paragraph in Paragraphs(doc))
    {
        if (paragraph.HeadingLevel > 0) list.Add(paragraph);
    }
    return list;
}

static List<Paragraph> ListParagraphs(OfficeDocument doc)
{
    List<Paragraph> list = new List<Paragraph>();
    foreach (Paragraph paragraph in Paragraphs(doc))
    {
        if (paragraph.ListMarker != null) list.Add(paragraph);
    }
    return list;
}

static List<Table> Tables(OfficeDocument doc)
{
    List<Table> list = new List<Table>();
    foreach (Section section in doc.Sections)
    {
        foreach (Block block in section.Blocks)
        {
            Table table = block as Table;
            if (table != null) list.Add(table);
        }
    }
    return list;
}

static int CountOf(string text, string needle)
{
    int count = 0;
    int index = text.IndexOf(needle, StringComparison.Ordinal);
    while (index >= 0)
    {
        count++;
        index = text.IndexOf(needle, index + needle.Length, StringComparison.Ordinal);
    }
    return count;
}

static string Snippet(string text, string needle)
{
    int index = text.IndexOf(needle, StringComparison.Ordinal);
    if (index < 0) return "未出现";
    int start = Math.Max(0, index - 60);
    int length = Math.Min(text.Length - start, 140);
    return "\"" + text.Substring(start, length).Replace("\n", " ") + "\"";
}
```

## 2. 行内格式落成什么

强调、行内代码、链接、图片都不建样式，而是拆成带格式的 `TextRun` / `ImageRun` 放进段落的 `Inlines`：

```csharp
int bold = 0, italic = 0, strike = 0, code = 0, links = 0, images = 0;
string codeFont = null, linkTarget = null, linkText = null;
int imageBytes = 0; double imageW = 0, imageH = 0;
double bodySize = 0;
foreach (Paragraph paragraph in Paragraphs(document))
{
    foreach (Inline inline in paragraph.Inlines)
    {
        TextRun run = inline as TextRun;
        if (run != null)
        {
            if (run.Format != null && run.Format.Bold && !run.Format.Italic) bold++;
            if (run.Format != null && run.Format.Italic && !run.Format.Bold) italic++;
            if (run.Format != null && run.Format.Strikeout) strike++;
            if (run.Format != null && run.Format.FontFamily == "Consolas") { code++; codeFont = run.Format.FontFamily; }
            if (run.Link != null) { links++; linkTarget = run.Link; linkText = run.Text; }
            if (bodySize == 0 && run.Format != null && paragraph.HeadingLevel == 0) bodySize = run.Format.FontSize;
        }
        ImageRun image = inline as ImageRun;
        if (image != null) { images++; imageBytes = image.Image.Data.Length; imageW = image.Width; imageH = image.Height; }
    }
}
Console.WriteLine("[2] 粗体=" + bold + " 斜体=" + italic + " 删除线=" + strike
    + " 行内代码=" + code + "(字体 " + codeFont + ") 链接=" + links + " 图片=" + images
    + " 正文字号=" + bodySize.ToString(CultureInfo.InvariantCulture));
Console.WriteLine("    链接文本=\"" + linkText + "\" 目标=" + linkTarget + " 图片字节=" + imageBytes
    + " 尺寸=" + ((int)imageW) + "x" + ((int)imageH));
```

实测输出（`linkText` / `linkTarget` 保留最后一个带 `Link` 的运行，即那句危险链接）：

```text
[2] 粗体=4 斜体=3 删除线=1 行内代码=3(字体 Consolas) 链接=2 图片=1 正文字号=11
    链接文本="点我" 目标=javascript:alert(1) 图片字节=70 尺寸=0x0
```

对账：粗体 4 = `**粗体**` + 粗斜体 + 表头两个单元格（表头整体加粗）；斜体 3 = `*斜体*` + 粗斜体 + 两个引用段落之一（引用单列在第 5 步）；行内代码 3 = 正文 1 个 + 代码块 2 行（代码块就是连续等宽段落）；链接 2 = 安全链接与 `javascript:` 链接，**模型里都存了目标**，安全过滤发生在渲染阶段（第 6 步）；正文字号 11 磅是 `MarkdownOptions.BaseFontSize` 的默认值。

`尺寸=0x0` 不是丢了：Markdown 没有尺寸语法，解析器把 `ImageRun.Width/Height` 留 0，导出器（PDF / Word / PPT）看到 0 就按 `OfficeImage` 自身的像素与 DPI 回填自然尺寸，图片字节 70 已经进了模型。

## 3. 表格：行数看 `Rows`，对齐看单元格段落

```csharp
List<Table> tables = Tables(document);
Table table = tables[0];
Console.WriteLine("[3] 表格=" + tables.Count + " 个，" + table.Rows.Count + " 行 x "
    + table.Rows[0].Cells.Count + " 列，表头 IsHeader=" + table.Rows[0].IsHeader
    + " 表头底色=" + (table.Rows[0].Cells[0].BackgroundColor == null ? "无" : table.Rows[0].Cells[0].BackgroundColor.Value.ToHex()));
string[] aligns = new string[table.Rows[0].Cells.Count];
for (int c = 0; c < aligns.Length; c++)
{
    Paragraph cell = table.Rows[1].Cells[c].Blocks[0] as Paragraph;
    aligns[c] = cell.Alignment.ToString();
}
Console.WriteLine("    列对齐=" + string.Join(",", aligns) + " 单元格(2,0)=\""
    + ((Paragraph)table.Rows[2].Cells[0].Blocks[0]).GetText() + "\"");
Console.WriteLine("    Columns(列宽定义)=" + table.Columns.Count + " 首列宽=" + table.Columns[0].Width);
```

实测输出：

```text
[3] 表格=1 个，3 行 x 3 列，表头 IsHeader=True 表头底色=D9D9D9
    列对齐=Left,Center,Right 单元格(2,0)="PPT"
    Columns(列宽定义)=3 首列宽=0
```

三点约定：

- 表头行是数据行的一部分，`Rows[0]` 就是它，靠 `IsHeader` 与灰底 `D9D9D9` + 加粗区分；`RepeatHeaderRow` 不设（Markdown 没有跨页重复表头语义）。
- `:---` / `:---:` / `---:` 三种对齐落在**每个单元格段落**的 `Alignment` 上，不是 `TableColumn` 上——`TableColumn` 只有 `Width`。
- Markdown 没有列宽概念，`Columns` 只按列数补了占位、`Width = 0`，导出时按内容自适应。数行列数请用 `Rows[0].Cells.Count`，别拿 `Columns` 当数据。

## 4. 列表、嵌套与任务项

```csharp
Console.WriteLine("[4] 列表段=" + ListParagraphs(document).Count);
foreach (Paragraph item in ListParagraphs(document))
{
    string text = item.GetText();
    string task = item.ListMarker.StartsWith("☑") ? "已完成" : (item.ListMarker.StartsWith("☐") ? "待办" : "-");
    Console.WriteLine("    标记=\"" + item.ListMarker + "\" 缩进=" + item.LeftIndent.ToString("0.#", CultureInfo.InvariantCulture)
        + " 任务=" + task + " 文本=\"" + text + "\"");
}
```

实测输出：

```text
[4] 列表段=8
    标记="•" 缩进=12 任务=- 文本="Excel 原生图表"
    标记="•" 缩进=12 任务=- 文本="条件格式"
    标记="◦" 缩进=36 任务=- 文本="数据条"
    标记="◦" 缩进=36 任务=- 文本="色阶"
    标记="1." 缩进=12 任务=- 文本="解析 Markdown"
    标记="2." 缩进=12 任务=- 文本="导出 PDF"
    标记="☑ •" 缩进=12 任务=已完成 文本="图表部件写入"
    标记="☐ •" 缩进=12 任务=待办 文本="图标集对齐值"
```

- 无序标记按层轮替：第 1 层 `•`、第 2 层 `◦`、第 3 层回到 `•`；有序列表直接把源文的 `1.` / `2.` 存进 `ListMarker`，序号不重算。
- 层级由 `LeftIndent` 表达：第 N 层（从 1 数）是 `12 + (N - 1) * 24` 磅，即 12 / 36 / 60 / 84 / 108。两个写器按同一个"每 24 磅一层"换算回去，五层嵌套也能原样还原（第 8 步实测）。
- **任务项的勾选状态在 `ListMarker` 里，不在正文里**：`- [x]` → 标记 `☑ •`、正文只剩"图表部件写入"。判完成度要读标记前缀，别读 `GetText()`。
- 有序列表与紧邻的无序列表在模型里就是相邻段落，靠 `ListMarker` 是数字还是符号区分。

## 5. 代码块与引用

```csharp
Paragraph codePara = null;
foreach (Paragraph paragraph in Paragraphs(document))
{
    if (codePara == null && paragraph.Inlines.Count > 0 && paragraph.Inlines[0] is TextRun
        && ((TextRun)paragraph.Inlines[0]).Format != null
        && ((TextRun)paragraph.Inlines[0]).Format.FontFamily == "Consolas" && paragraph.ListMarker == null
        && paragraph.GetText().Contains("ToDocument"))
    {
        codePara = paragraph;
    }
}
int codeParas = 0;
foreach (Paragraph paragraph in Paragraphs(document))
{
    if (paragraph.GetText().Contains("ToDocument") || paragraph.GetText().Contains("ToHtml")) codeParas++;
}
Console.WriteLine("[5] 代码行数=" + codeParas + " 代码字体=" + ((TextRun)codePara.Inlines[0]).Format.FontFamily
    + " 字号=" + ((TextRun)codePara.Inlines[0]).Format.FontSize.ToString(CultureInfo.InvariantCulture));
foreach (Paragraph paragraph in Paragraphs(document))
{
    if (paragraph.GetText().Contains("引用") && paragraph.HeadingLevel == 0)
    {
        Console.WriteLine("    引用段 斜体=" + ((TextRun)paragraph.Inlines[0]).Format.Italic
            + " 缩进=" + paragraph.LeftIndent.ToString("0.#", CultureInfo.InvariantCulture)
            + " 段前=" + paragraph.SpaceBefore.ToString("0.#", CultureInfo.InvariantCulture)
            + " 文本=\"" + paragraph.GetText() + "\"");
    }
}
```

实测输出：

```text
[5] 代码行数=2 代码字体=Consolas 字号=10.5
    引用段 斜体=True 缩进=36 段前=2 文本="第一层引用。"
    引用段 斜体=True 缩进=54 段前=2 文本="第二层引用。"
```

围栏代码块（```` ``` ```` 与 `~~~`）和 4 空格缩进块都拆成**一行一段**的等宽段落，字号比正文小 0.5 磅（11 → 10.5）；反引号里的行内代码同字体同字号。识别代码块靠"这一段所有文本运行都是等宽字体"，所以手写文档时把 `TextFormat.FontFamily` 设成 `Consolas`（或选项里的 `CodeFontFamily`）就能让连续段落被当代码块输出。

引用是"斜体 + 递进左缩进"：单层 36 磅、双层 54 磅，段前 2 磅。**模型里没有引用块类型**，也没有承载 `>` 的字段，所以回写 Markdown 与渲染 HTML 时引用退化成普通斜体段落（第 8、10 步实测）。

## 6. 渲染 HTML 片段

```csharp
string html = Markdown.ToHtml(source);
Console.WriteLine("[6] HTML 长度=" + html.Length + " 字符");
Console.WriteLine("    <h1>=" + CountOf(html, "<h1>") + " <h2>=" + CountOf(html, "<h2>")
    + " <h3>=" + CountOf(html, "<h3>") + " <ul>=" + CountOf(html, "<ul>")
    + " <ol>=" + CountOf(html, "<ol>") + " <li>=" + CountOf(html, "<li>")
    + " checkbox=" + CountOf(html, "<input type=\"checkbox\"")
    + " <pre><code>=" + CountOf(html, "<pre><code>") + " <table>=" + CountOf(html, "<table")
    + " <hr />=" + CountOf(html, "<hr />") + " 内联图片=" + CountOf(html, "<img src=\"data:image/png;base64,"));
Console.WriteLine("    安全 href=" + CountOf(html, "<a href=\"https://example.com/docs\"")
    + " javascript href 数=" + CountOf(html, "href=\"javascript")
    + " 裸标签转义=" + (html.Contains("&lt;script&gt;") ? "是" : "否")
    + " title 属性=" + (html.Contains("title=\"悬浮提示\"") ? "有" : "无"));
Console.WriteLine("    危险链接渲染片段=" + Snippet(html, "点我"));
```

实测输出：

```text
[6] HTML 长度=1564 字符
    <h1>=1 <h2>=1 <h3>=1 <ul>=3 <ol>=1 <li>=8 checkbox=2 <pre><code>=1 <table>=1 <hr />=1 内联图片=1
    安全 href=1 javascript href 数=0 裸标签转义=是 title 属性=无
    危险链接渲染片段="lert(1)&lt;/script&gt; 与危险链接 <span style="color:#0000FF"><u>点我</u></span>。</p>"
```

逐项解释：

- `<ul>=3` 是三个无序列表组（要点组、嵌套组、任务清单组），`<ol>=1` 是有序组，`<li>=8` 对上第 4 步的 8 个列表段。
- 任务项转 `<input type="checkbox" disabled="disabled" checked="checked">`，未完成只带 `disabled`。
- 连续等宽段落合并成一个 `<pre><code>`，内容整体转义，所以代码里的 `<` 不会破坏页面。
- 分割线 `---` 转成 `<hr />`——模型里那条居中的"＊ ＊ ＊"段落由写器按同一常量识别回来。
- 图片一律转 `data:image/png;base64,` 内联，产物单文件可离线打开；教程这颗 70 字节的 PNG 渲染后就在 `html` 里。
- **安全策略**：`<a href>` 只接受相对地址、`#锚点` 与 `http` / `https` / `mailto` / `ftp` / `ftps` / `tel`。`javascript:alert(1)` 不进 `href`，改成"蓝色下划线的纯文本 `<span>`"，看着像链接但点不动；`<script>` 这类裸标签在 HTML 阶段永远转义，`EnableRawHtml` 也管不到它。
- 链接的 `"悬浮提示"`（title）被解析但无处安放，模型没有该字段，HTML 里也就没有 `title` 属性。

拼成整页自己包一层即可：

```csharp
string page = "<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">"
    + "<title>" + document.Metadata.Title + "</title></head><body>"
    + html + "</body></html>";
```

## 7. 八个开关逐个实测

```csharp
MarkdownOptions noTables = new MarkdownOptions();
noTables.EnableTables = false;
OfficeDocument noTablesDoc = Markdown.ToDocument(source, noTables);
Console.WriteLine("[7] EnableTables=false：块=" + CountBlocks(noTablesDoc) + " 表格=" + Tables(noTablesDoc).Count
    + " HTML <table>=" + CountOf(Markdown.ToHtml(source, noTables), "<table"));

MarkdownOptions noTasks = new MarkdownOptions();
noTasks.EnableTaskLists = false;
OfficeDocument noTasksDoc = Markdown.ToDocument(source, noTasks);
string firstTask = "";
foreach (Paragraph item in ListParagraphs(noTasksDoc))
{
    if (item.GetText().Contains("图标集")) firstTask = item.GetText();
}
Console.WriteLine("    EnableTaskLists=false：任务项文本=\"" + firstTask + "\"");

MarkdownOptions hardBreaks = new MarkdownOptions();
hardBreaks.HardLineBreaks = true;
string breakMd = "第一行\n第二行\n带两空格  \n下一行";
Console.WriteLine("    HardLineBreaks=false <br />=" + CountOf(Markdown.ToHtml(breakMd), "<br />")
    + "，=true <br />=" + CountOf(Markdown.ToHtml(breakMd, hardBreaks), "<br />")
    + "；默认源码里的行尾两空格 <br />=" + CountOf(html, "<br />"));

MarkdownOptions fonts = new MarkdownOptions();
fonts.BaseFontSize = 14.0;
fonts.CodeFontFamily = "Courier New";
OfficeDocument fontDoc = Markdown.ToDocument(source, fonts);
string codeFont2 = null;
double bodySize2 = 0;
foreach (Paragraph paragraph in Paragraphs(fontDoc))
{
    foreach (Inline inline in paragraph.Inlines)
    {
        TextRun run = inline as TextRun;
        if (run == null || run.Format == null) continue;
        if (run.Format.FontFamily == "Courier New") codeFont2 = run.Format.FontFamily;
        if (bodySize2 == 0 && paragraph.HeadingLevel == 0 && paragraph.ListMarker == null) bodySize2 = run.Format.FontSize;
    }
}
Console.WriteLine("    BaseFontSize=14 + CodeFontFamily=Courier New：正文字号=" + bodySize2.ToString(CultureInfo.InvariantCulture)
    + " 代码字体=" + codeFont2);

MarkdownOptions rawHtml = new MarkdownOptions();
rawHtml.EnableRawHtml = true;
OfficeDocument styled = new OfficeDocument();
Section styledSection = styled.AddSection();
Paragraph styledPara = styledSection.AddParagraph("下划线与颜色");
styledPara.Inlines.Clear();
styledPara.AddText("红色下划线", new TextFormat { Bold = false, Underline = true, Color = new OfficeColor(0xC0, 0x00, 0x00) });
Console.WriteLine("    EnableRawHtml=false 回写=" + Markdown.ToMarkdown(styled));
Console.WriteLine("    EnableRawHtml=true  回写=" + Markdown.ToMarkdown(styled, rawHtml));
```

实测输出：

```text
[7] EnableTables=false：块=20 表格=0 HTML <table>=0
    EnableTaskLists=false：任务项文本="[ ] 图标集对齐值"
    HardLineBreaks=false <br />=1，=true <br />=3；默认源码里的行尾两空格 <br />=0
    BaseFontSize=14 + CodeFontFamily=Courier New：正文字号=14 代码字体=Courier New
    EnableRawHtml=false 回写=红色下划线
    EnableRawHtml=true  回写=<span style="color:#C00000"><u>红色下划线</u></span>
```

`EnableRawHtml` 那两行值得多说一句：段落正文是"下划线与颜色"，但 `Inlines.Clear()` 之后只剩那个带下划线与颜色的运行——默认选项下 Markdown 表达不了下划线和颜色，直接丢；开 `EnableRawHtml` 才以 `<u>` / `<span style>` 落进文本。

| 选项 | 默认 | 实测效果 |
| --- | --- | --- |
| `EnableTables` | `true` | 关掉后管道表整段当普通文本，块数不变（四行并成一段）、`Table` 数与 `<table>` 数都归 0 |
| `EnableTaskLists` | `true` | 关掉后 `[ ]` / `[x]` 不识别，复选框文本原样留在正文里 |
| `HardLineBreaks` | `false` | 默认只有"行尾两空格"产 `<br />`（四行样例 1 个）；开成 `true` 后每个换行都是 `<br />`（3 个） |
| `BaseFontSize` | `11.0` | 正文字号跟着变（14），代码仍比正文小 0.5 磅 |
| `CodeFontFamily` | `"Consolas"` | 代码块与行内代码换字体；等宽识别也认这个名，换成 `"Courier New"` 后原来的 Consolas 段落才算代码 |
| `EnableRawHtml` | `false` | **只影响回写 Markdown**：`true` 时下划线与颜色用 `<u>` / `<span style="color:#…">` 保留；HTML 渲染阶段一律转义 |
| `TrimTrailingWhitespace` | `true` | 去掉行尾空白，硬换行用的两个尾部空格除外 |
| `BodyFontFamily` | `null` | 给正文运行写 `FontFamily`；`null` 沿用文档默认字体 |

`options.Clone()` 复制一份再改，避免共用实例互相影响。

## 8. 文档模型 → Markdown

回写不要求文档来自 Markdown——手写模型、Word 读来的模型都行。模型侧的约定：`HeadingLevel` 出 `#`、`ListMarker` 出 `- ` / `1.`、任务状态看标记里的 `☑` / `☐`、连续等宽段落出围栏、`Table` 出管道表。

```csharp
OfficeDocument release = new OfficeDocument();
release.Metadata.Title = "发布说明";
Section releaseSection = release.AddSection();
Paragraph releaseHeading = releaseSection.AddParagraph("新版本要点");
releaseHeading.HeadingLevel = 2;
Paragraph releaseItem = releaseSection.AddParagraph("支持 Excel 原生图表");
releaseItem.ListMarker = "-";
Paragraph releaseDone = releaseSection.AddParagraph("PPTX 写出");
releaseDone.ListMarker = "-";
Table releaseTable = new Table();
releaseTable.AddRow("模块", "状态");
releaseTable.AddRow("Word", "已完成");
releaseSection.Blocks.Add(releaseTable);

string written = Markdown.ToMarkdown(release);
Console.WriteLine("[8] 手写文档 -> Markdown 长度=" + written.Length);
foreach (string line in written.Split('\n')) Console.WriteLine("    | " + line);
```

实测输出：

```text
[8] 手写文档 -> Markdown 长度=87
    | # 发布说明
    |
    | ## 新版本要点
    |
    | - 支持 Excel 原生图表
    | - PPTX 写出
    |
    | | 模块 | 状态 |
    | | :--- | :--- |
    | | Word | 已完成 |
```

`Metadata.Title`（"发布说明"）与首个标题（"新版本要点"）不同，写器在开头补了一行 `# 发布说明`；若首个标题已经就是它，则不重复补。表头行不再包 `**`——管道表头本身就是强调。

往返保真度实测：

```csharp
OfficeDocument roundTripA = Markdown.ToDocument(source);
string again = Markdown.ToMarkdown(roundTripA);
OfficeDocument roundTripB = Markdown.ToDocument(again);
Console.WriteLine("    往返：一次解析块=" + CountBlocks(roundTripA) + " 回写文本长度=" + again.Length
    + " 二次解析块=" + CountBlocks(roundTripB) + " 表格=" + Tables(roundTripB).Count
    + " 标题=" + Headings(roundTripB).Count);
Console.WriteLine("    表头对齐往返=" + DelimiterLine(again));
```

```text
    往返：一次解析块=20 回写文本长度=576 二次解析块=20 表格=1 标题=3
    表头对齐往返=| :--- | :---: | ---: |
```

`DelimiterLine` 只是把回写文本里的表格分隔行挑出来：

```csharp
static string DelimiterLine(string markdown)
{
    foreach (string line in markdown.Split('\n'))
    {
        if (line.Contains(":---") || line.Contains("-:") || line.Contains("| ---") || line.StartsWith("| :")) return line.Trim();
    }
    return "无对齐分隔行";
}
```

回写后的 576 字符再解析，块数 20 / 表格 1 / 标题 3 全部对上。`again` 全文实测如下：

```text
# 报表引擎发布说明

本段落含 **粗体**、*斜体*、***粗斜体***、~~删除线~~、`行内代码` 与 [官方文档](https://example.com/docs)。

## 新版要点

- Excel 原生图表
- 条件格式
  - 数据条
  - 色阶
1. 解析 Markdown
2. 导出 PDF

### 任务清单

- [x] 图表部件写入
- [ ] 图标集对齐值

```
var doc = Markdown.ToDocument(text);
Markdown.ToHtml(text);
```

*第一层引用。*

*第二层引用。*

| 模块 | 状态 | 进度 |
| :--- | :---: | ---: |
| Word | 已完成 | 100% |
| PPT | 进行中 | 60% |

![image](data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==)

---

结尾段落，含裸标签 \<script>alert(1)\</script> 与危险链接 [点我](javascript:alert(1))。
```

丢了什么、留了什么，一目了然：

- **留下**：标题层级、四种强调、行内代码、链接（含 URL）、嵌套列表层级、任务项 `[x]` / `[ ]`、代码围栏、表格三列对齐（`:---` / `:---:` / `---:`）、分割线 `---`、转义后的裸标签。
- **丢失**：围栏的语言标记（`csharp` 模型里没字段）、链接 title、front-matter（不回写）、引用的 `>`（只剩斜体）、图片的本地路径与 alt（模型不带 alt，统一写 `image`，路径变成 data URI）。
- 图片写成 data URI 后，再解析时按"不存在的本地路径"降级成 alt 文本，字节不再入库——想让图片继续存活就别做 md → md 往返，直接保源文。

深层嵌套也实测过（五层 + 分割线一次往返）：

```csharp
string deep = string.Join("\n", new string[]
{
    "- 第 1 层",
    "  - 第 2 层",
    "    - 第 3 层",
    "      - 第 4 层",
    "        - 第 5 层",
    "",
    "***",
    "",
    "正文"
});
OfficeDocument deepDoc = Markdown.ToDocument(deep);
Console.WriteLine("[11] 五层嵌套缩进=" + Indents(deepDoc));
string deepHtml = Markdown.ToHtml(deep);
Console.WriteLine("    HTML <ul>=" + CountOf(deepHtml, "<ul>") + " 分隔线 <hr />=" + CountOf(deepHtml, "<hr />"));
string deepMd = Markdown.ToMarkdown(deepDoc);
Console.WriteLine("    回写：");
foreach (string line in deepMd.Split('\n')) Console.WriteLine("    | " + line);
Console.WriteLine("    回写 <hr> 分隔线还原=" + (deepMd.Contains("---") || deepMd.Contains("***") ? "是" : "否"));
```

```text
[11] 五层嵌套缩进=12,36,60,84,108
    HTML <ul>=5 分隔线 <hr />=1
    回写：
    | - 第 1 层
    |   - 第 2 层
    |     - 第 3 层
    |       - 第 4 层
    |         - 第 5 层
    |
    | ---
    |
    | 正文
    回写 <hr> 分隔线还原=是
```

`Indents` 把每个列表段的左缩进拼成一行：

```csharp
static string Indents(OfficeDocument doc)
{
    List<string> parts = new List<string>();
    foreach (Paragraph item in ListParagraphs(doc))
    {
        parts.Add(item.LeftIndent.ToString("0", CultureInfo.InvariantCulture));
    }
    return string.Join(",", parts);
}
```

`***` 与 `---` 都能识别成分割线，回写统一输出 `---`。

## 9. 经 `Wisdom.Office` 门面读写

门面按扩展名识别 `.md` / `.markdown` / `.mkd`，之后 Markdown 就是"又一个可解析、可再导出的文档"：

```csharp
string mdPath = "out/article.md";
File.WriteAllText(mdPath, source, new UTF8Encoding(false));
OfficeFile loaded = OfficeFile.Load(mdPath);
Console.WriteLine("[9] 识别=" + OfficeFileFormats.GetDisplayName(loaded.Format) + " 块=" + CountBlocks(loaded.AsDocument())
    + " 纯文本长度=" + loaded.ExtractText().Length);
loaded.Save("out/article.pdf");
OfficeConverter.Convert(loaded, "out/article.docx");
OfficeConverter.ToMarkdown(loaded.AsDocument(), "out/article-back.md");
Console.WriteLine("    md -> PDF=" + new FileInfo("out/article.pdf").Length + " 字节"
    + "，md -> DOCX=" + new FileInfo("out/article.docx").Length + " 字节"
    + "，md -> md=" + new FileInfo("out/article-back.md").Length + " 字节");
OfficeFile fromDocx = OfficeFile.Load("out/article.docx");
string docxToMd = OfficeConverter.ToMarkdown(fromDocx.AsDocument());
Console.WriteLine("    docx 回转 Markdown 长度=" + docxToMd.Length
    + " 链接 URL 保留=" + (docxToMd.Contains("example.com") ? "是" : "否(退化纯文本)")
    + " 图片保留=" + (docxToMd.Contains("![") ? "是" : "否"));
OfficeFile mdFromDocx = OfficeFile.Load("out/article-back.md");
Console.WriteLine("    md -> md -> 再解析块=" + CountBlocks(mdFromDocx.AsDocument()));
```

实测输出：

```text
[9] 识别=Markdown 文档 块=20 纯文本长度=298
    md -> PDF=121895 字节，md -> DOCX=5004 字节，md -> md=804 字节
    docx 回转 Markdown 长度=525 链接 URL 保留=否(退化纯文本) 图片保留=是
    md -> md -> 再解析块=20
```

`md -> md` 的 804 是 UTF-8 字节数，第 8 步的 576 是字符数——同一串文本，中文每字 3 字节。`file.Save("out/article.pdf")` 与 `OfficeConverter.Convert(file, path)` 等价，前者按扩展名推断格式，后者还能带 Excel 排版参数。DOCX 再转回 Markdown 时链接只剩文字（`TextRun.Link` 只在 Markdown 侧填充，Word 里的超链接读入模型不入库 URL），这是门面的已知限制，不是 Markdown 引擎的。

顺带实测分割线在 Word / PDF 侧的样子：

```csharp
OfficeFile pdfBack = OfficeFile.Load("out/article.pdf");
OfficeFile docxBack = OfficeFile.Load("out/article.docx");
Console.WriteLine("[14] PDF 文本含 ＊ ＊ ＊=" + pdfBack.ExtractText().Contains("＊ ＊ ＊")
    + " DOCX 含=" + docxBack.ExtractText().Contains("＊ ＊ ＊")
    + " PDF 页数=" + pdfBack.Pdf.PageCount);
```

```text
[14] PDF 文本含 ＊ ＊ ＊=True DOCX 含=True PDF 页数=1
```

段落模型没有"下边框"属性，所以分割线在文档模型里是一条居中、写着 `＊ ＊ ＊` 的段落，Word / PDF 就排这行字符；只有 HTML 与 Markdown 两条出口会把它还原成 `<hr />` 与 `---`。

## 10. 降级行为实测：哪些写法会被改写、哪些留字面

```csharp
string limits = string.Join("\n", new string[]
{
    "引用式链接 [文档][ref] 与脚注[^1]。",
    "",
    "[^1]: 脚注内容",
    "",
    "- 列表首段",
    "  续写第二段落",
    "",
    "![远程](https://example.com/logo.png)",
    "",
    "| 单元格 | 内容 |",
    "|---|---|",
    "| - 内嵌列表 | `代码` |",
    "",
    "[文档]: https://example.com/a"
});
OfficeDocument limitDoc = Markdown.ToDocument(limits);
string limitHtml = Markdown.ToHtml(limits);
Console.WriteLine("[10] 块=" + CountBlocks(limitDoc)
    + " 首块=" + limitDoc.Sections[0].Blocks[0].GetType().Name
    + " <a>=" + CountOf(limitHtml, "<a href") + " HTML 长度=" + limitHtml.Length);
Console.WriteLine("    首段=\"" + ((Paragraph)limitDoc.Sections[0].Blocks[0]).GetText()
    + "\" 末段=\"" + LastText(limitDoc) + "\"");
Console.WriteLine("    表格单元格文本=\"" + CellText(Tables(limitDoc)[0], 1, 0)
    + "\" 列表项段数=" + ListParagraphs(limitDoc).Count);
foreach (Block block in limitDoc.Sections[0].Blocks)
{
    Paragraph p = block as Paragraph;
    Console.WriteLine("    - " + block.GetType().Name + " \"" + (p == null ? "" : p.GetText()) + "\" 标记="
        + (p == null ? "-" : (p.ListMarker ?? "无")));
}
```

实测输出：

```text
[10] 块=4 首块=Paragraph <a>=1 HTML 长度=365
    首段="引用式链接 [文档][ref] 与脚注^1。" 末段="远程"
    表格单元格文本="- 内嵌列表" 列表项段数=1
    - Paragraph "引用式链接 [文档][ref] 与脚注^1。" 标记=无
    - Paragraph "列表首段 续写第二段落" 标记=•
    - Paragraph "远程" 标记=无
    - Table "" 标记=-
```

两个辅助函数只取文本，注意 `TableCell` 没有 `GetText()`，单元格内容在 `Blocks` 里：

```csharp
static string LastText(OfficeDocument doc)
{
    List<Paragraph> list = Paragraphs(doc);
    return list[list.Count - 1].GetText();
}

static string CellText(Table table, int row, int column)
{
    Paragraph paragraph = table.Rows[row].Cells[column].Blocks[0] as Paragraph;
    return paragraph == null ? string.Empty : paragraph.GetText();
}
```

对应包说明里的「已知限制」，全部是不抛异常、只降级：

- **引用式链接要标签对上才生效**：`[文档][ref]` 找的是标签 `ref` 的定义，而这段里定义的是 `[文档]: …`，标签不匹配，整个结构按字面文本留着；对上了（如 `[a][1]` 配 `[1]: /url`）就渲染成 `<a>`，定义行本身不再占一个块（实测块数从 7 降到 4）。未定义的引用不抛异常，也不凭空造链接。
- **脚注 `[^1]` 不是脚注**：引用式的标签规则会把 `^1` 当成一个普通引用名，配套的 `[^1]: 脚注内容` 被收进定义表，实测渲染成 `<a href="脚注内容">^1</a>`——不是上标脚注，模型里那段文本也变成"引用式链接 [文档][ref] 与脚注^1。"（方括号被吃掉）。写文档时别把脚注定义行混进来。
- **列表项内的缩进续行并进同一个项**：`- 列表首段` 后面紧跟两空格缩进的"续写第二段落"，实测成一个 `<li>列表首段 续写第二段落</li>`、`列表项段数=1`，两段文字合成一段（回写 Markdown 也合成一行）。模型里列表项就是段落序列，没有"松散列表"这一层。
- **远程图片** `![远程](https://…)` 不下载，降级成 alt 文本"远程"；本地路径找不到同理；引用式图片 `![alt][pic]` 也只输出 alt 文本。
- **表格单元格内不解析块级语法**：`| - 内嵌列表 |` 里的 `- ` 只是文字（单元格文本实测就是 `- 内嵌列表`），行内 `` `代码` `` 之类仍按行内处理。
- **HTML 出口永远转义，原始 HTML 永不透传**：`EnableRawHtml` 只决定回写 Markdown 时用不用 `<u>` / `<span>` 表达下划线与颜色，对 HTML 渲染没有任何影响。

```csharp
MarkdownOptions rawOn = new MarkdownOptions();
rawOn.EnableRawHtml = true;
Console.WriteLine("    EnableRawHtml=true 的 HTML=" + Markdown.ToHtml("<b>粗</b> 与 [链接](https://example.com)", rawOn));
Console.WriteLine("    同输入默认选项 HTML=" + Markdown.ToHtml("<b>粗</b> 与 [链接](https://example.com)"));
Console.WriteLine("    <div> 独占一行=" + Markdown.ToHtml("<div>"));
Console.WriteLine("    <script> 一行=" + Markdown.ToHtml("<script>alert(1)</script>"));
```

```text
    EnableRawHtml=true 的 HTML=<p>&lt;b&gt;粗&lt;/b&gt; 与 <a href="https://example.com"><span style="color:#0000FF"><u>链接</u></span></a></p>
    同输入默认选项 HTML=<p>&lt;b&gt;粗&lt;/b&gt; 与 <a href="https://example.com"><span style="color:#0000FF"><u>链接</u></span></a></p>
    <div> 独占一行=<p>&lt;div&gt;</p>
    <script> 一行=<p>&lt;script&gt;alert(1)&lt;/script&gt;</p>
```

标签被转义成文本，同一行的行内 Markdown 照常解析（链接仍然是 `<a>`）；`<div>` 单独成行也只是变成一段转义文本，块里那一行仍按 Markdown 走——没有 HTML 块容器的概念。

## 11. 字符实体与资源上限

数字形式 `&#65;` / `&#x42;` 与 73 个常用命名字（`&amp;`、`&nbsp;`、`&copy;`、`&mdash;` 等）在解析阶段就解成字符，各写出口再按需转义；匹配不上的形式按字面保留，与 CommonMark 一致。

```csharp
string md = "文本 `&#65;` 与 &#65;&#x42;&copy;&nbsp;&mdash; 以及 &copy 和 &#0;";
Console.WriteLine(Markdown.ToDocument(md).GetPlainText().Trim());
Console.WriteLine(Markdown.ToHtml(md));
```

```text
文本 &#65; 与 AB© — 以及 &copy 和 &#0;
<p>文本 <code>&amp;#65;</code> 与 AB© — 以及 &amp;copy 和 &amp;#0;</p>
```

要点：等宽片段里的 `` `&#65;` `` 不解码（HTML 里保持 `&amp;#65;`）；`&copy` 缺分号、`&#0;` 码点为 0（同样非法的还有落在代理区、超过 U+10FFFF 的码点）都按字面留着；正文里的 `&#65;&#x42;&copy;` 解成 `AB©`，紧跟的 `&nbsp;` 是不换行空格 U+00A0（控制台看着像普通空格），`&mdash;` 解成长破折号。

解析不受信任的 Markdown 时有资源上限，超限抛 `InvalidDataException` 而不是无界增长：单张表格 4 096 列、累计 65 536 个单元格，单篇文档嵌入图片合计 256 MB。

```csharp
try { Markdown.ToDocument(new string('|', 4096) + "\n" + new string('-', 4096) + "|"); }
catch (Exception ex) { Console.WriteLine("列数：" + ex.GetType().Name + ": " + ex.Message); }

StringBuilder row = new StringBuilder();
StringBuilder sep = new StringBuilder();
for (int i = 0; i < 400; i++) { row.Append("| c" + i + " "); sep.Append("| --- "); }
row.Append('|'); sep.Append('|');
StringBuilder many = new StringBuilder();
many.Append(row).Append('\n').Append(sep).Append('\n');
for (int r = 0; r < 200; r++) many.Append(row).Append('\n');
try { Markdown.ToDocument(many.ToString()); }
catch (Exception ex) { Console.WriteLine("单元格：" + ex.GetType().Name + ": " + ex.Message); }
```

```text
列数：InvalidDataException: The Markdown table exceeds the 4096-column limit.
单元格：InvalidDataException: The Markdown table exceeds the 65536-cell limit.
```

合法边界不受影响：4 095 列的表照常解析，400 列 × 160 个数据行（连表头 161 行、64 400 个单元格）也在限内。同一张本地图片被引用几百次时按归一化绝对路径只读一次、只存一份，累计字节也只计一次；单张图另受 `MarkdownOptions.MaxImageBytes`（默认 16 MB）约束，超限与网络图片一样降级成 alt 文本。

## 12. 与 CommonMark 对齐的行内与块级细节

这一轮把引用式链接、标题锚点、列表起始编号、强调贴合规则、行内代码与块级优先级按 CommonMark 补齐。下面每组都是 `Markdown.ToHtml` 直接打印的实测结果。

### 引用式链接与定义行

```csharp
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "[a][1]", "", "[1]: /url")));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "[1]", "", "[1]: /url")));          // 简写式
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "[a][]", "", "[a]: /u2")));         // 折叠式
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "[a][REF]", "", "[ref]: /url3")));  // 标签不区分大小写
Console.WriteLine(Markdown.ToHtml("[a][zz]"));                                          // 引用没定义
Console.WriteLine(Markdown.ToMarkdown(Markdown.ToDocument(string.Join("\n", "[a][1]", "", "[1]: /url"))));
```

```text
<p><a href="/url"><span style="color:#0000FF"><u>a</u></span></a></p>
<p><a href="/url"><span style="color:#0000FF"><u>1</u></span></a></p>
<p><a href="/u2"><span style="color:#0000FF"><u>a</u></span></a></p>
<p><a href="/url3"><span style="color:#0000FF"><u>a</u></span></a></p>
<p>[a][zz]</p>
[a](/url)
```

定义行本身不再渲染成正文（第 10 节实测块数因此从 7 降到 4）；定义里带的 `"标题"` 只识别、不输出，模型没有放 title 的字段。四种形态里只有**标签对上**才成链接，对不上就整串留字面，绝不凭空造地址。引用式图片 `![alt][pic]` 只输出 alt 文本；定义行的地址是 `javascript:` 一类时同样拒绝进 `href`，降级成不可点击的蓝色下划线文本。回写 Markdown 时引用式统一落成内联式 `[a](/url)`。

### 标题锚点

```csharp
Console.WriteLine(Markdown.ToHtml("# Hello World"));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "# A", "", "# A")));
Console.WriteLine(Markdown.ToHtml("# C# & *fancy* name"));
Console.WriteLine(Markdown.ToHtml("# 中文标题"));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "line1", "line2", "===")));
```

```text
<h1 id="hello-world"><span style="color:#1F3A5F"><strong>Hello World</strong></span></h1>
<h1 id="a"><span style="color:#1F3A5F"><strong>A</strong></span></h1>
<h1 id="a-1"><span style="color:#1F3A5F"><strong>A</strong></span></h1>
<h1 id="c-fancy-name"><span style="color:#1F3A5F"><strong>C# &amp; <em>fancy</em> name</strong></span></h1>
<h1 id="中文标题"><span style="color:#1F3A5F"><strong>中文标题</strong></span></h1>
<h1 id="line1-line2"><span style="color:#1F3A5F"><strong>line1 line2</strong></span></h1>
```

`id` 取标题**正文**（强调标记不算字符）转小写、非字母数字并成单个 `-`、首尾的 `-` 去掉；同名标题依次追加 `-1` / `-2`。中日韩标题按原文取锚点——`id="中文标题"` 是能直接跳转的 id，把非 ASCII 一律压成 `section` 反而没法用。Setext 写法（`===` / `---` 压在下文上）同样生成锚点，且整段并成一个标题。锚点只在 HTML 出口出现，文档模型里没有这个字段。

### 有序列表起始编号

```csharp
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "5. five", "6. six")));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "0. zero", "1. one")));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "1. one", "2. two")));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "3) three", "4) four")));
```

```text
<ol start="5">
<li>five</li>
<li>six</li>
</ol>
<ol start="0">
<li>zero</li>
<li>one</li>
</ol>
<ol>
<li>one</li>
<li>two</li>
</ol>
<ol start="3">
<li>three</li>
<li>four</li>
</ol>
```

首项不是 1 才写 `start`（`0` 也照写），首项是 1 时不写这个属性；`1)` 这种括号分隔符一样识别。无序列表从不带 `start`。

### 强调的贴合与相邻分隔符串

```csharp
Console.WriteLine(Markdown.ToHtml("in_var_iable"));
Console.WriteLine(Markdown.ToHtml("2*3*4"));
Console.WriteLine(Markdown.ToHtml("foo * bar*"));
Console.WriteLine(Markdown.ToHtml("*a**b*"));
Console.WriteLine(Markdown.ToHtml("**a *b* c**"));
Console.WriteLine(Markdown.ToHtml("**a***b***"));
Console.WriteLine(Markdown.ToHtml("**bold *italic* bold**"));
```

```text
<p>in_var_iable</p>
<p>2<em>3</em>4</p>
<p>foo * bar*</p>
<p><em>a**b</em></p>
<p><strong>a <em>b</em> c</strong></p>
<p><strong>a</strong><em>b</em>**</p>
<p><strong>bold <em>italic</em> bold</strong></p>
```

词内的下划线不成强调（`in_var_iable` 原样），星号允许词内成对（`2*3*4` 出 `<em>3</em>`）；起始串后面必须紧跟内容，所以 `foo * bar*` 只是文本。相邻的分隔符串按 CommonMark 配对：单星不能抢既能开场又能闭合的双星（`*a**b*` 里那两个星是正文），内层成对的星号先拿走自己的闭合位、外层不越过去抢，最后**剥掉标签剩下的字符与原文一致**——星号只会变成标签，不会凭空消失。

### 行内代码与围栏代码的空白

```csharp
Console.WriteLine(Markdown.ToHtml("` a `"));
Console.WriteLine(Markdown.ToHtml("`a  b`"));
Console.WriteLine(Markdown.ToHtml("`` a ` b ``"));
Console.WriteLine(Markdown.ToHtml("`&#65;`"));
Console.WriteLine(Markdown.ToHtml("` `"));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "```", "int x;  ", "```")));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "```", "a", "", "b", "```")));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "```", "```")));
```

```text
<p><code>a</code></p>
<p><code>a  b</code></p>
<p><code>a ` b</code></p>
<p><code>&amp;#65;</code></p>
<p><code> </code></p>
<pre><code>int x;  
</code></pre>
<pre><code>a

b
</code></pre>
<pre><code></code></pre>
```

闭合串必须与起始串**等长**（` `` a ` b `` ` 里的单星号是内容），两侧各恰好一个空格时各去掉一格、其余空白一律保留（`` `a  b` `` 的双空格、`" "` 这种纯空格片段都原样出）；串内不做实体解码，`` `&#65;` `` 就是 `&#65;`——这正是反引号的用途。围栏代码的空白同样保真：行尾空格不裁、块内空行不会把它拆成两块、块尾留一个换行，回写成 Markdown 再解析结果不变；空围栏输出 `<pre><code></code></pre>` 而不是消失。

### 块级优先级与链接目标

```csharp
Console.WriteLine(Markdown.ToHtml("* * *"));
Console.WriteLine(Markdown.ToHtml("- - -"));
Console.WriteLine(Markdown.ToHtml("    - not a list"));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "para", "    still para")));
Console.WriteLine(Markdown.ToHtml("[a](/url with space)"));
Console.WriteLine(Markdown.ToHtml("[a](</url with space>)"));
Console.WriteLine(Markdown.ToHtml("[a](/u(te)f)"));
Console.WriteLine(Markdown.ToHtml("[a](/u \"the title\")"));
Console.WriteLine(Markdown.ToHtml(string.Join("\n", "line1  ", "line2")));
```

```text
<hr />
<hr />
<pre><code>- not a list
</code></pre>
<p>para still para</p>
<p>[a](/url with space)</p>
<p><a href="/url with space"><span style="color:#0000FF"><u>a</u></span></a></p>
<p><a href="/u(te)f"><span style="color:#0000FF"><u>a</u></span></a></p>
<p><a href="/u"><span style="color:#0000FF"><u>a</u></span></a></p>
<p>line1<br />
line2</p>
```

`* * *` 与 `- - -` 先判成分隔线，不会被误当成三个列表项；4 空格缩进的行先判成缩进代码（`    - not a list` 是代码不是列表），但缩进行**不能打断段落**——段落后面缩进一行仍是同一段。链接目标里出现未转义空格时整个结构退回字面文本，用 `<…>` 包裹就可以带空格，成对括号可直接写进目标，`"title"` 认出来但不输出。硬换行两种写法都认：行尾两个空格、行尾反斜杠，`<br />` 之后还跟着一个真正的换行；`HardLineBreaks = true` 时每个换行都是硬换行。

## 常见问题

**`table.Rows[0].Cells.Count` 才是列数。** `Table.Columns` 是列宽定义，Markdown 表格没有宽度概念，解析器只按列数补了 `Width = 0` 的占位。手写文档时也一样：`AddRow(params string[])` 填的是 `Rows[i].Cells`。

**任务项读不到勾选状态。** 状态在 `ListMarker`（`☑ •` / `☐ •`），不在正文；正文只剩条目文字。判状态用 `marker.StartsWith("☑")`。

**嵌套列表回写后被压平。** 层级靠 `LeftIndent`，写器按"每 24 磅一层"换算，所以自造文档要给第 N 层设 `12 + (N - 1) * 24` 磅；用别的磅值（比如每层 18）会在第 3 层以后被算成同一层。

**分割线在 Word 里是一串 ＊＊＊。** 模型没有分隔块，解析器用居中段落承载 `＊ ＊ ＊`，HTML / Markdown 出口识别回 `<hr />` / `---`，Word / PDF 就照排这几个字符（第 9 步实测）。

**代码块围栏没了语言。** `csharp` 这类信息模型里没字段，回写只输出空围栏；HTML 渲染也不加 `class`。

**`<span>`、`<u>` 只出现在回写的 Markdown 里。** `EnableRawHtml = true` 只影响 `ToMarkdown`；`ToHtml` 阶段所有文本一律转义，`<script>` 永远进不了输出。

**危险链接还是蓝色的。** 非白名单协议（`javascript:`、`data:` 等）不进 `href`，渲染成蓝 + 下划线的纯文本 `<span>`——视觉上像链接但没有跳转。白名单是相对地址、`#锚点`、`http`、`https`、`mailto`、`ftp`、`ftps`、`tel`。

**图片只支持 PNG / JPEG。** `OfficeImage` 按签名识别，`.gif`、`.svg`、`.emf` 不解析；网络地址与缺失文件降级成 alt 文本，不抛异常。

**要不要装 Office。** 不需要。解析是纯 C# 单遍扫描，HTML / Markdown 都是字符串产物；经门面转 PDF / Word 也走本仓库自研写器，全程零第三方依赖。
