# Wisdom.Word 使用教程

Word 侧的实战教程：生成并回读一份 DOCX，做模板填充，把格式混乱的文档批量整理成统一规范，再校验一致性、比对两版差异、读取旧版 `.doc`。安装方式见[包说明](../src/Wisdom.Word/README.md#安装)。

教程自带输入文件（第 1 步就用 `DocxWriter` 造出后面要处理的文档），照抄即可跑通。文中所有输出都是实测结果。

## 你将完成

```text
out/
  draft.docx                 第 1 步生成的"脏"文档（空段、字号混用、假标题、标题跳级、表格不齐）
  draft.backup-*.docx        第 5 步流水线自动创建的备份
  contract-template.docx     第 3 步生成的模板
  contract.docx              模板填充结果
  clean.docx                 整理后的文档
  clean.docx.report.txt      整理报告（变更清单 + 问题清单 + 前后统计）
  legacy.docx                由 .doc 转换而来
  tagged.docx                带自定义属性的文档
  labels.docx                标签纸
  student-labels.docx        带填写线的学生信息标签纸
  retention-labels.docx      A4 4×10 分格留样标签
  chemical-safety.docx       A5 化学品安全标签（含 PDF）
  envelope.docx              信封
  bookmarks.docx             第 13 步的书签文档（整段 + 范围）
  bookmarks.odt              同一文档转 ODT，范围书签保留
```

## 1. 生成一份待整理的草稿

故意留下常见问题：多余空段、用"加粗大字"冒充标题、标题层级从 1 直接跳到 4、表格没居中也没设重复表头。

```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Wisdom.Office.Core;
using Wisdom.Word;

OfficeDocument draft = new OfficeDocument();
draft.Metadata.Title = "项目实施方案（草稿）";

Section section = draft.AddSection(PageSetup.A4());

Paragraph title = new Paragraph();
title.Alignment = ParagraphAlignment.Center;
title.AddText("项目实施方案", new TextFormat("SimHei", 22) { Bold = true });   // 假标题：只是加粗大字
section.Blocks.Add(title);

section.Blocks.Add(new Paragraph(""));                    // 多余空段
section.Blocks.Add(new Paragraph(""));

Paragraph pseudo = new Paragraph();
pseudo.AddText("一、实施范围", new TextFormat("SimHei", 16) { Bold = true });   // 又一个假标题
section.Blocks.Add(pseudo);

section.AddParagraph("本项目覆盖 3 个楼栋的公区与地下车库，含 墙面涂料 、门窗 五金更换，以及 机电末端 检修。");
section.AddParagraph("工期： 90 日历天");
section.AddParagraph("");

Paragraph levelSkip = new Paragraph("（三）质量目标");
levelSkip.HeadingLevel = 4;                               // 前面没有 1/2/3 级，直接跳到 4
section.Blocks.Add(levelSkip);

Table table = new Table();
table.Columns.Add(new TableColumn(120));
table.Columns.Add(new TableColumn(70));
table.Alignment = TableAlignment.Left;                    // 没居中
table.RepeatHeaderRow = false;                            // 表头不跨页重复（该属性默认为 true）
table.AddRow("工序", "天数");
table.AddRow("基层处理", "12");
table.AddRow("涂料施工", "35");
section.Blocks.Add(table);

Directory.CreateDirectory("out");
DocxWriter.Write(draft, "out/draft.docx");
```

## 2. 读回来看看结构

```csharp
OfficeDocument document = DocxReader.Load("out/draft.docx");

Console.WriteLine(document.Metadata.Title);            // 项目实施方案（草稿）
Console.WriteLine(document.Sections.Count);            // 1

foreach (Block block in document.FirstSection.Blocks)
{
    if (block is Paragraph p)
    {
        Console.WriteLine($"[段落 H{p.HeadingLevel}] {p.GetText()}");
    }
    else if (block is Table t)
    {
        Console.WriteLine($"[表格 {t.Rows.Count}x{t.Columns.Count}] 对齐={t.Alignment} 重复表头={t.RepeatHeaderRow}");
    }
}
```

实测输出（注意末尾多出的一段空段）：

```text
[段落 H0] 项目实施方案
[段落 H0]
[段落 H0]
[段落 H0] 一、实施范围
[段落 H0] 本项目覆盖 3 个楼栋的公区与地下车库，含 墙面涂料 、门窗 五金更换，以及 机电末端 检修。
[段落 H0] 工期： 90 日历天
[段落 H0]
[段落 H4] （三）质量目标
[表格 3x2] 对齐=Left 重复表头=False
[段落 H0]
```

正文以表格结尾时，写出器会在表格后补一个空段落——OOXML 要求 `w:body` 的最后一个元素是段落，所以回读会比原模型多一块。这一空段会被下一步的"删除空段落"一并清掉。

只要纯文本时用 `DocxReader.ExtractText(path)`，不必构造完整模型（本文档的草稿提出 119 个字符）。

## 3. 模板填充

模板就是普通 DOCX，里面写占位符；跨 Run 拆开的占位符也能识别（Word 常把 `{{name}}` 拆成好几段）。

先生成模板。`{{#each 列表}}` 与 `{{/each}}` 各占一个段落，中间是要重复的内容：

```csharp
OfficeDocument template = new OfficeDocument();
Section tsection = template.AddSection(PageSetup.A4());

Paragraph greeting = new Paragraph();
greeting.AddText("{{customerName}}，您好：", new TextFormat("SimSun", 11));
tsection.Blocks.Add(greeting);

tsection.AddParagraph("我方确认以下订单：");

Paragraph loop = new Paragraph();
loop.AddText("{{#each items}}", new TextFormat("SimSun", 11));
tsection.Blocks.Add(loop);

Paragraph line = new Paragraph();
line.AddText("{{name}}　数量 {{qty}}　金额 {{amount:0.00}}", new TextFormat("SimSun", 11));
tsection.Blocks.Add(line);

Paragraph loopEnd = new Paragraph();
loopEnd.AddText("{{/each}}", new TextFormat("SimSun", 11));
tsection.Blocks.Add(loopEnd);

Paragraph total = new Paragraph();
total.AddText("合计：{{total:0.00}} 元　签订日期 {{signDate:yyyy-MM-dd}}", new TextFormat("SimSun", 11));
tsection.Blocks.Add(total);

DocxWriter.Write(template, "out/contract-template.docx");
```

再填充：

```csharp
TemplateContext context = new TemplateContext();
context["customerName"] = "示例客户";
context["total"] = 1268.0;
context["signDate"] = new DateTime(2026, 10, 21);
context["items"] = new[]
{
    new { name = "无线键盘", qty = 2, amount = 398.0 },
    new { name = "机械鼠标", qty = 5, amount = 449.5 },
    new { name = "清洁服务", qty = 1, amount = 420.5 }
};

TemplateOptions options = new TemplateOptions
{
    LeaveUnresolvedTokens = false,   // 没给值的占位符留空，而不是保留 {{xxx}}
    HtmlEncodeValues = false,
    Culture = System.Globalization.CultureInfo.InvariantCulture
};

DocxTemplate.Fill("out/contract-template.docx", "out/contract.docx", context, options);

string filled = DocxReader.ExtractText("out/contract.docx");
Console.WriteLine(filled.Contains("示例客户"));      // True
Console.WriteLine(filled.Contains("机械鼠标"));      // True
Console.WriteLine(filled.Contains("{{#each"));       // False —— 循环标记已被清掉
Console.WriteLine(filled.Contains("449.50"));        // True —— 按 :0.00 格式化
Console.WriteLine(filled.Contains("2026-10-21"));    // True —— 按 :yyyy-MM-dd 格式化
```

实测填出来的正文：

```text
示例客户，您好：
我方确认以下订单：
无线键盘　数量 2　金额 398.00
机械鼠标　数量 5　金额 449.50
清洁服务　数量 1　金额 420.50
合计：1268.00 元　签订日期 2026-10-21
```

段落级 `{{#each}}` 会把标记所在的段落整体换成 N 份内容；列表为空时循环块整块消失。表格行循环的写法相同，把 `{{#each}}` / `{{/each}}` 放在循环行的首列与末列。循环体内还能用 `@index`、`@number`、`@first`、`@last` 和 `.`（当前项本身）；取列表项属性用 `{{name}}`，对象嵌套用点路径 `{{order.id}}`，不支持 `[i]` 下标。

### 3.1 模板夹带的可执行部件会被剥掉

填充是整包搬运：模板里除了正文，宏工程、嵌入的 OLE / ActiveX 对象都会跟着进成品。收到别人发来的 `.docx` 当模板用时，`DocxTemplate` 会无条件剥掉这些部件，连带删除它们的关系项、`[Content_Types].xml` 里的 `Override`，以及正文里指向它们的 `r:id`（连同 `<w:object>` 这类承载元素），留下一个自洽的纯内容包。下面手工拼一个"带宏"的模板，六个部件填充后剩四个：

```csharp
// 六项：[Content_Types].xml、_rels/.rels、word/document.xml、word/_rels/document.xml.rels、
//       word/vbaProject.bin、word/embeddings/ole1.xlsx
// document.xml 里同时引用 rIdSafe（https 外链）、rIdFile（file: 外链）与 rIdEmbed（嵌入对象）
byte[] template = ZipBytes(entries);
Console.WriteLine("模板字节=" + template.Length + " 模板部件数=6");

TemplateContext ctx = new TemplateContext();
ctx.Set("customerName", "示例客户");
byte[] filledBytes = DocxTemplate.Fill(template, ctx);

Dictionary<string, byte[]> parts = ReadParts(filledBytes);
Console.WriteLine("填充后部件数=" + parts.Count);
Console.WriteLine("填充后部件=" + string.Join(" ", parts.Keys.OrderBy(k => k)));
string document = Text(parts["word/document.xml"]);
string rels = Text(parts["word/_rels/document.xml.rels"]);
Console.WriteLine("正文保留 rIdSafe=" + document.Contains("rIdSafe")
                  + " 保留 rIdFile=" + document.Contains("rIdFile")
                  + " 保留 rIdEmbed=" + document.Contains("rIdEmbed")
                  + " 保留 w:object=" + document.Contains("<w:object"));
Console.WriteLine("关系保留 https=" + rels.Contains("https://example.com/report")
                  + " 保留 file:=" + rels.Contains("file:")
                  + " 保留 vbaProject=" + rels.Contains("vbaProject"));
Console.WriteLine("正文=" + DocxReader.Load(filledBytes).GetPlainText().Trim());
```

```text
模板字节=1718 模板部件数=6
填充后部件数=4
填充后部件=_rels/.rels [Content_Types].xml word/_rels/document.xml.rels word/document.xml
正文保留 rIdSafe=True 保留 rIdFile=False 保留 rIdEmbed=False 保留 w:object=False
关系保留 https=True 保留 file:=False 保留 vbaProject=False
正文=客户：示例客户 官网本机
```

要点：宏部件与被嵌入的 OLE 都不再随成品带走，指向它们的引用一并清理，所以不会出现"点开是空白/损坏对象"的悬空 `r:id`；模板里 `file:` / `smb:` 这类指向本机或共享盘的超链接按同一份协议白名单被剥掉，`https:` 外链与正文文字都不受影响（`官网本机` 就是两个链接的文字部分）。相对目标如 `../embeddings/oleObject1.bin` 会先规范化再判定，不会因为字符串拼接漏剥。要保留宏请在填充之后自行按 `.docm` 组装。整包解压字节数（含原样搬运的条目）统一计入 `OfficePackageLimits` 的累计上限。

## 4. 批量整理

`WordNormalizeOptions` 的默认构造就是"中文公文常用值"，绝大多数开关**默认已打开**。所以下面这段是在改目标值、并关掉两处演示中不需要的行为：

```csharp
WordNormalizeOptions normalize = new WordNormalizeOptions
{
    MergeShortParagraphs = false,          // 默认 true：短于 ShortParagraphThreshold(18) 的段落会被并进相邻段落
    RegenerateTableOfContents = false,     // 默认 true：会按标题插入目录并加分页符
    RemoveTinyImages = false,              // 默认 true：边长 ≤ TinyImageMaxPixels(1) 的图片会被删除

    BodyEastAsiaFont = "宋体",
    BodyAsciiFont = "Times New Roman",
    BodyFontSize = 12,
    BodyLineSpacing = 1.5,                 // 行距是倍数，不是磅值
    FirstLineIndentCharacters = 2,

    PromoteFakeHeadings = true,
    FakeHeadingMinFontSize = 14,
    HeadingFontSizes = new double[] { 16, 15, 14, 13, 12, 12, 12, 12, 12 },

    NormalizeTables = true,
    UnifyTableAlignment = true,
    TableAlignment = TableAlignment.Center,
    RepeatTableHeaderRows = true
};

WordNormalizeResult result = WordNormalizer.Normalize(document, normalize);
OfficeDocument cleaned = result.Document;   // 与 document 是同一个实例：整理是就地修改

Console.WriteLine("变更 " + result.Changes.Count + " 项，问题 " + result.Issues.Count + " 项");
foreach (WordChange change in result.Changes)
{
    Console.WriteLine($"  [{change.Category}] {change.Description} @ {change.Location}");
}
```

实测输出（草稿经第 2 步回读后的那份模型）：

```text
变更 9 项，问题 1 项
  [段落整理] 删除空段落 @ 第1节 第10块
  [段落整理] 删除空段落 @ 第1节 第7块
  [段落整理] 删除空段落 @ 第1节 第3块
  [段落整理] 删除空段落 @ 第1节 第2块
  [标题整理] 伪标题升级为 1 级标题: 项目实施方案 @ 第1节 第1块
  [标题整理] 伪标题升级为 2 级标题: 一、实施范围 @ 第1节 第2块
  [标题整理] 标题层级跳跃 4 -> 3 @ 第1节 第5块
  [样式整理] 统一正文字体/字号，共 11 处文本 @
  [样式整理] 统一行距/段距，共 11 个段落 @
```

前后统计对比在 `result.Before` / `result.After`（都是 `WordDocumentMetrics`），完整报告文本用 `result.ToReport()`：

```csharp
Console.WriteLine("before: 段落=" + result.Before.ParagraphCount + " 空段=" + result.Before.EmptyParagraphCount + " 标题=" + result.Before.HeadingCount);
Console.WriteLine("after : 段落=" + result.After.ParagraphCount + " 空段=" + result.After.EmptyParagraphCount + " 标题=" + result.After.HeadingCount);
```

```text
before: 段落=15 空段=4 标题=1
after : 段落=11 空段=0 标题=3
```

`段落` 计数含表格单元格内的段落。整理后的表格与标题状态：

```csharp
Table cleanedTable = cleaned.FirstSection.Blocks.OfType<Table>().First();
Console.WriteLine("对齐=" + cleanedTable.Alignment + " 重复表头=" + cleanedTable.RepeatHeaderRow
                  + " 自适应=" + cleanedTable.Autofit + " 表头行=" + cleanedTable.Rows[0].IsHeader);
foreach (Paragraph p in cleaned.FirstSection.Blocks.OfType<Paragraph>().Where(x => x.HeadingLevel > 0))
{
    Console.WriteLine("  H" + p.HeadingLevel + " " + p.GetText());
}
```

```text
对齐=Center 重复表头=True 自适应=True 表头行=True
  H1 项目实施方案
  H2 一、实施范围
  H3 （三）质量目标
```

各步骤的判定规则：

- `PromoteFakeHeadings` 按"字号与正文的比值 + 是否加粗"识别假标题，再据此定级别：命中 `第X章/节/篇/部分` 句式或比值 ≥ 1.5 记 1 级，≥ 1.3 记 2 级，≥ 1.15 记 3 级，其余记 4 级。这里 22pt / 16pt 相对 12pt 正文分别是 1.83 与 1.33，所以得到 H1 与 H2。低于 `FakeHeadingMinFontSize` 且不满足比值条件的加粗段落不会被提升。
- `FixHeadingLevelSkips` 按文档顺序把跳级压回相邻级别（本文的 4 → 3）。
- `NormalizeStyles` 用 `HeadingFontSizes` 重写各级标题字号，所以 22pt 的标题最终是 16pt（见第 7 步的字号分布）。
- `NormalizeTables` 会把首行标记为表头行并开启跨页重复，其余行取消标记；`AutofitTables` 打开表格自适应。

## 5. 文件级流水线（含备份与回滚）

```csharp
WordProcessResult process = WordProcessor.Process("out/draft.docx", "out/clean.docx", normalize);

Console.WriteLine(process.OutputPath);   // out/clean.docx
Console.WriteLine(process.BackupPath);   // out/draft.backup-20261021081601.docx，与原文件同目录、带时间戳
Console.WriteLine(process.ReportPath);   // out/clean.docx.report.txt，固定为 outputPath + ".report.txt"
Console.WriteLine(process.Report);       // 与 Normalize 的 ToReport() 同一份文本
Console.WriteLine(process.Normalize.Changes.Count);
Console.WriteLine(process.Diff.ToText());
```

`Process` 支持 `.docx` / `.docm` / `.doc` 入口，内部顺序是"备份原文件 → 加载 → 就地整理 → 写出 → 落报告文件"。

需要还原时，`Rollback` 用备份**覆盖**指定的目标文件，通常就是覆盖回原文件：

```csharp
WordProcessor.Rollback(process.BackupPath, "out/draft.docx");
```

`Normalize` 就地修改模型且本包不公开深拷贝，所以试算时从文件重新加载一份，或先 `WordProcessor.CreateBackup` 备份文件：

```csharp
OfficeDocument probe = DocxReader.Load("out/draft.docx");
WordNormalizeResult trial = WordNormalizer.Normalize(probe, normalize);
Console.WriteLine(trial.Changes.Count);                  // 9
Console.WriteLine(probe.FirstSection.Blocks.Count);      // 6（整理前 10 块）
```

## 6. 一致性校验

不改文档，只报问题：

```csharp
OfficeDocument raw = DocxReader.Load("out/draft.docx");
List<WordIssue> issues = WordConsistencyChecker.Check(raw, normalize);
foreach (WordIssue issue in issues)
{
    Console.WriteLine($"{issue.Severity} {issue.Message} @ {issue.Location}");
}

Console.WriteLine("Error 级 " + issues.Count(i => i.Severity == WordIssueSeverity.Error) + " 项");
```

实测输出：

```text
Warning 仍存在 4 个空段落 @
Warning 正文存在 4 处非规范字体 @
Warning 正文存在 4 处非规范字号 @
Info 存在 8 个段落行距/段距与规范不一致 @
Info 统计: 标题 1 个, 表格 1 个, 目录条目 0 条 @
Error 级 0 项
```

`Severity` 分 `Info` / `Warning` / `Error`。当前会命中的检查项：空段落、非规范字体、非规范字号、行距/段距不一致、半角句末标点、不可见控制字符、标题层级跳跃、表格空行、图片超出正文宽度，最后一项是统计信息。校验基准来自传入的 `WordNormalizeOptions`——同一份文档配不同的目标字号，问题数量会不同。第 4 步的 `Normalize` 结尾也会自动跑一次校验，结果并进 `result.Issues`。

## 7. 两版差异与统计

```csharp
OfficeDocument before = DocxReader.Load("out/draft.docx");
OfficeDocument after = DocxReader.Load("out/clean.docx");

WordDocumentDiff diff = WordDocumentDiff.Compare(before, after);
Console.WriteLine(diff.HasDifferences);                       // False
Console.WriteLine($"+{diff.AddedCount} -{diff.RemovedCount} ={diff.UnchangedCount}");   // +0 -0 =8
Console.WriteLine(diff.ToText());
```

差异比对的是**正文文本行**，不含字体、字号、行距这些格式属性。本草稿的整理只改格式和删空段，所以 `HasDifferences` 为 `False`、未变化 8 行；要看到真正的增删，比较两份文本不同的版本即可：

```csharp
OfficeDocument revised = DocxReader.Load("out/clean.docx");
revised.FirstSection.Blocks.RemoveAt(revised.FirstSection.Blocks.Count - 1);
revised.FirstSection.AddParagraph("新增：验收标准见附表。");

WordDocumentDiff revision = WordDocumentDiff.Compare(after, revised);
Console.WriteLine(revision.HasDifferences + " +" + revision.AddedCount + " -" + revision.RemovedCount);
Console.WriteLine(revision.ToText());
```

```text
True +1 -0
=== 整理前/后差异对比 ===
未变化行: 8
删除行: 0
新增行: 1
--- 新增内容（前 200 条）---
  + 新增：验收标准见附表。
```

`AddedLines` / `RemovedLines` 是原始文本行，`ToText()` 最多列出前 200 条。

统计用 `WordDocumentMetrics.Measure`：

```csharp
WordDocumentMetrics metrics = WordDocumentMetrics.Measure(after);
Console.WriteLine($"段落 {metrics.ParagraphCount}，空段 {metrics.EmptyParagraphCount}，" +
                  $"字 {metrics.CharacterCount}，表 {metrics.TableCount}，图 {metrics.ImageCount}，标题 {metrics.HeadingCount}");
Console.WriteLine("各级标题: " + string.Join("/", Enumerable.Range(1, 4).Select(l => metrics.HeadingCounts[l])));
foreach (KeyValuePair<string, int> font in metrics.FontUsage)
{
    Console.WriteLine($"  字体 {font.Key} 用 {font.Value} 处");
}
foreach (KeyValuePair<string, int> size in metrics.FontSizeUsage)
{
    Console.WriteLine($"  字号 {size.Key} 用 {size.Value} 处");
}
```

```text
段落 12，空段 1，字 92，表 1，图 0，标题 3
各级标题: 1/1/1/0
  字体 黑体 用 3 处
  字体 宋体 用 8 处
  字号 16 用 1 处
  字号 15 用 1 处
  字号 14 用 1 处
  字号 12 用 8 处
```

`HeadingCounts` 是长度 10 的数组，下标即标题级别，下标 0 恒为 0（普通段落不计入，`HeadingCount` 只累加 1 级以上）。`字 92` 是正文、页眉页脚与表格单元格所有段落文本长度之和（含空格），`空段 1` 来自第 2 步说的表尾补段。字体名写成"黑体/宋体"是因为整理时用的是 `BodyEastAsiaFont`/`HeadingEastAsiaFont` 的中文字面值；西文用 `BodyAsciiFont`。

## 8. 旧版 .doc

```csharp
string path = "legacy.doc";                       // Word 97-2003，换成你自己的文件
OfficeDocument legacy = DocReader.Load(path);
string plain = DocReader.ExtractText(path);

Console.WriteLine(legacy.FirstSection.Blocks.Count);
Console.WriteLine(plain.Length);

DocxWriter.Write(legacy, "out/legacy.docx");      // 最常见用法：.doc -> .docx
```

用仓库里的 `tests/Wisdom.Office.Tests/fixtures/legacy-sample.doc` 实测：`Blocks.Count = 2`、纯文本 57 字符，转出的 DOCX 为 3760 字节；表格行内的单元格以制表符分隔。`.doc` 只读，复杂域、批注、修订不保证还原，转出结果保留正文段落、表格与文本格式。

## 9. 七种 Word 格式互转

同一份 `OfficeDocument` 可以写成 DOCX、DOTX、RTF、HTML、ODT 或纯文本；读取侧额外支持旧版 `.doc`。下面用一段代码走完一条格式链：

```csharp
using Wisdom.Word;

string input = "fixtures/legacy-sample.doc";   // 或 .docx / .dotx / .rtf / .html / .odt / .txt

WordFormatConverter.Convert(input, "out/chain.rtf");
WordFormatConverter.Convert("out/chain.rtf", "out/chain.html");
WordFormatConverter.Convert("out/chain.html", "out/chain.odt");
WordFormatConverter.Convert("out/chain.odt", "out/chain.docx");
WordFormatConverter.Convert("out/chain.docx", "out/chain.dotx");
WordFormatConverter.Convert("out/chain.dotx", "out/chain.txt");

WordDocumentFormat detected = WordFormatConverter.DetectFormat("out/chain.odt");
Console.WriteLine(detected);                    // Odt

OfficeDocument document = WordFormatConverter.Load("out/chain.odt");
byte[] html = WordFormatConverter.ToArray(document, WordDocumentFormat.Html);
```

7 种输入 × 6 种输出共 42 种组合，仓库测试会逐一转换并回读，断言标题、表格合并、书签、链接、图片、两节页眉页脚与元数据仍然存在（纯文本只校验正文）。RTF/HTML/ODT 与 DOCX 的差异主要在高级包结构：样式表、主题、修订、数字签名不会跨格式保留。

## 10. 修订、SmartArt 与数字签名

Word 审阅产生的插入/删除修订在 DOCX 读写中保留：插入内容带 `RevisionType.Insert`，被删除内容进入 `DeletedInlines`，段落标记修订记录在 `Paragraph.Revision`。

```csharp
OfficeDocument reviewed = DocxReader.Load("out/reviewed.docx");

int changes = WordRevisions.Count(reviewed);
Paragraph first = (Paragraph)reviewed.FirstSection.Blocks[0];
Console.WriteLine("删除内容：" + first.GetDeletedText());
foreach (Inline inline in first.Inlines)
{
    if (inline.Revision != null)
        Console.WriteLine(inline.Revision.Type + " by " + inline.Revision.Author);
}

// 二选一：接受后写回最终稿，或拒绝后回到原始稿
WordRevisions.AcceptAll(reviewed);
// WordRevisions.RejectAll(reviewed);
```

自己写入修订时直接给行内元素或段落赋 `RevisionInfo`，再用 `DocxWriter.Write` 输出；`w:author`、`w:date` 与 `w:delText` 会一并生成。接受/拒绝修订不会重建原始字符位置，拒绝时删除内容统一追加到段落末尾。

SmartArt 与数字签名属于包级结构，统一模型不承载图形布局或签名加密数据，因此提供只读检查：

```csharp
foreach (WordSmartArt diagram in WordSmartArtReader.Read("out/board.docx"))
{
    Console.WriteLine(diagram.NodeCount + " 个节点");
    Console.WriteLine(diagram.Text);
}

WordDigitalSignatureReport signatures = WordDigitalSignatureReader.Read("out/signed.docx");
Console.WriteLine(signatures.IsSigned + " / " + signatures.Signatures.Count);
foreach (WordDigitalSignature signature in signatures.Signatures)
{
    Console.WriteLine(signature.SignatureMethod + " " + signature.SignerSubject);
}
```

SmartArt 检查器提取节点文本，但重新生成的包不会保留图形结构；数字签名检查器不验证密码学有效性，且任何重新写包都会使原签名失效。

## 11. 自定义文档属性

自定义属性适合存放业务标签，例如项目编号、审批版本、归档时间。它们写在 `docProps/custom.xml`，Word 的“文件 → 信息 → 属性 → 高级属性 → 自定义”里可以看到。

```csharp
OfficeDocument tagged = new OfficeDocument();
tagged.AddSection().AddParagraph("带标签的正文。");

CustomPropertyCollection properties = tagged.Metadata.CustomProperties;
properties.Set("项目编号", "QF-2026-018");
properties.Set("版本", 3);
properties.Set("已审核", true);
properties.Set("抽检比例", 0.15);
properties.Set("归档时间", new DateTime(2026, 10, 20, 8, 30, 0, DateTimeKind.Utc));

DocxWriter.Write(tagged, "out/tagged.docx");

OfficeDocument readTagged = DocxReader.Load("out/tagged.docx");
Console.WriteLine(readTagged.Metadata.CustomProperties.GetString("项目编号")); // QF-2026-018
Console.WriteLine(readTagged.Metadata.CustomProperties.GetInt32("版本"));      // 3
Console.WriteLine(readTagged.Metadata.CustomProperties.GetBoolean("已审核"));  // True

// 名称不区分大小写，枚举保持写入顺序
foreach (string name in readTagged.Metadata.CustomProperties.Names)
{
    Console.WriteLine(name + " = " + readTagged.Metadata.CustomProperties[name]);
}
```

支持 `string`、`bool`、`int`、`long`、`double`、`decimal` 与 `DateTime`，分别写成 `vt:lpwstr`、`bool`、`i4`、`i8`、`r8`、`decimal`、`filetime`。写 `null` 会删除同名属性；`WordDocumentOperations.Clone` 会连自定义属性一起深拷贝。

属性只随 DOCX/DOTX 包保存，转成 RTF、HTML、ODT 或纯文本后不会出现。

## 12. 标签纸与信封

### 12.1 标签纸

`WordLabels` 直接输出可打印的标签纸。先用固定文本排一页 Avery L7160（A4，3 列 × 7 行）：

```csharp
string[] first = { "张三", "上海市浦东新区世纪大道 100 号", "200120" };
string[] second = { "李四", "北京市朝阳区建国路 8 号", "100022" };

OfficeDocument sheet = WordLabels.Create(
    WordLabelLayout.AveryL7160(),
    new[] { first, second },
    new WordLabelOptions { FontSize = 10.5 });

DocxWriter.Write(sheet, "out/labels.docx");
Console.WriteLine(WordLabelLayout.AveryL7160().LabelsPerPage);   // 21
```

超过每页容量会自动分页。数据来自集合时用模板：

```csharp
OfficeDocument fromData = WordLabels.CreateFromRecords(
    WordLabelLayout.Avery5163(),                 // US Letter，2 列 × 5 行
    new[]
    {
        new { Name = "张三", Address = "上海市浦东新区世纪大道 100 号", Zip = "200120" },
        new { Name = "李四", Address = "北京市朝阳区建国路 8 号", Zip = "100022" }
    },
    "{{Name}}\n{{Address}}\n{{Zip}}",
    new WordLabelOptions { FontSize = 9, Bold = true });

DocxWriter.Write(fromData, "out/labels-records.docx");
```

`WordLabelLayout.FromMillimeters` 可定义任意版式；`WordLabels.CreateBlank` 生成空白标签纸，便于套打。生成物是每页一张无边框表格，标签行高为“最小值”，内容超出单张高度时行高扩展、后续行下移。打印前建议先用普通纸试打一张，按打印机实际边距微调 `TopMargin` / `LeftMargin`。

### 12.2 表单标签纸

需要“姓名：____”这类可手写的标签时，用 `WordLabels.CreateForm`。下面生成一张 A4 2×5 的学生信息标签，带标签外框和三条填写线：

```csharp
WordLabelLayout studentLayout = WordLabelLayout.FromMillimeters(
    "学生信息标签 A4 2×5",
    210, 297, 2, 5, 82, 48, 25, 22, 4, 4);

OfficeDocument form = WordLabels.CreateForm(
    studentLayout,
    new[] { "姓名：", "年级：", "学号：" },
    new WordLabelOptions { ShowCellBorders = true });

DocxWriter.Write(form, "out/student-labels.docx");
```

`ShowCellBorders` 控制标签外框，邮件标签保持默认 `false` 即可。每条填写线是嵌套表格中列的下边框：左列放字段名，中列空白，右列留 6 磅空白，横线因此不会贴住标签边框；三列底对齐，横线紧贴文字下方。表格外框就是标签边框，表尾用空白行补足标签高度，垂直间隙留在框外，上下两张标签的边框不会粘在一起。`FieldSpacing` 调整相邻填写线的距离，`BorderColor` / `BorderWidth` 调整外框和线条的样式。空字段会被忽略，字段列表为空时抛 `ArgumentException`。

### 12.3 分格标签与表格标签

留样标签这种“每行文字之间画横线”的版式，把 `ShowRowSeparators` 打开。下面生成 A4 4×10、每张四行的留样标签：

```csharp
WordLabelLayout retentionLayout = WordLabelLayout.FromMillimeters(
    "留样标签 A4 4×10",
    210, 297, 4, 10, 45, 25, 12, 10.5, 3, 2);

string[] retentionLines =
{
    "201  年  月  日  点  分",
    "食品名：          （    克）",
    "留样人：",
    "审核人："
};
List<string[]> retention = new List<string[]>();
for (int i = 0; i < retentionLayout.LabelsPerPage; i++) retention.Add(retentionLines);

OfficeDocument retentionSheet = WordLabels.Create(
    retentionLayout,
    retention,
    new WordLabelOptions
    {
        FontFamily = "Arial",          // 西文用 Arial，空格更窄，长行不会被挤断
        EastAsiaFontFamily = "SimSun", // 中文用宋体
        FontSize = 9,
        ShowCellBorders = true,
        ShowRowSeparators = true,
        BorderWidth = 0.75
    });
DocxWriter.Write(retentionSheet, "out/retention-labels.docx");
```

`ShowRowSeparators` 把标签正文改成一列多行的嵌套表格：行高按标签高度均分，行间画横线；`ShowCellBorders` 提供外框，两者合起来就是完整的表格框。整页按最多文本行数定行数，行数少的标签也保留空白格，框线上下对齐；内部只有横线没有竖线，垂直间隙留在标签框外。

化学品安全标签这类合并单元格较多的版式直接用 `Table` 搭。下面这张 A5 横向标签有三列六行：纯度和图形跨行，底部两行左侧合并两列、内嵌联系表，右侧放 UN/CN 编号。把任意 GHS 腐蚀图形存为 `ghs.png` 放在当前目录：

```csharp
PageSetup page = new PageSetup();
page.Width = 210.0 * 72 / 25.4;
page.Height = 148.0 * 72 / 25.4;
page.MarginLeft = page.MarginRight = 6.0 * 72 / 25.4;
page.MarginTop = page.MarginBottom = 6.0 * 72 / 25.4;

OfficeDocument chemical = new OfficeDocument();
Section section = chemical.AddSection(page);

BorderLine edge = new BorderLine(BorderStyle.Single, 1.0, OfficeColor.Black);
Table box = new Table
{
    Borders = new TableBorders
    {
        Top = edge, Bottom = edge, Left = edge, Right = edge,
        InsideHorizontal = edge, InsideVertical = edge
    },
    CellPadding = 3,
    Autofit = false,
    RepeatHeaderRow = false
};
box.Columns.Add(new TableColumn(38.0 * 72 / 25.4));
box.Columns.Add(new TableColumn(120.0 * 72 / 25.4));
box.Columns.Add(new TableColumn(40.0 * 72 / 25.4));
section.Blocks.Add(box);

TableRow titleRow = box.AddRow();
titleRow.Height = 26.0 * 72 / 25.4;
TableCell purity = new TableCell { RowSpan = 2, VerticalAlignment = VerticalAlignment.Middle };
purity.Blocks.Add(Centered("纯度：", 8.5, false));
purity.Blocks.Add(Centered("≥30%", 11, true));
titleRow.Cells.Add(purity);
TableCell title = new TableCell { VerticalAlignment = VerticalAlignment.Middle };
title.Blocks.Add(Centered("hydrochloric acid", 10.5, true));
title.Blocks.Add(Centered("盐酸", 16, true));
title.Blocks.Add(Centered("HCl", 10, true));
titleRow.Cells.Add(title);
TableCell pictogram = new TableCell { RowSpan = 4, VerticalAlignment = VerticalAlignment.Middle };
Paragraph imageParagraph = Centered(string.Empty, 8.5, false);
imageParagraph.AddImage(new OfficeImage(File.ReadAllBytes("ghs.png")), 34.0 * 72 / 25.4, 34.0 * 72 / 25.4);
pictogram.Blocks.Add(imageParagraph);
titleRow.Cells.Add(pictogram);

// 纯度与图形仍在跨行，危险行只放中间一格
TableRow dangerRow = box.AddRow();
dangerRow.Height = 14.0 * 72 / 25.4;
TableCell danger = new TableCell { VerticalAlignment = VerticalAlignment.Middle };
danger.Blocks.Add(Centered("危  险", 12, true));
danger.Blocks.Add(Centered("酸性腐蚀，对人体有害", 9, true));
dangerRow.Cells.Add(danger);

TableRow safetyRow = box.AddRow();
safetyRow.Height = 55.0 * 72 / 25.4;
TableCell storage = new TableCell { RowSpan = 2, VerticalAlignment = VerticalAlignment.Middle };
storage.Blocks.Add(Centered("最大储存量：", 8.5, false));
storage.Blocks.Add(Centered("60 吨", 11, true));
safetyRow.Cells.Add(storage);
TableCell measures = new TableCell { VerticalAlignment = VerticalAlignment.Top };
measures.Blocks.Add(Left("安全措施：", 9, true));
measures.Blocks.Add(Bullet("◆ 皮肤接触：立即脱去污染的衣着，用大量流动清水冲洗至少 15 分钟。就医。"));
measures.Blocks.Add(Bullet("◆ 眼睛接触：立即提起眼睑，用大量流动清水或生理盐水彻底冲洗至少 15 分钟。就医。"));
measures.Blocks.Add(Left("灭火方法：", 9, true));
measures.Blocks.Add(Left("用碱性物质如碳酸氢钠、碳酸钠、消石灰等中和。", 8.5, false));
safetyRow.Cells.Add(measures);

TableRow noticeRow = box.AddRow();
noticeRow.Height = 8.0 * 72 / 25.4;
noticeRow.Cells.Add(CenteredCell("请参阅安全技术说明书", 8.5, false));

AddContactRow(box, "生产企业：南宁化工股份有限公司", "应急电话：0771-4818653", "UN NO.:1789");
AddContactRow(box, "应急电话：火警：3668119", "值班：3668181　化学值班：3668311", "CN NO.:81013");

DocxWriter.Write(chemical, "out/chemical-safety.docx");
OfficeConverter.ToPdf(chemical, "out/chemical-safety.pdf");

// ---- 以下为上面用到的辅助函数 ----
static TextFormat LabelFormat(double size, bool bold)
{
    TextFormat format = new TextFormat("Arial", size);
    format.EastAsiaFontFamily = "SimSun";
    format.Bold = bold;
    return format;
}

static Paragraph Centered(string text, double size, bool bold)
{
    Paragraph paragraph = new Paragraph { Alignment = ParagraphAlignment.Center };
    paragraph.AddText(text, LabelFormat(size, bold));
    return paragraph;
}

static Paragraph Left(string text, double size, bool bold)
{
    Paragraph paragraph = new Paragraph { Alignment = ParagraphAlignment.Left, LineSpacing = 1.1 };
    paragraph.AddText(text, LabelFormat(size, bold));
    return paragraph;
}

static Paragraph Bullet(string text)
{
    Paragraph paragraph = Left(text, 8.5, false);
    paragraph.SpaceAfter = 2;
    paragraph.LeftIndent = 11;
    paragraph.FirstLineIndent = -11;
    return paragraph;
}

static TableCell CenteredCell(string text, double size, bool bold)
{
    TableCell cell = new TableCell { VerticalAlignment = VerticalAlignment.Middle };
    cell.Blocks.Add(Centered(text, size, bold));
    return cell;
}

static void AddContactRow(Table box, string left, string right, string code)
{
    TableRow row = box.AddRow();
    row.Height = 8.0 * 72 / 25.4;
    TableCell merged = new TableCell { ColSpan = 2, VerticalAlignment = VerticalAlignment.Middle };
    Table contact = new Table
    {
        Borders = TableBorders.None(),
        CellPadding = 0,
        Autofit = false,
        RepeatHeaderRow = false
    };
    contact.Columns.Add(new TableColumn(92.0 * 72 / 25.4));
    contact.Columns.Add(new TableColumn(62.0 * 72 / 25.4));
    TableRow contactRow = contact.AddRow();
    TableCell company = new TableCell { VerticalAlignment = VerticalAlignment.Middle };
    company.Blocks.Add(Left(left, 8.5, false));
    contactRow.Cells.Add(company);
    TableCell phone = new TableCell { VerticalAlignment = VerticalAlignment.Middle };
    Paragraph phoneText = Left(right, 8.5, false);
    phoneText.Alignment = ParagraphAlignment.Right;
    phone.Blocks.Add(phoneText);
    contactRow.Cells.Add(phone);
    merged.Blocks.Add(contact);
    merged.Blocks.Add(new Paragraph { LineSpacing = 0.1 });
    row.Cells.Add(merged);
    row.Cells.Add(CenteredCell(code, 8, false));
}
```

要点：

- 被 `RowSpan` 覆盖的后续行不能再放该列单元格，被 `ColSpan` 覆盖的列同理，否则 Word 中会错位。DOCX 中会写出 `w:vMerge` / `w:gridSpan`，回读后 `RowSpan` / `ColSpan` 保持不变。
- 单元格内的嵌套表格用于底部左右对齐的联系信息；嵌套表格后补一个 `LineSpacing = 0.1` 的空段落，避免 Word 的结尾段落撑高行。
- 图片按磅指定显示尺寸，段落居中；转成 PDF 时作为嵌入对象绘制，与文本一起排版。

### 12.4 信封

```csharp
OfficeDocument envelope = WordEnvelopes.Create(new WordEnvelopeOptions
{
    Layout = WordEnvelopeLayout.Dl(),            // DL、C5、Chinese5 或自定义
    Sender = "智慧办公\n上海市浦东新区示例路 1 号\n200120",
    Recipient = "张三\n上海市浦东新区世纪大道 100 号\n200120"
});
DocxWriter.Write(envelope, "out/envelope.docx");
```

批量信封每条记录一页，收件人模板按记录替换，寄件人可固定也可用模板：

```csharp
OfficeDocument batch = WordEnvelopes.CreateMany(
    new[]
    {
        new { Name = "张三", Address = "上海市浦东新区世纪大道 100 号" },
        new { Name = "李四", Address = "北京市朝阳区建国路 8 号" }
    },
    "{{Name}}\n{{Address}}",
    new WordEnvelopeOptions { Layout = WordEnvelopeLayout.Chinese5() },
    "智慧办公");

DocxWriter.Write(batch, "out/envelopes.docx");
```

地址以段前间距和左缩进定位，输出是普通段落，Word/WPS 与 PDF 导出都能渲染；不包含回邮窗口、条码或邮票框。

## 13. 书签与文档内跳转

书签有两种形态：覆盖整段的书签（写在 `Paragraph.BookmarkName` 上，与旧版本 API 兼容）和任意位置的范围书签（`BookmarkStart`/`BookmarkEnd` 两个行内标记，可以落在段落中间、跨段落、也可以跨表格单元格）。`document.Bookmarks` 每次访问都重新配对，改完不用刷新。

```csharp
WordDocumentBuilder builder = new WordDocumentBuilder();
builder.Font.FontFamily = "Microsoft YaHei";
builder.HeadingLevel = 1;
builder.Write("年度总结");
builder.InsertBookmark("Top");                 // 整段书签
builder.HeadingLevel = 0;

builder.InsertParagraph();
builder.Write("营收 ");
builder.StartBookmark("Revenue");              // 范围书签：只框住金额
builder.Write("1,280 万元");
builder.EndBookmark("Revenue");
builder.Write("，同比增长 12%。");

builder.InsertParagraph();
builder.InsertHyperlink("回到标题", "#Top");    // 表内跳转指向书签

DocxWriter.Write(builder.Document, "out/bookmarks.docx");

OfficeDocument marked = DocxReader.Load("out/bookmarks.docx");
Console.WriteLine(marked.Bookmarks.Count);
foreach (Bookmark bookmark in marked.Bookmarks)
{
    Console.WriteLine(bookmark.Name + " | " + bookmark.IsWholeParagraph + " | [" + bookmark.Text + "]");
}

Console.WriteLine(marked.Bookmarks["Revenue"].IsValid);
marked.Bookmarks["Revenue"].Text = "1,350 万元";   // 只换正文，书签还在
Console.WriteLine(marked.Bookmarks["Revenue"].Text);
Console.WriteLine(marked.GetPlainText());
Console.WriteLine(marked.Bookmarks.Contains("Top"));
marked.Bookmarks.Remove("Top");                    // 删标记，不删正文
Console.WriteLine(marked.Bookmarks.Count);
Console.WriteLine(marked.GetPlainText());
```

实测输出：

```text
2
Top | True | [年度总结]
Revenue | False | [1,280 万元]
True
1,350 万元
年度总结
营收 1,350 万元，同比增长 12%。
回到标题

True
1
年度总结
营收 1,350 万元，同比增长 12%。
回到标题
```

把光标挪到书签端点旁边，就能就地续写；ODT 与 RTF 同样保留范围（`text:bookmark-start`/`\bkmkstart`）：

```csharp
OfficeDocument wordDocument = new OfficeDocument();
WordDocumentBuilder moved = new WordDocumentBuilder(wordDocument);
moved.Write("合计 ");
moved.StartBookmark("Total");
moved.Write("待填");
moved.EndBookmark("Total");
moved.Write(" 元");
moved.MoveToBookmark("Total", false, true);        // 终点之后
moved.Write("（已复核）");
Console.WriteLine(wordDocument.Bookmarks["Total"].Text);
Console.WriteLine(wordDocument.GetPlainText());

WordFormatConverter.Save(wordDocument, "out/bookmarks.odt", WordDocumentFormat.Odt);
OfficeDocument fromOdt = WordFormatConverter.Load("out/bookmarks.odt");
Console.WriteLine(fromOdt.Bookmarks["Total"].Text);
Console.WriteLine(fromOdt.GetPlainText());
```

实测输出：

```text
待填
合计 待填（已复核） 元

待填
合计 待填（已复核） 元
```

几点约定：

- DOCX 的 `w:bookmarkStart` 带 `w:id` 和 `w:name`，`w:bookmarkEnd` 只带 `w:id`。读取端按 id 反查名称配对，写出端复用起点的 id；起点没落笔时不会留下无主的 `w:bookmarkEnd`。
- id 在每个部件内各自唯一（正文、页眉、页脚分开计数），页眉页脚里的书签同样能取到。Word 自己的内部书签 `_GoBack` 会被忽略。
- 书签名 1–40 个字符，以字母或下划线开头，只允许字母、数字与下划线（中文算字母）；`BookmarkNames.IsValid` 用于判断，`Validate` 会去空白并在非法时抛 `ArgumentException`，`StartBookmark`/`EndBookmark` 内部就走这套校验。
- 起点与终点缺一半时 `IsValid` 为 `false`，`Text` 返回空字符串，赋值会抛 `InvalidOperationException`——范围不明时不做删除。
- 起点标记恰好是段内第一个元素、终点标记是最后一个时，`IsWholeParagraph` 为 `true` 且 `Paragraph` 指向该段；整段书签和"框住整段的范围书签"在 DOCX 里是同一对 `w:bookmarkStart/End`，往返不会变形。
- HTML 只有点锚：范围写出后回读会退化成整段书签，名字与位置还在，范围边界丢失。
- 正文与表格之间的块级标记（`w:p` 之外的 `w:bookmarkStart`）以及脚注、尾注里的书签暂不读取。

## 常见问题

**整理后没有变化。** `WordNormalizeOptions` 的默认构造几乎打开了所有开关，先确认没有显式置 `false`；若确实没变化，看 `result.Changes` 是否为空——文档已符合规范时不会有任何变更。`PromoteFakeHeadings` 依赖字号比值与 `FakeHeadingMinFontSize`，正文里 12pt 的加粗短句不会被当作标题。

**目录凭空多出来。** `RegenerateTableOfContents` 默认为 `true`，会按现有标题重建目录并在其后加分页符；不需要就显式关掉。若文档一个标题都没有，校验阶段会给出"未识别到任何标题，目录无法生成"的提示。

**短段落被并到一起。** `MergeShortParagraphs` 默认 `true`，`ShortParagraphThreshold`（默认 18 字）以下的段落会向前合并；只想删空段就把它关掉。

**合并单元格在 Word 里错位。** 被跨行单元格覆盖的那一行必须不再为该列放单元格；回读时看 `TableCell.RowSpan` / `ColSpan` 与每行 `Cells.Count` 是否自洽。

**页码没出现。** 本包写的是 `PAGE` / `NUMPAGES` 域，真实数字由 Word 打开时计算；导成 PDF 时才会由渲染器替换成实际页码。

**模板占位符没被替换。** 先用 `DocxReader.ExtractText` 确认模板里确实是 `{{key}}`（全角花括号、多余空格都会导致不匹配），再用 `TemplateEngine.ContainsToken` 快速判断某个字符串是否含占位符。
