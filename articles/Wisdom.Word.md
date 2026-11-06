# 零依赖生成 Word 文档：Wisdom.Word 能做到什么程度

> 「纯 .NET 操作 Office」系列 · Word 篇

## 为什么又要写一个 Word 库

服务端出 Word 文档，常见做法要么是 `Interop`（服务端装 Office，基本违规），要么是"拼 HTML 改后缀当 doc 发出去"（打开就排版崩），要么是商用库（Aspose.Words 的价格劝退个人项目）。

`Wisdom.Word` 的思路和 `Wisdom.Excel` 一致：**DOCX 本质是一个 ZIP + WordprocessingML XML，直接用 `System.IO.Compression.ZipArchive` + LINQ to XML 处理**。唯一依赖是自己家的数据模型包 `Wisdom.Office.Core`，无第三方库，不要求装 Office 或 WPS。

```bash
dotnet add package Wisdom.Word
```

支持 `net461` 到 `net10.0` 六份目标框架，同一份源码编译，老 .NET Framework 项目和 Unity 都能引。

这个包的能力面比"读写 docx"要宽不少，本文按场景展开。

## 一、读取与生成

```csharp
using Wisdom.Word;
using Wisdom.Office.Core;

OfficeDocument document = DocxReader.Load("report.docx");   // 亦支持 Stream / byte[]
Console.WriteLine(document.Metadata.Title);

foreach (Section section in document.Sections)
    foreach (Block block in section.Blocks)
    {
        if (block is Paragraph p) Console.WriteLine(p.GetText());
        if (block is Table t) Console.WriteLine($"{t.Rows.Count} 行 x {t.Columns.Count} 列");
    }

string text = DocxReader.ExtractText("report.docx");   // 只要纯文本
```

读取覆盖段落与行内格式、表格（含合并单元格 `ColSpan`/`RowSpan`）、页眉页脚、图片、超链接（外部 URL 与 `#书签` 锚点）、书签、标题级别、对齐缩进与元数据。

生成侧从模型开始：

```csharp
OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "季度总结";
Section section = document.AddSection(PageSetup.A4());

// 页脚用真实页码域，Word 打开时自动计算
Paragraph footer = section.Footer.AddParagraph(string.Empty);
TextFormat small = new TextFormat("SimSun", 8);
footer.AddText("第 ", small);
footer.Inlines.Add(new FieldRun("PAGE", "{page}", small));
footer.AddText(" 页　共 ", small);
footer.Inlines.Add(new FieldRun("NUMPAGES", "{pages}", small));
footer.AddText(" 页", small);

Paragraph title = new Paragraph("成品检验报告") { HeadingLevel = 1, Alignment = ParagraphAlignment.Center };
section.Blocks.Add(title);
DocxWriter.Write(document, "report.docx");
```

合并单元格是 Word 排版的老大难，这里的写法是直接给 `TableCell` 设跨度：

```csharp
TableCell cell = new TableCell { ColSpan = 2, RowSpan = 1, BackgroundColor = OfficeColor.LightGray };
cell.Blocks.Add(new Paragraph("实测结果"));
row.Cells.Add(cell);
```

写出的是 `w:gridSpan` / `w:vMerge`，回读后跨度保持不变。生成的包通过内容类型、关系目标与 XML 合法性校验，Word 2010+ 和 WPS 可直接打开。

要注意：**分页、页码这些运行时排版由 Word/WPS 决定**，本包输出的是 `PAGE`/`NUMPAGES` 域和分页属性，不做像素级分页模拟。目录（TOC）同理，写入的是域，条目要打开后按 F9 生成。

## 二、WordDocumentBuilder：Aspose 风格的流式写法

如果你写过 Aspose.Words，会认得这个 API。它维护一个写入游标和当前格式，按顺序追加，特别适合"拼接报告"这种过程式生成：

```csharp
WordDocumentBuilder builder = new WordDocumentBuilder();
builder.Font.FontFamily = "Microsoft YaHei";
builder.Font.EastAsiaFontFamily = "Microsoft YaHei";   // 中西文字体分开设置

builder.HeadingLevel = 1;
builder.WriteLine("2026 年度报告");
builder.HeadingLevel = 0;

builder.InsertParagraph("正文第一段。");
builder.InsertBookmark("Intro");
builder.InsertHyperlink("访问官网", "https://example.com/report");
builder.InsertHyperlink("回到引言", "#Intro");          // 文档内锚点

builder.StartTable();
builder.InsertCell("项目"); builder.InsertCell("金额"); builder.EndRow();
builder.InsertCell("合计"); builder.InsertCell("100"); builder.EndTable();

builder.InsertImage(new OfficeImage(File.ReadAllBytes("logo.png")), 120, 40);
builder.InsertTableOfContents(3);
builder.InsertField("PAGE", "1");

DocxWriter.Write(builder.Document, "report.docx");
```

还能设 `Alignment`、`SpaceBefore/SpaceAfter`、`LineSpacing`、各向缩进、`KeepWithNext`、`PageBreakBefore`。

## 三、书签：整段与范围两种形态

`Paragraph.BookmarkName` 那种"整段书签"很多库都支持，容易漏的是**只框住半句话的范围书签**：

```csharp
builder.Write("合计 ");
builder.StartBookmark("Total");
builder.Write("128 万元");
builder.EndBookmark("Total");
builder.Write("（含税）");

OfficeDocument document = builder.Document;
document.Bookmarks["Total"].Text = "130 万元";   // 就地替换范围内正文，书签保留
builder.MoveToBookmark("Total", false, true);    // 定位到终点之后继续写
document.Bookmarks.Remove("Total");              // 只删标记，不删正文
```

`document.Bookmarks` 每次访问即时配对，读改存之间不需要刷新。底层细节也处理得挺严谨：DOCX 里 `w:bookmarkStart` 带 `w:id` 和 `w:name`，`w:bookmarkEnd` 只带 `w:id`，读取端按 id 反查名称配对，写出端复用起点的 id，因此起点未落笔时不会留下无主的 `w:bookmarkEnd`；id 在正文 / 页眉 / 页脚各自计数；Word 内部的 `_GoBack` 会被忽略。RTF 与 ODT 同样保留范围，HTML 只有点锚，范围读取后退化成整段书签。

## 四、查找替换

```csharp
int count = WordFindReplace.Replace(document, "旧公司名", "新公司名");   // 跨 Run 命中

WordFindReplaceOptions options = new WordFindReplaceOptions
{
    MatchCase = false,
    WholeWord = true,
    IncludeHeadersAndFooters = true,
    IncludeTables = true
};
WordFindReplace.Replace(document, "公司", "集团", options);

WordFindReplaceOptions regex = new WordFindReplaceOptions { UseRegex = true };
WordFindReplace.Replace(document, @"\d{4}年", "2026年", regex);   // 内置 2 秒匹配超时
```

"跨 Run 命中"是关键：Word 常把一个词拆进多个 Run（拼写检查、格式微调都会导致），朴素的字符串替换会漏掉一大半。替换文字继承首个命中 Run 的格式。正则带 2 秒超时，防灾难性回溯把线程挂死。

## 五、邮件合并与区域循环

`WordMailMerge` 同时认 Word 的 `MERGEFIELD` 域和 `{{token}}` 标记，数据源可以是字典、`DataRow` 或 `TemplateContext`：

```csharp
TemplateContext context = new TemplateContext();
context["customer"] = new { name = "示例客户" };    // 点路径 customer.name
context["Balance"] = 9876.5m;                       // MERGEFIELD Balance \# "#,##0.00"
context["Logo"] = new OfficeImage(File.ReadAllBytes("logo.png"));

int merged = WordMailMerge.Execute(document, context, new WordMailMergeOptions
{
    LeaveUnresolvedFields = false,
    Culture = CultureInfo.InvariantCulture,
    IncludeHeadersAndFooters = true
});
```

合并后的 `MERGEFIELD` 会转成普通文本或图文，避免 Word 更新域时把结果清空——这个细节踩过的人都知道有多烦。

`ExecuteWithRegions` 支持 `TableStart` / `TableEnd` 区域和任意深度嵌套：

```text
MERGEFIELD TableStart:orders
{{orderNo}}        {{customer}}
MERGEFIELD TableStart:lines
{{sku}}            {{qty}}
MERGEFIELD TableEnd:lines
MERGEFIELD TableEnd:orders
```

```csharp
root["orders"] = new[]
{
    new { orderNo = "A-001", customer = "海森公司",
          lines = new[] { new { sku = "键盘", qty = 2 }, new { sku = "鼠标", qty = 4 } } }
};
WordMailMerge.ExecuteWithRegions(document, root);

// DataTable / DataSet：表名即区域名，DataRelation 的子表自动成为嵌套区域
WordMailMerge.ExecuteWithRegions(document, dataSet);
```

规则明确：`IsHeader` 行只输出一次；嵌套字段先按子行解析再回退父行 / 全局；集合为空则整块区域移除；嵌套深度上限 32；不支持同名区域自我嵌套。

## 六、标签纸与信封：意外的实用

这块通常是商用库的付费功能，这里做得挺完整。`WordLabels` 内置 Avery 5160（US Letter 3×10）、5163（2×5）、L7160（A4 3×7），也支持毫米自定义：

```csharp
OfficeDocument sheet = WordLabels.Create(
    WordLabelLayout.AveryL7160(),
    new[]
    {
        new[] { "张三", "上海市浦东新区世纪大道 100 号", "200120" },
        new[] { "李四", "北京市朝阳区建国路 8 号", "100022" }
    },
    new WordLabelOptions { FontSize = 10.5 });

// 数据记录 + 多行模板，超出每页容量自动分页
OfficeDocument fromData = WordLabels.CreateFromRecords(
    WordLabelLayout.Avery5163(), records, "{{Name}}\n{{Address}}\n{{Zip}}");
```

除了邮件标签，还有两类国内特色场景：

**表单标签**（学生信息标签这种要手写的），`CreateForm` 自动输出"字段名 + 填写横线"，横线延伸到标签框右侧并留 6 磅空白，可加外框：

```csharp
OfficeDocument formLabels = WordLabels.CreateForm(
    studentLayout,
    new[] { "姓名：", "年级：", "学号：" },
    new WordLabelOptions { FontSize = 10.5, ShowCellBorders = true, FieldSpacing = 6, BorderWidth = 0.5 });
```

**分格标签**（食堂留样标签、样品签这种"每行文字之间画一条横线"），打开 `ShowRowSeparators` 即可；同一页内所有标签取最大文本行数作为行数，格子上下对齐，空白行也保留边框：

```csharp
OfficeDocument retentionSheet = WordLabels.Create(retentionLayout, retentionLabels,
    new WordLabelOptions
    {
        FontFamily = "Arial",          // 西文用 Arial，空格更窄
        EastAsiaFontFamily = "SimSun", // 中文用宋体
        FontSize = 9,
        ShowCellBorders = true,
        ShowRowSeparators = true,
        BorderWidth = 0.75
    });
```

更复杂的版式（化学品安全标签那种带合并单元格、警示图形和底部联系信息的），直接用 `Table` 手排：`RowSpan`/`ColSpan` 做合并、单元格内放图片或嵌套表格。要点是被跨度覆盖的后续行不能再放该列单元格，否则 Word 里会错位。

信封走 `WordEnvelopes`，内置 DL、C5、中式 5 号，地址用段前间距控制纵向位置、左缩进控制横向位置，输出为普通段落而不是文本框，所以 Word/WPS 和 PDF 导出都能直接渲染。

## 七、7 种格式互转

`WordFormatConverter` 以统一文档模型为中转：

| 格式 | 读 | 写 | 说明 |
| --- | :---: | :---: | --- |
| `.docx` | ✓ | ✓ | 主格式 |
| `.dotx` | ✓ | ✓ | 模板 |
| `.doc` | ✓ | ✗ | Word 97-2003 二进制，只读 |
| `.rtf` | ✓ | ✓ | 多节、页眉页脚、表格合并、图片、书签与超链接 |
| `.html` | ✓ | ✓ | 自包含单文件，CSS 与 Base64 图片内嵌 |
| `.odt` | ✓ | ✓ | ODF 1.3 |
| `.txt` | ✓ | ✓ | UTF-8，仅段落文本 |

7 种输入 × 6 种输出 = 42 种组合，扩展名和文件头都参与识别：

```csharp
WordFormatConverter.Convert("legacy.doc", "report.odt");
WordFormatConverter.Convert("report.rtf", "report.html");
WordDocumentFormat format = WordFormatConverter.DetectFormat("report.odt");   // Odt
```

跨格式转换是"结构等价"而非像素级还原：样式表、主题、复杂域、修订、数字签名不进入统一模型，转换后不保留。

## 八、修订跟踪、SmartArt 与签名检查

```csharp
OfficeDocument document = DocxReader.Load("review.docx");
foreach (Section section in document.Sections)
    foreach (Block block in section.Blocks)
        if (block is Paragraph p)
            foreach (Inline inline in p.Inlines)
                if (inline.Revision != null)
                    Console.WriteLine($"{inline.Revision.Type} {inline.Revision.Author} {inline.Revision.Date}");

int changes = WordRevisions.Count(document);
WordRevisions.AcceptAll(document);     // 或 RejectAll
```

插入 / 删除修订进入统一模型（行内用 `Inline.Revision`，段落标记存 `Paragraph.Revision`，被删正文放 `Paragraph.DeletedInlines` 且不参与排版），`DocxWriter` 能重新写出 `w:ins`/`w:del`/`w:delText`——也就是说可以"读修订 → 处理 → 带修订写回"。主动写修订也支持：给 `TextRun.Revision` 赋一个 `RevisionInfo` 即可。

`WordSmartArtReader` 提取 SmartArt 节点文本（用于内容检查和抽取，重新生成的包不保留图形结构）；`WordDigitalSignatureReader` 读签名 / 摘要 / 规范化算法、被引用部件与证书主体（只做结构检查，不验密码学有效性；且任何"读取后重写"都会让原签名失效，写出器不复制签名部件）。

## 九、批量整理、校验与统计

`WordNormalizer` 把多人协作、格式混乱的文档拉回统一标准，默认构造就按中文公文常用值打开了大部分开关：

```csharp
WordNormalizeOptions options = new WordNormalizeOptions
{
    RemoveEmptyParagraphs = true,
    BodyEastAsiaFont = "宋体",
    BodyAsciiFont = "Times New Roman",
    BodyFontSize = 12,
    BodyLineSpacing = 1.5,             // 倍数，不是磅值
    FirstLineIndentCharacters = 2,
    PromoteFakeHeadings = true,        // 把"加粗大字"提升为真标题
    FixHeadingLevelSkips = true,       // 修正 1 级直接跳 3 级
    NormalizeTables = true,
    RepeatTableHeaderRows = true,
    MaxImageWidth = 420
};

WordNormalizeResult result = WordNormalizer.Normalize(document, options);
Console.WriteLine(result.ToReport());   // 变更列表 + 问题列表 + 前后统计
```

`WordProcessor` 是"读文件 → 备份 → 整理 → 写出 → 出报告"的一键流水线，带 `Rollback`：

```csharp
WordProcessResult r = WordProcessor.Process("draft.docx", "final.docx", options);
Console.WriteLine(r.BackupPath);
WordProcessor.Rollback(r.BackupPath, "draft.docx");
```

配套还有 `WordConsistencyChecker`（空段落、字体字号混用、行距段距不一致、半角句末标点、控制字符、标题跳级、表格空行、图片超宽）、`WordDocumentDiff`（两版本文本差异）、`WordDocumentMetrics`（段落数 / 字数 / 表格数 / 各级标题数 / 字体使用分布）。

注意 `Normalize` **就地修改**传入的模型（`result.Document` 就是它本身），想试算请重新加载一份或先备份。

## 其它实用件

- **自定义文档属性**（Word「文件 → 信息 → 属性 → 高级属性 → 自定义」）：`document.Metadata.CustomProperties.Set("项目编号", "QF-2026-018")`，强类型读写 7 种 CLR 类型，读取兼容 `lpstr`/`bstr`/`ui4`/`r4` 等变体，只写进 DOCX/DOTX 的 `docProps/custom.xml`。
- **文档克隆与合并**：`WordDocumentOperations.Clone` 深拷贝；`Append(target, source)` 按节追加并保留页面设置与页眉页脚；`Merge(report1, report2, report3)` 合成一份。

## 已知边界

图片按 PNG / JPEG / GIF / BMP / TIFF / EMF / WMF 识别并在 DOCX 包内原样往返（转 PDF 时只有 PNG / JPEG 会光栅化）；不处理批注、脚注、尾注、嵌入对象与 VBA；模型没有分栏、节类型与“首页不同”页眉页脚，读取时忽略；不支持写 `.doc`；移动修订（`w:moveFrom`/`w:moveTo`）读取为删除 / 插入；不支持表格行 / 单元格 / section 级修订与修订批注气泡；接受 / 拒绝修订不重建原始字符位置；查找替换只处理普通文本运行，不搜域显示文本；自定义属性只在 DOCX/DOTX 往返。

## 小结

`Wisdom.Word` 覆盖面比"生成个 docx"大得多：流式构建、书签、跨 Run 替换、邮件合并（含嵌套区域）、标签纸 / 信封、7 格式互转、修订处理、批量整理与校验。放弃的部分都很明确——批注、VBA、样式表继承、像素级分页。

如果你的场景是"合同 / 报告 / 表单套打"这类结构化文档生成，而不是"100% 还原设计师排好的复杂版式"，这个包基本够用，而且完全没有许可证和体积负担。
