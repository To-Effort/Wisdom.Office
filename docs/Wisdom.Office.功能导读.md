# 零依赖 C# Office 库到底怎么用——四个场景讲透 Wisdom.Office

这是一篇"功能导读"，不是 API 参考，也不是分步教程。参考文档回答"有什么"，教程回答"照着做一遍"，而这篇想回答的是另一个问题：**手上有真实业务需求时，该从哪个入口进来、四个场景之间怎么选、以及哪些行为只有真正跑过才知道。**

文中的代码全部按仓库里的公开签名写成，数字全部是实测输出（可在 `docs/` 下七份分步教程里复现）。

先给出这个库的画像：

- 纯 C# 自研，**不依赖任何第三方 NuGet 组件**，也不依赖 Microsoft Office 或 WPS。ZIP、XML、OLE2 复合文档、PDF、OOXML、BIFF8 都是自己实现的。
- 覆盖 PDF（读+写）、DOCX（读+写）、DOC（读）、XLSX（读+写）、XLS（读）、PPTX 与 POTX（读+写）、Markdown（读+写）。
- 每个包同时提供六份资产：`net461`、`netstandard2.0`、`net7.0`、`net8.0`、`net9.0`、`net10.0`，源码里没有一处 `#if NET*` 条件编译。
- 四个目标场景：报表导出、文档互转、模板填充、通用文档模型。PPTX 与 Markdown 不是凭空多出来的两个场景——Markdown 汇入统一文档模型，PPTX 自带演示模型，两者都从"文档互转"这个入口进来。

---

## 0. 先把层次看清楚，再决定装哪个包

七个包的分工是纵向的，不是横向平铺的：

```text
Wisdom.Office.Core   统一文档模型 + 工作簿模型 + 演示模型 + 模板引擎 + OLE2/ZIP/压缩基础设施
Wisdom.Pdf           PDF 1.7 生成与解析、TrueType 字体子集嵌入、文本与表格提取、对象级编辑（增删改对象/合并拆分/水印/AcroForm 填表与扁平化）、标准安全处理器加密读写
Wisdom.Word          DOCX 读写、模板填充、.doc 读取、文档整理/校验/差异/统计
Wisdom.Excel         XLSX 读写、模板填充、图表/条件格式/锚定图片、.xls(BIFF8) 读取
Wisdom.Ppt           PPTX / POTX 读写、版式与占位符、备注/切换/背景/自选图形、入场动画、主题深改
Wisdom.Markdown      Markdown 解析与渲染、front-matter、GFM 表格与任务列表、HTML 出口
Wisdom.Office        门面：格式检测、格式互转、Excel/PowerPoint→Word/PDF 排版、Report 报表 DSL
```

依赖方向只有一个：上面吃下面，`Core` 谁都不吃。所以选型规则很简单：

- **只做报表导出或格式转换**：装门面包 `Wisdom.Office` 就够了，它会传递带上其余六个。
- **只写 Excel**（服务里生成台账、看板）：单装 `Wisdom.Excel`，它只依赖 `Core`。
- **只解析 PDF 抽字段、填表单、盖章加水印**：单装 `Wisdom.Pdf`。
- **只生成或解析 PPTX**（内部述职、自动生成宣讲材料）：单装 `Wisdom.Ppt`，同样只依赖 `Core`。
- **只把 Markdown 转成 HTML / 反解析成文档模型**：单装 `Wisdom.Markdown`。
- **想自己控制排版**（比如做一套企业内部的版式规范）：直接用 `Wisdom.Office.Core` 的模型，把渲染交给 `Wisdom.Office` 的 `OfficeConverter.ToPdf/ToWord`。

安装始终是一条命令（版本与门面包保持一致，当前 `1.0.0`）：

```bash
dotnet add package Wisdom.Office
```

有一个容易被忽略的点：**六份资产由同一份源码编译**。这意味着升级到 .NET 11 时不必重新打包（NuGet 会选中 `net10.0` 或向前回退），也意味着 .NET Framework 4.6.1 上的行为和 .NET 10 上完全一致——这一点不是推论，端到端测试套件在 `net8.0` 和 `net461` 上各跑一遍，用同一套断言。

---

## 场景一：报表导出——一份数据，三种格式

这是最常进来的入口，也是这个库把"复杂度收起来"做得最彻底的地方。

### 1.1 数据只描述一次

`Report` 是一个构建在统一模型之上的轻量 DSL，它自己不做排版引擎，只描述"有哪些表、每列多宽、什么格式、怎么对齐"，导出时由对应格式的引擎去解释：

```csharp
using System;
using System.Collections.Generic;
using Wisdom.Office;
using Wisdom.Office.Core;

record Sale(string Product, int Qty, decimal Amount, DateTime Date);

List<Sale> sales = new List<Sale>
{
    new Sale("无线键盘", 12, 3594.00m, new DateTime(2026, 7, 5)),
    new Sale("机械鼠标", 30, 2697.00m, new DateTime(2026, 7, 6)),
    new Sale("27 英寸显示器", 5, 9495.00m, new DateTime(2026, 7, 7)),
};

Report report = new Report("季度销售报表");
report.Subtitle = "2026 年第三季度";
report.Author = "示例公司";
report.FooterText = "由 Wisdom.Office 自动生成";
report.GeneratedAt = new DateTime(2026, 10, 21, 9, 30, 0);
report.ShowGeneratedTime = true;

ReportTable table = report.AddTable("产品销售明细");
table.AddColumn("产品", 120, null, HorizontalAlignment.Left);
table.AddColumn("数量", 60, "0", HorizontalAlignment.Right);
table.AddColumn("金额", 100, "#,##0.00", HorizontalAlignment.Right);
table.AddColumn("日期", 90, "yyyy-MM-dd", HorizontalAlignment.Center);

foreach (Sale sale in sales) table.AddRow(sale.Product, sale.Qty, sale.Amount, sale.Date);

report.AddNote("金额单位：人民币元，均为含税价。");
```

`AddColumn(header, width, numberFormat, alignment)` 里的宽度是**磅**，`numberFormat` 用的是 Excel 那套格式串（`#,##0.00`、`yyyy-MM-dd`、`0;-0;"—"` 都支持），这意味着同一份报表定义在三种格式里的数字呈现是一致的——它不是"三个导出器各自猜"。

### 1.2 样式必须在导出之前定完

外观走 `report.Style`，一份样式同时作用于三种格式：

```csharp
report.Style.FontFamily = "Microsoft YaHei";
report.Style.BaseFontSize = 9;
report.Style.HeaderFill = OfficeColor.FromHtml("#1F4E79");
report.Style.HeaderTextColor = OfficeColor.White;
report.Style.AlternateRowFill = OfficeColor.FromHtml("#F2F7FC");
report.Style.ZebraStripes = true;

report.Export("out/sales.pdf");     // 按扩展名分派
report.Export("out/sales.docx");
report.Export("out/sales.xlsx");

byte[] pdfBytes = report.ToPdf();   // 或者直接拿字节
byte[] docxBytes = report.ToWord();
byte[] xlsxBytes = report.ToExcel();
```

`ToDocument()` / `ToWorkbook()` / `ToPdf()` 每次调用都从当前定义重新构建。所以顺序错了不会报错，只会静悄悄地出一份没样式的产物——这是这个入口唯一的"反直觉"之处，值得写进团队的 review 清单。

### 1.3 导出物各自落在哪里

`Report` 替你做掉的事，按格式分别是：

- **PDF**：标题、副标题、生成时间、表题、表头底纹、斑马纹、边框、备注，表格跨页时表头自动重复。
- **DOCX**：同一套内容映射成段落与表格，表头行标成 `IsHeader` 且 `RepeatHeaderRow = true`。
- **XLSX**：竖排的记录区，并给数据区加冻结窗格。

回读 XLSX 时会看到一个固定的行号结构（实测）：

```text
第 1 行  季度销售报表        标题
第 2 行  2026 年第三季度     副标题
第 3 行  生成时间 …          ShowGeneratedTime 打开时
第 4 行  （空行）
第 5 行  产品销售明细        表题
第 6 行  产品 数量 金额 日期  表头
第 7 行  无线键盘 …          第一条数据
```

`back.GetCell(7, 3).Value` 取回来是数值 `3594` 而不是 `"3,594.00"`——`decimal` 在写入时被转成 `double` 存成数字单元格，格式串只影响显示。写校验脚本时这点很关键：**你要断言的是值，不是文本**。

另外两条实测行为：

- DOCX 侧：当文档以表格结尾时，写出器会补一个空段落。OOXML 要求 body 以 `w:p` 结束，这是合规代价，不是 bug；但回读断言块数时会多出一块。
- PDF 侧：表头跨页重复是可以直接验证的。`docs/Wisdom.Pdf.Tutorial.md` 里那份长表格的实测结果是 `pages=2`、`hasLastRow=True`、`pagesWithHeader=2/2`，页脚每页都能匹配到 `第 N /`。

### 1.4 需要超出 Report 的表达力时

`Report` 故意不提供页眉页脚细节、图文混排、跨列合并。真要那些，就取中间模型继续手工加工：

```csharp
OfficeDocument document = report.ToDocument();
document.Sections[0].Blocks.Add(new PageBreak());   // 备注另起一页

OfficeWorkbook workbook = report.ToWorkbook();
workbook.Worksheets[0].SetColumnWidth(1, 28);       // 默认列宽可以再调
```

这一步之后你就脱离了"报表 DSL"，进入场景四的统一模型——但两者之间没有断层，`Report` 产出的就是同一套模型。

---

## 场景二：文档互转——按内容认格式，按选项排版

### 2.1 先别管扩展名

`OfficeFile.Load` 用文件内容签名判断格式，扩展名写错照样能读：

```csharp
OfficeFile file = OfficeFile.Load("out/sales.xlsx");

Console.WriteLine(file.Format);          // Xlsx
Console.WriteLine(file.Metadata.Title);  // 季度销售报表
Console.WriteLine(file.ExtractText());   // 纯文本摘要

file.Save("out/sales-copy.xlsx");        // 同格式回存
byte[] pdfBytes = file.ToArray(OfficeFileFormat.Pdf);
```

只想路由、不想构造完整模型时用 `OfficeFileFormats.Detect(path)`，返回 `Pdf / Docx / Doc / Xlsx / Xls / Pptx / Unknown`。这一步很轻，适合放在上传接口的前面做分派。

但有一个例外必须知道：**Markdown 是纯文本，没有内容签名可嗅**。实测 `Detect("article.md")` 返回 `Unknown`，而 `OfficeFile.Load("article.md")` 返回 `Markdown`——因为路径版 `Load` 在嗅探失败后会退回扩展名判断。推论是：`Load(byte[])` 与 `Load(Stream)` **不接受 `.md`**（前者抛 `NotSupportedException`，后者根本没有格式重载），走字节的上传接口要自己按 `Text` 分支处理，或改用 `Markdown.ToDocument(text)`。

按格式拿底层模型：

| 属性 / 方法 | 适用格式 | 得到 |
| --- | --- | --- |
| `file.Workbook` | Xlsx、Xls | `OfficeWorkbook` |
| `file.Document` | Docx、Doc、Pdf、Markdown | `OfficeDocument` |
| `file.Presentation` | Pptx | `OfficePresentation` |
| `file.Pdf` | Pdf | `PdfDocument` |
| `file.AsDocument()` | Xlsx、Xls | 按默认排版映射出的 `OfficeDocument` |
| `file.AsDocument(options)` | Xlsx、Xls | 按指定排版选项映射出的 `OfficeDocument` |
| `file.AsPresentation()` | Docx、Doc、Pdf、Markdown、Xlsx、Xls | 按标题拆页映射出的 `OfficePresentation` |

### 2.2 转换矩阵，以及它为什么长这样

`OfficeConverter.Convert(源, 目标)` 的实测矩阵（目标格式由扩展名决定，同一份源码、七种可读格式）：

| 源 \ 目标 | PDF | DOCX | XLSX | PPTX | Markdown |
| --- | --- | --- | --- | --- | --- |
| Markdown | 通过 | 通过 | 不支持 | 通过 | 通过 |
| DOCX / DOC | 通过 | 通过 | 不支持 | 通过 | 通过 |
| PDF | 通过 | 通过 | 不支持 | 通过 | 通过 |
| XLSX / XLS | 通过 | 通过 | 通过 | 通过 | 通过 |
| PPTX | 通过 | 通过 | 不支持 | 通过 | 通过 |

**只有"非电子表格 → XLSX"这一格会抛异常**，消息实测是 `Word/PDF 文档不能直接保存为 Excel 工作簿。`——文档模型里没有"单元格地址"这个概念，反向映射会丢语义，所以不提供。其余 24 格都是先汇入统一模型（`OfficeDocument` / `OfficeWorkbook` / `OfficePresentation`）再排版，所以任何两格之间的产物都共享同一套排版规则。

另外三条实测边界：

- **`.doc` / `.xls` 只能作为源。** 旧版二进制格式只实现了解析，写回会抛 `NotSupportedException`。而且 `.xls` 的公式**只保留缓存值**——用 `XlsReader` 读一份优化器求解结果工作簿，实测 `公式单元格数 = 0`，所有数值都是 BIFF8 里存的最后计算结果。指望读回公式再重算的，要另找方案。
- **PDF → PPTX 是通的，但只出一页。** 走的是 `Pdf → OfficeDocument → OfficePresentation`，拆页依据是标题层级，普通 PDF 没有标题就落在单页里。
- **Markdown 的目标只有三种可回头**：`md → pdf/docx/pptx/md`。反向 `xxx → md` 也全通（实测 `data.xlsx → md` 只有 63 字节——电子表格进 Markdown 就是一张 GFM 表），但 `.md` 作为源没有内容签名，见 2.1。

### 2.3 Excel → PDF/Word：真正需要你决策的地方

工作簿没有"页"的概念，`WorkbookToDocumentConverter` 必须替你决定怎么折。这就是 `WorkbookLayoutOptions` 存在的原因，它的默认值本身就是一份"常规打印"的策略：

| 字段 | 默认 | 作用 |
| --- | --- | --- |
| `IncludeSheetNames` | `true` | 每个工作表前输出一级标题（表名） |
| `IncludeHiddenWorksheets` | `false` | 隐藏工作表是否参与转换 |
| `UsePrintSetup` | `true` | 沿用纸张、方向、页边距、打印区域、分带 |
| `UseGridLines` | `true` | 用网格线代替逐格边框 |
| `PageSetup` | `null` | 非 null 时统一覆盖所有工作表的页面设置 |
| `FontFamily` / `FontSize` | `"Arial"` / `9` | 正文字体族与字号 |
| `GridColor` | 浅灰 `#D0D0D0` | 网格线颜色 |
| `EmulateNumberOverflow` | `true` | 列宽不足时按 Excel 规则输出 `####` |

横向大表的推荐组合不是在转换时硬调，而是**把排版意图留在源工作簿里**，转换时沿用：

```csharp
using Wisdom.Excel;

OfficeWorkbook data = XlsxReader.Load("out/sales.xlsx");

Worksheet sheet = data.Worksheets[0];
sheet.PrintSetup.IsLandscape = true;
sheet.PrintAreaFirstRow = 1;
sheet.PrintAreaLastRow = 60;
sheet.PrintAreaFirstColumn = 1;
sheet.PrintAreaLastColumn = 16;

OfficeConverter.WorkbookToPdf(data, "out/wide.pdf", new WorkbookLayoutOptions
{
    UsePrintSetup = true,
    IncludeSheetNames = true,
    IncludeHiddenWorksheets = false
});
```

### 2.4 三个高频"为什么输出不对"

**中文变方块。** PDF 生成需要从系统字体目录嵌入 TrueType 字体，且只支持 `glyf` 轮廓，不支持 CFF/OpenType（`OTTO`）。Windows 默认可用；精简容器里请装 Noto 或文泉驿，或用 `Wisdom.Pdf` 的 `FontResolver` 显式注册字体文件。字体族名匹配会做归一化——`FontResolver.NormalizeFamily("  Microsoft  YaHei ")` 得到 `microsoftyahei`（去空格、转小写），所以传 `"Microsoft YaHei"` 和 `"microsoftyahei"` 是同一件事。

**数字变成 `####`。** 这是 `EmulateNumberOverflow = true` 时按 Excel 规则的**模拟输出**，属于正确行为。要么加宽源列，要么把该选项设为 `false` 直接输出文本。注意区分：`Cell.DisplayText()` 本身**不会**输出 `####`，它只负责按数字格式渲染文本；`####` 是排版层的决定。

**隐藏内容忽隐忽现。** 两层默认值方向相反，这是最容易踩的：`ExcelReadOptions` 的 `IncludeHiddenWorksheets / IncludeHiddenRows / IncludeHiddenColumns` **默认全为 true**（读的时候什么都给），而 `WorkbookLayoutOptions.IncludeHiddenWorksheets` **默认 false**（转 Word/PDF 时丢掉隐藏表）。真实考勤报表里这一点很具体：那份工作簿的第一个工作表是隐藏的取数区，`1556 行 x 15 列`，关掉转换侧的隐藏表开关才能只拿到 4 张业务表。

### 2.5 PDF → Word 是"重新排版"，不是"还原"

`OfficeConverter.Convert("scan-output.pdf", "out/summary.docx")` 走的是解析文本与表格、再按统一模型重排。产出的是"内容 + 大致结构"的可编辑文档，不追求与原始 PDF 逐像素一致。表格识别基于文本位置与线条，属启发式，合并单元格、无边框表、跨页表需要按坐标自行校正。

提取侧还有一条实测细节会影响你的断言写法：`PdfTextLine.Text` 会把同一行的多个单元格用空格连起来，例如 `序号 巡检项 结果 用时(min)`。按整段原文 `==` 比较会失败，应该用 `Contains` 或者逐片段（`Fragments` 保留各自坐标与字号）判断。

如果目标不是"换个格式重排"，而是**在原文件上动刀且不许版式变**，就不要走 `OfficeConverter`，走 `PdfFile`：它直接读写交叉表，按对象号增删改，`SetPageRotation` / `AddWatermark` / `Merge` / `Split` 之后 `Save` 整量重写或 `SaveIncremental` 只追加改动对象（原字节保留）。水印是 Form XObject，用 `/Rotate` 反向补偿矩阵贴合页面旋转，图层为 `Background` 时挂 `PrependPageContent`、`Foreground` 时挂 `AppendPageContent`，透明度走 `/ExtGState` 的 `CA`+`ca`；只有当页面带 `/StructParents` 时才补 `/Artifact` 标记，避免给没有结构树的 PDF 写出幽灵 MCID。

填表也在这一层：`HasForm` / `GetForm()` 读字段树（`Fields` 是文档顺序的全部节点，含中间层父域；`GetField("父.子")` 按点号名查），`FillForm(values)` 批量写 `/V` 并统一生成 `/AP /N` 外观（中文按 Identity-H 子集嵌入并带 ToUnicode，所以阅读器显示与文本提取都读得出），`FlattenForm()` 再把外观按控件矩形画进页面内容、清掉控件注释与 `AcroForm`。写值时 `/NeedAppearances` 自动置 `true`、生成完外观复位，"值已改外观还是旧的"这个中间态不会显示错值。单选组的值只记在父域 `/V`，各选项仅同步自己的 `/AS`；密码域、没有 `/AP` 的字段、控件没挂进 `/Annots` 的字段以及签名域与普通按钮都不参与扁平化，值不会被静默丢掉。带 `/XFA` 的混合表单只能整体 `RemoveFormXfa()` 剥离，本库不解析 XFA 模板。

### 2.6 服务场景：全程字节

```csharp
// 单向、无需排版选项
public static byte[] ToDocx(byte[] sourceDocx)
{
    return OfficeConverter.Convert(sourceDocx, OfficeFileFormat.Docx);
}

// Excel -> PDF 且要控制排版
public static byte[] BuildPdf(byte[] sourceXlsx)
{
    OfficeWorkbook workbook = XlsxReader.Load(sourceXlsx);
    return OfficeConverter.WorkbookToPdf(workbook, new WorkbookLayoutOptions
    {
        IncludeSheetNames = true,
        UseGridLines = true
    });
}
```

这里有个容易写错的签名事实：`OfficeConverter.Convert` 的字节重载**只有** `(byte[], OfficeFileFormat)` 一个，不接受排版选项。要带 `WorkbookLayoutOptions` 走字节路径，就先 `XlsxReader.Load(bytes)` 再 `WorkbookToPdf(workbook, options)`；文件路径版才有 `(sourcePath, targetPath, layout)` 三参重载。

### 2.7 PPTX 与 Markdown：新入口的两条真实边界

**Markdown 是"最短路径"的格式。** `Markdown.ToDocument(text)` 直接产出统一文档模型，后面接什么导出器都行；实测那份 `article.md` 一次 `Convert` 出 PDF 72168 字节、DOCX 4052 字节、PPTX 10048 字节，`→ md` 只有 258 字节。它落到模型的形式都有明确对应（标题→`HeadingLevel`、GFM 表格→`Table` + `IsHeader`、任务列表→`ListMarker` 前缀 `☐`/`☑`、front-matter→`Metadata` 的 `title`/`author`/`date`），但也有三处**有意的降级**：引用块模型里没类型，只剩斜体加缩进，回写 Markdown 时引用标记丢失；分割线落成居中写着 `＊ ＊ ＊` 的段落（段落模型没有下边框属性），HTML 与 Markdown 出口按同一常量还原成 `<hr />` 和 `---`，Word/PDF 里排的就是这三个字符；链接的 `title` 字符串会被识别但模型没有字段，直接丢弃。

**PPTX 的"版式"是推断出来的，不是存下来的。** OOXML 里没有"版式类别"字段，`PptxReader.InferLayoutKind` 只能按形状组合猜。实测两个方向的偏差：`AddSlide(TitleAndTable)` 生成的表格占位符按 `obj` 类型写出，回读变成 `TitleAndTable → TitleAndContent`；反过来在 `AddSlide(TitleAndContent)` 上手工 `AddTable(...)`，回读会被猜成 `TitleAndTable`。七种版式逐个写读的实测对照是：

```text
写出 = Title, TitleAndContent, SectionHeader, TwoContent, TitleAndPicture, TitleAndTable, Blank
回读 = Title, TitleAndContent, Title,         TwoContent, TitleAndPicture, TitleAndContent, Blank
```

错的两格都是"只带标题的形状被当成 Title"和"表格/副标题占位符被并进 TitleAndContent"，占位符身份本身（`Title#0`、`Content#1`、`Content#2`、`Picture#1`、`Table#1`）和几何（实测 `Left = 60` 原样读回）没有丢。**所以别把 `slide.Layout` 当断言目标**，要断言形状类型、占位符索引和文本。

同一轮实测里，这些是**按原值往返**的：备注（`开场 2 分钟。`）、切换类型与时长（`Fade`/1.5、`Push`/2.25）、背景色（`#F2F7FF`）、自选图形（`RoundedRectangle` 填充 `#00B050`、`CornerRadius` 默认 0.15）、图片字节与 alt 文本（70 字节的 1x1 PNG、alt `示意图`、120x120）、入场动画（`FlyIn`/0.8s 时长 + 0.5s 延迟，回读一字不差）、`AdvanceAfterSeconds`（3 秒回读 `3s`，写成 `p:transition/@advTm="3000"`）。唯一要留心的是表格：`Table.Columns` 的列定义写出时不进 XML（实测写出为"2 行 0 列"），回读按每行单元格数补齐成"2 行 2 列"——列宽信息在 PPTX 里是逐单元格的 `a:gridCol`，模型那份列定义拿不回来。

**动画只有入场，而且延迟只对第一个形状有效。** `SlideAnimationKind` 是 `None` / `Fade` / `FlyIn` / `Wipe` / `Zoom` / `Split` 六个值，写出的是 PowerPoint 自己的 `p:timing` 主序列（presetId 10 / 2 / 22 / 23 / 26），所以放映时能看到原生效果；强调、退出、路径与媒体动画模型里没有位置，读第三方 PPTX 时也只按 `presetClass="entr"` 挑入场节点。`AnimationDurationSeconds = 0` 不是"没动画"而是"按类别取默认"（FlyIn / Zoom 各 1 秒，其余 0.5 秒）。`AnimationDelaySeconds` 更硬：写出器把第一个动画形状写成 `afterEvent`（配合本页 `AdvanceAfterSeconds`），其余形状一律 `indef`，实测两个形状都设 `1.2s` 时回读是 `1.2 / 0`。

**主题深改只开到色板与字体这一层。** `OfficePresentation.Theme`（`PresentationTheme`）覆盖 `theme1.xml` 的六个阶段（dk1 / lt1 / dk2 / lt2 / hlink / folHlink）与 `majorFont` / `minorFont` 的 latin 字体名，母版和版式经 `p:clrMap`（`bg1→lt1`、`tx1→dk1`）自动跟着变，不用逐个形状改填充；强调色仍走 `AccentColors`。读回时六个阶段与两个字体**全部显式赋值**，没设过主题的普通 PPTX 也读得出默认值（实测 `名=Office，dk1=#000000，lt1=#FFFFFF，主字体=Calibri，次字体=Calibri`），所以 `Theme.Text` 可以直接 `.Value`。母版里更深的东西（占位符精确坐标、效果/填充样式列表）仍由本库固定生成。

**`.potx` 与 `.pptx` 只差一个内容类型。** `PptxWriter.WriteTemplate(...)` 写出的包，`/ppt/presentation.xml` 的 `Override` 换成 `…presentationml.template.main+xml`，关系类型仍是 `officeDocument`；`OfficeFileFormats.Detect` 就是读这一处区分二者，实测识别为 `PowerPoint 模板`、`OfficeFile.Load` 后 `AsPresentation()` 正常拿到页表，按扩展名 `Save("deck.pptx")` 就转回普通演示文稿。普通 `.pptx` 不会被误判成模板。

Word/Excel/Md → PPTX 走 `file.AsPresentation()`，规则是**按标题层级拆页**：实测那份三个标题的 DOCX 出来正好 3 页、`960x540`、版式全是 `TitleAndContent`，表格降级成 `[表格 3x2]` 文本占位。反向 PPTX → Markdown 则是每页一个 `#` 标题加正文，表格还原成 GFM 管道表（实测 97 字符）。

---

## 场景三：模板填充——让业务方改模板，不改代码

前两个场景是"代码决定版式"，模板填充反过来：**版式冻结在模板文件里，代码只喂数据。** 这对合同、通知单、检验报告这类"格式是法务/质量部门定的"的场景更合适。

### 3.1 语法只有这么点

```text
{{key}}                 简单值
{{key:format}}          带格式串（数字、日期同 Excel 那套）
{{obj.prop}}            点路径
{{#each list}} ... {{/each}}   循环
```

循环体内可用 `@index`（从 0）、`@number`（从 1）、`@first`、`@last` 和 `.`（当前项本身）。**不支持 `[i]` 下标**——要按下标取，先把列表投影成需要的形状。另外，`TemplateEngine.Replace()` 只做标量替换，遇到 `{{#each}}` 原样返回标记；展开循环是 `DocxTemplate` / `XlsxTemplate` 的事，所以标记必须各自独占一段（Word）或一格一行（Excel）。

看起来太简单了，但"跨 Run"这件事只有自研解析才处理得干净：Word 会把一段文字按格式切成若干 `w:r`，`{{customerName}}` 经常被拆成 `{{`、`customer`、`Name}}` 三个 run。模板引擎在文本层面重组占位符，所以模板里手打的占位符不需要"保证在一个 run 内"。

### 3.2 Word：段落级与表格行循环

模板长这样：

```text
{{customerName}}，您好：
我方确认以下订单：
{{#each items}}
{{name}}　数量 {{qty}}　金额 {{amount:0.00}}
{{/each}}
合计：{{total:0.00}} 元　签订日期 {{signDate:yyyy-MM-dd}}
```

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

段落级 `{{#each}}` 会把标记所在段落整体换成 N 份内容，列表为空时循环块整块消失。表格行循环写法相同，把 `{{#each}}` / `{{/each}}` 放在循环行的首列与末列。`DocxTemplate.Fill` 提供文件、字节、流三种形态，字节版直接返回填充后的 `byte[]`。

### 3.3 Excel：行级循环，样式与合并跟着复制

Excel 侧循环以"行"为单位，标记可以跨一行或多行：

```csharp
sheet.GetCell(4, 1).Value = "{{#each items}}{{name}}";
sheet.GetCell(4, 2).Value = "{{qty}}";
sheet.GetCell(4, 3).Value = "{{amount:0.00}}";
sheet.GetCell(4, 4).Value = "{{@number}}";
sheet.GetCell(4, 5).Value = "{{/each}}";
sheet.GetCell(6, 1).Value = "合计";
sheet.GetCell(6, 3).Value = "{{total:0.00}}";
```

三条数据填充后（实测）：

```text
3: 品名 | 数量 | 金额 | 序号
4: 无线键盘 | 2 | 398.00 | 1
5: 机械鼠标 | 5 | 449.50 | 2
6: 清洁服务 | 1 | 420.50 | 3
7:  |  |  |
8: 合计 |  | 1268.00 |
lastRow=8
```

值得注意的两件事：模板行的**样式、行高、列宽、合并区域会逐份复制**，所以"金额列右对齐 + 千分位"只需在模板里做一次；`合计` 那行整体下移，地址不需要在代码里重算。列表为空时循环块删掉、后续行上移，同一份模板填零条数据得到 `lastRow=5`，`A5` 就是原来的 `合计`——**空列表不会留下一行带 `{{}}` 的残骸**，这对自动生成"零订单客户"对账单很重要。

内存版重载返回**新**工作簿，模板实例不被污染，可以在服务里加载一次模板、并发填充多次：

```csharp
OfficeWorkbook inMemory = XlsxTemplate.Fill(XlsxReader.Load("out/order-template.xlsx"), context);
```

实测对照：模板文件里的 `A1` 仍是 `客户：{{customerName}}`，结果里的 `A1` 是 `客户：示例客户`。

### 3.4 TemplateOptions 的三个开关各管一类事故

- `LeaveUnresolvedTokens`：默认 `true`，未给值的占位符原样留在文档里。给"人工补录"的流程用 `true`；给客户看的产物必须 `false`，否则 `{{invoiceNo}}` 会印在合同上。
- `HtmlEncodeValues`：值里含 `<`、`&` 时是否需要转义。打开它，转义由引擎负责，别在业务层重复处理。
- `Culture`：决定 `{{amount:0.00}}` 的小数点与千分位分隔符。跨环境一致输出请固定 `InvariantCulture`。

---

## 场景四：通用文档模型——前三个场景脚下的东西

`Report` 和转换器都是模型的消费者。当你需要的版式超出 DSL 能表达的范围（不规则表头、页眉里的公司标识与页码域、跨列合并、图注混排），就直接写模型。

### 4.1 文档侧：三层结构 + 三种块

```text
OfficeDocument
  └─ Section（PageSetup、Header/Footer）
       └─ Block：Paragraph | Table | PageBreak
            └─ Paragraph.Inlines：TextRun | FieldRun | ImageRun
```

```csharp
OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "示例文档";

Section section = document.AddSection(PageSetup.A4());
section.Blocks.Add(new Paragraph("标题").AddText(" - 副标题", new TextFormat("Arial", 10)));

Table table = new Table();
table.Columns.Add(new TableColumn(160));
table.Columns.Add(new TableColumn(100));
table.AddRow("项目", "金额");
TableRow row = table.AddRow("合计", "2,288.40");
row.IsHeader = false;
section.Blocks.Add(table);

OfficeConverter.ToPdf(document, "document.pdf");
OfficeConverter.ToWord(document, "document.docx");
```

几个绕不开的约定：

**页码是域，不是文本。** 用 `FieldRun`：

```csharp
new FieldRun("PAGE", "{page}", new TextFormat("Arial", 9))
```

Word 侧写成 `w:fldSimple`，由 Word 打开时计算；PDF 侧由生成器把 `{page}` / `{pages}` 替换成实际数字。同一份模型因此在两种格式里都有页码，而不是一句死的"第 1 页"。

**合并单元格的覆盖格要手动少放。** `TableCell.ColSpan / RowSpan` 只描述跨度，模型不会替你校验：`RowSpan = 2` 的单元格，它覆盖到的下一行**必须不再为该列放单元格**。写错不会抛异常，后果是 DOCX 的合并结构与 Word 里显示的不一致——这是最容易在上线后才被发现的一类缺陷。想知道某个坐标被哪个合并区覆盖，用 `FindMergeAnchor(row, column)` 取回锚点单元格；区域字符串则用 `SpreadsheetRange.TryParse("台账!$B$2:$D$9", out firstRow, out lastRow, out firstColumn, out lastColumn)` 解析（实测 `2,9,2,4`）。

**横向是交换宽高。** `PageSetup.IsLandscape = true` 把 `595.28 x 841.89` 变成 `841.89 x 595.28`（A4 的磅值），它不是渲染时的旋转标记。

**单位换算工具在 Core 里。** `UnitConverter.PointsToTwips(72) = 1440`（Word 的表宽与边距用 twips）、`UnitConverter.PointsToEmu(72) = 914400`（Drawing 锚定用 EMU）。模型层统一用磅，落盘时各自换算。

### 4.2 工作簿侧：模型不是二维数组

`OfficeWorkbook` / `Worksheet` / `Cell` / `CellStyle` 看着直观，但有四条语义必须提前知道，否则读回来的结果会和你预期的不一样：

- **`GetCell(r,c)` 会创建，`FindCell(r,c)` 只探测。** 用 `GetCell` 遍历稀疏表会凭空造出大量空单元格。
- **`LastRow` / `LastColumn` 把合并区和"手工创建但没赋值"的空单元格算进去。** 空表 `Sheet1` 的 `LastRow` 是 0，但被你 `GetCell` 过的区域会算数。
- **`DisplayText()` 认数字格式的分节语义**，包括 `正;负;零;文本` 四段式。`12800.50` 配 `#,##0.00` 显示成 `12,800.50`，配 `"0;-0;\"—\""` 时零值显示成 `—`；`[Red]` 之类的颜色指令只影响显示，值本身仍是数字。
- **列宽单位是 Excel 的"字符宽"，不是磅。** 要和页面排版对齐就用 `GetColumnWidthPoints` 换算。

### 4.3 演示侧：版面是算出来的，不是画出来的

第三个模型是 `OfficePresentation`（同样定义在 `Core`），它和前两个的最大差别：**你只决定"这一页放什么内容"，坐标由版式规则算**。

```csharp
using Wisdom.Ppt;
using Wisdom.Office.Core;

OfficePresentation deck = OfficePresentation.Widescreen();   // 960 x 540 磅
deck.Metadata.Title = "季度经营回顾";

Slide cover = deck.AddSlide(SlideLayoutKind.Title);
// Title 版式只有标题占位符；GetTitle() 返回的是纯文本，填内容要取形状
((TextBoxShape)cover.Shapes[0]).Paragraphs[0].AddText("2026 年第三季度");
cover.NotesText = "开场 2 分钟。";
cover.Transition = SlideTransitionKind.Fade;
cover.TransitionDuration = 1.5;

Slide body = deck.AddSlide(SlideLayoutKind.TitleAndContent);
((TextBoxShape)body.Shapes[0]).Paragraphs[0].AddText("业务进展");
((TextBoxShape)body.Shapes[1]).Paragraphs[0].AddText("完成报表引擎");

Slide free = deck.AddSlide(SlideLayoutKind.Blank);
free.AddShape(AutoShapeKind.RoundedRectangle, 60, 60, 200, 80, "圆角矩形");
free.AddImage(OfficeImage.FromFile("out/article.png"), 320, 60, 120, 120);
free.BackgroundColor = OfficeColor.FromHtml("#F2F7FF");

PptxWriter.Write(deck, "review.pptx");
```

几何规则是固定比例（`Standard` 是 720x540，`Widescreen` 是 960x540）：页边距 = 宽度的 8%，标题距顶 = 高度的 6%、占 20%；`SectionHeader` 的标题下移到 34%、占 26%；内容区从 30% 起。占位符在页内的身份是显式的——实测七种版式的形状序列为 `Title=[Title#0]`、`TitleAndContent=[Title#0 Content#1]`、`TwoContent=[Title#0 Content#1 Content#2]`、`TitleAndPicture=[Title#0 Picture#1]`、`TitleAndTable=[Title#0 Table#1]`、`Blank=[]`。想拿精确坐标只能自己 `AddShape(...)` 传绝对位置（实测写进去 `Left=60`，回读仍是 60）。

默认字体也在这个模型里定死：标题 Calibri 40 粗体、正文 Calibri 18。单个形状换字体就在 `TextBoxShape` 的 `Paragraph` 上带 `TextFormat`；整份稿子的主题字体（`theme1.xml` 的 `majorFont` / `minorFont`）由 `OfficePresentation.Theme` 的 `FontMajor` / `FontMinor` 覆盖，两者互不影响——形状上带显式 `TextFormat` 时以它为准。

### 4.4 Word 整理引擎：模型之上的批处理

这是场景四的一个延伸，也是实际项目里最常缺的一块：把业务方上传的"脏草稿"批量整理成规范文档。

```csharp
WordNormalizeOptions normalize = new WordNormalizeOptions
{
    MergeShortParagraphs = false,          // 默认 true
    RegenerateTableOfContents = false,     // 默认 true
    RemoveTinyImages = false,              // 默认 true

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
OfficeDocument cleaned = result.Document;   // 与 document 是同一个实例：就地修改
```

第一句要提醒的是：**默认构造就是"中文公文常用值"，绝大多数开关默认已打开**，不是全关。所以上面这段是在改目标值并关掉演示中不需要的三项。为什么要关 `MergeShortParagraphs`——实测它会把"项目实施方案"和"一、实施范围"并成 `项目实施方案一、实施范围`，还顺带产出一个假的二级标题。

伪标题升级是按字号比例判定：章级命名模式或比例 ≥1.5 记 1 级，≥1.3 记 2 级，≥1.15 记 3 级，其余 4 级；`FixHeadingLevelSkips` 再把跳跃的层级夹到"上一级 +1"。实测那份草稿的结果是 `变更 9 项，问题 1 项`，其中包含 `伪标题升级为 1 级标题: 项目实施方案` 和 `标题层级跳跃 4 -> 3`。表格则统一把第 0 行标成表头并打开 `RepeatHeaderRow`。

批处理走 `WordProcessor.Process(inputPath, outputPath, options)`：先在输入旁边生成 `.backup-<时间戳>.docx`，再加载、整理、写回，报告固定落在 `outputPath + ".report.txt"`。回滚是 `WordProcessor.Rollback(backupPath, targetPath)`，本质是覆盖式 `File.Copy`——**目标路径填原始文件**，不是填输出文件。

验证整理效果时有两个工具，各自有明确的适用面：

- `WordDocumentDiff.Compare` 只比文本行。纯格式整理后差异是 `HasDifferences=False`、`+0 -0 =8`；改了正文才会 `True +1 -0`。**所以它不能用来证明"整理生效了"**，那是 metrics 的活。
- `WordDocumentMetrics.Measure` 统计正文 + 页眉页脚 + 单元格段落。实测整理后：`段落 12，空段 1，字 92，表 1，图 0，标题 3`、`各级标题: 1/1/1/0`、字体分布 `黑体 3 处 / 宋体 8 处`、字号分布 `16/15/14/12`。这组数字前后对比才是格式整理的有效证据。

还有一点关于克隆：库内部有文档克隆器，但**没有公开的 Clone API**。要在不改原件的前提下试不同整理策略，就重新从文件加载一次，或者先 `CreateBackup`。

---

## 五、零依赖的另一面：边界条件得自己扛

用第三方库时，畸形文件的处理是别人的事；自研时，这就是你的事。这个库在这一点上的做法可以直接抄给自己的项目。

**把框架差异收敛成一种异常。.NET Framework 的 `ZipArchive` 延迟解析中央目录，损坏的 OOXML 条目名会在毫不相关的调用点抛 `ArgumentException`，绕过读取器外层的 `catch (IOException)`。**解决办法是新增 `PackageArchive.OpenForRead`：打开时立即物化中央目录，把这类异常统一收敛成 `InvalidDataException`，已应用于 XLSX/DOCX 读取、DOCX 模板填充与格式检测四处入口。收敛保证的是"同一扇门在六个框架上表现一致"，而异常类型仍取决于你从哪扇门进来——实测四类损坏输入：

```text
XlsxReader.Load(截断的 xlsx)      -> InvalidDataException（End of Central Directory record could not be found.）
OfficeFile.Load(截断的 pdf)       -> InvalidDataException（PDF 没有有效 catalog）
OfficeFile.Load(随机 8 字节)      -> NotSupportedException（无法识别文件格式…）
OfficeFileFormats.Detect(随机字节) -> Unknown，不抛异常
```

也就是说：**格式专属的读取器把损坏统一表达为 `InvalidDataException`；门面先做内容嗅探，嗅不出来的走 `NotSupportedException`；只做检测则永远返回 `Unknown`。** 服务侧要写"坏文件返回 400"的兜底，就照这三类分别处理，别只 catch 一种。

**给解压量和结构规模设硬上限。**

```text
OfficePackageLimits.MaxEntryBytes   = 256 MB
OfficePackageLimits.MaxTotalBytes   = 1 GB
PdfReadLimits.MaxDecodedStreamBytes = 256 MB
PdfReadLimits.MaxXrefEntries        = 2,000,000
PdfReadLimits.MaxObjectStreamObjects = 1,000,000
PdfReadLimits.MaxPageTreeDepth      = 64
```

流式读取用 `ReadEntryBytes(entry, ref totalBytes)` 累计，超限抛 `DecompressionLimitException`——它和 `InvalidDataException` 分开，是为了让调用方能区分"文件坏了"和"文件太大了"，这两者在服务侧的处置策略通常不同（重试/告警 vs. 拒绝/限流）。

**把攻击面写进测试而不是评论。**测试套件里包含 XXE、zip 炸弹、路径穿越、zlib 解压上限、模糊输入这类用例，且同一套断言在 `net8.0` 与 `net461` 上各跑一遍。安全回归和跨框架回归共用一套基线，是自研项目里性价比高的一件事。

---

## 六、十七条实测出来的行为，文档不会主动告诉你

这些是写教程时一条条撞出来、已经写进教程但值得单独列在这里的：

1. `report.Style` 必须在 `Export`/`ToXxx` 之前配置完，模型每次重新构建。
2. `Table.RepeatHeaderRow` 默认 `true`——你以为需要显式打开，其实往往需要显式关掉。
3. DOCX 以表格结尾时写出器补一个空段落，回读块数会多 1。
4. `ExcelReadOptions` 三个隐藏开关默认 `true`，`WorkbookLayoutOptions.IncludeHiddenWorksheets` 默认 `false`。
5. `Cell.DisplayText()` 不输出 `####`；`####` 来自 `EmulateNumberOverflow`。
6. `decimal` 写入时转成 `double`，回读断言要用值不是文本。
7. `WordNormalizeOptions` 默认几乎全开，`MergeShortParagraphs` 能粘出假标题。
8. `WordDocumentDiff` 只比文本行，格式整理的效果要靠 `WordDocumentMetrics` 看。
9. 字节版 `OfficeConverter.Convert` 不接受 `WorkbookLayoutOptions`；带选项的 Excel→PDF 字节路径要走 `XlsxReader` + `WorkbookToPdf`。
10. 图表按 OOXML 保存，只有 Excel/WPS 会渲染；导出 Word/PDF 时输出占位文本。需要图形请在源工作簿里改用 `ExcelImage` 锚定图片。
11. 坏文件的异常类型取决于入口：专属读取器给 `InvalidDataException`，门面嗅探失败给 `NotSupportedException`，`Detect` 只给 `Unknown`。
12. Markdown 没有内容签名：`Detect("a.md")` 是 `Unknown`，只有路径版 `Load` 会按扩展名兜底，字节/流入口拿不到 `.md`。
13. `slide.Layout` 是回读时推断出来的，不能作为断言目标（`SectionHeader`→`Title`、`TitleAndTable`→`TitleAndContent`）；要断言形状类型、占位符索引和文本。
14. PPTX 里 `Table.Columns` 的列定义不落 XML：写出实测"2 行 0 列"，回读按每行单元格数补成"2 行 2 列"，列宽信息回不来。
15. `PdfFile.AddWatermark("文字", 0)` 是给**全部页**加图层为背景的水印，不是给第 0 页加——整数字面量会隐式转成 `WatermarkLayer`，`(string, WatermarkLayer)` 重载比 `params int[]` 更具体所以胜出。只给指定页加水印必须写 `AddWatermark(WatermarkOptions, 0)`。
16. `PdfFormField.Value` 读取时按 `/V` → `/DV` 回退（与阅读器显示默认值的行为一致），所以"没填过"和"填成了默认值"读起来一模一样；要区分得看 `field.Dictionary.ContainsKey("V")`。
17. 单选组的选项节点通常没有 `/T`，`FullyQualifiedName` 是空串、也进不了 `GetField` 的名字索引——按 `form.GetField("order.plan")` 只能拿到组，选项要遍历 `Children` 取 `OnState`；组的 `/Rect` 是从选项控件回填补出来的，但组的 `PageIndex` 仍是 `-1`（控件挂在选项上，不在组上），扁平化因此逐个处理选项而不是处理组。

---

## 七、上手路径

建议的顺序是：先跑通门面那一份，再按你实际负责的格式深入一个包。

| 包 | 说明文档（也是包内 README） | 分步教程 |
| --- | --- | --- |
| `Wisdom.Office` | [src/Wisdom.Office/README.md](../src/Wisdom.Office/README.md) | [docs/Wisdom.Office.Tutorial.md](Wisdom.Office.Tutorial.md) |
| `Wisdom.Office.Core` | [src/Wisdom.Office.Core/README.md](../src/Wisdom.Office.Core/README.md) | [docs/Wisdom.Office.Core.Tutorial.md](Wisdom.Office.Core.Tutorial.md) |
| `Wisdom.Pdf` | [src/Wisdom.Pdf/README.md](../src/Wisdom.Pdf/README.md) | [docs/Wisdom.Pdf.Tutorial.md](Wisdom.Pdf.Tutorial.md) |
| `Wisdom.Word` | [src/Wisdom.Word/README.md](../src/Wisdom.Word/README.md) | [docs/Wisdom.Word.Tutorial.md](Wisdom.Word.Tutorial.md) |
| `Wisdom.Excel` | [src/Wisdom.Excel/README.md](../src/Wisdom.Excel/README.md) | [docs/Wisdom.Excel.Tutorial.md](Wisdom.Excel.Tutorial.md) |
| `Wisdom.Ppt` | [src/Wisdom.Ppt/README.md](../src/Wisdom.Ppt/README.md) | [docs/Wisdom.Ppt.Tutorial.md](Wisdom.Ppt.Tutorial.md) |
| `Wisdom.Markdown` | [src/Wisdom.Markdown/README.md](../src/Wisdom.Markdown/README.md) | [docs/Wisdom.Markdown.Tutorial.md](Wisdom.Markdown.Tutorial.md) |

每份教程的输入文件都由教程自身生成，代码片段拼起来就是一个可运行的控制台程序，文中输出全是实测结果——包括那些"看起来像笔误"的行号和字节数。

验收产物时推荐这个姿势：回读并断言标题存在、关键数据行存在、页数合理、没有空页；不要硬记行号，遍历 `sheet.Cells` 找目标文本更稳：

```csharp
foreach (Cell cell in back.Cells)
{
    if (cell.Value is string s && s == "无线键盘")
    {
        Console.WriteLine($"row {cell.Row}, col {cell.Column}");   // row 7, col 1
    }
}
```

库自身的回归就是靠财务报表、数据看板、检验/审核报告、考勤明细四类真实版式的断言兜住的。把你自己的产物也纳入同一套检查，是上线后最省事的护栏。

---

## 八、什么情况下不该选它

把边界说清楚比承诺能力更有用：

- **加密 PDF**：口令类（`/Filter /Standard`）读写都支持——RC4-40/128、AES-128、AES-256 四种算法已与第三方阅读器双向验证（写出用 R6，读取兼容 R5，也读得懂对方把 `/P` 写成无符号值的文档）；非 Standard 的处理器（公钥加密、厂商自定义）直接抛 `NotSupportedException`，R2–R6 以外的版本号也一样拒。
- **文字水印的旋转语义**：`AddWatermark` 会按页面 `/Rotate` 反向补偿矩阵，保证打印方向读起来是正的；但反过来说，页面本身带 `/Rotate` 时，`ExtractText()` 拿到的是"按流向合并后"的片段，行序仍按设备坐标排，跨页对齐别指望。
- **表单**：AcroForm 的读取、填写、外观生成与扁平化都支持；`/XFA` 只能整体剥离，不解析模板也不做字段级回退。本库不生成新的表单字段，给空白文档加控件要走 `PdfFile` 的对象级 API（教程第 9 章有完整示例）。签名域只读（赋值抛 `InvalidOperationException`）且不参与扁平化；字段里的 JavaScript 计算与校验脚本不执行；外观按控件矩形重排，是"看得懂、印得出"的近似，不追求与 Acrobat 逐像素一致。
- **CFF / OpenType（`OTTO`）字体嵌入**：不支持，只嵌 TrueType `glyf` 轮廓；隔行扫描 PNG 也不支持。
- **扫描件 OCR**：文本提取依赖文本操作符，图像里没有字就没有字。
- **`.doc` / `.xls` 写回**：只读；`.doc` 的复杂域、批注、修订不保证还原；`.xls` 只支持 BIFF8，不含加密工作簿。
- **`.ppt` 读写**：都不支持，只有 `.pptx`。PPTX 侧还不处理图表、SmartArt、嵌入 OLE、母版深层样式与动画序列，主题与 `slideMaster` 由本库生成一套固定的，不解析外部 `.potx` 模板。
- **完整 CommonMark / GFM**：Markdown 解析器是手写子集。引用式链接 `[text][ref]`、脚注 `[^1]`、HTML 块内套 Markdown、列表项内缩进续写的多段落都不支持——不抛异常，按字面文本保留。要做严格合规渲染请另找实现。
- **Word/PDF → Excel**：不提供。
- **像素级复刻 PDF**：PDF 保存是依据统一模型重新排版，不保留原字节和完整视觉结构。
- **图表进 Word/PDF**：见第 10 条，占位文本。
- **超链接的跨格式保真**：`TextRun.Link` 目前只在 Markdown 侧填充，DOCX/PPTX 读入模型时只留文字与格式、URL 不入库，所以 Word 转 Markdown 后链接退化成普通文本。

反过来说，如果你的需求是"服务器上不装 Office、不引第三方组件，把数据稳定导出成 PDF/Word/Excel/PPTX/Markdown，或者把上传的办公文档读成统一模型"，从 .NET Framework 4.6.1 到 .NET 10 都要覆盖——这个库的四个场景入口恰好就是为这四件事准备的。
