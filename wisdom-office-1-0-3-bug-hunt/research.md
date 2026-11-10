# Research: Wisdom.Office 1.0.3 修复实录

## 选题陈述

以 Wisdom.Office（纯 C# 自研 Office 文档库，七个 NuGet 包）在 1.0.3 发布前的一轮缺陷排查为线索，写一篇面向 .NET 开发者的中文技术长文。核心不是"我们做了个库"，而是"一批**不会报错、但会毁数据或毁文件**的缺陷是怎么被找出来、又是怎么被证明修好了的"。

## 读者假设

- 中文技术社区（掘金 / 博客园 / CSDN）的 .NET 开发者
- 大概率做过文件读写、导出、互转；对 OOXML / PDF 的内部结构不熟但知道它们存在
- 关心：怎么发现静默失败、怎么写测试才算数、性能问题怎么量化
- 不需要被科普"什么是 docx"

## 研究范围

- 来源：本仓库工作区在 2026-11 这一轮的全部改动、测试与提交记录
- 时间：2026-11-04 至 2026-11-06（1.0.3 打包完成）
- 不覆盖：与本次修复无关的库功能全景、性能基准的横向对比

## 关键事实（均已在本会话中被工具输出或提交记录证实）

### 起点

- 接手时测试工程**编译不过**：`PptxWriter.cs` 调用了一个从未落地的方法 `HasDeclaredHeader`，六个目标框架全部 `CS0103`。
- 已入库的一条红测（提交 `f9fc99a`）在等这个方法。

### 已修的条目（按现象归类，共 19 条）

数据丢失 / 结构破坏：

1. **E7 先截断再写**：`XlsxWriter` / `DocxWriter` / `PptxWriter` / `PdfRenderer` / `WordFormatConverter` / `OfficeFile` 共九个入口都用 `new FileStream(path, FileMode.Create, ...)`，而参数校验在**被调用的流重载内部**。于是 `Write(null, path)` 先把用户文件截成 0 字节，再抛 `ArgumentNullException`。修复前实测输出：`Test failed: a failed xlsx write left the original file with 0 byte(s) instead of the 13 it had`。
2. **WE9 ODT 超限图片**：单张超过 64 MB 的内嵌图片抛 `InvalidDataException`，整篇文档读不出来。修复前：`Unhandled exception. System.IO.InvalidDataException: ODT 内嵌图片超过安全上限。`
3. **PT2 `p:timing` 树形状**：一个 `p:par` 里并列两个 `p:cTn`、`seq` 与根 `cTn` 平级、两个节点同用 `id="1"`、`dur="indef"`（规范词法是 `indefinite`）、类别写进 `type` 而不是 `nodeType`、把 `a:clrMapOvr` 塞进 `p:setTo`。修复前：`Test failed: the root cTn must carry nodeType="tmRoot":`。
4. **PT3 表格结构**：没声明表头也恒写 `firstRow="1"` / `bandRow="1"`，首行刷成加粗白字，每个格子补 `a:noFill` 把 `tableStyleId` 的条纹堵死。修复前：`Test failed: firstRow was switched on for a table that never declared a header row`。
5. **PT3 形状填充**：没设填充/描边也补 `a:noFill` 与空的 `a:ln`。修复前：`a shape that never set a fill was pinned with a:noFill, silencing the inherited one`。
6. **WE10 可空模型字段**：`TextRun.Text` / `TextRun.Format` / `Table.Borders` 置空后，DOCX/RTF 崩、HTML/ODT 正常。修复前两条：`the Rtf writer failed on a null model field: NullReferenceException`、`the Html writer failed on a null model field: NullReferenceException`。
7. **PT10 `docProps/app.xml` 次序**：不满足 `CT_Properties` 的 sequence。修复前：`docProps/app.xml is not in CT_Properties order (position 0): Application, Slides, Company, AppVersion`。
8. **WE7 部件解析失败静默丢内容**：`word/header1.xml` 存在但 XML 非法时返回 `null`，"没有页眉"与"页眉被丢弃"不可区分。修复前：`expected exactly one downgrade notice for the corrupt header, got 0`。

静默算错：

9. **Excel 共享公式表名前缀**：`Sheet1!A2` 里的 `A2` 从不改写，每个从格都指向主格那一行。修复前：`a sheet-qualified relative reference did not shift with the dependent: Sheet1!A2+1`。
10. **PT9 切换效果**：`dir`/`spd`/`advClick` 写死，"向上推入"读回变向左推入、自动换片变必须点击。
11. **PT3 上下标**：读端按 `a:rPr/@baseline` 还原，写端从不输出。修复前：`superscript was lost: the writer never emits a:rPr/@baseline`。
12. **PT3 段落间距**：`lnSpc`/`spcBef`/`spcAft` 反过来——写端一直写、读端从不读。
13. **PT3 纵向合并**：只写 `gridSpan`，`RowSpan` 丢掉、续接格不占位，第二行数据整体左移一列。
14. **PT8 章节页版式**：`SectionHeader` 与「仅标题页」形状完全相同，形状推断永远分不开，读回被重绑成 `Title`。修复前：`the section header slide came back as Title`。
15. **PT7 零尺寸几何**：`cy="0"` 被当成几何缺失，水平分隔线读回变成默认大框。修复前：`a zero-height separator was inflated to 313.2 pt on the way back in`。
16. **Excel 越界日期序列**：`epoch.AddDays(serial)` 对 `1E+10` 抛 `ArgumentOutOfRangeException`，一个单元格废掉整本工作簿。修复前：`Unhandled exception. System.ArgumentOutOfRangeException: ... (Parameter 'serial')`。
17. **Core 模板坏格式串**：`TemplateEngine.Replace` 直接调 `ToString(format, culture)`，`{{when:'unclosed}}` 抛 `FormatException` 穿出，整份报表全废。

性能与并发：

18. **Excel 读表平方级**：`ExcelCollection` 把 `sheet.LastRow` 写在循环条件里，而 `LastRow` 每次访问都遍历全表单元格。修复前：`reading 8000 rows took 1600 ms; LastRow is still being evaluated inside the loop`。
19. **PD20 平铺水印退化**：步长极小时 600 次绘制预算全花在页面一角；再小下去 `double → int` 饱和成 `int.MinValue`。修复前：`a 0.01 pt block collapsed the tiling into a corner: 600 tiles, furthest at 33.6 x -0.0 (page is 595 x 842)`。

另外一条只做代码对齐、没有自动断言：**PD19** `FontResolver` 在锁外读字体字节缓存（并发时序不可稳定复现）；**WE8** ODT 打开入口绕过 `PackageArchive`（需要构造 .NET Framework 惰性解析中央目录的时机差）。

### 方法论

- **先红后绿**：每条修完都临时回退源码、取到真实失败原文、再恢复并转绿。文中引用的所有 `Test failed:` 都是这样取到的。
- **双目标**：`net8.0` 与 `net461` 各跑一遍，都到末行 `All tests finished.`；跨框架行为差异（如 `ZipArchive` 的惰性解析）是这条纪律存在的原因。
- **能自动化的才叫修好**：两条无断言的条目在提交与文档里明确写了原因，不假装已覆盖。
- **发布级验证**：1.0.3 发布前做了一次全新克隆 → 构建 → 跑 `net8.0` 全套，证明快照自包含。

### 两个"测量陷阱"（真实发生）

- **假绿**：PD20 第一版红测在"关掉修复"的树上也是绿的——因为 `WatermarkOptions` 默认 `TileSpacingX/Y` 是 60/40 磅，把退化路径兜住了。夹具必须显式把间距归零才测得到。
- **断言的可见性**：WE10 的红测按 `Docx → Rtf → Html → Odt` 顺序跑，第一版只暴露了 Rtf 与 Html；`OdtWriter` 那处是靠**代码形态审计**发现的，不是被测到的。

### 交付结果

- 1.0.3：七个包 + `.snupkg`，六份目标框架资产（net461 / netstandard2.0 / net7-10），依赖区间锁在同一版本。
- 仓库：GitHub 与 Gitee 双远程同步到同一提交，本次共 58 个提交。
- 打包与推送：包已构建并逐包核对 nuspec；推送到 nuget.org 因本机缺少 API Key 返回 401，未上传（这是发布流程的现状，不是代码问题）。

## 可见的取舍（写进文章的"没修的部分"）

- `PT2` 只做结构断言，**没有 PowerPoint 真机验证**。
- `PT3` 仍未补：`wrap`、主题 accent、超链接、组合与旋转翻转——要动模型或写端关系表。
- `WE7` 只给 Word 侧建了最小告警通道（`OfficeDocument.Warnings`），没有替其他包定契约。
- `PD19` / `WE8` 无自动断言，理由已记录。

## 可能的文章角度

1. 以"九条静默失败"为主线的工程叙事（选定）
2. 以"怎么写出会红的测试"为主的测试方法论
3. 以 OOXML 结构合规为主的格式科普

## 未决 / 不确定

- 反例的行业对比数据（如其他库的同类缺陷率）没有采集，不写。
- 性能数字只有本机一组（8000 行 1600 ms → 阈值内），不做跨机器比较。
- 发布到 nuget.org 的最终状态取决于用户是否提供 API Key。

## 风险与措辞

- 文中所有失败原文与数字均来自本会话工具输出，可直接引用。
- "修复前/修复后"只描述本仓库的代码，不外推到其他库。
- 标准条款（ISO/IEC 29500、ISO 32000-1）按通用知识引用条款主题，不逐字引条文。
