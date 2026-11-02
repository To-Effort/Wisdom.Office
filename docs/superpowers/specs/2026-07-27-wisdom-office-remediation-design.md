# Wisdom.Office 热点缺陷修复规范（Excel / 公式引擎 / Word）

日期：2026-07-27
范围：`src/Wisdom.Excel`、`src/Wisdom.Office.Core/Formula`、`src/Wisdom.Word`（含 Docx、Formats、Processing）。`Wisdom.Ppt`、`Wisdom.Pdf`、`Wisdom.Markdown` 本轮未审，见 §9。

---

## 1. 背景与当前基线

四个源码审计代理并行通读上述三个包，合计回报约 80 条发现。逐条对**当前工作区**复核后，至少 12 条已不成立——它们描述的是提交历史里的旧代码，而工作区有一批未提交改动（下称「在飞批次」）已经修掉了其中一部分。因此本规范不采信任何未经复现的结论。该批次在本规范撰写期间仍在被并行编辑，从约 692 行增长到 17 文件 921 增 / 185 删，所以下文凡引用其规模或状态，都以 Phase 0 起点的重测为准。

复核时确立的事实：

| 事实 | 依据 |
| --- | --- |
| 自检运行器是自制控制台断言器，不是 xUnit | `tests/Wisdom.Office.Tests/Wisdom.Office.Tests.csproj` 为 `OutputType=Exe`，入口 `tests/Wisdom.Office.Tests/Program.cs` |
| `dotnet test` 在本仓库等于空跑 | 只输出「所有项目均是最新的」，返回码 0，但一条测试都没执行 |
| 有效验证命令 | `dotnet run --project tests/Wisdom.Office.Tests -f net8.0`，以及 `-f net461` |
| 基线必须在 Phase 0 起点重测，不可沿用历史结论 | 本规范撰写过程中曾取得双目标全绿（72 行输出，含 `[hardening]`/`[security]`/`[limits]` 各组）；此后在飞批次继续增长到 17 文件 921 增 / 185 删，`dotnet run --project tests/Wisdom.Office.Tests -f net8.0` 转为编译失败：两处 `CS0122` 保护级别错误——`XlsxStyles.ApplyBorder` 声明为 `private` 却被 `XlsxDecorationReader` 调用，`XlsxWriter.BorderEntry.CreateEdge` 声明为 `private` 却被同文件外层写出逻辑调用。主题与边框那条改动线仍未收尾 |
| 在飞批次曾带一个硬崩溃 | `src/Wisdom.Word/Docx/DocxWriter.cs` 里出现过 `new XElement(WordNamespace + "p")`；`WordNamespace` 是字符串常量（`:20`），拼接得到 `...mainp` 而非「命名空间 + 本地名」，写出任何含相邻两张表的文档直接抛 `XmlException`，`Program.Main` 在 `:85` 未捕获中止，其后所有测试不执行。现工作区已改为方法内 `XNamespace w = WordNamespace;`（`:727`）并用 `w + "p"`（`:741`） |

Phase 0（§10）先把在飞批次按模块分组提交，取得干净基线，再动任何修复。

---

## 2. 取证方法与置信度分级

每条发现标注证据强度：

- **A 已核实**：本次直接在当前工作区读到该代码行，且逻辑后果明确。
- **B 待确证**：来自代理报告，行号可能漂移，后果合理但我未逐行读通。

A 类条目在进入修复前，同样必须先跑红（§3）。

---

## 3. 门禁规则：红测先行

1. 每条发现先在 `tests/Wisdom.Office.Tests` 对应测试文件里写一条最小复现断言，沿用现有 `Require(...)`、`RequireEqual(...)`、`RequireContains(...)` 风格，并输出一行 `[remediation] <编号> ...` 摘要，与现有 `[hardening]`、`[security]` 组保持一致。
2. 断言必须在未修改源码时失败。跑不出失败的条目判定为误报，移入 §7.2。
3. 修复后 net8.0 与 net461 双目标都要跑完整运行器到 `All tests finished.`；任何因未捕获异常导致的早退都视为基线不可信，须先解决早退。
4. 涉及读入畸形文件的条目，复现输入在测试内用 `MemoryStream` 与构造字节生成，禁止依赖仓库外的文件路径。
5. 代码注释遵循项目既有规范：中文与英文并列；单行 `if` 也独立成行；先写 guard 再写主体；新增公共成员补齐文档注释。

---

## 4. P0 —— 崩溃、挂起与产出错误数据

### E1｜写端完全不识别 1904 日期系统　A

- 位置：`src/Wisdom.Excel/Xlsx/XlsxWriter.cs` 的 `ToSerial`（基准固定 `new DateTime(1899, 12,t...)`）；该文件 `date1904` 命中数为 0。
- 现状：读端正确识别 `date1904="1"`（`src/Wisdom.Excel/Xlsx/XlsxReader.cs:940-945` 的 `FromSerial` 有 1904/1900 分支），模型也保留了该标记；写端输出 `<workbookPr>` 时不写该属性，且所有日期序列一律按 1900 基准计算。
- 后果：加载 1904 制工作簿再保存，每一个日期值与依赖日期的公式都偏移 4 年 + 1 天，且不报错。
- 目标行为：写端读取模型的 1904 标记，序列计算与 `<workbookPr date1904="1">` 同步输出；无标记时维持现有 1900 行为。
- 红测：内存构造 1904 工作簿 → `Load` → `Save` → 再 `Load`，断言 `DateTime` 值与原始一致，并断言输出 XML 含 `date1904="1"`。
- BREAKING：否（原本产出即错误）。

### E2｜日期序列无上限钳制，单个畸形值毁掉整次加载　A

- 位置：`src/Wisdom.Excel/Xlsx/XlsxReader.cs:940-945`（`FromSerial` 只钳负数，随后 `epoch.AddDays(serial)`）；`src/Wisdom.Excel/Legacy/XlsReader.cs:873-879`（同型）；异常边界见 `src/Wisdom.Excel/Xlsx/XlsxReader.cs:218-225`，那里只 `catch (XmlException)`。
- 现状：畸形但合法的 XML 中一个 `<v>1e10</v>` 就让 `AddDays` 抛 `ArgumentOutOfRangeException`，而外层只捕获 `XmlException`，异常穿透导致 `ExcelCollection.Load()` 对整个工作簿失败。
- 第二处不一致：`src/Wisdom.Excel/ExcelCollection.cs:524` 的转换路径已钳制 `serial >= 0 && serial < 2958466`，两条读取路径对越界序列处理不同；序列 60（Excel 幻象的 1900-02-29）未被特判。
- 目标行为：收敛到单一日期序列换算入口（含上下界与 60 特判），越界按既定策略降级而不是抛出，单元格级异常不得升级为整簿加载失败。
- 红测：内存构造含超大 `<v>` 的最小 xlsx → 断言当前 `Load` 抛出 → 修复后断言 `Load` 成功且其余单元格完好。
- BREAKING：否（防御性收紧）。

### F1｜`FACT` 与 `FACTDOUBLE` 迭代次数取自参数，无钳制　A

- 位置：`src/Wisdom.Office.Core/Formula/Functions.cs:144-145`（注册项只判 `x < 0`）与 `:369-381`（`Factorial`、`DoubleFactorial` 是 `for (double i = 2; i <= n; i++)` 裸循环）。
- 现状：`=FACT(1000000000)` 循环 10 亿次；Excel 在 170! 之后直接返回 `#NUM!`。
- 目标行为：超出溢出界返回 `#NUM!`，不进循环。
- 红测：`Evaluate("=FACT(2000000000)")` 断言返回 `#NUM!`（现状为挂起，测试内需超时包裹）。

### F2｜通配符匹配为无记忆化回溯，条件取自单元格文本时指数爆炸　A

- 位置：`src/Wisdom.Office.Core/Formula/Functions.Aggregates.cs:359-372`（`MatchWildcard` 的 `*` 分支对每个起点递归），调用点在 `COUNTIF`、`SUMIF`、`AVERAGEIF` 的逐格循环内。
- 现状：`COUNTIF(A1:A100000,"*a*b*c*d*e*f")` 永不返回；模式串可以来自不可信工作簿里的普通单元格文本，且对模式与文本长度无任何上限。
- 目标行为：改为线性匹配（贪心加单点回溯，或受限迭代 DP），并对模式与文本长度设上限，超限返回可预期错误。
- 红测：多 `*` 模式配中等长度串，测试内超时断言现状挂起、修复后在线性时间内给出正确结果。

---

## 5. P1 —— 产出被 Word/Excel 判损坏，或静默丢内容

### E3｜组合图里散点与气泡系列引用了未生成的坐标轴　A

- 位置：`src/Wisdom.Excel/Xlsx/XlsxChartParts.cs:31`、`:33`（`ScatterXAxisId = "555555555"`、`ScatterYAxisId = "666666666"`），组内固定发出这两个 `axId`（`:297-298`、`:309-310`）；而 `BuildPlotArea`（`:171-179` 区段）仅在散点族且无类目系列时才写这两根轴。
- 现状：混合图（散点系列 + 类目系列；系列级类型覆盖是被支持的，见同文件 `:126`）里散点组引用了不存在的轴 id，Excel 弹「需要修复」。
- 目标行为：轴集合按实际出现的系列种类生成；混合图要么补出散点轴对，要么让散点组复用已有轴 id。
- 红测：建柱加散点的组合图 → 写出 → 解析回 XML，断言每个 `axId` 都能在 `plotArea` 找到对应轴定义。

### E4｜模板嵌套 `{{#each}}` 被误配，内层静默丢失　A

- 位置：`src/Wisdom.Excel/Templates/XlsxTemplate.cs:314-338`（`DetectLoops` 只扫顶层，命中第一个 `{{/each}}` 即把行游标推进到该结束行）。
- 现状：外层循环体里的内层循环被当作字面文本吞掉，随后 `RemoveLoopMarkers` 清除标记，内层数据永不展开，且无任何错误。
- 目标行为：按名称与作用域正确配对 `{{#each}}` 与 `{{/each}}`；做不到就明确抛错，绝不静默。
- 红测：嵌套模板 → 断言现状展开结果缺少内层行 → 修复后断言展开正确或抛出可诊断异常。

### F3｜单元格链式求值无递归深度上限，可致进程级栈溢出　A

- 位置：`src/Wisdom.Office.Core/Formula/Evaluator.cs:432-437`（`_path` 只做环检测）；表达式嵌套另有 `Parser.cs` 的 `MaxDepth = 128`。
- 现状：`A2=A1+1` 直到 `A20000=A19999+1` 这类依赖链每级叠约 5 个栈帧（`CellValue` → `Evaluate` → `Eval` → `EvalBinary` → `CellValue`）。NET 的栈溢出不可捕获，宿主进程被杀，而 `FormulaEngine.Calculate` 会主动遍历全链。
- 目标行为：对依赖链深度设上限，超限返回明确错误值而不是继续下钻。
- 红测：内存构造 5 万级依赖链 → 断言现状崩溃或超时 → 修复后断言返回错误值且进程存活。

### F4｜`COUNTBLANK` 对多格区域恒为 0　A

- 位置：`src/Wisdom.Office.Core/Formula/Functions.Aggregates.cs:37-46`（对 `Evaluator.Scalar(a[0])` 的结果做 `Flatten`）配合 `src/Wisdom.Office.Core/Formula/Evaluator.cs:462-468`（多格 `Range` 经 `Scalar` 变成 `Err(Value)`）。
- 现状：区域先被压成单个错误值，循环里既不是 `Empty` 也不是空文本，任何区域都数出 0。
- 目标行为：直接展开区域计数，与 `COUNTA` 对空串与空格的处理口径一致。
- 红测：`COUNTBLANK(A1:A5)`，区域内混合空格子与空串，断言计数正确。

### F5｜`COUNT` 遇到区域内错误时吞掉错误并返回 0　A

- 位置：`src/Wisdom.Office.Core/Formula/Functions.cs:155`。
- 现状：`if (!CollectNumbers(a, out n, out err)) return FxValue.Num(0);`，而同文件 `SUM`（`:153`）与 `AVERAGE`（`:158`）都是 `return err;`。Excel 传播错误，这里给的是一个看似合理的 0。
- 目标行为：与兄弟聚合一致传播区域错误。
- 红测：区域内放一个 `#DIV/0!`，断言 `COUNT` 返回该错误。

### F6｜`AVERAGEIF` 分母把空白与非数值计入　A

- 位置：`src/Wisdom.Office.Core/Formula/Functions.Aggregates.cs:201`。
- 现状：`if (Evaluator.TryCoerceNumber(Evaluator.Scalar(avgValues[i]), out d)) { total += d; count++; }`，而写法正确的 `SUMIF`（`:166`）额外要求 `Kind == FxKind.Number`。`TryCoerceNumber` 对 `Empty` 返回真并置 0，空串与数值化文本同样进分母，均值偏低。
- 目标行为：与 `SUMIF` 取数口径一致。
- 红测：条件命中 3 行、求和区含 1 个空格子，断言分母为 2 而非 3。

### F7｜文本转数值走机器区域设置，跨主机结果不一致　A

- 位置：`src/Wisdom.Office.Core/Formula/Evaluator.cs:481`（`double.TryParse(v.Text, NumberStyles.Float, CultureInfo.CurrentCulture, ...)`）、`src/Wisdom.Office.Core/Formula/Functions.Advanced.cs:326`（`DateTime.TryParse(text, CultureInfo.CurrentCulture, ...)`）。
- 现状：同一个文件在 de-DE 上 `VALUE("1,5")` 得 1.5、`DATEVALUE("01/02/2021")` 得 1 月 2 日，在 en-US 上报 `#VALUE!` 或得 2 月 1 日。引擎其余部分刻意全部 Invariant，这是唯一的非确定路径。
- 目标行为：求值路径全部改 Invariant；如确有需求，另设显式的区域解析入口。
- 红测：测试内临时把 `CurrentCulture` 设为 de-DE 与 en-US，断言两种环境下求值结果相同。
- BREAKING：是。依赖本机区域设置的既有工作簿求值结果会变，需随 1.1.0 与「已知限制」声明。

### W1｜`Paragraph.StyleName` 只读不写，段落样式往返丢失　A

- 现状：读端填充（`src/Wisdom.Word/Docx/DocxStyles.cs:254`，取样式名否则取样式 id），写端全文 `StyleName` 命中数为 0；`pStyle` 只在标题层级时按 `Heading<n>` 生成。
- 后果：一次 `Load` 加 `Save` 之后，所有非标题段落样式（Normal、自定义样式、代码块、目录系列）被抹平。
- 目标行为：写端优先输出 `StyleName`，无样式表定义时回退到既有标题映射。
- 红测：带自定义段落样式的 docx 往返，断言 `StyleName` 保持。

### W2｜列表是单向的，编号定义从不写出　A

- 现状：`src/Wisdom.Word/Docx/DocxWriter.cs` 与 `DocxPartReader.cs` 中 `numbering` 命中数为 0；读端把 `w:numPr` 拍平成字面 `ListMarker` 文本，写端把该文本当普通 run 输出。
- 后果：往返后 Word 里所有自动编号列表变成手打文本，多级列表层级信息全丢。
- 目标行为：模型保留列表与编号语义，写端产出 `word/numbering.xml` 与 `w:numPr`，并在内容类型与关系里登记。
- 红测：含有序列表的 docx 往返，断言仍是编号而非文本。
- 说明：范围较大，作为 P1 内独立里程碑（§10 Phase 3）。

### W3｜复合域（`w:fldChar` 加 `w:instrText`）不解析，缓存结果被冻成静态文本　A

- 现状：全仓 `fldChar` 命中数为 0；读端 `src/Wisdom.Word/Docx/DocxPartReader.cs:435` 附近对 `instrText` 直接 break，写端只生成 `w:fldSimple`。
- 后果：`PAGE`、`NUMPAGES`、`DATE`、`TOC` 等域往返后退化为死文本，页码不再变、目录不再更新。
- 目标行为：解析域字符与指令序列为域模型，写出时还原为域字符序列。
- 红测：含 `PAGE` 域的 docx 往返，断言仍能识别为域。

### C1｜HTML 导出的 `style` 属性上下文未做属性转义　A

- 位置：`src/Wisdom.Word/Formats/HtmlWriter.cs:392-396`（`AppendFontFamily` 只转义反斜杠与单引号）与 `:297`（把 css 直接拼进 `<span style="`）、`:225`（段落 `style=`）。
- 现状：字体族名可以来自不可信输入文件（RTF、ODT、DOCX 的字体表）。名字里带一个双引号即可提前闭合 `style` 属性并注入事件处理属性。文本、`href`、`data-style`、书签名都正确走了 `WordFormatUtil.EscapeHtml`，唯独这里绕过。
- 目标行为：写入属性上下文前对 `& < > "` 做实体转义，或过滤非法字符后拒绝该字体名并回退默认字体；CSS 上下文转义保持不变。
- 红测：字体名含双引号与游离属性的文档导出 HTML，断言输出中不出现被提前闭合的属性；按现有 `[hardening]` 风格补一行 `[remediation] html attribute context escaped`。

### PR1｜规范化「拆长段落」压平各行内结构　A

- 位置：`src/Wisdom.Word/Processing/WordNormalizer.cs:468-492` 与 `:518-523`（`ReplaceParagraphText`，以及 `paragraph.Inlines.Clear()` 后只加一个 `TextRun`）。
- 现状：拆分只保护了域、图片、标题与列表，命中后整段折叠为单一 run，段中各处加粗与斜体、`TextRun.Link` 超链接、`BookmarkStart` 与 `BookmarkEnd` 范围全部丢失。
- 目标行为：按原 run 边界切分并保留各行内属性与书签范围；做不到就不拆，宁可少改也不丢结构。
- 红测：含书签范围、中部加粗与超链接的长段落触发拆分，断言往返后三类结构仍在。

---

## 6. P2 —— 健壮性与性能

| 编号 | 位置 | 现状与目标 | 红测要点 | 级 |
| --- | --- | --- | --- | --- |
| E5 | `src/Wisdom.Excel/ExcelCollection.cs:313`、`:370` | 循环上界 `row <= sheet.LastRow` 每次比较都触发全表扫描（`src/Wisdom.Office.Core/Model/SpreadsheetModel.cs:1054`），导出 N 行退化为 O(N²)。目标：进循环前取一次快照，同文件 CSV 路径已有正确写法。 | 用现有 `[complex-excel]` 计时风格做耗时阈值断言 | A |
| E6 | `src/Wisdom.Excel/Xlsx/XlsxStyles.cs:220-221` | 每个带样式单元格都 `return _styles[index].Clone()`。目标：读路径返回不可变共享实例，仅当调用方请求修改时才克隆。 | 计数型断言：样式表规模与克隆次数解耦 | A |
| E7 | `src/Wisdom.Excel/Xlsx/XlsxWriter.cs:52`；`src/Wisdom.Word/Docx/DocxWriter.cs:68`、`:116`；`src/Wisdom.Word/Formats/WordFormatConverter.cs:154` | `FileMode.Create` 先截断目标再边写边压 zip，中途任何异常都把用户原文件变成残缺。目标：同目录临时文件、写全成功、原子替换；失败保留原件并清理临时文件。 | 注入一次写盘中途异常，断言原文件字节未变 | A |
| E8 | `src/Wisdom.Excel/ExcelRange.cs`、`src/Wisdom.Excel/Xlsx/XlsxConditionalFormatBuilder.cs`（均为在飞改动） | 在飞批次已加区域面积上限、交叠粘贴快照、`sqref` 多区域分片。目标：Phase 0 提交后确认这些行为都有断言守着，缺则补。 | 交叠 `CopyTo`、多区域条件格式往返 | B |
| F8 | `src/Wisdom.Office.Core/Formula/Functions.Advanced.cs:586` 附近 | `reference.Kind == FxKind.Range ? reference.R1 : reference.R1` 两侧同值，是 `StampReference` 落地后的冗余残留。目标：直读 `R1`，并在文档注释说明标量已带坐标。 | 现有 `ROW`、`OFFSET` 单格断言须继续为绿 | A |
| W4 | `src/Wisdom.Word/Docx/DocxXml.cs:386-425` 配合 `src/Wisdom.Word/Docx/DocxReader.cs:212-219` | 无部件字节与 XML 缓存，每个继承页眉页脚的节都重新解压并重复累加 `_totalBytes`，合法多节文档可能触发 `MaxTotalBytes` 使 `Load()` 抛 `InvalidDataException`。目标：按部件名缓存解析结果，预算只计一次。 | 构造多节共享同一页眉的文档，断言可加载 | A |
| W5 | `src/Wisdom.Word/Docx/DocxTemplate.cs:112`（仅 `catch (IOException)`）对比 `:202`、`:251`、`:412`、`:443` 的 `Wml.LoadXml` | 畸形模板抛出的 `XmlException` 逃逸，违反声明的 `InvalidDataException` 契约。目标：统一包装为契约内异常类型。 | 截断或坏 XML 模板，断言抛 `InvalidDataException` | A |
| W6 | `src/Wisdom.Word/Docx/DocxPartReader.cs:745` 附近 | `w:gridCol` 宽度缺失或为 0 时列被丢弃，`table.Columns` 少于真实网格列数，后续 `gridSpan`、`gridBefore` 全部错位。目标：保留 0 宽列。 | 含 0 宽列的表格往返，断言列数 | B |
| C2 | `src/Wisdom.Word/Formats/HtmlReader.cs:1714-1728`（`DecodeText`） | 只认 BOM，其余一律按 UTF-8 解码，全文 `charset` 命中数为 0；声明 cp1252 或 GB2312 的无 BOM 页面静默乱码。目标：解析 meta charset，失败时走 `ResolveEncoding` 既有安全回退。 | cp1252 高位字节文本往返 | A |
| C3 | `src/Wisdom.Word/Formats/RtfReader.cs:267` | `\binN` 应按字节跳过，此处按已解码字符串的字符数推进；双字节代码页下后续组边界失步。目标：在字节层按声明长度消费。 | 多字节代码页加 `\bin` 的文档 | B |
| C4 | `src/Wisdom.Word/Formats/RtfWriter.cs:379-383`、`src/Wisdom.Word/Formats/OdtWriter.cs:480-483` | 列表标记写成字面文本加制表符，而两个 reader 都不回填 `ListMarker`，往返后变普通段落。HTML 路径已有 `.wisdom-list-marker` 正确处理，三格式不一致。目标：与 W2 对齐，或至少保证往返保真。 | 往返后断言段落仍为列表 | B |
| PR2 | `src/Wisdom.Word/WordFindReplace.cs:75` 对比 `:162` | 正则模式用随文化的 `RegexOptions.IgnoreCase`，普通模式用 `OrdinalIgnoreCase`，同一查询在 `UseRegex` 开关下匹配集合不同。目标：正则补 `RegexOptions.CultureInvariant`，两种模式语义统一，并在文档注释写明大小写规则。 | 设为土耳其语区域后两模式结果一致 | A |
| PR3 | `src/Wisdom.Word/WordMailMerge.cs:334`、`:406`（从 `i + 1`、`r + 1` 起找结束标记）、`:360-365` | 紧凑写法里 `MERGEFIELD TableStart:X` 与 `TableEnd:X` 同段或同行时永不配对，标记被清除、区域只渲染一次且无错误；落单标记还会抹掉同段其它区域。目标：先判同块结束、按名配对。 | 同段 TableStart 与 TableEnd 展开多行 | B |
| PR4 | `src/Wisdom.Word/WordTextRange.cs:109-113` | 多 run 替换把新文本写进首个 run 却保留其 `Link`，替换文本静默成为原超链接的一部分。目标：跨 run 替换时清除或显式处理链接。 | 命中已链接文本的替换，断言结果不挂链接 | B |
| PR5 | `src/Wisdom.Word/Processing/WordNormalizer.cs:831`、`:274` | `section.PageSetup` 与 `run.Text` 裸解引用，而同属性在 `src/Wisdom.Word/WordDocumentOperations.cs:50` 与 `WordTextRange.ReplaceRange` 处有判空；二者均可公开置空，规范化直接 NRE。目标：统一 guard，空值走默认。 | 置空 `PageSetup` 与 null 文本 run 后调用 | A |

---

## 7. P3（次要）与已排除误报

### 7.1 P3（逐条仍需红测确证后再修）

- `src/Wisdom.Excel/Xlsx/XlsxWriter.cs:294-298`：无条件写 `fullCalcOnLoad="1"`，大工作簿每次打开全量重算。目标：默认不写，仅在存在脏公式标记时写。BREAKING。
- `src/Wisdom.Excel/ExcelQuery.cs:408-411`：越界列把脏值写回 `lastColumnIndex`，同一行后续隐式列号整体错位。目标：丢弃该格但不污染游标。
- `src/Wisdom.Office.Core/Model/SpreadsheetModel.cs:1062-1065`：一个只含注释的空 `if` 死块。目标：删除或落实其语义。
- `src/Wisdom.Excel/Xlsx/XlsxWriter.cs:470` 附近：模型时间戳为默认值时仍写 `dcterms:created` 与 `modified`，输出字面 `0001-01-01T00:00:00Z`，严格 OOXML 校验拒收。目标：缺省不写。
- `CreateWorkbook` 零工作表时写出空 `<sheets/>`，Excel 拒开（OPC 要求至少一张）。目标：抛可预期异常，或补一张默认表。
- `src/Wisdom.Excel/Templates/XlsxTemplate.cs:311-313`、`:340-342`：`param` 文档注释只有中文，缺英文并列，违反项目注释约定。目标：补齐，并在 Phase 0 一并处理同批改动引入的同类缺口。
- `src/Wisdom.Word/Formats/HtmlReader.cs:229`：`!selfClosing && IsRawText(tagName)` 使自闭合 `script` 绕过 raw-text 通道，脚本源码被当正文并入模型。
- `src/Wisdom.Word/Formats/HtmlReader.cs:34`、`src/Wisdom.Word/Formats/WordFormatConverter.cs:25`、`src/Wisdom.Word/Docx/DocxReader.cs:41`：文件路径入口用无上限的 `File.ReadAllBytes`，而流式入口走 `OfficePackageLimits`。目标：入口统一限额。
- `src/Wisdom.Word/Formats/RtfWriter.cs:530-537`：图像逐字节 `ToString("x2")` 造成海量分配；`:84` 与 `:808-833` 的 ASCII 编码把非 ASCII 链接目标打成问号。
- `src/Wisdom.Word/Processing/WordLabels.cs:631-632`：填充高度为负时跳过填充行，标签盒溢出固定行高且后续页整体位移，且不记录任何 issue。
- `src/Wisdom.Word/Processing/WordDocumentDiff.cs:42-76`：实现是多重集比对而非 diff，没有「修改」概念、计数不构成划分、纯空白改动不可见；若保留该语义，命名与文档注释必须与之相符。
- `src/Wisdom.Word/Processing/WordRevisions.cs:112-121`、`:90-95`：`Count` 口径既不等于来源 `w:ins`/`w:del` 元素数也不等于写出数；`RejectAll` 把恢复内容追加到段尾，顺序错乱。
- `src/Wisdom.Word/WordMailMerge.cs:48` 对比 `:347-354`：`MaxRegionDepth` 只限层数不限展开总量，恶意模板仍可无界增长；且类型备注宣称任意嵌套深度，与实际静默降级相矛盾。
- `src/Wisdom.Office.Core/Utilities/OfficePackageLimits.cs:16-18`：公开的静态可变限额且无同步，并发下预算可被其他线程抬高，绕过解压炸弹防护。目标：改为不可变或按调用传入。

### 7.2 已排除（代理误报，工作区已修或本不成立）

| 代理结论 | 排除依据 |
| --- | --- |
| `ROW(A1)`、`OFFSET(A1,...)` 对单格引用恒报错 | `src/Wisdom.Office.Core/Formula/Evaluator.cs:390-417` 已引入 `StampReference`，把坐标盖到标量副本上 |
| `ROUND` 负 `num_digits` 被取绝对值当作 0 位 | `src/Wisdom.Office.Core/Formula/Functions.cs:249` 注释明确记为旧行为；`:256-275` 现为 `Math.Pow` 加正负 300 钳制 |
| `NumberAt` 把缺参当 0，导致所有可选默认值失效 | `src/Wisdom.Office.Core/Formula/Functions.cs:46-55` 现返回 `false`，注释说明正是为此前缺陷而改 |
| `MATCH` 省略 `match_type` 会走精确匹配 | 承接上一条：缺参返回 `false`，`Functions.Advanced.cs:463` 的默认 1 生效 |
| `RAND`、`RANDBETWEEN` 每次调用 `new Random()` | `Functions.cs:327-345` 已改为每线程实例，注释记录旧行为 |
| 空节导致同一 `pPr` 下出现两个 `sectPr` | 在飞批次改 `src/Wisdom.Word/Docx/DocxWriter.cs:388-411`，按本节是否写出块决定挂载点 |
| `MakeKey` 左移 20 位使列位串进行号 | `src/Wisdom.Office.Core/Model/SpreadsheetModel.cs:999-1013` 已改为 32 位，并附反例注释 |
| 写端可越界写出超 `XFD` 的列 | `SpreadsheetModel.cs:1015-1044` 已加 `RequireGrid` 系列，`GetCell`、宽高、隐藏行列入口拒绝越界（已发生的 BREAKING，需按 §11 声明） |
| 越界坐标写进 `dimension` 与 `row r` | 同上入口校验；反写的合并区域也已按 min/max 归一（`:1156-1167`） |
| `sqref` 多区域被整串 `LastIndexOf('!')` 截断只剩最后一片 | 在飞批次改 `src/Wisdom.Excel/Xlsx/XlsxConditionalFormatBuilder.cs:209-249`，逐片剥前缀 |
| 交叠 `CopyTo` 边读边写导致整体错位 | 在飞批次在 `src/Wisdom.Excel/ExcelRange.cs` 增加快照路径与 `Overlaps`、`CloneContent` |
| 逐格区域操作无面积上限 | `ExcelRange.cs:62-83` 的 `RequireWalkableArea`，上限与公式引擎一致 |
| 相邻两个 `w:tbl` 会被 Word 合并成一张表 | 在飞批次在 `DocxWriter.WriteBlocks` 插入空段落，即 §1 那条崩溃的出处 |
| CSV 与 HTML 导出未中和 `=cmd|` 类公式注入 | 运行器已有 `[hardening] csv formula switch verified`，`ExcelCsvOptions.EscapeFormulaLikeText` 存在 |
| `javascript:`、`data:` 等链接协议未过滤 | 运行器已有 `[hardening] file: and smb: links refused` |
| 全角数字与非法列字母等畸形引用被静默误解 | `SpreadsheetModel.cs:2152-2183` 已改为只认 ASCII 列字母与数字，并拒绝尾随垃圾 |
| `CompoundFile` 未释放等资源泄漏 | 该类以 `byte[]` 为底，不实现 `IDisposable` |
| 实例状态未同步的线程安全问题 | 唯一可变静态 `MapCache` 是 `ConcurrentDictionary`；代码与文档均无实例级并发承诺 |

---

## 8. 仓库卫生

| 现象 | 证据 | 建议处置 |
| --- | --- | --- |
| 一次失败的分支操作残留 | `stash@{0}`，标题为「正在签出 main 执行 main 签出前，暂存未提交的更改」 | 先比对其内容与在飞批次是否重复；确认无独有内容后丢弃，处理前导出留档 |
| 双解决方案文件 | 根目录 `Wisdom.Office.sln`（已跟踪）与 `Wisom.Office.sln`（未跟踪，目录名拼写为 `Wisom`） | 删除拼写错的未跟踪 sln，避免 IDE 载入错误的工程集合 |
| 拼写错的测试目录 | 未跟踪 `tests/Wisom.Office.Tests/` | 核对是否为历史副本，确认无用则删除 |
| 其它未跟踪产物 | `articles/`、`tmp-docverify/`、`bin/`、`obj/`、`out/`、`nupkg/`；八个包目录下各有 `.hbDebug/`（隔离构建输出，`git check-ignore` 确认未被忽略） | 确属构建产物的补进 `.gitignore`（`.hbDebug/` 应与 `bin/`、`obj/` 同列）；`tmp-docverify/` 此前已移出版本库；`articles/` 需明确是文档资产还是临时物 |
| 提交信息质量 | 历史含 `d554a22 123` 这类无信息量标题 | 本规范之后的提交沿用现有「scope: 描述」中文风格，不回改历史 |

---

## 9. 未审缺口

本轮只覆盖 `src/Wisdom.Excel`、`src/Wisdom.Office.Core/Formula`、`src/Wisdom.Word`。未审：`src/Wisdom.Ppt`、`src/Wisdom.Pdf`、`src/Wisdom.Markdown`，以及 `src/Wisdom.Office.Core` 其余子目录（`IO/PackageArchive`、`Templates/TemplateEngine`、`Utilities` 未涉及项）。这些包的审计应另起一轮，不与本规范混排，以免共享红测门禁时相互污染基线。

---

## 10. 实施阶段与依赖

- **Phase 0 建立基线**：前置条件是在飞批次能编译——当前它有两处 `CS0122`（§1），须由该改动线的作者收尾或明确弃写，本规范不代为改动未完成的特性。编译通过后把在飞批次按模块分组提交——公式引擎、`SpreadsheetModel` 与 `ExcelRange`、`XlsxConditionalFormatBuilder`、`DocxWriter` 与 `DocxPartReader`、主题与边框（`XlsxTheme`、`XlsxDecorationReader`、`XlsxStyles`、`XlsxWriter`）。每组提交前后各跑一次双目标运行器；随后按 §7.2 核对哪些发现已被这批改动覆盖，缺断言的按 E8 补齐。
- **Phase 1 P0**：E1、E2、F1、F2。四条相互独立；E2 与 F3 共用「单元格级异常不得升级为整簿失败」的收敛策略，建议同一批设计。
- **Phase 2 P1 引擎与 Excel**：F3 至 F7、E3、E4。F7 与 E1 都涉及日期与数值口径，先定口径再改调用点。
- **Phase 3 P1 Word**：W1、C1、PR1；W2 与 W3 属新增部件与域模型的里程碑级改造，单独排期。
- **Phase 4 P2**：§6 表中其余条目，建议顺序 E7、W4、C2、PR2、PR3、PR4、PR5、E5、E6、F8、W5、W6、C3、C4。E7 涉及文件替换语义，与 §7.1 的读入限额一并考虑。
- **Phase 5 P3 与文档同步**：§7.1 全表、注释双语缺口、§8 卫生项、§11 文档义务。

依赖提示：Phase 0 未完成前不要启动任何 P0 改动，否则无法区分「改坏的」和「本来就红的」。

---

## 11. 文档同步义务（行为变更必做）

下列变更合入同一里程碑（建议 1.1.0）时，必须同步改写文档；文档中只写仓库相对路径，不写发布步骤。

1. `GetCell`、宽高、隐藏行列对越界坐标改抛 `ArgumentOutOfRangeException`（在飞批次已引入）：`README.md:482` 的「已知限制」与 `docs/Wisdom.Excel.Tutorial.md` 不支持章节。
2. F7 数值与日期文本解析改 Invariant：`docs/Wisdom.Office.Tutorial.md`、`docs/Wisdom.Excel.Tutorial.md` 的公式求值说明。
3. E1 写出 1904 日期系统：Excel 教程的日期能力说明。
4. §7.1 去掉恒 `fullCalcOnLoad`：Excel 教程中关于打开即重算的描述。
5. W1、W2、W3 若本轮只做到部分保真：`docs/Wisdom.Word.Tutorial.md` 与 `README.md` 已知限制需明确「列表编号、复合域仍不支持」，或改述为已支持。
6. C1、PR3、PR4 等安全与保真变化：`articles/Wisdom.Word.md`、`articles/Wisdom.Excel.md` 对应能力段落。
7. 新增运行器 `[remediation]` 断言组：在 `docs/Wisdom.Office.功能导读.md` 的自验证说明处登记一条。

---

## 12. 验收清单

- [ ] Phase 0 分组提交完成，工作区不再有未解释的未跟踪源码改动。
- [ ] §4 至 §7.1 每条都有红测转绿测的记录，记录中写明修复前的失败信息原文。
- [ ] §7.2 每条都有守护断言防止回归；无法自动化守护的注明原因。
- [ ] `dotnet run --project tests/Wisdom.Office.Tests -f net8.0` 与 `-f net461` 均跑到末行 `All tests finished.`，全程无未捕获异常导致的早退。
- [ ] §11 文档同步逐条落地，文档内无本地绝对路径、无发布步骤。
- [ ] 新增与改动代码的注释中文与英文并列，公共成员文档注释完整。
- [ ] §8 卫生项逐条给出处置结论（保留、删除或进 `.gitignore`），空 stash 在确认内容后处理。
- [ ] 所有 BREAKING 条目（F7、E1 的行为修正、§7.1 的 `fullCalcOnLoad`、以及 §7.2 已发生的越界坐标拒绝）汇总进同一里程碑版本说明。
