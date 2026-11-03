# Wisdom.Office 热点缺陷修复规范 第二集（PDF / PPT / Word 出错路径）

日期：2026-11-02
范围：`src/Wisdom.Pdf`（读端、写端、字体、加密、渲染）、`src/Wisdom.Ppt`、`src/Wisdom.Word` 的**出错路径**。
与第一集的关系：第一集覆盖 Excel、公式引擎与 Word 的保真缺陷（编号 `E*`、`F*`、`W*`、`C*`、`PR*`），本集不重复那些条目，编号一律用 `PD*`（PDF）、`PT*`（PPT）、`WE*`（Word 新增），跨集衔接见 §9。门禁规则、置信度分级沿用第一集 §2、§3，本集只补一条崩溃类发现的方法学要求（§2）。

---

## 1. 背景与基线

四个审计代理并行通读上述范围，回报 56 条。逐条对当前工作区读码复核后，**49 条成立**（其中 A 级 41 条），6 条与第一集已登记条目同根，1 条因前提不可构造而排除。

| 事实 | 依据 |
| --- | --- |
| 本集红测必须等第一集 Phase 0 完成后才能跑 | 自检运行器 `tests/Wisdom.Office.Tests` 同时引用全部包；当前 Excel 在飞批次有两处 `CS0122`（第一集 §1），整个测试工程编译失败 |
| PDF 读端已有 14 个恶意输入用例 | `tests/Wisdom.Office.Tests/PdfHostileInputTests.cs`；语法嵌套、`/Length` 溢出、xref 条目号与字典规模等闸门**已核实存在并已覆盖**，本集不再重复申报（见 §8） |
| PPT 写端产出的结构错误未被现有断言发现 | 现有 PPT 断言按本库自己的私有形状自洽校验（例如动画只数 `presetClass="entr"`），文件级 XML 合法性从未断言 |
| Word 写端三个写出器的健壮性等级不一致 | `Formats/RtfWriter.cs:313`、`Formats/HtmlWriter.cs:178` 有 `depth > 16` 递归上限，`Formats/OdtWriter.cs` 与 `Docx/DocxWriter.cs` 完全没有；`HtmlWriter.cs:101` 对页眉判空，`RtfWriter.cs:126` 不判 |

---

## 2. 门禁补充：崩溃类发现要子进程隔离

栈溢出与挂起类条目（`PD3`、`PD7`、`WE1`）不能在主运行器里直接断言——`StackOverflowException` 不可捕获，会把整个测试进程带走，第一集 §3 第 3 条的「跑到末行」基线当场作废。约定：

1. 主运行器只断言**修复后**的正确行为（返回错误值或降级输出）。
2. 修复前的「会崩」证据由一条独立命令产生：把复现输入交给 `dotnet run --project tests/Wisdom.Office.Tests -f net8.0 -- --crash-probe <编号>` 这类子进程入口，主进程只检查其退出码非 0 且输出含指定标记行。
3. 挂起类用带超时的等待，超时才算红、不挂。

其余门禁（红测先行、双语注释、`[remediation]` 标记行、内存构造输入）照第一集 §3。

---

## 3. P0 —— 崩溃、挂起、进程死亡与静默毁数据

### PD1｜页树节点声明 `/Type /Pages` 却没有 `/Kids` 时直接空引用　A

- 位置：`src/Wisdom.Pdf/Reading/PdfDocument.cs:855-869`。
- 现状：`:856` `bool isPages = kids != null;`，`kids` 为 null 时 `:860` 仍可凭 `/Type` 把 `isPages` 置真，`:863` 的叶子分支被跳过，`:869` `for (int i = 0; i < kids.Count; i++)` 解引用 null。同包的兄弟实现 `src/Wisdom.Pdf/Document/PdfFile.Pages.cs:553-563` 对同一情形写了 `if (kids != null)`，说明这是漏守卫而不是设计。
- 后果：xref 与 trailer 完全合法、只是页树节点写成 `<< /Type /Pages /Count 0 >>`（手写或半截生成器的常见产物）的 PDF，`PdfDocument.Load()` 抛未处理的 `NullReferenceException`；`/Kids` 指向不存在对象（`Resolve` 返回 null，此时 `kids` 是 null 而非空数组）同样命中。
- 目标行为：`kids == null` 时按叶子页处理或跳过该节点，与 `PdfFile.Pages` 一致。
- 红测：内存拼一份最小 PDF（合法 xref + 无 `/Kids` 的 Pages 节点），断言 `Load` 不抛且页数为 0 或 1（按目标语义固定其一）。

### PD2｜对象流表头一处不合法就让整条流里的所有对象消失　A

- 位置：`src/Wisdom.Pdf/Reading/PdfDocument.cs:726-731`；抛点 `src/Wisdom.Pdf/Parsing/PdfParser.cs:607`（无数字即 `throw new InvalidDataException`）；吞点 `src/Wisdom.Pdf/Reading/PdfDocument.cs:693-696`（`LoadObject` 的 `catch (Exception) { return null; }`）。
- 现状：循环按声明的 `/N` 强行读 `2*count` 个整数，中途抛异常时 `:730` 已经写进 `offset` 的命中结果随异常一起丢掉，`LoadObject` 把整条流当成「对象不存在」。
- 后果：PDF 1.5+ 的对象流里 `/N` 与实际配对数不符（截断、增量更新工具的常见产物）→ 该流内所有对象（可能是 Catalog、页树、字体）全部读不出，页面缺失或整体退回全文扫描重建，且无任何提示。
- 目标行为：表头读到异常或非数字即停止迭代，保留已匹配到的偏移（读多少算多少）；只有目标对象确实没找到才返回 null。
- 红测：构造 `/N 3` 但表头只有 2 对的对象流，目标对象落在前 2 对内，断言修复前 `LoadObject` 返回 null、修复后返回该对象。

### PD3｜文本片段按字元累加拼接，二次方复制　A

- 位置：`src/Wisdom.Pdf/Reading/PdfContentInterpreter.cs:107` `_last.Text += text;`（`src/Wisdom.Pdf/Reading/PdfTextStructures.cs:11` 的 `Text` 是 `string` 自动属性）；同型问题在 Word 侧 `src/Wisdom.Word/Formats/OdtReader.cs:565` 与 `src/Wisdom.Word/Formats/RtfReader.cs:954`。
- 现状：`:99-105` 的合并条件对连续字元恒成立（`gap = flow - _right` 实际为 0，落在 `mergeGap` 内），于是一整行乃至一页字元全部并进同一个片段，每并进一个就整串重读重分配。
- 后果：一页内容流写 `BT /F1 12 Tf 20 700 Td` 后接 N 个 `(A) Tj` → N²/2 次字符复制加 N 次大对象分配。约 8MB 未压缩内容流（Flate 后可压到 100KB 级）即 N=10⁶ → 分钟级挂起或 OOM。文本片段路径没有任何数量上限：全文件只有 `src/Wisdom.Pdf/Reading/PdfContentInterpreter.cs:325` 对矢量段做 `if (_segments.Count > 250000) break;`，`Fragments.Add`（`:127`）无守卫。
- 目标行为：片段文本改 `StringBuilder` 累积、物化时一次性成型；给单页字元与片段总数设上限，超限按 `PD9` 的告警通道上报而不是静默截断。
- 红测：主运行器断言中等规模（如 5 万字元）提取耗时低于阈值；崩溃与挂起证据按 §2 走子进程。RTF 侧同样：`src/Wisdom.Word/Formats/RtfReader.cs:23` 的 `MaxNodes = 4000000` 是节点数闸门，不约束单节点字符量。

### PD6｜TrueType `maxp.numGlyphs` 未校验，为 0 时在保存路径抛越界　A

- 位置：`src/Wisdom.Pdf/Fonts/TrueTypeFace.cs:106`（`ReadU16` 越界返回 0，见 `:795`，构造函数不报错）与 `:327-328`（`bool[] include = new bool[GlyphCount]; include[0] = true;`）。
- 现状：字体构造被 `src/Wisdom.Pdf/Fonts/FontResolver.cs:617-626` 的 `catch (Exception)` 吞掉并降级，但 `BuildSubset` 不在任何 catch 内——它由 `src/Wisdom.Pdf/Fonts/PdfFont.cs:283` 在渲染与保存链路上调用。
- 后果：`numGlyphs` 字段被改成 0 或被截短的 TTF，任何走嵌入子集的产出（Word/文档转 PDF、`GenerateFormAppearances`）抛 `IndexOutOfRangeException`，整个保存失败。
- 目标行为：构造期 `GlyphCount <= 0` 即抛 `InvalidDataException` 交给解析器降级；`BuildSubset` 内再兜一层 `Math.Max(1, GlyphCount)`。
- 红测：内存构造一个 `numGlyphs=0` 的合法头字体，断言修复前抛 `IndexOutOfRangeException`、修复后按字体不可用的既定路径降级且保存成功。

### PD7｜自包含表格让 ODT 与 DOCX 写出递归到栈溢出，进程直接死亡　A

- 位置：`src/Wisdom.Word/Formats/OdtWriter.cs:716` `WriteBlock(xml, cell.Blocks[b], styles, null, false);`（`:442` `WriteBlock` → `:609` `WriteTable` → `:716` `WriteBlock`，全链无深度参数）；`src/Wisdom.Word/Docx/DocxWriter.cs:1536` `WriteBlocks(element, cell.Blocks, context, ...)` 同型，`CollectImages`（`:523`）也无深度。
- 现状：同库的 `src/Wisdom.Word/Formats/RtfWriter.cs:313` 与 `src/Wisdom.Word/Formats/HtmlWriter.cs:178` 都有 `if (blocks == null || depth > 16) return;`，读端 `src/Wisdom.Word/Docx/DocxPartReader.cs` 也有 `MaxBlockDepth`——只有这两个写出端没有。
- 后果：`TableCell.Blocks.Add(所属 Table)` 或任何深嵌套表格 → 递归到栈耗尽；`StackOverflowException` 不可捕获，宿主进程被杀；DOCX 侧还叠加第一集 `E7` 的「先截断目标文件」，用户的原文件同时变成残缺包。
- 目标行为：与 Rtf/Html 写侧对齐，加深度参数与 16–32 上限，超限输出空单元格并记告警。
- 红测：主运行器断言超限后不再无限递归（返回结果或抛声明内异常）；「修复前会崩」按 §2 用子进程退出码取证。

### WE1｜ODT 表格跨行跨列按乘积枚举槽位，单个畸形单元格即可假死　A

- 位置：`src/Wisdom.Word/Formats/OdtReader.cs:763-773`。
- 现状：`colSpan` 上限 `MaxColumns - column`（4096）、`rowSpan` 上限 4096，两个因子各自夹紧，`:767-772` 的双层展开却按乘积逐个 `grid.Covered.Add(...)`，乘积无预算；`:1450-1454` 的 `CountNode()` 只数 XML 元素，不数槽位。
- 后果：约 200 字节的 `content.xml` 里一个 `number-rows-spanned="4096" number-columns-spanned="4096"` 单元格即 16,777,216 次哈希插入；复制 20 个这样的单元格仍远低于 40 万节点闸门 → `OdtReader.Load` 假死或 OOM。
- 目标行为：给每表槽位总数设预算（超限停止展开、保留原跨度），或改为「起点 + 跨度」区间判定，不枚举槽位。
- 红测：内存构造上述 200 字节 ODT，断言加载在阈值内完成且表格结构正确。

### 崩溃类共性问题

`PD1`、`PD2`、`PD6`、`WE1` 同属「单个畸形元素毁掉整次读取」，与第一集 `E2`、`F3` 是同一收敛策略：**单元格/对象/元素级失败降级，不得升级为整簿整档失败**。建议这四条与 `E2`、`F3` 同一批设计错误上报通道（见 `PD9`）。

---

## 4. P1 —— PDF 写出端与结构

### PD4｜行归并对每个片段线性扫描全部已有行，二次方热点　A

- 位置：`src/Wisdom.Pdf/Reading/PdfPageExtractor.cs:74`（`for (int b = 0; b < builders.Count; b++)`）。
- 现状：命中即 `break`，不匹配时扫完全部行候选；片段已按 Y 升序排序（`:62`），「每片段各自成一行」必然全扫。
- 后果：20 万个不同基线的片段 → 约 2×10¹⁰ 次比较，数十秒到分钟级；与 `PD3` 叠加更糟。
- 目标行为：按基线容差投桶（`Dictionary<long, ...>` 或排序后只回看邻近 k 行），降到 O(n log n)。
- 红测：合成多行文本页，断言提取耗时随片段数近线性（阈值断言，沿用现有 `[complex-excel]` 计时风格改 `[remediation]`）。

### PD5｜聚合与重复解码没有任何总预算　A

- 位置：`src/Wisdom.Pdf/Reading/PdfPage.cs:92-100`（`/Contents` 数组逐个 `_document.DecodeStream(item)` 写进同一个 `MemoryStream`）；`src/Wisdom.Pdf/Reading/PdfContentInterpreter.cs:621`（Form XObject 每次 `Do` 重新解压，`depth` 上限只限层数不限扇出）。
- 现状：`src/Wisdom.Pdf/Reading/PdfDocument.cs:40-48` 只有 `_xref` 与 `_objects` 两个缓存，解码结果不进缓存；`src/Wisdom.Pdf/Parsing/PdfStreamDecoder.cs:46/75` 的 256MB 只约束**单个流**。
- 后果：一个解压后 200MB 的 Flate 流被 `/Contents [4 0 R × 50]` 引用 → 约 10GB 累计分配；或同一表单流被 `Do` 2000 次 → 2000 次完整 inflate。文件本身可只有几 KB。
- 目标行为：解码结果按对象号（或 `PdfStream` 实例）缓存；`GetContentBytes` 一类聚合出口共用一份文档级/页级预算，超限走 `PD9` 的告警与既定异常。
- 红测：内存构造重复引用同一压缩流的页面，断言解压次数等于流个数（计数型断言）且总字节预算生效。

### PD8｜交叉表修复后做增量保存，产出不可读文件　A

- 位置：`src/Wisdom.Pdf/Document/PdfFile.cs:958`（`RebuildXref()` 末尾 `_lastXrefOffset = 0;`，且该方法不设 `_rewriteRequired`）；`src/Wisdom.Pdf/Document/PdfFile.Writer.cs:397`（`if (incremental && _lastXrefOffset > 0) trailer.SetNumber("Prev", ...)`）；`:340` 增量段只写「已重写的变更行」。
- 现状：`_rewriteRequired` 只在源文档加密（`PdfFile.cs:460`）与 `SetEncryption`（`:474`）时置真。未加密但 `startxref` 损坏的文件走恢复路径后仍可增量保存：新段含少数对象、`/Size` 照全量取、**没有 `/Prev`**。
- 后果：按规范从最后一个 `startxref` 读起的阅读器看不到未改动的对象（页面树、内容流）→ 报文件损坏或显示空白。
- 目标行为：恢复路径要么把 `_lastXrefOffset` 修正为基段偏移，要么置 `_rewriteRequired` 强制全量重写。
- 红测：内存构造 `startxref` 被改坏的 PDF → `Load` → 删一页 → `SaveIncrementalToBytes()` → 用本库重开并断言页数正确；修复前断言页树不可达。

### PD10｜trailer 的 `/Root` 世代号被硬写成 0　A

- 位置：`src/Wisdom.Pdf/Document/PdfFile.Writer.cs:394`（`new PdfReference(_catalogNumber, 0)`，`:444` 同样）；世代丢失点 `src/Wisdom.Pdf/Document/PdfFile.cs:575`（只取 `rootReference.Number`）；对照恢复路径 `:952` 用的是真实世代，且对象头（`Writer.cs:187-188`）与 xref 行（`:155-156`）都保留真实世代。
- 后果：目录位于非 0 世代的文件（增量更新、签名文档、`RebuildXref` 找回的文件）保存后写出 `5 2 obj` + `00002 n` 行，trailer 却说 `/Root 5 0 R`。pdfium 按号命中而侥幸可开，按 (号, 世代) 严格校验的读者拒收。
- 目标行为：与编号一并保存 `_catalogGeneration`，`/Root` 用真实世代。
- 红测：内存构造目录世代为 2 的 PDF → 打开保存 → 解析输出 trailer，断言 `/Root` 世代等于对象头世代。

### PD11｜水印表单的 `/BBox` 用页面坐标，内容却画在阅读空间　A

- 位置：`src/Wisdom.Pdf/Document/PdfFile.Watermark.cs:299`（`BBox = [left bottom right top]`）对照 `:272-274`（`swapped`/`displayWidth`/`displayHeight`）与 `:348`（`cm` 把表单坐标映射到页面坐标）、`AnchorOf` 与 `DrawTiled` 全程不加 `left`/`bottom`。
- 后果：裁剪发生在表单坐标系的 `/BBox` 内。`[0 0 612 792]` 加 90° 旋转页：内容按 `displayWidth=792` 铺开，`/BBox` 只放到 612 → 水印右侧约 23% 被硬裁；`MediaBox` 原点非零（如 `[0 200 612 992]`）时下半部水印消失。现有断言 `tests/Wisdom.Office.Tests/PdfWatermarkTests.cs:122-126` 只查 `/Rotate`、补偿矩阵串与 q/Q 平衡，不模拟裁剪，所以测不出。
- 目标行为：`/BBox` 用表单空间 `[0 0 displayWidth displayHeight]`。
- 红测：旋转 90° 与非零原点页各加一条居中横排水印，断言输出 `/BBox` 与内容坐标范围一致。

### PD12｜`SetPageRotation` 写出非 90 倍数的 `/Rotate`，与自身文档注释和包内既有函数矛盾　A

- 位置：`src/Wisdom.Pdf/Document/PdfFile.Pages.cs:241-244`；同文件 `:234-235` 的文档注释声明「角度会规整到 0/90/180/270」，`:577` 已有 `NormalizeRotation` 且读端 `:551` 确实在用。
- 后果：`SetPageRotation(0, 45)` 写出 `/Rotate 45`（规范强制 90 的倍数，见 ISO 32000-1 §7.9.2.4），校验器直接判非法；读端 `:551` 会把它规整成 90，于是「写 45 读回 90」——往返断言今天就能通过，红测必须断言写出字节。水印按 `rotation == 90 || 270` 判方向，45 度页按不旋转算，与 `PD11` 同源放大。
- 目标行为：写端调用 `NormalizeRotation`。
- 红测：45/91/200/-90 四例，断言写出字节里的 `/Rotate` 值。

### PD13｜嵌入字体的 `/Flags` 把斜体写进 AllCap 位　A

- 位置：`src/Wisdom.Pdf/Fonts/PdfFont.cs:316-317` `int flags = 4 | (_face.IsItalic ? 64 : 0);`。
- 依据：ISO 32000-1 表 223（§9.6.2 字体描述符）按位值给出 `1 FixedPitch、2 Serif、4 Symbolic、8 Script、16 Nonsymbolic、32 Italic、64 AllCap、128 SmallCap、256 ForceBold`。本条初稿把 TrueType `OS/2 fsSelection` 的位义（`TrueTypeFace.cs:113-114` 用的 0x20/0x01）套到了 PDF 标志上，误记为「4 是 Italic、8 是 Bold」——**PDF 根本没有 Bold 位**，粗体由 `StemV`（本文件 `:328` 已按 `IsBold` 写 120/80）表达，`ForceBold` 是「强制加粗渲染」而非「该字体是粗体」，不可挪用。
- 后果：斜体子集写成 `/Flags 68` = Symbolic|AllCap，查看器按「全大写」处理该字体，跨平台文本检索与大小写还原会退化为大写；嵌入的是 Identity-H 自定义编码子集，`Symbolic`（4）本就是正确取值，无需再补 `Nonsymbolic`（16，两者互斥）。
- 目标行为：`4 | (_face.IsItalic ? 32 : 0)`。
- 红测：从样例 PDF 逐个抽出字体描述符的 `/FontName`、`/Flags`、`/ItalicAngle`，断言 `ItalicAngle != 0` 与 32 位互为充要、任何描述符都不得置 64 位、4 位恒置。用 `ItalicAngle` 而不是字体名做斜体判据，避免依赖解析器恰好选中哪个物理字体。

### PD14｜表单扁平化的外观资源名硬编码 `/Fm<序号>`，与同名条目相互顶替　A

- 位置：`src/Wisdom.Pdf/Document/PdfFile.Forms.cs:1127-1128`；`SetPageResource` 无条件覆盖（`src/Wisdom.Pdf/Document/PdfFile.Pages.cs:229`）；对照同库的正确写法 `src/Wisdom.Pdf/Document/PdfFile.Watermark.cs:577-589`（`while (group.ContainsKey(candidate))` 探测空闲名）。
- 后果：`FlattenForm()` 每轮序号从 0 起，两轮扁平化（第一次的结果再读入第二次扁平）后 `/Fm1` 覆盖 `/Fm1`，而第一轮追加的 `q…cm /Fm1 Do Q` 仍在 → 一个字段渲染成另一个字段的外观；第三方 PDF 常用 `/Fm0`、`/Fm1` 命名 XObject，一轮就能顶掉原页面图形。
- 目标行为：复用探测空闲资源名的逻辑，并避开注释原有 `/AP` 名。
- 红测：两轮扁平化后断言页面 `/Resources/XObject` 里两条目并存且各自内容流引用不混。

### PD15｜生成外观后未标记控件注释为已改，增量保存丢掉这次改动；`/AP` 读取漏 `Resolve`　A（显示后果 B）

- 位置：`src/Wisdom.Pdf/Document/PdfFile.Forms.cs:838`（`dictionary.Get("AP") as PdfDictionary`，不 `Resolve`）、`:844-845`（只 `MarkModified(field.ObjectNumber)`，控件号从未登记）；对照 `:673-674`、`:1252-1253` 都按注释号补标记。
- 后果：增量保存后文件里仍是替换前的 `/AP`，而 `/V` 已是新值；Acrobat 类阅读器优先取注释自身 `/AP`，显示旧文字。源 `/AP` 是间接对象时（Word/LibreOffice/Acrobat 产物普遍如此）判空并被内联字典整体替换，`/R`、`/D` 与状态映射丢弃。
- 目标行为：按控件号逐个 `MarkModified`；`/AP` 先 `Resolve` 再按需克隆回写。
- 红测：设值 + 生成外观 + 增量保存，断言输出字节里控件注释的外观流引用指向新对象。

### PD16｜PNG 预测器在正文短于一整行时返回空数组，整条流被清零　A

- 位置：`src/Wisdom.Pdf/Parsing/PdfStreamDecoder.cs:182`（`rowLength > data.Length` 才回退，等号放行）与 `:202-204`（`stride = rowLength + 1`，`rowCount = data.Length / stride` 得 0，`new byte[rowCount * rowLength]`）。
- 后果：xref 流或对象流的正文恰好等于一行长度（截断或生成器漏写滤波字节）时解码结果为空 → 交叉表 0 条条目 → 整档退回全文扫描，增量与已删除对象信息一并消失；`/ObjStm` 命中则该流内所有对象消失，与 `PD2` 后果叠合。
- 目标行为：`data.Length < stride` 时原样返回或按可用行解码，不得把「不完整」解码成「全零」。
- 红测：直接对解码器喂 predictor=12、正文恰为一行的输入，断言返回长度。

### PD17｜按偏移解析对象时不校验对象号，偏移错指即静默返回别的对象　A（触发概率 B）

- 位置：`src/Wisdom.Pdf/Parsing/PdfParser.cs:642-643`（两个 `ReadInteger()` 返回值直接丢弃），调用方 `src/Wisdom.Pdf/Reading/PdfDocument.cs:680`。
- 后果：重建路径 `ScanIndirectObjects`（`:570-603`）要求「空白+数字+空白+数字+空白+obj」，图像与 JPEG 正文里的伪命中会把某对象号指向流内部字节，`result[number]` 后写覆盖前写；`LoadObject` 于是交出从图像数据中部解析出的任意对象，全程静默。畸形 xref（两条同偏移）同理可把 `/Font 5 0 R` 换成任意内容。
- 目标行为：`ParseIndirectObjectAt(offset, expectedNumber)` 比对头部号，不一致即判条目失效并走重建。
- 红测：构造两条指向同一偏移的 xref 条目，断言后者不被当作目标对象返回。

### PD9｜提取器吞掉一切异常，超限与解密失败都伪装成「这页没有文字」　A

- 位置：`src/Wisdom.Pdf/Reading/PdfPageExtractor.cs:34-36`（空 `catch (Exception)`），随后 `:38-40` 照常用空片段回填，`src/Wisdom.Pdf/Reading/PdfPage.cs:160-166` 把空列表当「已解析完成」永久缓存。
- 后果：解压超限（`DecompressionLimitException`）、`PD2` 的资源取不到等，全部表现为「空白页」，`ExtractText()` 返回空串且无任何信号。现有断言 `tests/Wisdom.Office.Tests/PdfHostileInputTests.cs` 里的 `/Length` 溢出用例只断言「不得造假文本」，恰恰把静默空当成通过。
- 目标行为：区分「语法损坏（继续容错）」与「上限/解密/内存类失败（上抛或记入 `PdfPage`/`PdfDocument` 的告警集合）」；这条与 `PD1`/`PD2`/`PD6`/`WE1` 共用第一集 `E2`、`F3` 的降级通道设计。
- 红测：让某页内容流解压超限，断言 `ExtractText()` 返回空的同时告警集合里有一条可检记录。

### PD18｜命名 CMap 的 Type0 字体在无 `/ToUnicode` 时整页文本静默为空　A（代码路径）/B（命中率）

- 位置：`src/Wisdom.Pdf/Reading/PdfTextFont.cs:271-273`（`else if (_twoByte)` 分支在 `_toUnicode == null` 时无条件给 `string.Empty`）；`_twoByte` 仅由 Subtype 决定（`:77`），从不读 `/Encoding` 名称；证据是同文件 `:44` 声明、`:78` 赋值却全仓无读取点的死字段 `_identityLike`。
- 后果：`/Encoding /UniGB-UCS2-H` 这类「码位本身就是 UCS-2」的字体（部分国产驱动产物），文本能算出码位却一律给空串，用户看到「PDF 有字读不出」；改用 `/Identity-H` 时确实无法还原，属既定能力边界。
- 目标行为：对 `*-UCS2-*`、`*-UTF16-*`、`Identity-*` 按双字节大端直接产出字符（处理落单字节），其余非 Identity 编码显式记「无法还原」而非静默空。
- 红测：内存构造无 `/ToUnicode` 的 `/UniGB-UCS2-H` Type0 字体页，断言抽出预期字符。

---

## 5. P1 —— PPT：结构非法与读写不对称

### PT1｜备注母版写出非法元素 `p:nvGrpPr`，每个输出文件都带结构错误　A

- 位置：`src/Wisdom.Ppt/Pptx/PptxWriter.cs:1689`；同文件其余形状树一致用 `nvPr`（`:524`、`:561`），全仓 `nvGrpPr` 仅此一处。
- 现状：`CT_GroupShapeNonVisual` 要求子元素为 `cNvPr, cNvGrpSpPr, nvPr`。`notesMaster1.xml` 在 `WritePackage`（`:270`）无条件写出并被 `presentation.xml` 的 `notesMasterIdLst`（`:413`）引用。
- 后果：任何一次 `Save` 产物——哪怕只有一页空白幻灯片——都含该非法元素，PowerPoint 打开即命中校验/修复路径。
- 目标行为：改为 `P + "nvPr"`。
- 红测：最小一份单页演示文稿保存后解析 `notesMaster1.xml`，断言不存在 `nvGrpPr` 且 `p:grpSp` 的三个必需子元素齐备。这是本集**唯一一条改一行就消除全量产物错误**的条目，建议最先做。

### PT2｜`p:timing` 树形状不合 `CT_TL*`，带动画的输出大概率被判损坏　B

- 位置：`src/Wisdom.Ppt/Pptx/PptxWriter.cs:1820`（一个 `p:par` 里塞两个 `p:cTn`）、`:1833-1841`（`p:seq` 与根 `p:cTn` 平级挂在 `p:par` 下，PowerPoint 实际写 `par > cTn(nodeType="tmRoot") > childTnLst > seq`）、`:1827` 与 `:1837` 两个节点同用 `id="1"`、`:1814`/`:1828`/`:1838` 写 `dur="indef"`、`:1808`/`:1862`/`:1885` 写 `delay="indef"`（规范词法是 `indefinite`）、`:1805`/`:1831`/`:1869`/`:1882` 用 `type=` 承载 `clickEffect`/`mainSeq`（应为 `nodeType=`，仅 `:1883` 用对）、`:1872-1874` 把 `a:clrMapOvr` 放进 `p:setTo`。
- 后果：设置了动画的演示文稿在 PowerPoint 里触发修复；读端按同一套私有形状自洽解析（`src/Wisdom.Ppt/Pptx/PptxReader.cs:592-614` 向上找 `cond`），自产自检发现不了。
- 目标行为：按 `tnLst/par/cTn(nodeType="tmRoot")/childTnLst/seq(nodeType="mainSeq")` 重建，id 全局递增，`dur`/`delay` 用 `indefinite`，`set` 行为改 `p:attrNameLst` + `p:to/p:strVal`。
- 红测：给一个形状设淡入动画后保存，断言 `p:timing` 子元素路径与命名空间前缀符合上述形状（结构断言，不依赖 PowerPoint）；真机验证另列。

### PT3｜读端与写端字段不对称，「存进去再读回来就变」　A（逐条已核）

| 方向 | 字段 | 证据 |
| --- | --- | --- |
| 读不写 | `a:tc@rowSpan` 与 `hMerge`/`vMerge` | 读 `PptxReader.cs:868-869`、`:882`；写端只有 `gridSpan`（`PptxWriter.cs:1190`），全文件无 `rowSpan`/`vMerge` |
| 读不写 | `a:rPr@baseline` → `Superscript`/`Subscript` | 读 `PptxReader.cs:229-230`；写端仅在 `IsGenericDefault` 判定里出现（`PptxWriter.cs:1035`），从不输出 `baseline` |
| 写不读 | `a:pPr` 的 `lnSpc`/`spcBef`/`spcAft` | 写 `PptxWriter.cs:917-929`；读端 `PptxReader.cs:979-985` 只取 `marL`/`indent`/`algn`，全文件无这三个元素的读取 |
| 写死 | `a:bodyPr@wrap="square"` | 写 `PptxWriter.cs:798` 硬编码；读端不解析 `wrap`、`spAutoFit`、`noAutofit`、`fontScale`、`anchorCtr`（`:953` 只取 `normAutofit`） |
| 写不读 | `a:tblPr@firstRow`/`bandRow` 与 `tableStyleId` | 写 `PptxWriter.cs:1139-1141` 每张表恒写；读端不解析 `a:tableStyles`，`TableRow.IsHeader` 无来源 |
| 全不支 | `a:hlinkClick` 超链接 | `PptxReader.cs:221-246` 的字符属性全集里没有，写端也没有；模型只有主题色 `Hyperlink` |
| 全不支 | `p:style` 的 `fillRef`/`lnRef`/`fontRef` | 读端 `PptxReader.cs:170-176` 只在 `spPr` 下找 `fillRef`（真实文件里它是 `p:sp/p:style` 的子元素，永不命中）；写端把 null 落成 `a:noFill`（`PptxWriter.cs:750`、`:762`） |
| 静默降级 | `p:fld` 页码/日期字段 | 读 `PptxReader.cs:998-1002` 只取缓存文本；写端只有 `a:r` |
| 静默降级 | `p:grpSp` 组合与 `a:xfrm@rot`/`hflip`/`vflip` | 读 `PptxReader.cs:733-736` 展平并丢弃旋转翻转 |
| 静默降级 | 图表 / SmartArt / OLE / `p:cxnSp` 连线 | 读 `PptxReader.cs:699`、`:839` 整块跳过，部件、rels 与 Content_Types 在保存后全部消失 |
| 静默降级 | 外部关系（`r:link`）图片 | 读 `PptxReader.cs:798-805` 无字节即整形状丢弃 |
| 静默降级 | 版式 `SectionHeader` | `PptxReader.cs:1516-1546` 的 `InferLayoutKind` 从不返回该枚举，节标题页被重绑到别的版式 |
| 静默降级 | 主题 accent1-6 与 `fmtScheme` | 读端 `PptxReader.cs:353-358` 只登记 6 个基础相位、`:295-298` 解析成 RGB 后即丢；写端 `PptxWriter.cs:1559-1564` 一律用硬编码默认色（`:1409`） |

- 后果分档：`rowSpan`/`vMerge` 还额外造成**列错位**——补位空格一律追加到行尾（`PptxWriter.cs:1166-1170`）而不是按列游标插入，第 2 行数据整体左移一列；主题色被替换会让品牌模板在二次保存后全部配色走样。
- 目标行为（分档）：
  1. 结构错误先修（`PT1`、`PT2`、列错位补位）。
  2. 写端不要伪造模型未表达的意图：未显式声明表头就不写 `firstRow`/`bandRow` 与逐格填充（让 `tableStyleId` 生效）；无标记时不写 `buNone`（`PptxWriter.cs:942`，配合 `:932` 的 `any = true` 使 `pPr` 恒被写出）；`p:style` 未解析出的填充描边要区分「未设置」与「显式无填充」，不能一律 `a:noFill`。
  3. 读端补齐可对称的字段（`lnSpc`/`spcBef`/`spcAft`、`wrap`、`baseline`、`rowSpan`/`vMerge`、主题 accent），模型补 `Hyperlink`、`WordWrap`/`AutoFit`、`GroupShape` 与旋转翻转属里程碑级改造，单独排期。
  4. 不可还原部件（图表、连线、外链图片、被跳过的畸形页）必须**可见**：提供未知部件透传或返回 skipped-parts 清单，禁止把「跳过」变成「删除」。`PptxReader.cs:1726`、`:1733-1736` 现在遇到悬空 `r:id` 或畸形幻灯片直接 `continue`，页数无声缩水，随后 `Save` 把该页彻底写没。

### PT4｜非 PNG/JPEG 图片写成 `imageN.bin` + `application/octet-stream`　A

- 位置：`src/Wisdom.Office.Core/Model/OfficeImage.cs:86-88`（嗅探失败即 `.bin` / octet-stream）；`src/Wisdom.Ppt/Pptx/PptxWriter.cs:1082` 沿用该扩展名，`:2077` 据此登记 `Default`。
- 后果：GIF/BMP/TIFF/EMF/WMF 图片保存后在 PowerPoint 里显示为空或损坏（EMF/WMF 本可原生渲染）；读端同样只嗅探 PNG/JPEG，所以真实 GIF 走一遍 Load+Save 也变成 `.bin`。
- 目标行为：补常见魔数与 MIME，嗅探失败时保留原扩展名线索。
- 红测：GIF 与 BMP 往返，断言媒体部件扩展名与 `[Content_Types].xml` 的 `Default` 项。

### PT5｜只设了一处格式的 run 被钉死成默认字号与显式字体　A

- 位置：读 `src/Wisdom.Ppt/Pptx/PptxReader.cs:231-232`（`sz` 缺失时 `TextFormat` 保留模型默认）；写 `PptxWriter.cs:982-983`（`sz` 恒写）、`:989`（恒写 `a:latin`）；继承让位的判定是 `:1030-1036` 的 `IsGenericDefault`（要求全部属性为默认）。
- 后果：真实文件里 `<a:rPr b="1"/>`（只加粗、字号字体继承母版的 18pt）读入后 `Bold=true` 使格式不再是「纯默认」，写出时落 `sz="1100"` 与显式 Calibri → 该词从 18pt 缩成 11pt，占位符继承链被硬编码值覆盖。这与项目记忆里「构造器预置默认值击败 `??`」那类缺陷同源（第一集未覆盖 PPT 侧）。
- 目标行为：`TextFormat` 用可空字段表达「未设置」，只写真正设置过的属性。
- 红测：构造只带 `b="1"` 的 run 往返，断言输出 `a:rPr` 不含 `sz` 与 `a:latin`。

### PT6｜空文本占位符在保存时被整形丢弃　A

- 位置：`src/Wisdom.Ppt/Pptx/PptxWriter.cs:445`（`if (IsEmptyPlaceholder(shapes[i])) continue;`）与 `:492`（判据是文本为空或全空白）；模型侧 `src/Wisdom.Office.Core/Model/PresentationModel.cs:677` 里 `AddSlide(layout)` 自建的占位符就是空文本。
- 后果：只填标题、正文另设过底色或边框的版式页，保存后正文形状整体消失，形状计数与层叠次序随之改变。
- 目标行为：只跳过「由 `BuildPlaceholders` 自动补且未改动」的占位符（加来源标记），用户显式添加的形状必须写出。
- 红测：断言该形状往返仍存在且尺寸与填充保持。

### PT7｜零宽或零高的几何被判「读不到」，替换成默认大框　B

- 位置：`src/Wisdom.Ppt/Pptx/PptxReader.cs:1253`（`if (cx <= 0 || cy <= 0) return null;`）→ 回落 `:1266-1275` 的 `DefaultBox(phKey)`（`:1177-1188` 按幻灯片比例推算）。
- 后果：`<a:ext cx="7000000" cy="0"/>` 的水平分隔线保存后高度变成约 313pt（半个版面），叠加未知预设降矩形（`:783`）后细线变成铺半屏的矩形。
- 目标行为：允许 0 尺寸（负值才拒绝），区分「读不到」与「读到 0」；非占位符形状保留 0 尺寸。
- 红测：分隔线往返断言高度为 0。

---

## 6. P1 —— Word 新增出错路径（`WE*`）

### WE2｜ODT 读端 `content.xml` 的 `XmlException` 逃逸，破坏 `InvalidDataException` 契约　A

- 位置：`src/Wisdom.Word/Formats/OdtReader.cs:110`；同函数对缺失与为空都抛 `InvalidDataException`（`:108`、`:111`），对 `styles.xml`（`:117`）与 `meta.xml`（`:125-126`）都 `catch (XmlException)`；`Wml.LoadXml` 明确声明 `XmlException`（`src/Wisdom.Word/Docx/DocxXml.cs:285`）。
- 后果：`content.xml` 内容非法（如 `<office:text><<<`）的 .odt，`OdtReader.Load` / `WordFormatConverter.Load` 抛 `XmlException`，按 `InvalidDataException` 兜错的调用方（`src/Wisdom.Word/Formats/WordDocumentFormat.cs:171-179`）直接把异常漏给业务。与第一集 `W5`（`DocxTemplate.cs:112`）同一类，不同位置。
- 目标行为：包装为 `InvalidDataException` 并保留内层异常。
- 红测：内存 zip 构造非法 `content.xml`，断言异常类型。

### WE3｜页眉页脚在写出与统计路径上不判空，同库另有 6 处判空　A

- 位置（不判空）：`src/Wisdom.Word/Docx/DocxWriter.cs:465`、`:466`、`:668`、`:675`；`src/Wisdom.Word/Formats/RtfWriter.cs:126`、`:127`；`src/Wisdom.Word/Processing/WordNormalizeResult.cs:130`（同方法内部 `:138` 就有 `if (blocks == null) return;`）。
- 位置（判空）：`src/Wisdom.Word/Docx/DocxWriter.cs:368`、`Formats/HtmlWriter.cs:101`、`Formats/OdtWriter.cs:825`、`WordDocumentWalker.cs:38`、`WordMailMerge.cs:303`、`WordDocumentOperations.cs:50`。
- 后果：`Section.Header` 是公开可写属性（`src/Wisdom.Office.Core/Model/DocumentModel.cs:852`），置 null 后 DOCX 与 RTF 写出抛 `NullReferenceException`，而 HTML 与 ODT 同输入正常输出——四个写出器行为不一致；DOCX 侧叠加 `E7` 的文件截断。
- 目标行为：写出入口一次性把各节的页眉页脚规范化（缺失即视为禁用），六个直接解引用点消除。
- 红测：`Header = null` 的文档过四个写出器，断言全部成功输出。

### WE4｜`ImageRun.Image == null` 抛 `ArgumentNullException`，本意为「输出空运行」的降级分支不可达　A

- 位置：`src/Wisdom.Word/Docx/DocxWriter.cs:1277`（`context.GetImageRelationship(image.Image)`）→ `:226`（`_imageIds.TryGetValue(image, out id)`，字典 null 键在 net461 与 .NET 上都抛）；`:1278` 的 `if (relationshipId == null) return new XElement(w + "r");` 正是被跳过的降级分支。
- 对照：同文件 `:513` 与 `Formats/RtfWriter.cs:468`、`Formats/HtmlWriter.cs:316` 都显式判 `image.Image != null`，说明 null 是被预期可能的值。
- 目标行为：`GetImageRelationship` 首行判空返回 null，或调用点补判。
- 红测：`new ImageRun(null)` 进出 DOCX 写出，断言输出空运行且不抛。

### WE5｜ODT 写出器不防 `document.Metadata == null`，其余三个写出器都防　A

- 位置：`src/Wisdom.Word/Formats/OdtWriter.cs:70`（`BuildMeta(document.Metadata)`）→ `:118` 直接解引用。
- 对照：`Formats/HtmlWriter.cs:56` 与 `Docx/DocxWriter.cs:1870`、`:1900` 用 `?? new DocumentMetadata()`，`Formats/RtfWriter.cs:232` 用 `if (metadata == null) return;`；`OfficeDocument.Metadata` 公开可写（`DocumentModel.cs:925`）。
- 目标行为：`BuildMeta` 入口 `?? new DocumentMetadata()`。
- 红测：`Metadata = null` 过四个写出器，断言全部成功。

### WE6｜查找替换的正则超时以未声明异常穿出，且文档已处于半替换状态　A

- 位置：`src/Wisdom.Word/WordFindReplace.cs:75`（构造 `Regex` 时传入 2 秒 `RegexTimeout`，`:37`）；方法文档只声明 `ArgumentNullException` 与 `ArgumentException`（`:64-65`）。
- 后果：段落 1 `"ab"`、段落 2 `"aaaaaaaaaaaaaaaaaaaaaaaaaaaaab"`，`UseRegex` 下用 `(a+)+b` 替换——段落 1 已改，段落 2 触发 `RegexMatchTimeoutException`（派生自 `TimeoutException`）穿出，调用方拿到的是被就地改到一半的文档、丢失已完成计数、无法回滚（`WordDocumentWalker.ForEachParagraph` 无事务）。
- 目标行为：外层转成已声明的 `ArgumentException` 并说明超时；先在副本上完成全部替换再一次性提交，或显式提供原子性/`FirstOnly` 语义。
- 红测：上述两段文档 + 灾难性模式，断言抛 `ArgumentException` 且文档字节级未被改动。注意这条与第一集 `PR2`（随文化 `IgnoreCase`）同位置不同问题，两条应同批改动。

### WE7｜DOCX 非主部件解析失败一律降级为 null，内容静默丢失　A（机制）/B（影响面）

- 位置：`src/Wisdom.Word/Docx/DocxXml.cs:394-396`（`catch (XmlException)` 后 `return null`）；消费点把 null 当缺失并继续：`Docx/DocxReader.cs:133`（styles 与 numbering 一起丢）、`:285`（页眉页脚整块消失）、`:341`、`:381`（元数据丢弃）、`WordSmartArtReader.cs:91`/`:101`、`WordDigitalSignatureReader.cs:147`。
- 后果：把正常 docx 的 `word/header1.xml` 删掉一个 `>` 重新打包 → `Load` 成功返回，`Sections[0].Header.Blocks.Count == 0`，调用方无从分辨「本来没有页眉」与「页眉被丢弃」；`styles.xml` 同理丢全部样式派生格式。主部件非法时至少抛 `InvalidDataException`（`DocxReader.cs:123-125`）。
- 目标行为：区分「部件缺失」与「部件存在但解析失败」，后者计入可观测的 `Warnings` 或抛 `InvalidDataException`；不要用被吞掉的 null 兼作「不存在」判定。
- 红测：损坏 `header1.xml` 的 docx 往返，断言加载结果带一条告警。

### WE8｜ODT 写出器与读端绕过 `PackageArchive.OpenForRead`，net461 上异常类型分歧　B

- 位置：`src/Wisdom.Word/Formats/OdtReader.cs:70`、`:84`（两处 `new ZipArchive(...)`）；`src/Wisdom.Office.Core/IO/PackageArchive.cs:8-15` 的类注释正是为消除该分歧而写（.NET Framework 的 `ZipArchive` 在惰性枚举条目时才抛 `ArgumentException`，`:35-44` 统一转 `InvalidDataException`）。DOCX 侧全部走 `OpenForRead`（`Docx/DocxReader.cs`、`WordSmartArtReader.cs:84`、`WordDigitalSignatureReader.cs:98`、`Formats/WordDocumentFormat.cs:141`）。
- 后果：中央目录含非法条目名的 .odt，net461 上抛 `ArgumentException`、net7+ 上抛 `InvalidDataException`，双目标行为不一致（第一集 §3 第 3 条要求双目标同结果）。
- 目标行为：`OdtReader` 两个入口改用 `PackageArchive.OpenForRead`。
- 红测：构造非法条目名的 zip，断言 net8.0 与 net461 抛出同一类型。

### WE9｜单张内嵌图片超限抛异常中止整篇读取，与其余通道口径不一致　B

- 位置：`src/Wisdom.Word/Formats/OdtReader.cs:1482`、`:1487`（`throw new InvalidDataException`），同方法 `:1481` 的数量超限却是 `return null` 跳过；`src/Wisdom.Word/Formats/HtmlReader.cs:1219`、`:1229` 一律跳过（丢图不丢文档）。
- 后果：结构完全合法、含一张大于 64MB `draw:base64-image-data`（扫描大图）的 .odt 整篇读失败，正文与表格全部读不到。
- 目标行为：改为跳过该图并记告警，总量保护交给 `OfficePackageLimits` 的 `_totalBytes` 预算。
- 红测：内存构造超限图的 ODT，断言正文仍可读且带告警。

---

## 7. P2 与 P3

| 编号 | 位置 | 现状与目标 | 级 |
| --- | --- | --- | --- |
| PD19 | `src/Wisdom.Pdf/Fonts/FontResolver.cs:593` | 锁外读 `_fileBytes.TryGetValue`，而 `:596`、`:651` 在 `lock (_sync)` 内写；`:14-16` 的文档注释声明「线程安全由单一锁保证」，`FontResolver.Default`（`:93`）是进程级单例。并发转换不同文档时首次解析同一字体可撞内部扩容，读到撕裂状态。目标：读入锁或改 `ConcurrentDictionary`。 | A |
| PD20 | `src/Wisdom.Pdf/Document/PdfFile.Watermark.cs:496-499` | 平铺行列范围由 double 直接强转 int，`ImageWidth`/`FontSize` 取极小值时比值超过 `2^31` 饱和成 `int.MinValue`，前 600 个 tile（`MaxTileDraws`，`:18`、`:502`）全落页面外坐标 → 水印静默消失。入口只校验符号（`:472`、`PdfWatermark.cs:167`）。目标：步长设下限并在 int 化前夹紧。 | A |
| PD21 | `src/Wisdom.Pdf/Parsing/PdfParser.cs:463` → `src/Wisdom.Pdf/Syntax/PdfObjects.cs:379` | `PdfDictionary.Set`/`Get` 是插入序线性表全表扫描（文件头注释自认线性），解析大字典为二次方；trailer 或 `/Resources` 里塞十万键即在任何上限生效前挂住。目标：内部加名称索引（保留写出顺序）或对单字典键数设上限。 | A |
| PD22 | `src/Wisdom.Pdf/Reading/PdfContentInterpreter.cs:610`、`:621` | Form XObject 递归深度上限 12 却不限扇出，与 `PD5` 同根；目标：单页累计解码字节预算，与 `PD5` 同一改。 | A |
| PT8 | `src/Wisdom.Ppt/Pptx/PptxWriter.cs:1502-1503`、`PptxReader.cs:1516-1546` | 版式识别与版式几何只覆盖 6 种，`SectionHeader` 写端有专门几何却永远走不到（`PT3` 表末行）。目标：读端保留原始 `p:sldLayout@type` 与 `cSld@name`。 | A |
| PT9 | `src/Wisdom.Ppt/Pptx/PptxReader.cs:1126-1130`、`PptxWriter.cs:1728`、`:1755` | 切换效果只认 4 种，`dir`、`spd`、`advClick` 被写死（上推还原成左推、「与上一动画同时」变成必须点击）。目标：扩充枚举并原值往返。 | A |
| PT10 | `src/Wisdom.Ppt/Pptx/PptxWriter.cs:2014-2017` | `docProps/app.xml` 元素次序与 PowerPoint 不同；该部件不在打开校验关键路径，列为次要。 | A |
| WE10 | `src/Wisdom.Word/Formats/RtfWriter.cs:146`、`:152-153`、`:183`/`:741`、`:779`；`Docx/DocxWriter.cs:895`、`:991` | 写侧对公开可空的模型字段（`TextRun.Format`、`TextRun.Text`、`Table.Borders`）直接解引用，而同文件别处已按可空处理（`:195` `if (format == null) return "Arial";`、`:483` `format ?? new TextFormat()`、`:183` `if (borders == null) return;`）。目标：写侧入口对模型做一次规范化快照。触发需显式置 null，故列 P2。 | B |
| WE11 | `src/Wisdom.Word/Formats/RtfWriter.cs:801` | 非 BMP 字符写 `\u55357`，超出 RTF 有符号 16 位；本库读端（`RtfReader.cs:619-621` 负值加 65536）能自洽回读，属输出规范性问题。目标：改代理对或多字节通道。 | B |

### 7.1 本集排除项

| 代理或直觉结论 | 排除依据 |
| --- | --- |
| `OdtReader.cs:1475` 逐部件重读并重复累加 `_totalBytes` | 与第一集 `W4`（`DocxXml` 无部件缓存导致 `MaxTotalBytes` 误触发）同根同因，不重复登记 |
| `DocxTemplate.cs:251`、`:412`、`:443` 的 `Wml.LoadXml` 未 catch | 与第一集 `W5` 同根，仅位置不同 |
| `DocxWriter.cs:68`/`:116` 先截断目标 | 第一集 `E7` 已登记；本集只把「`document == null` 的校验发生在 `:98`，即截断之后」作为 `E7` 的升级证据（§9） |
| `DocxPartReader.cs:822` `column += cell.ColSpan` 缺 `Math.Max(1, …)` | `TableCell` 构造器把 `ColSpan` 置 1（`DocumentModel.cs:642`），`gridSpan` 分支 `:790` 已夹紧，不可能为 0 或负 |
| `DocxWriter.cs:1999` `Convert.ToString` 可能抛 `InvalidCastException` | 上游 `CustomPropertyCollection.SetValue` 用 `IsSupportedValue` 预校验并抛已声明的 `ArgumentException`（`Model/CustomProperties.cs:322-327`），无绕过入口 |
| `WordDocumentBuilder.cs:381` 书签索引器可能抛 `KeyNotFoundException` | `BookmarkCollection` 字符串索引器未命中返回 null（`Model/BookmarkModel.cs:342-352`），`:382` 随即判空 |
| PDF 语法嵌套、`/Length`、xref 条目号、字典规模无上限 | `PdfParser.cs:207`（深度 64，`finally` 回退）、`:688`/`:736`（先减后比）、`PdfReadLimits.MaxObjectNumber`/`MaxXrefEntries` 三处闸门（`PdfDocument.cs:341`、`:440`、`:516`、`:522`）齐备，且 `PdfHostileInputTests.cs` 已有对应断言 |
| PDF `/Prev` 链、页树、Form 引用成环死循环 | `PdfDocument.cs:233`、`:847` 的 `visited.Add` 闸门、`PdfContentInterpreter.cs:610` 深度上限、`LoadObject` 的 `_loading` 环检测（`:669`） |
| PDF 解压各分支缺上限 | `PdfStreamDecoder.cs:240` 的 `EnsureCapacity` 在 RunLength `:371`/`:379`、LZW `:440` 逐次计费，Flate 走带预算的 `Zlib.Decompress`（`Wisdom.Office.Core/Utilities/Zlib.cs:70`、`:87`、`:126`）；真正缺的是**聚合预算**，即 `PD5` |
| PDF 加密原语错误：RC4/AES、R6 迭代哈希、`/P` 掩码、口令非常数时间比较、`/ID` 缺失 | `PdfCrypto.cs:16-42`、`:194`、`:315` 与 `PdfStandardSecurityHandler.cs:666-686`、`:729`、`:162`、`:229` 逐条比对 pypdf 与 PDFBox 常量一致；`PdfFile.Writer.cs:398-401` 两条路径都写 `/ID` |
| PDF 写出结构错误：`/Length` 与实际不符、缺 `ToUnicode`/`/W`/`CIDToGIDMap`/`/Length1`、增量更新改世代 | `PdfObjectWriter.cs:219` 压缩后重算长度并先 `Clone()`；`PdfFont.cs:368`、`:394-423`、`:347-357`、`:358-359`、`:310` 齐备；`PdfFile.Writer.cs:71-77`、`:367-368` 符合 ISO 32000-1 §7.5.5 |
| TrueType 表目录越界崩溃、`hhea.numberOfHMetrics=0` 负索引、`UnitsPerEm=0` 除零、OS/2 字段偏移差 2 | `TrueTypeFace.cs:73`、`:83-87`、`:793-825`、`:261`、`:264`、`:91`、`:112-133` 逐项核对；越界读返回 0，构造期异常被 `FontResolver.cs:617-626` 吞掉。唯 `BuildSubset` 不在该保护内，即 `PD6` |
| PPT `mc:Choice Requires="p14"` 属性名错、rId 冲突、`EndsWith("/slide")` 误命中版式、CT 元素次序、单位换算 | `Requires` 是 `mc:Choice` 上的非限定属性（PowerPoint 自身亦如此写）；`PptXml.cs:342-356` 后缀须完整以 `/slide` 结尾；`PptxWriter.cs:389-411`、`:1432` 的 id 区间无冲突；CT 次序除 `PT1`/`PT2` 外逐项核对无误；EMU/磅、`sz` 百分磅、`lnSpc` 十万分、`adj` 万分换算对称 |
| PPT 无裸 `int.Parse`、无 zip 炸弹与 XXE 缺口 | `PptxReader.cs:1789-1797` 用 `ParseLong`，`:608` 的 `TryParse` 有 `&& delayMs > 0` 保护；`PptXml.cs:229-240` `DtdProcessing.Prohibit` + `XmlResolver = null`，`OfficePackageLimits.ValidateEntryName` 在 `PptxWriter.cs:329`、`:342` 逐条校验 |
| Word `HtmlReader` 的 data:URI 切片、`@page` 切片、CSS 声明解析可能越界 | `HtmlReader.cs:1213-1229` 先判 `comma < 0` 并 catch `FormatException`；`:957`、`:1509-1519`、`:1349-1355` 的起点判定均保证长度与冒号存在 |
| Word `WordNormalizer.cs:294` `while (true)` 不前进 | 每轮 `run.Text.Remove(previousChar, 1)` 严格减小总字符量，未命中时 `GetLastVisibleChar` 返回 `'\0'` 并在 `:283` 早退 |
| Word `Legacy/DocReader.cs:514`、`:282`、`:301` 空 `catch` 吞异常 | 吞掉的是元数据与编码探测的可选信息，属已知容错设计且无上报通道；上报通道由 `WE7`/`PD9` 那条统一改造引入 |

---

## 8. 未审缺口

本集覆盖 PDF 全量、PPT 全量、Word 的出错路径。仍未审：`src/Wisdom.Markdown`、`src/Wisdom.Office.Core` 的 `IO/PackageArchive`、`Templates/TemplateEngine`、`Utilities` 未涉及项；PDF 的**渲染保真**（`Rendering/PdfRenderer.cs`、`TextLayout.cs` 与 Word→PDF 的分页规则）只审了出错路径，版面对错需要真机比对，另起一轮；PPT 与真实 PowerPoint 的互操作（动画、切换、主题）需要真机验证，本集只做结构断言。

---

## 9. 与第一集的衔接

- `E7` 升级为 P0：`src/Wisdom.Word/Docx/DocxWriter.cs:68` 在校验之前截断目标，`document == null` 直到 `:98` 才判——`Write(null, path)` 直接把用户原文件变成 0 字节。本集 `PD7`、`WE3` 的崩溃都落在这条截断之后，危害叠加。改动顺序上先修 `E7`（内存构建成功后再落盘），再修各崩溃项，红测才有意义。
- `W4`、`W5`、`PR2` 与本集 `WE7`、`WE2`、`WE6` 是同族（部件缓存缺失、异常契约违规、正则语义），建议同批设计，共享一个「文档级告警与降级通道」的新抽象；`PD9` 也接进同一通道。
- 第一集 §3 的门禁与 §11 的文档同步义务对本集同样有效；本集新增的 BREAKING 条目见下表。

| 编号 | 变更 | 需同步的文档 |
| --- | --- | --- |
| `PD12` | `SetPageRotation` 从不规整改为规整（或抛异常），45 度输入的结果变化 | `docs/Wisdom.Pdf.Tutorial.md` 页面旋转段；`README.md:482` 已知限制 |
| `PD13` | 斜体嵌入子集的 `/Flags` 从 `68`（Symbolic|AllCap）改为 `36`（Symbolic|Italic），产出字节变化 | `docs/Wisdom.Pdf.Tutorial.md` 字体嵌入段无需新增承诺；仅在已知限制里说明粗体只由 `StemV` 表达、规范无 Bold 位 |
| `PD1` | 声明 `/Type /Pages` 却取不到 `/Kids` 的节点，从抛 `NullReferenceException` 改为零页降级 | `docs/Wisdom.Pdf.Tutorial.md` §10 与 `src/Wisdom.Pdf/README.md` 已补「空容器按零页降级」；与兄弟实现的语义分歧记在 §11.2 |
| `PD6` | 零字形字体从「保存链路中途越界」改为构造期抛 `InvalidDataException`，保存链路整体降级不再炸 | 内部类型，不新增公共承诺；`docs/Wisdom.Pdf.Tutorial.md` §10 与 `src/Wisdom.Pdf/README.md` 已补「构造期拒绝」 |
| `WE1` | ODT 畸形跨度乘积从假死/OOM 改为「预算用完停止登记覆盖」，声明的跨度保留 | `README.md` 不可信输入列举、`src/Wisdom.Word/README.md` 写明降级代价（已同步） |
| `PD16`、`PD20`、`PT4` | 超限与畸形输入从静默产出空结果改为可检测失败 | `docs/Wisdom.Pdf.Tutorial.md`、`docs/Wisdom.Ppt.Tutorial.md` 的容错说明 |
| `PD9`、`WE7`、`WE9` | 引入告警集合后，读入失败的返回语义变化 | 两个包的教程与 `articles/Wisdom.Pdf.md`、`articles/Wisdom.Ppt.md` |
| `PT1`、`PT2`、`PT3` 结构档 | 输出文件的 XML 形状改变（历史产物可能被新版本读不出） | `docs/Wisdom.Ppt.Tutorial.md` 需写明往返保真范围变化；若旧形状需保留读取兼容，一并声明 |

---

## 10. 实施阶段与依赖

- **前置**：第一集 Phase 0 完成、测试工程能编译。本集任何红测在此之前都跑不起来（§1）。
- **Phase 1 一行即消除全量产物错误**：`PT1`（`nvGrpPr` → `nvPr`）、`PD12`（改用 `NormalizeRotation`）、`PD13`（`/Flags` 位义）。三条都是确定性替换，先做，随后跑双目标全量断言。
- **Phase 2 崩溃与挂起**：`PD1`、`PD2`、`PD6`、`PD7`、`WE1`、`WE4`、`WE5`，连同升级后的 `E7`。`PD7`/`E7` 顺序上先 `E7`。
- **Phase 3 性能与资源上限**：`PD3`、`PD4`、`PD5`（含 `PD22`）、`PD21`；与第一集 `E5`、`E6`、`F2` 共用「快照 + 预算」的同一批工具。崩溃与挂起证据按 §2 走子进程。
- **Phase 4 产出损坏与静默丢失**：PDF `PD8`、`PD10`、`PD11`、`PD14`、`PD15`、`PD16`、`PD17`、`PD18`、`PD9`；PPT `PT2` 与 `PT3` 的结构档、列错位补位、`PT4`、`PT5`、`PT6`；Word `WE2`、`WE3`、`WE6`、`WE7`。
- **Phase 5 读写对称改造（里程碑级，可拆版）**：`PT3` 的读端补齐与模型扩展（超链接、组合与旋转、自动换行、主题 accent、图表透传）、`PT8`、`PT9`、`PT7`；与第一集 `W2`、`W3` 同属「模型要长新部件」一档，合并排期。
- **Phase 6 P2/P3 与文档**：`PD19`、`PD20`、`PT10`、`WE10`、`WE11`、`WE8`、`WE9` 与 §9 的文档同步义务。

依赖提示：`PD5` 与 `PD22`、`PD9` 必须同批设计（预算与告警是一件事的两半）；`PT3` 的「写端不伪造意图」不依赖模型改造，可以先于 Phase 5 落地。

---

## 11. 验收清单

- [ ] §3 至 §7 每条都有红测转绿测记录，记录写明修复前的失败信息原文；崩溃与挂起类附子进程退出码证据。
- [ ] `PT1` 落地后，最小单页产物断言 `notesMaster1.xml` 结构合法；PPT 增加一条「写出后可被本库完整读回」的字段级往返断言，作为 `PT3` 的长期守门。
- [ ] 崩溃类（`PD1`、`PD2`、`PD6`、`PD7`、`WE1`）修复后，同一输入在 net8.0 与 net461 上结果一致（`WE8` 正是这条的例外来源）。
- [ ] 双目标运行器跑到末行 `All tests finished.`，全程无未捕获异常导致的早退（第一集 §3 第 3 条）。
- [ ] §9 表中 BREAKING 条目逐条同步文档，文档内无本地绝对路径、无发布步骤。
- [ ] 新增与改动代码注释中英并列，公共成员文档注释完整；`WE6` 的方法文档补上其真实抛出的异常类型。
- [ ] §7.1 排除项中「已有闸门」的部分补一条守护断言防回归；无法自动化的注明原因。

### 11.1 Phase 1 落地记录（2026-11-02）

断言落在 `tests/Wisdom.Office.Tests/RemediationPhaseOneTests.cs`，标记行 `[remediation]`，接在 `Program.Main` 的 PDF 段之后。运行器是 fail-fast 的，所以三条修复前的失败信息是逐条放开取证的原文：

```text
Test failed: notes master writes the illegal element p:nvGrpPr
Test failed: SetPageRotation(45) did not write /Rotate 90
Test failed: font descriptor flags: /AAAAAE+Arial-ItalicMT forges the AllCap bit (/Flags 68) | /AAAAAE+Arial-ItalicMT is italic but lost the Italic bit (/Flags 68)
```

修复后 net8.0 与 net461 双目标均跑到末行 `All tests finished.`，产物字节复核：样例 PDF 的五个描述符由 `4 × /Flags 4 + 1 × /Flags 68` 变为 `4 × /Flags 4 + 1 × /Flags 36`；`SetPageRotation(0, 45)` 写出 `/Rotate 90`、`-90` 写出 `/Rotate 270`；`notesMaster1.xml` 的 `p:nvGrpSpPr` 现在是 `cNvPr / cNvGrpSpPr / nvPr` 三子元素。

两点副产物记录在案：

- `PD13` 初稿的位表是错的（把 TrueType `fsSelection` 当成了 PDF 标志），已就地改写；**PDF 没有 Bold 位**，粗体由 `StemV` 表达，本集不再追查「Bold 从不置」这条伪缺陷。
- 形状树兼容性：本库读端只按名字取 `p:nvPr`（`PptxReader.cs:1326-1336` 的占位符定位），从不解析 `p:nvGrpSpPr` 的子元素名，因此 `PT1` 不会打断本库自己的往返；风险只落在外部生产者写下非规范子元素、又指望本库宽容读取的文件上，这类兼容本集按「不承诺」处理（见 §9 的 `PT1` 行）。
- `PD1` 复核仍然成立：`PdfDocument.cs:855-869` 先以 `/Type /Pages` 判定 `isPages`，`kids` 为 null 时照样进 `kids.Count`，NRE 未被并行批次修掉（兄弟实现 `PdfFile.Pages.cs:554` 的守卫是另一处）。

### 11.2 Phase 2 落地记录（2026-11-03）

`PD1`、`PD2`、`PD6` 的断言落在 `tests/Wisdom.Office.Tests/RemediationPhaseTwoTests.cs`，逐条先红后再绿，分别由 `3568a0c`、`229cd2c`、`e50edb6` 落地。`WE5` 同一文件同一流程，修复与红测分两次提交（`ae7a687`、`4639930`）。`WE1` 属挂起类，按 §2 不进主运行器。

主运行器里的断言文本（修复前的失败形态见 §3 对应条目）：

```text
a Pages node without /Kids should yield zero pages, got N
a /Kids reference to a missing object should yield zero pages, got N
a short object-stream header dropped the catalog
a short object-stream header dropped the catalog in the object-level reader
a font with maxp.numGlyphs erased failed the save path with <Type>: <Message>
the degraded save produced no file
null metadata failed a writer with NullReferenceException: Object reference not set to an instance of an object.
```

`WE5` 的失败原文就是上面最后一行：四条 `Require(…Length > 0, …)` 一条都没触发，`OdtWriter.Write` 抛出的 `NullReferenceException` 先被测试的 `try` 收进 `failure`，再由末行断言拼出消息。四个写出器只有 ODT 这一支没防：RTF 的 `AppendInfo` 首行 `if (metadata == null) return;`，HTML 与 DOCX 各自 `?? new DocumentMetadata()`，所以修复也按这个已有口径落在 `BuildMeta` 入口一行，不新增公共承诺，§9 表中无需为 `WE5` 加行。

`WE1` 用同一条命令取两轮证据，只改 `MaxCoveredSlots` 一个常量：

```text
# 预算置为 int.MaxValue（修复前形态）：120 秒未返回，被 SIGKILL 收走
exit=137

# 预算 200000（d3c1080）：
[crash-probe] WE1 loaded a 4096x4096 span cell in 35 ms, rows=1, cells=1, colSpan=4096
exit=0
```

双目标复核：net8.0 与 net461 均跑到末行 `All tests finished.`，无未捕获异常早退。

遗留与更正：

- `PD1` 选了「零页」降级，与兄弟实现 `PdfFile.Pages.cs:554`（当作叶子页）不一致，登记为后续统一项；§11.1 末条那条「`PD1` 仍然成立」的复核记录已被 `3568a0c` 取代。
- `PD7` 与 `E7` 暂缓：两条都动 `DocxWriter.cs`，该文件正被并行批次改写，等其入库后再取红证，避免互相覆盖。
- ODT 跨度属性前缀：本库读写两端都用 `table:` 前缀读 `number-rows-spanned` / `number-columns-spanned`；真实 ODF 是否该带 `text:` 前缀尚未取证（OASIS 页面取不到），保留为待查项，本集不改代码。
- `WE1` 的预算按「单表」计数，多张表各自享有一份；极端构造下总量仍可放大，与 `MaxNodes` 的联动留作后续项。
- 文档同步已完成：`README.md` 的「不可信输入」列举了 ODT 表格覆盖登记，`src/Wisdom.Word/README.md` 写明降级代价是超大跨度之后的网格对齐不再精确。
- `WE1` 的套件内包装已撤除（`4639930`）：`7da0ad6` 入库时连带进了一段在跑测试时 fork 子进程再断言退出码的 `RunOdtSpanProductBudgetTest`，它同时犯了两处错——用 `Environment.ProcessPath`（net461 无此 API，`CS0117` 直接把 net461 编译打断）和把 `ReadToEnd()` 排在 `WaitForExit()` 之前（子进程一假死，主运行器就跟着一起不动，正是 §2 要避免的形态）。挂起类证据只由 `--crash-probe WE1` 加调用方超时产出。
- `WE5` 条目里引用的 `DocxWriter.cs:1870`、`:1900` 两处判空，在并行批次改写后已漂到 `:2281`、`:2311`；`RtfWriter.cs:232` 与 `HtmlWriter.cs:56` 未动。行号只是指路，结论仍成立。
