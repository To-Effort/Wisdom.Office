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
  2. 写端不要伪造模型未表达的意图（四档均已落地：表头一档 `04ff01b`、形状填充与描边一档 `6dce2b6`、项目符号与 `pPr` 一档 `adb0f30`、文本框折行一档 `5f15d03`）：未显式声明表头就不写 `firstRow`/`bandRow` 与逐格填充（让 `tableStyleId` 生效）；无标记时不写 `buNone`，一段连对齐缩进行距都没设时 `a:pPr` 整个省略；`p:style` 未解析出的填充描边要区分「未设置」与「显式无填充」，不能一律 `a:noFill`；`a:bodyPr@wrap` 只有 `none` 是一句真话，square 与缺失都算「没说过」。
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

### PT11｜双线与点线被塞进 `prstDash`，产出枚举外的线型　A

- 位置：写 `src/Wisdom.Ppt/Pptx/PptxWriter.cs:807-816`（`DashStyle` 给 `Double` 返回 `"dbl"`、给 `Dotted` 返回 `"sysDash"`），调用点 `:799-800`（形状描边 `a:ln`）与 `:1406-1407`（表格边线 `a:lnL`/`lnR`/`lnT`/`lnB`）；读 `PptxReader.cs:1193-1200`（`MapDashStyle` 的取值表里没有 `dbl`）。
- 后果：`ST_PresetLineDashVal` 根本没有 `dbl` 这一档——双线不是划线图案，是**复合线**，规范写法是 `a:ln@cmpd="dbl"`（同一个写出器给表格边线就写了 `cmpd="sng"`，可见属性位置本来就在手边）。文件里落下一个枚举外的值，严格校验按坏属性处理；更要紧的是读端认不出来，`MapDashStyle` 一律回落 `Single`，双线边框存一次盘静默降成单线。点线写成 `sysDash`（长划图案），画出来是虚线而不是点。
- 目标行为：`Double` 写 `cmpd="dbl"`（形状描边与表格边线两条路径同改），`Dotted` 写 `sysDot`；读端优先看 `cmpd`，同时继续认历史文件里的 `prstDash val="dbl"`，让 1.0.2 之前存出的双线仍然读得回来。实线与虚线的字节不动。
- 红测：一份双线描边 + 点线描边 + 双线表格边线的稿子，断言文件里不出现 `prstDash val="dbl"`、双线形状的 `a:ln@cmpd="dbl"` 在位、点线落在圆点档、表格某条 `a:lnX@cmpd="dbl"` 在位，`PptxReader.Load` 回读 `Style` 仍是 `Double` 与 `Dotted`。

### PT12｜`Table.Borders` 六个方向只有一个真正生效　A

- 位置：`src/Wisdom.Ppt/Pptx/PptxWriter.cs:1236`（`BorderLine inside = model.Borders != null ? model.Borders.InsideHorizontal : BorderLine.Thin();`）→ `CreateTcProperties` 把四条边一并初始化成这一个对象（`:1376`），只有 `cell.Borders` 能逐边覆盖。
- 后果：`TableBorders` 的 `Top`/`Bottom`/`Left`/`Right`/`InsideVertical` 五份声明写不出门。外圈 3 磅、内部 1 磅是演示稿表格最常见的写法，存一次盘四条边一样粗；读端从不解析 `a:lnL`/`lnR`/`lnT`/`lnB`（`PptxReader.cs:957-960` 自己写明），于是这份丢失在往返里连痕迹都不留。
- 目标行为：每条边按单元格落在网格的哪一格取声明——贴外圈的走六个方向各自的边，内部的走两条分隔线，跨行跨列的合并块按自身首末行列判定，`TableCell.Borders` 仍逐边优先。`TableBorders` 的默认六向同为 0.5 磅黑色细实线，没动过 `Borders` 的稿子字节不变。
- 红测：2 行 3 列表格，六个方向各给不同颜色与线宽，逐格断言 `a:lnT`/`lnB`/`lnL`/`lnR` 的 `srgbClr@val` 与 `w` 落在该落的那份声明上。

### PT13｜`BorderStyle.None` 的表格边线仍写 `a:solidFill`，无边框被反着执行　A

- 位置：`src/Wisdom.Ppt/Pptx/PptxWriter.cs:1409-1420`（`CreateCellBorder` 取线色时无条件 `ln.Add(solidFill)`，从不看 `line.Style`）。两处对照就在同仓：`PptxWriter.cs:789-802` 给形状描边的 `None` 写 `a:noFill`；`src/Wisdom.Word/Docx/DocxWriter.cs:2040-2044` 给 `None` 写 `w:val="nil"`。
- 后果：`Table.Borders = TableBorders.None()` 是统一模型里「去边框」的唯一写法（`WordLabels.cs:441`、`WorkbookToDocumentConverter.cs:355` 都在用它），存进 `.pptx` 得到的却是四条 0.5 磅黑色实线——声明被反着执行。读端又不解析 `a:lnL`/`lnR`/`lnT`/`lnB`，往返一路看不出问题。
- 目标行为：`BorderStyle.None` 写 `<a:lnX w cmpd algn><a:noFill/></a:lnX>`，其余线型的字节不变。
- 红测：一张 `Borders = TableBorders.None()`、只把 `Top` 设回细线的表，逐格断言首行 `lnT` 是黑色实线、`lnB`/`lnL`/`lnR` 都是 `a:noFill`。

### PT14｜表格逐格边线只出不进，载进来的边框全丢、自己写的也活不过第二次保存　A

- 位置：`src/Wisdom.Ppt/Pptx/PptxReader.cs`（`ReadTableCell` 取过 `tcPr` 的底色与 `anchor`，却从不看 `a:lnL`/`lnR`/`lnT`/`lnB`，原注释自写「不读入」）。写出侧 `PptxWriter.cs:1387-1397` 每个单元格都写四条边，`TableCellBorders` 四个字段在模型里本来就有（`DocumentModel.cs:671-681`）。同库 Word 侧早有对照：`DocxPartReader.cs:856-870` 的 `ReadCellBorders` 逐边解析、四条全缺才返回 null。
- 后果：别人的表格载进来，外圈粗线与内部淡线一律按「没说过」处理，回写拿到 `TableBorders` 默认的 0.5 磅黑色细实线——整张表的边框排版塌回一种线；本库自己写出的逐格 `TableCell.Borders` 同样在第二次保存时被抹平。读侧不解析，往返一路连痕迹都不留，`PT11`/`PT12` 那两批只能把这条写进「已知限制」。
- 目标行为：四条边逐格读回 `TableCell.Borders`；`w` 转磅、颜色取 `a:solidFill`、双线看 `cmpd="dbl"`、`a:noFill` 认成 `BorderStyle.None`（「这条边不画」）而不是「没说过」；四条全未声明时 `Borders` 保持 null，让表级六向做主。第二次保存逐边字节不变。
- 红测：两张表（六向各异的一份 + `TableBorders.None()` 只留 `Top` 的一份），`PptxReader.Load` 后逐格断言四边颜色/线型回到声明值，再把载入的稿子写出一次，逐格四条边的 `w` 与颜色签名必须与第一次一致。

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

### WE12｜ODT 页码域按 `IndexOf("PAGE")` 判定，`NUMPAGES` 被写成当前页码　A

- 位置：`src/Wisdom.Word/Formats/OdtWriter.cs:545`（改前是 `field.Instruction.IndexOf("PAGE", StringComparison.OrdinalIgnoreCase) >= 0`）；读端 `src/Wisdom.Word/Formats/OdtReader.cs:559` 有 `text:page-number` 分支，没有 `text:page-count` 分支。
- 后果：`NUMPAGES` 这个词里正含 `PAGE` 四个字母，所以「第 X 页 共 N 页」（`src/Wisdom.Word/README.md` 页脚示例的写法）写成 ODT 是两个 `<text:page-number text:select-page="current"/>`，总页数被画成当前页码，LibreOffice 打开即见错值；载入再存一次也修不回来，因为指令在这一路已经被丢掉。DOCX（`src/Wisdom.Word/Docx/DocxWriter.cs:1534`）、RTF（`src/Wisdom.Word/Formats/RtfWriter.cs:465`）、HTML（`src/Wisdom.Word/Formats/HtmlWriter.cs:309`）都把 `FieldRun.Instruction` 原样带出去，只有 ODT 塌。
- 目标行为：按整词判定——`NUMPAGES`/`PAGES` 写 `text:page-count`，整词 `PAGE` 写 `text:page-number`；读端补 `text:page-count` → `NUMPAGES`。`FieldRun.Instruction` 本来就存着指令，模型无需改动。
- 红测：一段里放 `PAGE` 与 `NUMPAGES` 两个域写 ODT，断言 `content.xml` 出现 `text:page-count` 且 `text:page-number` 只出现一次；回读断言指令含 PAGE 与 NUMPAGES；DOCX 同稿作对照。

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
| `PD8` | 交叉表由全文扫描救回的文档，增量保存从「追加一段」改为全量重写；原字节不再逐字保留 | `docs/Wisdom.Pdf.Tutorial.md` 增量保存段、`src/Wisdom.Pdf/README.md` 已知限制、`docs/Wisdom.Office.功能导读.md` 2.4 的「只追加改动对象」承诺（三处已同步） |
| `PD17` | 交叉表条目指到别人的对象头时，从「静默把该偏移上的对象当成目标号交回」改为「按条修正后重读，修不回即当作缺失」 | `src/Wisdom.Pdf/README.md` 已知限制（已同步）；`docs/Wisdom.Pdf.Tutorial.md` 的容错说明已随 `PD9` 一并补（`PD9` 取上抛路由，未建告警通道） |
| `PD9`、`WE7`、`WE9` | `PD9` 落地上抛路由：解码超限与内存失败不再被页提取吞掉，`ExtractText()` / `GetContentBytes()` 新增可抛 `DecompressionLimitException`（未建告警集合，统一通道留给 `WE7`/`WE9` 那一批） | `docs/Wisdom.Pdf.Tutorial.md` §10、`src/Wisdom.Pdf/README.md`、`articles/Wisdom.Pdf.md` 已补「资源上限不走静默降级」；`docs/Wisdom.Ppt.Tutorial.md` 与 Word 侧待 `WE7`/`WE9` |
| `PT1`、`PT2` | 输出文件的 XML 形状改变（历史产物可能被新版本读不出） | `docs/Wisdom.Ppt.Tutorial.md` 需写明往返保真范围变化；若旧形状需保留读取兼容，一并声明 |
| `PT6` | 空占位符的省略条件从「占位符且无文本」收窄为「本库按版式补出、且到写出时仍没有任何可见效果」；从文件读回的空占位符与调用方自建的空框现在照常写出，往返形状数与层叠次序稳定 | `src/Wisdom.Ppt/README.md`「版式与占位符」补了省略规则与新增的来源标记 `SlideShape.IsGeneratedPlaceholder`（public 读、internal 写）；`docs/Wisdom.Ppt.Tutorial.md` 第 3 节「空框叠在表格上」的说法按新规则重写过 |
| `PT5` | 输出 `a:rPr` 的形状改变：字号、颜色、字体三项只在**显式设过**时才写，只加粗的 run 从 `sz="1100"` + `a:solidFill` + `a:latin`/`a:ea` 缩成只剩 `b="1"`；这类稿子二次写出（Load+Save）的字节随之变化，渲染改由版式与母版继承决定 | `src/Wisdom.Ppt/README.md`「文本、图形、图片与表格」写明这条规则与 `TextFormat` 新增的三个只读标记；`docs/Wisdom.Ppt.Tutorial.md` 常见问题补「只想改字号，保存后连字体一起被换了」 |
| `PT3` 结构档 | `a:tblPr` 从每张表恒写 `firstRow="1" bandRow="1"` 改为只在声明过表头（某行 `IsHeader` 或给过 `HeaderColor`）时写 `firstRow`，`bandRow` 因模型没有带状行字段而不再输出；未设底色的单元格从写 `a:noFill` 改为不写填充，补齐用的空格在数据行也不再写白色，`tableStyleId` 本该生效的整表外观因此重新可见。新写出的无人声明表头的表格字节随之变化 | `src/Wisdom.Ppt/README.md`「文本、图形、图片与表格」新增表格结构档一段；`docs/Wisdom.Ppt.Tutorial.md` 第 3 节注意事项补「表头要声明才有表头」一条、常见问题补「表格首行没有表头的样子」；读端新增把 `a:tblPr@firstRow` 读回 `TableRow.IsHeader`，旧形状（恒写 firstRow 的历史文件）照常读取，其首行按表头认，二次写出不掉样式 |
| `PT3` 结构档·段落符号 | 每段末尾的 `a:buNone` 不再无条件写出：无 `ListMarker` 时不写任何 bullet 子元素，全段无属性时连 `a:pPr` 一并省略。自家母版与版式不定义符号，故新建产物字节不变；差别在载入第三方稿再存盘——过去那份逐段压制等于关掉母版里的项目符号，现在符号继承照常生效。代价是「显式关掉符号」与「未设置」在模型里仍同一个值，来稿中的 `buNone` 存盘后不保留 | `src/Wisdom.Ppt/README.md`「文本、图形、图片与表格」新增段落属性一段（含这条边界）；`docs/Wisdom.Ppt.Tutorial.md` 常见问题补「用本库打开别人的演示稿再存盘，列表的项目符号全消失了」 |
| `PT3` 结构档·文本框折行 | `TextBoxShape` 新增 public 可空字段 `WrapText`（纯增项，不破坏既有调用）。写端仍按「未设置 → `wrap="square"`」落，自家产物字节不变；差别在载入第三方稿：过去 `wrap="none"` 被读端完全忽略、写端一律钉成 square，那条不许折行的单行标题存一次盘就开始在框内折行。现在 `false` 写 `none`、`none` 读回 `false`；`square`/`emergency`/属性缺失一律读成 null（文件里没有任何一句「要折行」值得当成声明，`emergency` 本库模型也表达不了） | `src/Wisdom.Ppt/README.md`「文本、图形、图片与表格」段落属性一段后补折行一段、「已知限制」补 `emergency` 与 true 不往返的边界；`docs/Wisdom.Ppt.Tutorial.md` 常见问题补「载入别人的稿子后那一行标题开始折行」 |
| `PT11` | 双线不再写进 `a:prstDash`（`ST_PresetLineDashVal` 没有 `dbl` 这一档），改写成 `a:ln@cmpd="dbl"`；点线从 `sysDash`（长划）改成 `sysDot`。用过这两种线型的稿件在双线/点线两处字节变化，实线与虚线不动；读端补 `cmpd` 识别，并继续认历史文件里的 `prstDash val="dbl"`，旧稿读回仍是双线 | `src/Wisdom.Ppt/README.md`「文本、图形、图片与表格」新增线型一段、「已知限制」补「表格边线只出不进」；`docs/Wisdom.Ppt.Tutorial.md` 常见问题补「双线边框存一次盘变成单线」 |
| `PT12` | 表格四条边不再一律取 `InsideHorizontal`：贴外圈的那条边取 `Top`/`Bottom`/`Left`/`Right`，内部取两条分隔线，合并块按自身首末行列判定，`TableCell.Borders` 逐边优先。`Table()` 的默认六向同为细黑线，没动过 `Borders` 的稿子一个字节不变；只有声明过非默认方向的稿子随之变化 | `src/Wisdom.Ppt/README.md`「文本、图形、图片与表格」新增表格边框一段，「已知限制」的「只出不进」改写为按六向落位；`docs/Wisdom.Ppt.Tutorial.md` 常见问题补「表格外圈加粗，保存后四条边一样粗」 |
| `PT13` | 表格的「无边框」不再补一条实线：`BorderStyle.None` 的那条边写 `a:noFill`。用过 `TableBorders.None()` 的稿子字节随之变化（变化本身就是修复），其余线型一个字节不动；形状描边一直是这个口径，两条路径由此一致 | `src/Wisdom.Ppt/README.md`「文本、图形、图片与表格」的表格边框一段补这句；`docs/Wisdom.Ppt.Tutorial.md` 常见问题补「`Table.Borders` 设成 `TableBorders.None()`，保存后边框还在」 |
| `PT14` | 表格逐格四条边线从「只出不进」改为双向：读端解析 `a:lnL`/`lnR`/`lnT`/`lnB` 逐格落进 `TableCell.Borders`，`a:noFill` 认成 `BorderStyle.None`，四条全缺仍为 null。载入别人的表格再存盘不再把边框塌回默认细线，这一类稿子的字节随之变化（变化本身就是修复）；本库新建稿第二次保存逐边字节不变 | `src/Wisdom.Ppt/README.md` 表格边框一段改写为双向，「已知限制」的「只出不进」条目替换为「逐格往返、表级六向不回」；`docs/Wisdom.Ppt.Tutorial.md` 常见问题补「别人的表格打开再存一次，边框全没了」，并把 `PT11`/`PT12` 两条里「只出不进」的边界句改掉 |
| `WE12` | ODT 页码域从 `IndexOf("PAGE")` 改为整词判定：`NUMPAGES`/`PAGES` 写 `text:page-count`、整词 `PAGE` 写 `text:page-number`，读端新增 `text:page-count` → `NUMPAGES`（此前没有这一分支）。同时用到两个域的 ODT 产物在总页数那一处字节变化（变化本身就是修复），只用 `PAGE` 的稿子一个字节不动 | `src/Wisdom.Word/README.md`「在格式之间保留什么」的页码一行改为「页码与总页数域」并新增一段说明 ODT 只认这两个指令、其余指令写的是显示文本；`docs/Wisdom.Word.Tutorial.md` 常见问题补「『第 X 页 共 N 页』存成 ODT 变成两个页码」 |

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

### 11.3 Phase 3 落地记录（2026-11-03）

断言落在 `tests/Wisdom.Office.Tests/RemediationPhaseThreeTests.cs`，接在 Phase 2 之后。三条修复前的失败原文：

```text
Test failed: 200000 glyphs on one line took 7622 ms to extract
Test failed: doubling 20000 lines to 40000 multiplied extraction from 2860 ms to 11376 ms
Test failed: doubling a dictionary from 20000 to 40000 keys multiplied parsing from 2828 ms to 14489 ms
```

修复后同一批断言的实测：`PD3` 116 ms（net8.0）与 153 ms（net461），阈值 1200 ms 留十倍余量；`PD4` 103/203 ms 与 81/183 ms，倍率上限「三倍加 200 ms」；`PD21` 26/39 ms 与 27/39 ms。形状同时守住：两万行仍归并成两万行、四万行仍四万行，四万键字典照旧读出 1 页，`/Kids` 与 `/Count` 没被索引读丢。

提交逐条配对：`57a4c35`+`ea8f115`（`PD3`）、`f2826ca`+`9c39e68`（`PD4`）、`ce05890`+`fedb7f8`（`PD21`）。

两处与规范目标写法的差异，按落地事实登记：

- `PD3` 的目标行为还要求「给单页字元与片段总数设上限，超限按 `PD9` 的告警通道上报而不是静默截断」。上限这一半依赖尚不存在的告警通道，按 §10「`PD5` 与 `PD22`、`PD9` 必须同批设计」并入 `PD9` 批次；本相只消掉二次方本身（`_last.Text += text` 改缓冲累积，取用 `Fragments` 前由提取器 `Flush` 一次成型）。同型的 Word 侧两处（`OdtReader.cs:565`、`RtfReader.cs:954`）累加在公共模型属性 `TextRun.Text` 上，改成缓冲要先决定「公共可写属性由 StringBuilder 支撑」的语义（外部写 `Text` 就得作废缓冲），不是同一次改动的规模，登记为独立条目。
- `PD4` 的目标行为写的是按基线容差投桶。实际落的是单调游标：片段本就按 Y 升序，一条行的 `Bottom` 一旦越过当前片段，与它和其后任何片段都不可能满足「至少重叠半格」，把它们划出窗口不改变「第一个命中」的结果，归并形状完全守恒，代价是实现更简单。已知残留：若同一竖直窗口内堆满角度互不相同的片段，窗口仍大、扫描仍是窗口内线性，投桶写法留作后续加强。

`PD5` 与 `PD22` 本相未动，理由同上（预算与告警是一件事的两半）。

### 11.4 Phase 4 落地记录（2026-11-03 起）

- `PD16`（`c6d755b` + `cede32d`）：把 `BT /F1 12 Tf 20 700 Td (Hi) Tj ET` 用 Flate 加 `/Predictor 12 /Columns <正文字节数>` 打包，正文恰好等于一整行，滤波字节无从读起。修复前的失败原文：

  ```text
  Test failed: a predictor body of exactly one row lost the content stream, got ''
  ```

  断言用的 fixture 需要二进制流正文，而 `AssembleClassicPdf` 全程走 `Encoding.ASCII`，会把 zlib 字节打成 `?`，因此另写一个按字节计偏移的 `AssembleBinaryPdf`。修复后双目标均跑到末行 `All tests finished.`，无新增公共承诺：预测器只在「正文不足滤波一整行」时改为原样交回，其余分支字节不变。
- `PD10`（`cfa7146` + `be31658`）：目标行为「与编号一并保存 `_catalogGeneration`，`/Root` 用真实世代」照原样落地。fixture 用新的 `AssembleGenerationsPdf` 把目录对象写成 `1 2 obj`、交叉表行 `0000000009 00002 n`、trailer `/Root 1 2 R`，保存后重新解析输出。修复前的失败原文：

  ```text
  Unhandled exception. System.InvalidOperationException: Test failed: the trailer /Root is '1 0 R' while the object header is '1 2 obj'
  ```

  改动点比条目多一处：`Reading/PdfDocument.cs:556` 的读取端救援路径同样把世代硬写成 0，而同名救援路径在 `PdfFile.cs` 里一直用的是扫描到的真实世代，两条现在口径一致。读取端这一处没有单独断言——它写出的引用只用于定位目录，`Wisdom.Pdf.PdfDocument` 不暴露 trailer，也没有按 (号, 世代) 严格查对象的调用点，因此改成真实世代不改变任何可观测行为，只是消除同一个硬编码；写出端五处引用构造（`PdfFile.cs:159`、`:577` 与 `PdfFile.Writer.cs:394`、`:444`、`:516`）加上世代回收点 `PdfFile.cs:582` 都在断言覆盖的保存路径上。无新增公共承诺：新文档的目录仍由 `AddObject` 分配、世代为 0，只有源文档目录本身住在非 0 世代时输出才变。
- 验证方式更正一条：并行批次正把 `DocxWriter.cs` 改到语法不通（`CS1031` 等），共享工作副本在此期间无法编译，因此 `Phase 3`/`Phase 4` 的双目标运行是在一棵与工作副本隔离的临时构建树里完成的——该树取当前受版本控制的文件，仅 `DocxWriter.cs` 退回入库版本、测试入口退回不含并行挂接行的版本。**该副本已恢复可编译，`PD3`/`PD4`/`PD21`/`PD16`/`PD10` 已在共享工作副本里重跑，`net8.0` 与 `net461` 均到末行 `All tests finished.`，隔离树兜底作废。**
- `PD17`（`c174e85` + `490dd31`）：目标行为「比对头部号，不一致即判条目失效并走重建」落地为**按条修正**——`ParseIndirectObjectAt(offset, expectedNumber)` 不符即抛 `PdfObjectHeaderMismatchException`，两个 `LoadObject`（写出端 `PdfFile` 与读取端 `PdfDocument`）catch 到它之后跑一次全文扫描，只用扫到的真头替换**这一条**条目再读一次。不整表重建：`RebuildXref()` 会连 `_trailer` 一起换成文件末尾的 trailer，一条坏条目没有资格触发那种改动；扫描每个文档至多跑一次并留用结果，避免畸形文件按失效次数反复全文扫描。`startxref` 与 `/XRefStm` 两个调用点只知道偏移不知道号，沿用不校验的重载。修复前的失败原文：

  ```text
  Unhandled exception. System.InvalidOperationException: Test failed: the xref row for object 5 points at object 3's header and the reader handed that object back: PdfDictionary /Type=PdfName /Parent=PdfReference /MediaBox=PdfArray /Contents=PdfReference /Resources=PdfDictionary
  ```

  fixture 把 5 号条目改指 3 号对象的头（两条落同一偏移），断言 `GetObject(5)` 不再吐出页节点，且读取端 `ExtractText()` 仍能拿到页里的 `Hi`——即条目被修好后内容回来了，不是静默变成空白页。`InvalidDataException` 是密封类，异常基类只能用 `Exception`，该类型 internal，不越出程序集。
- 验证方式再记一条：共享工作副本此刻能编译，但并行批次正在 `ExcelEnhancementTests.cs:197` 加打印标题断言（`Report!$1:$2`），该断言在其功能完成前必红，`--no-build` 的共享副本跑到那里就停。因此 `PD17` 的双目标末行是在 `D:/tmp/hb/verify2` 里取得的——该树取当前受版本控制文件的工作副本内容，只有测试入口 `Program.cs` 与 `ExcelEnhancementTests.cs` 退回入库版本（入库版不含那条挂接），`net8.0` 与 `net461` 均到 `All tests finished.`；共享副本里 `phase-four corrupt output and silent loss ok` 已经打出来，即本条目自身在共享副本也是绿的。**并行批次已收尾，`PD8` 之后共享工作副本整跑过一遍（见下），这条兜底作废。**
- `PD8`（`587ae3a` + `22f2fbb`）：目标行为「恢复路径要么修正 `_lastXrefOffset`，要么置 `_rewriteRequired`」取后半句——`RebuildXref()` 末尾补 `_rewriteRequired = true`，`SaveIncrementalToBytes()` 已有的分流（`_isNew || _data == null || _pendingEncryption != null || _rewriteRequired` 落到 `WriteFile(false)`）于是自动兜住这条路径，不需要新增判断点；不选「修正偏移」那一半，因为扫描出来的条目本来就散落在文件各处、不存在一个可作为 `/Prev` 目标的基段。修复前的失败原文：

  ```text
  Unhandled exception. System.InvalidOperationException: Test failed: the incremental update on a scan-recovered document dropped the untouched objects: second page text reads ''
  ```

  断言前先立两条控制：扫描救回的文档本身必须 `PageCount == 2`、原文里读得到 `Second`，否则红的是 fixture 而不是增量段。第一版只有 4 个对象，`RemovePage(0)` 之后余下对象全落在 `_changes` 里，追加段自给自足，测试假绿；补成 7 个对象（两条正文流加一个共用字体）后，未改动的 6、7 号必须由重开的阅读器按交叉表找回，才真打出上面这条红。`_rewriteRequired` 的字段注释同时改写：它原先只讲加密两种来源，现在多一条「交叉表来自扫描兜底」。
- `PD8`、`PD17` 的文档义务按 §9 新增两行落到位：`docs/Wisdom.Pdf.Tutorial.md` 在 `SaveIncremental` 代码块后补「只追加有前提」一段，`src/Wisdom.Pdf/README.md` 已知限制补两条（扫描救回后增量保存退化为全量重写、条目编号错位的作废与按条修正），`docs/Wisdom.Office.功能导读.md` 2.4 节那句「`SaveIncremental` 只追加改动对象（原字节保留）」就地加了同一例外。三处都只写可观测行为，不提标志位与内部方法名，也未出现本地绝对路径或发布步骤。
- 验证：本条文档改动前后各跑了一次共享工作副本整跑，输出分别落在 `D:/tmp/hb/run_net8.log` 与 `D:/tmp/hb/run_net461.log`，`net8.0` 与 `net461` 均打出 `[remediation] phase-four corrupt output and silent loss ok` 并到末行 `All tests finished.`；构建输出用 `.hb/Debug/<tfm>` 隔离目录，避开并行会话对 `bin/` 里 exe 的文件锁。
- `PD11`（`c237cb2` + `0b460b1`）：目标行为「`/BBox` 用表单空间 `[0 0 displayWidth displayHeight]`」照原样落地，改的就是 `PdfFile.Watermark.cs:299` 这一行——同一方法里 `displayWidth`/`displayHeight` 早已算好（`:273-274`），只是从未被 `/BBox` 用过。红测按条目说的两个用例各加一条居中以外的横排：`Free` 摆放 `AutoFit=false`，一个 90° 旋转的 612×792 页把锚点摆到 x≈735，一个 `[0 200 612 992]` 的原点抬高页把锚点摆到 y≈170，断言两件事——`/BBox` 的宽高必须等于阅读方向的显示宽高，正文第一条 `cm` 的平移分量必须落在 `/BBox` 内（即 `Do` 的裁剪框得住）。第一个用例的失败原文：

  ```text
  Unhandled exception. System.InvalidOperationException: Test failed: a page rotated to 90: the form /BBox is [0 0 612 792] but the watermark draws in a 792 x 612 reading direction
  ```

  第二个用例在此之前不会单独报红（宽高恰好相同，只有原点差 200），所以断言顺序把「锚点落在框外」也立成独立一条，`[0 200 612 992]` 那页的 `Do` 裁掉下半部才真的被测到。
- `PD11` 的产出字节核对：不旋转且 `MediaBox` 原点为 0 时 `displayWidth=right-left`、`displayHeight=top-bottom`，新旧 `/BBox` 完全相同，绝大多数存量文档的输出一个字节都不变；只有旋转页与抬高原点的页会变化，且变化后才是规范里表单空间该有的样子。§9 不需要新行——水印 API 与文档承诺都没动，教程的旋转页段落描述的正是「阅读方向」这一套坐标。
- `PD14`（`25d5ee8` + `047893a`）：目标行为「复用探测空闲资源名的逻辑」落地为把 `WatermarkResourceName` 的探测抽成 `PdfFile.Pages.cs` 的 `FreeResourceName(index, category, baseName)`（internal），水印与扁平化同用一个：`FlattenField` 先探测再 `SetPageResource`，`WatermarkResourceName` 只剩一行委托。抽出来而不是各写一份，是因为旧探测顺手兼顾了页面**继承**来的资源字典——第三方 PDF 的 `/Fm0`、`/Fm1` 常常正住在那里，只查页面自己那份会漏掉冲突；`FlattenField` 原先连这一步都没有。红测按条目说的两轮形态做：第一轮放一个已挂进 `/Annots` 的文本域、再放一个故意不挂的（本轮画不出、留在表单里），第二轮把后者的控件补挂进 `/Annots` 再扁平，序号于是又从 1 数起。断言取「页面内容里每个 `Do` 真正解析到的不同对象号必须是 2 个」。修复前的失败原文：

  ```text
  Unhandled exception. System.InvalidOperationException: Test failed: two flatten rounds left only 1 appearance(s) reachable from the page: the second round reused the first round's resource name
  ```

  补挂 `/Annots` 之后要 `SaveToBytes` 再 `Load` 一次才能进第二轮：表单模型是按 `/Annots` 建的，同一份内存里的旧模型不知道控件刚被挂上去，`ResolveWidgetPage` 仍返回 `-1`，第一轮的红测因此先撞在「第二轮没画出任何域」这条控制断言上（`the deferred field must flatten once it is placed on the page, got 0`），加上重开一份才拿到上面那条真红。产出字节只在名字真的撞车时才变（空闲名字照旧是 `/Fm1`），表单字段数不变、页面资源条目只会多不会少。
- `PD15`（`4982a31` + `0b6f061`）：`targets` 旁边并记一份 `targetNumbers`，字段本体与每个控件注释各按自己的号 `MarkModified`；`/AP` 先判引用再 `Resolve`，命中间接对象时编辑那一个字典并登记它的号，不再用全新的内联字典整块顶掉。红测的第一版用 `AddObjectReference` 现场搭控件，结果修复前就是绿的——新建对象本来就在 `_changes` 里，增量保存必然把它写出去，少掉的那次标记根本显不出来。改成 `AssembleClassicPdf` 按字节拼样例（控件 6 号、字段 7 号、旧外观流 10 号都是文件里原有的间接对象）才拿到真红，修复前失败原文：

  ```text
  Unhandled exception. System.InvalidOperationException: Test failed: the incremental save left the widget's old appearance in the file: /AP /N should be 18 0 R but the saved widget carries 10 0 R
  ```

  第二条用例（`/AP` 住在间接对象里）单独验 `Resolve`：修复前它撞的是「`/R` 一起丢了」，与第一条互不遮盖。产出字节只在两种情形变化：控件与字段分住两个对象且生成过外观（增量段多写控件的新字节，这正是「显示旧文字」的成因）、`/AP` 是间接对象（改为原地编辑，`/R`、`/D` 与状态映射保留）。§9 不需要新行——`SaveIncremental` 与 `GenerateFormAppearances` 的公开承诺本来就是「改动全部落盘」，这次只是让它兑现。
- `PD18`（`64fb595` + `e688ad2`）：目标行为「按编码名分派」落地为构造期读 `dictionary.Get("Encoding")`，`*-UCS2-*`、`*-UTF16-*` 记进 `_codesAreUnicode`，`AddGlyph` 的两字节分支在缺 `ToUnicode` 时按大端码元直接产出字符；`Decode` 给 `AddGlyph` 多带一个 `paired`，落单的残字节凑不出码元，宁可留空也不伪造。Identity 一族维持「取不到文字」——条目正文自己就写了「改用 `/Identity-H` 时确实无法还原，属既定能力边界」，把 CID 当字符编号吐出一堆假字比空串更坏，`UniGB-UTF8-H` 这类变长编码同样不在本条范围。死字段 `_identityLike` 由 `_codesAreUnicode` 接替，注释里那句「两字节 CID 字体若无 ToUnicode 则只能得到空文本」一并改成按编码名分派。夹具用 `AssembleClassicPdf`，显示串写成十六进制 `<4E2D65876D4B8BD5>`——纯 ASCII，既不依赖源文件编码，也免去 CJK 在 PDF 字面串里的转义。修复前的失败原文：

  ```text
  Unhandled exception. System.InvalidOperationException: Test failed: a /UniGB-UCS2-H page without /ToUnicode extracted '' instead of 中文测试
  ```

  两条控制断言（`Identity-H`、`UniGB-UTF8-H` 不许出现「中文测试」）修复前就是绿的，它们钉的是这次改动可能越过的界：UTF-16 码元逐个产出、由调用方拼接，因此代理对天然还原成增补平面字符，不需要在读字元时前瞻。产出字节不受影响——本条在读取端，只改 `ExtractText` 与依赖它的渲染取字结果；§9 不需要新行，README 与教程的清单从来没承诺「两字节字体无 ToUnicode 必定抽不出字」，这次是少抽变多抽。
- `PD9` + `PD5` + `PD22`（`b0c7a0a` + `c5b1d4a`）：三条按 §10 的「必须同批设计」合批落地。
  - **路由取上抛，不建告警集合**。条目给的分支是「上抛或记入告警集合」。选上抛有两条理由：其一，`src/` 全仓没有 `InternalsVisibleTo`，告警集合做成 internal 就测不到、做成 public 就是新增一整个公共契约（要按 §9 同步两个包的教程与 `articles/`），而 §9 已经把 `PD9`、`WE7`、`WE9` 点名为「共享一个文档级告警与降级通道的新抽象」——统一通道不该由 PDF 单方面先定形状，留给 Word 那一批共同设计；其二，本条丢的是一件**确定的事实**（这一页解码超限），异常本身就能表达，告警集合该留给「部分丢失但整体可读」那一类。统一通道因此登记为独立后续条目。
  - `PD9`：`PdfPageExtractor.Extract` 原来一把 `catch (Exception)` 全吞，超限因此伪装成「这页没字」，还被 `PdfPage.EnsureExtracted` 永久缓存成空文本；现在 `DecompressionLimitException` / `OutOfMemoryException` 先原样上抛，其余照旧按语法损坏静默降级。控制断言留在红测里：把一小段压缩正文的中间字节改坏，仍不许有任何异常逃出 `ExtractText()`——静默那一半不能被这次改动一起丢掉。`PdfContentInterpreter.Run` 的算符级 catch 与 `ExecuteXObject` 的解码 catch 同口径。
  - `PD5`：`DecodeStream` 按 `PdfStream` 实例缓存解码正文。条目写「按对象号（或 `PdfStream` 实例）」，取实例这一半——解码调用点手里只有实例，按号缓存要把号穿到每一个解码点，而读取端一个号本来就只建一个实例。留存上限 64 MB，超了整表作废重解，不做 LRU：要挡的是「读完一本书把全书正文留在内存里」，逐条淘汰换来的复杂度换不到额外的正确性。解压放在锁外，两个线程同时未命中各解各的、内容相同。`GetContentBytes` 的 `/Contents` 数组出口逐段核对合计，共用单流上限。
  - `PD22`：新增 internal `PdfDecodeBudget`，给单页累计解码量按流实例记账，越界即抛；嵌套 Form 沿用同一个实例。**不新增公共上限字段**——扇出预算和单流预算量的是同一格「一页能占多少内存」，多一个字段就得向调用方多解释一次两者怎么配比。按实例记账保证同一个表单被 `Do` 两千次只算一份（这条控制断言单独立着，正常模板不许误伤），深度上限只挡得住自引用，一层铺开的扇出由累计预算补。
  - 断言形态偏离条目一处，记录原因：条目建议「计数型断言」数解码调用次数，但 `DecodeStream` 是 internal 且无 `InternalsVisibleTo`，测试数不到。改用计时：240 个对象、4.6 MB 的对象流，修复前 2 468 ms、修复后 84 ms（`net461` 127 ms），断言上界 700 ms。**首版这个夹具是假绿的**——200 页共用同一段填充数字，重复字节被 deflate 压得极小也 inflate 得极快，修复前只用 556 ms 就撞不到 2000 ms 的红线；改成每页独占一段不可压缩填充、页数升到 240、上界收到 700 ms 才拉开差距。计时夹具必须自己证明「每页数据互不相同」，否则测的是压缩率不是缓存。
  - 刻意不动的一处：`PdfTextFont.cs:88-104` 的 ToUnicode 流解码仍按字体粒度静默降级。单个字体映射失败只让这一支字体抽不出字，页面其余文字照常，正是「部分丢失但整体可读」那一类，要不要上报等统一告警通道一起定。
  - §9 文档同步到位：`docs/Wisdom.Pdf.Tutorial.md` 与 `src/Wisdom.Pdf/README.md` 的「结构损坏静默降级」段落各加一句「资源上限不走这条降级」，`articles/Wisdom.Pdf.md` 的读取上限一节同样补上。三处只写可观测行为（哪两个方法、抛哪一类异常、批量索引该怎么接），不提缓存字段与预算类名，未出现本地绝对路径或发布步骤。`docs/Wisdom.Office.功能导读.md` 那句「`DecompressionLimitException` 与 `InvalidDataException` 分开，好让调用方区分『文件坏了』和『文件太大了』」本来就以上抛为前提，这次是让它第一次真的成立，无需改动。
  - 验证：`net8.0` 与 `net461` 双目标各整跑一遍，均打出 `[remediation] decode-budget PD9 PD5 PD22 ok` 并到末行 `All tests finished.`。
- `WE2`（`c7426d6` + `7d7f644`）：目标行为「包装为 `InvalidDataException` 并保留内层异常」照原样落地，改的就是 `OdtReader.LoadPackage` 里 `Wml.LoadXml(contentBytes)` 这一行——同一个函数对 `styles.xml`、`meta.xml` 早就 `catch (XmlException)` 按支降级，对缺部件也抛 `InvalidDataException`，只有主部件这一行让解析器类型直接穿出契约。主部件没有降级余地：正文读不出来就没有可重建的对象，整个文档一无所有，只能算坏文件。
  - 断言把「只有主部件致命」这条界三面钉住：正文合法而 `styles.xml` 坏掉必须照常出文（否则红的是夹具不是本条）；缺 `content.xml` 本来就抛 `InvalidDataException`，用它确认这一族被当成「坏文件」而非「读不出来」；门面 `WordFormatConverter.Load(byte[])` 单独一条，因为条目点名的漏网路径正是按契约兜错的调用方。最后一句钉 `InnerException is XmlException`——换类型不许把行号列号一起换掉。
  - 夹具现用 `ZipArchive(ZipArchiveMode.Create)` 在内存里拼包（`mimetype` + `META-INF/manifest.xml` + `content.xml`），ODT 侧此前没有拼包工具。`manifest` 必须在：`Detect` 要同时见到 `content.xml` 与 `META-INF/manifest.xml` 才认 ODT，缺了它门面那一路测的就是「格式没认出来」，与 `XmlException` 无关。
  - §9 不需要新行：`docs/Wisdom.Office.功能导读.md` 那句「格式专属的读取器把损坏统一表达为 `InvalidDataException`」本来就承诺了这件事，这次是让它对 ODT 第一次真的成立；顺带给 `OdtReader` 三个公共 `Load` 重载补 `<exception cref="InvalidDataException">`，把承诺写回方法文档。
  - 验证：`net8.0` 与 `net461` 双目标均打出 `[remediation] word phase-four contracts ok` 并到末行 `All tests finished.`。
- `WE6`（`3f3a411` + `ea370ee`）：目标行为给了两句——「外层转成已声明的 `ArgumentException` 并说明超时」和「先在副本上完成全部替换再一次性提交」。两句都兑现了，但第二句不照字面做副本：替换是逐处落在 `TextRun.Text` 上的，深拷贝整篇文档等于内存翻倍，而且模型没有现成的 `Clone`。改成**计划/落地两趟**拿到同样的原子性而不加内存——正则是唯一可能中途失败的步骤，只跑在计划趟；落地趟只是给运行赋新文本，不再有能失败的解析。超时因此必然发生在「一个字都没改」的时刻。
  - 内层异常留在包装里：超时发生在哪一段只能靠 `RegexMatchTimeoutException` 自带的模式与输入定位。
  - 两趟改造要对齐的旧语义逐条核对过：`FirstOnly` 旧实现命中第一处后不论 `ReplaceRange` 成败都停止，新实现计划里只留一条范围、落地尝试一次，计数同样可能为 0；空运行清理的触发条件由「本段落的累计计数变过」变为「本段落至少落地一处」，并从「段落处理完就清」改为「全部落地完再按段落清一次」（同一段落可能横跨多个分组，用 touched 列表去重）。存量 `WordEnhancementTests` 的跨运行、全词、首处、页眉页脚与表格范围断言全绿，是这次语义对齐的主要证据。
  - §9 不需要新行：`ArgumentException` 本来就是方法文档声明的类型，这次是兑现承诺而非改变契约；两个公共 `Replace` 重载的 `<exception>` 文案补上了「匹配超时」这一触发条件。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`（超时那一趟实测各约 2 秒，即 `RegexTimeout` 的设定值）。
- `PT6`（`52f66bd` + `158b3f2`）：目标行为「只跳过由 `BuildPlaceholders` 自动补且未改动的占位符（加来源标记）」落地为 `SlideShape.IsGeneratedPlaceholder`——public 读、internal 写，来源只能由本库置位，外部伪造不了；`AddSlide(layout)` 补出来的那一份在 `Placeholder(...)` 里带标记，`PptxWriter.IsEmptyPlaceholder` 从「占位符且无文本」收紧为「带标记且没有任何可见效果」。
  - 「未改动」没有做几何快照去比对 `Left/Top/Width/Height`：一个空框能不能被看见由文本、底色、描边、入场动画四样决定，四样全空时它在页面上画不出任何东西（PDF 侧同理，`PdfRenderer` 的 `DrawSlideText` 在没有非空段落时直接返回），省掉不丢内容；任一样被设过就照常写出。红测因此设的是底色与描边，而不是挪坐标。
  - 修复前失败原文：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: the styled placeholder vanished on save: the slide carries 1 shape(s) instead of 2
    ```

  - 控制断言钉反方向：只填标题的那一页，先立住模型里确实有两个自动占位符，往返后回读只剩 1 个——「程序生成的稿子不该留一个等用户点进去的空框」这条旧承诺不能被这次改动一起丢掉。夹具走内存流写出再 `PptxReader.Load(bytes)` 回读，不碰磁盘；底色断言复用 `Batch4PptTests.cs` 里的 `RequireHex`。
  - 产出字节的变化不止红测那一页：**从文件读回的 pptx 里空着的内容占位符过去会被写出器删掉，现在原样保留**（读端不置标记），代价是这类稿子在 PowerPoint 里看得见"单击此处添加"的空框，收益是往返的形状数与层叠次序不再被静默改写。§9 因此补一行，`src/Wisdom.Ppt/README.md`「版式与占位符」写清省略规则，`docs/Wisdom.Ppt.Tutorial.md` 第 3 节那句「直接 `AddTable` 会让这一页多出一个叠在表格上的空框」按新规则重写——空框在 `.pptx` 与 PDF 里都画不出东西，它占的是模型里的一格。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`；存量 PPT 套件（`Batch4PptTests`、`PptMarkdownTests`、`WriteOutputTests` 的形状计数与往返断言）全绿，说明收紧后的省略条件没改动正常稿子的输出。
- `PT5`（`d3dbe35` + `7e0f7db`）：目标行为「`TextFormat` 表达『未设置』，只写真正设置过的属性」落地为三个只读标记 + 写端三处门控，而不是把公共属性改成可空。
  - 为什么不走可空字段：`TextFormat.FontSize` 是 `public double`、`Color` 是 `public OfficeColor`（结构体），被 Word / Excel / Pdf 三条链路同时消费，把类型改成可空是一次跨包 BREAKING，且要动的 `DocxWriter.cs` 正被另一路会话改写；本次只在类内加 `private` 状态与三个 `public bool` 只读属性，两条构造器与全部属性签名字面未变，读取方拿到的 `FontSize` 仍是 11、`Color` 仍是黑，行为差异只在「写不写得出这个属性」。标记的可见性与 `PT6` 同源：public 读，写权由赋值本身掌握。
  - 无参构造器改成直接写 backing field（`_fontFamily = "Arial"` 等），预置的 Arial/11 磅/黑色因此不点亮标记——这正是「构造器预置默认值击败 `??`」那一类缺陷在 PPT 侧的最后一个入口；两参构造 `TextFormat(fontFamily, fontSize)` 按它自己文档承诺的「字体、字号、黑色」三项全算显式，`HasExplicitFontFamily` 只认非空字体名，`HasExplicitFontSize` 只认正字号。`Clone()` 走 `MemberwiseClone`，标记随格式一起复制。
  - 写端只留 `lang` 是恒写的：`CharProps` 里 `sz` 受 `HasExplicitFontSize` 门控、`a:solidFill` 受 `HasExplicitColor`、`a:latin`/`a:ea` 受 `HasExplicitFontFamily`（两者同进同出，只设东亚字体时西文仍按 `GetAsciiFontFamily()` 的回退写，与修复前一致）。`ResolveFormat`/`IsGenericDefault` 原样保留：整段没设过任何属性的 run 仍走「取基础格式」那条路，`_bodyFormat`/`_titleFormat` 由 `new TextFormat("Calibri", 18/40)` 造出、三项俱为显式，普通稿子的输出因此不变——变化只落在「设了一处、其余留空」的 run 上，`Batch4PptTests` 的主题字体断言（`typeface="微软雅黑"`、`typeface="Consolas"`）与 `DefaultTextFormat` 优先断言全绿就是这条边界未被推开的证据。
  - 修复前失败原文（`net8.0` 与 `net461` 同一条）：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: a run that never set a size was pinned to the model default: sz=1100
       at Wisdom.Office.Tests.Program.RunPartialRunFormatStaysInheritedTest() ... RemediationPptPhaseFourTests.cs:line 107
    ```

  - 夹具走内存流，`a:rPr` 直接从 zip 部件里按文档顺序取（`a:endParaRPr` 不在 `a:r` 下，天然排除）：三条 run 分别只设粗体、只设 22 磅、只设「宋体」，断言第一份既无 `sz` 也无 `a:latin`/`a:ea`/`a:solidFill`，后两份各只写自己那一项、不许臆造另一项；再 `PptxReader.Load` 回读并二次写出，验证读端没在文件里见到的属性不会被当成显式设过——模型侧回读断言 `Bold`、`FontSize=22`、`FontFamily=宋体` 三项仍在，省下来的属性按模型默认值读回，不是丢失。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`。
- `PT3` 结构档（`f9fc99a` + `04ff01b`）：目标行为「写端不伪造模型未表达的意图」（`PT3` 目标 2 的表头一档）落地为四件小事——声明式表头判据、`bandRow` 下线、单元格填充可缺省、`firstRow` 有读回的路。
  - 旧 `HasHeaderRow(rows)` 循环查完显式表头后 `return rows.Count > 0`，等于宣布「只要非空，首行就是表头」；`IsHeaderRowAt` 又把这个条件逐格套用，于是两行数据表出门自带主题 accent1 底色 + 加粗白字。新判据 `HasDeclaredHeader(table, rows)` 只认模型里真说过的两句话：某行 `TableRow.IsHeader`，或给过 `TableShape.HeaderColor`。`IsHeaderRowAt` 的「显式行优先，否则首行」不变，但它现在只在门控通过后才可达——门内不再有猜测。
  - `bandRow` 直接不写而不是补字段：统一模型里没有带状行这一项，恒写 `bandRow="1"` 是替调用者做样式决定；把它建模属 Phase 5 的模型扩展（与 `PT3` 的超链接、组合、`rowSpan` 同批排期）。这次先让写端闭嘴。
  - `a:tcPr` 的填充从「没底色就写 `a:noFill`」改为「没底色就不写」：`a:noFill` 是显式断言「这格不许有底色」，`tableStyleId` 本来要给条纹与整表填充被逐格堵死，这正是 §PT3 里「未设置」与「显式无填充」必须分开的同一件事。同一理由管到补齐用的空格（`placed < gridCols` 那一段）：它在表头行借表头色，免得一条表头被补位格截成两截；在数据行什么都不写。
  - 读端这一次只补 `firstRow` 一条：`a:tblPr@firstRow` 是 `TableRow.IsHeader` 唯一能回来的路（读端从不解析 `a:tableStyles`），不读它，声明过的表头在第二次保存被整段抹平——写端不再猜测之后，这条读回路径就成了保真的一部分，不是可选优化。`lnSpc`/`wrap`/`baseline`/`rowSpan` 与主题 accent 的读端补齐仍留在 Phase 5。
  - 修复前失败原文（`net8.0` 与 `net461` 同一条）：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: firstRow was switched on for a table that never declared a header row
       at Wisdom.Office.Tests.Program.RunUndeclaredTableHeaderStaysPlainTest() ... RemediationPptPhaseFourTests.cs:line 179
    ```

  - 断言分正反两面：未声明的那张既不得出现 `firstRow`/`bandRow`，首格也既无 `a:solidFill` 也无 `a:noFill`，运行不带 `b="1"`；设了 `HeaderColor` 的那张必须写得出 `firstRow="1"`、`1F3A5F` 底色与加粗表头——把「不伪造」做成「该写的还写得出」的反证，避免修复退化成整段不写样式。再加一条往返：`PptxReader.Load` 后声明过的首行 `IsHeader` 仍在，未声明的仍是普通行。夹具继续走内存流，`a:tbl` 从 zip 部件按文档顺序取。
  - `§9` 单独补一行（原先与 `PT1`/`PT2` 合并一行）：输出形状变了，`tableStyleId` 的话头重新生效。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`；`PptMarkdownTests` 的表格往返断言（`HeaderColor` + 三行，回读 3 行、单元格文本 `"华东"`）与 Word 侧共用表格夹具的套件全绿，说明声明过表头的常用写法输出不变。
  - 同批目标 2 还剩两项的收口：`p:style` 未解析出的填充与描边一律落成 `a:noFill`（`PptxWriter.cs:757`、`:769`）已由 `6dce2b6` 落地为「未设置就不写、显式 `BorderStyle.None` 才写空描边」；段落无标记仍恒写 `pPr` 与 `buNone` 一档见下一条。
- `PT3` 结构档·项目符号（`19cee6b` + `adb0f30`）：目标 2 的第三档——`a:buNone` 不再占位，全段无属性的 `a:pPr` 随之省略。
  - `CreateParagraphProperties` 结尾两行是 `pPr.Add(CreateBullet(paragraph.ListMarker)); any = true;`，而 `CreateBullet` 对空标记返回 `a:buNone`：任何段落都带 `pPr`，任何段落都带一条「关掉本层符号」。`buNone` 不是留白而是显式指令，本库自建的母版与版式从不定义符号所以自家产物看不出问题，一旦 `PptxReader.Load` 别人的稿子再 `Write` 回去，母版与版式里定义好的项目符号被逐段压掉——这与 `a:tcPr` 那档「未设置写成显式无填充」是同一个病。
  - 修法只在写端：`CreateBullet` 无标记返回 `null`，调用点判空才 `Add` 并置 `any`；`pPr` 的「什么都没有就返回 null」从此真的成立（方法文档早就这么承诺，调用点 `CreateParagraph:875-876` 也一直容忍 null）。读端不动：`ReadListMarker` 把 `buNone` 映射成 `null`，往返稳定。
  - 修复前失败原文（`net8.0` 与 `net461` 同一条）：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: a paragraph that set nothing still got an a:pPr, housing a fabricated a:buNone
       at Wisdom.Office.Tests.Program.RunUnmarkedParagraphDoesNotSuppressBulletsTest() ... RemediationPptPhaseFourTests.cs:line 321
    ```

  - 夹具是一个文本框三段：第一段声明 `ListMarker = "•"`（对照组，必须仍然写得出 `buChar`）、第二段什么都不设、第三段只设居中正。断言三件事——对照组符号仍在、第二段连 `a:pPr` 都不出现、第三段有 `algn="ctr"` 却没有任何 bullet 子元素；再 `Load` 回读断言标记仍是 `"•"`，二次写出不得把丢掉的 `pPr`/`buNone` 补回来。
  - 已知边界（写进 README 与教程 FAQ）：`Paragraph.ListMarker` 一个值承载两种意思，「显式关掉符号」与「没想过」不可分，载入带 `buNone` 的外来稿再存盘时那份显式关闭不被保留；建模它需要新字段，属 Phase 5 的读端补齐与模型扩展（与 `lnSpc`/`wrap`/`baseline`/`rowSpan` 同批）。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`；`[remediation] ppt phase-four shapes and media ok`、`[ppt-anim]`、`[ppt-theme]` 与 PPT 往返套件全绿，说明自家产物的段落属性写出与回读未受影响。
- `PT3` 读端补齐·文本框折行（`d836400` + `5f15d03`）：`a:bodyPr@wrap` 从「写端钉死 square、读端从不看」变成一句可以往返的声明。
  - 这一档与前三档的病因不同：缺的不是「少写了一句」，而是模型里没地方放。`wrap` 的两种真实取值里，`none`（不许折行——单行标题、竖排标签就是这么写的）在 `TextBoxShape` 上没有任何字段承载，读端读到也只能丢。所以先补一个 public 可空字段 `WrapText`（纯增项，不破坏既有调用），再把两侧接线。
  - 读端只把 `none` 当成一句话：`square` 与属性缺失在文件里都表示「按默认折行」，把它读成 `true` 等于替调用者造一条他没做过的声明（正是 `PT5` 那条教训的形状），一律记 null；`emergency`（紧急收缩）模型无法表达，同样 null 并写进「已知限制」。
  - 写端维持「未设置就写 square」不动——那正是 PowerPoint 自己在属性缺失时的行为，自家产物一个字节都没变。与 `buNone` 那档必须改字节的情形不同，这里「省略」与「写默认值」等价，选了风险低的写法。
  - 修复前失败原文（`net8.0` 与 `net461` 同一条）：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: wrap="none" was flattened to square: the box that must not fold started folding
       at Wisdom.Office.Tests.Program.RunTextBoxWrapNoneSurvivesTest() ... RemediationPptPhaseFourTests.cs:line 390
    ```

  - 夹具做法：新 API 在修复前不存在，红测不能靠「引用一个还不存在的字段」，于是从文件侧取证——写一份两个文本框的稿子，把第一个 `bodyPr` 的属性原文改成 `none`（内存重打包，新增 `SlideEntryText`、`RewriteZipEntryText`、`ReadBodyPrs` 三个助手），`Load` 后二次写出，断言第一个仍是 `none`、对照格仍是 `square`。转绿之后再补模型侧：声明 `false` 写得出 `none`，声明 `true` 与没说过都写 `square` 且都读回 null，第二次写出与第一次同形。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`。本库 PPT→PDF 渲染器按绝对坐标逐段画行、本就不做折行，`wrap` 不参与自家渲染，这个字段只影响写出的 `.pptx` 在 PowerPoint / WPS 里的表现。
  - `PT3` 的读端补齐到这里收口为「`wrap` 已补」；主题 accent（`a:schemeClr`）、超链接、组合形状与旋转翻转仍按 `§11.13` 的排期另开条目。
- `PT11` 线型编码（`805bfdf` + `a5c4c0e`）：条目为本轮新发现，§5 补的是同一件事的登记——`DashStyle` 把两种线型塞进了一个装不下它们的位置。
  - 双线不是划线图案而是**复合线**，规范编码在 `a:ln@cmpd="dbl"`；`ST_PresetLineDashVal` 的取值集合里没有 `dbl`，旧写法既产出一个枚举外的值，又让读端 `MapDashStyle`（取值表里同样没有 `dbl`）回落 `Single`——双线存一次盘成单线，与 `PT3` 那批「静默改版式」同一种后果，只是这次坏在写端自己造了一个非法值。点线更直白：`sysDash` 是长划图案，点线该写 `sysDot`，画出来一直是虚线。
  - 修法两条路径同改：形状描边 `CreateLine` 在 `Double` 时补 `cmpd="dbl"`（其余线型不写这个属性，缺省即 `sng`，实线/虚线字节不动），表格边线 `CreateCellBorder` 本就以 `cmpd="sng"` 起头，改成双线时落 `dbl`；`DashStyle` 只留 `dash` 与 `sysDot` 两档，双线返回 null。读端 `ReadOutline` 先看 `cmpd="dbl"`，`MapDashStyle` 继续把 `dbl` 认成 `Double`，为的是历史文件——1.0.2 之前存出的那份非法值读回来仍然是双线，不改掉用户的稿子。
  - 修复前失败原文（`net8.0`，`net461` 同一条）：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: a:prstDash/@val was written as "dbl", which is a compound-line value, not a dash pattern
       at Wisdom.Office.Tests.Program.RunRemediationPptPhaseFourTest() ... RemediationPptPhaseFourTests.cs:line 28
    ```

  - 夹具是三件东西同页：双线描边矩形、点线描边矩形、一张把 `Borders.InsideHorizontal` 设成双线的表格。断言分四层——整页 `a:prstDash` 里不得出现 `dbl`（枚举外值本身就是缺陷证据，不必等读端）、双线形状的 `a:ln@cmpd` 必须是 `dbl`、点线形状的 `prstDash val` 必须落在圆点档、表格里至少一条 `a:lnL/lnR/lnT/lnB` 带 `cmpd="dbl"`；最后 `PptxReader.Load` 回读断言两个 `Outline.Style` 仍是 `Double` 与 `Dotted`。表格那一路顺便暴露一件事：单元格的 `a:lnL/lnR/lnT/lnB` 读端从不解析（`PptxReader.cs:957-960` 自己写明），所以表格边框只出不进，这条已写进「已知限制」。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`。
- `PT12` 表格边框方向（`32b33a4` + `ec24b6f`）：`Table.Borders` 的六个方向里只有一个真正写得出门。
  - `CreateTableElement` 只留一条 `BorderLine inside = model.Borders.InsideHorizontal`，`CreateTcProperties` 又把 `left/right/top/bottom` 四个变量一并初始化成它，于是 `Top`/`Bottom`/`Left`/`Right`/`InsideVertical` 五份声明在文件里查无此线——外圈 3 磅、内部 1 磅这种写法存一次盘塌成四条同粗的边。它与 `PT3` 那批不同源：那里是写端替调用者多说话，这里是少说话，声明在模型里、落不进文件。
  - 修法是按位置取边而不是加字段：新增 `ResolveEdges(borders, row, rowSpan, totalRows, column, colSpan, gridCols)` 返回一个四边对象，贴外圈的那条边用对应方向，内部用两条分隔线；合并块用 `coveredUntil` 记下块内最末一行，续接格与补齐空格也各自按所在网格位置判定，不再共用一条线。`TableCell.Borders` 的逐边覆盖照旧优先。
  - 字节口径：`Table()` 构造器给的是 `new TableBorders()`，六个方向同为 0.5 磅黑色细实线，所以没动过 `Borders` 的稿子逐字节不变；只有声明过非默认方向的稿子开始按声明画线。读端仍不解析 `a:lnL/lnR/lnT/lnB`（`PptxReader.cs:957-960`），这条留在「已知限制」里，本批只是让写端不再丢声明。
  - 修复前失败原文（`net8.0`，`net461` 同一条）：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: row 0 cell 0 lnT was 808080, expected FF0000
       at Wisdom.Office.Tests.Program.RunTableBorderDirectionsLandTest() ... RemediationPptPhaseFourTests.cs:line 557
    ```

  - 夹具是一张 2 行 3 列的表，六个方向各给一种颜色与线宽（外圈 3 磅四种色、内部 1 磅两色），逐格断言 `a:lnT`/`lnB`/`lnL`/`lnR` 的 `srgbClr@val` 与该取的色一致、`w` 落在 38100 或 12700 上——一张表同时盯住「方向有没有落位」和「线宽有没有落位」，也顺便钉住首行与末行、首列与末列的判定边界。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`。
- `PT13` 表格无边框（`3986130` + `f19f842`）：`CreateCellBorder` 从不看线型，`BorderStyle.None` 也照样补 `a:solidFill`。
  - 这条与 `PT3` 那批形状填充一档是同一个语义在两条路径上的落差：`CreateLine` 早在 `6dce2b6` 就学会「`None` 写 `a:noFill`」，表格边线留在原地——于是统一模型里唯一的「去边框」写法（`TableBorders.None()`，Word 侧 `WordLabels.cs:441` 与门面包 `WorkbookToDocumentConverter.cs:355` 都用它）进了 `.pptx` 变成四条黑色实线，声明被反着执行。同库 DOCX 给 `None` 写 `w:val="nil"`，仓内口径本来就一致，缺的只是这一处实现。
  - 修法三行：`line.Style == BorderStyle.None` 时写 `a:noFill` 并就此收尾，`w`/`cmpd`/`algn` 三个属性照写（PowerPoint 也是这么写无填充的边）。顺带删掉 `line == null` 的分支——`ResolveEdges` 经 `PT12` 之后每条边都拿得到线，那个 null 是上一批留下的死路，留着只会让人以为还能没线。
  - 修复前失败原文（`net8.0`，`net461` 同一条）：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: row 0 cell 0 lnB was 000000
       at Wisdom.Office.Tests.Program.RunCellBorderNoneMeansNoLineTest() ... RemediationPptPhaseFourTests.cs:line 634
    ```

  - 夹具故意把「无线」与「有线」混在一张表里：`Borders = TableBorders.None()` 之后只把 `Top` 设回细黑线。断言首行 `lnT` 仍是 `000000`（证明没把「有线」那一档一起修没），其余三条边一律是 `<a:noFill/>`。`EdgeHex` 这个助手把缺失的边、无填充的边与没颜色的边写成不同占位串，失败消息直接说清文件里是什么。
  - 字节口径：只有设过 `TableBorders.None()` 或某方向 `BorderStyle.None` 的稿子会变，那正是本条要修的差别；默认六向细实线的表与形状描边输出不动。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`。
- 后续按 §10 顺序推进；`PD3` 的单页字元与片段上限、`PdfTextFont` 的 ToUnicode 降级上报，都并入「文档级告警与降级通道」那一批（与 Word `WE7`/`WE9` 同批设计）。
- `PT14` 表格边线读回（`9cbc757` + `a96a0e1`）：写出侧逐格写四条边，读侧从不解析，表格边框只出不进。
  - 这一条是 `PT11`/`PT12` 两批各自留下的那句「要说边界」：那两批只把写端修对了，读端仍写着「不读入」（`PptxReader.cs` 的旧注释自己点名 `a:lnL`/`lnR`/`lnT`/`lnB`）。差别不在自家稿子——新建的表在内存里带着 `Table.Borders`，存几次都按六向落位——在载入：`PptxReader.Load` 出来的表 `TableCell.Borders` 全是 null，`Table.Borders` 是构造器默认的六向细黑线，于是别人的外圈 3 磅四色 + 内部 1 磅淡线，再存一次就是一张均匀黑格网。同库 Word 侧 `DocxPartReader.cs:856-870` 的 `ReadCellBorders` 早就逐边解析、四条全缺才返回 null，本次按同一口径补齐 PPT 读端（`ReadCellBorders` + `ReadCellEdge`）。
  - 两个语义要点值得记下：`a:noFill` 必须读成 `BorderStyle.None` 而不是 null——null 在模型里是「这条边没说过」，会被表级默认线填成一条实线，正好把 `PT13` 修掉的那件事反着做一遍；`algn`、逐格 `marT`/`marB`/`marR` 不建模，读侧不取，写出仍是固定值，这条留在文档的边界句里。
  - 修复前失败原文（`net8.0`，`net461` 同一条）：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: the loaded table kept no cell borders: row 0 cell 0 has Borders=<null>
       at Wisdom.Office.Tests.Program.RunTableCellBordersComeBackTest() ... RemediationPptPhaseFourTests.cs:line 677
    ```

  - 夹具是两张表同页：一张六向各异（外圈 3 磅四色、内部 1 磅灰/紫、`InsideVertical` 给双线），一张 `TableBorders.None()` 只留 `Top` 细线。断言分三层——载入后逐格四边的颜色与线型回到声明值（含双线 `Double` 与 `None`）、末行末列的 ring 判定边界、最后把载入的稿子再写出一次，逐格四条边的 `w` + 颜色签名必须与第一次逐字相同。最后一层才是「往返」的真正断言：只测读回来的对象等于声明，测不出写出侧会不会另取一份默认值。
  - 字节口径：本库新建稿第二次保存逐边不变（测试直接比对两次写出的边线签名）；变化只发生在「载入别人的表格再存盘」这一类稿子上，而那正是本条要修的丢失。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`。
- `WE12` ODT 页码与总页数域（`fbd2c6e` 取证 + `f5bbd98` 修复）：`OdtWriter` 认页码用的是 `Instruction.IndexOf("PAGE")`，而 `NUMPAGES` 里正含 `PAGE` 四个字母——`第 X 页 共 N 页` 这份最常见的页脚写成 ODT 就成了两个当前页码。改成整词判定：`PageFieldTokens` 把指令按字母成词切出并大写，`NUMPAGES`/`PAGES` 先认走写 `text:page-count`，剩下整词 `PAGE` 写 `text:page-number`；`OdtReader` 补 `text:page-count` → `NUMPAGES`。
  - 修复前失败原文：

    ```text
    Unhandled exception. System.InvalidOperationException: Test failed: NUMPAGES wrote no text:page-count, so the total is drawn as a page number: ...<text:page-number text:select-page="current"/>...<text:page-number text:select-page="current"/>...
    at Wisdom.Office.Tests.Program.RunOdtPageCountFieldTest() ... RemediationWordPhaseFourTests.cs:line 397
    ```

  - 只有 ODT 这一路塌：DOCX / RTF / HTML 三个写出器都把 `FieldRun.Instruction` 原样带出去，`FieldRun` 本来就存着指令，所以本轮不动模型；测试用同一份稿子的 DOCX 做对照，断言它读回 PAGE 与 NUMPAGES 两条指令，把「模型没丢」和「ODT 写错了」分开。
  - 顺手记下的边界：ODT 只认这两个指令，其它指令（`SECTIONPAGES`、`DATE` 之类）写出的是域的显示文本而不是指令本身，载回后不再是一个域；`OdtReader` 有 `text:date`/`text:time` 分支而 `OdtWriter` 不产它们，这条读写不对称留给后续条目。
  - 字节口径：只用 `PAGE` 的稿子输出不变（那一条元素的文本与属性都没动），差别只在同时出现总页数的 ODT。
  - 验证：`net8.0` 与 `net461` 双目标均到末行 `All tests finished.`。
