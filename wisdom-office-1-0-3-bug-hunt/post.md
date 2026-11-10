# 一个 `Write(null, path)` 把用户文件变成 0 字节：Wisdom.Office 1.0.3 的十九条缺陷排查实录

> 这是一篇修复实录，不是功能介绍。文里出现的每一段 `Test failed:` 都是真跑出来的，不是手写的示意。

## 引子：三种不会报错的失败

做文档库久了会发现，最贵的 bug 不是抛异常的那种，而是下面这三种：

1. **写文件写失败了，但用户的原文件先被清空了。** 调用方看到的是 `ArgumentNullException`，以为自己传错了参数，回头一看磁盘上的报表已经变成 0 字节。
2. **公式没报错，但每一行都算同一个值。** Excel 里一张跨行填充的表格，重写之后第二行开始全部指向第一行的数据，界面上数字都在，只是全错。
3. **文件能打开，但 PowerPoint 弹"发现需要修复的内容"。** 结构违规不会让库自己报错，因为库按自己的私有形状自洽地读写。

1.0.3 发布前，我把这三类一共十九条缺陷清了一遍。下面是其中五条完整的推理链，以及一套我认为值得抄走的验证纪律。

## 背景：这是什么，为什么会攒下这些债

Wisdom.Office 是一个纯 C# 自研的 Office 文档库，覆盖 Word / Excel / PowerPoint / PDF / Markdown，零第三方依赖——ZIP、XML、OLE2 复合文档、PDF 对象与内容流、OOXML、BIFF8、PresentationML、Markdown 解析器全部手写，同时产出 `net461`、`netstandard2.0`、`net7.0`–`net10.0` 六份资产。

手写格式的直接代价是：**每个格式的错误处理都要自己设计**。用现成库时，"读到坏文件该怎么办"是别人的问题；自己写，就得逐条决定：是抛异常、是跳过、还是降级。而"决定"这件事，很容易在不同模块之间做出不一致的选择——这正是这批缺陷的主要来源。

接手时的第一个发现就很典型：**测试工程根本编不过**。`PptxWriter.cs` 里已经在调用 `HasDeclaredHeader`，但这个方法从未落地，六个目标框架全部 `CS0103`。跟着这条线查下去，才看到它背后是一整批"上一轮改到一半"的入口。

## 纪律一：先让它红，再让它绿

修一个 bug 之前，我要先看到它**红**。不是"我觉得这里有 bug"，而是"我把它改回去，测试真的挂在这个字符串上"。

以那条 313.2 pt 的分隔线为例。现象是：用 `AddShape(AutoShapeKind.Line, 60, 120, 700, 0, null)` 画一条水平线，写出再读回，高度变成 313.2 pt——半个版面。

根因在读端：

```csharp
long cx = ParseLong(Ppt.Attribute(ext, "cx"), 0);
long cy = ParseLong(Ppt.Attribute(ext, "cy"), 0);
if (cx <= 0 || cy <= 0) return null;   // 0 被当成"读不到"
```

`0` 和"属性不存在"被压成了同一个结论，于是几何回落到按版面比例推算的默认框（`0.58 × 540 = 313.2`）。修法是先判属性在不在，再判值：

```csharp
string cxText = Ppt.Attribute(ext, "cx");
string cyText = Ppt.Attribute(ext, "cy");
if (cxText == null || cyText == null) return null;  // 缺失才回落
long cx = ParseLong(cxText, 0);
long cy = ParseLong(cyText, 0);
if (cx < 0 || cy < 0) return null;                  // 负值才是坏几何
```

取证时把 `cx < 0` 那行改回 `cx <= 0`，重跑，得到：

```text
Test failed: a zero-height separator was inflated to 313.2 pt on the way back in
```

这句话本身就是根因说明——313.2 正好是 `0.58 × 540`，看到数字就知道它走了哪条回落路径。

## 纪律二：红测第一个版本经常是假绿的

这是本次最有价值的一课。

平铺水印有个退化步长问题：请求的步长小到整页需要上千格时，600 次绘制预算会全花在页面左下角的一小块里；步长再小下去，`double → int` 还会饱和成 `int.MinValue`，所有 tile 叠到同一坐标。

我写了红测，把修复关掉，跑——**绿的**。

原因不在代码，在夹具：`WatermarkOptions` 的默认 `TileSpacingX/Y` 是 60/40 磅，只要间距不为零，步长就有兜底，退化路径根本走不到。夹具必须显式把两个间距归零：

```csharp
options.FontSize = 0.01;
options.TileSpacingX = 0;   // 少了这两行，测的是压缩率之外的另一条正常路径
options.TileSpacingY = 0;
```

改成零之后才拿到真红：

```text
Test failed: a 0.01 pt block collapsed the tiling into a corner: 600 tiles, furthest at 33.6 x -0.0 (page is 595 x 842)
```

`600 tiles` 说明预算确实花光了，`furthest at 33.6 x -0.0` 说明它们全挤在左下角一条 33 pt 宽的缝里。两个数字一起读，才知道"消失"是挤没的，不是画到页面外。

**教训**：红测首先证明的是"这个夹具能触发这条路径"。默认值会把路径兜住，而默认值写在构造器里，读代码时最不容易注意到。

## 纪律三：一个失败点背后通常有一组同类调用点

`WE10` 是"公开可空的模型字段被写端直接解引用"。文献里列的是 `TextRun.Format`、`TextRun.Text`、`Table.Borders`，但真正的教训在覆盖面。

红测按 `Docx → Rtf → Html → Odt` 顺序跑，因为 fail-fast，第一轮只暴露了两个：

```text
Test failed: the Rtf writer failed on a null model field: NullReferenceException: ...
Test failed: the Html writer failed on a null model field: NullReferenceException: ...
```

修完这两个，测试绿了。但 `OdtWriter` 里有一模一样的四个三元表达式，只是没被测到——它是靠**代码形态审计**（`table.Borders` 在四个分支里各解引用一次）发现的，不是被测试抓到的。

所以修法不是"改到测试不红"，而是"把这一形态的所有出现点一起改掉"：

```csharp
// 先归一，再取边，四个三元表达式不再各解引用一次
TableBorders tableBorders = table.Borders ?? new TableBorders();
```

同类的事情在 E7 上更极端：条目点名四处写出入口，实际有**九个**——`DocxWriter.Write` / `WriteTemplate`、`WordFormatConverter.Save`、`XlsxWriter.Write`、`PptxWriter.Write` / `WriteTemplate`、`PdfRenderer.Render` / `RenderPresentation`、`OfficeFile.Save`。

## 案例拆解

### 一、先截断再写：最贵的一条

九个入口都是同一句话：

```csharp
using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
{
    Write(workbook, stream);   // 参数校验在流重载内部
}
```

`FileMode.Create` 在 `Write` 被调用之前就已经把目标文件截断了。于是 `Write(null, path)` 的语义变成"先把用户的文件清空，再告诉用户参数不对"。

修复是一个统一入口：同目录临时文件 → 写完 `Flush` → `File.Replace`（同卷原子替换）或 `File.Move`（目标不存在时）。失败只删临时文件，目标一个字节都不动。

```csharp
string target = Path.GetFullPath(path);
string temporary = Path.Combine(directory, "." + Path.GetFileName(target) + "." + Guid.NewGuid().ToString("N") + ".tmp");
// ...写 temporary...
if (File.Exists(target)) File.Replace(temporary, target, null);
else File.Move(temporary, target);
```

红证是把 `XlsxWriter` 临时退回 `FileMode.Create`：

```text
Test failed: a failed xlsx write left the original file with 0 byte(s) instead of the 13 it had
```

**必须一起说的副作用**：目标文件现在是"被替换"而不是"原地覆盖"，所以文件系统层面的 inode / 创建时间会变；被其他进程独占打开的目标，现在会以 `IOException` 失败而**不破坏原文件**。这是有意的取舍，但它确实改变了可观测行为，得写进文档。

### 二、把一个属性当字段用在循环条件里

Excel 读表有这么一段：

```csharp
for (int row = firstDataRow; row <= sheet.LastRow; row++)   // LastRow 每次都遍历全表单元格
```

`LastRow` 的实现是遍历 `_cells` 取最大行号，所以循环体里每判断一次就扫一遍全表，复杂度是"行数 × 单元格数"。8000 行、两列时：

```text
Test failed: reading 8000 rows took 1600 ms; LastRow is still being evaluated inside the loop
```

修法是一行快照：

```csharp
int lastRow = sheet.LastRow;
for (int row = firstDataRow; row <= lastRow; row++)
```

**这类缺陷不好用单元断言描述，只能用"数量级断言"**：给一个足够大的输入和一个留足余量的阈值（这里 600 ms），并在断言消息里把实测毫秒打出来。这样即使将来变慢，失败信息也能告诉你"慢了多少"，而不是只说"超时了"。

### 三、`Sheet1!A2` 里的 `A2` 为什么没跟着走

共享公式（`<f t="shared" si="0">`）只有主格带公式文本，从格只带编号，读取时按"从格相对主格的偏移"平移相对引用。判断"这是不是一个引用起点"的函数里有一行：

```csharp
if (IsLetter(previous) || IsDigit(previous) || previous == '_' || previous == '.' || previous == '!')
    return false;
```

`!` 也在排除表里。于是 `Sheet1!A2` 的 `A2` 前面是 `!`，直接跳过——公式还在，但每个从格都指向主格那一行，**静默算错**。

红证：

```text
Test failed: a sheet-qualified relative reference did not shift with the dependent: Sheet1!A2+1
```

把 `!` 从排除表去掉就行，但要先确认它当初为什么被加进去：表名本身（`Sheet1`）不会误判，因为它要么有 4 个以上字母被 `letters >= 4` 挡掉，要么后面紧跟 `!` 被"下一个字符不能是感叹号"挡掉。删掉这一项是有依据的，不是碰运气。

### 四、一份"看起来对"的 XML

`p:timing`（PowerPoint 动画时间线）是最容易自洽地写错的地方，因为它只在 PowerPoint 里才会被严格检查。

旧实现的问题一次能数出七处：一个 `p:par` 里并列两个 `p:cTn`（`CT_TLTimeNodeParallel` 只允许一个）；`seq` 与根 `cTn` 平级（应住在根 `cTn` 的 `childTnLst` 里）；两个节点同用 `id="1"`；`dur="indef"`（规范词法是 `indefinite`）；类别写进 `type` 而规范字段是 `nodeType`；`p:set` 里塞了 `a:clrMapOvr`（那是颜色映射，不是 `p:set` 的内容模型）；属性名写成 `presetId` 而规范是 `presetID`。

重写后的骨架是：

```text
p:timing
└─ p:tnLst
   └─ p:par
      └─ p:cTn(nodeType="tmRoot", dur="indefinite", id=1)
         └─ p:childTnLst
            └─ p:seq
               └─ p:cTn(nodeType="mainSeq", id=2)
                  └─ p:childTnLst
                     └─ p:par › p:cTn(触发) › p:childTnLst › p:par › p:cTn(延迟) › ... › p:cTn(效果)
```

延迟单独占一个 `cTn`，是 PowerPoint 表达"单击后延迟 0.25 秒"的真实写法；"每一层都用 `par > cTn` 包"也不是啰嗦——`CT_TLTimeNodeList` 只接受 `par` / `seq` 这类时间节点，裸的 `cTn` 不是合法子元素。

顺带暴露了读端一个隐藏依赖：它按"父节点的父节点"两跳去找延迟，而我新加的中间层正好让它跨过了延迟节点。改成逐级向上检查每个 `cTn` 的 `stCondLst` 才正确——**写端结构一变，读端的"快捷跳转"就会失效**，这类耦合只能靠往返断言兜住。

```text
Test failed: the root cTn must carry nodeType="tmRoot":
```

**诚实的边界**：这一条只做了结构断言与往返，没有 PowerPoint 真机验证。术语表能证明"符合架构"，不能证明"PowerPoint 一定接受"。这种验证缺口必须写出来，不能当成已解决。

### 五、写端不要替调用者表态

有一组缺陷是同一个毛病的不同表现：**写端把"没设置"写成了"显式不要"**。

| 写法 | 实际语义 | 后果 |
| --- | --- | --- |
| 表格没声明表头，仍写 `firstRow="1"` 并把首行刷成加粗白字 | "首行是表头" | 两行数据表也自带表头样式 |
| 单元格没底色，补 `a:noFill` | "这格不许有底色" | 把 `tableStyleId` 的条纹堵死 |
| 形状没设填充/描边，补 `a:noFill` + 空 `a:ln` | "这形状不许有填充" | 主题与版式继承的外观被压掉 |
| 只加粗的 run，写出 `sz="1100"` + 显式字体 | "字号 11、字体 Calibri" | 版式里 18pt 的正文缩成 11pt |
| 上下标没写 `baseline` | （读端按 `baseline` 还原） | 往返一次上下标就丢 |

对应的三条红证，形态各异但同源：

```text
Test failed: firstRow was switched on for a table that never declared a header row
Test failed: a shape that never set a fill was pinned with a:noFill, silencing the inherited one
Test failed: superscript was lost: the writer never emits a:rPr/@baseline
```

修法统一成一句话：**只在调用者真的设过时才写**；`null` 表示"没说过"，写端回落到该元素的原生默认，而不是替用户拍板。这条规则后来也用来修 `PT9` 的切换方向/速度/点击语义——`TransitionDirection` 做成可空，`null` 时按效果的原生默认方向写，输出与旧稿逐字节一致。

## 怎么证明"真的修好了"

三条硬要求，缺一条我都不认这一条修完：

**1. 回退源码能复现。** 每条的失败原文都是把修复改回去跑出来的，不是根据代码推断的。这也是为什么全文的失败信息都带着具体数字（`0 byte(s) instead of the 13`、`1600 ms`、`313.2 pt`、`33.6 x -0.0`）——数字能反推路径。

**2. 双目标都到末行。** `net8.0` 与 `net461` 各跑一遍，都以 `All tests finished.` 结束。跨框架差异不是理论问题：`.NET Framework` 的 `ZipArchive` 惰性解析中央目录，非法条目名要等到枚举时才抛 `ArgumentException`，而 .NET 上是 `InvalidDataException`。同一份代码在两个目标上给出不同异常类型，调用方的 `catch` 就会漏。

**3. 发布前做一次全新克隆验证。** 这一条抓到了一个真实问题：仓库里已入库的测试文件引用了**三个一直没被跟踪的文件**（`XlsxReadFidelityTests.cs`、`ImageFidelityTests.cs`、`XlsxTheme.cs`），只推已提交的历史，别人克隆下来会直接编译失败。补提交之后，把仓库克隆到临时目录 → 构建 0 警告 → 跑 `net8.0` 全套到末行，才敢说"发布快照是自包含的"。

**能自动化的才叫修好。** 这次有两条没做自动断言，原因写在文档里：`FontResolver` 在锁外读缓存（并发时序不可稳定复现，需要注入调度器）、ODT 打开入口绕过 `PackageArchive`（要构造 .NET Framework 惰性解析的时机差）。这两条只做代码对齐，不假装覆盖。

## 没修的部分，以及为什么值得写出来

一份只讲"全修好了"的修复记录没有参考价值。这次明确留了四类：

- **要真机验证的**：`p:timing` 只做结构断言，PowerPoint 的实际接受度需要人工抽查。
- **要动模型的**：`PT3` 还剩 `wrap`、主题 accent、超链接、组合与旋转翻转——它们要么给模型加类型，要么改写端的关系表，属里程碑级。
- **要定契约的**：部件解析失败的告警只给 Word 侧建了最小通道（`OfficeDocument.Warnings`），没有替其他包先定形状。异常还是告警，是个需要统一设计的决定。
- **改了也只是删代码的**：`UnitConverter` 里两分支同值的死代码，记录但不改——它不改变任何可观测行为，顺手清理可以留给下一次。

把"没做"写进文档，比假装做完有用得多——下一个人接手时，不会再花半天确认"到底测没测过"。

## 结果

1.0.3 是这一轮的交付：七个包（`Wisdom.Office` 门面 + Core / Pdf / Word / Excel / Ppt / Markdown），每个包六份目标框架资产、包内自带 README 与 LICENSE、依赖区间锁在同一版本（避免门面包拉到旧依赖）；`.snupkg` 符号包一并产出。打完之后逐包解包核对 nuspec，确认版本号、六份 `lib/`、依赖区间都对。

仓库侧同步到 GitHub 与 Gitee 两个远程，本次共 58 个提交。

一个不在计划内的结尾：推送到 nuget.org 时返回 `401`——这台机器上没有配置 API Key。包已经备好并核对过，卡在凭证而不是代码。这种"发布流程本身也会失败"的事实，同样值得写进记录里。

## 给做文件格式库的六条

1. **先红后绿，而且要先证明夹具能红。** 默认值会把退化路径兜住，假绿比没测更危险。
2. **修一组同类调用点，不是修一个失败点。** 红测只暴露它跑到的那一条；其余靠形态审计。
3. **写端不要替调用者表态。** "没设置"和"显式不要"在 OOXML 里是两种字节，混了就是静默改样式。
4. **"写不进去"和"写坏了"要分开。** 前者该抛，后者该降级并留痕；最糟的是两者都表现成"这页没有内容"。
5. **目标文件永远不要先截断。** 临时文件 + 原子替换，失败时用户的原件还在。
6. **把没测的部分写出来。** 没有真机验证、没有并发断言，就明写；这比一句"已修复"可信得多。
