# Wisdom.Pdf

`Wisdom.Pdf` 是**自研的 PDF 引擎**：按统一文档模型排版生成 PDF 1.7（含 TrueType 子集字体嵌入与图片），解析既有 PDF 并提取文本与表格，还能在对象级编辑文档——合并拆分页面、页面旋转、文字与图像水印、口令加密读写、AcroForm 填表与扁平化。不依赖任何第三方 PDF 库，也不调用打印驱动。

- 包名：`Wisdom.Pdf`
- 版本：`1.0.2`
- 依赖：`Wisdom.Office.Core`（唯一依赖）
- 第三方依赖：无
- 命名空间：`Wisdom.Pdf`

本文是能力与 API 参考；分步实战教程见源码仓库的 `docs/Wisdom.Pdf.Tutorial.md`。

## 支持的框架

| 目标框架 | 包内资产 | 典型运行时 |
| --- | --- | --- |
| `net461` | `lib/net461` | .NET Framework 4.6.1 / 4.7.2 / 4.8.x |
| `netstandard2.0` | `lib/netstandard2.0` | .NET Core 2.0-3.1、.NET 5/6、Mono、Unity |
| `net7.0` | `lib/net7.0` | .NET 7 |
| `net8.0` | `lib/net8.0` | .NET 8 |
| `net9.0` | `lib/net9.0` | .NET 9 |
| `net10.0` | `lib/net10.0` | .NET 10 / 11 |

ZIP、XML、压缩流全部自行实现；.NET Framework 与 .NET 上的 inflate 行为已在测试中对齐。

## 安装

```bash
dotnet add package Wisdom.Pdf --version 1.0.2
```

只需要「任意文档 → PDF」的转换或报表导出时，请引用 `Wisdom.Office` 门面包；本包接收的是 `Wisdom.Office.Core` 的 `OfficeDocument` 模型。

## 生成 PDF

```csharp
using Wisdom.Pdf;
using Wisdom.Office.Core;

OfficeDocument document = new OfficeDocument();
document.Metadata.Title = "季度总结";

Section section = document.AddSection(PageSetup.A4());

Paragraph title = new Paragraph("季度总结");
title.HeadingLevel = 1;
title.Alignment = ParagraphAlignment.Center;
title.KeepWithNext = true;
section.Blocks.Add(title);

Table table = new Table();
table.Columns.Add(new TableColumn(160));
table.Columns.Add(new TableColumn(100));
table.AddRow("项目", "金额");
table.AddRow("合计", "2,288.40");
table.RepeatHeaderRow = true;
section.Blocks.Add(table);

PdfRenderer renderer = new PdfRenderer();
renderer.Render(document, "out.pdf");              // 亦支持 Stream
byte[] bytes = renderer.RenderToBytes(document);
```

排版能力：自动分页与跨页表格、`IsHeader` 行在每页顶部重复、合并单元格（`ColSpan`/`RowSpan`）的跨度原子性拆分（不会把一组跨列单元拆到两页）、页眉页脚、`{page}`/`{pages}` 页码替换、段落对齐与缩进、行距与段前后距、`PageBreakBefore`、`PageBreak`、图片、CJK 与西文混排换行。

```csharp
// 页脚页码：PDF 渲染时替换成真实页码，同一份模型写 DOCX 则输出 PAGE 域
Paragraph footer = section.Footer.AddParagraph(string.Empty);
footer.AddText("第 ");
footer.Inlines.Add(new FieldRun("PAGE", "{page}"));
footer.AddText(" 页　共 ");
footer.Inlines.Add(new FieldRun("NUMPAGES", "{pages}"));
footer.AddText(" 页");

// 横向 A4 / 自定义页边距
PageSetup wide = PageSetup.A4();
wide.IsLandscape = true;                 // 赋值即自动交换宽高
wide.MarginLeft = wide.MarginRight = 56;
wide.HeaderDistance = 36;
Section s2 = document.AddSection(wide);
```

## 字体

生成 PDF 必须嵌入 TrueType/TTC 字体。`FontResolver` 默认登记各系统的常见字体目录（Windows Fonts、`%LOCALAPPDATA%\Microsoft\Windows\Fonts`、`/usr/share/fonts`、macOS 目录等），缺失目录自动跳过；找不到请求字体时按码位回退到能覆盖该字符的已注册字体（含 Noto、Liberation、文泉驿等常见命名）。

```csharp
FontResolver resolver = new FontResolver();

// 显式注册（文件路径或字节；TTC 可指定 faceIndex）
resolver.RegisterFont("Source Han Sans", false, false, @"C:\fonts\SourceHanSansSC-Regular.ttf");
resolver.RegisterFont("My Font", true, false, File.ReadAllBytes("bold.ttf"), faceIndex: 1);

PdfRenderer renderer = new PdfRenderer(resolver);
renderer.Render(document, "out.pdf");

// 查询与手工解析
TrueTypeFace face = resolver.Resolve("Microsoft YaHei", bold: false, italic: false);
TrueTypeFace fallback = resolver.ResolveFallback(0x4E2D, false, false);
bool isCjk = FontResolver.IsCjkCodePoint(0x4E2D);
```

`TextFormat` 的西文/中文族名分开设置，便于中英混排：

```csharp
TextFormat body = new TextFormat("Times New Roman", 10.5f);
body.EastAsiaFontFamily = "宋体";
```

字体只写出实际用到的字形（子集化），文档体积与用到的字符数相关，而不是字体文件大小。

## 解析 PDF

```csharp
using Wisdom.Pdf;

PdfDocument pdf = PdfDocument.Load("input.pdf");   // 亦支持 Stream / byte[]

Console.WriteLine(pdf.PageCount);
Console.WriteLine(pdf.Info.Title);
Console.WriteLine(pdf.Info.Author);
Console.WriteLine(pdf.Info.CreationDate);
Console.WriteLine(pdf.ExtractText());              // 全文

foreach (PdfPage page in pdf.Pages)
{
    Console.WriteLine($"{page.Number}: {page.Width:F0}x{page.Height:F0} rot={page.Rotation}");

    foreach (PdfTextLine line in page.GetTextLines())
    {
        Console.WriteLine($"{line.Y:F1} {line.Text}");
        foreach (PdfTextFragment f in line.Fragments)   // 逐片段带坐标、字号
            Console.WriteLine($"    {f.X:F1} w={f.Width:F1} size={f.FontSize:F1} {f.Text}");
    }

    foreach (PdfTable table in page.GetTables())
    {
        Console.WriteLine($"表格 {table.RowCount} x {table.ColumnCount}");
        Console.WriteLine(table.GetCell(0, 0));
        Console.WriteLine(table.ToText("\t", "\n"));
    }
}
```

转成统一模型后即可交给 Word 侧输出：

```csharp
OfficeDocument document = pdf.ToOfficeDocument();
```

## 编辑既有 PDF 与水印

`PdfFile` 是对象级编辑入口：直接读写交叉表、按对象号增删改、合并拆分页面，不需要重建文档模型。

```csharp
using Wisdom.Pdf;

PdfFile file = PdfFile.Load("contract.pdf");        // 加密文档传口令：Load(path, "pwd")
file.SetPageRotation(0, 90);                        // 首页转成横向阅读
file.RemovePage(file.PageCount - 1);                // 删除末页
file.Save("contract-edited.pdf");                   // 整量重写
file.SaveIncremental("contract-lite.pdf");          // 只追加改动对象，原字节保留
```

跨文档取页与拆分：

```csharp
PdfFile cover = PdfFile.Load("cover.pdf");
PdfFile body = PdfFile.Load("body.pdf");
cover.AppendPagesFrom(body, 0, 1, 2);               // 把 body 的前三页追加过来
PdfFile merged = PdfFile.Merge(cover, body);
IReadOnlyList<PdfFile> parts = merged.Split(3);     // 每 3 页一份
```

口令加密读写走标准安全处理器（`/Filter /Standard`）：

```csharp
PdfFile file = PdfFile.Load("plain.pdf");
PdfEncryptionOptions options = new PdfEncryptionOptions();
options.Algorithm = PdfEncryptionAlgorithm.Aes128;            // Rc4_40 / Rc4_128 / Aes128 / Aes256
options.UserPassword = "reader-pwd";                          // 打开文档的口令
options.OwnerPassword = "owner-pwd";                          // 修改权限的口令，留空等同用户口令
options.Permissions = PdfPermissions.Print | PdfPermissions.FillForms;
options.EncryptMetadata = true;
file.SetEncryption(options);
file.Save("secret.pdf");

PdfFile opened = PdfFile.Load("secret.pdf", "reader-pwd");    // 口令不对抛 PdfPasswordException
```

写出端默认用 AES-256（R6，`/CFM /AESV3`），读取端覆盖 RC4-40（R2）、RC4-128（R3）、AES-128（R4）与 AES-256（R5/R6）；
四种算法的产物都能被第三方阅读器口令开箱，第三方加密的文档也能被本库读回。

文字水印一行即可，默认斜跨 45°、浅灰、30% 不透明度、落在背景层：

```csharp
PdfFile file = PdfFile.Load("report.pdf");
file.AddWatermark("机密文件 CONFIDENTIAL");          // 中文自动回退到中文字体
file.Save("report-watermarked.pdf");
```

需要更多控制时使用 `WatermarkOptions`：

```csharp
WatermarkOptions options = WatermarkOptions.ForText("内部资料 INTERNAL USE");
options.Layer = WatermarkLayer.Foreground;          // 前景层压在正文之上
options.Placement = WatermarkPlacement.Tile;        // Diagonal/Center/Tile/四角/Free
options.FontSize = 28;
options.Bold = true;
options.Color = new OfficeColor(0xFF, 0x33, 0x33);
options.Opacity = 0.15;
options.Angle = 30;                                 // 逆时针角度，斜跨与平铺默认 45
options.AutoFit = true;                             // 文字块超出页面时自动缩小
options.Resolver = FontResolver.Default;            // 指定字体来源，默认已带中文回退
file.AddWatermark(options, 0, 1);                   // 只加给第 1、2 页（下标从 0 起）
```

图像水印：

```csharp
WatermarkOptions stamp = WatermarkOptions.ForImage(File.ReadAllBytes("seal.png"));
stamp.Placement = WatermarkPlacement.BottomRight;
stamp.ImageWidth = 120;                             // 只给宽时高度按原始比例推出
stamp.Opacity = 0.6;
file.AddWatermark(stamp);
file.AddImageWatermark("seal.png");                 // 便捷重载，直接读图片文件
```

第二个参数是图层而不是页码：`AddWatermark("文字", 0)` 里的 `0` 会隐式转成 `WatermarkLayer.Background`，
要给指定页加水印请使用 `AddWatermark(options, 0)`。

实现方式：全部页面共用同一份字体与图像资源（只嵌入一次），每页新增一个与 MediaBox 等大的 Form XObject
和一段 `q … /Wm Do … Q` 内容流；背景层把这段插到 `/Contents` 前面，前景层追加到后面。页面自带 `/Rotate`
时写出端预旋转补偿，水印始终相对阅读方向摆放。带 `/StructParents` 的标签 PDF 会把水印标成 `/Artifact`，
不破坏无障碍阅读顺序；不透明度小于 1 时自动把版本升到 1.4 并写入 `/ExtGState` 与透明 `/Group`。

## 表单填写与扁平化

`PdfFile` 直接读写 AcroForm：读出字段树、写入值、自行生成外观流，或把填好的表单扁平化成固定版式。

```csharp
PdfFile file = PdfFile.Load("form.pdf");
if (file.HasForm)
{
    PdfForm form = file.GetForm();
    foreach (PdfFormField field in form.Fields)
        Console.WriteLine(field.FullyQualifiedName + " " + field.Kind + " = " + (field.Value ?? ""));
}
```

批量填写：

```csharp
Dictionary<string, string> values = new Dictionary<string, string>();
values["applicant.name"] = "王小明 Wang Xiaoming";  // 层级字段用 "父域.子域" 的点号名
values["applicant.agree"] = "true";                 // 复选框：true/状态名 勾选，false/Off/"" 取消
values["order.plan"] = "pro";                       // 单选组：状态名、选项名或 1 起的序号
values["applicant.city"] = "shenzhen";              // 组合框：/Opt 里的导出值
int filled = file.FillForm(values);                 // 填完统一生成外观，中文字体自动子集嵌入
file.Save("form-filled.pdf");
```

单字段用 `file.SetFormFieldValue("applicant.name", "李四")`，或直接 `field.Value = "李四"`；给签名域赋值会抛
`InvalidOperationException`。`FillForm(values, false)` 只写值不生成外观，此时 `/NeedAppearances` 置为 `true` 交给阅读器排版，
调用 `GenerateFormAppearances()` 后自动复位——所以"填了值但外观还是旧的"这种中间态不会在阅读器里显示错值。

扁平化：把值烧进页面内容，控件与字段一并移除。

```csharp
int flattened = file.FlattenForm();   // 内部先生成外观，再把 /AP /N 画进页面
file.Save("form-flat.pdf");           // 之后 HasForm 为 false，文档不再是可填表单
```

画不出外观的字段（密码域、没有 `/AP` 的字段、控件没挂进 `/Annots` 的字段）会原样保留，扁平化绝不会把它们的值丢掉；
签名域与普通按钮同样保持交互态。

混合表单（AcroForm + XFA）里阅读器优先渲染 XFA，填进去的 AcroForm 值可能看不见，用 `file.RemoveFormXfa()` 整体剥离。

字段模型可读出的信息：`Kind`（Text / CheckBox / RadioButton / PushButton / ComboBox / ListBox / Signature / Unknown）、
`IsMultiline`、`IsPassword`、`IsReadOnly`、`IsTerminal`、`MaximumLength`、`Options`、`SelectedValues`、`OnState`、
`PageIndex`、`Rect`、`DefaultAppearance`、`Alignment`、`Parent`/`Children`、`Flags`，以及对象级的 `ObjectNumber` 与
`Dictionary`（用于检视 `/AP` 这类尚未封装的原始条目）。

实现要点：值写入字段 `/V`（文本域为 UTF-16 字符串，按钮与选项域为状态名）；单选组的值记在父域，各选项只同步自己的 `/AS`；
外观流按控件矩形排版，`/DA` 字号为 0 时自动缩字号、多行域自动折行，勾选标记用矢量边框与勾/圆点画出，不依赖 ZapfDingbats；
文字走本库既有的 Identity-H 子集嵌入与 ToUnicode 映射，因此中文值在阅读器显示和文本提取里都可读。

## 读取上限与安全

解析不受信任的 PDF 时可下调上限，超限会中止而不是无界占用内存（解压超限抛出 `DecompressionLimitException`，结构非法抛出 `InvalidDataException`）：

```csharp
PdfReadLimits.MaxDecodedStreamBytes = 64L * 1024 * 1024;  // 单流解压后字节上限，默认 256 MB
PdfReadLimits.MaxObjectNumber = 5000000;                  // 对象号上限，默认 100 000 000
PdfReadLimits.MaxXrefEntries = 200000;                    // 交叉引用表条目上限，默认 2 000 000
PdfReadLimits.MaxObjectStreamObjects = 100000;            // 对象流内对象上限，默认 1 000 000
PdfReadLimits.MaxPageTreeDepth = 32;                      // 页面树遍历深度上限，默认 64
PdfReadLimits.MaxObjectDepth = 16;                        // 数组/字典嵌套深度上限，默认 64
PdfReadLimits.MaxGlyphMapEntries = 50000;                 // 单个字体 CMap 与 /W 宽度表条目上限，默认 300 000
```

这些上限不只是"报错前少分配点内存"：流内容按 `OfficePackageLimits.CopyBounded` 分块复制，**不会按对象声明的 `/Length` 预分配缓冲区**，声明长度伪造到 `int.MaxValue` 也拿不到那么多内存；嵌套深度上限挡住畸形词元把调用栈打满；`MaxObjectNumber` 同时约束对象号计数器与 `/Size`，谎报 `2147483647` 的文件不再能把编辑后的输出搅坏；畸形 CMap 用大量各自合法的跨度堆映射表时，超出 `MaxGlyphMapEntries` 的区间直接丢弃。

口令加密的取舍：`/CFM` 解析不出加密算法时（`/StmF`、`/StrF` 指向不存在的密码筛选器）在派生密钥**之前**就报错，而不是让算法混淆伪装成"口令错误"；RC4 / AES 解密后的 PKCS#7 填充按内容校验，填充不合法时保留原文而不是截掉一段乱字节。

库内自行实现 inflate（`Zlib` / `Adler32`）与对象/词法解析，遇到结构损坏时抛 `InvalidDataException`，不会越界读取。

## 类型一览

| 类型 | 用途 |
| --- | --- |
| `PdfRenderer` | `OfficeDocument` → PDF（文件 / 流 / 字节） |
| `PdfFile` | 对象级编辑：`Load`、`Save`、`SaveIncremental`、`GetObject`/`SetObject`/`AddObject`/`RemoveObject`、页面与元数据读写、加密 |
| `PdfFile`（页面） | `PageCount`、`GetPageDictionary`、`GetPageContentBytes`、`SetPageContent`、`AppendPageContent`、`PrependPageContent`、`SetPageResource`、`SetPageRotation`、`SetPageCropBox`、`RemovePage`、`MovePage`、`ReorderPages`、`InsertPagesFrom`、`AppendPagesFrom`、`Merge`、`Split`、`ExtractPages` |
| `PdfFile`（水印） | `AddWatermark`、`AddImageWatermark` |
| `PdfFile`（表单） | `HasForm`、`GetForm`、`SetFormFieldValue`、`FillForm`、`GenerateFormAppearances`、`FlattenForm`、`RemoveFormXfa` |
| `PdfForm` / `PdfFormField` / `PdfFormFieldKind` / `PdfFormFieldRect` | AcroForm 字段树、单个字段的读取与赋值 |
| `WatermarkOptions` / `WatermarkLayer` / `WatermarkPlacement` | 水印内容、图层与摆放 |
| `PdfEncryptionOptions` / `PdfEncryptionAlgorithm` / `PdfPermissions` / `PdfPasswordException` | 口令与权限加密 |
| `PdfDocument` | 解析入口：`Load`、`PageCount`、`Pages`、`Info`、`ExtractText`、`ToOfficeDocument` |
| `PdfPage` | `Number`、`Width/Height`、`Rotation`、`GetTextLines`、`GetTables`、`ExtractText`、`GetContentBytes` |
| `PdfTextLine` / `PdfTextFragment` | 带坐标与字号的文本行、文本片段 |
| `PdfTable` | 表格识别结果：`RowCount`、`ColumnCount`、`Cells`、`GetCell`、`ToText` |
| `PdfDocumentInfo` | 信息字典元数据 |
| `FontResolver` / `TrueTypeFace` / `PdfFont` / `StandardPdfFont` / `TrueTypePdfFont` / `SubsetResult` | 字体解析、嵌入与子集化 |
| `PdfReadLimits` | 解析资源上限 |
| `DecompressionLimitException`（来自 `Wisdom.Office.Core`） | 解压超限信号 |
| `PdfObject` 族（`PdfDictionary`、`PdfArray`、`PdfName`、`PdfString`、`PdfNumber`、`PdfStream`、`PdfReference`、`PdfBoolean`、`PdfNull`）、`PdfObjectWriter` | 底层对象模型与序列化，供需要直接操作 PDF 结构的场景使用 |

## 已知限制

- 加密读写支持标准安全处理器（RC4-40、RC4-128、AES-128、AES-256），四种算法已与第三方阅读器双向验证；写出统一用 R6，读取兼容 R5。公钥加密与 Adobe 自定义安全处理器不在范围内。
- 表单：支持 AcroForm 的读取、填写、外观生成与扁平化；XFA 只能整体剥离，不解析 XFA 模板；本包不生成新的表单字段，要给空白文档加控件请使用 `PdfFile` 的对象级 API；签名域与普通按钮不参与扁平化。
- 字体嵌入支持 TrueType `glyf` 轮廓；CFF/OpenType（`OTTO`）字体不能作为嵌入源（部分厂商以 `.otf` 后缀发布 TrueType 轮廓的字体可正常使用）。
- PNG 不支持隔行扫描（interlaced）图片；图片支持 PNG 与 JPEG。
- `GetTables()` 是启发式识别（基于文本位置与线条），复杂版式可能需要人工校正。
- 文本提取依赖文本操作符，扫描件需要额外 OCR；本包不含 OCR。旋转文字按流向归并，但旋转页面上的表格识别仍以轴对齐坐标为准。
- `PdfFile` 做的是对象级改写，签名文档与带对象流的文档在整量保存时会重建交叉表；数字签名（增量保存后校验原字节）尚不支持。
- `OfficeDocument` → PDF 走 `PdfRenderer` 重新排版，不保留来源文档的字节级布局；需要保留原版面请用 `PdfFile`。
- 无桌面字体的精简容器环境需预先安装字体（如 Noto、文泉驿），否则生成阶段无法解析字体。
