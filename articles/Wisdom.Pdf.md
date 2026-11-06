# 不装 Office、不引第三方：Wisdom.Pdf 自研 PDF 引擎的边界在哪里

> 「纯 .NET 操作 Office」系列 · PDF 篇

## 为什么要自研一个 PDF 库

服务端生成或处理 PDF，能选的方案看着不少，实际每一条都有代价。iText 功能最全，但商用授权这一关就把个人项目和小团队劝退；PdfPig 之类偏重"读"，排版生成和对象级编辑基本得自己补；最省事的是调 Word/Excel 转 PDF，但那要求机器上装 Office 或者跑打印驱动，容器里根本走不通。

`Wisdom.Pdf` 是第四条路：**自己实现一套 PDF 引擎**。按统一文档模型排版生成 PDF 1.7（含 TrueType 子集字体嵌入与图片），解析既有 PDF 提取文本与表格，还能在对象级编辑——合并拆分页面、页面旋转、文字与图像水印、口令加密读写、AcroForm 填表与扁平化。全文只有一个内部依赖 `Wisdom.Office.Core`（数据模型包），**不依赖任何第三方 PDF 库，也不调用打印驱动**。NuGet 安装：

```bash
dotnet add package Wisdom.Pdf
```

支持 `net461` / `netstandard2.0` / `net7.0` / `net8.0` / `net9.0` / `net10.0` 六份目标框架，ZIP、XML、压缩流全部自行实现，.NET Framework 与 .NET 上的 inflate 行为已在测试中对齐。需要「任意文档 → PDF」转换或报表导出的话引用 `Wisdom.Office` 门面包更合适；本包接收的是 `Wisdom.Office.Core` 的 `OfficeDocument` 模型。

这篇按场景覆盖：生成、字体、解析、对象级编辑、加密、水印、表单，以及读取上限。

## 一、生成 PDF：排版能力是重点

从模型开始，`PdfRenderer` 负责把 `OfficeDocument` 落成 PDF：

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

这块真正花力气的是排版：**自动分页与跨页表格**、`IsHeader` 行在每页顶部重复、合并单元格（`ColSpan`/`RowSpan`）的跨度原子性拆分——一组跨列单元不会被拆到两页上，页眉页脚、`{page}`/`{pages}` 页码替换、段落对齐与缩进、行距与段前后距、`PageBreakBefore`、`PageBreak`、图片，以及 CJK 与西文混排换行。

页脚页码是最能体现"统一模型"设计的地方。同一份模型，PDF 侧渲染成真实页码，写 DOCX 侧则输出 `PAGE` 域：

```csharp
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

## 二、字体：子集嵌入与按码位回退

生成 PDF 必须嵌入 TrueType/TTC 字体。`FontResolver` 默认登记各系统常见字体目录（Windows Fonts、`%LOCALAPPDATA%\Microsoft\Windows\Fonts`、`/usr/share/fonts`、macOS 目录等），缺失目录自动跳过；**找不到请求字体时按码位回退**到能覆盖该字符的已注册字体（含 Noto、Liberation、文泉驿等常见命名）。

```csharp
FontResolver resolver = new FontResolver();

// 显式注册（文件路径或字节；TTC 可指定 faceIndex）
resolver.RegisterFont("Source Han Sans", false, false, "/usr/share/fonts/SourceHanSansSC-Regular.ttf");
resolver.RegisterFont("My Font", true, false, File.ReadAllBytes("bold.ttf"), faceIndex: 1);

PdfRenderer renderer = new PdfRenderer(resolver);
renderer.Render(document, "out.pdf");

// 查询与手工解析
TrueTypeFace face = resolver.Resolve("Microsoft YaHei", bold: false, italic: false);
TrueTypeFace fallback = resolver.ResolveFallback(0x4E2D, false, false);
bool isCjk = FontResolver.IsCjkCodePoint(0x4E2D);
```

`TextFormat` 把西文与中文族名分开设置，中英混排时各走各的字体，和 Word 篇里的写法一致：

```csharp
TextFormat body = new TextFormat("Times New Roman", 10.5f);
body.EastAsiaFontFamily = "宋体";
```

字体只写出实际用到的字形（**子集化**），所以文档体积跟用到的字符数相关，而不是字体文件本身有多大。

## 三、解析 PDF：带坐标的文本行与表格

读取入口是 `PdfDocument`，文本行与片段都带坐标和字号，表格识别基于文本位置与线条：

```csharp
using Wisdom.Pdf;

PdfDocument pdf = PdfDocument.Load("input.pdf");   // 亦支持 Stream / byte[]

Console.WriteLine(pdf.PageCount);
Console.WriteLine(pdf.Info.Title);
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

转成统一模型后即可交给 Word 侧输出，这也是 PDF → Word 这条链的中转点：

```csharp
OfficeDocument document = pdf.ToOfficeDocument();
```

## 四、对象级编辑：PdfFile

`PdfFile` 是另一条路径：**直接读写交叉表、按对象号增删改、合并拆分页面，不重建文档模型**。需要保留原版面、只做局部改动时用它，而不是走 `PdfRenderer` 重排。

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

`SaveIncremental` 只追加改动的对象、保留原始字节，这点在对"改动越小越好"的场景（比如加个书签、转一页方向）很实用。

## 五、口令加密读写

加密走标准安全处理器（`/Filter /Standard`）：

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

写出端默认用 AES-256（R6，`/CFM /AESV3`）；读取端覆盖 RC4-40（R2）、RC4-128（R3）、AES-128（R4）与 AES-256（R5/R6）。四种算法的产物都能被第三方阅读器口令开箱，第三方加密的文档也能被本库读回——这是双向验证过的。

## 六、水印：一个必须记住的坑

文字水印一行即可，默认斜跨 45°、浅灰、30% 不透明度、落在背景层，中文自动回退到中文字体：

```csharp
PdfFile file = PdfFile.Load("report.pdf");
file.AddWatermark("机密文件 CONFIDENTIAL");
file.Save("report-watermarked.pdf");
```

需要更多控制时用 `WatermarkOptions`，图层、摆放、角度、AutoFit、字体来源都能指定：

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

**这里有个务必注意的坑：`AddWatermark` 的第二个参数是图层而不是页码。** `AddWatermark("文字", 0)` 里的 `0` 会隐式转成 `WatermarkLayer.Background`，不是"第 0 页"。要给指定页加水印，请使用 `AddWatermark(options, 0)` 这个显式带页码的重载。

实现方式值得一提：全部页面共用同一份字体与图像资源（只嵌入一次），每页新增一个与 MediaBox 等大的 Form XObject 和一段 `q … /Wm Do … Q` 内容流；**背景层把这段插到 `/Contents` 前面，前景层追加到后面**。页面自带 `/Rotate` 时写出端做预旋转补偿，水印始终相对阅读方向摆放。带 `/StructParents` 的标签 PDF 会把水印标成 `/Artifact`，不破坏无障碍阅读顺序；不透明度小于 1 时自动把版本升到 1.4 并写入 `/ExtGState` 与透明 `/Group`。

## 七、表单填写与扁平化

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

批量填写时，层级字段用点号名，各类控件的取值语义也各不相同：

```csharp
Dictionary<string, string> values = new Dictionary<string, string>();
values["applicant.name"] = "王小明 Wang Xiaoming";  // 层级字段用 "父域.子域" 的点号名
values["applicant.agree"] = "true";                 // 复选框：true/状态名 勾选，false/Off/"" 取消
values["order.plan"] = "pro";                       // 单选组：状态名、选项名或 1 起的序号
values["applicant.city"] = "shenzhen";              // 组合框：/Opt 里的导出值
int filled = file.FillForm(values);                 // 填完统一生成外观，中文字体自动子集嵌入
file.Save("form-filled.pdf");
```

单字段用 `file.SetFormFieldValue("applicant.name", "李四")`，或直接 `field.Value = "李四"`；给签名域赋值会抛 `InvalidOperationException`。`FillForm(values, false)` 只写值不生成外观，此时 `/NeedAppearances` 置为 `true` 交给阅读器排版，调用 `GenerateFormAppearances()` 后自动复位——所以"填了值但外观还是旧的"这种中间态不会在阅读器里显示错值。

扁平化把值烧进页面内容，控件与字段一并移除：

```csharp
int flattened = file.FlattenForm();   // 内部先生成外观，再把 /AP /N 画进页面
file.Save("form-flat.pdf");           // 之后 HasForm 为 false，文档不再是可填表单
```

画不出外观的字段（密码域、没有 `/AP` 的字段、控件没挂进 `/Annots` 的字段）会原样保留，扁平化绝不会把它们的值丢掉；签名域与普通按钮同样保持交互态。混合表单（AcroForm + XFA）里阅读器优先渲染 XFA，填进去的 AcroForm 值可能看不见，用 `file.RemoveFormXfa()` 整体剥离。

实现要点：值写入字段 `/V`（文本域为 UTF-16 字符串，按钮与选项域为状态名）；单选组的值记在父域，各选项只同步自己的 `/AS`；外观流按控件矩形排版，`/DA` 字号为 0 时自动缩字号、多行域自动折行，勾选标记用矢量边框与勾/圆点画出、不依赖 ZapfDingbats；文字走本库既有的 Identity-H 子集嵌入与 ToUnicode 映射，**因此中文值在阅读器显示和文本提取里都可读**。

## 八、读取上限与安全

解析不受信任的 PDF 时，可以下调资源上限，超限会中止而不是无界占用内存（解压超限抛 `DecompressionLimitException`，结构非法抛 `InvalidDataException`）：

```csharp
PdfReadLimits.MaxDecodedStreamBytes = 64L * 1024 * 1024;  // 解压后字节上限，默认 256 MB
PdfReadLimits.MaxXrefEntries = 200000;                    // 交叉引用表条目上限
PdfReadLimits.MaxObjectStreamObjects = 100000;            // 对象流内对象上限
PdfReadLimits.MaxPageTreeDepth = 32;                      // 页面树递归深度上限
```

库内自行实现 inflate（`Zlib` / `Adler32`）与对象/词法解析，遇到结构损坏时抛 `InvalidDataException`，不会越界读取。

上限这条降级路径**不经过**文本提取的静默容错：一页正文可以由多个流引用、表单可以被反复绘制，单看每一份都合法、合起来是几十倍体积，所以 `ExtractText()` / `GetContentBytes()` 会在单条流超限、整页累计超限或各段合计超限时抛 `DecompressionLimitException`，不会返回一个空字符串冒充"这页没字"。批量索引不可信文件时按页包一层 catch。

## 已知边界

- 字体嵌入支持 TrueType `glyf` 轮廓；**CFF/OpenType（`OTTO`）字体不能作为嵌入源**（部分厂商以 `.otf` 后缀发布、实为 TrueType 轮廓的字体可正常使用）。
- PNG 不支持隔行扫描（interlaced）；图片支持 PNG 与 JPEG。
- `GetTables()` 是启发式识别（基于文本位置与线条），复杂版式可能需要人工校正。
- 文本提取依赖文本操作符，**扫描件需要额外 OCR，本包不含 OCR**；旋转文字按流向归并，但旋转页面上的表格识别仍以轴对齐坐标为准。
- 加密读写支持标准安全处理器（RC4-40 / RC4-128 / AES-128 / AES-256），写出统一 R6、读取兼容 R5；公钥加密与 Adobe 自定义安全处理器不在范围内。**数字签名（增量保存后校验原字节）尚不支持**；`PdfFile` 做的是对象级改写，签名文档与带对象流的文档在整量保存时会重建交叉表。
- 表单支持 AcroForm 的读取、填写、外观生成与扁平化；XFA 只能整体剥离、不解析模板；本包不生成新的表单字段，要给空白文档加控件请使用 `PdfFile` 的对象级 API。
- `OfficeDocument` → PDF 走 `PdfRenderer` 重新排版，不保留来源文档的字节级布局；需要保留原版面请用 `PdfFile`。
- **无桌面字体的精简容器环境需预先安装字体**（如 Noto、文泉驿），否则生成阶段无法解析字体。

## 小结

`Wisdom.Pdf` 的定位是自研 PDF 引擎的"三合一"：`PdfRenderer` 从统一模型排版生成、`PdfDocument` 解析回模型、`PdfFile` 在对象级直接改交叉表。生成侧把分页、跨页表格、表头重复、合并单元格跨度原子性拆分这些排版硬骨头啃了下来，编辑侧覆盖了旋转、增删页、增量保存、加解密、文字/图像水印和 AcroForm 填表扁平化。代价也很明确——放弃 CFF 嵌入、OCR、数字签名和 XFA 模板解析。

如果你的场景是"报表导出成 PDF""给已有 PDF 加水印/加密/填表""把 PDF 抽成模型再转 Word"这类结构化处理，而不是"像素级还原复杂原版面"或"处理扫描图像"，这个包够用，而且没有第三方 PDF 库的授权和体积负担。只处理 Word/Excel 想顺带出 PDF 的话，去看同系列的 `Wisdom.Office` 门面包。
