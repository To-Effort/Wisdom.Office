using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Ppt
{
    /// <summary>
    /// 将统一演示文稿模型（<see cref="OfficePresentation"/>）写出为 PresentationML（.pptx）的 OPC 包。
    /// Writes the unified presentation model (<see cref="OfficePresentation"/>) as a PresentationML (.pptx) OPC package.
    /// </summary>
    public static class PptxWriter
    {
        /// <summary>将演示文稿写入指定路径的文件（覆盖已存在文件）。 Writes the presentation to a file at the given path, overwriting an existing file.</summary>
        /// <param name="presentation">要写出的演示文稿模型。 Presentation model to write.</param>
        /// <param name="path">输出 .pptx 文件路径。 Output .pptx file path.</param>
        public static void Write(OfficePresentation presentation, string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            // 先写同目录临时文件再原子替换：中途失败不会把用户原文件截成 0 字节。
            // Write a temporary file beside the target and swap it in atomically, so a mid-write failure
            // never leaves the user's original file truncated to zero bytes.
            AtomicFileWriter.Write(path, delegate(Stream stream) { Write(presentation, stream); });
        }

        /// <summary>将演示文稿写出的字节数组形式（内存 .pptx 包）。 Writes the presentation as a byte array (an in-memory .pptx package).</summary>
        /// <param name="presentation">要写出的演示文稿模型。 Presentation model to write.</param>
        /// <returns>.pptx 包的完整字节。 The complete bytes of the .pptx package.</returns>
        public static byte[] Write(OfficePresentation presentation)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Write(presentation, stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// 将演示文稿写入流：单趟流式写出——逐张幻灯片构建后即序列化进复用缓冲并写入 zip 条目，
        /// 媒体在登记时直接落位，母版、版式、主题、presentation 与内容类型在尾部补齐。
        /// Writes the presentation to a stream in a single streaming pass: each slide is serialized into a
        /// reused scratch buffer and written to the zip right after it is built, media is stored as it is
        /// registered, and the master, layouts, theme, presentation and content types are emitted at the end.
        /// </summary>
        /// <param name="presentation">要写出的演示文稿模型。 Presentation model to write.</param>
        /// <param name="stream">目标流；调用方负责关闭。 Target stream; the caller is responsible for closing it.</param>
        public static void Write(OfficePresentation presentation, Stream stream)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            if (stream == null) throw new ArgumentNullException("stream");
            if (!stream.CanWrite) throw new ArgumentException("Stream is not writable.", "stream");
            WriterState state = new WriterState(presentation, false);
            state.WritePackage(stream);
        }

        /// <summary>将演示文稿写出为 .potx 模板文件：包结构与 .pptx 相同，仅 presentation 部件的
        /// 内容类型换为 template。 Writes the presentation as a .potx template: the package layout is identical to
        /// .pptx, only the presentation part's content type becomes the template one.</summary>
        /// <param name="presentation">要写出的演示文稿模型。 Presentation model to write.</param>
        /// <param name="path">输出 .potx 文件路径。 Output .potx file path.</param>
        public static void WriteTemplate(OfficePresentation presentation, string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            AtomicFileWriter.Write(path, delegate(Stream stream) { WriteTemplate(presentation, stream); });
        }

        /// <summary>将演示文稿写出为 .potx 模板的字节数组。 Writes the presentation as .potx template bytes.</summary>
        /// <param name="presentation">要写出的演示文稿模型。 Presentation model to write.</param>
        /// <returns>.potx 包的完整字节。 The complete bytes of the .potx package.</returns>
        public static byte[] WriteTemplate(OfficePresentation presentation)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                WriteTemplate(presentation, stream);
                return stream.ToArray();
            }
        }

        /// <summary>将演示文稿写出为 .potx 模板包并写入流。 Writes the presentation as a .potx template package into a stream.</summary>
        /// <param name="presentation">要写出的演示文稿模型。 Presentation model to write.</param>
        /// <param name="stream">目标流；调用方负责关闭。 Target stream; the caller is responsible for closing it.</param>
        public static void WriteTemplate(OfficePresentation presentation, Stream stream)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            if (stream == null) throw new ArgumentNullException("stream");
            if (!stream.CanWrite) throw new ArgumentException("Stream is not writable.", "stream");
            WriterState state = new WriterState(presentation, true);
            state.WritePackage(stream);
        }

        /// <summary>
        /// 一次写出过程的内部状态：集中生成母版、版式、主题、幻灯片、备注与各关系部件，
        /// 并对媒体图片去重登记，最终按 OPC 结构写入 zip 归档。
        /// Internal state of one write operation: generates the master, layouts, theme, slides, notes and every
        /// relationship part, de-duplicating image media, and finally writes all parts into the zip archive.
        /// </summary>
        private sealed class WriterState
        {
            /// <summary>演示部件主命名空间前缀 p（等价 <see cref="Ppt.Main"/>）。 Presentation main namespace prefix p (alias of <see cref="Ppt.Main"/>).</summary>
            private static readonly XNamespace P = Ppt.Main;
            /// <summary>图形命名空间前缀 a（等价 <see cref="Ppt.Drawing"/>）。 Drawing namespace prefix a (alias of <see cref="Ppt.Drawing"/>).</summary>
            private static readonly XNamespace A = Ppt.Drawing;
            /// <summary>关系命名空间前缀 r（等价 <see cref="Ppt.R"/>）。 Relationship namespace prefix r (alias of <see cref="Ppt.R"/>).</summary>
            private static readonly XNamespace R = Ppt.R;
            /// <summary>OPC 包关系命名空间。 OPC package relationships namespace.</summary>
            private static readonly XNamespace Rels = Ppt.PackageRels;
            /// <summary>OPC 内容类型命名空间。 OPC content types namespace.</summary>
            private static readonly XNamespace Ctypes = Ppt.ContentTypes;
            /// <summary>XML 实例命名空间，用于 <c>xsi:type</c>。 XML-schema-instance namespace, used for <c>xsi:type</c>.</summary>
            private static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";
            /// <summary>标记兼容性命名空间前缀 mc（等价 <see cref="Ppt.MarkupCompat"/>）。 Markup-compatibility namespace prefix mc (alias of <see cref="Ppt.MarkupCompat"/>).</summary>
            private static readonly XNamespace Mc = Ppt.MarkupCompat;
            /// <summary>PowerPoint 2010 扩展命名空间前缀 p14（等价 <see cref="Ppt.PowerPoint2010"/>）。 PowerPoint 2010 extension namespace prefix p14 (alias of <see cref="Ppt.PowerPoint2010"/>).</summary>
            private static readonly XNamespace P14 = Ppt.PowerPoint2010;

            /// <summary>关系类型的完整命名空间前缀。 Full namespace prefix of officeDocument relationship types.</summary>
            private const string RelBase = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            /// <summary>officeDocument 主文档关系类型。 officeDocument relationship type.</summary>
            private const string RelOfficeDocument = RelBase + "/officeDocument";
            /// <summary>幻灯片关系类型。 Slide relationship type.</summary>
            private const string RelSlide = RelBase + "/slide";
            /// <summary>幻灯片版式关系类型。 Slide layout relationship type.</summary>
            private const string RelSlideLayout = RelBase + "/slideLayout";
            /// <summary>幻灯片母版关系类型。 Slide master relationship type.</summary>
            private const string RelSlideMaster = RelBase + "/slideMaster";
            /// <summary>主题关系类型。 Theme relationship type.</summary>
            private const string RelTheme = RelBase + "/theme";
            /// <summary>图片关系类型。 Image relationship type.</summary>
            private const string RelImage = RelBase + "/image";
            /// <summary>备注母版关系类型。 Notes master relationship type.</summary>
            private const string RelNotesMaster = RelBase + "/notesMaster";
            /// <summary>备注页关系类型。 Notes slide relationship type.</summary>
            private const string RelNotesSlide = RelBase + "/notesSlide";
            /// <summary>核心属性关系类型。 Core properties relationship type.</summary>
            private const string RelCoreProperties = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
            /// <summary>扩展属性关系类型。 Extended properties relationship type.</summary>
            private const string RelExtendedProperties = RelBase + "/extended-properties";

            /// <summary>PresentationML 部件内容类型前缀。 PresentationML part content-type prefix.</summary>
            private const string Ctp = "application/vnd.openxmlformats-officedocument.presentationml.";
            /// <summary>rels 部件内容类型。 Content type of rels parts.</summary>
            private const string CtRels = "application/vnd.openxmlformats-package.relationships+xml";
            /// <summary>普通 XML 部件内容类型。 Default content type for plain XML parts.</summary>
            private const string CtXml = "application/xml";
            /// <summary>presentation 部件内容类型。 Content type of the presentation part.</summary>
            private const string CtPresentation = Ctp + "presentation.main+xml";
            /// <summary>presentation 部件在 .potx 模板中的内容类型。 Presentation-part content type used by a .potx template.</summary>
            private const string CtPresentationTemplate = Ctp + "template.main+xml";
            /// <summary>slideMaster 部件内容类型。 Content type of the slide master part.</summary>
            private const string CtSlideMaster = Ctp + "slideMaster+xml";
            /// <summary>slideLayout 部件内容类型。 Content type of slide layout parts.</summary>
            private const string CtSlideLayout = Ctp + "slideLayout+xml";
            /// <summary>slide 部件内容类型。 Content type of slide parts.</summary>
            private const string CtSlide = Ctp + "slide+xml";
            /// <summary>notesSlide 部件内容类型。 Content type of notes slide parts.</summary>
            private const string CtNotesSlide = Ctp + "notesSlide+xml";
            /// <summary>notesMaster 部件内容类型。 Content type of the notes master part.</summary>
            private const string CtNotesMaster = Ctp + "notesMaster+xml";
            /// <summary>theme 部件内容类型。 Content type of the theme part.</summary>
            private const string CtTheme = "application/vnd.openxmlformats-officedocument.theme+xml";
            /// <summary>core 部件内容类型。 Content type of the core properties part.</summary>
            private const string CtCore = "application/vnd.openxmlformats-package.core-properties+xml";
            /// <summary>app 部件内容类型。 Content type of the extended properties part.</summary>
            private const string CtApp = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

            /// <summary>图片内容类型：PNG。 Image content type for PNG.</summary>
            private const string CtPng = "image/png";
            /// <summary>图片内容类型：JPEG。 Image content type for JPEG.</summary>
            private const string CtJpeg = "image/jpeg";
            /// <summary>图片内容类型：GIF。 Image content type for GIF.</summary>
            private const string CtGif = "image/gif";

            /// <summary>DrawingML 表格图形数据类型 URI。 DrawingML table graphic-data URI.</summary>
            private const string TableGraphicUri = "http://schemas.openxmlformats.org/drawingml/2006/table";
            /// <summary>内置表格 Medium Style 2 的样式标识。 Built-in Medium Style 2 table style id.</summary>
            private const string DefaultTableStyleId = "{5C22544A-7EE3-442F-B6AB-1F4F63D42E33}";

            /// <summary>待写出的演示文稿模型。 Presentation model being written.</summary>
            private readonly OfficePresentation _p;
            /// <summary>幻灯片宽度（磅）。 Slide width in points.</summary>
            private readonly double _w;
            /// <summary>幻灯片高度（磅）。 Slide height in points.</summary>
            private readonly double _h;
            /// <summary>正文默认字体格式（回退 Calibri 18）。 Default body format (falls back to Calibri 18).</summary>
            private readonly TextFormat _bodyFormat;
            /// <summary>标题默认字体格式（回退 Calibri 40 粗体）。 Default title format (falls back to bold Calibri 40).</summary>
            private readonly TextFormat _titleFormat;
            /// <summary>是否按 .potx 模板写出（仅 presentation 内容类型不同）。 Whether the package is written as a .potx template (only the presentation content type differs).</summary>
            private readonly bool _template;

            /// <summary>按首次出现顺序排列、实际用到的版式类别。 Used layout kinds in first-appearance order.</summary>
            private readonly List<SlideLayoutKind> _layoutKinds = new List<SlideLayoutKind>();
            /// <summary>版式类别到 1 基版式编号的映射。 Map from layout kind to its 1-based layout number.</summary>
            private readonly Dictionary<SlideLayoutKind, int> _layoutNumber = new Dictionary<SlideLayoutKind, int>();

            /// <summary>媒体去重表：同一图片在整包中只写一份。 Image de-duplication table: each image produces one media file per package.</summary>
            private readonly Dictionary<OfficeImage, MediaInfo> _media = new Dictionary<OfficeImage, MediaInfo>();
            /// <summary>已登记的媒体数量（用于编号）。 Count of registered media, used for numbering.</summary>
            private int _mediaCount;
            /// <summary>未在 Default 中覆盖的媒体内容类型（扩展名到内容类型）。 Media content types not covered by a Default (extension to content type).</summary>
            private readonly Dictionary<string, string> _extraMediaTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            /// <summary>保存过程中的目标归档。 Archive being written during the save pass.</summary>
            private ZipArchive _archive;
            /// <summary>归档中的幻灯片总数（预扫描得出）。 Total slide count in the archive, from the metadata pre-scan.</summary>
            private int _slideCount;
            /// <summary>归档中的备注页总数（预扫描得出）。 Total notes-slide count in the archive, from the metadata pre-pass.</summary>
            private int _notesCount;
            /// <summary>流式写出过程中已写出的备注页数。 Notes-slide parts written so far during the streaming pass.</summary>
            private int _notesWritten;

            /// <summary>基于待写出的演示文稿创建写出状态。 Creates the write state for a presentation model.</summary>
            /// <param name="presentation">演示文稿模型。 Presentation model.</param>
            /// <param name="template">是否按 .potx 模板写出。 Whether to write the package as a .potx template.</param>
            public WriterState(OfficePresentation presentation, bool template)
            {
                _p = presentation;
                _template = template;
                _w = presentation.SlideWidth > 0 ? presentation.SlideWidth : 960.0;
                _h = presentation.SlideHeight > 0 ? presentation.SlideHeight : 540.0;
                // 形状文本的默认字体优先取模型显式设定，其次取主题深改的主/次字体，最后才是 Calibri；
                // 每个 run 都写出显式 a:latin，不改这一步的话主题字体只存在于 theme1.xml，看不见效果。
                // Default run fonts prefer the model's explicit setting, then the deep-edited theme fonts,
                // then Calibri; every run writes an explicit a:latin, so without this the theme typefaces
                // would live only in theme1.xml and never show up in the text.
                _bodyFormat = presentation.GetEffectiveTextFormat() ?? new TextFormat("Calibri", 18.0);
                _titleFormat = presentation.GetEffectiveTitleFormat() ?? new TextFormat("Calibri", 40.0) { Bold = true };
            }

            /// <summary>
            /// 单趟写出：先轻量预扫描（版式、页数、备注数、媒体扩展名）以便把 [Content_Types].xml
            /// 按 OPC 规范放在包首，再逐张幻灯片构建后立即序列化写入归档条目，媒体登记时直接落位，
            /// 版式、母版、主题与 presentation 在尾部补齐。
            /// A single streaming pass: a light metadata pre-scan (layouts, counts, media extensions) first
            /// so [Content_Types].xml can be the first entry as the OPC spec demands, then each slide is
            /// serialized into the archive right after it is built, media is stored as it registers, and
            /// the layouts, master, theme and presentation parts are emitted at the end.
            /// </summary>
            /// <param name="stream">目标流。 Target stream.</param>
            public void WritePackage(Stream stream)
            {
                List<Slide> slides = _p.Slides ?? new List<Slide>();
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
                {
                    _archive = archive;
                    ScanMetadata(slides);
                    WriteEntryXml("[Content_Types].xml", CreateContentTypes());
                    WriteEntryXml("_rels/.rels", CreateRootRelationships());
                    WriteEntryXml("docProps/core.xml", CreateCoreProperties());
                    WriteEntryXml("docProps/app.xml", CreateAppProperties());
                    for (int i = 0; i < slides.Count; i++)
                    {
                        Slide slide = slides[i] ?? new Slide();
                        int layoutNumber = EnsureLayout(slide.Layout);
                        BuildSlide(slide, i + 1, layoutNumber);
                    }

                    WriteLayouts();
                    WriteEntryXml("ppt/theme/theme1.xml", CreateTheme());
                    WriteEntryXml("ppt/slideMasters/slideMaster1.xml", CreateMaster());
                    WriteEntryXml("ppt/slideMasters/_rels/slideMaster1.xml.rels", CreateMasterRelationships());
                    WriteEntryXml("ppt/notesMasters/notesMaster1.xml", CreateNotesMaster());
                    WriteEntryXml("ppt/notesMasters/_rels/notesMaster1.xml.rels", CreateNotesMasterRelationships());
                    WritePresentation(slides.Count);
                    _archive = null;
                }
            }

            /// <summary>预扫描模型：登记版式、统计页数与备注数、收集媒体扩展名，全部供首部内容类型使用。
            /// Pre-scans the model: registers layouts, counts slides and notes, and collects media
            /// extensions, everything the leading content-types part needs.</summary>
            /// <param name="slides">幻灯片列表。 The slides.</param>
            private void ScanMetadata(List<Slide> slides)
            {
                _slideCount = slides.Count;
                for (int i = 0; i < slides.Count; i++)
                {
                    Slide slide = slides[i] ?? new Slide();
                    EnsureLayout(slide.Layout);
                    if (!string.IsNullOrEmpty(slide.NotesText)) _notesCount++;
                    List<SlideShape> shapes = slide.Shapes ?? new List<SlideShape>();
                    for (int s = 0; s < shapes.Count; s++)
                    {
                        PictureShape picture = shapes[s] as PictureShape;
                        if (picture == null || picture.Image == null) continue;
                        OfficeImage image = picture.Image;
                        string extension = string.IsNullOrEmpty(image.FileExtension) ? ".png" : image.FileExtension.ToLowerInvariant();
                        if (!IsCoveredDefault(extension))
                        {
                            _extraMediaTypes[extension.TrimStart('.')] =
                                string.IsNullOrEmpty(image.ContentType) ? CtPng : image.ContentType;
                        }
                    }
                }
            }

            /// <summary>保存阶段跨部件复用的序列化缓冲区。 Serialization buffer reused across parts during the save phase.</summary>
            private MemoryStream _scratch;

            /// <summary>把 XML 部件序列化进复用缓冲后一次性写入归档条目（最优压缩）：既免掉逐部件的
            /// byte[] 中转拷贝，又避免 XmlWriter 小步直写压缩流。
            /// Serializes an XML document into the shared scratch buffer and writes it to the entry in one
            /// shot (optimal compression): no per-part byte round trip and no trickle writes into the deflate stream.</summary>
            /// <param name="name">部件路径名。 Part path name.</param>
            /// <param name="document">部件文档。 Part document.</param>
            private void WriteEntryXml(string name, XDocument document)
            {
                XmlWriterSettings settings = new XmlWriterSettings();
                settings.Encoding = new UTF8Encoding(false);
                settings.Indent = false;
                settings.OmitXmlDeclaration = document.Declaration == null;
                // 与 SerializeXml 同一道防护：控制字符会让 XmlWriter 抛异常，整个包写坏。
                // The same guard as SerializeXml: a control character makes XmlWriter throw and wrecks the package.
                OfficeXmlText.Sanitize(document);
                if (_scratch == null) _scratch = new MemoryStream(65536);
                _scratch.SetLength(0);
                using (XmlWriter writer = XmlWriter.Create(_scratch, settings))
                {
                    document.Save(writer);
                }
                ZipArchiveEntry entry = _archive.CreateEntry(OfficePackageLimits.ValidateEntryName(name), CompressionLevel.Optimal);
                using (Stream stream = entry.Open())
                {
                    stream.Write(_scratch.GetBuffer(), 0, (int)_scratch.Length);
                }
            }

            /// <summary>向归档写入一个原始字节部件。 Writes one raw-byte part into the archive.</summary>
            /// <param name="name">部件路径名。 Part path name.</param>
            /// <param name="data">部件字节内容。 Part byte content.</param>
            /// <param name="level">压缩级别。 Compression level.</param>
            private void WriteEntryBytes(string name, byte[] data, CompressionLevel level)
            {
                ZipArchiveEntry entry = _archive.CreateEntry(OfficePackageLimits.ValidateEntryName(name), level);
                using (Stream stream = entry.Open())
                {
                    stream.Write(data, 0, data.Length);
                }
            }

            /// <summary>登记一个版式类别（首次出现时分配编号）。 Registers a layout kind, allocating a number on first use.</summary>
            /// <param name="kind">版式类别。 Layout kind.</param>
            /// <returns>该版式的 1 基编号。 The 1-based layout number.</returns>
            private int EnsureLayout(SlideLayoutKind kind)
            {
                int number;
                if (_layoutNumber.TryGetValue(kind, out number)) return number;
                number = _layoutKinds.Count + 1;
                _layoutKinds.Add(kind);
                _layoutNumber[kind] = number;
                return number;
            }

            /// <summary>为每个已登记的版式类别生成并写出 slideLayout 部件与其指向母版的关系。 Generates and writes the slideLayout part and its master relationship for every registered layout kind.</summary>
            private void WriteLayouts()
            {
                for (int i = 0; i < _layoutKinds.Count; i++)
                {
                    SlideLayoutKind kind = _layoutKinds[i];
                    int number = i + 1;
                    XElement layout = Declare(RelationshipRoot(P + "sldLayout",
                        new XElement(P + "cSld",
                            new XAttribute("name", LayoutName(kind)),
                            CreatePlaceholderTree(kind, false))), true);
                    // 版式类别写进 p:sldLayout@type：章节过渡页只有这一处证据能区别于「仅标题页」。
                    // The layout kind goes onto p:sldLayout@type: that is the only evidence that tells a
                    // section header apart from a title-only slide.
                    layout.SetAttributeValue("type", LayoutTypeCode(kind));
                    layout.Add(new XElement(P + "clrMapOvr", new XElement(A + "masterClrMapping")));
                    layout.Add(new XElement(P + "timing"));
                    XDocument rels = BuildRels(new List<RelEntry>
                    {
                        new RelEntry("rId1", RelSlideMaster, "../slideMasters/slideMaster1.xml")
                    });
                    WriteEntryXml("ppt/slideLayouts/slideLayout" + IntStr(number) + ".xml", WithDeclaration(layout));
                    WriteEntryXml("ppt/slideLayouts/_rels/slideLayout" + IntStr(number) + ".xml.rels", rels);
                }
            }

            /// <summary>生成并写出 presentation.xml 与其关系。 Builds and writes presentation.xml and its relationships.</summary>
            /// <param name="slideCount">幻灯片总数。 Total number of slides.</param>
            private void WritePresentation(int slideCount)
            {
                List<RelEntry> presentationRels = new List<RelEntry>();
                presentationRels.Add(new RelEntry("rId1", RelSlideMaster, "slideMasters/slideMaster1.xml"));
                XElement sldIdLst = new XElement(P + "sldIdLst");
                int nextId = 256;
                for (int i = 0; i < slideCount; i++)
                {
                    string rid = "rId" + IntStr(i + 2);
                    presentationRels.Add(new RelEntry(rid, RelSlide, "slides/slide" + IntStr(i + 1) + ".xml"));
                    sldIdLst.Add(new XElement(P + "sldId",
                        new XAttribute("id", IntStr(nextId)),
                        new XAttribute(R + "id", rid)));
                    nextId++;
                }
                if (slideCount == 0)
                {
                    sldIdLst = new XElement(P + "sldIdLst");
                }
                string notesMasterRid = "rId" + IntStr(slideCount + 2);
                presentationRels.Add(new RelEntry(notesMasterRid, RelNotesMaster, "notesMasters/notesMaster1.xml"));

                XElement presentation = Declare(RelationshipRoot(P + "presentation",
                    new XElement(P + "sldMasterIdLst",
                        new XElement(P + "sldMasterId",
                            new XAttribute("id", "2147483648"),
                            new XAttribute(R + "id", "rId1"))),
                    new XElement(P + "notesMasterIdLst",
                        new XElement(P + "notesMasterId", new XAttribute(R + "id", notesMasterRid))),
                    sldIdLst,
                    new XElement(P + "sldSz",
                        new XAttribute("cx", IntStr(Ppt.Emu(_w))),
                        new XAttribute("cy", IntStr(Ppt.Emu(_h)))),
                    new XElement(P + "notesSz",
                        new XAttribute("cx", "6858000"),
                        new XAttribute("cy", "9144000"))), true);
                WriteEntryXml("ppt/presentation.xml", WithDeclaration(presentation));
                WriteEntryXml("ppt/_rels/presentation.xml.rels", BuildRels(presentationRels));
            }

            /// <summary>构建并写出单张幻灯片部件及其关系集合（版式、图片、备注）。 Builds and writes one slide part and its relationship collection (layout, images, notes).</summary>
            /// <param name="slide">幻灯片模型。 Slide model.</param>
            /// <param name="number">幻灯片 1 基编号。 1-based slide number.</param>
            /// <param name="layoutNumber">所用版式编号。 Used layout number.</param>
            private void BuildSlide(Slide slide, int number, int layoutNumber)
            {
                List<RelEntry> rels = new List<RelEntry>();
                rels.Add(new RelEntry("rId1", RelSlideLayout, "../slideLayouts/slideLayout" + IntStr(layoutNumber) + ".xml"));
                SlidePart part = new SlidePart(number, slide, rels);

                XElement cSld = new XElement(P + "cSld");
                XElement bg = CreateBackground(slide);
                if (bg != null) cSld.Add(bg);
                XElement spTree = CreateSpTree();
                List<SlideShape> shapes = slide.Shapes ?? new List<SlideShape>();
                List<int> animationIds = new List<int>();
                List<SlideShape> animationShapes = new List<SlideShape>();
                for (int i = 0; i < shapes.Count; i++)
                {
                    if (IsEmptyPlaceholder(shapes[i])) continue;
                    XElement element = CreateShape(shapes[i], part, part.NextShapeId++);
                    if (element == null) continue;
                    spTree.Add(element);
                    if (shapes[i].Animation != SlideAnimationKind.None)
                    {
                        animationIds.Add(part.NextShapeId - 1);
                        animationShapes.Add(shapes[i]);
                    }
                }
                cSld.Add(spTree);

                XElement root = RelationshipRoot(P + "sld", cSld);
                if (!string.IsNullOrEmpty(slide.NotesText))
                {
                    int notesNumber = _notesWritten + 1;
                    _notesWritten++;
                    WriteEntryXml("ppt/notesSlides/notesSlide" + IntStr(notesNumber) + ".xml",
                        WithDeclaration(CreateNotesSlide(slide)));
                    List<RelEntry> notesRels = new List<RelEntry>();
                    notesRels.Add(new RelEntry("rId1", RelNotesMaster, "../notesMasters/notesMaster1.xml"));
                    notesRels.Add(new RelEntry("rId2", RelSlide, "../slides/slide" + IntStr(number) + ".xml"));
                    WriteEntryXml("ppt/notesSlides/_rels/notesSlide" + IntStr(notesNumber) + ".xml.rels", BuildRels(notesRels));
                    part.AddRel(RelNotesSlide, "../notesSlides/notesSlide" + IntStr(notesNumber) + ".xml");
                }
                XElement transition = CreateTransition(slide);
                if (transition != null) root.Add(transition);
                XElement timing = CreateTiming(animationIds, animationShapes, slide);
                if (timing != null) root.Add(timing);
                WriteEntryXml("ppt/slides/slide" + IntStr(number) + ".xml", WithDeclaration(Declare(root, true)));
                WriteEntryXml("ppt/slides/_rels/slide" + IntStr(number) + ".xml.rels", BuildRels(part.Rels));
                _slideCount++;
            }

            /// <summary>判断形状是否可以按「程序生成的稿子不该留一个等用户点进去的空框」省略掉。
            /// 只有本库自己补出来、且至今没有任何可见效果的占位符才算：文本为空、没有底色、没有描边、
            /// 没有入场动画。从文件读回的或调用方自己添加的占位符不在其列——它们即便没有文本也是
            /// 页面上的一个形状，省掉就是静默丢失。
            /// Reports whether a shape may be dropped under the rule that a generated deck should not leave
            /// a "click to add" box behind. Only a placeholder this library generated and that still has no
            /// visible effect qualifies: no text, no fill, no outline, no entrance animation. One read back
            /// from a file or added by the caller never does — even empty, it is a shape on the page, and
            /// dropping it silently loses content.</summary>
            /// <param name="shape">待判断的形状。 The shape to test.</param>
            /// <returns>是可省略的空占位符时为 true。 True for a droppable text-less generated placeholder.</returns>
            private static bool IsEmptyPlaceholder(SlideShape shape)
            {
                if (shape == null || shape.Placeholder == PlaceholderKind.None) return false;
                if (!shape.IsGeneratedPlaceholder) return false;
                if (shape.FillColor != null || shape.Outline != null) return false;
                if (shape.Animation != SlideAnimationKind.None) return false;
                TextBoxShape box = shape as TextBoxShape;
                if (box == null) return false;
                string text = box.GetText();
                return string.IsNullOrEmpty(text) || text.Trim().Length == 0;
            }

            /// <summary>生成单张备注页的 <code>p:notes</code> 根元素（把备注文本按换行拆成段落）。 Builds one notes slide's <code>p:notes</code> root, splitting the notes text into paragraphs by line breaks.</summary>
            /// <param name="slide">幻灯片。 Slide.</param>
            /// <returns>声明了命名空间的备注根元素。 The namespace-declared notes root.</returns>
            private XElement CreateNotesSlide(Slide slide)
            {
                List<Paragraph> paragraphs = new List<Paragraph>();
                string[] lines = (slide.NotesText ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
                for (int i = 0; i < lines.Length; i++)
                {
                    paragraphs.Add(new Paragraph(lines[i]));
                }
                XElement bodyTx = new XElement(P + "txBody",
                    new XElement(A + "bodyPr", new XAttribute("anchor", "t")),
                    new XElement(A + "lstStyle"));
                AppendParagraphs(bodyTx, paragraphs, _bodyFormat);
                XElement sp = new XElement(P + "sp",
                    new XElement(P + "nvSpPr",
                        new XElement(P + "cNvPr", new XAttribute("id", "2"), new XAttribute("name", "备注占位符 1")),
                        new XElement(P + "cNvSpPr", new XElement(A + "spLocks", new XAttribute("noGrp", "1"))),
                        new XElement(P + "nvPr", new XElement(P + "ph", new XAttribute("type", "body"), new XAttribute("idx", "1")))),
                    new XElement(P + "spPr",
                        new XElement(A + "xfrm",
                            new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "4572000")),
                            new XElement(A + "ext", new XAttribute("cx", "6858000"), new XAttribute("cy", "4572000")))),
                    bodyTx);
                XElement spTree = new XElement(P + "spTree",
                    new XElement(P + "nvGrpSpPr",
                        new XElement(P + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", string.Empty)),
                        new XElement(P + "cNvGrpSpPr"),
                        new XElement(P + "nvPr")),
                    new XElement(P + "grpSpPr",
                        new XElement(A + "xfrm",
                            new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                            new XElement(A + "ext", new XAttribute("cx", "6858000"), new XAttribute("cy", "9144000")),
                            new XElement(A + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")),
                            new XElement(A + "chExt", new XAttribute("cx", "6858000"), new XAttribute("cy", "9144000")))),
                    sp);
                XElement root = RelationshipRoot(P + "notes",
                    new XElement(P + "cSld", spTree),
                    new XElement(P + "clrMapOvr", new XElement(A + "masterClrMapping")));
                return Declare(root, true);
            }

            /// <summary>生成 <code>p:bg</code> 背景元素；无背景色时返回 null。 Builds the <code>p:bg</code> background element, or null when there is no background color.</summary>
            /// <param name="slide">幻灯片。 Slide.</param>
            /// <returns>背景元素或 null。 The background element or null.</returns>
            private static XElement CreateBackground(Slide slide)
            {
                if (!slide.BackgroundColor.HasValue) return null;
                OfficeColor color = slide.BackgroundColor.Value;
                XElement bg = new XElement(P + "bg",
                    new XElement(P + "bgPr",
                        new XElement(A + "solidFill",
                            new XElement(A + "srgbClr", new XAttribute("val", color.ToHex()))),
                        new XElement(A + "effectLst")));
                return bg;
            }

            /// <summary>创建 <code>p:spTree</code> 并预置组级 <code>nvGrpSpPr</code> 与 <code>grpSpPr</code>。 Creates the <code>p:spTree</code> with the group <code>nvGrpSpPr</code> and <code>grpSpPr</code> prefilled.</summary>
            /// <returns>形状树元素。 The shape-tree element.</returns>
            private XElement CreateSpTree()
            {
                return new XElement(P + "spTree",
                    new XElement(P + "nvGrpSpPr",
                        new XElement(P + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", string.Empty)),
                        new XElement(P + "cNvGrpSpPr"),
                        new XElement(P + "nvPr")),
                    new XElement(P + "grpSpPr",
                        new XElement(A + "xfrm",
                            new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                            new XElement(A + "ext", new XAttribute("cx", IntStr(Ppt.Emu(_w))), new XAttribute("cy", IntStr(Ppt.Emu(_h)))),
                            new XElement(A + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")),
                            new XElement(A + "chExt", new XAttribute("cx", IntStr(Ppt.Emu(_w))), new XAttribute("cy", IntStr(Ppt.Emu(_h)))))));
            }

            /// <summary>按形状类型分派生成对应的 XML 元素（p:sp、p:pic 或 p:graphicFrame）。 Dispatches a shape to its XML element (p:sp, p:pic or p:graphicFrame).</summary>
            /// <param name="shape">形状模型。 Shape model.</param>
            /// <param name="part">所属幻灯片条目，用于登记图片关系。 Owning slide entry used to register image relationships.</param>
            /// <param name="id">该形状在本张幻灯片内的唯一 <c>cNvPr id</c>。 The shape's unique cNvPr id within this slide.</param>
            /// <returns>形状元素；无法处理时返回 null。 The shape element, or null when it cannot be handled.</returns>
            private XElement CreateShape(SlideShape shape, SlidePart part, int id)
            {
                if (shape == null) return null;
                PictureShape picture = shape as PictureShape;
                if (picture != null) return CreatePicture(picture, part, id);
                TableShape table = shape as TableShape;
                if (table != null) return CreateTable(table, id);
                TextBoxShape text = shape as TextBoxShape;
                if (text != null) return CreateAutoOrTextBox(text, id);
                return CreateEmptyShape(shape, id);
            }

            /// <summary>生成文本框或自选图形的 <code>p:sp</code> 元素。 Builds the <code>p:sp</code> element of a text box or auto shape.</summary>
            /// <param name="shape">文本形状。 Text-bearing shape.</param>
            /// <param name="id">形状唯一 <c>cNvPr id</c>。 The unique cNvPr id.</param>
            /// <returns>p:sp 元素。 The p:sp element.</returns>
            private XElement CreateAutoOrTextBox(TextBoxShape shape, int id)
            {
                AutoShape auto = shape as AutoShape;
                XElement sp = new XElement(P + "sp");
                sp.Add(CreateShapeNonVisual(shape, auto == null, id));
                sp.Add(CreateShapeProperties(shape, auto));
                sp.Add(CreateTextBody(shape, ResolveBaseFormat(shape)));
                return sp;
            }

            /// <summary>为未知形状生成一个最小 <code>p:sp</code>，保证关系与 z 序不丢失。 Builds a minimal <code>p:sp</code> for an unrecognized shape so the relationship and z-order are preserved.</summary>
            /// <param name="shape">形状模型。 Shape model.</param>
            /// <param name="id">形状唯一 <c>cNvPr id</c>。 The unique cNvPr id.</param>
            /// <returns>p:sp 元素。 The p:sp element.</returns>
            private XElement CreateEmptyShape(SlideShape shape, int id)
            {
                XElement nvPr = new XElement(P + "nvPr");
                XElement ph = CreatePlaceholder(shape);
                if (ph != null) nvPr.Add(ph);
                return new XElement(P + "sp",
                    new XElement(P + "nvSpPr",
                        CreateCnvPr(shape, null, id),
                        new XElement(P + "cNvSpPr"),
                        nvPr),
                    new XElement(P + "spPr",
                        CreateXfrm(shape),
                        new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst")),
                        CreateOptionalFill(shape.FillColor),
                        CreateOptionalLine(shape.Outline)),
                    new XElement(P + "txBody",
                        new XElement(A + "bodyPr"),
                        new XElement(A + "lstStyle"),
                        new XElement(A + "p", new XElement(A + "endParaRPr", new XAttribute("lang", "zh-CN")))));
            }

            /// <summary>生成 <code>p:nvSpPr</code>（含可选占位符 <code>p:ph</code>）。 Builds <code>p:nvSpPr</code> with an optional placeholder <code>p:ph</code>.</summary>
            /// <param name="shape">形状模型。 Shape model.</param>
            /// <param name="isTextBox">是否为纯文本框（自选图形为 false；仅文本框才写 txBox="1"）。 Whether it is a plain text box (false for auto shapes; only text boxes carry txBox="1").</param>
            /// <param name="id">形状唯一 <c>cNvPr id</c>。 The unique cNvPr id.</param>
            /// <returns>nvSpPr 元素。 The nvSpPr element.</returns>
            private static XElement CreateShapeNonVisual(SlideShape shape, bool isTextBox, int id)
            {
                XElement nvPr = new XElement(P + "nvPr");
                XElement ph = CreatePlaceholder(shape);
                if (ph != null) nvPr.Add(ph);
                XElement cNvSpPr = new XElement(P + "cNvSpPr");
                if (isTextBox) cNvSpPr.Add(new XAttribute("txBox", "1"));
                return new XElement(P + "nvSpPr",
                    CreateCnvPr(shape, null, id),
                    cNvSpPr,
                    nvPr);
            }

            /// <summary>生成 <code>p:cNvPr</code>，图片形状附带备选文字描述。 Builds <code>p:cNvPr</code>, adding alternative text for pictures.</summary>
            /// <param name="shape">形状模型。 Shape model.</param>
            /// <param name="descr">备选文字，可为 null。 Alternative text, may be null.</param>
            /// <param name="id">形状唯一 <c>cNvPr id</c>。 The unique cNvPr id.</param>
            /// <returns>cNvPr 元素。 The cNvPr element.</returns>
            private static XElement CreateCnvPr(SlideShape shape, string descr, int id)
            {
                XElement element = new XElement(P + "cNvPr",
                    new XAttribute("id", IntStr(Math.Max(2, id))),
                    new XAttribute("name", string.IsNullOrEmpty(shape.Name) ? "Shape" : shape.Name));
                if (!string.IsNullOrEmpty(descr)) element.Add(new XAttribute("descr", descr));
                return element;
            }

            /// <summary>按占位符类别生成 <code>p:ph</code>；非占位符返回 null。 Builds the placeholder <code>p:ph</code> element; returns null for non-placeholders.</summary>
            /// <param name="shape">形状模型。 Shape model.</param>
            /// <returns>占位符元素或 null。 The placeholder element or null.</returns>
            private static XElement CreatePlaceholder(SlideShape shape)
            {
                if (shape.Placeholder == PlaceholderKind.None) return null;
                XElement ph = new XElement(P + "ph", new XAttribute("type", PlaceholderType(shape.Placeholder)));
                if (shape.Placeholder != PlaceholderKind.Title && shape.PlaceholderIndex >= 0)
                {
                    ph.Add(new XAttribute("idx", IntStr(shape.PlaceholderIndex)));
                }
                return ph;
            }

            /// <summary>把占位符类别映射为 PresentationML 的 <c>ph type</c> 文本。 Maps a placeholder kind to its PresentationML <c>ph type</c> text.</summary>
            /// <param name="kind">占位符类别。 Placeholder kind.</param>
            /// <returns>type 文本。 The type text.</returns>
            private static string PlaceholderType(PlaceholderKind kind)
            {
                switch (kind)
                {
                    case PlaceholderKind.Title: return "title";
                    case PlaceholderKind.Subtitle: return "subTitle";
                    case PlaceholderKind.Body: return "body";
                    case PlaceholderKind.Date: return "dt";
                    case PlaceholderKind.Footer: return "ftr";
                    case PlaceholderKind.SlideNumber: return "sldNum";
                    case PlaceholderKind.Picture: return "pic";
                    default: return "obj";
                }
            }

            /// <summary>生成 <code>p:spPr</code>：xfrm、prstGeom、填充与描边。 Builds <code>p:spPr</code>: xfrm, prstGeom, fill and outline.</summary>
            /// <param name="shape">形状模型。 Shape model.</param>
            /// <param name="auto">自选图形（可为 null）。 The auto shape, or null for a plain text box.</param>
            /// <returns>spPr 元素。 The spPr element.</returns>
            private XElement CreateShapeProperties(SlideShape shape, AutoShape auto)
            {
                XElement spPr = new XElement(P + "spPr", CreateXfrm(shape));
                XElement geom = new XElement(A + "prstGeom", new XAttribute("prst", auto != null ? PrimitiveGeometry(auto.Kind) : "rect"));
                if (auto != null && auto.Kind == AutoShapeKind.RoundedRectangle)
                {
                    int adj = ClampInt((int)RoundToInt(auto.CornerRadius * 100000.0), 0, 50000);
                    geom.Add(new XElement(A + "avLst", new XElement(A + "gd",
                        new XAttribute("name", "adj"), new XAttribute("fmla", "val " + IntStr(adj)))));
                }
                else
                {
                    geom.Add(new XElement(A + "avLst"));
                }
                spPr.Add(geom);
                // 没设过填充/描边就什么都不写：补 a:noFill 等于替调用者宣布「这形状不许有底色」，
                // 会把 p:style 与主题本来要给的外观一并抹掉。显式 BorderStyle.None 仍然写 a:ln + a:noFill。
                // Nothing is written when neither fill nor outline was set: an a:noFill would announce "this
                // shape may have no fill" on the caller's behalf and silence whatever p:style and the theme
                // would supply. An explicit BorderStyle.None still reaches the file as a:ln + a:noFill.
                XElement fill = CreateOptionalFill(shape.FillColor);
                if (fill != null) spPr.Add(fill);
                XElement line = CreateOptionalLine(shape.Outline);
                if (line != null) spPr.Add(line);
                return spPr;
            }

            /// <summary>只在调用者设过填充时返回 <c>a:solidFill</c>，否则返回 null（不写任何填充元素）。Returns <c>a:solidFill</c> only when the caller set a fill, otherwise null so no fill element is written.</summary>
            /// <param name="color">填充色，null 表示没设过。 Fill colour; null means it was never set.</param>
            /// <returns>填充元素或 null。 The fill element, or null.</returns>
            private static XElement CreateOptionalFill(OfficeColor? color)
            {
                if (!color.HasValue) return null;
                return new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", color.Value.ToHex())));
            }

            /// <summary>只在调用者设过描边时返回 <c>a:ln</c>，否则返回 null。Returns <c>a:ln</c> only when the caller set an outline, otherwise null.</summary>
            /// <param name="line">描边，null 表示没设过。 The outline; null means it was never set.</param>
            /// <returns>描边元素或 null。 The line element, or null.</returns>
            private XElement CreateOptionalLine(BorderLine line)
            {
                return line == null ? null : CreateLine(line);
            }

            /// <summary>生成 <code>a:xfrm</code> 的位置与大小。 Builds the <code>a:xfrm</code> offset and extent.</summary>
            /// <param name="shape">形状模型。 Shape model.</param>
            /// <returns>xfrm 元素。 The xfrm element.</returns>
            private XElement CreateXfrm(SlideShape shape)
            {
                return new XElement(A + "xfrm",
                    new XElement(A + "off", new XAttribute("x", IntStr(Ppt.Emu(shape.Left))), new XAttribute("y", IntStr(Ppt.Emu(shape.Top)))),
                    new XElement(A + "ext", new XAttribute("cx", IntStr(Ppt.Emu(shape.Width))), new XAttribute("cy", IntStr(Ppt.Emu(shape.Height)))));
            }

            /// <summary>自选图形类别到 <c>prst</c> 名称的映射。 Maps an auto-shape kind to its <c>prst</c> name.</summary>
            /// <param name="kind">图形类别。 Shape kind.</param>
            /// <returns>预设几何名称。 The preset-geometry name.</returns>
            private static string PrimitiveGeometry(AutoShapeKind kind)
            {
                switch (kind)
                {
                    case AutoShapeKind.Rectangle: return "rect";
                    case AutoShapeKind.RoundedRectangle: return "roundRect";
                    case AutoShapeKind.Ellipse: return "ellipse";
                    case AutoShapeKind.Triangle: return "triangle";
                    case AutoShapeKind.Diamond: return "diamond";
                    case AutoShapeKind.Pentagon: return "pentagon";
                    case AutoShapeKind.Hexagon: return "hexagon";
                    case AutoShapeKind.RightArrow: return "rightArrow";
                    case AutoShapeKind.Chevron: return "chevron";
                    case AutoShapeKind.Line: return "line";
                    default: return "rect";
                }
            }

            /// <summary>生成实色填充或 <code>a:noFill</code>。 Builds a solid fill or <code>a:noFill</code>.</summary>
            /// <param name="color">填充色，null 表示无填充。 Fill color; null means no fill.</param>
            /// <returns>填充元素。 The fill element.</returns>
            private static XElement CreateSolidOrNoFill(OfficeColor? color)
            {
                if (!color.HasValue) return new XElement(A + "noFill");
                return new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", color.Value.ToHex())));
            }

            /// <summary>生成 <code>a:ln</code> 描边：无线输出 noFill，其余按线型加 prstDash，双线改用 <c>cmpd="dbl"</c>。 Builds the <code>a:ln</code> outline: noFill when line-less, otherwise width, fill and prstDash; a double line becomes <c>cmpd="dbl"</c> instead.</summary>
            /// <param name="outline">描边描述。 Outline description.</param>
            /// <returns>ln 元素。 The ln element.</returns>
            private XElement CreateLine(BorderLine outline)
            {
                XElement ln = new XElement(A + "ln");
                if (outline == null || outline.Style == BorderStyle.None)
                {
                    ln.Add(new XElement(A + "noFill"));
                    return ln;
                }
                ln.Add(new XAttribute("w", IntStr(Math.Max(1, Ppt.Emu(outline.Width)))));
                if (outline.Style == BorderStyle.Double) ln.Add(new XAttribute("cmpd", "dbl"));
                ln.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", outline.Color.ToHex()))));
                string dash = DashStyle(outline.Style);
                if (dash != null) ln.Add(new XElement(A + "prstDash", new XAttribute("val", dash)));
                return ln;
            }

            /// <summary>
            /// 线型到 <c>prstDash</c> 值的映射。双线不是虚线图案，<c>ST_PresetLineDashVal</c> 里没有 dbl 这一档，
            /// 它靠 <c>a:ln@cmpd="dbl"</c> 表达，所以这里返回 null；实线同样返回 null（缺省即 solid）。
            /// Maps a border style to a <c>prstDash</c> value. A double line is not a dash pattern —
            /// <c>ST_PresetLineDashVal</c> has no dbl, it lives on <c>a:ln@cmpd="dbl"</c> — so it returns null here,
            /// and so does a solid line (the implicit value).
            /// </summary>
            /// <param name="style">边框线型。 Border style.</param>
            /// <returns>prstDash 值或 null。 The prstDash value or null.</returns>
            private static string DashStyle(BorderStyle style)
            {
                switch (style)
                {
                    case BorderStyle.Dashed: return "dash";
                    case BorderStyle.Dotted: return "sysDot";
                    default: return null;
                }
            }

            /// <summary>生成 <code>p:txBody</code>：bodyPr、lstStyle 与段落。 Builds <code>p:txBody</code>: bodyPr, lstStyle and paragraphs.</summary>
            /// <param name="shape">文本形状。 Text-bearing shape.</param>
            /// <param name="base">基础字体格式（未显式设置的运行按此渲染）。 Base format used for runs left at the model default.</param>
            /// <returns>txBody 元素。 The txBody element.</returns>
            private XElement CreateTextBody(TextBoxShape shape, TextFormat @base)
            {
                XElement bodyPr = new XElement(A + "bodyPr",
                    new XAttribute("lIns", IntStr(Ppt.Emu(shape.MarginLeft))),
                    new XAttribute("tIns", IntStr(Ppt.Emu(shape.MarginTop))),
                    new XAttribute("rIns", IntStr(Ppt.Emu(shape.MarginRight))),
                    new XAttribute("bIns", IntStr(Ppt.Emu(shape.MarginBottom))),
                    new XAttribute("anchor", AnchorValue(shape.VerticalAlignment)),
                    new XAttribute("wrap", WrapValue(shape.WrapText)));
                if (shape.ShrinkToFit) bodyPr.Add(new XElement(A + "normAutofit"));
                XElement txBody = new XElement(P + "txBody", bodyPr, new XElement(A + "lstStyle"));
                AppendParagraphs(txBody, shape.Paragraphs, @base);
                return txBody;
            }

            /// <summary>折行开关到 <c>wrap</c> 值的映射：只有显式「不折行」写成 none，其余（含未设置）都是 square——那正是 PowerPoint 自己不写这个属性时的行为。 Maps the wrapping switch onto <c>wrap</c>: only an explicit "do not fold" becomes none, everything else (including unset) is square, which is what PowerPoint itself means when the attribute is absent.</summary>
            /// <param name="wrap">模型里的折行开关。 The model's wrapping switch.</param>
            /// <returns>none 或 square。 Either none or square.</returns>
            private static string WrapValue(bool? wrap)
            {
                return wrap == false ? "none" : "square";
            }

            /// <summary>垂直对齐到 <c>anchor</c> 值的映射。 Maps the vertical anchor to its <c>anchor</c> value.</summary>
            /// <param name="anchor">垂直对齐。 Vertical alignment.</param>
            /// <returns>t、ctr 或 b。 One of t, ctr or b.</returns>
            private static string AnchorValue(SlideTextAnchor anchor)
            {
                switch (anchor)
                {
                    case SlideTextAnchor.Middle: return "ctr";
                    case SlideTextAnchor.Bottom: return "b";
                    default: return "t";
                }
            }

            /// <summary>把段落集合逐个转成 <code>a:p</code> 追加到容器。 Converts each paragraph into an <code>a:p</code> and appends it to the container.</summary>
            /// <param name="container">目标容器（txBody 或单元格 txBody）。 Target container (a txBody or a cell txBody).</param>
            /// <param name="paragraphs">段落列表。 Paragraph list.</param>
            /// <param name="base">基础字体格式。 Base format.</param>
            private void AppendParagraphs(XElement container, List<Paragraph> paragraphs, TextFormat @base)
            {
                if (paragraphs == null || paragraphs.Count == 0)
                {
                    container.Add(CreateParagraph(new Paragraph(string.Empty), @base));
                    return;
                }
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    Paragraph paragraph = paragraphs[i] ?? new Paragraph(string.Empty);
                    container.Add(CreateParagraph(paragraph, @base));
                }
            }

            /// <summary>生成单个 <code>a:p</code>：pPr、文本运行与 endParaRPr。 Builds one <code>a:p</code>: pPr, text runs and endParaRPr.</summary>
            /// <param name="paragraph">段落模型。 Paragraph model.</param>
            /// <param name="base">基础字体格式。 Base format.</param>
            /// <returns>a:p 元素。 The a:p element.</returns>
            private XElement CreateParagraph(Paragraph paragraph, TextFormat @base)
            {
                XElement p = new XElement(A + "p");
                XElement pPr = CreateParagraphProperties(paragraph, @base);
                if (pPr != null) p.Add(pPr);
                List<Inline> inlines = paragraph.Inlines ?? new List<Inline>();
                TextFormat last = @base;
                bool any = false;
                for (int i = 0; i < inlines.Count; i++)
                {
                    TextRun run = inlines[i] as TextRun;
                    if (run != null)
                    {
                        if (string.IsNullOrEmpty(run.Text)) continue;
                        TextFormat fmt = ResolveFormat(run.Format, @base);
                        last = fmt;
                        p.Add(CreateRun(run.Text, fmt));
                        any = true;
                        continue;
                    }
                    // FieldRun 与 ImageRun 在幻灯片文本里降级为纯文本占位（图片输出 [image]），
                    // 因为 PresentationML 的形状文本不支持内联图片或域代码。
                    // FieldRun and ImageRun degrade to plain text on slides (an image becomes "[image]"),
                    // because PresentationML shape text supports neither inline pictures nor fields.
                    FieldRun field = inlines[i] as FieldRun;
                    if (field != null)
                    {
                        if (!string.IsNullOrEmpty(field.DisplayText))
                        {
                            p.Add(CreateRun(field.DisplayText, ResolveFormat(field.Format, @base)));
                            any = true;
                        }
                        continue;
                    }
                    if (inlines[i] is ImageRun)
                    {
                        p.Add(CreateRun("[image]", @base));
                        any = true;
                    }
                }
                if (!any)
                {
                    XElement endR = CharProps("endParaRPr", @base);
                    p.Add(endR);
                }
                else
                {
                    p.Add(CharProps("endParaRPr", last));
                }
                return p;
            }

            /// <summary>生成 <code>a:pPr</code>（顺序：lnSpc、spcBef、spcAft、bullet、defRPr）。 Builds <code>a:pPr</code> in schema order: lnSpc, spcBef, spcAft, bullet, defRPr.</summary>
            /// <param name="paragraph">段落模型。 Paragraph model.</param>
            /// <param name="base">基础字体格式。 Base format.</param>
            /// <returns>pPr 元素；无任何属性时返回 null。 The pPr element, or null when nothing to emit.</returns>
            private XElement CreateParagraphProperties(Paragraph paragraph, TextFormat @base)
            {
                XElement pPr = new XElement(A + "pPr");
                bool any = false;
                if (paragraph.Alignment != ParagraphAlignment.Left)
                {
                    pPr.Add(new XAttribute("algn", AlignmentValue(paragraph.Alignment)));
                    any = true;
                }
                if (paragraph.LeftIndent != 0)
                {
                    pPr.Add(new XAttribute("marL", IntStr(Ppt.Emu(paragraph.LeftIndent))));
                    any = true;
                }
                if (paragraph.FirstLineIndent != 0)
                {
                    pPr.Add(new XAttribute("indent", IntStr(Ppt.Emu(paragraph.FirstLineIndent))));
                    any = true;
                }
                if (paragraph.LineSpacing > 0 && Math.Abs(paragraph.LineSpacing - 1.0) > 0.0001)
                {
                    pPr.Add(new XElement(A + "lnSpc",
                        new XElement(A + "spcPct", new XAttribute("val", IntStr(RoundToInt(paragraph.LineSpacing * 100000.0))))));
                    any = true;
                }
                if (paragraph.SpaceBefore > 0)
                {
                    pPr.Add(new XElement(A + "spcBef", new XElement(A + "spcPts", new XAttribute("val", IntStr(RoundToInt(paragraph.SpaceBefore * 100.0))))));
                    any = true;
                }
                if (paragraph.SpaceAfter > 0)
                {
                    pPr.Add(new XElement(A + "spcAft", new XElement(A + "spcPts", new XAttribute("val", IntStr(RoundToInt(paragraph.SpaceAfter * 100.0))))));
                    any = true;
                }
                XElement bullet = CreateBullet(paragraph.ListMarker);
                if (bullet != null)
                {
                    pPr.Add(bullet);
                    any = true;
                }
                if (!any) return null;
                return pPr;
            }

            /// <summary>
            /// 生成项目符号子元素：无序用 buChar，有序（数字开头）用 buAutoNum；没标记时返回 null，什么都不写。
            /// 这里不能用 a:buNone 占位——buNone 是一条「关掉上层符号」的显式指令，写进每一段就等于把
            /// 第三方母版与版式里定义好的项目符号全数压制；而「没标记」只是模型里没有这句话，
            /// 该让继承链自己说话。本库自建的母版与版式从不定义符号，省掉这一句对自家产物毫无影响。
            /// Builds the bullet child: buChar for unordered, buAutoNum for ordered (digit-prefixed) markers,
            /// null for no marker. A buNone placeholder is not an option: buNone is an explicit instruction to
            /// switch off the level's bullet, so emitting it on every paragraph suppresses bullets defined by
            /// third-party masters and layouts. "No marker" means the model says nothing, and the inheritance
            /// chain should be left to speak. This library's own masters and layouts never define bullets, so
            /// omitting the child changes nothing in self-generated decks.
            /// </summary>
            /// <param name="marker">列表标记文本。 List marker text.</param>
            /// <returns>bullet 元素；无标记时为 null。 The bullet element, or null when there is no marker.</returns>
            private static XElement CreateBullet(string marker)
            {
                if (string.IsNullOrEmpty(marker)) return null;
                char first = marker[0];
                if (first >= '0' && first <= '9')
                {
                    return new XElement(A + "buAutoNum", new XAttribute("type", "arabicPeriod"));
                }
                return new XElement(A + "buChar", new XAttribute("char", marker));
            }

            /// <summary>段落对齐到 <c>algn</c> 值的映射。 Maps paragraph alignment to its <c>algn</c> value.</summary>
            /// <param name="alignment">对齐方式。 Alignment.</param>
            /// <returns>l、ctr、r 或 just。 One of l, ctr, r or just.</returns>
            private static string AlignmentValue(ParagraphAlignment alignment)
            {
                switch (alignment)
                {
                    case ParagraphAlignment.Center: return "ctr";
                    case ParagraphAlignment.Right: return "r";
                    case ParagraphAlignment.Justify: return "just";
                    default: return "l";
                }
            }

            /// <summary>生成一个 <code>a:r</code> 文本运行。 Builds one <code>a:r</code> text run.</summary>
            /// <param name="text">运行文本。 Run text.</param>
            /// <param name="format">运行格式。 Run format.</param>
            /// <returns>a:r 元素。 The a:r element.</returns>
            private static XElement CreateRun(string text, TextFormat format)
            {
                XElement rPr = CharProps("rPr", format);
                return new XElement(A + "r", rPr, new XElement(A + "t", text));
            }

            /// <summary>
            /// 生成字符属性元素（a:rPr 或 a:endParaRPr）：lang 恒写，字号、加粗等开关、颜色与字体只写显式设过的。
            /// 模型无参构造预置的 Arial/11 磅/黑色不是意图：把「等于默认值」当成「设过」，一个只加粗的词会被钉成
            /// 11 磅的 Calibri，母版继承链上 18pt 的正文就此缩掉。
            /// Builds a character-properties element (a:rPr or a:endParaRPr): lang always, then the size, the
            /// b/i/u/strike toggles, the colour and the fonts — but only those actually set. The Arial/11pt/black
            /// the model's parameterless constructor presets is not intent; writing it as if it were pins a
            /// bold-only word to 11pt Calibri and shrinks the 18pt body it inherited from the master.
            /// </summary>
            /// <param name="tag">元素标签名。 Element tag.</param>
            /// <param name="format">文本格式。 Text format.</param>
            /// <returns>字符属性元素。 The character-properties element.</returns>
            private static XElement CharProps(string tag, TextFormat format)
            {
                XElement el = new XElement(A + tag, new XAttribute("lang", "zh-CN"));
                if (format.HasExplicitFontSize)
                {
                    int size = (int)RoundToInt(format.FontSize * 100.0);
                    if (size > 0) el.Add(new XAttribute("sz", IntStr(size)));
                }
                if (format.Bold) el.Add(new XAttribute("b", "1"));
                if (format.Italic) el.Add(new XAttribute("i", "1"));
                if (format.Underline) el.Add(new XAttribute("u", "sng"));
                if (format.Strikeout) el.Add(new XAttribute("strike", "sngStrike"));
                // 上下标由 a:rPr/@baseline 表达：上标 +30%、下标 -25%，与 PowerPoint 的取值一致。
                // 读端一直按这个属性还原 Superscript/Subscript，写端过去从不输出，于是往返一次就丢。
                // Super- and subscript are expressed by a:rPr/@baseline (+30% and -25%, the values PowerPoint
                // uses). The reader has always restored Superscript/Subscript from that attribute, but the
                // writer never emitted it, so a round trip dropped it.
                if (format.Superscript) el.Add(new XAttribute("baseline", "30000"));
                else if (format.Subscript) el.Add(new XAttribute("baseline", "-25000"));
                if (format.HasExplicitColor)
                {
                    el.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", format.Color.ToHex()))));
                }
                if (format.HasExplicitFontFamily)
                {
                    el.Add(new XElement(A + "latin", new XAttribute("typeface", SafeFamily(format.GetAsciiFontFamily()))));
                    el.Add(new XElement(A + "ea", new XAttribute("typeface", SafeFamily(format.GetEastAsiaFontFamily()))));
                }
                return el;
            }

            /// <summary>字体名为空时回退 Calibri。 Falls back to Calibri when the font family name is empty.</summary>
            /// <param name="family">字体名。 Family name.</param>
            /// <returns>非空字体名。 A non-empty family name.</returns>
            private static string SafeFamily(string family)
            {
                return string.IsNullOrEmpty(family) ? "Calibri" : family;
            }

            /// <summary>按占位符类别确定基础字体格式（标题用标题默认，其余用正文默认）。 Chooses the base format by placeholder kind (title uses the title default, others the body default).</summary>
            /// <param name="shape">形状模型。 Shape model.</param>
            /// <returns>基础格式。 The base format.</returns>
            private TextFormat ResolveBaseFormat(SlideShape shape)
            {
                return shape.Placeholder == PlaceholderKind.Title ? _titleFormat : _bodyFormat;
            }

            /// <summary>
            /// 若运行未显式设置格式（与模型纯默认值一致），则采用基础格式，从而尊重演示文稿的默认字体；
            /// 否则保留运行自带格式。
            /// Applies the base format when a run is left at the model's bare default (so the presentation's default
            /// fonts win); otherwise the run's own format is preserved.
            /// </summary>
            /// <param name="runFormat">运行格式。 Run format.</param>
            /// <param name="base">基础格式。 Base format.</param>
            /// <returns>生效格式。 The effective format.</returns>
            private static TextFormat ResolveFormat(TextFormat runFormat, TextFormat @base)
            {
                if (runFormat == null) return @base;
                return IsGenericDefault(runFormat) ? @base : runFormat;
            }

            /// <summary>判断格式是否等于模型的“纯默认”（Arial/11/黑色、无任何加粗等）。 Checks whether a format equals the model's bare default (Arial/11/black with no emphasis).</summary>
            /// <param name="format">待判定的格式。 Format to test.</param>
            /// <returns>是纯默认返回 true。 True when it is the bare default.</returns>
            private static bool IsGenericDefault(TextFormat format)
            {
                return format.FontFamily == "Arial"
                    && string.IsNullOrEmpty(format.AsciiFontFamily)
                    && string.IsNullOrEmpty(format.EastAsiaFontFamily)
                    && Math.Abs(format.FontSize - 11.0) < 0.0001
                    && !format.Bold && !format.Italic && !format.Underline && !format.Strikeout
                    && !format.Superscript && !format.Subscript
                    && format.Color == OfficeColor.Black;
            }

            /// <summary>生成图片形状的 <code>p:pic</code>；图片字节去重后登记到幻灯片关系。 Builds a picture's <code>p:pic</code>, de-duplicating the image bytes and registering a slide relationship.</summary>
            /// <param name="picture">图片形状。 Picture shape.</param>
            /// <param name="part">所属幻灯片条目。 Owning slide entry.</param>
            /// <param name="id">形状唯一 <c>cNvPr id</c>。 The unique cNvPr id.</param>
            /// <returns>p:pic 元素。 The p:pic element.</returns>
            private XElement CreatePicture(PictureShape picture, SlidePart part, int id)
            {
                XElement nvPicPr = new XElement(P + "nvPicPr",
                    CreateCnvPr(picture, picture.AlternativeText, id),
                    new XElement(P + "cNvPicPr", new XElement(A + "picLocks", new XAttribute("noChangeAspect", "1"))));
                XElement nvPr = new XElement(P + "nvPr");
                XElement ph = CreatePlaceholder(picture);
                if (ph != null) nvPr.Add(ph);
                nvPicPr.Add(nvPr);

                XElement blipFill = new XElement(P + "blipFill");
                XElement blip = new XElement(A + "blip");
                if (picture.Image != null && picture.Image.Data != null && picture.Image.Data.Length > 0)
                {
                    MediaInfo media = RegisterMedia(picture.Image);
                    string rid = part.AddRel(RelImage, media.Target);
                    blip.Add(new XAttribute(R + "embed", rid));
                }
                blipFill.Add(blip);
                blipFill.Add(new XElement(A + "stretch", new XElement(A + "fillRect")));

                XElement spPr = new XElement(P + "spPr",
                    CreateXfrm(picture),
                    new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst")));
                if (picture.FillColor.HasValue) spPr.Add(CreateSolidOrNoFill(picture.FillColor));
                // 没设过描边就整条不写：a:ln/noFill 是一句「本图不要线」的真话资格在调用者身上，
                // 写端替它说出口，版式与母版本来的描边就被压掉——自选图形与文本框已按这条修过。
                // An outline never set stays unwritten: only the caller may say "no line for this picture",
                // and saying it on their behalf suppresses the line the layout or master supplies. Shapes
                // and text boxes already follow this rule.
                XElement pictureLine = CreateOptionalLine(picture.Outline);
                if (pictureLine != null) spPr.Add(pictureLine);

                return new XElement(P + "pic", nvPicPr, blipFill, spPr);
            }

            /// <summary>登记图片到全局媒体表（按引用去重），返回其媒体信息。 Registers an image in the global media table (de-duplicated by reference) and returns its info.</summary>
            /// <param name="image">图片对象。 Image.</param>
            /// <returns>媒体登记信息。 The media registration.</returns>
            private MediaInfo RegisterMedia(OfficeImage image)
            {
                MediaInfo existing;
                if (_media.TryGetValue(image, out existing)) return existing;
                _mediaCount++;
                string extension = string.IsNullOrEmpty(image.FileExtension) ? ".png" : image.FileExtension.ToLowerInvariant();
                string fileName = "image" + IntStr(_mediaCount) + extension;
                MediaInfo media = new MediaInfo();
                media.Target = "../media/" + fileName;
                media.Extension = extension;
                media.ContentType = string.IsNullOrEmpty(image.ContentType) ? CtPng : image.ContentType;
                _media[image] = media;
                // PNG/JPEG 已是压缩格式，zip 内直接存储，省掉无效压缩。
                // PNG/JPEG are already compressed; store them raw instead of deflating again.
                WriteEntryBytes("ppt/media/" + fileName, image.Data, CompressionLevel.NoCompression);
                if (!IsCoveredDefault(extension))
                {
                    _extraMediaTypes[extension.TrimStart('.')] = media.ContentType;
                }
                return media;
            }

            /// <summary>判断扩展名是否已被 [Content_Types].xml 的 Default 覆盖。 Checks whether an extension is already covered by a Default entry.</summary>
            /// <param name="extension">含点的扩展名。 Extension with the leading dot.</param>
            /// <returns>已覆盖返回 true。 True when covered.</returns>
            private static bool IsCoveredDefault(string extension)
            {
                return extension == ".png" || extension == ".jpg" || extension == ".jpeg" || extension == ".gif";
            }

            /// <summary>生成表格形状的 <code>p:graphicFrame</code>。 Builds a table shape's <code>p:graphicFrame</code>.</summary>
            /// <param name="table">表格形状。 Table shape.</param>
            /// <param name="id">形状唯一 <c>cNvPr id</c>。 The unique cNvPr id.</param>
            /// <returns>p:graphicFrame 元素。 The graphic-frame element.</returns>
            private XElement CreateTable(TableShape table, int id)
            {
                XElement nvPr = new XElement(P + "nvPr");
                XElement ph = CreatePlaceholder(table);
                if (ph != null) nvPr.Add(ph);
                XElement nv = new XElement(P + "nvGraphicFramePr",
                    CreateCnvPr(table, null, id),
                    new XElement(P + "cNvGraphicFramePr", new XElement(A + "graphicFrameLocks", new XAttribute("noGrp", "1"))),
                    nvPr);
                XElement xfrm = new XElement(P + "xfrm",
                    new XElement(A + "off", new XAttribute("x", IntStr(Ppt.Emu(table.Left))), new XAttribute("y", IntStr(Ppt.Emu(table.Top)))),
                    new XElement(A + "ext", new XAttribute("cx", IntStr(Ppt.Emu(table.Width))), new XAttribute("cy", IntStr(Ppt.Emu(table.Height)))));
                XElement graphic = new XElement(A + "graphic",
                    new XElement(A + "graphicData", new XAttribute("uri", TableGraphicUri), CreateTableElement(table)));
                return new XElement(P + "graphicFrame", nv, xfrm, graphic);
            }

            /// <summary>生成 <code>a:tbl</code>：tblPr、tblGrid 与逐行单元格。 Builds <code>a:tbl</code>: tblPr, tblGrid and rows of cells.</summary>
            /// <param name="table">表格形状。 Table shape.</param>
            /// <returns>a:tbl 元素。 The table element.</returns>
            private XElement CreateTableElement(TableShape table)
            {
                Table model = table.Table ?? new Table();
                List<TableRow> rows = model.Rows ?? new List<TableRow>();
                int gridCols = ComputeGridColumns(model, rows);

                // tblPr 的开关只写模型说过的话：firstRow 跟着显式表头走，带状行（bandRow）在统一模型里
                // 没有对应字段，因此不再恒写——过去每张表出门都自带一套没人要过的条纹。
                // The tblPr switches now say only what the model said: firstRow follows a declared header,
                // and bandRow, which has no counterpart in the unified model, is no longer written for every
                // table — until now every table left the file already striped.
                XElement tblPr = new XElement(A + "tblPr", new XElement(A + "tableStyleId", DefaultTableStyleId));
                bool styledHeader = HasDeclaredHeader(table, rows);
                if (styledHeader) tblPr.Add(new XAttribute("firstRow", "1"));

                XElement tbl = new XElement(A + "tbl", tblPr,
                    CreateGrid(model, rows, gridCols, Ppt.Emu(table.Width)));

                OfficeColor headerFill = table.HeaderColor.HasValue ? table.HeaderColor.Value : AccentColor(0);
                double padding = model.CellPadding > 0 ? model.CellPadding : 4.0;
                TableBorders borders = model.Borders ?? new TableBorders();
                int totalRows = rows.Count;
                // 纵向合并要占住后续行：过去只写 gridSpan，续接格被整行丢掉，
                // 于是第 2 行的数据整体左移一列。这里按网格记录每列被谁占到第几行。
                // A vertical merge occupies the following rows: only gridSpan used to be written and the
                // continuation cells were dropped, shifting the next row's data one column to the left.
                // The grid below records, per column, which row the merge runs through and where it started.
                int[] coveredUntil = new int[gridCols + 1];
                int[] mergeOrigin = new int[gridCols + 1];
                for (int c = 0; c <= gridCols; c++)
                {
                    coveredUntil[c] = -1;
                    mergeOrigin[c] = 0;
                }
                for (int r = 0; r < totalRows; r++)
                {
                    TableRow row = rows[r];
                    if (row == null) continue;
                    bool isHeader = styledHeader && IsHeaderRowAt(model, rows, r);
                    int h = Ppt.Emu(row.Height > 0 ? row.Height : DefaultRowHeight(table.Height, totalRows));
                    XElement tr = new XElement(A + "tr", new XAttribute("h", IntStr(Math.Max(1, h))));
                    List<TableCell> cells = row.Cells ?? new List<TableCell>();
                    OfficeColor? fillerFill = isHeader ? (OfficeColor?)headerFill : null;
                    int column = 1;
                    int source = 0;
                    while (column <= gridCols)
                    {
                        if (coveredUntil[column] >= r)
                        {
                            // 续接格：纵向合并块的第一格写 vMerge="1"，同一块里其余列写 hMerge="1"。
                            // Continuation cell: the block's first column gets vMerge="1" and the remaining
                            // columns of the same block get hMerge="1".
                            bool blockStart = column == mergeOrigin[column];
                            int blockSpan = coveredUntil[column] - r + 1;
                            tr.Add(CreateMergedCell(blockStart, fillerFill, isHeader, padding,
                                ResolveEdges(borders, r, blockSpan, totalRows, column, 1, gridCols)));
                            column++;
                            continue;
                        }
                        if (source >= cells.Count)
                        {
                            // 补齐用的空格跟着所在行的待遇走：表头行里的空格借表头色，免得一条表头被补位格截成两截；
                            // 数据行里的空格什么都不写，让 tableStyleId 说话。
                            // A filler cell follows its row's treatment: inside a header row it takes the header
                            // fill so the band is not cut in two, and in a data row it writes no fill at all and
                            // lets the tableStyleId speak.
                            tr.Add(new XElement(A + "tc", new XElement(A + "txBody",
                                new XElement(A + "bodyPr"), new XElement(A + "lstStyle"), new XElement(A + "p")),
                                CreateTcProperties(null, fillerFill, isHeader, padding,
                                    ResolveEdges(borders, r, 1, totalRows, column, 1, gridCols), 1)));
                            column++;
                            continue;
                        }
                        TableCell cell = cells[source++];
                        if (cell == null) continue;
                        int span = ClampInt(cell.ColSpan, 1, gridCols - column + 1);
                        int rowSpan = ClampInt(cell.RowSpan, 1, totalRows - r);
                        tr.Add(CreateTableCell(cell, _bodyFormat, isHeader, headerFill, padding,
                            ResolveEdges(borders, r, rowSpan, totalRows, column, span, gridCols), span, rowSpan));
                        for (int k = 0; k < span; k++)
                        {
                            coveredUntil[column + k] = r + rowSpan - 1;
                            mergeOrigin[column + k] = column;
                        }
                        column += span;
                    }
                    tbl.Add(tr);
                }
                return tbl;
            }

            /// <summary>
            /// 生成一个合并续接格 <code>a:tc</code>：块首列 <c>vMerge="1"</c>，块内其余列 <c>hMerge="1"</c>，
            /// 内容为空（内容在原格上）。缺了这些格子，后续行的真实单元格会顶到错误的列上。
            /// Builds one merge-continuation <code>a:tc</code>: <c>vMerge="1"</c> on the block's first column and
            /// <c>hMerge="1"</c> on the rest, with an empty body (the content lives on the origin cell). Without
            /// these cells the following rows' real cells land on the wrong columns.
            /// </summary>
            /// <param name="blockStart">是否为合并块的首列。Whether this is the block's first column.</param>
            /// <param name="fill">底色。Fill colour.</param>
            /// <param name="header">是否表头行。Whether the row is a header row.</param>
            /// <param name="padding">内边距（磅）。Padding in points.</param>
            /// <param name="border">该格四边的边框线。 The four border lines for this cell.</param>
            /// <returns>a:tc 元素。 The cell element.</returns>
            private XElement CreateMergedCell(bool blockStart, OfficeColor? fill, bool header, double padding, CellBorderEdges border)
            {
                XElement tc = new XElement(A + "tc");
                tc.Add(new XAttribute(blockStart ? "vMerge" : "hMerge", "1"));
                tc.Add(new XElement(A + "txBody",
                    new XElement(A + "bodyPr"), new XElement(A + "lstStyle"), new XElement(A + "p")));
                tc.Add(CreateTcProperties(null, fill, header, padding, border, 1));
                return tc;
            }

            /// <summary>生成单个表格单元格 <code>a:tc</code>。 Builds one table cell <code>a:tc</code>.</summary>
            /// <param name="cell">单元格模型。 Cell model.</param>
            /// <param name="base">基础格式。 Base format.</param>
            /// <param name="header">是否为表头单元格。 Whether it is a header cell.</param>
            /// <param name="headerFill">表头底色。 Header fill color.</param>
            /// <param name="padding">内边距（磅）。 Padding in points.</param>
            /// <param name="border">该格四边的边框线。 The four border lines for this cell.</param>
            /// <param name="span">横向跨度。 Horizontal span.</param>
            /// <param name="rowSpan">纵向跨度。 Vertical span.</param>
            /// <returns>a:tc 元素。 The cell element.</returns>
            private XElement CreateTableCell(TableCell cell, TextFormat @base, bool header, OfficeColor headerFill, double padding, CellBorderEdges border, int span, int rowSpan)
            {
                XElement tc = new XElement(A + "tc");
                if (span > 1) tc.Add(new XAttribute("gridSpan", IntStr(span)));
                if (rowSpan > 1) tc.Add(new XAttribute("rowSpan", IntStr(rowSpan)));
                XElement txBody = new XElement(A + "txBody", new XElement(A + "bodyPr"), new XElement(A + "lstStyle"));
                List<Paragraph> paragraphs = CollectCellParagraphs(cell);
                TextFormat cellBase = @base;
                if (header)
                {
                    cellBase = @base.Clone();
                    cellBase.Bold = true;
                    cellBase.Color = OfficeColor.White;
                }
                AppendParagraphs(txBody, paragraphs, cellBase);
                tc.Add(txBody);
                OfficeColor? fill = header ? (OfficeColor?)headerFill : cell.BackgroundColor;
                tc.Add(CreateTcProperties(cell, fill, header, padding, border, span));
                return tc;
            }

            /// <summary>
            /// 生成 <code>a:tcPr</code>：内边距、锚点、底色与四边线。底色没设过就什么都不写——
            /// 补一个 <code>a:noFill</code> 等于替调用者宣布「这格不许有底色」，会把 <code>tableStyleId</code>
            /// 本来要给的条纹与整表填充一并抹掉。
            /// Builds <code>a:tcPr</code>: insets, anchor, fill and four border edges. No fill at all when none
            /// was set — emitting <code>a:noFill</code> would announce "this cell may not be filled" on the
            /// caller's behalf and wipe out the banding and whole-table fill the <code>tableStyleId</code>
            /// was about to give.
            /// </summary>
            /// <param name="cell">单元格模型（可为 null）。 Cell model, may be null.</param>
            /// <param name="fill">底色覆盖；表头单元格使用表头色。 Fill override; header cells use the header color.</param>
            /// <param name="header">是否表头。 Whether it is a header cell.</param>
            /// <param name="padding">内边距（磅）。 Padding in points.</param>
            /// <param name="border">该格四边的边框线。 The four border lines for this cell.</param>
            /// <param name="span">横向跨度。 Horizontal span.</param>
            /// <returns>tcPr 元素。 The tcPr element.</returns>
            private XElement CreateTcProperties(TableCell cell, OfficeColor? fill, bool header, double padding, CellBorderEdges border, int span)
            {
                string inset = IntStr(Ppt.Emu(padding));
                XElement tcPr = new XElement(A + "tcPr",
                    new XAttribute("marL", inset),
                    new XAttribute("marR", inset),
                    new XAttribute("marT", inset),
                    new XAttribute("marB", inset),
                    new XAttribute("anchor", cell != null ? CellAnchor(cell.VerticalAlignment) : "ctr"));
                BorderLine left = border.Left, right = border.Right, top = border.Top, bottom = border.Bottom;
                if (cell != null && cell.Borders != null)
                {
                    if (cell.Borders.Left != null) left = cell.Borders.Left;
                    if (cell.Borders.Right != null) right = cell.Borders.Right;
                    if (cell.Borders.Top != null) top = cell.Borders.Top;
                    if (cell.Borders.Bottom != null) bottom = cell.Borders.Bottom;
                }
                tcPr.Add(CreateCellBorder("lnL", left));
                tcPr.Add(CreateCellBorder("lnR", right));
                tcPr.Add(CreateCellBorder("lnT", top));
                tcPr.Add(CreateCellBorder("lnB", bottom));
                if (fill.HasValue)
                {
                    tcPr.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", fill.Value.ToHex()))));
                }
                return tcPr;
            }

            /// <summary>
            /// 生成单元格的单边线元素（lnL/lnR/lnT/lnB）。双线走 <c>cmpd="dbl"</c>，不是 prstDash 取值；
            /// <c>BorderStyle.None</c> 写 <c>a:noFill</c>——补一条 solidFill 等于把调用者撤掉的线画回来。
            /// Builds one cell border edge element (lnL/lnR/lnT/lnB). A double line goes on <c>cmpd="dbl"</c>,
            /// never on prstDash; <c>BorderStyle.None</c> writes <c>a:noFill</c>, because a solidFill there would
            /// redraw the very line the caller removed.
            /// </summary>
            /// <param name="tag">边元素标签。 Edge element tag.</param>
            /// <param name="line">边线描述。 Border line.</param>
            /// <returns>边线元素。 The border element.</returns>
            private XElement CreateCellBorder(string tag, BorderLine line)
            {
                XElement ln = new XElement(A + tag);
                ln.Add(new XAttribute("w", IntStr(Math.Max(1, Ppt.Emu(line.Width > 0 ? line.Width : 0.5)))));
                ln.Add(new XAttribute("cmpd", line.Style == BorderStyle.Double ? "dbl" : "sng"), new XAttribute("algn", "ctr"));
                if (line.Style == BorderStyle.None)
                {
                    ln.Add(new XElement(A + "noFill"));
                    return ln;
                }
                ln.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", line.Color.ToHex()))));
                string dash = DashStyle(line.Style);
                if (dash != null) ln.Add(new XElement(A + "prstDash", new XAttribute("val", dash)));
                return ln;
            }

            /// <summary>
            /// 按单元格在网格里的位置取出它四边的线：落在外圈的那条边用 <c>TableBorders</c> 的 Top/Bottom/
            /// Left/Right，内部的分隔用 InsideHorizontal/InsideVertical。从前这四个位置一律回落到
            /// <c>InsideHorizontal</c>，于是另外五份声明连一个字节都没写下。
            /// Picks the four lines for a cell by its grid position: an edge on the ring takes Top/Bottom/
            /// Left/Right from <c>TableBorders</c>, an inner rule takes InsideHorizontal/InsideVertical.
            /// Every edge used to fall back to <c>InsideHorizontal</c>, so the other five declarations never
            /// reached a single byte.
            /// </summary>
            /// <param name="borders">表格六向边框。 The table's six-way border set.</param>
            /// <param name="row">首行下标（0 起）。 Zero-based first row.</param>
            /// <param name="rowSpan">纵向跨度。 Vertical span.</param>
            /// <param name="totalRows">总行数。 Total row count.</param>
            /// <param name="column">首列（1 起的网格列号）。 One-based first grid column.</param>
            /// <param name="colSpan">横向跨度。 Horizontal span.</param>
            /// <param name="gridCols">网格列数。 Grid column count.</param>
            /// <returns>该格四边的线。 The four lines for this cell.</returns>
            private static CellBorderEdges ResolveEdges(TableBorders borders, int row, int rowSpan, int totalRows, int column, int colSpan, int gridCols)
            {
                BorderLine thin = BorderLine.Thin();
                bool firstRow = row <= 0;
                bool lastRow = row + Math.Max(1, rowSpan) - 1 >= totalRows - 1;
                bool firstColumn = column <= 1;
                bool lastColumn = column + Math.Max(1, colSpan) - 1 >= gridCols;
                return new CellBorderEdges(
                    firstColumn ? LineOrDefault(borders.Left, thin) : LineOrDefault(borders.InsideVertical, thin),
                    lastColumn ? LineOrDefault(borders.Right, thin) : LineOrDefault(borders.InsideVertical, thin),
                    firstRow ? LineOrDefault(borders.Top, thin) : LineOrDefault(borders.InsideHorizontal, thin),
                    lastRow ? LineOrDefault(borders.Bottom, thin) : LineOrDefault(borders.InsideHorizontal, thin));
            }

            /// <summary>声明过就用，没声明就退回细线。 Uses the declared line, or the thin default when none was set.</summary>
            /// <param name="declared">声明过的边框线。 The declared border line.</param>
            /// <param name="fallback">退回用的线。 The fallback line.</param>
            /// <returns>要写出的边框线。 The line to write.</returns>
            private static BorderLine LineOrDefault(BorderLine declared, BorderLine fallback)
            {
                return declared != null ? declared : fallback;
            }

            /// <summary>一个单元格四边各自取到的边框线。 The four border lines resolved for one cell.</summary>
            private sealed class CellBorderEdges
            {
                /// <summary>左边线。 The left edge line.</summary>
                public readonly BorderLine Left;
                /// <summary>右边线。 The right edge line.</summary>
                public readonly BorderLine Right;
                /// <summary>上边线。 The top edge line.</summary>
                public readonly BorderLine Top;
                /// <summary>下边线。 The bottom edge line.</summary>
                public readonly BorderLine Bottom;

                /// <summary>按左右上下顺序装好四条线。 Stores the four lines in left, right, top, bottom order.</summary>
                /// <param name="left">左边线。 The left edge line.</param>
                /// <param name="right">右边线。 The right edge line.</param>
                /// <param name="top">上边线。 The top edge line.</param>
                /// <param name="bottom">下边线。 The bottom edge line.</param>
                public CellBorderEdges(BorderLine left, BorderLine right, BorderLine top, BorderLine bottom)
                {
                    Left = left;
                    Right = right;
                    Top = top;
                    Bottom = bottom;
                }
            }

            /// <summary>单元格垂直对齐到 <c>anchor</c> 值的映射。 Maps a cell's vertical alignment to its <c>anchor</c> value.</summary>
            /// <param name="alignment">垂直对齐。 Vertical alignment.</param>
            /// <returns>t、ctr 或 b。 One of t, ctr or b.</returns>
            private static string CellAnchor(VerticalAlignment alignment)
            {
                switch (alignment)
                {
                    case VerticalAlignment.Middle: return "ctr";
                    case VerticalAlignment.Bottom: return "b";
                    default: return "t";
                }
            }

            /// <summary>收集单元格内的段落（其余块级元素降级为跳过）。 Collects the paragraphs inside a cell (other block elements are skipped).</summary>
            /// <param name="cell">单元格。 Cell.</param>
            /// <returns>段落列表；无段落时返回空列表。 The paragraph list; empty when there is none.</returns>
            private static List<Paragraph> CollectCellParagraphs(TableCell cell)
            {
                List<Paragraph> result = new List<Paragraph>();
                if (cell == null || cell.Blocks == null) return result;
                for (int i = 0; i < cell.Blocks.Count; i++)
                {
                    Paragraph paragraph = cell.Blocks[i] as Paragraph;
                    if (paragraph != null) result.Add(paragraph);
                }
                return result;
            }

            /// <summary>计算表格网格列数（各行跨度和的最大值，至少 1）。 Computes the grid column count (the maximum per-row span sum, at least 1).</summary>
            /// <param name="model">表格模型。 Table model.</param>
            /// <param name="rows">行列表。 Rows.</param>
            /// <returns>网格列数。 The grid column count.</returns>
            private static int ComputeGridColumns(Table model, List<TableRow> rows)
            {
                int gridCols = model.Columns != null ? model.Columns.Count : 0;
                for (int r = 0; r < rows.Count; r++)
                {
                    TableRow row = rows[r];
                    if (row == null || row.Cells == null) continue;
                    int sum = 0;
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        TableCell cell = row.Cells[c];
                        sum += cell != null && cell.ColSpan > 1 ? cell.ColSpan : 1;
                    }
                    if (sum > gridCols) gridCols = sum;
                }
                return Math.Max(1, gridCols);
            }

            /// <summary>生成 <code>a:tblGrid</code>：列宽按模型列宽比例分摊总宽度，末列补足取整误差。 Builds <code>a:tblGrid</code>: widths shared by the column proportions, with rounding drift absorbed by the last column.</summary>
            /// <param name="model">表格模型。 Table model.</param>
            /// <param name="rows">行列表（列定义缺失时据此推断）。 Rows (used when column definitions are missing).</param>
            /// <param name="gridCols">网格列数。 Grid column count.</param>
            /// <param name="totalEmu">表格总宽度（EMU）。 Total table width in EMU.</param>
            /// <returns>tblGrid 元素。 The tblGrid element.</returns>
            private static XElement CreateGrid(Table model, List<TableRow> rows, int gridCols, int totalEmu)
            {
                double[] weights = new double[gridCols];
                double sum = 0;
                for (int c = 0; c < gridCols; c++)
                {
                    double weight = 1.0;
                    if (model.Columns != null && c < model.Columns.Count && model.Columns[c] != null && model.Columns[c].Width > 0)
                    {
                        weight = model.Columns[c].Width;
                    }
                    weights[c] = weight;
                    sum += weight;
                }
                XElement grid = new XElement(A + "tblGrid");
                int used = 0;
                for (int c = 0; c < gridCols; c++)
                {
                    int width;
                    if (c == gridCols - 1)
                    {
                        width = totalEmu - used;
                    }
                    else
                    {
                        width = (int)RoundToInt((double)totalEmu * (weights[c] / sum));
                        used += width;
                    }
                    grid.Add(new XElement(A + "gridCol", new XAttribute("w", IntStr(Math.Max(1, width)))));
                }
                return grid;
            }

            /// <summary>
            /// 判断调用者是否真的声明过表头：某一行带了 IsHeader，或者给过表头底色。
            /// 两者都没有时一律不立 firstRow——「首行长得像表头」是写端自己的假设，不是模型里的意图；
            /// 旧实现让每张表都当自己的首行是表头，于是没人要过的加粗白字与条带跟着进了文件。
            /// Determines whether the caller actually declared a header: some row carries IsHeader, or a
            /// header colour was set. With neither, firstRow stays off — "the first row looks like a header"
            /// is the writer's own assumption, not intent held in the model; the old implementation treated
            /// every table's first row as a header and shipped bold white text and banding nobody asked for.
            /// </summary>
            /// <param name="table">表格形状。 Table shape.</param>
            /// <param name="rows">行列表。 Rows.</param>
            /// <returns>声明过返回 true。 True when a header was declared.</returns>
            private static bool HasDeclaredHeader(TableShape table, List<TableRow> rows)
            {
                if (table != null && table.HeaderColor.HasValue) return true;
                if (rows == null) return false;
                for (int r = 0; r < rows.Count; r++)
                {
                    if (rows[r] != null && rows[r].IsHeader) return true;
                }
                return false;
            }

            /// <summary>
            /// 判断指定行是否按表头样式渲染。条件已经由 <see cref="HasDeclaredHeader"/> 把过一遍：
            /// 走到这里说明调用者声明过表头，显式的行优先，没有显式行时把首行当表头。
            /// Determines whether a row is rendered with the header style. The caller has already gated this
            /// on <see cref="HasDeclaredHeader"/>: reaching here means a header was declared, an explicitly
            /// marked row wins, and otherwise the first row carries the header treatment.
            /// </summary>
            /// <param name="model">表格模型。 Table model.</param>
            /// <param name="rows">行列表。 Rows.</param>
            /// <param name="index">行下标。 Row index.</param>
            /// <returns>按表头渲染返回 true。 True when styled as a header.</returns>
            private static bool IsHeaderRowAt(Table model, List<TableRow> rows, int index)
            {
                if (rows[index] != null && rows[index].IsHeader) return true;
                bool anyExplicit = false;
                for (int r = 0; r < rows.Count; r++)
                {
                    if (rows[r] != null && rows[r].IsHeader) { anyExplicit = true; break; }
                }
                if (!anyExplicit && index == 0) return true;
                return false;
            }

            /// <summary>按行数均摊表格高度得到默认行高（磅）。 Derives the default row height in points by sharing the table height across rows.</summary>
            /// <param name="tableHeight">表格总高度（磅）。 Total table height in points.</param>
            /// <param name="rowCount">行数。 Row count.</param>
            /// <returns>每行高度（磅）。 The per-row height in points.</returns>
            private static double DefaultRowHeight(double tableHeight, int rowCount)
            {
                if (rowCount <= 0 || tableHeight <= 0) return 18.0;
                double height = tableHeight / rowCount;
                return height < 6 ? 6 : height;
            }

            /// <summary>取得主题强调色（越界时回退默认 Office 强调色）。 Gets a theme accent color, falling back to the default Office palette when out of range.</summary>
            /// <param name="index">0 基序号。 Zero-based index.</param>
            /// <returns>强调色。 The accent color.</returns>
            private OfficeColor AccentColor(int index)
            {
                List<OfficeColor> colors = _p.AccentColors;
                if (colors != null && index >= 0 && index < colors.Count) return colors[index];
                OfficeColor color;
                if (!OfficeColor.TryParse(DefaultAccent(index), out color)) color = OfficeColor.Blue;
                return color;
            }

            /// <summary>默认 Office 强调色的十六进制文本表。 Table of the default Office accent colors as hex text.</summary>
            /// <param name="index">0 基序号。 Zero-based index.</param>
            /// <returns>RRGGBB 文本。 The RRGGBB text.</returns>
            private static string DefaultAccent(int index)
            {
                string[] defaults = { "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47" };
                if (index < 0 || index >= defaults.Length) return defaults[0];
                return defaults[index];
            }

            /// <summary>生成母版 <code>p:sldMaster</code>：含 title/body 两个占位符、clrMap 与 sldLayoutIdLst。 Builds the master <code>p:sldMaster</code>: title/body placeholders, clrMap and sldLayoutIdLst.</summary>
            /// <returns>母版文档。 The master document.</returns>
            private XDocument CreateMaster()
            {
                XElement tree = RelationshipRoot(P + "sldMaster",
                    new XElement(P + "cSld", CreatePlaceholderTree(SlideLayoutKind.TitleAndContent, true)));
                XElement clrMap = new XElement(P + "clrMap",
                    new XAttribute("bg1", "lt1"), new XAttribute("tx1", "dk1"),
                    new XAttribute("bg2", "lt2"), new XAttribute("tx2", "dk2"),
                    new XAttribute("accent1", "accent1"), new XAttribute("accent2", "accent2"),
                    new XAttribute("accent3", "accent3"), new XAttribute("accent4", "accent4"),
                    new XAttribute("accent5", "accent5"), new XAttribute("accent6", "accent6"),
                    new XAttribute("hlink", "hlink"), new XAttribute("folHlink", "folHlink"));
                tree.Add(clrMap);
                XElement layoutIdLst = new XElement(P + "sldLayoutIdLst");
                for (int i = 0; i < _layoutKinds.Count; i++)
                {
                    layoutIdLst.Add(new XElement(P + "sldLayoutId",
                        new XAttribute("id", (2147483649L + i).ToString(CultureInfo.InvariantCulture)),
                        new XAttribute(R + "id", "rId" + IntStr(i + 1))));
                }
                tree.Add(layoutIdLst);
                return WithDeclaration(Declare(tree, true));
            }

            /// <summary>生成母版的关系：各版式与主题。 Builds the master relationships: every layout plus the theme.</summary>
            /// <returns>关系文档。 The relationships document.</returns>
            private XDocument CreateMasterRelationships()
            {
                List<RelEntry> rels = new List<RelEntry>();
                for (int i = 0; i < _layoutKinds.Count; i++)
                {
                    rels.Add(new RelEntry("rId" + IntStr(i + 1), RelSlideLayout, "../slideLayouts/slideLayout" + IntStr(i + 1) + ".xml"));
                }
                rels.Add(new RelEntry("rId" + IntStr(_layoutKinds.Count + 1), RelTheme, "../theme/theme1.xml"));
                return BuildRels(rels);
            }

            /// <summary>为母版或版式生成占位符形状树（title 与 body/content）。 Builds the placeholder shape tree for the master or a layout (title plus body/content).</summary>
            /// <param name="kind">版式类别（母版按 TitleAndContent 处理）。 Layout kind (the master is treated as TitleAndContent).</param>
            /// <param name="isMaster">是否为母版（母版 body 用 type=body，版式随类别）。 Whether it is the master (master body uses type=body; layouts follow the kind).</param>
            /// <returns>p:spTree 元素。 The spTree element.</returns>
            private XElement CreatePlaceholderTree(SlideLayoutKind kind, bool isMaster)
            {
                XElement tree = CreateSpTree();
                List<PlaceholderGeometry> list = LayoutGeometry(kind, isMaster);
                for (int i = 0; i < list.Count; i++)
                {
                    tree.Add(CreatePlaceholderSp(list[i], isMaster));
                }
                return tree;
            }

            /// <summary>生成一个占位符 <code>p:sp</code>（含 xfrm 与空文本体）。 Builds a placeholder <code>p:sp</code> with its xfrm and an empty text body.</summary>
            /// <param name="geo">占位符几何。 Placeholder geometry.</param>
            /// <param name="isMaster">是否为母版占位符。 Whether it is a master placeholder.</param>
            /// <returns>p:sp 元素。 The sp element.</returns>
            private XElement CreatePlaceholderSp(PlaceholderGeometry geo, bool isMaster)
            {
                XElement ph = new XElement(P + "ph", new XAttribute("type", geo.Type));
                if (geo.Kind != PlaceholderKind.Title && geo.Index >= 0) ph.Add(new XAttribute("idx", IntStr(geo.Index)));
                XElement nvPr = new XElement(P + "nvPr", ph);
                XElement sp = new XElement(P + "sp",
                    new XElement(P + "nvSpPr",
                        new XElement(P + "cNvPr", new XAttribute("id", IntStr(2 + geo.Index)), new XAttribute("name", geo.Name)),
                        new XElement(P + "cNvSpPr", new XAttribute("spcc", "1")),
                        nvPr),
                    new XElement(P + "spPr",
                        new XElement(A + "xfrm",
                            new XElement(A + "off", new XAttribute("x", IntStr(Ppt.Emu(geo.Left))), new XAttribute("y", IntStr(Ppt.Emu(geo.Top)))),
                            new XElement(A + "ext", new XAttribute("cx", IntStr(Ppt.Emu(geo.Width))), new XAttribute("cy", IntStr(Ppt.Emu(geo.Height))))),
                        new XElement(A + "prstGeom", new XAttribute("prst", "rect"), new XElement(A + "avLst"))),
                    new XElement(P + "txBody",
                        new XElement(A + "bodyPr", new XAttribute("anchor", geo.Kind == PlaceholderKind.Title ? "ctr" : "t")),
                        new XElement(A + "lstStyle"),
                        new XElement(A + "p", new XElement(A + "endParaRPr", new XAttribute("lang", "zh-CN")))));
                return sp;
            }

            /// <summary>按版式类别推算占位符几何，比例与 <c>BuildPlaceholders</c> 一致。 Derives the placeholder geometry for a layout kind using the same proportions as <c>BuildPlaceholders</c>.</summary>
            /// <param name="kind">版式类别。 Layout kind.</param>
            /// <param name="isMaster">是否母版（母版 content 记为 body）。 Whether it is the master (its content placeholder is typed body).</param>
            /// <returns>占位符几何列表。 The placeholder geometry list.</returns>
            private List<PlaceholderGeometry> LayoutGeometry(SlideLayoutKind kind, bool isMaster)
            {
                List<PlaceholderGeometry> list = new List<PlaceholderGeometry>();
                double marginX = _w * 0.08;
                double usable = _w - marginX * 2.0;
                double titleTop = kind == SlideLayoutKind.SectionHeader ? _h * 0.34 : _h * 0.06;
                double titleHeight = kind == SlideLayoutKind.SectionHeader ? _h * 0.26 : _h * 0.20;
                if (kind != SlideLayoutKind.Blank)
                {
                    list.Add(new PlaceholderGeometry("title", PlaceholderKind.Title, 0, "Title 1", marginX, titleTop, usable, titleHeight));
                }
                if (kind == SlideLayoutKind.Title || kind == SlideLayoutKind.Blank || kind == SlideLayoutKind.SectionHeader)
                {
                    return list;
                }
                if (kind == SlideLayoutKind.TwoContent)
                {
                    list.Add(new PlaceholderGeometry(isMaster ? "body" : "obj", PlaceholderKind.Content, 1, "Content Placeholder 2", marginX, _h * 0.30, usable / 2.0 - 9.0, _h * 0.58));
                    list.Add(new PlaceholderGeometry("obj", PlaceholderKind.Content, 2, "Content Placeholder 3", marginX + usable / 2.0 + 9.0, _h * 0.30, usable / 2.0 - 9.0, _h * 0.58));
                    return list;
                }
                if (isMaster)
                {
                    list.Add(new PlaceholderGeometry("body", PlaceholderKind.Body, 1, "Content Placeholder 2", marginX, _h * 0.30, usable, _h * 0.58));
                    return list;
                }
                string type = "obj";
                string name = "Content Placeholder 2";
                PlaceholderKind pk = PlaceholderKind.Content;
                if (kind == SlideLayoutKind.TitleAndPicture) { type = "pic"; name = "Picture Placeholder 2"; pk = PlaceholderKind.Picture; }
                if (kind == SlideLayoutKind.TitleAndTable) { name = "Table Placeholder 2"; }
                list.Add(new PlaceholderGeometry(type, pk, 1, name, marginX, _h * 0.30, usable, _h * 0.58));
                return list;
            }

            /// <summary>版式类别的名称文本。 The display name of a layout kind.</summary>
            /// <param name="kind">版式类别。 Layout kind.</param>
            /// <returns>名称。 The name.</returns>
            private static string LayoutName(SlideLayoutKind kind)
            {
                switch (kind)
                {
                    case SlideLayoutKind.Title: return "标题幻灯片";
                    case SlideLayoutKind.TitleAndContent: return "标题和内容";
                    case SlideLayoutKind.SectionHeader: return "节标题";
                    case SlideLayoutKind.TwoContent: return "两栏内容";
                    case SlideLayoutKind.TitleAndPicture: return "标题和内容（图片）";
                    case SlideLayoutKind.TitleAndTable: return "标题和内容（表格）";
                    default: return "空白";
                }
            }

            /// <summary>
            /// 版式类别到 ST_SlideLayoutType 的映射，写到 <c>p:sldLayout@type</c> 上。
            /// 这是版式意图唯一能被读端还原的地方：章节过渡页与「仅标题页」的形状完全一样（都只有一个标题占位符），
            /// 只靠形状推断永远认不出来。
            /// Maps a layout kind to ST_SlideLayoutType for <c>p:sldLayout@type</c>. This is the only place the
            /// layout intent survives a round trip: a section header and a title-only slide have identical
            /// shapes (a single title placeholder), so shape inference alone can never tell them apart.
            /// </summary>
            /// <param name="kind">版式类别。 Layout kind.</param>
            /// <returns>ST_SlideLayoutType 取值。 The ST_SlideLayoutType token.</returns>
            private static string LayoutTypeCode(SlideLayoutKind kind)
            {
                switch (kind)
                {
                    case SlideLayoutKind.Title: return "title";
                    case SlideLayoutKind.SectionHeader: return "secHead";
                    case SlideLayoutKind.TwoContent: return "twoColTx";
                    case SlideLayoutKind.TitleAndPicture: return "picTx";
                    case SlideLayoutKind.TitleAndTable: return "tbl";
                    case SlideLayoutKind.Blank: return "blank";
                    default: return "obj";
                }
            }

            /// <summary>生成最小合法的 <code>a:theme</code>（clrScheme/fontScheme/fmtScheme）。 Builds a minimal valid <code>a:theme</code> (clrScheme/fontScheme/fmtScheme).</summary>
            /// <returns>主题文档。 The theme document.</returns>
            private XDocument CreateTheme()
            {
                PresentationTheme theme = _p.Theme;
                XElement clrScheme = new XElement(A + "clrScheme", new XAttribute("name", theme != null && !string.IsNullOrEmpty(theme.Name) ? theme.Name : "Office"),
                    ColorPhase("dk1", PhaseColor("dk1", theme != null ? theme.Text : null, new XElement(A + "sysClr", new XAttribute("val", "windowText"), new XAttribute("lastClr", "000000")))),
                    ColorPhase("lt1", PhaseColor("lt1", theme != null ? theme.Background : null, new XElement(A + "sysClr", new XAttribute("val", "background1"), new XAttribute("lastClr", "FFFFFF")))),
                    ColorPhase("dk2", PhaseColor("dk2", theme != null ? theme.TextSecondary : null, Srgb("44546A"))),
                    ColorPhase("lt2", PhaseColor("lt2", theme != null ? theme.BackgroundSecondary : null, Srgb("E7E6E6"))),
                    ColorPhase("accent1", Srgb(AccentHex(0))),
                    ColorPhase("accent2", Srgb(AccentHex(1))),
                    ColorPhase("accent3", Srgb(AccentHex(2))),
                    ColorPhase("accent4", Srgb(AccentHex(3))),
                    ColorPhase("accent5", Srgb(AccentHex(4))),
                    ColorPhase("accent6", Srgb(AccentHex(5))),
                    ColorPhase("hlink", PhaseColor("hlink", theme != null ? theme.Hyperlink : null, Srgb("0563C1"))),
                    ColorPhase("folHlink", PhaseColor("folHlink", theme != null ? theme.FollowedHyperlink : null, Srgb("954F72"))));

                string major = SafeFamily(theme != null && !string.IsNullOrEmpty(theme.FontMajor)
                    ? theme.FontMajor
                    : _titleFormat.GetAsciiFontFamily());
                string minor = SafeFamily(theme != null && !string.IsNullOrEmpty(theme.FontMinor)
                    ? theme.FontMinor
                    : _bodyFormat.GetAsciiFontFamily());
                string schemeName = theme != null && !string.IsNullOrEmpty(theme.Name) ? theme.Name : "Office";
                string themeName = theme != null && !string.IsNullOrEmpty(theme.Name) ? theme.Name : "Office Theme";
                XElement fontScheme = new XElement(A + "fontScheme", new XAttribute("name", schemeName),
                    new XElement(A + "majorFont",
                        new XElement(A + "latin", new XAttribute("typeface", major)),
                        new XElement(A + "ea", new XAttribute("typeface", major)),
                        new XElement(A + "cs", new XAttribute("typeface", major))),
                    new XElement(A + "minorFont",
                        new XElement(A + "latin", new XAttribute("typeface", minor)),
                        new XElement(A + "ea", new XAttribute("typeface", minor)),
                        new XElement(A + "cs", new XAttribute("typeface", minor))));

                XElement fillStyles = new XElement(A + "fillStyleLst",
                    SolidStyle(AccentHex(0), "FFFFFF"), SolidStyle("000000", "FFFFFF"), SolidStyle(AccentHex(4), "FFFFFF"));
                XElement lineStyles = new XElement(A + "lnStyleLst",
                    LineStyle(6350, "A0A0A0"), LineStyle(12700, "A0A0A0"), LineStyle(19050, "A0A0A0"));
                XElement effectStyles = new XElement(A + "effectStyleLst",
                    new XElement(A + "effectStyle", new XElement(A + "effectLst")),
                    new XElement(A + "effectStyle", new XElement(A + "effectLst")),
                    new XElement(A + "effectStyle", new XElement(A + "effectLst")));
                XElement bgStyles = new XElement(A + "bgFillStyleLst",
                    SolidStyle("FFFFFF", "FFFFFF"), SolidStyle("FFFFFF", "FFFFFF"), SolidStyle("FFFFFF", "FFFFFF"));

                XElement themeElements = new XElement(A + "themeElements",
                    new XAttribute("name", themeName), clrScheme, fontScheme,
                    new XElement(A + "fmtScheme", new XAttribute("name", schemeName), fillStyles, lineStyles, effectStyles, bgStyles));
                XElement root = new XElement(A + "theme",
                    new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                    new XAttribute("name", themeName), themeElements);
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>生成主题色彩阶段元素 <code>a:&lt;name&gt;</code>。 Builds a theme color-phase element <code>a:&lt;name&gt;</code>.</summary>
            /// <param name="name">阶段名。 Phase name.</param>
            /// <param name="child">颜色值子元素。 Color-value child.</param>
            /// <returns>阶段元素。 The phase element.</returns>
            private static XElement ColorPhase(string name, XElement child)
            {
                return new XElement(A + name, child);
            }

            /// <summary>主题阶段的颜色值：未覆盖时返回默认子元素（dk1/lt1 保持 sysClr，与历史输出一致），
            /// 覆盖后统一写 <code>a:srgbClr</code> 常量；覆盖值恰等于 sysClr 的 lastClr 时仍写回 sysClr，
            /// 这样读入内置默认色再存盘不会改变 theme1.xml。母版与版式经 clrMap 的 bg1/tx1 映射自动跟随该色。
            /// The phase colour value: the default child when not overridden (dk1/lt1 keep their sysClr, matching
            /// historic output), an <code>a:srgbClr</code> constant once overridden; an override that equals the
            /// sysClr lastClr is written back as sysClr, so loading a built-in default colour and resaving leaves
            /// theme1.xml unchanged. The master and layouts follow the new colour through the clrMap bg1/tx1 mapping.</summary>
            /// <param name="phase">阶段名（dk1/lt1/dk2/lt2/hlink/folHlink）。 Phase name.</param>
            /// <param name="color">覆盖色，null 表示沿用默认值。 The override, or null to keep the default.</param>
            /// <param name="defaultChild">默认颜色子元素。 The default colour child.</param>
            /// <returns>颜色值子元素。 The colour-value child.</returns>
            private static XElement PhaseColor(string phase, OfficeColor? color, XElement defaultChild)
            {
                if (color == null) return defaultChild;
                string hex = color.Value.ToHex();
                if (defaultChild != null && defaultChild.Name == A + "sysClr"
                    && string.Equals((string)defaultChild.Attribute("lastClr"), hex, StringComparison.OrdinalIgnoreCase))
                {
                    return defaultChild;
                }
                return Srgb(hex);
            }

            /// <summary>生成 <code>a:srgbClr</code> 值元素。 Builds an <code>a:srgbClr</code> value element.</summary>
            /// <param name="hex">RRGGBB 十六进制。 The RRGGBB hex.</param>
            /// <returns>srgbClr 元素。 The srgbClr element.</returns>
            private static XElement Srgb(string hex)
            {
                return new XElement(A + "srgbClr", new XAttribute("val", hex));
            }

            /// <summary>取强调色的十六进制文本（越界回退默认）。 Gets the accent hex text, falling back to defaults out of range.</summary>
            /// <param name="index">0 基序号。 Zero-based index.</param>
            /// <returns>RRGGBB 文本。 The RRGGBB text.</returns>
            private string AccentHex(int index)
            {
                List<OfficeColor> colors = _p.AccentColors;
                if (colors != null && index >= 0 && index < colors.Count) return colors[index].ToHex();
                return DefaultAccent(index);
            }

            /// <summary>生成主题填充样式（solidFill）。 Builds a theme fill style (solidFill).</summary>
            /// <param name="hex">前景色十六进制。 Foreground hex.</param>
            /// <param name="unusedBg">保留参数，兼容签名。 Reserved for signature compatibility.</param>
            /// <returns>a:solidFill 元素。 The solidFill element.</returns>
            private static XElement SolidStyle(string hex, string unusedBg)
            {
                return new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", hex)));
            }

            /// <summary>生成主题线条样式（a:ln）。 Builds a theme line style (a:ln).</summary>
            /// <param name="widthEmu">线宽（EMU）。 Line width in EMU.</param>
            /// <param name="hex">颜色十六进制。 Color hex.</param>
            /// <returns>a:ln 元素。 The ln element.</returns>
            private static XElement LineStyle(int widthEmu, string hex)
            {
                return new XElement(A + "ln",
                    new XAttribute("w", IntStr(widthEmu)),
                    new XAttribute("cap", "flat"), new XAttribute("cmpd", "sng"), new XAttribute("algn", "ctr"),
                    new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", hex))),
                    new XElement(A + "prstDash", new XAttribute("val", "solid")),
                    new XElement(A + "miter", new XAttribute("lim", "800000")));
            }

            /// <summary>生成 <code>p:notesMaster</code>：最小形状树与 clrMap。 Builds <code>p:notesMaster</code>: a minimal shape tree and clrMap.</summary>
            /// <returns>备注母版文档。 The notes-master document.</returns>
            private XDocument CreateNotesMaster()
            {
                XElement tree = RelationshipRoot(P + "notesMaster",
                    new XElement(P + "cSld",
                        new XElement(P + "spTree",
                            new XElement(P + "nvGrpSpPr",
                                new XElement(P + "cNvPr", new XAttribute("id", "1"), new XAttribute("name", string.Empty)),
                                new XElement(P + "cNvGrpSpPr"),
                                new XElement(P + "nvPr")),
                            new XElement(P + "grpSpPr",
                                new XElement(A + "xfrm",
                                    new XElement(A + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                    new XElement(A + "ext", new XAttribute("cx", "6858000"), new XAttribute("cy", "9144000")),
                                    new XElement(A + "chOff", new XAttribute("x", "0"), new XAttribute("y", "0")),
                                    new XElement(A + "chExt", new XAttribute("cx", "6858000"), new XAttribute("cy", "9144000")))))));
                tree.Add(new XElement(P + "clrMap",
                    new XAttribute("bg1", "lt1"), new XAttribute("tx1", "dk1"),
                    new XAttribute("bg2", "lt2"), new XAttribute("tx2", "dk2"),
                    new XAttribute("accent1", "accent1"), new XAttribute("accent2", "accent2"),
                    new XAttribute("accent3", "accent3"), new XAttribute("accent4", "accent4"),
                    new XAttribute("accent5", "accent5"), new XAttribute("accent6", "accent6"),
                    new XAttribute("hlink", "hlink"), new XAttribute("folHlink", "folHlink")));
                return WithDeclaration(Declare(tree, true));
            }

            /// <summary>生成备注母版的关系（指向主题）。 Builds the notes-master relationship pointing at the theme.</summary>
            /// <returns>关系文档。 The relationships document.</returns>
            private XDocument CreateNotesMasterRelationships()
            {
                return BuildRels(new List<RelEntry> { new RelEntry("rId1", RelTheme, "../theme/theme1.xml") });
            }

            /// <summary>
            /// 按幻灯片生成切换效果部件；无切换效果返回 null。时长不在 ISO/IEC 29500 的 CT_SlideTransition 里，
            /// 因此按微软文档的做法用 mc:AlternateContent 包一份带 p14:dur 的 Choice 与一份不带时长的 Fallback。
            /// Builds a slide's transition part; returns null when there is no transition. The duration is not part of
            /// CT_SlideTransition in ISO/IEC 29500, so, as Microsoft documents it, an mc:AlternateContent wraps a
            /// mc:Choice carrying p14:dur and an mc:Fallback without it.
            /// </summary>
            /// <param name="slide">幻灯片。 Slide.</param>
            /// <returns>transition 或 AlternateContent 元素，无切换时 null。 The transition or AlternateContent element, null without a transition.</returns>
            private static XElement CreateTransition(Slide slide)
            {
                string speed = slide.TransitionSpeed.HasValue ? SpeedCode(slide.TransitionSpeed.Value) : "med";
                string click = slide.AdvanceOnClick ? "1" : "0";
                XElement effect = null;
                switch (slide.Transition)
                {
                    case SlideTransitionKind.Fade: effect = new XElement(P + "fade"); break;
                    case SlideTransitionKind.Push: effect = new XElement(P + "push", new XAttribute("dir", DirectionCode(slide.TransitionDirection, SlideTransitionDirection.Left))); break;
                    case SlideTransitionKind.Wipe: effect = new XElement(P + "wipe", new XAttribute("dir", DirectionCode(slide.TransitionDirection, SlideTransitionDirection.Down))); break;
                    case SlideTransitionKind.Split: effect = new XElement(P + "split", new XAttribute("orient", "vert"), new XAttribute("dir", DirectionCode(slide.TransitionDirection, SlideTransitionDirection.Down))); break;
                    default: break;
                }
                if (effect == null)
                {
                    // 无切换效果但设了自动换页：写不带效果节点的 transition，仅由 advTm 表达计时。
                    // No visual effect but auto-advance set: a transition without an effect child, carried by advTm alone.
                    if (slide.AdvanceAfterSeconds <= 0) return null;
                    XElement idle = new XElement(P + "transition",
                        new XAttribute("advTm", IntStr(RoundToInt(slide.AdvanceAfterSeconds * 1000.0))));
                    if (!slide.AdvanceOnClick) idle.Add(new XAttribute("advClick", click));
                    return idle;
                }
                string advTm = slide.AdvanceAfterSeconds > 0
                    ? IntStr(RoundToInt(slide.AdvanceAfterSeconds * 1000.0))
                    : null;
                XElement transition = new XElement(P + "transition");
                if (slide.TransitionDuration <= 0)
                {
                    if (advTm != null) transition.Add(new XAttribute("advTm", advTm));
                    if (slide.TransitionSpeed.HasValue) transition.Add(new XAttribute("spd", speed));
                    if (!slide.AdvanceOnClick) transition.Add(new XAttribute("advClick", click));
                    transition.Add(new XElement(effect));
                    return transition;
                }
                string dur = IntStr(RoundToInt(slide.TransitionDuration * 1000.0));
                XElement choiceTransition = new XElement(P + "transition",
                    new XAttribute("spd", speed),
                    new XAttribute(P14 + "dur", dur),
                    new XAttribute("advClick", click),
                    new XElement(effect));
                XElement fallbackTransition = new XElement(P + "transition",
                    new XAttribute("spd", speed),
                    new XAttribute("advClick", click),
                    new XElement(effect));
                if (advTm != null)
                {
                    choiceTransition.Add(new XAttribute("advTm", advTm));
                    fallbackTransition.Add(new XAttribute("advTm", advTm));
                }
                XElement choice = new XElement(Mc + "Choice",
                    new XAttribute("Requires", "p14"),
                    choiceTransition);
                XElement fallback = new XElement(Mc + "Fallback",
                    fallbackTransition);
                return new XElement(Mc + "AlternateContent",
                    new XAttribute(XNamespace.Xmlns + "mc", Mc.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "p14", P14.NamespaceName),
                    choice,
                    fallback);
            }

            /// <summary>把切换方向映射成 <c>dir</c> 的 l/r/u/d；未设置时用该效果的原生默认方向。Maps a transition direction to dir's l/r/u/d, falling back to the effect's native default.</summary>
            /// <param name="direction">模型里的方向，可为 null。The direction from the model, possibly null.</param>
            /// <param name="fallback">未设置时的默认方向。The default direction when unset.</param>
            /// <returns>l、r、u 或 d。One of l, r, u or d.</returns>
            private static string DirectionCode(SlideTransitionDirection? direction, SlideTransitionDirection fallback)
            {
                switch (direction ?? fallback)
                {
                    case SlideTransitionDirection.Right: return "r";
                    case SlideTransitionDirection.Up: return "u";
                    case SlideTransitionDirection.Down: return "d";
                    default: return "l";
                }
            }

            /// <summary>把切换速度映射成 <c>spd</c> 的 slow/med/fast。Maps a transition speed to spd's slow/med/fast.</summary>
            /// <param name="speed">速度。The speed.</param>
            /// <returns>slow、med 或 fast。One of slow, med or fast.</returns>
            private static string SpeedCode(SlideTransitionSpeed speed)
            {
                switch (speed)
                {
                    case SlideTransitionSpeed.Slow: return "slow";
                    case SlideTransitionSpeed.Fast: return "fast";
                    default: return "med";
                }
            }

            /// <summary>生成 <code>p:timing</code> 动画时间线：每个带动画的形状是一个顺序节点，按 z 序排列，
            /// 延迟在时间线上累计（与 PowerPoint「单击时」序列语义一致）；设置了自动切换的幻灯片把首个
            /// 节点的触发条件从 <c>st="indef"</c> 换成 <c>st="afterEvent"</c> 加累计毫秒延迟。
            /// Builds the <code>p:timing</code> timeline: one sequence node per animated shape in z-order, with
            /// delays accumulating along the timeline (matching PowerPoint's "On Click" semantics). When the
            /// slide advances automatically the first node's trigger becomes <c>st="afterEvent"</c> plus the
            /// accumulated delay in milliseconds instead of <c>st="indef"</c>.</summary>
            /// <param name="ids">带动画形状已写出的 cNvPr id。 Written cNvPr ids of the animated shapes.</param>
            /// <param name="shapes">与 id 对应的形状。 The shapes matching those ids.</param>
            /// <param name="slide">所在幻灯片。 The owning slide.</param>
            /// <returns>timing 元素；没有动画形状时为 null（自动切换由 p:transition 的 advTm 表达）。 The timing
            /// element, or null without animated shapes (auto-advance alone is carried by p:transition's advTm).</returns>
            private static XElement CreateTiming(List<int> ids, List<SlideShape> shapes, Slide slide)
            {
                if (shapes == null || shapes.Count == 0) return null;
                bool autoAdvance = slide.AdvanceAfterSeconds > 0;
                // id 全局唯一：过去根节点与主序列节点都用 "1"，严格的 CT_TLTimeNode 校验直接判重。
                // Ids are unique across the tree; the root and the main sequence used to share "1", which a
                // strict CT_TLTimeNode check rejects as a duplicate.
                int nextId = 3;
                XElement groups = new XElement(P + "childTnLst");
                for (int i = 0; i < shapes.Count; i++)
                {
                    SlideShape shape = shapes[i];
                    int delay = shape.AnimationDelaySeconds > 0 ? RoundToInt(shape.AnimationDelaySeconds * 1000.0) : 0;
                    bool autoFirst = i == 0 && autoAdvance;

                    int groupId = nextId++;
                    int delayId = nextId++;
                    int effectId = nextId++;
                    int setId = nextId++;
                    int behaviorId = nextId++;
                    XElement effect = CreateEffectNode(shape, ids[i], effectId, setId, behaviorId,
                        autoFirst ? "afterEffect" : "clickEffect");

                    // 三层：触发节点（单击=indefinite / 自动=0）→ 延迟节点（数值毫秒）→ 效果节点。
                    // PowerPoint 对「单击后延迟」正是这样写：延迟单独占一个 cTn，挂在触发节点与效果节点之间。
                    // 每一层都用 par > cTn 包装，因为 CT_TLTimeNodeList 只接受 par/seq，裸 cTn 不是合法子元素。
                    // Three levels: the trigger node (indefinite on click, 0 for automatic), the delay node
                    // (milliseconds), and the effect node. That is exactly how PowerPoint encodes "delay after
                    // click": the delay gets its own cTn between the trigger and the effect. Each level is
                    // wrapped in par > cTn because CT_TLTimeNodeList only accepts par/seq, never a bare cTn.
                    XElement delayNode = new XElement(P + "cTn",
                        new XAttribute("id", IntStr(delayId)),
                        new XAttribute("fill", "hold"),
                        new XElement(P + "stCondLst",
                            new XElement(P + "cond", new XAttribute("delay", IntStr(delay)))),
                        new XElement(P + "childTnLst", new XElement(P + "par", effect)));
                    XElement group = new XElement(P + "cTn",
                        new XAttribute("id", IntStr(groupId)),
                        new XAttribute("fill", "hold"),
                        new XElement(P + "stCondLst",
                            new XElement(P + "cond", new XAttribute("delay", autoFirst ? "0" : "indefinite"))),
                        new XElement(P + "childTnLst", new XElement(P + "par", delayNode)));
                    groups.Add(new XElement(P + "par", group));
                }

                XElement seq = new XElement(P + "seq",
                    new XAttribute("concurrent", "1"),
                    new XAttribute("nextAc", "seek"),
                    new XElement(P + "cTn",
                        new XAttribute("id", "2"),
                        new XAttribute("dur", "indefinite"),
                        new XAttribute("nodeType", "mainSeq"),
                        groups),
                    new XElement(P + "prevCondLst",
                        new XElement(P + "cond",
                            new XAttribute("evt", "onPrev"),
                            new XAttribute("delay", "0"),
                            new XElement(P + "tgtEl", new XElement(P + "sldTgt")))),
                    new XElement(P + "nextCondLst",
                        new XElement(P + "cond",
                            new XAttribute("evt", "onNext"),
                            new XAttribute("delay", "0"),
                            new XElement(P + "tgtEl", new XElement(P + "sldTgt")))));
                // 根节点：nodeType="tmRoot" 加 childTnLst，seq 必须住在它里面而不是与它平级。
                // The root node: nodeType="tmRoot" with a childTnLst, and seq lives inside it rather than
                // sitting beside it.
                return new XElement(P + "timing",
                    new XElement(P + "tnLst",
                        new XElement(P + "par",
                            new XElement(P + "cTn",
                                new XAttribute("id", "1"),
                                new XAttribute("dur", "indefinite"),
                                new XAttribute("restart", "never"),
                                new XAttribute("nodeType", "tmRoot"),
                                new XElement(P + "childTnLst", seq)))));
            }

            /// <summary>
            /// 生成单个形状的入场行为 <code>p:cTn</code>：preset 属性、子节点列表（可见性 set 加效果行为）。
            /// Builds one shape's entrance behaviour <code>p:cTn</code>: the preset attributes and the child
            /// list holding the visibility set plus the effect behaviour.
            /// </summary>
            /// <param name="shape">带动画的形状。 The animated shape.</param>
            /// <param name="spid">形状的 cNvPr id。 The shape's cNvPr id.</param>
            /// <param name="effectId">行为节点的 id。 The behaviour node id.</param>
            /// <param name="setId">可见性 set 节点内 cTn 的 id。 The cTn id inside the visibility set.</param>
            /// <param name="behaviorId">效果行为内 cTn 的 id。 The cTn id inside the effect behaviour.</param>
            /// <param name="nodeType">clickEffect 或 afterEffect。 Either clickEffect or afterEffect.</param>
            /// <returns>p:cTn 行为元素。 The behaviour cTn element.</returns>
            private static XElement CreateEffectNode(SlideShape shape, int spid, int effectId, int setId,
                int behaviorId, string nodeType)
            {
                int duration = shape.AnimationDurationSeconds > 0
                    ? RoundToInt(shape.AnimationDurationSeconds * 1000.0)
                    : DefaultAnimationDuration(shape.Animation);
                string spidText = IntStr(spid);
                XElement children = new XElement(P + "childTnLst",
                    CreateVisibilitySet(spidText, setId),
                    CreateEffectBehavior(shape.Animation, spidText, behaviorId, duration));
                return new XElement(P + "cTn",
                    new XAttribute("id", IntStr(effectId)),
                    new XAttribute("presetID", AnimationPresetId(shape.Animation)),
                    new XAttribute("presetClass", "entr"),
                    new XAttribute("presetSubtype", AnimationPresetSubtype(shape.Animation)),
                    // dur 写在行为节点上：读端按 cTn/@dur 还原时长，PowerPoint 对入场效果也这样落。
                    // dur sits on the behaviour node: the reader restores the duration from cTn/@dur, and
                    // PowerPoint writes entrance effects the same way.
                    new XAttribute("dur", IntStr(duration)),
                    new XAttribute("fill", "hold"),
                    new XAttribute("nodeType", nodeType),
                    // 触发延迟记在外层组节点上（PowerPoint 就是这样写的），效果节点自己的 cond 恒为 0。
                    // The trigger delay lives on the outer group node (which is how PowerPoint writes it), and
                    // the effect node's own cond stays at zero.
                    new XElement(P + "stCondLst",
                        new XElement(P + "cond", new XAttribute("delay", "0"))),
                    children);
            }

            /// <summary>
            /// 生成入场前的可见性 <code>p:set</code>：cBhvr 里给 cTn、目标（<c>p:spTgt/@spid</c>）与
            /// <c>p:attrNameLst</c>，再由 <c>p:to/p:strVal</c> 给出 "visible"。
            /// 过去这里是 <c>p:setTo</c> 里塞一个 <c>a:clrMapOvr</c>，既不是 <c>p:set</c> 的内容模型，也没有意义。
            /// Builds the pre-entrance visibility <code>p:set</code>: a cBhvr carrying the cTn, the target
            /// (<c>p:spTgt/@spid</c>) and a <c>p:attrNameLst</c>, with <c>p:to/p:strVal</c> supplying "visible".
            /// It used to put an <c>a:clrMapOvr</c> inside a <c>p:setTo</c>, which is neither part of the
            /// <c>p:set</c> content model nor meaningful.
            /// </summary>
            /// <param name="spidText">形状 id 文本。 The shape id as text.</param>
            /// <param name="setId">set 内 cTn 的 id。 The cTn id inside the set.</param>
            /// <returns>p:set 元素。 The p:set element.</returns>
            private static XElement CreateVisibilitySet(string spidText, int setId)
            {
                return new XElement(P + "set",
                    new XElement(P + "cBhvr",
                        new XElement(P + "cTn",
                            new XAttribute("id", IntStr(setId)),
                            new XAttribute("dur", "1"),
                            new XAttribute("fill", "hold"),
                            new XElement(P + "stCondLst",
                                new XElement(P + "cond", new XAttribute("delay", "0")))),
                        new XElement(P + "tgtEl",
                            new XElement(P + "spTgt", new XAttribute("spid", spidText))),
                        new XElement(P + "attrNameLst",
                            new XElement(P + "attrName", "style.visibility"))),
                    new XElement(P + "to",
                        new XElement(P + "strVal", new XAttribute("val", "visible"))));
            }

            /// <summary>
            /// 生成效果行为元素（animEffect / animMotion / animScale），统一挂上 cBhvr 的 cTn 与目标。
            /// Split 用 <c>animEffect filter="barn(inVertical)"</c>——<c>p:bldP</c> 属于 <c>p:bldLst</c> 的构建清单，
            /// 不是时间节点，塞进 childTnLst 同样不合法。
            /// Builds the effect behaviour element (animEffect / animMotion / animScale) with the cBhvr's cTn
            /// and target. Split uses <c>animEffect filter="barn(inVertical)"</c> because <c>p:bldP</c> belongs
            /// to the <c>p:bldLst</c> build list rather than the time node list, so it is equally illegal here.
            /// </summary>
            /// <param name="kind">动画类别。 Animation kind.</param>
            /// <param name="spidText">形状 id 文本。 The shape id as text.</param>
            /// <param name="nodeId">行为 cTn 的 id。 The behaviour cTn id.</param>
            /// <param name="duration">时长（毫秒）。Duration in milliseconds.</param>
            /// <returns>行为元素。 The behaviour element.</returns>
            private static XElement CreateEffectBehavior(SlideAnimationKind kind, string spidText, int nodeId, int duration)
            {
                XElement behavior;
                switch (kind)
                {
                    case SlideAnimationKind.FlyIn:
                        behavior = new XElement(P + "animMotion",
                            new XAttribute("path", "M 0 0 L 0 0.02"),
                            new XAttribute("pathKind", "line"),
                            new XAttribute("ptsTypes", "VV"),
                            new XAttribute("filter", "shadow(w=0.75,h=0.75)"));
                        break;
                    case SlideAnimationKind.Zoom:
                        behavior = new XElement(P + "animScale",
                            new XAttribute("attrName", "r@x,r@y"),
                            new XElement(A + "from",
                                new XAttribute("x", "1000"), new XAttribute("y", "1000")),
                            new XElement(A + "to",
                                new XAttribute("x", "100000"), new XAttribute("y", "100000")));
                        break;
                    case SlideAnimationKind.Split:
                        behavior = new XElement(P + "animEffect",
                            new XAttribute("transition", "in"),
                            new XAttribute("filter", "barn(inVertical)"));
                        break;
                    case SlideAnimationKind.Wipe:
                        behavior = new XElement(P + "animEffect",
                            new XAttribute("transition", "in"),
                            new XAttribute("filter", "wipe(down)"));
                        break;
                    default:
                        behavior = new XElement(P + "animEffect",
                            new XAttribute("transition", "in"),
                            new XAttribute("filter", "fade"));
                        break;
                }
                behavior.Add(new XElement(P + "cBhvr",
                    new XElement(P + "cTn",
                        new XAttribute("id", IntStr(nodeId)),
                        new XAttribute("dur", IntStr(duration)),
                        new XElement(P + "stCondLst",
                            new XElement(P + "cond", new XAttribute("delay", "0")))),
                    new XElement(P + "tgtEl",
                        new XElement(P + "spTgt", new XAttribute("spid", spidText)))));
                return behavior;
            }

            /// <summary>取动画类别的 OOXML presetId。 Gets the OOXML presetId of the animation kind.</summary>
            /// <param name="kind">动画类别。 Animation kind.</param>
            /// <returns>presetId 文本。 The presetId text.</returns>
            private static string AnimationPresetId(SlideAnimationKind kind)
            {
                switch (kind)
                {
                    case SlideAnimationKind.Fade: return "10";
                    case SlideAnimationKind.FlyIn: return "2";
                    case SlideAnimationKind.Wipe: return "22";
                    case SlideAnimationKind.Zoom: return "23";
                    default: return "26";
                }
            }

            /// <summary>取动画类别的 OOXML presetSubtype（方向位）。 Gets the OOXML presetSubtype direction bits.</summary>
            /// <param name="kind">动画类别。 Animation kind.</param>
            /// <returns>presetSubtype 文本。 The presetSubtype text.</returns>
            private static string AnimationPresetSubtype(SlideAnimationKind kind)
            {
                switch (kind)
                {
                    case SlideAnimationKind.FlyIn: return "4";
                    case SlideAnimationKind.Wipe: return "4";
                    case SlideAnimationKind.Split: return "16";
                    default: return "0";
                }
            }

            /// <summary>取动画类别的默认时长毫秒。 Gets the default duration in milliseconds for the kind.</summary>
            /// <param name="kind">动画类别。 Animation kind.</param>
            /// <returns>默认毫秒数。 The default milliseconds.</returns>
            private static int DefaultAnimationDuration(SlideAnimationKind kind)
            {
                switch (kind)
                {
                    case SlideAnimationKind.FlyIn: return 1000;
                    case SlideAnimationKind.Zoom: return 1000;
                    default: return 500;
                }
            }

            /// <summary>生成 <code>docProps/core.xml</code>。 Builds <code>docProps/core.xml</code>.</summary>
            /// <returns>核心属性文档。 The core-properties document.</returns>
            private XDocument CreateCoreProperties()
            {
                DocumentMetadata metadata = _p.Metadata ?? new DocumentMetadata();
                XNamespace cp = Ppt.CoreProperties;
                XNamespace dc = Ppt.Dc;
                XNamespace dcterms = Ppt.DcTerms;
                XElement root = new XElement(cp + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp", cp.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "dc", dc.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "dcterms", dcterms.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "xsi", Xsi.NamespaceName),
                    new XElement(dc + "title", metadata.Title ?? string.Empty),
                    new XElement(dc + "subject", metadata.Subject ?? string.Empty),
                    new XElement(dc + "creator", metadata.Author ?? string.Empty),
                    new XElement(cp + "keywords", metadata.Keywords ?? string.Empty),
                    new XElement(cp + "lastModifiedBy", metadata.Author ?? string.Empty),
                    new XElement(dcterms + "created", new XAttribute(Xsi + "type", "dcterms:W3CDTF"), FormatTimestamp(metadata.Created)),
                    new XElement(dcterms + "modified", new XAttribute(Xsi + "type", "dcterms:W3CDTF"), FormatTimestamp(metadata.Modified)));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 生成 <code>docProps/app.xml</code>：应用名与幻灯片数。
            /// 元素次序按 ISO/IEC 29500 的 <c>CT_Properties</c> 序列给出（Company → Slides → Application → AppVersion），
            /// 而不是 PowerPoint 实测写出的那种「Application 提前」次序——后者不满足架构定义的 sequence，
            /// 严格校验会报错，而本部件只承载统计信息，按架构写不会有任何消费方损失。
            /// Builds <code>docProps/app.xml</code> with the application name and slide count.
            /// The elements follow the <c>CT_Properties</c> sequence in ISO/IEC 29500
            /// (Company → Slides → Application → AppVersion) rather than the "Application first" order
            /// PowerPoint actually emits; the latter violates the schema's sequence and fails strict
            /// validation, while this part only carries statistics, so following the schema costs no consumer
            /// anything.
            /// </summary>
            /// <returns>扩展属性文档。 The extended-properties document.</returns>
            private XDocument CreateAppProperties()
            {
                XNamespace ep = Ppt.ExtendedProperties;
                XElement root = new XElement(ep + "Properties",
                    new XAttribute(XNamespace.Xmlns + "ep", ep.NamespaceName),
                    new XElement(ep + "Company", string.Empty),
                    new XElement(ep + "Slides", IntStr(_slideCount)),
                    new XElement(ep + "Application", "Wisdom.Ppt"),
                    new XElement(ep + "AppVersion", "1.0"));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>判断扩展名是否已由固定的 Default 条目声明。 Reports whether a fixed Default entry already declares this extension.</summary>
            /// <param name="extension">不含点的扩展名。 Extension without the leading dot.</param>
            /// <returns>已声明返回 true。 True when it is already covered.</returns>
            private static bool IsKnownExtension(string extension)
            {
                return string.Equals(extension, "rels", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(extension, "xml", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(extension, "png", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(extension, "jpeg", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(extension, "jpg", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(extension, "gif", StringComparison.OrdinalIgnoreCase);
            }

            /// <summary>生成 <code>[Content_Types].xml</code>。 Builds <code>[Content_Types].xml</code>.</summary>
            /// <returns>内容类型文档。 The content-types document.</returns>
            private XDocument CreateContentTypes()
            {
                XElement root = new XElement(Ctypes + "Types",
                    new XElement(Ctypes + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", CtRels)),
                    new XElement(Ctypes + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", CtXml)),
                    new XElement(Ctypes + "Default", new XAttribute("Extension", "png"), new XAttribute("ContentType", CtPng)),
                    new XElement(Ctypes + "Default", new XAttribute("Extension", "jpeg"), new XAttribute("ContentType", CtJpeg)),
                    new XElement(Ctypes + "Default", new XAttribute("Extension", "jpg"), new XAttribute("ContentType", CtJpeg)),
                    new XElement(Ctypes + "Default", new XAttribute("Extension", "gif"), new XAttribute("ContentType", CtGif)),
                    new XElement(Ctypes + "Override", new XAttribute("PartName", "/ppt/presentation.xml"),
                        new XAttribute("ContentType", _template ? CtPresentationTemplate : CtPresentation)),
                    new XElement(Ctypes + "Override", new XAttribute("PartName", "/ppt/slideMasters/slideMaster1.xml"), new XAttribute("ContentType", CtSlideMaster)),
                    new XElement(Ctypes + "Override", new XAttribute("PartName", "/ppt/theme/theme1.xml"), new XAttribute("ContentType", CtTheme)),
                    new XElement(Ctypes + "Override", new XAttribute("PartName", "/ppt/notesMasters/notesMaster1.xml"), new XAttribute("ContentType", CtNotesMaster)),
                    new XElement(Ctypes + "Override", new XAttribute("PartName", "/docProps/core.xml"), new XAttribute("ContentType", CtCore)),
                    new XElement(Ctypes + "Override", new XAttribute("PartName", "/docProps/app.xml"), new XAttribute("ContentType", CtApp)));
                for (int i = 1; i <= _layoutKinds.Count; i++)
                {
                    root.Add(new XElement(Ctypes + "Override",
                        new XAttribute("PartName", "/ppt/slideLayouts/slideLayout" + IntStr(i) + ".xml"),
                        new XAttribute("ContentType", CtSlideLayout)));
                }
                for (int i = 1; i <= _slideCount; i++)
                {
                    root.Add(new XElement(Ctypes + "Override",
                        new XAttribute("PartName", "/ppt/slides/slide" + IntStr(i) + ".xml"),
                        new XAttribute("ContentType", CtSlide)));
                }
                for (int i = 1; i <= _notesCount; i++)
                {
                    root.Add(new XElement(Ctypes + "Override",
                        new XAttribute("PartName", "/ppt/notesSlides/notesSlide" + IntStr(i) + ".xml"),
                        new XAttribute("ContentType", CtNotesSlide)));
                }
                foreach (KeyValuePair<string, string> pair in _extraMediaTypes)
                {
                    // OPC 规定 Override 只带 PartName；按扩展名声明的类型必须写成 Default，
                    // 且不能与上面已有的固定扩展名重复。
                    // OPC allows Override elements to carry PartName only, so extension-based types have to be
                    // declared as Defaults, skipping the fixed extensions written above.
                    if (IsKnownExtension(pair.Key)) continue;
                    root.Add(new XElement(Ctypes + "Default", new XAttribute("Extension", pair.Key), new XAttribute("ContentType", pair.Value)));
                }
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>生成包级 <code>_rels/.rels</code>：主文档、核心属性与扩展属性。 Builds the package-level <code>_rels/.rels</code>: main document, core and extended properties.</summary>
            /// <returns>关系文档。 The relationships document.</returns>
            private XDocument CreateRootRelationships()
            {
                List<RelEntry> rels = new List<RelEntry>
                {
                    new RelEntry("rId1", RelOfficeDocument, "ppt/presentation.xml"),
                    new RelEntry("rId2", RelCoreProperties, "docProps/core.xml"),
                    new RelEntry("rId3", RelExtendedProperties, "docProps/app.xml")
                };
                return BuildRels(rels);
            }

            /// <summary>由关系条目列表构造 <code>Relationships</code> 文档。 Builds a <code>Relationships</code> document from relationship entries.</summary>
            /// <param name="entries">关系条目列表。 Relationship entries.</param>
            /// <returns>关系文档。 The relationships document.</returns>
            private static XDocument BuildRels(List<RelEntry> entries)
            {
                XElement root = new XElement(Rels + "Relationships");
                for (int i = 0; i < entries.Count; i++)
                {
                    RelEntry entry = entries[i];
                    root.Add(new XElement(Rels + "Relationship",
                        new XAttribute("Id", entry.Id),
                        new XAttribute("Type", entry.Type),
                        new XAttribute("Target", entry.Target)));
                }
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 为部件根元素声明 p/a/r 命名空间：把原始根替换为带 xmlns 声明的等价根，
            /// 避免 LINQ to XML 为每个 a: 元素补写冗余的 xmlns 属性。
            /// Declares the p/a/r namespaces on a part root: replaces the raw root with an equivalent root
            /// carrying the xmlns declarations, so LINQ to XML does not emit a redundant xmlns on every a: element.
            /// </summary>
            /// <param name="root">已构建的根元素。 The constructed root.</param>
            /// <param name="withR">是否声明 r 前缀。 Whether to declare the r prefix.</param>
            /// <returns>带命名空间声明的根元素。 The namespace-declared root.</returns>
            private static XElement Declare(XElement root, bool withR)
            {
                List<XObject> content = new List<XObject>();
                content.Add(new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName));
                content.Add(new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName));
                if (withR) content.Add(new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName));
                foreach (XAttribute attribute in root.Attributes()) content.Add(attribute);
                foreach (XElement child in root.Elements()) content.Add(child);
                return new XElement(root.Name, content);
            }

            /// <summary>构造声明了 p/a/r 命名空间的关系型根元素（母版/版式/幻灯片/备注母版）。 Builds a relationship-style root that declares p/a/r namespaces (master/layout/slide/notes-master).</summary>
            /// <param name="name">根元素名。 Root element name.</param>
            /// <param name="children">子元素。 Children.</param>
            /// <returns>根元素。 The root element.</returns>
            private static XElement RelationshipRoot(XName name, params XElement[] children)
            {
                XElement root = new XElement(name);
                for (int i = 0; i < children.Length; i++)
                {
                    if (children[i] != null) root.Add(children[i]);
                }
                return root;
            }

            /// <summary>给根元素加上 XML 声明封装为文档。 Wraps a root into a document with the XML declaration.</summary>
            /// <param name="root">根元素。 Root element.</param>
            /// <returns>带声明的文档。 The declared document.</returns>
            private static XDocument WithDeclaration(XElement root)
            {
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>把时间格式化为 UTC 的 W3CDTF 文本（yyyy-MM-ddTHH:mm:ssZ）。 Formats a timestamp as UTC W3CDTF text (yyyy-MM-ddTHH:mm:ssZ).</summary>
            /// <param name="value">时间值。 Timestamp.</param>
            /// <returns>ISO 8601 文本。 The ISO 8601 text.</returns>
            private static string FormatTimestamp(DateTime value)
            {
                DateTime utc = value.Kind == DateTimeKind.Utc ? value : value.ToUniversalTime();
                return utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
            }

            /// <summary>生成不变区域的整数字符串。 Builds an invariant-culture integer string.</summary>
            /// <param name="value">整数值。 Integer value.</param>
            /// <returns>十进制文本。 The decimal text.</returns>
            private static string IntStr(int value)
            {
                return value.ToString(CultureInfo.InvariantCulture);
            }

            /// <summary>四舍五入到最接近的整数（中点远离零）。 Rounds to the nearest integer with midpoints away from zero.</summary>
            /// <param name="value">数值。 Value.</param>
            /// <returns>取整结果。 The rounded value.</returns>
            private static int RoundToInt(double value)
            {
                return (int)Math.Round(value, MidpointRounding.AwayFromZero);
            }

            /// <summary>把整数限制到闭区间 [min, max]。 Clamps an integer into the inclusive range [min, max].</summary>
            /// <param name="value">待限制的值。 Value to clamp.</param>
            /// <param name="min">下界。 Lower bound.</param>
            /// <param name="max">上界。 Upper bound.</param>
            /// <returns>限制后的值。 The clamped value.</returns>
            private static int ClampInt(int value, int min, int max)
            {
                if (value < min) return min;
                if (value > max) return max;
                return value;
            }

            /// <summary>一条 OPC 关系条目。 A single OPC relationship entry.</summary>
            private struct RelEntry
            {
                /// <summary>关系 Id。 Relationship id.</summary>
                public string Id;
                /// <summary>关系类型完整 URL。 Full relationship-type URL.</summary>
                public string Type;
                /// <summary>关系目标（相对路径）。 Relationship target (relative path).</summary>
                public string Target;

                /// <summary>创建关系条目。 Creates a relationship entry.</summary>
                /// <param name="id">关系 Id。 Relationship id.</param>
                /// <param name="type">关系类型。 Relationship type.</param>
                /// <param name="target">关系目标。 Relationship target.</param>
                public RelEntry(string id, string type, string target)
                {
                    Id = id;
                    Type = type;
                    Target = target;
                }
            }

            /// <summary>媒体去重登记信息。 Image de-duplication registration record.</summary>
            private sealed class MediaInfo
            {
                /// <summary>相对幻灯片目录的目标路径（如 ../media/image1.png）。 Target path relative to the slides directory.</summary>
                public string Target;
                /// <summary>含点的扩展名。 Extension with the leading dot.</summary>
                public string Extension;
                /// <summary>媒体内容类型。 Media content type.</summary>
                public string ContentType;
            }

            /// <summary>单个占位符的几何描述。 Geometry description of one placeholder.</summary>
            private sealed class PlaceholderGeometry
            {
                /// <summary>占位符 type 文本。 Placeholder type text.</summary>
                public string Type;
                /// <summary>占位符类别。 Placeholder kind.</summary>
                public PlaceholderKind Kind;
                /// <summary>占位符序号。 Placeholder index.</summary>
                public int Index;
                /// <summary>占位符名称。 Placeholder name.</summary>
                public string Name;
                /// <summary>左边距（磅）。 Left offset in points.</summary>
                public double Left;
                /// <summary>上边距（磅）。 Top offset in points.</summary>
                public double Top;
                /// <summary>宽度（磅）。 Width in points.</summary>
                public double Width;
                /// <summary>高度（磅）。 Height in points.</summary>
                public double Height;

                /// <summary>创建占位符几何描述。 Creates a placeholder geometry record.</summary>
                /// <param name="type">占位符 type 文本。 Placeholder type text.</param>
                /// <param name="kind">占位符类别。 Placeholder kind.</param>
                /// <param name="index">占位符序号。 Placeholder index.</param>
                /// <param name="name">占位符名称。 Placeholder name.</param>
                /// <param name="left">左边距（磅）。 Left offset in points.</param>
                /// <param name="top">上边距（磅）。 Top offset in points.</param>
                /// <param name="width">宽度（磅）。 Width in points.</param>
                /// <param name="height">高度（磅）。 Height in points.</param>
                public PlaceholderGeometry(string type, PlaceholderKind kind, int index, string name, double left, double top, double width, double height)
                {
                    Type = type;
                    Kind = kind;
                    Index = index;
                    Name = name;
                    Left = left;
                    Top = top;
                    Width = width;
                    Height = height;
                }
            }

            /// <summary>单张幻灯片构建期间的写出条目。 In-progress output entry for one slide.</summary>
            private sealed class SlidePart
            {
                /// <summary>幻灯片 1 基编号。 1-based slide number.</summary>
                public int Number;
                /// <summary>幻灯片模型。 Slide model.</summary>
                public Slide Slide;
                /// <summary>关系条目列表。 Relationship entries.</summary>
                public List<RelEntry> Rels;
                /// <summary>下一个可用关系编号。 Next available relationship number.</summary>
                public int NextRel;
                /// <summary>下一个可用的形状 <c>cNvPr id</c>（从 2 起，1 留给组）。 Next shape cNvPr id (from 2, since 1 is the group).</summary>
                public int NextShapeId;

                /// <summary>创建幻灯片条目。 Creates a slide entry.</summary>
                /// <param name="number">1 基编号。 1-based number.</param>
                /// <param name="slide">幻灯片模型。 Slide model.</param>
                /// <param name="rels">初始关系列表。 Initial relationship list.</param>
                public SlidePart(int number, Slide slide, List<RelEntry> rels)
                {
                    Number = number;
                    Slide = slide;
                    Rels = rels;
                    NextRel = rels.Count + 1;
                    NextShapeId = 2;
                }

                /// <summary>追加一条关系并返回其 r:id。 Appends a relationship and returns its r:id.</summary>
                /// <param name="type">关系类型。 Relationship type.</param>
                /// <param name="target">关系目标。 Relationship target.</param>
                /// <returns>新关系的 r:id。 The new relationship id.</returns>
                public string AddRel(string type, string target)
                {
                    string id = "rId" + IntStrSafe(NextRel);
                    Rels.Add(new RelEntry(id, type, target));
                    NextRel++;
                    return id;
                }

                /// <summary>整数字符串。 Integer to string helper.</summary>
                /// <param name="value">数值。 Value.</param>
                /// <returns>十进制文本。 Decimal text.</returns>
                private static string IntStrSafe(int value)
                {
                    return value.ToString(CultureInfo.InvariantCulture);
                }
            }
        }
    }
}
