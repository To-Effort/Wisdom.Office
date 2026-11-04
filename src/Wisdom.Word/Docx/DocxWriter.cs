using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// DOCX 写出器：将统一文档模型序列化为 WordprocessingML（.docx）压缩包。
    /// DOCX writer: serialises the unified document model into a WordprocessingML (.docx) package.
    /// </summary>
    public static class DocxWriter
    {
        // ---- OOXML 命名空间常量 ----
        private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string OfficeRelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string CorePropertiesNamespace = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        private const string ExtendedPropertiesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        private const string DocPropsVectorNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
        private const string DcNamespace = "http://purl.org/dc/elements/1.1/";
        private const string DcTermsNamespace = "http://purl.org/dc/terms/";
        private const string XsiNamespace = "http://www.w3.org/2001/XMLSchema-instance";

        // ---- 关系类型 URI 常量 ----
        private const string DocumentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string CorePropertiesRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        private const string ExtendedPropertiesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        private const string CustomPropertiesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
        private const string StylesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        private const string SettingsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
        private const string NumberingRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering";
        private const string ImageRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        private const string HeaderRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
        private const string FooterRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
        private const string HyperlinkRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        // ---- 内容类型（[Content_Types].xml）常量 ----
        private const string ContentTypeRelationships = "application/vnd.openxmlformats-package.relationships+xml";
        private const string ContentTypeXml = "application/xml";
        private const string ContentTypeDocument = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
        private const string ContentTypeTemplate = "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml";
        private const string ContentTypeStyles = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
        private const string ContentTypeSettings = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";
        private const string ContentTypeNumbering = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
        private const string ContentTypeHeader = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml";
        private const string ContentTypeFooter = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
        private const string ContentTypeCoreProperties = "application/vnd.openxmlformats-package.core-properties+xml";
        private const string ContentTypeExtendedProperties = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        private const string ContentTypeCustomProperties = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
        private const string CustomPropertiesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
        private const string CustomPropertiesFormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

        /// <summary>
        /// 将文档模型写出为磁盘上的 .docx 文件（覆盖已有文件）。
        /// Writes the document model to a .docx file on disk, overwriting an existing file.
        /// </summary>
        /// <param name="document">待写出的文档模型。The document model to write.</param>
        /// <param name="path">输出文件路径。The output file path.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> 为 null。Thrown when <paramref name="path"/> is null.</exception>
        public static void Write(OfficeDocument document, string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            // 先写同目录临时文件再原子替换：中途失败不会把用户原文件截成 0 字节。
            // Write a temporary file beside the target and swap it in atomically, so a mid-write failure
            // never leaves the user's original file truncated to zero bytes.
            AtomicFileWriter.Write(path, delegate(Stream stream) { Write(document, stream); });
        }

        /// <summary>
        /// 将文档模型序列化为 .docx 字节数组。
        /// Serialises the document model into a .docx byte array.
        /// </summary>
        /// <param name="document">待写出的文档模型。The document model to write.</param>
        /// <returns>完整的 .docx 包字节内容。The bytes of the complete .docx package.</returns>
        public static byte[] Write(OfficeDocument document)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Write(document, stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// 将文档模型写出到可 Seek 的输出流（如内存流或文件流）。
        /// Writes the document model into a seekable output stream such as a memory or file stream.
        /// </summary>
        /// <param name="document">待写出的文档模型。The document model to write.</param>
        /// <param name="stream">接收 .docx 字节的输出流。The output stream receiving the .docx bytes.</param>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="stream"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="stream"/> is null.</exception>
        public static void Write(OfficeDocument document, Stream stream)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (stream == null) throw new ArgumentNullException("stream");
            WriterState state = new WriterState(document, false);
            state.Build();
            state.Save(stream);
        }

        /// <summary>
        /// 将文档模型写出为磁盘上的 .dotx 模板文件（包结构与 .docx 相同，主部件内容类型为模板）。
        /// Writes the document model to a .dotx template file on disk; the package layout matches .docx and only
        /// the main-part content type differs.
        /// </summary>
        /// <param name="document">待写出的文档模型。The document model to write.</param>
        /// <param name="path">输出文件路径。The output file path.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> 为 null。Thrown when <paramref name="path"/> is null.</exception>
        public static void WriteTemplate(OfficeDocument document, string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            AtomicFileWriter.Write(path, delegate(Stream stream) { WriteTemplate(document, stream); });
        }

        /// <summary>
        /// 将文档模型序列化为 .dotx 模板字节数组。
        /// Serialises the document model into a .dotx template byte array.
        /// </summary>
        /// <param name="document">待写出的文档模型。The document model to write.</param>
        /// <returns>完整的 .dotx 包字节内容。The bytes of the complete .dotx package.</returns>
        public static byte[] WriteTemplate(OfficeDocument document)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                WriteTemplate(document, stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// 将文档模型写出为 .dotx 模板到可写流。
        /// Writes the document model as a .dotx template into a writable stream.
        /// </summary>
        /// <param name="document">待写出的文档模型。The document model to write.</param>
        /// <param name="stream">接收 .dotx 字节的输出流。The output stream receiving the .dotx bytes.</param>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="stream"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="stream"/> is null.</exception>
        public static void WriteTemplate(OfficeDocument document, Stream stream)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (stream == null) throw new ArgumentNullException("stream");
            WriterState state = new WriterState(document, true);
            state.Build();
            state.Save(stream);
        }

        /// <summary>关系条目（Id/Type/Target）的内存表示，用于生成 .rels 部件。In-memory relationship entry (Id/Type/Target) used to generate a .rels part.</summary>
        private sealed class RelationshipEntry
        {
            /// <summary>关系 ID（如 rId1）。The relationship id (for example rId1).</summary>
            public string Id;
            /// <summary>关系类型 URI。The relationship type URI.</summary>
            public string Type;
            /// <summary>关系目标路径。The relationship target path.</summary>
            public string Target;
            /// <summary>是否为外部关系（TargetMode="External"）。Whether the relationship is external (TargetMode="External").</summary>
            public bool External;
        }

        /// <summary>
        /// 从段落标记文本还原出的编号形状：编号格式、lvlText 模板与起始值，决定 abstractNum 里第 0 级的写法。
        /// A numbering shape recovered from a paragraph's marker text: the format, the lvlText template and the
        /// start value, which together define level 0 of an abstractNum.
        /// </summary>
        private sealed class NumberingShape
        {
            /// <summary>w:numFmt 值（decimal、lowerLetter、lowerRoman、bullet 等）。The w:numFmt value (decimal, lowerLetter, lowerRoman, bullet, ...).</summary>
            public string Format;
            /// <summary>w:lvlText 模板，数字位置以 %1 表示，如 "%1."。The w:lvlText template with %1 standing for the number, e.g. "%1.".</summary>
            public string LevelText;
            /// <summary>起始序号（w:start），由标记里的数字反推。The start number (w:start), recovered from the marker.</summary>
            public int Start = 1;
            /// <summary>是否为项目符号（项目符号没有起始值，lvlText 就是符号本身）。Whether this is a bullet (bullets have no start value and the lvlText is the symbol itself).</summary>
            public bool Bullet;
        }

        /// <summary>
        /// 一条编号序列：numId 指向可复用的 abstractNumId，一段连续的列表占一个 numId，从而各自从起始值重新计数。
        /// One numbering sequence: a numId pointing at a reusable abstractNumId. Every contiguous list run gets its
        /// own numId so each run counts from its own start value.
        /// </summary>
        private sealed class NumberingDefinition
        {
            /// <summary>编号序列 ID（w:num/@w:numId）。The numbering id (w:num/@w:numId).</summary>
            public int NumId;
            /// <summary>指向的抽象编号定义 ID。The abstract numbering definition it points at.</summary>
            public int AbstractNumId;
            /// <summary>本序列的编号形状。The numbering shape of this sequence.</summary>
            public NumberingShape Shape;
        }

        /// <summary>
        /// 单个 XML 部件（正文/页眉/页脚）的写出上下文：累积该部件的关系条目，
        /// 并为图片提供按部件去重的关系 ID 分配。
        /// Writing context for one XML part (body, header or footer): accumulates the part's relationship
        /// entries and assigns image relationship ids de-duplicated per part.
        /// </summary>
        private sealed class PartContext
        {
            /// <summary>图片到本部件内关系 ID 的缓存，同一图片在同部件只建一条关系。Cache of images to relationship ids in this part; one relationship per image per part.</summary>
            private readonly Dictionary<OfficeImage, string> _imageIds = new Dictionary<OfficeImage, string>();
            /// <summary>外部超链接目标到本部件内关系 ID 的缓存，同一目标只建一条关系。Cache of external hyperlink targets to relationship ids in this part; one relationship per target.</summary>
            private readonly Dictionary<string, string> _hyperlinkIds = new Dictionary<string, string>(StringComparer.Ordinal);
            /// <summary>下一个可用的关系编号。The next relationship number to allocate.</summary>
            private int _nextId;

            /// <summary>
            /// 创建绑定到指定部件路径的写出上下文。
            /// Creates a writing context bound to the given part path.
            /// </summary>
            /// <param name="partPath">部件路径（如 word/document.xml）。The part path (for example word/document.xml).</param>
            public PartContext(string partPath)
            {
                PartPath = partPath;
            }

            /// <summary>本上下文对应的部件路径。The part path this context belongs to.</summary>
            public string PartPath { get; private set; }
            /// <summary>本部件累积的关系条目列表。The relationship entries accumulated for this part.</summary>
            public List<RelationshipEntry> Relationships { get; private set; } = new List<RelationshipEntry>();
            /// <summary>组装阶段算好的部件字节。Part bytes produced during the assembly phase.</summary>
            public byte[] PartBytes;
            /// <summary>组装阶段算好的关系部件字节，null 表示本部件没有关系。Relationship part bytes, null when the part has no relationships.</summary>
            public byte[] RelationshipsBytes;
            /// <summary>所属的写出状态，用于查询全局媒体信息。The owning writer state, used to look up global media information.</summary>
            public WriterState Owner;

            /// <summary>
            /// 追加一条关系并自动分配递增的 rId 编号。
            /// Appends a relationship and allocates the next incremental rId automatically.
            /// </summary>
            /// <param name="type">关系类型 URI。The relationship type URI.</param>
            /// <param name="target">关系目标路径。The relationship target path.</param>
            /// <returns>分配的关系 ID。The allocated relationship id.</returns>
            public string AddRelationship(string type, string target)
            {
                RelationshipEntry entry = new RelationshipEntry();
                entry.Id = "rId" + (++_nextId).ToString(CultureInfo.InvariantCulture);
                entry.Type = type;
                entry.Target = target;
                Relationships.Add(entry);
                return entry.Id;
            }

            /// <summary>
            /// 获取图片在本部件中的关系 ID：已存在则复用，否则依据全局媒体信息新建关系；
            /// 图片未登记到媒体表时返回 null。
            /// Gets the relationship id of an image within this part, reusing an existing one or creating a new
            /// relationship from the global media table; returns null when the image is not registered.
            /// </summary>
            /// <param name="image">图片对象。The image object.</param>
            /// <returns>关系 ID 或 null。The relationship id, or null.</returns>
            public string GetImageRelationship(OfficeImage image)
            {
                string id;
                if (_imageIds.TryGetValue(image, out id)) return id;
                MediaInfo media;
                if (!Owner.Media.TryGetValue(image, out media)) return null;
                id = AddRelationship(ImageRelationshipType, media.Target);
                _imageIds[image] = id;
                return id;
            }

            /// <summary>
            /// 获取外部超链接在本部件中的关系 ID：同一目标复用已有关系，否则新建 External 关系。
            /// Gets the relationship id of an external hyperlink in this part, reusing the existing relationship
            /// for the same target or creating a new External relationship.
            /// </summary>
            /// <param name="target">外部链接目标。The external hyperlink target.</param>
            /// <returns>关系 ID。The relationship id.</returns>
            public string GetHyperlinkRelationship(string target)
            {
                string id;
                if (_hyperlinkIds.TryGetValue(target, out id)) return id;
                RelationshipEntry entry = new RelationshipEntry();
                entry.Id = "rId" + (++_nextId).ToString(CultureInfo.InvariantCulture);
                entry.Type = HyperlinkRelationshipType;
                entry.Target = target;
                entry.External = true;
                Relationships.Add(entry);
                _hyperlinkIds[target] = entry.Id;
                return entry.Id;
            }
        }

        /// <summary>图片媒体登记信息：文件名、word/ 下相对目标路径、内容类型与扩展名。Registration record for image media: file name, target path relative to word/, content type and extension.</summary>
        private sealed class MediaInfo
        {
            /// <summary>媒体文件名（如 image1.png）。The media file name (for example image1.png).</summary>
            public string FileName;
            /// <summary>相对 word/ 的目标路径（如 media/image1.png）。The target path relative to word/ (for example media/image1.png).</summary>
            public string Target;
            /// <summary>媒体内容类型（如 image/png）。The media content type (for example image/png).</summary>
            public string ContentType;
            /// <summary>小写扩展名（含点，如 .png）。The lower-case extension including the dot (for example .png).</summary>
            public string Extension;
            /// <summary>关联的图片对象。The image object this record belongs to.</summary>
            public OfficeImage Image;
        }

        /// <summary>一节对应的页眉/页脚部件及其关系 ID 与该节内容宽度。The header and footer parts of a section, their relationship ids and the section content width.</summary>
        private sealed class SectionParts
        {
            /// <summary>页眉部件上下文，未启用页眉时为 null。The header part context, null when the section has no header.</summary>
            public PartContext Header;
            /// <summary>页脚部件上下文，未启用页脚时为 null。The footer part context, null when the section has no footer.</summary>
            public PartContext Footer;
            /// <summary>正文部件中指向页眉的关系 ID，可为 null。The relationship id pointing at the header from the body part, may be null.</summary>
            public string HeaderRelationshipId;
            /// <summary>正文部件中指向页脚的关系 ID，可为 null。The relationship id pointing at the footer from the body part, may be null.</summary>
            public string FooterRelationshipId;
            /// <summary>本节内容宽度（磅，页面宽减去左右边距），用于表格列宽分配。The content width of the section in points (page width minus margins), used to distribute table column widths.</summary>
            public double ContentWidth;
        }

        /// <summary>
        /// 写出过程的可变状态：持有文档模型、媒体登记表、各节部件、正文元素以及
        /// 图形/书签/id 计数器；先 <see cref="Build"/> 组装 XML，再 <see cref="Save"/> 打包 ZIP。
        /// Mutable state of a write operation: holds the document model, the media table, the per-section parts,
        /// the body element and the drawing/bookmark id counters; <see cref="Build"/> assembles the XML and
        /// <see cref="Save"/> packs the ZIP.
        /// </summary>
        private sealed class WriterState
        {
            /// <summary>待写出的文档模型。The document model being written.</summary>
            private readonly OfficeDocument _document;
            /// <summary>图片去重登记表：同一图片在整包中只写一份媒体文件。Image de-duplication table: each image produces one media file per package.</summary>
            private readonly Dictionary<OfficeImage, MediaInfo> _media = new Dictionary<OfficeImage, MediaInfo>();
            /// <summary>与文档各节一一对应的页眉页脚部件信息。Header/footer part information aligned one-to-one with the document sections.</summary>
            private readonly List<SectionParts> _sections = new List<SectionParts>();
            /// <summary>主文档部件的写出上下文。The writing context of the main document part.</summary>
            private readonly PartContext _documentPart = new PartContext("word/document.xml");
            /// <summary>组装中的 w:body 元素。The w:body element under construction.</summary>
            private XElement _body;
            /// <summary>DrawingML docPr 图形 ID 计数器。Counter for DrawingML docPr graphic ids.</summary>
            private int _drawingId;
            /// <summary>书签 ID 计数器。Counter for bookmark ids.</summary>
            private int _bookmarkId;
            /// <summary>
            /// 书签名到本次写出分配的 id 的映射，供 <c>w:bookmarkEnd</c> 复用起点的 id；
            /// 全局计数保证同一部件内 id 唯一（OOXML 只要求部件内唯一）。
            /// Maps bookmark names to the ids assigned during this write so <c>w:bookmarkEnd</c> reuses its start's
            /// id; the global counter keeps ids unique within every part (OOXML only requires uniqueness per part).
            /// </summary>
            private readonly Dictionary<string, int> _bookmarkIdByName = new Dictionary<string, int>(StringComparer.Ordinal);
            /// <summary>修订 ID 计数器，写出时重新分配唯一 ID。Counter for revision ids; fresh unique ids are assigned on write.</summary>
            private int _revisionId;
            /// <summary>styles.xml 的关系 ID（构建时分配）。The relationship id of styles.xml, allocated while building.</summary>
            private string _stylesRelationshipId;
            /// <summary>settings.xml 的关系 ID（构建时分配）。The relationship id of settings.xml, allocated while building.</summary>
            private string _settingsRelationshipId;
            /// <summary>编号形状（格式 + 模板 + 起始值）到 abstractNumId 的映射；形状相同的列表复用同一 abstractNum。Map from a numbering shape (format, template and start value) to its abstractNumId; lists of the same shape share one abstractNum.</summary>
            private readonly Dictionary<string, int> _abstractNumbersByShape = new Dictionary<string, int>(StringComparer.Ordinal);
            /// <summary>本次写出分配出的编号序列，按 numId 顺序写入 numbering.xml。The numbering sequences allocated by this write, emitted into numbering.xml in numId order.</summary>
            private readonly List<NumberingDefinition> _numberingDefinitions = new List<NumberingDefinition>();
            /// <summary>下一个可用的 numId（从 1 开始）。The next numId to allocate, starting at 1.</summary>
            private int _nextNumId = 1;
            /// <summary>下一个可用的 abstractNumId（从 0 开始）。The next abstractNumId to allocate, starting at 0.</summary>
            private int _nextAbstractNumberId;
            /// <summary>当前是否处于一段连续的列表段落中。Whether the writer is inside a contiguous run of list paragraphs.</summary>
            private bool _listRunOpen;
            /// <summary>当前列表段的编号形状键。The shape key of the active list run.</summary>
            private string _activeListShape;
            /// <summary>当前列表段所在的部件；跨部件（正文到页眉）必须重新起号。The part owning the active list run; crossing parts (body to header) must restart the numbering.</summary>
            private PartContext _activeListPart;

            /// <summary>主文档部件的内容类型（.docx 或 .dotx）。The content type of the main document part (.docx or .dotx).</summary>
            private readonly string _mainContentType;

            /// <summary>
            /// 以文档模型创建写出状态并绑定主文档部件上下文。
            /// Creates the writer state from a document model and binds the main document part context to it.
            /// </summary>
            /// <param name="document">待写出的文档模型。The document model to write.</param>
            /// <param name="template">true 写出 .dotx 模板内容类型，false 写出普通 .docx。True writes the .dotx template content type; false writes a regular .docx.</param>
            public WriterState(OfficeDocument document, bool template)
            {
                _document = document;
                _mainContentType = template ? ContentTypeTemplate : ContentTypeDocument;
                _documentPart.Owner = this;
            }

            /// <summary>全局图片媒体登记表，供各部件上下文查询关系目标。The global image media table that part contexts use to resolve relationship targets.</summary>
            public Dictionary<OfficeImage, MediaInfo> Media { get { return _media; } }

            /// <summary>
            /// 组装正文：登记图片与样式/设置关系，为启用页眉页脚的节创建独立部件，
            /// 写出各节块内容；中间节以末段 pPr 内的 sectPr 结束，最后一节的 sectPr 直接挂在 body 末尾
            /// （若正文以表格结尾则先补一个空段落，符合 Word 规范）。
            /// Assembles the body: registers images plus the styles and settings relationships, creates separate
            /// parts for sections with headers or footers and emits every section's blocks. Intermediate sections
            /// end with a sectPr inside the last paragraph's pPr, while the final sectPr hangs directly on the
            /// body (an empty paragraph is appended first when the body ends with a table, as Word requires).
            /// </summary>
            public void Build()
            {
                CollectImages(_document);
                _stylesRelationshipId = _documentPart.AddRelationship(StylesRelationshipType, "styles.xml");
                _settingsRelationshipId = _documentPart.AddRelationship(SettingsRelationshipType, "settings.xml");

                XNamespace w = WordNamespace;
                _body = new XElement(w + "body");

                int headerCount = 0;
                int footerCount = 0;
                for (int i = 0; i < _document.Sections.Count; i++)
                {
                    Section section = _document.Sections[i];
                    SectionParts parts = new SectionParts();
                    PageSetup setup = section.PageSetup ?? new PageSetup();
                    parts.ContentWidth = Math.Max(36, setup.Width - setup.MarginLeft - setup.MarginRight);
                    if (section.Header != null && section.Header.IsEnabled)
                    {
                        headerCount++;
                        string path = "word/header" + headerCount.ToString(CultureInfo.InvariantCulture) + ".xml";
                        parts.Header = new PartContext(path);
                        parts.Header.Owner = this;
                        parts.HeaderRelationshipId = _documentPart.AddRelationship(HeaderRelationshipType, "header" + headerCount.ToString(CultureInfo.InvariantCulture) + ".xml");
                    }
                    if (section.Footer != null && section.Footer.IsEnabled)
                    {
                        footerCount++;
                        string path = "word/footer" + footerCount.ToString(CultureInfo.InvariantCulture) + ".xml";
                        parts.Footer = new PartContext(path);
                        parts.Footer.Owner = this;
                        parts.FooterRelationshipId = _documentPart.AddRelationship(FooterRelationshipType, "footer" + footerCount.ToString(CultureInfo.InvariantCulture) + ".xml");
                    }
                    _sections.Add(parts);
                }

                for (int i = 0; i < _document.Sections.Count; i++)
                {
                    Section section = _document.Sections[i];
                    SectionParts parts = _sections[i];
                    // sectPr 只能挂在本节自己的末段上：早先直接取 body 的最后一个元素，空节（Blocks 为空）
                    // 会把分节符塞进上一节末段的 pPr，轻则空节的页面设置丢失，重则同一段落下出现两个 sectPr
                    // （pPr 只允许一个 sectPr，Word 会判文件损坏）。
                    // A section's sectPr may only attach to that section's own last paragraph: the old code took
                    // the body's last element, so an empty section pushed its break into the previous section's
                    // final paragraph — losing the empty section's page setup and, when two empty sections lined
                    // up, writing two sectPr under one pPr (only one is allowed, and Word flags the file as corrupt).
                    int bodyBefore = _body.Elements().Count();
                    WriteBlocks(_body, section.Blocks, _documentPart, parts.ContentWidth);
                    bool wroteBlocks = _body.Elements().Count() > bodyBefore;

                    XElement sectionProperties = CreateSectionProperties(section, parts);
                    if (i < _document.Sections.Count - 1)
                    {
                        XElement last = wroteBlocks ? _body.Elements().LastOrDefault() : null;
                        if (last != null && last.Name == w + "p")
                        {
                            XElement properties = last.Element(w + "pPr");
                            if (properties == null)
                            {
                                properties = new XElement(w + "pPr");
                                last.AddFirst(properties);
                            }
                            properties.Add(sectionProperties);
                        }
                        else
                        {
                            _body.Add(new XElement(w + "p", new XElement(w + "pPr", sectionProperties)));
                        }
                    }
                    else
                    {
                        if (_body.Elements().LastOrDefault() != null
                            && _body.Elements().LastOrDefault().Name == w + "tbl")
                        {
                            _body.Add(new XElement(w + "p"));
                        }
                        _body.Add(sectionProperties);
                    }
                }

                if (_document.Sections.Count == 0)
                {
                    _body.Add(CreateSectionProperties(new Section(), new SectionParts()));
                }

                // 页眉页脚的正文必须在 numbering 定稿之前组装：页眉里的列表同样会登记编号定义，
                // 若等到写包时才生成，numbering.xml、它的关系与内容类型都已按"没有编号"定稿，
                // 页眉里的 w:numPr 会指向不存在的序列（Word 打开即提示修复）。
                // Header/footer content has to be assembled before numbering is finalised: a list inside a header
                // registers numbering definitions too, and generating it during the zip write would leave
                // numbering.xml, its relationship and its content-type already settled as "no numbering", so the
                // w:numPr in the header would reference a sequence that does not exist (Word reports a repair).
                for (int i = 0; i < _sections.Count; i++)
                {
                    SectionParts parts = _sections[i];
                    if (parts.Header != null) BuildHeaderFooterPart(parts.Header, true);
                    if (parts.Footer != null) BuildHeaderFooterPart(parts.Footer, false);
                }

                // numbering.xml 只有写出过编号段落时才存在，关系也必须同步缺席，否则 Word 会报包结构损坏。
                // numbering.xml only exists when numbered paragraphs were written, and the relationship has to be
                // absent with it; an orphan relationship makes Word report a broken package.
                if (_numberingDefinitions.Count > 0)
                {
                    _documentPart.AddRelationship(NumberingRelationshipType, "numbering.xml");
                }
            }

            /// <summary>
            /// 将已组装的全部部件写入 ZIP 包（创建模式，流保持打开）。
            /// Writes all assembled parts into the ZIP package in create mode, leaving the stream open.
            /// </summary>
            /// <param name="stream">输出流。The output stream.</param>
            public void Save(Stream stream)
            {
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
                {
                    WritePackageParts(archive);
                }
            }

            /// <summary>
            /// 遍历正文、页眉、页脚收集全部图片，按出现顺序去重登记为 media/imageN.ext，
            /// 同时按扩展名汇总媒体内容类型。
            /// Walks the body, headers and footers collecting every image, registering each one de-duplicated in
            /// order of appearance as media/imageN.ext and aggregating media content types by extension.
            /// </summary>
            /// <param name="document">文档模型。The document model.</param>
            private void CollectImages(OfficeDocument document)
            {
                Dictionary<string, string> contentTypesByExtension = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                List<OfficeImage> ordered = new List<OfficeImage>();
                for (int s = 0; s < document.Sections.Count; s++)
                {
                    CollectImages(document.Sections[s].Blocks, ordered);
                    // Section.Header / Footer 有公开 setter，宿主可以置空；这里先判空再取块。
                    // Section.Header and Section.Footer expose public setters, so either can be null here.
                    if (document.Sections[s].Header != null) CollectImages(document.Sections[s].Header.Blocks, ordered);
                    if (document.Sections[s].Footer != null) CollectImages(document.Sections[s].Footer.Blocks, ordered);
                }

                int index = 0;
                for (int i = 0; i < ordered.Count; i++)
                {
                    OfficeImage image = ordered[i];
                    if (_media.ContainsKey(image)) continue;
                    index++;
                    string extension = string.IsNullOrEmpty(image.FileExtension) ? ".bin" : image.FileExtension.ToLowerInvariant();
                    if (!extension.StartsWith(".", StringComparison.Ordinal)) extension = "." + extension;
                    MediaInfo media = new MediaInfo();
                    media.Image = image;
                    media.Extension = extension;
                    media.FileName = "image" + index.ToString(CultureInfo.InvariantCulture) + extension;
                    media.Target = "media/" + media.FileName;
                    media.ContentType = string.IsNullOrEmpty(image.ContentType) ? "application/octet-stream" : image.ContentType;
                    _media[image] = media;
                    if (!contentTypesByExtension.ContainsKey(extension))
                    {
                        contentTypesByExtension[extension] = media.ContentType;
                    }
                }
                MediaContentTypes = contentTypesByExtension;
            }

            /// <summary>媒体扩展名到内容类型的映射（[Content_Types].xml 的 Default 条目来源）。Map from media extension to content type, the source of the Default entries of [Content_Types].xml.</summary>
            public Dictionary<string, string> MediaContentTypes { get; private set; } = new Dictionary<string, string>();

            /// <summary>
            /// 递归收集块列表（含表格单元格内块）中全部 <see cref="ImageRun"/> 引用的图片。
            /// Recursively collects the images referenced by every <see cref="ImageRun"/> in a block list,
            /// including blocks inside table cells.
            /// </summary>
            /// <param name="blocks">块列表，可为 null。The block list, may be null.</param>
            /// <param name="images">按出现顺序收集的图片列表。The images collected in order of appearance.</param>
            private static void CollectImages(List<Block> blocks, List<OfficeImage> images)
            {
                if (blocks == null) return;
                for (int i = 0; i < blocks.Count; i++)
                {
                    Paragraph paragraph = blocks[i] as Paragraph;
                    if (paragraph != null)
                    {
                        for (int r = 0; r < paragraph.Inlines.Count; r++)
                        {
                            ImageRun image = paragraph.Inlines[r] as ImageRun;
                            if (image != null && image.Image != null) images.Add(image.Image);
                        }
                        continue;
                    }
                    Table table = blocks[i] as Table;
                    if (table == null) continue;
                    for (int row = 0; row < table.Rows.Count; row++)
                    {
                        for (int cell = 0; cell < table.Rows[row].Cells.Count; cell++)
                        {
                            CollectImages(table.Rows[row].Cells[cell].Blocks, images);
                        }
                    }
                }
            }

            /// <summary>
            /// 生成 w:sectPr 节属性：页眉页脚关系引用、页面尺寸（磅转缇，宽大于高时标记 landscape）、
            /// 页边距与单栏布局。
            /// Builds the w:sectPr section properties: header/footer relationship references, page size (points
            /// converted to twips, landscape flagged when width exceeds height), margins and a single-column layout.
            /// </summary>
            /// <param name="section">文档节。The document section.</param>
            /// <param name="parts">该节对应的页眉页脚部件信息。The header/footer part information of the section.</param>
            /// <returns>sectPr 元素。The sectPr element.</returns>
            private XElement CreateSectionProperties(Section section, SectionParts parts)
            {
                XNamespace w = WordNamespace;
                XNamespace r = OfficeRelationshipNamespace;
                XElement properties = new XElement(w + "sectPr");
                if (parts.HeaderRelationshipId != null)
                {
                    properties.Add(new XElement(w + "headerReference",
                        new XAttribute(w + "type", "default"),
                        new XAttribute(r + "id", parts.HeaderRelationshipId)));
                }
                if (parts.FooterRelationshipId != null)
                {
                    properties.Add(new XElement(w + "footerReference",
                        new XAttribute(w + "type", "default"),
                        new XAttribute(r + "id", parts.FooterRelationshipId)));
                }

                PageSetup setup = section.PageSetup ?? new PageSetup();
                XElement pageSize = new XElement(w + "pgSz",
                    new XAttribute(w + "w", ToTwips(setup.Width)),
                    new XAttribute(w + "h", ToTwips(setup.Height)));
                if (setup.Width > setup.Height)
                {
                    pageSize.Add(new XAttribute(w + "orient", "landscape"));
                }
                properties.Add(pageSize);
                properties.Add(new XElement(w + "pgMar",
                    new XAttribute(w + "top", ToTwips(setup.MarginTop)),
                    new XAttribute(w + "right", ToTwips(setup.MarginRight)),
                    new XAttribute(w + "bottom", ToTwips(setup.MarginBottom)),
                    new XAttribute(w + "left", ToTwips(setup.MarginLeft)),
                    new XAttribute(w + "header", ToTwips(setup.HeaderDistance)),
                    new XAttribute(w + "footer", ToTwips(setup.FooterDistance)),
                    new XAttribute(w + "gutter", "0")));
                properties.Add(new XElement(w + "cols", new XAttribute(w + "space", "708")));
                return properties;
            }

            /// <summary>
            /// 写出包内全部部件：主文档、样式、设置及其关系、各节页眉页脚部件、
            /// 媒体图片、[Content_Types].xml、根关系与 core/app 属性部件。
            /// Writes every part of the package: main document, styles, settings and their relationships, the
            /// header/footer parts of each section, media images, [Content_Types].xml, the root relationships
            /// and the core/app property parts.
            /// </summary>
            /// <param name="archive">目标 ZIP 包。The target ZIP package.</param>
            private void WritePackageParts(ZipArchive archive)
            {
                XNamespace w = WordNamespace;
                XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
                XNamespace r = OfficeRelationshipNamespace;

                XDocument documentXml = new XDocument(
                    new XDeclaration("1.0", "UTF-8", "yes"),
                    new XElement(w + "document",
                        new XAttribute(XNamespace.Xmlns + "w", w),
                        new XAttribute(XNamespace.Xmlns + "r", r),
                        new XAttribute(XNamespace.Xmlns + "wp", wp),
                        new XAttribute(XNamespace.Xmlns + "a", a),
                        new XAttribute(XNamespace.Xmlns + "pic", pic),
                        _body));

                WriteEntry(archive, "word/document.xml", Wml.SerializeXml(documentXml));
                WriteEntry(archive, "word/styles.xml", Wml.SerializeXml(CreateStylesXml()));
                WriteEntry(archive, "word/settings.xml", Wml.SerializeXml(CreateSettingsXml()));
                if (_numberingDefinitions.Count > 0)
                {
                    WriteEntry(archive, "word/numbering.xml", Wml.SerializeXml(CreateNumberingXml()));
                }
                WriteEntry(archive, "word/_rels/document.xml.rels", Wml.SerializeXml(CreateRelationshipsXml(_documentPart.Relationships)));

                for (int i = 0; i < _document.Sections.Count; i++)
                {
                    SectionParts parts = _sections[i];
                    if (parts.Header != null)
                    {
                        WriteStoredHeaderFooter(archive, parts.Header);
                    }
                    if (parts.Footer != null)
                    {
                        WriteStoredHeaderFooter(archive, parts.Footer);
                    }
                }

                foreach (KeyValuePair<OfficeImage, MediaInfo> pair in _media)
                {
                    MediaInfo media = pair.Value;
                    WriteEntry(archive, "word/" + media.Target, media.Image.Data);
                }

                WriteEntry(archive, "[Content_Types].xml", Wml.SerializeXml(CreateContentTypesXml()));
                WriteEntry(archive, "_rels/.rels", Wml.SerializeXml(CreateRootRelationshipsXml()));
                WriteEntry(archive, "docProps/core.xml", Wml.SerializeXml(CreateCorePropertiesXml()));
                WriteEntry(archive, "docProps/app.xml", Wml.SerializeXml(CreateAppPropertiesXml()));
                if (HasCustomProperties())
                {
                    WriteEntry(archive, "docProps/custom.xml", Wml.SerializeXml(CreateCustomPropertiesXml()));
                }
            }

            /// <summary>
            /// 组装单个页眉或页脚部件（w:hdr / w:ftr 根元素）：定位其所属节以取正确块内容与宽度，
            /// 生成部件 XML 与存在图片关系时的关系 XML，字节暂存留给写包阶段落盘。
            /// Assembles one header or footer part (w:hdr / w:ftr root): locates the owning section to take the
            /// right blocks and width, builds the part XML plus the relationship XML when image relationships
            /// exist, and stashes the bytes for the zip-writing phase.
            /// </summary>
            /// <param name="context">页眉/页脚部件上下文。The header/footer part context.</param>
            /// <param name="header">true 为页眉，false 为页脚。True for a header, false for a footer.</param>
            private void BuildHeaderFooterPart(PartContext context, bool header)
            {
                XNamespace w = WordNamespace;
                XNamespace r = OfficeRelationshipNamespace;
                XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
                string rootName = header ? "hdr" : "ftr";

                XElement root = new XElement(w + rootName,
                    new XAttribute(XNamespace.Xmlns + "w", w),
                    new XAttribute(XNamespace.Xmlns + "r", r),
                    new XAttribute(XNamespace.Xmlns + "wp", wp),
                    new XAttribute(XNamespace.Xmlns + "a", a),
                    new XAttribute(XNamespace.Xmlns + "pic", pic));
                double width = 468;
                List<Block> blocks = null;
                for (int i = 0; i < _document.Sections.Count; i++)
                {
                    SectionParts parts = _sections[i];
                    if ((header && parts.Header == context) || (!header && parts.Footer == context))
                    {
                        Section section = _document.Sections[i];
                        blocks = header ? section.Header.Blocks : section.Footer.Blocks;
                        width = parts.ContentWidth;
                        break;
                    }
                }
                if (blocks == null && _document.Sections.Count > 0)
                {
                    Section first = _document.Sections[0];
                    PageSetup setup = first.PageSetup;
                    width = Math.Max(36, setup.Width - setup.MarginLeft - setup.MarginRight);
                    blocks = header
                        ? (first.Header == null ? null : first.Header.Blocks)
                        : (first.Footer == null ? null : first.Footer.Blocks);
                }
                WriteBlocks(root, blocks, context, width);

                XDocument xml = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
                context.PartBytes = Wml.SerializeXml(xml);
                if (context.Relationships.Count > 0)
                {
                    string file = context.PartPath.Substring(context.PartPath.LastIndexOf('/') + 1);
                    context.RelationshipsBytes = Wml.SerializeXml(CreateRelationshipsXml(context.Relationships));
                }
            }

            /// <summary>
            /// 把组装阶段暂存的页眉/页脚字节落盘，并写出其关系部件（若有）。
            /// Flushes the header/footer bytes stashed during assembly, plus the relationship part when present.
            /// </summary>
            /// <param name="archive">目标 ZIP 包。The target ZIP package.</param>
            /// <param name="context">页眉/页脚部件上下文。The header/footer part context.</param>
            private void WriteStoredHeaderFooter(ZipArchive archive, PartContext context)
            {
                WriteEntry(archive, context.PartPath, context.PartBytes);
                if (context.RelationshipsBytes != null)
                {
                    string file = context.PartPath.Substring(context.PartPath.LastIndexOf('/') + 1);
                    WriteEntry(archive, "word/_rels/" + file + ".rels", context.RelationshipsBytes);
                }
            }

            /// <summary>
            /// 以最佳压缩级别向包中写入一个条目。
            /// Writes a single entry into the package at the optimal compression level.
            /// </summary>
            /// <param name="archive">目标 ZIP 包。The target ZIP package.</param>
            /// <param name="path">条目路径。The entry path.</param>
            /// <param name="data">条目字节内容。The entry bytes.</param>
            private static void WriteEntry(ZipArchive archive, string path, byte[] data)
            {
                ZipArchiveEntry entry = archive.CreateEntry(OfficePackageLimits.ValidateEntryName(path), CompressionLevel.Optimal);
                using (Stream stream = entry.Open())
                {
                    stream.Write(data, 0, data.Length);
                }
            }

            /// <summary>
            /// 按块类型分派写出：段落生成 w:p、表格生成 w:tbl、分页符生成含 w:br type=page 的空段落。
            /// Dispatches writing by block type: paragraphs become w:p, tables become w:tbl and page breaks
            /// become a paragraph holding w:br type=page.
            /// </summary>
            /// <param name="parent">接收元素的父 XML 节点。The parent XML node receiving the elements.</param>
            /// <param name="blocks">块列表，可为 null。The block list, may be null.</param>
            /// <param name="context">当前部件写出上下文。The current part writing context.</param>
            /// <param name="contentWidth">可用于表格列宽分配的内容宽度（磅）。The content width in points available for table column sizing.</param>
            /// <summary>
            /// 顺序写出块级元素。相邻两张表之间必须插一个空段落：OOXML 里直接相连的 w:tbl 会被 Word 合并成
            /// 一张表（第二张表的行会被塞进第一张表的列数里），读回来时行列都错了。
            /// Writes the block elements in order. Two adjacent tables need an empty paragraph between them:
            /// back-to-back w:tbl elements are merged into a single table by Word, so the second table's rows end
            /// up inside the first table's column count and both tables read back wrong.
            /// </summary>
            private void WriteBlocks(XElement parent, List<Block> blocks, PartContext context, double contentWidth)
            {
                if (blocks == null) return;
                XNamespace w = WordNamespace;
                bool lastWasTable = false;
                for (int i = 0; i < blocks.Count; i++)
                {
                    Paragraph paragraph = blocks[i] as Paragraph;
                    if (paragraph != null)
                    {
                        parent.Add(CreateParagraph(paragraph, context));
                        lastWasTable = false;
                        continue;
                    }
                    Table table = blocks[i] as Table;
                    if (table != null)
                    {
                        if (lastWasTable) parent.Add(new XElement(w + "p"));
                        // 表格会打断列表：表内段落即使带标记也自成一段序列，表后段落重新起号。
                        // A table interrupts the list: even marked paragraphs inside it form their own runs, and
                        // the paragraphs after the table restart the numbering.
                        _listRunOpen = false;
                        parent.Add(CreateTable(table, context, contentWidth));
                        _listRunOpen = false;
                        lastWasTable = true;
                        continue;
                    }
                    if (blocks[i] is PageBreak)
                    {
                        parent.Add(CreatePageBreakParagraph());
                        lastWasTable = false;
                    }
                    // 表格与分页符会打断列表：下一段列表必须重新起号，不能接着上一段计数。
                    // Tables and page breaks interrupt a list, so the next list paragraph must restart its count.
                    _listRunOpen = false;
                }
            }

            /// <summary>
            /// 生成段落元素 w:p：写出段落属性（标题样式、间距、缩进、对齐、编号引用）、书签包裹、
            /// 以及各行内内容（文本/域/图片）；无法映射到 numbering.xml 的列表标记才写成近似文本。
            /// Builds the w:p paragraph element: paragraph properties (heading style, spacing, indents, alignment,
            /// numbering reference), a bookmark wrapper and every inline (text, field or image); only list markers
            /// that cannot be mapped onto numbering.xml are written as approximate text.
            /// </summary>
            /// <param name="paragraph">段落模型。The paragraph model.</param>
            /// <param name="context">当前部件写出上下文。The current part writing context.</param>
            /// <returns>w:p 元素。The w:p element.</returns>
            private XElement CreateParagraph(Paragraph paragraph, PartContext context)
            {
                XNamespace w = WordNamespace;
                XElement element = new XElement(w + "p");
                XElement properties = new XElement(w + "pPr");
                NumberingDefinition numbering = AssignNumbering(paragraph, context);

                if (paragraph.HeadingLevel > 0)
                {
                    properties.Add(new XElement(w + "pStyle", new XAttribute(w + "val", "Heading" + Math.Min(9, paragraph.HeadingLevel).ToString(CultureInfo.InvariantCulture))));
                }
                if (paragraph.KeepWithNext) properties.Add(new XElement(w + "keepNext"));
                if (paragraph.PageBreakBefore) properties.Add(new XElement(w + "pageBreakBefore"));
                // numPr 在 pPr 里必须排在 spacing/ind/jc 之前，否则 Word 判为无效顺序。
                // numPr has to precede spacing/ind/jc inside pPr, otherwise Word rejects the child order.
                if (numbering != null)
                {
                    properties.Add(new XElement(w + "numPr",
                        new XElement(w + "ilvl", new XAttribute(w + "val", "0")),
                        new XElement(w + "numId", new XAttribute(w + "val", numbering.NumId.ToString(CultureInfo.InvariantCulture)))));
                }
                if (paragraph.SpaceBefore > 0 || paragraph.SpaceAfter > 0 || Math.Abs(paragraph.LineSpacing - 1.0) > 0.001)
                {
                    XElement spacing = new XElement(w + "spacing");
                    if (paragraph.SpaceBefore > 0)
                        spacing.Add(new XAttribute(w + "before", ToTwips(paragraph.SpaceBefore)));
                    if (paragraph.SpaceAfter > 0)
                        spacing.Add(new XAttribute(w + "after", ToTwips(paragraph.SpaceAfter)));
                    if (Math.Abs(paragraph.LineSpacing - 1.0) > 0.001)
                    {
                        spacing.Add(new XAttribute(w + "line", ClampToInt(240 * Math.Max(0.1, paragraph.LineSpacing), 24, 2147483000).ToString(CultureInfo.InvariantCulture)));
                        spacing.Add(new XAttribute(w + "lineRule", "auto"));
                    }
                    properties.Add(spacing);
                }
                double leftIndent = paragraph.LeftIndent;
                double firstLineIndent = paragraph.FirstLineIndent;
                if (numbering != null && leftIndent == 0 && firstLineIndent == 0)
                {
                    // 编号段落没有显式缩进时补一套标准悬挂缩进：Word 从 numbering.xml 拿到同样的缩进，
                    // 而我们自己的读取端与 PDF/HTML 渲染只看段落属性，补上后三处显示一致。
                    // A numbered paragraph without explicit indents gets the standard hanging indent: Word takes the
                    // same indent from numbering.xml, while our reader and the PDF/HTML renderers only look at
                    // paragraph properties, so writing it keeps all three views consistent.
                    leftIndent = 36;
                    firstLineIndent = -18;
                }
                if (leftIndent != 0 || paragraph.RightIndent != 0 || firstLineIndent != 0)
                {
                    XElement indent = new XElement(w + "ind");
                    if (leftIndent != 0) indent.Add(new XAttribute(w + "left", ToTwips(leftIndent)));
                    if (paragraph.RightIndent != 0) indent.Add(new XAttribute(w + "right", ToTwips(paragraph.RightIndent)));
                    if (firstLineIndent > 0) indent.Add(new XAttribute(w + "firstLine", ToTwips(firstLineIndent)));
                    else if (firstLineIndent < 0) indent.Add(new XAttribute(w + "hanging", ToTwips(-firstLineIndent)));
                    properties.Add(indent);
                }
                switch (paragraph.Alignment)
                {
                    case ParagraphAlignment.Center:
                        properties.Add(new XElement(w + "jc", new XAttribute(w + "val", "center")));
                        break;
                    case ParagraphAlignment.Right:
                        properties.Add(new XElement(w + "jc", new XAttribute(w + "val", "right")));
                        break;
                    case ParagraphAlignment.Justify:
                        properties.Add(new XElement(w + "jc", new XAttribute(w + "val", "both")));
                        break;
                }

                if (paragraph.Revision != null)
                {
                    properties.Add(new XElement(w + "rPr",
                        CreateRevisionElement(
                            paragraph.Revision.Type == RevisionType.Insert ? "ins" : "del",
                            paragraph.Revision)));
                }

                if (properties.HasElements)
                {
                    element.Add(properties);
                }

                int bookmarkId = 0;
                // 段落书签名与显式起点标记是同一书签的两种形态：标记已存在时按标记就地写出，不再补整段对。
                // The paragraph name and an explicit start marker are two shapes of the same bookmark; when the
                // marker exists it is written in place, so no whole-paragraph pair is added.
                if (!string.IsNullOrEmpty(paragraph.BookmarkName) && !HasStartMarker(paragraph, paragraph.BookmarkName))
                {
                    bookmarkId = ++_bookmarkId;
                    _bookmarkIdByName[paragraph.BookmarkName] = bookmarkId;
                    element.Add(new XElement(w + "bookmarkStart",
                        new XAttribute(w + "id", bookmarkId.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute(w + "name", paragraph.BookmarkName)));
                }

                // 能还原成 numbering.xml 的标记交给 Word 自己编号（numPr），只有无法映射的标记才退回把标记写成正文。
                // Markers that map onto numbering.xml are left for Word to number through numPr; only markers that
                // cannot be mapped fall back to writing the marker as literal text.
                if (numbering == null && !string.IsNullOrEmpty(paragraph.ListMarker))
                {
                    TextFormat markerFormat = new TextFormat();
                    for (int i = 0; i < paragraph.Inlines.Count; i++)
                    {
                        TextRun candidate = paragraph.Inlines[i] as TextRun;
                        if (candidate != null) { markerFormat = candidate.Format; break; }
                    }
                    string marker = paragraph.ListMarker;
                    if (!marker.EndsWith(" ", StringComparison.Ordinal)) marker += " ";
                    element.Add(CreateRun(marker, markerFormat, context));
                }

                for (int i = 0; i < paragraph.Inlines.Count; i++)
                {
                    AddInline(element, paragraph.Inlines[i], context);
                }
                AddDeletedInlines(element, paragraph.DeletedInlines, context);
                if (bookmarkId != 0)
                {
                    _bookmarkIdByName.Remove(paragraph.BookmarkName);
                    element.Add(new XElement(w + "bookmarkEnd",
                        new XAttribute(w + "id", bookmarkId.ToString(CultureInfo.InvariantCulture))));
                }
                return element;
            }

            /// <summary>
            /// 为带列表标记的段落分配编号引用：连续且形状相同的标记共用一条 numId 序列，
            /// 形状不同、列表被打断或换到另一个部件时新建一条序列（各自从标记里的起始值重新计数）。
            /// Assigns a numbering reference to a paragraph with a list marker: contiguous markers of the same shape
            /// share one numId sequence, while a different shape, an interrupted run or another part starts a new
            /// sequence that counts again from the start value recovered from the marker.
            /// </summary>
            /// <param name="paragraph">目标段落。The paragraph.</param>
            /// <param name="context">当前部件写出上下文。The writing context of the current part.</param>
            /// <returns>分配的编号序列；标记无法映射到 numbering.xml 时返回 null。The assigned sequence; null when the marker cannot be mapped onto numbering.xml.</returns>
            private NumberingDefinition AssignNumbering(Paragraph paragraph, PartContext context)
            {
                NumberingShape shape;
                string key;
                if (!TryCreateNumberingShape(paragraph, out shape, out key))
                {
                    _listRunOpen = false;
                    return null;
                }
                if (_listRunOpen && _activeListPart == context
                    && string.Equals(_activeListShape, key, StringComparison.Ordinal))
                {
                    return _numberingDefinitions[_numberingDefinitions.Count - 1];
                }

                int abstractNumId;
                if (!_abstractNumbersByShape.TryGetValue(key, out abstractNumId))
                {
                    abstractNumId = _nextAbstractNumberId++;
                    _abstractNumbersByShape[key] = abstractNumId;
                }
                NumberingDefinition definition = new NumberingDefinition();
                definition.NumId = _nextNumId++;
                definition.AbstractNumId = abstractNumId;
                definition.Shape = shape;
                _numberingDefinitions.Add(definition);
                _listRunOpen = true;
                _activeListShape = key;
                _activeListPart = context;
                return definition;
            }

            /// <summary>
            /// 从标记文本还原编号形状：项目符号、十进制（含补零）、字母或罗马数字，连同前后缀组成 lvlText 模板。
            /// 多级标记（如 "1.1"）与无法复现的写法返回 false，由调用方退回把标记写成正文文本。
            /// Recovers a numbering shape from the marker text: a bullet, decimal (including zero-padded), letter or
            /// Roman number, with its prefix and suffix forming the lvlText template. Multilevel markers such as
            /// "1.1" and shapes that cannot be reproduced return false so the caller writes the marker as text.
            /// </summary>
            /// <param name="paragraph">目标段落。The paragraph.</param>
            /// <param name="shape">还原出的编号形状。The recovered numbering shape.</param>
            /// <param name="key">形状键，相同键复用同一 abstractNum。The shape key; equal keys share one abstractNum.</param>
            /// <returns>可映射返回 true。True when the marker can be mapped.</returns>
            private static bool TryCreateNumberingShape(Paragraph paragraph, out NumberingShape shape, out string key)
            {
                shape = null;
                key = null;
                string marker = paragraph.ListMarker;
                if (string.IsNullOrEmpty(marker)) return false;
                string trimmed = marker.Trim();
                if (trimmed.Length == 0 || trimmed.IndexOf('%') >= 0) return false;

                int start = 0;
                while (start < trimmed.Length && !IsNumberTokenStart(trimmed[start])) start++;
                if (start >= trimmed.Length)
                {
                    if (!IsBulletMarker(trimmed)) return false;
                    shape = new NumberingShape();
                    shape.Bullet = true;
                    shape.Format = "bullet";
                    shape.LevelText = trimmed;
                    key = BuildNumberingShapeKey(shape);
                    return true;
                }

                string prefix = trimmed.Substring(0, start);
                int end = start;
                if (IsAsciiDigit(trimmed[start]))
                {
                    while (end < trimmed.Length && IsAsciiDigit(trimmed[end])) end++;
                }
                else
                {
                    while (end < trimmed.Length && IsAsciiLetter(trimmed[end])) end++;
                }
                string token = trimmed.Substring(start, end - start);
                string suffix = trimmed.Substring(end);
                // 后缀里再出现数字或字母说明这是多级编号（"1.1"、"1.a"），单一 ilvl=0 的编号复现不了。
                // A digit or letter after the token means a multilevel marker ("1.1", "1.a"), which a single
                // ilvl=0 sequence cannot reproduce.
                for (int i = 0; i < suffix.Length; i++)
                {
                    if (IsNumberTokenStart(suffix[i])) return false;
                }

                int value;
                string format;
                if (IsAsciiDigit(token[0]))
                {
                    if (!int.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out value) || value <= 0)
                    {
                        return false;
                    }
                    if (token.Length == 1)
                    {
                        format = "decimal";
                    }
                    else if (token.Length == 2 && token[0] == '0')
                    {
                        format = "decimalZero";
                    }
                    else
                    {
                        return false;
                    }
                }
                else if (token.Length == 1)
                {
                    format = char.IsUpper(token[0]) ? "upperLetter" : "lowerLetter";
                    value = char.ToLowerInvariant(token[0]) - 'a' + 1;
                }
                else if (TryParseRomanNumeral(token, out value))
                {
                    format = char.IsUpper(token[0]) ? "upperRoman" : "lowerRoman";
                }
                else
                {
                    return false;
                }

                shape = new NumberingShape();
                shape.Format = format;
                shape.LevelText = prefix + "%1" + suffix;
                shape.Start = value;
                key = BuildNumberingShapeKey(shape);
                return true;
            }

            /// <summary>把编号形状压成可比较的字典键（格式、lvlText、起始值三元组）。Flattens a numbering shape into a comparable dictionary key (format, lvlText and start value).</summary>
            /// <param name="shape">编号形状。The numbering shape.</param>
            /// <returns>形状键。The shape key.</returns>
            private static string BuildNumberingShapeKey(NumberingShape shape)
            {
                return shape.Format + "\u0001" + shape.LevelText + "\u0001"
                    + shape.Start.ToString(CultureInfo.InvariantCulture);
            }

            /// <summary>字符能否作为一个序号的开头（ASCII 数字或字母）。Whether the character can start a sequence number (an ASCII digit or letter).</summary>
            /// <param name="value">待判断字符。The character to test.</param>
            /// <returns>是返回 true。True when it can.</returns>
            private static bool IsNumberTokenStart(char value)
            {
                return IsAsciiDigit(value) || IsAsciiLetter(value);
            }

            /// <summary>是否为 ASCII 数字 0-9。Whether the character is an ASCII digit 0-9.</summary>
            /// <param name="value">待判断字符。The character to test.</param>
            /// <returns>是返回 true。True when it is.</returns>
            private static bool IsAsciiDigit(char value)
            {
                return value >= '0' && value <= '9';
            }

            /// <summary>是否为 ASCII 字母 a-z/A-Z。Whether the character is an ASCII letter a-z or A-Z.</summary>
            /// <param name="value">待判断字符。The character to test.</param>
            /// <returns>是返回 true。True when it is.</returns>
            private static bool IsAsciiLetter(char value)
            {
                return (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z');
            }

            /// <summary>标记是否完全由标点/符号组成，从而可以当作项目符号的 lvlText。Whether the marker consists only of punctuation or symbols and can serve as a bullet lvlText.</summary>
            /// <param name="marker">去掉首尾空白的标记。The marker with surrounding whitespace removed.</param>
            /// <returns>是返回 true。True when it can.</returns>
            private static bool IsBulletMarker(string marker)
            {
                for (int i = 0; i < marker.Length; i++)
                {
                    if (!char.IsPunctuation(marker[i]) && !char.IsSymbol(marker[i])) return false;
                }
                return true;
            }

            /// <summary>按罗马数字语法解析序号，并要求重新写出的文本与原标记一致（"IIII" 之类不可复现的写法被拒绝）。Parses a Roman numeral and requires the canonical rendering to match the marker, rejecting shapes like "IIII" that cannot be reproduced.</summary>
            /// <param name="token">标记里的序号文本。The number text inside the marker.</param>
            /// <param name="value">解析出的序号。The parsed number.</param>
            /// <returns>解析成功返回 true。True when the token is a canonical Roman numeral.</returns>
            private static bool TryParseRomanNumeral(string token, out int value)
            {
                value = 0;
                string upper = token.ToUpperInvariant();
                for (int i = 0; i < upper.Length; i++)
                {
                    int current = RomanDigitValue(upper[i]);
                    if (current == 0) return false;
                    int next = i + 1 < upper.Length ? RomanDigitValue(upper[i + 1]) : 0;
                    value += current < next ? -current : current;
                }
                if (value <= 0 || value > 3999) return false;
                return string.Equals(ToRomanText(value), upper, StringComparison.Ordinal);
            }

            /// <summary>单个罗马数字字符对应的数值，非罗马数字返回 0。The value of one Roman digit character; 0 for non-Roman characters.</summary>
            /// <param name="value">罗马数字字符。The Roman digit character.</param>
            /// <returns>数值。The digit value.</returns>
            private static int RomanDigitValue(char value)
            {
                switch (value)
                {
                    case 'I': return 1;
                    case 'V': return 5;
                    case 'X': return 10;
                    case 'L': return 50;
                    case 'C': return 100;
                    case 'D': return 500;
                    case 'M': return 1000;
                    default: return 0;
                }
            }

            /// <summary>把 1-3999 写成规范的大写罗马数字，用于校验标记能否复现。Renders 1-3999 as a canonical upper-case Roman numeral so the marker can be verified as reproducible.</summary>
            /// <param name="value">序号。The number.</param>
            /// <returns>罗马数字文本。The Roman numeral text.</returns>
            private static string ToRomanText(int value)
            {
                int[] values = { 1000, 900, 500, 400, 100, 90, 50, 40, 10, 9, 5, 4, 1 };
                string[] symbols = { "M", "CM", "D", "CD", "C", "XC", "L", "XL", "X", "IX", "V", "IV", "I" };
                System.Text.StringBuilder builder = new System.Text.StringBuilder();
                int remaining = value;
                for (int i = 0; i < values.Length; i++)
                {
                    while (remaining >= values[i])
                    {
                        builder.Append(symbols[i]);
                        remaining -= values[i];
                    }
                }
                return builder.ToString();
            }

            /// <summary>
            /// 生成 numbering.xml：每个编号形状一个 abstractNum（9 级，第 0 级按形状定义，深层级按缩进递增），
            /// 每段连续列表一个 num 指向对应 abstractNum。abstractNum 必须全部排在 num 之前。
            /// Builds numbering.xml: one abstractNum per numbering shape (nine levels, level 0 following the shape
            /// and deeper levels stepping the indent) and one num per contiguous list run pointing at it. Every
            /// abstractNum has to precede the num elements.
            /// </summary>
            /// <returns>numbering.xml 文档。The numbering.xml document.</returns>
            private XDocument CreateNumberingXml()
            {
                XNamespace w = WordNamespace;
                XElement root = new XElement(w + "numbering",
                    new XAttribute(XNamespace.Xmlns + "w", w));
                Dictionary<int, NumberingShape> shapes = new Dictionary<int, NumberingShape>();
                List<int> abstractIds = new List<int>();
                for (int i = 0; i < _numberingDefinitions.Count; i++)
                {
                    NumberingDefinition definition = _numberingDefinitions[i];
                    if (shapes.ContainsKey(definition.AbstractNumId)) continue;
                    shapes[definition.AbstractNumId] = definition.Shape;
                    abstractIds.Add(definition.AbstractNumId);
                }
                abstractIds.Sort();
                for (int i = 0; i < abstractIds.Count; i++)
                {
                    root.Add(CreateAbstractNumbering(abstractIds[i], shapes[abstractIds[i]]));
                }
                for (int i = 0; i < _numberingDefinitions.Count; i++)
                {
                    NumberingDefinition definition = _numberingDefinitions[i];
                    XElement num = new XElement(w + "num",
                        new XAttribute(w + "numId", definition.NumId.ToString(CultureInfo.InvariantCulture)),
                        new XElement(w + "abstractNumId",
                            new XAttribute(w + "val", definition.AbstractNumId.ToString(CultureInfo.InvariantCulture))));
                    // Word 的计数器实际记在 abstractNum 上：同一个 abstractNum 被第二个 numId 复用时，
                    // 不写 w:startOverride 就会接着上一段列表继续排（打断后重新起号变成 4、5）。
                    // Word keeps the counter on the abstractNum, so when a second numId reuses the same
                    // abstractNum the sequence continues unless w:startOverride restarts it — otherwise a
                    // list resumed after an interruption would carry on at 4, 5 instead of starting again.
                    if (!definition.Shape.Bullet)
                    {
                        num.Add(new XElement(w + "lvlOverride",
                            new XAttribute(w + "ilvl", "0"),
                            new XElement(w + "startOverride",
                                new XAttribute(w + "val", definition.Shape.Start.ToString(CultureInfo.InvariantCulture)))));
                    }
                    root.Add(num);
                }
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 生成一个 abstractNum：9 个 w:lvl，每级写出起始值、格式、lvlText、对齐与逐级递增的悬挂缩进。
            /// Builds one abstractNum: nine w:lvl entries, each with its start value, format, lvlText, alignment
            /// and a hanging indent that steps out with the level.
            /// </summary>
            /// <param name="abstractNumId">抽象编号定义 ID。The abstract numbering definition id.</param>
            /// <param name="shape">本抽象定义的编号形状。The numbering shape of this definition.</param>
            /// <returns>w:abstractNum 元素。The w:abstractNum element.</returns>
            private static XElement CreateAbstractNumbering(int abstractNumId, NumberingShape shape)
            {
                XNamespace w = WordNamespace;
                XElement abstractNum = new XElement(w + "abstractNum",
                    new XAttribute(w + "abstractNumId", abstractNumId.ToString(CultureInfo.InvariantCulture)),
                    new XElement(w + "multiLevelType", new XAttribute(w + "val", "hybridMultilevel")));
                for (int level = 0; level < 9; level++)
                {
                    string index = (level + 1).ToString(CultureInfo.InvariantCulture);
                    XElement lvl = new XElement(w + "lvl", new XAttribute(w + "ilvl", level.ToString(CultureInfo.InvariantCulture)));
                    if (!shape.Bullet)
                    {
                        lvl.Add(new XElement(w + "start",
                            new XAttribute(w + "val", (level == 0 ? shape.Start : 1).ToString(CultureInfo.InvariantCulture))));
                    }
                    lvl.Add(new XElement(w + "numFmt", new XAttribute(w + "val", shape.Format)));
                    string levelText = level == 0 || shape.Bullet
                        ? shape.LevelText
                        : shape.LevelText.Replace("%1", "%" + index);
                    lvl.Add(new XElement(w + "lvlText", new XAttribute(w + "val", levelText)));
                    lvl.Add(new XElement(w + "lvlJc", new XAttribute(w + "val", "left")));
                    lvl.Add(new XElement(w + "pPr", new XElement(w + "ind",
                        new XAttribute(w + "left", (720 * (level + 1)).ToString(CultureInfo.InvariantCulture)),
                        new XAttribute(w + "hanging", "360"))));
                    abstractNum.Add(lvl);
                }
                return abstractNum;
            }

            /// <summary>段落内是否已有该名称的显式书签起点。Determines whether the paragraph already holds an explicit start marker of the name.</summary>
            /// <param name="paragraph">目标段落。The paragraph.</param>
            /// <param name="name">书签名。The bookmark name.</param>
            /// <returns>存在返回 true。True when present.</returns>
            private static bool HasStartMarker(Paragraph paragraph, string name)
            {
                List<Inline> inlines = paragraph.Inlines;
                for (int i = 0; i < inlines.Count; i++)
                {
                    BookmarkStart start = inlines[i] as BookmarkStart;
                    if (start != null && string.Equals(start.Name, name, StringComparison.Ordinal)) return true;
                }
                return false;
            }

            /// <summary>
            /// 按行内类型写出单个元素，并按需包裹 w:ins 修订；无法识别的类型被忽略。
            /// Writes a single inline element by type, wrapping it in a w:ins revision when required; unrecognised
            /// types are ignored.
            /// </summary>
            /// <param name="parent">接收元素的父节点。The parent node receiving the element.</param>
            /// <param name="inline">行内元素。The inline element.</param>
            /// <param name="context">当前部件写出上下文。The current part writing context.</param>
            private void AddInline(XElement parent, Inline inline, PartContext context)
            {
                XNamespace w = WordNamespace;
                RevisionInfo revision = inline == null ? null : inline.Revision;
                if (revision != null && revision.Type == RevisionType.Delete)
                {
                    TextRun deletedText = inline as TextRun;
                    if (deletedText != null && !string.IsNullOrEmpty(deletedText.Text))
                    {
                        XElement deletedRun = CreateDeletedRun(deletedText.Text, deletedText.Format);
                        XElement deletedLink = string.IsNullOrEmpty(deletedText.Link) ? null : CreateHyperlink(deletedText.Link, deletedRun, context);
                        XElement wrapper = CreateRevisionElement("del", revision);
                        wrapper.Add(deletedLink ?? deletedRun);
                        parent.Add(wrapper);
                    }
                    return;
                }

                XElement content = null;
                BookmarkStart markerStart = inline as BookmarkStart;
                if (markerStart != null)
                {
                    int startId = ++_bookmarkId;
                    if (!string.IsNullOrEmpty(markerStart.Name)) _bookmarkIdByName[markerStart.Name] = startId;
                    parent.Add(new XElement(w + "bookmarkStart",
                        new XAttribute(w + "id", startId.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute(w + "name", markerStart.Name ?? string.Empty)));
                    return;
                }
                BookmarkEnd markerEnd = inline as BookmarkEnd;
                if (markerEnd != null)
                {
                    // 终点只带 id：复用起点分配的 id，起点未写出（跨部件或落单）时整个终点略过，
                    // 以免产生无主 w:bookmarkEnd。The end carries only an id: it reuses the start's id, and an end
                    // whose start was never written (another part, or unpaired) is dropped instead of orphaned.
                    int endId;
                    if (!string.IsNullOrEmpty(markerEnd.Name)
                        && _bookmarkIdByName.TryGetValue(markerEnd.Name, out endId))
                    {
                        _bookmarkIdByName.Remove(markerEnd.Name);
                        parent.Add(new XElement(w + "bookmarkEnd",
                            new XAttribute(w + "id", endId.ToString(CultureInfo.InvariantCulture))));
                    }
                    return;
                }
                TextRun text = inline as TextRun;
                if (text != null)
                {
                    // TextRun.Text 是公开可写属性，宿主可以置空；这里按空串处理而不是解引用。
                    // TextRun.Text exposes a public setter, so a host can null it; treat that as empty
                    // text instead of dereferencing it.
                    if (string.IsNullOrEmpty(text.Text)) return;
                    XElement run = CreateRun(text.Text, text.Format, context);
                    XElement hyperlink = string.IsNullOrEmpty(text.Link) ? null : CreateHyperlink(text.Link, run, context);
                    content = hyperlink ?? run;
                }
                else
                {
                    FieldRun field = inline as FieldRun;
                    if (field != null)
                    {
                        content = CreateFieldRun(field, context);
                    }
                    else
                    {
                        ImageRun image = inline as ImageRun;
                        if (image != null) content = CreateImageRun(image, context);
                    }
                }
                if (content == null) return;
                if (revision != null && revision.Type == RevisionType.Insert)
                {
                    XElement wrapper = CreateRevisionElement("ins", revision);
                    wrapper.Add(content);
                    parent.Add(wrapper);
                    return;
                }
                parent.Add(content);
            }

            /// <summary>
            /// 写出被修订删除的行内内容：连续且修订元数据相同的片段合并到一个 w:del 中，
            /// 文本改用 w:delText（图片不写出）。
            /// Writes inline content removed by tracked changes: consecutive fragments sharing the same revision
            /// metadata are merged into one w:del, and text is emitted as w:delText (images are not written).
            /// </summary>
            /// <param name="parent">接收元素的父节点。The parent node receiving the elements.</param>
            /// <param name="deletedInlines">被删除的行内元素列表。The list of deleted inline elements.</param>
            /// <param name="context">当前部件写出上下文。The current part writing context.</param>
            private void AddDeletedInlines(XElement parent, List<Inline> deletedInlines, PartContext context)
            {
                if (deletedInlines == null) return;
                XElement wrapper = null;
                RevisionInfo current = null;
                for (int i = 0; i < deletedInlines.Count; i++)
                {
                    TextRun text = deletedInlines[i] as TextRun;
                    string value = text != null ? text.Text : null;
                    TextFormat format = text != null ? text.Format : null;
                    if (text == null)
                    {
                        FieldRun field = deletedInlines[i] as FieldRun;
                        if (field == null || string.IsNullOrEmpty(field.DisplayText)) continue;
                        value = field.DisplayText;
                        format = field.Format;
                    }
                    if (value.Length == 0) continue;

                    RevisionInfo revision = deletedInlines[i].Revision ?? new RevisionInfo(RevisionType.Delete, string.Empty, null);
                    if (wrapper == null || !SameRevision(current, revision))
                    {
                        wrapper = CreateRevisionElement("del", revision);
                        parent.Add(wrapper);
                        current = revision;
                    }
                    XElement run = CreateDeletedRun(value, format);
                    XElement hyperlink = text == null || string.IsNullOrEmpty(text.Link) ? null : CreateHyperlink(text.Link, run, context);
                    wrapper.Add(hyperlink ?? run);
                }
            }

            /// <summary>
            /// 判断两条修订是否具有相同的作者、时间与类型（用于合并相邻删除片段）。
            /// Determines whether two revisions share the author, timestamp and type (used to merge adjacent
            /// deleted fragments).
            /// </summary>
            /// <param name="left">左侧修订，可为 null。The left revision, may be null.</param>
            /// <param name="right">右侧修订。The right revision.</param>
            /// <returns>元数据一致时返回 true。True when the metadata matches.</returns>
            private static bool SameRevision(RevisionInfo left, RevisionInfo right)
            {
                if (left == null || right == null) return false;
                return left.Type == right.Type
                    && string.Equals(left.Author ?? string.Empty, right.Author ?? string.Empty, StringComparison.Ordinal)
                    && left.Date == right.Date;
            }

            /// <summary>
            /// 生成修订容器元素（w:ins / w:del），写入重新分配的唯一 ID、作者与 UTC 时间。
            /// Builds a revision container element (w:ins / w:del) carrying a freshly assigned unique id, the
            /// author and a UTC timestamp.
            /// </summary>
            /// <param name="name">元素本地名称（ins 或 del）。The element local name (ins or del).</param>
            /// <param name="revision">修订元数据，可为 null。The revision metadata, may be null.</param>
            /// <returns>修订容器元素。The revision container element.</returns>
            private XElement CreateRevisionElement(string name, RevisionInfo revision)
            {
                XNamespace w = WordNamespace;
                RevisionInfo info = revision ?? new RevisionInfo();
                DateTime date = info.Date.HasValue ? info.Date.Value : DateTime.UtcNow;
                if (date.Kind == DateTimeKind.Unspecified) date = DateTime.SpecifyKind(date, DateTimeKind.Utc);
                string author = string.IsNullOrEmpty(info.Author) ? "Wisdom.Office" : info.Author;
                return new XElement(w + name,
                    new XAttribute(w + "id", (++_revisionId).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute(w + "author", author),
                    new XAttribute(w + "date", date.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)));
            }

            /// <summary>
            /// 生成被删除文本的运行：与普通运行结构一致，但 w:t 全部替换为 w:delText。
            /// Builds a run for deleted text: same structure as a normal run, but every w:t is replaced by w:delText.
            /// </summary>
            /// <param name="text">被删除的文本。The deleted text.</param>
            /// <param name="format">文字格式，可为 null。The text format, may be null.</param>
            /// <returns>含 delText 的 w:r 元素。The w:r element carrying delText.</returns>
            private XElement CreateDeletedRun(string text, TextFormat format)
            {
                XNamespace w = WordNamespace;
                XElement run = CreateRun(text, format, null);
                foreach (XElement child in run.Elements().ToList())
                {
                    if (child.Name != w + "t") continue;
                    XElement deleted = new XElement(w + "delText", child.Value);
                    XAttribute space = child.Attribute(XNamespace.Xml + "space");
                    if (space != null) deleted.Add(new XAttribute(XNamespace.Xml + "space", space.Value));
                    child.ReplaceWith(deleted);
                }
                return run;
            }

            /// <summary>
            /// 生成 w:fldSimple 简单域元素，标记 dirty=true 以便 Word 打开时刷新域结果。
            /// Builds a w:fldSimple field element marked dirty=true so Word refreshes the field result on open.
            /// </summary>
            /// <param name="field">域模型。The field model.</param>
            /// <param name="context">当前部件写出上下文。The current part writing context.</param>
            /// <returns>fldSimple 元素。The fldSimple element.</returns>
            private XElement CreateFieldRun(FieldRun field, PartContext context)
            {
                XNamespace w = WordNamespace;
                XElement simple = new XElement(w + "fldSimple",
                    new XAttribute(w + "instr", field.Instruction ?? string.Empty),
                    new XAttribute(w + "dirty", "true"));
                if (!string.IsNullOrEmpty(field.DisplayText))
                {
                    simple.Add(CreateRun(field.DisplayText, field.Format, context));
                }
                return simple;
            }

            /// <summary>
            /// 预留判断段落是否为「仅含换行的空段」；当前实现恒返回 false。
            /// Reserved check for whether a paragraph is an empty break-only paragraph; the current
            /// implementation always returns false.
            /// </summary>
            /// <param name="paragraph">待检测段落。The paragraph to test.</param>
            /// <returns>恒为 false。Always false.</returns>
            private static bool IsEmptyBreak(Paragraph paragraph)
            {
                return false;
            }

            /// <summary>
            /// 生成仅含分页符（w:br type="page"）的段落，用于 <see cref="PageBreak"/> 块。
            /// Builds a paragraph containing only a page break (w:br type="page") for <see cref="PageBreak"/> blocks.
            /// </summary>
            /// <returns>带分页运行的 w:p 元素。The w:p element carrying the page-break run.</returns>
            private XElement CreatePageBreakParagraph()
            {
                XNamespace w = WordNamespace;
                return new XElement(w + "p",
                    new XElement(w + "r", new XElement(w + "br", new XAttribute(w + "type", "page"))));
            }

            /// <summary>
            /// 生成运行元素 w:r：按格式生成 rPr，并把文本中的 \n、\t 拆分为 w:br / w:tab 与多个 w:t。
            /// Builds the w:r run element: renders rPr from the format and splits \n and \t in the text into
            /// w:br / w:tab elements and several w:t nodes.
            /// </summary>
            /// <param name="text">运行文本。The run text.</param>
            /// <param name="format">文字格式，可为 null。The text format, may be null.</param>
            /// <param name="context">当前部件写出上下文（图片关系需要，此处仅透传）。The current part writing context (needed for image relationships, passed through here).</param>
            /// <returns>w:r 元素。The w:r element.</returns>
            private XElement CreateRun(string text, TextFormat format, PartContext context)
            {
                XNamespace w = WordNamespace;
                text = text ?? string.Empty;
                XElement run = new XElement(w + "r");
                if (format != null) run.Add(CreateRunProperties(format));
                string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
                StringBuilder buffer = new StringBuilder();
                for (int i = 0; i < normalized.Length; i++)
                {
                    char c = normalized[i];
                    if (c == '\n' || c == '\t')
                    {
                        FlushText(run, buffer);
                        run.Add(c == '\n' ? new XElement(w + "br") : new XElement(w + "tab"));
                        continue;
                    }
                    buffer.Append(c);
                }
                FlushText(run, buffer);
                return run;
            }

            /// <summary>
            /// 将缓冲区文本作为带 xml:space="preserve" 的 w:t 元素刷入运行并清空缓冲区。
            /// Flushes the buffered text into the run as a w:t element with xml:space="preserve" and clears
            /// the buffer.
            /// </summary>
            /// <param name="run">目标 w:r 元素。The target w:r element.</param>
            /// <param name="buffer">文本缓冲区。The text buffer.</param>
            private static void FlushText(XElement run, StringBuilder buffer)
            {
                if (buffer.Length == 0) return;
                XNamespace w = WordNamespace;
                XElement text = new XElement(w + "t", buffer.ToString());
                text.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                run.Add(text);
                buffer.Length = 0;
            }

            /// <summary>
            /// 用 w:hyperlink 包裹运行：以 "#" 开头的目标写成表内锚点 w:anchor，其余经协议校验后
            /// 建立 External 关系并引用 r:id；目标为空或不安全时返回 null，由调用方退回普通运行。
            /// Wraps a run in w:hyperlink: targets starting with "#" become an in-document w:anchor, other
            /// targets are validated and written as an External relationship referenced by r:id. Returns null for
            /// empty or unsafe targets so the caller can fall back to a plain run.
            /// </summary>
            /// <param name="target">超链接目标。The hyperlink target.</param>
            /// <param name="run">被包裹的运行元素。The run element to wrap.</param>
            /// <param name="context">当前部件写出上下文。The current part writing context.</param>
            /// <returns>w:hyperlink 元素，或 null。The w:hyperlink element, or null.</returns>
            private static XElement CreateHyperlink(string target, XElement run, PartContext context)
            {
                XNamespace w = WordNamespace;
                XNamespace r = OfficeRelationshipNamespace;
                string trimmed = target.Trim();
                if (trimmed.Length == 0) return null;
                XElement hyperlink = new XElement(w + "hyperlink");

                if (trimmed[0] == '#')
                {
                    string anchor = trimmed.Substring(1).Trim();
                    if (anchor.Length == 0) return null;
                    hyperlink.Add(new XAttribute(w + "anchor", anchor));
                    hyperlink.Add(new XAttribute(w + "history", "1"));
                    hyperlink.Add(run);
                    return hyperlink;
                }

                if (!WordFormatUtil.IsSafeHyperlinkTarget(trimmed)) return null;
                hyperlink.Add(new XAttribute(r + "id", context.GetHyperlinkRelationship(trimmed)));
                hyperlink.Add(new XAttribute(w + "history", "1"));
                hyperlink.Add(run);
                return hyperlink;
            }

            /// <summary>
            /// 生成运行属性 w:rPr：中西文字体四件套（ascii/hAnsi/eastAsia/cs）、加粗倾斜删除线、
            /// 颜色、半磅字号、下划线与上下标。字体、字号、颜色与文档默认值相同时一律省略，让样式链生效。
            /// Builds the run properties w:rPr: the four East-Asian/Latin font slots (ascii/hAnsi/eastAsia/cs),
            /// bold/italic/strike-through, colour, half-point size, underline and super/subscript. Font, size and
            /// colour equal to the document default are omitted so the style chain can still apply.
            /// </summary>
            /// <remarks>
            /// 每个 TextRun 都由 <c>new TextFormat()</c> 起步（Arial、11 磅、黑色），早先无条件写出 rFonts/sz/color，
            /// 运行级属性优先级高于段落样式，标题段因此被压回正文的 11 磅黑色 Arial——标题样式里的字号和
            /// 1F3A5F 颜色全部失效。docDefaults 已经写了同样的默认字体与字号，所以与默认值一致的覆盖可以直接省略。
            /// Every TextRun starts from <c>new TextFormat()</c> (Arial, 11pt, black), and the old code always wrote
            /// rFonts/sz/color. Run-level properties beat paragraph styles, so headings were flattened back to
            /// 11pt black Arial — the heading style's size and 1F3A5F colour never applied. docDefaults already
            /// carries the same default font and size, so overrides equal to the default are safe to omit.
            /// </remarks>
            /// <param name="format">文字格式。The text format.</param>
            /// <returns>w:rPr 元素。The w:rPr element.</returns>
            private XElement CreateRunProperties(TextFormat format)
            {
                XNamespace w = WordNamespace;
                XElement properties = new XElement(w + "rPr");
                TextFormat defaults = _document != null && _document.DefaultTextFormat != null
                    ? _document.DefaultTextFormat
                    : new TextFormat();
                double defaultSize = defaults.FontSize > 0 ? defaults.FontSize : 11;
                string asciiFont = format.GetAsciiFontFamily();
                string eastAsiaFont = format.GetEastAsiaFontFamily();
                string defaultFont = string.IsNullOrEmpty(defaults.FontFamily) ? "Calibri" : defaults.FontFamily;
                string defaultAscii = defaults.GetAsciiFontFamily();
                if (string.IsNullOrEmpty(defaultAscii)) defaultAscii = defaultFont;
                string defaultEastAsia = defaults.GetEastAsiaFontFamily();
                if (string.IsNullOrEmpty(defaultEastAsia)) defaultEastAsia = defaultFont;
                if ((!string.IsNullOrEmpty(asciiFont) || !string.IsNullOrEmpty(eastAsiaFont))
                    && (!string.Equals(asciiFont, defaultAscii, StringComparison.Ordinal)
                        || !string.Equals(eastAsiaFont, defaultEastAsia, StringComparison.Ordinal)))
                {
                    string latin = string.IsNullOrEmpty(asciiFont) ? eastAsiaFont : asciiFont;
                    string cjk = string.IsNullOrEmpty(eastAsiaFont) ? asciiFont : eastAsiaFont;
                    properties.Add(new XElement(w + "rFonts",
                        new XAttribute(w + "ascii", latin),
                        new XAttribute(w + "hAnsi", latin),
                        new XAttribute(w + "eastAsia", cjk),
                        new XAttribute(w + "cs", latin)));
                }
                if (format.Bold) properties.Add(new XElement(w + "b"));
                if (format.Italic) properties.Add(new XElement(w + "i"));
                if (format.Strikeout) properties.Add(new XElement(w + "strike"));
                string colorHex = format.Color.ToHex();
                string defaultColorHex = defaults.Color.ToHex();
                if (!string.Equals(colorHex, defaultColorHex, StringComparison.OrdinalIgnoreCase))
                {
                    properties.Add(new XElement(w + "color", new XAttribute(w + "val", colorHex)));
                }
                if (format.FontSize > 0 && Math.Abs(format.FontSize - defaultSize) > 0.001)
                {
                    string halfPoints = ClampToInt(format.FontSize * 2, 2, 32767).ToString(CultureInfo.InvariantCulture);
                    properties.Add(new XElement(w + "sz", new XAttribute(w + "val", halfPoints)));
                    properties.Add(new XElement(w + "szCs", new XAttribute(w + "val", halfPoints)));
                }
                if (format.Underline) properties.Add(new XElement(w + "u", new XAttribute(w + "val", "single")));
                if (format.Superscript) properties.Add(new XElement(w + "vertAlign", new XAttribute(w + "val", "superscript")));
                else if (format.Subscript) properties.Add(new XElement(w + "vertAlign", new XAttribute(w + "val", "subscript")));
                return properties;
            }

            /// <summary>
            /// 生成内联图片运行：磅尺寸换算为 EMU，构造 wp:inline + a:graphic + pic:pic 的
            /// DrawingML 结构并引用部件关系 ID；图片无关系时返回空运行。
            /// Builds the inline image run: converts point sizes to EMUs, constructs the wp:inline + a:graphic
            /// + pic:pic DrawingML structure referencing the part relationship id, and returns an empty run
            /// when the image has no relationship.
            /// </summary>
            /// <param name="image">图片运行模型。The image run model.</param>
            /// <param name="context">当前部件写出上下文（用于分配/复用图片关系）。The current part writing context (allocates or reuses image relationships).</param>
            /// <returns>含 w:drawing 的 w:r 元素。The w:r element carrying w:drawing.</returns>
            private XElement CreateImageRun(ImageRun image, PartContext context)
            {
                XNamespace w = WordNamespace;
                XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
                XNamespace r = OfficeRelationshipNamespace;

                string relationshipId = context.GetImageRelationship(image.Image);
                if (relationshipId == null) return new XElement(w + "r");
                double width = image.Width > 0 ? image.Width : image.Image.DefaultWidthPoints;
                double height = image.Height > 0 ? image.Height : image.Image.DefaultHeightPoints;
                long cx = (long)Math.Round(UnitConverter.PointsToEmu(width));
                long cy = (long)Math.Round(UnitConverter.PointsToEmu(height));
                int id = ++_drawingId;
                string name = "Picture " + id.ToString(CultureInfo.InvariantCulture);

                string idText = id.ToString(CultureInfo.InvariantCulture);
                string cxText = cx.ToString(CultureInfo.InvariantCulture);
                string cyText = cy.ToString(CultureInfo.InvariantCulture);

                XElement pictureShape = new XElement(pic + "spPr",
                    new XElement(a + "xfrm",
                        new XElement(a + "off", new XAttribute("x", "0"), new XAttribute("y", "0")),
                        new XElement(a + "ext", new XAttribute("cx", cxText), new XAttribute("cy", cyText))),
                    new XElement(a + "prstGeom", new XAttribute("prst", "rect"),
                        new XElement(a + "avLst")));
                XElement picture = new XElement(pic + "pic",
                    new XElement(pic + "nvPicPr",
                        new XElement(pic + "cNvPr", new XAttribute("id", idText), new XAttribute("name", name)),
                        new XElement(pic + "cNvPicPr")),
                    new XElement(pic + "blipFill",
                        new XElement(a + "blip", new XAttribute(r + "embed", relationshipId)),
                        new XElement(a + "stretch", new XElement(a + "fillRect"))),
                    pictureShape);
                XElement graphic = new XElement(a + "graphic",
                    new XElement(a + "graphicData",
                        new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/picture"),
                        picture));
                XElement inline = new XElement(wp + "inline",
                    new XAttribute("distT", "0"),
                    new XAttribute("distB", "0"),
                    new XAttribute("distL", "0"),
                    new XAttribute("distR", "0"),
                    new XElement(wp + "extent", new XAttribute("cx", cxText), new XAttribute("cy", cyText)),
                    new XElement(wp + "effectExtent", new XAttribute("l", "0"), new XAttribute("t", "0"), new XAttribute("r", "0"), new XAttribute("b", "0")),
                    new XElement(wp + "docPr", new XAttribute("id", idText), new XAttribute("name", name)),
                    new XElement(wp + "cNvGraphicFramePr",
                        new XElement(a + "graphicFrameLocks", new XAttribute("noChangeAspect", "1"))),
                    graphic);
                XElement drawing = new XElement(w + "drawing", inline);
                return new XElement(w + "r", drawing);
            }

            /// <summary>
            /// 生成表格元素 w:tbl：表格属性（宽度/对齐/边框/边距/布局）、网格列定义，
            /// 并按 pendingMerges 状态机把纵向合并展开为后续行的 vMerge continue 单元格。
            /// Builds the w:tbl table element: table properties (width/alignment/borders/margins/layout), the
            /// grid column definitions, and the pendingMerges state machine expanding vertical merges into
            /// vMerge continue cells in the following rows.
            /// </summary>
            /// <param name="table">表格模型。The table model.</param>
            /// <param name="context">当前部件写出上下文。The current part writing context.</param>
            /// <param name="contentWidth">可用内容宽度（磅），用于列宽分配。The content width in points used for column sizing.</param>
            /// <returns>w:tbl 元素。The w:tbl element.</returns>
            private XElement CreateTable(Table table, PartContext context, double contentWidth)
            {
                XNamespace w = WordNamespace;
                double[] grid = ComputeGrid(table, contentWidth);
                double tableWidth = 0;
                for (int i = 0; i < grid.Length; i++) tableWidth += grid[i];

                XElement properties = new XElement(w + "tblPr");
                if (table.Autofit)
                {
                    properties.Add(new XElement(w + "tblW",
                        new XAttribute(w + "w", "5000"),
                        new XAttribute(w + "type", "pct")));
                }
                else
                {
                    properties.Add(new XElement(w + "tblW",
                        new XAttribute(w + "w", ToTwips(tableWidth)),
                        new XAttribute(w + "type", "dxa")));
                }
                // 表格对齐的元素名是 w:jc（CT_JcTable），不是 w:tblJc——后者不在 ECMA-376 词汇表里，
                // Word 会忽略它（表格仍然左对齐），严格校验器还会判为非法。
                // The table-alignment element is w:jc (CT_JcTable), not w:tblJc — the latter is not in the ECMA-376
                // vocabulary, so Word ignores it (the table stays left-aligned) and a strict validator flags it.
                if (table.Alignment == TableAlignment.Center)
                {
                    properties.Add(new XElement(w + "jc", new XAttribute(w + "val", "center")));
                }
                else if (table.Alignment == TableAlignment.Right)
                {
                    properties.Add(new XElement(w + "jc", new XAttribute(w + "val", "right")));
                }
                properties.Add(CreateTableBorders(table.Borders));
                properties.Add(new XElement(w + "tblLayout",
                    new XAttribute(w + "type", table.Autofit ? "autofit" : "fixed")));
                XElement margins = new XElement(w + "tblCellMar");
                double padding = Math.Max(0, table.CellPadding);
                margins.Add(new XElement(w + "top", new XAttribute(w + "w", ToTwips(padding)), new XAttribute(w + "type", "dxa")));
                margins.Add(new XElement(w + "left", new XAttribute(w + "w", ToTwips(padding)), new XAttribute(w + "type", "dxa")));
                margins.Add(new XElement(w + "bottom", new XAttribute(w + "w", ToTwips(padding)), new XAttribute(w + "type", "dxa")));
                margins.Add(new XElement(w + "right", new XAttribute(w + "w", ToTwips(padding)), new XAttribute(w + "type", "dxa")));
                properties.Add(margins);
                properties.Add(new XElement(w + "tblLook",
                    new XAttribute(w + "val", "04A0"),
                    new XAttribute(w + "firstRow", "1"),
                    new XAttribute(w + "lastRow", "0"),
                    new XAttribute(w + "firstColumn", "0"),
                    new XAttribute(w + "lastColumn", "0"),
                    new XAttribute(w + "noHBand", "0"),
                    new XAttribute(w + "noVBand", "1")));

                XElement element = new XElement(w + "tbl", properties);
                XElement gridElement = new XElement(w + "tblGrid");
                for (int i = 0; i < grid.Length; i++)
                {
                    gridElement.Add(new XElement(w + "gridCol", new XAttribute(w + "w", ToTwips(grid[i]))));
                }
                element.Add(gridElement);

                Dictionary<int, MergeState> pendingMerges = new Dictionary<int, MergeState>();
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    TableRow row = table.Rows[r];
                    XElement rowElement = new XElement(w + "tr");
                    XElement rowProperties = new XElement(w + "trPr");
                    if (row.Height > 0)
                    {
                        rowProperties.Add(new XElement(w + "trHeight",
                            new XAttribute(w + "val", ToTwips(row.Height)),
                            new XAttribute(w + "hRule", "atLeast")));
                    }
                    if (row.IsHeader || (r == 0 && table.RepeatHeaderRow))
                    {
                        rowProperties.Add(new XElement(w + "tblHeader"));
                    }
                    if (rowProperties.HasElements) rowElement.Add(rowProperties);

                    int cellIndex = 0;
                    for (int column = 0; column < grid.Length; )
                    {
                        MergeState merge;
                        if (pendingMerges.TryGetValue(column, out merge) && merge.Remaining > 0)
                        {
                            rowElement.Add(CreateCell(null, merge.Span, grid, column, true, context));
                            if (merge.Remaining > 1) merge.Remaining--;
                            else pendingMerges.Remove(column);
                            column += Math.Max(1, merge.Span);
                            continue;
                        }

                        if (cellIndex >= row.Cells.Count)
                        {
                            rowElement.Add(CreateCell(null, 1, grid, column, false, context));
                            column++;
                            continue;
                        }
                        TableCell cell = row.Cells[cellIndex++];
                        int span = Math.Max(1, cell.ColSpan);
                        if (column + span > grid.Length) span = Math.Max(1, grid.Length - column);
                        rowElement.Add(CreateCell(cell, span, grid, column, false, context));

                        int rowSpan = Math.Max(1, cell.RowSpan);
                        if (rowSpan > 1)
                        {
                            MergeState state = new MergeState();
                            state.Remaining = rowSpan - 1;
                            state.Span = span;
                            pendingMerges[column] = state;
                        }
                        column += span;
                    }
                    element.Add(rowElement);
                }
                return element;
            }

            /// <summary>纵向合并进行中的状态：剩余待填充的行数与占用的网格列数。State of an in-progress vertical merge: the rows still to fill and the grid columns occupied.</summary>
            private sealed class MergeState
            {
                /// <summary>后续仍需生成 continue 单元格的行数。The number of following rows that still need a continue cell.</summary>
                public int Remaining;
                /// <summary>该合并占用的横向网格跨度。The horizontal grid span occupied by the merge.</summary>
                public int Span;
            }

            /// <summary>
            /// 生成单元格元素 w:tc：宽度（按网格跨度累加）、gridSpan、vMerge restart/continue、
            /// 边框、底色、垂直对齐及块内容；cell 为 null 时生成空占位单元格。
            /// Builds the w:tc cell element: width (summed over the grid span), gridSpan, vMerge
            /// restart/continue, borders, shading, vertical alignment and block content; a null cell
            /// produces an empty placeholder.
            /// </summary>
            /// <param name="cell">单元格模型，null 表示占位/合并延续单元格。The cell model; null marks a placeholder or merge-continuation cell.</param>
            /// <param name="span">横向网格跨度（≥1）。The horizontal grid span (at least 1).</param>
            /// <param name="grid">表格网格列宽数组（磅）。The grid column widths in points.</param>
            /// <param name="column">起始网格列号（从 0 开始）。The zero-based starting grid column.</param>
            /// <param name="verticalMergeContinue">true 表示该单元格是纵向合并的延续部分。True when the cell continues a vertical merge.</param>
            /// <param name="context">当前部件写出上下文。The current part writing context.</param>
            /// <returns>w:tc 元素。The w:tc element.</returns>
            private XElement CreateCell(
                TableCell cell,
                int span,
                double[] grid,
                int column,
                bool verticalMergeContinue,
                PartContext context)
            {
                XNamespace w = WordNamespace;
                XElement element = new XElement(w + "tc");
                XElement properties = new XElement(w + "tcPr");

                double width = 0;
                for (int i = column; i < column + span && i < grid.Length; i++) width += grid[i];
                properties.Add(new XElement(w + "tcW",
                    new XAttribute(w + "w", ToTwips(width)),
                    new XAttribute(w + "type", "dxa")));
                if (span > 1)
                {
                    properties.Add(new XElement(w + "gridSpan", new XAttribute(w + "val", span.ToString(CultureInfo.InvariantCulture))));
                }
                if (verticalMergeContinue)
                {
                    properties.Add(new XElement(w + "vMerge", new XAttribute(w + "val", "continue")));
                }
                else if (cell != null && cell.RowSpan > 1)
                {
                    properties.Add(new XElement(w + "vMerge", new XAttribute(w + "val", "restart")));
                }

                if (cell != null)
                {
                    if (cell.Borders != null)
                    {
                        XElement borders = new XElement(w + "tcBorders");
                        bool any = false;
                        if (cell.Borders.Top != null) { borders.Add(CreateBorderElement("top", cell.Borders.Top)); any = true; }
                        if (cell.Borders.Left != null) { borders.Add(CreateBorderElement("left", cell.Borders.Left)); any = true; }
                        if (cell.Borders.Bottom != null) { borders.Add(CreateBorderElement("bottom", cell.Borders.Bottom)); any = true; }
                        if (cell.Borders.Right != null) { borders.Add(CreateBorderElement("right", cell.Borders.Right)); any = true; }
                        if (any) properties.Add(borders);
                    }
                    if (cell.BackgroundColor.HasValue)
                    {
                        properties.Add(new XElement(w + "shd",
                            new XAttribute(w + "val", "clear"),
                            new XAttribute(w + "color", "auto"),
                            new XAttribute(w + "fill", cell.BackgroundColor.Value.ToHex())));
                    }
                    switch (cell.VerticalAlignment)
                    {
                        case VerticalAlignment.Middle:
                            properties.Add(new XElement(w + "vAlign", new XAttribute(w + "val", "center")));
                            break;
                        case VerticalAlignment.Bottom:
                            properties.Add(new XElement(w + "vAlign", new XAttribute(w + "val", "bottom")));
                            break;
                    }
                }
                element.Add(properties);

                if (cell != null)
                {
                    WriteBlocks(element, cell.Blocks, context, Math.Max(18, width - 8));
                    if (cell.Blocks == null || cell.Blocks.Count == 0) element.Add(new XElement(w + "p"));
                }
                else
                {
                    element.Add(new XElement(w + "p"));
                }
                return element;
            }

            /// <summary>
            /// 生成表格边框集合 w:tblBorders：外框四边加内外分隔线，缺省线条写为 nil。
            /// Builds the w:tblBorders table border set: the four outer sides plus the inside horizontal and
            /// vertical rules, with missing lines written as nil.
            /// </summary>
            /// <param name="borders">表格边框定义，可为 null。The table border definition, may be null.</param>
            /// <returns>w:tblBorders 元素。The w:tblBorders element.</returns>
            private XElement CreateTableBorders(TableBorders borders)
            {
                XNamespace w = WordNamespace;
                XElement element = new XElement(w + "tblBorders");
                TableBorders source = borders ?? new TableBorders();
                element.Add(CreateBorderElement("top", source.Top));
                element.Add(CreateBorderElement("left", source.Left));
                element.Add(CreateBorderElement("bottom", source.Bottom));
                element.Add(CreateBorderElement("right", source.Right));
                element.Add(CreateBorderElement("insideH", source.InsideHorizontal));
                element.Add(CreateBorderElement("insideV", source.InsideVertical));
                return element;
            }

            /// <summary>
            /// 生成单条边框元素：样式映射为 OOXML 值（single/double/dashed/dotted），
            /// 宽度由磅换算为 1/8 磅单位；无边框写为 val="nil"。
            /// Builds one border element: the style is mapped to an OOXML value
            /// (single/double/dashed/dotted) and the width converted from points to eighths of a point; no
            /// border is written as val="nil".
            /// </summary>
            /// <param name="name">边框元素名（top/left/bottom/right/insideH/insideV）。The border element name (top/left/bottom/right/insideH/insideV).</param>
            /// <param name="border">边框定义，可为 null。The border definition, may be null.</param>
            /// <returns>对应名称的边框元素。The border element with the given name.</returns>
            private XElement CreateBorderElement(string name, BorderLine border)
            {
                XNamespace w = WordNamespace;
                XElement element = new XElement(w + name);
                if (border == null || border.Style == BorderStyle.None)
                {
                    element.Add(new XAttribute(w + "val", "nil"));
                    return element;
                }
                string value;
                switch (border.Style)
                {
                    case BorderStyle.Double: value = "double"; break;
                    case BorderStyle.Dashed: value = "dashed"; break;
                    case BorderStyle.Dotted: value = "dotted"; break;
                    default: value = "single"; break;
                }
                element.Add(new XAttribute(w + "val", value));
                element.Add(new XAttribute(w + "sz", ClampToInt(border.Width * 8, 2, 255).ToString(CultureInfo.InvariantCulture)));
                element.Add(new XAttribute(w + "space", "0"));
                element.Add(new XAttribute(w + "color", border.Color.ToHex()));
                return element;
            }

            /// <summary>
            /// 计算表格网格列宽（磅）：列数取声明列数与各行跨度和的最大值；
            /// 全部未指定时均分内容宽度，部分指定时未指定的列按已指定列的平均宽补齐。
            /// Computes the table grid column widths in points: the column count is the maximum of the
            /// declared columns and the summed spans of each row; when no width is set the content width is
            /// split equally, otherwise unset columns take the average of the set ones.
            /// </summary>
            /// <param name="table">表格模型。The table model.</param>
            /// <param name="contentWidth">可用内容宽度（磅）。The content width in points.</param>
            /// <returns>每列宽度数组。The per-column width array.</returns>
            private static double[] ComputeGrid(Table table, double contentWidth)
            {
                int count = table.Columns.Count;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    int span = 0;
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        span += Math.Max(1, table.Rows[r].Cells[c].ColSpan);
                    }
                    if (span > count) count = span;
                }
                if (count <= 0) count = 1;
                if (contentWidth < 36) contentWidth = 468;

                double[] grid = new double[count];
                double specified = 0;
                int specifiedCount = 0;
                for (int i = 0; i < count; i++)
                {
                    if (i < table.Columns.Count && table.Columns[i].Width > 0)
                    {
                        grid[i] = table.Columns[i].Width;
                        specified += grid[i];
                        specifiedCount++;
                    }
                }
                if (specifiedCount == 0)
                {
                    double equal = contentWidth / count;
                    for (int i = 0; i < count; i++) grid[i] = equal;
                    return grid;
                }

                double average = specified / specifiedCount;
                for (int i = 0; i < count; i++)
                {
                    if (grid[i] <= 0) grid[i] = average;
                }
                return grid;
            }

            /// <summary>
            /// 生成 styles.xml：docDefaults 默认字体（含中西文）与字号、默认 Normal 段落样式，
            /// 以及 Heading1-9 标题样式（保持下段同页、大纲级别、加粗、主题色与递减字号）。
            /// Builds styles.xml: the docDefaults default fonts (East-Asian and Latin) and size, the default
            /// Normal paragraph style, and the Heading1-9 styles (keep-with-next, outline level, bold, theme
            /// colour and decreasing sizes).
            /// </summary>
            /// <returns>styles.xml 文档。The styles.xml document.</returns>
            private XDocument CreateStylesXml()
            {
                XNamespace w = WordNamespace;
                TextFormat defaults = _document.DefaultTextFormat ?? new TextFormat();
                string font = string.IsNullOrEmpty(defaults.FontFamily) ? "Calibri" : defaults.FontFamily;
                string asciiFont = defaults.GetAsciiFontFamily();
                if (string.IsNullOrEmpty(asciiFont)) asciiFont = font;
                string eastAsiaFont = defaults.GetEastAsiaFontFamily();
                if (string.IsNullOrEmpty(eastAsiaFont)) eastAsiaFont = font;
                double size = defaults.FontSize > 0 ? defaults.FontSize : 11;

                XElement runDefaults = new XElement(w + "rPrDefault",
                    new XElement(w + "rPr",
                        new XElement(w + "rFonts",
                            new XAttribute(w + "ascii", asciiFont),
                            new XAttribute(w + "hAnsi", asciiFont),
                            new XAttribute(w + "eastAsia", eastAsiaFont),
                            new XAttribute(w + "cs", asciiFont)),
                        // 运行级颜色在与文档默认色一致时会被省略，所以默认色必须在这里落地，否则非黑默认色
                        // 的文档省略后按 Word 的 auto（黑）渲染。黑色等同 auto，无需写出。
                        // A run-level colour equal to the document default is now omitted, so the default has to
                        // land here — otherwise a document whose default colour is not black would render black
                        // after the omission. Black equals Word's auto, so it needs no element.
                        !string.Equals(defaults.Color.ToHex(), OfficeColor.Black.ToHex(), StringComparison.OrdinalIgnoreCase)
                            ? new XElement(w + "color", new XAttribute(w + "val", defaults.Color.ToHex()))
                            : null,
                        new XElement(w + "sz", new XAttribute(w + "val", ClampToInt(size * 2, 2, 32767).ToString(CultureInfo.InvariantCulture)))));

                XElement root = new XElement(w + "styles",
                    new XAttribute(XNamespace.Xmlns + "w", w),
                    new XElement(w + "docDefaults", runDefaults),
                    new XElement(w + "style",
                        new XAttribute(w + "type", "paragraph"),
                        new XAttribute(w + "default", "1"),
                        new XAttribute(w + "styleId", "Normal"),
                        new XElement(w + "name", new XAttribute(w + "val", "Normal")),
                        new XElement(w + "qFormat")));

                double[] headingSizes = { 20, 15, 13, 12, 11, 10, 10, 10, 10 };
                for (int level = 1; level <= 9; level++)
                {
                    root.Add(new XElement(w + "style",
                        new XAttribute(w + "type", "paragraph"),
                        new XAttribute(w + "styleId", "Heading" + level.ToString(CultureInfo.InvariantCulture)),
                        new XElement(w + "name", new XAttribute(w + "val", "heading " + level.ToString(CultureInfo.InvariantCulture))),
                        new XElement(w + "basedOn", new XAttribute(w + "val", "Normal")),
                        new XElement(w + "next", new XAttribute(w + "val", "Normal")),
                        new XElement(w + "qFormat"),
                        new XElement(w + "pPr",
                            new XElement(w + "keepNext"),
                            new XElement(w + "keepLines"),
                            new XElement(w + "spacing", new XAttribute(w + "before", "240"), new XAttribute(w + "after", "60")),
                            new XElement(w + "outlineLvl", new XAttribute(w + "val", (level - 1).ToString(CultureInfo.InvariantCulture)))),
                        new XElement(w + "rPr",
                            new XElement(w + "b"),
                            new XElement(w + "color", new XAttribute(w + "val", "1F3A5F")),
                            new XElement(w + "sz", new XAttribute(w + "val", ClampToInt(headingSizes[level - 1] * 2, 2, 32767).ToString(CultureInfo.InvariantCulture))),
                            new XElement(w + "szCs", new XAttribute(w + "val", ClampToInt(headingSizes[level - 1] * 2, 2, 32767).ToString(CultureInfo.InvariantCulture))))));
                }
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 生成 settings.xml：默认制表位、打开时更新域、不压缩字符间距，并声明兼容模式为 15（Word 2013+）。
            /// Builds settings.xml: the default tab stop, update-fields-on-open, no character-spacing
            /// compression and compatibility mode 15 (Word 2013+).
            /// </summary>
            /// <returns>settings.xml 文档。The settings.xml document.</returns>
            private XDocument CreateSettingsXml()
            {
                XNamespace w = WordNamespace;
                XElement root = new XElement(w + "settings",
                    new XAttribute(XNamespace.Xmlns + "w", w),
                    new XElement(w + "defaultTabStop", new XAttribute(w + "val", "420")),
                    new XElement(w + "updateFields", new XAttribute(w + "val", "true")),
                    new XElement(w + "characterSpacingControl", new XAttribute(w + "val", "doNotCompress")),
                    new XElement(w + "compat",
                        new XElement(w + "compatSetting",
                            new XAttribute(w + "name", "compatibilityMode"),
                            new XAttribute(w + "uri", "http://schemas.microsoft.com/office/word"),
                            new XAttribute(w + "val", "15"))));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 生成 [Content_Types].xml：rels/xml/媒体扩展名的 Default 条目，
            /// 以及主文档、样式、设置、页眉页脚与属性部件的 Override 条目。
            /// Builds [Content_Types].xml: Default entries for the rels, xml and media extensions plus
            /// Override entries for the main document, styles, settings, header/footer and property parts.
            /// </summary>
            /// <returns>[Content_Types].xml 文档。The [Content_Types].xml document.</returns>
            private XDocument CreateContentTypesXml()
            {
                XNamespace ct = ContentTypesNamespace;
                XElement root = new XElement(ct + "Types",
                    new XAttribute(XNamespace.Xmlns + "ct", ct));
                HashSet<string> defaults = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                AddDefault(root, defaults, "rels", ContentTypeRelationships);
                AddDefault(root, defaults, "xml", ContentTypeXml);
                foreach (KeyValuePair<string, string> pair in MediaContentTypes)
                {
                    AddDefault(root, defaults, pair.Key.TrimStart('.'), pair.Value);
                }
                AddOverride(root, "/word/document.xml", _mainContentType);
                AddOverride(root, "/word/styles.xml", ContentTypeStyles);
                AddOverride(root, "/word/settings.xml", ContentTypeSettings);
                if (_numberingDefinitions.Count > 0)
                {
                    AddOverride(root, "/word/numbering.xml", ContentTypeNumbering);
                }
                for (int i = 0; i < _sections.Count; i++)
                {
                    SectionParts parts = _sections[i];
                    if (parts.Header != null) AddOverride(root, "/" + parts.Header.PartPath, ContentTypeHeader);
                    if (parts.Footer != null) AddOverride(root, "/" + parts.Footer.PartPath, ContentTypeFooter);
                }
                AddOverride(root, "/docProps/core.xml", ContentTypeCoreProperties);
                AddOverride(root, "/docProps/app.xml", ContentTypeExtendedProperties);
                if (HasCustomProperties())
                {
                    AddOverride(root, "/docProps/custom.xml", ContentTypeCustomProperties);
                }
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 追加按扩展名的 Default 内容类型条目；扩展名为空或已添加过则跳过（去重）。
            /// Appends a Default content-type entry by extension; empty or already-added extensions are
            /// skipped for de-duplication.
            /// </summary>
            /// <param name="root">Types 根元素。The Types root element.</param>
            /// <param name="added">已添加扩展名集合。The set of extensions already added.</param>
            /// <param name="extension">文件扩展名（不含点）。The file extension without the dot.</param>
            /// <param name="contentType">对应内容类型。The matching content type.</param>
            private static void AddDefault(XElement root, HashSet<string> added, string extension, string contentType)
            {
                if (string.IsNullOrEmpty(extension) || !added.Add(extension)) return;
                XNamespace ct = ContentTypesNamespace;
                root.Add(new XElement(ct + "Default",
                    new XAttribute("Extension", extension),
                    new XAttribute("ContentType", contentType)));
            }

            /// <summary>
            /// 追加按部件路径的 Override 内容类型条目。
            /// Appends an Override content-type entry addressed by part path.
            /// </summary>
            /// <param name="root">Types 根元素。The Types root element.</param>
            /// <param name="partName">以 / 开头的部件路径。The part path starting with /.</param>
            /// <param name="contentType">对应内容类型。The matching content type.</param>
            private static void AddOverride(XElement root, string partName, string contentType)
            {
                XNamespace ct = ContentTypesNamespace;
                root.Add(new XElement(ct + "Override",
                    new XAttribute("PartName", partName),
                    new XAttribute("ContentType", contentType)));
            }

            /// <summary>
            /// 生成包级根关系 _rels/.rels：固定三条关系指向主文档、核心属性与扩展属性部件。
            /// Builds the package-level root relationships _rels/.rels: three fixed relationships pointing at
            /// the main document, core-properties and extended-properties parts.
            /// </summary>
            /// <returns>根关系文档。The root relationships document.</returns>
            private XDocument CreateRootRelationshipsXml()
            {
                List<RelationshipEntry> relationships = new List<RelationshipEntry>();
                relationships.Add(new RelationshipEntry { Id = "rId1", Type = DocumentRelationshipType, Target = "word/document.xml" });
                relationships.Add(new RelationshipEntry { Id = "rId2", Type = CorePropertiesRelationshipType, Target = "docProps/core.xml" });
                relationships.Add(new RelationshipEntry { Id = "rId3", Type = ExtendedPropertiesRelationshipType, Target = "docProps/app.xml" });
                if (HasCustomProperties())
                {
                    relationships.Add(new RelationshipEntry { Id = "rId4", Type = CustomPropertiesRelationshipType, Target = "docProps/custom.xml" });
                }
                return CreateRelationshipsXml(relationships);
            }

            /// <summary>
            /// 将关系条目列表序列化为 Relationships 根元素的 XML 文档。
            /// Serialises a list of relationship entries into an XML document rooted at Relationships.
            /// </summary>
            /// <param name="relationships">关系条目列表。The list of relationship entries.</param>
            /// <returns>.rels 部件文档。The .rels part document.</returns>
            private XDocument CreateRelationshipsXml(List<RelationshipEntry> relationships)
            {
                XNamespace rel = PackageRelationshipNamespace;
                XElement root = new XElement(rel + "Relationships",
                    new XAttribute(XNamespace.Xmlns + "rel", rel));
                for (int i = 0; i < relationships.Count; i++)
                {
                    RelationshipEntry entry = relationships[i];
                    XElement relationship = new XElement(rel + "Relationship",
                        new XAttribute("Id", entry.Id),
                        new XAttribute("Type", entry.Type),
                        new XAttribute("Target", entry.Target));
                    if (entry.External) relationship.Add(new XAttribute("TargetMode", "External"));
                    root.Add(relationship);
                }
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 生成 docProps/core.xml 核心属性部件：标题、作者、最后修改者、主题、关键词与 W3CDTF 格式的创建/修改时间。
            /// Builds the docProps/core.xml core-properties part: title, author, last modified by, subject,
            /// keywords and the W3CDTF created/modified timestamps.
            /// </summary>
            /// <returns>core.xml 文档。The core.xml document.</returns>
            private XDocument CreateCorePropertiesXml()
            {
                XNamespace cp = CorePropertiesNamespace;
                XNamespace dc = DcNamespace;
                XNamespace dcterms = DcTermsNamespace;
                XNamespace xsi = XsiNamespace;
                DocumentMetadata metadata = _document.Metadata ?? new DocumentMetadata();
                XElement root = new XElement(cp + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp", cp),
                    new XAttribute(XNamespace.Xmlns + "dc", dc),
                    new XAttribute(XNamespace.Xmlns + "dcterms", dcterms),
                    new XAttribute(XNamespace.Xmlns + "xsi", xsi));
                root.Add(new XElement(dc + "title", metadata.Title ?? string.Empty));
                root.Add(new XElement(dc + "creator", metadata.Author ?? string.Empty));
                root.Add(new XElement(cp + "lastModifiedBy", metadata.Author ?? string.Empty));
                if (!string.IsNullOrEmpty(metadata.Subject)) root.Add(new XElement(dc + "subject", metadata.Subject));
                if (!string.IsNullOrEmpty(metadata.Keywords)) root.Add(new XElement(cp + "keywords", metadata.Keywords));
                root.Add(new XElement(dcterms + "created",
                    new XAttribute(xsi + "type", "dcterms:W3CDTF"),
                    FormatTimestamp(metadata.Created)));
                root.Add(new XElement(dcterms + "modified",
                    new XAttribute(xsi + "type", "dcterms:W3CDTF"),
                    FormatTimestamp(metadata.Modified)));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 生成 docProps/app.xml 扩展属性部件：应用程序名固定为 Wisdom.Office，公司名取元数据创建者。
            /// Builds the docProps/app.xml extended-properties part: the application name is fixed to
            /// Wisdom.Office and the company name comes from the metadata creator.
            /// </summary>
            /// <returns>app.xml 文档。The app.xml document.</returns>
            private XDocument CreateAppPropertiesXml()
            {
                XNamespace ep = ExtendedPropertiesNamespace;
                XNamespace vt = DocPropsVectorNamespace;
                DocumentMetadata metadata = _document.Metadata ?? new DocumentMetadata();
                XElement root = new XElement(ep + "Properties",
                    new XAttribute(XNamespace.Xmlns + "vt", vt),
                    new XElement(ep + "Application", "Wisdom.Office"),
                    new XElement(ep + "DocSecurity", "0"),
                    new XElement(ep + "ScaleCrop", "false"),
                    new XElement(ep + "Company", metadata.Creator ?? string.Empty),
                    new XElement(ep + "LinksUpToDate", "false"),
                    new XElement(ep + "SharedDoc", "false"),
                    new XElement(ep + "HyperlinksChanged", "false"),
                    new XElement(ep + "AppVersion", "16.0000"));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 判断文档是否含有需要写出的自定义属性。
            /// Determines whether the document has any custom properties to write.
            /// </summary>
            /// <returns>存在自定义属性时返回 true。True when custom properties are present.</returns>
            private bool HasCustomProperties()
            {
                return _document != null
                    && _document.Metadata != null
                    && _document.Metadata.CustomProperties != null
                    && _document.Metadata.CustomProperties.Count > 0;
            }

            /// <summary>
            /// 生成 docProps/custom.xml 自定义属性部件：每个属性写入 fmtid/pid/name 与对应的 vt 类型值。
            /// Builds the docProps/custom.xml custom-properties part: each property carries fmtid/pid/name and
            /// its matching vt-typed value.
            /// </summary>
            /// <returns>custom.xml 文档。The custom.xml document.</returns>
            private XDocument CreateCustomPropertiesXml()
            {
                XNamespace cp = CustomPropertiesNamespace;
                XNamespace vt = DocPropsVectorNamespace;
                XElement root = new XElement(cp + "Properties",
                    new XAttribute(XNamespace.Xmlns + "vt", vt));
                int pid = 2;
                foreach (KeyValuePair<string, object> pair in _document.Metadata.CustomProperties)
                {
                    root.Add(new XElement(cp + "property",
                        new XAttribute("fmtid", CustomPropertiesFormatId),
                        new XAttribute("pid", pid.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("name", pair.Key),
                        CreateCustomPropertyValue(pair.Value, vt)));
                    pid++;
                }
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 按 .NET 值类型生成对应的 vt 元素：string→lpwstr、bool→bool、int→i4、long→i8、
            /// double→r8、decimal→decimal、DateTime→filetime（UTC）。
            /// Maps a .NET value onto its vt element: string→lpwstr, bool→bool, int→i4, long→i8,
            /// double→r8, decimal→decimal and DateTime→filetime (UTC).
            /// </summary>
            /// <param name="value">属性值。The property value.</param>
            /// <param name="vt">docPropsVTypes 命名空间。The docPropsVTypes namespace.</param>
            /// <returns>vt 值元素。The vt value element.</returns>
            private static XElement CreateCustomPropertyValue(object value, XNamespace vt)
            {
                if (value is bool)
                {
                    return new XElement(vt + "bool", (bool)value ? "true" : "false");
                }
                if (value is int)
                {
                    return new XElement(vt + "i4", ((int)value).ToString(CultureInfo.InvariantCulture));
                }
                if (value is long)
                {
                    return new XElement(vt + "i8", ((long)value).ToString(CultureInfo.InvariantCulture));
                }
                if (value is double)
                {
                    return new XElement(vt + "r8", ((double)value).ToString("R", CultureInfo.InvariantCulture));
                }
                if (value is decimal)
                {
                    return new XElement(vt + "decimal", ((decimal)value).ToString(CultureInfo.InvariantCulture));
                }
                if (value is DateTime)
                {
                    DateTime timestamp = (DateTime)value;
                    DateTime utc = timestamp.Kind == DateTimeKind.Utc
                        ? timestamp
                        : timestamp.Kind == DateTimeKind.Unspecified
                            ? DateTime.SpecifyKind(timestamp, DateTimeKind.Utc)
                            : timestamp.ToUniversalTime();
                    string fileTime = utc.ToString("yyyy-MM-dd'T'HH:mm:ss", CultureInfo.InvariantCulture);
                    long fraction = utc.Ticks % TimeSpan.TicksPerSecond;
                    if (fraction != 0)
                    {
                        fileTime += "." + fraction.ToString("D7", CultureInfo.InvariantCulture).TrimEnd('0');
                    }
                    return new XElement(vt + "filetime", fileTime + "Z");
                }
                return new XElement(vt + "lpwstr", Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
            }

            /// <summary>
            /// 将时间统一转换为 UTC 并按 W3CDTF（yyyy-MM-ddTHH:mm:ssZ）格式输出。
            /// Converts the value to UTC and formats it as W3CDTF (yyyy-MM-ddTHH:mm:ssZ).
            /// </summary>
            /// <param name="value">本地或 UTC 时间。A local or UTC time.</param>
            /// <returns>ISO 8601 格式的 UTC 时间字符串。The UTC timestamp in ISO 8601 form.</returns>
            private static string FormatTimestamp(DateTime value)
            {
                DateTime utc = value.Kind == DateTimeKind.Utc ? value : value.ToUniversalTime();
                return utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
            }

            /// <summary>
            /// 把任意浮点量（可能来自被污染的模型值或 HTML 内联样式）取整并夹到指定区间，
            /// NaN / 无穷 / 越界一律落到边界，避免 (int) 转换回绕成负数后写出非法属性。
            /// Rounds an arbitrary double — which may come from a poisoned model value or an HTML inline style —
            /// and clamps it into the range. NaN, infinity and out-of-range values land on a boundary instead of
            /// the (int) cast wrapping around and emitting a negative attribute Word rejects.
            /// </summary>
            /// <param name="value">待取整的值。The value to round.</param>
            /// <param name="minimum">区间下界。The lower bound.</param>
            /// <param name="maximum">区间上界。The upper bound.</param>
            /// <returns>区间内的整数。The integer inside the range.</returns>
            private static int ClampToInt(double value, int minimum, int maximum)
            {
                if (double.IsNaN(value) || double.IsInfinity(value)) return minimum;
                double rounded = Math.Round(value);
                if (rounded <= minimum) return minimum;
                if (rounded >= maximum) return maximum;
                return (int)rounded;
            }

            private static string ToTwips(double points)
            {
                return ClampToInt(points * UnitConverter.TwipsPerPoint, -2147483000, 2147483000)
                    .ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}
