using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// DOCX 读取器：将 WordprocessingML（.docx）压缩包解析为统一文档模型。
    /// DOCX reader that parses a WordprocessingML (.docx) package into the unified document model.
    /// </summary>
    public static class DocxReader
    {
        /// <summary>分节解析期间的临时草稿：保存该节的块内容、节属性以及页眉页脚部件路径。Working draft used while splitting the body into sections: holds the section's blocks, its section properties and the header/footer part paths.</summary>
        private sealed class SectionDraft
        {
            /// <summary>本节已解析的块级内容列表。The blocks parsed so far for this section.</summary>
            public List<Block> Blocks = new List<Block>();
            /// <summary>本节的 w:sectPr 节属性元素，可为 null。The section's w:sectPr element, may be null.</summary>
            public XElement Properties;
            /// <summary>本节引用的页眉部件路径，可为 null（表示沿用上一节）。The header part referenced by this section, or null to inherit the previous section's.</summary>
            public string HeaderPart;
            /// <summary>本节引用的页脚部件路径，可为 null（表示沿用上一节）。The footer part referenced by this section, or null to inherit the previous section's.</summary>
            public string FooterPart;
        }

        /// <summary>
        /// 从磁盘路径加载 .docx 文档。
        /// Loads a .docx document from a disk path.
        /// </summary>
        /// <param name="path">.docx 文件路径。The path of the .docx file.</param>
        /// <returns>解析得到的统一文档模型。The parsed unified document model.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> 为 null。Thrown when <paramref name="path"/> is null.</exception>
        /// <exception cref="System.IO.InvalidDataException">文件不是有效的 .docx 包（或为旧版 .doc）。Thrown when the file is not a valid .docx package (or is a legacy .doc).</exception>
        public static OfficeDocument Load(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(File.ReadAllBytes(path));
        }

        /// <summary>
        /// 从可读流加载 .docx 文档（流会先完整复制到内存再解析）。
        /// Loads a .docx document from a readable stream (the stream is copied to memory before parsing).
        /// </summary>
        /// <param name="stream">包含 .docx 字节内容的流。A stream containing the .docx bytes.</param>
        /// <returns>解析得到的统一文档模型。The parsed unified document model.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> 为 null。Thrown when <paramref name="stream"/> is null.</exception>
        public static OfficeDocument Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return Load(OfficePackageLimits.ReadAllBytes(stream));
        }

        /// <summary>
        /// 从字节数组加载 .docx 文档：先校验文件头（拒绝 OLE 复合文档即旧版 .doc，要求 ZIP 签名），
        /// 再解压并解析整个包。
        /// Loads a .docx document from a byte array: validates the header first (OLE compound files, i.e.
        /// legacy .doc, are rejected and a ZIP signature is required), then decompresses and parses the package.
        /// </summary>
        /// <param name="data">.docx 文件的完整字节内容。The full bytes of the .docx file.</param>
        /// <returns>解析得到的统一文档模型。The parsed unified document model.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null。Thrown when <paramref name="data"/> is null.</exception>
        /// <exception cref="InvalidDataException">数据不是有效的 .docx 包，或包已损坏/截断。Thrown when the data is not a valid .docx package or the package is corrupt or truncated.</exception>
        public static OfficeDocument Load(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length >= 8
                && data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0
                && data[4] == 0xA1 && data[5] == 0xB1 && data[6] == 0x1A && data[7] == 0xE1)
            {
                throw new InvalidDataException("The file is a legacy .doc document. Use DocReader instead.");
            }
            if (data.Length < 4 || data[0] != 0x50 || data[1] != 0x4B)
            {
                throw new InvalidDataException("The file is not a valid .docx package.");
            }

            try
            {
                using (MemoryStream memory = new MemoryStream(data, false))
                using (ZipArchive archive = PackageArchive.OpenForRead(memory, false))
                {
                    DocxPackage package = new DocxPackage(archive);
                    return ReadPackage(package);
                }
            }
            catch (IOException ex)
            {
                throw new InvalidDataException("The .docx package is corrupt or truncated.", ex);
            }
        }

        /// <summary>
        /// 快速提取指定 .docx 文件的纯文本内容（等效于 Load 后调用 GetPlainText）。
        /// Extracts the plain text of a .docx file (equivalent to calling GetPlainText after Load).
        /// </summary>
        /// <param name="path">.docx 文件路径。The path of the .docx file.</param>
        /// <returns>文档的纯文本。The plain text of the document.</returns>
        public static string ExtractText(string path)
        {
            return Load(path).GetPlainText();
        }

        /// <summary>
        /// 解析整个 DOCX 包：定位主文档/样式/编号部件，按 w:sectPr 将正文切分为多个节草稿，
        /// 随后顺序构建各节并处理页眉页脚与页面设置的向后继承（未显式指定时沿用上一节的值）。
        /// Parses the whole DOCX package: locates the main document, styles and numbering parts, splits the body
        /// into section drafts at each w:sectPr, then builds the sections in order while inheriting headers,
        /// footers and page setup from the previous section when they are not explicitly set.
        /// </summary>
        /// <param name="package">已打开的 DOCX 包。The opened DOCX package.</param>
        /// <returns>组装完成的统一文档模型。The assembled unified document model.</returns>
        /// <exception cref="InvalidDataException">包中缺少主文档部件或正文元素。Thrown when the package has no main document part or no body element.</exception>
        private static OfficeDocument ReadPackage(DocxPackage package)
        {
            string documentPart = package.GetRelationshipTargetByType(string.Empty, "/officeDocument");
            if (string.IsNullOrEmpty(documentPart)) documentPart = "word/document.xml";

            XDocument documentXml = package.GetXml(documentPart);
            if (documentXml == null || documentXml.Root == null)
            {
                throw new InvalidDataException("The .docx package does not contain a main document part.");
            }

            string stylesPart = package.GetRelationshipTargetByType(documentPart, "/styles");
            if (stylesPart == null) stylesPart = DocxPackage.ResolvePartPath(documentPart, "styles.xml");
            string numberingPart = package.GetRelationshipTargetByType(documentPart, "/numbering");
            if (numberingPart == null) numberingPart = DocxPackage.ResolvePartPath(documentPart, "numbering.xml");

            DocxStyles styles = new DocxStyles(package.GetXml(stylesPart), package.GetXml(numberingPart));

            OfficeDocument document = new OfficeDocument();
            document.DefaultTextFormat = styles.DefaultTextFormat.Clone();
            ReadMetadata(package, document.Metadata);

            XElement body = Wml.Child(documentXml.Root, "body");
            if (body == null) throw new InvalidDataException("The Word document body is missing.");

            List<SectionDraft> drafts = new List<SectionDraft>();
            SectionDraft current = new SectionDraft();
            XElement bodyProperties = null;
            DocxPartReader reader = new DocxPartReader(package, styles, documentPart);

            foreach (XElement child in body.Elements())
            {
                if (child.Name == Wml.W + "p")
                {
                    Paragraph paragraph = reader.ReadParagraph(child);
                    AppendParagraphBlocks(current.Blocks, paragraph);
                    XElement paragraphProperties = Wml.Child(child, "pPr");
                    XElement sectionProperties = Wml.Child(paragraphProperties, "sectPr");
                    if (sectionProperties != null)
                    {
                        current.Properties = sectionProperties;
                        ReadHeaderFooterReferences(package, documentPart, sectionProperties, current);
                        drafts.Add(current);
                        current = new SectionDraft();
                    }
                    continue;
                }

                if (child.Name == Wml.W + "tbl")
                {
                    Table table = reader.ReadTable(child);
                    if (table != null) current.Blocks.Add(table);
                    continue;
                }

                if (child.Name == Wml.W + "sectPr")
                {
                    bodyProperties = child;
                    continue;
                }

                if (child.Name == Wml.W + "sdt")
                {
                    reader.ReadBlocks(Wml.Child(child, "sdtContent"), current.Blocks);
                    continue;
                }

                if (child.Name == Wml.W + "customXml")
                {
                    reader.ReadBlocks(child, current.Blocks);
                }
            }

            current.Properties = bodyProperties;
            if (bodyProperties != null)
            {
                ReadHeaderFooterReferences(package, documentPart, bodyProperties, current);
            }
            drafts.Add(current);

            string inheritedHeader = null;
            string inheritedFooter = null;
            PageSetup inheritedSetup = null;
            for (int i = 0; i < drafts.Count; i++)
            {
                SectionDraft draft = drafts[i];
                if (draft.HeaderPart != null) inheritedHeader = draft.HeaderPart;
                if (draft.FooterPart != null) inheritedFooter = draft.FooterPart;

                PageSetup setup = inheritedSetup == null ? new PageSetup() : inheritedSetup.Clone();
                ReadPageSetup(draft.Properties, setup);
                inheritedSetup = setup;

                Section section = document.AddSection(setup);
                section.Blocks.AddRange(draft.Blocks);
                if (inheritedHeader != null)
                {
                    section.Header.Blocks.AddRange(ReadHeaderFooterBlocks(package, styles, inheritedHeader));
                }
                if (inheritedFooter != null)
                {
                    section.Footer.Blocks.AddRange(ReadHeaderFooterBlocks(package, styles, inheritedFooter));
                }
            }

            return document;
        }

        /// <summary>
        /// 从节属性中读取 w:headerReference / w:footerReference 关系引用，
        /// 解析出实际引用的页眉、页脚部件路径（优先 default 类型，其次首个引用）。
        /// Reads the w:headerReference / w:footerReference relationship entries of a section and resolves the
        /// header and footer part paths they point at (the default type wins, otherwise the first reference).
        /// </summary>
        /// <param name="package">DOCX 包，用于查询关系目标。The DOCX package used to look up relationship targets.</param>
        /// <param name="documentPart">当前部件路径（关系解析基准）。The current part path used as the relationship base.</param>
        /// <param name="sectionProperties">节属性元素 w:sectPr。The w:sectPr section-properties element.</param>
        /// <param name="draft">接收页眉页脚路径的节草稿。The section draft that receives the header/footer paths.</param>
        private static void ReadHeaderFooterReferences(
            DocxPackage package,
            string documentPart,
            XElement sectionProperties,
            SectionDraft draft)
        {
            string defaultHeader = null;
            string firstHeader = null;
            foreach (XElement reference in Wml.Children(sectionProperties, "headerReference"))
            {
                string type = Wml.AttributeValue(reference, "type") ?? "default";
                string relationshipId = Wml.AttributeValue(reference, Wml.R + "id");
                string target = package.GetRelationshipTarget(documentPart, relationshipId);
                if (target == null) continue;
                if (type == "default" && defaultHeader == null) defaultHeader = target;
                else if (firstHeader == null) firstHeader = target;
            }
            if (defaultHeader != null || firstHeader != null)
            {
                draft.HeaderPart = defaultHeader ?? firstHeader;
            }

            string defaultFooter = null;
            string firstFooter = null;
            foreach (XElement reference in Wml.Children(sectionProperties, "footerReference"))
            {
                string type = Wml.AttributeValue(reference, "type") ?? "default";
                string relationshipId = Wml.AttributeValue(reference, Wml.R + "id");
                string target = package.GetRelationshipTarget(documentPart, relationshipId);
                if (target == null) continue;
                if (type == "default" && defaultFooter == null) defaultFooter = target;
                else if (firstFooter == null) firstFooter = target;
            }
            if (defaultFooter != null || firstFooter != null)
            {
                draft.FooterPart = defaultFooter ?? firstFooter;
            }
        }

        /// <summary>
        /// 读取页眉/页脚部件的全部块内容；部件缺失时返回空列表。
        /// Reads all blocks of a header or footer part and returns an empty list when the part is missing.
        /// </summary>
        /// <param name="package">DOCX 包。The DOCX package.</param>
        /// <param name="styles">文档样式表。The document style table.</param>
        /// <param name="partPath">页眉或页脚部件路径。The header or footer part path.</param>
        /// <returns>部件解析出的块列表。The blocks parsed from the part.</returns>
        private static List<Block> ReadHeaderFooterBlocks(DocxPackage package, DocxStyles styles, string partPath)
        {
            List<Block> blocks = new List<Block>();
            XDocument xml = package.GetXml(partPath);
            if (xml == null || xml.Root == null) return blocks;
            DocxPartReader reader = new DocxPartReader(package, styles, partPath);
            reader.ReadBlocks(xml.Root, blocks);
            return blocks;
        }

        /// <summary>
        /// 从 w:sectPr 读取页面尺寸/方向与页边距（缇换算为磅）并写入 <see cref="PageSetup"/>；
        /// 方向与宽高不一致时自动交换宽高。
        /// Reads the page size, orientation and margins from w:sectPr (twips converted to points) into a
        /// <see cref="PageSetup"/>; width and height are swapped when they disagree with the orientation.
        /// </summary>
        /// <param name="sectionProperties">节属性元素，可为 null。The section-properties element, may be null.</param>
        /// <param name="setup">接收页面设置的目标对象。The target object that receives the page setup.</param>
        private static void ReadPageSetup(XElement sectionProperties, PageSetup setup)
        {
            if (sectionProperties == null) return;
            XElement size = Wml.Child(sectionProperties, "pgSz");
            if (size != null)
            {
                double width = Wml.TwipsToPoints(Wml.AttributeValue(size, "w"), 0);
                double height = Wml.TwipsToPoints(Wml.AttributeValue(size, "h"), 0);
                if (width > 0) setup.Width = width;
                if (height > 0) setup.Height = height;
                string orientation = Wml.AttributeValue(size, "orient");
                if (orientation == "landscape" && setup.Width < setup.Height
                    || orientation == "portrait" && setup.Width > setup.Height)
                {
                    double swap = setup.Width;
                    setup.Width = setup.Height;
                    setup.Height = swap;
                }
            }

            XElement margins = Wml.Child(sectionProperties, "pgMar");
            if (margins != null)
            {
                setup.MarginTop = Wml.TwipsToPoints(Wml.AttributeValue(margins, "top"), setup.MarginTop);
                setup.MarginBottom = Wml.TwipsToPoints(Wml.AttributeValue(margins, "bottom"), setup.MarginBottom);
                setup.MarginLeft = Wml.TwipsToPoints(Wml.AttributeValue(margins, "left"), setup.MarginLeft);
                setup.MarginRight = Wml.TwipsToPoints(Wml.AttributeValue(margins, "right"), setup.MarginRight);
                setup.HeaderDistance = Wml.TwipsToPoints(Wml.AttributeValue(margins, "header"), setup.HeaderDistance);
                setup.FooterDistance = Wml.TwipsToPoints(Wml.AttributeValue(margins, "footer"), setup.FooterDistance);
            }
        }

        /// <summary>
        /// 从 docProps/core.xml 核心属性部件读取标题、作者、主题、关键词与创建/修改时间等元数据。
        /// Reads metadata (title, author, subject, keywords and created/modified timestamps) from the
        /// docProps/core.xml core-properties part.
        /// </summary>
        /// <param name="package">DOCX 包。The DOCX package.</param>
        /// <param name="metadata">接收元数据的目标对象。The target object that receives the metadata.</param>
        private static void ReadMetadata(DocxPackage package, DocumentMetadata metadata)
        {
            XDocument core = package.GetXml("docProps/core.xml");
            if (core != null && core.Root != null)
            {
                XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
                XNamespace dc = "http://purl.org/dc/elements/1.1/";
                XNamespace dcterms = "http://purl.org/dc/terms/";

                XElement title = core.Root.Element(dc + "title");
                if (title != null) metadata.Title = title.Value;
                XElement creator = core.Root.Element(dc + "creator");
                if (creator != null) metadata.Author = creator.Value;
                XElement subject = core.Root.Element(dc + "subject");
                if (subject != null) metadata.Subject = subject.Value;
                XElement keywords = core.Root.Element(cp + "keywords");
                if (keywords != null) metadata.Keywords = keywords.Value;
                XElement lastModifiedBy = core.Root.Element(cp + "lastModifiedBy");
                if (lastModifiedBy != null && string.IsNullOrEmpty(metadata.Author)) metadata.Author = lastModifiedBy.Value;

                DateTime created;
                if (TryParseTimestamp(core.Root.Element(dcterms + "created"), out created)) metadata.Created = created;
                DateTime modified;
                if (TryParseTimestamp(core.Root.Element(dcterms + "modified"), out modified)) metadata.Modified = modified;
            }

            ReadCustomProperties(package, metadata);
        }

        /// <summary>
        /// 从 docProps/custom.xml 自定义属性部件读取键值：先按包根关系定位部件，缺失时回退到固定路径；
        /// 支持文本、布尔、整数、浮点、十进制与 filetime 日期，无法识别的 vt 类型按文本保留。
        /// Reads custom properties from the docProps/custom.xml part: locates the part through the package
        /// relationship and falls back to the fixed path; supports text, Boolean, integer, floating-point,
        /// decimal and filetime values, while unrecognised vt types are preserved as text.
        /// </summary>
        /// <param name="package">DOCX 包。The DOCX package.</param>
        /// <param name="metadata">接收自定义属性的元数据对象。The metadata object receiving the custom properties.</param>
        private static void ReadCustomProperties(DocxPackage package, DocumentMetadata metadata)
        {
            string partPath = package.GetRelationshipTargetByType(string.Empty, "/custom-properties");
            if (string.IsNullOrEmpty(partPath)) partPath = "docProps/custom.xml";
            XDocument custom = package.GetXml(partPath);
            if (custom == null || custom.Root == null) return;

            XNamespace propertiesNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties";
            foreach (XElement property in custom.Root.Elements(propertiesNamespace + "property"))
            {
                string name = (string)property.Attribute("name");
                if (string.IsNullOrEmpty(name)) continue;

                XElement valueElement = null;
                foreach (XElement child in property.Elements())
                {
                    valueElement = child;
                    break;
                }
                if (valueElement == null)
                {
                    metadata.CustomProperties.Set(name, string.Empty);
                    continue;
                }

                string text = valueElement.Value ?? string.Empty;
                switch (valueElement.Name.LocalName.ToLowerInvariant())
                {
                    case "lpwstr":
                    case "lpstr":
                    case "bstr":
                    case "string":
                        metadata.CustomProperties.Set(name, text);
                        break;
                    case "bool":
                        metadata.CustomProperties.Set(name, ParseCustomBoolean(text));
                        break;
                    case "i1":
                    case "i2":
                    case "i4":
                    case "int":
                    case "ui1":
                    case "ui2":
                    case "ui4":
                        {
                            int number;
                            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out number))
                            {
                                metadata.CustomProperties.Set(name, number);
                            }
                            else
                            {
                                metadata.CustomProperties.Set(name, text);
                            }
                        }
                        break;
                    case "i8":
                    case "ui8":
                        {
                            long number;
                            if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out number))
                            {
                                metadata.CustomProperties.Set(name, number);
                            }
                            else
                            {
                                metadata.CustomProperties.Set(name, text);
                            }
                        }
                        break;
                    case "r4":
                    case "r8":
                        {
                            double number;
                            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                            {
                                metadata.CustomProperties.Set(name, number);
                            }
                            else
                            {
                                metadata.CustomProperties.Set(name, text);
                            }
                        }
                        break;
                    case "decimal":
                        {
                            decimal number;
                            if (decimal.TryParse(text, NumberStyles.Number, CultureInfo.InvariantCulture, out number))
                            {
                                metadata.CustomProperties.Set(name, number);
                            }
                            else
                            {
                                metadata.CustomProperties.Set(name, text);
                            }
                        }
                        break;
                    case "filetime":
                    case "date":
                        {
                            DateTime timestamp;
                            if (DateTime.TryParse(
                                text,
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                                out timestamp))
                            {
                                metadata.CustomProperties.Set(name, timestamp);
                            }
                            else
                            {
                                metadata.CustomProperties.Set(name, text);
                            }
                        }
                        break;
                    default:
                        metadata.CustomProperties.Set(name, text);
                        break;
                }
            }
        }

        /// <summary>解析 OOXML 布尔文本：1/true 为真，0/false 为假，其他值按 false 处理。Parses OOXML Boolean text: 1/true is true, 0/false is false, and anything else is false.</summary>
        /// <param name="text">布尔文本。The Boolean text.</param>
        /// <returns>解析结果。The parsed value.</returns>
        private static bool ParseCustomBoolean(string text)
        {
            return string.Equals(text, "1", StringComparison.Ordinal)
                || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 尝试将元素文本按不变文化区域、按 UTC 解释并调整为通用时间，解析为 <see cref="DateTime"/>。
        /// Tries to parse the element text as a <see cref="DateTime"/> using the invariant culture, treating it
        /// as UTC and adjusting it to universal time.
        /// </summary>
        /// <param name="element">含时间戳文本的元素，可为 null。The element holding the timestamp text, may be null.</param>
        /// <param name="value">解析结果；失败时为 DateTime.MinValue。The parsed value; DateTime.MinValue on failure.</param>
        /// <returns>解析成功返回 true。True when parsing succeeded.</returns>
        private static bool TryParseTimestamp(XElement element, out DateTime value)
        {
            value = DateTime.MinValue;
            if (element == null || string.IsNullOrEmpty(element.Value)) return false;
            if (DateTime.TryParse(
                element.Value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                out value))
            {
                return true;
            }
            return false;
        }

        /// <summary>
        /// 将段落追加到块列表：若段落文本含分页符（\f），则按分页符拆分为多个段落，
        /// 段间插入 <see cref="PageBreak"/>，各拆分段落保留原段落的格式外壳。
        /// Appends a paragraph to the block list: when its text contains a form feed (\f) the paragraph is split
        /// into several paragraphs separated by <see cref="PageBreak"/> blocks, each keeping the formatting
        /// shell of the original paragraph.
        /// </summary>
        /// <param name="blocks">目标块列表。The target block list.</param>
        /// <param name="paragraph">待追加的段落。The paragraph to append.</param>
        private static void AppendParagraphBlocks(List<Block> blocks, Paragraph paragraph)
        {
            if (paragraph == null) return;
            if (!ContainsFormFeed(paragraph))
            {
                blocks.Add(paragraph);
                return;
            }

            Paragraph current = CloneShell(paragraph);
            if (paragraph.DeletedInlines.Count > 0)
            {
                current.DeletedInlines.AddRange(paragraph.DeletedInlines);
            }
            bool emittedContent = false;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run == null)
                {
                    current.Inlines.Add(paragraph.Inlines[i]);
                    emittedContent = true;
                    continue;
                }

                string[] parts = run.Text.Split('\f');
                for (int p = 0; p < parts.Length; p++)
                {
                    if (p > 0)
                    {
                        if (current.Inlines.Count > 0)
                        {
                            blocks.Add(current);
                        }
                        blocks.Add(new PageBreak());
                        current = CloneShell(paragraph);
                        emittedContent = false;
                    }
                    if (parts[p].Length > 0)
                    {
                        TextRun split = new TextRun(parts[p], run.Format.Clone());
                        split.Link = run.Link;
                        split.Revision = run.Revision == null ? null : run.Revision.Clone();
                        current.Inlines.Add(split);
                        emittedContent = true;
                    }
                }
            }

            if (emittedContent || current.Inlines.Count > 0)
            {
                blocks.Add(current);
            }
        }

        /// <summary>
        /// 判断段落的任一文本运行中是否包含分页符字符（\f）。
        /// Detects whether any text run of the paragraph contains the form-feed character (\f).
        /// </summary>
        /// <param name="paragraph">待检测的段落。The paragraph to test.</param>
        /// <returns>包含分页符时返回 true。True when a page break is present.</returns>
        private static bool ContainsFormFeed(Paragraph paragraph)
        {
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null && run.Text.IndexOf('\f') >= 0) return true;
            }
            return false;
        }

        /// <summary>
        /// 复制段落的格式外壳（对齐、间距、缩进、标题级别等），不包含行内内容，用于分页符拆分。
        /// Copies only the formatting shell of a paragraph (alignment, spacing, indents, heading level, ...)
        /// without its inline content; used when splitting on page breaks.
        /// </summary>
        /// <param name="source">源段落。The source paragraph.</param>
        /// <returns>仅带格式属性的新段落副本。A new paragraph carrying only the formatting properties.</returns>
        private static Paragraph CloneShell(Paragraph source)
        {
            Paragraph copy = new Paragraph();
            copy.Alignment = source.Alignment;
            copy.SpaceBefore = source.SpaceBefore;
            copy.SpaceAfter = source.SpaceAfter;
            copy.LineSpacing = source.LineSpacing;
            copy.LeftIndent = source.LeftIndent;
            copy.RightIndent = source.RightIndent;
            copy.FirstLineIndent = source.FirstLineIndent;
            copy.HeadingLevel = source.HeadingLevel;
            copy.KeepWithNext = source.KeepWithNext;
            copy.PageBreakBefore = source.PageBreakBefore;
            copy.StyleName = source.StyleName;
            copy.ListMarker = source.ListMarker;
            copy.BookmarkName = source.BookmarkName;
            copy.Revision = source.Revision == null ? null : source.Revision.Clone();
            return copy;
        }
    }
}
