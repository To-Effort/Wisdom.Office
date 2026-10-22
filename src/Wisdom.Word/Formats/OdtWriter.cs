using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 将统一文档模型写成 OpenDocument 文本（.odt）：生成 mimetype、manifest、content.xml、
    /// styles.xml 与 meta.xml，输出段落/字符样式、页眉页脚、分页、表格（含合并与底色）、
    /// 书签、超链接与内嵌 PNG/JPEG 图片。
    /// Writes the unified document model as OpenDocument Text (.odt): produces mimetype, manifest, content.xml,
    /// styles.xml and meta.xml, with paragraph/character styles, headers and footers, page breaks, tables
    /// (including merges and shading), bookmarks, hyperlinks and embedded PNG/JPEG images.
    /// </summary>
    public static class OdtWriter
    {
        /// <summary>ODF 1.3 文本文档的 MIME 类型。The MIME type of an ODF 1.3 text document.</summary>
        private const string MimeType = "application/vnd.oasis.opendocument.text";

        /// <summary>将文档保存为 ODT 文件。Saves the document as an ODT file.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <param name="path">目标文件路径。The target file path.</param>
        public static void Write(OfficeDocument document, string path)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (path == null) throw new ArgumentNullException("path");
            File.WriteAllBytes(path, Write(document));
        }

        /// <summary>将文档写成 ODT 字节并写入流。Writes the document as ODT bytes into a stream.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <param name="stream">目标流。The target stream.</param>
        public static void Write(OfficeDocument document, Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            byte[] bytes = Write(document);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>将文档写成 ODT 字节数组。Writes the document as an ODT byte array.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <returns>ODT 文件字节。The ODT file bytes.</returns>
        public static byte[] Write(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            using (MemoryStream memory = new MemoryStream())
            {
                using (ZipArchive archive = new ZipArchive(memory, ZipArchiveMode.Create, true))
                {
                    WritePackage(document, archive);
                }
                return memory.ToArray();
            }
        }

        /// <summary>组装 ODT 压缩包的全部部件。Assembles all parts of the ODT package.</summary>
        private static void WritePackage(OfficeDocument document, ZipArchive archive)
        {
            StyleContext bodyStyles = new StyleContext(string.Empty);
            StyleContext headerStyles = new StyleContext("H");

            List<Section> sections = document.Sections.Count > 0 ? document.Sections : new List<Section> { new Section() };
            StringBuilder masterPages = BuildMasterPages(sections, headerStyles);
            string content = BuildContent(document, sections, bodyStyles);
            string styles = BuildStyles(document, sections, headerStyles, masterPages);
            string meta = BuildMeta(document.Metadata);

            ZipArchiveEntry mimetype = archive.CreateEntry("mimetype", CompressionLevel.NoCompression);
            using (Stream stream = mimetype.Open())
            {
                byte[] bytes = Encoding.ASCII.GetBytes(MimeType);
                stream.Write(bytes, 0, bytes.Length);
            }
            WriteEntry(archive, "content.xml", content);
            WriteEntry(archive, "styles.xml", styles);
            WriteEntry(archive, "meta.xml", meta);
            WriteEntry(archive, "META-INF/manifest.xml", BuildManifest());
        }

        /// <summary>写入一个 UTF-8 XML 部件。Writes one UTF-8 XML part.</summary>
        private static void WriteEntry(ZipArchive archive, string name, string content)
        {
            ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream stream = entry.Open())
            {
                byte[] bytes = new UTF8Encoding(false).GetBytes(content);
                stream.Write(bytes, 0, bytes.Length);
            }
        }

        /// <summary>输出内容类型清单。Writes the content-type manifest.</summary>
        private static string BuildManifest()
        {
            StringBuilder xml = new StringBuilder(512);
            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            xml.Append("<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\">\n");
            xml.Append("  <manifest:file-entry manifest:full-path=\"/\" manifest:version=\"1.3\" manifest:media-type=\"").Append(MimeType).Append("\"/>\n");
            xml.Append("  <manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>\n");
            xml.Append("  <manifest:file-entry manifest:full-path=\"styles.xml\" manifest:media-type=\"text/xml\"/>\n");
            xml.Append("  <manifest:file-entry manifest:full-path=\"meta.xml\" manifest:media-type=\"text/xml\"/>\n");
            xml.Append("</manifest:manifest>\n");
            return xml.ToString();
        }

        /// <summary>输出文档元数据部件。Writes the document metadata part.</summary>
        private static string BuildMeta(DocumentMetadata metadata)
        {
            StringBuilder xml = new StringBuilder(1024);
            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            xml.Append("<office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"")
                .Append(" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\"")
                .Append(" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" office:version=\"1.3\">\n");
            xml.Append("<office:meta>");
            if (!string.IsNullOrEmpty(metadata.Title)) xml.Append("<dc:title>").Append(WordFormatUtil.EscapeXml(metadata.Title)).Append("</dc:title>");
            if (!string.IsNullOrEmpty(metadata.Author)) xml.Append("<dc:creator>").Append(WordFormatUtil.EscapeXml(metadata.Author)).Append("</dc:creator>");
            if (!string.IsNullOrEmpty(metadata.Subject)) xml.Append("<dc:subject>").Append(WordFormatUtil.EscapeXml(metadata.Subject)).Append("</dc:subject>");
            if (!string.IsNullOrEmpty(metadata.Keywords)) xml.Append("<meta:keyword>").Append(WordFormatUtil.EscapeXml(metadata.Keywords)).Append("</meta:keyword>");
            if (!string.IsNullOrEmpty(metadata.Creator)) xml.Append("<meta:generator>").Append(WordFormatUtil.EscapeXml(metadata.Creator)).Append("</meta:generator>");
            if (metadata.Created != default(DateTime))
            {
                xml.Append("<meta:creation-date>").Append(metadata.Created.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture)).Append("</meta:creation-date>");
            }
            if (metadata.Modified != default(DateTime))
            {
                xml.Append("<dc:date>").Append(metadata.Modified.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture)).Append("</dc:date>");
            }
            xml.Append("</office:meta>\n</office:document-meta>\n");
            return xml.ToString();
        }

        // ---------------------------------------------------------------- styles

        /// <summary>
        /// 样式生成器：按需生成段落、文本、表格单元格、行与列样式，并缓存重复定义。
        /// Style generator: produces paragraph, text, table-cell, row and column styles on demand, caching repeated definitions.
        /// </summary>
        private sealed class StyleContext
        {
            /// <summary>样式名前缀，用于区分正文与页眉页脚的自动样式。Style-name prefix separating body automatic styles from header/footer ones.</summary>
            private readonly string _prefix;
            /// <summary>自动样式 XML 片段。The automatic-style XML fragments.</summary>
            public readonly StringBuilder Automatic = new StringBuilder();
            /// <summary>字体声明 XML 片段。The font-face declaration XML fragments.</summary>
            public readonly StringBuilder FontFaces = new StringBuilder();
            private readonly Dictionary<string, string> _paragraphStyles = new Dictionary<string, string>();
            private readonly Dictionary<string, string> _textStyles = new Dictionary<string, string>();
            private readonly Dictionary<string, string> _cellStyles = new Dictionary<string, string>();
            private readonly Dictionary<string, string> _rowStyles = new Dictionary<string, string>();
            private readonly Dictionary<string, string> _columnStyles = new Dictionary<string, string>();
            private readonly HashSet<string> _fonts = new HashSet<string>(StringComparer.Ordinal);
            private int _paragraphIndex;
            private int _textIndex;
            private int _cellIndex;
            private int _rowIndex;
            private int _columnIndex;
            private int _tableIndex;

            /// <summary>以指定前缀创建样式生成器。Creates a style generator with the given prefix.</summary>
            public StyleContext(string prefix)
            {
                _prefix = prefix ?? string.Empty;
            }

            /// <summary>注册字体并输出字体声明。Registers a font and emits its font-face declaration.</summary>
            public void AddFont(string name)
            {
                if (string.IsNullOrEmpty(name) || !_fonts.Add(name)) return;
                FontFaces.Append("<style:font-face style:name=\"").Append(WordFormatUtil.EscapeXml(name))
                    .Append("\" svg:font-family=\"").Append(WordFormatUtil.EscapeXml(name)).Append("\"/>");
            }

            /// <summary>生成段落样式并返回样式名。Builds a paragraph style and returns its name.</summary>
            public string ParagraphStyle(Paragraph paragraph, string masterPage, bool breakBefore)
            {
                string properties = BuildParagraphProperties(paragraph, breakBefore);
                string key = properties + "|" + (masterPage ?? string.Empty);
                string name;
                if (_paragraphStyles.TryGetValue(key, out name)) return name;
                _paragraphIndex++;
                name = _prefix + "P" + _paragraphIndex.ToString(CultureInfo.InvariantCulture);
                _paragraphStyles[key] = name;
                Automatic.Append("<style:style style:name=\"").Append(name).Append("\" style:family=\"paragraph\"");
                if (!string.IsNullOrEmpty(masterPage)) Automatic.Append(" style:master-page-name=\"").Append(masterPage).Append("\"");
                Automatic.Append('>');
                if (properties.Length > 0) Automatic.Append("<style:paragraph-properties ").Append(properties).Append("/>");
                Automatic.Append("</style:style>");
                return name;
            }

            /// <summary>生成分页空段落的样式；可同时指定该段开始的主页面。Builds the style for an empty page-break paragraph, optionally naming the master page that starts there.</summary>
            public string PageBreakStyle(string masterPage)
            {
                string key = "fo:break-before=\"page\"|" + (masterPage ?? string.Empty);
                string name;
                if (_paragraphStyles.TryGetValue(key, out name)) return name;
                _paragraphIndex++;
                name = _prefix + "P" + _paragraphIndex.ToString(CultureInfo.InvariantCulture);
                _paragraphStyles[key] = name;
                Automatic.Append("<style:style style:name=\"").Append(name).Append("\" style:family=\"paragraph\"");
                if (!string.IsNullOrEmpty(masterPage)) Automatic.Append(" style:master-page-name=\"").Append(masterPage).Append("\"");
                Automatic.Append("><style:paragraph-properties fo:break-before=\"page\"/></style:style>");
                return name;
            }

            /// <summary>生成一个表格编号。Produces a table name sequence value.</summary>
            public int NextTableIndex()
            {
                _tableIndex++;
                return _tableIndex;
            }

            /// <summary>构造段落的属性文本。Builds the paragraph property text.</summary>
            private static string BuildParagraphProperties(Paragraph paragraph, bool breakBefore)
            {
                StringBuilder properties = new StringBuilder();
                switch (paragraph.Alignment)
                {
                    case ParagraphAlignment.Center: AppendProperty(properties, "fo:text-align", "center"); break;
                    case ParagraphAlignment.Right: AppendProperty(properties, "fo:text-align", "end"); break;
                    case ParagraphAlignment.Justify: AppendProperty(properties, "fo:text-align", "justify"); break;
                }
                if (paragraph.SpaceBefore != 0) AppendProperty(properties, "fo:margin-top", Points(paragraph.SpaceBefore));
                if (paragraph.SpaceAfter != 0) AppendProperty(properties, "fo:margin-bottom", Points(paragraph.SpaceAfter));
                if (paragraph.LeftIndent != 0) AppendProperty(properties, "fo:margin-left", Points(paragraph.LeftIndent));
                if (paragraph.RightIndent != 0) AppendProperty(properties, "fo:margin-right", Points(paragraph.RightIndent));
                if (paragraph.FirstLineIndent != 0) AppendProperty(properties, "fo:text-indent", Points(paragraph.FirstLineIndent));
                if (paragraph.LineSpacing > 0 && Math.Abs(paragraph.LineSpacing - 1) > 0.001)
                {
                    AppendProperty(properties, "fo:line-height", WordFormatUtil.Number(paragraph.LineSpacing * 100) + "%");
                }
                if (paragraph.KeepWithNext) AppendProperty(properties, "fo:keep-with-next", "always");
                if (breakBefore || paragraph.PageBreakBefore) AppendProperty(properties, "fo:break-before", "page");
                return properties.ToString();
            }

            /// <summary>生成文本运行样式并返回样式名；默认格式也生成样式以保证读回一致。Builds a text-run style and returns its name; even default formatting gets a style so it reads back consistently.</summary>
            public string TextStyle(TextFormat format)
            {
                format = format ?? new TextFormat();
                string ascii = format.GetAsciiFontFamily();
                string eastAsia = format.GetEastAsiaFontFamily();
                AddFont(ascii);
                if (!string.Equals(ascii, eastAsia, StringComparison.Ordinal)) AddFont(eastAsia);
                StringBuilder properties = new StringBuilder();
                if (format.FontSize > 0) AppendProperty(properties, "fo:font-size", Points(format.FontSize));
                AppendProperty(properties, "fo:font-family", ascii);
                if (!string.Equals(ascii, eastAsia, StringComparison.Ordinal))
                {
                    AppendProperty(properties, "style:font-family-asian", eastAsia);
                }
                if (format.Bold) AppendProperty(properties, "fo:font-weight", "bold");
                if (format.Italic) AppendProperty(properties, "fo:font-style", "italic");
                if (format.Underline) AppendProperty(properties, "style:text-underline-style", "solid");
                if (format.Strikeout) AppendProperty(properties, "style:text-line-through-style", "solid");
                if (format.Superscript) AppendProperty(properties, "style:text-position", "super 58%");
                else if (format.Subscript) AppendProperty(properties, "style:text-position", "sub 58%");
                OfficeColor color = format.Color;
                if (!(color.R == 0 && color.G == 0 && color.B == 0)) AppendProperty(properties, "fo:color", Hex(color));
                string key = properties.ToString();
                string name;
                if (_textStyles.TryGetValue(key, out name)) return name;
                _textIndex++;
                name = _prefix + "T" + _textIndex.ToString(CultureInfo.InvariantCulture);
                _textStyles[key] = name;
                Automatic.Append("<style:style style:name=\"").Append(name).Append("\" style:family=\"text\"><style:text-properties ")
                    .Append(key).Append("/></style:style>");
                return name;
            }

            /// <summary>生成单元格样式并返回样式名。Builds a table-cell style and returns its name.</summary>
            public string CellStyle(TableCell cell, Table table, BorderLine top, BorderLine bottom, BorderLine left, BorderLine right)
            {
                StringBuilder properties = new StringBuilder();
                if (cell.BackgroundColor.HasValue) AppendProperty(properties, "fo:background-color", Hex(cell.BackgroundColor.Value));
                properties.Append("fo:border-top=\"").Append(BorderText(top)).Append("\" ");
                properties.Append("fo:border-bottom=\"").Append(BorderText(bottom)).Append("\" ");
                properties.Append("fo:border-left=\"").Append(BorderText(left)).Append("\" ");
                properties.Append("fo:border-right=\"").Append(BorderText(right)).Append("\" ");
                if (table.CellPadding > 0) AppendProperty(properties, "fo:padding", Points(table.CellPadding));
                switch (cell.VerticalAlignment)
                {
                    case VerticalAlignment.Middle: AppendProperty(properties, "style:vertical-align", "middle"); break;
                    case VerticalAlignment.Bottom: AppendProperty(properties, "style:vertical-align", "bottom"); break;
                }
                string key = properties.ToString();
                string name;
                if (_cellStyles.TryGetValue(key, out name)) return name;
                _cellIndex++;
                name = _prefix + "C" + _cellIndex.ToString(CultureInfo.InvariantCulture);
                _cellStyles[key] = name;
                Automatic.Append("<style:style style:name=\"").Append(name).Append("\" style:family=\"table-cell\"><style:table-cell-properties ")
                    .Append(key).Append("/></style:style>");
                return name;
            }

            /// <summary>生成表行样式并返回样式名。Builds a table-row style and returns its name.</summary>
            public string RowStyle(double height)
            {
                string key = Points(height);
                string name;
                if (_rowStyles.TryGetValue(key, out name)) return name;
                _rowIndex++;
                name = _prefix + "R" + _rowIndex.ToString(CultureInfo.InvariantCulture);
                _rowStyles[key] = name;
                Automatic.Append("<style:style style:name=\"").Append(name).Append("\" style:family=\"table-row\"><style:table-row-properties style:row-height=\"")
                    .Append(key).Append("\"/></style:style>");
                return name;
            }

            /// <summary>生成表格列样式并返回样式名。Builds a table-column style and returns its name.</summary>
            public string ColumnStyle(double width)
            {
                string key = Points(width);
                string name;
                if (_columnStyles.TryGetValue(key, out name)) return name;
                _columnIndex++;
                name = _prefix + "Col" + _columnIndex.ToString(CultureInfo.InvariantCulture);
                _columnStyles[key] = name;
                Automatic.Append("<style:style style:name=\"").Append(name).Append("\" style:family=\"table-column\"><style:table-column-properties style:column-width=\"")
                    .Append(key).Append("\"/></style:style>");
                return name;
            }

            /// <summary>输出一个属性（键值对），并自动添加空格分隔。Writes one property (key-value pair) with automatic space separation.</summary>
            private static void AppendProperty(StringBuilder builder, string name, string value)
            {
                builder.Append(name).Append("=\"").Append(WordFormatUtil.EscapeXml(value)).Append("\" ");
            }

            /// <summary>格式化为磅值文本。Formats a value as points.</summary>
            private static string Points(double value)
            {
                return WordFormatUtil.Number(value) + "pt";
            }

            /// <summary>输出 #RRGGBB 颜色。Writes a #RRGGBB colour.</summary>
            private static string Hex(OfficeColor color)
            {
                return "#" + color.R.ToString("X2", CultureInfo.InvariantCulture)
                    + color.G.ToString("X2", CultureInfo.InvariantCulture)
                    + color.B.ToString("X2", CultureInfo.InvariantCulture);
            }

            /// <summary>输出 ODF 边框文本。Writes the ODF border text.</summary>
            private static string BorderText(BorderLine line)
            {
                if (line == null || line.Style == BorderStyle.None) return "none";
                string style = "solid";
                if (line.Style == BorderStyle.Dashed) style = "dashed";
                else if (line.Style == BorderStyle.Dotted) style = "dotted";
                else if (line.Style == BorderStyle.Double) style = "double";
                double width = line.Width > 0 ? line.Width : 0.5;
                return WordFormatUtil.Number(width) + "pt " + style + " " + Hex(line.Color);
            }
        }

        // ---------------------------------------------------------------- content

        /// <summary>构建 content.xml。Builds content.xml.</summary>
        private static string BuildContent(OfficeDocument document, List<Section> sections, StyleContext styles)
        {
            StringBuilder body = new StringBuilder(8192);
            bool firstSection = true;
            for (int s = 0; s < sections.Count; s++)
            {
                Section section = sections[s];
                string masterPage = "Standard" + (s == 0 ? string.Empty : (s + 1).ToString(CultureInfo.InvariantCulture));
                bool needMaster = true;
                bool breakBefore = !firstSection;
                for (int i = 0; i < section.Blocks.Count; i++)
                {
                    Block block = section.Blocks[i];
                    if (needMaster)
                    {
                        Paragraph firstParagraph = block as Paragraph;
                        if (firstParagraph != null)
                        {
                            WriteParagraph(body, firstParagraph, styles, masterPage, breakBefore);
                            needMaster = false;
                            breakBefore = false;
                            continue;
                        }
                        if (block is Table)
                        {
                            body.Append("<text:p text:style-name=\"").Append(styles.ParagraphStyle(new Paragraph(), masterPage, breakBefore)).Append("\"/>");
                            needMaster = false;
                            breakBefore = false;
                        }
                        else if (block is PageBreak)
                        {
                            body.Append("<text:p text:style-name=\"").Append(styles.PageBreakStyle(breakBefore ? masterPage : null)).Append("\"/>");
                            needMaster = false;
                            breakBefore = false;
                            continue;
                        }
                    }
                    WriteBlock(body, block, styles, null, false);
                }
                if (needMaster && !firstSection)
                {
                    body.Append("<text:p text:style-name=\"").Append(styles.PageBreakStyle(masterPage)).Append("\"/>");
                }
                firstSection = false;
            }

            StringBuilder xml = new StringBuilder(body.Length + styles.Automatic.Length + 2048);
            AppendContentHeader(xml);
            if (styles.FontFaces.Length > 0)
            {
                xml.Append("<office:font-face-decls>").Append(styles.FontFaces).Append("</office:font-face-decls>\n");
            }
            xml.Append("<office:automatic-styles>").Append(styles.Automatic).Append("</office:automatic-styles>\n");
            xml.Append("<office:body><office:text>");
            xml.Append(body);
            xml.Append("</office:text></office:body>\n</office:document-content>\n");
            return xml.ToString();
        }

        /// <summary>输出 content.xml 根元素与命名空间声明。Writes the content.xml root element and namespace declarations.</summary>
        private static void AppendContentHeader(StringBuilder xml)
        {
            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            xml.Append("<office:document-content")
                .Append(" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"")
                .Append(" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\"")
                .Append(" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"")
                .Append(" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"")
                .Append(" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\"")
                .Append(" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\"")
                .Append(" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\"")
                .Append(" xmlns:xlink=\"http://www.w3.org/1999/xlink\"")
                .Append(" xmlns:meta=\"urn:oasis:names:tc:opendocument:xmlns:meta:1.0\"")
                .Append(" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"")
                .Append(" office:version=\"1.3\">\n");
        }

        /// <summary>输出一个块级元素。Writes one block element.</summary>
        private static void WriteBlock(StringBuilder xml, Block block, StyleContext styles, string masterPage, bool breakBefore)
        {
            Paragraph paragraph = block as Paragraph;
            if (paragraph != null)
            {
                WriteParagraph(xml, paragraph, styles, masterPage, breakBefore);
                return;
            }
            Table table = block as Table;
            if (table != null)
            {
                WriteTable(xml, table, styles);
                return;
            }
            if (block is PageBreak)
            {
                xml.Append("<text:p text:style-name=\"").Append(styles.PageBreakStyle(null)).Append("\"/>");
            }
        }

        /// <summary>输出一个段落或标题。Writes one paragraph or heading.</summary>
        private static void WriteParagraph(StringBuilder xml, Paragraph paragraph, StyleContext styles, string masterPage, bool breakBefore)
        {
            string style = styles.ParagraphStyle(paragraph, masterPage, breakBefore);
            bool heading = paragraph.HeadingLevel > 0;
            if (heading)
            {
                xml.Append("<text:h text:outline-level=\"").Append(Math.Min(10, paragraph.HeadingLevel).ToString(CultureInfo.InvariantCulture))
                    .Append("\" text:style-name=\"").Append(style).Append("\">");
            }
            else
            {
                xml.Append("<text:p text:style-name=\"").Append(style).Append("\">");
            }
            if (!string.IsNullOrEmpty(paragraph.BookmarkName) && !HasStartMarker(paragraph, paragraph.BookmarkName))
            {
                xml.Append("<text:bookmark text:name=\"").Append(WordFormatUtil.EscapeXml(paragraph.BookmarkName)).Append("\"/>");
            }
            if (!string.IsNullOrEmpty(paragraph.ListMarker))
            {
                xml.Append(EncodeText(paragraph.ListMarker)).Append("<text:tab/>");
            }
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                WriteInline(xml, paragraph.Inlines[i], styles);
            }
            xml.Append(heading ? "</text:h>" : "</text:p>");
        }

        /// <summary>段落内是否已有该名称的显式书签起点。Determines whether the paragraph holds an explicit start marker of the name.</summary>
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

        /// <summary>输出一个行内元素。Writes one inline element.</summary>
        private static void WriteInline(StringBuilder xml, Inline inline, StyleContext styles)
        {
            BookmarkStart startMarker = inline as BookmarkStart;
            if (startMarker != null)
            {
                xml.Append("<text:bookmark-start text:name=\"")
                    .Append(WordFormatUtil.EscapeXml(startMarker.Name)).Append("\"/>");
                return;
            }
            BookmarkEnd endMarker = inline as BookmarkEnd;
            if (endMarker != null)
            {
                xml.Append("<text:bookmark-end text:name=\"")
                    .Append(WordFormatUtil.EscapeXml(endMarker.Name)).Append("\"/>");
                return;
            }

            TextRun run = inline as TextRun;
            if (run != null)
            {
                string style = styles.TextStyle(run.Format);
                bool linked = !string.IsNullOrEmpty(run.Link)
                    && (run.Link.Trim().StartsWith("#", StringComparison.Ordinal) || WordFormatUtil.IsSafeHyperlinkTarget(run.Link.Trim()));
                if (linked)
                {
                    xml.Append("<text:a xlink:type=\"simple\" xlink:href=\"").Append(WordFormatUtil.EscapeXml(run.Link.Trim())).Append("\">");
                }
                xml.Append("<text:span text:style-name=\"").Append(style).Append("\">");
                xml.Append(EncodeText(run.Text));
                xml.Append("</text:span>");
                if (linked) xml.Append("</text:a>");
                return;
            }

            FieldRun field = inline as FieldRun;
            if (field != null)
            {
                if (field.Instruction != null && field.Instruction.IndexOf("PAGE", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    xml.Append("<text:page-number text:select-page=\"current\"/>");
                    return;
                }
                string style = styles.TextStyle(field.Format);
                xml.Append("<text:span text:style-name=\"").Append(style).Append("\">")
                    .Append(EncodeText(field.DisplayText)).Append("</text:span>");
                return;
            }

            ImageRun image = inline as ImageRun;
            if (image != null && image.Image != null)
            {
                double width = image.Width > 0 ? image.Width : image.Image.DefaultWidthPoints;
                double height = image.Height > 0 ? image.Height : image.Image.DefaultHeightPoints;
                xml.Append("<draw:frame draw:name=\"Image\" text:anchor-type=\"as-char\" svg:width=\"")
                    .Append(WordFormatUtil.Number(width)).Append("pt\" svg:height=\"").Append(WordFormatUtil.Number(height)).Append("pt\">")
                    .Append("<draw:image><office:binary-data>")
                    .Append(Convert.ToBase64String(image.Image.Data))
                    .Append("</office:binary-data></draw:image></draw:frame>");
            }
        }

        /// <summary>把文本编码为 ODF 行内标记：多空格写成 text:s、制表符与换行写成对应元素。Encodes text into ODF inline markup: space runs become text:s, tabs and newlines become their elements.</summary>
        private static string EncodeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            StringBuilder xml = new StringBuilder(text.Length + 16);
            int index = 0;
            while (index < text.Length)
            {
                char current = text[index];
                if (current == ' ')
                {
                    int start = index;
                    while (index < text.Length && text[index] == ' ') index++;
                    int count = index - start;
                    if (count == 1 && start > 0 && index < text.Length) xml.Append(' ');
                    else if (count <= 1) xml.Append("<text:s/>");
                    else xml.Append("<text:s text:c=\"").Append(count.ToString(CultureInfo.InvariantCulture)).Append("\"/>");
                    continue;
                }
                if (current == '\t')
                {
                    xml.Append("<text:tab/>");
                    index++;
                    continue;
                }
                if (current == '\n')
                {
                    xml.Append("<text:line-break/>");
                    index++;
                    continue;
                }
                int plainStart = index;
                while (index < text.Length && text[index] != ' ' && text[index] != '\t' && text[index] != '\n') index++;
                xml.Append(WordFormatUtil.EscapeXml(text.Substring(plainStart, index - plainStart)));
            }
            return xml.ToString();
        }

        // ---------------------------------------------------------------- tables

        /// <summary>输出一个表格：列定义、按合并关系展开的行列网格、单元格内容。Writes a table: column definitions, a row/column grid expanded for merges, and cell content.</summary>
        private static void WriteTable(StringBuilder xml, Table table, StyleContext styles)
        {
            int rowCount = table.Rows.Count;
            if (rowCount == 0) return;
            int columnCount = 0;
            for (int r = 0; r < rowCount; r++)
            {
                int count = 0;
                for (int c = 0; c < table.Rows[r].Cells.Count; c++) count += Math.Max(1, table.Rows[r].Cells[c].ColSpan);
                columnCount = Math.Max(columnCount, count);
            }
            columnCount = Math.Max(columnCount, table.Columns.Count);
            if (columnCount <= 0) columnCount = 1;

            int[,] grid = new int[rowCount, columnCount];
            TableCell[,] origins = new TableCell[rowCount, columnCount];
            for (int r = 0; r < rowCount; r++)
            {
                int column = 0;
                TableRow row = table.Rows[r];
                for (int c = 0; c < row.Cells.Count; c++)
                {
                    while (column < columnCount && grid[r, column] != 0) column++;
                    if (column >= columnCount) break;
                    TableCell cell = row.Cells[c];
                    int colSpan = Math.Min(Math.Max(1, cell.ColSpan), columnCount - column);
                    int rowSpan = Math.Min(Math.Max(1, cell.RowSpan), rowCount - r);
                    for (int dr = 0; dr < rowSpan; dr++)
                    {
                        for (int dc = 0; dc < colSpan; dc++)
                        {
                            grid[r + dr, column + dc] = dr == 0 && dc == 0 ? 1 : 2;
                            if (dr == 0 && dc == 0) origins[r, column] = cell;
                        }
                    }
                    column += colSpan;
                }
            }

            xml.Append("<table:table table:name=\"Table").Append(styles.NextTableIndex().ToString(CultureInfo.InvariantCulture)).Append("\"");
            if (table.Alignment == TableAlignment.Center) xml.Append(" table:align=\"center\"");
            else if (table.Alignment == TableAlignment.Right) xml.Append(" table:align=\"right\"");
            xml.Append('>');
            if (table.Columns.Count > 0)
            {
                for (int c = 0; c < columnCount; c++)
                {
                    double width = c < table.Columns.Count ? table.Columns[c].Width : 0;
                    if (width > 0)
                    {
                        xml.Append("<table:table-column table:style-name=\"").Append(styles.ColumnStyle(width)).Append("\"/>");
                    }
                    else
                    {
                        xml.Append("<table:table-column/>");
                    }
                }
            }
            else
            {
                xml.Append("<table:table-column table:number-columns-repeated=\"").Append(columnCount.ToString(CultureInfo.InvariantCulture)).Append("\"/>");
            }

            bool inHeaderGroup = false;
            for (int r = 0; r < rowCount; r++)
            {
                TableRow row = table.Rows[r];
                if (row.IsHeader && !inHeaderGroup)
                {
                    xml.Append("<table:table-header-rows>");
                    inHeaderGroup = true;
                }
                else if (!row.IsHeader && inHeaderGroup)
                {
                    xml.Append("</table:table-header-rows>");
                    inHeaderGroup = false;
                }
                xml.Append("<table:table-row");
                if (row.Height > 0) xml.Append(" table:style-name=\"").Append(styles.RowStyle(row.Height)).Append("\"");
                xml.Append('>');
                for (int c = 0; c < columnCount; c++)
                {
                    int kind = grid[r, c];
                    if (kind == 2)
                    {
                        xml.Append("<table:covered-table-cell/>");
                        continue;
                    }
                    if (kind == 0)
                    {
                        xml.Append("<table:table-cell/>");
                        continue;
                    }
                    TableCell cell = origins[r, c];
                    if (cell == null) continue;
                    int colSpan = Math.Max(1, cell.ColSpan);
                    int rowSpan = Math.Max(1, cell.RowSpan);
                    BorderLine top = CellBorder(cell, table, r, c, rowSpan, colSpan, 0);
                    BorderLine bottom = CellBorder(cell, table, r, c, rowSpan, colSpan, 1);
                    BorderLine left = CellBorder(cell, table, r, c, rowSpan, colSpan, 2);
                    BorderLine right = CellBorder(cell, table, r, c, rowSpan, colSpan, 3);
                    xml.Append("<table:table-cell table:style-name=\"").Append(styles.CellStyle(cell, table, top, bottom, left, right)).Append("\"");
                    if (colSpan > 1) xml.Append(" table:number-columns-spanned=\"").Append(colSpan.ToString(CultureInfo.InvariantCulture)).Append("\"");
                    if (rowSpan > 1) xml.Append(" table:number-rows-spanned=\"").Append(rowSpan.ToString(CultureInfo.InvariantCulture)).Append("\"");
                    xml.Append('>');
                    for (int b = 0; b < cell.Blocks.Count; b++)
                    {
                        WriteBlock(xml, cell.Blocks[b], styles, null, false);
                    }
                    xml.Append("</table:table-cell>");
                }
                xml.Append("</table:table-row>");
            }
            if (inHeaderGroup) xml.Append("</table:table-header-rows>");
            xml.Append("</table:table>");
        }

        /// <summary>计算单元格某条边的实际边框：单元格覆盖优先，否则按表格外框/内线。Computes one cell edge's effective border: the cell override wins, otherwise table outer/inner lines apply.</summary>
        private static BorderLine CellBorder(TableCell cell, Table table, int row, int column, int rowSpan, int colSpan, int edge)
        {
            int rowCount = table.Rows.Count;
            int columnCount = table.Columns.Count;
            for (int r = 0; columnCount == 0 && r < rowCount; r++)
            {
                int count = 0;
                for (int c = 0; c < table.Rows[r].Cells.Count; c++) count += Math.Max(1, table.Rows[r].Cells[c].ColSpan);
                columnCount = Math.Max(columnCount, count);
            }
            if (columnCount <= 0) columnCount = 1;
            int endRow = row + rowSpan - 1;
            int endColumn = column + colSpan - 1;
            if (edge == 0)
            {
                return cell.Borders != null && cell.Borders.Top != null ? cell.Borders.Top
                    : (row == 0 ? table.Borders.Top : table.Borders.InsideHorizontal);
            }
            if (edge == 1)
            {
                return cell.Borders != null && cell.Borders.Bottom != null ? cell.Borders.Bottom
                    : (endRow >= rowCount - 1 ? table.Borders.Bottom : table.Borders.InsideHorizontal);
            }
            if (edge == 2)
            {
                return cell.Borders != null && cell.Borders.Left != null ? cell.Borders.Left
                    : (column == 0 ? table.Borders.Left : table.Borders.InsideVertical);
            }
            return cell.Borders != null && cell.Borders.Right != null ? cell.Borders.Right
                : (endColumn >= columnCount - 1 ? table.Borders.Right : table.Borders.InsideVertical);
        }

        // ---------------------------------------------------------------- styles.xml

        /// <summary>构建 styles.xml，包含页面布局、主页面以及页眉页脚内容。Builds styles.xml with page layouts, master pages and header/footer content.</summary>
        private static string BuildStyles(OfficeDocument document, List<Section> sections, StyleContext styles, StringBuilder masterPages)
        {
            StringBuilder xml = new StringBuilder(masterPages.Length + styles.Automatic.Length + 4096);
            xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
            xml.Append("<office:document-styles")
                .Append(" xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"")
                .Append(" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\"")
                .Append(" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"")
                .Append(" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"")
                .Append(" xmlns:draw=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\"")
                .Append(" xmlns:fo=\"urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0\"")
                .Append(" xmlns:svg=\"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0\"")
                .Append(" xmlns:xlink=\"http://www.w3.org/1999/xlink\"")
                .Append(" office:version=\"1.3\">\n");
            if (styles.FontFaces.Length > 0)
            {
                xml.Append("<office:font-face-decls>").Append(styles.FontFaces).Append("</office:font-face-decls>\n");
            }
            xml.Append("<office:styles>");
            xml.Append("<style:default-style style:family=\"paragraph\"><style:paragraph-properties fo:orphans=\"2\" fo:widows=\"2\"/>")
                .Append("<style:text-properties fo:font-size=\"11pt\"/></style:default-style>");
            xml.Append("<style:style style:name=\"Standard\" style:family=\"paragraph\" style:class=\"text\"/>");
            for (int level = 1; level <= 6; level++)
            {
                xml.Append("<style:style style:name=\"Heading").Append(level).Append("\" style:family=\"paragraph\"><style:paragraph-properties fo:keep-with-next=\"always\"/>")
                    .Append("<style:text-properties fo:font-weight=\"bold\" fo:font-size=\"").Append((26 - level * 2).ToString(CultureInfo.InvariantCulture)).Append("pt\"/></style:style>");
            }
            xml.Append("</office:styles>\n");
            xml.Append("<office:automatic-styles>");
            for (int s = 0; s < sections.Count; s++)
            {
                PageSetup setup = sections[s].PageSetup ?? new PageSetup();
                xml.Append("<style:page-layout style:name=\"pm").Append((s + 1).ToString(CultureInfo.InvariantCulture)).Append("\">")
                    .Append("<style:page-layout-properties fo:page-width=\"").Append(WordFormatUtil.Number(setup.Width)).Append("pt\"")
                    .Append(" fo:page-height=\"").Append(WordFormatUtil.Number(setup.Height)).Append("pt\"")
                    .Append(" style:print-orientation=\"").Append(setup.Width > setup.Height ? "landscape" : "portrait").Append("\"")
                    .Append(" fo:margin-top=\"").Append(WordFormatUtil.Number(setup.MarginTop)).Append("pt\"")
                    .Append(" fo:margin-bottom=\"").Append(WordFormatUtil.Number(setup.MarginBottom)).Append("pt\"")
                    .Append(" fo:margin-left=\"").Append(WordFormatUtil.Number(setup.MarginLeft)).Append("pt\"")
                    .Append(" fo:margin-right=\"").Append(WordFormatUtil.Number(setup.MarginRight)).Append("pt\"/>")
                    .Append("<style:header-style><style:header-footer-properties fo:min-height=\"").Append(WordFormatUtil.Number(setup.HeaderDistance)).Append("pt\"/></style:header-style>")
                    .Append("<style:footer-style><style:header-footer-properties fo:min-height=\"").Append(WordFormatUtil.Number(setup.FooterDistance)).Append("pt\"/></style:footer-style>")
                    .Append("</style:page-layout>");
            }
            xml.Append(styles.Automatic);
            xml.Append("</office:automatic-styles>\n");
            xml.Append("<office:master-styles>");
            xml.Append(masterPages);
            xml.Append("</office:master-styles>\n");
            xml.Append("</office:document-styles>\n");
            return xml.ToString();
        }

        /// <summary>构建全部主页面及其页眉页脚。Builds every master page together with its header and footer.</summary>
        private static StringBuilder BuildMasterPages(List<Section> sections, StyleContext styles)
        {
            StringBuilder pages = new StringBuilder(2048);
            for (int s = 0; s < sections.Count; s++)
            {
                Section section = sections[s];
                string name = "Standard" + (s == 0 ? string.Empty : (s + 1).ToString(CultureInfo.InvariantCulture));
                pages.Append("<style:master-page style:name=\"").Append(name).Append("\" style:page-layout-name=\"pm")
                    .Append((s + 1).ToString(CultureInfo.InvariantCulture)).Append("\">");
                if (section.Header != null && section.Header.IsEnabled)
                {
                    pages.Append("<style:header>");
                    WriteBlocks(pages, section.Header.Blocks, styles);
                    pages.Append("</style:header>");
                }
                if (section.Footer != null && section.Footer.IsEnabled)
                {
                    pages.Append("<style:footer>");
                    WriteBlocks(pages, section.Footer.Blocks, styles);
                    pages.Append("</style:footer>");
                }
                pages.Append("</style:master-page>");
            }
            return pages;
        }

        /// <summary>输出一组块（用于页眉页脚），没有内容时补一个空段落。Writes a block list (for headers/footers), adding an empty paragraph when there is no content.</summary>
        private static void WriteBlocks(StringBuilder xml, List<Block> blocks, StyleContext styles)
        {
            bool wrote = false;
            for (int i = 0; i < blocks.Count; i++)
            {
                Block block = blocks[i];
                if (block is Paragraph || block is Table || block is PageBreak)
                {
                    WriteBlock(xml, block, styles, null, false);
                    wrote = true;
                }
            }
            if (!wrote) xml.Append("<text:p/>");
        }
    }
}
