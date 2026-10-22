using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 将统一文档模型写成 RTF 1.9：输出字体表、颜色表、页面设置、页眉页脚、段落与字符格式、
    /// 表格（含跨行跨列与底色）、超链接域、书签与内嵌 PNG/JPEG 图片。
    /// Writes the unified document model as RTF 1.9: font and colour tables, page setup, headers/footers,
    /// paragraph and character formatting, tables (with spans and shading), hyperlink fields, bookmarks
    /// and embedded PNG/JPEG images.
    /// </summary>
    public static class RtfWriter
    {
        /// <summary>写入过程中的字体与颜色注册表。Font and colour registry used while writing.</summary>
        private sealed class WriterState
        {
            /// <summary>字体名列表，索引即 \fN。The font names; the index is the \fN value.</summary>
            public List<string> Fonts = new List<string>();
            /// <summary>字体名到索引的映射（忽略大小写）。Font name to index map (case-insensitive).</summary>
            public Dictionary<string, int> FontIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            /// <summary>颜色列表，索引 0 为自动色，从 1 开始是实际颜色。The colours; index 0 is automatic and real colours start at 1.</summary>
            public List<OfficeColor> Colors = new List<OfficeColor>();
            /// <summary>颜色到索引的映射。Colour to index map.</summary>
            public Dictionary<int, int> ColorIndex = new Dictionary<int, int>();
            /// <summary>当前页设置，供表格宽度计算。The current page setup, used for table width calculation.</summary>
            public PageSetup Page = new PageSetup();

            /// <summary>注册字体并返回索引。Registers a font and returns its index.</summary>
            public int AddFont(string name)
            {
                if (string.IsNullOrEmpty(name)) name = "Arial";
                int index;
                if (FontIndex.TryGetValue(name, out index)) return index;
                index = Fonts.Count;
                Fonts.Add(name);
                FontIndex[name] = index;
                return index;
            }

            /// <summary>注册颜色并返回颜色表索引。Registers a colour and returns its colour-table index.</summary>
            public int AddColor(OfficeColor color)
            {
                int key = (color.A << 24) | (color.R << 16) | (color.G << 8) | color.B;
                int index;
                if (ColorIndex.TryGetValue(key, out index)) return index;
                index = Colors.Count + 1;
                Colors.Add(color);
                ColorIndex[key] = index;
                return index;
            }
        }

        /// <summary>将文档保存为 RTF 文件。Saves the document as an RTF file.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <param name="path">目标文件路径。The target file path.</param>
        public static void Write(OfficeDocument document, string path)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (path == null) throw new ArgumentNullException("path");
            File.WriteAllBytes(path, Write(document));
        }

        /// <summary>将文档写成 RTF 字节并写入流。Writes the document as RTF bytes into a stream.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <param name="stream">目标流。The target stream.</param>
        public static void Write(OfficeDocument document, Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            byte[] bytes = Write(document);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>将文档写成 RTF 字节数组（ASCII 编码，非 ASCII 字符使用 \uN 转义）。Writes the document as an RTF byte array (ASCII encoding, non-ASCII characters escaped as \uN).</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <returns>RTF 文件字节。The RTF file bytes.</returns>
        public static byte[] Write(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            return Encoding.ASCII.GetBytes(ToRtf(document));
        }

        /// <summary>将文档渲染为 RTF 文本。Renders the document as RTF text.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <returns>完整的 RTF 文本。The complete RTF text.</returns>
        public static string ToRtf(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            WriterState state = new WriterState();
            if (document.DefaultTextFormat != null)
            {
                state.AddFont(document.DefaultTextFormat.GetAsciiFontFamily());
            }
            state.AddFont("Arial");
            CollectResources(document, state);

            StringBuilder body = new StringBuilder(16384);
            WriteSections(document, body, state);

            StringBuilder rtf = new StringBuilder(body.Length + 2048);
            rtf.Append("{\\rtf1\\ansi\\ansicpg65001\\uc1\\deff0\n");
            AppendFontTable(rtf, state);
            AppendColorTable(rtf, state);
            AppendInfo(rtf, document.Metadata);
            AppendPageSetup(rtf, document.Sections.Count > 0 ? document.Sections[0].PageSetup : new PageSetup());
            rtf.Append(body);
            rtf.Append("}\n");
            return rtf.ToString();
        }

        /// <summary>预扫描全文，收集字体与颜色，保证字体表/颜色表写在正文之前。Pre-scans the document for fonts and colours so the tables can precede the body.</summary>
        private static void CollectResources(OfficeDocument document, WriterState state)
        {
            if (document.DefaultTextFormat != null)
            {
                state.AddColor(document.DefaultTextFormat.Color);
            }
            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                CollectBlocks(section.Blocks, state, 0);
                CollectBlocks(section.Header.Blocks, state, 0);
                CollectBlocks(section.Footer.Blocks, state, 0);
            }
        }

        /// <summary>递归收集块中的字体、文字颜色、表格底色与边框颜色。Recursively collects fonts, text colours, table shading and border colours from blocks.</summary>
        private static void CollectBlocks(List<Block> blocks, WriterState state, int depth)
        {
            if (blocks == null || depth > 16) return;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    for (int r = 0; r < paragraph.Inlines.Count; r++)
                    {
                        TextRun run = paragraph.Inlines[r] as TextRun;
                        if (run != null)
                        {
                            state.AddFont(SelectFont(run.Format, run.Text));
                            state.AddColor(run.Format.Color);
                            continue;
                        }
                        FieldRun field = paragraph.Inlines[r] as FieldRun;
                        if (field != null)
                        {
                            state.AddFont(SelectFont(field.Format, field.DisplayText));
                            state.AddColor(field.Format.Color);
                        }
                    }
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                CollectBorderColors(table.Borders, state);
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        TableCell cell = table.Rows[r].Cells[c];
                        if (cell.BackgroundColor.HasValue) state.AddColor(cell.BackgroundColor.Value);
                        if (cell.Borders != null)
                        {
                            if (cell.Borders.Top != null) state.AddColor(cell.Borders.Top.Color);
                            if (cell.Borders.Bottom != null) state.AddColor(cell.Borders.Bottom.Color);
                            if (cell.Borders.Left != null) state.AddColor(cell.Borders.Left.Color);
                            if (cell.Borders.Right != null) state.AddColor(cell.Borders.Right.Color);
                        }
                        CollectBlocks(cell.Blocks, state, depth + 1);
                    }
                }
            }
        }

        /// <summary>收集表格六条边框的颜色。Collects the colours of a table's six borders.</summary>
        private static void CollectBorderColors(TableBorders borders, WriterState state)
        {
            if (borders == null) return;
            if (borders.Top != null) state.AddColor(borders.Top.Color);
            if (borders.Bottom != null) state.AddColor(borders.Bottom.Color);
            if (borders.Left != null) state.AddColor(borders.Left.Color);
            if (borders.Right != null) state.AddColor(borders.Right.Color);
            if (borders.InsideHorizontal != null) state.AddColor(borders.InsideHorizontal.Color);
            if (borders.InsideVertical != null) state.AddColor(borders.InsideVertical.Color);
        }

        /// <summary>按文本内容选择西文字体或东亚字体。Selects the Latin or East-Asian font based on the text content.</summary>
        private static string SelectFont(TextFormat format, string text)
        {
            if (format == null) return "Arial";
            return WordFormatUtil.ContainsEastAsian(text) ? format.GetEastAsiaFontFamily() : format.GetAsciiFontFamily();
        }

        /// <summary>输出字体表。Writes the font table.</summary>
        private static void AppendFontTable(StringBuilder rtf, WriterState state)
        {
            rtf.Append("{\\fonttbl");
            for (int i = 0; i < state.Fonts.Count; i++)
            {
                string name = state.Fonts[i];
                rtf.Append("{\\f").Append(i.ToString(CultureInfo.InvariantCulture)).Append("\\fnil");
                rtf.Append(WordFormatUtil.ContainsEastAsian(name) ? "\\fcharset134" : "\\fcharset0");
                rtf.Append(' ');
                AppendEscapedText(rtf, name);
                rtf.Append(";}");
            }
            rtf.Append("}\n");
        }

        /// <summary>输出颜色表；索引 0 为 RTF 的自动色占位。Writes the colour table; index 0 is the RTF automatic-colour placeholder.</summary>
        private static void AppendColorTable(StringBuilder rtf, WriterState state)
        {
            rtf.Append("{\\colortbl;");
            for (int i = 0; i < state.Colors.Count; i++)
            {
                OfficeColor color = state.Colors[i];
                rtf.Append("\\red").Append(color.R.ToString(CultureInfo.InvariantCulture))
                    .Append("\\green").Append(color.G.ToString(CultureInfo.InvariantCulture))
                    .Append("\\blue").Append(color.B.ToString(CultureInfo.InvariantCulture)).Append(';');
            }
            rtf.Append("}\n");
        }

        /// <summary>输出文档属性组（标题、作者、主题、关键词与创建/修改时间）。Writes the document property group (title, author, subject, keywords and created/modified times).</summary>
        private static void AppendInfo(StringBuilder rtf, DocumentMetadata metadata)
        {
            if (metadata == null) return;
            bool hasContent = !string.IsNullOrEmpty(metadata.Title) || !string.IsNullOrEmpty(metadata.Author)
                || !string.IsNullOrEmpty(metadata.Subject) || !string.IsNullOrEmpty(metadata.Keywords)
                || metadata.Created != default(DateTime) || metadata.Modified != default(DateTime);
            if (!hasContent) return;
            rtf.Append("{\\info");
            AppendInfoText(rtf, "title", metadata.Title);
            AppendInfoText(rtf, "author", metadata.Author);
            AppendInfoText(rtf, "subject", metadata.Subject);
            AppendInfoText(rtf, "keywords", metadata.Keywords);
            if (metadata.Created != default(DateTime)) AppendInfoDate(rtf, "creatim", metadata.Created);
            if (metadata.Modified != default(DateTime)) AppendInfoDate(rtf, "revtim", metadata.Modified);
            rtf.Append("}\n");
        }

        /// <summary>输出一个文本属性；空值跳过。Writes one text property, skipping empty values.</summary>
        private static void AppendInfoText(StringBuilder rtf, string name, string value)
        {
            if (string.IsNullOrEmpty(value)) return;
            rtf.Append("{\\").Append(name).Append(' ');
            AppendEscapedText(rtf, value);
            rtf.Append('}');
        }

        /// <summary>输出一个日期属性。Writes one date property.</summary>
        private static void AppendInfoDate(StringBuilder rtf, string name, DateTime value)
        {
            rtf.Append("{\\").Append(name)
                .Append("\\yr").Append(value.Year.ToString(CultureInfo.InvariantCulture))
                .Append("\\mo").Append(value.Month.ToString(CultureInfo.InvariantCulture))
                .Append("\\dy").Append(value.Day.ToString(CultureInfo.InvariantCulture))
                .Append("\\hr").Append(value.Hour.ToString(CultureInfo.InvariantCulture))
                .Append("\\min").Append(value.Minute.ToString(CultureInfo.InvariantCulture))
                .Append('}');
        }

        /// <summary>输出纸张尺寸、页边距与页眉页脚距离（单位为缇）。Writes paper size, margins and header/footer distances (in twips).</summary>
        private static void AppendPageSetup(StringBuilder rtf, PageSetup setup)
        {
            if (setup == null) return;
            rtf.Append("\\paperw").Append(WordFormatUtil.PointsToTwips(setup.Width))
                .Append("\\paperh").Append(WordFormatUtil.PointsToTwips(setup.Height))
                .Append("\\margl").Append(WordFormatUtil.PointsToTwips(setup.MarginLeft))
                .Append("\\margr").Append(WordFormatUtil.PointsToTwips(setup.MarginRight))
                .Append("\\margt").Append(WordFormatUtil.PointsToTwips(setup.MarginTop))
                .Append("\\margb").Append(WordFormatUtil.PointsToTwips(setup.MarginBottom))
                .Append("\\headery").Append(WordFormatUtil.PointsToTwips(Math.Max(0, setup.HeaderDistance)))
                .Append("\\footery").Append(WordFormatUtil.PointsToTwips(Math.Max(0, setup.FooterDistance)))
                .Append("\n");
        }

        /// <summary>输出全部节；节之间用 \sect 分隔并重设页面参数。Writes all sections, separated by \sect with page parameters re-applied.</summary>
        private static void WriteSections(OfficeDocument document, StringBuilder rtf, WriterState state)
        {
            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                state.Page = section.PageSetup ?? new PageSetup();
                if (s > 0)
                {
                    rtf.Append("\\sect\n");
                    AppendPageSetup(rtf, state.Page);
                }
                WriteHeaderFooter(rtf, section.Header, state, "header");
                WriteHeaderFooter(rtf, section.Footer, state, "footer");
                WriteBlocks(rtf, section.Blocks, state, 0, false, false);
            }
        }

        /// <summary>输出页眉或页脚组。Writes a header or footer group.</summary>
        private static void WriteHeaderFooter(StringBuilder rtf, HeaderFooter headerFooter, WriterState state, string destination)
        {
            if (headerFooter == null || !headerFooter.IsEnabled) return;
            rtf.Append("{\\").Append(destination);
            WriteBlocks(rtf, headerFooter.Blocks, state, 0, true, false);
            rtf.Append("}\n");
        }

        /// <summary>输出块级内容；inHeaderFooter 控制是否补充末尾段落，inTable 时段落标记为表内段落。Writes block content; inHeaderFooter controls whether a trailing paragraph is added, and inTable marks paragraphs as belonging to a table.</summary>
        private static void WriteBlocks(StringBuilder rtf, List<Block> blocks, WriterState state, int depth, bool inHeaderFooter, bool inTable)
        {
            if (blocks == null || depth > 16) return;
            bool wroteContent = false;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    WriteParagraph(rtf, paragraph, state, inTable);
                    wroteContent = true;
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    WriteTable(rtf, table, state, depth + 1);
                    wroteContent = true;
                    continue;
                }
                if (blocks[i] is PageBreak)
                {
                    rtf.Append("\\pard\\page\\par\n");
                    wroteContent = true;
                }
            }
            if (inHeaderFooter && !wroteContent)
            {
                rtf.Append("\\pard\\par\n");
            }
        }

        /// <summary>输出一个段落：对齐、间距、缩进、行距、分页/孤行控制、书签、列表标记与行内运行。Writes one paragraph: alignment, spacing, indents, line spacing, page-break/keep controls, bookmark, list marker and inline runs.</summary>
        private static void WriteParagraph(StringBuilder rtf, Paragraph paragraph, WriterState state, bool inTable)
        {
            if (paragraph.PageBreakBefore) rtf.Append("\\page\n");
            rtf.Append("\\pard");
            if (inTable) rtf.Append("\\intbl");
            switch (paragraph.Alignment)
            {
                case ParagraphAlignment.Center: rtf.Append("\\qc"); break;
                case ParagraphAlignment.Right: rtf.Append("\\qr"); break;
                case ParagraphAlignment.Justify: rtf.Append("\\qj"); break;
                default: rtf.Append("\\ql"); break;
            }
            rtf.Append("\\sb").Append(WordFormatUtil.PointsToTwips(paragraph.SpaceBefore))
                .Append("\\sa").Append(WordFormatUtil.PointsToTwips(paragraph.SpaceAfter))
                .Append("\\li").Append(WordFormatUtil.PointsToTwips(paragraph.LeftIndent))
                .Append("\\ri").Append(WordFormatUtil.PointsToTwips(paragraph.RightIndent))
                .Append("\\fi").Append(WordFormatUtil.PointsToTwips(paragraph.FirstLineIndent));
            if (paragraph.LineSpacing > 0)
            {
                int twips = (int)Math.Round(240 * paragraph.LineSpacing);
                if (twips < 1) twips = 240;
                rtf.Append("\\sl").Append(twips.ToString(CultureInfo.InvariantCulture)).Append("\\slmult1");
            }
            if (paragraph.HeadingLevel > 0 && paragraph.HeadingLevel <= 9)
            {
                rtf.Append("\\outlinelevel").Append((paragraph.HeadingLevel - 1).ToString(CultureInfo.InvariantCulture));
            }
            if (paragraph.KeepWithNext) rtf.Append("\\keepn");
            rtf.Append(' ');

            bool hasBookmark = !string.IsNullOrEmpty(paragraph.BookmarkName) && !HasStartMarker(paragraph, paragraph.BookmarkName);
            if (hasBookmark)
            {
                rtf.Append("{\\*\\bkmkstart ").Append(EscapeBookmark(paragraph.BookmarkName)).Append('}');
            }
            if (!string.IsNullOrEmpty(paragraph.ListMarker))
            {
                AppendEscapedText(rtf, paragraph.ListMarker);
                rtf.Append("\\tab ");
            }
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                WriteInline(rtf, paragraph.Inlines[i], state);
            }
            if (hasBookmark)
            {
                rtf.Append("{\\*\\bkmkend ").Append(EscapeBookmark(paragraph.BookmarkName)).Append('}');
            }
            rtf.Append("\\par\n");
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

        /// <summary>转义书签名，去掉空格与控制字符。Escapes a bookmark name, dropping spaces and control characters.</summary>
        private static string EscapeBookmark(string name)
        {
            StringBuilder builder = new StringBuilder(name.Length);
            for (int i = 0; i < name.Length; i++)
            {
                char value = name[i];
                if (char.IsLetterOrDigit(value) || value == '_' || value == '-' || value == '.') builder.Append(value);
            }
            return builder.Length == 0 ? "bookmark" : builder.ToString();
        }

        /// <summary>输出一个行内元素；带链接的文本写成 HYPERLINK 域，图片写成 \pict。Writes one inline element; linked text becomes a HYPERLINK field and images become \pict groups.</summary>
        private static void WriteInline(StringBuilder rtf, Inline inline, WriterState state)
        {
            BookmarkStart startMarker = inline as BookmarkStart;
            if (startMarker != null)
            {
                rtf.Append("{\\*\\bkmkstart ").Append(EscapeBookmark(startMarker.Name)).Append('}');
                return;
            }
            BookmarkEnd endMarker = inline as BookmarkEnd;
            if (endMarker != null)
            {
                rtf.Append("{\\*\\bkmkend ").Append(EscapeBookmark(endMarker.Name)).Append('}');
                return;
            }

            TextRun run = inline as TextRun;
            if (run != null)
            {
                bool hasLink = !string.IsNullOrEmpty(run.Link)
                    && (run.Link.Trim().StartsWith("#", StringComparison.Ordinal)
                        || WordFormatUtil.IsSafeHyperlinkTarget(run.Link.Trim()));
                if (hasLink)
                {
                    rtf.Append("{\\field{\\*\\fldinst{HYPERLINK \"")
                        .Append(EscapeFieldArgument(run.Link.Trim())).Append("\"}}{\\fldrslt");
                    WriteRun(rtf, run, state);
                    rtf.Append("}}");
                    return;
                }
                WriteRun(rtf, run, state);
                return;
            }

            FieldRun field = inline as FieldRun;
            if (field != null)
            {
                rtf.Append("{\\field{\\*\\fldinst{");
                AppendEscapedText(rtf, field.Instruction);
                rtf.Append("}}{\\fldrslt");
                WriteRunGroup(rtf, field.DisplayText, field.Format, state);
                rtf.Append("}}");
                return;
            }

            ImageRun image = inline as ImageRun;
            if (image != null && image.Image != null)
            {
                WriteImage(rtf, image, state);
            }
        }

        /// <summary>输出一个文本运行，按需包裹 HYPERLINK 域结果。Writes a text run, optionally nested in a HYPERLINK field result.</summary>
        private static void WriteRun(StringBuilder rtf, TextRun run, WriterState state)
        {
            WriteRunGroup(rtf, run.Text, run.Format, state);
        }

        /// <summary>输出一个带字符格式的 RTF 组；换行与制表符写成控制字。Writes an RTF group with character formatting; newlines and tabs become control words.</summary>
        private static void WriteRunGroup(StringBuilder rtf, string text, TextFormat format, WriterState state)
        {
            format = format ?? new TextFormat();
            rtf.Append('{');
            AppendRunProperties(rtf, format, text, state);
            AppendEscapedText(rtf, text);
            rtf.Append('}');
        }

        /// <summary>输出字符格式控制字：字体、字号、颜色、粗斜体、下划线、删除线与上下标。Writes character-format control words: font, size, colour, bold/italic, underline, strike-through and super/subscript.</summary>
        private static void AppendRunProperties(StringBuilder rtf, TextFormat format, string text, WriterState state)
        {
            string font = SelectFont(format, text);
            rtf.Append("\\f").Append(state.AddFont(font).ToString(CultureInfo.InvariantCulture));
            if (format.FontSize > 0)
            {
                int halfPoints = (int)Math.Round(format.FontSize * 2);
                if (halfPoints < 2) halfPoints = 2;
                rtf.Append("\\fs").Append(halfPoints.ToString(CultureInfo.InvariantCulture));
            }
            rtf.Append("\\cf").Append(state.AddColor(format.Color).ToString(CultureInfo.InvariantCulture));
            if (format.Bold) rtf.Append("\\b");
            if (format.Italic) rtf.Append("\\i");
            if (format.Underline) rtf.Append("\\ul");
            if (format.Strikeout) rtf.Append("\\strike");
            if (format.Superscript) rtf.Append("\\super");
            else if (format.Subscript) rtf.Append("\\sub");
            rtf.Append(' ');
        }

        /// <summary>输出内嵌图片组：PNG/JPEG 以十六进制数据写入，尺寸按目标磅值换算为缇。Writes an embedded image group: PNG/JPEG hex data with target sizes converted to twips.</summary>
        private static void WriteImage(StringBuilder rtf, ImageRun run, WriterState state)
        {
            OfficeImage image = run.Image;
            if (image.Format != OfficeImageFormat.Png && image.Format != OfficeImageFormat.Jpeg) return;
            double width = run.Width > 0 ? run.Width : image.DefaultWidthPoints;
            double height = run.Height > 0 ? run.Height : image.DefaultHeightPoints;
            int pixelWidth = image.PixelWidth > 0 ? image.PixelWidth : 96;
            int pixelHeight = image.PixelHeight > 0 ? image.PixelHeight : 96;
            rtf.Append("{\\pict\\").Append(image.Format == OfficeImageFormat.Png ? "pngblip" : "jpegblip")
                .Append("\\picw").Append(pixelWidth.ToString(CultureInfo.InvariantCulture))
                .Append("\\pich").Append(pixelHeight.ToString(CultureInfo.InvariantCulture))
                .Append("\\picwgoal").Append(WordFormatUtil.PointsToTwips(width))
                .Append("\\pichgoal").Append(WordFormatUtil.PointsToTwips(height)).Append('\n');
            AppendHex(rtf, image.Data);
            rtf.Append("}\n");
        }

        /// <summary>输出十六进制图片数据，每 128 字节换行（RTF 忽略换行）。Writes hex image data with a line break every 128 bytes (RTF ignores the breaks).</summary>
        private static void AppendHex(StringBuilder rtf, byte[] data)
        {
            for (int i = 0; i < data.Length; i++)
            {
                rtf.Append(data[i].ToString("x2", CultureInfo.InvariantCulture));
                if ((i & 127) == 127) rtf.Append('\n');
            }
        }

        /// <summary>输出表格：行定义（单元格边界、边框、底色、纵向对齐、横纵合并）与单元格内容。Writes a table: row definitions (cell boundaries, borders, shading, vertical alignment, merges) and cell content.</summary>
        private static void WriteTable(StringBuilder rtf, Table table, WriterState state, int depth)
        {
            if (depth > 16) return;
            int rowCount = table.Rows.Count;
            int columnCount = table.Columns.Count;
            for (int r = 0; columnCount == 0 && r < rowCount; r++)
            {
                columnCount = Math.Max(columnCount, EffectiveColumnCount(table.Rows[r]));
            }
            if (columnCount <= 0) columnCount = 1;
            double available = Math.Max(72, state.Page.Width - state.Page.MarginLeft - state.Page.MarginRight);
            double[] widths = ComputeColumnWidths(table, columnCount, available);
            TableCell[] verticalCells = new TableCell[columnCount + 2];
            int[] verticalStartRows = new int[columnCount + 2];
            int[] verticalStartColumns = new int[columnCount + 2];
            int[] verticalRemaining = new int[columnCount + 2];

            for (int r = 0; r < rowCount; r++)
            {
                TableRow row = table.Rows[r];
                if (row == null) continue;
                List<CellSlot> slots = new List<CellSlot>();
                int gridColumn = 1;
                int cellIndex = 0;
                while (gridColumn <= columnCount)
                {
                    if (verticalCells[gridColumn] != null)
                    {
                        TableCell active = verticalCells[gridColumn];
                        slots.Add(new CellSlot
                        {
                            Column = gridColumn,
                            Content = null,
                            Props = BuildCellProperties(table, active, r, gridColumn,
                                verticalStartColumns[gridColumn], verticalStartRows[gridColumn], state)
                        });
                        verticalRemaining[gridColumn]--;
                        if (verticalRemaining[gridColumn] <= 0)
                        {
                            verticalCells[gridColumn] = null;
                        }
                        gridColumn++;
                        continue;
                    }

                    TableCell cell = cellIndex < row.Cells.Count ? row.Cells[cellIndex] : null;
                    if (cell != null) cellIndex++;
                    int span = cell != null ? Math.Max(1, cell.ColSpan) : 1;
                    if (gridColumn + span - 1 > columnCount) span = columnCount - gridColumn + 1;
                    bool verticalStart = cell != null && cell.RowSpan > 1;
                    for (int k = 0; k < span; k++)
                    {
                        bool first = k == 0;
                        slots.Add(new CellSlot
                        {
                            Column = gridColumn + k,
                            Content = first ? cell : null,
                            Props = BuildCellProperties(table, cell, r, gridColumn + k, gridColumn, r, state)
                        });
                        if (verticalStart)
                        {
                            verticalCells[gridColumn + k] = cell;
                            verticalStartRows[gridColumn + k] = r;
                            verticalStartColumns[gridColumn + k] = gridColumn;
                            verticalRemaining[gridColumn + k] = Math.Max(1, cell.RowSpan) - 1;
                        }
                    }
                    gridColumn += span;
                }

                rtf.Append("\\trowd");
                if (row.IsHeader) rtf.Append("\\trhdr");
                if (row.Height > 0) rtf.Append("\\trrh").Append(WordFormatUtil.PointsToTwips(row.Height));
                rtf.Append("\\trgaph").Append(WordFormatUtil.PointsToTwips(Math.Max(0, table.CellPadding))).Append('\n');
                int boundary = 0;
                for (int i = 0; i < slots.Count; i++)
                {
                    boundary += WordFormatUtil.PointsToTwips(widths[Math.Min(slots[i].Column - 1, widths.Length - 1)]);
                    rtf.Append(slots[i].Props).Append("\\cellx").Append(boundary.ToString(CultureInfo.InvariantCulture)).Append('\n');
                }
                for (int i = 0; i < slots.Count; i++)
                {
                    if (slots[i].Content != null)
                    {
                        WriteBlocks(rtf, slots[i].Content.Blocks, state, depth + 1, false, true);
                    }
                    else
                    {
                        rtf.Append("\\pard\\intbl\\par");
                    }
                    rtf.Append("\\cell\n");
                }
                rtf.Append("\\row\n");
            }
        }

        /// <summary>表格行的有效列数（按横向合并展开）。The effective column count of a row after expanding horizontal spans.</summary>
        private static int EffectiveColumnCount(TableRow row)
        {
            int count = 0;
            for (int i = 0; i < row.Cells.Count; i++) count += Math.Max(1, row.Cells[i].ColSpan);
            return count;
        }

        /// <summary>计算列宽：优先使用显式列宽，其次等分可用宽度。Computes column widths: explicit widths first, otherwise equal division of the available width.</summary>
        private static double[] ComputeColumnWidths(Table table, int columnCount, double available)
        {
            double[] widths = new double[columnCount];
            double total = 0;
            for (int i = 0; i < table.Columns.Count && i < columnCount; i++)
            {
                widths[i] = Math.Max(0, table.Columns[i].Width);
                total += widths[i];
            }
            for (int i = 0; i < columnCount; i++)
            {
                if (widths[i] <= 0) widths[i] = 0;
            }
            if (total <= 0)
            {
                double each = available / columnCount;
                for (int i = 0; i < columnCount; i++) widths[i] = each;
            }
            else if (total < available)
            {
                double remaining = available - total;
                int unspecified = 0;
                for (int i = 0; i < columnCount; i++) if (widths[i] <= 0) unspecified++;
                if (unspecified > 0)
                {
                    for (int i = 0; i < columnCount; i++) if (widths[i] <= 0) widths[i] = remaining / unspecified;
                }
            }
            return widths;
        }

        /// <summary>表格单元格的行定义属性片段。The row-definition property fragment for one table cell.</summary>
        private sealed class CellSlot
        {
            /// <summary>网格列（从 1 开始）。The grid column (1-based).</summary>
            public int Column;
            /// <summary>单元格内容；合并占位与空单元格为 null。The cell content; null for merge placeholders and empty cells.</summary>
            public TableCell Content;
            /// <summary>\cellx 之前的属性文本。The property text preceding \cellx.</summary>
            public string Props;
        }

        /// <summary>
        /// 构造一个物理单元格的 RTF 行定义属性：合并标记、底色、纵向对齐与四边边框。
        /// 合并标记按单元格起始行列推导：首列为 \clmgf、后续列为 \clmrg；首行为 \clvmgf、后续行为 \clvmrg。
        /// Builds the RTF row-definition properties for one physical cell: merge markers, shading, vertical alignment and four borders.
        /// Merge markers derive from the cell's start row/column: the first column gets \clmgf and later ones \clmrg; the first row gets \clvmgf and later ones \clvmrg.
        /// </summary>
        private static string BuildCellProperties(Table table, TableCell cell, int row, int column, int startColumn, int startRow, WriterState state)
        {
            StringBuilder props = new StringBuilder();
            if (cell != null)
            {
                int colSpan = Math.Max(1, cell.ColSpan);
                int rowSpan = Math.Max(1, cell.RowSpan);
                if (colSpan > 1)
                {
                    props.Append(column == startColumn ? "\\clmgf" : "\\clmrg");
                }
                if (rowSpan > 1)
                {
                    props.Append(row == startRow ? "\\clvmgf" : "\\clvmrg");
                }
                if (cell.BackgroundColor.HasValue)
                {
                    props.Append("\\clcbpat").Append(state.AddColor(cell.BackgroundColor.Value).ToString(CultureInfo.InvariantCulture));
                }
                if (cell.VerticalAlignment == VerticalAlignment.Middle) props.Append("\\clvertalc");
                else if (cell.VerticalAlignment == VerticalAlignment.Bottom) props.Append("\\clvertalb");
                else props.Append("\\clvertalt");
                AppendCellBorders(props, table, cell, row, column, startColumn, startRow);
            }
            return props.ToString();
        }

        /// <summary>
        /// 输出单元格四边边框：只在合并块的真实外缘绘制，内部边界留空；外缘再按表格外框或内线及单元格覆盖决定。
        /// Writes the cell's four borders: only the true outer edges of a merged block are drawn, internal boundaries stay empty;
        /// outer edges then follow table outer/inner lines with cell overrides applied.
        /// </summary>
        private static void AppendCellBorders(StringBuilder props, Table table, TableCell cell, int row, int column, int startColumn, int startRow)
        {
            int colSpan = Math.Max(1, cell.ColSpan);
            int rowSpan = Math.Max(1, cell.RowSpan);
            int endColumn = startColumn + colSpan - 1;
            int endRow = startRow + rowSpan - 1;
            int rowCount = table.Rows.Count;
            int columnCount = table.Columns.Count;
            for (int c = 0; columnCount == 0 && c < rowCount; c++)
            {
                columnCount = Math.Max(columnCount, EffectiveColumnCount(table.Rows[c]));
            }
            if (columnCount <= 0) columnCount = 1;

            BorderLine top = row != startRow ? null
                : cell.Borders != null && cell.Borders.Top != null ? cell.Borders.Top
                : (startRow == 0 ? table.Borders.Top : table.Borders.InsideHorizontal);
            BorderLine bottom = row != endRow ? null
                : cell.Borders != null && cell.Borders.Bottom != null ? cell.Borders.Bottom
                : (endRow >= rowCount - 1 ? table.Borders.Bottom : table.Borders.InsideHorizontal);
            BorderLine left = column != startColumn ? null
                : cell.Borders != null && cell.Borders.Left != null ? cell.Borders.Left
                : (startColumn == 1 ? table.Borders.Left : table.Borders.InsideVertical);
            BorderLine right = column != endColumn ? null
                : cell.Borders != null && cell.Borders.Right != null ? cell.Borders.Right
                : (endColumn >= columnCount ? table.Borders.Right : table.Borders.InsideVertical);
            AppendBorder(props, "clbrdrt", top);
            AppendBorder(props, "clbrdrb", bottom);
            AppendBorder(props, "clbrdrl", left);
            AppendBorder(props, "clbrdrr", right);
        }

        /// <summary>输出单条 RTF 单元格边框控制字。Writes one RTF cell-border control sequence.</summary>
        private static void AppendBorder(StringBuilder props, string target, BorderLine line)
        {
            if (line == null || line.Style == BorderStyle.None)
            {
                props.Append('\\').Append(target).Append("\\brdrnil");
                return;
            }
            props.Append('\\').Append(target);
            switch (line.Style)
            {
                case BorderStyle.Dashed: props.Append("\\brdrdash"); break;
                case BorderStyle.Dotted: props.Append("\\brdrdot"); break;
                case BorderStyle.Double: props.Append("\\brdrdb"); break;
                default: props.Append("\\brdrs"); break;
            }
            props.Append("\\brdrw").Append(Math.Max(1, WordFormatUtil.PointsToTwips(line.Width)));
        }

        /// <summary>转义 RTF 正文：反斜杠、花括号、换行、制表符与非 ASCII 字符。Escapes RTF body text: backslashes, braces, newlines, tabs and non-ASCII characters.</summary>
        private static void AppendEscapedText(StringBuilder rtf, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            for (int i = 0; i < text.Length; i++)
            {
                char value = text[i];
                switch (value)
                {
                    case '\\': rtf.Append("\\\\"); break;
                    case '{': rtf.Append("\\{"); break;
                    case '}': rtf.Append("\\}"); break;
                    case '\r': break;
                    case '\n': rtf.Append("\\line "); break;
                    case '\t': rtf.Append("\\tab "); break;
                    default:
                        if (value < 32)
                        {
                            break;
                        }
                        if (value < 128)
                        {
                            rtf.Append(value);
                            break;
                        }
                        rtf.Append("\\u").Append(((int)value).ToString(CultureInfo.InvariantCulture)).Append('?');
                        break;
                }
            }
        }

        /// <summary>转义 HYPERLINK 域的带引号参数：反斜杠与花括号转义，双引号与控制字符写成百分号编码。Escapes the quoted HYPERLINK field argument: backslashes and braces are escaped, double quotes and control characters are percent-encoded.</summary>
        private static string EscapeFieldArgument(string value)
        {
            StringBuilder builder = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                char current = value[i];
                switch (current)
                {
                    case '\\': builder.Append("\\\\"); break;
                    case '{': builder.Append("\\{"); break;
                    case '}': builder.Append("\\}"); break;
                    case '"': builder.Append("%22"); break;
                    default:
                        if (current < 32)
                        {
                            builder.Append('%').Append(((int)current).ToString("x2", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            builder.Append(current);
                        }
                        break;
                }
            }
            return builder.ToString();
        }
    }
}
