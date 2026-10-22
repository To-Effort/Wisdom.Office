using System;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 将统一文档模型写成自包含的 HTML 文档：正文、标题、表格、图片（Base64 内嵌）、超链接、
    /// 书签、分页符、页眉页脚与页面尺寸都会输出为 HTML/CSS，不引用任何外部资源。
    /// Writes the unified document model as a self-contained HTML document: body text, headings, tables,
    /// images (embedded as Base64), hyperlinks, bookmarks, page breaks, headers/footers and page size are
    /// all emitted as HTML/CSS without referencing any external resource.
    /// </summary>
    public static class HtmlWriter
    {
        /// <summary>将文档保存为 UTF-8 HTML 文件。Saves the document as a UTF-8 HTML file.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <param name="path">目标文件路径。The target file path.</param>
        public static void Write(OfficeDocument document, string path)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (path == null) throw new ArgumentNullException("path");
            File.WriteAllBytes(path, Write(document));
        }

        /// <summary>将文档写成 UTF-8 HTML 字节并写入流。Writes the document as UTF-8 HTML bytes into a stream.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <param name="stream">目标流。The target stream.</param>
        public static void Write(OfficeDocument document, Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            byte[] bytes = Write(document);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>将文档写成 UTF-8 HTML 字节数组。Writes the document as a UTF-8 HTML byte array.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <returns>HTML 文件字节（不含 BOM）。The HTML file bytes (without a BOM).</returns>
        public static byte[] Write(OfficeDocument document)
        {
            return new UTF8Encoding(false).GetBytes(ToHtml(document));
        }

        /// <summary>将文档渲染为 HTML 文本。Renders the document as HTML text.</summary>
        /// <param name="document">待输出的文档。The document to write.</param>
        /// <returns>完整的 HTML 文档文本。The complete HTML document text.</returns>
        public static string ToHtml(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            StringBuilder builder = new StringBuilder(16384);
            builder.Append("<!DOCTYPE html>\n<html lang=\"zh-CN\">\n<head>\n");
            builder.Append("<meta charset=\"utf-8\" />\n");
            builder.Append("<meta name=\"generator\" content=\"Wisdom.Office\" />\n");
            DocumentMetadata metadata = document.Metadata ?? new DocumentMetadata();
            if (!string.IsNullOrEmpty(metadata.Title))
            {
                builder.Append("<title>").Append(WordFormatUtil.EscapeHtml(metadata.Title)).Append("</title>\n");
            }
            else
            {
                builder.Append("<title>Wisdom.Office Document</title>\n");
            }
            AppendMeta(builder, "author", metadata.Author);
            AppendMeta(builder, "description", metadata.Subject);
            AppendMeta(builder, "keywords", metadata.Keywords);
            if (metadata.Created != default(DateTime))
            {
                AppendMeta(builder, "dcterms.created", metadata.Created.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture));
            }
            if (metadata.Modified != default(DateTime))
            {
                AppendMeta(builder, "dcterms.modified", metadata.Modified.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture));
            }

            builder.Append("<style>\n");
            AppendPageCss(builder, document);
            builder.Append("body{margin:0;padding:0;background:#ffffff;color:#000000;}\n");
            builder.Append(".wisdom-document{margin:0 auto;}\n");
            builder.Append(".wisdom-section{margin:0 auto;}\n");
            builder.Append(".wisdom-header{font-size:9pt;color:#555555;border-bottom:0.5pt solid #cccccc;padding-bottom:4pt;margin-bottom:10pt;}\n");
            builder.Append(".wisdom-footer{font-size:9pt;color:#555555;border-top:0.5pt solid #cccccc;padding-top:4pt;margin-top:10pt;}\n");
            builder.Append("table.wisdom-table{border-collapse:collapse;table-layout:auto;}\n");
            builder.Append("table.wisdom-table td,table.wisdom-table th{padding:3pt 4pt;}\n");
            builder.Append(".wisdom-list-marker{margin-right:0.35em;}\n");
            builder.Append(".wisdom-pagebreak{page-break-after:always;break-after:page;height:0;margin:0;border:0;}\n");
            builder.Append("</style>\n</head>\n<body>\n<div class=\"wisdom-document\">\n");

            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                string className = "wisdom-section wisdom-section-" + (s + 1).ToString(CultureInfo.InvariantCulture);
                builder.Append("<section class=\"").Append(className).Append('"');
                AppendSectionAttributes(builder, section.PageSetup);
                if (s > 0)
                {
                    builder.Append(" style=\"page-break-before:always;break-before:page;\"");
                }
                builder.Append(">\n");
                if (section.Header != null && section.Header.IsEnabled)
                {
                    builder.Append("<header class=\"wisdom-header\">\n");
                    WriteBlocks(builder, section.Header.Blocks, true, 0);
                    builder.Append("</header>\n");
                }
                WriteBlocks(builder, section.Blocks, false, 0);
                if (section.Footer != null && section.Footer.IsEnabled)
                {
                    builder.Append("<footer class=\"wisdom-footer\">\n");
                    WriteBlocks(builder, section.Footer.Blocks, true, 0);
                    builder.Append("</footer>\n");
                }
                builder.Append("</section>\n");
            }

            builder.Append("</div>\n</body>\n</html>\n");
            return builder.ToString();
        }

        /// <summary>追加一个可选的 meta 标签；值为空时跳过。Appends an optional meta tag, skipped when the value is empty.</summary>
        private static void AppendMeta(StringBuilder builder, string name, string value)
        {
            if (string.IsNullOrEmpty(value)) return;
            builder.Append("<meta name=\"").Append(name).Append("\" content=\"")
                .Append(WordFormatUtil.EscapeHtml(value)).Append("\" />\n");
        }

        /// <summary>输出各节的 @page 规则与节专用页面名，供打印时按节设置纸张和页边距。Emits the @page rules and per-section page names so printing honours each section's paper size and margins.</summary>
        private static void AppendPageCss(StringBuilder builder, OfficeDocument document)
        {
            if (document.Sections.Count == 0)
            {
                builder.Append("@page{size:595.28pt 841.89pt;margin:72pt;}\n");
                return;
            }
            PageSetup first = document.Sections[0].PageSetup ?? new PageSetup();
            builder.Append("@page{size:").Append(WordFormatUtil.Number(first.Width)).Append("pt ")
                .Append(WordFormatUtil.Number(first.Height)).Append("pt;margin:")
                .Append(WordFormatUtil.Number(first.MarginTop)).Append("pt ")
                .Append(WordFormatUtil.Number(first.MarginRight)).Append("pt ")
                .Append(WordFormatUtil.Number(first.MarginBottom)).Append("pt ")
                .Append(WordFormatUtil.Number(first.MarginLeft)).Append("pt;}\n");
            for (int s = 0; s < document.Sections.Count; s++)
            {
                PageSetup setup = document.Sections[s].PageSetup ?? new PageSetup();
                int number = s + 1;
                builder.Append("@page wisdompage").Append(number.ToString(CultureInfo.InvariantCulture)).Append("{size:")
                    .Append(WordFormatUtil.Number(setup.Width)).Append("pt ").Append(WordFormatUtil.Number(setup.Height))
                    .Append("pt;margin:").Append(WordFormatUtil.Number(setup.MarginTop)).Append("pt ")
                    .Append(WordFormatUtil.Number(setup.MarginRight)).Append("pt ")
                    .Append(WordFormatUtil.Number(setup.MarginBottom)).Append("pt ")
                    .Append(WordFormatUtil.Number(setup.MarginLeft)).Append("pt;}\n");
                builder.Append(".wisdom-section-").Append(number.ToString(CultureInfo.InvariantCulture))
                    .Append("{page:wisdompage").Append(number.ToString(CultureInfo.InvariantCulture)).Append(";")
                    .Append("max-width:").Append(WordFormatUtil.Number(setup.Width - setup.MarginLeft - setup.MarginRight))
                    .Append("pt;}\n");
            }
        }

        /// <summary>在节元素上输出页面尺寸与页边距数据属性，读取器据此恢复 PageSetup。Writes page size and margin data attributes on a section element so the reader can restore the PageSetup.</summary>
        private static void AppendSectionAttributes(StringBuilder builder, PageSetup setup)
        {
            if (setup == null) return;
            builder.Append(" data-page-width=\"").Append(WordFormatUtil.Number(setup.Width)).Append('"')
                .Append(" data-page-height=\"").Append(WordFormatUtil.Number(setup.Height)).Append('"')
                .Append(" data-margin-left=\"").Append(WordFormatUtil.Number(setup.MarginLeft)).Append('"')
                .Append(" data-margin-top=\"").Append(WordFormatUtil.Number(setup.MarginTop)).Append('"')
                .Append(" data-margin-right=\"").Append(WordFormatUtil.Number(setup.MarginRight)).Append('"')
                .Append(" data-margin-bottom=\"").Append(WordFormatUtil.Number(setup.MarginBottom)).Append('"')
                .Append(" data-header-distance=\"").Append(WordFormatUtil.Number(setup.HeaderDistance)).Append('"')
                .Append(" data-footer-distance=\"").Append(WordFormatUtil.Number(setup.FooterDistance)).Append('"');
        }

        /// <summary>输出块级内容；compact 为 true 时段落使用零外边距（用于表格单元格与页眉页脚）。Writes block content; when compact is true, paragraphs use zero margins (for table cells and headers/footers).</summary>
        private static void WriteBlocks(StringBuilder builder, System.Collections.Generic.List<Block> blocks, bool compact, int depth)
        {
            if (blocks == null || depth > 16) return;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    WriteParagraph(builder, paragraph, compact);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    WriteTable(builder, table, depth + 1);
                    continue;
                }
                if (blocks[i] is PageBreak)
                {
                    builder.Append("<div class=\"wisdom-pagebreak\"></div>\n");
                }
            }
        }

        /// <summary>输出单个段落；标题级别映射到 h1~h6，其余使用 p。Writes one paragraph; heading levels map to h1-h6 and everything else uses p.</summary>
        private static void WriteParagraph(StringBuilder builder, Paragraph paragraph, bool compact)
        {
            int level = paragraph.HeadingLevel;
            string tag = level >= 1 && level <= 6 ? "h" + level.ToString(CultureInfo.InvariantCulture) : "p";
            builder.Append('<').Append(tag);
            StringBuilder style = new StringBuilder();
            if (compact) style.Append("margin:0;");
            else
            {
                style.Append("margin-top:").Append(WordFormatUtil.Number(paragraph.SpaceBefore)).Append("pt;");
                style.Append("margin-bottom:").Append(WordFormatUtil.Number(paragraph.SpaceAfter)).Append("pt;");
            }
            if (paragraph.Alignment == ParagraphAlignment.Center) style.Append("text-align:center;");
            else if (paragraph.Alignment == ParagraphAlignment.Right) style.Append("text-align:right;");
            else if (paragraph.Alignment == ParagraphAlignment.Justify) style.Append("text-align:justify;");
            if (paragraph.LineSpacing > 0 && Math.Abs(paragraph.LineSpacing - 1.0) > 0.001)
            {
                style.Append("line-height:").Append(WordFormatUtil.Number(paragraph.LineSpacing)).Append(';');
            }
            if (paragraph.LeftIndent > 0) style.Append("margin-left:").Append(WordFormatUtil.Number(paragraph.LeftIndent)).Append("pt;");
            if (paragraph.RightIndent > 0) style.Append("margin-right:").Append(WordFormatUtil.Number(paragraph.RightIndent)).Append("pt;");
            if (paragraph.FirstLineIndent != 0) style.Append("text-indent:").Append(WordFormatUtil.Number(paragraph.FirstLineIndent)).Append("pt;");
            if (paragraph.PageBreakBefore) style.Append("page-break-before:always;break-before:page;");
            if (paragraph.KeepWithNext) style.Append("page-break-after:avoid;break-after:avoid;");
            if (style.Length > 0) builder.Append(" style=\"").Append(style).Append('"');
            if (!string.IsNullOrEmpty(paragraph.StyleName))
            {
                builder.Append(" data-style=\"").Append(WordFormatUtil.EscapeHtml(paragraph.StyleName)).Append('"');
            }
            builder.Append('>');

            if (!string.IsNullOrEmpty(paragraph.BookmarkName) && !HasStartMarker(paragraph, paragraph.BookmarkName))
            {
                AppendAnchor(builder, paragraph.BookmarkName);
            }
            if (!string.IsNullOrEmpty(paragraph.ListMarker))
            {
                builder.Append("<span class=\"wisdom-list-marker\">")
                    .Append(WordFormatUtil.EscapeHtml(paragraph.ListMarker)).Append("</span>");
            }
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                WriteInline(builder, paragraph.Inlines[i]);
            }
            builder.Append("</").Append(tag).Append(">\n");
        }

        /// <summary>输出一个空锚点作为书签位置；HTML 只有点锚，范围终点不产生标记。Writes an empty anchor as the bookmark position; HTML has point anchors only, so a range end emits nothing.</summary>
        /// <param name="builder">输出构建器。The output builder.</param>
        /// <param name="name">书签名。The bookmark name.</param>
        private static void AppendAnchor(StringBuilder builder, string name)
        {
            string escaped = WordFormatUtil.EscapeHtml(name);
            builder.Append("<a id=\"").Append(escaped).Append("\" name=\"").Append(escaped).Append("\"></a>");
        }

        /// <summary>段落内是否已有该名称的显式书签起点。Determines whether the paragraph holds an explicit start marker of the name.</summary>
        /// <param name="paragraph">目标段落。The paragraph.</param>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <returns>存在返回 true。True when present.</returns>
        private static bool HasStartMarker(Paragraph paragraph, string name)
        {
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                BookmarkStart start = paragraph.Inlines[i] as BookmarkStart;
                if (start != null && string.Equals(start.Name, name, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        /// <summary>输出一个行内元素：文本运行附带字体与装饰 CSS，域输出显示值，图片以 Base64 内嵌。Writes one inline element: text runs carry font and decoration CSS, fields output their display value, and images are embedded as Base64.</summary>
        private static void WriteInline(StringBuilder builder, Inline inline)
        {
            BookmarkStart startMarker = inline as BookmarkStart;
            if (startMarker != null)
            {
                AppendAnchor(builder, startMarker.Name);
                return;
            }
            if (inline is BookmarkEnd) return;

            TextRun run = inline as TextRun;
            if (run != null)
            {
                string css = BuildRunCss(run.Format);
                bool hasLink = !string.IsNullOrEmpty(run.Link) && WordFormatUtil.IsSafeHyperlinkTarget(run.Link.Trim())
                    || (!string.IsNullOrEmpty(run.Link) && run.Link.Trim().StartsWith("#", StringComparison.Ordinal));
                if (hasLink)
                {
                    builder.Append("<a href=\"").Append(WordFormatUtil.EscapeHtml(run.Link.Trim()))
                        .Append("\" target=\"_blank\" rel=\"noopener noreferrer\"");
                    if (css.Length > 0) builder.Append(" style=\"").Append(css).Append('"');
                    builder.Append('>');
                }
                else if (css.Length > 0)
                {
                    builder.Append("<span style=\"").Append(css).Append("\">");
                }
                AppendTextWithBreaks(builder, run.Text);
                if (hasLink) builder.Append("</a>");
                else if (css.Length > 0) builder.Append("</span>");
                return;
            }

            FieldRun field = inline as FieldRun;
            if (field != null)
            {
                string fieldCss = BuildRunCss(field.Format);
                builder.Append("<span data-field=\"").Append(WordFormatUtil.EscapeHtml(field.Instruction)).Append('"');
                if (fieldCss.Length > 0) builder.Append(" style=\"").Append(fieldCss).Append('"');
                builder.Append('>').Append(WordFormatUtil.EscapeHtml(field.DisplayText)).Append("</span>");
                return;
            }

            ImageRun image = inline as ImageRun;
            if (image != null && image.Image != null)
            {
                if (image.Image.Format != OfficeImageFormat.Png && image.Image.Format != OfficeImageFormat.Jpeg) return;
                double width = image.Width > 0 ? image.Width : image.Image.DefaultWidthPoints;
                double height = image.Height > 0 ? image.Height : image.Image.DefaultHeightPoints;
                builder.Append("<img alt=\"\" src=\"data:").Append(image.Image.ContentType).Append(";base64,")
                    .Append(Convert.ToBase64String(image.Image.Data)).Append("\" style=\"width:")
                    .Append(WordFormatUtil.Number(width)).Append("pt;height:").Append(WordFormatUtil.Number(height)).Append("pt;\" />");
            }
        }

        /// <summary>把文本中的换行输出为 &lt;br /&gt;，其余字符按 HTML 转义。Emits newlines as &lt;br /&gt; and escapes the remaining characters for HTML.</summary>
        private static void AppendTextWithBreaks(StringBuilder builder, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            int start = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] != '\n') continue;
                if (i > start) builder.Append(WordFormatUtil.EscapeHtml(text.Substring(start, i - start)));
                builder.Append("<br />");
                start = i + 1;
            }
            if (start < text.Length) builder.Append(WordFormatUtil.EscapeHtml(text.Substring(start)));
        }

        /// <summary>根据文本格式生成内联 CSS（字体、字号、颜色、粗斜体、下划线、删除线与上下标）。Builds inline CSS from a text format (font, size, colour, bold/italic, underline, strike-through and super/subscript).</summary>
        private static string BuildRunCss(TextFormat format)
        {
            if (format == null) return string.Empty;
            StringBuilder css = new StringBuilder();
            string ascii = format.GetAsciiFontFamily();
            string eastAsia = format.GetEastAsiaFontFamily();
            if (!string.IsNullOrEmpty(ascii) || !string.IsNullOrEmpty(eastAsia))
            {
                css.Append("font-family:");
                AppendFontFamily(css, ascii);
                if (!string.IsNullOrEmpty(eastAsia)
                    && !string.Equals(eastAsia, ascii, StringComparison.OrdinalIgnoreCase))
                {
                    css.Append(',');
                    AppendFontFamily(css, eastAsia);
                }
                css.Append(",sans-serif;");
            }
            if (format.FontSize > 0)
            {
                css.Append("font-size:").Append(WordFormatUtil.Number(format.FontSize)).Append("pt;");
            }
            if (format.Bold) css.Append("font-weight:bold;");
            if (format.Italic) css.Append("font-style:italic;");
            if (format.Underline || format.Strikeout)
            {
                css.Append("text-decoration:");
                if (format.Underline) css.Append("underline");
                if (format.Underline && format.Strikeout) css.Append(' ');
                if (format.Strikeout) css.Append("line-through");
                css.Append(';');
            }
            if (format.Color.A < 255)
            {
                css.Append("color:rgba(").Append(format.Color.R.ToString(CultureInfo.InvariantCulture)).Append(',')
                    .Append(format.Color.G.ToString(CultureInfo.InvariantCulture)).Append(',')
                    .Append(format.Color.B.ToString(CultureInfo.InvariantCulture)).Append(',')
                    .Append(WordFormatUtil.Number(format.Color.A / 255.0)).Append(");");
            }
            else
            {
                css.Append("color:#").Append(format.Color.ToHex()).Append(';');
            }
            if (format.Superscript) css.Append("vertical-align:super;font-size:smaller;");
            else if (format.Subscript) css.Append("vertical-align:sub;font-size:smaller;");
            return css.ToString();
        }

        /// <summary>输出一个用单引号包裹并转义的 CSS 字体名。Writes a CSS font family wrapped in single quotes and escaped.</summary>
        private static void AppendFontFamily(StringBuilder css, string family)
        {
            if (string.IsNullOrEmpty(family)) return;
            css.Append('\'').Append(family.Replace("\\", "\\\\").Replace("'", "\\'")).Append('\'');
        }

        /// <summary>输出表格：列宽、外框与内线、跨行跨列、单元格底色与纵向对齐。Writes a table: column widths, outer and inner borders, spans, cell shading and vertical alignment.</summary>
        private static void WriteTable(StringBuilder builder, Table table, int depth)
        {
            if (depth > 16) return;
            builder.Append("<table class=\"wisdom-table\" style=\"border-collapse:collapse;");
            if (table.Alignment == TableAlignment.Center) builder.Append("margin-left:auto;margin-right:auto;");
            else if (table.Alignment == TableAlignment.Right) builder.Append("margin-left:auto;");
            builder.Append("\">\n");
            if (table.Columns.Count > 0)
            {
                builder.Append("<colgroup>");
                for (int c = 0; c < table.Columns.Count; c++)
                {
                    double width = table.Columns[c] != null ? table.Columns[c].Width : 0;
                    builder.Append("<col");
                    if (width > 0) builder.Append(" style=\"width:").Append(WordFormatUtil.Number(width)).Append("pt\"");
                    builder.Append(" />");
                }
                builder.Append("</colgroup>\n");
            }
            for (int r = 0; r < table.Rows.Count; r++)
            {
                TableRow row = table.Rows[r];
                builder.Append("<tr");
                if (row.IsHeader) builder.Append(" data-header=\"true\"");
                if (row.Height > 0) builder.Append(" style=\"height:").Append(WordFormatUtil.Number(row.Height)).Append("pt\"");
                builder.Append('>');
                for (int c = 0; c < row.Cells.Count; c++)
                {
                    TableCell cell = row.Cells[c];
                    bool header = row.IsHeader;
                    builder.Append(header ? "<th" : "<td");
                    if (cell.ColSpan > 1) builder.Append(" colspan=\"").Append(cell.ColSpan.ToString(CultureInfo.InvariantCulture)).Append('"');
                    if (cell.RowSpan > 1) builder.Append(" rowspan=\"").Append(cell.RowSpan.ToString(CultureInfo.InvariantCulture)).Append('"');
                    AppendCellStyle(builder, table, cell, r, c);
                    builder.Append('>');
                    WriteBlocks(builder, cell.Blocks, true, depth + 1);
                    builder.Append(header ? "</th>" : "</td>");
                }
                builder.Append("</tr>\n");
            }
            builder.Append("</table>\n");
        }

        /// <summary>输出单元格内联样式：四边边框（按表格外框/内线与单元格覆盖计算）、内边距、底色与纵向对齐。Writes the cell's inline style: four borders (derived from table outer/inner lines and cell overrides), padding, shading and vertical alignment.</summary>
        private static void AppendCellStyle(StringBuilder builder, Table table, TableCell cell, int row, int column)
        {
            builder.Append(" style=\"");
            int rowCount = table.Rows.Count;
            int columnCount = table.Columns.Count;
            for (int i = 0; columnCount == 0 && i < table.Rows.Count; i++)
            {
                columnCount = Math.Max(columnCount, table.Rows[i].Cells.Count);
            }
            if (columnCount == 0) columnCount = 1;
            BorderLine top = cell.Borders != null && cell.Borders.Top != null ? cell.Borders.Top
                : (row == 0 ? table.Borders.Top : table.Borders.InsideHorizontal);
            BorderLine bottom = cell.Borders != null && cell.Borders.Bottom != null ? cell.Borders.Bottom
                : (row + Math.Max(1, cell.RowSpan) >= rowCount ? table.Borders.Bottom : table.Borders.InsideHorizontal);
            BorderLine left = cell.Borders != null && cell.Borders.Left != null ? cell.Borders.Left
                : (column == 0 ? table.Borders.Left : table.Borders.InsideVertical);
            BorderLine right = cell.Borders != null && cell.Borders.Right != null ? cell.Borders.Right
                : (column + Math.Max(1, cell.ColSpan) >= columnCount ? table.Borders.Right : table.Borders.InsideVertical);
            AppendBorderCss(builder, "border-top", top);
            AppendBorderCss(builder, "border-bottom", bottom);
            AppendBorderCss(builder, "border-left", left);
            AppendBorderCss(builder, "border-right", right);
            if (table.CellPadding > 0)
            {
                builder.Append("padding:").Append(WordFormatUtil.Number(table.CellPadding)).Append("pt;");
            }
            if (cell.BackgroundColor.HasValue)
            {
                OfficeColor background = cell.BackgroundColor.Value;
                builder.Append("background-color:#").Append(background.ToHex()).Append(';');
            }
            if (cell.VerticalAlignment == VerticalAlignment.Middle) builder.Append("vertical-align:middle;");
            else if (cell.VerticalAlignment == VerticalAlignment.Bottom) builder.Append("vertical-align:bottom;");
            else builder.Append("vertical-align:top;");
            builder.Append('"');
        }

        /// <summary>输出单条边框 CSS；无线条时写 none。Writes one border CSS declaration; a missing line becomes none.</summary>
        private static void AppendBorderCss(StringBuilder builder, string property, BorderLine line)
        {
            builder.Append(property).Append(':');
            if (line == null || line.Style == BorderStyle.None)
            {
                builder.Append("none;");
                return;
            }
            builder.Append(WordFormatUtil.Number(line.Width > 0 ? line.Width : 0.5)).Append("pt ");
            switch (line.Style)
            {
                case BorderStyle.Dashed: builder.Append("dashed "); break;
                case BorderStyle.Dotted: builder.Append("dotted "); break;
                case BorderStyle.Double: builder.Append("double "); break;
                default: builder.Append("solid "); break;
            }
            builder.Append('#').Append(line.Color.ToHex()).Append(';');
        }
    }
}
