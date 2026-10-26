using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Markdown
{
    /// <summary>
    /// 将统一文档模型渲染为 HTML 正文片段的内部写入器（不含 html/head 外壳）。
    /// 所有文本一律做 HTML 转义；图片刻意内联为 base64 data URI，保证片段脱离任何资源目录也能离线可读。
    /// Internal writer that renders the unified document model as an HTML body fragment (no html/head shell).
    /// All text is HTML-escaped; images are deliberately inlined as base64 data URIs so the fragment stays readable offline.
    /// </summary>
    internal static class MarkdownHtmlWriter
    {
        /// <summary>每层列表嵌套对应的左缩进磅数。The left indent, in points, that corresponds to one level of list nesting.</summary>
        private const double IndentPerLevel = 24.0;

        /// <summary>
        /// 渲染入口：段落转 p、标题转 h1~h6、列表转嵌套 ul/ol、连续等宽段落转 pre/code、表格转 table、图片转内联 data URI 的 img。
        /// Entry point: paragraphs become p, headings h1-h6, list paragraphs nested ul/ol, consecutive monospaced paragraphs pre/code, tables table, images img with inline data URIs.
        /// </summary>
        /// <param name="document">源文档；null 时返回空字符串。The source document; null returns an empty string.</param>
        /// <param name="options">渲染选项；null 时回退默认值。The render options; null falls back to defaults.</param>
        /// <returns>HTML 正文片段。The HTML body fragment.</returns>
        internal static string Render(OfficeDocument document, MarkdownOptions options)
        {
            if (document == null)
            {
                return string.Empty;
            }
            if (options == null)
            {
                options = new MarkdownOptions();
            }

            List<Block> blocks = FlattenBlocks(document);
            StringBuilder builder = new StringBuilder();
            ListStack stacks = new ListStack(builder);
            // 标题锚点：与 Markdig 的 AutoIdentifier 同样保证一篇文档内不重名。
            // Heading anchors: like Markdig's AutoIdentifier, unique within one document.
            Dictionary<string, int> usedSlugs = new Dictionary<string, int>(StringComparer.Ordinal);
            int index = 0;
            while (index < blocks.Count)
            {
                Paragraph paragraph = blocks[index] as Paragraph;
                if (paragraph == null)
                {
                    stacks.CloseAll();
                    Table table = blocks[index] as Table;
                    if (table != null)
                    {
                        WriteTable(builder, table, options);
                    }
                    else if (blocks[index] is PageBreak)
                    {
                        builder.Append("<hr />\n");
                    }
                    index++;
                    continue;
                }

                if (IsCodeParagraph(paragraph, options))
                {
                    stacks.CloseAll();
                    int start = index;
                    while (index < blocks.Count)
                    {
                        Paragraph next = blocks[index] as Paragraph;
                        if (next == null || !IsCodeParagraph(next, options))
                        {
                            break;
                        }
                        index++;
                    }
                    WriteCodeBlock(builder, blocks, start, index);
                    continue;
                }

                if (IsSeparator(paragraph))
                {
                    stacks.CloseAll();
                    builder.Append("<hr />\n");
                    index++;
                    continue;
                }

                string listKind = GetListKind(paragraph);
                if (listKind != null)
                {
                    stacks.OpenItem(GetLevel(paragraph), listKind == "ol", GetListItemNumber(paragraph));
                    WriteListItem(builder, paragraph, options);
                    index++;
                    continue;
                }

                stacks.CloseAll();
                WriteTextParagraph(builder, paragraph, options, usedSlugs);
                index++;
            }
            stacks.CloseAll();

            string html = builder.ToString();
            if (options.TrimTrailingWhitespace)
            {
                html = TrimLines(html);
            }
            return html.TrimEnd('\r', '\n', ' ', '\t');
        }

        /// <summary>把所有节的正文块按顺序摊平为一个列表。Flattens the body blocks of every section into one ordered list.</summary>
        /// <param name="document">源文档。The source document.</param>
        /// <returns>摊平后的块列表。The flattened block list.</returns>
        private static List<Block> FlattenBlocks(OfficeDocument document)
        {
            List<Block> blocks = new List<Block>();
            if (document.Sections == null)
            {
                return blocks;
            }
            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                if (section == null || section.Blocks == null)
                {
                    continue;
                }
                for (int i = 0; i < section.Blocks.Count; i++)
                {
                    if (section.Blocks[i] != null)
                    {
                        blocks.Add(section.Blocks[i]);
                    }
                }
            }
            return blocks;
        }

        /// <summary>计算列表嵌套层级：每 24 磅左缩进一层，负值取 0、超过 8 层封顶。Computes the list nesting level: one level per 24 points of left indent, floored at 0 and capped at 8.</summary>
        /// <param name="paragraph">列表段落。The list paragraph.</param>
        /// <returns>层级（从 0 开始）。The zero-based level.</returns>
        private static int GetLevel(Paragraph paragraph)
        {
            int level = (int)Math.Floor(paragraph.LeftIndent / IndentPerLevel);
            if (level < 0)
            {
                level = 0;
            }
            return level > 8 ? 8 : level;
        }

        /// <summary>判断列表段落类型：标记以数字加点/括号视为有序，否则视为无序；非列表返回 null。Classifies a list paragraph: digits followed by "." or ")" mean ordered, any other marker means unordered; null for non-list paragraphs.</summary>
        /// <param name="paragraph">待判断段落。The paragraph to classify.</param>
        /// <returns>"ol"、"ul" 或 null。"ol", "ul", or null.</returns>
        private static string GetListKind(Paragraph paragraph)
        {
            string marker = paragraph.ListMarker;
            if (string.IsNullOrEmpty(marker))
            {
                return null;
            }
            int pos = 0;
            while (pos < marker.Length && marker[pos] >= '0' && marker[pos] <= '9')
            {
                pos++;
            }
            if (pos > 0 && pos < marker.Length && (marker[pos] == '.' || marker[pos] == ')'))
            {
                return "ol";
            }
            return "ul";
        }

        /// <summary>
        /// 读取有序列表项的起始编号：解析器把 "5." 这样的标记存进 ListMarker，非数字或超过 9 位（溢出风险）返回 -1。
        /// Reads an ordered item's start number: the parser keeps markers such as "5." in ListMarker; non-numeric or
        /// longer than nine digits (an overflow risk) returns -1.
        /// </summary>
        /// <param name="paragraph">列表段落。The list paragraph.</param>
        /// <returns>起始编号，无法判定时 -1。The start number, -1 when undeterminable.</returns>
        private static int GetListItemNumber(Paragraph paragraph)
        {
            string marker = paragraph.ListMarker;
            if (string.IsNullOrEmpty(marker))
            {
                return -1;
            }
            int value = 0;
            int digits = 0;
            while (digits < marker.Length && marker[digits] >= '0' && marker[digits] <= '9')
            {
                digits++;
                if (digits > 9)
                {
                    return -1;
                }
                value = value * 10 + (marker[digits - 1] - '0');
            }
            if (digits == 0 || digits >= marker.Length || (marker[digits] != '.' && marker[digits] != ')'))
            {
                return -1;
            }
            return value;
        }

        /// <summary>判断分隔线段落：去空白后的正文恰为 em 划线或星号/下划线三连等常见分隔写法。Detects a separator paragraph: trimmed text that is exactly an em dash or the common triple star/dash/underscore forms.</summary>
        /// <param name="paragraph">待判断段落。The paragraph to test.</param>
        /// <returns>是否为分隔线。Whether it is a separator.</returns>
        private static bool IsSeparator(Paragraph paragraph)
        {
            if (!string.IsNullOrEmpty(paragraph.ListMarker) || paragraph.HeadingLevel > 0)
            {
                return false;
            }
            string text = paragraph.GetText().Trim();
            return text == "\u2014" || text == MarkdownParser.ThematicBreakText || text == "* * *"
                || text == "---" || text == "***" || text == "___";
        }

        /// <summary>判断代码段落：先认解析器打的结构标记，没有标记的手写文档才退回等宽字体推断（空文本的等宽运行也算，代码块内的空行与空围栏靠它保住）。Detects a code paragraph: the parser's structural marker wins, and only a hand-built document with no marker falls back to the monospaced-font test (an empty monospaced run counts too, since blank lines inside a block and an empty fence depend on it).</summary>
        /// <param name="paragraph">待判断段落。The paragraph to test.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <returns>是否为代码段落。Whether it is a code paragraph.</returns>
        private static bool IsCodeParagraph(Paragraph paragraph, MarkdownOptions options)
        {
            if (paragraph.HeadingLevel > 0 || !string.IsNullOrEmpty(paragraph.ListMarker))
            {
                return false;
            }
            if (string.Equals(paragraph.StyleName, MarkdownParser.CodeStyleName, StringComparison.Ordinal))
            {
                return true;
            }
            if (string.Equals(paragraph.StyleName, MarkdownParser.InlineStyleName, StringComparison.Ordinal))
            {
                return false;
            }
            bool sawCodeRun = false;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null)
                {
                    string fontFamily = run.Format != null ? run.Format.GetAsciiFontFamily() : null;
                    if (!IsMonospace(fontFamily, options))
                    {
                        // 空白的正文字体运行只是段落里的空行，不影响判定。
                        // A blank run in the body font is only an empty line; it does not disqualify the block.
                        if (!string.IsNullOrEmpty(run.Text))
                        {
                            return false;
                        }
                        continue;
                    }
                    sawCodeRun = true;
                    continue;
                }
                if (!(paragraph.Inlines[i] is FieldRun))
                {
                    return false;
                }
            }
            return sawCodeRun;
        }

        /// <summary>判断字体名是否等宽：Consolas/Courier/Monospace 前缀匹配，或等于选项指定的代码字体。Tests whether a font family is monospaced: prefix match on Consolas/Courier/Monospace, or equal to the configured code font.</summary>
        /// <param name="fontFamily">字体名。The font family.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <returns>是否等宽。Whether it is monospaced.</returns>
        private static bool IsMonospace(string fontFamily, MarkdownOptions options)
        {
            if (string.IsNullOrEmpty(fontFamily))
            {
                return false;
            }
            string name = fontFamily.Trim();
            if (name.StartsWith("Consolas", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("Courier", StringComparison.OrdinalIgnoreCase)
                || name.StartsWith("Monospace", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
            return options != null && !string.IsNullOrEmpty(options.CodeFontFamily)
                && name.Equals(options.CodeFontFamily, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>输出普通段落或标题；空内容段落跳过，非左对齐时附带 text-align 样式，标题带唯一锚点 id。Writes a plain paragraph or heading; empty paragraphs are skipped, non-left alignment adds a text-align style, and headings carry a unique anchor id.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="paragraph">源段落。The source paragraph.</param>
        /// <param name="options">渲染选项。The render options.</param>
        /// <param name="usedSlugs">已用过的标题锚点，用于去重。Slugs used so far, for de-duplication.</param>
        private static void WriteTextParagraph(StringBuilder builder, Paragraph paragraph, MarkdownOptions options, Dictionary<string, int> usedSlugs)
        {
            StringBuilder content = new StringBuilder();
            WriteInlines(content, paragraph.Inlines, options);
            if (content.Length == 0)
            {
                return;
            }
            bool heading = paragraph.HeadingLevel > 0;
            string tag = heading
                ? "h" + (paragraph.HeadingLevel > 6 ? 6 : paragraph.HeadingLevel).ToString(CultureInfo.InvariantCulture)
                : "p";
            builder.Append('<').Append(tag);
            if (heading)
            {
                builder.Append(" id=\"").Append(Escape(NextSlug(usedSlugs, Slugify(paragraph.GetText())))).Append('"');
            }
            builder.Append(GetStyleAttribute(paragraph.Alignment, null, null)).Append('>')
                   .Append(content.ToString()).Append("</").Append(tag).Append(">\n");
        }

        /// <summary>
        /// 生成标题锚点：转小写、空白变连字符、只留字母数字与 - 和 _（中日韩字符按 Unicode 字母原样保留）。
        /// Builds a heading anchor: case-folded, whitespace becomes hyphens, and only letters, digits, hyphen and
        /// underscore survive (CJK ideographs are letters in Unicode and are kept verbatim).
        /// </summary>
        /// <param name="text">标题正文。The heading text.</param>
        /// <returns>锚点串，可能为空。The slug, possibly empty.</returns>
        private static string Slugify(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            StringBuilder slug = new StringBuilder(text.Length);
            bool pendingHyphen = false;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (char.IsWhiteSpace(c))
                {
                    pendingHyphen = slug.Length > 0;
                    continue;
                }
                if (c == '-' || c == '_' || char.IsLetterOrDigit(c))
                {
                    if (pendingHyphen)
                    {
                        slug.Append('-');
                        pendingHyphen = false;
                    }
                    slug.Append(char.ToLowerInvariant(c));
                }
            }
            return slug.ToString();
        }

        /// <summary>
        /// 取一个未被占用的锚点：同名时依次追加 -1、-2……
        /// Takes an unused anchor, appending -1, -2 and so on while the slug collides.
        /// </summary>
        /// <param name="usedSlugs">已用锚点计数表。The table of slugs used so far.</param>
        /// <param name="slug">候选锚点。The candidate slug.</param>
        /// <returns>唯一锚点。A unique anchor.</returns>
        private static string NextSlug(Dictionary<string, int> usedSlugs, string slug)
        {
            if (slug.Length == 0)
            {
                slug = "section";
            }
            int count;
            if (!usedSlugs.TryGetValue(slug, out count))
            {
                usedSlugs[slug] = 1;
                return slug;
            }
            usedSlugs[slug] = count + 1;
            return slug + "-" + count.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>输出一个列表项的正文（开闭标签由层级栈管理）；启用任务列表时识别行首复选框。Writes a list item's body (the tags are managed by the level stack); with task lists enabled a leading checkbox is recognized.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="paragraph">列表段落。The list paragraph.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteListItem(StringBuilder builder, Paragraph paragraph, MarkdownOptions options)
        {
            if (options.EnableTaskLists)
            {
                bool isChecked;
                if (TryGetTaskMarkerFromList(paragraph.ListMarker, out isChecked))
                {
                    AppendCheckbox(builder, isChecked);
                    WriteInlines(builder, paragraph.Inlines, options);
                    return;
                }
                string text = paragraph.GetText();
                string rest;
                if (TryGetTaskMarker(text, out rest, out isChecked))
                {
                    AppendCheckbox(builder, isChecked);
                    builder.Append(' ').Append(Escape(rest));
                    return;
                }
            }
            WriteInlines(builder, paragraph.Inlines, options);
        }

        /// <summary>输出禁用态的复选框标签。Emits a disabled checkbox tag.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="isChecked">是否已勾选。Whether it is checked.</param>
        private static void AppendCheckbox(StringBuilder builder, bool isChecked)
        {
            builder.Append("<input type=\"checkbox\" disabled=\"disabled\"");
            if (isChecked)
            {
                builder.Append(" checked=\"checked\"");
            }
            builder.Append(" />");
        }

        /// <summary>
        /// 从列表符号里读取任务状态：解析器把 <c>☐</c> / <c>☑</c> 前缀写进 ListMarker。
        /// Reads the task state from the list marker: the parser stores the ☐ / ☑ prefix in ListMarker.
        /// </summary>
        /// <param name="marker">列表符号。The list marker.</param>
        /// <param name="isChecked">输出是否已勾选。Reports whether the item is checked.</param>
        /// <returns>是否为任务项。Whether the item is a task.</returns>
        private static bool TryGetTaskMarkerFromList(string marker, out bool isChecked)
        {
            isChecked = false;
            if (string.IsNullOrEmpty(marker))
            {
                return false;
            }
            if (marker[0] == '\u2611')
            {
                isChecked = true;
                return true;
            }
            return marker[0] == '\u2610';
        }

        /// <summary>识别任务列表标记 "[ ]"/"[x]"/"[X]" 并拆出正文与勾选状态。Recognizes the task markers "[ ]"/"[x]"/"[X]" and splits the body text and checked state.</summary>
        /// <param name="text">列表项全文。The full item text.</param>
        /// <param name="rest">去掉标记后的正文。The body without the marker.</param>
        /// <param name="isChecked">是否已勾选。Whether it is checked.</param>
        /// <returns>是否命中任务标记。Whether a task marker was found.</returns>
        private static bool TryGetTaskMarker(string text, out string rest, out bool isChecked)
        {
            rest = null;
            isChecked = false;
            if (text == null || text.Length < 3 || text[0] != '[')
            {
                return false;
            }
            if (text[1] != ' ' && text[1] != 'x' && text[1] != 'X')
            {
                return false;
            }
            if (text[2] != ']' || (text.Length > 3 && text[3] != ' '))
            {
                return false;
            }
            isChecked = text[1] == 'x' || text[1] == 'X';
            rest = text.Length > 4 ? text.Substring(4) : string.Empty;
            return true;
        }

        /// <summary>输出连续等宽段落组成的代码块，块内文本整体 HTML 转义。Writes consecutive monospaced paragraphs as one code block with fully escaped content.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="blocks">摊平的块列表。The flattened block list.</param>
        /// <param name="start">代码段起始下标。The inclusive start index.</param>
        /// <param name="end">代码段结束下标（不含）。The exclusive end index.</param>
        private static void WriteCodeBlock(StringBuilder builder, List<Block> blocks, int start, int end)
        {
            StringBuilder body = new StringBuilder();
            for (int i = start; i < end; i++)
            {
                if (i > start)
                {
                    body.Append('\n');
                }
                body.Append(Escape(GetRawText((Paragraph)blocks[i])));
            }
            // 空代码块不带收尾换行，与 CommonMark 的 <pre><code></code></pre> 一致。
            // An empty block gets no trailing newline, matching CommonMark's <pre><code></code></pre>.
            builder.Append("<pre><code>");
            if (body.Length > 0)
            {
                builder.Append(body.ToString()).Append('\n');
            }
            builder.Append("</code></pre>\n");
        }

        /// <summary>输出表格；未启用表格时降级为逐行的纯文本段落。Writes a table; when tables are disabled each row degrades to a plain text paragraph.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="table">源表格。The source table.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteTable(StringBuilder builder, Table table, MarkdownOptions options)
        {
            if (table.Rows == null || table.Rows.Count == 0)
            {
                return;
            }
            if (!options.EnableTables)
            {
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    TableRow row = table.Rows[r];
                    if (row == null || row.Cells == null)
                    {
                        continue;
                    }
                    StringBuilder line = new StringBuilder();
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        if (c > 0)
                        {
                            line.Append(" | ");
                        }
                        line.Append(Escape(GetCellText(row.Cells[c])));
                    }
                    builder.Append("<p>").Append(line.ToString()).Append("</p>\n");
                }
                return;
            }

            int headerCount = 0;
            while (headerCount < table.Rows.Count && table.Rows[headerCount] != null && table.Rows[headerCount].IsHeader)
            {
                headerCount++;
            }

            builder.Append("<table>\n");
            if (headerCount > 0)
            {
                builder.Append("<thead>\n");
                for (int r = 0; r < headerCount; r++)
                {
                    WriteTableRow(builder, table.Rows[r], true, options);
                }
                builder.Append("</thead>\n");
            }
            if (headerCount < table.Rows.Count)
            {
                builder.Append("<tbody>\n");
                for (int r = headerCount; r < table.Rows.Count; r++)
                {
                    TableRow row = table.Rows[r];
                    if (row != null)
                    {
                        WriteTableRow(builder, row, false, options);
                    }
                }
                builder.Append("</tbody>\n");
            }
            builder.Append("</table>\n");
        }

        /// <summary>输出一行表格行（th 或 td），含 colspan/rowspan 与对齐、底色样式。Writes one table row (th or td) including colspan/rowspan plus alignment and background styles.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="row">源行。The source row.</param>
        /// <param name="header">是否为表头行。Whether the row is a header row.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteTableRow(StringBuilder builder, TableRow row, bool header, MarkdownOptions options)
        {
            builder.Append("<tr>");
            if (row.Cells != null)
            {
                for (int c = 0; c < row.Cells.Count; c++)
                {
                    TableCell cell = row.Cells[c];
                    if (cell == null)
                    {
                        continue;
                    }
                    string tag = header ? "th" : "td";
                    builder.Append('<').Append(tag);
                    if (cell.ColSpan > 1)
                    {
                        builder.Append(" colspan=\"").Append(cell.ColSpan.ToString(CultureInfo.InvariantCulture)).Append('"');
                    }
                    if (cell.RowSpan > 1)
                    {
                        builder.Append(" rowspan=\"").Append(cell.RowSpan.ToString(CultureInfo.InvariantCulture)).Append('"');
                    }
                    ParagraphAlignment alignment = ParagraphAlignment.Left;
                    Paragraph first = cell.Blocks != null ? cell.Blocks[0] as Paragraph : null;
                    if (first != null)
                    {
                        alignment = first.Alignment;
                    }
                    string background = cell.BackgroundColor.HasValue && cell.BackgroundColor.Value.A != 0
                        ? "#" + cell.BackgroundColor.Value.ToHex()
                        : null;
                    builder.Append(GetStyleAttribute(alignment, null, background)).Append('>');
                    WriteCellContent(builder, cell, options);
                    builder.Append("</").Append(tag).Append('>');
                }
            }
            builder.Append("</tr>\n");
        }

        /// <summary>输出单元格内容：多个段落的行内结果以 br 连接。Writes the cell content: the inline output of its paragraphs joined by br.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="cell">源单元格。The source cell.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteCellContent(StringBuilder builder, TableCell cell, MarkdownOptions options)
        {
            if (cell.Blocks == null)
            {
                return;
            }
            bool firstParagraph = true;
            for (int b = 0; b < cell.Blocks.Count; b++)
            {
                Paragraph paragraph = cell.Blocks[b] as Paragraph;
                if (paragraph == null)
                {
                    continue;
                }
                StringBuilder content = new StringBuilder();
                WriteInlines(content, paragraph.Inlines, options);
                if (content.Length == 0)
                {
                    continue;
                }
                if (!firstParagraph)
                {
                    builder.Append("<br />");
                }
                builder.Append(content.ToString());
                firstParagraph = false;
            }
        }

        /// <summary>
        /// 输出段落内全部行内元素：链接转 a、颜色转 span、下划线转 u、删除线转 del、粗体转 strong、斜体转 em、等宽转 code。
        /// 相邻同格式的运行共用一对标签（连续层用状态机开合），因此 "bold <em>italic</em> bold" 不再被拆成三段 strong。
        /// Writes a paragraph's inlines: links to a, color to span, underline to u, strikeout to del, bold to strong,
        /// italic to em and monospace to code. Adjacent runs sharing a format reuse one pair of tags - the layers open
        /// and close through a state machine - so "bold italic bold" is no longer split into three strong elements.
        /// </summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="inlines">行内元素列表。The inline elements.</param>
        /// <param name="options">渲染选项。The render options.</param>
        private static void WriteInlines(StringBuilder builder, List<Inline> inlines, MarkdownOptions options)
        {
            if (inlines == null)
            {
                return;
            }
            InlineFrame frame = new InlineFrame(builder, options);
            for (int i = 0; i < inlines.Count; i++)
            {
                Inline inline = inlines[i];
                TextRun text = inline as TextRun;
                if (text != null)
                {
                    frame.Apply(text.Text, text.Link, text.Format);
                    continue;
                }
                FieldRun field = inline as FieldRun;
                if (field != null)
                {
                    frame.Apply(field.DisplayText, null, field.Format);
                    continue;
                }
                ImageRun image = inline as ImageRun;
                if (image != null)
                {
                    frame.AppendImage(image);
                }
            }
            frame.Close();
        }

        /// <summary>
        /// 行内标签层的状态机：按 a &gt; span &gt; u &gt; del &gt; strong &gt; em &gt; code 的固定顺序开合标签
        /// （粗斜体合并成一个同时带两面格式的运行，这个顺序让 <c>strong</c> 能横跨整段而不被 <c>em</c> 切开），
        /// 只在相邻运行的层序真正变化时才闭合再重开，避免同一段格式被切成多个同名元素。
        /// The inline tag state machine: layers open and close in the fixed order a &gt; span &gt; u &gt; del &gt;
        /// strong &gt; em &gt; code, which keeps one &lt;strong&gt; stretch unbroken around the &lt;em&gt; inside it, and only the
        /// part of the stack that actually changed is closed and reopened, so one formatted stretch is never cut into
        /// several identically named elements.
        /// </summary>
        private sealed class InlineFrame
        {
            private const int LayerLink = 0;
            private const int LayerSpan = 1;
            private const int LayerUnderline = 2;
            private const int LayerDelete = 3;
            private const int LayerStrong = 4;
            private const int LayerEmphasis = 5;
            private const int LayerCode = 6;
            private const int LayerCount = 7;

            private readonly StringBuilder builder;
            private readonly MarkdownOptions options;
            private readonly int[] openLayers = new int[LayerCount];
            private readonly string[] openKeys = new string[LayerCount];
            private readonly int[] wantedLayers = new int[LayerCount];
            private readonly string[] wantedKeys = new string[LayerCount];
            private int depth;
            private int wantedDepth;

            /// <summary>绑定输出缓冲区与渲染选项。Binds the output buffer and the render options.</summary>
            /// <param name="builder">输出缓冲区。The output buffer.</param>
            /// <param name="options">渲染选项。The render options.</param>
            public InlineFrame(StringBuilder builder, MarkdownOptions options)
            {
                this.builder = builder;
                this.options = options;
            }

            /// <summary>写一段文本：先按该运行的格式调整标签层，再输出转义后的正文。Writes a text run: the tag layers are adjusted to the run's format, then the escaped text is emitted.</summary>
            /// <param name="text">文本内容。The text content.</param>
            /// <param name="link">超链接地址，可为空。The link target, possibly empty.</param>
            /// <param name="format">文本格式。The text format.</param>
            public void Apply(string text, string link, TextFormat format)
            {
                if (string.IsNullOrEmpty(text))
                {
                    return;
                }
                if (IsNewLineOnly(text))
                {
                    builder.Append("<br />\n");
                    return;
                }
                Build(format, link);
                Sync();
                WriteText(text);
            }

            /// <summary>在当前仍打开的标签层里写图片。Writes an image inside the layers still open.</summary>
            /// <param name="image">图片运行。The image run.</param>
            public void AppendImage(ImageRun image)
            {
                WriteImage(builder, image);
            }

            /// <summary>闭合所有仍然打开的标签层。Closes every layer still open.</summary>
            public void Close()
            {
                CloseTo(0);
            }

            /// <summary>把该运行的格式换算成期望的层序列（顺序即嵌套顺序）。Turns a run's format into the wanted layer sequence, in nesting order.</summary>
            /// <param name="format">文本格式。The text format.</param>
            /// <param name="link">超链接地址。The link target.</param>
            private void Build(TextFormat format, string link)
            {
                wantedDepth = 0;
                TextFormat effective = format ?? new TextFormat();
                if (!string.IsNullOrEmpty(link) && IsSafeLink(link))
                {
                    Add(LayerLink, "<a href=\"" + Escape(link) + "\">");
                }
                if (effective.Color != OfficeColor.Black && effective.Color.A != 0)
                {
                    Add(LayerSpan, "<span style=\"color:#" + effective.Color.ToHex() + "\">");
                }
                if (effective.Underline)
                {
                    Add(LayerUnderline, "<u>");
                }
                if (effective.Strikeout)
                {
                    Add(LayerDelete, "<del>");
                }
                if (effective.Bold)
                {
                    Add(LayerStrong, "<strong>");
                }
                if (effective.Italic)
                {
                    Add(LayerEmphasis, "<em>");
                }
                if (IsMonospace(effective.GetAsciiFontFamily(), options))
                {
                    Add(LayerCode, "<code>");
                }
            }

            /// <summary>追加一层。Appends one layer.</summary>
            /// <param name="layer">层序号。The layer index.</param>
            /// <param name="openTag">起始标签。The opening tag.</param>
            private void Add(int layer, string openTag)
            {
                wantedLayers[wantedDepth] = layer;
                wantedKeys[wantedDepth] = openTag;
                wantedDepth++;
            }

            /// <summary>把已打开的层调整到期望状态：公共前缀保持不动，多出来的闭合、缺少的打开。Adjusts the open stack to the wanted state, keeping the common prefix, closing what is left over and opening what is missing.</summary>
            private void Sync()
            {
                int common = 0;
                while (common < depth && common < wantedDepth
                       && openLayers[common] == wantedLayers[common]
                       && string.Equals(openKeys[common], wantedKeys[common], StringComparison.Ordinal))
                {
                    common++;
                }
                CloseTo(common);
                for (int i = common; i < wantedDepth; i++)
                {
                    builder.Append(wantedKeys[i]);
                    openLayers[depth] = wantedLayers[i];
                    openKeys[depth] = wantedKeys[i];
                    depth++;
                }
            }

            /// <summary>闭合到指定深度。Closes layers down to the given depth.</summary>
            /// <param name="common">要保留的层数。The number of layers to keep.</param>
            private void CloseTo(int common)
            {
                while (depth > common)
                {
                    depth--;
                    builder.Append(CloseTag(openLayers[depth]));
                }
            }

            /// <summary>层序号对应的闭合标签。The closing tag matching a layer index.</summary>
            /// <param name="layer">层序号。The layer index.</param>
            /// <returns>闭合标签。The closing tag.</returns>
            private static string CloseTag(int layer)
            {
                switch (layer)
                {
                    case LayerLink:
                        return "</a>";
                    case LayerSpan:
                        return "</span>";
                    case LayerUnderline:
                        return "</u>";
                    case LayerDelete:
                        return "</del>";
                    case LayerStrong:
                        return "</strong>";
                    case LayerEmphasis:
                        return "</em>";
                    default:
                        return "</code>";
                }
            }

            /// <summary>输出转义后的正文，硬换行选项决定文本内嵌换行的处理。Emits the escaped text; the hard-break option governs embedded newlines.</summary>
            /// <param name="text">原始文本。The raw text.</param>
            private void WriteText(string text)
            {
                string content = Escape(text);
                if (options != null && options.HardLineBreaks)
                {
                    content = content.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", "<br />\n");
                }
                else
                {
                    content = content.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ');
                }
                builder.Append(content);
            }
        }


        /// <summary>输出图片为内联 base64 data URI 的 img 标签；尺寸取显式磅值，缺省用图片默认磅值并换算为像素。刻意为之：内联字节使 HTML 片段无需外部资源即可独立查看。Writes the image as an img tag with an inline base64 data URI; sizes use explicit points or the defaults converted to pixels. Deliberate: inlining the bytes lets the fragment stand alone without external resources.</summary>
        /// <param name="builder">输出缓冲区。The output buffer.</param>
        /// <param name="image">图片运行。The image run.</param>
        private static void WriteImage(StringBuilder builder, ImageRun image)
        {
            OfficeImage data = image.Image;
            if (data == null || data.Data == null || data.Data.Length == 0)
            {
                return;
            }
            double widthPoints = image.Width > 0 ? image.Width : data.DefaultWidthPoints;
            double heightPoints = image.Height > 0 ? image.Height : data.DefaultHeightPoints;
            int widthPixels = (int)Math.Round(widthPoints * 96.0 / 72.0, MidpointRounding.AwayFromZero);
            int heightPixels = (int)Math.Round(heightPoints * 96.0 / 72.0, MidpointRounding.AwayFromZero);
            builder.Append("<img src=\"data:").Append(GetMimeType(data)).Append(";base64,")
                   .Append(Convert.ToBase64String(data.Data))
                   .Append("\" width=\"").Append(widthPixels.ToString(CultureInfo.InvariantCulture))
                   .Append("\" height=\"").Append(heightPixels.ToString(CultureInfo.InvariantCulture))
                   .Append("\" alt=\"\" />");
        }

        /// <summary>按图片格式返回 MIME 类型；未知格式回退 image/png。Returns the MIME type for the image format; unknown formats fall back to image/png.</summary>
        /// <param name="image">图片数据。The image data.</param>
        /// <returns>MIME 类型字符串。The MIME type string.</returns>
        private static string GetMimeType(OfficeImage image)
        {
            if (image.Format == OfficeImageFormat.Jpeg)
            {
                return "image/jpeg";
            }
            return "image/png";
        }

        /// <summary>生成 style 属性文本（对齐、字号、底色），全部缺省时返回空串。Builds the style attribute text (alignment, font size, background), returning an empty string when all are absent.</summary>
        /// <param name="alignment">段落对齐；Left 表示不输出。The paragraph alignment; Left outputs nothing.</param>
        /// <param name="fontFamily">字体名；空表示不输出。The font family; empty outputs nothing.</param>
        /// <param name="background">底色十六进制；空表示不输出。The background hex; empty outputs nothing.</param>
        /// <returns>形如 ' style="..."' 的属性文本。Attribute text such as ' style="..."'.</returns>
        private static string GetStyleAttribute(ParagraphAlignment alignment, string fontFamily, string background)
        {
            StringBuilder style = new StringBuilder();
            if (alignment == ParagraphAlignment.Center)
            {
                style.Append("text-align:center;");
            }
            else if (alignment == ParagraphAlignment.Right)
            {
                style.Append("text-align:right;");
            }
            else if (alignment == ParagraphAlignment.Justify)
            {
                style.Append("text-align:justify;");
            }
            if (!string.IsNullOrEmpty(fontFamily))
            {
                style.Append("font-family:").Append(Escape(fontFamily)).Append(';');
            }
            if (!string.IsNullOrEmpty(background))
            {
                style.Append("background-color:").Append(background).Append(';');
            }
            if (style.Length == 0)
            {
                return string.Empty;
            }
            return " style=\"" + style.ToString().TrimEnd(';').ToString() + "\"";
        }

        /// <summary>提取段落原始文本：文本运行取正文、域取显示文本、图片忽略。Extracts the raw paragraph text: text runs contribute their body, fields their display text, images are skipped.</summary>
        /// <param name="paragraph">源段落。The source paragraph.</param>
        /// <returns>原始文本。The raw text.</returns>
        private static string GetRawText(Paragraph paragraph)
        {
            StringBuilder text = new StringBuilder();
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null)
                {
                    text.Append(run.Text);
                    continue;
                }
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field != null)
                {
                    text.Append(field.DisplayText);
                }
            }
            return text.ToString();
        }

        /// <summary>提取单元格纯文本（段落间以空格连接）。Extracts the cell's plain text (paragraphs joined by spaces).</summary>
        /// <param name="cell">源单元格。The source cell.</param>
        /// <returns>纯文本。The plain text.</returns>
        private static string GetCellText(TableCell cell)
        {
            StringBuilder text = new StringBuilder();
            if (cell == null || cell.Blocks == null)
            {
                return text.ToString();
            }
            for (int b = 0; b < cell.Blocks.Count; b++)
            {
                Paragraph paragraph = cell.Blocks[b] as Paragraph;
                if (paragraph == null)
                {
                    continue;
                }
                if (text.Length > 0)
                {
                    text.Append(' ');
                }
                text.Append(GetRawText(paragraph));
            }
            return text.ToString();
        }

        /// <summary>判断文本是否仅由换行符组成。Tests whether the text consists solely of line-break characters.</summary>
        /// <param name="text">待判断文本。The text to test.</param>
        /// <returns>是否为纯换行。Whether it is a pure line break.</returns>
        private static bool IsNewLineOnly(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] != '\n' && text[i] != '\r')
                {
                    return false;
                }
            }
            return text.Length > 0;
        }

        /// <summary>
        /// 判断链接可否写进 href：允许相对地址、锚点与 http/https/mailto/ftp/tel 协议，其余（javascript:、data: 等）按纯文本输出。
        /// 以 "//" 或反斜杠变体开头的协议相对地址会被拒绝，否则会静默跳到外部域。
        /// Decides whether a link may go into href: relative addresses, anchors and http/https/mailto/ftp/tel schemes are allowed;
        /// anything else (javascript:, data:, and so on) is rendered as plain text. Protocol-relative targets opening with "//"
        /// (or a backslash variant) are rejected, since they would silently jump to an external domain.
        /// </summary>
        /// <param name="link">链接目标。The link target.</param>
        /// <returns>是否安全。Whether it is safe to emit.</returns>
        private static bool IsSafeLink(string link)
        {
            if (string.IsNullOrEmpty(link))
            {
                return false;
            }
            if (IsProtocolRelative(link))
            {
                return false;
            }
            int colon = link.IndexOf(':');
            if (colon < 0)
            {
                return true;
            }
            int slash = link.IndexOf('/');
            if (slash >= 0 && slash < colon)
            {
                return true;
            }
            if (link.StartsWith("#", StringComparison.Ordinal))
            {
                return true;
            }
            string scheme = link.Substring(0, colon).ToLowerInvariant();
            return scheme == "http" || scheme == "https" || scheme == "mailto" || scheme == "ftp" || scheme == "ftps" || scheme == "tel";
        }

        /// <summary>
        /// 判断地址是否为协议相对形式（"//host"、"/\host"、"\host" 等浏览器会补当前协议的外链）。
        /// Decides whether a target is protocol-relative ("//host", "/\host", "\host" and similar forms browsers complete with the current scheme).
        /// </summary>
        /// <param name="link">链接目标。The link target.</param>
        /// <returns>是否协议相对。Whether the target is protocol-relative.</returns>
        private static bool IsProtocolRelative(string link)
        {
            if (link.Length < 2)
            {
                return false;
            }
            char first = link[0];
            char second = link[1];
            if (first != '/' && first != '\\')
            {
                return false;
            }
            return second == '/' || second == '\\';
        }

        /// <summary>HTML 转义 &amp;、&lt;、&gt;、"（硬要求：所有文本一律转义，原始 HTML 永不透传）。Escapes &amp;, &lt;, &gt; and " (hard requirement: all text is escaped, raw HTML is never passed through).</summary>
        /// <param name="text">待转义文本。The text to escape.</param>
        /// <returns>转义后的文本。The escaped text.</returns>
        private static string Escape(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }
            StringBuilder escaped = new StringBuilder(text.Length + 8);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '&')
                {
                    escaped.Append("&amp;");
                }
                else if (c == '<')
                {
                    escaped.Append("&lt;");
                }
                else if (c == '>')
                {
                    escaped.Append("&gt;");
                }
                else if (c == '"')
                {
                    escaped.Append("&quot;");
                }
                else
                {
                    escaped.Append(c);
                }
            }
            return escaped.ToString();
        }

        /// <summary>逐行去掉行尾空白（HTML 的换行由 br 标签表达，尾部空格无特殊含义），并去掉整体尾部多余空行。Trims trailing whitespace per line (HTML breaks lines with br tags, so trailing spaces carry no meaning) and removes trailing blank lines.</summary>
        /// <param name="text">原始文本。The raw text.</param>
        /// <returns>整理后的文本。The tidied text.</returns>
        private static string TrimLines(string text)
        {
            string[] lines = text.Split('\n');
            StringBuilder trimmed = new StringBuilder(text.Length);
            bool inPre = false;
            for (int i = 0; i < lines.Length; i++)
            {
                // <pre> 里的行尾空白是代码内容本身，裁掉就改变了导出的字节序列；
                // 起始标签与正文同处一行，所以含 <pre 的那一行也整体保留。
                // Trailing whitespace inside <pre> is code content; trimming it changes the exported bytes. The
                // opening tag shares a line with the first body line, so a line containing <pre is kept as well.
                bool keep = inPre || lines[i].IndexOf("<pre", StringComparison.Ordinal) >= 0;
                string line = keep ? lines[i] : lines[i].TrimEnd('\r', ' ', '\t');
                int open = lines[i].IndexOf("<pre", StringComparison.Ordinal);
                int close = lines[i].IndexOf("</pre", StringComparison.Ordinal);
                if (open >= 0 || close >= 0)
                {
                    inPre = open >= 0 && (close < 0 || open > close);
                }
                if (i < lines.Length - 1)
                {
                    trimmed.Append(line).Append('\n');
                }
                else
                {
                    trimmed.Append(line);
                }
            }
            return trimmed.ToString();
        }

        /// <summary>
        /// 跟踪未闭合的 ul/ol 层级栈：在层级下降、同级换列表或列表结束时补发闭合标签。
        /// Tracks the stack of open ul/ol levels, emitting closing tags when the level drops, a same-level list changes kind, or the list ends.
        /// </summary>
        private sealed class ListStack
        {
            private readonly StringBuilder builder;
            private readonly List<bool> orderedFlags = new List<bool>();
            private bool itemOpen;

            /// <summary>绑定输出缓冲区。Binds the output buffer.</summary>
            /// <param name="builder">输出缓冲区。The output buffer.</param>
            public ListStack(StringBuilder builder)
            {
                this.builder = builder;
            }

            /// <summary>在指定层级开一个新列表项，必要时先闭合更深层级或补开中间层级；新开一层的有序表带上起始编号。Opens a list item at the given level, closing deeper levels or opening intermediate ones as needed; a freshly opened ordered list carries its start number.</summary>
            /// <param name="level">层级（从 0 开始）。The zero-based level.</param>
            /// <param name="ordered">是否为有序列表。Whether the list is ordered.</param>
            /// <param name="number">本项的编号，0 表示未知或无序。This item's number, 0 when unknown or unordered.</param>
            public void OpenItem(int level, bool ordered, int number)
            {
                while (orderedFlags.Count > level + 1)
                {
                    CloseLevel();
                }
                if (orderedFlags.Count == level + 1)
                {
                    int top = orderedFlags.Count - 1;
                    if (orderedFlags[top] != ordered)
                    {
                        CloseLevel();
                    }
                    else if (itemOpen)
                    {
                        builder.Append("</li>\n");
                        itemOpen = false;
                    }
                }
                while (orderedFlags.Count < level + 1)
                {
                    if (orderedFlags.Count > 0 && !itemOpen)
                    {
                        builder.Append("<li>\n");
                        itemOpen = true;
                    }
                    bool nestedOrdered = orderedFlags.Count == level && ordered;
                    if (nestedOrdered && number >= 0 && number != 1)
                    {
                        // CommonMark 只在起始编号不是 1 时输出 start 属性（0 也要输出）。
                        // CommonMark emits the start attribute whenever the first number is not 1, zero included.
                        builder.Append("<ol start=\"").Append(number.ToString(CultureInfo.InvariantCulture)).Append("\">\n");
                    }
                    else
                    {
                        builder.Append(nestedOrdered ? "<ol>\n" : "<ul>\n");
                    }
                    orderedFlags.Add(nestedOrdered);
                    itemOpen = false;
                }
                builder.Append("<li>");
                itemOpen = true;
            }

            /// <summary>闭合全部仍未关闭的列表。Closes every still-open list level.</summary>
            public void CloseAll()
            {
                while (orderedFlags.Count > 0)
                {
                    CloseLevel();
                }
            }

            /// <summary>闭合最深一层列表：补发 li 与 ul/ol 闭合标签，并把父层 li 标记为仍然打开。Closes the deepest list level, emitting the li and ul/ol end tags and marking the parent's li as still open.</summary>
            private void CloseLevel()
            {
                if (itemOpen)
                {
                    builder.Append("</li>\n");
                    itemOpen = false;
                }
                bool ordered = orderedFlags[orderedFlags.Count - 1];
                builder.Append(ordered ? "</ol>\n" : "</ul>\n");
                orderedFlags.RemoveAt(orderedFlags.Count - 1);
                if (orderedFlags.Count > 0)
                {
                    itemOpen = true;
                }
            }
        }
    }
}
