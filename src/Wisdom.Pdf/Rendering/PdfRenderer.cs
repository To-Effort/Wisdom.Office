using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 对统一文档模型进行排版并写出 PDF 1.7 文件。
    /// Lays out the unified document model and writes a PDF 1.7 file.
    /// </summary>
    public sealed partial class PdfRenderer
    {
        /// <summary>单元格排版结果：位置、跨度与内容尺寸。A laid-out cell: position, spans, and content size.</summary>
        private sealed class CellLayout
        {
            /// <summary>被排版的单元格。The cell being laid out.</summary>
            public TableCell Cell;
            /// <summary>起始行号。The starting row index.</summary>
            public int Row;
            /// <summary>纵向合并的行数，默认 1。The number of rows spanned, default 1.</summary>
            public int RowSpan = 1;
            /// <summary>起始列号。The starting column index.</summary>
            public int Column;
            /// <summary>横向合并的列数，默认 1。The number of columns spanned, default 1.</summary>
            public int ColumnSpan = 1;
            /// <summary>相对表格左边的 X 坐标。The X offset relative to the table's left edge.</summary>
            public double X;
            /// <summary>单元格总宽度。The full cell width.</summary>
            public double Width;
            /// <summary>扣除内边距后的内容宽度。The content width after padding.</summary>
            public double ContentWidth;
            /// <summary>测量得到的内容高度。The measured content height.</summary>
            public double ContentHeight;
            /// <summary>按 <see cref="TableCell.Blocks"/> 下标缓存的断行结果，绘制时直接复用，非段落项为 null。
            /// Wrapped lines cached per block index so drawing reuses the measurement pass; null for non-paragraph blocks.</summary>
            public List<PdfLayoutLine>[] BlockLines;
        }

        /// <summary>表格排版结果：列宽、行高、覆盖矩阵与总高度。A laid-out table: column widths, row heights, the coverage grid, and the total height.</summary>
        private sealed class TableLayout
        {
            /// <summary>被排版的表格。The table being laid out.</summary>
            public Table Table;
            /// <summary>单元格内边距。The cell padding.</summary>
            public double Padding;
            /// <summary>表格左边的绝对 X 坐标。The absolute left X of the table.</summary>
            public double OriginX;
            /// <summary>各列宽度。The width of each column.</summary>
            public double[] Widths;
            /// <summary>各列相对左边的 X。The relative left X of each column.</summary>
            public double[] ColumnX;
            /// <summary>各行高度。The height of each row.</summary>
            public double[] Heights;
            /// <summary>每行实际放置的单元格。The cells actually placed in each row.</summary>
            public List<CellLayout>[] RowCells;
            /// <summary>覆盖网格：每格指向占据它的单元格。The coverage grid; each slot points to the cell occupying it.</summary>
            public CellLayout[][] Coverage;
            /// <summary>表格总高度。The total table height.</summary>
            public double TotalHeight;
        }

        /// <summary>匹配页码占位符 {page}。Matches the {page} page-number token.</summary>
        private static readonly Regex PageToken = new Regex("\\{page\\}", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        /// <summary>匹配总页数占位符 {pages}。Matches the {pages} total-page-count token.</summary>
        private static readonly Regex PagesToken = new Regex("\\{pages\\}", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>字体解析器，null 时取默认实例。The font resolver; the default instance when null.</summary>
        private readonly FontResolver _resolver;
        /// <summary>已排版页面内容列表。The list of laid-out page contents.</summary>
        private readonly List<PdfPageContent> _pages = new List<PdfPageContent>();
        /// <summary>本轮渲染的字体集合。The font set for the current render pass.</summary>
        private PdfFontSet _fonts;
        /// <summary>本轮渲染的图片集合。The image collection for the current render pass.</summary>
        private PdfImageCollection _images;
        /// <summary>待渲染的文档模型。The document model being rendered.</summary>
        private OfficeDocument _document;
        /// <summary>默认文本格式。The default text format.</summary>
        private TextFormat _defaultFormat;
        /// <summary>当前处理的节。The section currently being processed.</summary>
        private Section _section;
        /// <summary>当前绘制的页面。The page currently being drawn.</summary>
        private PdfPageContent _page;
        /// <summary>当前纵向游标（页面坐标）。The current vertical cursor (page coordinates).</summary>
        private double _cursorY;
        /// <summary>内容区左边 X。The content-area left X.</summary>
        private double _contentLeft;
        /// <summary>内容区宽度。The content-area width.</summary>
        private double _contentWidth;
        /// <summary>内容区顶部 Y。The content-area top Y.</summary>
        private double _contentTop;
        /// <summary>内容区底部 Y。The content-area bottom Y.</summary>
        private double _contentBottom;

        /// <summary>
        /// 使用默认字体解析器创建渲染器。
        /// Creates a renderer using the default font resolver.
        /// </summary>
        public PdfRenderer()
            : this(null)
        {
        }

        /// <summary>
        /// 使用指定字体解析器创建渲染器。
        /// Creates a renderer with the given font resolver.
        /// </summary>
        /// <param name="resolver">字体解析器，null 时取 <see cref="FontResolver.Default"/>。The font resolver; <see cref="FontResolver.Default"/> when null.</param>
        public PdfRenderer(FontResolver resolver)
        {
            _resolver = resolver;
        }

        /// <summary>生效的字体解析器。The effective font resolver.</summary>
        public FontResolver Resolver
        {
            get { return _resolver ?? FontResolver.Default; }
        }

        /// <summary>
        /// 渲染文档并写入指定路径的文件。
        /// Renders the document and writes it to the given file path.
        /// </summary>
        /// <param name="document">待渲染的文档模型。The document model to render.</param>
        /// <param name="path">输出文件路径。The output file path.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> 为 null 时抛出。Thrown when <paramref name="path"/> is null.</exception>
        public void Render(OfficeDocument document, string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                Render(document, stream);
            }
        }

        /// <summary>
        /// 渲染文档并返回 PDF 字节。
        /// Renders the document and returns the PDF bytes.
        /// </summary>
        /// <param name="document">待渲染的文档模型。The document model to render.</param>
        /// <returns>完整 PDF 文件的字节。The bytes of the complete PDF file.</returns>
        public byte[] RenderToBytes(OfficeDocument document)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Render(document, stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// 渲染文档并写入输出流：先排版与绘制，再序列化对象。
        /// Renders the document into the output stream: lays out and draws, then serializes the objects.
        /// </summary>
        /// <param name="document">待渲染的文档模型。The document model to render.</param>
        /// <param name="output">接收 PDF 的流。The stream receiving the PDF.</param>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="output"/> 为 null 时抛出。Thrown when <paramref name="document"/> or <paramref name="output"/> is null.</exception>
        public void Render(OfficeDocument document, Stream output)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (output == null) throw new ArgumentNullException("output");

            _document = document;
            _defaultFormat = document.DefaultTextFormat ?? new TextFormat();
            if (_defaultFormat.FontSize <= 0) _defaultFormat.FontSize = 11;
            _fonts = new PdfFontSet(_resolver);
            _images = new PdfImageCollection();
            _pages.Clear();

            LayoutDocument();
            DrawHeadersAndFooters();
            WriteDocument(output);
        }

        /// <summary>
        /// 逐节排版：无节时补一个空节，逐块绘制到当前页面游标。
        /// Lays out section by section: adds an empty section when none exist and draws each block onto the current page cursor.
        /// </summary>
        private void LayoutDocument()
        {
            if (_document.Sections.Count == 0)
            {
                _document.Sections.Add(new Section());
            }

            for (int s = 0; s < _document.Sections.Count; s++)
            {
                _section = _document.Sections[s];
                StartNewPage();
                for (int b = 0; b < _section.Blocks.Count; b++)
                {
                    DrawBlock(_section.Blocks[b]);
                }
            }
        }

        /// <summary>
        /// 依据节的页面设置新建一页并重置内容边界与游标。
        /// Starts a new page from the section's setup, resetting the content bounds and cursor.
        /// </summary>
        private void StartNewPage()
        {
            PageSetup setup = _section.PageSetup ?? new PageSetup();
            _page = new PdfPageContent(setup.Width, setup.Height, _section);
            _page.PageNumber = _pages.Count + 1;
            _pages.Add(_page);

            _contentLeft = Math.Max(0, setup.MarginLeft);
            _contentWidth = Math.Max(36, setup.Width - setup.MarginLeft - setup.MarginRight);
            _contentTop = Math.Max(0, setup.MarginTop);
            _contentBottom = setup.Height - Math.Max(0, setup.MarginBottom);
            if (_contentBottom < _contentTop + 36) _contentBottom = _contentTop + 36;
            _cursorY = _contentTop;
        }

        /// <summary>
        /// 按块类型分派绘制：段落、表格或分页符。
        /// Dispatches drawing by block type: paragraph, table, or page break.
        /// </summary>
        /// <param name="block">待绘制的内容块，null 时忽略。The block to draw; null is ignored.</param>
        private void DrawBlock(Block block)
        {
            if (block == null) return;
            Paragraph paragraph = block as Paragraph;
            if (paragraph != null)
            {
                DrawParagraphFlow(paragraph);
                return;
            }
            Table table = block as Table;
            if (table != null)
            {
                DrawTableFlow(table);
                return;
            }
            if (block is PageBreak)
            {
                if (_cursorY > _contentTop + 0.01 && _section.AllowAutoPageBreak)
                {
                    StartNewPage();
                }
            }
        }

        /// <summary>
        /// 绘制段落流：先换行排版，再按页高在行间自动分页。
        /// Draws a paragraph flow: wraps it first, then breaks across pages between lines as the page fills.
        /// </summary>
        /// <param name="paragraph">待绘制的段落。The paragraph to draw.</param>
        private void DrawParagraphFlow(Paragraph paragraph)
        {
            if (paragraph.PageBreakBefore && _cursorY > _contentTop + 0.01 && _section.AllowAutoPageBreak)
            {
                StartNewPage();
            }

            List<PdfLayoutLine> lines = ParagraphLayouter.Wrap(paragraph, _contentWidth, _defaultFormat, _fonts);
            if (_cursorY > _contentTop + 0.01 && paragraph.SpaceBefore > 0)
            {
                _cursorY += paragraph.SpaceBefore;
            }

            for (int i = 0; i < lines.Count; i++)
            {
                PdfLayoutLine line = lines[i];
                if (_cursorY + line.Height > _contentBottom + 0.01 && _cursorY > _contentTop + 0.01 && _section.AllowAutoPageBreak)
                {
                    StartNewPage();
                }
                DrawTextLine(line, paragraph, _contentLeft, _cursorY, i == lines.Count - 1);
                _cursorY += line.Height;
            }

            if (paragraph.SpaceAfter > 0)
            {
                _cursorY += paragraph.SpaceAfter;
            }
        }

        /// <summary>
        /// 绘制一行：按对齐与两端对齐调整横向位置，把连续同属性字元合并成绘制命令。
        /// Draws one line: adjusts the horizontal position by alignment and justification, then groups consecutive same-style characters into draw commands.
        /// </summary>
        /// <param name="line">排版后的行。The laid-out line.</param>
        /// <param name="paragraph">所属段落，提供对齐方式。The owning paragraph, supplying the alignment.</param>
        /// <param name="originX">行起始 X。The line's start X.</param>
        /// <param name="lineTop">行顶部 Y。The line's top Y.</param>
        /// <param name="isLastLine">是否为末段行（两端对齐时不拉伸）。Whether this is the last line (justification does not stretch it).</param>
        private void DrawTextLine(PdfLayoutLine line, Paragraph paragraph, double originX, double lineTop, bool isLastLine)
        {
            ParagraphAlignment alignment = paragraph.Alignment;
            double x = originX + line.LeftIndent;
            double available = line.AvailableWidth;
            double extraSpacing = 0;

            if (alignment == ParagraphAlignment.Justify && !isLastLine && line.Width < available)
            {
                int spaceCount = 0;
                for (int a = 0; a < line.Atoms.Count; a++)
                {
                    if (line.Atoms[a].IsSpace) spaceCount += Math.Max(1, line.Atoms[a].Chars.Count);
                }
                if (spaceCount > 0)
                {
                    extraSpacing = (available - line.Width) / spaceCount;
                    if (extraSpacing < 0) extraSpacing = 0;
                }
            }
            else if (alignment == ParagraphAlignment.Center)
            {
                x += Math.Max(0, (available - line.Width) / 2);
            }
            else if (alignment == ParagraphAlignment.Right)
            {
                x += Math.Max(0, available - line.Width);
            }

            for (int a = 0; a < line.Atoms.Count; a++)
            {
                LayoutAtom atom = line.Atoms[a];
                if (atom.IsImage)
                {
                    double imageY = lineTop + Math.Max(0, (line.Height - atom.ImageHeight) / 2);
                    if (atom.Image.Format != OfficeImageFormat.Unknown)
                    {
                        PdfImage pdfImage = _images.Get(atom.Image);
                        _page.Add(new PdfImageCommand
                        {
                            Image = pdfImage,
                            X = x,
                            Y = imageY,
                            Width = atom.ImageWidth,
                            Height = atom.ImageHeight
                        });
                        _page.AddImage(pdfImage);
                    }
                    x += atom.Width;
                    continue;
                }

                int index = 0;
                while (index < atom.Chars.Count)
                {
                    ShapedChar first = atom.Chars[index];
                    if (first.IsNewline)
                    {
                        index++;
                        continue;
                    }

                    int end = index + 1;
                    if (!(extraSpacing > 0 && first.IsSpace))
                    {
                        while (end < atom.Chars.Count && SameRun(first, atom.Chars[end])
                            && !(extraSpacing > 0 && atom.Chars[end].IsSpace))
                        {
                            end++;
                        }
                    }

                    StringBuilder text = new StringBuilder(end - index);
                    double groupWidth = 0;
                    for (int i = index; i < end; i++)
                    {
                        atom.Chars[i].AppendUtf16(text);
                        groupWidth += atom.Chars[i].Width;
                    }

                    double baseline = lineTop + line.Ascent + first.BaselineShift;
                    _page.Add(new PdfTextCommand
                    {
                        Font = first.Font,
                        FontSize = first.FontSize,
                        X = x,
                        Y = baseline,
                        Text = text.ToString(),
                        Color = first.Format.Color,
                        Width = groupWidth
                    });
                    _page.AddFont(first.Font);

                    if (first.Format.Underline)
                    {
                        DrawDecoration(x, baseline + Math.Max(1.0, first.FontSize * 0.1), x + groupWidth, first.Format.Color);
                    }
                    if (first.Format.Strikeout)
                    {
                        DrawDecoration(x, baseline - first.FontSize * 0.3, x + groupWidth, first.Format.Color);
                    }

                    x += groupWidth;
                    if (extraSpacing > 0 && first.IsSpace && end == index + 1)
                    {
                        x += extraSpacing;
                    }
                    index = end;
                }
            }
        }

        /// <summary>
        /// 判断两个相邻字元是否同字体、同字号、同位移、同颜色与装饰，可合并为一次绘制。
        /// Determines whether two adjacent characters share font, size, shift, color, and decorations and can be merged into one draw.
        /// </summary>
        /// <param name="a">前一个字元。The earlier character.</param>
        /// <param name="b">后一个字元。The later character.</param>
        /// <returns>可合并返回 true。True when they can be merged.</returns>
        private static bool SameRun(ShapedChar a, ShapedChar b)
        {
            return ReferenceEquals(a.Font, b.Font)
                && Math.Abs(a.FontSize - b.FontSize) < 0.0001
                && Math.Abs(a.BaselineShift - b.BaselineShift) < 0.0001
                && a.Format.Color.Equals(b.Format.Color)
                && a.Format.Underline == b.Format.Underline
                && a.Format.Strikeout == b.Format.Strikeout
                && !b.IsNewline;
        }

        /// <summary>
        /// 绘制下划线或删除线等装饰线。
        /// Draws a decoration line such as an underline or strikeout.
        /// </summary>
        /// <param name="x1">起点 X。The start X.</param>
        /// <param name="y">线所在的 Y。The Y at which the line sits.</param>
        /// <param name="x2">终点 X。The end X.</param>
        /// <param name="color">线条颜色。The line color.</param>
        private void DrawDecoration(double x1, double y, double x2, OfficeColor color)
        {
            if (x2 <= x1) return;
            _page.Add(new PdfLineCommand
            {
                X1 = x1,
                Y1 = y,
                X2 = x2,
                Y2 = y,
                Width = 0.5,
                Color = color,
                Style = BorderStyle.Single
            });
        }

        /// <summary>
        /// 在给定位置绘制段落并返回其占用高度。
        /// Draws a paragraph at the given position and returns the height it occupies.
        /// </summary>
        /// <param name="paragraph">待绘制的段落。The paragraph to draw.</param>
        /// <param name="x">起始 X。The start X.</param>
        /// <param name="width">可用宽度。The available width.</param>
        /// <param name="top">顶部 Y。The top Y.</param>
        /// <returns>段落占用的总高度。The total height the paragraph occupies.</returns>
        private double DrawParagraphAt(Paragraph paragraph, double x, double width, double top)
        {
            if (paragraph == null) return 0;
            List<PdfLayoutLine> lines = ParagraphLayouter.Wrap(paragraph, width, _defaultFormat, _fonts);
            double y = top + Math.Max(0, paragraph.SpaceBefore);
            for (int i = 0; i < lines.Count; i++)
            {
                DrawTextLine(lines[i], paragraph, x, y, i == lines.Count - 1);
                y += lines[i].Height;
            }
            y += Math.Max(0, paragraph.SpaceAfter);
            return y - top;
        }

        /// <summary>
        /// 测量段落在指定宽度下的高度。
        /// Measures the paragraph's height at the given width.
        /// </summary>
        /// <param name="paragraph">待测量的段落。The paragraph to measure.</param>
        /// <param name="width">可用宽度。The available width.</param>
        /// <returns>含段前段后间距的高度。The height including space-before and space-after.</returns>
        private double MeasureParagraph(Paragraph paragraph, double width)
        {
            if (paragraph == null) return 0;
            List<PdfLayoutLine> lines = ParagraphLayouter.Wrap(paragraph, width, _defaultFormat, _fonts);
            return ParagraphHeight(paragraph, lines);
        }

        /// <summary>
        /// 在给定位置依次绘制内容块列表并返回总高度。
        /// Draws a block list sequentially at the given position and returns the total height.
        /// </summary>
        /// <param name="blocks">内容块列表。The list of blocks.</param>
        /// <param name="x">起始 X。The start X.</param>
        /// <param name="top">顶部 Y。The top Y.</param>
        /// <param name="width">可用宽度。The available width.</param>
        /// <returns>块序列占用的总高度。The total height the sequence occupies.</returns>
        private double DrawBlocksAt(List<Block> blocks, double x, double top, double width)
        {
            double y = top;
            if (blocks == null) return 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    y += DrawParagraphAt(paragraph, x, width, y);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    y += DrawNestedTable(table, x, y, width);
                }
            }
            return y - top;
        }

        /// <summary>
        /// 测量内容块列表在指定宽度下的总高度。
        /// Measures the total height of a block list at the given width.
        /// </summary>
        /// <param name="blocks">内容块列表。The list of blocks.</param>
        /// <param name="width">可用宽度。The available width.</param>
        /// <returns>块序列的总高度。The total height of the sequence.</returns>
        private double MeasureBlocks(List<Block> blocks, double width)
        {
            double height = 0;
            if (blocks == null) return 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    height += MeasureParagraph(paragraph, width);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    height += MeasureTable(table, width);
                }
            }
            return height;
        }

        /// <summary>
        /// 在单元格内绘制块列表：优先复用测量阶段缓存的断行结果。
        /// Draws a block list within a cell, reusing the line breaks cached during measurement.
        /// </summary>
        /// <param name="cell">单元格排版。The laid-out cell.</param>
        /// <param name="x">起始 X。The start X.</param>
        /// <param name="top">顶部 Y。The top Y.</param>
        /// <param name="width">可用宽度。The available width.</param>
        /// <returns>块序列占用的总高度。The total height the sequence occupies.</returns>
        private double DrawCellBlocks(CellLayout cell, double x, double top, double width)
        {
            List<Block> blocks = cell.Cell.Blocks;
            if (blocks == null) return 0;
            List<PdfLayoutLine>[] cache = cell.BlockLines;
            double y = top;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    List<PdfLayoutLine> lines = cache != null && i < cache.Length ? cache[i] : null;
                    y += lines != null
                        ? DrawCachedParagraph(paragraph, lines, x, y)
                        : DrawParagraphAt(paragraph, x, width, y);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    y += DrawNestedTable(table, x, y, width);
                }
            }
            return y - top;
        }

        /// <summary>
        /// 用已排好的行绘制段落并返回高度（与 <see cref="DrawParagraphAt"/> 等价，只是跳过重新断行）。
        /// Draws a paragraph from already-wrapped lines and returns its height; equivalent to
        /// <see cref="DrawParagraphAt"/> but skips re-wrapping.
        /// </summary>
        /// <param name="paragraph">目标段落。The paragraph to draw.</param>
        /// <param name="lines">缓存的断行结果。The cached wrapped lines.</param>
        /// <param name="x">起始 X。The start X.</param>
        /// <param name="top">顶部 Y。The top Y.</param>
        /// <returns>段落占用的高度。The height the paragraph occupies.</returns>
        private double DrawCachedParagraph(Paragraph paragraph, List<PdfLayoutLine> lines, double x, double top)
        {
            double y = top + Math.Max(0, paragraph.SpaceBefore);
            for (int i = 0; i < lines.Count; i++)
            {
                DrawTextLine(lines[i], paragraph, x, y, i == lines.Count - 1);
                y += lines[i].Height;
            }
            y += Math.Max(0, paragraph.SpaceAfter);
            return y - top;
        }

        /// <summary>
        /// 测量单元格内容块的高度，同时缓存段落的断行结果供绘制复用。
        /// Measures a cell's block list and caches each paragraph's wrapped lines for the draw pass.
        /// </summary>
        /// <param name="cell">单元格排版。The laid-out cell.</param>
        /// <param name="width">可用宽度。The available width.</param>
        /// <returns>块序列的总高度。The total height of the sequence.</returns>
        private double MeasureCellBlocks(CellLayout cell, double width)
        {
            List<Block> blocks = cell.Cell.Blocks;
            if (blocks == null) return 0;

            List<PdfLayoutLine>[] cache = new List<PdfLayoutLine>[blocks.Count];
            double height = 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    List<PdfLayoutLine> lines = ParagraphLayouter.Wrap(paragraph, width, _defaultFormat, _fonts);
                    cache[i] = lines;
                    height += ParagraphHeight(paragraph, lines);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    height += MeasureTable(table, width);
                }
            }
            cell.BlockLines = cache;
            return height;
        }

        /// <summary>
        /// 由已断行的结果计算段落高度，含段前段后间距。
        /// Computes a paragraph's height from its wrapped lines, including space-before and space-after.
        /// </summary>
        /// <param name="paragraph">目标段落。The paragraph.</param>
        /// <param name="lines">断行结果。The wrapped lines.</param>
        /// <returns>含间距的高度。The height including spacing.</returns>
        private double ParagraphHeight(Paragraph paragraph, List<PdfLayoutLine> lines)
        {
            double height = Math.Max(0, paragraph.SpaceBefore) + Math.Max(0, paragraph.SpaceAfter);
            for (int i = 0; i < lines.Count; i++)
            {
                height += lines[i].Height;
            }
            if (lines.Count == 0) height += DefaultLineHeight();
            return height;
        }

        /// <summary>
        /// 计算表格各列宽度：结合显式列宽与 ColSpan 撑出的列数，按可用宽度缩放。
        /// Computes the table's column widths: reconciles explicit widths with columns expanded by ColSpan and scales them to the available width.
        /// </summary>
        /// <param name="table">目标表格。The target table.</param>
        /// <param name="totalWidth">表格可用总宽度。The total width available to the table.</param>
        /// <returns>每列的最终宽度。The final width of each column.</returns>
        private double[] ComputeColumnWidths(Table table, double totalWidth)
        {
            int count = table.Columns.Count;
            for (int r = 0; r < table.Rows.Count; r++)
            {
                int span = 0;
                TableRow row = table.Rows[r];
                for (int c = 0; c < row.Cells.Count; c++)
                {
                    span += Math.Max(1, row.Cells[c].ColSpan);
                }
                if (span > count) count = span;
            }
            if (count <= 0) count = 1;
            totalWidth = Math.Max(36, totalWidth);

            double[] widths = new double[count];
            double specified = 0;
            int specifiedCount = 0;
            for (int i = 0; i < count; i++)
            {
                if (i < table.Columns.Count && table.Columns[i].Width > 0)
                {
                    widths[i] = table.Columns[i].Width;
                    specified += widths[i];
                    specifiedCount++;
                }
            }

            if (specifiedCount == 0)
            {
                double equal = totalWidth / count;
                for (int i = 0; i < count; i++) widths[i] = equal;
                return widths;
            }

            double average = specified / specifiedCount;
            double total = 0;
            for (int i = 0; i < count; i++)
            {
                if (widths[i] <= 0) widths[i] = average;
                total += widths[i];
            }
            double scale = totalWidth / total;
            if (!table.Autofit && scale > 1) scale = 1;
            for (int i = 0; i < count; i++)
            {
                widths[i] *= scale;
                if (widths[i] < 6) widths[i] = 6;
            }
            return widths;
        }

        /// <summary>
        /// 构建表格排版：放置单元格并处理合并跨度，测量每行高度并补齐跨行单元格所需空间。
        /// Builds the table layout: places cells handling merge spans, measures each row height, and grows rows to fit row-spanning cells.
        /// </summary>
        /// <param name="table">目标表格。The target table.</param>
        /// <param name="width">表格可用总宽度。The total width available to the table.</param>
        /// <returns>完整的表格排版结果。The complete table layout.</returns>
        private TableLayout BuildTableLayout(Table table, double width)
        {
            TableLayout layout = new TableLayout();
            layout.Table = table;
            layout.Padding = Math.Max(0, table.CellPadding);
            layout.Widths = ComputeColumnWidths(table, width);
            int rowCount = table.Rows.Count;
            int columnCount = layout.Widths.Length;
            double padding = layout.Padding;

            layout.ColumnX = new double[columnCount];
            double cumulative = 0;
            for (int c = 0; c < columnCount; c++)
            {
                layout.ColumnX[c] = cumulative;
                cumulative += layout.Widths[c];
            }

            layout.RowCells = new List<CellLayout>[rowCount];
            layout.Coverage = new CellLayout[rowCount][];
            for (int r = 0; r < rowCount; r++)
            {
                layout.Coverage[r] = new CellLayout[columnCount];
                layout.RowCells[r] = new List<CellLayout>();
            }

            for (int r = 0; r < rowCount; r++)
            {
                TableRow row = table.Rows[r];
                int column = 0;
                int cellIndex = 0;
                while (column < columnCount && cellIndex < row.Cells.Count)
                {
                    if (layout.Coverage[r][column] != null)
                    {
                        column++;
                        continue;
                    }

                    TableCell cell = row.Cells[cellIndex++];
                    if (cell == null)
                    {
                        column++;
                        continue;
                    }

                    int span = Math.Max(1, cell.ColSpan);
                    if (column + span > columnCount) span = columnCount - column;
                    while (span > 1 && HasCoverage(layout.Coverage[r], column, span)) span--;

                    int rowSpan = Math.Max(1, cell.RowSpan);
                    if (r + rowSpan > rowCount) rowSpan = rowCount - r;

                    CellLayout cellLayout = new CellLayout();
                    cellLayout.Cell = cell;
                    cellLayout.Row = r;
                    cellLayout.RowSpan = rowSpan;
                    cellLayout.Column = column;
                    cellLayout.ColumnSpan = span;
                    cellLayout.X = layout.ColumnX[column];
                    double cellWidth = 0;
                    for (int c = column; c < column + span; c++) cellWidth += layout.Widths[c];
                    cellLayout.Width = cellWidth;
                    cellLayout.ContentWidth = Math.Max(1, cellWidth - padding * 2);
                    cellLayout.ContentHeight = MeasureCellBlocks(cellLayout, cellLayout.ContentWidth);
                    layout.RowCells[r].Add(cellLayout);

                    for (int rr = r; rr < r + rowSpan; rr++)
                    {
                        for (int cc = column; cc < column + span; cc++)
                        {
                            layout.Coverage[rr][cc] = cellLayout;
                        }
                    }
                    column += span;
                }
            }

            layout.Heights = new double[rowCount];
            for (int r = 0; r < rowCount; r++)
            {
                // 有内容的行按实际内容测量；只有空单元格（没有段落）才退回默认行高，避免小字号或
                // 显式小行高的表格被 11 磅默认行高撑大。
                // Rows with content use the measured content; only empty cells (no paragraphs) fall back
                // to the default line height, so small fonts or explicit small row heights are not inflated
                // by the 11-point document default.
                double height = 0;
                bool hasContentCell = false;
                List<CellLayout> cells = layout.RowCells[r];
                for (int i = 0; i < cells.Count; i++)
                {
                    if (cells[i].RowSpan != 1) continue;
                    hasContentCell = true;
                    double cellHeight = cells[i].ContentHeight + padding * 2;
                    if (cells[i].ContentHeight <= 0.001)
                    {
                        double emptyHeight = DefaultLineHeight() + padding * 2;
                        if (emptyHeight > cellHeight) cellHeight = emptyHeight;
                    }
                    if (cellHeight > height) height = cellHeight;
                }
                if (!hasContentCell) height = DefaultLineHeight() + padding * 2;
                TableRow row = table.Rows[r];
                if (row.Height > height) height = row.Height;
                layout.Heights[r] = Math.Max(1, height);
            }

            for (int r = 0; r < rowCount; r++)
            {
                List<CellLayout> cells = layout.RowCells[r];
                for (int i = 0; i < cells.Count; i++)
                {
                    CellLayout cell = cells[i];
                    if (cell.RowSpan <= 1) continue;
                    double required = cell.ContentHeight + padding * 2;
                    double available = 0;
                    for (int rr = cell.Row; rr < cell.Row + cell.RowSpan; rr++) available += layout.Heights[rr];
                    if (required > available)
                    {
                        double extra = (required - available) / cell.RowSpan;
                        for (int rr = cell.Row; rr < cell.Row + cell.RowSpan; rr++) layout.Heights[rr] += extra;
                    }
                }
            }

            layout.TotalHeight = 0;
            for (int r = 0; r < rowCount; r++) layout.TotalHeight += layout.Heights[r];
            return layout;
        }

        /// <summary>
        /// 判断给定跨度区间内是否已有单元格被覆盖，用于收缩 ColSpan。
        /// Determines whether a cell is already covered within the given span, used to shrink a ColSpan.
        /// </summary>
        /// <param name="row">该行的覆盖数组。The row's coverage array.</param>
        /// <param name="column">起始列。The starting column.</param>
        /// <param name="span">候选跨度。The candidate span.</param>
        /// <returns>区间内有已覆盖单元格返回 true。True when the range contains a covered cell.</returns>
        private static bool HasCoverage(CellLayout[] row, int column, int span)
        {
            for (int c = column + 1; c < column + span; c++)
            {
                if (row[c] != null) return true;
            }
            return false;
        }

        /// <summary>
        /// 求从某行起因 RowSpan 相互牵连而必须整体保持的行组结束下标。
        /// Finds the end row index of the group that must be kept together from a start row due to interlocking RowSpans.
        /// </summary>
        /// <param name="layout">表格排版。The table layout.</param>
        /// <param name="startRow">起始行。The starting row.</param>
        /// <returns>行组的最后一行下标（夹到表末）。The last row index of the group, clamped to the table's end.</returns>
        private int GetSpanGroupEnd(TableLayout layout, int startRow)
        {
            int end = startRow;
            bool expanded = true;
            while (expanded)
            {
                expanded = false;
                for (int r = startRow; r <= end; r++)
                {
                    List<CellLayout> cells = layout.RowCells[r];
                    for (int i = 0; i < cells.Count; i++)
                    {
                        int cellEnd = cells[i].Row + cells[i].RowSpan - 1;
                        if (cellEnd > end)
                        {
                            end = cellEnd;
                            expanded = true;
                        }
                    }
                }
            }
            return Math.Min(end, layout.RowCells.Length - 1);
        }

        /// <summary>
        /// 跨页绘制表格：按行组决定是否换页，并在换页时按设置重复表头。
        /// Draws a table across pages: decides page breaks by row group and repeats the header on each new page when configured.
        /// </summary>
        /// <param name="table">目标表格。The target table.</param>
        private void DrawTableFlow(Table table)
        {
            if (table == null || table.Rows.Count == 0) return;
            TableLayout layout = BuildTableLayout(table, _contentWidth);
            layout.OriginX = _contentLeft;
            if (table.Alignment == TableAlignment.Center)
            {
                double tableWidth = 0;
                for (int c = 0; c < layout.Widths.Length; c++) tableWidth += layout.Widths[c];
                double shift = (_contentWidth - tableWidth) / 2;
                if (shift > 0) layout.OriginX += shift;
            }

            int headerCount = 0;
            while (headerCount < table.Rows.Count && table.Rows[headerCount].IsHeader)
            {
                headerCount++;
            }

            int rowIndex = 0;
            int noBreakUntil = -1;
            double pageHeight = Math.Max(1, _contentBottom - _contentTop);
            while (rowIndex < table.Rows.Count)
            {
                double height = layout.Heights[rowIndex];
                bool mayBreak = rowIndex > noBreakUntil && _section.AllowAutoPageBreak;
                if (mayBreak && _cursorY > _contentTop + 0.01)
                {
                    int groupEnd = GetSpanGroupEnd(layout, rowIndex);
                    double required = 0;
                    for (int r = rowIndex; r <= groupEnd; r++) required += layout.Heights[r];
                    if (required > pageHeight)
                    {
                        groupEnd = rowIndex;
                        required = height;
                    }
                    if (_cursorY + required > _contentBottom + 0.01)
                    {
                        StartNewPage();
                        if (table.RepeatHeaderRow && headerCount > 0 && rowIndex >= headerCount)
                        {
                            DrawRepeatedHeaderRows(layout, headerCount);
                        }
                        noBreakUntil = groupEnd;
                        continue;
                    }
                    noBreakUntil = groupEnd;
                }

                DrawTableRow(layout, rowIndex, layout.OriginX, _cursorY);
                _cursorY += height;
                rowIndex++;
            }
        }

        /// <summary>
        /// 在新页顶部重绘表头行，空间不足时停止。
        /// Redraws the header rows at the top of a new page, stopping when space runs out.
        /// </summary>
        /// <param name="layout">表格排版。The table layout.</param>
        /// <param name="headerCount">表头行数。The number of header rows.</param>
        private void DrawRepeatedHeaderRows(TableLayout layout, int headerCount)
        {
            for (int h = 0; h < headerCount; h++)
            {
                double height = layout.Heights[h];
                if (_cursorY + height > _contentBottom + 0.01 && _cursorY > _contentTop + 0.01)
                {
                    break;
                }
                DrawTableRow(layout, h, layout.OriginX, _cursorY);
                _cursorY += height;
            }
        }

        /// <summary>
        /// 绘制一行：先填充背景，再解析每格四边的边框来源并绘制，最后按垂直对齐放置单元格内容。
        /// Draws one row: fills backgrounds, resolves each cell's four border sources and paints them, then places cell content by vertical alignment.
        /// </summary>
        /// <param name="layout">表格排版。The table layout.</param>
        /// <param name="rowIndex">要绘制的行下标。The row index to draw.</param>
        /// <param name="x">表格左边绝对 X。The table's absolute left X.</param>
        /// <param name="top">该行的顶部 Y。The row's top Y.</param>
        private void DrawTableRow(TableLayout layout, int rowIndex, double x, double top)
        {
            Table table = layout.Table;
            double padding = layout.Padding;
            double rowHeight = layout.Heights[rowIndex];
            int columnCount = layout.Widths.Length;
            CellLayout[] coverage = layout.Coverage[rowIndex];

            for (int c = 0; c < columnCount; )
            {
                CellLayout covered = coverage[c];
                int span = 1;
                while (c + span < columnCount && coverage[c + span] == covered) span++;
                if (covered != null)
                {
                    OfficeColor? background = covered.Cell.BackgroundColor;
                    if (background.HasValue && background.Value.A > 0)
                    {
                        double width = 0;
                        for (int i = c; i < c + span; i++) width += layout.Widths[i];
                        _page.Add(new PdfFillRectCommand
                        {
                            X = x + layout.ColumnX[c],
                            Y = top,
                            Width = width,
                            Height = rowHeight,
                            Color = background.Value
                        });
                    }
                }
                c += span;
            }

            TableBorders tableBorders = table.Borders ?? new TableBorders();
            List<CellLayout> cells = layout.RowCells[rowIndex];
            for (int i = 0; i < cells.Count; i++)
            {
                CellLayout cellLayout = cells[i];
                TableCellBorders cellBorders = cellLayout.Cell.Borders;
                double right = x + cellLayout.X + cellLayout.Width;
                double cellTop = top;
                double cellHeight = 0;
                int lastRow = cellLayout.Row + cellLayout.RowSpan - 1;
                for (int r = cellLayout.Row; r <= lastRow; r++) cellHeight += layout.Heights[r];
                double cellBottom = cellTop + cellHeight;

                BorderLine topLine = cellBorders != null && cellBorders.Top != null
                    ? cellBorders.Top
                    : (rowIndex == 0 ? tableBorders.Top : tableBorders.InsideHorizontal);
                BorderLine bottomLine = cellBorders != null && cellBorders.Bottom != null
                    ? cellBorders.Bottom
                    : (lastRow == table.Rows.Count - 1 ? tableBorders.Bottom : tableBorders.InsideHorizontal);
                BorderLine leftLine = cellBorders != null && cellBorders.Left != null
                    ? cellBorders.Left
                    : (cellLayout.Column == 0 ? tableBorders.Left : tableBorders.InsideVertical);
                BorderLine rightLine = cellBorders != null && cellBorders.Right != null
                    ? cellBorders.Right
                    : (cellLayout.Column + cellLayout.ColumnSpan >= columnCount ? tableBorders.Right : tableBorders.InsideVertical);

                DrawBorderLine(topLine, x + cellLayout.X, cellTop, right, cellTop);
                DrawBorderLine(bottomLine, x + cellLayout.X, cellBottom, right, cellBottom);
                DrawBorderLine(leftLine, x + cellLayout.X, cellTop, x + cellLayout.X, cellBottom);
                DrawBorderLine(rightLine, right, cellTop, right, cellBottom);
            }

            for (int i = 0; i < cells.Count; i++)
            {
                CellLayout cellLayout = cells[i];
                double availableHeight = 0;
                int lastRow = cellLayout.Row + cellLayout.RowSpan - 1;
                for (int r = cellLayout.Row; r <= lastRow; r++) availableHeight += layout.Heights[r];
                availableHeight = Math.Max(0, availableHeight - padding * 2);
                double offset = 0;
                if (cellLayout.Cell.VerticalAlignment == VerticalAlignment.Middle)
                {
                    offset = Math.Max(0, (availableHeight - cellLayout.ContentHeight) / 2);
                }
                else if (cellLayout.Cell.VerticalAlignment == VerticalAlignment.Bottom)
                {
                    offset = Math.Max(0, availableHeight - cellLayout.ContentHeight);
                }
                DrawCellBlocks(cellLayout, x + cellLayout.X + padding, top + padding + offset, cellLayout.ContentWidth);
            }
        }

        /// <summary>
        /// 绘制一条边框线：双线样式用两条偏移单线合成，None 样式跳过。
        /// Draws a border line: a double style is synthesized from two offset single lines, and a None style is skipped.
        /// </summary>
        /// <param name="border">边框定义。The border definition.</param>
        /// <param name="x1">起点 X。The start X.</param>
        /// <param name="y1">起点 Y。The start Y.</param>
        /// <param name="x2">终点 X。The end X.</param>
        /// <param name="y2">终点 Y。The end Y.</param>
        private void DrawBorderLine(BorderLine border, double x1, double y1, double x2, double y2)
        {
            if (border == null || border.Style == BorderStyle.None) return;
            double width = border.Width > 0 ? border.Width : 0.5;
            if (border.Style == BorderStyle.Double)
            {
                bool horizontal = Math.Abs(y1 - y2) < 0.01;
                double offset = Math.Max(1.0, width);
                if (horizontal)
                {
                    AddBorderCommand(x1, y1 - offset / 2, x2, y2 - offset / 2, width, border.Color, BorderStyle.Single);
                    AddBorderCommand(x1, y1 + offset / 2, x2, y2 + offset / 2, width, border.Color, BorderStyle.Single);
                }
                else
                {
                    AddBorderCommand(x1 - offset / 2, y1, x2 - offset / 2, y2, width, border.Color, BorderStyle.Single);
                    AddBorderCommand(x1 + offset / 2, y1, x2 + offset / 2, y2, width, border.Color, BorderStyle.Single);
                }
                return;
            }
            AddBorderCommand(x1, y1, x2, y2, width, border.Color, border.Style);
        }

        /// <summary>
        /// 向当前页面追加一条线条绘制命令。
        /// Appends a line draw command to the current page.
        /// </summary>
        /// <param name="x1">起点 X。The start X.</param>
        /// <param name="y1">起点 Y。The start Y.</param>
        /// <param name="x2">终点 X。The end X.</param>
        /// <param name="y2">终点 Y。The end Y.</param>
        /// <param name="width">线宽。The line width.</param>
        /// <param name="color">线条颜色。The line color.</param>
        /// <param name="style">线条样式。The line style.</param>
        private void AddBorderCommand(double x1, double y1, double x2, double y2, double width, OfficeColor color, BorderStyle style)
        {
            _page.Add(new PdfLineCommand
            {
                X1 = x1,
                Y1 = y1,
                X2 = x2,
                Y2 = y2,
                Width = width,
                Color = color,
                Style = style
            });
        }

        /// <summary>
        /// 在给定位置绘制嵌套（单元格内）表格并返回其高度。
        /// Draws a nested (in-cell) table at the given position and returns its height.
        /// </summary>
        /// <param name="table">目标表格。The target table.</param>
        /// <param name="x">起始 X。The start X.</param>
        /// <param name="top">顶部 Y。The top Y.</param>
        /// <param name="width">可用宽度。The available width.</param>
        /// <returns>嵌套表格占用的总高度。The total height the nested table occupies.</returns>
        private double DrawNestedTable(Table table, double x, double top, double width)
        {
            if (table == null || table.Rows.Count == 0) return 0;
            TableLayout layout = BuildTableLayout(table, width);
            double y = top;
            for (int r = 0; r < table.Rows.Count; r++)
            {
                DrawTableRow(layout, r, x, y);
                y += layout.Heights[r];
            }
            return y - top;
        }

        /// <summary>
        /// 测量表格在指定宽度下的总高度。
        /// Measures a table's total height at the given width.
        /// </summary>
        /// <param name="table">目标表格。The target table.</param>
        /// <param name="width">可用宽度。The available width.</param>
        /// <returns>表格总高度。The total table height.</returns>
        private double MeasureTable(Table table, double width)
        {
            if (table == null || table.Rows.Count == 0) return 0;
            return BuildTableLayout(table, width).TotalHeight;
        }

        /// <summary>
        /// 排版完成后为每页绘制其节的页眉与页脚。
        /// After layout, draws each page's section header and footer.
        /// </summary>
        private void DrawHeadersAndFooters()
        {
            int totalPages = _pages.Count;
            for (int i = 0; i < _pages.Count; i++)
            {
                PdfPageContent page = _pages[i];
                Section section = page.Section;
                if (section == null) continue;
                if (section.Header != null && section.Header.IsEnabled)
                {
                    DrawHeaderFooter(page, section.Header, true, totalPages);
                }
                if (section.Footer != null && section.Footer.IsEnabled)
                {
                    DrawHeaderFooter(page, section.Footer, false, totalPages);
                }
            }
        }

        /// <summary>
        /// 在页边区域绘制页眉或页脚，临时切换当前页与节后再恢复。
        /// Draws a header or footer in the margin area, temporarily switching the current page and section before restoring them.
        /// </summary>
        /// <param name="page">目标页面。The target page.</param>
        /// <param name="content">页眉/页脚定义。The header/footer definition.</param>
        /// <param name="isHeader">true 为页眉，false 为页脚。True for a header, false for a footer.</param>
        /// <param name="totalPages">总页数，用于替换 {pages}。The total page count used to replace {pages}.</param>
        private void DrawHeaderFooter(PdfPageContent page, HeaderFooter content, bool isHeader, int totalPages)
        {
            Section section = page.Section;
            PageSetup setup = section.PageSetup ?? new PageSetup();
            double width = Math.Max(36, page.Width - setup.MarginLeft - setup.MarginRight);
            List<Block> blocks = CloneBlocksWithTokens(content.Blocks, page.PageNumber, totalPages);
            double height = MeasureBlocks(blocks, width);
            if (height <= 0) return;

            double top = isHeader
                ? Math.Max(0, setup.HeaderDistance)
                : page.Height - Math.Max(0, setup.FooterDistance) - height;

            PdfPageContent savedPage = _page;
            Section savedSection = _section;
            _page = page;
            _section = section;
            try
            {
                DrawBlocksAt(blocks, Math.Max(0, setup.MarginLeft), top, width);
            }
            finally
            {
                _page = savedPage;
                _section = savedSection;
            }
        }

        /// <summary>
        /// 克隆页眉页脚块并把页码占位符替换为当前值。
        /// Clones header/footer blocks and substitutes page-number tokens with the current values.
        /// </summary>
        /// <param name="blocks">原始块列表。The original block list.</param>
        /// <param name="pageNumber">当前页码。The current page number.</param>
        /// <param name="totalPages">总页数。The total page count.</param>
        /// <returns>替换占位符后的新块列表。A new block list with tokens replaced.</returns>
        private List<Block> CloneBlocksWithTokens(List<Block> blocks, int pageNumber, int totalPages)
        {
            List<Block> result = new List<Block>();
            if (blocks == null) return result;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    result.Add(CloneParagraphWithTokens(paragraph, pageNumber, totalPages));
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table != null)
                {
                    result.Add(CloneTableWithTokens(table, pageNumber, totalPages));
                }
            }
            return result;
        }

        /// <summary>
        /// 克隆段落并在其文本域中替换页码占位符。
        /// Clones a paragraph and replaces page-number tokens in its text fields.
        /// </summary>
        /// <param name="paragraph">原始段落。The source paragraph.</param>
        /// <param name="pageNumber">当前页码。The current page number.</param>
        /// <param name="totalPages">总页数。The total page count.</param>
        /// <returns>替换后的段落副本。The copied paragraph with tokens replaced.</returns>
        private Paragraph CloneParagraphWithTokens(Paragraph paragraph, int pageNumber, int totalPages)
        {
            Paragraph copy = paragraph.Clone();
            for (int i = 0; i < copy.Inlines.Count; i++)
            {
                TextRun run = copy.Inlines[i] as TextRun;
                if (run != null)
                {
                    run.Text = ReplacePageTokens(run.Text, pageNumber, totalPages);
                    continue;
                }
                FieldRun field = copy.Inlines[i] as FieldRun;
                if (field != null)
                {
                    field.DisplayText = ReplacePageTokens(field.DisplayText, pageNumber, totalPages);
                }
            }
            return copy;
        }

        /// <summary>
        /// 克隆表格（含嵌套表）并在各单元格段落中替换页码占位符。
        /// Clones a table (including nested tables) and replaces page-number tokens in each cell's paragraphs.
        /// </summary>
        /// <param name="table">原始表格。The source table.</param>
        /// <param name="pageNumber">当前页码。The current page number.</param>
        /// <param name="totalPages">总页数。The total page count.</param>
        /// <returns>替换后的表格副本。The copied table with tokens replaced.</returns>
        private Table CloneTableWithTokens(Table table, int pageNumber, int totalPages)
        {
            Table copy = new Table();
            copy.Borders = table.Borders;
            copy.CellPadding = table.CellPadding;
            copy.RepeatHeaderRow = table.RepeatHeaderRow;
            for (int i = 0; i < table.Columns.Count; i++)
            {
                copy.Columns.Add(new TableColumn(table.Columns[i].Width));
            }
            for (int r = 0; r < table.Rows.Count; r++)
            {
                TableRow sourceRow = table.Rows[r];
                TableRow row = new TableRow();
                row.IsHeader = sourceRow.IsHeader;
                row.Height = sourceRow.Height;
                for (int c = 0; c < sourceRow.Cells.Count; c++)
                {
                    TableCell sourceCell = sourceRow.Cells[c];
                    TableCell cell = new TableCell();
                    cell.ColSpan = sourceCell.ColSpan;
                    cell.RowSpan = sourceCell.RowSpan;
                    cell.BackgroundColor = sourceCell.BackgroundColor;
                    cell.Borders = sourceCell.Borders;
                    cell.VerticalAlignment = sourceCell.VerticalAlignment;
                    for (int b = 0; b < sourceCell.Blocks.Count; b++)
                    {
                        Paragraph paragraph = sourceCell.Blocks[b] as Paragraph;
                        if (paragraph != null)
                        {
                            cell.Blocks.Add(CloneParagraphWithTokens(paragraph, pageNumber, totalPages));
                            continue;
                        }
                        Table nested = sourceCell.Blocks[b] as Table;
                        if (nested != null)
                        {
                            cell.Blocks.Add(CloneTableWithTokens(nested, pageNumber, totalPages));
                        }
                    }
                    row.Cells.Add(cell);
                }
                copy.Rows.Add(row);
            }
            return copy;
        }

        /// <summary>
        /// 把文本中的 {page} 与 {pages} 替换为当前页码与总页数。
        /// Replaces {page} and {pages} in the text with the current page number and total count.
        /// </summary>
        /// <param name="text">原始文本。The source text.</param>
        /// <param name="pageNumber">当前页码。The current page number.</param>
        /// <param name="totalPages">总页数。The total page count.</param>
        /// <returns>替换后的文本；不含左花括号时原样返回。The replaced text; returned unchanged when it has no brace.</returns>
        private static string ReplacePageTokens(string text, int pageNumber, int totalPages)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOf('{') < 0) return text ?? string.Empty;
            string result = PageToken.Replace(text, pageNumber.ToString(CultureInfo.InvariantCulture));
            result = PagesToken.Replace(result, totalPages.ToString(CultureInfo.InvariantCulture));
            return result;
        }

        /// <summary>
        /// 默认格式下的标准行高，作为空内容的兜底高度。
        /// The standard line height under the default format, used as a fallback height for empty content.
        /// </summary>
        /// <returns>行高（点）。The line height in points.</returns>
        private double DefaultLineHeight()
        {
            PdfFont font = _fonts.GetPrimary(_defaultFormat);
            double size = _defaultFormat.FontSize > 0 ? _defaultFormat.FontSize : 11;
            return Math.Max(size * 1.15, font.GetLineHeight(size));
        }

        /// <summary>
        /// 序列化文档：准备字体与图片对象、逐页构建资源与内容流，最后写入目录、信息并输出。
        /// Serializes the document: prepares font and image objects, builds each page's resources and content stream, then writes the catalog, info, and trailer.
        /// </summary>
        /// <param name="output">接收 PDF 的流。The stream receiving the PDF.</param>
        private void WriteDocument(Stream output)
        {
            PdfObjectWriter writer = new PdfObjectWriter();

            PdfDictionary fontResources = new PdfDictionary();
            _fonts.PrepareAll(writer, fontResources);

            PdfDictionary imageResources = new PdfDictionary();
            foreach (PdfImage image in _images.All)
            {
                image.Prepare(writer, imageResources);
            }

            PdfDictionary pagesDictionary = new PdfDictionary();
            pagesDictionary.SetName("Type", "Pages");
            PdfArray kids = new PdfArray();
            pagesDictionary.Set("Kids", kids);
            pagesDictionary.SetNumber("Count", _pages.Count);
            PdfReference pagesReference = writer.Add(pagesDictionary);

            for (int i = 0; i < _pages.Count; i++)
            {
                PdfPageContent page = _pages[i];
                PdfDictionary resources = new PdfDictionary();

                PdfDictionary pageFonts = new PdfDictionary();
                foreach (PdfFont font in page.Fonts)
                {
                    PdfObject reference = fontResources.Get(font.ResourceName);
                    if (reference != null) pageFonts.Set(font.ResourceName, reference);
                }
                if (pageFonts.Count > 0) resources.Set("Font", pageFonts);

                PdfDictionary pageImages = new PdfDictionary();
                foreach (PdfImage image in page.Images)
                {
                    PdfObject reference = imageResources.Get(image.ResourceName);
                    if (reference != null) pageImages.Set(image.ResourceName, reference);
                }
                if (pageImages.Count > 0) resources.Set("XObject", pageImages);

                PdfDictionary pageDictionary = new PdfDictionary();
                pageDictionary.SetName("Type", "Page");
                pageDictionary.Set("Parent", pagesReference);
                pageDictionary.Set("MediaBox", new PdfArray()
                    .AddNumber(0).AddNumber(0).AddNumber(page.Width).AddNumber(page.Height));
                pageDictionary.Set("Resources", resources);
                pageDictionary.Set("Contents", writer.AddStream(page.BuildContentStream(), true));
                kids.Add(writer.Add(pageDictionary));
            }

            PdfDictionary catalog = new PdfDictionary();
            catalog.SetName("Type", "Catalog");
            catalog.Set("Pages", pagesReference);
            PdfReference catalogReference = writer.Add(catalog);

            DocumentMetadata metadata = _document.Metadata ?? new DocumentMetadata();
            PdfReference infoReference = writer.Add(CreateInfo(metadata));
            writer.Write(output, catalogReference, infoReference);
        }

        /// <summary>
        /// 由元数据构建 PDF Info 字典（含标题、作者、日期等）。
        /// Builds the PDF Info dictionary from the metadata (title, author, dates, and so on).
        /// </summary>
        /// <param name="metadata">文档元数据。The document metadata.</param>
        /// <returns>信息字典对象。The information dictionary object.</returns>
        private static PdfDictionary CreateInfo(DocumentMetadata metadata)
        {
            PdfDictionary info = new PdfDictionary();
            if (!string.IsNullOrEmpty(metadata.Title)) info.SetString("Title", metadata.Title);
            if (!string.IsNullOrEmpty(metadata.Author)) info.SetString("Author", metadata.Author);
            if (!string.IsNullOrEmpty(metadata.Subject)) info.SetString("Subject", metadata.Subject);
            if (!string.IsNullOrEmpty(metadata.Keywords)) info.SetString("Keywords", metadata.Keywords);
            info.SetString("Creator", string.IsNullOrEmpty(metadata.Creator) ? "Wisdom.Office" : metadata.Creator);
            info.SetString("Producer", "Wisdom.Office PDF");
            info.Set("CreationDate", PdfString.Literal(PdfObjectWriter.FormatPdfDate(metadata.Created)));
            info.Set("ModDate", PdfString.Literal(PdfObjectWriter.FormatPdfDate(metadata.Modified)));
            return info;
        }
    }
}
