using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>页面绘制指令的抽象基类。Abstract base class for a single drawing command on a page.</summary>
    internal abstract class PdfDrawCommand
    {
    }

    /// <summary>绘制一段文本的指令。A command that draws a run of text.</summary>
    internal sealed class PdfTextCommand : PdfDrawCommand
    {
        /// <summary>使用的字体。The font used for this run.</summary>
        public PdfFont Font;
        /// <summary>字号（PDF 用户单位）。Font size in PDF user units.</summary>
        public double FontSize;
        /// <summary>文本起点横坐标（左上原点坐标系）。X coordinate of the text origin in the top-left coordinate system.</summary>
        public double X;
        /// <summary>文本起点纵坐标（左上原点坐标系）。Y coordinate of the text origin in the top-left coordinate system.</summary>
        public double Y;
        /// <summary>待绘制的文本内容。The text to draw.</summary>
        public string Text;
        /// <summary>文字填充颜色。The fill color of the glyphs.</summary>
        public OfficeColor Color;
        /// <summary>文本测量宽度。The measured width of the text.</summary>
        public double Width;
    }

    /// <summary>绘制一条线段的指令，兼作表格边框。A command that draws a line segment, also used for table borders.</summary>
    internal sealed class PdfLineCommand : PdfDrawCommand
    {
        /// <summary>起点横坐标。X coordinate of the start point.</summary>
        public double X1;
        /// <summary>起点纵坐标。Y coordinate of the start point.</summary>
        public double Y1;
        /// <summary>终点横坐标。X coordinate of the end point.</summary>
        public double X2;
        /// <summary>终点纵坐标。Y coordinate of the end point.</summary>
        public double Y2;
        /// <summary>线宽。The stroke width.</summary>
        public double Width;
        /// <summary>线条颜色。The stroke color.</summary>
        public OfficeColor Color;
        /// <summary>线型（实线、虚线、点线）。The line style (solid, dashed or dotted).</summary>
        public BorderStyle Style;
    }

    /// <summary>绘制实心矩形的指令，用于单元格底纹等背景。A command that fills a rectangle, used for cell shading and other backgrounds.</summary>
    internal sealed class PdfFillRectCommand : PdfDrawCommand
    {
        /// <summary>左上角横坐标。X coordinate of the top-left corner.</summary>
        public double X;
        /// <summary>左上角纵坐标。Y coordinate of the top-left corner.</summary>
        public double Y;
        /// <summary>矩形宽度。The width of the rectangle.</summary>
        public double Width;
        /// <summary>矩形高度。The height of the rectangle.</summary>
        public double Height;
        /// <summary>填充颜色。The fill color.</summary>
        public OfficeColor Color;
    }

    /// <summary>绘制图像的指令。A command that draws an image.</summary>
    internal sealed class PdfImageCommand : PdfDrawCommand
    {
        /// <summary>待绘制的图像对象。The image object to draw.</summary>
        public PdfImage Image;
        /// <summary>放置位置左上角横坐标。X coordinate of the top-left placement corner.</summary>
        public double X;
        /// <summary>放置位置左上角纵坐标。Y coordinate of the top-left placement corner.</summary>
        public double Y;
        /// <summary>绘制宽度。The rendered width.</summary>
        public double Width;
        /// <summary>绘制高度。The rendered height.</summary>
        public double Height;
    }

    /// <summary>几何形状类型，决定 <see cref="PdfShapeCommand"/> 生成的路径。Geometry kind that decides the path emitted by <see cref="PdfShapeCommand"/>.</summary>
    internal enum PdfShapeKind
    {
        /// <summary>矩形。Rectangle.</summary>
        Rectangle = 0,
        /// <summary>圆角矩形。Rounded rectangle.</summary>
        RoundedRectangle = 1,
        /// <summary>椭圆。Ellipse.</summary>
        Ellipse = 2,
        /// <summary>多边形（顶点由 <see cref="PdfShapeCommand.Points"/> 给出）。Polygon (vertices from <see cref="PdfShapeCommand.Points"/>).</summary>
        Polygon = 3
    }

    /// <summary>
    /// 绘制几何形状的指令：填充与描边可各自省略，用于幻灯片里的色块、圆角矩形与装饰多边形。
    /// A command that draws a geometry with optional fill and stroke, used for slide blocks, rounded
    /// rectangles and decorative polygons.
    /// </summary>
    internal sealed class PdfShapeCommand : PdfDrawCommand
    {
        /// <summary>形状类型。The geometry kind.</summary>
        public PdfShapeKind Kind;
        /// <summary>外接矩形左上角横坐标。X coordinate of the bounding box's top-left corner.</summary>
        public double X;
        /// <summary>外接矩形左上角纵坐标。Y coordinate of the bounding box's top-left corner.</summary>
        public double Y;
        /// <summary>外接矩形宽度。The bounding box width.</summary>
        public double Width;
        /// <summary>外接矩形高度。The bounding box height.</summary>
        public double Height;
        /// <summary>圆角半径（磅），仅圆角矩形使用。Corner radius in points, rounded rectangles only.</summary>
        public double CornerRadius;
        /// <summary>多边形顶点，成对的 x,y（左上原点坐标系）。Polygon vertices as x,y pairs in the top-left coordinate system.</summary>
        public double[] Points;
        /// <summary>填充色，null 表示不填充。Fill color; null means no fill.</summary>
        public OfficeColor? FillColor;
        /// <summary>描边色，null 表示不描边。Stroke color; null means no stroke.</summary>
        public OfficeColor StrokeColor;
        /// <summary>描边宽度（磅）。The stroke width in points.</summary>
        public double StrokeWidth;
        /// <summary>描边线型。The stroke style.</summary>
        public BorderStyle StrokeStyle;
    }

    /// <summary>一页渲染结果的绘制指令集合，可序列化为 PDF 内容流。The collection of drawing commands for one rendered page, serializable into a PDF content stream.</summary>
    /// <remarks>
    /// 内部坐标采用左上原点、Y 轴向下的约定，生成内容流时统一翻转回 PDF 的左下原点坐标系。
    /// Coordinates inside the page use a top-left origin with the Y axis pointing down; they are flipped
    /// back to the PDF bottom-left origin when the content stream is built.
    /// </remarks>
    internal sealed class PdfPageContent
    {
        /// <summary>按添加顺序存放的绘制指令。Drawing commands in insertion order.</summary>
        private readonly List<PdfDrawCommand> _commands = new List<PdfDrawCommand>();
        /// <summary>本页引用到的去重字体集合。Deduplicated set of fonts referenced by this page.</summary>
        private readonly HashSet<PdfFont> _fonts = new HashSet<PdfFont>();
        /// <summary>本页引用到的去重图像集合。Deduplicated set of images referenced by this page.</summary>
        private readonly HashSet<PdfImage> _images = new HashSet<PdfImage>();

        /// <summary>页面宽度（PDF 用户单位）。Page width in PDF user units.</summary>
        public double Width { get; private set; }
        /// <summary>页面高度（PDF 用户单位）。Page height in PDF user units.</summary>
        public double Height { get; private set; }
        /// <summary>页面所属的节，提供页面设置来源。The section this page belongs to, providing the page-setup source.</summary>
        public Section Section { get; private set; }
        /// <summary>页码，写入时回填。The page number, filled in at write time.</summary>
        public int PageNumber { get; set; }
        /// <summary>全部绘制指令。All drawing commands on the page.</summary>
        public IReadOnlyList<PdfDrawCommand> Commands { get { return _commands; } }
        /// <summary>本页引用到的字体。The fonts referenced by this page.</summary>
        public IEnumerable<PdfFont> Fonts { get { return _fonts; } }
        /// <summary>本页引用到的图像。The images referenced by this page.</summary>
        public IEnumerable<PdfImage> Images { get { return _images; } }

        /// <summary>
        /// 创建指定尺寸的页面内容。
        /// Creates page content with the given size.
        /// </summary>
        /// <param name="width">页面宽度（PDF 用户单位）。Page width in PDF user units.</param>
        /// <param name="height">页面高度（PDF 用户单位）。Page height in PDF user units.</param>
        /// <param name="section">所属节。The owning section.</param>
        public PdfPageContent(double width, double height, Section section)
        {
            Width = width;
            Height = height;
            Section = section;
        }

        /// <summary>
        /// 追加一条绘制指令，null 会被忽略。
        /// Appends a drawing command; null is ignored.
        /// </summary>
        /// <param name="command">绘制指令。The command to append.</param>
        public void Add(PdfDrawCommand command)
        {
            if (command != null) _commands.Add(command);
        }

        /// <summary>
        /// 登记本页使用到的字体，供资源字典写出。
        /// Registers a font used by this page so it appears in the resource dictionary.
        /// </summary>
        /// <param name="font">字体对象，null 会被忽略。The font object; null is ignored.</param>
        public void AddFont(PdfFont font)
        {
            if (font != null) _fonts.Add(font);
        }

        /// <summary>
        /// 登记本页使用到的图像，供资源字典写出。
        /// Registers an image used by this page so it appears in the resource dictionary.
        /// </summary>
        /// <param name="image">图像对象，null 会被忽略。The image object; null is ignored.</param>
        public void AddImage(PdfImage image)
        {
            if (image != null) _images.Add(image);
        }

        /// <summary>
        /// 把全部指令序列化为 PDF 内容流字节。
        /// Serializes all commands into PDF content stream bytes.
        /// </summary>
        /// <returns>ASCII 编码的内容流字节；Y 坐标在此翻转成 PDF 的左下原点约定。The ASCII-encoded content stream bytes; Y coordinates are flipped here to the PDF bottom-left origin convention.</returns>
        public byte[] BuildContentStream()
        {
            StringBuilder builder = new StringBuilder(Math.Max(64, _commands.Count * 40));
            char[] scratch = new char[24];

            // PDF 图形状态在整页内延续，相同颜色、线宽、线型与字体只需在变化时重新设置。
            // PDF graphics state persists across the page, so repeated colors, widths, dash patterns and
            // fonts only need to be emitted when they change.
            int fillKey = -1;
            int strokeKey = -1;
            double lastWidth = double.NaN;
            int lastDash = -1;
            PdfFont lastFont = null;
            double lastFontSize = double.NaN;

            for (int i = 0; i < _commands.Count; i++)
            {
                PdfTextCommand text = _commands[i] as PdfTextCommand;
                if (text != null)
                {
                    double y = Height - text.Y;
                    int colorKey = ColorKey(text.Color);
                    if (colorKey != fillKey)
                    {
                        fillKey = colorKey;
                        AppendColor(builder, "rg", text.Color, scratch);
                    }
                    builder.Append("BT\n");
                    if (!ReferenceEquals(lastFont, text.Font) || lastFontSize != text.FontSize)
                    {
                        lastFont = text.Font;
                        lastFontSize = text.FontSize;
                        builder.Append("/").Append(text.Font.ResourceName).Append(" ");
                        AppendNumber(builder, text.FontSize, scratch);
                        builder.Append(" Tf\n");
                    }
                    builder.Append("1 0 0 1 ");
                    AppendNumber(builder, text.X, scratch);
                    builder.Append(' ');
                    AppendNumber(builder, y, scratch);
                    builder.Append(" Tm\n");
                    builder.Append("<").Append(text.Font.EncodeText(text.Text)).Append("> Tj\n");
                    builder.Append("ET\n");
                    continue;
                }

                PdfLineCommand line = _commands[i] as PdfLineCommand;
                if (line != null)
                {
                    if (line.Width <= 0) line.Width = 0.5;
                    int colorKey = ColorKey(line.Color);
                    if (colorKey != strokeKey)
                    {
                        strokeKey = colorKey;
                        AppendColor(builder, "RG", line.Color, scratch);
                    }
                    if (line.Width != lastWidth)
                    {
                        lastWidth = line.Width;
                        AppendNumber(builder, line.Width, scratch);
                        builder.Append(" w\n");
                    }
                    int dash = line.Style == BorderStyle.Dashed ? 1 : (line.Style == BorderStyle.Dotted ? 2 : 0);
                    if (dash != lastDash)
                    {
                        lastDash = dash;
                        if (dash == 1) builder.Append("[3 2] 0 d\n");
                        else if (dash == 2) builder.Append("[1 2] 0 d\n");
                        else builder.Append("[] 0 d\n");
                    }
                    AppendNumber(builder, line.X1, scratch);
                    builder.Append(' ');
                    AppendNumber(builder, Height - line.Y1, scratch);
                    builder.Append(" m ");
                    AppendNumber(builder, line.X2, scratch);
                    builder.Append(' ');
                    AppendNumber(builder, Height - line.Y2, scratch);
                    builder.Append(" l S\n");
                    continue;
                }

                PdfFillRectCommand rect = _commands[i] as PdfFillRectCommand;
                if (rect != null)
                {
                    int colorKey = ColorKey(rect.Color);
                    if (colorKey != fillKey)
                    {
                        fillKey = colorKey;
                        AppendColor(builder, "rg", rect.Color, scratch);
                    }
                    AppendRectPath(builder, rect.X, rect.Y, rect.Width, rect.Height, scratch);
                    builder.Append(" f\n");
                    continue;
                }

                PdfShapeCommand shape = _commands[i] as PdfShapeCommand;
                if (shape != null)
                {
                    bool fill = shape.FillColor.HasValue;
                    bool stroke = shape.StrokeWidth > 0;
                    if (fill)
                    {
                        OfficeColor fillColor = shape.FillColor.Value;
                        int colorKey = ColorKey(fillColor);
                        if (colorKey != fillKey)
                        {
                            fillKey = colorKey;
                            AppendColor(builder, "rg", fillColor, scratch);
                        }
                    }
                    if (stroke)
                    {
                        int colorKey = ColorKey(shape.StrokeColor);
                        if (colorKey != strokeKey)
                        {
                            strokeKey = colorKey;
                            AppendColor(builder, "RG", shape.StrokeColor, scratch);
                        }
                        if (shape.StrokeWidth != lastWidth)
                        {
                            lastWidth = shape.StrokeWidth;
                            AppendNumber(builder, shape.StrokeWidth, scratch);
                            builder.Append(" w\n");
                        }
                        int dash = shape.StrokeStyle == BorderStyle.Dashed ? 1 : (shape.StrokeStyle == BorderStyle.Dotted ? 2 : 0);
                        if (dash != lastDash)
                        {
                            lastDash = dash;
                            if (dash == 1) builder.Append("[3 2] 0 d\n");
                            else if (dash == 2) builder.Append("[1 2] 0 d\n");
                            else builder.Append("[] 0 d\n");
                        }
                    }
                    AppendShapePath(builder, shape, scratch);
                    builder.Append(fill && stroke ? " B\n" : (fill ? " f\n" : " S\n"));
                    continue;
                }

                PdfImageCommand image = _commands[i] as PdfImageCommand;
                if (image != null)
                {
                    builder.Append("q\n");
                    AppendNumber(builder, image.Width, scratch);
                    builder.Append(" 0 0 ");
                    AppendNumber(builder, image.Height, scratch);
                    builder.Append(' ');
                    AppendNumber(builder, image.X, scratch);
                    builder.Append(' ');
                    AppendNumber(builder, Height - image.Y - image.Height, scratch);
                    builder.Append(" cm\n");
                    builder.Append("/").Append(image.Image.ResourceName).Append(" Do\n");
                    builder.Append("Q\n");
                    lastFont = null;
                    lastFontSize = double.NaN;
                    fillKey = -1;
                    strokeKey = -1;
                    lastWidth = double.NaN;
                    lastDash = -1;
                }
            }
            return Encoding.ASCII.GetBytes(builder.ToString());
        }

        /// <summary>
        /// 写出形状路径（不含绘图算子）：矩形用 re，圆角矩形与椭圆用贝塞尔曲线，多边形用折线。
        /// Writes the shape path without the painting operator: rectangles use re, rounded rectangles and
        /// ellipses use Bézier curves, polygons use a polyline.
        /// </summary>
        /// <param name="builder">内容流构建器。The content stream builder.</param>
        /// <param name="shape">形状指令。The shape command.</param>
        /// <param name="scratch">数字格式化的临时缓冲。The scratch buffer for number formatting.</param>
        private void AppendShapePath(StringBuilder builder, PdfShapeCommand shape, char[] scratch)
        {
            double w = shape.Width;
            double h = shape.Height;

            if (shape.Kind == PdfShapeKind.Polygon)
            {
                double[] points = shape.Points;
                if (points == null || points.Length < 6) return;
                AppendPoint(builder, points[0], points[1], scratch);
                builder.Append(" m");
                for (int i = 2; i + 1 < points.Length; i += 2)
                {
                    builder.Append(' ');
                    AppendPoint(builder, points[i], points[i + 1], scratch);
                    builder.Append(" l");
                }
                builder.Append(" h");
                return;
            }

            if (w <= 0 || h <= 0 || shape.Kind == PdfShapeKind.Rectangle)
            {
                AppendRectPath(builder, shape.X, shape.Y, w, h, scratch);
                return;
            }

            if (shape.Kind == PdfShapeKind.Ellipse)
            {
                const double Kappa = 0.5522847498307936;
                double cx = shape.X + w / 2;
                double cy = shape.Y + h / 2;
                double rx = w / 2;
                double ry = h / 2;
                double kx = rx * Kappa;
                double ky = ry * Kappa;
                AppendPoint(builder, cx + rx, cy, scratch);
                builder.Append(" m");
                AppendCurve(builder, cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry, scratch);
                AppendCurve(builder, cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy, scratch);
                AppendCurve(builder, cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry, scratch);
                AppendCurve(builder, cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy, scratch);
                builder.Append(" h");
                return;
            }

            double radius = shape.CornerRadius;
            double cap = Math.Min(w, h) / 2;
            if (radius > cap) radius = cap;
            if (radius <= 0)
            {
                // 圆角半径为 0 时退化为直角矩形。A zero radius degrades to square corners.
                AppendRectPath(builder, shape.X, shape.Y, w, h, scratch);
                return;
            }
            double control = radius * 0.5522847498307936;
            double x1 = shape.X;
            double y1 = shape.Y;
            double x2 = x1 + w;
            double y2 = y1 + h;
            AppendPoint(builder, x1 + radius, y1, scratch);
            builder.Append(" m");
            builder.Append(' ');
            AppendPoint(builder, x2 - radius, y1, scratch);
            builder.Append(" l");
            AppendCurve(builder, x2 - radius + control, y1, x2, y1 + radius - control, x2, y1 + radius, scratch);
            builder.Append(' ');
            AppendPoint(builder, x2, y2 - radius, scratch);
            builder.Append(" l");
            AppendCurve(builder, x2, y2 - radius + control, x2 - radius + control, y2, x2 - radius, y2, scratch);
            builder.Append(' ');
            AppendPoint(builder, x1 + radius, y2, scratch);
            builder.Append(" l");
            AppendCurve(builder, x1 + radius - control, y2, x1, y2 - radius + control, x1, y2 - radius, scratch);
            builder.Append(' ');
            AppendPoint(builder, x1, y1 + radius, scratch);
            builder.Append(" l");
            AppendCurve(builder, x1, y1 + radius - control, x1 + radius - control, y1, x1 + radius, y1, scratch);
            builder.Append(" h");
        }

        /// <summary>追加矩形路径（re 算子）。Appends a rectangle path (the re operator).</summary>
        /// <param name="builder">内容流构建器。The content stream builder.</param>
        /// <param name="x">左上角横坐标。Top-left X.</param>
        /// <param name="y">左上角纵坐标。Top-left Y.</param>
        /// <param name="width">宽度。The width.</param>
        /// <param name="height">高度。The height.</param>
        /// <param name="scratch">数字格式化的临时缓冲。The scratch buffer for number formatting.</param>
        private void AppendRectPath(StringBuilder builder, double x, double y, double width, double height, char[] scratch)
        {
            AppendNumber(builder, x, scratch);
            builder.Append(' ');
            AppendNumber(builder, Height - y - height, scratch);
            builder.Append(' ');
            AppendNumber(builder, width, scratch);
            builder.Append(' ');
            AppendNumber(builder, height, scratch);
            builder.Append(" re");
        }

        /// <summary>追加一个顶点，Y 在此翻转为 PDF 左下原点。Appends one vertex, flipping Y to the PDF bottom-left origin.</summary>
        /// <param name="builder">内容流构建器。The content stream builder.</param>
        /// <param name="x">横坐标（左上原点）。X in the top-left system.</param>
        /// <param name="y">纵坐标（左上原点）。Y in the top-left system.</param>
        /// <param name="scratch">数字格式化的临时缓冲。The scratch buffer for number formatting.</param>
        private void AppendPoint(StringBuilder builder, double x, double y, char[] scratch)
        {
            AppendNumber(builder, x, scratch);
            builder.Append(' ');
            AppendNumber(builder, Height - y, scratch);
        }

        /// <summary>追加一段三次贝塞尔曲线（c 算子，起点取当前点）。Appends one cubic Bézier segment (the c operator, starting from the current point).</summary>
        /// <param name="builder">内容流构建器。The content stream builder.</param>
        /// <param name="x1">第一控制点横坐标。First control point X.</param>
        /// <param name="y1">第一控制点纵坐标。First control point Y.</param>
        /// <param name="x2">第二控制点横坐标。Second control point X.</param>
        /// <param name="y2">第二控制点纵坐标。Second control point Y.</param>
        /// <param name="x3">终点横坐标。End X.</param>
        /// <param name="y3">终点纵坐标。End Y.</param>
        /// <param name="scratch">数字格式化的临时缓冲。The scratch buffer for number formatting.</param>
        private void AppendCurve(StringBuilder builder, double x1, double y1,
            double x2, double y2, double x3, double y3, char[] scratch)
        {
            builder.Append(' ');
            AppendPoint(builder, x1, y1, scratch);
            builder.Append(' ');
            AppendPoint(builder, x2, y2, scratch);
            builder.Append(' ');
            AppendPoint(builder, x3, y3, scratch);
            builder.Append(" c");
        }

        /// <summary>把颜色折叠成可比较的整数键（内容流只按 RGB 写出，不含通道值以外的差异）。
        /// Collapses a color into a comparable integer key (the content stream writes RGB only).</summary>
        /// <param name="color">颜色。The color.</param>
        /// <returns>24 位 RGB 键。The 24-bit RGB key.</returns>
        private static int ColorKey(OfficeColor color)
        {
            return (color.R << 16) | (color.G << 8) | color.B;
        }

        /// <summary>
        /// 写出 RGB 颜色算子（rg 或 RG）。
        /// Emits an RGB color operator (rg or RG).
        /// </summary>
        /// <param name="builder">内容流构建器。The content stream builder.</param>
        /// <param name="operator">算子名，"rg" 为填充色、"RG" 为描边色。The operator name: "rg" fills, "RG" strokes.</param>
        /// <param name="color">颜色，各通道按 0-255 归一化。The color, with each 0-255 channel normalized.</param>
        /// <param name="scratch">数字格式化的临时缓冲。The scratch buffer for number formatting.</param>
        private static void AppendColor(StringBuilder builder, string @operator, OfficeColor color, char[] scratch)
        {
            AppendNumber(builder, color.R / 255.0, scratch);
            builder.Append(' ');
            AppendNumber(builder, color.G / 255.0, scratch);
            builder.Append(' ');
            AppendNumber(builder, color.B / 255.0, scratch);
            builder.Append(' ').Append(@operator).Append('\n');
        }

        /// <summary>
        /// 把数值直接追加为 PDF 操作数，最多 4 位小数，不产生中间字符串；NaN/Infinity 写作 0。
        /// Appends a value straight into the builder as a PDF operand with at most 4 decimals and no
        /// intermediate string; NaN and Infinity become 0.
        /// </summary>
        /// <param name="builder">内容流构建器。The content stream builder.</param>
        /// <param name="value">待格式化的数值。The value to format.</param>
        /// <param name="scratch">整数位使用的临时缓冲。The scratch buffer for the integer digits.</param>
        private static void AppendNumber(StringBuilder builder, double value, char[] scratch)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                builder.Append('0');
                return;
            }

            bool negative = value < 0;
            if (negative) value = -value;
            if (value > 4.0e14)
            {
                builder.Append(Format(value));
                return;
            }

            long scaled = (long)(value * 10000.0 + 0.5);
            if (scaled == 0)
            {
                builder.Append('0');
                return;
            }

            long whole = scaled / 10000;
            int fraction = (int)(scaled - whole * 10000);
            if (negative) builder.Append('-');

            if (whole > 0)
            {
                int position = scratch.Length;
                while (whole > 0)
                {
                    position--;
                    scratch[position] = (char)('0' + (int)(whole % 10));
                    whole /= 10;
                }
                builder.Append(scratch, position, scratch.Length - position);
            }
            else
            {
                builder.Append('0');
            }

            if (fraction > 0)
            {
                int digits = 4;
                if (fraction % 1000 == 0) digits = 1;
                else if (fraction % 100 == 0) digits = 2;
                else if (fraction % 10 == 0) digits = 3;
                builder.Append('.');
                int divisor = 1000;
                for (int i = 0; i < digits; i++)
                {
                    builder.Append((char)('0' + fraction / divisor));
                    fraction %= divisor;
                    divisor /= 10;
                }
            }
        }

        /// <summary>
        /// 把数值格式化为 PDF 操作数，最多 4 位小数；NaN/Infinity 一律写成 0。
        /// Formats a number as a PDF operand with at most 4 decimals; NaN and Infinity are written as 0.
        /// </summary>
        /// <param name="value">待格式化的数值。The value to format.</param>
        /// <returns>不变文化（invariant-culture）的数字字符串。The culture-invariant numeric string.</returns>
        private static string Format(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return "0";
            return value.ToString("0.####", CultureInfo.InvariantCulture);
        }
    }
}
