using System;
using System.Collections.Generic;
using System.IO;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// PdfRenderer 的幻灯片固定版式模式：每张幻灯片渲染为一页，形状按绝对坐标与 z 序绘制，
    /// 保留位置、尺寸、填充、描边、圆角、背景色、内边距、垂直对齐与项目符号。
    /// The fixed-layout slide mode of PdfRenderer: every slide becomes one page and shapes are painted at
    /// their absolute coordinates in z order, preserving position, size, fill, outline, corner radius,
    /// background, insets, vertical anchoring and list markers.
    /// </summary>
    public sealed partial class PdfRenderer
    {
        /// <summary>无圆角时的兜底半径比例，与 PresentationML 预设几何的默认值一致。The fallback corner-radius ratio, matching the PresentationML preset default.</summary>
        private const double DefaultCornerRatio = 0.15;

        /// <summary>
        /// 以固定版式渲染演示文稿并写入文件。
        /// Renders a presentation in fixed layout and writes it to a file.
        /// </summary>
        /// <param name="presentation">待渲染的演示文稿。The presentation to render.</param>
        /// <param name="path">输出文件路径。The output file path.</param>
        /// <exception cref="ArgumentNullException"><paramref name="presentation"/> 或 <paramref name="path"/> 为 null 时抛出。Thrown when <paramref name="presentation"/> or <paramref name="path"/> is null.</exception>
        public void RenderPresentation(OfficePresentation presentation, string path)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            if (path == null) throw new ArgumentNullException("path");
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                RenderPresentation(presentation, stream);
            }
        }

        /// <summary>
        /// 以固定版式渲染演示文稿并返回 PDF 字节。
        /// Renders a presentation in fixed layout and returns the PDF bytes.
        /// </summary>
        /// <param name="presentation">待渲染的演示文稿。The presentation to render.</param>
        /// <returns>完整 PDF 文件的字节。The bytes of the complete PDF file.</returns>
        public byte[] RenderPresentationToBytes(OfficePresentation presentation)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                RenderPresentation(presentation, stream);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// 以固定版式渲染演示文稿并写入输出流：不走流式分页，一页幻对应一页，超出页面的内容按裁切处理。
        /// Renders a presentation in fixed layout into the output stream: no flow pagination, one page per
        /// slide, and content beyond the page edge is simply clipped.
        /// </summary>
        /// <param name="presentation">待渲染的演示文稿。The presentation to render.</param>
        /// <param name="output">接收 PDF 的流。The stream receiving the PDF.</param>
        /// <exception cref="ArgumentNullException"><paramref name="presentation"/> 或 <paramref name="output"/> 为 null 时抛出。Thrown when <paramref name="presentation"/> or <paramref name="output"/> is null.</exception>
        public void RenderPresentation(OfficePresentation presentation, Stream output)
        {
            if (presentation == null) throw new ArgumentNullException("presentation");
            if (output == null) throw new ArgumentNullException("output");

            _document = new OfficeDocument();
            CopyPresentationMetadata(presentation.Metadata, _document.Metadata);
            TextFormat bodyDefault = presentation.GetEffectiveTextFormat();
            _defaultFormat = bodyDefault != null ? bodyDefault.Clone() : new TextFormat("Calibri", 18.0);
            if (_defaultFormat.FontSize <= 0) _defaultFormat.FontSize = 18;
            _fonts = new PdfFontSet(_resolver);
            _images = new PdfImageCollection();
            _pages.Clear();

            double width = presentation.SlideWidth > 0 ? presentation.SlideWidth : 960.0;
            double height = presentation.SlideHeight > 0 ? presentation.SlideHeight : 540.0;
            PageSetup setup = new PageSetup();
            setup.Width = width;
            setup.Height = height;
            _section = new Section(setup);
            _section.AllowAutoPageBreak = false;

            List<Slide> slides = presentation.Slides ?? new List<Slide>();
            for (int i = 0; i < slides.Count; i++)
            {
                _page = new PdfPageContent(width, height, _section);
                _page.PageNumber = _pages.Count + 1;
                _pages.Add(_page);
                DrawSlide(slides[i] ?? new Slide(), presentation, width, height);
            }

            if (_pages.Count == 0)
            {
                _page = new PdfPageContent(width, height, _section);
                _page.PageNumber = 1;
                _pages.Add(_page);
            }

            WriteDocument(output);
        }

        /// <summary>绘制一张幻灯片：背景色优先，其后按 z 序绘制形状。Draws one slide: the background first, then shapes in z order.</summary>
        /// <param name="slide">目标幻灯片。The slide to draw.</param>
        /// <param name="presentation">所属演示文稿，提供默认字体。The owning presentation, supplying default fonts.</param>
        /// <param name="width">页面宽度（磅）。The page width in points.</param>
        /// <param name="height">页面高度（磅）。The page height in points.</param>
        private void DrawSlide(Slide slide, OfficePresentation presentation, double width, double height)
        {
            if (slide.BackgroundColor.HasValue)
            {
                _page.Add(new PdfFillRectCommand
                {
                    X = 0,
                    Y = 0,
                    Width = width,
                    Height = height,
                    Color = slide.BackgroundColor.Value
                });
            }

            List<SlideShape> shapes = slide.Shapes ?? new List<SlideShape>();
            for (int i = 0; i < shapes.Count; i++)
            {
                DrawSlideShape(shapes[i], presentation);
            }
        }

        /// <summary>按形状类型分派绘制。Dispatches shape drawing by type.</summary>
        /// <param name="shape">目标形状，null 时忽略。The shape to draw; null is ignored.</param>
        /// <param name="presentation">所属演示文稿。The owning presentation.</param>
        private void DrawSlideShape(SlideShape shape, OfficePresentation presentation)
        {
            if (shape == null) return;

            TableShape table = shape as TableShape;
            if (table != null)
            {
                DrawSlideTable(table);
                return;
            }

            PictureShape picture = shape as PictureShape;
            if (picture != null)
            {
                DrawShapeGeometry(picture);
                DrawSlidePicture(picture);
                return;
            }

            TextBoxShape box = shape as TextBoxShape;
            if (box != null)
            {
                DrawShapeGeometry(box);
                DrawSlideText(box, presentation);
            }
        }

        /// <summary>
        /// 绘制形状自身的填充与描边：矩形、圆角矩形、椭圆、多边形与直线分别落到路径指令上。
        /// Draws the shape's own fill and outline: rectangles, rounded rectangles, ellipses, polygons and
        /// straight lines each map onto a path command.
        /// </summary>
        /// <param name="shape">目标形状。The shape.</param>
        private void DrawShapeGeometry(SlideShape shape)
        {
            OfficeColor? fill = shape.FillColor;
            BorderLine outline = shape.Outline;
            bool stroke = outline != null && outline.Style != BorderStyle.None && outline.Width > 0;

            AutoShape auto = shape as AutoShape;
            AutoShapeKind kind = auto != null ? auto.Kind : AutoShapeKind.Rectangle;
            if (kind == AutoShapeKind.Line)
            {
                // 直线没有填充，只用描边；无描边时不产生任何图形。A line has no fill and needs a stroke to appear at all.
                if (stroke)
                {
                    AddBorderCommand(shape.Left, shape.Top, shape.Left + shape.Width, shape.Top + shape.Height,
                        outline.Width, outline.Color, outline.Style);
                }
                return;
            }

            if (!fill.HasValue && !stroke) return;

            PdfShapeCommand command = new PdfShapeCommand();
            command.X = shape.Left;
            command.Y = shape.Top;
            command.Width = shape.Width;
            command.Height = shape.Height;
            command.FillColor = fill;
            command.Kind = PdfShapeKind.Rectangle;
            if (stroke)
            {
                command.StrokeColor = outline.Color;
                command.StrokeWidth = outline.Width;
                command.StrokeStyle = outline.Style;
            }

            if (kind == AutoShapeKind.RoundedRectangle)
            {
                command.Kind = PdfShapeKind.RoundedRectangle;
                double ratio = auto != null && auto.CornerRadius > 0 ? auto.CornerRadius : DefaultCornerRatio;
                command.CornerRadius = ratio * Math.Min(shape.Width, shape.Height);
            }
            else if (kind == AutoShapeKind.Ellipse)
            {
                command.Kind = PdfShapeKind.Ellipse;
            }
            else
            {
                double[] points = AutoShapePoints(kind, shape.Left, shape.Top, shape.Width, shape.Height);
                if (points != null)
                {
                    command.Kind = PdfShapeKind.Polygon;
                    command.Points = points;
                }
            }

            _page.Add(command);
        }

        /// <summary>
        /// 计算非矩形预设几何的顶点；未覆盖的类型返回 null，由调用方退化为矩形。
        /// Computes the vertices of a non-rectangular preset geometry; uncovered kinds return null so the
        /// caller degrades to a rectangle.
        /// </summary>
        /// <param name="kind">图形类型。The shape kind.</param>
        /// <param name="x">外接矩形左边。The bounding box left.</param>
        /// <param name="y">外接矩形上边。The bounding box top.</param>
        /// <param name="width">外接矩形宽度。The bounding box width.</param>
        /// <param name="height">外接矩形高度。The bounding box height.</param>
        /// <returns>成对的顶点坐标（左上原点），或 null。The x,y vertex pairs in the top-left system, or null.</returns>
        private static double[] AutoShapePoints(AutoShapeKind kind, double x, double y, double width, double height)
        {
            switch (kind)
            {
                case AutoShapeKind.Triangle:
                    return new double[] { x + width / 2, y, x + width, y + height, x, y + height };
                case AutoShapeKind.Diamond:
                    return new double[] { x + width / 2, y, x + width, y + height / 2, x + width / 2, y + height, x, y + height / 2 };
                case AutoShapeKind.Pentagon:
                    return RegularPolygonPoints(5, x, y, width, height, -Math.PI / 2);
                case AutoShapeKind.Hexagon:
                    return RegularPolygonPoints(6, x, y, width, height, 0);
                case AutoShapeKind.RightArrow:
                    return new double[]
                    {
                        x, y + height * 0.25, x + width * 0.5, y + height * 0.25, x + width * 0.5, y,
                        x + width, y + height / 2, x + width * 0.5, y + height, x + width * 0.5, y + height * 0.75,
                        x, y + height * 0.75
                    };
                case AutoShapeKind.Chevron:
                    return new double[]
                    {
                        x, y, x + width * 0.5, y, x + width, y + height / 2,
                        x + width * 0.5, y + height, x, y + height, x + width * 0.5, y + height / 2
                    };
                default:
                    return null;
            }
        }

        /// <summary>生成内接于外接矩形的正多边形顶点。Builds the vertices of a regular polygon inscribed in the bounding box.</summary>
        /// <param name="sides">边数。The number of sides.</param>
        /// <param name="x">外接矩形左边。The bounding box left.</param>
        /// <param name="y">外接矩形上边。The bounding box top.</param>
        /// <param name="width">外接矩形宽度。The bounding box width.</param>
        /// <param name="height">外接矩形高度。The bounding box height.</param>
        /// <param name="startAngle">起始角（弧度）。The start angle in radians.</param>
        /// <returns>成对的顶点坐标。The x,y vertex pairs.</returns>
        private static double[] RegularPolygonPoints(int sides, double x, double y, double width, double height, double startAngle)
        {
            double cx = x + width / 2;
            double cy = y + height / 2;
            double rx = width / 2;
            double ry = height / 2;
            double[] points = new double[sides * 2];
            for (int i = 0; i < sides; i++)
            {
                double angle = startAngle + i * 2 * Math.PI / sides;
                points[i * 2] = cx + rx * Math.Cos(angle);
                points[i * 2 + 1] = cy + ry * Math.Sin(angle);
            }
            return points;
        }

        /// <summary>按形状位置与尺寸绘制图片。Draws the picture at the shape's position and size.</summary>
        /// <param name="picture">图片形状。The picture shape.</param>
        private void DrawSlidePicture(PictureShape picture)
        {
            OfficeImage image = picture.Image;
            if (image == null || image.Format == OfficeImageFormat.Unknown) return;
            double width = picture.Width;
            double height = picture.Height;
            if (width <= 0 && height <= 0)
            {
                width = image.DefaultWidthPoints;
                height = image.DefaultHeightPoints;
            }
            else if (width <= 0 && image.PixelHeight > 0)
            {
                width = height * image.PixelWidth / image.PixelHeight;
            }
            else if (height <= 0 && image.PixelWidth > 0)
            {
                height = width * image.PixelHeight / image.PixelWidth;
            }
            if (width <= 0 || height <= 0) return;

            PdfImage pdfImage = _images.Get(image);
            _page.Add(new PdfImageCommand
            {
                Image = pdfImage,
                X = picture.Left,
                Y = picture.Top,
                Width = width,
                Height = height
            });
            _page.AddImage(pdfImage);
        }

        /// <summary>
        /// 在形状位置绘制表格：先铺表头底色，再逐行调用与正文一致的表格绘制。
        /// Draws a table at the shape's position: header shading first, then rows through the same painter
        /// the body text uses.
        /// </summary>
        /// <param name="shape">表格形状。The table shape.</param>
        private void DrawSlideTable(TableShape shape)
        {
            Table table = shape.Table;
            if (table == null || table.Rows.Count == 0) return;

            double width = shape.Width > 0 ? shape.Width : _page.Width;
            TableLayout layout = BuildTableLayout(table, width);
            layout.OriginX = shape.Left;

            if (shape.HeaderColor.HasValue)
            {
                int headerRows = 0;
                while (headerRows < table.Rows.Count && table.Rows[headerRows].IsHeader) headerRows++;
                double headerHeight = 0;
                for (int r = 0; r < headerRows; r++) headerHeight += layout.Heights[r];
                if (headerHeight > 0)
                {
                    _page.Add(new PdfFillRectCommand
                    {
                        X = shape.Left,
                        Y = shape.Top,
                        Width = layout.Widths.Length > 0 ? TotalWidth(layout) : width,
                        Height = headerHeight,
                        Color = shape.HeaderColor.Value
                    });
                }
            }

            double y = shape.Top;
            for (int r = 0; r < table.Rows.Count; r++)
            {
                DrawTableRow(layout, r, shape.Left, y);
                y += layout.Heights[r];
            }
        }

        /// <summary>表格各列宽度之和。The summed column widths of a layout.</summary>
        /// <param name="layout">表格排版。The table layout.</param>
        /// <returns>总宽度（磅）。The total width in points.</returns>
        private static double TotalWidth(TableLayout layout)
        {
            double width = 0;
            for (int i = 0; i < layout.Widths.Length; i++) width += layout.Widths[i];
            return width;
        }

        /// <summary>
        /// 在文本框位置绘制文本：套用内边距、按垂直锚点起排，必要时整体缩放字号。
        /// Draws the text at the box's position: applies insets, starts from the vertical anchor and scales
        /// the font down as a whole when the box is set to shrink text to fit.
        /// </summary>
        /// <param name="box">文本框或自选图形。The text box or auto shape.</param>
        /// <param name="presentation">所属演示文稿，提供占位符默认字体。The owning presentation, supplying placeholder default fonts.</param>
        private void DrawSlideText(TextBoxShape box, OfficePresentation presentation)
        {
            if (box.Paragraphs.Count == 0) return;

            bool isTitle = box.Placeholder == PlaceholderKind.Title || box.Placeholder == PlaceholderKind.Subtitle;
            TextFormat fallback = isTitle ? presentation.GetEffectiveTitleFormat() : presentation.GetEffectiveTextFormat();
            if (fallback == null) fallback = _defaultFormat;

            double innerWidth = box.Width - box.MarginLeft - box.MarginRight;
            if (innerWidth < 10) innerWidth = Math.Max(10, box.Width * 0.8);
            double innerX = box.Left + box.MarginLeft;
            double innerTop = box.Top + box.MarginTop;
            double innerHeight = box.Height - box.MarginTop - box.MarginBottom;
            if (innerHeight < 0) innerHeight = 0;

            List<Paragraph> paragraphs = new List<Paragraph>();
            for (int i = 0; i < box.Paragraphs.Count; i++)
            {
                Paragraph source = box.Paragraphs[i];
                if (source == null) continue;
                if (source.Inlines.Count == 0 && source.GetText().Length == 0 && string.IsNullOrEmpty(source.ListMarker)) continue;
                Paragraph copy = source.Clone();
                ApplySlideDefaultFont(copy, fallback);
                paragraphs.Add(copy);
            }
            if (paragraphs.Count == 0) return;

            double total = 0;
            for (int i = 0; i < paragraphs.Count; i++) total += MeasureParagraph(paragraphs[i], innerWidth);

            if (box.ShrinkToFit && innerHeight > 0 && total > innerHeight)
            {
                double scale = innerHeight / total;
                if (scale < 0.5) scale = 0.5;
                if (scale < 1)
                {
                    for (int i = 0; i < paragraphs.Count; i++) ScaleParagraphFont(paragraphs[i], scale);
                    total = 0;
                    for (int i = 0; i < paragraphs.Count; i++) total += MeasureParagraph(paragraphs[i], innerWidth);
                }
            }

            double y = innerTop;
            if (box.VerticalAlignment == SlideTextAnchor.Middle) y += Math.Max(0, (innerHeight - total) / 2);
            else if (box.VerticalAlignment == SlideTextAnchor.Bottom) y += Math.Max(0, innerHeight - total);

            for (int i = 0; i < paragraphs.Count; i++)
            {
                y += DrawParagraphAt(paragraphs[i], innerX, innerWidth, y);
            }
        }

        /// <summary>
        /// 给未显式设置字体或字号的运行补上占位符默认格式，避免继承链缺失时字号忽大忽小。
        /// Fills the placeholder default font into runs without an explicit family or size so a missing
        /// inheritance chain does not produce erratic text.
        /// </summary>
        /// <param name="paragraph">目标段落（必须是副本）。The paragraph, which must be a copy.</param>
        /// <param name="fallback">默认文字格式。The fallback text format.</param>
        private void ApplySlideDefaultFont(Paragraph paragraph, TextFormat fallback)
        {
            if (paragraph == null || fallback == null) return;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run == null || run.Format == null) continue;
                if (string.IsNullOrEmpty(run.Format.FontFamily))
                {
                    run.Format.FontFamily = fallback.FontFamily;
                    run.Format.AsciiFontFamily = fallback.AsciiFontFamily;
                    run.Format.EastAsiaFontFamily = fallback.EastAsiaFontFamily;
                }
                if (run.Format.FontSize <= 0) run.Format.FontSize = fallback.FontSize;
            }
        }

        /// <summary>按系数缩放段落内所有运行的字号（含图片行内的高宽）。Scales every run's font size in the paragraph by a factor, images included.</summary>
        /// <param name="paragraph">目标段落（必须是副本）。The paragraph, which must be a copy.</param>
        /// <param name="scale">缩放系数。The scale factor.</param>
        private static void ScaleParagraphFont(Paragraph paragraph, double scale)
        {
            if (paragraph == null || scale <= 0 || scale >= 1) return;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                TextRun run = paragraph.Inlines[i] as TextRun;
                if (run != null && run.Format != null && run.Format.FontSize > 0)
                {
                    run.Format.FontSize *= scale;
                    continue;
                }
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field != null && field.Format != null && field.Format.FontSize > 0)
                {
                    field.Format.FontSize *= scale;
                    continue;
                }
                ImageRun image = paragraph.Inlines[i] as ImageRun;
                if (image != null)
                {
                    if (image.Width > 0) image.Width *= scale;
                    if (image.Height > 0) image.Height *= scale;
                }
            }
        }

        /// <summary>把演示文稿元数据复制进 PDF 信息字典的来源对象。Copies presentation metadata into the document that feeds the PDF Info dictionary.</summary>
        /// <param name="source">源元数据。Source metadata.</param>
        /// <param name="target">目标元数据。Target metadata.</param>
        private static void CopyPresentationMetadata(DocumentMetadata source, DocumentMetadata target)
        {
            if (source == null || target == null) return;
            target.Title = source.Title;
            target.Author = source.Author;
            target.Subject = source.Subject;
            target.Keywords = source.Keywords;
            target.Creator = source.Creator;
            target.Created = source.Created;
            target.Modified = source.Modified;
        }
    }
}
