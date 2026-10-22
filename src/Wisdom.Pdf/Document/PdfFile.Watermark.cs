using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 页面水印：把文字或图像作为 Form XObject 加到一个或全部页面的背景/前景层。
    /// Page watermarks: text or images added as a Form XObject to the background or foreground layer of one or
    /// all pages.
    /// </summary>

    public sealed partial class PdfFile
    {
        /// <summary>平铺水印单页最多绘制的次数，超出后停止以避免内容流过大。Maximum tile draws per page, beyond which tiling stops to keep the stream small.</summary>
        private const int MaxTileDraws = 600;

        /// <summary>
        /// 给全部页面加文字水印，使用默认样式（斜跨、灰色、30% 不透明度、背景层）。
        /// Adds a text watermark to every page with the default style (diagonal, gray, 30% opacity, background).
        /// </summary>
        /// <param name="text">水印文字，支持中文（自动回退到中文字体）。The watermark text; Chinese works through font fallback.</param>
        public void AddWatermark(string text)
        {
            AddWatermark(text, WatermarkLayer.Background);
        }

        /// <summary>
        /// 给全部页面加文字水印并指定图层。
        /// Adds a text watermark to every page on the chosen layer.
        /// </summary>
        /// <param name="text">水印文字。The watermark text.</param>
        /// <param name="layer">背景层或前景层。The background or foreground layer.</param>
        /// <remarks>第二个参数是图层而非页码：整数字面量会隐式转成 <see cref="WatermarkLayer"/>，
        /// 只给指定页加水印请改用 <see cref="AddWatermark(WatermarkOptions, int[])"/>。
        /// The second argument is the layer, not a page index: an integer literal converts implicitly to
        /// <see cref="WatermarkLayer"/>. To target specific pages use
        /// <see cref="AddWatermark(WatermarkOptions, int[])"/> instead.</remarks>
        public void AddWatermark(string text, WatermarkLayer layer)
        {
            WatermarkOptions options = WatermarkOptions.ForText(text);
            options.Layer = layer;
            AddWatermark(options);
        }

        /// <summary>
        /// 给全部页面加水印。
        /// Adds a watermark to every page.
        /// </summary>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <exception cref="ArgumentNullException"><paramref name="options"/> 为 null 时抛出。Thrown when <paramref name="options"/> is null.</exception>
        /// <exception cref="InvalidOperationException">选项既无文字又无图像时抛出。Thrown when the options carry neither text nor image.</exception>
        public void AddWatermark(WatermarkOptions options)
        {
            ApplyWatermark(options, null);
        }

        /// <summary>
        /// 给指定页加水印；下标从 0 开始，越界下标会被忽略。
        /// Adds a watermark to the given pages (zero-based); out-of-range indexes are ignored.
        /// </summary>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <param name="pageIndexes">目标页面下标，为空时处理全部页面。The target page indexes; all pages when empty.</param>
        public void AddWatermark(WatermarkOptions options, params int[] pageIndexes)
        {
            ApplyWatermark(options, pageIndexes);
        }

        /// <summary>
        /// 从图片文件给全部页面加图像水印。
        /// Adds an image watermark to every page from an image file.
        /// </summary>
        /// <param name="imagePath">PNG 或 JPEG 图片路径。The PNG or JPEG image path.</param>
        /// <param name="layer">背景层或前景层。The background or foreground layer.</param>
        public void AddImageWatermark(string imagePath, WatermarkLayer layer = WatermarkLayer.Background)
        {
            WatermarkOptions options = new WatermarkOptions();
            options.ImagePath = imagePath;
            options.Layer = layer;
            options.Placement = WatermarkPlacement.Diagonal;
            AddWatermark(options);
        }

        /// <summary>
        /// 执行水印写出：准备共享资源，再逐页生成 Form XObject 与内容流。
        /// Performs the watermark write: the shared resources are prepared once, then each page gets its own
        /// Form XObject and content stream.
        /// </summary>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <param name="pageIndexes">目标页面下标，null 或空表示全部页面。Target page indexes; null or empty means every page.</param>
        private void ApplyWatermark(WatermarkOptions options, IList<int> pageIndexes)
        {
            if (options == null) throw new ArgumentNullException("options");
            options.Validate();

            List<int> targets = ResolveWatermarkPages(pageIndexes);
            if (targets.Count == 0) return;

            if (options.Opacity < 1) EnsureTransparencyVersion();

            WatermarkContent content = BuildWatermarkContent(options);

            double angle = options.ResolveAngle() * Math.PI / 180.0;
            for (int i = 0; i < targets.Count; i++)
            {
                DrawWatermarkPage(content, options, targets[i], angle, i + 1);
            }
        }

        /// <summary>
        /// 规整目标页面下标：去重、忽略越界，空集合表示全部页面。
        /// Normalizes the target page indexes: de-duplicates, drops out-of-range values, and expands an empty
        /// set to every page.
        /// </summary>
        /// <param name="pageIndexes">请求的下标。The requested indexes.</param>
        /// <returns>有效页面下标。The usable page indexes.</returns>
        private List<int> ResolveWatermarkPages(IList<int> pageIndexes)
        {
            int count = PageCount;
            List<int> result = new List<int>();
            if (pageIndexes == null || pageIndexes.Count == 0)
            {
                for (int i = 0; i < count; i++) result.Add(i);
                return result;
            }

            HashSet<int> seen = new HashSet<int>();
            for (int i = 0; i < pageIndexes.Count; i++)
            {
                int index = pageIndexes[i];
                if (index < 0 || index >= count) continue;
                if (!seen.Add(index)) continue;
                result.Add(index);
            }
            result.Sort();
            return result;
        }

        /// <summary>
        /// 准备水印用到的字体、图像与透明度资源，并量出文字块尺寸。
        /// Prepares the fonts, images and transparency resources used by the watermark and measures the text block.
        /// </summary>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <returns>可复用的水印内容。The reusable watermark content.</returns>
        private WatermarkContent BuildWatermarkContent(WatermarkOptions options)
        {
            PdfFileObjectWriter writer = new PdfFileObjectWriter(this);

            PdfDictionary resources = new PdfDictionary();
            resources.Set("ProcSet", new PdfArray()
                .AddName("PDF").AddName("Text").AddName("ImageB").AddName("ImageC").AddName("ImageI"));

            WatermarkContent content = new WatermarkContent();
            content.Resources = resources;

            if (options.Opacity < 1)
            {
                PdfDictionary extGState = new PdfDictionary();
                extGState.SetName("Type", "ExtGState");
                extGState.SetNumber("CA", options.Opacity);
                extGState.SetNumber("ca", options.Opacity);
                PdfDictionary states = new PdfDictionary();
                states.Set("WmAlpha", writer.Add(extGState));
                resources.Set("ExtGState", states);
                content.AlphaName = "WmAlpha";
            }

            if (!string.IsNullOrEmpty(options.Text))
            {
                PdfFontSet fonts = new PdfFontSet(options.Resolver);
                TextFormat format = new TextFormat();
                format.FontFamily = string.IsNullOrEmpty(options.FontFamily) ? "Arial" : options.FontFamily;
                format.FontSize = options.FontSize;
                format.Bold = options.Bold;
                format.Italic = options.Italic;

                double ascent = 0;
                double descent = 0;
                PdfFont runFont = null;
                StringBuilder run = new StringBuilder();
                foreach (int codePoint in UnicodeUtil.EnumerateCodePoints(options.Text))
                {
                    PdfFont font = fonts.GetFontForCodePoint(format, codePoint);
                    if (runFont != null && font != runFont)
                    {
                        content.Runs.Add(new WatermarkTextRun(runFont, run.ToString()));
                        run.Length = 0;
                    }
                    runFont = font;
                    font.MarkUsed(codePoint);
                    run.Append(UnicodeUtil.ToUtf16String(codePoint));
                    content.TextWidth += font.GetAdvance(codePoint, options.FontSize);
                    double fontAscent = font.GetAscent(options.FontSize);
                    double fontDescent = font.GetDescent(options.FontSize);
                    if (fontAscent > ascent) ascent = fontAscent;
                    if (fontDescent > descent) descent = fontDescent;
                }
                if (run.Length > 0 && runFont != null) content.Runs.Add(new WatermarkTextRun(runFont, run.ToString()));

                PdfDictionary fontResources = new PdfDictionary();
                HashSet<PdfFont> prepared = new HashSet<PdfFont>();
                for (int i = 0; i < content.Runs.Count; i++)
                {
                    PdfFont font = content.Runs[i].Font;
                    if (prepared.Add(font)) font.Prepare(writer, fontResources);
                }
                resources.Set("Font", fontResources);

                content.Size = options.FontSize;
                content.Ascent = ascent;
                content.Descent = descent;
                if (content.TextWidth <= 0) content.TextWidth = options.FontSize * 0.5 * options.Text.Length;
            }

            if (options.ImageData != null && options.ImageData.Length > 0)
            {
                OfficeImage image = new OfficeImage(options.ImageData);
                PdfImage pdfImage = new PdfImage(image);
                pdfImage.ResourceName = "WmImage";
                PdfDictionary imageResources = new PdfDictionary();
                pdfImage.Prepare(writer, imageResources);
                resources.Set("XObject", imageResources);

                double width = options.ImageWidth;
                double height = options.ImageHeight;
                if (width <= 0) width = image.DefaultWidthPoints;
                if (height <= 0) height = width * image.PixelHeight / Math.Max(1.0, image.PixelWidth);
                content.HasImage = true;
                content.ImageWidth = width;
                content.ImageHeight = height;
            }

            content.Color = options.Color;
            return content;
        }

        /// <summary>
        /// 在一页上画出水印：算出该页的显示尺寸与旋转补偿，生成 Form XObject 并挂到页面内容上。
        /// Draws the watermark on one page: the display size and rotation compensation are computed, the Form
        /// XObject is written, and the page content is wired to it.
        /// </summary>
        /// <param name="content">共享的水印内容。The shared watermark content.</param>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="angle">旋转弧度（逆时针）。The rotation in radians, counter-clockwise.</param>
        /// <param name="serial">本页水印序号，用于生成不冲突的资源名。The watermark serial used to build a unique resource name.</param>
        private void DrawWatermarkPage(WatermarkContent content, WatermarkOptions options, int index, double angle, int serial)
        {
            PageNode node = GetPageNode(index);
            double left = 0, bottom = 0, right = 612, top = 792;
            PdfArray box = Resolve(node.MediaBox) as PdfArray;
            if (box != null && box.Count >= 4)
            {
                PdfNumber n1 = Resolve(box[0]) as PdfNumber;
                PdfNumber n2 = Resolve(box[1]) as PdfNumber;
                PdfNumber n3 = Resolve(box[2]) as PdfNumber;
                PdfNumber n4 = Resolve(box[3]) as PdfNumber;
                if (n1 != null && n2 != null && n3 != null && n4 != null)
                {
                    left = Math.Min(n1.Value, n3.Value);
                    bottom = Math.Min(n2.Value, n4.Value);
                    right = Math.Max(n1.Value, n3.Value);
                    top = Math.Max(n2.Value, n4.Value);
                }
            }

            double width = right - left;
            double height = top - bottom;
            int rotation = node.Rotation;
            bool swapped = rotation == 90 || rotation == 270;
            double displayWidth = swapped ? height : width;
            double displayHeight = swapped ? width : height;
            if (displayWidth <= 0 || displayHeight <= 0) return;

            double scale = options.AutoFit ? FitScale(content, options, displayWidth, displayHeight, angle) : 1;

            StringBuilder body = new StringBuilder();
            if (content.AlphaName != null) body.Append("/" + content.AlphaName + " gs\n");
            body.Append(ColorOperator(content.Color));

            if (options.Placement == WatermarkPlacement.Tile)
            {
                DrawTiled(body, content, options, displayWidth, displayHeight, angle, scale);
            }
            else
            {
                double centerX, centerY;
                AnchorOf(options, content, displayWidth, displayHeight, angle, scale, out centerX, out centerY);
                DrawItems(body, content, scale, centerX, centerY, angle);
            }

            if (body.Length == 0) return;

            PdfDictionary formDictionary = new PdfDictionary();
            formDictionary.SetName("Type", "XObject");
            formDictionary.SetName("Subtype", "Form");
            formDictionary.Set("BBox", new PdfArray().AddNumber(left).AddNumber(bottom).AddNumber(right).AddNumber(top));
            formDictionary.Set("Resources", content.Resources);
            if (content.AlphaName != null)
            {
                PdfDictionary group = new PdfDictionary();
                group.SetName("Type", "Group");
                group.SetName("S", "Transparency");
                group.SetName("CS", "DeviceRGB");
                formDictionary.Set("Group", group);
            }

            byte[] formBytes = Encoding.ASCII.GetBytes(body.ToString());
            PdfReference formReference = AddObjectReference(new PdfStream(formDictionary, formBytes, true));

            string resourceName = WatermarkResourceName(index, serial);
            SetPageResource(index, "XObject", resourceName, formReference);

            StringBuilder pageStream = new StringBuilder();
            PdfDictionary page = GetPageDictionary(index);
            bool tagged = page.ContainsKey("StructParents");
            if (tagged)
            {
                pageStream.Append("/Artifact << /Type /Pagination /Subtype /Watermark /Attached /Top >> BDC\n");
            }
            pageStream.Append(RotationCompensation(rotation, left, bottom, right, top));
            pageStream.Append("/" + resourceName + " Do\nQ\n");
            if (tagged) pageStream.Append("EMC\n");

            byte[] pageBytes = Encoding.ASCII.GetBytes(pageStream.ToString());
            if (options.Layer == WatermarkLayer.Background) PrependPageContent(index, pageBytes, true);
            else AppendPageContent(index, pageBytes, true);
        }

        /// <summary>
        /// 生成页面旋转补偿指令：先 q，再按 /Rotate 把「阅读方向坐标」映射回未旋转的页面坐标。
        /// Emits the rotation compensation: an opening q plus the matrix that maps reading-direction
        /// coordinates back into the unrotated page space for the given /Rotate.
        /// </summary>
        /// <param name="rotation">页面旋转角度。The page rotation.</param>
        /// <param name="left">MediaBox 左边界。The MediaBox left edge.</param>
        /// <param name="bottom">MediaBox 下边界。The MediaBox bottom edge.</param>
        /// <param name="right">MediaBox 右边界。The MediaBox right edge.</param>
        /// <param name="top">MediaBox 上边界。The MediaBox top edge.</param>
        /// <returns>q 加 cm 两行。The q plus cm lines.</returns>
        private static string RotationCompensation(int rotation, double left, double bottom, double right, double top)
        {
            string matrix;
            switch (rotation)
            {
                case 90: matrix = "0 1 -1 0 " + Num(right) + " " + Num(bottom); break;
                case 180: matrix = "-1 0 0 -1 " + Num(right) + " " + Num(top); break;
                case 270: matrix = "0 -1 1 0 " + Num(left) + " " + Num(top); break;
                default: matrix = "1 0 0 1 " + Num(left) + " " + Num(bottom); break;
            }
            return "q\n" + matrix + " cm\n";
        }

        /// <summary>
        /// 计算把水印缩到页面内的比例（只缩小不放大）。Computes the shrink-only factor that fits the watermark onto the page.
        /// </summary>
        /// <param name="content">水印内容。The watermark content.</param>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <param name="displayWidth">阅读方向宽度。The reading-direction width.</param>
        /// <param name="displayHeight">阅读方向高度。The reading-direction height.</param>
        /// <param name="angle">旋转弧度。The rotation in radians.</param>
        /// <returns>0~1 之间的缩放比例。A factor between 0 and 1.</returns>
        private static double FitScale(WatermarkContent content, WatermarkOptions options,
            double displayWidth, double displayHeight, double angle)
        {
            double blockWidth, blockHeight;
            RotatedExtents(content, options, 1, angle, out blockWidth, out blockHeight);
            if (blockWidth <= 0 || blockHeight <= 0) return 1;
            double availableWidth = displayWidth - 2 * options.Margin;
            double availableHeight = displayHeight - 2 * options.Margin;
            if (availableWidth <= 0 || availableHeight <= 0) return 1;
            double scale = Math.Min(availableWidth / blockWidth, availableHeight / blockHeight);
            if (scale >= 1) return 1;
            return Math.Max(0.05, scale);
        }

        /// <summary>
        /// 求出水印块旋转后的外接矩形尺寸。Measures the rotated bounding box of the watermark block.
        /// </summary>
        /// <param name="content">水印内容。The watermark content.</param>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <param name="scale">缩放比例。The scale factor.</param>
        /// <param name="angle">旋转弧度。The rotation in radians.</param>
        /// <param name="width">旋转后宽度。The rotated width.</param>
        /// <param name="height">旋转后高度。The rotated height.</param>
        private static void RotatedExtents(WatermarkContent content, WatermarkOptions options, double scale,
            double angle, out double width, out double height)
        {
            double cos = Math.Abs(Math.Cos(angle));
            double sin = Math.Abs(Math.Sin(angle));
            double blockWidth = content.TextWidth * scale;
            double blockHeight = (content.Ascent + content.Descent) * scale;
            if (content.HasImage)
            {
                if (content.ImageWidth * scale > blockWidth) blockWidth = content.ImageWidth * scale;
                if (content.ImageHeight * scale > blockHeight) blockHeight = content.ImageHeight * scale;
            }
            width = blockWidth * cos + blockHeight * sin;
            height = blockWidth * sin + blockHeight * cos;
        }

        /// <summary>
        /// 按摆放方式算出水印中心点（阅读方向坐标，原点左下）。Computes the watermark centre in reading-direction coordinates, origin bottom-left.
        /// </summary>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <param name="content">水印内容。The watermark content.</param>
        /// <param name="displayWidth">阅读方向宽度。The reading-direction width.</param>
        /// <param name="displayHeight">阅读方向高度。The reading-direction height.</param>
        /// <param name="angle">旋转弧度。The rotation in radians.</param>
        /// <param name="scale">缩放比例。The scale factor.</param>
        /// <param name="centerX">水印中心 X。The centre X.</param>
        /// <param name="centerY">水印中心 Y。The centre Y.</param>
        private static void AnchorOf(WatermarkOptions options, WatermarkContent content, double displayWidth,
            double displayHeight, double angle, double scale, out double centerX, out double centerY)
        {
            double rotatedWidth, rotatedHeight;
            RotatedExtents(content, options, scale, angle, out rotatedWidth, out rotatedHeight);
            double margin = options.Margin;
            switch (options.Placement)
            {
                case WatermarkPlacement.TopLeft:
                    centerX = margin + rotatedWidth / 2;
                    centerY = displayHeight - margin - rotatedHeight / 2;
                    return;
                case WatermarkPlacement.TopRight:
                    centerX = displayWidth - margin - rotatedWidth / 2;
                    centerY = displayHeight - margin - rotatedHeight / 2;
                    return;
                case WatermarkPlacement.BottomLeft:
                    centerX = margin + rotatedWidth / 2;
                    centerY = margin + rotatedHeight / 2;
                    return;
                case WatermarkPlacement.BottomRight:
                    centerX = displayWidth - margin - rotatedWidth / 2;
                    centerY = margin + rotatedHeight / 2;
                    return;
                case WatermarkPlacement.Free:
                    centerX = options.Left + rotatedWidth / 2;
                    centerY = displayHeight - options.Top - rotatedHeight / 2;
                    return;
                default:
                    centerX = displayWidth / 2;
                    centerY = displayHeight / 2;
                    return;
            }
        }

        /// <summary>
        /// 平铺水印：把页面四角投影到文字方向坐标系，按步长覆盖整个页面。Tiles the watermark by projecting the page corners onto the text axes and stepping across them.
        /// </summary>
        /// <param name="body">水印内容流累加器。The watermark stream accumulator.</param>
        /// <param name="content">水印内容。The watermark content.</param>
        /// <param name="options">水印选项。The watermark options.</param>
        /// <param name="displayWidth">阅读方向宽度。The reading-direction width.</param>
        /// <param name="displayHeight">阅读方向高度。The reading-direction height.</param>
        /// <param name="angle">旋转弧度。The rotation in radians.</param>
        /// <param name="scale">缩放比例。The scale factor.</param>
        private static void DrawTiled(StringBuilder body, WatermarkContent content, WatermarkOptions options,
            double displayWidth, double displayHeight, double angle, double scale)
        {
            double blockWidth = content.TextWidth * scale;
            double blockHeight = (content.Ascent + content.Descent) * scale;
            if (content.HasImage)
            {
                if (content.ImageWidth * scale > blockWidth) blockWidth = content.ImageWidth * scale;
                if (content.ImageHeight * scale > blockHeight) blockHeight = content.ImageHeight * scale;
            }
            double stepU = blockWidth + Math.Max(0, options.TileSpacingX);
            double stepV = blockHeight + Math.Max(0, options.TileSpacingY);
            if (stepU <= 0 || stepV <= 0) return;

            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double[] cornerU = new double[4];
            double[] cornerV = new double[4];
            double[,] corners = new double[,] { { 0, 0 }, { displayWidth, 0 }, { displayWidth, displayHeight }, { 0, displayHeight } };
            double minU = double.MaxValue, maxU = double.MinValue;
            double minV = double.MaxValue, maxV = double.MinValue;
            for (int i = 0; i < 4; i++)
            {
                double x = corners[i, 0];
                double y = corners[i, 1];
                cornerU[i] = x * cos + y * sin;
                cornerV[i] = -x * sin + y * cos;
            }
            for (int i = 0; i < 4; i++)
            {
                if (cornerU[i] < minU) minU = cornerU[i];
                if (cornerU[i] > maxU) maxU = cornerU[i];
                if (cornerV[i] < minV) minV = cornerV[i];
                if (cornerV[i] > maxV) maxV = cornerV[i];
            }

            int iMin = (int)Math.Floor((minU - blockWidth) / stepU);
            int iMax = (int)Math.Ceiling((maxU + blockWidth) / stepU);
            int jMin = (int)Math.Floor((minV - blockHeight) / stepV);
            int jMax = (int)Math.Ceiling((maxV + blockHeight) / stepV);

            int drawn = 0;
            for (int j = jMin; j <= jMax && drawn < MaxTileDraws; j++)
            {
                for (int i = iMin; i <= iMax && drawn < MaxTileDraws; i++)
                {
                    double centerU = i * stepU;
                    double centerV = j * stepV;
                    double centerX = centerU * cos - centerV * sin;
                    double centerY = centerU * sin + centerV * cos;
                    DrawItems(body, content, scale, centerX, centerY, angle);
                    drawn++;
                }
            }
        }

        /// <summary>
        /// 在给定中心点画出图像与文字两项内容。Draws the image and text items centred at the given point.
        /// </summary>
        /// <param name="body">水印内容流累加器。The watermark stream accumulator.</param>
        /// <param name="content">水印内容。The watermark content.</param>
        /// <param name="scale">缩放比例。The scale factor.</param>
        /// <param name="centerX">中心 X（阅读方向坐标）。The centre X in reading-direction coordinates.</param>
        /// <param name="centerY">中心 Y。The centre Y.</param>
        /// <param name="angle">旋转弧度。The rotation in radians.</param>
        private static void DrawItems(StringBuilder body, WatermarkContent content, double scale,
            double centerX, double centerY, double angle)
        {
            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);

            if (content.HasImage)
            {
                double iw = content.ImageWidth * scale;
                double ih = content.ImageHeight * scale;
                double a = iw * cos;
                double b = iw * sin;
                double c = -ih * sin;
                double d = ih * cos;
                body.Append("q\n" + Num(a) + " " + Num(b) + " " + Num(c) + " " + Num(d) + " "
                    + Num(centerX - (a + c) / 2) + " " + Num(centerY - (b + d) / 2) + " cm\n/WmImage Do\nQ\n");
            }

            if (content.Runs.Count == 0) return;
            double size = content.Size * scale;
            double textWidth = content.TextWidth * scale;
            double baseline = (content.Ascent - content.Descent) * scale / 2;
            body.Append("q\n" + Num(cos) + " " + Num(sin) + " " + Num(-sin) + " " + Num(cos) + " "
                + Num(centerX) + " " + Num(centerY) + " cm\n");
            body.Append("BT\n1 0 0 1 " + Num(-textWidth / 2) + " " + Num(-baseline) + " Tm\n");
            for (int i = 0; i < content.Runs.Count; i++)
            {
                WatermarkTextRun run = content.Runs[i];
                body.Append("/" + run.Font.ResourceName + " " + Num(size) + " Tf\n");
                body.Append("<" + run.Font.EncodeText(run.Text) + "> Tj\n");
            }
            body.Append("ET\nQ\n");
        }

        /// <summary>
        /// 生成设置填充色的 rg 指令。Builds the rg operator that sets the fill colour.
        /// </summary>
        /// <param name="color">水印颜色。The watermark colour.</param>
        /// <returns>一行 rg 指令。One rg line.</returns>
        private static string ColorOperator(OfficeColor color)
        {
            return Num(color.R / 255.0) + " " + Num(color.G / 255.0) + " " + Num(color.B / 255.0) + " rg\n";
        }

        /// <summary>
        /// 选取一个与页面既有 XObject 不冲突的水印资源名。Picks a watermark resource name that does not clash with the page's XObjects.
        /// </summary>
        /// <param name="index">页面下标。The page index.</param>
        /// <param name="serial">本页水印序号。The watermark serial on this page.</param>
        /// <returns>资源名（不含斜杠）。The resource name without the slash.</returns>
        private string WatermarkResourceName(int index, int serial)
        {
            string baseName = "Wm" + serial.ToString(CultureInfo.InvariantCulture);
            PdfDictionary page = GetPageDictionary(index);
            PdfDictionary resources = Resolve(page.Get("Resources")) as PdfDictionary;
            PdfDictionary group = resources == null ? null : Resolve(resources.Get("XObject")) as PdfDictionary;
            if (group == null)
            {
                PdfDictionary inherited = GetPageResources(index);
                if (inherited != null) group = Resolve(inherited.Get("XObject")) as PdfDictionary;
            }
            string candidate = baseName;
            int suffix = 1;
            while (group != null && group.ContainsKey(candidate))
            {
                candidate = baseName + "_" + suffix.ToString(CultureInfo.InvariantCulture);
                suffix++;
            }
            return candidate;
        }

        /// <summary>
        /// 把数值格式化为 PDF 内容流可接受的 ASCII 文本（最多四位小数，避免出现 -0）。
        /// Formats a number as ASCII accepted by a content stream (up to four decimals, never "-0").
        /// </summary>
        /// <param name="value">数值。The value.</param>
        /// <returns>数值文本。The formatted text.</returns>
        private static string Num(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return "0";
            if (value == 0) value = 0;
            return value.ToString("0.####", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 透明水印要求文件头版本不低于 1.4，必要时提升版本。Transparency needs at least version 1.4, so the header version is raised when lower.
        /// </summary>
        private void EnsureTransparencyVersion()
        {
            string version = _version ?? "1.7";
            int dot = version.IndexOf('.');
            int major;
            int minor = 4;
            if (!int.TryParse(dot > 0 ? version.Substring(0, dot) : version, NumberStyles.Integer,
                CultureInfo.InvariantCulture, out major))
            {
                return;
            }
            if (dot > 0) int.TryParse(version.Substring(dot + 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out minor);
            if (major > 1 || minor >= 4) return;
            SetVersion("1.4");
        }

        /// <summary>
        /// 一次水印调用共享的资源与度量。The resources and metrics shared by one watermark call.
        /// </summary>
        private sealed class WatermarkContent
        {
            /// <summary>全部页面共用的 Form XObject 资源字典。The resource dictionary shared by every form XObject.</summary>
            public PdfDictionary Resources;
            /// <summary>透明度 ExtGState 的资源名，不透明时为 null。The alpha ExtGState resource name; null when fully opaque.</summary>
            public string AlphaName;
            /// <summary>按字体切分后的文字段。The text split into per-font runs.</summary>
            public readonly List<WatermarkTextRun> Runs = new List<WatermarkTextRun>();
            /// <summary>字号（点）。The font size in points.</summary>
            public double Size;
            /// <summary>文字总宽（点）。The total text width in points.</summary>
            public double TextWidth;
            /// <summary>最大 ascent（点）。The largest ascent in points.</summary>
            public double Ascent;
            /// <summary>最大 descent（点，正值）。The largest descent in points, positive.</summary>
            public double Descent;
            /// <summary>是否带图像水印。Whether an image watermark is present.</summary>
            public bool HasImage;
            /// <summary>图像宽度（点）。The image width in points.</summary>
            public double ImageWidth;
            /// <summary>图像高度（点）。The image height in points.</summary>
            public double ImageHeight;
            /// <summary>填充颜色。The fill colour.</summary>
            public OfficeColor Color;
        }

        /// <summary>
        /// 水印文字中使用同一字体的连续片段。A run of watermark text sharing one font.
        /// </summary>
        private sealed class WatermarkTextRun
        {
            /// <summary>本段使用的字体。The font used by the run.</summary>
            public readonly PdfFont Font;
            /// <summary>本段文本。The run text.</summary>
            public readonly string Text;

            /// <summary>创建一段水印文字。Creates a watermark text run.</summary>
            /// <param name="font">字体。The font.</param>
            /// <param name="text">文本。The text.</param>
            public WatermarkTextRun(PdfFont font, string text)
            {
                Font = font;
                Text = text;
            }
        }
    }
}
