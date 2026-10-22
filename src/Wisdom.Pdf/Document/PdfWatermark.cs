using System;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 水印所处的图层。The layer a watermark is drawn on.
    /// </summary>
    public enum WatermarkLayer
    {
        /// <summary>背景层：水印内容流排在页面原内容之前，不遮挡正文， readability 最佳。
        /// Background: the watermark stream is placed before the page content, so it never covers the text.</summary>
        Background = 0,
        /// <summary>前景层：水印内容流排在页面原内容之后，覆盖正文，防伪提示更醒目。
        /// Foreground: the watermark stream is placed after the page content, covering it for a stronger notice.</summary>
        Foreground = 1
    }

    /// <summary>
    /// 水印在页面上的摆放方式。The placement of a watermark on the page.
    /// </summary>
    public enum WatermarkPlacement
    {
        /// <summary>斜跨页面中央，默认 45°，按 <see cref="WatermarkOptions.AutoFit"/> 自动缩放。
        /// Across the page centre at 45 degrees by default, scaled to fit when <see cref="WatermarkOptions.AutoFit"/> is on.</summary>
        Diagonal = 0,
        /// <summary>页面正中，默认不旋转。Centre of the page, unrotated by default.</summary>
        Center = 1,
        /// <summary>平铺整页，按 <see cref="WatermarkOptions.TileSpacingX"/> 与
        /// <see cref="WatermarkOptions.TileSpacingY"/> 重复。Tiled across the whole page.</summary>
        Tile = 2,
        /// <summary>左上角，向内缩进 <see cref="WatermarkOptions.Margin"/>。Top-left corner, inset by the margin.</summary>
        TopLeft = 3,
        /// <summary>右上角。Top-right corner.</summary>
        TopRight = 4,
        /// <summary>左下角。Bottom-left corner.</summary>
        BottomLeft = 5,
        /// <summary>右下角。Bottom-right corner.</summary>
        BottomRight = 6,
        /// <summary>自由定位，以 <see cref="WatermarkOptions.Left"/> 与 <see cref="WatermarkOptions.Top"/>
        /// 从页面左上角量起。Free positioning measured from the top-left corner.</summary>
        Free = 7
    }

    /// <summary>
    /// 水印参数：文字或图像内容、字体字号、颜色透明度、旋转角度与摆放方式。
    /// Watermark settings: text or image content, font and size, colour and opacity, rotation and placement.
    /// </summary>
    /// <remarks>
    /// 参照 Aspose.PDF 的 WatermarkArtifact/PageStamp、iText 7 的页面事件画布与 QuestPDF 的页面背景/前景层，
    /// 统一为一个可复用的选项对象：同一次调用里字体与图像只嵌入一份，逐页只生成一个 Form XObject。
    /// Modelled on Aspose.PDF's WatermarkArtifact/PageStamp, iText 7's page-event canvas and QuestPDF's page
    /// background/foreground layers, but folded into one reusable options object: one call embeds each font and
    /// image once and adds only a single Form XObject per page.
    /// </remarks>
    public sealed class WatermarkOptions
    {
        /// <summary>水印文字，可为 null；与 <see cref="ImageData"/> 同时给出时先画图后写字。
        /// The watermark text; may be null. With <see cref="ImageData"/> both are drawn, image first.</summary>
        public string Text { get; set; }
        /// <summary>水印图像字节（PNG 或 JPEG），可为 null。The watermark image bytes (PNG or JPEG); may be null.</summary>
        public byte[] ImageData { get; set; }
        /// <summary>水印图像文件路径，设置后读取为 <see cref="ImageData"/>。The image file path, loaded into <see cref="ImageData"/>.</summary>
        public string ImagePath
        {
            get { return _imagePath; }
            set
            {
                _imagePath = value;
                ImageData = string.IsNullOrEmpty(value) ? null : System.IO.File.ReadAllBytes(value);
            }
        }
        /// <summary>图像宽度（点），0 表示按像素密度自动取值。Image width in points; 0 means automatic.</summary>
        public double ImageWidth { get; set; }
        /// <summary>图像高度（点），0 表示按宽高比自动取值。Image height in points; 0 keeps the aspect ratio.</summary>
        public double ImageHeight { get; set; }

        /// <summary>字体家族名，默认 Arial，中文会自动回退到中文字体。The font family, Arial by default; Chinese falls back automatically.</summary>
        public string FontFamily { get; set; }
        /// <summary>是否粗体。Whether the text is bold.</summary>
        public bool Bold { get; set; }
        /// <summary>是否斜体。Whether the text is italic.</summary>
        public bool Italic { get; set; }
        /// <summary>字号（点），默认 60。The font size in points, 60 by default.</summary>
        public double FontSize { get; set; }
        /// <summary>水印颜色，默认中性灰。The watermark colour, neutral gray by default.</summary>
        public OfficeColor Color { get; set; }
        /// <summary>不透明度 0~1，默认 0.30；小于 1 时使用 ExtGState alpha。Opacity from 0 to 1, 0.30 by default.</summary>
        public double Opacity { get; set; }
        /// <summary>旋转角度（度，逆时针），null 时由摆放方式决定默认值。Rotation in degrees, counter-clockwise; null uses the placement default.</summary>
        public double? Angle { get; set; }
        /// <summary>所在图层，默认背景层。The layer, background by default.</summary>
        public WatermarkLayer Layer { get; set; }
        /// <summary>摆放方式，默认斜跨中央。The placement, diagonal by default.</summary>
        public WatermarkPlacement Placement { get; set; }
        /// <summary>页边距（点）：角部摆放的内缩量，也是自动缩放时预留的空白。Page margin in points, used for corner insets and fit padding.</summary>
        public double Margin { get; set; }
        /// <summary>自由摆放时距页面左边的距离（点）。Distance from the left edge for free placement.</summary>
        public double Left { get; set; }
        /// <summary>自由摆放时距页面上边的距离（点）。Distance from the top edge for free placement.</summary>
        public double Top { get; set; }
        /// <summary>平铺时沿文字基线方向的间距（点）。Tile spacing along the baseline direction.</summary>
        public double TileSpacingX { get; set; }
        /// <summary>平铺时垂直于基线方向的间距（点）。Tile spacing across the baseline direction.</summary>
        public double TileSpacingY { get; set; }
        /// <summary>是否自动缩小以放进页面，默认 true。Whether the watermark shrinks to fit the page, true by default.</summary>
        public bool AutoFit { get; set; }
        /// <summary>字体解析器，null 时使用进程级默认解析器。The font resolver; null uses the process-wide default.</summary>
        public FontResolver Resolver { get; set; }

        /// <summary>最近一次设置的文件路径原文。The last image path that was set.</summary>
        private string _imagePath;

        /// <summary>
        /// 创建带默认值的水印选项。
        /// Creates watermark options with the default values.
        /// </summary>
        public WatermarkOptions()
        {
            FontFamily = "Arial";
            FontSize = 60;
            Color = new OfficeColor(189, 189, 189);
            Opacity = 0.3;
            Layer = WatermarkLayer.Background;
            Placement = WatermarkPlacement.Diagonal;
            Margin = 24;
            TileSpacingX = 60;
            TileSpacingY = 40;
            AutoFit = true;
        }

        /// <summary>
        /// 创建文字水印选项，除文本外全部使用默认值。
        /// Creates text watermark options with defaults apart from the text.
        /// </summary>
        /// <param name="text">水印文字。The watermark text.</param>
        /// <returns>可直接传给 <see cref="Wisdom.Pdf.PdfFile.AddWatermark(WatermarkOptions)"/> 的选项。Options ready for <see cref="Wisdom.Pdf.PdfFile.AddWatermark(WatermarkOptions)"/>.</returns>
        public static WatermarkOptions ForText(string text)
        {
            WatermarkOptions options = new WatermarkOptions();
            options.Text = text;
            return options;
        }

        /// <summary>
        /// 创建图像水印选项。
        /// Creates image watermark options from image bytes.
        /// </summary>
        /// <param name="imageData">PNG 或 JPEG 字节。The PNG or JPEG bytes.</param>
        /// <returns>图像水印选项。The image watermark options.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="imageData"/> 为 null 时抛出。Thrown when <paramref name="imageData"/> is null.</exception>
        public static WatermarkOptions ForImage(byte[] imageData)
        {
            if (imageData == null) throw new ArgumentNullException("imageData");
            WatermarkOptions options = new WatermarkOptions();
            options.ImageData = imageData;
            return options;
        }

        /// <summary>校验选项是否可写出。Validates that the options can be written.</summary>
        internal void Validate()
        {
            if (string.IsNullOrEmpty(Text) && (ImageData == null || ImageData.Length == 0))
            {
                throw new InvalidOperationException("A watermark needs text or an image.");
            }
            if (FontSize <= 0) throw new ArgumentOutOfRangeException("FontSize");
            if (Opacity < 0 || Opacity > 1) throw new ArgumentOutOfRangeException("Opacity");
            if (Margin < 0) throw new ArgumentOutOfRangeException("Margin");
        }

        /// <summary>
        /// 取得生效的旋转角度：显式设置优先，否则斜跨与平铺用 45°，其余摆放用 0°。
        /// Gets the effective angle: the explicit value wins, otherwise diagonal and tile use 45 degrees and the
        /// other placements use 0.
        /// </summary>
        /// <returns>角度（度，逆时针）。The angle in degrees, counter-clockwise.</returns>
        internal double ResolveAngle()
        {
            if (Angle.HasValue) return Angle.Value;
            if (Placement == WatermarkPlacement.Diagonal) return 45;
            if (Placement == WatermarkPlacement.Tile) return 45;
            return 0;
        }
    }
}
