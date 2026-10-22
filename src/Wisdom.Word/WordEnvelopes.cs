using System;
using System.Collections.Generic;
using System.Globalization;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 信封版式：纸张尺寸、页边距以及寄件人/收件人地址距页面边缘的位置，长度单位为磅。
    /// Envelope layout: page size, margins, and the positions of the sender and recipient addresses measured
    /// from the page edges, with every length in points.
    /// </summary>
    /// <remarks>
    /// 内置 DL、C5 与中式 5 号信封；可用 <see cref="FromMillimeters"/> 定义自定义尺寸。生成器用段前间距与
    /// 左缩进定位地址，因此不需要文本框或图形对象。
    /// DL, C5 and the Chinese No. 5 envelope are built in; use FromMillimeters for custom sizes. The generator
    /// positions addresses with paragraph spacing and left indentation, so no text boxes or drawing objects are needed.
    /// </remarks>
    public sealed class WordEnvelopeLayout
    {
        /// <summary>版式名称。The layout name.</summary>
        public string Name { get; set; }
        /// <summary>信封宽度，单位为磅。The envelope width, in points.</summary>
        public double PageWidth { get; set; }
        /// <summary>信封高度，单位为磅。The envelope height, in points.</summary>
        public double PageHeight { get; set; }
        /// <summary>左边距，单位为磅。The left margin, in points.</summary>
        public double MarginLeft { get; set; }
        /// <summary>上边距，单位为磅。The top margin, in points.</summary>
        public double MarginTop { get; set; }
        /// <summary>右边距，单位为磅。The right margin, in points.</summary>
        public double MarginRight { get; set; }
        /// <summary>下边距，单位为磅。The bottom margin, in points.</summary>
        public double MarginBottom { get; set; }
        /// <summary>寄件人地址距页面左侧的距离，单位为磅。The sender address's distance from the page left edge, in points.</summary>
        public double SenderLeft { get; set; }
        /// <summary>寄件人地址距页面顶部的距离，单位为磅。The sender address's distance from the page top, in points.</summary>
        public double SenderTop { get; set; }
        /// <summary>收件人地址距页面左侧的距离，单位为磅。The recipient address's distance from the page left edge, in points.</summary>
        public double RecipientLeft { get; set; }
        /// <summary>收件人地址距页面顶部的距离，单位为磅。The recipient address's distance from the page top, in points.</summary>
        public double RecipientTop { get; set; }

        /// <summary>默认构造：DL 信封（220 × 110 毫米）。Default constructor: a DL envelope (220 by 110 mm).</summary>
        public WordEnvelopeLayout()
        {
            Name = "DL";
            PageWidth = MillimetersToPoints(220);
            PageHeight = MillimetersToPoints(110);
            MarginLeft = MillimetersToPoints(10);
            MarginTop = MillimetersToPoints(10);
            MarginRight = MillimetersToPoints(10);
            MarginBottom = MillimetersToPoints(10);
            SenderLeft = MillimetersToPoints(15);
            SenderTop = MillimetersToPoints(12);
            RecipientLeft = MillimetersToPoints(105);
            RecipientTop = MillimetersToPoints(45);
        }

        /// <summary>DL 信封：220 × 110 毫米，收件人地址位于右侧窗口中。DL envelope: 220 by 110 mm with the recipient address in the right-hand window area.</summary>
        /// <returns>新的信封版式。The new envelope layout.</returns>
        public static WordEnvelopeLayout Dl()
        {
            return FromMillimeters("DL", 220, 110, 10, 15, 12, 105, 45);
        }

        /// <summary>C5 信封：229 × 162 毫米。C5 envelope: 229 by 162 mm.</summary>
        /// <returns>新的信封版式。The new envelope layout.</returns>
        public static WordEnvelopeLayout C5()
        {
            return FromMillimeters("C5", 229, 162, 10, 15, 12, 110, 60);
        }

        /// <summary>中式 5 号信封：220 × 110 毫米。Chinese No. 5 envelope: 220 by 110 mm.</summary>
        /// <returns>新的信封版式。The new envelope layout.</returns>
        public static WordEnvelopeLayout Chinese5()
        {
            return FromMillimeters("Chinese No. 5", 220, 110, 10, 15, 12, 100, 45);
        }

        /// <summary>
        /// 以毫米为单位创建自定义信封版式。
        /// Creates a custom envelope layout expressed in millimetres.
        /// </summary>
        /// <param name="name">版式名称。The layout name.</param>
        /// <param name="pageWidthMm">信封宽度（毫米）。The envelope width in millimetres.</param>
        /// <param name="pageHeightMm">信封高度（毫米）。The envelope height in millimetres.</param>
        /// <param name="marginMm">四边页边距（毫米）。The page margin on all four sides, in millimetres.</param>
        /// <param name="senderLeftMm">寄件人左侧位置（毫米）。The sender's left position in millimetres.</param>
        /// <param name="senderTopMm">寄件人顶部位置（毫米）。The sender's top position in millimetres.</param>
        /// <param name="recipientLeftMm">收件人左侧位置（毫米）。The recipient's left position in millimetres.</param>
        /// <param name="recipientTopMm">收件人顶部位置（毫米）。The recipient's top position in millimetres.</param>
        /// <returns>新的信封版式。The new envelope layout.</returns>
        public static WordEnvelopeLayout FromMillimeters(
            string name,
            double pageWidthMm,
            double pageHeightMm,
            double marginMm,
            double senderLeftMm,
            double senderTopMm,
            double recipientLeftMm,
            double recipientTopMm)
        {
            double margin = MillimetersToPoints(marginMm);
            return new WordEnvelopeLayout
            {
                Name = name ?? string.Empty,
                PageWidth = MillimetersToPoints(pageWidthMm),
                PageHeight = MillimetersToPoints(pageHeightMm),
                MarginLeft = margin,
                MarginTop = margin,
                MarginRight = margin,
                MarginBottom = margin,
                SenderLeft = MillimetersToPoints(senderLeftMm),
                SenderTop = MillimetersToPoints(senderTopMm),
                RecipientLeft = MillimetersToPoints(recipientLeftMm),
                RecipientTop = MillimetersToPoints(recipientTopMm)
            };
        }

        /// <summary>
        /// 生成信封的页面设置：纸张取版式尺寸，四边距取版式边距，页眉页脚距离置零。
        /// Builds the envelope page setup: page size and margins come from the layout and header/footer
        /// distances are zero.
        /// </summary>
        /// <returns>新的页面设置。The new page setup.</returns>
        public PageSetup CreatePageSetup()
        {
            Validate();
            PageSetup setup = new PageSetup();
            setup.Width = PageWidth;
            setup.Height = PageHeight;
            setup.MarginLeft = MarginLeft;
            setup.MarginTop = MarginTop;
            setup.MarginRight = MarginRight;
            setup.MarginBottom = MarginBottom;
            setup.HeaderDistance = 0;
            setup.FooterDistance = 0;
            return setup;
        }

        /// <summary>
        /// 校验版式：尺寸为正、地址位置落在信封范围内。
        /// Validates the layout: positive sizes and address positions inside the envelope.
        /// </summary>
        /// <exception cref="ArgumentException">版式数值非法。Thrown when a layout value is invalid.</exception>
        public void Validate()
        {
            if (PageWidth <= 0 || PageHeight <= 0)
            {
                throw new ArgumentException("Envelope page width and height must be positive.");
            }
            if (MarginLeft < 0 || MarginTop < 0 || MarginRight < 0 || MarginBottom < 0)
            {
                throw new ArgumentException("Envelope margins must not be negative.");
            }
            if (MarginLeft + MarginRight >= PageWidth || MarginTop + MarginBottom >= PageHeight)
            {
                throw new ArgumentException("Envelope margins leave no printable area.");
            }
            if (SenderLeft < 0 || SenderTop < 0 || RecipientLeft < 0 || RecipientTop < 0
                || SenderLeft >= PageWidth || RecipientLeft >= PageWidth
                || SenderTop >= PageHeight || RecipientTop >= PageHeight)
            {
                throw new ArgumentException("Envelope address positions must fall inside the page.");
            }
        }

        /// <summary>把毫米换算为磅（1 毫米 = 72 / 25.4 磅）。Converts millimetres to points (1 mm = 72 / 25.4 points).</summary>
        /// <param name="millimeters">毫米值。The value in millimetres.</param>
        /// <returns>磅值。The value in points.</returns>
        private static double MillimetersToPoints(double millimeters)
        {
            return millimeters * 72.0 / 25.4;
        }
    }

    /// <summary>
    /// 信封内容选项：版式、寄件人/收件人地址文本与字体。
    /// Envelope content options: layout, sender/recipient address text and fonts.
    /// </summary>
    public sealed class WordEnvelopeOptions
    {
        /// <summary>信封版式，默认 DL。The envelope layout; DL by default.</summary>
        public WordEnvelopeLayout Layout { get; set; }
        /// <summary>寄件人地址，可含换行；空字符串表示不打印寄件人。The sender address, which may contain line breaks; empty means no sender block.</summary>
        public string Sender { get; set; }
        /// <summary>收件人地址，可含换行。The recipient address, which may contain line breaks.</summary>
        public string Recipient { get; set; }
        /// <summary>字体名。The font family.</summary>
        public string FontFamily { get; set; }
        /// <summary>中文字体名，为空时回退到 <see cref="FontFamily"/>。The East-Asian font family; falls back to FontFamily when empty.</summary>
        public string EastAsiaFontFamily { get; set; }
        /// <summary>收件人字号，单位为磅。The recipient font size, in points.</summary>
        public double FontSize { get; set; }
        /// <summary>寄件人字号，单位为磅。The sender font size, in points.</summary>
        public double SenderFontSize { get; set; }
        /// <summary>行距倍数。The line-spacing multiple.</summary>
        public double LineSpacing { get; set; }
        /// <summary>是否加粗收件人地址。Whether the recipient address is bold.</summary>
        public bool Bold { get; set; }

        /// <summary>默认构造：DL 版式、无地址、宋体 12 磅（寄件人 9 磅）、单倍行距。Default constructor: DL layout, no addresses, SimSun at 12 points (sender 9 points) and single spacing.</summary>
        public WordEnvelopeOptions()
        {
            Layout = WordEnvelopeLayout.Dl();
            Sender = string.Empty;
            Recipient = string.Empty;
            FontFamily = "SimSun";
            EastAsiaFontFamily = "SimSun";
            FontSize = 12;
            SenderFontSize = 9;
            LineSpacing = 1.0;
        }
    }

    /// <summary>
    /// 信封生成器：按版式放置寄件人与收件人地址，支持单个信封与按记录批量生成。
    /// Envelope generator: places sender and recipient addresses according to a layout, for a single
    /// envelope or a batch built from data records.
    /// </summary>
    /// <remarks>
    /// 地址用段前间距和左缩进定位，输出为普通段落，Word/WPS 与 PDF 导出都能直接使用。版式位置小于页边距时
    /// 会贴齐页边距。
    /// Addresses are positioned with paragraph spacing and left indentation as ordinary paragraphs, so Word,
    /// WPS and PDF export all render them. Positions closer to the edge than the margin are clamped to the margin.
    /// </remarks>
    public static class WordEnvelopes
    {
        /// <summary>
        /// 生成单个信封。
        /// Creates a single envelope.
        /// </summary>
        /// <param name="options">信封内容与版式，不能为 null。The envelope content and layout; must not be null.</param>
        /// <returns>信封文档模型。The envelope document model.</returns>
        public static OfficeDocument Create(WordEnvelopeOptions options)
        {
            if (options == null) throw new ArgumentNullException("options");
            WordEnvelopeLayout layout = options.Layout ?? WordEnvelopeLayout.Dl();
            layout.Validate();
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection(layout.CreatePageSetup());
            AppendEnvelope(section, layout, options, options.Sender, options.Recipient);
            return document;
        }

        /// <summary>
        /// 用数据记录批量生成信封：每条记录一页，收件人模板中的 {{字段}} 按记录解析。
        /// Creates a batch of envelopes from data records: one envelope per page, with {{field}} tokens in
        /// the recipient template resolved against each record.
        /// </summary>
        /// <typeparam name="T">记录类型，可为字典、DataRow、DataRowView、TemplateContext 或普通对象。The record type: a dictionary, DataRow, DataRowView, TemplateContext or plain object.</typeparam>
        /// <param name="records">数据记录集合。The data records.</param>
        /// <param name="recipientTemplate">收件人模板，可含换行。The recipient template, which may contain line breaks.</param>
        /// <param name="options">字体、版式与固定寄件人；null 时使用默认值。Fonts, layout and fixed sender; defaults are used when null.</param>
        /// <param name="senderTemplate">可选的寄件人模板；为 null 时使用 <see cref="WordEnvelopeOptions.Sender"/>。An optional sender template; null uses WordEnvelopeOptions.Sender.</param>
        /// <returns>信封文档模型。The envelope document model.</returns>
        public static OfficeDocument CreateMany<T>(
            IEnumerable<T> records,
            string recipientTemplate,
            WordEnvelopeOptions options = null,
            string senderTemplate = null)
        {
            if (records == null) throw new ArgumentNullException("records");
            if (recipientTemplate == null) throw new ArgumentNullException("recipientTemplate");
            options = options ?? new WordEnvelopeOptions();
            WordEnvelopeLayout layout = options.Layout ?? WordEnvelopeLayout.Dl();
            layout.Validate();

            TemplateOptions templateOptions = new TemplateOptions
            {
                LeaveUnresolvedTokens = false,
                Culture = CultureInfo.InvariantCulture
            };
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection(layout.CreatePageSetup());
            bool any = false;
            foreach (T record in records)
            {
                if (any) section.Blocks.Add(new PageBreak());
                any = true;
                TemplateContext context = WordTemplateText.CreateContext(record);
                string recipient = TemplateEngine.Replace(recipientTemplate, context, templateOptions);
                string sender = senderTemplate == null
                    ? options.Sender
                    : TemplateEngine.Replace(senderTemplate, context, templateOptions);
                AppendEnvelope(section, layout, options, sender, recipient);
            }
            if (!any)
            {
                AppendEnvelope(section, layout, options, options.Sender, options.Recipient);
            }
            return document;
        }

        /// <summary>
        /// 把寄件人与收件人地址块追加到节中：用段前间距控制纵向位置，用左缩进控制横向位置。
        /// Appends the sender and recipient blocks to the section, using space-before for vertical
        /// placement and left indentation for horizontal placement.
        /// </summary>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="layout">信封版式。The envelope layout.</param>
        /// <param name="options">字体与行距选项。The font and spacing options.</param>
        /// <param name="sender">寄件人文本。The sender text.</param>
        /// <param name="recipient">收件人文本。The recipient text.</param>
        private static void AppendEnvelope(
            Section section,
            WordEnvelopeLayout layout,
            WordEnvelopeOptions options,
            string sender,
            string recipient)
        {
            bool hasSender = !string.IsNullOrEmpty(sender);
            bool hasRecipient = !string.IsNullOrEmpty(recipient);
            TextFormat recipientFormat = CreateFormat(options, options.FontSize, options.Bold);
            TextFormat senderFormat = CreateFormat(options, options.SenderFontSize, false);

            double senderStart = Math.Max(layout.MarginTop, layout.SenderTop);
            double senderLeft = Math.Max(0, layout.SenderLeft - layout.MarginLeft);
            double recipientLeft = Math.Max(0, layout.RecipientLeft - layout.MarginLeft);
            double senderBlockHeight = 0;

            if (hasSender)
            {
                string[] senderLines = WordTemplateText.SplitLines(sender);
                senderBlockHeight = senderLines.Length * options.SenderFontSize * 1.2 * options.LineSpacing;
                for (int i = 0; i < senderLines.Length; i++)
                {
                    Paragraph paragraph = CreateParagraph(
                        senderLines[i],
                        senderFormat,
                        options.LineSpacing,
                        senderLeft,
                        i == 0 ? Math.Max(0, layout.SenderTop - layout.MarginTop) : 0);
                    section.Blocks.Add(paragraph);
                }
            }

            if (hasRecipient)
            {
                string[] recipientLines = WordTemplateText.SplitLines(recipient);
                double recipientSpaceBefore = hasSender
                    ? Math.Max(0, layout.RecipientTop - (senderStart + senderBlockHeight))
                    : Math.Max(0, layout.RecipientTop - layout.MarginTop);
                for (int i = 0; i < recipientLines.Length; i++)
                {
                    Paragraph paragraph = CreateParagraph(
                        recipientLines[i],
                        recipientFormat,
                        options.LineSpacing,
                        recipientLeft,
                        i == 0 ? recipientSpaceBefore : 0);
                    section.Blocks.Add(paragraph);
                }
            }

            if (!hasSender && !hasRecipient)
            {
                section.Blocks.Add(new Paragraph());
            }
        }

        /// <summary>创建信封文本格式。Creates the text format used by envelope addresses.</summary>
        /// <param name="options">信封选项。The envelope options.</param>
        /// <param name="fontSize">字号。The font size.</param>
        /// <param name="bold">是否加粗。Whether to use bold text.</param>
        /// <returns>文本格式。The text format.</returns>
        private static TextFormat CreateFormat(WordEnvelopeOptions options, double fontSize, bool bold)
        {
            TextFormat format = new TextFormat(options.FontFamily, fontSize);
            format.EastAsiaFontFamily = options.EastAsiaFontFamily;
            format.Bold = bold;
            return format;
        }

        /// <summary>创建带定位属性的地址段落。Creates an address paragraph with positioning properties.</summary>
        /// <param name="text">段落文本。The paragraph text.</param>
        /// <param name="format">字符格式。The character format.</param>
        /// <param name="lineSpacing">行距倍数。The line-spacing multiple.</param>
        /// <param name="leftIndent">左缩进。The left indent.</param>
        /// <param name="spaceBefore">段前间距。The space before the paragraph.</param>
        /// <returns>地址段落。The address paragraph.</returns>
        private static Paragraph CreateParagraph(
            string text,
            TextFormat format,
            double lineSpacing,
            double leftIndent,
            double spaceBefore)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.LeftIndent = leftIndent;
            paragraph.SpaceBefore = spaceBefore;
            paragraph.SpaceAfter = 0;
            paragraph.LineSpacing = lineSpacing;
            paragraph.AddText(text, format);
            return paragraph;
        }
    }
}
