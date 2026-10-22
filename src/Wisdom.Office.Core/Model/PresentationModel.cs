using System;
using System.Collections.Generic;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// 幻灯片版式类别，决定占位符的默认位置与用途。
    /// Slide layout category, determining the default placeholder positions and purpose.
    /// </summary>
    public enum SlideLayoutKind
    {
        /// <summary>仅标题页。Title-only slide.</summary>
        Title = 0,
        /// <summary>标题加正文。Title with body text.</summary>
        TitleAndContent = 1,
        /// <summary>章节过渡页。Section header slide.</summary>
        SectionHeader = 2,
        /// <summary>标题加左右两栏内容。Title with two content columns.</summary>
        TwoContent = 3,
        /// <summary>标题加图片。Title with a picture placeholder.</summary>
        TitleAndPicture = 4,
        /// <summary>标题加表格。Title with a table placeholder.</summary>
        TitleAndTable = 5,
        /// <summary>空白页。Blank slide.</summary>
        Blank = 6
    }

    /// <summary>
    /// 占位符类型，对应 PresentationML 的 <c>ph type</c> 属性。
    /// Placeholder type, mapping to the PresentationML <c>ph type</c> attribute.
    /// </summary>
    public enum PlaceholderKind
    {
        /// <summary>非占位符（普通文本框、图片等）。Not a placeholder (plain text box, picture, and so on).</summary>
        None = 0,
        /// <summary>标题占位符。Title placeholder.</summary>
        Title = 1,
        /// <summary>居中副标题占位符。Centered subtitle placeholder.</summary>
        Subtitle = 2,
        /// <summary>正文占位符。Body placeholder.</summary>
        Body = 3,
        /// <summary>内容占位符。Content placeholder.</summary>
        Content = 4,
        /// <summary>日期占位符。Date placeholder.</summary>
        Date = 5,
        /// <summary>页脚占位符。Footer placeholder.</summary>
        Footer = 6,
        /// <summary>幻灯片编号占位符。Slide number placeholder.</summary>
        SlideNumber = 7,
        /// <summary>图片占位符。Picture placeholder.</summary>
        Picture = 8,
        /// <summary>表格占位符。Table placeholder.</summary>
        Table = 9,
        /// <summary>图表占位符。Chart placeholder.</summary>
        Chart = 10
    }

    /// <summary>
    /// 自选图形类型。PPTX 只写出本库能映射到 <c>prstGeom</c> 的形状。
    /// Auto shape kind; PPTX emits only the shapes this library can map to a <c>prstGeom</c>.
    /// </summary>
    public enum AutoShapeKind
    {
        /// <summary>矩形。Rectangle.</summary>
        Rectangle = 0,
        /// <summary>圆角矩形。Rounded rectangle.</summary>
        RoundedRectangle = 1,
        /// <summary>椭圆。Ellipse.</summary>
        Ellipse = 2,
        /// <summary>三角形。Triangle.</summary>
        Triangle = 3,
        /// <summary>菱形。Diamond.</summary>
        Diamond = 4,
        /// <summary>五边形。Pentagon.</summary>
        Pentagon = 5,
        /// <summary>六边形。Hexagon.</summary>
        Hexagon = 6,
        /// <summary>右箭头。Right arrow.</summary>
        RightArrow = 7,
        /// <summary>燕尾形。Chevron.</summary>
        Chevron = 8,
        /// <summary>直线。Line.</summary>
        Line = 9
    }

    /// <summary>
    /// 文本在形状内的垂直对齐方式。Vertical alignment of text inside a shape.
    /// </summary>
    public enum SlideTextAnchor
    {
        /// <summary>顶部对齐。Aligned to the top.</summary>
        Top = 0,
        /// <summary>垂直居中。Vertically centered.</summary>
        Middle = 1,
        /// <summary>底部对齐。Aligned to the bottom.</summary>
        Bottom = 2
    }

    /// <summary>幻灯片切换效果。Slide transition effect.</summary>
    public enum SlideTransitionKind
    {
        /// <summary>无切换效果。No transition.</summary>
        None = 0,
        /// <summary>淡入淡出。Fade.</summary>
        Fade = 1,
        /// <summary>推入。Push.</summary>
        Push = 2,
        /// <summary>擦除。Wipe.</summary>
        Wipe = 3,
        /// <summary>分割。Split.</summary>
        Split = 4
    }

    /// <summary>形状入场动画类别。Shape entrance animation kind.</summary>
    public enum SlideAnimationKind
    {
        /// <summary>无动画：形状随幻灯片直接出现。No animation: the shape appears with the slide.</summary>
        None = 0,
        /// <summary>淡入。Fade in.</summary>
        Fade = 1,
        /// <summary>自底飞入。Fly in from the bottom.</summary>
        FlyIn = 2,
        /// <summary>自底擦除。Wipe from the bottom.</summary>
        Wipe = 3,
        /// <summary>缩放浮现。Grow and turn on.</summary>
        Zoom = 4,
        /// <summary>分割展开。Split open.</summary>
        Split = 5
    }

    /// <summary>
    /// 演示文稿主题深改面：名称、核心色板与主/次字体方案。所有字段默认 null 表示沿用内置默认
    /// （色板取 Office 默认，字体取 <see cref="OfficePresentation.DefaultTitleFormat"/> 与
    /// <see cref="OfficePresentation.DefaultTextFormat"/>）；强调色序列仍由
    /// <see cref="OfficePresentation.AccentColors"/> 提供，两者保持一致。
    /// Deep-edit surface of a presentation theme: name, core colour palette and major/minor font scheme.
    /// Every field defaults to null, meaning the built-in default (Office palette, or the fonts of
    /// <see cref="OfficePresentation.DefaultTitleFormat"/> and <see cref="OfficePresentation.DefaultTextFormat"/>);
    /// the accent sequence keeps coming from <see cref="OfficePresentation.AccentColors"/> so both stay in sync.
    /// </summary>
    public sealed class PresentationTheme
    {
        /// <summary>主题名称，null 使用默认 "Office Theme"。Theme name; null uses the default "Office Theme".</summary>
        public string Name { get; set; }
        /// <summary>主要文本色（clrScheme 的 dk1），null 使用默认。Primary text colour (dk1); null uses the default.</summary>
        public OfficeColor? Text { get; set; }
        /// <summary>主要背景色（clrScheme 的 lt1），null 使用默认。Primary background colour (lt1); null uses the default.</summary>
        public OfficeColor? Background { get; set; }
        /// <summary>次要文本色（dk2），null 使用默认。Secondary text colour (dk2); null uses the default.</summary>
        public OfficeColor? TextSecondary { get; set; }
        /// <summary>次要背景色（lt2），null 使用默认。Secondary background colour (lt2); null uses the default.</summary>
        public OfficeColor? BackgroundSecondary { get; set; }
        /// <summary>超链接色（hlink），null 使用默认。Hyperlink colour; null uses the default.</summary>
        public OfficeColor? Hyperlink { get; set; }
        /// <summary>已访问超链接色（folHlink），null 使用默认。Followed-hyperlink colour; null uses the default.</summary>
        public OfficeColor? FollowedHyperlink { get; set; }
        /// <summary>主字体（标题）族名，null 取标题默认格式。Major (title) font family; null falls back to the default title format.</summary>
        public string FontMajor { get; set; }
        /// <summary>次字体（正文）族名，null 取正文默认格式。Minor (body) font family; null falls back to the default body format.</summary>
        public string FontMinor { get; set; }
    }

    /// <summary>
    /// 幻灯片内所有图形对象的基类：几何位置、填充与描边。
    /// Base class for every shape on a slide: geometry, fill, and outline.
    /// </summary>
    public abstract class SlideShape
    {
        /// <summary>形状标识，写出时作为 <c>id</c> 属性；由模型自动分配。Shape identifier written as the <c>id</c> attribute; assigned by the model.</summary>
        public int Id { get; set; }
        /// <summary>形状名称，用于可访问性与编辑界面显示。Shape name, shown to editors and used for accessibility.</summary>
        public string Name { get; set; }
        /// <summary>左边距，单位为磅。Left offset in points.</summary>
        public double Left { get; set; }
        /// <summary>上边距，单位为磅。Top offset in points.</summary>
        public double Top { get; set; }
        /// <summary>宽度，单位为磅。Width in points.</summary>
        public double Width { get; set; }
        /// <summary>高度，单位为磅。Height in points.</summary>
        public double Height { get; set; }
        /// <summary>占位符类型；<see cref="PlaceholderKind.None"/> 表示自由摆放的形状。Placeholder type; <see cref="PlaceholderKind.None"/> means a freely placed shape.</summary>
        public PlaceholderKind Placeholder { get; set; }
        /// <summary>占位符序号，与版式中的占位符配对。Placeholder index used to pair with the layout placeholder.</summary>
        public int PlaceholderIndex { get; set; }
        /// <summary>填充色，null 表示无填充。Fill color; null means no fill.</summary>
        public OfficeColor? FillColor { get; set; }
        /// <summary>描边样式，null 表示无描边。Outline style; null means no outline.</summary>
        public BorderLine Outline { get; set; }
        /// <summary>入场动画类别；<see cref="SlideAnimationKind.None"/> 表示不参与动画时间线。Entrance animation kind; <see cref="SlideAnimationKind.None"/> keeps the shape out of the timing timeline.</summary>
        public SlideAnimationKind Animation { get; set; }
        /// <summary>动画相对上一动画的延迟，单位为秒。Delay after the previous animation, in seconds.</summary>
        public double AnimationDelaySeconds { get; set; }
        /// <summary>动画持续时间，单位为秒；0 表示使用默认时长。Animation duration in seconds; 0 uses the default duration.</summary>
        public double AnimationDurationSeconds { get; set; }

        /// <summary>创建默认形状的基类实例。Creates the base instance with default values.</summary>
        protected SlideShape()
        {
            Name = string.Empty;
            Placeholder = PlaceholderKind.None;
            PlaceholderIndex = -1;
        }
    }

    /// <summary>
    /// 文本框形状：段落集合与内边距。Text box shape: a paragraph collection plus inner padding.
    /// </summary>
    public class TextBoxShape : SlideShape
    {
        /// <summary>段落列表，复用统一文档模型的段落与内联元素。Paragraphs reusing the unified document model inlines.</summary>
        public List<Paragraph> Paragraphs { get; private set; }
        /// <summary>文本垂直对齐方式。Vertical text alignment.</summary>
        public SlideTextAnchor VerticalAlignment { get; set; }
        /// <summary>左侧内边距，单位为磅。Left inset in points.</summary>
        public double MarginLeft { get; set; }
        /// <summary>右侧内边距，单位为磅。Right inset in points.</summary>
        public double MarginRight { get; set; }
        /// <summary>顶部内边距，单位为磅。Top inset in points.</summary>
        public double MarginTop { get; set; }
        /// <summary>底部内边距，单位为磅。Bottom inset in points.</summary>
        public double MarginBottom { get; set; }
        /// <summary>是否在内容超出时自动缩放文字。Whether the text shrinks to fit the shape.</summary>
        public bool ShrinkToFit { get; set; }

        /// <summary>创建空文本框。Creates an empty text box.</summary>
        public TextBoxShape()
        {
            Paragraphs = new List<Paragraph>();
            MarginLeft = 9.0;
            MarginRight = 9.0;
            MarginTop = 4.5;
            MarginBottom = 4.5;
        }

        /// <summary>按几何位置创建文本框。Creates a text box at the given geometry.</summary>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        public TextBoxShape(double left, double top, double width, double height) : this()
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        /// <summary>追加一个纯文本段落。Appends a plain-text paragraph.</summary>
        /// <param name="text">段落文本。Paragraph text.</param>
        /// <returns>新段落，便于继续设置格式。The new paragraph for further formatting.</returns>
        public Paragraph AddParagraph(string text)
        {
            Paragraph paragraph = new Paragraph(text);
            Paragraphs.Add(paragraph);
            return paragraph;
        }

        /// <summary>追加一个已构建好的段落。Appends an already built paragraph.</summary>
        /// <param name="paragraph">段落对象。The paragraph.</param>
        /// <returns>传入的段落。The supplied paragraph.</returns>
        public Paragraph AddParagraph(Paragraph paragraph)
        {
            if (paragraph == null) throw new ArgumentNullException("paragraph");
            Paragraphs.Add(paragraph);
            return paragraph;
        }

        /// <summary>返回全部段落拼接出的纯文本。Returns the plain text of all paragraphs.</summary>
        /// <returns>以换行符分隔的文本。Text separated by line breaks.</returns>
        public string GetText()
        {
            System.Text.StringBuilder builder = new System.Text.StringBuilder();
            for (int i = 0; i < Paragraphs.Count; i++)
            {
                if (i > 0) builder.Append('\n');
                builder.Append(Paragraphs[i].GetText());
            }
            return builder.ToString();
        }
    }

    /// <summary>
    /// 自选图形：矩形、椭圆等带文本的几何形状。Auto shape: rectangle, ellipse and other geometric shapes carrying text.
    /// </summary>
    public sealed class AutoShape : TextBoxShape
    {
        /// <summary>图形类型。Shape geometry kind.</summary>
        public AutoShapeKind Kind { get; set; }
        /// <summary>圆角半径比例（0-0.5），仅对 <see cref="AutoShapeKind.RoundedRectangle"/> 生效。Corner radius ratio (0-0.5), only for <see cref="AutoShapeKind.RoundedRectangle"/>.</summary>
        public double CornerRadius { get; set; }

        /// <summary>创建默认矩形。Creates a default rectangle.</summary>
        public AutoShape()
        {
            Kind = AutoShapeKind.Rectangle;
            CornerRadius = 0.15;
        }

        /// <summary>按类型与几何位置创建自选图形。Creates an auto shape by kind and geometry.</summary>
        /// <param name="kind">图形类型。Shape kind.</param>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        public AutoShape(AutoShapeKind kind, double left, double top, double width, double height) : this()
        {
            Kind = kind;
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }
    }

    /// <summary>
    /// 图片形状。Picture shape.
    /// </summary>
    public sealed class PictureShape : SlideShape
    {
        /// <summary>图片数据。Image data.</summary>
        public OfficeImage Image { get; set; }
        /// <summary>备选文字（alt text）。Alternative text.</summary>
        public string AlternativeText { get; set; }

        /// <summary>创建空图片形状。Creates an empty picture shape.</summary>
        public PictureShape()
        {
        }

        /// <summary>按图片与几何位置创建图片形状。Creates a picture shape from an image and geometry.</summary>
        /// <param name="image">图片。The image.</param>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        public PictureShape(OfficeImage image, double left, double top, double width, double height)
        {
            Image = image;
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }
    }

    /// <summary>
    /// 表格形状，复用统一文档模型的表格。Table shape reusing the unified document model table.
    /// </summary>
    public sealed class TableShape : SlideShape
    {
        /// <summary>表格内容。Table content.</summary>
        public Table Table { get; set; }
        /// <summary>表头行底色，null 表示不额外着色。Header row background color; null leaves the table styles in charge.</summary>
        public OfficeColor? HeaderColor { get; set; }

        /// <summary>创建空表格形状。Creates an empty table shape.</summary>
        public TableShape()
        {
            Table = new Table();
        }

        /// <summary>按几何位置创建表格形状。Creates a table shape at the given geometry.</summary>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        public TableShape(double left, double top, double width, double height) : this()
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }
    }

    /// <summary>
    /// 单张幻灯片：形状集合、背景、备注与切换效果。One slide: shapes, background, notes and transition.
    /// </summary>
    public sealed class Slide
    {
        /// <summary>幻灯片内的形状列表，按添加顺序即 z 序。Shapes on the slide; insertion order is the z-order.</summary>
        public List<SlideShape> Shapes { get; private set; }
        /// <summary>版式类别。Layout category.</summary>
        public SlideLayoutKind Layout { get; set; }
        /// <summary>背景色，null 表示继承母版。Background color; null inherits the master.</summary>
        public OfficeColor? BackgroundColor { get; set; }
        /// <summary>演讲者备注纯文本。Speaker notes as plain text.</summary>
        public string NotesText { get; set; }
        /// <summary>切换效果。Transition effect.</summary>
        public SlideTransitionKind Transition { get; set; }
        /// <summary>切换持续时间，单位为秒。Transition duration in seconds.</summary>
        public double TransitionDuration { get; set; }
        /// <summary>无人工操作时自动切换前的停留时间，单位为秒；0 表示仅点击切换。Seconds to wait before advancing automatically; 0 means click-only advance.</summary>
        public double AdvanceAfterSeconds { get; set; }

        /// <summary>创建空白幻灯片。Creates a blank slide.</summary>
        public Slide()
        {
            Shapes = new List<SlideShape>();
            Layout = SlideLayoutKind.Blank;
        }

        /// <summary>按版式创建幻灯片。Creates a slide with the given layout.</summary>
        /// <param name="layout">版式类别。Layout category.</param>
        public Slide(SlideLayoutKind layout) : this()
        {
            Layout = layout;
        }

        /// <summary>添加文本框。Adds a text box.</summary>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        /// <param name="text">文本内容。Text content.</param>
        /// <returns>新文本框。The new text box.</returns>
        public TextBoxShape AddTextBox(double left, double top, double width, double height, string text)
        {
            TextBoxShape box = new TextBoxShape(left, top, width, height);
            box.AddParagraph(text);
            Shapes.Add(box);
            return box;
        }

        /// <summary>添加自选图形。Adds an auto shape.</summary>
        /// <param name="kind">图形类型。Shape kind.</param>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        /// <param name="text">图形内文本，可为空。Optional text inside the shape.</param>
        /// <returns>新图形。The new shape.</returns>
        public AutoShape AddShape(AutoShapeKind kind, double left, double top, double width, double height, string text)
        {
            AutoShape shape = new AutoShape(kind, left, top, width, height);
            if (!string.IsNullOrEmpty(text)) shape.AddParagraph(text);
            Shapes.Add(shape);
            return shape;
        }

        /// <summary>添加图片。Adds a picture.</summary>
        /// <param name="image">图片数据。Image data.</param>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        /// <returns>新图片形状。The new picture shape.</returns>
        public PictureShape AddImage(OfficeImage image, double left, double top, double width, double height)
        {
            PictureShape picture = new PictureShape(image, left, top, width, height);
            Shapes.Add(picture);
            return picture;
        }

        /// <summary>添加表格。Adds a table shape.</summary>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        /// <returns>新表格形状。The new table shape.</returns>
        public TableShape AddTable(double left, double top, double width, double height)
        {
            TableShape table = new TableShape(left, top, width, height);
            Shapes.Add(table);
            return table;
        }

        /// <summary>取得标题占位符中的文本。Gets the text of the title placeholder.</summary>
        /// <returns>标题文本；没有标题占位符时返回空字符串。The title text, or an empty string when no title placeholder exists.</returns>
        public string GetTitle()
        {
            for (int i = 0; i < Shapes.Count; i++)
            {
                TextBoxShape box = Shapes[i] as TextBoxShape;
                if (box != null && box.Placeholder == PlaceholderKind.Title) return box.GetText();
            }
            return string.Empty;
        }
    }

    /// <summary>
    /// 演示文稿统一模型（对应 .pptx）。Unified presentation model (the .pptx counterpart).
    /// </summary>
    public sealed class OfficePresentation
    {
        /// <summary>文档元数据。Document metadata.</summary>
        public DocumentMetadata Metadata { get; set; }
        /// <summary>幻灯片宽度，单位为磅。Slide width in points.</summary>
        public double SlideWidth { get; set; }
        /// <summary>幻灯片高度，单位为磅。Slide height in points.</summary>
        public double SlideHeight { get; set; }
        /// <summary>幻灯片列表。Slides.</summary>
        public List<Slide> Slides { get; private set; }
        /// <summary>正文默认字体格式。设置它即视为用户显式指定，优先级高于主题字体；构造函数的内置值不算显式设置。
        /// Default body text format. Assigning it marks the value as user-set, which outranks the theme fonts;
        /// the built-in value from the constructor does not count as explicit.</summary>
        public TextFormat DefaultTextFormat
        {
            get { return _defaultTextFormat; }
            set { _defaultTextFormat = value; _defaultTextFormatSet = value != null; }
        }
        /// <summary>标题默认字体格式，显式设置规则同 <see cref="DefaultTextFormat"/>。
        /// Default title text format; explicit-set rules match <see cref="DefaultTextFormat"/>.</summary>
        public TextFormat DefaultTitleFormat
        {
            get { return _defaultTitleFormat; }
            set { _defaultTitleFormat = value; _defaultTitleFormatSet = value != null; }
        }
        /// <summary>用户是否显式设置过正文默认格式。Whether the user explicitly set the body default format.</summary>
        public bool HasExplicitDefaultTextFormat { get { return _defaultTextFormatSet; } }
        /// <summary>用户是否显式设置过标题默认格式。Whether the user explicitly set the title default format.</summary>
        public bool HasExplicitDefaultTitleFormat { get { return _defaultTitleFormatSet; } }
        /// <summary>主题强调色序列，形状与表格按序取色。Theme accent colors, consumed in order by shapes and tables.</summary>
        public List<OfficeColor> AccentColors { get; private set; }
        /// <summary>主题深改面，null 表示完全使用内置默认。Theme deep-edit surface; null means fully default theme.</summary>
        public PresentationTheme Theme { get; set; }

        private TextFormat _defaultTextFormat;
        private TextFormat _defaultTitleFormat;
        private bool _defaultTextFormatSet;
        private bool _defaultTitleFormatSet;

        /// <summary>
        /// 取正文的实际生效默认格式：用户显式设置优先，其次主题次字体（沿用内置 defaults 的字号与样式），
        /// 最后回退内置默认。
        /// Resolves the effective body default format: an explicit setting wins, then the theme's minor
        /// typeface (keeping the built-in size and style), then the built-in default.
        /// </summary>
        /// <returns>生效的正文默认格式。The effective body default format.</returns>
        public TextFormat GetEffectiveTextFormat()
        {
            return ApplyThemeFont(_defaultTextFormat, _defaultTextFormatSet,
                Theme != null ? Theme.FontMinor : null, false);
        }

        /// <summary>
        /// 取标题的实际生效默认格式，规则同 <see cref="GetEffectiveTextFormat"/>，改用主题主字体。
        /// Resolves the effective title default format the same way as <see cref="GetEffectiveTextFormat"/>,
        /// using the theme's major typeface.
        /// </summary>
        /// <returns>生效的标题默认格式。The effective title default format.</returns>
        public TextFormat GetEffectiveTitleFormat()
        {
            return ApplyThemeFont(_defaultTitleFormat, _defaultTitleFormatSet,
                Theme != null ? Theme.FontMajor : null, true);
        }

        /// <summary>按优先级把主题字体并入内置默认格式；显式设置时原样返回。 Merges a theme typeface into the
        /// built-in default by precedence, returning the format untouched when it was explicitly set.</summary>
        /// <param name="current">当前默认格式。The current default format.</param>
        /// <param name="explicitlySet">是否被显式设置。Whether it was explicitly set.</param>
        /// <param name="themeFont">主题字体名，可为空。The theme typeface; may be empty.</param>
        /// <param name="bold">未显式设置时标题是否需要粗体。Whether the title is bold when not explicit.</param>
        /// <returns>生效格式。The effective format.</returns>
        private TextFormat ApplyThemeFont(TextFormat current, bool explicitlySet, string themeFont, bool bold)
        {
            if (explicitlySet) return current;
            if (string.IsNullOrEmpty(themeFont) || current == null) return current;
            TextFormat themed = current.Clone();
            themed.FontFamily = themeFont;
            if (bold) themed.Bold = true;
            return themed;
        }

        /// <summary>创建 16:9 演示文稿。Creates a 16:9 presentation.</summary>
        public OfficePresentation()
        {
            Metadata = new DocumentMetadata();
            Slides = new List<Slide>();
            SlideWidth = 960.0;
            SlideHeight = 540.0;
            _defaultTextFormat = new TextFormat("Calibri", 18.0);
            _defaultTitleFormat = new TextFormat("Calibri", 40.0) { Bold = true };
            AccentColors = new List<OfficeColor>
            {
                new OfficeColor(0x44, 0x72, 0xC4),
                new OfficeColor(0xED, 0x7D, 0x31),
                new OfficeColor(0xA5, 0xA5, 0xA5),
                new OfficeColor(0xFF, 0xC0, 0x00),
                new OfficeColor(0x5B, 0x9B, 0xD5),
                new OfficeColor(0x70, 0xAD, 0x47)
            };
        }

        /// <summary>16:9 宽屏尺寸（960 x 540 磅）。Widescreen size (960 x 540 points).</summary>
        public static OfficePresentation Widescreen()
        {
            return new OfficePresentation();
        }

        /// <summary>4:3 标准尺寸（720 x 540 磅）。Standard 4:3 size (720 x 540 points).</summary>
        public static OfficePresentation Standard()
        {
            OfficePresentation presentation = new OfficePresentation();
            presentation.SlideWidth = 720.0;
            return presentation;
        }

        /// <summary>追加一张幻灯片。Appends a slide.</summary>
        /// <returns>新幻灯片。The new slide.</returns>
        public Slide AddSlide()
        {
            Slide slide = new Slide();
            Slides.Add(slide);
            return slide;
        }

        /// <summary>按版式追加一张幻灯片，并放入该版式的占位符。Appends a slide of the given layout with its placeholders.</summary>
        /// <param name="layout">版式类别。Layout category.</param>
        /// <returns>新幻灯片。The new slide.</returns>
        public Slide AddSlide(SlideLayoutKind layout)
        {
            Slide slide = new Slide(layout);
            Slides.Add(slide);
            BuildPlaceholders(slide);
            return slide;
        }

        /// <summary>
        /// 依据版式为幻灯片生成默认占位符（标题、正文等），位置按当前幻灯片尺寸等比推算。
        /// 占位符内不预置文本运行，字体由写出器按 <see cref="DefaultTitleFormat"/> 与
        /// <see cref="DefaultTextFormat"/> 决定。
        /// Creates the default placeholders (title, body, and so on) for the slide layout, positioned
        /// proportionally to the current slide size. Placeholders carry no text runs; the writer applies
        /// <see cref="DefaultTitleFormat"/> and <see cref="DefaultTextFormat"/>.
        /// </summary>
        /// <param name="slide">目标幻灯片。Target slide.</param>
        private void BuildPlaceholders(Slide slide)
        {
            double marginX = SlideWidth * 0.08;
            double usable = SlideWidth - marginX * 2.0;
            TextBoxShape title = Placeholder("标题", PlaceholderKind.Title, 0, marginX, SlideHeight * 0.06, usable, SlideHeight * 0.20);

            switch (slide.Layout)
            {
                case SlideLayoutKind.Title:
                    slide.Shapes.Add(title);
                    break;
                case SlideLayoutKind.SectionHeader:
                    title.Top = SlideHeight * 0.34;
                    title.Height = SlideHeight * 0.26;
                    slide.Shapes.Add(title);
                    break;
                case SlideLayoutKind.TwoContent:
                    slide.Shapes.Add(title);
                    slide.Shapes.Add(Placeholder("内容占位符 1", PlaceholderKind.Content, 1, marginX, SlideHeight * 0.30, usable / 2.0 - 9.0, SlideHeight * 0.58));
                    slide.Shapes.Add(Placeholder("内容占位符 2", PlaceholderKind.Content, 2, marginX + usable / 2.0 + 9.0, SlideHeight * 0.30, usable / 2.0 - 9.0, SlideHeight * 0.58));
                    break;
                case SlideLayoutKind.Blank:
                    break;
                default:
                    PlaceholderKind kind = PlaceholderKind.Content;
                    if (slide.Layout == SlideLayoutKind.TitleAndPicture) kind = PlaceholderKind.Picture;
                    if (slide.Layout == SlideLayoutKind.TitleAndTable) kind = PlaceholderKind.Table;
                    slide.Shapes.Add(title);
                    slide.Shapes.Add(Placeholder("内容占位符", kind, 1, marginX, SlideHeight * 0.30, usable, SlideHeight * 0.58));
                    break;
            }
        }

        /// <summary>构造一个占位符文本框。Builds a placeholder text box.</summary>
        /// <param name="name">占位符名称。Placeholder name.</param>
        /// <param name="kind">占位符类型。Placeholder kind.</param>
        /// <param name="index">占位符序号。Placeholder index.</param>
        /// <param name="left">左边距（磅）。Left offset in points.</param>
        /// <param name="top">上边距（磅）。Top offset in points.</param>
        /// <param name="width">宽度（磅）。Width in points.</param>
        /// <param name="height">高度（磅）。Height in points.</param>
        /// <returns>构造出的占位符文本框。The placeholder text box.</returns>
        private static TextBoxShape Placeholder(string name, PlaceholderKind kind, int index, double left, double top, double width, double height)
        {
            TextBoxShape box = new TextBoxShape(left, top, width, height);
            box.Name = name;
            box.Placeholder = kind;
            box.PlaceholderIndex = index;
            box.VerticalAlignment = kind == PlaceholderKind.Title ? SlideTextAnchor.Middle : SlideTextAnchor.Top;
            box.AddParagraph(string.Empty);
            return box;
        }
    }
}
