using System;
using System.Collections.Generic;
using System.Text;

namespace Wisdom.Office.Core
{
    /// <summary>段落水平对齐方式。The horizontal alignment of a paragraph.</summary>
    public enum ParagraphAlignment
    {
        /// <summary>左对齐。Left-aligned.</summary>
        Left = 0,
        /// <summary>居中对齐。Center-aligned.</summary>
        Center = 1,
        /// <summary>右对齐。Right-aligned.</summary>
        Right = 2,
        /// <summary>两端对齐。Justified.</summary>
        Justify = 3
    }

    /// <summary>垂直对齐方式（主要用于表格单元格）。The vertical alignment (primarily for table cells).</summary>
    public enum VerticalAlignment
    {
        /// <summary>顶部对齐。Aligned to the top.</summary>
        Top = 0,
        /// <summary>垂直居中。Vertically centered.</summary>
        Middle = 1,
        /// <summary>底部对齐。Aligned to the bottom.</summary>
        Bottom = 2
    }

    /// <summary>表格在页面中的水平对齐方式。The horizontal alignment of a table on the page.</summary>
    public enum TableAlignment
    {
        /// <summary>表格左对齐。The table is left-aligned.</summary>
        Left = 0,
        /// <summary>表格居中。The table is centered.</summary>
        Center = 1,
        /// <summary>表格右对齐。The table is right-aligned.</summary>
        Right = 2
    }

    /// <summary>边框线型。The border line style.</summary>
    public enum BorderStyle
    {
        /// <summary>无边框。No border.</summary>
        None = 0,
        /// <summary>单实线。A single solid line.</summary>
        Single = 1,
        /// <summary>虚线。A dashed line.</summary>
        Dashed = 2,
        /// <summary>点线。A dotted line.</summary>
        Dotted = 3,
        /// <summary>双线。A double line.</summary>
        Double = 4
    }

    /// <summary>一条边框线的完整描述（线型、宽度、颜色）。A complete description of a single border line (style, width, color).</summary>
    public sealed class BorderLine
    {
        /// <summary>边框线型。The border line style.</summary>
        public BorderStyle Style { get; set; }
        /// <summary>边框线宽，单位为磅（point）。The border line width, in points.</summary>
        public double Width { get; set; }
        /// <summary>边框颜色。The border color.</summary>
        public OfficeColor Color { get; set; }

        /// <summary>默认构造：无线条、0.5 磅宽、黑色。Default constructor: no line, 0.5-point width, black.</summary>
        public BorderLine()
        {
            Style = BorderStyle.None;
            Width = 0.5;
            Color = OfficeColor.Black;
        }

        /// <summary>以指定线型、宽度和颜色创建边框线。Creates a border line with the given style, width and color.</summary>
        /// <param name="style">边框线型。The border line style.</param>
        /// <param name="width">线宽，单位为磅。The line width, in points.</param>
        /// <param name="color">边框颜色。The border color.</param>
        public BorderLine(BorderStyle style, double width, OfficeColor color)
        {
            Style = style;
            Width = width;
            Color = color;
        }

        /// <summary>创建常用细边框：0.5 磅黑色单实线。Creates the common thin border: a 0.5-point black single solid line.</summary>
        /// <returns>新的细边框线实例。A new thin border line instance.</returns>
        public static BorderLine Thin()
        {
            return new BorderLine(BorderStyle.Single, 0.5, OfficeColor.Black);
        }

        /// <summary>返回当前边框线的副本。Returns a copy of the current border line.</summary>
        /// <returns>属性相同的新实例。A new instance with the same properties.</returns>
        public BorderLine Clone()
        {
            return new BorderLine(Style, Width, Color);
        }
    }

    /// <summary>文本字符格式（字体、字号、粗斜体等），供文本运行与默认样式使用。Character formatting for text (font, size, bold/italic, etc.), used by text runs and default styles.</summary>
    public sealed class TextFormat
    {
        /// <summary>统一字体名；未分别指定西文/中文字体时作为回退值。The unified font family; the fallback when no separate Latin/East-Asian font is specified.</summary>
        public string FontFamily { get; set; }
        /// <summary>西文（ASCII）字体名，可为空。The Latin (ASCII) font family; may be empty.</summary>
        public string AsciiFontFamily { get; set; }
        /// <summary>东亚（中文）字体名，可为空。The East-Asian (Chinese) font family; may be empty.</summary>
        public string EastAsiaFontFamily { get; set; }
        /// <summary>字号，单位为磅。The font size, in points.</summary>
        public double FontSize { get; set; }
        /// <summary>是否加粗。Whether the text is bold.</summary>
        public bool Bold { get; set; }
        /// <summary>是否倾斜。Whether the text is italic.</summary>
        public bool Italic { get; set; }
        /// <summary>是否带下划线。Whether the text is underlined.</summary>
        public bool Underline { get; set; }
        /// <summary>是否带删除线。Whether the text is struck through.</summary>
        public bool Strikeout { get; set; }
        /// <summary>字体颜色。The font color.</summary>
        public OfficeColor Color { get; set; }
        /// <summary>是否上标。Whether the text is superscript.</summary>
        public bool Superscript { get; set; }
        /// <summary>是否下标。Whether the text is subscript.</summary>
        public bool Subscript { get; set; }

        /// <summary>默认构造：Arial、11 磅、黑色。Default constructor: Arial, 11 points, black.</summary>
        public TextFormat()
        {
            FontFamily = "Arial";
            FontSize = 11;
            Color = OfficeColor.Black;
        }

        /// <summary>以指定字体与字号创建格式，颜色为黑色。Creates a format with the given font and size, using black.</summary>
        /// <param name="fontFamily">字体名。The font family.</param>
        /// <param name="fontSize">字号，单位为磅。The font size, in points.</param>
        public TextFormat(string fontFamily, double fontSize)
        {
            FontFamily = fontFamily;
            FontSize = fontSize;
            Color = OfficeColor.Black;
        }

        /// <summary>取得实际西文字体名：优先使用 <see cref="AsciiFontFamily"/>，为空时回退到 <see cref="FontFamily"/>。Gets the effective Latin font family: prefers <see cref="AsciiFontFamily"/> and falls back to <see cref="FontFamily"/> when empty.</summary>
        /// <returns>西文字体名。The Latin font family.</returns>
        public string GetAsciiFontFamily()
        {
            return string.IsNullOrEmpty(AsciiFontFamily) ? FontFamily : AsciiFontFamily;
        }

        /// <summary>取得实际中文字体名：优先使用 <see cref="EastAsiaFontFamily"/>，为空时回退到 <see cref="FontFamily"/>。Gets the effective East-Asian font family: prefers <see cref="EastAsiaFontFamily"/> and falls back to <see cref="FontFamily"/> when empty.</summary>
        /// <returns>东亚字体名。The East-Asian font family.</returns>
        public string GetEastAsiaFontFamily()
        {
            return string.IsNullOrEmpty(EastAsiaFontFamily) ? FontFamily : EastAsiaFontFamily;
        }

        /// <summary>浅拷贝返回一个新的格式实例。Returns a new format instance via a shallow copy.</summary>
        /// <returns>属性相同的副本。A copy with the same properties.</returns>
        public TextFormat Clone()
        {
            return (TextFormat)MemberwiseClone();
        }

        /// <summary>按标题级别生成预设格式：加粗、深蓝色，字号随级别递减（一级 20、二级 15、其余 12 磅）。Builds a preset format for a heading level: bold, dark blue, with the size decreasing by level (20 for level 1, 15 for level 2, otherwise 12 points).</summary>
        /// <param name="level">标题级别，1 表示一级标题。The heading level; 1 means a top-level heading.</param>
        /// <returns>标题文本格式。The heading text format.</returns>
        public static TextFormat Heading(int level)
        {
            TextFormat format = new TextFormat("Arial", level <= 1 ? 20 : level == 2 ? 15 : 12);
            format.Bold = true;
            format.Color = new OfficeColor(0x1F, 0x3A, 0x5F);
            return format;
        }
    }

    /// <summary>修订类型：插入或删除。The revision type: an insertion or a deletion.</summary>
    public enum RevisionType
    {
        /// <summary>插入的修订。An insertion revision.</summary>
        Insert = 0,
        /// <summary>删除的修订。A deletion revision.</summary>
        Delete = 1
    }

    /// <summary>一条修订的元数据（类型、作者、时间与原始 ID）。Metadata of a single tracked change: type, author, timestamp and original id.</summary>
    public sealed class RevisionInfo
    {
        /// <summary>修订类型。The revision type.</summary>
        public RevisionType Type { get; set; }
        /// <summary>修订作者。The author of the revision.</summary>
        public string Author { get; set; }
        /// <summary>修订时间，可为 null（原始文档未记录时间）。The revision timestamp, or null when the source document recorded none.</summary>
        public DateTime? Date { get; set; }
        /// <summary>原始文档中的修订 ID（w:id），写回时可重新分配。The revision id from the source document (w:id); a fresh id may be assigned on write.</summary>
        public string Id { get; set; }

        /// <summary>默认构造：插入修订、作者与 ID 为空。Default constructor: an insertion revision with an empty author and id.</summary>
        public RevisionInfo()
        {
            Type = RevisionType.Insert;
            Author = string.Empty;
            Id = string.Empty;
        }

        /// <summary>以类型、作者与时间创建修订元数据。Creates revision metadata from the type, author and timestamp.</summary>
        /// <param name="type">修订类型。The revision type.</param>
        /// <param name="author">修订作者。The revision author.</param>
        /// <param name="date">修订时间，可为 null。The revision timestamp, may be null.</param>
        public RevisionInfo(RevisionType type, string author, DateTime? date)
        {
            Type = type;
            Author = author ?? string.Empty;
            Date = date;
            Id = string.Empty;
        }

        /// <summary>返回修订元数据的副本。Returns a copy of the revision metadata.</summary>
        /// <returns>属性相同的新实例。A new instance with the same properties.</returns>
        public RevisionInfo Clone()
        {
            RevisionInfo copy = new RevisionInfo(Type, Author, Date);
            copy.Id = Id ?? string.Empty;
            return copy;
        }
    }

    /// <summary>段落内行内元素的基类（文本、域、图片等的公共父类）。The base class for inline elements within a paragraph (the common parent of text, fields, images, etc.).</summary>
    public abstract class Inline
    {
        /// <summary>该行内元素所属的修订；null 表示未跟踪修订的普通内容。The tracked revision this inline belongs to; null for normal, untracked content.</summary>
        public RevisionInfo Revision { get; set; }
    }

    /// <summary>一段带格式的纯文本行内元素。A run of formatted plain text as an inline element.</summary>
    public sealed class TextRun : Inline
    {
        /// <summary>文本内容。The text content.</summary>
        public string Text { get; set; }
        /// <summary>文本格式。The text format.</summary>
        public TextFormat Format { get; set; }
        /// <summary>
        /// 超链接目标（URL、mailto: 地址或文档内锚点）；null 表示普通文本。
        /// The hyperlink target (a URL, a mailto: address, or an in-document anchor); null means plain text.
        /// </summary>
        public string Link { get; set; }

        /// <summary>以默认格式包装文本。Wraps text with the default format.</summary>
        /// <param name="text">文本内容，null 视为空字符串。The text content; null is treated as an empty string.</param>
        public TextRun(string text)
            : this(text, new TextFormat())
        {
        }

        /// <summary>以指定格式包装文本。Wraps text with the given format.</summary>
        /// <param name="text">文本内容，null 视为空字符串。The text content; null is treated as an empty string.</param>
        /// <param name="format">文本格式，null 时回退默认格式。The text format; falls back to the default format when null.</param>
        public TextRun(string text, TextFormat format)
        {
            Text = text ?? string.Empty;
            Format = format ?? new TextFormat();
        }
    }

    /// <summary>Word 域行内元素（如页码 PAGE、日期 DATE 等）。A Word field inline element (such as PAGE for page numbers or DATE).</summary>
    public sealed class FieldRun : Inline
    {
        /// <summary>域指令文本，如 "PAGE"。The field instruction text, such as "PAGE".</summary>
        public string Instruction { get; set; }
        /// <summary>域的显示文本（占位结果）。The display text of the field (a placeholder result).</summary>
        public string DisplayText { get; set; }
        /// <summary>域的文本格式。The text format of the field.</summary>
        public TextFormat Format { get; set; }

        /// <summary>仅指定域指令，显示文本为空、格式默认。Specifies only the field instruction, with empty display text and default format.</summary>
        /// <param name="instruction">域指令文本。The field instruction text.</param>
        public FieldRun(string instruction)
            : this(instruction, string.Empty, new TextFormat())
        {
        }

        /// <summary>指定域指令、显示文本与格式。Specifies the field instruction, display text and format.</summary>
        /// <param name="instruction">域指令文本，null 视为空字符串。The field instruction text; null is treated as an empty string.</param>
        /// <param name="displayText">显示文本，null 视为空字符串。The display text; null is treated as an empty string.</param>
        /// <param name="format">文本格式，null 时回退默认格式。The text format; falls back to the default format when null.</param>
        public FieldRun(string instruction, string displayText, TextFormat format)
        {
            Instruction = instruction ?? string.Empty;
            DisplayText = displayText ?? string.Empty;
            Format = format ?? new TextFormat();
        }
    }

    /// <summary>段落内嵌图片行内元素。An inline element for an image embedded within a paragraph.</summary>
    public sealed class ImageRun : Inline
    {
        /// <summary>图片数据。The image data.</summary>
        public OfficeImage Image { get; set; }
        /// <summary>显示宽度，单位为磅；0 表示使用图片默认宽度。The display width, in points; 0 means the image's default width.</summary>
        public double Width { get; set; }
        /// <summary>显示高度，单位为磅；0 表示使用图片默认高度。The display height, in points; 0 means the image's default height.</summary>
        public double Height { get; set; }

        /// <summary>插入图片，尺寸使用图片默认值。Inserts an image using its default size.</summary>
        /// <param name="image">图片数据。The image data.</param>
        public ImageRun(OfficeImage image)
            : this(image, 0, 0)
        {
        }

        /// <summary>插入图片并指定显示尺寸。Inserts an image with an explicit display size.</summary>
        /// <param name="image">图片数据。The image data.</param>
        /// <param name="width">显示宽度（磅），0 表示默认。The display width in points; 0 means the default.</param>
        /// <param name="height">显示高度（磅），0 表示默认。The display height in points; 0 means the default.</param>
        public ImageRun(OfficeImage image, double width, double height)
        {
            Image = image;
            Width = width;
            Height = height;
        }
    }

    /// <summary>文档段落块：由行内元素列表与段落级排版属性组成。A document paragraph block: a list of inline elements plus paragraph-level layout properties.</summary>
    public sealed class Paragraph : Block
    {
        /// <summary>段落包含的行内元素列表。The list of inline elements contained in the paragraph.</summary>
        public List<Inline> Inlines { get; private set; }
        /// <summary>段落水平对齐方式。The horizontal alignment of the paragraph.</summary>
        public ParagraphAlignment Alignment { get; set; }
        /// <summary>段前间距，单位为磅。The space before the paragraph, in points.</summary>
        public double SpaceBefore { get; set; }
        /// <summary>段后间距，单位为磅。The space after the paragraph, in points.</summary>
        public double SpaceAfter { get; set; }
        /// <summary>行距倍数（1.0 表示单倍行距）。The line-spacing multiple (1.0 means single spacing).</summary>
        public double LineSpacing { get; set; }
        /// <summary>左缩进，单位为磅。The left indent, in points.</summary>
        public double LeftIndent { get; set; }
        /// <summary>右缩进，单位为磅。The right indent, in points.</summary>
        public double RightIndent { get; set; }
        /// <summary>首行缩进，单位为磅。The first-line indent, in points.</summary>
        public double FirstLineIndent { get; set; }
        /// <summary>标题级别：0 表示正文，1/2/3… 表示一/二/三级标题。The heading level: 0 is body text, 1/2/3... are first/second/third-level headings.</summary>
        public int HeadingLevel { get; set; }
        /// <summary>是否与下一段保持在同一页（避免孤行）。Whether to keep the paragraph on the same page as the next one (avoiding orphan lines).</summary>
        public bool KeepWithNext { get; set; }
        /// <summary>段落前是否强制分页。Whether a page break is forced before the paragraph.</summary>
        public bool PageBreakBefore { get; set; }
        /// <summary>样式名，导出时可映射到文档内置或自定义样式。The style name; on export it can map to a built-in or custom document style.</summary>
        public string StyleName { get; set; }
        /// <summary>列表标记文本（如 "•"、"1."），为空表示非列表段落。The list marker text (such as "•" or "1."); empty means a non-list paragraph.</summary>
        public string ListMarker { get; set; }
        /// <summary>书签名称，用于生成书签与交叉引用。The bookmark name, used to generate bookmarks and cross references.</summary>
        public string BookmarkName { get; set; }
        /// <summary>段落标记（段落结束符）所属的修订；null 表示段落标记未被跟踪。The revision of the paragraph mark; null when the mark is not tracked.</summary>
        public RevisionInfo Revision { get; set; }
        /// <summary>被删除的行内内容（修订跟踪中保留原文），不参与 <see cref="GetText"/> 与布局。Deleted inline content kept for tracked-change history; it does not contribute to <see cref="GetText"/> or layout.</summary>
        public List<Inline> DeletedInlines { get; private set; }

        /// <summary>默认构造：空段落、单倍行距。Default constructor: an empty paragraph with single line spacing.</summary>
        public Paragraph()
        {
            Inlines = new List<Inline>();
            DeletedInlines = new List<Inline>();
            LineSpacing = 1.0;
        }

        /// <summary>创建包含单个默认格式文本运行的段落。Creates a paragraph containing a single default-formatted text run.</summary>
        /// <param name="text">段落文本。The paragraph text.</param>
        public Paragraph(string text)
            : this()
        {
            AddText(text);
        }

        /// <summary>以默认格式追加文本。Appends text with the default format.</summary>
        /// <param name="text">要追加的文本。The text to append.</param>
        /// <returns>当前段落，便于链式调用。The current paragraph, for fluent chaining.</returns>
        public Paragraph AddText(string text)
        {
            return AddText(text, new TextFormat());
        }

        /// <summary>以指定格式追加文本；空文本被忽略。Appends text with the given format; empty text is ignored.</summary>
        /// <param name="text">要追加的文本。The text to append.</param>
        /// <param name="format">文本格式，null 时回退默认格式。The text format; falls back to the default format when null.</param>
        /// <returns>当前段落，便于链式调用。The current paragraph, for fluent chaining.</returns>
        public Paragraph AddText(string text, TextFormat format)
        {
            if (!string.IsNullOrEmpty(text))
            {
                Inlines.Add(new TextRun(text, format ?? new TextFormat()));
            }
            return this;
        }

        /// <summary>追加已构建的文本运行；null 被忽略。Appends an already-built text run; null is ignored.</summary>
        /// <param name="run">文本运行。The text run.</param>
        /// <returns>当前段落，便于链式调用。The current paragraph, for fluent chaining.</returns>
        public Paragraph AddRun(TextRun run)
        {
            if (run != null)
            {
                Inlines.Add(run);
            }
            return this;
        }

        /// <summary>追加图片行内元素。Appends an image inline element.</summary>
        /// <param name="image">图片数据。The image data.</param>
        /// <param name="width">显示宽度（磅），0 表示默认。The display width in points; 0 means the default.</param>
        /// <param name="height">显示高度（磅），0 表示默认。The display height in points; 0 means the default.</param>
        /// <returns>当前段落，便于链式调用。The current paragraph, for fluent chaining.</returns>
        public Paragraph AddImage(OfficeImage image, double width, double height)
        {
            Inlines.Add(new ImageRun(image, width, height));
            return this;
        }

        /// <summary>追加软换行（以 "\n" 文本运行实现）。Appends a soft line break (implemented as a "\n" text run).</summary>
        /// <returns>当前段落，便于链式调用。The current paragraph, for fluent chaining.</returns>
        public Paragraph AddLineBreak()
        {
            Inlines.Add(new TextRun("\n"));
            return this;
        }

        /// <summary>提取段落纯文本：文本运行取正文、域取显示文本、图片以 [image] 占位。Extracts the paragraph's plain text: text runs contribute their body, fields contribute display text, and images are placeholders written as [image].</summary>
        /// <returns>段落的纯文本表示。The plain-text representation of the paragraph.</returns>
        public string GetText()
        {
            StringBuilder builder = new StringBuilder();
            AppendInlineText(Inlines, builder);
            return builder.ToString();
        }

        /// <summary>提取被修订删除的内容纯文本（不包含当前可见正文）。Extracts the plain text of content removed by tracked changes, excluding the currently visible body.</summary>
        /// <returns>删除内容的纯文本表示。The plain-text representation of the deleted content.</returns>
        public string GetDeletedText()
        {
            StringBuilder builder = new StringBuilder();
            AppendInlineText(DeletedInlines, builder);
            return builder.ToString();
        }

        /// <summary>把行内列表的纯文本追加到构建器：文本取正文、域取显示文本、图片写占位符。Appends the plain text of an inline list to the builder: text runs contribute their body, fields their display text, and images a placeholder.</summary>
        /// <param name="inlines">行内元素列表。The inline elements.</param>
        /// <param name="builder">接收文本的构建器。The builder receiving the text.</param>
        private static void AppendInlineText(List<Inline> inlines, StringBuilder builder)
        {
            for (int i = 0; i < inlines.Count; i++)
            {
                TextRun run = inlines[i] as TextRun;
                if (run != null)
                {
                    builder.Append(run.Text);
                    continue;
                }
                FieldRun field = inlines[i] as FieldRun;
                if (field != null)
                {
                    builder.Append(field.DisplayText);
                }
                else if (inlines[i] is ImageRun)
                {
                    builder.Append("[image]");
                }
            }
        }

        /// <summary>深拷贝段落：复制全部排版属性，并逐项复制行内元素。Deep-copies the paragraph: duplicates all layout properties and copies each inline element.</summary>
        /// <returns>段落的副本。A copy of the paragraph.</returns>
        public Paragraph Clone()
        {
            Paragraph copy = new Paragraph();
            copy.Alignment = Alignment;
            copy.SpaceBefore = SpaceBefore;
            copy.SpaceAfter = SpaceAfter;
            copy.LineSpacing = LineSpacing;
            copy.LeftIndent = LeftIndent;
            copy.RightIndent = RightIndent;
            copy.FirstLineIndent = FirstLineIndent;
            copy.HeadingLevel = HeadingLevel;
            copy.KeepWithNext = KeepWithNext;
            copy.PageBreakBefore = PageBreakBefore;
            copy.StyleName = StyleName;
            copy.ListMarker = ListMarker;
            copy.BookmarkName = BookmarkName;
            copy.Revision = Revision == null ? null : Revision.Clone();
            for (int i = 0; i < Inlines.Count; i++)
            {
                Inline clone = CloneInline(Inlines[i]);
                if (clone != null) copy.Inlines.Add(clone);
            }
            for (int i = 0; i < DeletedInlines.Count; i++)
            {
                Inline clone = CloneInline(DeletedInlines[i]);
                if (clone != null) copy.DeletedInlines.Add(clone);
            }
            return copy;
        }

        /// <summary>复制单个行内元素（含其修订标记），无法识别的类型返回 null。Copies a single inline element including its revision marker; unrecognised types return null.</summary>
        /// <param name="inline">源行内元素。The source inline element.</param>
        /// <returns>行内元素的副本，或 null。A copy of the inline element, or null.</returns>
        private static Inline CloneInline(Inline inline)
        {
            TextRun text = inline as TextRun;
            if (text != null)
            {
                TextRun clone = new TextRun(text.Text, text.Format.Clone());
                clone.Link = text.Link;
                clone.Revision = text.Revision == null ? null : text.Revision.Clone();
                return clone;
            }
            FieldRun field = inline as FieldRun;
            if (field != null)
            {
                FieldRun clone = new FieldRun(field.Instruction, field.DisplayText, field.Format.Clone());
                clone.Revision = field.Revision == null ? null : field.Revision.Clone();
                return clone;
            }
            ImageRun image = inline as ImageRun;
            if (image != null)
            {
                ImageRun clone = new ImageRun(image.Image, image.Width, image.Height);
                clone.Revision = image.Revision == null ? null : image.Revision.Clone();
                return clone;
            }
            BookmarkStart start = inline as BookmarkStart;
            if (start != null) return new BookmarkStart(start.Name);
            BookmarkEnd end = inline as BookmarkEnd;
            if (end != null) return new BookmarkEnd(end.Name);
            return null;
        }
    }

    /// <summary>块级内容基类（段落、表格、分页符等的公共父类）。The base class for block-level content (the common parent of paragraphs, tables, page breaks, etc.).</summary>
    public abstract class Block
    {
    }

    /// <summary>强制分页块。A forced page-break block.</summary>
    public sealed class PageBreak : Block
    {
    }

    /// <summary>表格列定义。The definition of a table column.</summary>
    public sealed class TableColumn
    {
        /// <summary>列宽，单位为磅。The column width, in points.</summary>
        public double Width { get; set; }

        /// <summary>默认构造，列宽为 0（由布局自动计算）。Default constructor with a width of 0 (computed automatically by layout).</summary>
        public TableColumn()
        {
        }

        /// <summary>以指定列宽创建列定义。Creates a column definition with the given width.</summary>
        /// <param name="width">列宽，单位为磅。The column width, in points.</param>
        public TableColumn(double width)
        {
            Width = width;
        }
    }

    /// <summary>表格边框集合：外框四边加内部水平/垂直线。A table's border set: the four outer edges plus internal horizontal/vertical lines.</summary>
    public sealed class TableBorders
    {
        /// <summary>表格上边框。The table's top border.</summary>
        public BorderLine Top { get; set; }
        /// <summary>表格下边框。The table's bottom border.</summary>
        public BorderLine Bottom { get; set; }
        /// <summary>表格左边框。The table's left border.</summary>
        public BorderLine Left { get; set; }
        /// <summary>表格右边框。The table's right border.</summary>
        public BorderLine Right { get; set; }
        /// <summary>内部水平线（行分隔线）。The internal horizontal lines (row separators).</summary>
        public BorderLine InsideHorizontal { get; set; }
        /// <summary>内部垂直线（列分隔线）。The internal vertical lines (column separators).</summary>
        public BorderLine InsideVertical { get; set; }

        /// <summary>默认构造：六个方向均为 0.5 磅黑色单实线。Default constructor: all six directions use a 0.5-point black single solid line.</summary>
        public TableBorders()
        {
            Top = BorderLine.Thin();
            Bottom = BorderLine.Thin();
            Left = BorderLine.Thin();
            Right = BorderLine.Thin();
            InsideHorizontal = BorderLine.Thin();
            InsideVertical = BorderLine.Thin();
        }

        /// <summary>创建六个方向均无边框的边框集合。Creates a border set with no borders in any of the six directions.</summary>
        /// <returns>无边框配置。A borderless configuration.</returns>
        public static TableBorders None()
        {
            TableBorders borders = new TableBorders();
            borders.Top = new BorderLine();
            borders.Bottom = new BorderLine();
            borders.Left = new BorderLine();
            borders.Right = new BorderLine();
            borders.InsideHorizontal = new BorderLine();
            borders.InsideVertical = new BorderLine();
            return borders;
        }
    }

    /// <summary>单元格四边自定义边框，未设置的边可继承表格边框。Per-edge custom borders for a cell; unset edges can inherit the table's borders.</summary>
    public sealed class TableCellBorders
    {
        /// <summary>单元格上边框。The cell's top border.</summary>
        public BorderLine Top { get; set; }
        /// <summary>单元格下边框。The cell's bottom border.</summary>
        public BorderLine Bottom { get; set; }
        /// <summary>单元格左边框。The cell's left border.</summary>
        public BorderLine Left { get; set; }
        /// <summary>单元格右边框。The cell's right border.</summary>
        public BorderLine Right { get; set; }
    }

    /// <summary>表格单元格：内部可包含多个块级元素，并支持行列合并。A table cell: it can contain multiple block elements and supports row/column merging.</summary>
    public sealed class TableCell
    {
        /// <summary>单元格内的块元素列表。The list of block elements inside the cell.</summary>
        public List<Block> Blocks { get; private set; }
        /// <summary>横向合并的单元格数（至少为 1）。The number of horizontally merged cells (at least 1).</summary>
        public int ColSpan { get; set; }
        /// <summary>纵向合并的单元格数（至少为 1）。The number of vertically merged cells (at least 1).</summary>
        public int RowSpan { get; set; }
        /// <summary>单元格背景色，null 表示无填充。The cell's background color; null means no fill.</summary>
        public OfficeColor? BackgroundColor { get; set; }
        /// <summary>单元格四边边框。The cell's four-edge borders.</summary>
        public TableCellBorders Borders { get; set; }
        /// <summary>单元格内容垂直对齐方式。The vertical alignment of the cell content.</summary>
        public VerticalAlignment VerticalAlignment { get; set; }

        /// <summary>默认构造：跨度 1、顶部对齐、无内容。Default constructor: span 1, top-aligned and empty.</summary>
        public TableCell()
        {
            Blocks = new List<Block>();
            ColSpan = 1;
            RowSpan = 1;
            VerticalAlignment = VerticalAlignment.Top;
        }

        /// <summary>创建包含单个文本段落的单元格。Creates a cell containing a single text paragraph.</summary>
        /// <param name="text">单元格文本。The cell text.</param>
        public TableCell(string text)
            : this()
        {
            Blocks.Add(new Paragraph(text));
        }

        /// <summary>向单元格追加一个文本段落。Appends a text paragraph to the cell.</summary>
        /// <param name="text">段落文本。The paragraph text.</param>
        /// <returns>新追加的段落。The newly appended paragraph.</returns>
        public Paragraph AddParagraph(string text)
        {
            Paragraph paragraph = new Paragraph(text);
            Blocks.Add(paragraph);
            return paragraph;
        }
    }

    /// <summary>表格行。A table row.</summary>
    public sealed class TableRow
    {
        /// <summary>行内单元格列表。The list of cells in the row.</summary>
        public List<TableCell> Cells { get; private set; }
        /// <summary>是否为表头行。Whether this is a header row.</summary>
        public bool IsHeader { get; set; }
        /// <summary>行高，单位为磅；0 表示自动高度。The row height, in points; 0 means auto height.</summary>
        public double Height { get; set; }

        /// <summary>默认构造：空行。Default constructor: an empty row.</summary>
        public TableRow()
        {
            Cells = new List<TableCell>();
        }
    }

    /// <summary>表格块级元素：由列定义与行组成，支持边框、内边距、跨页重复表头等。A table block element: composed of column definitions and rows, supporting borders, padding, and header repetition across pages.</summary>
    public sealed class Table : Block
    {
        /// <summary>列定义列表。The list of column definitions.</summary>
        public List<TableColumn> Columns { get; private set; }
        /// <summary>行列表。The list of rows.</summary>
        public List<TableRow> Rows { get; private set; }
        /// <summary>表格边框配置。The table's border configuration.</summary>
        public TableBorders Borders { get; set; }
        /// <summary>单元格内容内边距，单位为磅。The padding inside cells, in points.</summary>
        public double CellPadding { get; set; }
        /// <summary>表格跨页时是否重复表头行。Whether to repeat the header row when the table spans pages.</summary>
        public bool RepeatHeaderRow { get; set; }
        /// <summary>表格在页面中的水平对齐方式。The table's horizontal alignment on the page.</summary>
        public TableAlignment Alignment { get; set; }
        /// <summary>是否按内容自动调整列宽。Whether to auto-fit column widths to the content.</summary>
        public bool Autofit { get; set; }

        /// <summary>默认构造：细边框、内边距 4 磅、重复表头、左对齐、自动列宽。Default constructor: thin borders, 4-point padding, repeated header, left-aligned and auto-fit columns.</summary>
        public Table()
        {
            Columns = new List<TableColumn>();
            Rows = new List<TableRow>();
            Borders = new TableBorders();
            CellPadding = 4;
            RepeatHeaderRow = true;
            Alignment = TableAlignment.Left;
            Autofit = true;
        }

        /// <summary>追加一个空行。Appends an empty row.</summary>
        /// <returns>新追加的行。The newly appended row.</returns>
        public TableRow AddRow()
        {
            TableRow row = new TableRow();
            Rows.Add(row);
            return row;
        }

        /// <summary>按文本值追加一行，每个值生成一个单元格；null 转为空字符串。Appends a row from text values, one cell per value; null becomes an empty string.</summary>
        /// <param name="values">各单元格的文本值。The text values for each cell.</param>
        /// <returns>新追加的行。The newly appended row.</returns>
        public TableRow AddRow(params string[] values)
        {
            TableRow row = new TableRow();
            if (values != null)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    row.Cells.Add(new TableCell(values[i] ?? string.Empty));
                }
            }
            Rows.Add(row);
            return row;
        }
    }

    /// <summary>页面设置：纸张尺寸与页边距，所有长度单位为磅（point）。Page setup: paper size and margins, all lengths in points.</summary>
    public sealed class PageSetup
    {
        /// <summary>页面宽度，单位为磅。The page width, in points.</summary>
        public double Width { get; set; }
        /// <summary>页面高度，单位为磅。The page height, in points.</summary>
        public double Height { get; set; }
        /// <summary>左边距，单位为磅。The left margin, in points.</summary>
        public double MarginLeft { get; set; }
        /// <summary>上边距，单位为磅。The top margin, in points.</summary>
        public double MarginTop { get; set; }
        /// <summary>右边距，单位为磅。The right margin, in points.</summary>
        public double MarginRight { get; set; }
        /// <summary>下边距，单位为磅。The bottom margin, in points.</summary>
        public double MarginBottom { get; set; }
        /// <summary>页眉基线距页面顶部的距离，单位为磅。The distance from the header baseline to the page top, in points.</summary>
        public double HeaderDistance { get; set; }
        /// <summary>页脚基线距页面底部的距离，单位为磅。The distance from the footer baseline to the page bottom, in points.</summary>
        public double FooterDistance { get; set; }

        /// <summary>默认构造：A4 纵向（595.28×841.89 磅），四边距 72 磅。Default constructor: A4 portrait (595.28 by 841.89 points) with 72-point margins on all sides.</summary>
        public PageSetup()
        {
            Width = 595.28;
            Height = 841.89;
            MarginLeft = 72;
            MarginTop = 72;
            MarginRight = 72;
            MarginBottom = 72;
            HeaderDistance = 36;
            FooterDistance = 36;
        }

        /// <summary>是否为横向页面；赋值 true/false 会自动交换宽高，其余情况不变。Whether the page is landscape; setting true or false swaps width and height automatically, other values leave them unchanged.</summary>
        public bool IsLandscape
        {
            get { return Width > Height; }
            set
            {
                if (value && Width < Height)
                {
                    double temp = Width;
                    Width = Height;
                    Height = temp;
                }
                else if (!value && Width > Height)
                {
                    double temp = Width;
                    Width = Height;
                    Height = temp;
                }
            }
        }

        /// <summary>返回 A4 纵向默认页面设置。Returns the default A4 portrait page setup.</summary>
        /// <returns>新的 PageSetup 实例。A new PageSetup instance.</returns>
        public static PageSetup A4()
        {
            return new PageSetup();
        }

        /// <summary>返回 Letter 纸（612×792 磅）页面设置。Returns the Letter page setup (612 by 792 points).</summary>
        /// <returns>新的 PageSetup 实例。A new PageSetup instance.</returns>
        public static PageSetup Letter()
        {
            PageSetup setup = new PageSetup();
            setup.Width = 612;
            setup.Height = 792;
            return setup;
        }

        /// <summary>浅拷贝返回当前页面设置的副本。Returns a copy of the current page setup via a shallow copy.</summary>
        /// <returns>属性相同的新实例。A new instance with the same properties.</returns>
        public PageSetup Clone()
        {
            return (PageSetup)MemberwiseClone();
        }
    }

    /// <summary>页眉或页脚内容容器。A container for header or footer content.</summary>
    public sealed class HeaderFooter
    {
        /// <summary>页眉/页脚内的块元素列表。The list of block elements inside the header/footer.</summary>
        public List<Block> Blocks { get; private set; }
        /// <summary>包含内容时视为启用页眉/页脚。The header/footer is considered enabled when it contains content.</summary>
        public bool IsEnabled { get { return Blocks.Count > 0; } }

        /// <summary>默认构造：内容为空。Default constructor: the content is empty.</summary>
        public HeaderFooter()
        {
            Blocks = new List<Block>();
        }

        /// <summary>追加一个文本段落。Appends a text paragraph.</summary>
        /// <param name="text">段落文本。The paragraph text.</param>
        /// <returns>新追加的段落。The newly appended paragraph.</returns>
        public Paragraph AddParagraph(string text)
        {
            Paragraph paragraph = new Paragraph(text);
            Blocks.Add(paragraph);
            return paragraph;
        }
    }

    /// <summary>文档节：拥有独立的页面设置、正文块与页眉页脚。A document section: it owns its page setup, body blocks and header/footer.</summary>
    public sealed class Section
    {
        /// <summary>本节页面设置。This section's page setup.</summary>
        public PageSetup PageSetup { get; set; }
        /// <summary>节内正文块列表。The list of body blocks within the section.</summary>
        public List<Block> Blocks { get; private set; }
        /// <summary>页眉内容。The header content.</summary>
        public HeaderFooter Header { get; set; }
        /// <summary>页脚内容。The footer content.</summary>
        public HeaderFooter Footer { get; set; }
        /// <summary>是否允许正文超出页面后自动续到新一页；置 false 时内容固定留在当前页（幻灯片转 PDF 用，保证一页幻对应一页）。Whether body content may flow onto auto-started pages when it exceeds the page; false keeps all content on the current page (used for slide-to-PDF so one slide maps to exactly one page).</summary>
        public bool AllowAutoPageBreak { get; set; }

        /// <summary>默认构造：使用默认 A4 页面设置，正文与页眉页脚为空。Default constructor: uses the default A4 page setup with empty body, header and footer.</summary>
        public Section()
        {
            PageSetup = new PageSetup();
            Blocks = new List<Block>();
            Header = new HeaderFooter();
            Footer = new HeaderFooter();
            AllowAutoPageBreak = true;
        }

        /// <summary>以指定页面设置创建节。Creates a section with the given page setup.</summary>
        /// <param name="setup">页面设置，null 时回退默认值。The page setup; falls back to the default when null.</param>
        public Section(PageSetup setup)
            : this()
        {
            PageSetup = setup ?? new PageSetup();
        }

        /// <summary>向节正文追加一个文本段落。Appends a text paragraph to the section body.</summary>
        /// <param name="text">段落文本。The paragraph text.</param>
        /// <returns>新追加的段落。The newly appended paragraph.</returns>
        public Paragraph AddParagraph(string text)
        {
            Paragraph paragraph = new Paragraph(text);
            Blocks.Add(paragraph);
            return paragraph;
        }
    }

    /// <summary>文档核心属性（元数据）。The core document properties (metadata).</summary>
    public sealed class DocumentMetadata
    {
        /// <summary>文档标题。The document title.</summary>
        public string Title { get; set; }
        /// <summary>作者。The author.</summary>
        public string Author { get; set; }
        /// <summary>主题。The subject.</summary>
        public string Subject { get; set; }
        /// <summary>关键词。The keywords.</summary>
        public string Keywords { get; set; }
        /// <summary>创建者应用程序名。The name of the creating application.</summary>
        public string Creator { get; set; }
        /// <summary>创建时间。The creation time.</summary>
        public DateTime Created { get; set; }
        /// <summary>最后修改时间。The last-modified time.</summary>
        public DateTime Modified { get; set; }
        /// <summary>自定义文档属性（标签），支持文本、布尔、整数、浮点、十进制与日期时间值。Custom document properties (tags) supporting text, Boolean, integer, floating-point, decimal and date/time values.</summary>
        public CustomPropertyCollection CustomProperties { get; private set; }

        /// <summary>默认构造：文本属性为空、Creator 为 Wisdom.Office、时间取当前时刻。Default constructor: text properties are empty, Creator is Wisdom.Office, and the times are set to now.</summary>
        public DocumentMetadata()
        {
            Title = string.Empty;
            Author = string.Empty;
            Subject = string.Empty;
            Keywords = string.Empty;
            Creator = "Wisdom.Office";
            Created = DateTime.Now;
            Modified = DateTime.Now;
            CustomProperties = new CustomPropertyCollection();
        }
    }

    /// <summary>统一的内存文档模型（Word/PDF 等导出器共享的中间表示），由若干节组成。The unified in-memory document model (the intermediate representation shared by Word/PDF and other exporters), composed of one or more sections.</summary>
    public sealed class OfficeDocument
    {
        /// <summary>文档元数据。The document metadata.</summary>
        public DocumentMetadata Metadata { get; set; }
        /// <summary>文档节列表。The list of document sections.</summary>
        public List<Section> Sections { get; private set; }
        /// <summary>未显式指定格式时使用的默认文本格式。The default text format used when no format is explicitly given.</summary>
        public TextFormat DefaultTextFormat { get; set; }
        /// <summary>
        /// 文档书签集合，覆盖正文、表格与页眉页脚中的范围书签和整段书签。
        /// The document bookmark collection, covering range and whole-paragraph bookmarks across body,
        /// tables, headers and footers.
        /// </summary>
        public BookmarkCollection Bookmarks { get; private set; }

        /// <summary>默认构造：空文档、默认元数据与文本格式。Default constructor: an empty document with default metadata and text format.</summary>
        public OfficeDocument()
        {
            Metadata = new DocumentMetadata();
            Sections = new List<Section>();
            DefaultTextFormat = new TextFormat();
            Bookmarks = new BookmarkCollection(this);
        }

        /// <summary>追加一个使用默认 A4 设置的节。Appends a section using the default A4 setup.</summary>
        /// <returns>新追加的节。The newly appended section.</returns>
        public Section AddSection()
        {
            return AddSection(new PageSetup());
        }

        /// <summary>以指定页面设置追加一个节。Appends a section with the given page setup.</summary>
        /// <param name="setup">页面设置。The page setup.</param>
        /// <returns>新追加的节。The newly appended section.</returns>
        public Section AddSection(PageSetup setup)
        {
            Section section = new Section(setup);
            Sections.Add(section);
            return section;
        }

        /// <summary>第一个节；尚不存在时自动创建默认节。The first section; a default section is created automatically when none exists.</summary>
        public Section FirstSection
        {
            get
            {
                if (Sections.Count == 0)
                {
                    AddSection();
                }
                return Sections[0];
            }
        }

        /// <summary>将整个文档渲染为纯文本：段落逐行输出，表格行内单元格以制表符分隔。Renders the whole document as plain text: paragraphs are emitted line by line and table cells within a row are separated by tabs.</summary>
        /// <returns>文档的纯文本表示。The plain-text representation of the document.</returns>
        public string GetPlainText()
        {
            StringBuilder builder = new StringBuilder();
            for (int s = 0; s < Sections.Count; s++)
            {
                Section section = Sections[s];
                for (int i = 0; i < section.Blocks.Count; i++)
                {
                    Paragraph paragraph = section.Blocks[i] as Paragraph;
                    if (paragraph != null)
                    {
                        builder.AppendLine(paragraph.GetText());
                        continue;
                    }
                    Table table = section.Blocks[i] as Table;
                    if (table != null)
                    {
                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                            {
                                if (c > 0)
                                {
                                    builder.Append('\t');
                                }
                                TableCell cell = table.Rows[r].Cells[c];
                                for (int b = 0; b < cell.Blocks.Count; b++)
                                {
                                    Paragraph cellParagraph = cell.Blocks[b] as Paragraph;
                                    if (cellParagraph != null)
                                    {
                                        builder.Append(cellParagraph.GetText());
                                    }
                                }
                            }
                            builder.AppendLine();
                        }
                    }
                }
            }
            return builder.ToString();
        }

        /// <summary>由纯文本创建单节文档：按行拆分（兼容 \r\n、\r、\n），每行一个段落。Creates a single-section document from plain text: splits by lines (accepting \r\n, \r and \n), one paragraph per line.</summary>
        /// <param name="text">纯文本内容，null 视为空字符串。The plain-text content; null is treated as an empty string.</param>
        /// <returns>新建的文档。The newly created document.</returns>
        public static OfficeDocument FromText(string text)
        {
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();
            string[] lines = (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                section.AddParagraph(lines[i]);
            }
            return document;
        }
    }
}
