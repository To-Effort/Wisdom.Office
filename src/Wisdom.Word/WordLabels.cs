using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 标签纸版式：页面尺寸、行列数、单张标签尺寸、间隙与页边距，所有长度单位为磅。
    /// Label-sheet layout: page size, row/column count, label size, gaps and page margins, with every
    /// length expressed in points.
    /// </summary>
    /// <remarks>
    /// 使用 <see cref="FromMillimeters"/> 以毫米创建自定义版式；内置 <see cref="Avery5160"/>、
    /// <see cref="Avery5163"/> 与 <see cref="AveryL7160"/> 三种常用规格。生成器按版式放置无边框表格，
    /// 每页 <see cref="LabelsPerPage"/> 张标签。
    /// Use FromMillimeters to build a custom layout in millimetres; Avery5160, Avery5163 and AveryL7160 are
    /// built in. The generator lays out a borderless table per page with LabelsPerPage labels.
    /// </remarks>
    public sealed class WordLabelLayout
    {
        /// <summary>版式名称。The layout name.</summary>
        public string Name { get; set; }
        /// <summary>纸张宽度，单位为磅。The page width, in points.</summary>
        public double PageWidth { get; set; }
        /// <summary>纸张高度，单位为磅。The page height, in points.</summary>
        public double PageHeight { get; set; }
        /// <summary>每行标签数。The number of labels per row.</summary>
        public int Columns { get; set; }
        /// <summary>每列标签数。The number of labels per column.</summary>
        public int Rows { get; set; }
        /// <summary>单张标签宽度，单位为磅。The width of one label, in points.</summary>
        public double LabelWidth { get; set; }
        /// <summary>单张标签高度，单位为磅。The height of one label, in points.</summary>
        public double LabelHeight { get; set; }
        /// <summary>标签之间的水平间隙，单位为磅。The horizontal gap between labels, in points.</summary>
        public double HorizontalGap { get; set; }
        /// <summary>标签之间的垂直间隙，单位为磅。The vertical gap between labels, in points.</summary>
        public double VerticalGap { get; set; }
        /// <summary>第一行标签距页面顶部的距离，单位为磅。The distance from the page top to the first label row, in points.</summary>
        public double TopMargin { get; set; }
        /// <summary>第一列标签距页面左侧的距离，单位为磅。The distance from the page left edge to the first label column, in points.</summary>
        public double LeftMargin { get; set; }
        /// <summary>标签内容与单元格边缘的内边距，单位为磅。The content inset from the label edges, in points.</summary>
        public double CellPadding { get; set; }

        /// <summary>每页标签数（列数乘行数）。The number of labels per page (columns times rows).</summary>
        public int LabelsPerPage
        {
            get { return Math.Max(1, Columns) * Math.Max(1, Rows); }
        }

        /// <summary>默认构造：A4 三列七行（Avery L7160 规格），内边距 2.88 磅。Default constructor: A4 with three columns and seven rows (Avery L7160 dimensions) and a 2.88-point content inset.</summary>
        public WordLabelLayout()
        {
            Name = "Avery L7160";
            PageWidth = MillimetersToPoints(210);
            PageHeight = MillimetersToPoints(297);
            Columns = 3;
            Rows = 7;
            LabelWidth = MillimetersToPoints(63.5);
            LabelHeight = MillimetersToPoints(38.1);
            HorizontalGap = MillimetersToPoints(2.5);
            VerticalGap = 0;
            TopMargin = MillimetersToPoints(15.1);
            LeftMargin = MillimetersToPoints(7);
            CellPadding = MillimetersToPoints(1);
        }

        /// <summary>
        /// 以毫米为单位创建自定义版式；内部会换算成磅。
        /// Creates a custom layout expressed in millimetres; values are converted to points internally.
        /// </summary>
        /// <param name="name">版式名称。The layout name.</param>
        /// <param name="pageWidthMm">纸张宽度（毫米）。The page width in millimetres.</param>
        /// <param name="pageHeightMm">纸张高度（毫米）。The page height in millimetres.</param>
        /// <param name="columns">每行标签数。The number of labels per row.</param>
        /// <param name="rows">每列标签数。The number of labels per column.</param>
        /// <param name="labelWidthMm">标签宽度（毫米）。The label width in millimetres.</param>
        /// <param name="labelHeightMm">标签高度（毫米）。The label height in millimetres.</param>
        /// <param name="topMarginMm">首行标签距页面顶部的距离（毫米）。The distance from the page top to the first label row, in millimetres.</param>
        /// <param name="leftMarginMm">首列标签距页面左侧的距离（毫米）。The distance from the page left edge to the first label column, in millimetres.</param>
        /// <param name="horizontalGapMm">标签水平间隙（毫米）。The horizontal gap between labels, in millimetres.</param>
        /// <param name="verticalGapMm">标签垂直间隙（毫米）。The vertical gap between labels, in millimetres.</param>
        /// <returns>新的标签版式。The new label layout.</returns>
        public static WordLabelLayout FromMillimeters(
            string name,
            double pageWidthMm,
            double pageHeightMm,
            int columns,
            int rows,
            double labelWidthMm,
            double labelHeightMm,
            double topMarginMm,
            double leftMarginMm,
            double horizontalGapMm = 0,
            double verticalGapMm = 0)
        {
            return new WordLabelLayout
            {
                Name = name ?? string.Empty,
                PageWidth = MillimetersToPoints(pageWidthMm),
                PageHeight = MillimetersToPoints(pageHeightMm),
                Columns = columns,
                Rows = rows,
                LabelWidth = MillimetersToPoints(labelWidthMm),
                LabelHeight = MillimetersToPoints(labelHeightMm),
                HorizontalGap = MillimetersToPoints(horizontalGapMm),
                VerticalGap = MillimetersToPoints(verticalGapMm),
                TopMargin = MillimetersToPoints(topMarginMm),
                LeftMargin = MillimetersToPoints(leftMarginMm),
                CellPadding = MillimetersToPoints(1)
            };
        }

        /// <summary>Avery 5160：US Letter，3 列 × 10 行，标签 2.625 × 1 英寸。Avery 5160: US Letter, 3 columns by 10 rows, 2.625 by 1 inch labels.</summary>
        /// <returns>新的标签版式。The new label layout.</returns>
        public static WordLabelLayout Avery5160()
        {
            return FromMillimeters("Avery 5160", 215.9, 279.4, 3, 10, 66.675, 25.4, 12.7, 4.7625, 3.175, 0);
        }

        /// <summary>Avery 5163：US Letter，2 列 × 5 行，标签 4 × 2 英寸。Avery 5163: US Letter, 2 columns by 5 rows, 4 by 2 inch labels.</summary>
        /// <returns>新的标签版式。The new label layout.</returns>
        public static WordLabelLayout Avery5163()
        {
            return FromMillimeters("Avery 5163", 215.9, 279.4, 2, 5, 101.6, 50.8, 12.7, 3.96875, 4.7625, 0);
        }

        /// <summary>Avery L7160：A4，3 列 × 7 行，标签 63.5 × 38.1 毫米。Avery L7160: A4, 3 columns by 7 rows, 63.5 by 38.1 mm labels.</summary>
        /// <returns>新的标签版式。The new label layout.</returns>
        public static WordLabelLayout AveryL7160()
        {
            return FromMillimeters("Avery L7160", 210, 297, 3, 7, 63.5, 38.1, 15.1, 7, 2.5, 0);
        }

        /// <summary>
        /// 生成标签纸的页面设置：左右上下边距精确对齐首张标签，页眉页脚距离置零。
        /// Builds the page setup of the label sheet: margins align with the first label and header/footer
        /// distances are zero.
        /// </summary>
        /// <returns>新的页面设置。The new page setup.</returns>
        public PageSetup CreatePageSetup()
        {
            Validate();
            double usedWidth = Columns * LabelWidth + Math.Max(0, Columns - 1) * HorizontalGap;
            double usedHeight = Rows * LabelHeight + Math.Max(0, Rows - 1) * VerticalGap;
            PageSetup setup = new PageSetup();
            setup.Width = PageWidth;
            setup.Height = PageHeight;
            setup.MarginLeft = LeftMargin;
            setup.MarginTop = TopMargin;
            setup.MarginRight = Math.Max(0, PageWidth - LeftMargin - usedWidth);
            setup.MarginBottom = Math.Max(0, PageHeight - TopMargin - usedHeight);
            setup.HeaderDistance = 0;
            setup.FooterDistance = 0;
            return setup;
        }

        /// <summary>
        /// 校验版式的数值：尺寸为正、行列至少为 1、标签区域不超出纸张。
        /// Validates the layout: positive sizes, at least one row and column, and a label grid that fits the page.
        /// </summary>
        /// <exception cref="ArgumentException">版式数值非法。Thrown when a layout value is invalid.</exception>
        public void Validate()
        {
            if (PageWidth <= 0 || PageHeight <= 0)
            {
                throw new ArgumentException("Label page width and height must be positive.");
            }
            if (Columns < 1 || Rows < 1)
            {
                throw new ArgumentException("A label layout needs at least one column and one row.");
            }
            if (LabelWidth <= 0 || LabelHeight <= 0)
            {
                throw new ArgumentException("Label width and height must be positive.");
            }
            if (HorizontalGap < 0 || VerticalGap < 0 || TopMargin < 0 || LeftMargin < 0 || CellPadding < 0)
            {
                throw new ArgumentException("Label gaps, margins and padding must not be negative.");
            }
            double usedWidth = Columns * LabelWidth + Math.Max(0, Columns - 1) * HorizontalGap;
            double usedHeight = Rows * LabelHeight + Math.Max(0, Rows - 1) * VerticalGap;
            if (LeftMargin + usedWidth > PageWidth + 0.5)
            {
                throw new ArgumentException("The label grid is wider than the page.");
            }
            if (TopMargin + usedHeight > PageHeight + 0.5)
            {
                throw new ArgumentException("The label grid is taller than the page.");
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
    /// 标签文字选项：字体、字号、加粗、行距、对齐与模板文化。
    /// Label text options: font, size, bold, line spacing, alignment and template culture.
    /// </summary>
    public sealed class WordLabelOptions
    {
        /// <summary>字体名，同时作为未单独指定中西文字体时的回退值。The font family, also the fallback when separate Latin/East-Asian fonts are not set.</summary>
        public string FontFamily { get; set; }
        /// <summary>中文字体名，为空时回退到 <see cref="FontFamily"/>。The East-Asian font family; falls back to FontFamily when empty.</summary>
        public string EastAsiaFontFamily { get; set; }
        /// <summary>字号，单位为磅。The font size, in points.</summary>
        public double FontSize { get; set; }
        /// <summary>是否加粗。Whether the label text is bold.</summary>
        public bool Bold { get; set; }
        /// <summary>行距倍数。The line-spacing multiple.</summary>
        public double LineSpacing { get; set; }
        /// <summary>段落对齐方式。The paragraph alignment.</summary>
        public ParagraphAlignment Alignment { get; set; }
        /// <summary>模板中无法解析的标记是否原样保留，默认 false（置空）。Whether unresolved template tokens are kept; false by default (removed).</summary>
        public bool LeaveUnresolvedTokens { get; set; }
        /// <summary>模板格式化的区域文化，默认不变文化。The culture used for template formatting; invariant by default.</summary>
        public CultureInfo Culture { get; set; }
        /// <summary>是否绘制标签边框，默认 false（无边框）。Whether label borders are drawn; false by default (borderless).</summary>
        public bool ShowCellBorders { get; set; }
        /// <summary>
        /// 是否在标签的每行文字之间绘制分隔线，默认 false；启用后标签内容放入一列多行嵌套表格，
        /// 各行等高并铺满分隔线，配合 <see cref="ShowCellBorders"/> 形成带内线的完整标签框。
        /// Whether to draw a separator between label text lines; false by default. When enabled the label
        /// content becomes a one-column, multi-row nested table with equal row heights and full-width
        /// separators, so combined with ShowCellBorders it forms a boxed label with internal rules.
        /// </summary>
        public bool ShowRowSeparators { get; set; }
        /// <summary>标签边框与表单填写线的颜色，默认黑色。The colour of label borders and form rules; black by default.</summary>
        public OfficeColor BorderColor { get; set; }
        /// <summary>标签边框与表单填写线的宽度（磅），默认 0.5。The width of label borders and form rules in points; 0.5 by default.</summary>
        public double BorderWidth { get; set; }
        /// <summary>
        /// 表单标签中每条填写行的垂直间距（磅），默认 6；填写线画在行底，文字紧贴线的上方。
        /// The vertical spacing of each fill-in row in form labels, in points; 6 by default. The rule is drawn
        /// at the row bottom so the text sits directly above it.
        /// </summary>
        public double FieldSpacing { get; set; }

        /// <summary>默认构造：宋体 10.5 磅、单倍行距、左对齐、未解析标记置空、黑色 0.5 磅边框、行间距 6 磅。Default constructor: SimSun at 10.5 points, single spacing, left aligned, unresolved tokens removed, black 0.5-point borders and a 6-point field spacing.</summary>
        public WordLabelOptions()
        {
            FontFamily = "SimSun";
            EastAsiaFontFamily = "SimSun";
            FontSize = 10.5;
            LineSpacing = 1.0;
            Alignment = ParagraphAlignment.Left;
            Culture = CultureInfo.InvariantCulture;
            BorderColor = OfficeColor.Black;
            BorderWidth = 0.5;
            FieldSpacing = 6;
        }
    }

    /// <summary>
    /// 邮件标签生成器：把多行文本或数据记录排成标准标签纸，支持按页自动分页。
    /// Mailing-label generator: lays out multi-line text or data records on standard label sheets with
    /// automatic pagination.
    /// </summary>
    /// <remarks>
    /// 输出使用无边框固定列宽表格，每页一张表；页与页之间插入分页符。标签内容超出单张高度时行高自动扩展。
    /// The output uses a borderless fixed-width table per page with page breaks between pages; a row grows
    /// when its content exceeds the label height.
    /// </remarks>
    public static class WordLabels
    {
        /// <summary>
        /// 生成空白标签纸（一页，所有标签为空）。
        /// Creates a blank label sheet (one page with empty labels).
        /// </summary>
        /// <param name="layout">标签版式。The label layout.</param>
        /// <returns>标签文档模型。The label document model.</returns>
        public static OfficeDocument CreateBlank(WordLabelLayout layout)
        {
            return Create(layout, new List<string[]>(), null);
        }

        /// <summary>
        /// 生成带填写横线的表单标签纸（如学生信息标签）：每张标签依次输出“标签文字 + 横线”，
        /// 横线从文字右侧延伸到标签框右侧、并留出 6 磅空白。
        /// Creates a form label sheet with fill-in rules (for example student information labels): every label
        /// renders each label text followed by a rule stretching from the text to the right box edge, leaving a
        /// six-point margin.
        /// </summary>
        /// <param name="layout">标签版式。The label layout.</param>
        /// <param name="fieldLabels">每行填写线的标签文字，如 "姓名："、"年级："。The label text of each fill-in row, such as "姓名：" or "年级：".</param>
        /// <param name="options">文字、边框与行距选项，null 时使用默认值。Text, border and spacing options; defaults are used when null.</param>
        /// <returns>表单标签文档模型。The form label document model.</returns>
        /// <exception cref="ArgumentException"><paramref name="fieldLabels"/> 为空或只含空字符串。Thrown when <paramref name="fieldLabels"/> is empty or contains only blank labels.</exception>
        public static OfficeDocument CreateForm(
            WordLabelLayout layout,
            IEnumerable<string> fieldLabels,
            WordLabelOptions options = null)
        {
            if (layout == null) throw new ArgumentNullException("layout");
            if (fieldLabels == null) throw new ArgumentNullException("fieldLabels");
            layout.Validate();
            options = options ?? new WordLabelOptions();

            List<string> labels = new List<string>();
            foreach (string label in fieldLabels)
            {
                if (!string.IsNullOrWhiteSpace(label)) labels.Add(label);
            }
            if (labels.Count == 0)
            {
                throw new ArgumentException("A form label sheet needs at least one non-empty field label.", "fieldLabels");
            }
            return Build(layout, new List<string[]>(), options, labels);
        }

        /// <summary>
        /// 按已排好的多行文本生成标签纸；标签按行优先顺序填充，超出每页容量时自动分页。
        /// Creates a label sheet from pre-formatted multi-line labels; labels fill row-major order and
        /// overflow onto new pages automatically.
        /// </summary>
        /// <param name="layout">标签版式。The label layout.</param>
        /// <param name="labels">每个标签的多行文本；null 元素按空标签处理。The lines of each label; null entries become empty labels.</param>
        /// <param name="options">文字选项，null 时使用默认值。The text options; defaults are used when null.</param>
        /// <returns>标签文档模型。The label document model.</returns>
        public static OfficeDocument Create(
            WordLabelLayout layout,
            IEnumerable<string[]> labels,
            WordLabelOptions options = null)
        {
            if (layout == null) throw new ArgumentNullException("layout");
            if (labels == null) throw new ArgumentNullException("labels");
            layout.Validate();
            options = options ?? new WordLabelOptions();

            List<string[]> contents = new List<string[]>();
            foreach (string[] label in labels)
            {
                contents.Add(label ?? new string[0]);
            }
            return Build(layout, contents, options);
        }

        /// <summary>
        /// 用数据记录与多行模板生成标签纸：模板按换行拆成多行，每行内的 {{字段}} 标记按记录解析。
        /// Creates a label sheet from data records and a multi-line template: the template is split on line
        /// breaks and {{field}} tokens are resolved against each record.
        /// </summary>
        /// <typeparam name="T">记录类型，可为字典、DataRow、DataRowView、TemplateContext 或普通对象。The record type: a dictionary, DataRow, DataRowView, TemplateContext or plain object.</typeparam>
        /// <param name="layout">标签版式。The label layout.</param>
        /// <param name="records">数据记录集合。The data records.</param>
        /// <param name="template">多行模板，例如 "{{name}}\n{{address}}"。The multi-line template, such as "{{name}}\n{{address}}".</param>
        /// <param name="options">文字与模板选项，null 时使用默认值。Text and template options; defaults are used when null.</param>
        /// <returns>标签文档模型。The label document model.</returns>
        public static OfficeDocument CreateFromRecords<T>(
            WordLabelLayout layout,
            IEnumerable<T> records,
            string template,
            WordLabelOptions options = null)
        {
            if (layout == null) throw new ArgumentNullException("layout");
            if (records == null) throw new ArgumentNullException("records");
            if (template == null) throw new ArgumentNullException("template");
            layout.Validate();
            options = options ?? new WordLabelOptions();

            List<string[]> contents = new List<string[]>();
            TemplateOptions templateOptions = new TemplateOptions
            {
                LeaveUnresolvedTokens = options.LeaveUnresolvedTokens,
                Culture = options.Culture ?? CultureInfo.InvariantCulture
            };
            foreach (T record in records)
            {
                TemplateContext context = WordTemplateText.CreateContext(record);
                contents.Add(WordTemplateText.ReplaceLines(template, context, templateOptions));
            }
            return Build(layout, contents, options);
        }

        /// <summary>
        /// 组装标签文档：逐页生成无边框表格，页间插入分页符；指定表单字段时每张标签输出填写行。
        /// Assembles the label document: one borderless table per page with page breaks between them; when form
        /// fields are supplied, every label renders fill-in rows.
        /// </summary>
        /// <param name="layout">标签版式。The label layout.</param>
        /// <param name="contents">标签内容。The label contents.</param>
        /// <param name="options">文字选项。The text options.</param>
        /// <param name="formFields">表单填写行的标签文字，null 表示普通文本标签。The label text of form fill-in rows, or null for plain text labels.</param>
        /// <returns>标签文档模型。The label document model.</returns>
        private static OfficeDocument Build(
            WordLabelLayout layout,
            List<string[]> contents,
            WordLabelOptions options,
            List<string> formFields = null)
        {
            OfficeDocument document = new OfficeDocument();
            PageSetup setup = layout.CreatePageSetup();
            if (formFields != null || options.ShowRowSeparators)
            {
                // Word 会在每个单元格末尾的嵌套表格后补一个空段落；把版心下沿下移一段，
                // 让这些不可见的段落留在页内，同时不改变标签框的位置和尺寸。
                // Word appends an empty paragraph after each nested table that ends a cell; extend the text
                // area by one reserve so those invisible paragraphs stay on the page without moving or
                // resizing the label boxes.
                setup.MarginBottom = Math.Max(0, setup.MarginBottom - 18);
            }
            Section section = document.AddSection(setup);
            int perPage = layout.LabelsPerPage;
            int pageCount = Math.Max(1, (contents.Count + perPage - 1) / perPage);
            for (int page = 0; page < pageCount; page++)
            {
                if (page > 0) section.Blocks.Add(new PageBreak());
                section.Blocks.Add(CreatePageTable(layout, contents, page * perPage, options, formFields));
            }
            return document;
        }

        /// <summary>
        /// 生成单页标签表格：列宽固定、无边框，标签之间按版式留出间隙列。
        /// Builds one page of labels: fixed column widths, no borders, and gap columns between labels.
        /// </summary>
        /// <param name="layout">标签版式。The label layout.</param>
        /// <param name="contents">全部标签内容。All label contents.</param>
        /// <param name="start">本页第一个标签的索引。The index of the first label on this page.</param>
        /// <param name="options">文字选项。The text options.</param>
        /// <param name="formFields">表单填写行的标签文字，null 表示普通文本标签。The form field labels, or null for plain text labels.</param>
        /// <returns>标签表格。The label table.</returns>
        private static Table CreatePageTable(
            WordLabelLayout layout,
            List<string[]> contents,
            int start,
            WordLabelOptions options,
            List<string> formFields)
        {
            bool form = formFields != null;
            Table table = new Table();
            table.Borders = TableBorders.None();
            // 标签内容由内边距或嵌套表格定位；页面表格本身不留边距，填写横线才能按标签宽度伸展。
            // Label content is positioned by padding or the nested table; the page table itself has no inset
            // so the fill-in rules can span the label width.
            table.CellPadding = 0;
            table.Autofit = false;
            table.Alignment = TableAlignment.Left;
            table.RepeatHeaderRow = false;

            for (int column = 0; column < layout.Columns; column++)
            {
                table.Columns.Add(new TableColumn(layout.LabelWidth));
                if (column < layout.Columns - 1 && layout.HorizontalGap > 0)
                {
                    table.Columns.Add(new TableColumn(layout.HorizontalGap));
                }
            }

            TextFormat format = new TextFormat(options.FontFamily, options.FontSize);
            format.EastAsiaFontFamily = options.EastAsiaFontFamily;
            format.Bold = options.Bold;

            // 分隔线模式让同一页的所有标签拥有相同的行数，空缺行用空段落补齐，标签框才整齐一致。
            // Separator mode gives every label on the page the same row count, filling missing rows with
            // empty paragraphs so the boxes stay aligned.
            int separatorRows = 1;
            if (!form && options.ShowRowSeparators)
            {
                int pageEnd = Math.Min(contents.Count, start + layout.LabelsPerPage);
                for (int index = start; index < pageEnd; index++)
                {
                    string[] pageLines = contents[index];
                    int count = pageLines == null ? 1 : Math.Max(1, pageLines.Length);
                    if (count > separatorRows) separatorRows = count;
                }
            }

            for (int rowIndex = 0; rowIndex < layout.Rows; rowIndex++)
            {
                TableRow row = table.AddRow();
                // 垂直间隙并入行高（最后一行除外），避免用空段落垫高时被默认字号撑开。
                // The vertical gap is folded into the row height (except the last row) so that spacer
                // paragraphs cannot inflate the gap with their default font size.
                row.Height = layout.LabelHeight
                    + (layout.VerticalGap > 0 && rowIndex < layout.Rows - 1 ? layout.VerticalGap : 0);
                for (int column = 0; column < layout.Columns; column++)
                {
                    int index = start + rowIndex * layout.Columns + column;
                    string[] lines = index < contents.Count ? contents[index] : null;
                    TableCell cell = new TableCell();
                    cell.VerticalAlignment = VerticalAlignment.Top;
                    // 表单模式把外框画在标签内部的嵌套表格上，这样垂直间隙落在框外，标签之间才有可见留白。
                    // Form mode draws the box on the nested table inside the label so the vertical gap stays
                    // outside the box and labels are visibly separated.
                    if (options.ShowCellBorders && !form && !options.ShowRowSeparators)
                    {
                        ApplyCellBorders(cell, options);
                    }
                    if (form) AppendForm(cell, formFields, format, options, layout);
                    else if (options.ShowRowSeparators)
                    {
                        AppendSeparatedLines(cell, lines, separatorRows, format, options, layout);
                    }
                    else
                    {
                        AppendLines(cell, lines, format, options, layout.CellPadding);
                    }
                    row.Cells.Add(cell);
                    if (column < layout.Columns - 1 && layout.HorizontalGap > 0)
                    {
                        row.Cells.Add(new TableCell());
                    }
                }
            }
            return table;
        }

        /// <summary>
        /// 给标签单元格加四边细边框；标签列之间的间隙单元格保持无边框。
        /// Applies thin borders to a label cell; the gap cells between label columns stay borderless.
        /// </summary>
        /// <param name="cell">标签单元格。The label cell.</param>
        /// <param name="options">字符与边框选项。The text and border options.</param>
        private static void ApplyCellBorders(TableCell cell, WordLabelOptions options)
        {
            double width = options.BorderWidth > 0 ? options.BorderWidth : 0.5;
            TableCellBorders borders = new TableCellBorders();
            borders.Top = new BorderLine(BorderStyle.Single, width, options.BorderColor);
            borders.Bottom = new BorderLine(BorderStyle.Single, width, options.BorderColor);
            borders.Left = new BorderLine(BorderStyle.Single, width, options.BorderColor);
            borders.Right = new BorderLine(BorderStyle.Single, width, options.BorderColor);
            cell.Borders = borders;
        }

        /// <summary>
        /// 向标签单元格写入表单填写行：输出一个三列嵌套表格（字段名、填写线、右侧留白），每个字段占一行，
        /// 表尾用空白行补足标签高度；左侧放标签文字，中列为空白单元格并只画下边框，
        /// 行高取“字高 + FieldSpacing”，文字与填写线都底对齐。启用边框时由嵌套表格的四周外框构成标签框，
        /// 因此垂直间隙留在框外，标签之间保持可见留白。
        /// Writes form fill-in rows into a label cell: a three-column nested table (label text, fill-in rule and
        /// right inset) with one row per field plus a blank filler row completing the label height. The left cell
        /// holds the label text, the middle cell is blank with only a bottom border, the row height is font line
        /// height plus FieldSpacing, and the text and rule are bottom-aligned. When borders are enabled the
        /// table's outer edges form the label box, so the vertical gap stays outside the box and labels remain
        /// visibly separated.
        /// </summary>
        /// <param name="cell">标签单元格。The label cell.</param>
        /// <param name="fieldLabels">填写行的标签文字。The label text of each fill-in row.</param>
        /// <param name="format">字符格式。The character format.</param>
        /// <param name="options">文字与边框选项。The text and border options.</param>
        /// <param name="layout">标签版式，提供标签尺寸。The label layout supplying the label dimensions.</param>
        private static void AppendForm(
            TableCell cell,
            List<string> fieldLabels,
            TextFormat format,
            WordLabelOptions options,
            WordLabelLayout layout)
        {
            double padding = Math.Max(0, layout.CellPadding);
            double boxWidth = Math.Max(8, layout.LabelWidth);
            double contentWidth = Math.Max(1, boxWidth - padding * 2);
            double estimatedLabelWidth = 0;
            for (int i = 0; i < fieldLabels.Count; i++)
            {
                double width = WordTemplateText.EstimateTextWidth(fieldLabels[i], options.FontSize) + 2;
                if (width > estimatedLabelWidth) estimatedLabelWidth = width;
            }
            double labelWidth = Math.Min(Math.Max(8, estimatedLabelWidth), contentWidth * 0.65);
            // 末尾留一列空白，填写横线在标签框右侧内缩，不会与边框粘在一起。
            // 列宽取 6 磅，同时满足 PDF 渲染器的最小列宽，Word 与 PDF 的框宽因此保持一致。
            // A trailing blank column insets the fill-in rules from the right box edge so they do not
            // merge with the border. Six points also satisfies the PDF renderer's minimum column width,
            // keeping the box width identical in Word and PDF.
            double rightInset = Math.Min(Math.Max(padding, 6), contentWidth * 0.2);
            double ruleWidth = Math.Max(1, contentWidth - labelWidth - rightInset);
            double lineHeight = Math.Max(8, options.FontSize * 1.4);
            double rowHeight = lineHeight + Math.Max(0, options.FieldSpacing);
            double borderWidth = options.BorderWidth > 0 ? options.BorderWidth : 0.5;

            // 所有字段共用一个嵌套表格：相邻的独立表格会在 Word 打开时被合并，单表加多行可保持结构稳定。
            // All fields share one nested table: adjacent standalone tables are merged by Word on open, while
            // a single table with multiple rows keeps the structure stable.
            Table lines = new Table();
            lines.Borders = options.ShowCellBorders
                ? CreateLabelBoxBorders(borderWidth, options.BorderColor)
                : TableBorders.None();
            lines.CellPadding = padding;
            lines.Autofit = false;
            lines.Alignment = TableAlignment.Left;
            lines.RepeatHeaderRow = false;
            // 列宽包含单元格内边距：文字列、填写线列、右侧留白列合计等于标签框宽度。
            // Column widths include cell padding: the text column, the rule column and the right inset
            // column add up to the label box width.
            lines.Columns.Add(new TableColumn(labelWidth + padding * 2));
            lines.Columns.Add(new TableColumn(ruleWidth));
            lines.Columns.Add(new TableColumn(rightInset));

            for (int i = 0; i < fieldLabels.Count; i++)
            {
                TableRow row = lines.AddRow();
                row.Height = rowHeight;

                TableCell labelCell = new TableCell();
                labelCell.VerticalAlignment = VerticalAlignment.Bottom;
                Paragraph labelParagraph = new Paragraph();
                ApplyLabelParagraph(labelParagraph, options);
                labelParagraph.AddText(fieldLabels[i], format);
                labelCell.Blocks.Add(labelParagraph);

                TableCell ruleCell = new TableCell();
                ruleCell.VerticalAlignment = VerticalAlignment.Bottom;
                ruleCell.Borders = new TableCellBorders();
                ruleCell.Borders.Bottom = new BorderLine(BorderStyle.Single, borderWidth, options.BorderColor);
                Paragraph ruleParagraph = new Paragraph();
                ApplyLabelParagraph(ruleParagraph, options);
                ruleCell.Blocks.Add(ruleParagraph);

                TableCell insetCell = new TableCell();
                Paragraph insetParagraph = new Paragraph();
                ApplyLabelParagraph(insetParagraph, options);
                insetCell.Blocks.Add(insetParagraph);

                row.Cells.Add(labelCell);
                row.Cells.Add(ruleCell);
                row.Cells.Add(insetCell);
            }

            // 表尾补一行空白，让标签框达到版式高度，最后一条填写线上方留出可书写空间。
            // A blank filler row completes the label box to the layout height and leaves writing space below
            // the last rule.
            double fillerHeight = layout.LabelHeight - fieldLabels.Count * rowHeight;
            if (fillerHeight > 0.01)
            {
                TableRow filler = lines.AddRow();
                filler.Height = fillerHeight;
                for (int column = 0; column < 3; column++)
                {
                    TableCell fillerCell = new TableCell();
                    Paragraph fillerParagraph = new Paragraph();
                    fillerParagraph.LineSpacing = 0.1;
                    fillerCell.Blocks.Add(fillerParagraph);
                    filler.Cells.Add(fillerCell);
                }
            }

            cell.Blocks.Add(lines);
            // Word 会在单元格末尾的嵌套表格后补一个空段落；这里显式补一个极小的段落，
            // 避免默认字号的空段把标签行撑高、挤乱整个分页。
            // Word appends an empty paragraph after a nested table that ends a cell; add a minimal one
            // explicitly so a default-sized empty paragraph cannot grow the label row and break pagination.
            Paragraph tail = new Paragraph();
            tail.LineSpacing = 0.1;
            tail.SpaceBefore = 0;
            tail.SpaceAfter = 0;
            cell.Blocks.Add(tail);
        }

        /// <summary>
        /// 把标签行写入一列多行嵌套表格：各行等高、文字顶对齐，行与行之间画满宽分隔线；
        /// 启用单元格边框时嵌套表格的四周边框构成标签外框。
        /// Writes the label lines into a one-column, multi-row nested table: equal row heights, top-aligned
        /// text and full-width separators between rows. When cell borders are enabled the nested table's
        /// four outer edges form the label box.
        /// </summary>
        /// <param name="cell">标签单元格。The label cell.</param>
        /// <param name="lines">标签文字行，可为 null。The label text lines; may be null.</param>
        /// <param name="rowCount">本页统一的标签行数。The page-wide label row count.</param>
        /// <param name="format">字符格式。The character format.</param>
        /// <param name="options">文字与边框选项。The text and border options.</param>
        /// <param name="layout">标签版式，提供标签尺寸。The label layout supplying the label dimensions.</param>
        private static void AppendSeparatedLines(
            TableCell cell,
            string[] lines,
            int rowCount,
            TextFormat format,
            WordLabelOptions options,
            WordLabelLayout layout)
        {
            int count = Math.Max(1, rowCount);
            double padding = Math.Max(0, layout.CellPadding);
            double borderWidth = options.BorderWidth > 0 ? options.BorderWidth : 0.5;
            double rowHeight = Math.Max(1, layout.LabelHeight / count);

            Table body = new Table();
            body.Borders = CreateRowSeparatorBorders(options.ShowCellBorders, borderWidth, options.BorderColor);
            body.CellPadding = padding;
            body.Autofit = false;
            body.Alignment = TableAlignment.Left;
            body.RepeatHeaderRow = false;
            // 列宽等于标签宽度：嵌套表格铺满单元格，外框与标签框重合，分隔线横贯整个标签。
            // The column spans the label width so the nested table fills the cell, its outer edges coincide
            // with the label box and the separators run across the whole label.
            body.Columns.Add(new TableColumn(Math.Max(8, layout.LabelWidth)));

            for (int i = 0; i < count; i++)
            {
                TableRow row = body.AddRow();
                row.Height = rowHeight;
                TableCell lineCell = new TableCell();
                lineCell.VerticalAlignment = VerticalAlignment.Top;
                Paragraph paragraph = new Paragraph();
                ApplyLabelParagraph(paragraph, options);
                string text = lines != null && i < lines.Length ? lines[i] : null;
                paragraph.AddText(text ?? string.Empty, format);
                lineCell.Blocks.Add(paragraph);
                row.Cells.Add(lineCell);
            }

            cell.Blocks.Add(body);
            // 嵌套表格后补一个极小的段落，避免 Word 补入的默认空段把标签行撑高。
            // Add a minimal paragraph after the nested table so Word's own trailing empty paragraph cannot
            // inflate the label row.
            Paragraph tail = new Paragraph();
            tail.LineSpacing = 0.1;
            tail.SpaceBefore = 0;
            tail.SpaceAfter = 0;
            cell.Blocks.Add(tail);
        }

        /// <summary>
        /// 创建带行分隔线的标签边框：内部横线始终绘制，外框由 <paramref name="showBox"/> 决定，内部竖线不绘制。
        /// Builds the label borders with row separators: internal horizontal lines are always drawn, the outer
        /// box follows <paramref name="showBox"/>, and internal vertical lines are omitted.
        /// </summary>
        /// <param name="showBox">是否绘制四周外框。Whether to draw the four outer edges.</param>
        /// <param name="width">边框宽度（磅）。The border width in points.</param>
        /// <param name="color">边框颜色。The border color.</param>
        /// <returns>带行分隔线的边框集合。The border set with row separators.</returns>
        private static TableBorders CreateRowSeparatorBorders(bool showBox, double width, OfficeColor color)
        {
            BorderLine line = new BorderLine(BorderStyle.Single, width, color);
            BorderLine none = new BorderLine();
            TableBorders borders = new TableBorders();
            borders.Top = showBox ? line : none;
            borders.Bottom = showBox ? line : none;
            borders.Left = showBox ? line : none;
            borders.Right = showBox ? line : none;
            borders.InsideHorizontal = line;
            borders.InsideVertical = none;
            return borders;
        }

        /// <summary>
        /// 创建标签框边框：只画表格四周外框，不画内部横线与竖线，使填写横线成为框内唯一的分隔线。
        /// Builds the label box border: only the four outer edges are drawn and internal horizontal and vertical
        /// lines are omitted, leaving the fill-in rules as the only dividers inside the box.
        /// </summary>
        /// <param name="width">边框宽度（磅）。The border width in points.</param>
        /// <param name="color">边框颜色。The border color.</param>
        /// <returns>四周外框的边框集合。The border set with the four outer edges.</returns>
        private static TableBorders CreateLabelBoxBorders(double width, OfficeColor color)
        {
            TableBorders borders = new TableBorders();
            borders.Top = new BorderLine(BorderStyle.Single, width, color);
            borders.Bottom = new BorderLine(BorderStyle.Single, width, color);
            borders.Left = new BorderLine(BorderStyle.Single, width, color);
            borders.Right = new BorderLine(BorderStyle.Single, width, color);
            borders.InsideHorizontal = new BorderLine();
            borders.InsideVertical = new BorderLine();
            return borders;
        }

        /// <summary>
        /// 应用表单标签的段落格式：无段前段后距、单倍行距，对齐方式沿用选项。
        /// Applies the form-label paragraph format: no space before or after, single spacing, option alignment.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="options">文字选项。The text options.</param>
        private static void ApplyLabelParagraph(Paragraph paragraph, WordLabelOptions options)
        {
            paragraph.Alignment = options.Alignment;
            paragraph.LineSpacing = 1.0;
            paragraph.SpaceBefore = 0;
            paragraph.SpaceAfter = 0;
        }

        /// <summary>
        /// 向标签单元格写入多行文本：首行留出上内边距，所有行统一左内边距与行距。
        /// Writes the multi-line text into a label cell: the first line gets the top inset, and every line
        /// shares the left inset and line spacing.
        /// </summary>
        /// <param name="cell">目标单元格。The target cell.</param>
        /// <param name="lines">文本行，可为 null。The text lines; may be null.</param>
        /// <param name="format">字符格式。The character format.</param>
        /// <param name="options">文字选项。The text options.</param>
        /// <param name="padding">内容内边距。The content inset.</param>
        private static void AppendLines(
            TableCell cell,
            string[] lines,
            TextFormat format,
            WordLabelOptions options,
            double padding)
        {
            if (lines == null || lines.Length == 0)
            {
                cell.Blocks.Add(new Paragraph());
                return;
            }
            for (int i = 0; i < lines.Length; i++)
            {
                Paragraph paragraph = new Paragraph();
                paragraph.Alignment = options.Alignment;
                paragraph.LineSpacing = options.LineSpacing;
                paragraph.LeftIndent = padding;
                paragraph.SpaceBefore = i == 0 ? padding : 0;
                paragraph.SpaceAfter = 0;
                paragraph.AddText(lines[i] ?? string.Empty, format);
                cell.Blocks.Add(paragraph);
            }
        }
    }

    /// <summary>
    /// 标签与信封共用的模板文本工具：按换行拆行、由记录构建上下文并替换 {{字段}} 标记。
    /// Template text helpers shared by labels and envelopes: split on line breaks, build a record context
    /// and replace {{field}} tokens.
    /// </summary>
    internal static class WordTemplateText
    {
        /// <summary>把文本按 CRLF/LF/CR 拆成行；null 或空串返回单个空行。Splits text on CRLF/LF/CR; null or empty yields one empty line.</summary>
        /// <param name="text">待拆分的文本。The text to split.</param>
        /// <returns>文本行数组。The text lines.</returns>
        public static string[] SplitLines(string text)
        {
            if (string.IsNullOrEmpty(text)) return new string[] { string.Empty };
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            return normalized.Split('\n');
        }

        /// <summary>
        /// 粗略估算文本在指定字号下的显示宽度（磅）：全角字符按 1 个字宽、空白按 0.3 个字宽、
        /// 其余半角字符按 0.55 个字宽计算，用于给表单标签预留文字列宽。
        /// Roughly estimates the display width of text at a given font size in points: full-width characters
        /// take one em, whitespace 0.3 em and other half-width characters 0.55 em, which sizes the label
        /// column of form labels.
        /// </summary>
        /// <param name="text">待估算的文本，可为 null。The text to measure; may be null.</param>
        /// <param name="fontSize">字号（磅）。The font size in points.</param>
        /// <returns>估算宽度（磅）。The estimated width in points.</returns>
        public static double EstimateTextWidth(string text, double fontSize)
        {
            if (string.IsNullOrEmpty(text)) return 0;
            if (fontSize <= 0) fontSize = 10.5;
            double em = fontSize;
            double width = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (char.IsWhiteSpace(ch)) width += em * 0.3;
                else if (IsFullWidth(ch)) width += em;
                else width += em * 0.55;
            }
            return width;
        }

        /// <summary>
        /// 判断字符是否按全角字宽排版：覆盖中日韩文字、假名、谚文、全角标点与兼容表意文字。
        /// Determines whether a character is laid out at full width: CJK ideographs, kana, Hangul,
        /// full-width punctuation and compatibility ideographs.
        /// </summary>
        /// <param name="ch">待判断的字符。The character to test.</param>
        /// <returns>全角字符返回 true。True for full-width characters.</returns>
        private static bool IsFullWidth(char ch)
        {
            return (ch >= '\u1100' && ch <= '\u115F')
                || (ch >= '\u2E80' && ch <= '\uA4CF')
                || (ch >= '\uAC00' && ch <= '\uD7A3')
                || (ch >= '\uF900' && ch <= '\uFAFF')
                || (ch >= '\uFE30' && ch <= '\uFE4F')
                || (ch >= '\uFF00' && ch <= '\uFF60')
                || (ch >= '\uFFE0' && ch <= '\uFFE6');
        }

        /// <summary>
        /// 逐行替换模板标记。
        /// Replaces template tokens line by line.
        /// </summary>
        /// <param name="template">多行模板。The multi-line template.</param>
        /// <param name="context">数据上下文。The data context.</param>
        /// <param name="options">模板选项。The template options.</param>
        /// <returns>替换后的文本行。The substituted text lines.</returns>
        public static string[] ReplaceLines(string template, TemplateContext context, TemplateOptions options)
        {
            string[] lines = SplitLines(template);
            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = TemplateEngine.Replace(lines[i], context, options);
            }
            return lines;
        }

        /// <summary>
        /// 由记录对象构建模板上下文：TemplateContext 直接使用，DataRow/DataRowView 按列名展开，其他对象走反射。
        /// Builds a template context from a record: TemplateContext is used as-is, DataRow/DataRowView expand by
        /// column name, and other objects are reflected.
        /// </summary>
        /// <param name="record">记录对象，可为 null。The record object; may be null.</param>
        /// <returns>模板上下文。The template context.</returns>
        public static TemplateContext CreateContext(object record)
        {
            TemplateContext context = record as TemplateContext;
            if (context != null) return context;

            DataRowView view = record as DataRowView;
            if (view != null) record = view.Row;

            DataRow row = record as DataRow;
            if (row != null && row.Table != null)
            {
                TemplateContext rowContext = new TemplateContext();
                DataColumnCollection columns = row.Table.Columns;
                for (int i = 0; i < columns.Count; i++)
                {
                    rowContext.Set(columns[i].ColumnName, row[columns[i]]);
                }
                return rowContext;
            }
            return TemplateEngine.CreateRowContext(record);
        }
    }
}
