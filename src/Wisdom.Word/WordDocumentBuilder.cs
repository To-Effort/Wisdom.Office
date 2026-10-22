using System;
using System.Collections.Generic;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 流式文档构建器，风格参考 Aspose.Words 的 DocumentBuilder：维护一个写入游标与当前字符/段落格式，
    /// 通过 Write/InsertXxx 顺序追加段落、文本、表格、图片、书签、超链接、域与目录。
    /// Fluent document builder modelled on the Aspose.Words DocumentBuilder: it keeps a write cursor plus the
    /// current character and paragraph formatting, and appends paragraphs, text, tables, images, bookmarks,
    /// hyperlinks, fields and tables of contents through Write/InsertXxx calls.
    /// </summary>
    /// <remarks>
    /// 构建器只操作统一文档模型，不直接接触 DOCX 包；完成后调用 <see cref="DocxWriter"/> 写出。
    /// The builder only manipulates the unified document model and never touches the DOCX package directly;
    /// write the result with <see cref="DocxWriter"/>.
    /// </remarks>
    public sealed class WordDocumentBuilder
    {
        /// <summary>正在构建的文档。The document under construction.</summary>
        private readonly OfficeDocument _document;
        /// <summary>当前节的索引。The index of the current section.</summary>
        private int _sectionIndex;
        /// <summary>当前节内的插入位置（块列表下标）。The insertion position within the current section's block list.</summary>
        private int _blockIndex;
        /// <summary>接收 Write 文字的当前段落；null 表示需要新建。The paragraph receiving Write text; null means one must be created.</summary>
        private Paragraph _paragraph;
        /// <summary>行内插入游标（当前段落 Inlines 下标）；-1 表示追加到末尾。The inline insertion cursor (an index into the current paragraph's inlines); -1 appends at the end.</summary>
        private int _inlineIndex = -1;
        /// <summary>游标所属段落；段落一变游标即失效。The paragraph owning the cursor; changing paragraph invalidates it.</summary>
        private Paragraph _cursorParagraph;
        /// <summary>正在构建的表格；null 表示光标在正文。The table under construction; null means the cursor is in the body.</summary>
        private Table _table;
        /// <summary>正在构建的表格行。The table row under construction.</summary>
        private TableRow _row;
        /// <summary>正在写入的单元格。The cell being written to.</summary>
        private TableCell _cell;

        /// <summary>当前段落对齐方式。The current paragraph alignment.</summary>
        private ParagraphAlignment _alignment;
        /// <summary>当前段前间距（磅）。The current space-before value in points.</summary>
        private double _spaceBefore;
        /// <summary>当前段后间距（磅）。The current space-after value in points.</summary>
        private double _spaceAfter;
        /// <summary>当前行距倍数。The current line-spacing multiple.</summary>
        private double _lineSpacing = 1.0;
        /// <summary>当前左缩进（磅）。The current left indent in points.</summary>
        private double _leftIndent;
        /// <summary>当前右缩进（磅）。The current right indent in points.</summary>
        private double _rightIndent;
        /// <summary>当前首行缩进（磅）。The current first-line indent in points.</summary>
        private double _firstLineIndent;
        /// <summary>当前标题级别，0 表示正文。The current heading level; 0 means body text.</summary>
        private int _headingLevel;
        /// <summary>当前段落是否与下段同页。Whether the current paragraph keeps with the next one.</summary>
        private bool _keepWithNext;
        /// <summary>当前段落是否段前分页。Whether the current paragraph starts on a new page.</summary>
        private bool _pageBreakBefore;

        /// <summary>
        /// 在已有文档末尾创建构建器。
        /// Creates a builder positioned at the end of an existing document.
        /// </summary>
        /// <param name="document">目标文档，不能为 null。The target document; must not be null.</param>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 为 null。Thrown when <paramref name="document"/> is null.</exception>
        public WordDocumentBuilder(OfficeDocument document)
        {
            if (document == null) throw new ArgumentNullException("document");
            _document = document;
            Section section = document.FirstSection;
            _sectionIndex = 0;
            _blockIndex = section.Blocks.Count;
        }

        /// <summary>创建一个空文档并绑定构建器。Creates an empty document and binds the builder to it.</summary>
        public WordDocumentBuilder()
            : this(new OfficeDocument())
        {
        }

        /// <summary>正在构建的文档。The document under construction.</summary>
        public OfficeDocument Document
        {
            get { return _document; }
        }

        /// <summary>当前字符格式；Write/Insert 时按当前值复制到每个运行。The current character format; copied to every run written by Write/Insert calls.</summary>
        public TextFormat Font { get; private set; } = new TextFormat();

        /// <summary>当前段落对齐方式，赋值同时应用到已打开的段落。The current paragraph alignment; setting it also updates the open paragraph.</summary>
        public ParagraphAlignment Alignment
        {
            get { return _alignment; }
            set { _alignment = value; if (_paragraph != null) _paragraph.Alignment = value; }
        }

        /// <summary>当前段前间距（磅）。The current space before the paragraph, in points.</summary>
        public double SpaceBefore
        {
            get { return _spaceBefore; }
            set { _spaceBefore = value; if (_paragraph != null) _paragraph.SpaceBefore = value; }
        }

        /// <summary>当前段后间距（磅）。The current space after the paragraph, in points.</summary>
        public double SpaceAfter
        {
            get { return _spaceAfter; }
            set { _spaceAfter = value; if (_paragraph != null) _paragraph.SpaceAfter = value; }
        }

        /// <summary>当前行距倍数（1.0 为单倍）。The current line-spacing multiple (1.0 is single).</summary>
        public double LineSpacing
        {
            get { return _lineSpacing; }
            set { _lineSpacing = value; if (_paragraph != null) _paragraph.LineSpacing = value; }
        }

        /// <summary>当前左缩进（磅）。The current left indent, in points.</summary>
        public double LeftIndent
        {
            get { return _leftIndent; }
            set { _leftIndent = value; if (_paragraph != null) _paragraph.LeftIndent = value; }
        }

        /// <summary>当前右缩进（磅）。The current right indent, in points.</summary>
        public double RightIndent
        {
            get { return _rightIndent; }
            set { _rightIndent = value; if (_paragraph != null) _paragraph.RightIndent = value; }
        }

        /// <summary>当前首行缩进（磅），负值表示悬挂缩进。The current first-line indent in points; a negative value means hanging.</summary>
        public double FirstLineIndent
        {
            get { return _firstLineIndent; }
            set { _firstLineIndent = value; if (_paragraph != null) _paragraph.FirstLineIndent = value; }
        }

        /// <summary>当前标题级别（0 正文，1~9 标题）。The current heading level (0 body, 1-9 headings).</summary>
        public int HeadingLevel
        {
            get { return _headingLevel; }
            set
            {
                if (value < 0 || value > 9) throw new ArgumentOutOfRangeException("value");
                _headingLevel = value;
                if (_paragraph != null) _paragraph.HeadingLevel = value;
            }
        }

        /// <summary>当前段落是否与下段同页。Whether the current paragraph keeps with the next one.</summary>
        public bool KeepWithNext
        {
            get { return _keepWithNext; }
            set { _keepWithNext = value; if (_paragraph != null) _paragraph.KeepWithNext = value; }
        }

        /// <summary>当前段落是否强制段前分页。Whether the current paragraph starts on a new page.</summary>
        public bool PageBreakBefore
        {
            get { return _pageBreakBefore; }
            set { _pageBreakBefore = value; if (_paragraph != null) _paragraph.PageBreakBefore = value; }
        }

        /// <summary>当前节。The current section.</summary>
        public Section CurrentSection
        {
            get { return _document.Sections[_sectionIndex]; }
        }

        /// <summary>
        /// 把光标移到文档开头（第一节第一个块之前）。
        /// Moves the cursor to the very beginning of the document (before the first block of the first section).
        /// </summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder MoveToDocumentStart()
        {
            if (_document.Sections.Count == 0) _document.AddSection();
            _sectionIndex = 0;
            _blockIndex = 0;
            ResetParagraph();
            return this;
        }

        /// <summary>
        /// 把光标移到文档末尾（最后一节最后一个块之后）。
        /// Moves the cursor to the end of the document (after the last block of the last section).
        /// </summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder MoveToDocumentEnd()
        {
            _sectionIndex = Math.Max(0, _document.Sections.Count - 1);
            _blockIndex = CurrentSection.Blocks.Count;
            ResetParagraph();
            return this;
        }

        /// <summary>
        /// 把光标移到指定节的末尾。
        /// Moves the cursor to the end of the given section.
        /// </summary>
        /// <param name="sectionIndex">节下标。The zero-based section index.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="ArgumentOutOfRangeException">节下标越界。Thrown when the section index is out of range.</exception>
        public WordDocumentBuilder MoveToSection(int sectionIndex)
        {
            if (sectionIndex < 0 || sectionIndex >= _document.Sections.Count) throw new ArgumentOutOfRangeException("sectionIndex");
            _sectionIndex = sectionIndex;
            _blockIndex = CurrentSection.Blocks.Count;
            ResetParagraph();
            return this;
        }

        /// <summary>
        /// 把光标移到当前节指定块之前。
        /// Moves the cursor before the given block of the current section.
        /// </summary>
        /// <param name="blockIndex">块下标；等于块数时表示末尾。The zero-based block index; the block count means the end.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="ArgumentOutOfRangeException">块下标越界。Thrown when the block index is out of range.</exception>
        public WordDocumentBuilder MoveToParagraph(int blockIndex)
        {
            if (blockIndex < 0 || blockIndex > CurrentSection.Blocks.Count) throw new ArgumentOutOfRangeException("blockIndex");
            _blockIndex = blockIndex;
            ResetParagraph();
            return this;
        }

        /// <summary>
        /// 追加一个新节并把光标移到该节开头；已有内容会自然形成分节符。
        /// Appends a new section and moves the cursor to its beginning; existing content naturally forms a
        /// section break.
        /// </summary>
        /// <param name="setup">页面设置；null 时使用 A4 默认值。The page setup; A4 defaults are used when null.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder AddSection(PageSetup setup)
        {
            _document.AddSection(setup == null ? new PageSetup() : setup);
            _sectionIndex = _document.Sections.Count - 1;
            _blockIndex = 0;
            ResetParagraph();
            return this;
        }

        /// <summary>
        /// 写入文字，使用当前字符格式的副本；文字中的 \n、\t 由写出器转换为换行与制表符。
        /// Writes text using a copy of the current character format; \n and \t are converted to line breaks and
        /// tabs by the writer.
        /// </summary>
        /// <param name="text">要写入的文字；null 或空串不产生任何内容。The text to write; null or empty produces nothing.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder Write(string text)
        {
            if (string.IsNullOrEmpty(text)) return this;
            InsertInlineAtCursor(new TextRun(text, Font.Clone()));
            return this;
        }

        /// <summary>
        /// 写入文字并结束当前段落，后续 Write 会落到新段落。
        /// Writes text and ends the current paragraph, so subsequent Write calls start a new paragraph.
        /// </summary>
        /// <param name="text">要写入的文字。The text to write.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder WriteLine(string text)
        {
            Write(text);
            return InsertParagraph();
        }

        /// <summary>
        /// 插入一个新段落并把它设为当前段落；可同时写入文字。
        /// Inserts a new paragraph and makes it the current one, optionally writing text into it.
        /// </summary>
        /// <param name="text">新段落的文字，null 表示空段落。The new paragraph's text; null creates an empty paragraph.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder InsertParagraph(string text)
        {
            Paragraph paragraph = new Paragraph();
            ApplyParagraphState(paragraph);
            if (!string.IsNullOrEmpty(text)) paragraph.AddText(text, Font.Clone());
            InsertBlockCore(paragraph);
            _paragraph = paragraph;
            return this;
        }

        /// <summary>
        /// 插入一个空段落并把它设为当前段落。
        /// Inserts an empty paragraph and makes it the current one.
        /// </summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder InsertParagraph()
        {
            return InsertParagraph(null);
        }

        /// <summary>
        /// 插入分页符块。
        /// Inserts a page-break block.
        /// </summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder InsertPageBreak()
        {
            InsertBlockCore(new PageBreak());
            _paragraph = null;
            return this;
        }

        /// <summary>
        /// 在当前段落插入内联图片；宽高为 0 时使用图片按 96 DPI 换算的默认尺寸。
        /// Inserts an inline image into the current paragraph; zero width or height uses the image's default size
        /// converted at 96 DPI.
        /// </summary>
        /// <param name="image">图片数据，不能为 null。The image data; must not be null.</param>
        /// <param name="width">显示宽度（磅），0 表示默认。The display width in points; 0 means the default.</param>
        /// <param name="height">显示高度（磅），0 表示默认。The display height in points; 0 means the default.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="image"/> 为 null。Thrown when <paramref name="image"/> is null.</exception>
        public WordDocumentBuilder InsertImage(OfficeImage image, double width, double height)
        {
            if (image == null) throw new ArgumentNullException("image");
            InsertInlineAtCursor(new ImageRun(image, width, height));
            return this;
        }

        /// <summary>
        /// 在当前段落插入书签。书签覆盖整个段落（与统一模型的书签模型一致）。
        /// Inserts a bookmark on the current paragraph. The bookmark covers the whole paragraph, matching the
        /// bookmark model of the unified document.
        /// </summary>
        /// <param name="name">书签名：以字母或下划线开头，只含字母、数字与下划线，最长 40 字符。The bookmark name: starts with a letter or underscore, contains only letters, digits and underscores, and is at most 40 characters long.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="ArgumentException">书签名非法。Thrown when the bookmark name is invalid.</exception>
        public WordDocumentBuilder InsertBookmark(string name)
        {
            string value = ValidateBookmarkName(name);
            EnsureParagraph().BookmarkName = value;
            return this;
        }

        /// <summary>
        /// 在光标处插入书签起点，与 <see cref="EndBookmark"/> 配对即可覆盖段落中间的任意范围、跨段落甚至跨表格。
        /// Inserts a bookmark start at the cursor; paired with <see cref="EndBookmark"/> it can cover any range inside
        /// a paragraph, across paragraphs, or across tables.
        /// </summary>
        /// <param name="name">书签名：以字母或下划线开头，只含字母、数字与下划线，最长 40 字符。The bookmark name: starts with a letter or underscore, contains only letters, digits and underscores, and is at most 40 characters long.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="ArgumentException">书签名非法。Thrown when the bookmark name is invalid.</exception>
        public WordDocumentBuilder StartBookmark(string name)
        {
            InsertInlineAtCursor(new BookmarkStart(ValidateBookmarkName(name)));
            return this;
        }

        /// <summary>
        /// 在光标处插入书签终点，闭合由 <see cref="StartBookmark"/> 打开的范围。
        /// Inserts a bookmark end at the cursor, closing the range opened by <see cref="StartBookmark"/>.
        /// </summary>
        /// <param name="name">与起点同名的书签名。The bookmark name matching the start.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="ArgumentException">书签名非法。Thrown when the bookmark name is invalid.</exception>
        public WordDocumentBuilder EndBookmark(string name)
        {
            InsertInlineAtCursor(new BookmarkEnd(ValidateBookmarkName(name)));
            return this;
        }

        /// <summary>
        /// 把光标移动到书签端点旁边，后续 Write/Insert 就地落入该段落；书签不存在时保持原位。
        /// Moves the cursor next to a bookmark endpoint so later Write/Insert calls land inside that paragraph;
        /// the cursor stays put when the bookmark does not exist.
        /// </summary>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <param name="isStart">true 定位到起点，false 定位到终点。True targets the start, false the end.</param>
        /// <param name="isAfter">true 落在端点之后，false 落在之前。True lands after the endpoint, false before it.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder MoveToBookmark(string name, bool isStart, bool isAfter)
        {
            string value = name == null ? string.Empty : name.Trim();
            Bookmark bookmark = value.Length == 0 ? null : _document.Bookmarks[value];
            if (bookmark == null) return this;
            if (bookmark.IsWholeParagraph)
            {
                MoveCursorTo(bookmark.Paragraph, null, true);
                return this;
            }
            Inline marker = isStart ? bookmark.Start : bookmark.End;
            if (marker == null) marker = isStart ? bookmark.End : bookmark.Start;
            Paragraph owner = _document.Bookmarks.ParagraphOf(marker);
            if (owner == null) owner = bookmark.Paragraph;
            MoveCursorTo(owner, marker, isAfter);
            return this;
        }

        /// <summary>
        /// 把光标移动到书签起点之前。Convenience overload that moves the cursor before a bookmark's start.
        /// </summary>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder MoveToBookmark(string name)
        {
            return MoveToBookmark(name, true, false);
        }

        /// <summary>
        /// 插入超链接文字。外部地址使用蓝色下划线样式（当前格式为默认黑色时），以 "#" 开头的目标视为表内锚点。
        /// Inserts hyperlink text. External targets get the blue underline look when the current format is plain
        /// black; targets starting with "#" are treated as in-document anchors.
        /// </summary>
        /// <param name="text">链接显示文字。The link display text.</param>
        /// <param name="target">外部 URL 或以 "#" 开头的书签/锚点。The external URL or an in-document anchor starting with "#".</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="ArgumentException"><paramref name="target"/> 为空。Thrown when <paramref name="target"/> is empty.</exception>
        public WordDocumentBuilder InsertHyperlink(string text, string target)
        {
            if (string.IsNullOrEmpty(target)) throw new ArgumentException("Hyperlink target cannot be empty.", "target");
            string value = target.Trim();
            if (value.Length == 0) throw new ArgumentException("Hyperlink target cannot be empty.", "target");

            TextFormat format = Font.Clone();
            if (!format.Underline) format.Underline = true;
            if (string.Equals(format.Color.ToArgbHex(), OfficeColor.Black.ToArgbHex(), StringComparison.Ordinal))
            {
                format.Color = new OfficeColor(0x05, 0x63, 0xC1);
            }
            TextRun run = new TextRun(string.IsNullOrEmpty(text) ? value : text, format);
            run.Link = value;
            InsertInlineAtCursor(run);
            return this;
        }

        /// <summary>
        /// 插入 Word 域（如 PAGE、NUMPAGES、MERGEFIELD），写出为打开时自动刷新的 fldSimple。
        /// Inserts a Word field (such as PAGE, NUMPAGES or MERGEFIELD), written as an fldSimple that refreshes on open.
        /// </summary>
        /// <param name="instruction">域指令。The field instruction.</param>
        /// <param name="displayText">域结果的占位显示文字。The placeholder display text for the field result.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder InsertField(string instruction, string displayText)
        {
            InsertInlineAtCursor(new FieldRun(instruction, displayText, Font.Clone()));
            return this;
        }

        /// <summary>
        /// 插入目录域（TOC）。标题需使用 <see cref="HeadingLevel"/> 标记，Word 打开时会按标题级别生成目录。
        /// Inserts a table-of-contents field (TOC). Headings must be marked with <see cref="HeadingLevel"/> so Word
        /// can build the contents from them when the document is opened.
        /// </summary>
        /// <param name="levels">收录的标题级别数（1~9）。The number of heading levels to include (1-9).</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder InsertTableOfContents(int levels)
        {
            if (levels < 1 || levels > 9) throw new ArgumentOutOfRangeException("levels");
            Paragraph paragraph = new Paragraph();
            ApplyParagraphState(paragraph);
            paragraph.Inlines.Add(new FieldRun(
                "TOC \\o \"1-" + levels.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\" \\h \\z \\u",
                "右键更新目录",
                Font.Clone()));
            InsertBlockCore(paragraph);
            _paragraph = paragraph;
            return this;
        }

        /// <summary>插入三级目录域。Inserts a three-level table of contents.</summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder InsertTableOfContents()
        {
            return InsertTableOfContents(3);
        }

        /// <summary>
        /// 插入一个已构建好的表格块。
        /// Inserts a pre-built table block.
        /// </summary>
        /// <param name="table">表格模型。The table model.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="table"/> 为 null。Thrown when <paramref name="table"/> is null.</exception>
        public WordDocumentBuilder InsertTable(Table table)
        {
            if (table == null) throw new ArgumentNullException("table");
            InsertBlockCore(table);
            _paragraph = null;
            return this;
        }

        /// <summary>
        /// 开始构建表格：后续 InsertCell/EndRow 在此表格内操作。
        /// Starts building a table; subsequent InsertCell/EndRow calls operate inside it.
        /// </summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder StartTable()
        {
            Table table = new Table();
            InsertBlockCore(table);
            _table = table;
            _row = null;
            _cell = null;
            _paragraph = null;
            return this;
        }

        /// <summary>
        /// 在当前行追加一个单元格并开始写入；没有行时自动创建一行。
        /// Appends a cell to the current row and starts writing into it; a row is created automatically when
        /// none is open.
        /// </summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="InvalidOperationException">未调用 <see cref="StartTable"/>。Thrown when no table is being built.</exception>
        public WordDocumentBuilder InsertCell()
        {
            if (_table == null) throw new InvalidOperationException("InsertCell requires StartTable first.");
            if (_row == null) _row = _table.AddRow();
            _cell = new TableCell();
            _row.Cells.Add(_cell);
            _paragraph = new Paragraph();
            ApplyParagraphState(_paragraph);
            _cell.Blocks.Add(_paragraph);
            return this;
        }

        /// <summary>
        /// 在当前行追加一个单元格并写入文字。
        /// Appends a cell to the current row and writes text into it.
        /// </summary>
        /// <param name="text">单元格文字。The cell text.</param>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder InsertCell(string text)
        {
            InsertCell();
            return Write(text);
        }

        /// <summary>
        /// 结束当前行；下一次 InsertCell 会创建新行。
        /// Ends the current row; the next InsertCell call creates a new row.
        /// </summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        public WordDocumentBuilder EndRow()
        {
            _row = null;
            _cell = null;
            _paragraph = null;
            return this;
        }

        /// <summary>
        /// 结束当前表格，光标回到正文末尾。
        /// Ends the current table and returns the cursor to the body.
        /// </summary>
        /// <returns>当前构建器，便于链式调用。The current builder, for fluent chaining.</returns>
        /// <exception cref="InvalidOperationException">未调用 <see cref="StartTable"/>。Thrown when no table is being built.</exception>
        public WordDocumentBuilder EndTable()
        {
            if (_table == null) throw new InvalidOperationException("EndTable requires StartTable first.");
            _table = null;
            _row = null;
            _cell = null;
            _paragraph = null;
            return this;
        }

        /// <summary>
        /// 确保当前写入位置存在一个段落：正文中新建块，表格中新建单元格段落。
        /// Ensures a paragraph exists at the current write position, creating a body block or a cell paragraph
        /// as needed.
        /// </summary>
        /// <returns>接收后续文字/行内元素的段落。The paragraph that receives subsequent text and inlines.</returns>
        private Paragraph EnsureParagraph()
        {
            if (_table != null)
            {
                if (_cell == null) InsertCell();
                if (_paragraph == null)
                {
                    _paragraph = new Paragraph();
                    ApplyParagraphState(_paragraph);
                    _cell.Blocks.Add(_paragraph);
                }
                return _paragraph;
            }

            if (_paragraph == null)
            {
                Paragraph paragraph = new Paragraph();
                ApplyParagraphState(paragraph);
                InsertBlockCore(paragraph);
                _paragraph = paragraph;
            }
            return _paragraph;
        }

        /// <summary>
        /// 把块插入当前节的光标位置并前移光标；表格单元格内的段落直接追加到单元格。
        /// Inserts a block at the cursor in the current section and advances the cursor; paragraphs inside a
        /// table cell are appended directly to the cell.
        /// </summary>
        /// <param name="block">待插入的块。The block to insert.</param>
        private void InsertBlockCore(Block block)
        {
            if (_table != null && _cell != null)
            {
                _cell.Blocks.Add(block);
                return;
            }
            Section section = CurrentSection;
            if (_blockIndex < 0) _blockIndex = 0;
            if (_blockIndex > section.Blocks.Count) _blockIndex = section.Blocks.Count;
            section.Blocks.Insert(_blockIndex, block);
            _blockIndex++;
        }

        /// <summary>
        /// 把构建器的段落格式状态应用到新段落。
        /// Applies the builder's paragraph-format state to a new paragraph.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        private void ApplyParagraphState(Paragraph paragraph)
        {
            paragraph.Alignment = _alignment;
            paragraph.SpaceBefore = _spaceBefore;
            paragraph.SpaceAfter = _spaceAfter;
            paragraph.LineSpacing = _lineSpacing;
            paragraph.LeftIndent = _leftIndent;
            paragraph.RightIndent = _rightIndent;
            paragraph.FirstLineIndent = _firstLineIndent;
            paragraph.HeadingLevel = _headingLevel;
            paragraph.KeepWithNext = _keepWithNext;
            paragraph.PageBreakBefore = _pageBreakBefore;
        }

        /// <summary>
        /// 清空当前段落引用（不影响已写入的块）。
        /// Clears the current paragraph reference without affecting blocks already written.
        /// </summary>
        private void ResetParagraph()
        {
            _paragraph = null;
            _inlineIndex = -1;
            _table = null;
            _row = null;
            _cell = null;
        }

        /// <summary>
        /// 校验书签名：非空、最长 40 字符、首字符为字母或下划线、其余为字母/数字/下划线。
        /// Validates a bookmark name: non-empty, at most 40 characters, starting with a letter or underscore and
        /// containing only letters, digits and underscores.
        /// </summary>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <returns>去除首尾空白后的书签名。The trimmed bookmark name.</returns>
        /// <exception cref="ArgumentException">书签名非法。Thrown when the name is invalid.</exception>
        private static string ValidateBookmarkName(string name)
        {
            return BookmarkNames.Validate(name);
        }

        /// <summary>
        /// 在光标处插入行内元素：游标未设置时追加到段落末尾，已设置时插到游标处并前移游标。
        /// Inserts an inline at the cursor: appended to the end of the paragraph when no cursor is set, otherwise
        /// placed at the cursor, which then advances.
        /// </summary>
        /// <param name="target">待插入行内元素。The inline to insert.</param>
        private void InsertInlineAtCursor(Inline target)
        {
            Paragraph paragraph = EnsureParagraph();
            if (!ReferenceEquals(paragraph, _cursorParagraph)) _inlineIndex = -1;
            _cursorParagraph = paragraph;
            if (_inlineIndex < 0)
            {
                paragraph.Inlines.Add(target);
                return;
            }
            if (_inlineIndex > paragraph.Inlines.Count) _inlineIndex = paragraph.Inlines.Count;
            paragraph.Inlines.Insert(_inlineIndex, target);
            _inlineIndex++;
        }

        /// <summary>
        /// 把写入光标定位到段落内的行内元素旁边。
        /// Moves the write cursor next to an inline element inside its paragraph.
        /// </summary>
        /// <param name="paragraph">目标段落。The paragraph holding the inline.</param>
        /// <param name="target">行内元素。The inline element.</param>
        /// <param name="after">true 插入该元素之后，false 插入之前。True inserts after the element, false before it.</param>
        private void MoveCursorTo(Paragraph paragraph, Inline target, bool after)
        {
            if (paragraph == null) return;
            int index = paragraph.Inlines.IndexOf(target);
            _paragraph = paragraph;
            _cursorParagraph = paragraph;
            _inlineIndex = index < 0 ? -1 : (after ? index + 1 : index);
            _table = null;
            _row = null;
            _cell = null;
            for (int s = 0; s < _document.Sections.Count; s++)
            {
                List<Block> blocks = _document.Sections[s].Blocks;
                for (int b = 0; b < blocks.Count; b++)
                {
                    if (!ReferenceEquals(blocks[b], paragraph)) continue;
                    // 新建段落落在书签段落之后，与就地续写的直觉一致。New paragraphs land after the bookmark paragraph.
                    _sectionIndex = s;
                    _blockIndex = b + 1;
                    return;
                }
            }
        }
    }
}
