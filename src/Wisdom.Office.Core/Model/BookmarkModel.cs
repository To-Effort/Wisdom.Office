using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// 书签范围起点行内标记，成对配合 <see cref="BookmarkEnd"/> 使用，可落在段落任意位置、跨段落甚至跨表格。
    /// The inline marker that opens a bookmark range, paired with <see cref="BookmarkEnd"/>; it can sit anywhere
    /// inside a paragraph, including across paragraphs and across tables.
    /// </summary>
    public sealed class BookmarkStart : Inline
    {
        /// <summary>书签名。The bookmark name.</summary>
        public string Name { get; set; }

        /// <summary>以指定名称创建起点标记。Creates a start marker with the given name.</summary>
        /// <param name="name">书签名，null 视为空字符串。The bookmark name; null is treated as an empty string.</param>
        public BookmarkStart(string name)
        {
            Name = name ?? string.Empty;
        }
    }

    /// <summary>书签范围终点行内标记。The inline marker that closes a bookmark range.</summary>
    public sealed class BookmarkEnd : Inline
    {
        /// <summary>书签名；读取 DOCX 时由 <c>w:bookmarkEnd</c> 的 <c>w:id</c> 反查填充。
        /// The bookmark name; when reading DOCX it is resolved from the matching <c>w:bookmarkEnd</c> <c>w:id</c>.</summary>
        public string Name { get; set; }

        /// <summary>以指定名称创建终点标记。Creates an end marker with the given name.</summary>
        /// <param name="name">书签名，null 视为空字符串。The bookmark name; null is treated as an empty string.</param>
        public BookmarkEnd(string name)
        {
            Name = name ?? string.Empty;
        }
    }

    /// <summary>书签名规则校验：Word 要求 1–40 个字符、以字母或下划线开头、仅含字母数字下划线。
    /// Validates bookmark names against the Word rules: 1 to 40 characters, starting with a letter or underscore,
    /// containing only letters, digits and underscores.</summary>
    public static class BookmarkNames
    {
        /// <summary>判断名称是否为合法 Word 书签名。Determines whether the name is a legal Word bookmark name.</summary>
        /// <param name="name">待校验名称。The name to test.</param>
        /// <returns>合法返回 true。True when legal.</returns>
        public static bool IsValid(string name)
        {
            if (string.IsNullOrEmpty(name) || name.Length > 40) return false;
            if (!IsLetter(name[0]) && name[0] != '_') return false;
            for (int i = 1; i < name.Length; i++)
            {
                char c = name[i];
                if (!IsLetter(c) && !(c >= '0' && c <= '9') && c != '_') return false;
            }
            return true;
        }

        /// <summary>校验名称，非法时抛出 <see cref="ArgumentException"/>。Validates a name, throwing ArgumentException when illegal.</summary>
        /// <param name="name">待校验名称。The name to validate.</param>
        /// <returns>去除首尾空白后的名称。The trimmed name.</returns>
        public static string Validate(string name)
        {
            string trimmed = name == null ? string.Empty : name.Trim();
            if (!IsValid(trimmed))
            {
                throw new ArgumentException("书签名须为 1-40 个字符、以字母或下划线开头、仅含字母数字下划线。"
                    + "A bookmark name must be 1-40 characters, start with a letter or underscore, "
                    + "and contain only letters, digits and underscores.", "name");
            }
            return trimmed;
        }

        /// <summary>判断是否为字母（含中日韩等文字，便于中文书签名）。Determines whether the character is a letter,
        /// accepting CJK so Chinese bookmark names remain usable.</summary>
        /// <param name="value">待判断字符。The character.</param>
        /// <returns>是字母返回 true。True for a letter.</returns>
        private static bool IsLetter(char value)
        {
            return (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z') || char.IsLetter(value);
        }
    }

    /// <summary>
    /// 一个已配对的书签：名称、范围端点与可见文本。整段书签（仅 <see cref="Paragraph.BookmarkName"/>）
    /// 视为覆盖整段的特殊形态。
    /// A resolved bookmark: its name, range endpoints and visible text. A whole-paragraph bookmark
    /// (carried only by <see cref="Paragraph.BookmarkName"/>) is the special case covering the entire paragraph.
    /// </summary>
    public sealed class Bookmark
    {
        private readonly BookmarkCollection _owner;
        private readonly Paragraph _paragraph;

        internal Bookmark(BookmarkCollection owner, string name, BookmarkStart start, BookmarkEnd end,
            Paragraph paragraph)
        {
            _owner = owner;
            Name = name;
            Start = start;
            End = end;
            _paragraph = paragraph;
        }

        /// <summary>书签名。The bookmark name.</summary>
        public string Name { get; set; }

        /// <summary>范围起点标记；整段书签为 null。The start marker; null for a whole-paragraph bookmark.</summary>
        public BookmarkStart Start { get; private set; }

        /// <summary>范围终点标记；源文件中 start/end 未配对时为 null。The end marker; null when the source left the range unpaired.</summary>
        public BookmarkEnd End { get; private set; }

        /// <summary>整段书签所属段落；非覆盖整段的范围书签为 null。The paragraph when the bookmark covers a whole paragraph; null for a range that does not.</summary>
        public Paragraph Paragraph { get { return _paragraph; } }

        /// <summary>是否为覆盖整段的书签：仅有段落名，或起点是段内第一个元素且终点是最后一个。Whether the bookmark covers
        /// a whole paragraph: either the name-only form, or a range whose start is the first inline and whose end is the last.</summary>
        public bool IsWholeParagraph { get { return _paragraph != null; } }

        /// <summary>范围是否完整：整段书签恒为完整，范围书签需两端齐备。
        /// Whether the range is complete: always true for whole-paragraph bookmarks, both ends required for ranges.</summary>
        public bool IsValid
        {
            get { return _paragraph != null || (Start != null && End != null); }
        }

        /// <summary>
        /// 读取书签覆盖的可见文本；跨段时以换行分隔，图片按 <c>[image]</c> 占位。
        /// Reads the visible text covered by the bookmark; paragraphs are joined with a newline and images
        /// appear as <c>[image]</c>.
        /// </summary>
        /// <returns>书签文本；未配对的空范围返回空字符串。The bookmark text; an unpaired empty range returns an empty string.</returns>
        public string Text
        {
            get { return GetText(); }
            set { SetText(value); }
        }

        /// <summary>取书签文本。Reads the bookmark text.</summary>
        /// <returns>文本内容。The text.</returns>
        private string GetText()
        {
            if (_paragraph != null && Start == null) return _paragraph.GetText();
            if (End == null) return string.Empty;
            List<Paragraph> paragraphs = _owner.Paragraphs();
            int startIndex = IndexOfStart(paragraphs);
            if (startIndex < 0) return string.Empty;
            int endIndex = startIndex;
            int endOffset = 0;
            if (!FindInline(paragraphs, End, startIndex, out endIndex, out endOffset)) return string.Empty;

            StringBuilder builder = new StringBuilder();
            for (int p = startIndex; p <= endIndex; p++)
            {
                List<Inline> inlines = paragraphs[p].Inlines;
                int from = p == startIndex ? InlineOffset(paragraphs[p], Start) + 1 : 0;
                int to = p == endIndex ? endOffset : inlines.Count;
                for (int i = from; i < to && i < inlines.Count; i++)
                {
                    AppendInlineText(inlines[i], builder);
                }
                if (p < endIndex) builder.Append('\n');
            }
            return builder.ToString();
        }

        /// <summary>
        /// 用新文本替换书签内容：书签本身保留，范围内原有行内元素（含嵌套书签标记）被移除，
        /// 新文本以单个文本运行插入起点之后。整段书签则替换整段内容。
        /// Replaces the bookmark content: the bookmark survives, every inline inside the range (including nested
        /// bookmark markers) is removed, and the new text is inserted as a single run after the start.
        /// A whole-paragraph bookmark replaces the whole paragraph.
        /// </summary>
        /// <param name="value">新文本，null 视为空字符串。The new text; null is treated as an empty string.</param>
        private void SetText(string value)
        {
            string text = value ?? string.Empty;
            if (_paragraph != null && Start == null)
            {
                _paragraph.Inlines.Clear();
                if (text.Length > 0) _paragraph.Inlines.Add(new TextRun(text));
                return;
            }
            if (Start == null || End == null)
            {
                throw new InvalidOperationException("书签范围未配对，无法替换内容。"
                    + "The bookmark range is unpaired, so its content cannot be replaced.");
            }

            List<Paragraph> paragraphs = _owner.Paragraphs();
            int startIndex;
            int startOffset;
            if (!FindInline(paragraphs, Start, 0, out startIndex, out startOffset)) return;
            int endIndex;
            int endOffset;
            if (!FindInline(paragraphs, End, startIndex, out endIndex, out endOffset)) return;
            if (endIndex < startIndex || (endIndex == startIndex && endOffset < startOffset)) return;

            // 从后往前删，避免删除过程中下标漂移。Delete back to front so indices cannot drift while removing.
            for (int p = endIndex; p >= startIndex; p--)
            {
                List<Inline> inlines = paragraphs[p].Inlines;
                int from = p == startIndex ? startOffset + 1 : 0;
                int to = p == endIndex ? endOffset - 1 : inlines.Count - 1;
                for (int i = to; i >= from && i < inlines.Count; i--)
                {
                    inlines.RemoveAt(i);
                }
            }

            if (text.Length > 0)
            {
                List<Inline> startInlines = paragraphs[startIndex].Inlines;
                int insertAt = InlineOffset(paragraphs[startIndex], Start) + 1;
                if (insertAt < 0) insertAt = startInlines.Count;
                startInlines.Insert(insertAt, new TextRun(text));
            }
        }

        /// <summary>
        /// 从文档中移除书签：删除两端标记但保留其间正文，整段书签清除段落书签名。
        /// Removes the bookmark from the document: both markers are deleted while the covered content stays,
        /// and a whole-paragraph bookmark loses its paragraph name.
        /// </summary>
        public void Remove()
        {
            // 先定位所属段落：标记一旦从树里摘掉就再也查不到。Locate the owning paragraph first: detached
            // markers can no longer be found in the tree.
            Paragraph owner = _paragraph;
            if (owner == null) owner = _owner.ParagraphOf(Start) ?? _owner.ParagraphOf(End);
            if (Start != null) RemoveFromParagraph(Start);
            if (End != null) RemoveFromParagraph(End);
            // 范围书签所在段可能同时带着同名的段落书签名（读取时一并填充），否则移除后书签会复活。
            // The start paragraph can also carry the same name (the reader fills both), so clear it too or the
            // bookmark would reappear.
            if (owner != null && string.Equals(owner.BookmarkName, Name, StringComparison.Ordinal))
            {
                owner.BookmarkName = null;
            }
        }

        /// <summary>把标记从其所在段落移除。Removes a marker from the paragraph holding it.</summary>
        /// <param name="marker">待移除标记。The marker to remove.</param>
        private void RemoveFromParagraph(Inline marker)
        {
            Paragraph paragraph = _owner.ParagraphOf(marker);
            if (paragraph == null) return;
            int index = paragraph.Inlines.IndexOf(marker);
            if (index >= 0) paragraph.Inlines.RemoveAt(index);
        }

        /// <summary>取起点所在段落序号。Finds the index of the paragraph holding the start marker.</summary>
        /// <param name="paragraphs">段落序列。The paragraph sequence.</param>
        /// <returns>段落序号，未找到返回 -1。The paragraph index, or -1 when absent.</returns>
        private int IndexOfStart(List<Paragraph> paragraphs)
        {
            int index;
            int offset;
            return FindInline(paragraphs, Start, 0, out index, out offset) ? index : -1;
        }

        /// <summary>取行内元素在段中的下标。Finds the index of an inline within its paragraph.</summary>
        /// <param name="paragraph">所在段落。The owning paragraph.</param>
        /// <param name="inline">目标元素。The target inline.</param>
        /// <returns>下标，未找到返回 -1。The index, or -1 when absent.</returns>
        private static int InlineOffset(Paragraph paragraph, Inline inline)
        {
            return paragraph.Inlines.IndexOf(inline);
        }

        /// <summary>
        /// 在段落序列中定位行内元素。Locates an inline element inside a paragraph sequence.
        /// </summary>
        /// <param name="paragraphs">段落序列。The paragraph sequence.</param>
        /// <param name="inline">目标元素。The target inline.</param>
        /// <param name="fromParagraph">起始段落序号。The first paragraph index to scan.</param>
        /// <param name="paragraphIndex">输出：所在段落序号。Output: the paragraph index.</param>
        /// <param name="inlineIndex">输出：段内下标。Output: the index inside the paragraph.</param>
        /// <returns>找到返回 true。True when found.</returns>
        private static bool FindInline(List<Paragraph> paragraphs, Inline inline, int fromParagraph,
            out int paragraphIndex, out int inlineIndex)
        {
            for (int p = fromParagraph; p < paragraphs.Count; p++)
            {
                int index = paragraphs[p].Inlines.IndexOf(inline);
                if (index >= 0)
                {
                    paragraphIndex = p;
                    inlineIndex = index;
                    return true;
                }
            }
            paragraphIndex = 0;
            inlineIndex = 0;
            return false;
        }

        /// <summary>把单个行内元素的可读文本追加到构建器。Appends the readable text of one inline to the builder.</summary>
        /// <param name="inline">行内元素。The inline.</param>
        /// <param name="builder">接收文本的构建器。The builder receiving text.</param>
        private static void AppendInlineText(Inline inline, StringBuilder builder)
        {
            TextRun run = inline as TextRun;
            if (run != null)
            {
                builder.Append(run.Text);
                return;
            }
            FieldRun field = inline as FieldRun;
            if (field != null)
            {
                builder.Append(field.DisplayText);
                return;
            }
            if (inline is ImageRun) builder.Append("[image]");
        }
    }

    /// <summary>
    /// 文档书签集合，按文档顺序遍历正文与页眉页脚；每次访问即时解析，
    /// 因此新增/删除标记后无需手动刷新。The document bookmark collection: it walks body, header and footer
    /// content in document order and resolves pairs on every access, so markers need no manual refresh.
    /// </summary>
    public sealed class BookmarkCollection : IEnumerable<Bookmark>
    {
        private readonly OfficeDocument _document;

        internal BookmarkCollection(OfficeDocument document)
        {
            _document = document;
        }

        /// <summary>书签数量。The number of bookmarks.</summary>
        public int Count { get { return Resolve().Count; } }

        /// <summary>按名称取书签，未找到返回 null。Gets a bookmark by name, returning null when absent.</summary>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <returns>书签或 null。The bookmark or null.</returns>
        public Bookmark this[string name]
        {
            get
            {
                List<Bookmark> all = Resolve();
                for (int i = 0; i < all.Count; i++)
                {
                    if (string.Equals(all[i].Name, name, StringComparison.Ordinal)) return all[i];
                }
                return null;
            }
        }

        /// <summary>按文档顺序取书签，越界返回 null。Gets a bookmark by document order, returning null out of range.</summary>
        /// <param name="index">从零开始的序号。The zero-based index.</param>
        /// <returns>书签或 null。The bookmark or null.</returns>
        public Bookmark this[int index]
        {
            get
            {
                List<Bookmark> all = Resolve();
                return index >= 0 && index < all.Count ? all[index] : null;
            }
        }

        /// <summary>判断是否存在同名书签。Determines whether a bookmark with the name exists.</summary>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <returns>存在返回 true。True when present.</returns>
        public bool Contains(string name)
        {
            return this[name] != null;
        }

        /// <summary>全部书签的数组快照。An array snapshot of every bookmark.</summary>
        /// <returns>书签数组。The bookmarks.</returns>
        public Bookmark[] ToArray()
        {
            return Resolve().ToArray();
        }

        /// <summary>
        /// 用已存在的端点登记一个书签；两端必须已经在文档树中。
        /// Registers a bookmark from existing endpoints; both must already live in the document tree.
        /// </summary>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <param name="start">起点标记。The start marker.</param>
        /// <param name="end">终点标记。The end marker.</param>
        /// <returns>登记后的书签。The resulting bookmark.</returns>
        public Bookmark Add(string name, BookmarkStart start, BookmarkEnd end)
        {
            string validated = BookmarkNames.Validate(name);
            if (start == null) throw new ArgumentNullException("start");
            if (end == null) throw new ArgumentNullException("end");
            List<Paragraph> paragraphs = Paragraphs();
            int paragraphIndex;
            int inlineIndex;
            if (!FindInlineHelper(paragraphs, start, out paragraphIndex, out inlineIndex))
            {
                throw new ArgumentException("起点必须已在文档中。The start marker must already be in the document.", "start");
            }
            if (!FindInlineHelper(paragraphs, end, out paragraphIndex, out inlineIndex))
            {
                throw new ArgumentException("终点必须已在文档中。The end marker must already be in the document.", "end");
            }
            start.Name = validated;
            end.Name = validated;
            return new Bookmark(this, validated, start, end, null);
        }

        /// <summary>
        /// 覆盖起止两个段落（含其间所有段落）并登记书签，端点分别插入两段的首尾。
        /// Spans two paragraphs (and everything between), inserting the endpoints at their edges.
        /// </summary>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <param name="startParagraph">起始段落。The first paragraph.</param>
        /// <param name="endParagraph">结束段落。The last paragraph.</param>
        /// <returns>登记后的书签。The resulting bookmark.</returns>
        public Bookmark Add(string name, Paragraph startParagraph, Paragraph endParagraph)
        {
            string validated = BookmarkNames.Validate(name);
            if (startParagraph == null) throw new ArgumentNullException("startParagraph");
            if (endParagraph == null) throw new ArgumentNullException("endParagraph");
            BookmarkStart start = new BookmarkStart(validated);
            BookmarkEnd end = new BookmarkEnd(validated);
            startParagraph.Inlines.Insert(0, start);
            endParagraph.Inlines.Add(end);
            return new Bookmark(this, validated, start, end, null);
        }

        /// <summary>移除指定名称的书签，内容保留。Removes a named bookmark, keeping its content.</summary>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <returns>移除成功返回 true。True when a bookmark was removed.</returns>
        public bool Remove(string name)
        {
            Bookmark bookmark = this[name];
            if (bookmark == null) return false;
            bookmark.Remove();
            return true;
        }

        /// <summary>移除全部书签，正文保留。Removes every bookmark, keeping all content.</summary>
        public void Clear()
        {
            Bookmark[] all = ToArray();
            for (int i = 0; i < all.Length; i++) all[i].Remove();
        }

        /// <summary>枚举书签。Enumerates the bookmarks.</summary>
        /// <returns>书签枚举器。The bookmark enumerator.</returns>
        public IEnumerator<Bookmark> GetEnumerator()
        {
            return Resolve().GetEnumerator();
        }

        /// <summary>非泛型枚举。Non-generic enumeration.</summary>
        /// <returns>书签枚举器。The bookmark enumerator.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// 查找行内元素所属段落。
        /// Finds the paragraph owning an inline element.
        /// </summary>
        /// <param name="inline">目标行内元素。The inline to locate.</param>
        /// <returns>所属段落；元素不在文档中或入参为 null 时返回 null。The owning paragraph, or null when the element is absent or null.</returns>
        public Paragraph ParagraphOf(Inline inline)
        {
            if (inline == null) return null;
            List<Paragraph> paragraphs = Paragraphs();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (paragraphs[i].Inlines.IndexOf(inline) >= 0) return paragraphs[i];
            }
            return null;
        }

        /// <summary>
        /// 文档内全部段落，含表格单元格内的段落，按正文、页眉、页脚的阅读顺序返回。
        /// Every paragraph in the document, including table cell paragraphs, in reading order across body,
        /// header and footer.
        /// </summary>
        /// <returns>段落序列。The paragraphs.</returns>
        internal List<Paragraph> Paragraphs()
        {
            List<Paragraph> result = new List<Paragraph>();
            List<Section> sections = _document.Sections;
            for (int i = 0; i < sections.Count; i++)
            {
                Section section = sections[i];
                if (section == null) continue;
                CollectParagraphs(section.Blocks, result);
                if (section.Header != null) CollectParagraphs(section.Header.Blocks, result);
                if (section.Footer != null) CollectParagraphs(section.Footer.Blocks, result);
            }
            return result;
        }

        /// <summary>递归收集块列表中的段落。Collects paragraphs from a block list, recursing into tables.</summary>
        /// <param name="blocks">块列表。The blocks.</param>
        /// <param name="result">累加容器。The accumulator.</param>
        private static void CollectParagraphs(List<Block> blocks, List<Paragraph> result)
        {
            if (blocks == null) return;
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    result.Add(paragraph);
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null || table.Rows == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    TableRow row = table.Rows[r];
                    if (row == null || row.Cells == null) continue;
                    for (int c = 0; c < row.Cells.Count; c++)
                    {
                        TableCell cell = row.Cells[c];
                        if (cell != null) CollectParagraphs(cell.Blocks, result);
                    }
                }
            }
        }

        /// <summary>解析文档中的全部书签。Resolves every bookmark present in the document.</summary>
        /// <returns>书签列表。The bookmarks.</returns>
        private List<Bookmark> Resolve()
        {
            List<Bookmark> result = new List<Bookmark>();
            List<Paragraph> paragraphs = Paragraphs();
            for (int p = 0; p < paragraphs.Count; p++)
            {
                Paragraph paragraph = paragraphs[p];
                List<Inline> inlines = paragraph.Inlines;

                for (int i = 0; i < inlines.Count; i++)
                {
                    BookmarkStart start = inlines[i] as BookmarkStart;
                    if (start == null) continue;
                    BookmarkEnd end = FindEnd(paragraphs, p, i, start.Name);
                    // 起点是段内第一个、终点是最后一个时，这个范围就是"整段书签"。
                    // A range whose start is the first inline and end the last is a whole-paragraph bookmark.
                    Paragraph covered = null;
                    int endParagraph;
                    int endOffset;
                    if (end != null && i == 0
                        && FindInlineHelper(paragraphs, end, out endParagraph, out endOffset)
                        && endParagraph == p
                        && endOffset == paragraph.Inlines.Count - 1)
                    {
                        covered = paragraph;
                    }
                    result.Add(new Bookmark(this, start.Name, start, end, covered));
                }

                // 整段书签：仅有段落名、段内没有对应的显式起点时才算独立形态。
                // Whole-paragraph bookmark: counted only when no explicit start marker of the same name exists.
                string legacy = paragraph.BookmarkName;
                if (!string.IsNullOrEmpty(legacy) && !HasNamedStart(paragraph, legacy))
                {
                    result.Add(new Bookmark(this, legacy, null, null, paragraph));
                }
            }
            return result;
        }

        /// <summary>判断段落是否已带指定名称的显式起点。Determines whether the paragraph holds an explicit start of the name.</summary>
        /// <param name="paragraph">目标段落。The paragraph.</param>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <returns>存在返回 true。True when present.</returns>
        private static bool HasNamedStart(Paragraph paragraph, string name)
        {
            List<Inline> inlines = paragraph.Inlines;
            for (int i = 0; i < inlines.Count; i++)
            {
                BookmarkStart start = inlines[i] as BookmarkStart;
                if (start != null && string.Equals(start.Name, name, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        /// <summary>向后查找与起点配对的终点。Finds the end marker paired with a start marker.</summary>
        /// <param name="paragraphs">段落序列。The paragraph sequence.</param>
        /// <param name="fromParagraph">起点段落序号。The start paragraph index.</param>
        /// <param name="fromInline">起点行内下标。The start inline index.</param>
        /// <param name="name">书签名。The bookmark name.</param>
        /// <returns>终点标记，未配对返回 null。The end marker, or null when unpaired.</returns>
        private static BookmarkEnd FindEnd(List<Paragraph> paragraphs, int fromParagraph, int fromInline, string name)
        {
            List<Inline> startInlines = paragraphs[fromParagraph].Inlines;
            for (int i = fromInline + 1; i < startInlines.Count; i++)
            {
                BookmarkEnd end = startInlines[i] as BookmarkEnd;
                if (end != null && string.Equals(end.Name, name, StringComparison.Ordinal)) return end;
            }
            for (int p = fromParagraph + 1; p < paragraphs.Count; p++)
            {
                List<Inline> inlines = paragraphs[p].Inlines;
                for (int i = 0; i < inlines.Count; i++)
                {
                    BookmarkEnd end = inlines[i] as BookmarkEnd;
                    if (end != null && string.Equals(end.Name, name, StringComparison.Ordinal)) return end;
                }
            }
            return null;
        }

        /// <summary>在段落序列中定位行内元素。Locates an inline inside the paragraph sequence.</summary>
        /// <param name="paragraphs">段落序列。The paragraph sequence.</param>
        /// <param name="inline">目标元素。The target inline.</param>
        /// <param name="paragraphIndex">输出：段落序号。Output: the paragraph index.</param>
        /// <param name="inlineIndex">输出：段内下标。Output: the inline index.</param>
        /// <returns>找到返回 true。True when found.</returns>
        private static bool FindInlineHelper(List<Paragraph> paragraphs, Inline inline,
            out int paragraphIndex, out int inlineIndex)
        {
            for (int p = 0; p < paragraphs.Count; p++)
            {
                int index = paragraphs[p].Inlines.IndexOf(inline);
                if (index >= 0)
                {
                    paragraphIndex = p;
                    inlineIndex = index;
                    return true;
                }
            }
            paragraphIndex = 0;
            inlineIndex = 0;
            return false;
        }
    }
}
