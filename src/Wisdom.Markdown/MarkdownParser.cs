using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Markdown
{
    /// <summary>
    /// Markdown 解析器：把 CommonMark 主体 + GFM 扩展的纯文本解析为统一文档模型 <see cref="OfficeDocument"/>。
    /// 行内解析采用单遍扫描并带明确回退（未闭合的分隔符按字面文本输出），不使用嵌套量词正则，确保超长输入不会指数爆炸。
    /// A Markdown parser that turns CommonMark plus GFM plain text into the unified document model <see cref="OfficeDocument"/>.
    /// Inline parsing is a single pass with explicit fallback (unclosed delimiters emit as literal text); no nested-quantifier regex, so long inputs cannot blow up exponentially.
    /// </summary>
    internal sealed class MarkdownParser
    {
        /// <summary>列表每嵌套一层增加的左缩进（磅），与两个写器的层级换算一致。Left indent added per list nesting level, in points; matches the level conversion used by both writers.</summary>
        private const double IndentPerLevel = 24.0;

        /// <summary>分割线在文档模型里的承载文本；两个写器按同一常量识别回 `---` 与 <c>&lt;hr /&gt;</c>。The text a thematic break carries in the document model; both writers match this same constant to restore <c>---</c> and <c>&lt;hr /&gt;</c>.</summary>
        internal const string ThematicBreakText = "＊ ＊ ＊";

        /// <summary>
        /// 围栏与缩进代码块段落的结构标记。模型里没有独立的代码块节点，等宽字体是唯一可见信号，
        /// 而这个信号与行内代码用的完全相同（整段只有一个行内代码的段落会被误判成代码块），
        /// 所以解析端在此打标记，两个写器优先读它，字体推断只留给手写文档兜底。
        /// Structural marker of a fenced or indented code-block paragraph. The model has no code-block node, so the
        /// monospaced font is the only visible signal - and it is indistinguishable from an inline code span, which
        /// turns a paragraph holding nothing but one code span into a bogus code block. The parser marks these
        /// paragraphs and both writers trust the marker first, keeping the font heuristic as a fallback for
        /// hand-built documents.
        /// </summary>
        internal const string CodeStyleName = "MarkdownCode";

        /// <summary>经行内解析生成的段落标记，含义是「这是正文，不是代码块」。两处标记互斥，其它样式名一律走字体推断。The marker of a paragraph built by the inline scanner, meaning body text rather than a code block. The two markers are exclusive; any other style name falls through to the font heuristic.</summary>
        internal const string InlineStyleName = "MarkdownInline";

        /// <summary>行内元素最大嵌套深度，超出后按字面文本输出，防止栈溢出。The maximum inline nesting depth; deeper content becomes literal text, preventing a stack overflow.</summary>
        private const int MaxInlineDepth = 64;

        /// <summary>
        /// 单张表格的列数上限：列数取自表头行的竖线，其后每一行都要补齐到这个列数，
        /// 所以一行几万个竖线就能把单元格数量按行数乘出来。
        /// Ceiling for the columns of one table: the column count comes from the pipes of the header line and
        /// every later row is padded to it, so a line with tens of thousands of pipes multiplies the cell count
        /// by the row count.
        /// </summary>
        private const int MaxTableColumns = 4096;

        /// <summary>
        /// 单张表格的单元格总数上限，与 XlsxReader 的合并区域同一量级：列数合规而行数无限时，
        /// 乘积照样把模型吃穿，而下游的 Word 与 PDF 写出早在几万个格子之前就撑不住了。
        /// Ceiling for the cells of one table, at the same order as the merged-range cap in XlsxReader: legal
        /// columns over an endless number of rows still eat the model, and the Word and PDF writers give up long
        /// before tens of thousands of cells.
        /// </summary>
        private const int MaxTableCells = 65536;

        /// <summary>
        /// 单篇文档嵌入图片的字节总上限：单张封顶挡不住几十万个不同路径的引用，
        /// 而每一份都会整份留在模型里，HTML 写出时还要再 base64 放大四分之三。
        /// Ceiling for the image bytes of one document: the per-file cap cannot stop a hundred thousand
        /// references to distinct paths, and every one of them stays in the model, then grows by three quarters
        /// again when the HTML writer base64-encodes it.
        /// </summary>
        private const int MaxDocumentImageBytes = 256 * 1024 * 1024;

        private readonly MarkdownOptions options;

        /// <summary>
        /// 链接引用定义表（CommonMark 的 [label]: url "title"），键为归一化后的标签；整篇解析开始前一次性收集。
        /// Link reference definitions (CommonMark's [label]: url "title"), keyed by the normalized label and
        /// collected once before the block pass starts.
        /// </summary>
        private readonly Dictionary<string, LinkDefinition> _definitions = new Dictionary<string, LinkDefinition>(StringComparer.Ordinal);

        /// <summary>
        /// 引用块嵌套层数上限：每层引用都要把剥掉 &gt; 的内部行再解析一遍，一行几万个 &gt;
        /// 就会递归同样多次而爆栈。超过上限的引用标记按字面文本留在段落里。
        /// Ceiling for block-quote nesting: every level re-parses its '>'-stripped lines, so a line of tens of
        /// thousands of '&gt;' would recurse that many times and overflow the stack. Beyond the ceiling the remaining
        /// markers stay as literal text in the paragraph.
        /// </summary>
        private const int MaxQuoteDepth = 24;

        /// <summary>当前引用嵌套深度，仅解析期间使用。The block-quote nesting depth currently in effect, used only during a parse.</summary>
        private int _quoteDepth;

        /// <summary>引用块在正文左侧增加的一格缩进（磅），与旧版单层引用的观感保持一致。The one column of left indent a block quote adds, in points, matching how a single-level quote used to look.</summary>
        private const double QuoteIndent = 18.0;

        /// <summary>
        /// 单篇文档链接引用定义的数量上限：定义表整篇常驻内存，无上限时几万个不同标签即可把模型撑大。
        /// Ceiling for the link reference definitions of one document: the table stays resident for the whole parse,
        /// so tens of thousands of distinct labels would bloat the model.
        /// </summary>
        private const int MaxLinkDefinitions = 4096;

        /// <summary>链接引用定义里的标签长度上限（字符）。Ceiling for the label length of a link reference definition, in characters.</summary>
        private const int MaxDefinitionLabelChars = 1000;

        /// <summary>
        /// 本次解析已读入的图片，按归一化后的绝对路径缓存：同一张图被引用几百次也只读一次、只存一份。
        /// Images loaded by this parse, keyed by normalized absolute path: a picture referenced hundreds of times
        /// is read once and stored once.
        /// </summary>
        private readonly Dictionary<string, OfficeImage> _imageCache = new Dictionary<string, OfficeImage>(StringComparer.Ordinal);

        /// <summary>本次解析累计读入的图片字节数。The image bytes loaded so far by this parse.</summary>
        private long _imageBytes;

        /// <summary>一条链接引用定义的目标与可选标题。The destination and optional title of one link reference definition.</summary>
        private sealed class LinkDefinition
        {
            /// <summary>链接目标（已去尖括号与转义）。The destination, angle brackets and backslash escapes removed.</summary>
            public string Destination;

            /// <summary>可选标题，未提供为 null。The optional title, null when absent.</summary>
            public string Title;
        }

        /// <summary>
        /// 以指定选项构造解析器。
        /// Constructs the parser with the given options.
        /// </summary>
        /// <param name="options">
        /// 解析选项，null 时使用默认选项。The parse options; default options are used when null.
        /// </param>
        internal MarkdownParser(MarkdownOptions options)
        {
            this.options = options ?? new MarkdownOptions();
        }

        /// <summary>
        /// 将 Markdown 文本解析为文档模型。null/空文本返回一个仅含空节的合法文档，绝不抛未捕获异常。
        /// Parses Markdown text into the document model. Null/empty input returns a valid document with one empty section and never throws.
        /// </summary>
        /// <param name="markdown">
        /// Markdown 源文本。The Markdown source text.
        /// </param>
        /// <returns>
        /// 解析得到的文档。The parsed document.
        /// </returns>
        internal OfficeDocument Parse(string markdown)
        {
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();

            if (string.IsNullOrEmpty(markdown))
            {
                return document;
            }

            string[] lines = NormalizeLines(markdown);
            int index = 0;

            // 处理 YAML front-matter（首行是 --- 的块），仅识别 title/author/date。
            // Handle YAML front matter (a leading --- block); only title/author/date are recognised.
            index = ParseFrontMatter(lines, index, document);

            // 先一次性收集链接引用定义：定义可出现在文档任意块起始处，必须早于行内解析才知道。
            // Collect link reference definitions first: they may appear at any block start, and inline parsing has
            // to know them before it runs.
            bool[] definitionLines = CollectLinkDefinitions(lines, index, true);

            ParseBlocks(lines, index, section, document, definitionLines);

            return document;
        }

        /// <summary>
        /// 块级主循环：按 CommonMark 的优先级依次判定围栏代码、ATX 标题、引用、分隔线、缩进代码、列表、表格、段落。
        /// 引用块会把内部行交给本方法递归解析，所以循环体不留在 <see cref="Parse"/> 里。
        /// The block-level main loop: CommonMark precedence, testing fenced code, ATX headings, block quotes, thematic
        /// breaks, indented code, lists, tables and paragraphs in turn. Block quotes hand their stripped inner lines
        /// back to this method, which is why the loop lives here rather than in <see cref="Parse"/>.
        /// </summary>
        /// <param name="lines">待解析行。The lines to parse.</param>
        /// <param name="start">起始行。The start index.</param>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="document">所属文档，用于元数据标题。The owning document, used for the metadata title.</param>
        /// <param name="definitionLines">与 <paramref name="lines"/> 等长的链接定义标记。Definition markers, as long as <paramref name="lines"/>.</param>
        private void ParseBlocks(string[] lines, int start, Section section, OfficeDocument document, bool[] definitionLines)
        {
            int index = start;
            int nextHeadingLevel = -1;

            while (index < lines.Length)
            {
                string line = lines[index];

                if (line.Trim().Length == 0)
                {
                    index++;
                    continue;
                }

                // 引用定义行只贡献给链接表，自身不产生任何块。
                // A definition line only feeds the link table and produces no block of its own.
                if (definitionLines[index])
                {
                    index++;
                    continue;
                }

                // 围栏代码块 ``` 或 ~~~。
                // Fenced code block ``` or ~~~.
                string fenceInfo;
                if (IsFenceStart(line, out fenceInfo))
                {
                    index = ParseFencedCode(lines, index, section, fenceInfo);
                    continue;
                }

                // ATX 标题 #..######。
                // ATX heading #..######.
                int atxLevel;
                string atxText;
                if (IsAtxHeading(line, out atxLevel, out atxText))
                {
                    AddHeading(section, atxLevel, atxText);
                    if (atxLevel == 1 && nextHeadingLevel == -1)
                    {
                        nextHeadingLevel = 0;
                        if (string.IsNullOrEmpty(document.Metadata.Title))
                        {
                            document.Metadata.Title = PlainFromInline(atxText);
                        }
                    }
                    index++;
                    continue;
                }

                // 引用块 >。
                // Block quote >.
                if (StartsQuote(line))
                {
                    index = ParseQuote(lines, index, section, document);
                    continue;
                }

                // 独立水平分隔线（≥3 个 -/*/_，可含空格）。必须先于列表判定，
                // 否则 "* * *" 与 "- - -" 会被当成三个空列表项。
                // Standalone thematic break (three or more -/*/_ with optional spaces). Must precede the list
                // test, otherwise "* * *" and "- - -" would read as a list of empty items.
                if (IsThematicBreak(line))
                {
                    AddThematicBreak(section);
                    index++;
                    continue;
                }

                // 4 空格（或制表符）缩进代码块。同样先于列表：顶层没有打开的列表时，
                // 缩进四格只可能是代码，CommonMark 不允许它作为列表开始。
                // Indented code block (4 spaces or a tab), also ahead of the list test: with no list in flight an
                // four-space indent can only be code, since CommonMark refuses it as a list start.
                if (IsIndentedCode(line))
                {
                    index = ParseIndentedCode(lines, index, section);
                    continue;
                }

                // 列表项（无序 -/*/+ 与有序 1. / 1)），支持缩进嵌套。
                // List item (unordered -/*/+ and ordered 1. / 1)), with indentation nesting.
                if (IsListItem(line))
                {
                    index = ParseList(lines, index, section);
                    continue;
                }

                // GFM 管道表格。
                // GFM pipe table.
                if (options.EnableTables && IsTableHeader(lines, index))
                {
                    index = ParseTable(lines, index, section);
                    continue;
                }

                // 普通段落（含 Setext 标题的合并判定）。
                // A normal paragraph (including Setext-heading detection for the following line).
                index = ParseParagraph(lines, index, section, document);
            }
        }

        /// <summary>
        /// 统一换行并按 \n 拆分为行数组。
        /// Normalises line endings and splits the text into lines by '\n'.
        /// </summary>
        /// <param name="markdown">源文本。The source text.</param>
        /// <returns>行数组。The array of lines.</returns>
        private static string[] NormalizeLines(string markdown)
        {
            string text = markdown.Replace("\r\n", "\n").Replace('\r', '\n');
            return text.Split('\n');
        }

        /// <summary>
        /// 解析开头的 YAML front-matter 块并写入元数据，返回块之后的起始行号。
        /// Parses the leading YAML front-matter block into metadata and returns the index of the first line after it.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">当前起始行。The current start index.</param>
        /// <param name="document">目标文档。The target document.</param>
        /// <returns>front-matter 之后的行号。The line index following the front matter.</returns>
        private static int ParseFrontMatter(string[] lines, int start, OfficeDocument document)
        {
            if (start >= lines.Length)
            {
                return start;
            }
            if (lines[start].Trim() != "---")
            {
                return start;
            }

            int end = -1;
            for (int i = start + 1; i < lines.Length; i++)
            {
                string trimmed = lines[i].Trim();
                if (trimmed == "---" || trimmed == "...")
                {
                    end = i;
                    break;
                }
            }
            if (end < 0)
            {
                // 没有闭合分隔符，按普通内容处理，不当作 front-matter。
                // No closing delimiter: treat as normal content, not front matter.
                return start;
            }

            for (int i = start + 1; i < end; i++)
            {
                string line = lines[i];
                int colon = line.IndexOf(':');
                if (colon <= 0)
                {
                    continue;
                }
                string key = line.Substring(0, colon).Trim();
                string value = Unquote(line.Substring(colon + 1).Trim());

                if (string.Equals(key, "title", StringComparison.OrdinalIgnoreCase))
                {
                    document.Metadata.Title = value;
                }
                else if (string.Equals(key, "author", StringComparison.OrdinalIgnoreCase))
                {
                    document.Metadata.Author = value;
                }
                else if (string.Equals(key, "date", StringComparison.OrdinalIgnoreCase))
                {
                    DateTime parsed;
                    if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsed))
                    {
                        document.Metadata.Created = parsed;
                        document.Metadata.Modified = parsed;
                    }
                }
                // 其余键忽略。Other keys are ignored.
            }

            return end + 1;
        }

        /// <summary>
        /// 去除字符串两端成对的单/双引号。
        /// Removes a matched pair of leading/trailing single or double quotes.
        /// </summary>
        /// <param name="value">原始值。The raw value.</param>
        /// <returns>去引号后的值。The value without surrounding quotes.</returns>
        private static string Unquote(string value)
        {
            if (value.Length >= 2)
            {
                if ((value[0] == '"' && value[value.Length - 1] == '"')
                    || (value[0] == '\'' && value[value.Length - 1] == '\''))
                {
                    return value.Substring(1, value.Length - 2);
                }
            }
            return value;
        }

        /// <summary>
        /// 判断行是否为围栏代码块的起始行并取出信息串（语言标注）。
        /// Determines whether the line opens a fenced code block and extracts the info string (language hint).
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <param name="info">输出的语言标注（去反引号后）。The language hint on output (after the fence marks).</param>
        /// <returns>是否为起始行。Whether it is a fence start.</returns>
        private static bool IsFenceStart(string line, out string info)
        {
            info = string.Empty;
            string trimmed = TrimLeftUpTo3(line);
            if (trimmed.Length < 3)
            {
                return false;
            }
            char c = trimmed[0];
            if (c != '`' && c != '~')
            {
                return false;
            }
            int run = 0;
            for (int i = 0; i < trimmed.Length && trimmed[i] == c; i++)
            {
                run++;
            }
            if (run < 3)
            {
                return false;
            }
            if (c == '`' && trimmed.Substring(run).IndexOf('`') >= 0)
            {
                return false;
            }
            info = trimmed.Substring(run).Trim();
            return true;
        }

        /// <summary>
        /// 解析围栏代码块，逐行生成等宽字体段落，返回块之后的行号。语言标注仅记录不用于高亮。
        /// Parses a fenced code block, emitting one monospaced paragraph per line, and returns the index after the block. The language hint is recorded but not highlighted.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">起始行。The start index.</param>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="info">信息串。The info string.</param>
        /// <returns>块之后的行号。The index following the block.</returns>
        private int ParseFencedCode(string[] lines, int start, Section section, string info)
        {
            string opener = TrimLeftUpTo3(lines[start]);
            char fenceChar = opener[0];
            int fenceLen = 0;
            for (int i = 0; i < opener.Length && opener[i] == fenceChar; i++)
            {
                fenceLen++;
            }
            // 围栏自身的缩进要一并从内容行去掉（CommonMark：正文按起始围栏的缩进对齐），
            // 否则列表项里的代码块会带着两格缩进原样输出。
            // The opening fence's own indent is stripped from the content lines (CommonMark aligns the body with the
            // opener), otherwise a code block inside a list item would keep its two-space indentation verbatim.
            int openIndent = LeadingSpaceCount(lines[start]);

            int i2 = start + 1;
            bool anyLine = false;
            while (i2 < lines.Length)
            {
                string candidate = TrimLeftUpTo3(lines[i2]);
                if (IsFenceClose(candidate, fenceChar, fenceLen))
                {
                    i2++;
                    break;
                }
                AddCodeParagraph(section, StripLeadingSpaces(lines[i2], openIndent));
                anyLine = true;
                i2++;
            }
            // 空围栏照样是一个代码块：模型没有 CodeBlock 类型，补一条空白的等宽段落，
            // 否则相邻的 ``` 与 ``` 会让整块凭空消失。
            // An empty fence is still a code block: the model has no CodeBlock type, so emit one blank monospaced
            // paragraph, otherwise adjacent ``` and ``` lines would make the whole block vanish.
            if (!anyLine)
            {
                AddCodeParagraph(section, string.Empty);
            }
            return i2;
        }

        /// <summary>
        /// 判断已去左空的行是否为对应围栏的闭合行。
        /// Determines whether the left-trimmed line closes the corresponding fence.
        /// </summary>
        /// <param name="trimmed">去左空后的行。The left-trimmed line.</param>
        /// <param name="fenceChar">围栏字符。The fence character.</param>
        /// <param name="fenceLen">围栏长度。The fence length.</param>
        /// <returns>是否为闭合行。Whether it closes the fence.</returns>
        private static bool IsFenceClose(string trimmed, char fenceChar, int fenceLen)
        {
            if (trimmed.Length == 0 || trimmed[0] != fenceChar)
            {
                return false;
            }
            int run = 0;
            for (int i = 0; i < trimmed.Length && trimmed[i] == fenceChar; i++)
            {
                run++;
            }
            if (run < fenceLen)
            {
                return false;
            }
            for (int i = run; i < trimmed.Length; i++)
            {
                if (trimmed[i] != ' ' && trimmed[i] != '\t')
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 解析 4 空格/制表符缩进代码块，去除一层缩进，返回块之后的行号。
        /// Parses an indented (4-space/tab) code block, removing one indent level, and returns the index after the block.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">起始行。The start index.</param>
        /// <param name="section">目标节。The target section.</param>
        /// <returns>块之后的行号。The index following the block.</returns>
        private int ParseIndentedCode(string[] lines, int start, Section section)
        {
            int i = start;
            while (i < lines.Length)
            {
                if (lines[i].Trim().Length == 0)
                {
                    // 允许代码块内的空行；若后续不再缩进则结束。
                    // Allow blank lines inside; end when the next non-blank line is not indented.
                    int peek = i;
                    while (peek < lines.Length && lines[peek].Trim().Length == 0)
                    {
                        peek++;
                    }
                    if (peek < lines.Length && IsIndentedCode(lines[peek]))
                    {
                        for (int b = i; b < peek; b++)
                        {
                            AddCodeParagraph(section, string.Empty);
                        }
                        i = peek;
                        continue;
                    }
                    break;
                }
                if (!IsIndentedCode(lines[i]))
                {
                    break;
                }
                AddCodeParagraph(section, DedentCode(lines[i]));
                i++;
            }
            return i;
        }

        /// <summary>
        /// 判断行是否有 4 个空格或制表符缩进（作为缩进代码块）。
        /// Determines whether the line is indented by four spaces or a tab (an indented code block).
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <returns>是否为缩进代码。Whether it is indented code.</returns>
        private static bool IsIndentedCode(string line)
        {
            if (line.Length >= 4 && line[0] == ' ' && line[1] == ' ' && line[2] == ' ' && line[3] == ' ')
            {
                return true;
            }
            return line.Length >= 1 && line[0] == '\t';
        }

        /// <summary>
        /// 去除缩进代码行的一层前导空白（4 空格或 1 制表符）。
        /// Removes one indent level (four spaces or one tab) from an indented code line.
        /// </summary>
        /// <param name="line">原始行。The raw line.</param>
        /// <returns>去缩进后的行。The dedented line.</returns>
        private static string DedentCode(string line)
        {
            if (line.Length >= 1 && line[0] == '\t')
            {
                return line.Substring(1);
            }
            if (line.Length >= 4 && line[0] == ' ' && line[1] == ' ' && line[2] == ' ' && line[3] == ' ')
            {
                return line.Substring(4);
            }
            return line.TrimStart(' ');
        }

        /// <summary>行首连续空格数（制表符按 4 格折算，与缩进判定一致）。The count of leading spaces, with a tab worth four to match the indent tests.</summary>
        /// <param name="line">原始行。The raw line.</param>
        /// <returns>前导缩进格数。The leading indent width.</returns>
        private static int LeadingSpaceCount(string line)
        {
            int n = 0;
            while (n < line.Length)
            {
                if (line[n] == ' ')
                {
                    n++;
                    continue;
                }
                if (line[n] == '\t')
                {
                    n += 4;
                    continue;
                }
                break;
            }
            return n;
        }

        /// <summary>
        /// 从行首最多去掉 count 个空白（制表符整块去掉），用于还原围栏代码块的相对缩进。
        /// Removes at most &lt;count&gt; leading spaces (a tab is removed whole), restoring a fenced block's relative indent.
        /// </summary>
        /// <param name="line">原始行。The raw line.</param>
        /// <param name="count">最多去掉的格数。The maximum number of columns to remove.</param>
        /// <returns>去缩进后的行。The dedented line.</returns>
        private static string StripLeadingSpaces(string line, int count)
        {
            if (count <= 0 || line.Length == 0)
            {
                return line;
            }
            int n = 0;
            while (n < line.Length && n < count)
            {
                if (line[n] == ' ')
                {
                    n++;
                    continue;
                }
                if (line[n] == '\t')
                {
                    n++;
                    break;
                }
                break;
            }
            return line.Substring(n);
        }

        /// <summary>
        /// 预扫一遍收集链接引用定义（<c>[label]: url "title"</c>），并标记被定义占用的行。
        /// 只有块起始位置（空行或另一条定义之后）的行才算定义，缩进不超过 3 格，围栏代码块内部的行一律跳过
        /// ——CommonMark 规定定义不能打断段落，也不存在于代码块里。
        /// Pre-scans the lines for link reference definitions (<c>[label]: url "title"</c>) and marks the lines they
        /// occupy. Only a line at a block start (after a blank line or another definition), indented by at most three
        /// columns, counts as a definition, and lines inside a fenced block are skipped outright - CommonMark forbids a
        /// definition from interrupting a paragraph and never recognizes one inside code.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">正文起始行。The first body line.</param>
        /// <param name="reset">是否清空已有链接表；仅顶层为 true，引用块内部的定义做增量合并。Whether to clear the table; only the top level does, so definitions inside a block quote merge in.</param>
        /// <returns>与 <paramref name="lines"/> 等长的定义行标记。Markers, one per line of <paramref name="lines"/>.</returns>
        private bool[] CollectLinkDefinitions(string[] lines, int start, bool reset)
        {
            if (reset)
            {
                _definitions.Clear();
            }
            bool[] marked = new bool[lines.Length];
            bool blockStart = true;
            for (int i = start; i < lines.Length; i++)
            {
                string line = lines[i];
                if (line.Trim().Length == 0)
                {
                    blockStart = true;
                    continue;
                }
                string fenceInfo;
                if (IsFenceStart(line, out fenceInfo))
                {
                    i = FindFenceEnd(lines, i);
                    blockStart = false;
                    continue;
                }
                if (!blockStart)
                {
                    continue;
                }
                string label;
                string destination;
                string title;
                if (TryParseLinkDefinition(line, out label, out destination, out title))
                {
                    if (_definitions.Count < MaxLinkDefinitions && !_definitions.ContainsKey(label))
                    {
                        LinkDefinition definition = new LinkDefinition();
                        definition.Destination = destination;
                        definition.Title = title;
                        _definitions[label] = definition;
                    }
                    // 上限之外的行同样不渲染：识别不出来的定义会退化成段落文本，比凭空多出一张表更保守。
                    // Lines beyond the cap are still swallowed: an unrecognized definition degrades into paragraph
                    // text, which is worse than a table that silently never appears, so the line stays consumed.
                    marked[i] = true;
                    continue;
                }
                blockStart = false;
            }
            return marked;
        }

        /// <summary>
        /// 解析一行链接引用定义：缩进 ≤3、形如 <c>[label]:</c>、目标为尖括号或无空白串、可选标题为成对引号或括号。
        /// Parses one definition line: indent of at most three, the form <c>[label]:</c>, a destination in angle
        /// brackets or a whitespace-free run, and an optional title in matching quotes or parentheses.
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <param name="label">输出的归一化标签。The normalized label on output.</param>
        /// <param name="destination">输出的链接目标。The destination on output.</param>
        /// <param name="title">输出的可选标题，缺省为 null。The optional title, null when absent.</param>
        /// <returns>是否为合法定义。Whether the line is a valid definition.</returns>
        private static bool TryParseLinkDefinition(string line, out string label, out string destination, out string title)
        {
            label = null;
            destination = null;
            title = null;
            if (LeadingSpaceCount(line) > 3)
            {
                return false;
            }
            int i = 0;
            while (i < line.Length && (line[i] == ' ' || line[i] == '\t'))
            {
                i++;
            }
            if (i >= line.Length || line[i] != '[')
            {
                return false;
            }
            int close = IndexUnescaped(line, ']', i + 1);
            if (close < 0 || close - i - 1 <= 0 || close - i - 1 > MaxDefinitionLabelChars)
            {
                return false;
            }
            string rawLabel = NormalizeLabel(line.Substring(i + 1, close - i - 1));
            if (rawLabel.Length == 0 || rawLabel.IndexOf('[') >= 0)
            {
                return false;
            }
            i = close + 1;
            if (i >= line.Length || line[i] != ':')
            {
                return false;
            }
            i++;
            while (i < line.Length && (line[i] == ' ' || line[i] == '\t'))
            {
                i++;
            }
            if (i >= line.Length)
            {
                return false;
            }

            int rest;
            string dest = ReadDefinitionDestination(line, i, out rest);
            if (dest == null)
            {
                return false;
            }
            i = rest;
            while (i < line.Length && (line[i] == ' ' || line[i] == '\t'))
            {
                i++;
            }
            if (i < line.Length)
            {
                string inlineTitle;
                int titleEnd;
                if (!ReadDefinitionTitle(line, i, out inlineTitle, out titleEnd))
                {
                    return false;
                }
                title = inlineTitle;
                i = titleEnd;
                while (i < line.Length && (line[i] == ' ' || line[i] == '\t'))
                {
                    i++;
                }
                if (i < line.Length)
                {
                    return false;
                }
            }
            label = rawLabel;
            destination = dest;
            return true;
        }

        /// <summary>
        /// 读取定义里的链接目标：尖括号形式可含空白与转义尖括号，普通形式遇到空白即止。
        /// Reads a definition's destination: the angle-bracket form may contain spaces and escaped brackets, the bare
        /// form ends at the first whitespace.
        /// </summary>
        /// <param name="line">原始行。The raw line.</param>
        /// <param name="start">目标起始位置。The start index.</param>
        /// <param name="end">目标结束后的位置。The index just after the destination.</param>
        /// <returns>链接目标，非法时为 null。The destination, or null when invalid.</returns>
        private static string ReadDefinitionDestination(string line, int start, out int end)
        {
            end = start;
            if (line[start] == '<')
            {
                StringBuilder sb = new StringBuilder();
                for (int i = start + 1; i < line.Length; i++)
                {
                    char c = line[i];
                    if (c == '\n')
                    {
                        return null;
                    }
                    if (c == '\\' && i + 1 < line.Length && IsEscapable(line[i + 1]))
                    {
                        sb.Append(line[i + 1]);
                        i++;
                        continue;
                    }
                    if (c == '>')
                    {
                        end = i + 1;
                        return sb.ToString();
                    }
                    sb.Append(c);
                }
                return null;
            }
            int j = start;
            while (j < line.Length && !char.IsWhiteSpace(line[j]))
            {
                if (line[j] == '\\' && j + 1 < line.Length)
                {
                    j += 2;
                    continue;
                }
                if (line[j] == '<')
                {
                    return null;
                }
                j++;
            }
            if (j == start)
            {
                return null;
            }
            end = j;
            return UnescapeBackslashes(line.Substring(start, j - start));
        }

        /// <summary>
        /// 读取定义的可选标题：双引号、单引号或圆括号成对包裹，必须延伸到行尾。
        /// Reads an optional title: matched double quotes, single quotes or parentheses running to the end of the line.
        /// </summary>
        /// <param name="line">原始行。The raw line.</param>
        /// <param name="start">标题起始位置。The start index.</param>
        /// <param name="value">输出的标题文本。The title text on output.</param>
        /// <param name="end">标题结束后的位置。The index just after the title.</param>
        /// <returns>是否为合法标题。Whether a valid title was read.</returns>
        private static bool ReadDefinitionTitle(string line, int start, out string value, out int end)
        {
            value = null;
            end = start;
            char open = line[start];
            char closeChar;
            if (open == '"')
            {
                closeChar = '"';
            }
            else if (open == '\'')
            {
                closeChar = '\'';
            }
            else if (open == '(')
            {
                closeChar = ')';
            }
            else
            {
                return false;
            }
            int last = line.Length - 1;
            if (last <= start || line[last] != closeChar)
            {
                return false;
            }
            value = UnescapeBackslashes(line.Substring(start + 1, last - start - 1));
            end = last + 1;
            return true;
        }

        /// <summary>
        /// 反斜杠转义还原：把 <c>\&gt;</c> 这类可转义字符还原成字符本身，其余反斜杠原样保留。
        /// Resolves backslash escapes: an escapable character following a backslash becomes that character, any other
        /// backslash stays literal.
        /// </summary>
        /// <param name="text">原始串。The raw string.</param>
        /// <returns>还原后的串。The resolved string.</returns>
        private static string UnescapeBackslashes(string text)
        {
            if (text.IndexOf('\\') < 0)
            {
                return text;
            }
            StringBuilder sb = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\\' && i + 1 < text.Length && IsEscapable(text[i + 1]))
                {
                    sb.Append(text[i + 1]);
                    i++;
                    continue;
                }
                sb.Append(text[i]);
            }
            return sb.ToString();
        }

        /// <summary>
        /// 归一化引用标签：连续空白并成一格、去首尾空白、转小写（CommonMark 的标签大小写不敏感）。
        /// Normalizes a reference label: runs of whitespace collapse to one space, the ends are trimmed and the case
        /// is folded, since CommonMark labels are case-insensitive.
        /// </summary>
        /// <param name="raw">原始标签。The raw label.</param>
        /// <returns>归一化标签。The normalized label.</returns>
        private static string NormalizeLabel(string raw)
        {
            StringBuilder sb = new StringBuilder(raw.Length);
            bool pendingSpace = false;
            for (int i = 0; i < raw.Length; i++)
            {
                if (char.IsWhiteSpace(raw[i]))
                {
                    pendingSpace = sb.Length > 0;
                    continue;
                }
                if (pendingSpace)
                {
                    sb.Append(' ');
                    pendingSpace = false;
                }
                sb.Append(char.ToLowerInvariant(raw[i]));
            }
            return sb.ToString();
        }

        /// <summary>查找未被反斜杠转义的目标字符位置，找不到返回 -1。Finds the next unescaped occurrence of a character, or -1 when absent.</summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="target">目标字符。The character to find.</param>
        /// <param name="from">起始位置。The start index.</param>
        /// <returns>目标位置或 -1。The found index or -1.</returns>
        private static int IndexUnescaped(string text, char target, int from)
        {
            for (int i = from; i < text.Length; i++)
            {
                if (text[i] == '\\')
                {
                    i++;
                    continue;
                }
                if (text[i] == target)
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// 返回围栏块闭合行的下标；未闭合时返回末行下标。
        /// Returns the index of a fence's closing line, or the last line index when the fence is never closed.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">起始围栏行。The opening fence line.</param>
        /// <returns>闭合行下标。The closing line index.</returns>
        private static int FindFenceEnd(string[] lines, int start)
        {
            string opener = TrimLeftUpTo3(lines[start]);
            char fenceChar = opener[0];
            int fenceLen = 0;
            for (int i = 0; i < opener.Length && opener[i] == fenceChar; i++)
            {
                fenceLen++;
            }
            for (int i = start + 1; i < lines.Length; i++)
            {
                if (IsFenceClose(TrimLeftUpTo3(lines[i]), fenceChar, fenceLen))
                {
                    return i;
                }
            }
            return lines.Length - 1;
        }

        /// <summary>
        /// 追加一个等宽字体代码段落。刻意不设背景色（模型仅表格单元格支持底纹，代码段落无背景属性）。
        /// Appends a monospaced code paragraph. No background is set deliberately (the model supports shading only on table cells; paragraphs have no background property).
        /// </summary>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="text">代码文本。The code text.</param>
        private void AddCodeParagraph(Section section, string text)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.StyleName = CodeStyleName;
            paragraph.LeftIndent = 12;
            paragraph.LineSpacing = 1.0;
            TextFormat format = NewCodeFormat();
            // 空行也保留为空段落，确保代码块行数与原文一致。
            // Blank lines are kept as empty paragraphs so the code-block line count matches the source.
            paragraph.Inlines.Add(new TextRun(text ?? string.Empty, format));
            section.Blocks.Add(paragraph);
        }

        /// <summary>
        /// 判断行是否为 ATX 标题。
        /// Determines whether the line is an ATX heading.
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <param name="level">输出的标题级别 1~6。The heading level 1-6 on output.</param>
        /// <param name="text">输出的标题文本。The heading text on output.</param>
        /// <returns>是否为 ATX 标题。Whether it is an ATX heading.</returns>
        private static bool IsAtxHeading(string line, out int level, out string text)
        {
            level = 0;
            text = string.Empty;
            string trimmed = TrimLeftUpTo3(line);
            if (trimmed.Length == 0 || trimmed[0] != '#')
            {
                return false;
            }
            int hashes = 0;
            for (int i = 0; i < trimmed.Length && trimmed[i] == '#'; i++)
            {
                hashes++;
            }
            if (hashes > 6)
            {
                return false;
            }
            if (hashes == trimmed.Length)
            {
                level = hashes;
                return true;
            }
            char after = trimmed[hashes];
            if (after != ' ' && after != '\t')
            {
                return false;
            }
            level = hashes;
            string rest = trimmed.Substring(hashes).Trim();
            // 去除尾部的 ### 关闭序列。
            // Strip a trailing ### closing sequence.
            rest = rest.TrimEnd('#', ' ');
            text = rest;
            return true;
        }

        /// <summary>
        /// 追加一个标题段落，套用递增字号的标题格式。
        /// Appends a heading paragraph using the size-increasing heading format.
        /// </summary>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="level">标题级别 1~6。The heading level 1-6.</param>
        /// <param name="text">标题文本（含行内语法）。The heading text (may contain inline syntax).</param>
        private void AddHeading(Section section, int level, string text)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.HeadingLevel = level;
            paragraph.SpaceBefore = Math.Max(0, this.options.BaseFontSize - 6);
            paragraph.SpaceAfter = 4;
            paragraph.KeepWithNext = true;
            TextFormat heading = TextFormat.Heading(level);
            if (!string.IsNullOrEmpty(this.options.BodyFontFamily))
            {
                heading.FontFamily = this.options.BodyFontFamily;
            }
            AddInlinesFromText(paragraph, text, heading);
            section.Blocks.Add(paragraph);
        }

        /// <summary>
        /// 判断行是否为 Setext 下划线（仅 = 或仅 -）。
        /// Determines whether the line is a Setext underline (only '=' or only '-').
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <param name="level">输出级别：=为1、-为2。The level on output: '=' is 1, '-' is 2.</param>
        /// <returns>是否为 Setext 下划线。Whether it is a Setext underline.</returns>
        private static bool IsSetextUnderline(string line, out int level)
        {
            level = 0;
            string trimmed = line.Trim();
            if (trimmed.Length == 0)
            {
                return false;
            }
            bool allEq = true;
            bool allDash = true;
            for (int i = 0; i < trimmed.Length; i++)
            {
                if (trimmed[i] != '=')
                {
                    allEq = false;
                }
                if (trimmed[i] != '-')
                {
                    allDash = false;
                }
            }
            if (allEq)
            {
                level = 1;
                return true;
            }
            if (allDash)
            {
                level = 2;
                return true;
            }
            return false;
        }

        /// <summary>
        /// 解析段落（含行尾两空格硬换行、Setext 标题合并），返回块之后的行号。
        /// Parses a paragraph (hard breaks from two trailing spaces, Setext-heading merging) and returns the index after the block.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">起始行。The start index.</param>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="document">目标文档（用于一级标题写元数据）。The target document (level-1 titles also set metadata).</param>
        /// <returns>块之后的行号。The index following the block.</returns>
        private int ParseParagraph(string[] lines, int start, Section section, OfficeDocument document)
        {
            List<string> buffer = new List<string>();
            buffer.Add(lines[start]);
            int i = start + 1;

            while (i < lines.Length)
            {
                string next = lines[i];
                if (next.Trim().Length == 0)
                {
                    break;
                }

                int setextLevel;
                if (IsSetextUnderline(next, out setextLevel))
                {
                    // CommonMark 把下划线之前的整段文字作为标题正文，而不只是紧邻下划线的那一行；
                    // 行间软换行在这里折成空格，Word 与 HTML 都不需要保留换行符。
                    // CommonMark makes the whole paragraph above the underline the heading content, not just the
                    // adjacent line; the soft breaks between those lines fold into spaces here.
                    StringBuilder headingText = new StringBuilder();
                    for (int k = 0; k < buffer.Count; k++)
                    {
                        if (k > 0)
                        {
                            headingText.Append(' ');
                        }
                        headingText.Append(StripHardBreakMarker(buffer[k]).Trim());
                    }
                    string title = headingText.ToString();
                    AddHeading(section, setextLevel, title);
                    if (setextLevel == 1 && string.IsNullOrEmpty(document.Metadata.Title))
                    {
                        document.Metadata.Title = PlainFromInline(title);
                    }
                    return i + 1;
                }

                if (IsParagraphStop(next, lines, i))
                {
                    break;
                }

                buffer.Add(next);
                i++;
            }

            FlushParagraph(section, buffer);
            return i;
        }

        /// <summary>
        /// 判断某行是否会中断段落（成为新的块级起点）。
        /// Determines whether a line interrupts the paragraph (starts a new block).
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="index">当前行号。The current index.</param>
        /// <returns>是否中断段落。Whether it interrupts the paragraph.</returns>
        private bool IsParagraphStop(string line, string[] lines, int index)
        {
            string info;
            if (IsFenceStart(line, out info))
            {
                return true;
            }
            int level;
            string text;
            if (IsAtxHeading(line, out level, out text))
            {
                return true;
            }
            if (StartsQuote(line))
            {
                return true;
            }
            // 分隔线要先于列表判定，否则 "* * *" 会被当作列表项而中断段落的方式不对。
            // Test the thematic break before the list, otherwise "* * *" interrupts the paragraph as list items.
            if (IsThematicBreak(line))
            {
                return true;
            }
            if (IsListItem(line))
            {
                return true;
            }
            if (options.EnableTables && IsTableHeader(lines, index))
            {
                return true;
            }
            // 缩进代码不能中断段落：CommonMark 把段落后的 4 空格行视为同一段的续行。
            // Indented code cannot interrupt a paragraph: CommonMark keeps a 4-space line after a paragraph as a
            // continuation line of that paragraph.
            return false;
        }

        /// <summary>
        /// 将缓冲的多行输出为一个段落，行间按硬换行或软空格连接。
        /// Emits the buffered lines as one paragraph, joining lines with a hard break or a soft space.
        /// </summary>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="buffer">缓冲行列表。The buffered lines.</param>
        private void FlushParagraph(Section section, List<string> buffer)
        {
            if (buffer.Count == 0)
            {
                return;
            }

            // 以 < 开头的行过去在这里绕过行内解析、原样留成文本，结果把同一段里的自动链接
            // 与强调一起吞掉了。写器对文本一律转义（原始 HTML 永不透传），这里就按普通段落走。
            // Lines starting with '<' used to bypass inline parsing here and were kept as bare text, which also
            // swallowed the autolinks and emphasis in the same run. The writers escape all text unconditionally
            // (raw HTML is never passed through), so treat these as ordinary paragraphs.
            Paragraph paragraph = new Paragraph();
            for (int k = 0; k < buffer.Count; k++)
            {
                if (k > 0)
                {
                    bool prevHard = IsHardBreak(buffer[k - 1]);
                    if (prevHard || options.HardLineBreaks)
                    {
                        paragraph.AddLineBreak();
                    }
                    else
                    {
                        paragraph.AddText(" ", NewBodyFormat());
                    }
                }

                string content = StripHardBreakMarker(buffer[k]);
                AddInlinesFromText(paragraph, content, NewBodyFormat());
            }
            section.Blocks.Add(paragraph);
        }

        /// <summary>
        /// 判断行尾是否为硬换行（两个及以上尾随空格，或结尾反斜杠）。
        /// Determines whether the line ends in a hard break (two or more trailing spaces, or a trailing backslash).
        /// </summary>
        /// <param name="line">原始行。The raw line.</param>
        /// <returns>是否硬换行。Whether it is a hard break.</returns>
        private static bool IsHardBreak(string line)
        {
            string right = line.TrimEnd();
            if (right.Length != line.Length && line.Length - right.Length >= 2)
            {
                return true;
            }
            return right.EndsWith("\\", StringComparison.Ordinal);
        }

        /// <summary>
        /// 去除行尾的硬换行标记（尾随空格或结尾反斜杠）以便行内解析。
        /// Strips the trailing hard-break marker (trailing spaces or trailing backslash) before inline parsing.
        /// </summary>
        /// <param name="line">原始行。The raw line.</param>
        /// <returns>清理后的行内容。The cleaned line content.</returns>
        private string StripHardBreakMarker(string line)
        {
            string right = line.TrimEnd();
            if (right.EndsWith("\\", StringComparison.Ordinal))
            {
                right = right.Substring(0, right.Length - 1);
            }
            return options.TrimTrailingWhitespace ? right.Trim() : right;
        }

        /// <summary>
        /// 判断行是否为水平分隔线（≥3 个连续 -、* 或 _，可含空格）。
        /// Determines whether the line is a thematic break (three or more consecutive -, *, or _, spaces allowed).
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <returns>是否为分隔线。Whether it is a thematic break.</returns>
        private static bool IsThematicBreak(string line)
        {
            string trimmed = TrimLeftUpTo3(line).Trim();
            if (trimmed.Length < 3)
            {
                return false;
            }
            char kind = trimmed[0];
            if (kind != '-' && kind != '*' && kind != '_')
            {
                return false;
            }
            int count = 0;
            for (int i = 0; i < trimmed.Length; i++)
            {
                char c = trimmed[i];
                if (c == kind)
                {
                    count++;
                }
                else if (c == ' ' || c == '\t' || c == '\r')
                {
                    continue;
                }
                else
                {
                    return false;
                }
            }
            return count >= 3;
        }

        /// <summary>
        /// 追加一条水平分隔线。由于模型的段落没有下边框属性，这里以最不过分的方式表达：
        /// 一个居中、承载 "＊ ＊ ＊" 的段落并设置段前/段后间距。
        /// Appends a thematic break. Since a model paragraph has no bottom-border property, this uses the least-surprising representation:
        /// a centered paragraph carrying "＊ ＊ ＊" with space before/after.
        /// </summary>
        /// <param name="section">目标节。The target section.</param>
        private void AddThematicBreak(Section section)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.Alignment = ParagraphAlignment.Center;
            paragraph.SpaceBefore = 6;
            paragraph.SpaceAfter = 6;
            TextFormat format = NewBodyFormat();
            paragraph.Inlines.Add(new TextRun(ThematicBreakText, format));
            section.Blocks.Add(paragraph);
        }

        /// <summary>
        /// 判断行是否为引用块（去掉≤3前导空格后以 &gt; 开头）。
        /// Determines whether the line starts a block quote (begins with '&gt;' after up to three leading spaces).
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <returns>是否为引用行。Whether it is a quote line.</returns>
        private static bool StartsQuote(string line)
        {
            string trimmed = TrimLeftUpTo3(line);
            return trimmed.Length > 0 && trimmed[0] == '>';
        }

        /// <summary>
        /// 解析引用块：连续的 &gt; 行（含 CommonMark 的懒惰续行）先各剥掉一层标记攒成缓冲区，
        /// 再当作普通块递归解析，最后统一补缩进与斜体。逐行成段会把引用里的列表、代码块、标题压成字面文本；
        /// 嵌套的 &gt;&gt; 不需要特判，内层行仍以 &gt; 开头，交给递归自然多缩一层。
        /// Parses a block quote: consecutive '&gt;' lines, plus CommonMark lazy continuations, each lose one marker
        /// level into a buffer that is parsed recursively as ordinary blocks; the resulting paragraphs then gain the
        /// indent and italics. Turning every line into its own paragraph used to flatten lists, code and headings
        /// inside a quote into literal text, and nested '&gt;&gt;' needs no special case - those lines still begin
        /// with '&gt;' and the recursion adds another level of indent.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">起始行。The start index.</param>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="document">所属文档。The owning document.</param>
        /// <returns>块之后的行号。The index following the block.</returns>
        private int ParseQuote(string[] lines, int start, Section section, OfficeDocument document)
        {
            List<string> buffer = new List<string>();
            int i = start;
            while (i < lines.Length)
            {
                string line = lines[i];
                if (StartsQuote(line))
                {
                    buffer.Add(RemoveOneQuoteMarker(TrimLeftUpTo3(line)));
                    i++;
                    continue;
                }
                // 懒惰续行：引用里未闭合的段落可以直接吃下后面这个普通行。
                // Lazy continuation: an open paragraph inside the quote swallows the next plain line.
                if (buffer.Count > 0 && IsLazyQuoteLine(line, buffer[buffer.Count - 1]))
                {
                    buffer.Add(line);
                    i++;
                    continue;
                }
                break;
            }

            string[] inner = buffer.ToArray();
            if (_quoteDepth >= MaxQuoteDepth)
            {
                // 超深嵌套不再递归，逐行按字面文本收成斜体段落，引用标记本身留在文本里。
                // Beyond the ceiling the buffer stops recursing: each line becomes an italic paragraph and the
                // remaining markers stay as literal text.
                AddQuoteParagraphs(section, inner, 0);
                return i;
            }

            Section quoted = new Section();
            bool[] definitions = CollectLinkDefinitions(inner, 0, false);
            _quoteDepth++;
            try
            {
                ParseBlocks(inner, 0, quoted, document, definitions);
            }
            finally
            {
                _quoteDepth--;
            }

            for (int b = 0; b < quoted.Blocks.Count; b++)
            {
                Paragraph paragraph = quoted.Blocks[b] as Paragraph;
                if (paragraph == null)
                {
                    // 表格没有可供引用表达的属性，原样接过来。Tables carry no place to express a quote, so they move over as built.
                    section.Blocks.Add(quoted.Blocks[b]);
                    continue;
                }
                // 列表项的缩进就是层级本身，加一格会让 HTML 写器把整项算成更深一层。
                // A list item's indent is its nesting level, so adding a column would make the writers read the item one level too deep.
                if (string.IsNullOrEmpty(paragraph.ListMarker))
                {
                    paragraph.LeftIndent += QuoteIndent;
                    if (paragraph.HeadingLevel == 0)
                    {
                        paragraph.SpaceBefore = 2;
                        paragraph.SpaceAfter = 2;
                    }
                }
                ApplyQuoteItalic(paragraph);
                section.Blocks.Add(paragraph);
            }
            return i;
        }

        /// <summary>
        /// 把缓冲区行逐行收成斜体段落，供超过嵌套上限的引用兜底。Collapses the buffer line by line into italic paragraphs, the fallback past the nesting ceiling.
        /// </summary>
        /// <param name="section">目标节。The target section.</param>
        /// <param name="inner">已剥掉一层标记的行。The lines with one marker level stripped.</param>
        /// <param name="extraIndent">本层之外还要补上的缩进。The indent to add on top of this level.</param>
        private void AddQuoteParagraphs(Section section, string[] inner, double extraIndent)
        {
            for (int k = 0; k < inner.Length; k++)
            {
                if (inner[k].Trim().Length == 0)
                {
                    continue;
                }
                Paragraph paragraph = new Paragraph();
                paragraph.LeftIndent = extraIndent + QuoteIndent;
                paragraph.SpaceBefore = 2;
                paragraph.SpaceAfter = 2;
                TextFormat format = NewBodyFormat();
                format.Italic = true;
                AddInlinesFromText(paragraph, inner[k].Trim(), format);
                section.Blocks.Add(paragraph);
            }
        }

        /// <summary>
        /// 给引用里的段落补斜体；等宽运行（代码块与行内代码）保持原样，斜体只让代码更难认。
        /// Italicises the runs of a quoted paragraph; monospaced runs (code blocks and spans) stay as they are,
        /// since italics only make code harder to read.
        /// </summary>
        /// <param name="paragraph">待处理段落。The paragraph to process.</param>
        private void ApplyQuoteItalic(Paragraph paragraph)
        {
            bool isCodeBlock = string.Equals(paragraph.StyleName, CodeStyleName, StringComparison.Ordinal);
            for (int r = 0; r < paragraph.Inlines.Count; r++)
            {
                TextRun run = paragraph.Inlines[r] as TextRun;
                if (run == null || run.Format == null)
                {
                    continue;
                }
                string family = run.Format.GetAsciiFontFamily();
                if (isCodeBlock || IsCodeFamily(family))
                {
                    continue;
                }
                run.Format.Italic = true;
            }
        }

        /// <summary>字体名是否代码字体：解析器自己的代码族，外加 Consolas/Courier 前缀。Whether a family is the code font: the parser's own code family plus the Consolas/Courier prefixes.</summary>
        /// <param name="family">字体名。The font family.</param>
        /// <returns>是否等宽代码字体。Whether it is a monospaced code font.</returns>
        private bool IsCodeFamily(string family)
        {
            if (string.IsNullOrEmpty(family))
            {
                return false;
            }
            string code = string.IsNullOrEmpty(options.CodeFontFamily) ? "Consolas" : options.CodeFontFamily;
            return family.Equals(code, StringComparison.OrdinalIgnoreCase)
                || family.StartsWith("Consolas", StringComparison.OrdinalIgnoreCase)
                || family.StartsWith("Courier", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 去掉一行最前面的一层 &gt; 标记（连同其后恰好一个空白）。Removes the leading '&gt;' marker of a line, together with the single space that may follow it.
        /// </summary>
        /// <param name="content">已去掉≤3 前导空白的行。The line with its up-to-three leading spaces already trimmed.</param>
        /// <returns>剥掉一层标记的行。The line with one marker level removed.</returns>
        private static string RemoveOneQuoteMarker(string content)
        {
            int pos = content.Length > 0 && content[0] == '>' ? 1 : 0;
            if (pos < content.Length && (content[pos] == ' ' || content[pos] == '\t'))
            {
                pos++;
            }
            return content.Substring(pos).TrimEnd('\r', ' ', '\t');
        }

        /// <summary>
        /// 判断一个非引用行能否作为懒惰续行接进上方未闭合的引用段落：空行、围栏、ATX 标题、分隔线、缩进代码都不能。
        /// Decides whether a non-quote line may join the open quote paragraph above it as a lazy continuation: blank
        /// lines, fences, ATX headings, thematic breaks and indented code may not.
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <param name="previous">缓冲区里的上一行。The previous buffered line.</param>
        /// <returns>是否接续引用段落。Whether it continues the quote.</returns>
        private static bool IsLazyQuoteLine(string line, string previous)
        {
            if (line.Trim().Length == 0 || StartsQuote(line))
            {
                return false;
            }
            if (IsIndentedCode(line) || IsThematicBreak(line))
            {
                return false;
            }
            int level;
            string text;
            if (IsAtxHeading(line, out level, out text))
            {
                return false;
            }
            string fenceInfo;
            if (IsFenceStart(line, out fenceInfo))
            {
                return false;
            }
            // 上一行本身若是围栏、缩进代码或空行，引用段落早已闭合，普通行只是新块。
            // A fence, indented code or blank previous line means no paragraph is open, so the plain line starts a block.
            if (previous.Trim().Length == 0 || IsIndentedCode(previous))
            {
                return false;
            }
            string previousFence;
            return !IsFenceStart(previous, out previousFence);
        }

        /// <summary>
        /// 判断行是否为列表项（无序 -/*/+ 或有序 数字. / 数字)）。
        /// Determines whether the line is a list item (unordered -/*/+ or ordered digits. / digits)).
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <returns>是否为列表项。Whether it is a list item.</returns>
        private static bool IsListItem(string line)
        {
            if (line.Trim().Length == 0)
            {
                return false;
            }
            string rest = line.Substring(Math.Min(LeadingCharIndex(line), line.Length));

            if (rest.Length >= 1 && (rest[0] == '-' || rest[0] == '*' || rest[0] == '+'))
            {
                // 单个 "-" 是合法的空列表项；带正文时必须有一个空格或制表符分隔。
                // A lone "-" is a legal empty list item; with a body the marker needs a space or tab separator.
                if (rest.Length == 1)
                {
                    return true;
                }
                return rest[1] == ' ' || rest[1] == '\t';
            }

            int d = 0;
            while (d < rest.Length && rest[d] >= '0' && rest[d] <= '9')
            {
                d++;
            }
            if (d > 0 && d < rest.Length && (rest[d] == '.' || rest[d] == ')'))
            {
                int after = d + 1;
                if (after >= rest.Length || rest[after] == ' ' || rest[after] == '\t')
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 解析列表块（连续列表项、支持缩进嵌套与任务列表），返回块之后的行号。
        /// Parses a list block (consecutive items, indentation nesting and task lists) and returns the index after the block.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">起始行。The start index.</param>
        /// <param name="section">目标节。The target section.</param>
        /// <returns>块之后的行号。The index following the block.</returns>
        private int ParseList(string[] lines, int start, Section section)
        {
            List<int> indentStack = new List<int>();
            int i = start;
            Paragraph lastItem = null;
            int lastItemIndent = -1;

            while (i < lines.Length)
            {
                if (lines[i].Trim().Length == 0)
                {
                    int peek = i;
                    while (peek < lines.Length && lines[peek].Trim().Length == 0)
                    {
                        peek++;
                    }
                    if (peek >= lines.Length)
                    {
                        break;
                    }
                    if (IsListItem(lines[peek]))
                    {
                        i = peek;
                        continue;
                    }
                    // 空行之后仍比本项更深缩进且不是新块的行，是本项的第二段（松散列表项）；
                    // 模型没有列表容器，故折回同一段落文本。
                    // A deeper-indented non-block line after a blank line is this item's second paragraph (a loose
                    // item); the model has no list container, so it folds back into the same paragraph text.
                    if (lastItem != null && CountLeadingSpaces(lines[peek]) > lastItemIndent
                        && !IsParagraphStop(lines[peek], lines, peek))
                    {
                        AddInlinesFromText(lastItem, " " + lines[peek].Trim(), NewBodyFormat());
                        i = peek + 1;
                        continue;
                    }
                    break;
                }

                if (!IsListItem(lines[i]))
                {
                    // 缩进续行与懒延续（CommonMark 都把它留在本项内）；新块起点则结束列表。
                    // An indented continuation line and a lazy continuation both stay inside the item (as in
                    // CommonMark); any new block start ends the list.
                    if (lastItem != null && !IsParagraphStop(lines[i], lines, i))
                    {
                        AddInlinesFromText(lastItem, " " + lines[i].Trim(), NewBodyFormat());
                        i++;
                        continue;
                    }
                    break;
                }

                int indent = CountLeadingSpaces(lines[i]);
                while (indentStack.Count > 0 && indent < indentStack[indentStack.Count - 1])
                {
                    indentStack.RemoveAt(indentStack.Count - 1);
                }
                if (indentStack.Count == 0 || indent > indentStack[indentStack.Count - 1])
                {
                    indentStack.Add(indent);
                }
                int level = indentStack.Count - 1;

                string rest = lines[i].Substring(Math.Min(LeadingCharIndex(lines[i]), lines[i].Length));
                bool ordered;
                string marker;
                string content = SplitListMarker(rest, out ordered, out marker);

                Paragraph paragraph = new Paragraph();
                paragraph.LeftIndent = 12 + level * IndentPerLevel;
                paragraph.SpaceBefore = 1;
                paragraph.SpaceAfter = 1;

                string bullet = ordered ? marker : BulletForLevel(level);
                if (options.EnableTaskLists)
                {
                    string checkbox = ExtractTaskCheckbox(ref content);
                    if (checkbox != null)
                    {
                        bullet = checkbox + " " + bullet;
                    }
                }
                paragraph.ListMarker = bullet;

                TextFormat format = NewBodyFormat();
                AddInlinesFromText(paragraph, content.Trim(), format);
                section.Blocks.Add(paragraph);
                lastItem = paragraph;
                lastItemIndent = indent;
                i++;
            }
            return i;
        }

        /// <summary>
        /// 拆分列表标记，返回项目正文；输出是否为有序与标记文本。
        /// Splits the list marker, returning the item body and reporting whether it is ordered plus the marker text.
        /// </summary>
        /// <param name="rest">去前导空白后的行。The line after leading spaces.</param>
        /// <param name="ordered">输出是否有序。Whether the item is ordered on output.</param>
        /// <param name="marker">输出有序编号标记（如 "3."）。The ordered marker text (such as "3.") on output.</param>
        /// <returns>项目正文。The item body.</returns>
        private static string SplitListMarker(string rest, out bool ordered, out string marker)
        {
            ordered = false;
            marker = string.Empty;
            char first = rest[0];
            if (first == '-' || first == '*' || first == '+')
            {
                return rest.Substring(1).TrimStart(' ', '\t');
            }

            int d = 0;
            while (d < rest.Length && rest[d] >= '0' && rest[d] <= '9')
            {
                d++;
            }
            char delimit = rest[d];
            string number = rest.Substring(0, d);
            ordered = true;
            marker = number + ".";
            return rest.Substring(d + 1).TrimStart(' ', '\t');
        }

        /// <summary>
        /// 若正文以 [ ] 或 [x] 起始则识别为任务项并去除它，返回前缀 "☐ " 或 "☑ "（非任务项返回 null）。
        /// Recognises a task item when the body starts with [ ] or [x], strips it, and returns the "☐ " or "☑ " prefix (null when not a task).
        /// </summary>
        /// <param name="content">项目正文（可被就地裁剪）。The item body (may be trimmed in place).</param>
        /// <returns>任务前缀或 null。The task prefix or null.</returns>
        private static string ExtractTaskCheckbox(ref string content)
        {
            if (content.Length < 3 || content[0] != '[')
            {
                return null;
            }
            char state = content[1];
            if (state != ' ' && state != 'x' && state != 'X')
            {
                return null;
            }
            if (content[2] != ']')
            {
                return null;
            }
            content = content.Substring(3).TrimStart(' ', '\t');
            return (state == ' ') ? "☐" : "☑";
        }

        /// <summary>
        /// 依据层级返回无序列表符号（• ◦ ▪）。
        /// Returns the unordered bullet for the given nesting level (• ◦ ▪).
        /// </summary>
        /// <param name="level">层级（0 起）。The nesting level (0-based).</param>
        /// <returns>列表符号。The bullet character.</returns>
        private static string BulletForLevel(int level)
        {
            if (level <= 0)
            {
                return "•";
            }
            if (level == 1)
            {
                return "◦";
            }
            return "▪";
        }

        /// <summary>
        /// 统计行首连续空格数（制表符视为 4 空格）。
        /// Counts the leading spaces at the line start (a tab counts as four spaces).
        /// </summary>
        /// <param name="line">待统计行。The line to measure.</param>
        /// <returns>前导空格数。The number of leading spaces.</returns>
        private static int CountLeadingSpaces(string line)
        {
            int count = 0;
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == ' ')
                {
                    count++;
                }
                else if (line[i] == '\t')
                {
                    count += 4;
                }
                else
                {
                    break;
                }
            }
            return count;
        }

        /// <summary>
        /// 返回行首连续空白（空格或制表符）之后的字符下标（按字符数，非制表符展开宽度）。
        /// Returns the character index just after the leading run of whitespace (spaces or tabs), counting characters rather than tab-expanded width.
        /// </summary>
        /// <param name="line">待处理行。The line to inspect.</param>
        /// <returns>首个非空白字符的下标。The index of the first non-whitespace character.</returns>
        private static int LeadingCharIndex(string line)
        {
            int i = 0;
            while (i < line.Length && (line[i] == ' ' || line[i] == '\t'))
            {
                i++;
            }
            return i;
        }

        /// <summary>
        /// 判断 index 行是否为表格表头（下一行是分隔行）。
        /// Determines whether the line at index is a table header (the following line is a delimiter row).
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="index">表头候选行号。The candidate header index.</param>
        /// <returns>是否为表格起点。Whether a table starts here.</returns>
        private static bool IsTableHeader(string[] lines, int index)
        {
            if (index + 1 >= lines.Length)
            {
                return false;
            }
            if (lines[index].IndexOf('|') < 0)
            {
                return false;
            }
            return IsTableDelimiter(lines[index + 1]);
        }

        /// <summary>
        /// 判断行是否为表格分隔行（仅由 |、:、-、空格构成且含至少一个 -）。
        /// Determines whether the line is a table delimiter row (only |, :, -, and spaces, containing at least one -).
        /// </summary>
        /// <param name="line">待判定行。The line to test.</param>
        /// <returns>是否为分隔行。Whether it is a delimiter row.</returns>
        private static bool IsTableDelimiter(string line)
        {
            string trimmed = line.Trim();
            if (trimmed.Length == 0 || trimmed.IndexOf('-') < 0 || trimmed.IndexOf('|') < 0)
            {
                return false;
            }
            for (int i = 0; i < trimmed.Length; i++)
            {
                char c = trimmed[i];
                if (c != '|' && c != ':' && c != '-' && c != ' ' && c != '\t')
                {
                    return false;
                }
            }
            return true;
        }

        /// <summary>
        /// 解析 GFM 管道表格：表头行加粗灰底且 IsHeader=true，按表头列数对数据行补齐/截断，
        /// 单元格支持行内语法，从分隔行解析左/中/右对齐。返回块之后的行号。
        /// Parses a GFM pipe table: the header row is bold, shaded and marked IsHeader; data rows are padded/truncated to the header column count;
        /// cells support inline syntax and alignment (left/center/right) comes from the delimiter row. Returns the index after the block.
        /// </summary>
        /// <param name="lines">全部行。All lines.</param>
        /// <param name="start">表头行号。The header row index.</param>
        /// <param name="section">目标节。The target section.</param>
        /// <returns>块之后的行号。The index following the block.</returns>
        private int ParseTable(string[] lines, int start, Section section)
        {
            RequireTableWidth(lines[start]);
            List<string> headerCells = SplitTableRow(lines[start]);
            List<ParagraphAlignment> aligns = ParseAlignments(lines[start + 1], headerCells.Count);
            int columnCount = headerCells.Count;
            long cellCount = columnCount;

            Table table = new Table();
            table.Autofit = true;
            for (int c = 0; c < columnCount; c++)
            {
                table.Columns.Add(new TableColumn());
            }

            TableRow header = table.AddRow();
            header.IsHeader = true;
            for (int c = 0; c < columnCount; c++)
            {
                TableCell cell = new TableCell();
                cell.BackgroundColor = OfficeColor.LightGray;
                Paragraph paragraph = new Paragraph();
                paragraph.Alignment = CellAlign(aligns, c);
                TextFormat format = NewBodyFormat();
                format.Bold = true;
                AddInlinesFromText(paragraph, c < headerCells.Count ? headerCells[c].Trim() : string.Empty, format);
                cell.Blocks.Add(paragraph);
                header.Cells.Add(cell);
            }

            int i = start + 2;
            while (i < lines.Length)
            {
                string rowLine = lines[i];
                if (rowLine.Trim().Length == 0 || rowLine.IndexOf('|') < 0)
                {
                    break;
                }
                RequireTableWidth(rowLine);
                // 比较用减法而不是先乘：列数与行数都来自文件，乘法本身可能溢出成负数而放行。
                // Compare by subtraction instead of multiplying: both factors come from the file, and the
                // product itself can overflow into a negative that the check would then accept.
                if (cellCount > MaxTableCells - columnCount)
                {
                    throw new InvalidDataException("The Markdown table exceeds the " + MaxTableCells + "-cell limit.");
                }
                cellCount += columnCount;
                List<string> cells = SplitTableRow(rowLine);
                TableRow row = table.AddRow();
                for (int c = 0; c < columnCount; c++)
                {
                    TableCell cell = new TableCell();
                    Paragraph paragraph = new Paragraph();
                    paragraph.Alignment = CellAlign(aligns, c);
                    string text = c < cells.Count ? cells[c].Trim() : string.Empty;
                    AddInlinesFromText(paragraph, text, NewBodyFormat());
                    cell.Blocks.Add(paragraph);
                    row.Cells.Add(cell);
                }
                i++;
            }

            section.Blocks.Add(table);
            return i;
        }

        /// <summary>
        /// 取指定列的对齐（越界时回退左对齐）。
        /// Gets the alignment for a column (falls back to left when out of range).
        /// </summary>
        /// <param name="aligns">对齐列表。The alignment list.</param>
        /// <param name="index">列索引。The column index.</param>
        /// <returns>段落对齐方式。The paragraph alignment.</returns>
        private static ParagraphAlignment CellAlign(List<ParagraphAlignment> aligns, int index)
        {
            if (index < 0 || index >= aligns.Count)
            {
                return ParagraphAlignment.Left;
            }
            return aligns[index];
        }

        /// <summary>
        /// 从分隔行解析每列对齐：:--- 左、---: 右、:---: 中、--- 左。
        /// Reads each column's alignment from the delimiter row: :--- left, ---: right, :---: center, --- left.
        /// </summary>
        /// <param name="delimiter">分隔行。The delimiter row.</param>
        /// <param name="columns">列数。The column count.</param>
        /// <returns>逐列对齐列表。The per-column alignment list.</returns>
        private static List<ParagraphAlignment> ParseAlignments(string delimiter, int columns)
        {
            List<string> parts = SplitTableRow(delimiter);
            List<ParagraphAlignment> aligns = new List<ParagraphAlignment>();
            for (int c = 0; c < columns; c++)
            {
                string cell = c < parts.Count ? parts[c].Trim() : string.Empty;
                bool left = cell.StartsWith(":", StringComparison.Ordinal);
                bool right = cell.EndsWith(":", StringComparison.Ordinal);
                if (left && right)
                {
                    aligns.Add(ParagraphAlignment.Center);
                }
                else if (right)
                {
                    aligns.Add(ParagraphAlignment.Right);
                }
                else
                {
                    aligns.Add(ParagraphAlignment.Left);
                }
            }
            return aligns;
        }

        /// <summary>
        /// 在物化单元格之前先数一遍竖线：SplitTableRow 会把每一列都存成字符串，等它返回再检查就晚了。
        /// Counts the pipes before the cells are materialised: SplitTableRow stores every column as a string, so a
        /// check after it returns would be too late.
        /// </summary>
        /// <param name="line">表格行。The table row.</param>
        private static void RequireTableWidth(string line)
        {
            int pipes = 0;
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == '|')
                {
                    pipes++;
                }
            }
            // 列数最多是竖线数加一（首尾不带竖线的写法），据此给出宽限一列的边界。
            // A row yields at most one more column than it has pipes (the form without outer pipes), so the
            // boundary grants that one column.
            if (pipes >= MaxTableColumns)
            {
                throw new InvalidDataException("The Markdown table exceeds the " + MaxTableColumns + "-column limit.");
            }
        }

        /// <summary>
        /// 按未转义的竖线拆分表格行为单元格文本（\| 视为字面竖线）。
        /// Splits a table row into cell texts on unescaped pipes (treating \| as a literal pipe).
        /// </summary>
        /// <param name="line">表格行。The table row.</param>
        /// <returns>单元格文本列表。The list of cell texts.</returns>
        private static List<string> SplitTableRow(string line)
        {
            List<string> cells = new List<string>();
            string trimmed = line.Trim();
            if (trimmed.StartsWith("|", StringComparison.Ordinal))
            {
                trimmed = trimmed.Substring(1);
            }
            if (trimmed.EndsWith("|", StringComparison.Ordinal))
            {
                trimmed = trimmed.Substring(0, trimmed.Length - 1);
            }

            StringBuilder current = new StringBuilder();
            for (int i = 0; i < trimmed.Length; i++)
            {
                char c = trimmed[i];
                if (c == '\\' && i + 1 < trimmed.Length && trimmed[i + 1] == '|')
                {
                    current.Append('|');
                    i++;
                }
                else if (c == '|')
                {
                    cells.Add(current.ToString());
                    current.Length = 0;
                }
                else
                {
                    current.Append(c);
                }
            }
            cells.Add(current.ToString());
            return cells;
        }

        /// <summary>
        /// 解析文本并追加为段落行内元素。
        /// Parses the text and appends the resulting inline elements to the paragraph.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="text">行内文本。The inline text.</param>
        /// <param name="baseFormat">基准文本格式。The base text format.</param>
        private void AddInlinesFromText(Paragraph paragraph, string text, TextFormat baseFormat)
        {
            // 代码块的标记一旦打上就不再改写；其余经行内解析的段落一律记为正文，
            // 免得「整段只有一个行内代码」被写器的字体推断读成代码块。
            // A code-block marker is never overwritten; every other paragraph built here is recorded as body text,
            // so a paragraph containing nothing but one code span cannot be misread as a code block.
            if (string.IsNullOrEmpty(paragraph.StyleName))
            {
                paragraph.StyleName = InlineStyleName;
            }
            List<Inline> inlines = ParseInlines(text ?? string.Empty, baseFormat);
            for (int i = 0; i < inlines.Count; i++)
            {
                paragraph.Inlines.Add(inlines[i]);
            }
        }

        /// <summary>
        /// 提取文本的纯文本形式（用于元数据标题），丢弃行内格式。
        /// Extracts the plain-text form of the given text (for the metadata title), discarding inline formatting.
        /// </summary>
        /// <param name="text">行内文本。The inline text.</param>
        /// <returns>纯文本。The plain text.</returns>
        private string PlainFromInline(string text)
        {
            Paragraph scratch = new Paragraph();
            AddInlinesFromText(scratch, text, NewBodyFormat());
            return scratch.GetText();
        }

        /// <summary>
        /// 行内单遍扫描解析：转义、行内代码、粗体/斜体/粗斜体、删除线、链接、图片、自动链接。
        /// 未闭合的分隔符按字面文本输出，全程线性、无嵌套回溯。
        /// A single-pass inline scanner: escapes, code spans, bold/italic/bold-italic, strikethrough, links, images, autolinks.
        /// Unclosed delimiters emit as literal text; the scan is linear with no nested backtracking.
        /// </summary>
        /// <param name="text">待解析文本。The text to parse.</param>
        /// <param name="baseFormat">基准文本格式。The base text format.</param>
        /// <returns>行内元素列表。The list of inline elements.</returns>
        private List<Inline> ParseInlines(string text, TextFormat baseFormat)
        {
            return ParseInlines(text, baseFormat, 0);
        }

        /// <summary>
        /// 带深度上限的行内解析：超过 <see cref="MaxInlineDepth"/> 层的强调或链接内容按字面文本输出，
        /// 避免恶意深层嵌套构造（如 <c>[[[[[…</c>）造成栈溢出。
        /// Depth-capped inline parsing: emphasis or link content nested beyond <see cref="MaxInlineDepth"/> is emitted as
        /// literal text, so a hostile deep nesting (such as <c>[[[[[…</c>) cannot overflow the stack.
        /// </summary>
        /// <param name="text">待解析文本。The text to parse.</param>
        /// <param name="baseFormat">基准文本格式。The base text format.</param>
        /// <param name="depth">当前嵌套深度。The current nesting depth.</param>
        /// <returns>行内元素列表。The list of inline elements.</returns>
        private List<Inline> ParseInlines(string text, TextFormat baseFormat, int depth)
        {
            List<Inline> result = new List<Inline>();
            if (depth > MaxInlineDepth)
            {
                if (!string.IsNullOrEmpty(text))
                {
                    result.Add(new TextRun(text, baseFormat));
                }
                return result;
            }
            StringBuilder buffer = new StringBuilder();
            int i = 0;
            int length = text.Length;
            ScanMap map = new ScanMap(text);

            while (i < length)
            {
                char c = text[i];

                if (c == '\\' && i + 1 < length && IsEscapable(text[i + 1]))
                {
                    buffer.Append(text[i + 1]);
                    i += 2;
                    continue;
                }

                if (c == '&')
                {
                    // 字符实体引用：&#65; 与 &#x41; 与 &copy; 先解成字符，写出口再按需转义；
                    // 匹配不上的形式按字面保留，与 CommonMark 一致。
                    // Character references: &#65;, &#x41; and &copy; decode to characters first, then the writers
                    // escape as usual; anything unrecognised stays literal, matching CommonMark.
                    int refEnd;
                    string decoded = TryDecodeCharRef(text, i, out refEnd);
                    if (decoded != null)
                    {
                        buffer.Append(decoded);
                        i = refEnd;
                        continue;
                    }
                }

                if (c == '`')
                {
                    // 闭合串必须与起始串等长，否则这些反引号只是正文；两侧各恰一个空格时去掉一格。
                    // The closer must be a backtick run of exactly the same length, otherwise those backticks are
                    // content; one space on each side is stripped.
                    int openLen = 1;
                    while (i + openLen < length && text[i + openLen] == '`')
                    {
                        openLen++;
                    }
                    int closer = -1;
                    int scan = i + openLen;
                    while (scan < length)
                    {
                        int at = map.Next('`', scan);
                        if (at < 0)
                        {
                            break;
                        }
                        int run = 1;
                        while (at + run < length && text[at + run] == '`')
                        {
                            run++;
                        }
                        if (run == openLen)
                        {
                            closer = at;
                            break;
                        }
                        scan = at + run;
                    }
                    if (closer < 0)
                    {
                        buffer.Append(text.Substring(i, openLen));
                        i += openLen;
                        continue;
                    }
                    string code = text.Substring(i + openLen, closer - i - openLen);
                    if (code.Length >= 2 && code[0] == ' ' && code[code.Length - 1] == ' ' && HasNonSpace(code))
                    {
                        code = code.Substring(1, code.Length - 2);
                    }
                    FlushBuffer(buffer, result, baseFormat);
                    result.Add(new TextRun(code, NewCodeFormat()));
                    i = closer + openLen;
                    continue;
                }

                if (c == '!' && i + 1 < length && text[i + 1] == '[')
                {
                    int consumed;
                    List<Inline> image = TryParseImage(text, i, baseFormat, map, out consumed);
                    if (image != null)
                    {
                        FlushBuffer(buffer, result, baseFormat);
                        result.AddRange(image);
                        i = consumed;
                        continue;
                    }
                    buffer.Append(c);
                    i++;
                    continue;
                }

                if (c == '[')
                {
                    int consumed;
                    List<Inline> link = TryParseLink(text, i, baseFormat, depth + 1, map, out consumed);
                    if (link != null)
                    {
                        FlushBuffer(buffer, result, baseFormat);
                        result.AddRange(link);
                        i = consumed;
                        continue;
                    }
                    buffer.Append(c);
                    i++;
                    continue;
                }

                if (c == '<')
                {
                    int consumed;
                    List<Inline> auto = TryParseAutolink(text, i, baseFormat, map, out consumed);
                    if (auto != null)
                    {
                        FlushBuffer(buffer, result, baseFormat);
                        result.AddRange(auto);
                        i = consumed;
                        continue;
                    }
                    buffer.Append(c);
                    i++;
                    continue;
                }

                if (c == '~' && i + 1 < length && text[i + 1] == '~')
                {
                    int closer = FindDelim(text, "~~", i + 2);
                    // 空内容不成删除线，且与 * 和 _ 分支一致把整串波浪号按字面输出：
                    // 只跳过两个波浪号会让 "~~~~" 的四个全部消失，文本凭空少掉。
                    // An empty body is not a strikeout and, like the * and _ branches, the whole tilde run falls
                    // back to literals: skipping only two tildes makes "~~~~" swallow all four and the text
                    // silently disappears.
                    if (closer < 0 || closer == i + 2)
                    {
                        int run = 1;
                        while (i + run < length && text[i + run] == '~')
                        {
                            run++;
                        }
                        buffer.Append(text.Substring(i, run));
                        i += run;
                        continue;
                    }
                    string inner = text.Substring(i + 2, closer - i - 2);
                    TextFormat strike = baseFormat.Clone();
                    strike.Strikeout = true;
                    FlushBuffer(buffer, result, baseFormat);
                    result.AddRange(ParseInlines(inner, strike, depth + 1));
                    i = closer + 2;
                    continue;
                }

                if (c == '*' || c == '_')
                {
                    int run = 1;
                    while (i + run < length && text[i + run] == c)
                    {
                        run++;
                    }
                    // CommonMark 的两侧贴合规则：起始串整串后面必须紧跟内容（非空白），闭合串前面同理；
                    // 下划线还禁止词内强调（in_var_iable、foo_bar_baz 都保持字面）。
                    // CommonMark's flanking rules: the whole opener run must be followed by content and the closer
                    // run preceded by content; underscores additionally refuse intraword emphasis (in_var_iable,
                    // foo_bar_baz stay literal).
                    bool leftFlank = IsLeftFlankingRun(text, c, i, run);
                    int closer = -1;
                    int closerRun = 0;
                    if (leftFlank)
                    {
                        // CommonMark 规则 11：一方能同时开合强调、另一方不能时，这一对不成立。
                        // 这里用 "两侧都能开又能合" 的布尔值相等来近似：一端是纯起始串、另一端是纯闭合串才配对，
                        // 像 *a**b* 里既是起始又是闭合的 ** 就不去抢那对单星。
                        // CommonMark rule 11: when one delimiter can both open and close emphasis and the other
                        // cannot, the pair is invalid. Approximated here by requiring both sides to agree on
                        // "can open and can close", so the ** in *a**b* (which both opens and closes) does not steal
                        // the single-star pair.
                        bool openerBoth = IsRightFlankingRun(text, c, i, run);
                        // 第一轮只接受长度不小于起始串的闭合串：`**a *b* c**` 中间那对单星属于内层强调，
                        // 不能当成外层 `**` 的闭合。第一轮空手而归才允许半配对，`***a**` 用掉两枚成粗体，
                        // 起始侧多出的那一枚退回字面文本。
                        // The first pass only accepts a closer at least as long as the opener: inside `**a *b* c**`
                        // the single asterisks belong to the inner emphasis and cannot close the outer pair. Only if
                        // that finds nothing does the second pass allow a partial pair, so `***a**` spends two markers
                        // on the strong and leaves the surplus opener as literal text.
                        for (int pass = 0; pass < 2 && closer < 0; pass++)
                        {
                            int need = pass == 0 ? run : 1;
                            int matched;
                            closer = FindEmphasisCloser(text, c, i, run, need, openerBoth, 0, out matched);
                            if (closer >= 0)
                            {
                                closerRun = matched;
                            }
                        }
                    }
                    if (closer < 0)
                    {
                        buffer.Append(text.Substring(i, run));
                        i += run;
                        continue;
                    }
                    // 一对分隔符各吃掉 min(两侧长度) 个字符：`**a***` 用掉两枚成粗体，余下一枚退回字面。
                    // Each pair consumes min(run lengths) markers, so `**a***` spends two on the strong and leaves one
                    // literal.
                    int used = Math.Min(run, closerRun);
                    TextFormat em = baseFormat.Clone();
                    if (used >= 2)
                    {
                        em.Bold = true;
                    }
                    // 四枚是两层粗体，模型里粗体不可叠加，斜体只在层数为奇数时才出现。
                    // Four markers mean two nested strongs, which the model flattens to bold alone; italics only come
                    // with an odd number of levels.
                    if ((used & 1) == 1)
                    {
                        em.Italic = true;
                    }
                    for (int lead = run - used; lead > 0; lead--)
                    {
                        buffer.Append(c);
                    }
                    FlushBuffer(buffer, result, baseFormat);
                    string inner = text.Substring(i + run, closer - i - run);
                    result.AddRange(ParseInlines(inner, em, depth + 1));
                    i = closer + used;
                    continue;
                }

                buffer.Append(c);
                i++;
            }

            FlushBuffer(buffer, result, baseFormat);
            return result;
        }

        /// <summary>
        /// 尝试解析图片 ![alt](src "title")：本地文件存在才嵌图，否则降级为 alt 文本。
        /// Attempts to parse an image ![alt](src "title"): embeds it only when the local file exists, otherwise downgrades to the alt text.
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="start">'!' 的位置。The index of '!'.</param>
        /// <param name="baseFormat">基准格式。The base format.</param>
        /// <param name="map">本段文本的配对查找表。The pairing table for this text.</param>
        /// <param name="consumed">输出消费到的位置。The consumed end index on output.</param>
        /// <returns>行内元素或 null。The inline elements or null.</returns>
        private List<Inline> TryParseImage(string text, int start, TextFormat baseFormat, ScanMap map, out int consumed)
        {
            consumed = 0;
            int bracketOpen = start + 1;
            int bracketClose = map.BracketClose(bracketOpen);
            if (bracketClose < 0)
            {
                return null;
            }
            string alt = text.Substring(bracketOpen + 1, bracketClose - bracketOpen - 1);
            string src;
            int end;
            if (!TryReadLinkTarget(text, alt, bracketClose, map, out src, out end))
            {
                return null;
            }

            consumed = end;

            OfficeImage image = TryLoadLocalImage(src);
            if (image != null)
            {
                List<Inline> list = new List<Inline>();
                list.Add(new ImageRun(image, image.DefaultWidthPoints, image.DefaultHeightPoints));
                return list;
            }

            // 网络/不存在的图片：降级为 alt 文本（无 alt 时用方括号原样）。
            // Network or missing images: downgrade to the alt text (or the bracketed form when there is none).
            List<Inline> fallback = new List<Inline>();
            if (!string.IsNullOrEmpty(alt))
            {
                fallback.Add(new TextRun(alt, baseFormat));
            }
            else
            {
                fallback.Add(new TextRun(text.Substring(start, end - start), baseFormat));
            }
            return fallback;
        }

        /// <summary>
        /// 读取 <c>]</c> 之后的链接目标：内联形式 <c>(url "title")</c> 与引用形式 <code>[label]</code> / 简写 <code>[label]</code>
        /// 都支持；内联目标含未转义空白或尾随杂串时整体退回字面文本（CommonMark 判定不合法）。
        /// Reads the link target following <c>]</c>: the inline form <c>(url "title")</c> and the reference forms
        /// <code>[text][label]</code>, <code>[text][]</code> and the shortcut <code>[label]</code> all resolve; an
        /// inline destination with unescaped whitespace or trailing junk makes the whole construct literal (CommonMark
        /// rejects it).
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="linkText">方括号内的文本，引用简写时兼作标签。The bracketed text, also the label for a shortcut reference.</param>
        /// <param name="bracketClose">闭合 <c>]</c> 的下标。The index of the closing ']'.</param>
        /// <param name="map">本段文本的配对查找表。The pairing table for this text.</param>
        /// <param name="target">输出的链接目标。The destination on output.</param>
        /// <param name="consumed">输出消费到的位置。The consumed end index on output.</param>
        /// <returns>是否解析成功。Whether a target was resolved.</returns>
        private bool TryReadLinkTarget(string text, string linkText, int bracketClose, ScanMap map, out string target, out int consumed)
        {
            target = null;
            consumed = 0;
            int after = bracketClose + 1;
            while (after < text.Length && (text[after] == ' ' || text[after] == '\t'))
            {
                after++;
            }

            if (after < text.Length && text[after] == '(')
            {
                int parenClose = map.ParenClose(after);
                if (parenClose < 0)
                {
                    return false;
                }
                string raw = text.Substring(after + 1, parenClose - after - 1);
                string dest;
                if (!TrySplitLinkTarget(raw, out dest))
                {
                    return false;
                }
                target = dest;
                consumed = parenClose + 1;
                return true;
            }

            if (_definitions.Count == 0)
            {
                return false;
            }

            string label = NormalizeLabel(linkText);
            int end = after;
            if (after < text.Length && text[after] == '[')
            {
                int refClose = IndexUnescaped(text, ']', after + 1);
                if (refClose < 0)
                {
                    return false;
                }
                string explicitLabel = NormalizeLabel(text.Substring(after + 1, refClose - after - 1));
                if (explicitLabel.Length > 0)
                {
                    label = explicitLabel;
                }
                end = refClose + 1;
            }
            LinkDefinition definition;
            if (label.Length == 0 || label.Length > MaxDefinitionLabelChars || !_definitions.TryGetValue(label, out definition))
            {
                return false;
            }
            target = definition.Destination;
            consumed = end;
            return true;
        }


        /// <summary>
        /// 尝试解析链接：内联形式 <c>[text](url "title")</c> 与引用形式 <c>[text][ref]</c>、<c>[text][]</c>、<c>[ref]</c>。
        /// 链接文本以蓝色下划线呈现，并把原始地址记在 TextRun.Link 上供写器输出超链接。
        /// Attempts to parse a link: the inline form <c>[text](url "title")</c> and the reference forms
        /// <c>[text][ref]</c>, <c>[text][]</c> and <c>[ref]</c>. The text renders in blue underline and the resolved
        /// destination is recorded on TextRun.Link so writers can emit a hyperlink.
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="start">'[' 的位置。The index of '['.</param>
        /// <param name="baseFormat">基准格式。The base format.</param>
        /// <param name="depth">当前嵌套深度。The current nesting depth.</param>
        /// <param name="map">本段文本的配对查找表。The pairing table for this text.</param>
        /// <param name="consumed">输出消费到的位置。The consumed end index on output.</param>
        /// <returns>行内元素或 null。The inline elements or null.</returns>
        private List<Inline> TryParseLink(string text, int start, TextFormat baseFormat, int depth, ScanMap map, out int consumed)
        {
            consumed = 0;
            int bracketClose = map.BracketClose(start);
            if (bracketClose < 0)
            {
                return null;
            }
            string linkText = text.Substring(start + 1, bracketClose - start - 1);
            string url;
            int end;
            if (!TryReadLinkTarget(text, linkText, bracketClose, map, out url, out end))
            {
                return null;
            }

            TextFormat linkFormat = baseFormat.Clone();
            linkFormat.Underline = true;
            linkFormat.Color = OfficeColor.Blue;

            List<Inline> inner = ParseInlines(linkText, linkFormat, depth + 1);
            if (inner.Count == 0 && !string.IsNullOrEmpty(url))
            {
                inner.Add(new TextRun(url, linkFormat));
            }
            for (int i = 0; i < inner.Count; i++)
            {
                TextRun run = inner[i] as TextRun;
                if (run != null && string.IsNullOrEmpty(run.Link))
                {
                    run.Link = url;
                }
            }
            consumed = end;
            return inner;
        }

        /// <summary>
        /// 尝试解析自动链接 &lt;http://...&gt; / &lt;mailto:...&gt;。
        /// Attempts to parse an autolink &lt;http://...&gt; / &lt;mailto:...&gt;.
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="start">'&lt;' 的位置。The index of '&lt;'.</param>
        /// <param name="baseFormat">基准格式。The base format.</param>
        /// <param name="map">本段文本的配对查找表。The pairing table for this text.</param>
        /// <param name="consumed">输出消费到的位置。The consumed end index on output.</param>
        /// <returns>行内元素或 null。The inline elements or null.</returns>
        private List<Inline> TryParseAutolink(string text, int start, TextFormat baseFormat, ScanMap map, out int consumed)
        {
            consumed = 0;
            int close = map.Next('>', start + 1);
            if (close < 0)
            {
                return null;
            }
            string body = text.Substring(start + 1, close - start - 1);
            if (body.Length == 0 || body.IndexOf(' ') >= 0)
            {
                return null;
            }
            bool looksUrl = body.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || body.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
                || body.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)
                || body.StartsWith("www.", StringComparison.OrdinalIgnoreCase);
            if (!looksUrl)
            {
                return null;
            }

            TextFormat linkFormat = baseFormat.Clone();
            linkFormat.Underline = true;
            linkFormat.Color = OfficeColor.Blue;
            List<Inline> list = new List<Inline>();
            TextRun auto = new TextRun(body, linkFormat);
            auto.Link = body;
            list.Add(auto);
            consumed = close + 1;
            return list;
        }

        /// <summary>
        /// 校验并拆出括号内的链接目标：目标必须不含未转义空白（尖括号包裹除外），
        /// 其后可跟一个成对引号标题；任何多余内容都让整个结构退回字面文本（CommonMark 的判定）。
        /// Validates and splits the parenthesized target: the destination may not contain unescaped whitespace unless
        /// wrapped in angle brackets, an optional quoted title may follow, and any leftover text makes the whole
        /// construct fall back to literals (CommonMark's verdict).
        /// </summary>
        /// <param name="raw">括号内的原始内容。The raw content between the parentheses.</param>
        /// <param name="destination">输出的链接目标。The destination on output.</param>
        /// <returns>是否为合法目标。Whether the target is valid.</returns>
        private static bool TrySplitLinkTarget(string raw, out string destination)
        {
            destination = null;
            int i = 0;
            while (i < raw.Length && char.IsWhiteSpace(raw[i]))
            {
                i++;
            }
            if (i >= raw.Length)
            {
                // [a]() 是合法的空目标。An empty target, as in [a](), is valid.
                destination = string.Empty;
                return true;
            }
            int end;
            string dest = ReadDefinitionDestination(raw, i, out end);
            if (dest == null)
            {
                return false;
            }
            i = end;
            while (i < raw.Length && char.IsWhiteSpace(raw[i]))
            {
                i++;
            }
            if (i >= raw.Length)
            {
                destination = dest;
                return true;
            }
            string title;
            int titleEnd;
            if (!ReadDefinitionTitle(raw, i, out title, out titleEnd))
            {
                return false;
            }
            i = titleEnd;
            while (i < raw.Length && char.IsWhiteSpace(raw[i]))
            {
                i++;
            }
            if (i < raw.Length)
            {
                return false;
            }
            destination = dest;
            return true;
        }

        /// <summary>
        /// 在基准目录沙箱内解析并读取本地图片：目录未设置、地址带协议、路径为绝对/UNC、
        /// 归一化后逃逸出目录、或文件超过上限时都安全返回 null（降级为替代文本），全程不抛异常。
        /// Resolves and reads a local image inside the base-directory sandbox: returns null (so the caller degrades to
        /// the alt text) when no base directory is set, the source carries a scheme, the path is absolute or UNC,
        /// the normalized path escapes the directory, or the file exceeds the size cap. Never throws.
        /// </summary>
        /// <param name="src">图片来源字符串。The image source string.</param>
        /// <returns>图片实例或 null。An image instance or null.</returns>
        private OfficeImage TryLoadLocalImage(string src)
        {
            if (string.IsNullOrEmpty(src))
            {
                return null;
            }
            string baseDirectory = options.ImageBaseDirectory;
            if (string.IsNullOrEmpty(baseDirectory))
            {
                return null;
            }
            if (src.IndexOf("://", StringComparison.Ordinal) >= 0
                || src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            if (Path.IsPathRooted(src) || (src.Length >= 2 && src[1] == ':'))
            {
                return null;
            }
            if (src.IndexOfAny(Path.GetInvalidPathChars()) >= 0)
            {
                return null;
            }
            try
            {
                string root = Path.GetFullPath(baseDirectory);
                string full = Path.GetFullPath(Path.Combine(root, src));
                if (!IsInsideDirectory(full, root))
                {
                    return null;
                }
                OfficeImage cached;
                if (this._imageCache.TryGetValue(full, out cached))
                {
                    return cached;
                }
                FileInfo info = new FileInfo(full);
                if (!info.Exists || info.Length == 0)
                {
                    return null;
                }
                int cap = options.MaxImageBytes > 0 ? options.MaxImageBytes : MarkdownOptions.DefaultMaxImageBytes;
                if (info.Length > cap)
                {
                    return null;
                }
                // 用减法比较总量上限：累计值与声明长度都是 int，直接相加可能回绕成负数而放行。
                // Compare the total by subtraction: both the running sum and the declared length are ints, and a
                // plain addition can wrap into a negative that the check would then accept.
                if (this._imageBytes > MaxDocumentImageBytes - info.Length)
                {
                    return null;
                }
                byte[] bytes = File.ReadAllBytes(full);
                if (bytes.Length == 0)
                {
                    return null;
                }
                OfficeImage image = new OfficeImage(bytes);
                this._imageCache[full] = image;
                this._imageBytes += bytes.Length;
                return image;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 判断已归一化的绝对路径是否位于目录内部（Windows 忽略大小写，其余平台区分大小写）。
        /// Decides whether a normalized absolute path lies inside a directory (case-insensitive on Windows, case-sensitive elsewhere).
        /// </summary>
        /// <param name="path">归一化后的文件路径。The normalized file path.</param>
        /// <param name="directory">归一化后的目录路径。The normalized directory path.</param>
        /// <returns>是否被目录包含。Whether the directory contains the path.</returns>
        private static bool IsInsideDirectory(string path, string directory)
        {
            string prefix = directory.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? directory
                : directory + Path.DirectorySeparatorChar;
            StringComparison comparison = Path.DirectorySeparatorChar == '\\'
                ? StringComparison.OrdinalIgnoreCase
                : StringComparison.Ordinal;
            return path.StartsWith(prefix, comparison);
        }

        /// <summary>
        /// 判断字符是否可被反斜杠转义。
        /// Determines whether a character can be backslash-escaped.
        /// </summary>
        /// <param name="c">待判定字符。The character to test.</param>
        /// <returns>是否可转义。Whether it is escapable.</returns>
        private static bool IsEscapable(char c)
        {
            switch (c)
            {
                case '\\':
                case '`':
                case '*':
                case '_':
                case '{':
                case '}':
                case '[':
                case ']':
                case '(':
                case ')':
                case '#':
                case '+':
                case '-':
                case '.':
                case '!':
                case '<':
                case '>':
                case '=':
                case '|':
                case '~':
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 判断串中是否含有非空白字符（行内代码去首尾空格的先决条件：全空白的内容不许裁切）。
        /// Tests whether a string contains any non-whitespace character (the precondition for stripping one space
        /// from each end of a code span; an all-whitespace body must not be trimmed).
        /// </summary>
        /// <param name="text">待判定串。The string to test.</param>
        /// <returns>是否含非空白字符。Whether it contains a non-whitespace character.</returns>
        private static bool HasNonSpace(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (!char.IsWhiteSpace(text[i]))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 从 &lt;start&gt; 处的 '&amp;' 起解析字符实体引用：数字形式 &amp;#65; / &amp;#x41; 与命名字形式 &amp;copy;。
        /// 命名字必须全为 ASCII 字母数字且以 ';' 结尾；码点为 0、落在代理区或超过 U+10FFFF 视为非法。
        /// 非法时返回 null 并把 '&amp;' 按字面保留，与 CommonMark 一致。
        /// Parses a character reference starting at the '&amp;' at &lt;start&gt;: numeric &amp;#65; / &amp;#x41; and named
        /// &amp;copy;. A name must be ASCII letters/digits and end with ';'; code points of 0, in the surrogate range or
        /// above U+10FFFF are rejected. Returns null on any rejection so the '&amp;' stays literal, matching CommonMark.
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="start">'&amp;' 的位置。The index of '&amp;'.</param>
        /// <param name="end">成功时置为实体结束后的位置。On success set to the index just after the entity.</param>
        /// <returns>解码得到的文本，非法或未知时为 null。The decoded text, or null when invalid or unknown.</returns>
        private static string TryDecodeCharRef(string text, int start, out int end)
        {
            end = start;
            int length = text.Length;
            int i = start + 1;
            if (i >= length)
            {
                return null;
            }

            if (text[i] == '#')
            {
                i++;
                bool hex = i < length && (text[i] == 'x' || text[i] == 'X');
                if (hex)
                {
                    i++;
                }
                int digitsStart = i;
                int value = 0;
                bool overflow = false;
                while (i < length)
                {
                    int digit;
                    char c = text[i];
                    if (c >= '0' && c <= '9')
                    {
                        digit = c - '0';
                    }
                    else if (hex && c >= 'a' && c <= 'f')
                    {
                        digit = c - 'a' + 10;
                    }
                    else if (hex && c >= 'A' && c <= 'F')
                    {
                        digit = c - 'A' + 10;
                    }
                    else
                    {
                        break;
                    }
                    int radix = hex ? 16 : 10;
                    if (value > (0x10FFFF - digit) / radix)
                    {
                        overflow = true;
                    }
                    else
                    {
                        value = value * radix + digit;
                    }
                    i++;
                }
                // 无数字、缺分号、溢出或落在保留码点区间都算非法引用。
                // No digits, a missing ';', overflow or a reserved code point all make the reference invalid.
                if (i == digitsStart || overflow || i >= length || text[i] != ';')
                {
                    return null;
                }
                if (value == 0 || (value >= 0xD800 && value <= 0xDFFF) || value > 0x10FFFF)
                {
                    return null;
                }
                end = i + 1;
                return char.ConvertFromUtf32(value);
            }

            int nameStart = i;
            while (i < length)
            {
                char c = text[i];
                bool asciiName = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9');
                if (!asciiName)
                {
                    break;
                }
                i++;
            }
            if (i == nameStart || i >= length || text[i] != ';')
            {
                return null;
            }
            string named;
            if (!CharEntityTable.TryGetValue(text.Substring(nameStart, i - nameStart), out named))
            {
                return null;
            }
            end = i + 1;
            return named;
        }

        /// <summary>
        /// 常用命名字符实体表：ASCII 标点、排版标点、货币与常用符号/数学记号。键不含 '&amp;' 与 ';'。
        /// Curated table of common named character references: ASCII punctuation, typographic punctuation, currency
        /// and frequent symbols/math marks. Keys exclude the leading '&amp;' and trailing ';'.
        /// </summary>
        private static readonly Dictionary<string, string> CharEntityTable = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "amp", "&" },
            { "lt", "<" },
            { "gt", ">" },
            { "quot", "\"" },
            { "apos", "'" },
            { "nbsp", "\u00A0" },
            { "shy", "\u00AD" },
            { "copy", "\u00A9" },
            { "reg", "\u00AE" },
            { "trade", "\u2122" },
            { "pound", "\u00A3" },
            { "cent", "\u00A2" },
            { "yen", "\u00A5" },
            { "currency", "\u00A4" },
            { "euro", "\u20AC" },
            { "deg", "\u00B0" },
            { "micro", "\u00B5" },
            { "para", "\u00B6" },
            { "sect", "\u00A7" },
            { "middot", "\u00B7" },
            { "bull", "\u2022" },
            { "hellip", "\u2026" },
            { "mdash", "\u2014" },
            { "ndash", "\u2013" },
            { "lsquo", "\u2018" },
            { "rsquo", "\u2019" },
            { "ldquo", "\u201C" },
            { "rdquo", "\u201D" },
            { "laquo", "\u00AB" },
            { "raquo", "\u00BB" },
            { "lsaquo", "\u2039" },
            { "rsaquo", "\u203A" },
            { "times", "\u00D7" },
            { "divide", "\u00F7" },
            { "plusmn", "\u00B1" },
            { "mp", "\u2213" },
            { "frac12", "\u00BD" },
            { "frac14", "\u00BC" },
            { "frac34", "\u00BE" },
            { "sup1", "\u00B9" },
            { "sup2", "\u00B2" },
            { "sup3", "\u00B3" },
            { "dagger", "\u2020" },
            { "Dagger", "\u2021" },
            { "prime", "\u2032" },
            { "Prime", "\u2033" },
            { "larr", "\u2190" },
            { "uarr", "\u2191" },
            { "rarr", "\u2192" },
            { "darr", "\u2193" },
            { "harr", "\u2194" },
            { "ne", "\u2260" },
            { "le", "\u2264" },
            { "ge", "\u2265" },
            { "minus", "\u2212" },
            { "infin", "\u221E" },
            { "sum", "\u2211" },
            { "prod", "\u220F" },
            { "alpha", "\u03B1" },
            { "beta", "\u03B2" },
            { "gamma", "\u03B3" },
            { "delta", "\u03B4" },
            { "epsilon", "\u03B5" },
            { "pi", "\u03C0" },
            { "sigma", "\u03C3" },
            { "tau", "\u03C4" },
            { "phi", "\u03C6" },
            { "omega", "\u03C9" },
            { "ensp", "\u2002" },
            { "emsp", "\u2003" },
            { "thinsp", "\u2009" },
            { "zwnj", "\u200C" },
            { "zwj", "\u200D" },
        };

        /// <summary>
        /// 从 from 起查找目标分隔串（跳过转义）的首个位置，找不到返回 -1。
        /// Finds the first unescaped occurrence of the delimiter string from the given index, or -1 when absent.
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="delim">目标分隔串。The delimiter string.</param>
        /// <param name="from">起始查找位置。The start search index.</param>
        /// <returns>目标位置或 -1。The found index or -1.</returns>
        private static int FindDelim(string text, string delim, int from)
        {
            int i = from;
            while (i <= text.Length - delim.Length)
            {
                if (text[i] == '\\')
                {
                    i += 2;
                    continue;
                }
                if (string.CompareOrdinal(text, i, delim, 0, delim.Length) == 0)
                {
                    return i;
                }
                i++;
            }
            return -1;
        }

        /// <summary>
        /// 从 from 起查找下一段连续的同种强调字符（跳过转义），并输出这段的长度。
        /// Finds the next run of consecutive emphasis characters from the given index, skipping escapes, and reports
        /// how long that run is.
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="delimiter">目标字符。The character to look for.</param>
        /// <param name="from">起始查找位置。The start search index.</param>
        /// <param name="length">输出的连续长度。The length of the run on output.</param>
        /// <returns>这段的起始位置或 -1。The start index of the run or -1.</returns>
        private static int NextDelimRun(string text, char delimiter, int from, out int length)
        {
            length = 0;
            int i = from;
            while (i < text.Length)
            {
                if (text[i] == '\\')
                {
                    i += 2;
                    continue;
                }
                if (text[i] == delimiter)
                {
                    int run = 1;
                    while (i + run < text.Length && text[i + run] == delimiter)
                    {
                        run++;
                    }
                    length = run;
                    return i;
                }
                i++;
            }
            return -1;
        }

        /// <summary>
        /// 判断 &lt;start&gt; 处长为 length 的分隔符串能否开启强调：整串后紧跟内容，下划线另禁词内起始。
        /// Tests whether the delimiter run of the given length at &lt;start&gt; can open emphasis: the whole run must be
        /// followed by content, and an underscore run additionally refuses to open inside a word.
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="delimiter">分隔符字符。The delimiter character.</param>
        /// <param name="start">分隔符串起点。The start index of the run.</param>
        /// <param name="length">分隔符串长度。The length of the run.</param>
        /// <returns>能否开启强调。Whether the run can open emphasis.</returns>
        private static bool IsLeftFlankingRun(string text, char delimiter, int start, int length)
        {
            int after = start + length;
            if (after >= text.Length || char.IsWhiteSpace(text[after]))
            {
                return false;
            }
            if (delimiter == '_' && start > 0 && char.IsLetterOrDigit(text[start - 1]))
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 判断 &lt;start&gt; 处长为 length 的分隔符串能否闭合强调：整串前紧跟内容，下划线另禁词内闭合。
        /// Tests whether the delimiter run of the given length at &lt;start&gt; can close emphasis: the whole run must be
        /// preceded by content, and an underscore run additionally refuses to close inside a word.
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="delimiter">分隔符字符。The delimiter character.</param>
        /// <param name="start">分隔符串起点。The start index of the run.</param>
        /// <param name="length">分隔符串长度。The length of the run.</param>
        /// <returns>能否闭合强调。Whether the run can close emphasis.</returns>
        private static bool IsRightFlankingRun(string text, char delimiter, int start, int length)
        {
            if (start <= 0 || char.IsWhiteSpace(text[start - 1]))
            {
                return false;
            }
            int after = start + length;
            if (delimiter == '_' && after < text.Length && char.IsLetterOrDigit(text[after]))
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 为一个强调起始串寻找闭合串：先替内层起始串找它自己的闭合串并整对跳过，
        /// 再按 CommonMark 规则 11 拒绝"一端能同时开合、另一端不能"的配对。
        /// Finds the closer for an emphasis opener: an inner run that can itself open emphasis gets its own closer
        /// searched first so the whole inner pair is skipped, and per CommonMark rule 11 a pair is rejected when one
        /// side can both open and close emphasis while the other cannot (so the ** in *a**b* does not steal the
        /// single-star pair).
        /// </summary>
        /// <param name="text">全文。The full text.</param>
        /// <param name="delimiter">分隔符字符。The delimiter character.</param>
        /// <param name="openStart">起始串起点。The start index of the opener run.</param>
        /// <param name="openLength">起始串长度。The length of the opener run.</param>
        /// <param name="minLength">闭合串允许的最短长度。The shortest closer run accepted.</param>
        /// <param name="openerBoth">起始串是否既能开又能合。Whether the opener can both open and close emphasis.</param>
        /// <param name="depth">内层递归深度，超上限即停止下探。The inner recursion depth; scanning stops at the cap.</param>
        /// <param name="matchedLength">输出闭合串实测长度。The measured length of the closer run on output.</param>
        /// <returns>闭合串起点或 -1。The start index of the closer, or -1.</returns>
        private static int FindEmphasisCloser(string text, char delimiter, int openStart, int openLength, int minLength, bool openerBoth, int depth, out int matchedLength)
        {
            matchedLength = 0;
            int length = text.Length;
            int search = openStart + openLength;
            while (search < length)
            {
                int runLength;
                int at = NextDelimRun(text, delimiter, search, out runLength);
                if (at < 0)
                {
                    return -1;
                }
                int after = at + runLength;
                bool canClose = IsRightFlankingRun(text, delimiter, at, runLength);
                bool canOpen = IsLeftFlankingRun(text, delimiter, at, runLength);
                // 单星起始串不抢两枚以上、且自身还能开场的闭合串：*a**b* 里的 ** 属于内容，
                // 抢过来会让中间两个星凭空消失。等长的正常配对（*a*、**a**）不受影响。
                // A one-star opener must not steal a longer run that can itself open emphasis: in *a**b* the ** is
                // content, and pairing with it would silently delete the two markers. Equal-length pairs (*a*,
                // **a**) are unaffected.
                bool ruleElevenBlocked = canOpen && runLength >= 2 && openLength == 1 && !openerBoth;
                // 反向的同形陷阱：a**b* 里既能开场又能闭合的双星去配一个更短的单星，用掉一枚就剩不下闭合位，
                // 整串只能退回字面文本，否则尾部那颗星会凭空消失。等长或更长的闭合串照常配对。
                // The mirror-image trap: in a**b* a run that can both open and close tries to close with a shorter
                // single star, and spending one marker leaves nothing to close with, so the whole run must fall back
                // to literals or the trailing star silently disappears. Equal-or-longer closers pair as usual.
                bool partialCloserBlocked = canClose && !canOpen && openerBoth && runLength < openLength;
                if (ruleElevenBlocked || partialCloserBlocked)
                {
                    // 被规则挡住的这串按内容跳过，不再替它找内层配对：*a**b* 的 ** 就此留在正文里。
                    // A run blocked by the rule is skipped as content instead of being given an inner partner, so
                    // the ** in *a**b* stays in the text.
                    search = after;
                    continue;
                }
                if (canClose && runLength >= minLength)
                {
                    matchedLength = runLength;
                    return at;
                }
                int innerMatched;
                if (canOpen && depth < MaxDelimiterScanDepth)
                {
                    int inner = FindEmphasisCloser(text, delimiter, at, runLength, 1, canClose, depth + 1, out innerMatched);
                    if (inner >= 0)
                    {
                        search = inner + innerMatched;
                        continue;
                    }
                }
                search = after > search ? after : search + 1;
            }
            return -1;
        }

        /// <summary>内层强调配对搜索的最大递归深度，避免超长分隔符串把扫描拖成平方级。Maximum nesting of the inner emphasis pair search, so long delimiter runs cannot degrade scanning into quadratic time.</summary>
        private const int MaxDelimiterScanDepth = 32;

        /// <summary>
        /// 将缓冲区文本作为带基准格式的运行刷出（空缓冲不产生元素）。
        /// Flushes the buffered text as a base-formatted run (an empty buffer produces no element).
        /// </summary>
        /// <param name="buffer">文本缓冲。The text buffer.</param>
        /// <param name="result">行内结果列表。The inline result list.</param>
        /// <param name="baseFormat">基准格式。The base format.</param>
        private static void FlushBuffer(StringBuilder buffer, List<Inline> result, TextFormat baseFormat)
        {
            if (buffer.Length == 0)
            {
                return;
            }
            result.Add(new TextRun(buffer.ToString(), baseFormat));
            buffer.Length = 0;
        }

        /// <summary>
        /// 去除最多三个前导空格（CommonMark 允许块起始≤3 空格缩进）。
        /// Trims up to three leading spaces (CommonMark allows block starts indented by at most three spaces).
        /// </summary>
        /// <param name="line">原始行。The raw line.</param>
        /// <returns>去缩进后的行。The line after indent removal.</returns>
        private static string TrimLeftUpTo3(string line)
        {
            int i = 0;
            while (i < line.Length && i < 3 && line[i] == ' ')
            {
                i++;
            }
            return line.Substring(i);
        }

        /// <summary>
        /// 创建正文文本格式（字号取基准，字体为 BodyFontFamily 或回退默认）。
        /// Creates a body text format (base size; BodyFontFamily when set, otherwise the default font).
        /// </summary>
        /// <returns>新的正文格式。A new body format.</returns>
        private TextFormat NewBodyFormat()
        {
            TextFormat format = new TextFormat();
            format.FontSize = this.options.BaseFontSize;
            if (!string.IsNullOrEmpty(this.options.BodyFontFamily))
            {
                format.FontFamily = this.options.BodyFontFamily;
            }
            format.Bold = false;
            format.Italic = false;
            format.Underline = false;
            return format;
        }

        /// <summary>
        /// 创建代码文本格式（等宽字体）。浅灰底纹在本模型段落中不可用，故仅使用等宽字体表达。
        /// Creates a code text format (monospaced font). The light-gray shading is unavailable on model paragraphs, so only the monospaced font is used to express code.
        /// </summary>
        /// <returns>新的代码格式。A new code format.</returns>
        private TextFormat NewCodeFormat()
        {
            TextFormat format = new TextFormat();
            string family = string.IsNullOrEmpty(this.options.CodeFontFamily) ? "Consolas" : this.options.CodeFontFamily;
            format.FontFamily = family;
            format.FontSize = Math.Max(8, this.options.BaseFontSize - 0.5);
            format.Bold = false;
            format.Italic = false;
            return format;
        }

        /// <summary>
        /// 行内文本的配对查找表。首次用到某类字符时一次性扫描全文并缓存结果，之后每个位置的查询都是 O(1)。
        /// 若沿用逐个位置向后扫描，一段大量未闭合的 <c>[</c>、<c>(</c> 或 <c>&lt;</c> 会让每个起始位置都重扫剩余文本，
        /// 把行内解析退化成 O(N²) 停顿；本表把该成本压回线性。
        /// A pairing lookup table for one inline text. The first query for a character class scans the whole text once
        /// and caches the result, so every later position query is O(1). Rescanning from each position would make a run
        /// of unmatched <c>[</c>, <c>(</c> or <c>&lt;</c> revisit the remaining text every time, turning inline parsing
        /// into an O(N²) stall; the table keeps that cost linear.
        /// </summary>
        private sealed class ScanMap
        {
            /// <summary>本表对应的文本。The text this table covers.</summary>
            private readonly string _text;
            /// <summary>每个 <c>[</c> 位置匹配的 <c>]</c> 下标。The matching ']' index for every '['.</summary>
            private int[] _brackets;
            /// <summary>每个 <c>(</c> 位置匹配的 <c>)</c> 下标。The matching ')' index for every '('.</summary>
            private int[] _parens;
            /// <summary>按字符缓存的“下一个未转义出现位置”表。The cached "next unescaped occurrence" tables keyed by character.</summary>
            private Dictionary<char, int[]> _nexts;

            /// <summary>创建查找表。Creates the lookup table.</summary>
            /// <param name="text">对应文本。The text to cover.</param>
            public ScanMap(string text)
            {
                _text = text ?? string.Empty;
            }

            /// <summary>查找 open 处 <c>[</c> 匹配的 <c>]</c>。Finds the ']' matching the '[' at the open index.</summary>
            /// <param name="open">起始下标。The start index.</param>
            /// <returns>匹配位置或 -1。The matched index or -1.</returns>
            public int BracketClose(int open)
            {
                int[] pairs = _brackets ?? (_brackets = BuildPairs('[', ']'));
                return open >= 0 && open < pairs.Length ? pairs[open] : -1;
            }

            /// <summary>查找 open 处 <c>(</c> 匹配的 <c>)</c>。Finds the ')' matching the '(' at the open index.</summary>
            /// <param name="open">起始下标。The start index.</param>
            /// <returns>匹配位置或 -1。The matched index or -1.</returns>
            public int ParenClose(int open)
            {
                int[] pairs = _parens ?? (_parens = BuildPairs('(', ')'));
                return open >= 0 && open < pairs.Length ? pairs[open] : -1;
            }

            /// <summary>查找 from 起下一个未转义的 target。Finds the next unescaped target at or after the index.</summary>
            /// <param name="target">目标字符。The target character.</param>
            /// <param name="from">起始下标。The start index.</param>
            /// <returns>目标位置或 -1。The found index or -1.</returns>
            public int Next(char target, int from)
            {
                if (_nexts == null) _nexts = new Dictionary<char, int[]>();
                int[] next;
                if (!_nexts.TryGetValue(target, out next))
                {
                    next = BuildNext(target);
                    _nexts.Add(target, next);
                }
                return from >= 0 && from < next.Length ? next[from] : -1;
            }

            /// <summary>单趟栈配对：为每个开字符记录其配对的闭字符下标，未配对的保持 -1。One stack pass pairing every opener with its closer; unmatched entries stay -1.</summary>
            /// <param name="open">开字符。The opener.</param>
            /// <param name="close">闭字符。The closer.</param>
            /// <returns>逐位置的配对表。The per-position pairing table.</returns>
            private int[] BuildPairs(char open, char close)
            {
                int[] pairs = new int[_text.Length];
                for (int i = 0; i < pairs.Length; i++) pairs[i] = -1;
                Stack<int> opens = new Stack<int>();
                for (int i = 0; i < _text.Length; i++)
                {
                    char c = _text[i];
                    if (c == '\\')
                    {
                        i++;
                        continue;
                    }
                    if (c == open)
                    {
                        opens.Push(i);
                    }
                    else if (c == close && opens.Count > 0)
                    {
                        pairs[opens.Pop()] = i;
                    }
                }
                return pairs;
            }

            /// <summary>两趟构建后继表：正向标记未转义的目标位，反向回填“其后的第一个目标位”。Two passes build the successor table: the forward pass marks unescaped targets, the backward pass fills in the next target index.</summary>
            /// <param name="target">目标字符。The target character.</param>
            /// <returns>逐位置的后继表。The per-position successor table.</returns>
            private int[] BuildNext(char target)
            {
                int[] next = new int[_text.Length];
                for (int i = 0; i < next.Length; i++) next[i] = -1;
                int scan = 0;
                while (scan < next.Length)
                {
                    char c = _text[scan];
                    if (c == '\\')
                    {
                        scan += 2;
                        continue;
                    }
                    if (c == target) next[scan] = scan;
                    scan++;
                }
                int found = -1;
                for (int i = next.Length - 1; i >= 0; i--)
                {
                    if (next[i] >= 0) found = i;
                    else next[i] = found;
                }
                return next;
            }
        }
    }
}
