using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// HTML 读取器：把常见的语义化 HTML（标题、段落、表格、列表、强调、超链接、图片与内联 CSS）
    /// 解析为统一文档模型。出于安全考虑只读取 data: 内嵌图片，绝不请求网络资源，并限制节点数与嵌套深度。
    /// HTML reader that parses common semantic HTML (headings, paragraphs, tables, lists, emphasis,
    /// hyperlinks, images and inline CSS) into the unified document model. For safety it only decodes
    /// data: embedded images, never fetches network resources, and caps the node count and nesting depth.
    /// </summary>
    public static class HtmlReader
    {
        /// <summary>DOM 最大嵌套深度，超出后按兄弟节点处理。The maximum DOM nesting depth; deeper elements are treated as siblings.</summary>
        private const int MaxDepth = 256;
        /// <summary>单个文档最多创建的 DOM 节点数。The maximum number of DOM nodes created for one document.</summary>
        private const int MaxNodes = 500000;
        /// <summary>单张内嵌图片的最大字节数（32 MB）。The maximum byte size of one embedded image (32 MB).</summary>
        private const int MaxImageBytes = 32 * 1024 * 1024;
        /// <summary>单个文档最多读取的内嵌图片数。The maximum number of embedded images read from one document.</summary>
        private const int MaxImages = 512;

        /// <summary>从磁盘文件加载 HTML 文档（UTF-8，自动识别 BOM）。Loads an HTML document from a disk file (UTF-8 with BOM detection).</summary>
        /// <param name="path">HTML 文件路径。The HTML file path.</param>
        /// <returns>解析得到的文档模型。The parsed document model.</returns>
        public static OfficeDocument Load(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(File.ReadAllBytes(path));
        }

        /// <summary>从流加载 HTML 文档。Loads an HTML document from a stream.</summary>
        /// <param name="stream">包含 HTML 字节的流。A stream containing HTML bytes.</param>
        /// <returns>解析得到的文档模型。The parsed document model.</returns>
        public static OfficeDocument Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return Load(OfficePackageLimits.ReadAllBytes(stream));
        }

        /// <summary>从字节数组加载 HTML 文档，按 BOM 或 UTF-8 解码。Loads an HTML document from a byte array, decoding by BOM or UTF-8.</summary>
        /// <param name="data">HTML 字节内容。The HTML byte content.</param>
        /// <returns>解析得到的文档模型。The parsed document model.</returns>
        public static OfficeDocument Load(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            return Parse(DecodeText(data));
        }

        /// <summary>解析 HTML 文本为文档模型。Parses HTML text into the document model.</summary>
        /// <param name="html">HTML 文本。The HTML text.</param>
        /// <returns>解析得到的文档模型。The parsed document model.</returns>
        public static OfficeDocument Parse(string html)
        {
            if (html == null) throw new ArgumentNullException("html");
            Parser parser = new Parser(html);
            ElementNode root = parser.ParseDom(out string styleText);
            ElementNode body = FindFirst(root, "body") ?? root;
            OfficeDocument document = new OfficeDocument();

            ReadMetadata(root, document.Metadata);
            PageSetup pageDefaults = ParsePageCss(styleText);

            List<ElementNode> sections = new List<ElementNode>();
            CollectSections(body, sections, 0);
            if (sections.Count == 0)
            {
                Section section = document.AddSection(pageDefaults.Clone());
                ExtractHeaderFooter(body, section);
                ParseBlocks(body, section.Blocks, new Context(), 0);
            }
            else
            {
                for (int i = 0; i < sections.Count; i++)
                {
                    ElementNode element = sections[i];
                    PageSetup setup = ReadSectionPageSetup(element, pageDefaults);
                    Section section = document.AddSection(setup);
                    ExtractHeaderFooter(element, section);
                    ParseBlocks(element, section.Blocks, new Context(), 0);
                }
            }

            if (document.Sections.Count == 0) document.AddSection();
            return document;
        }

        /// <summary>解析过程中的共享上下文：图片计数与标题默认格式。Shared parsing context: the image counter and the inherited format.</summary>
        private sealed class Context
        {
            /// <summary>已读取的图片数量。The number of images read so far.</summary>
            public int ImageCount;
        }

        /// <summary>轻量 DOM 节点基类。The lightweight DOM node base class.</summary>
        private abstract class Node
        {
        }

        /// <summary>文本节点。A text node.</summary>
        private sealed class TextNode : Node
        {
            /// <summary>文本内容。The text content.</summary>
            public string Text;
        }

        /// <summary>元素节点。An element node.</summary>
        private sealed class ElementNode : Node
        {
            /// <summary>小写标签名。The lower-case tag name.</summary>
            public string Name;
            /// <summary>小写属性字典。The dictionary of lower-case attributes.</summary>
            public Dictionary<string, string> Attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            /// <summary>子节点列表。The child nodes.</summary>
            public List<Node> Children = new List<Node>();
        }

        /// <summary>容错的 HTML 词法/语法分析器：处理注释、脚本、样式、自闭合标签与常见实体。A tolerant HTML lexer/parser that handles comments, scripts, styles, self-closing tags and common entities.</summary>
        private sealed class Parser
        {
            /// <summary>待解析的 HTML 文本。The HTML text to parse.</summary>
            private readonly string _html;
            /// <summary>当前扫描位置。The current scan position.</summary>
            private int _position;
            /// <summary>已创建的节点数。The number of nodes created so far.</summary>
            private int _nodeCount;
            /// <summary>&lt;style&gt; 元素内的 CSS 文本。The CSS text inside style elements.</summary>
            private readonly StringBuilder _styleText = new StringBuilder();

            /// <summary>创建解析器。Creates the parser.</summary>
            /// <param name="html">HTML 文本。The HTML text.</param>
            public Parser(string html)
            {
                _html = html;
            }

            /// <summary>解析整个文档并返回根节点，同时输出收集到的 CSS 文本。Parses the whole document, returning the root node and the collected CSS text.</summary>
            /// <param name="styleText">收集到的 style 元素内 CSS。The CSS collected from style elements.</param>
            /// <returns>文档根节点。The document root node.</returns>
            public ElementNode ParseDom(out string styleText)
            {
                ElementNode root = new ElementNode { Name = "#document" };
                _nodeCount = 1;
                List<ElementNode> stack = new List<ElementNode>();
                stack.Add(root);
                StringBuilder text = new StringBuilder();

                while (_position < _html.Length)
                {
                    char value = _html[_position];
                    if (value != '<')
                    {
                        text.Append(value);
                        _position++;
                        continue;
                    }
                    FlushText(stack, text);

                    if (StartsWith("<!--", _position))
                    {
                        int end = _html.IndexOf("-->", _position + 4, StringComparison.Ordinal);
                        _position = end < 0 ? _html.Length : end + 3;
                        continue;
                    }
                    if (StartsWith("<![CDATA[", _position))
                    {
                        int end = _html.IndexOf("]]>", _position + 9, StringComparison.Ordinal);
                        string content = end < 0 ? string.Empty : _html.Substring(_position + 9, end - _position - 9);
                        if (content.Length > 0) AddNode(stack, new TextNode { Text = content });
                        _position = end < 0 ? _html.Length : end + 3;
                        continue;
                    }
                    if (StartsWith("<!", _position) || StartsWith("<?", _position))
                    {
                        int end = _html.IndexOf('>', _position);
                        _position = end < 0 ? _html.Length : end + 1;
                        continue;
                    }

                    if (StartsWith("</", _position))
                    {
                        int end = _html.IndexOf('>', _position);
                        if (end < 0)
                        {
                            _position = _html.Length;
                            break;
                        }
                        int nameStart = _position + 2;
                        int nameEnd = nameStart;
                        while (nameEnd < end && !char.IsWhiteSpace(_html[nameEnd]) && _html[nameEnd] != '/') nameEnd++;
                        string name = _html.Substring(nameStart, nameEnd - nameStart).Trim().ToLowerInvariant();
                        _position = end + 1;
                        for (int i = stack.Count - 1; i >= 1; i--)
                        {
                            if (string.Equals(stack[i].Name, name, StringComparison.Ordinal))
                            {
                                stack.RemoveRange(i, stack.Count - i);
                                break;
                            }
                        }
                        continue;
                    }

                    // 开始标签。Opening tag.
                    int tagStart = _position + 1;
                    int cursor = tagStart;
                    while (cursor < _html.Length && (char.IsLetterOrDigit(_html[cursor]) || _html[cursor] == ':' || _html[cursor] == '-' || _html[cursor] == '_'))
                    {
                        cursor++;
                    }
                    if (cursor == tagStart)
                    {
                        // 孤立的 "<" 当作文本。A stray "<" is treated as text.
                        text.Append('<');
                        _position++;
                        continue;
                    }
                    string tagName = _html.Substring(tagStart, cursor - tagStart).ToLowerInvariant();
                    ElementNode element = new ElementNode { Name = tagName };
                    bool selfClosing = ParseAttributes(element, ref cursor);
                    _position = cursor;
                    AddNode(stack, element);

                    if (!selfClosing && IsRawText(tagName))
                    {
                        string closing = "</" + tagName;
                        int end = _html.IndexOf(closing, _position, StringComparison.OrdinalIgnoreCase);
                        int contentEnd = end < 0 ? _html.Length : end;
                        if (tagName == "style" && contentEnd > _position)
                        {
                            _styleText.Append(_html, _position, contentEnd - _position).Append('\n');
                        }
                        _position = end < 0 ? _html.Length : end;
                        int closeEnd = _html.IndexOf('>', _position);
                        _position = closeEnd < 0 ? _html.Length : closeEnd + 1;
                        continue;
                    }
                    if (!selfClosing && !IsVoid(tagName) && stack.Count < MaxDepth)
                    {
                        stack.Add(element);
                    }
                }
                FlushText(stack, text);
                styleText = _styleText.ToString();
                return root;
            }

            /// <summary>把累积文本挂到当前元素。Attaches the accumulated text to the current element.</summary>
            private void FlushText(List<ElementNode> stack, StringBuilder text)
            {
                if (text.Length == 0) return;
                AddNode(stack, new TextNode { Text = DecodeEntities(text.ToString()) });
                text.Length = 0;
            }

            /// <summary>把节点挂到栈顶元素；超过节点上限后丢弃。Attaches a node to the top element, dropping it once the node cap is reached.</summary>
            private void AddNode(List<ElementNode> stack, Node node)
            {
                if (_nodeCount >= MaxNodes) return;
                stack[stack.Count - 1].Children.Add(node);
                _nodeCount++;
            }

            /// <summary>从当前位置解析属性列表，直到标签结束；返回是否自闭合。Parses the attribute list up to the end of the tag and returns whether it is self-closing.</summary>
            private bool ParseAttributes(ElementNode element, ref int cursor)
            {
                bool selfClosing = false;
                while (cursor < _html.Length)
                {
                    while (cursor < _html.Length && char.IsWhiteSpace(_html[cursor])) cursor++;
                    if (cursor >= _html.Length) break;
                    if (_html[cursor] == '>')
                    {
                        cursor++;
                        break;
                    }
                    if (_html[cursor] == '/' && cursor + 1 < _html.Length && _html[cursor + 1] == '>')
                    {
                        selfClosing = true;
                        cursor += 2;
                        break;
                    }
                    int nameStart = cursor;
                    while (cursor < _html.Length && !char.IsWhiteSpace(_html[cursor])
                        && _html[cursor] != '=' && _html[cursor] != '>' && _html[cursor] != '/')
                    {
                        cursor++;
                    }
                    if (cursor == nameStart)
                    {
                        cursor++;
                        continue;
                    }
                    string name = _html.Substring(nameStart, cursor - nameStart).ToLowerInvariant();
                    while (cursor < _html.Length && char.IsWhiteSpace(_html[cursor])) cursor++;
                    string value = string.Empty;
                    if (cursor < _html.Length && _html[cursor] == '=')
                    {
                        cursor++;
                        while (cursor < _html.Length && char.IsWhiteSpace(_html[cursor])) cursor++;
                        if (cursor < _html.Length && (_html[cursor] == '"' || _html[cursor] == '\''))
                        {
                            char quote = _html[cursor];
                            cursor++;
                            int valueStart = cursor;
                            while (cursor < _html.Length && _html[cursor] != quote) cursor++;
                            value = cursor > valueStart ? _html.Substring(valueStart, cursor - valueStart) : string.Empty;
                            if (cursor < _html.Length) cursor++;
                        }
                        else
                        {
                            int valueStart = cursor;
                            while (cursor < _html.Length && !char.IsWhiteSpace(_html[cursor]) && _html[cursor] != '>') cursor++;
                            value = _html.Substring(valueStart, cursor - valueStart);
                        }
                    }
                    if (!element.Attributes.ContainsKey(name))
                    {
                        element.Attributes[name] = DecodeEntities(value);
                    }
                }
                return selfClosing;
            }

            /// <summary>判断是否从指定位置开始出现给定文本（忽略大小写）。Checks whether the given text occurs at the position (case-insensitive).</summary>
            private bool StartsWith(string value, int position)
            {
                if (position + value.Length > _html.Length) return false;
                return string.Compare(_html, position, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
            }
        }

        /// <summary>判断标签是否为原始文本元素（内容不做标签解析）。Checks whether a tag is a raw-text element whose content is not parsed as tags.</summary>
        private static bool IsRawText(string name)
        {
            return name == "script" || name == "style";
        }

        /// <summary>判断标签是否为 HTML 空元素（不产生闭合标签）。Checks whether a tag is an HTML void element (no closing tag).</summary>
        private static bool IsVoid(string name)
        {
            switch (name)
            {
                case "area":
                case "base":
                case "br":
                case "col":
                case "embed":
                case "hr":
                case "img":
                case "input":
                case "link":
                case "meta":
                case "param":
                case "source":
                case "track":
                case "wbr":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>深度优先查找第一个指定标签名的元素。Finds the first element with the given tag name in depth-first order.</summary>
        private static ElementNode FindFirst(Node node, string name)
        {
            ElementNode element = node as ElementNode;
            if (element == null) return null;
            if (string.Equals(element.Name, name, StringComparison.Ordinal)) return element;
            for (int i = 0; i < element.Children.Count; i++)
            {
                ElementNode found = FindFirst(element.Children[i], name);
                if (found != null) return found;
            }
            return null;
        }

        /// <summary>查找直接包含"文档节"标记的 section 元素（类名含 wisdom-section 或带页面尺寸属性）。Collects section elements that mark a document section (a wisdom-section class or page-size data attributes).</summary>
        private static void CollectSections(ElementNode container, List<ElementNode> sections, int depth)
        {
            if (depth > 32) return;
            for (int i = 0; i < container.Children.Count; i++)
            {
                ElementNode child = container.Children[i] as ElementNode;
                if (child == null) continue;
                string className = GetAttribute(child, "class");
                bool isSection = string.Equals(child.Name, "section", StringComparison.Ordinal)
                    && (className.IndexOf("wisdom-section", StringComparison.OrdinalIgnoreCase) >= 0
                        || child.Attributes.ContainsKey("data-page-width"));
                if (isSection)
                {
                    sections.Add(child);
                    continue;
                }
                if (child.Name == "body" || child.Name == "div" || child.Name == "main" || child.Name == "article")
                {
                    CollectSections(child, sections, depth + 1);
                }
            }
        }

        /// <summary>读取节元素上的页面设置数据属性；缺失时回退到 @page 默认值。Reads page-setup data attributes from a section element, falling back to the @page defaults.</summary>
        private static PageSetup ReadSectionPageSetup(ElementNode element, PageSetup fallback)
        {
            PageSetup setup = fallback.Clone();
            double value;
            if (TryReadDoubleAttribute(element, "data-page-width", out value)) setup.Width = value;
            if (TryReadDoubleAttribute(element, "data-page-height", out value)) setup.Height = value;
            if (TryReadDoubleAttribute(element, "data-margin-left", out value)) setup.MarginLeft = value;
            if (TryReadDoubleAttribute(element, "data-margin-top", out value)) setup.MarginTop = value;
            if (TryReadDoubleAttribute(element, "data-margin-right", out value)) setup.MarginRight = value;
            if (TryReadDoubleAttribute(element, "data-margin-bottom", out value)) setup.MarginBottom = value;
            if (TryReadDoubleAttribute(element, "data-header-distance", out value)) setup.HeaderDistance = value;
            if (TryReadDoubleAttribute(element, "data-footer-distance", out value)) setup.FooterDistance = value;
            return setup;
        }

        /// <summary>读取数值属性。Reads a numeric attribute.</summary>
        private static bool TryReadDoubleAttribute(ElementNode element, string name, out double value)
        {
            value = 0;
            string text;
            if (!element.Attributes.TryGetValue(name, out text)) return false;
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        /// <summary>把节或正文里的直接 header/footer 子元素解析到页眉页脚，并从后续正文解析中排除。Parses direct header/footer children into the header/footer and excludes them from body parsing.</summary>
        private static void ExtractHeaderFooter(ElementNode container, Section section)
        {
            for (int i = 0; i < container.Children.Count; i++)
            {
                ElementNode child = container.Children[i] as ElementNode;
                if (child == null) continue;
                if (string.Equals(child.Name, "header", StringComparison.Ordinal))
                {
                    ParseBlocks(child, section.Header.Blocks, new Context(), 0);
                }
                else if (string.Equals(child.Name, "footer", StringComparison.Ordinal))
                {
                    ParseBlocks(child, section.Footer.Blocks, new Context(), 0);
                }
            }
        }

        /// <summary>解析块级内容：段落、标题、表格、列表与分页符。Parses block-level content: paragraphs, headings, tables, lists and page breaks.</summary>
        private static void ParseBlocks(ElementNode container, List<Block> blocks, Context context, int depth)
        {
            if (depth > 24) return;
            Paragraph pending = null;
            for (int i = 0; i < container.Children.Count; i++)
            {
                Node node = container.Children[i];
                TextNode textNode = node as TextNode;
                if (textNode != null)
                {
                    string text = NormalizeText(textNode.Text, false);
                    if (string.IsNullOrEmpty(text)) continue;
                    EnsureParagraph(ref pending, container);
                    AppendText(pending, text, new TextFormat());
                    continue;
                }

                ElementNode element = node as ElementNode;
                if (element == null) continue;
                if (element.Name == "header" || element.Name == "footer" || element.Name == "script" || element.Name == "style"
                    || element.Name == "meta" || element.Name == "link" || element.Name == "title")
                {
                    continue;
                }

                if (element.Name == "table")
                {
                    FlushParagraph(blocks, ref pending);
                    Table table = ParseTable(element, context, depth + 1);
                    if (table != null) blocks.Add(table);
                    continue;
                }
                if (element.Name == "hr" || IsPageBreakDiv(element))
                {
                    FlushParagraph(blocks, ref pending);
                    blocks.Add(new PageBreak());
                    continue;
                }
                if (element.Name == "br")
                {
                    EnsureParagraph(ref pending, container);
                    pending.AddLineBreak();
                    continue;
                }
                if (element.Name == "img")
                {
                    Inline image = CreateImageInline(element, context);
                    if (image != null)
                    {
                        EnsureParagraph(ref pending, container);
                        pending.Inlines.Add(image);
                    }
                    continue;
                }
                if (IsParagraphTag(element.Name))
                {
                    FlushParagraph(blocks, ref pending);
                    Paragraph paragraph = CreateParagraph(element);
                    AppendInlineChildren(element, paragraph, GetBaseFormat(element), null, context, false);
                    blocks.Add(paragraph);
                    continue;
                }
                if (IsBlockContainer(element) && ContainsBlockChild(element))
                {
                    FlushParagraph(blocks, ref pending);
                    ParseBlocks(element, blocks, context, depth + 1);
                    continue;
                }
                if (element.Name == "ul" || element.Name == "ol")
                {
                    FlushParagraph(blocks, ref pending);
                    ParseList(element, blocks, context, depth + 1);
                    continue;
                }
                if (IsInlineTag(element.Name))
                {
                    EnsureParagraph(ref pending, container);
                    AppendInlineElement(element, pending, GetBaseFormat(element), null, context, false);
                    continue;
                }

                // 其他容器标签按块容器递归。Other container tags are recursed into as block containers.
                if (element.Name == "div" || element.Name == "section" || element.Name == "article"
                    || element.Name == "blockquote" || element.Name == "center" || element.Name == "body"
                    || element.Name == "main" || element.Name == "figure" || element.Name == "figcaption"
                    || element.Name == "dl" || element.Name == "dd" || element.Name == "dt")
                {
                    FlushParagraph(blocks, ref pending);
                    ParseBlocks(element, blocks, context, depth + 1);
                }
                else
                {
                    // 未知标签按内联处理，保留其中的文本。Unknown tags are treated as inline, keeping their text.
                    EnsureParagraph(ref pending, container);
                    AppendInlineElement(element, pending, GetBaseFormat(element), null, context, false);
                }
            }
            FlushParagraph(blocks, ref pending);
        }

        /// <summary>确保存在待写入的段落，并用容器默认格式初始化。Ensures a pending paragraph exists, initialised with the container's default format.</summary>
        private static void EnsureParagraph(ref Paragraph paragraph, ElementNode container)
        {
            if (paragraph == null) paragraph = new Paragraph();
        }

        /// <summary>把待写入段落加入块列表。Appends the pending paragraph to the block list.</summary>
        private static void FlushParagraph(List<Block> blocks, ref Paragraph paragraph)
        {
            if (paragraph != null)
            {
                blocks.Add(paragraph);
                paragraph = null;
            }
        }

        /// <summary>向段落追加文本运行，自动处理空文本。Appends a text run to a paragraph, skipping empty text.</summary>
        private static void AppendText(Paragraph paragraph, string text, TextFormat format)
        {
            if (paragraph == null || string.IsNullOrEmpty(text)) return;
            paragraph.Inlines.Add(new TextRun(text, format));
        }

        /// <summary>判断标签是否直接生成段落。Checks whether a tag directly produces a paragraph.</summary>
        private static bool IsParagraphTag(string name)
        {
            return name == "p" || name == "h1" || name == "h2" || name == "h3" || name == "h4"
                || name == "h5" || name == "h6" || name == "pre";
        }

        /// <summary>判断标签是否适合作为块容器。Checks whether a tag can act as a block container.</summary>
        private static bool IsBlockContainer(ElementNode element)
        {
            switch (element.Name)
            {
                case "div":
                case "section":
                case "article":
                case "blockquote":
                case "center":
                case "figure":
                case "figcaption":
                case "main":
                case "body":
                case "li":
                case "dd":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>判断容器是否直接包含已知块级子元素。Checks whether a container directly contains a known block-level child.</summary>
        private static bool ContainsBlockChild(ElementNode element)
        {
            for (int i = 0; i < element.Children.Count; i++)
            {
                ElementNode child = element.Children[i] as ElementNode;
                if (child == null) continue;
                if (child.Name == "table" || child.Name == "p" || child.Name == "ul" || child.Name == "ol"
                    || IsParagraphTag(child.Name) || child.Name == "hr" || child.Name == "div")
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>判断标签是否按内联方式处理。Checks whether a tag is handled as inline content.</summary>
        private static bool IsInlineTag(string name)
        {
            switch (name)
            {
                case "a":
                case "span":
                case "strong":
                case "b":
                case "em":
                case "i":
                case "u":
                case "s":
                case "strike":
                case "del":
                case "ins":
                case "sub":
                case "sup":
                case "small":
                case "big":
                case "mark":
                case "code":
                case "tt":
                case "kbd":
                case "samp":
                case "var":
                case "font":
                case "label":
                case "abbr":
                case "cite":
                case "q":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>判断 div 是否为分页符（类名含 pagebreak 或 CSS 指定前后分页）。Checks whether a div represents a page break (a pagebreak class or a CSS break before/after).</summary>
        private static bool IsPageBreakDiv(ElementNode element)
        {
            if (element.Name != "div" && element.Name != "hr") return false;
            string className = GetAttribute(element, "class");
            if (className.IndexOf("pagebreak", StringComparison.OrdinalIgnoreCase) >= 0
                || className.IndexOf("page-break", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }
            Dictionary<string, string> style = ParseStyle(GetAttribute(element, "style"));
            string breakBefore = GetStyle(style, "page-break-before", "break-before");
            string breakAfter = GetStyle(style, "page-break-after", "break-after");
            return string.Equals(breakBefore, "always", StringComparison.OrdinalIgnoreCase)
                || string.Equals(breakBefore, "page", StringComparison.OrdinalIgnoreCase)
                || string.Equals(breakAfter, "always", StringComparison.OrdinalIgnoreCase)
                || string.Equals(breakAfter, "page", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>从 HTML 元素构造段落：对齐、间距、缩进、标题级别、样式名与分页属性。Builds a paragraph from an HTML element: alignment, spacing, indents, heading level, style name and page-break flags.</summary>
        private static Paragraph CreateParagraph(ElementNode element)
        {
            Paragraph paragraph = new Paragraph();
            Dictionary<string, string> style = ParseStyle(GetAttribute(element, "style"));
            switch (GetStyle(style, "text-align"))
            {
                case "center": paragraph.Alignment = ParagraphAlignment.Center; break;
                case "right": paragraph.Alignment = ParagraphAlignment.Right; break;
                case "justify": paragraph.Alignment = ParagraphAlignment.Justify; break;
            }
            string align = GetAttribute(element, "align");
            if (align == "center") paragraph.Alignment = ParagraphAlignment.Center;
            else if (align == "right") paragraph.Alignment = ParagraphAlignment.Right;
            else if (align == "justify") paragraph.Alignment = ParagraphAlignment.Justify;

            double value;
            if (TryGetMargin(style, "top", out value)) paragraph.SpaceBefore = value;
            if (TryGetMargin(style, "bottom", out value)) paragraph.SpaceAfter = value;
            if (TryGetMargin(style, "left", out value)) paragraph.LeftIndent = value;
            if (TryGetMargin(style, "right", out value)) paragraph.RightIndent = value;
            if (TryGetStyleLength(style, out value, "text-indent")) paragraph.FirstLineIndent = value;
            string lineHeight = GetStyle(style, "line-height");
            if (!string.IsNullOrEmpty(lineHeight))
            {
                if (lineHeight.EndsWith("%", StringComparison.Ordinal))
                {
                    double percent;
                    if (double.TryParse(lineHeight.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out percent))
                    {
                        paragraph.LineSpacing = percent / 100.0;
                    }
                }
                else if (lineHeight.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                {
                    paragraph.LineSpacing = WordFormatUtil.ParseLengthPoints(lineHeight, 11) / 11.0;
                }
                else
                {
                    double multiple;
                    if (double.TryParse(lineHeight, NumberStyles.Float, CultureInfo.InvariantCulture, out multiple)
                        && multiple > 0 && multiple < 10)
                    {
                        paragraph.LineSpacing = multiple;
                    }
                }
            }
            string breakBefore = GetStyle(style, "page-break-before", "break-before");
            if (string.Equals(breakBefore, "always", StringComparison.OrdinalIgnoreCase)
                || string.Equals(breakBefore, "page", StringComparison.OrdinalIgnoreCase))
            {
                paragraph.PageBreakBefore = true;
            }
            string breakAfter = GetStyle(style, "page-break-after", "break-after");
            if (string.Equals(breakAfter, "avoid", StringComparison.OrdinalIgnoreCase))
            {
                paragraph.KeepWithNext = true;
            }
            string styleName = GetAttribute(element, "data-style");
            if (!string.IsNullOrEmpty(styleName)) paragraph.StyleName = styleName;

            if (element.Name.Length == 2 && element.Name[0] == 'h'
                && element.Name[1] >= '1' && element.Name[1] <= '6')
            {
                paragraph.HeadingLevel = element.Name[1] - '0';
            }
            else if (element.Name == "pre")
            {
                paragraph.StyleName = paragraph.StyleName ?? "Preformatted";
            }
            return paragraph;
        }

        /// <summary>根据元素标签与内联样式计算段落级基础文本格式（标题默认加粗放大，代码块使用等宽字体）。Computes the paragraph-level base text format from the element tag and inline style (headings default to bold and larger, code blocks use a monospace font).</summary>
        private static TextFormat GetBaseFormat(ElementNode element)
        {
            TextFormat format = new TextFormat();
            if (element.Name.Length == 2 && element.Name[0] == 'h'
                && element.Name[1] >= '1' && element.Name[1] <= '6')
            {
                format = TextFormat.Heading(element.Name[1] - '0');
            }
            else if (element.Name == "pre" || element.Name == "code" || element.Name == "tt" || element.Name == "kbd")
            {
                format = new TextFormat("Courier New", 10);
            }
            ApplyInlineStyle(element, format);
            return format;
        }

        /// <summary>解析列表：每个 li 生成一个段落，前缀为项目符号或编号。Parses a list: each li becomes one paragraph prefixed with a bullet or number.</summary>
        private static void ParseList(ElementNode list, List<Block> blocks, Context context, int depth)
        {
            bool ordered = string.Equals(list.Name, "ol", StringComparison.Ordinal);
            int index = 1;
            for (int i = 0; i < list.Children.Count; i++)
            {
                ElementNode item = list.Children[i] as ElementNode;
                if (item == null || item.Name != "li") continue;
                if (ContainsBlockChild(item))
                {
                    ParseBlocks(item, blocks, context, depth + 1);
                    index++;
                    continue;
                }
                Paragraph paragraph = CreateParagraph(item);
                paragraph.ListMarker = ordered
                    ? index.ToString(CultureInfo.InvariantCulture) + "."
                    : "\u2022";
                AppendInlineChildren(item, paragraph, GetBaseFormat(item), null, context, false);
                blocks.Add(paragraph);
                index++;
            }
        }

        /// <summary>把元素的内联子节点写入段落，继承当前格式与链接。Appends an element's inline children to a paragraph, inheriting the current format and link.</summary>
        private static void AppendInlineChildren(ElementNode element, Paragraph paragraph, TextFormat inherited, string link, Context context, bool pre)
        {
            for (int i = 0; i < element.Children.Count; i++)
            {
                Node node = element.Children[i];
                TextNode textNode = node as TextNode;
                if (textNode != null)
                {
                    string text = NormalizeText(textNode.Text, pre);
                    if (text.Length == 0) continue;
                    if (IsWhitespace(text))
                    {
                        if (paragraph.Inlines.Count > 0) paragraph.Inlines.Add(new TextRun(" ", inherited));
                        continue;
                    }
                    TextRun run = new TextRun(text, inherited.Clone());
                    run.Link = link;
                    paragraph.Inlines.Add(run);
                    continue;
                }
                ElementNode child = node as ElementNode;
                if (child == null) continue;
                AppendInlineElement(child, paragraph, inherited, link, context, pre);
            }
        }

        /// <summary>把单个内联元素（含标签语义与 CSS）写入段落。Appends one inline element, combining tag semantics with CSS.</summary>
        private static void AppendInlineElement(ElementNode element, Paragraph paragraph, TextFormat inherited, string link, Context context, bool pre)
        {
            if (element.Name == "br")
            {
                paragraph.AddLineBreak();
                return;
            }
            if (element.Name == "img")
            {
                Inline image = CreateImageInline(element, context);
                if (image != null) paragraph.Inlines.Add(image);
                return;
            }
            if (element.Name == "script" || element.Name == "style") return;
            if (element.Name == "a")
            {
                string name = GetAttribute(element, "name");
                if (string.IsNullOrEmpty(name)) name = GetAttribute(element, "id");
                string href = GetAttribute(element, "href").Trim();
                if (!string.IsNullOrEmpty(name) && string.IsNullOrEmpty(href))
                {
                    paragraph.BookmarkName = name;
                }
                if (!string.IsNullOrEmpty(href) && (href.StartsWith("#", StringComparison.Ordinal)
                    || WordFormatUtil.IsSafeHyperlinkTarget(href)))
                {
                    TextFormat linkFormat = inherited.Clone();
                    linkFormat.Underline = true;
                    linkFormat.Color = OfficeColor.Blue;
                    AppendInlineChildren(element, paragraph, linkFormat, href, context, pre);
                }
                else
                {
                    AppendInlineChildren(element, paragraph, inherited.Clone(), null, context, pre);
                }
                return;
            }
            if (element.Name == "span")
            {
                string className = GetAttribute(element, "class");
                if (className.IndexOf("wisdom-list-marker", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    string marker = GetElementText(element).Trim();
                    if (marker.Length > 0) paragraph.ListMarker = marker;
                    return;
                }
                string field = GetAttribute(element, "data-field");
                if (!string.IsNullOrEmpty(field))
                {
                    TextFormat fieldFormat = inherited.Clone();
                    ApplyInlineStyle(element, fieldFormat);
                    paragraph.Inlines.Add(new FieldRun(field, GetElementText(element), fieldFormat));
                    return;
                }
            }

            TextFormat format = inherited.Clone();
            ApplyTagSemantics(element.Name, format);
            ApplyInlineStyle(element, format);
            AppendInlineChildren(element, paragraph, format, link, context, pre || element.Name == "pre");
        }

        /// <summary>应用 HTML 标签自带的语义格式（strong/b、em/i、u、s、sup/sub、code 等）。Applies the semantic formatting implied by an HTML tag (strong/b, em/i, u, s, sup/sub, code and the like).</summary>
        private static void ApplyTagSemantics(string name, TextFormat format)
        {
            switch (name)
            {
                case "strong":
                case "b":
                    format.Bold = true;
                    break;
                case "em":
                case "i":
                case "cite":
                    format.Italic = true;
                    break;
                case "u":
                case "ins":
                    format.Underline = true;
                    break;
                case "s":
                case "strike":
                case "del":
                    format.Strikeout = true;
                    break;
                case "sup":
                    format.Superscript = true;
                    format.Subscript = false;
                    break;
                case "sub":
                    format.Subscript = true;
                    format.Superscript = false;
                    break;
                case "code":
                case "tt":
                case "kbd":
                case "samp":
                    format.FontFamily = "Courier New";
                    format.AsciiFontFamily = "Courier New";
                    format.EastAsiaFontFamily = "Courier New";
                    break;
                case "small":
                    format.FontSize = format.FontSize * 0.85;
                    break;
                case "big":
                    format.FontSize = format.FontSize * 1.2;
                    break;
            }
        }

        /// <summary>解析并应用内联 CSS 与旧式 font 属性到文本格式。Parses and applies inline CSS plus legacy font attributes to a text format.</summary>
        private static void ApplyInlineStyle(ElementNode element, TextFormat format)
        {
            Dictionary<string, string> style = ParseStyle(GetAttribute(element, "style"));
            string fontFamily = GetStyle(style, "font-family");
            if (!string.IsNullOrEmpty(fontFamily))
            {
                string first = FirstFontFamily(fontFamily);
                if (first.Length > 0)
                {
                    format.FontFamily = first;
                    format.AsciiFontFamily = first;
                    format.EastAsiaFontFamily = first;
                }
            }
            string fontSize = GetStyle(style, "font-size");
            if (!string.IsNullOrEmpty(fontSize))
            {
                if (fontSize.EndsWith("%", StringComparison.Ordinal))
                {
                    double percent;
                    if (double.TryParse(fontSize.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out percent))
                    {
                        format.FontSize = format.FontSize * percent / 100.0;
                    }
                }
                else if (fontSize.EndsWith("em", StringComparison.OrdinalIgnoreCase))
                {
                    double multiple;
                    if (double.TryParse(fontSize.Substring(0, fontSize.Length - 2), NumberStyles.Float,
                        CultureInfo.InvariantCulture, out multiple))
                    {
                        format.FontSize = format.FontSize * multiple;
                    }
                }
                else
                {
                    format.FontSize = WordFormatUtil.ParseLengthPoints(fontSize, format.FontSize);
                }
            }
            string fontWeight = GetStyle(style, "font-weight");
            if (fontWeight == "bold" || fontWeight == "bolder" || fontWeight == "700"
                || fontWeight == "800" || fontWeight == "900")
            {
                format.Bold = true;
            }
            else if (fontWeight == "normal" || fontWeight == "lighter" || fontWeight == "400")
            {
                format.Bold = false;
            }
            string fontStyle = GetStyle(style, "font-style");
            if (fontStyle == "italic" || fontStyle == "oblique") format.Italic = true;
            else if (fontStyle == "normal") format.Italic = false;
            string decoration = GetStyle(style, "text-decoration", "text-decoration-line");
            if (decoration.IndexOf("underline", StringComparison.OrdinalIgnoreCase) >= 0) format.Underline = true;
            if (decoration.IndexOf("line-through", StringComparison.OrdinalIgnoreCase) >= 0) format.Strikeout = true;
            string color = GetStyle(style, "color");
            OfficeColor parsed;
            if (ParseColor(color, out parsed)) format.Color = parsed;
            string vertical = GetStyle(style, "vertical-align");
            if (vertical == "super")
            {
                format.Superscript = true;
                format.Subscript = false;
            }
            else if (vertical == "sub")
            {
                format.Subscript = true;
                format.Superscript = false;
            }

            string face = GetAttribute(element, "face");
            if (!string.IsNullOrEmpty(face))
            {
                format.FontFamily = face.Trim();
                format.AsciiFontFamily = face.Trim();
                format.EastAsiaFontFamily = face.Trim();
            }
            string colorAttribute = GetAttribute(element, "color");
            if (ParseColor(colorAttribute, out parsed)) format.Color = parsed;
            string sizeAttribute = GetAttribute(element, "size");
            int legacySize;
            if (int.TryParse(sizeAttribute, NumberStyles.Integer, CultureInfo.InvariantCulture, out legacySize))
            {
                double[] sizes = { 8, 10, 12, 14, 18, 24, 36 };
                if (legacySize >= 1 && legacySize <= 7) format.FontSize = sizes[legacySize - 1];
            }
        }

        /// <summary>取 CSS font-family 的第一个字体名并去掉引号。Takes the first CSS font-family name and strips quotes.</summary>
        private static string FirstFontFamily(string value)
        {
            int comma = value.IndexOf(',');
            string first = comma < 0 ? value : value.Substring(0, comma);
            first = first.Trim().Trim('\'', '"');
            return first.Replace("\\'", "'");
        }

        /// <summary>解析表格：列定义、行、单元格跨度、底色、纵向对齐与表头行。Parses a table: column definitions, rows, cell spans, shading, vertical alignment and header rows.</summary>
        private static Table ParseTable(ElementNode element, Context context, int depth)
        {
            if (depth > 24) return null;
            Table table = new Table();
            string tableStyle = GetAttribute(element, "style");
            Dictionary<string, string> style = ParseStyle(tableStyle);
            if (GetStyle(style, "border").StartsWith("none", StringComparison.OrdinalIgnoreCase))
            {
                table.Borders = TableBorders.None();
            }
            OfficeColor color;
            if (ParseColor(GetStyle(style, "background-color", "background"), out color))
            {
                // 表格底色在模型中不单独建模，写入首行单元格以外会失真，这里忽略。Table background has no dedicated model slot; ignore it.
            }

            List<ElementNode> columns = FindChildren(element, "col");
            List<ElementNode> colGroups = FindChildren(element, "colgroup");
            for (int i = 0; i < colGroups.Count; i++)
            {
                columns.AddRange(FindChildren(colGroups[i], "col"));
            }
            for (int i = 0; i < columns.Count; i++)
            {
                double width = 0;
                Dictionary<string, string> columnStyle = ParseStyle(GetAttribute(columns[i], "style"));
                TryGetStyleLength(columnStyle, out width, "width");
                table.Columns.Add(new TableColumn(width));
            }

            List<RowInfo> rowInfos = new List<RowInfo>();
            CollectRows(element, rowInfos, null, 0);
            int maxColumns = 0;
            for (int r = 0; r < rowInfos.Count; r++)
            {
                TableRow row = new TableRow();
                row.IsHeader = rowInfos[r].IsHeader;
                List<ElementNode> cells = rowInfos[r].Cells;
                for (int c = 0; c < cells.Count; c++)
                {
                    ElementNode cellElement = cells[c];
                    TableCell cell = new TableCell();
                    int span;
                    if (int.TryParse(GetAttribute(cellElement, "colspan"), out span) && span > 1) cell.ColSpan = span;
                    if (int.TryParse(GetAttribute(cellElement, "rowspan"), out span) && span > 1) cell.RowSpan = span;
                    Dictionary<string, string> cellStyle = ParseStyle(GetAttribute(cellElement, "style"));
                    OfficeColor background;
                    if (ParseColor(GetStyle(cellStyle, "background-color", "background"), out background))
                    {
                        cell.BackgroundColor = background;
                    }
                    else if (ParseColor(GetAttribute(cellElement, "bgcolor"), out background))
                    {
                        cell.BackgroundColor = background;
                    }
                    string vertical = GetStyle(cellStyle, "vertical-align");
                    if (vertical == "middle") cell.VerticalAlignment = VerticalAlignment.Middle;
                    else if (vertical == "bottom") cell.VerticalAlignment = VerticalAlignment.Bottom;
                    ParseCellBorders(cell, cellStyle);
                    ParseBlocks(cellElement, cell.Blocks, context, depth + 1);
                    if (cell.Blocks.Count == 0) cell.Blocks.Add(new Paragraph());
                    row.Cells.Add(cell);
                }
                if (!row.IsHeader && row.Cells.Count > 0)
                {
                    bool allHeader = true;
                    for (int c = 0; c < cells.Count; c++)
                    {
                        if (!string.Equals(cells[c].Name, "th", StringComparison.Ordinal))
                        {
                            allHeader = false;
                            break;
                        }
                    }
                    row.IsHeader = allHeader;
                }
                if (row.Cells.Count > 0) table.Rows.Add(row);
                maxColumns = Math.Max(maxColumns, CountEffectiveColumns(row));
            }
            if (table.Columns.Count == 0)
            {
                int columnCount = Math.Max(1, maxColumns);
                for (int i = 0; i < columnCount; i++) table.Columns.Add(new TableColumn());
            }
            return table;
        }

        /// <summary>表格行信息：行元素、单元格集合与是否为表头。Row information: the row element, its cells and whether it is a header row.</summary>
        private sealed class RowInfo
        {
            /// <summary>该行的单元格。The cells in the row.</summary>
            public List<ElementNode> Cells = new List<ElementNode>();
            /// <summary>是否位于 thead。Whether the row is inside thead.</summary>
            public bool IsHeader;
        }

        /// <summary>收集表格行，递归进入 thead/tbody/tfoot。Collects table rows, descending into thead/tbody/tfoot.</summary>
        private static void CollectRows(ElementNode container, List<RowInfo> rows, bool? header, int depth)
        {
            if (depth > 8) return;
            for (int i = 0; i < container.Children.Count; i++)
            {
                ElementNode child = container.Children[i] as ElementNode;
                if (child == null) continue;
                if (child.Name == "thead" || child.Name == "tbody" || child.Name == "tfoot")
                {
                    CollectRows(child, rows, header ?? child.Name == "thead" || header == true, depth + 1);
                    continue;
                }
                if (child.Name != "tr") continue;
                RowInfo info = new RowInfo();
                info.IsHeader = header == true;
                for (int c = 0; c < child.Children.Count; c++)
                {
                    ElementNode cell = child.Children[c] as ElementNode;
                    if (cell == null) continue;
                    if (cell.Name != "td" && cell.Name != "th") continue;
                    info.Cells.Add(cell);
                }
                rows.Add(info);
            }
        }

        /// <summary>统计一行按 colspan 展开后的有效列数。Counts a row's effective columns after expanding colspan.</summary>
        private static int CountEffectiveColumns(TableRow row)
        {
            int count = 0;
            for (int i = 0; i < row.Cells.Count; i++) count += Math.Max(1, row.Cells[i].ColSpan);
            return count;
        }

        /// <summary>解析单元格 CSS 四边边框。Parses per-edge cell borders from CSS.</summary>
        private static void ParseCellBorders(TableCell cell, Dictionary<string, string> style)
        {
            TableCellBorders borders = null;
            borders = AddBorder(borders, "top", GetStyle(style, "border-top", "border"));
            borders = AddBorder(borders, "bottom", GetStyle(style, "border-bottom", "border"));
            borders = AddBorder(borders, "left", GetStyle(style, "border-left", "border"));
            borders = AddBorder(borders, "right", GetStyle(style, "border-right", "border"));
            if (borders != null) cell.Borders = borders;
        }

        /// <summary>解析单条边框 CSS 文本并挂到指定边；无法解析时返回原集合。Parses one border CSS value and attaches it to the given edge, returning the original set when it cannot be parsed.</summary>
        private static TableCellBorders AddBorder(TableCellBorders borders, string edge, string value)
        {
            if (string.IsNullOrEmpty(value) || value.StartsWith("none", StringComparison.OrdinalIgnoreCase)) return borders;
            string[] parts = value.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            BorderStyle style = BorderStyle.Single;
            double width = 0.5;
            OfficeColor color = OfficeColor.Black;
            for (int i = 0; i < parts.Length; i++)
            {
                string part = parts[i].Trim();
                if (part == "dashed") style = BorderStyle.Dashed;
                else if (part == "dotted") style = BorderStyle.Dotted;
                else if (part == "double") style = BorderStyle.Double;
                else if (part == "solid") style = BorderStyle.Single;
                else if (part.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    || part.EndsWith("px", StringComparison.OrdinalIgnoreCase))
                {
                    width = WordFormatUtil.ParseLengthPoints(part, width);
                }
                else
                {
                    OfficeColor parsed;
                    if (ParseColor(part, out parsed)) color = parsed;
                }
            }
            if (borders == null) borders = new TableCellBorders();
            BorderLine line = new BorderLine(style, width, color);
            if (edge == "top") borders.Top = line;
            else if (edge == "bottom") borders.Bottom = line;
            else if (edge == "left") borders.Left = line;
            else borders.Right = line;
            return borders;
        }

        /// <summary>创建图片行内元素：只接受 data: 内嵌 PNG/JPEG，外部 URL 不请求、仅回退 alt 文本。Creates an image inline: only data: embedded PNG/JPEG is accepted; external URLs are never requested and only the alt text is kept.</summary>
        private static Inline CreateImageInline(ElementNode element, Context context)
        {
            string source = GetAttribute(element, "src").Trim();
            string alt = GetAttribute(element, "alt");
            if (context.ImageCount >= MaxImages || source.Length == 0 || !source.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                return string.IsNullOrEmpty(alt) ? null : new TextRun(alt);
            }
            int comma = source.IndexOf(',');
            if (comma < 0) return string.IsNullOrEmpty(alt) ? null : new TextRun(alt);
            string header = source.Substring(5, comma - 5).ToLowerInvariant();
            if (!header.EndsWith(";base64", StringComparison.Ordinal)) return string.IsNullOrEmpty(alt) ? null : new TextRun(alt);
            string mime = header.Substring(0, header.Length - 7).Trim();
            if (mime != "image/png" && mime != "image/jpeg" && mime != "image/jpg") return null;
            if (source.Length > MaxImageBytes * 2) return null;
            byte[] data;
            try
            {
                data = Convert.FromBase64String(source.Substring(comma + 1));
            }
            catch (FormatException)
            {
                return null;
            }
            if (data.Length == 0 || data.Length > MaxImageBytes) return null;
            context.ImageCount++;
            OfficeImage image = new OfficeImage(data);
            double width = ReadImageLength(element, "width");
            double height = ReadImageLength(element, "height");
            return new ImageRun(image, width, height);
        }

        /// <summary>读取图片的显示宽高（属性或 CSS），换算为磅。Reads an image's display width/height from attributes or CSS, converted to points.</summary>
        private static double ReadImageLength(ElementNode element, string name)
        {
            Dictionary<string, string> style = ParseStyle(GetAttribute(element, "style"));
            string value = GetStyle(style, name);
            if (!string.IsNullOrEmpty(value)) return WordFormatUtil.ParseLengthPoints(value, 0);
            value = GetAttribute(element, name);
            if (string.IsNullOrEmpty(value)) return 0;
            double pixels;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out pixels))
            {
                return pixels * 72.0 / 96.0;
            }
            return WordFormatUtil.ParseLengthPoints(value, 0);
        }

        /// <summary>读取文档元数据：title、author、description、keywords 与生成器。Reads document metadata: title, author, description, keywords and generator.</summary>
        private static void ReadMetadata(ElementNode root, DocumentMetadata metadata)
        {
            ElementNode title = FindFirst(root, "title");
            if (title != null) metadata.Title = GetElementText(title).Trim();
            List<ElementNode> metas = new List<ElementNode>();
            CollectElements(root, "meta", metas, 0);
            for (int i = 0; i < metas.Count; i++)
            {
                string name = GetAttribute(metas[i], "name").ToLowerInvariant();
                string content = GetAttribute(metas[i], "content");
                if (string.IsNullOrEmpty(content)) continue;
                if (name == "author") metadata.Author = content;
                else if (name == "description") metadata.Subject = content;
                else if (name == "keywords") metadata.Keywords = content;
                else if (name == "generator") metadata.Creator = content;
                else if (name == "dcterms.created" || name == "dcterms.modified")
                {
                    DateTime date;
                    if (DateTime.TryParse(content, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                    {
                        if (name == "dcterms.created") metadata.Created = date;
                        else metadata.Modified = date;
                    }
                }
            }
        }

        /// <summary>收集指定标签名的全部元素。Collects all elements with the given tag name.</summary>
        private static void CollectElements(Node node, string name, List<ElementNode> result, int depth)
        {
            if (depth > 32) return;
            ElementNode element = node as ElementNode;
            if (element == null) return;
            if (string.Equals(element.Name, name, StringComparison.Ordinal)) result.Add(element);
            for (int i = 0; i < element.Children.Count; i++)
            {
                CollectElements(element.Children[i], name, result, depth + 1);
            }
        }

        /// <summary>查找直接子元素中指定标签名的元素。Finds direct child elements with the given tag name.</summary>
        private static List<ElementNode> FindChildren(ElementNode container, string name)
        {
            List<ElementNode> result = new List<ElementNode>();
            for (int i = 0; i < container.Children.Count; i++)
            {
                ElementNode child = container.Children[i] as ElementNode;
                if (child != null && child.Name == name) result.Add(child);
            }
            return result;
        }

        /// <summary>提取元素内全部文本。Extracts all text inside an element.</summary>
        private static string GetElementText(Node node)
        {
            StringBuilder builder = new StringBuilder();
            AppendElementText(node, builder, 0);
            return builder.ToString();
        }

        /// <summary>递归收集文本。Recursively collects text.</summary>
        private static void AppendElementText(Node node, StringBuilder builder, int depth)
        {
            if (depth > 32) return;
            TextNode text = node as TextNode;
            if (text != null)
            {
                builder.Append(text.Text);
                return;
            }
            ElementNode element = node as ElementNode;
            if (element == null) return;
            if (element.Name == "br") builder.Append('\n');
            for (int i = 0; i < element.Children.Count; i++)
            {
                AppendElementText(element.Children[i], builder, depth + 1);
            }
        }

        /// <summary>读取属性值，未设置时返回空字符串。Reads an attribute value, returning an empty string when absent.</summary>
        private static string GetAttribute(ElementNode element, string name)
        {
            string value;
            return element.Attributes.TryGetValue(name, out value) ? value : string.Empty;
        }

        /// <summary>解析内联 style 属性为小写属性字典。Parses an inline style attribute into a lower-case property dictionary.</summary>
        private static Dictionary<string, string> ParseStyle(string style)
        {
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(style)) return result;
            string[] declarations = style.Split(';');
            for (int i = 0; i < declarations.Length; i++)
            {
                string declaration = declarations[i];
                int colon = declaration.IndexOf(':');
                if (colon <= 0) continue;
                string name = declaration.Substring(0, colon).Trim().ToLowerInvariant();
                string value = declaration.Substring(colon + 1).Trim();
                int important = value.IndexOf("!important", StringComparison.OrdinalIgnoreCase);
                if (important >= 0) value = value.Substring(0, important).Trim();
                if (name.Length > 0 && value.Length > 0) result[name] = value;
            }
            return result;
        }

        /// <summary>按顺序查找第一个存在的样式值。Gets the first style value present among the given property names.</summary>
        private static string GetStyle(Dictionary<string, string> style, params string[] names)
        {
            for (int i = 0; i < names.Length; i++)
            {
                string value;
                if (style.TryGetValue(names[i], out value)) return value.Trim();
            }
            return string.Empty;
        }

        /// <summary>读取样式长度并限制为非负；无法解析时返回 false。Reads a style length clamped to non-negative, returning false when it cannot be parsed.</summary>
        private static bool TryGetStyleLength(Dictionary<string, string> style, out double value, params string[] names)
        {
            value = 0;
            for (int i = 0; i < names.Length; i++)
            {
                string text = GetStyle(style, names[i]);
                if (string.IsNullOrEmpty(text)) continue;
                double parsed = WordFormatUtil.ParseLengthPoints(text, double.NaN);
                if (double.IsNaN(parsed)) continue;
                value = Math.Max(0, parsed);
                return true;
            }
            return false;
        }

        /// <summary>读取一条页边距：优先 margin-边，其次 margin 简写（1~4 个值按上右下左展开）。Reads one margin edge: prefers margin-edge, falling back to the margin shorthand (one to four values expanded as top/right/bottom/left).</summary>
        private static bool TryGetMargin(Dictionary<string, string> style, string edge, out double value)
        {
            value = 0;
            double parsed;
            if (TryGetStyleLength(style, out parsed, "margin-" + edge))
            {
                value = parsed;
                return true;
            }
            string shorthand = GetStyle(style, "margin");
            if (string.IsNullOrEmpty(shorthand)) return false;
            string[] parts = shorthand.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return false;
            string text;
            switch (parts.Length)
            {
                case 1: text = parts[0]; break;
                case 2: text = edge == "left" || edge == "right" ? parts[1] : parts[0]; break;
                case 3: text = edge == "left" || edge == "right" ? parts[1] : edge == "bottom" ? parts[2] : parts[0]; break;
                default:
                    text = edge == "top" ? parts[0] : edge == "right" ? parts[1] : edge == "bottom" ? parts[2] : parts[3];
                    break;
            }
            parsed = WordFormatUtil.ParseLengthPoints(text, double.NaN);
            if (double.IsNaN(parsed)) return false;
            value = Math.Max(0, parsed);
            return true;
        }

        /// <summary>解析颜色文本：支持 #rgb/#rrggbb、rgb()/rgba()、少量颜色名与 transparent。Parses color text: #rgb/#rrggbb, rgb()/rgba(), a few color names and transparent.</summary>
        private static bool ParseColor(string value, out OfficeColor color)
        {
            color = OfficeColor.Black;
            if (string.IsNullOrEmpty(value)) return false;
            string text = value.Trim();
            if (text == "transparent") return false;
            if (text.StartsWith("#", StringComparison.Ordinal))
            {
                return OfficeColor.TryParse(text, out color);
            }
            if (text.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
            {
                int open = text.IndexOf('(');
                int close = text.LastIndexOf(')');
                if (open < 0 || close <= open) return false;
                string[] parts = text.Substring(open + 1, close - open - 1).Split(',');
                if (parts.Length < 3) return false;
                byte r, g, b;
                byte a = 255;
                if (!TryChannel(parts[0], out r) || !TryChannel(parts[1], out g) || !TryChannel(parts[2], out b))
                {
                    return false;
                }
                if (parts.Length >= 4)
                {
                    double alpha;
                    if (double.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out alpha))
                    {
                        a = (byte)Math.Max(0, Math.Min(255, Math.Round(alpha * 255)));
                    }
                }
                color = new OfficeColor(a, r, g, b);
                return true;
            }
            switch (text.ToLowerInvariant())
            {
                case "black": color = OfficeColor.Black; return true;
                case "white": color = OfficeColor.White; return true;
                case "red": color = OfficeColor.Red; return true;
                case "green": color = OfficeColor.Green; return true;
                case "blue": color = OfficeColor.Blue; return true;
                case "gray":
                case "grey": color = OfficeColor.Gray; return true;
                case "silver": color = new OfficeColor(192, 192, 192); return true;
                case "lightgray":
                case "lightgrey": color = new OfficeColor(211, 211, 211); return true;
                case "maroon": color = new OfficeColor(128, 0, 0); return true;
                case "navy": color = new OfficeColor(0, 0, 128); return true;
                case "teal": color = new OfficeColor(0, 128, 128); return true;
                case "olive": color = new OfficeColor(128, 128, 0); return true;
                case "purple": color = new OfficeColor(128, 0, 128); return true;
                case "yellow": color = new OfficeColor(255, 255, 0); return true;
                case "lime": color = new OfficeColor(0, 255, 0); return true;
                case "aqua": color = new OfficeColor(0, 255, 255); return true;
                case "fuchsia": color = new OfficeColor(255, 0, 255); return true;
                case "orange": color = new OfficeColor(255, 165, 0); return true;
                default: return false;
            }
        }

        /// <summary>解析 rgb() 的单个通道：支持 0~255 整数与百分比。Parses one rgb() channel: an integer 0-255 or a percentage.</summary>
        private static bool TryChannel(string text, out byte value)
        {
            value = 0;
            string trimmed = text.Trim();
            if (trimmed.EndsWith("%", StringComparison.Ordinal))
            {
                double percent;
                if (!double.TryParse(trimmed.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out percent)) return false;
                value = (byte)Math.Max(0, Math.Min(255, Math.Round(percent * 2.55)));
                return true;
            }
            double number;
            if (!double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) return false;
            value = (byte)Math.Max(0, Math.Min(255, Math.Round(number)));
            return true;
        }

        /// <summary>解析 @page 规则的纸张尺寸与页边距；无法解析时返回 A4 默认值。Parses paper size and margins from the @page rule, returning A4 defaults when it cannot be parsed.</summary>
        private static PageSetup ParsePageCss(string css)
        {
            PageSetup setup = new PageSetup();
            if (string.IsNullOrEmpty(css)) return setup;
            int index = 0;
            while (index < css.Length)
            {
                int at = css.IndexOf("@page", index, StringComparison.OrdinalIgnoreCase);
                if (at < 0) break;
                int open = css.IndexOf('{', at);
                int close = open < 0 ? -1 : css.IndexOf('}', open);
                if (open < 0 || close < 0) break;
                string selector = css.Substring(at + 5, open - at - 5).Trim();
                string body = css.Substring(open + 1, close - open - 1);
                index = close + 1;
                if (selector.Length > 0) continue; // 具名页面规则不参与默认值。Named page rules are ignored for the default.

                int sizeAt = IndexOfProperty(body, "size");
                if (sizeAt >= 0)
                {
                    int colon = body.IndexOf(':', sizeAt);
                    int end = body.IndexOf(';', colon);
                    string size = end < 0 ? body.Substring(colon + 1) : body.Substring(colon + 1, end - colon - 1);
                    ApplyPageSize(setup, size.Trim());
                }
                int marginAt = IndexOfProperty(body, "margin");
                if (marginAt >= 0)
                {
                    int colon = body.IndexOf(':', marginAt);
                    int end = body.IndexOf(';', colon);
                    string margin = end < 0 ? body.Substring(colon + 1) : body.Substring(colon + 1, end - colon - 1);
                    ApplyMargins(setup, margin.Trim());
                }
                break;
            }
            return setup;
        }

        /// <summary>查找 CSS 声明名（确保是完整属性名而非另一个属性的后缀）。Finds a CSS declaration name, ensuring it is a complete property name rather than a suffix of another.</summary>
        private static int IndexOfProperty(string body, string name)
        {
            int index = 0;
            while (index < body.Length)
            {
                int found = body.IndexOf(name, index, StringComparison.OrdinalIgnoreCase);
                if (found < 0) return -1;
                int before = found - 1;
                while (before >= 0 && (body[before] == ' ' || body[before] == '\t' || body[before] == '\r' || body[before] == '\n')) before--;
                int after = found + name.Length;
                while (after < body.Length && (body[after] == ' ' || body[after] == '\t')) after++;
                if ((before < 0 || body[before] == ';') && after < body.Length && body[after] == ':') return found;
                index = found + name.Length;
            }
            return -1;
        }

        /// <summary>应用 @page size：支持两个长度、A4/Letter/A3/Legal 关键字。Applies @page size: two lengths or the A4/Letter/A3/Legal keywords.</summary>
        private static void ApplyPageSize(PageSetup setup, string size)
        {
            string[] parts = size.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 1)
            {
                string name = parts[0].Trim().ToLowerInvariant();
                if (name == "a4") { setup.Width = 595.28; setup.Height = 841.89; }
                else if (name == "a3") { setup.Width = 841.89; setup.Height = 1190.55; }
                else if (name == "letter") { setup.Width = 612; setup.Height = 792; }
                else if (name == "legal") { setup.Width = 612; setup.Height = 1008; }
                return;
            }
            if (parts.Length >= 2)
            {
                double width = WordFormatUtil.ParseLengthPoints(parts[0], setup.Width);
                double height = WordFormatUtil.ParseLengthPoints(parts[1], setup.Height);
                if (width > 0 && height > 0)
                {
                    setup.Width = width;
                    setup.Height = height;
                }
            }
        }

        /// <summary>应用 @page margin 简写（1~4 个值）。Applies the @page margin shorthand (one to four values).</summary>
        private static void ApplyMargins(PageSetup setup, string margin)
        {
            string[] parts = margin.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return;
            double top = WordFormatUtil.ParseLengthPoints(parts[0], setup.MarginTop);
            double right = parts.Length >= 2 ? WordFormatUtil.ParseLengthPoints(parts[1], top) : top;
            double bottom = parts.Length >= 3 ? WordFormatUtil.ParseLengthPoints(parts[2], top) : top;
            double left = parts.Length >= 4 ? WordFormatUtil.ParseLengthPoints(parts[3], right) : right;
            setup.MarginTop = Math.Max(0, top);
            setup.MarginRight = Math.Max(0, right);
            setup.MarginBottom = Math.Max(0, bottom);
            setup.MarginLeft = Math.Max(0, left);
        }

        /// <summary>规范化文本空白：非预格式内容把换行/制表折叠为单个空格，预格式内容只统一换行。Normalises text whitespace: non-preformatted content folds newlines/tabs into single spaces while preformatted content only normalises line endings.</summary>
        private static string NormalizeText(string text, bool pre)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            if (pre) return text.Replace("\r\n", "\n").Replace('\r', '\n');
            StringBuilder builder = new StringBuilder(text.Length);
            bool lastSpace = false;
            for (int i = 0; i < text.Length; i++)
            {
                char value = text[i];
                if (value == '\r' || value == '\n' || value == '\t' || value == '\f' || value == ' ')
                {
                    if (!lastSpace) builder.Append(' ');
                    lastSpace = true;
                    continue;
                }
                if (value < 32) continue;
                builder.Append(value);
                lastSpace = false;
            }
            return builder.ToString();
        }

        /// <summary>判断文本是否只含空白。Checks whether text consists solely of whitespace.</summary>
        private static bool IsWhitespace(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (!char.IsWhiteSpace(text[i])) return false;
            }
            return text.Length > 0;
        }

        /// <summary>解码 HTML 实体（命名实体与十进制/十六进制数字实体）；未知实体原样保留。Decodes HTML entities (named entities and decimal/hex numeric entities), leaving unknown entities unchanged.</summary>
        private static string DecodeEntities(string text)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOf('&') < 0) return text ?? string.Empty;
            StringBuilder builder = new StringBuilder(text.Length);
            int i = 0;
            while (i < text.Length)
            {
                char value = text[i];
                if (value != '&')
                {
                    builder.Append(value);
                    i++;
                    continue;
                }
                int semicolon = text.IndexOf(';', i + 1);
                if (semicolon < 0 || semicolon - i > 12)
                {
                    builder.Append('&');
                    i++;
                    continue;
                }
                string entity = text.Substring(i + 1, semicolon - i - 1);
                string replacement = DecodeEntity(entity);
                if (replacement == null)
                {
                    builder.Append('&');
                    i++;
                    continue;
                }
                builder.Append(replacement);
                i = semicolon + 1;
            }
            return builder.ToString();
        }

        /// <summary>解码单个实体名；无法识别时返回 null。Decodes a single entity name, returning null when unrecognized.</summary>
        private static string DecodeEntity(string entity)
        {
            if (entity.Length == 0) return null;
            if (entity[0] == '#')
            {
                int code;
                bool parsed = entity.Length > 1 && (entity[1] == 'x' || entity[1] == 'X')
                    ? int.TryParse(entity.Substring(2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code)
                    : int.TryParse(entity.Substring(1), NumberStyles.Integer, CultureInfo.InvariantCulture, out code);
                if (!parsed || code <= 0 || code > 0x10FFFF) return null;
                if (code >= 0xD800 && code <= 0xDFFF) return null;
                if (code <= 0xFFFF) return ((char)code).ToString();
                return char.ConvertFromUtf32(code);
            }
            switch (entity.ToLowerInvariant())
            {
                case "amp": return "&";
                case "lt": return "<";
                case "gt": return ">";
                case "quot": return "\"";
                case "apos": return "'";
                case "nbsp": return "\u00A0";
                case "copy": return "\u00A9";
                case "reg": return "\u00AE";
                case "trade": return "\u2122";
                case "ndash": return "\u2013";
                case "mdash": return "\u2014";
                case "hellip": return "\u2026";
                case "ldquo": return "\u201C";
                case "rdquo": return "\u201D";
                case "lsquo": return "\u2018";
                case "rsquo": return "\u2019";
                case "times": return "\u00D7";
                case "divide": return "\u00F7";
                case "euro": return "\u20AC";
                case "pound": return "\u00A3";
                case "yen": return "\u00A5";
                case "sect": return "\u00A7";
                case "deg": return "\u00B0";
                case "plusmn": return "\u00B1";
                case "middot": return "\u00B7";
                case "bull": return "\u2022";
                case "larr": return "\u2190";
                case "rarr": return "\u2192";
                case "uarr": return "\u2191";
                case "darr": return "\u2193";
                default: return null;
            }
        }

        /// <summary>按 BOM 或 UTF-8 解码 HTML 字节。Decodes HTML bytes by BOM or UTF-8.</summary>
        private static string DecodeText(byte[] data)
        {
            if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
            {
                return Encoding.UTF8.GetString(data, 3, data.Length - 3);
            }
            if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xFE)
            {
                return Encoding.Unicode.GetString(data, 2, data.Length - 2);
            }
            if (data.Length >= 2 && data[0] == 0xFE && data[1] == 0xFF)
            {
                return Encoding.BigEndianUnicode.GetString(data, 2, data.Length - 2);
            }
            return Encoding.UTF8.GetString(data);
        }
    }
}
