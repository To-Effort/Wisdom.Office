using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 读取 RTF 文档：解析字体表、颜色表、文档属性、分节页面设置、页眉页脚、段落与字符格式、
    /// 表格（含横向/纵向合并、表头行）、超链接域、书签与内嵌 PNG/JPEG 图片。
    /// Reads RTF documents: font and colour tables, document properties, per-section page setup, headers/footers,
    /// paragraph and character formatting, tables (with horizontal/vertical merges and header rows), hyperlink fields,
    /// bookmarks and embedded PNG/JPEG images.
    /// </summary>
    public static class RtfReader
    {
        /// <summary>允许的最大分组嵌套深度。The maximum permitted group nesting depth.</summary>
        private const int MaxDepth = 256;
        /// <summary>允许的最大节点（分组/控制字/文本）数量。The maximum permitted number of nodes (groups/control words/text runs).</summary>
        private const int MaxNodes = 4000000;
        /// <summary>单张图片的最大字节数。The maximum byte size of a single image.</summary>
        private const int MaxImageBytes = 64 * 1024 * 1024;
        /// <summary>允许的最大图片数量。The maximum permitted number of images.</summary>
        private const int MaxImages = 512;
        /// <summary>读取输入的最大字节数。The maximum number of input bytes read.</summary>
        private const long MaxInputBytes = 256L * 1024 * 1024;

        /// <summary>从磁盘读取 RTF 文件并载入文档模型。Reads an RTF file from disk and loads it into the document model.</summary>
        /// <param name="path">RTF 文件路径。The RTF file path.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(ReadFile(path));
        }

        /// <summary>从流读取 RTF 字节并载入文档模型；流可以不可定位。Reads RTF bytes from a stream and loads the document model; the stream need not be seekable.</summary>
        /// <param name="stream">源流。The source stream.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            using (MemoryStream buffer = new MemoryStream())
            {
                byte[] chunk = new byte[81920];
                long total = 0;
                int read;
                while ((read = stream.Read(chunk, 0, chunk.Length)) > 0)
                {
                    total += read;
                    if (total > MaxInputBytes) throw new InvalidDataException("RTF 输入超过大小限制。RTF input exceeds the size limit.");
                    buffer.Write(chunk, 0, read);
                }
                return Load(buffer.ToArray());
            }
        }

        /// <summary>从字节数组载入 RTF 文档，自动识别声明的 ANSI 代码页。Loads an RTF document from a byte array, honouring the declared ANSI code page.</summary>
        /// <param name="data">RTF 文件字节。The RTF file bytes.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length > MaxInputBytes) throw new InvalidDataException("RTF 输入超过大小限制。RTF input exceeds the size limit.");
            Encoding encoding = ResolveEncoding(DetectCodePage(data));
            string text = encoding.GetString(data);
            return Parse(text, encoding);
        }

        /// <summary>解析 RTF 文本并返回文档模型。Parses RTF text and returns the document model.</summary>
        /// <param name="rtf">RTF 文本。The RTF text.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Parse(string rtf)
        {
            if (rtf == null) throw new ArgumentNullException("rtf");
            return Parse(rtf, null);
        }

        /// <summary>解析 RTF 文本；encoding 为空时按 \ansicpg 声明解析 \'hh 字节转义。Parses RTF text; when encoding is null, \'hh byte escapes are decoded using the declared \ansicpg.</summary>
        private static OfficeDocument Parse(string rtf, Encoding encoding)
        {
            if (rtf == null) throw new ArgumentNullException("rtf");
            if (rtf.Length > MaxInputBytes) throw new InvalidDataException("RTF 输入超过大小限制。RTF input exceeds the size limit.");
            RtfGroup root = FindDocumentGroup(Tokenize(rtf));
            int codePage = FindCodePage(root);
            Encoding ansi = encoding ?? ResolveEncoding(codePage);
            DecodeByteNodes(root, ansi);

            ParseContext context = new ParseContext();
            context.Encoding = ansi;
            ParseFontTable(root, context);
            ParseColorTable(root, context);

            OfficeDocument document = new OfficeDocument();
            context.Document = document;
            ParseInfo(root, document.Metadata);
            ParseSections(root, context, document);
            return document;
        }

        /// <summary>读取磁盘文件，超过大小限制时拒绝。Reads a disk file, refusing input beyond the size limit.</summary>
        private static byte[] ReadFile(string path)
        {
            FileInfo info = new FileInfo(path);
            if (info.Length > MaxInputBytes) throw new InvalidDataException("RTF 输入超过大小限制。RTF input exceeds the size limit.");
            return File.ReadAllBytes(path);
        }

        // ---------------------------------------------------------------- tokenizer

        /// <summary>轻量 RTF 节点基类。Base class for lightweight RTF nodes.</summary>
        private abstract class RtfNode
        {
        }

        /// <summary>RTF 控制字或控制符号。An RTF control word or control symbol.</summary>
        private sealed class RtfControl : RtfNode
        {
            /// <summary>控制字名称（不含反斜杠）。The control word name (without the backslash).</summary>
            public string Name;
            /// <summary>可选数字参数。The optional numeric parameter.</summary>
            public int? Parameter;
        }

        /// <summary>RTF 普通文本。Plain RTF text.</summary>
        private sealed class RtfText : RtfNode
        {
            /// <summary>文本内容。The text content.</summary>
            public string Text;
        }

        /// <summary>由 \'hh 字节转义组成的原始字节串，稍后按 ANSI 代码页解码。A raw byte run made of \'hh escapes, decoded later using the ANSI code page.</summary>
        private sealed class RtfBytes : RtfNode
        {
            /// <summary>原始字节。The raw bytes.</summary>
            public List<byte> Bytes = new List<byte>();
        }

        /// <summary>RTF 分组。An RTF group.</summary>
        private sealed class RtfGroup : RtfNode
        {
            /// <summary>分组内的节点。The nodes within the group.</summary>
            public List<RtfNode> Items = new List<RtfNode>();
        }

        /// <summary>定位最外层的 \rtf 文档分组；多段文本或畸形输入时返回原始根分组。Locates the outermost \rtf document group; returns the original root for concatenated or malformed input.</summary>
        /// <param name="root">词法分析得到的根分组。The root group produced by tokenizing.</param>
        /// <returns>文档分组或原始根分组。The document group, or the original root group.</returns>
        private static RtfGroup FindDocumentGroup(RtfGroup root)
        {
            for (int i = 0; i < root.Items.Count; i++)
            {
                RtfGroup group = root.Items[i] as RtfGroup;
                if (group == null || group.Items.Count == 0) continue;
                RtfControl first = group.Items[0] as RtfControl;
                if (first != null && first.Name == "rtf") return group;
            }
            return root;
        }

        /// <summary>将 RTF 文本拆分为分组、控制字与文本节点。Splits RTF text into groups, control words and text nodes.</summary>
        private static RtfGroup Tokenize(string text)
        {
            RtfGroup root = new RtfGroup();
            Stack<RtfGroup> stack = new Stack<RtfGroup>();
            stack.Push(root);
            RtfBytes pendingBytes = null;
            int nodeCount = 0;
            int index = 0;
            while (index < text.Length)
            {
                char current = text[index];
                if (current == '{')
                {
                    FlushBytes(stack, ref pendingBytes);
                    if (stack.Count >= MaxDepth) throw new InvalidDataException("RTF 嵌套过深。RTF nesting is too deep.");
                    RtfGroup group = new RtfGroup();
                    AddNode(stack.Peek(), group, ref nodeCount);
                    stack.Push(group);
                    index++;
                    continue;
                }
                if (current == '}')
                {
                    FlushBytes(stack, ref pendingBytes);
                    if (stack.Count > 1) stack.Pop();
                    index++;
                    continue;
                }
                if (current == '\r' || current == '\n')
                {
                    index++;
                    continue;
                }
                if (current != '\\')
                {
                    int start = index;
                    while (index < text.Length)
                    {
                        char value = text[index];
                        if (value == '{' || value == '}' || value == '\\' || value == '\r' || value == '\n') break;
                        index++;
                    }
                    if (index > start)
                    {
                        AddNode(stack.Peek(), new RtfText { Text = text.Substring(start, index - start) }, ref nodeCount);
                    }
                    continue;
                }

                index++;
                if (index >= text.Length) break;
                char next = text[index];
                if (next == '\'')
                {
                    if (index + 2 >= text.Length)
                    {
                        index = text.Length;
                        break;
                    }
                    int high = HexValue(text[index + 1]);
                    int low = HexValue(text[index + 2]);
                    if (high < 0 || low < 0)
                    {
                        index += 3;
                        continue;
                    }
                    if (pendingBytes == null) pendingBytes = new RtfBytes();
                    pendingBytes.Bytes.Add((byte)((high << 4) | low));
                    index += 3;
                    continue;
                }

                if (char.IsLetter(next))
                {
                    FlushBytes(stack, ref pendingBytes);
                    int start = index;
                    while (index < text.Length && char.IsLetter(text[index])) index++;
                    string name = text.Substring(start, index - start);
                    int? parameter = null;
                    if (index < text.Length && (text[index] == '-' || char.IsDigit(text[index])))
                    {
                        bool negative = false;
                        if (text[index] == '-')
                        {
                            negative = true;
                            index++;
                        }
                        long value = 0;
                        bool any = false;
                        while (index < text.Length && char.IsDigit(text[index]))
                        {
                            value = value * 10 + (text[index] - '0');
                            if (value > int.MaxValue) value = int.MaxValue;
                            any = true;
                            index++;
                        }
                        if (any) parameter = (int)(negative ? -value : value);
                    }
                    if (index < text.Length && text[index] == ' ') index++;
                    AddNode(stack.Peek(), new RtfControl { Name = name, Parameter = parameter }, ref nodeCount);
                    if (name == "bin" && parameter.HasValue && parameter.Value > 0)
                    {
                        index += Math.Min(parameter.Value, text.Length - index);
                    }
                    continue;
                }

                FlushBytes(stack, ref pendingBytes);
                AddNode(stack.Peek(), new RtfControl { Name = next.ToString() }, ref nodeCount);
                index++;
            }
            FlushBytes(stack, ref pendingBytes);
            return root;
        }

        /// <summary>把累积的字节转义写入当前分组。Emits accumulated byte escapes into the current group.</summary>
        private static void FlushBytes(Stack<RtfGroup> stack, ref RtfBytes pending)
        {
            if (pending == null || pending.Bytes.Count == 0)
            {
                pending = null;
                return;
            }
            stack.Peek().Items.Add(pending);
            pending = null;
        }

        /// <summary>添加节点并检查数量上限。Adds a node while enforcing the node-count limit.</summary>
        private static void AddNode(RtfGroup group, RtfNode node, ref int count)
        {
            count++;
            if (count > MaxNodes) throw new InvalidDataException("RTF 节点数量超过限制。RTF node count exceeds the limit.");
            group.Items.Add(node);
        }

        /// <summary>十六进制字符转数值，非法字符返回 -1。Converts a hex digit to a value; returns -1 for invalid characters.</summary>
        private static int HexValue(char value)
        {
            if (value >= '0' && value <= '9') return value - '0';
            if (value >= 'a' && value <= 'f') return value - 'a' + 10;
            if (value >= 'A' && value <= 'F') return value - 'A' + 10;
            return -1;
        }

        /// <summary>按 \ansicpgN 声明扫描 ANSI 代码页；未声明时返回 1252。Scans the declared ANSI code page from \ansicpgN; returns 1252 when absent.</summary>
        private static int DetectCodePage(byte[] data)
        {
            int limit = Math.Min(data.Length, 4096);
            for (int i = 0; i + 8 < limit; i++)
            {
                if (data[i] == (byte)'\\' && data[i + 1] == (byte)'a' && data[i + 2] == (byte)'n'
                    && data[i + 3] == (byte)'s' && data[i + 4] == (byte)'i' && data[i + 5] == (byte)'c'
                    && data[i + 6] == (byte)'p' && data[i + 7] == (byte)'g')
                {
                    int value = 0;
                    bool any = false;
                    int position = i + 8;
                    while (position < limit && data[position] >= (byte)'0' && data[position] <= (byte)'9')
                    {
                        value = value * 10 + (data[position] - (byte)'0');
                        any = true;
                        position++;
                    }
                    if (any && value > 0 && value < 65536) return value;
                }
            }
            return 1252;
        }

        /// <summary>在解析树中查找 \ansicpg 参数。Finds the \ansicpg parameter in the parsed tree.</summary>
        private static int FindCodePage(RtfGroup root)
        {
            for (int i = 0; i < root.Items.Count; i++)
            {
                RtfControl control = root.Items[i] as RtfControl;
                if (control != null && control.Name == "ansicpg" && control.Parameter.HasValue
                    && control.Parameter.Value > 0 && control.Parameter.Value < 65536)
                {
                    return control.Parameter.Value;
                }
            }
            return 1252;
        }

        /// <summary>解析代码页对应的编码；无法获得时使用 Latin-1。Resolves the encoding for a code page, falling back to Latin-1.</summary>
        private static Encoding ResolveEncoding(int codePage)
        {
            if (codePage <= 0) codePage = 1252;
            try
            {
                return Encoding.GetEncoding(codePage);
            }
            catch (ArgumentException)
            {
            }
            catch (NotSupportedException)
            {
            }
            try
            {
                Type provider = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages", false);
                if (provider != null)
                {
                    PropertyInfo instanceProperty = provider.GetProperty("Instance", BindingFlags.Public | BindingFlags.Static);
                    MethodInfo register = typeof(Encoding).GetMethod("RegisterProvider", BindingFlags.Public | BindingFlags.Static);
                    if (instanceProperty != null && register != null)
                    {
                        register.Invoke(null, new object[] { instanceProperty.GetValue(null, null) });
                        return Encoding.GetEncoding(codePage);
                    }
                }
            }
            catch (Exception)
            {
            }
            return Encoding.GetEncoding(28591);
        }

        /// <summary>把树中的字节转义节点替换为按 ANSI 编码解码后的文本节点。Replaces byte-escape nodes in the tree with text nodes decoded using the ANSI encoding.</summary>
        private static void DecodeByteNodes(RtfGroup group, Encoding encoding)
        {
            for (int i = 0; i < group.Items.Count; i++)
            {
                RtfBytes bytes = group.Items[i] as RtfBytes;
                if (bytes != null)
                {
                    group.Items[i] = new RtfText { Text = encoding.GetString(bytes.Bytes.ToArray()) };
                    continue;
                }
                RtfGroup child = group.Items[i] as RtfGroup;
                if (child != null) DecodeByteNodes(child, encoding);
            }
        }

        // ---------------------------------------------------------------- resources and metadata

        /// <summary>解析上下文：共享的字体、颜色、编码与默认格式。Parse context: shared fonts, colours, encoding and default format.</summary>
        private sealed class ParseContext
        {
            /// <summary>字体索引到字体名的映射。Font index to font name map.</summary>
            public Dictionary<int, string> Fonts = new Dictionary<int, string>();
            /// <summary>颜色表，索引 0 为自动色。The colour table; index 0 is automatic.</summary>
            public List<OfficeColor> Colors = new List<OfficeColor> { OfficeColor.Black };
            /// <summary>ANSI 编码。The ANSI encoding.</summary>
            public Encoding Encoding;
            /// <summary>正在构建的文档。The document being built.</summary>
            public OfficeDocument Document;
        }

        /// <summary>查找指定目标的子分组（忽略大小写与可忽略标记）。Finds a child group with the given destination (case-insensitive, skipping the ignorable marker).</summary>
        private static RtfGroup FindGroup(RtfGroup parent, string destination)
        {
            for (int i = 0; i < parent.Items.Count; i++)
            {
                RtfGroup child = parent.Items[i] as RtfGroup;
                if (child == null) continue;
                if (string.Equals(GetDestinationName(child), destination, StringComparison.OrdinalIgnoreCase)) return child;
            }
            return null;
        }

        /// <summary>取得分组的首个有效控制字名称（跳过 \* 标记）。Gets the first significant control word name of a group (skipping the \* marker).</summary>
        private static string GetDestinationName(RtfGroup group)
        {
            for (int i = 0; i < group.Items.Count; i++)
            {
                RtfControl control = group.Items[i] as RtfControl;
                if (control == null) continue;
                if (control.Name == "*") continue;
                return control.Name;
            }
            return null;
        }

        /// <summary>判断分组是否以 \* 可忽略标记开头。Whether the group starts with the \* ignorable marker.</summary>
        private static bool HasIgnorableMarker(RtfGroup group)
        {
            if (group.Items.Count == 0) return false;
            RtfControl control = group.Items[0] as RtfControl;
            return control != null && control.Name == "*";
        }

        /// <summary>解析字体表，把 \fN 映射到字体名。Parses the font table, mapping \fN to font names.</summary>
        private static void ParseFontTable(RtfGroup root, ParseContext context)
        {
            RtfGroup fontTable = FindGroup(root, "fonttbl");
            if (fontTable == null) return;
            for (int i = 0; i < fontTable.Items.Count; i++)
            {
                RtfGroup entry = fontTable.Items[i] as RtfGroup;
                if (entry == null) continue;
                int index = 0;
                bool found = false;
                StringBuilder name = new StringBuilder();
                for (int n = 0; n < entry.Items.Count; n++)
                {
                    RtfControl control = entry.Items[n] as RtfControl;
                    if (control != null)
                    {
                        if (control.Name == "f" && control.Parameter.HasValue && !found)
                        {
                            index = control.Parameter.Value;
                            found = true;
                        }
                        continue;
                    }
                    RtfText text = entry.Items[n] as RtfText;
                    if (text != null) name.Append(text.Text);
                }
                if (!found) continue;
                string fontName = name.ToString().Trim();
                if (fontName.EndsWith(";", StringComparison.Ordinal)) fontName = fontName.Substring(0, fontName.Length - 1).Trim();
                if (fontName.Length > 0) context.Fonts[index] = fontName;
            }
        }

        /// <summary>解析颜色表；索引 0 保留为自动色。Parses the colour table; index 0 is reserved for the automatic colour.</summary>
        private static void ParseColorTable(RtfGroup root, ParseContext context)
        {
            RtfGroup colorTable = FindGroup(root, "colortbl");
            if (colorTable == null) return;
            int red = 0;
            int green = 0;
            int blue = 0;
            bool pending = false;
            for (int i = 0; i < colorTable.Items.Count; i++)
            {
                RtfControl control = colorTable.Items[i] as RtfControl;
                if (control != null)
                {
                    if (control.Name == "red" && control.Parameter.HasValue) { red = ClampColor(control.Parameter.Value); pending = true; }
                    else if (control.Name == "green" && control.Parameter.HasValue) { green = ClampColor(control.Parameter.Value); pending = true; }
                    else if (control.Name == "blue" && control.Parameter.HasValue) { blue = ClampColor(control.Parameter.Value); pending = true; }
                    else if ((control.Name == ";" || control.Name == "semicolon") && pending)
                    {
                        context.Colors.Add(new OfficeColor((byte)red, (byte)green, (byte)blue));
                        red = 0;
                        green = 0;
                        blue = 0;
                        pending = false;
                    }
                    continue;
                }
                RtfText text = colorTable.Items[i] as RtfText;
                if (text != null && pending && text.Text.IndexOf(';') >= 0)
                {
                    context.Colors.Add(new OfficeColor((byte)red, (byte)green, (byte)blue));
                    red = 0;
                    green = 0;
                    blue = 0;
                    pending = false;
                }
            }
            if (pending) context.Colors.Add(new OfficeColor((byte)red, (byte)green, (byte)blue));
        }

        /// <summary>把颜色分量限制在 0~255。Clamps a colour component to 0-255.</summary>
        private static int ClampColor(int value)
        {
            return value < 0 ? 0 : value > 255 ? 255 : value;
        }

        /// <summary>解析 \info 文档属性。Parses the \info document properties.</summary>
        private static void ParseInfo(RtfGroup root, DocumentMetadata metadata)
        {
            RtfGroup info = FindGroup(root, "info");
            if (info == null) return;
            for (int i = 0; i < info.Items.Count; i++)
            {
                RtfGroup entry = info.Items[i] as RtfGroup;
                if (entry == null) continue;
                string name = GetDestinationName(entry);
                if (name == null) continue;
                if (name == "title") metadata.Title = GetPlainText(entry).Trim();
                else if (name == "author") metadata.Author = GetPlainText(entry).Trim();
                else if (name == "subject") metadata.Subject = GetPlainText(entry).Trim();
                else if (name == "keywords") metadata.Keywords = GetPlainText(entry).Trim();
                else if (name == "creatim") metadata.Created = ReadInfoDate(entry, metadata.Created);
                else if (name == "revtim") metadata.Modified = ReadInfoDate(entry, metadata.Modified);
            }
        }

        /// <summary>读取信息组中的日期控制字。Reads the date control words in an info group.</summary>
        private static DateTime ReadInfoDate(RtfGroup group, DateTime fallback)
        {
            int year = 0;
            int month = 1;
            int day = 1;
            int hour = 0;
            int minute = 0;
            for (int i = 0; i < group.Items.Count; i++)
            {
                RtfControl control = group.Items[i] as RtfControl;
                if (control == null || !control.Parameter.HasValue) continue;
                switch (control.Name)
                {
                    case "yr": year = control.Parameter.Value; break;
                    case "mo": month = control.Parameter.Value; break;
                    case "dy": day = control.Parameter.Value; break;
                    case "hr": hour = control.Parameter.Value; break;
                    case "min": minute = control.Parameter.Value; break;
                }
            }
            if (year < 1 || year > 9999) return fallback;
            try
            {
                return new DateTime(year, Math.Min(12, Math.Max(1, month)), Math.Min(28, Math.Max(1, day)),
                    Math.Min(23, Math.Max(0, hour)), Math.Min(59, Math.Max(0, minute)), 0);
            }
            catch (ArgumentOutOfRangeException)
            {
                return fallback;
            }
        }

        /// <summary>提取分组中的纯文本：文本节点直连，控制字按段落/制表/换行语义转换，嵌套目标分组忽略。Extracts plain text from a group: text nodes pass through, control words map to paragraph/tab/line semantics and nested destinations are ignored.</summary>
        private static string GetPlainText(RtfGroup group)
        {
            StringBuilder builder = new StringBuilder();
            int pendingSkip = 0;
            AppendPlainText(group, builder, 0, 1, ref pendingSkip);
            return builder.ToString();
        }

        /// <summary>递归累积纯文本；unicodeSkip 为当前 \ucN 的替代字符数，pendingSkip 为 \u 之后尚未跳过的替代字符数。Recursively accumulates plain text; unicodeSkip is the active \ucN fallback width and pendingSkip counts fallback characters still to skip after \u.</summary>
        private static void AppendPlainText(RtfGroup group, StringBuilder builder, int depth, int unicodeSkip, ref int pendingSkip)
        {
            if (depth > MaxDepth) return;
            for (int i = 0; i < group.Items.Count; i++)
            {
                RtfText text = group.Items[i] as RtfText;
                if (text != null)
                {
                    int start = 0;
                    if (pendingSkip > 0)
                    {
                        start = Math.Min(pendingSkip, text.Text.Length);
                        pendingSkip -= start;
                    }
                    if (start < text.Text.Length) builder.Append(text.Text, start, text.Text.Length - start);
                    continue;
                }
                RtfControl control = group.Items[i] as RtfControl;
                if (control != null)
                {
                    if (control.Name == "uc" && control.Parameter.HasValue)
                    {
                        unicodeSkip = Math.Max(0, control.Parameter.Value);
                        pendingSkip = 0;
                    }
                    else if (control.Name == "par" || control.Name == "line") builder.Append('\n');
                    else if (control.Name == "tab") builder.Append('\t');
                    else if (control.Name == "u" && control.Parameter.HasValue)
                    {
                        int value = control.Parameter.Value;
                        if (value < 0) value += 65536;
                        builder.Append((char)value);
                        pendingSkip = unicodeSkip;
                    }
                    else if (control.Name == "\\") builder.Append('\\');
                    else if (control.Name == "{") builder.Append('{');
                    else if (control.Name == "}") builder.Append('}');
                    continue;
                }
                RtfGroup child = group.Items[i] as RtfGroup;
                if (child != null)
                {
                    int childSkip = pendingSkip;
                    AppendPlainText(child, builder, depth + 1, unicodeSkip, ref childSkip);
                    pendingSkip = childSkip;
                }
            }
        }

        // ---------------------------------------------------------------- sections

        /// <summary>把根分组拆分为若干节，并解析页眉页脚与正文。Splits the root group into sections and parses headers/footers and body content.</summary>
        private static void ParseSections(RtfGroup root, ParseContext context, OfficeDocument document)
        {
            Section section = new Section();
            PageSetup pageSetup = section.PageSetup;
            List<RtfNode> body = new List<RtfNode>();
            for (int i = 0; i < root.Items.Count; i++)
            {
                RtfNode node = root.Items[i];
                RtfControl control = node as RtfControl;
                if (control != null)
                {
                    if (control.Name == "sect")
                    {
                        FlushBody(body, context, section);
                        section.PageSetup = pageSetup;
                        document.Sections.Add(section);
                        section = new Section(pageSetup.Clone());
                        pageSetup = section.PageSetup;
                        continue;
                    }
                    if (ApplyPageSetup(control, pageSetup)) continue;
                    if (control.Name == "deff") continue;
                }

                RtfGroup group = node as RtfGroup;
                if (group != null)
                {
                    string destination = GetDestinationName(group);
                    if (destination == "header")
                    {
                        new ContentParser(context, section.Header.Blocks).Parse(group.Items);
                        continue;
                    }
                    if (destination == "footer")
                    {
                        new ContentParser(context, section.Footer.Blocks).Parse(group.Items);
                        continue;
                    }
                    if (IsResourceDestination(destination)) continue;
                }

                body.Add(node);
            }

            FlushBody(body, context, section);
            section.PageSetup = pageSetup;
            if (section.Blocks.Count > 0 || section.Header.IsEnabled || section.Footer.IsEnabled || document.Sections.Count == 0)
            {
                document.Sections.Add(section);
            }
            if (document.Sections.Count == 0) document.AddSection();
        }

        /// <summary>将当前节收集到的正文节点交给内容解析器，并清空缓冲。Hands the body nodes collected for the current section to a content parser and clears the buffer.</summary>
        private static void FlushBody(List<RtfNode> body, ParseContext context, Section section)
        {
            if (body.Count == 0) return;
            new ContentParser(context, section.Blocks).Parse(body);
            body.Clear();
        }

        /// <summary>判断是否为无需进入正文的资源分组。Whether the destination is a resource group that should not enter body content.</summary>
        private static bool IsResourceDestination(string destination)
        {
            if (destination == null) return false;
            switch (destination.ToLowerInvariant())
            {
                case "fonttbl":
                case "colortbl":
                case "stylesheet":
                case "info":
                case "listtable":
                case "listoverridetable":
                case "generator":
                case "xmlnstbl":
                case "themedata":
                case "colorschememapping":
                case "datastore":
                case "latentstyles":
                case "rsidtbl":
                case "revtbl":
                case "filetbl":
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>应用页面设置控制字；返回是否已消费该控制字。Applies a page-setup control word; returns whether the control word was consumed.</summary>
        private static bool ApplyPageSetup(RtfControl control, PageSetup setup)
        {
            if (!control.Parameter.HasValue) return false;
            double points = control.Parameter.Value / 20.0;
            switch (control.Name)
            {
                case "paperw":
                    if (points > 0) setup.Width = points;
                    return true;
                case "paperh":
                    if (points > 0) setup.Height = points;
                    return true;
                case "margl":
                    if (points >= 0) setup.MarginLeft = points;
                    return true;
                case "margr":
                    if (points >= 0) setup.MarginRight = points;
                    return true;
                case "margt":
                    if (points >= 0) setup.MarginTop = points;
                    return true;
                case "margb":
                    if (points >= 0) setup.MarginBottom = points;
                    return true;
                case "headery":
                    if (points >= 0) setup.HeaderDistance = points;
                    return true;
                case "footery":
                    if (points >= 0) setup.FooterDistance = points;
                    return true;
                default:
                    return false;
            }
        }

        // ---------------------------------------------------------------- content parser

        /// <summary>段落属性累积器。Accumulates paragraph properties.</summary>
        private sealed class ParagraphProperties
        {
            /// <summary>水平对齐。Horizontal alignment.</summary>
            public ParagraphAlignment Alignment = ParagraphAlignment.Left;
            /// <summary>段前间距（磅）。Space before, in points.</summary>
            public double SpaceBefore;
            /// <summary>段后间距（磅）。Space after, in points.</summary>
            public double SpaceAfter;
            /// <summary>左缩进（磅）。Left indent, in points.</summary>
            public double LeftIndent;
            /// <summary>右缩进（磅）。Right indent, in points.</summary>
            public double RightIndent;
            /// <summary>首行缩进（磅）。First-line indent, in points.</summary>
            public double FirstLineIndent;
            /// <summary>行距倍数。Line-spacing multiple.</summary>
            public double LineSpacing;
            /// <summary>标题级别。Heading level.</summary>
            public int HeadingLevel;
            /// <summary>与下段同页。Keep with next.</summary>
            public bool KeepWithNext;
            /// <summary>段前分页。Page break before.</summary>
            public bool PageBreakBefore;

            /// <summary>返回属性副本。Returns a copy of the properties.</summary>
            public ParagraphProperties Clone()
            {
                return (ParagraphProperties)MemberwiseClone();
            }
        }

        /// <summary>表格物理单元格的定义与属性。Definition and properties of one physical table cell.</summary>
        private sealed class CellSlot
        {
            /// <summary>右边界位置（缇）。The right boundary position (twips).</summary>
            public int Boundary;
            /// <summary>横向合并的起始单元格。The first cell of a horizontal merge.</summary>
            public bool HorizontalFirst;
            /// <summary>横向合并的后续单元格。A continuation cell of a horizontal merge.</summary>
            public bool HorizontalContinuation;
            /// <summary>纵向合并的起始单元格。The first cell of a vertical merge.</summary>
            public bool VerticalFirst;
            /// <summary>纵向合并的后续单元格。A continuation cell of a vertical merge.</summary>
            public bool VerticalContinuation;
            /// <summary>底色在颜色表中的索引。The shading colour index in the colour table.</summary>
            public int? BackgroundIndex;
            /// <summary>纵向对齐。Vertical alignment.</summary>
            public VerticalAlignment VerticalAlignment = VerticalAlignment.Top;
            /// <summary>上边框。The top border.</summary>
            public BorderLine TopBorder;
            /// <summary>下边框。The bottom border.</summary>
            public BorderLine BottomBorder;
            /// <summary>左边框。The left border.</summary>
            public BorderLine LeftBorder;
            /// <summary>右边框。The right border.</summary>
            public BorderLine RightBorder;
        }

        /// <summary>解析一段块级内容：正文、页眉/页脚或表格单元格内部。Parses one block-content destination: body, header/footer or a table cell interior.</summary>
        private sealed class ContentParser
        {
            private readonly ParseContext _context;
            private readonly List<Block> _rootTarget;

            private TextFormat _format;
            private ParagraphProperties _props = new ParagraphProperties();
            private List<Block> _blocks;
            private Paragraph _paragraph;
            private int _ucSkip = 1;
            private int _skipChars;
            private bool _afterPageBreak;

            private Table _table;
            private TableRow _row;
            private List<CellSlot> _slots;
            private CellSlot _slot;
            private List<TableCell> _physicalCells;
            private TableCell _cell;
            private List<Block> _outerBlocks;
            private int _cellIndex;
            private TableCell[] _previousGrid = new TableCell[0];
            private int _columnCount;
            private int _imageCount;

            /// <summary>以目标块列表创建内容解析器。Creates a content parser writing into the target block list.</summary>
            public ContentParser(ParseContext context, List<Block> target)
            {
                _context = context;
                _rootTarget = target;
                _blocks = target;
                _format = new TextFormat();
            }

            /// <summary>解析节点列表并结束最后一个段落。Parses a node list and finishes the trailing paragraph.</summary>
            public void Parse(List<RtfNode> items)
            {
                ProcessNodes(items, null, 0);
                FlushTable();
                FlushParagraph();
            }

            /// <summary>处理节点列表；格式与段落属性在分组范围内保存并恢复。Processes nodes; formatting and paragraph properties are saved and restored within a group scope.</summary>
            private void ProcessNodes(List<RtfNode> items, string link, int depth)
            {
                if (depth > MaxDepth) return;
                TextFormat savedFormat = _format.Clone();
                ParagraphProperties savedProps = _props.Clone();
                for (int i = 0; i < items.Count; i++)
                {
                    RtfText text = items[i] as RtfText;
                    if (text != null)
                    {
                        AppendText(text.Text, link);
                        continue;
                    }
                    RtfControl control = items[i] as RtfControl;
                    if (control != null)
                    {
                        HandleControl(control, link);
                        continue;
                    }
                    RtfGroup group = items[i] as RtfGroup;
                    if (group != null) HandleGroup(group, link, depth + 1);
                }
                _format = savedFormat;
                _props = savedProps;
            }

            // ------------------------------------------------------------ paragraphs and runs

            /// <summary>创建当前段落（若尚未创建）。Ensures the current paragraph exists.</summary>
            private Paragraph EnsureParagraph()
            {
                if (_table != null && ReferenceEquals(_blocks, _outerBlocks))
                {
                    FlushTable();
                }
                if (_paragraph == null)
                {
                    _paragraph = new Paragraph();
                    ApplyParagraphProperties(_paragraph, _props);
                    _blocks.Add(_paragraph);
                }
                return _paragraph;
            }

            /// <summary>把累积的段落属性写入段落。Applies accumulated paragraph properties to a paragraph.</summary>
            private static void ApplyParagraphProperties(Paragraph paragraph, ParagraphProperties props)
            {
                paragraph.Alignment = props.Alignment;
                paragraph.SpaceBefore = props.SpaceBefore;
                paragraph.SpaceAfter = props.SpaceAfter;
                paragraph.LeftIndent = props.LeftIndent;
                paragraph.RightIndent = props.RightIndent;
                paragraph.FirstLineIndent = props.FirstLineIndent;
                paragraph.LineSpacing = props.LineSpacing > 0 ? props.LineSpacing : 1;
                paragraph.HeadingLevel = props.HeadingLevel;
                paragraph.KeepWithNext = props.KeepWithNext;
                paragraph.PageBreakBefore = props.PageBreakBefore;
            }

            /// <summary>结束当前段落。Finishes the current paragraph.</summary>
            private void FlushParagraph()
            {
                _paragraph = null;
                _skipChars = 0;
            }

            /// <summary>追加文本，按需跳过 \uN 的 ANSI 回退字符，并合并同格式相邻文本。Appends text, skipping ANSI fallback characters of \uN as needed and merging adjacent text with identical formatting.</summary>
            private void AppendText(string text, string link)
            {
                if (string.IsNullOrEmpty(text)) return;
                if (_skipChars > 0)
                {
                    int skip = Math.Min(_skipChars, text.Length);
                    _skipChars -= skip;
                    if (skip >= text.Length) return;
                    text = text.Substring(skip);
                }
                _afterPageBreak = false;
                Paragraph paragraph = EnsureParagraph();
                if (paragraph.Inlines.Count > 0)
                {
                    TextRun previous = paragraph.Inlines[paragraph.Inlines.Count - 1] as TextRun;
                    if (previous != null && previous.Link == link && SameFormat(previous.Format, _format))
                    {
                        previous.Text += text;
                        return;
                    }
                }
                paragraph.Inlines.Add(new TextRun(text, _format.Clone()) { Link = link });
            }

            /// <summary>比较两个文本格式是否等价，用于合并相邻运行。Compares two text formats for equivalence, used to merge adjacent runs.</summary>
            private static bool SameFormat(TextFormat left, TextFormat right)
            {
                if (left == null || right == null) return left == right;
                return left.GetAsciiFontFamily() == right.GetAsciiFontFamily()
                    && left.GetEastAsiaFontFamily() == right.GetEastAsiaFontFamily()
                    && left.FontSize == right.FontSize
                    && left.Bold == right.Bold
                    && left.Italic == right.Italic
                    && left.Underline == right.Underline
                    && left.Strikeout == right.Strikeout
                    && left.Superscript == right.Superscript
                    && left.Subscript == right.Subscript
                    && left.Color.R == right.Color.R
                    && left.Color.G == right.Color.G
                    && left.Color.B == right.Color.B
                    && left.Color.A == right.Color.A;
            }

            /// <summary>把字体索引应用到当前字符格式。Applies a font index to the current character format.</summary>
            private void ApplyFont(int index)
            {
                string name;
                if (_context.Fonts.TryGetValue(index, out name))
                {
                    _format.FontFamily = name;
                    if (WordFormatUtil.ContainsEastAsian(name)) _format.EastAsiaFontFamily = name;
                    else _format.AsciiFontFamily = name;
                }
            }

            /// <summary>取得颜色表中的颜色；索引越界时返回黑色。Gets a colour from the colour table; returns black when out of range.</summary>
            private OfficeColor GetColor(int index)
            {
                if (index > 0 && index < _context.Colors.Count) return _context.Colors[index];
                return OfficeColor.Black;
            }

            // ------------------------------------------------------------ control words

            /// <summary>处理单个控制字或控制符号。Handles one control word or control symbol.</summary>
            private void HandleControl(RtfControl control, string link)
            {
                switch (control.Name)
                {
                    case "par":
                        if (_paragraph == null && _afterPageBreak)
                        {
                            _afterPageBreak = false;
                            return;
                        }
                        EnsureParagraph();
                        FlushParagraph();
                        return;
                    case "pard":
                        _props = new ParagraphProperties();
                        return;
                    case "plain":
                        _format = new TextFormat();
                        return;
                    case "ql": _props.Alignment = ParagraphAlignment.Left; return;
                    case "qc": _props.Alignment = ParagraphAlignment.Center; return;
                    case "qr": _props.Alignment = ParagraphAlignment.Right; return;
                    case "qj": _props.Alignment = ParagraphAlignment.Justify; return;
                    case "sb": _props.SpaceBefore = Twips(control); return;
                    case "sa": _props.SpaceAfter = Twips(control); return;
                    case "li": _props.LeftIndent = Twips(control); return;
                    case "ri": _props.RightIndent = Twips(control); return;
                    case "fi": _props.FirstLineIndent = Twips(control); return;
                    case "sl":
                        if (control.Parameter.HasValue && control.Parameter.Value > 0) _props.LineSpacing = control.Parameter.Value / 240.0;
                        return;
                    case "slmult":
                        return;
                    case "outlinelevel":
                        _props.HeadingLevel = control.Parameter.HasValue ? Math.Max(0, control.Parameter.Value) + 1 : 0;
                        return;
                    case "keepn":
                        _props.KeepWithNext = control.Parameter.GetValueOrDefault() != 0;
                        return;
                    case "pagebb":
                        _props.PageBreakBefore = control.Parameter.GetValueOrDefault() != 0;
                        return;
                    case "page":
                        if (_table == null)
                        {
                            if (_paragraph != null) FlushParagraph();
                            _blocks.Add(new PageBreak());
                            _afterPageBreak = true;
                        }
                        return;
                    case "f":
                        if (control.Parameter.HasValue) ApplyFont(control.Parameter.Value);
                        return;
                    case "fs":
                        if (control.Parameter.HasValue && control.Parameter.Value > 0) _format.FontSize = control.Parameter.Value / 2.0;
                        return;
                    case "cf":
                        if (control.Parameter.HasValue) _format.Color = GetColor(control.Parameter.Value);
                        return;
                    case "b":
                        _format.Bold = control.Parameter.GetValueOrDefault() != 0;
                        return;
                    case "i":
                        _format.Italic = control.Parameter.GetValueOrDefault() != 0;
                        return;
                    case "ul":
                        _format.Underline = control.Parameter.GetValueOrDefault() != 0;
                        return;
                    case "ulnone":
                        _format.Underline = false;
                        return;
                    case "strike":
                    case "striked":
                        _format.Strikeout = control.Parameter.GetValueOrDefault() != 0;
                        return;
                    case "super":
                        _format.Superscript = true;
                        _format.Subscript = false;
                        return;
                    case "sub":
                        _format.Subscript = true;
                        _format.Superscript = false;
                        return;
                    case "nosupersub":
                        _format.Superscript = false;
                        _format.Subscript = false;
                        return;
                    case "u":
                        if (control.Parameter.HasValue)
                        {
                            int value = control.Parameter.Value;
                            if (value < 0) value += 65536;
                            AppendText(((char)value).ToString(), link);
                            _skipChars = _ucSkip;
                        }
                        return;
                    case "uc":
                        _ucSkip = control.Parameter.HasValue ? Math.Max(0, control.Parameter.Value) : 1;
                        return;
                    case "tab":
                        AppendText("\t", link);
                        return;
                    case "line":
                        AppendText("\n", link);
                        return;
                    case "emdash": AppendText("\u2014", link); return;
                    case "endash": AppendText("\u2013", link); return;
                    case "bullet": AppendText("\u2022", link); return;
                    case "lquote": AppendText("\u2018", link); return;
                    case "rquote": AppendText("\u2019", link); return;
                    case "ldblquote": AppendText("\u201C", link); return;
                    case "rdblquote": AppendText("\u201D", link); return;
                    case "\\": AppendText("\\", link); return;
                    case "{": AppendText("{", link); return;
                    case "}": AppendText("}", link); return;
                    case "~": AppendText("\u00A0", link); return;
                    case "_": AppendText("-", link); return;
                    case "-": return;
                    case "trowd": StartRow(); return;
                    case "cellx": AddCellBoundary(control); return;
                    case "cell": FinishCell(); return;
                    case "row": FinishRow(); return;
                    case "trhdr": if (_row != null) _row.IsHeader = true; return;
                    case "trrh": if (_row != null) _row.Height = Math.Abs(Twips(control)); return;
                    case "trgaph": if (_table != null) _table.CellPadding = Math.Max(0, Twips(control)); return;
                    case "clmgf": Slot().HorizontalFirst = true; return;
                    case "clmrg": Slot().HorizontalContinuation = true; return;
                    case "clvmgf": Slot().VerticalFirst = true; return;
                    case "clvmrg": Slot().VerticalContinuation = true; return;
                    case "clcbpat":
                        if (control.Parameter.HasValue) Slot().BackgroundIndex = control.Parameter.Value;
                        return;
                    case "clvertalt": Slot().VerticalAlignment = VerticalAlignment.Top; return;
                    case "clvertalc": Slot().VerticalAlignment = VerticalAlignment.Middle; return;
                    case "clvertalb": Slot().VerticalAlignment = VerticalAlignment.Bottom; return;
                    case "clbrdrt": _borderTarget = 1; _borderStyle = null; _borderWidth = 0.5; return;
                    case "clbrdrb": _borderTarget = 2; _borderStyle = null; _borderWidth = 0.5; return;
                    case "clbrdrl": _borderTarget = 3; _borderStyle = null; _borderWidth = 0.5; return;
                    case "clbrdrr": _borderTarget = 4; _borderStyle = null; _borderWidth = 0.5; return;
                    case "brdrs": case "brdrth": case "brdrsh": SetBorder(BorderStyle.Single); return;
                    case "brdrdash": case "brdrdashsm": SetBorder(BorderStyle.Dashed); return;
                    case "brdrdot": SetBorder(BorderStyle.Dotted); return;
                    case "brdrdb": SetBorder(BorderStyle.Double); return;
                    case "brdrnil": case "brdrnone": SetBorder(BorderStyle.None); return;
                    case "brdrw":
                        if (control.Parameter.HasValue)
                        {
                            _borderWidth = Math.Max(0, control.Parameter.Value / 20.0);
                            ApplyPendingBorder();
                        }
                        return;
                }
            }

            /// <summary>把缇换算为磅。Converts twips to points.</summary>
            private static double Twips(RtfControl control)
            {
                return control.Parameter.HasValue ? control.Parameter.Value / 20.0 : 0;
            }

            // ------------------------------------------------------------ border state

            private int _borderTarget;
            private BorderStyle? _borderStyle;
            private double _borderWidth = 0.5;

            /// <summary>取得当前单元格定义槽，必要时创建。Gets the current cell-definition slot, creating it when needed.</summary>
            private CellSlot Slot()
            {
                if (_slot == null) _slot = new CellSlot();
                return _slot;
            }

            /// <summary>记录当前边框样式并写入目标边。Records the current border style and writes it to the target edge.</summary>
            private void SetBorder(BorderStyle style)
            {
                _borderStyle = style;
                ApplyPendingBorder();
            }

            /// <summary>把当前边框目标、样式与宽度落到槽位上。Commits the current border target, style and width to the slot.</summary>
            private void ApplyPendingBorder()
            {
                if (_borderTarget == 0) return;
                CellSlot slot = Slot();
                BorderLine line = new BorderLine(_borderStyle ?? BorderStyle.Single, _borderWidth, OfficeColor.Black);
                if (_borderTarget == 1) slot.TopBorder = line;
                else if (_borderTarget == 2) slot.BottomBorder = line;
                else if (_borderTarget == 3) slot.LeftBorder = line;
                else slot.RightBorder = line;
            }

            // ------------------------------------------------------------ groups

            /// <summary>处理嵌套分组：域、图片、书签与普通格式作用域。Handles nested groups: fields, images, bookmarks and ordinary formatting scopes.</summary>
            private void HandleGroup(RtfGroup group, string link, int depth)
            {
                string destination = GetDestinationName(group);
                if (destination == "field")
                {
                    HandleField(group, link, depth);
                    return;
                }
                if (destination == "pict")
                {
                    HandlePicture(group);
                    return;
                }
                if (destination == "bkmkstart")
                {
                    string name = GetPlainText(group).Trim();
                    name = CleanBookmarkName(name);
                    if (name.Length > 0)
                    {
                        // 起点既是行内范围标记，也沿用段落书签名兼容旧读取语义。
                        // The start becomes an inline range marker and still fills the paragraph bookmark name.
                        Paragraph bookmarked = EnsureParagraph();
                        bookmarked.BookmarkName = name;
                        bookmarked.Inlines.Add(new BookmarkStart(name));
                    }
                    return;
                }
                if (destination == "bkmkend")
                {
                    string ended = CleanBookmarkName(GetPlainText(group).Trim());
                    if (ended.Length > 0) EnsureParagraph().Inlines.Add(new BookmarkEnd(ended));
                    return;
                }
                if (destination == "fldinst" || destination == "fldrslt")
                {
                    ProcessNodes(group.Items, link, depth + 1);
                    return;
                }
                if (HasIgnorableMarker(group))
                {
                    return;
                }
                if (IsNestedDestination(destination)) return;
                ProcessNodes(group.Items, link, depth + 1);
            }

            /// <summary>判断分组的首个控制字是否为需要跳过的目标（样式表、脚注、批注等）。Whether the group's first control word is a destination to skip (stylesheet, footnote, annotation, etc.).</summary>
            private static bool IsNestedDestination(string destination)
            {
                if (destination == null) return false;
                switch (destination.ToLowerInvariant())
                {
                    case "header":
                    case "footer":
                    case "footnote":
                    case "annotation":
                    case "atnid":
                    case "atnauthor":
                    case "stylesheet":
                    case "listtext":
                    case "pntext":
                    case "pntxta":
                    case "pntxtb":
                    case "object":
                    case "shp":
                    case "shpinst":
                    case "nonshppict":
                    case "result":
                    case "generator":
                    case "info":
                    case "fonttbl":
                    case "colortbl":
                    case "filetbl":
                    case "listtable":
                    case "listoverridetable":
                        return true;
                    default:
                        return false;
                }
            }

            /// <summary>清理书签名，只保留字母、数字、下划线、连字符与点。Cleans a bookmark name, keeping letters, digits, underscores, hyphens and dots.</summary>
            private static string CleanBookmarkName(string value)
            {
                StringBuilder builder = new StringBuilder(value.Length);
                for (int i = 0; i < value.Length; i++)
                {
                    char current = value[i];
                    if (char.IsLetterOrDigit(current) || current == '_' || current == '-' || current == '.') builder.Append(current);
                }
                return builder.ToString();
            }

            /// <summary>解析 HYPERLINK 等域；超链接结果作为带链接的文本运行写入。Parses HYPERLINK and other fields; hyperlink results become text runs carrying the link.</summary>
            private void HandleField(RtfGroup group, string link, int depth)
            {
                RtfGroup instructionGroup = FindGroup(group, "fldinst");
                RtfGroup resultGroup = FindGroup(group, "fldrslt");
                string instruction = instructionGroup != null ? GetPlainText(instructionGroup).Trim() : string.Empty;
                string resultLink = null;
                if (instruction.StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase))
                {
                    string target = ExtractFieldArgument(instruction);
                    bool isAnchor = ContainsControl(instructionGroup, "l");
                    if (isAnchor && target.Length > 0 && target[0] != '#') target = "#" + target;
                    if (target.Length > 0 && (target[0] == '#' || WordFormatUtil.IsSafeHyperlinkTarget(target))) resultLink = target;
                }
                else if (instruction.Length > 0)
                {
                    string display = resultGroup != null ? GetPlainText(resultGroup) : string.Empty;
                    EnsureParagraph().Inlines.Add(new FieldRun(instruction, display, _format.Clone()));
                    return;
                }
                if (resultGroup != null)
                {
                    ProcessNodes(resultGroup.Items, resultLink ?? link, depth + 1);
                }
            }

            /// <summary>提取域指令中的第一个带引号参数；没有引号时取第一个空格后的内容。Extracts the first quoted argument from a field instruction, or the text after the first space when unquoted.</summary>
            private static string ExtractFieldArgument(string instruction)
            {
                int start = instruction.IndexOf('"');
                if (start >= 0)
                {
                    int end = instruction.IndexOf('"', start + 1);
                    if (end > start) return instruction.Substring(start + 1, end - start - 1);
                }
                int space = instruction.IndexOf(' ');
                return space >= 0 ? instruction.Substring(space + 1).Trim() : string.Empty;
            }

            /// <summary>判断分组内是否含有指定控制字。Whether the group contains the given control word.</summary>
            private static bool ContainsControl(RtfGroup group, string name)
            {
                if (group == null) return false;
                for (int i = 0; i < group.Items.Count; i++)
                {
                    RtfControl control = group.Items[i] as RtfControl;
                    if (control != null && string.Equals(control.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
                    RtfGroup child = group.Items[i] as RtfGroup;
                    if (child != null && ContainsControl(child, name)) return true;
                }
                return false;
            }

            // ------------------------------------------------------------ pictures

            /// <summary>解析 \pict 图片组：仅接受 PNG/JPEG 十六进制数据。Parses a \pict group: only PNG/JPEG hexadecimal data is accepted.</summary>
            private void HandlePicture(RtfGroup group)
            {
                if (_imageCount >= MaxImages) return;
                bool png = false;
                bool jpeg = false;
                int pixelWidth = 0;
                int pixelHeight = 0;
                int goalWidth = 0;
                int goalHeight = 0;
                StringBuilder hex = new StringBuilder();
                for (int i = 0; i < group.Items.Count; i++)
                {
                    RtfControl control = group.Items[i] as RtfControl;
                    if (control != null)
                    {
                        if (control.Name == "pngblip") png = true;
                        else if (control.Name == "jpegblip" || control.Name == "jpgblip") jpeg = true;
                        else if (control.Name == "picw" && control.Parameter.HasValue) pixelWidth = control.Parameter.Value;
                        else if (control.Name == "pich" && control.Parameter.HasValue) pixelHeight = control.Parameter.Value;
                        else if (control.Name == "picwgoal" && control.Parameter.HasValue) goalWidth = control.Parameter.Value;
                        else if (control.Name == "pichgoal" && control.Parameter.HasValue) goalHeight = control.Parameter.Value;
                        continue;
                    }
                    RtfText text = group.Items[i] as RtfText;
                    if (text != null) hex.Append(text.Text);
                }
                if (!png && !jpeg) return;
                if (hex.Length == 0) return;
                int capacity = hex.Length / 2;
                if (capacity > MaxImageBytes) return;
                byte[] data = new byte[capacity];
                int written = 0;
                int high = -1;
                for (int i = 0; i < hex.Length; i++)
                {
                    int value = HexValue(hex[i]);
                    if (value < 0) continue;
                    if (high < 0)
                    {
                        high = value;
                    }
                    else
                    {
                        data[written++] = (byte)((high << 4) | value);
                        high = -1;
                    }
                }
                if (written == 0) return;
                if (written != data.Length) Array.Resize(ref data, written);
                OfficeImage image;
                try
                {
                    image = new OfficeImage(data);
                }
                catch (ArgumentException)
                {
                    return;
                }
                if (image.Format != OfficeImageFormat.Png && image.Format != OfficeImageFormat.Jpeg) return;
                double width = goalWidth > 0 ? goalWidth / 20.0 : pixelWidth > 0 ? pixelWidth * 72.0 / 96.0 : 0;
                double height = goalHeight > 0 ? goalHeight / 20.0 : pixelHeight > 0 ? pixelHeight * 72.0 / 96.0 : 0;
                EnsureParagraph().Inlines.Add(new ImageRun(image, width, height));
                _imageCount++;
            }

            // ------------------------------------------------------------ tables

            /// <summary>开始一个新表格行；首次调用时创建表格并把输出切换到单元格。Starts a new table row; the first call creates the table and redirects output into cells.</summary>
            private void StartRow()
            {
                if (_row != null) FinishRow();
                if (_table == null)
                {
                    if (_paragraph != null) FlushParagraph();
                    _outerBlocks = _blocks;
                    _table = new Table();
                    _table.Borders = TableBorders.None();
                    _blocks.Add(_table);
                    _previousGrid = new TableCell[0];
                    _columnCount = 0;
                }
                _row = new TableRow();
                _slots = new List<CellSlot>();
                _slot = new CellSlot();
                _physicalCells = new List<TableCell>();
                _cell = new TableCell();
                _cellIndex = 0;
                _blocks = _cell.Blocks;
                _paragraph = null;
            }

            /// <summary>记录一个单元格右边界。Records a cell right boundary.</summary>
            private void AddCellBoundary(RtfControl control)
            {
                if (_row == null) return;
                CellSlot slot = _slot ?? new CellSlot();
                ApplyPendingBorder();
                slot.Boundary = control.Parameter.GetValueOrDefault();
                _slots.Add(slot);
                _slot = new CellSlot();
                _borderTarget = 0;
                _borderStyle = null;
            }

            /// <summary>结束当前单元格，应用其定义属性并开始下一个单元格。Finishes the current cell, applies its definition properties and starts the next one.</summary>
            private void FinishCell()
            {
                if (_row == null) return;
                FlushParagraph();
                if (_cellIndex < _slots.Count) ApplySlotProperties(_cell, _slots[_cellIndex]);
                _physicalCells.Add(_cell);
                _cellIndex++;
                _cell = new TableCell();
                _blocks = _cell.Blocks;
                _paragraph = null;
            }

            /// <summary>结束当前行：按合并标记重建逻辑单元格并写入表格。Finishes the current row: rebuilds logical cells from merge markers and writes it into the table.</summary>
            private void FinishRow()
            {
                if (_row == null) return;
                FlushParagraph();
                if (_cellIndex < _slots.Count && _cell.Blocks.Count > 0)
                {
                    ApplySlotProperties(_cell, _slots[_cellIndex]);
                    _physicalCells.Add(_cell);
                }
                if (_table != null && _slots != null)
                {
                    BuildRowCells();
                    _table.Rows.Add(_row);
                    if (_columnCount == 0 && _slots.Count > 0)
                    {
                        _columnCount = _slots.Count;
                        int previous = 0;
                        for (int i = 0; i < _slots.Count; i++)
                        {
                            int boundary = _slots[i].Boundary;
                            _table.Columns.Add(new TableColumn(Math.Max(0, (boundary - previous) / 20.0)));
                            previous = boundary;
                        }
                    }
                    else if (_slots.Count > _columnCount)
                    {
                        _columnCount = _slots.Count;
                    }
                }
                _row = null;
                _slots = null;
                _slot = null;
                _physicalCells = null;
                _cell = null;
                _cellIndex = 0;
                _blocks = _outerBlocks ?? _rootTarget;
                _paragraph = null;
            }

            /// <summary>按横向/纵向合并标记把物理单元格折叠为逻辑单元格，并重建行网格映射。Collapses physical cells into logical cells using horizontal/vertical merge markers and rebuilds the row grid map.</summary>
            private void BuildRowCells()
            {
                int total = Math.Max(_columnCount, _slots.Count);
                if (total <= 0) total = Math.Max(1, _slots.Count);
                if (_previousGrid.Length < total) Array.Resize(ref _previousGrid, total);
                TableCell[] grid = new TableCell[total];
                List<TableCell> cells = new List<TableCell>();
                HashSet<TableCell> extended = new HashSet<TableCell>();
                int index = 0;
                int column = 0;
                while (index < _slots.Count && column < total)
                {
                    CellSlot slot = _slots[index];
                    int span = 1;
                    while (index + span < _slots.Count && _slots[index + span].HorizontalContinuation) span++;
                    if (column + span > total) span = Math.Max(1, total - column);

                    TableCell target = null;
                    if (slot.VerticalContinuation && column < _previousGrid.Length)
                    {
                        target = _previousGrid[column];
                    }
                    if (target == null)
                    {
                        target = index < _physicalCells.Count ? _physicalCells[index] : new TableCell();
                        ApplySlotProperties(target, slot);
                        target.ColSpan = span;
                        cells.Add(target);
                    }
                    else if (extended.Add(target))
                    {
                        target.RowSpan = target.RowSpan + 1;
                    }
                    for (int k = 0; k < span && column + k < total; k++) grid[column + k] = target;
                    index += span;
                    column += span;
                }
                _row.Cells.AddRange(cells);
                _previousGrid = grid;
            }

            /// <summary>把单元格定义属性（底色、纵向对齐、边框）写入逻辑单元格。Applies cell definition properties (shading, vertical alignment, borders) to a logical cell.</summary>
            private void ApplySlotProperties(TableCell cell, CellSlot slot)
            {
                if (cell == null || slot == null) return;
                if (slot.BackgroundIndex.HasValue && slot.BackgroundIndex.Value > 0 && slot.BackgroundIndex.Value < _context.Colors.Count)
                {
                    cell.BackgroundColor = _context.Colors[slot.BackgroundIndex.Value];
                }
                cell.VerticalAlignment = slot.VerticalAlignment;
                if (slot.TopBorder != null || slot.BottomBorder != null || slot.LeftBorder != null || slot.RightBorder != null)
                {
                    cell.Borders = new TableCellBorders
                    {
                        Top = slot.TopBorder,
                        Bottom = slot.BottomBorder,
                        Left = slot.LeftBorder,
                        Right = slot.RightBorder
                    };
                }
            }

            /// <summary>收尾当前表格：结束未完成的行并恢复外层输出目标。Finishes the active table: closes any unfinished row and restores the outer output target.</summary>
            private void FlushTable()
            {
                if (_table == null) return;
                if (_row != null) FinishRow();
                _table = null;
                _previousGrid = new TableCell[0];
                _columnCount = 0;
                if (_outerBlocks != null)
                {
                    _blocks = _outerBlocks;
                    _outerBlocks = null;
                }
                _paragraph = null;
            }
        }
    }
}
