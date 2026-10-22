using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// DOCX 模板引擎：在保持运行格式的前提下填充 .docx 包中的 {{token}} 占位符。
    /// 支持跨运行合并的标记、{{key:format}} 格式化说明符，以及由 {{#each name}} 与
    /// {{/each}} 界定的表格行循环和段落循环。
    /// DOCX template engine that fills {{token}} placeholders inside a .docx package while preserving run
    /// formatting. It supports markers split across runs, {{key:format}} format specifiers, and table-row and
    /// paragraph loops delimited by {{#each name}} and {{/each}}.
    /// </summary>
    public static class DocxTemplate
    {
        /// <summary>匹配 {{key}} 或 {{key:format}} 占位符（key 以 # 或 / 开头的循环标记除外）。Matches {{key}} or {{key:format}} placeholders, excluding loop markers whose key starts with # or /.</summary>
        private static readonly Regex TokenRegex = new Regex(
            @"\{\{\s*([^#/{}][^{}:]*?)\s*(?::\s*([^{}]+?))?\s*\}\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>匹配 {{#each name}} 循环起始标记。Matches the opening {{#each name}} loop marker.</summary>
        private static readonly Regex LoopStartRegex = new Regex(
            @"\{\{\s*#each\s+([^{}]+?)\s*\}\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// <summary>匹配 {{/each}} 循环结束标记。Matches the closing {{/each}} loop marker.</summary>
        private static readonly Regex LoopEndRegex = new Regex(
            @"\{\{\s*/each\s*\}\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// <summary>段落内单个 w:t 文本节点在「合并文本」中的位置记录，用于跨运行替换。Records the position of one w:t text node inside a paragraph within its merged text, used to replace markers that span several runs.</summary>
        private sealed class TextNode
        {
            /// <summary>对应的 w:t 元素。The w:t element this node refers to.</summary>
            public XElement Element;
            /// <summary>该节点文本在段落合并文本中的起始偏移。The start offset of the node text within the merged paragraph text.</summary>
            public int Start;
            /// <summary>该节点文本长度。The length of the node text.</summary>
            public int Length;
        }

        /// <summary>
        /// 用默认模板选项将上下文数据填充到模板文件并写出到目标文件。
        /// Fills a template file with the context data using the default template options and writes the result
        /// to the target file.
        /// </summary>
        /// <param name="templatePath">模板 .docx 文件路径。The path of the template .docx file.</param>
        /// <param name="outputPath">输出 .docx 文件路径。The path of the output .docx file.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        public static void Fill(string templatePath, string outputPath, TemplateContext context)
        {
            Fill(templatePath, outputPath, context, new TemplateOptions());
        }

        /// <summary>
        /// 使用指定选项将上下文数据填充到模板文件并写出到目标文件。
        /// Fills a template file with the context data using the given options and writes the result to the
        /// target file.
        /// </summary>
        /// <param name="templatePath">模板 .docx 文件路径。The path of the template .docx file.</param>
        /// <param name="outputPath">输出 .docx 文件路径。The path of the output .docx file.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项（未解析标记处理、格式化区域等）。The template options (unresolved-marker handling, format culture, ...).</param>
        /// <exception cref="ArgumentNullException"><paramref name="templatePath"/> 或 <paramref name="outputPath"/> 为 null。Thrown when <paramref name="templatePath"/> or <paramref name="outputPath"/> is null.</exception>
        public static void Fill(string templatePath, string outputPath, TemplateContext context, TemplateOptions options)
        {
            if (templatePath == null) throw new ArgumentNullException("templatePath");
            if (outputPath == null) throw new ArgumentNullException("outputPath");
            byte[] result = Fill(File.ReadAllBytes(templatePath), context, options);
            File.WriteAllBytes(outputPath, result);
        }

        /// <summary>
        /// 使用默认选项填充模板字节数据并返回结果包。
        /// Fills the template bytes with the default options and returns the resulting package.
        /// </summary>
        /// <param name="templateData">模板 .docx 字节内容。The bytes of the template .docx package.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <returns>填充后的 .docx 字节内容。The filled .docx bytes.</returns>
        public static byte[] Fill(byte[] templateData, TemplateContext context)
        {
            return Fill(templateData, context, new TemplateOptions());
        }

        /// <summary>
        /// 使用指定选项填充模板字节数据并返回结果包；包损坏时抛出 <see cref="InvalidDataException"/>。
        /// Fills the template bytes with the given options and returns the resulting package; throws
        /// <see cref="InvalidDataException"/> when the package is corrupt.
        /// </summary>
        /// <param name="templateData">模板 .docx 字节内容。The bytes of the template .docx package.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项，为 null 时使用默认选项。The template options; the defaults are used when null.</param>
        /// <returns>填充后的 .docx 字节内容。The filled .docx bytes.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="templateData"/> 为 null。Thrown when <paramref name="templateData"/> is null.</exception>
        public static byte[] Fill(byte[] templateData, TemplateContext context, TemplateOptions options)
        {
            if (templateData == null) throw new ArgumentNullException("templateData");
            if (options == null) options = new TemplateOptions();
            try
            {
                return FillCore(templateData, context, options);
            }
            catch (IOException ex)
            {
                throw new InvalidDataException("The .docx template package is corrupt or truncated.", ex);
            }
        }

        /// <summary>
        /// 核心填充流程：逐条复制 ZIP 条目，仅对主文档及 word/header*、word/footer* 部件执行
        /// 循环展开与占位符替换，其余部件原样写出。
        /// Core fill routine: copies every ZIP entry, expanding loops and replacing placeholders only in the
        /// main document and the word/header*, word/footer* parts; all other parts are written out unchanged.
        /// </summary>
        /// <param name="templateData">模板 .docx 字节内容。The bytes of the template .docx package.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        /// <returns>填充后的 .docx 字节内容。The filled .docx bytes.</returns>
        private static byte[] FillCore(byte[] templateData, TemplateContext context, TemplateOptions options)
        {
            using (MemoryStream input = new MemoryStream(templateData, false))
            using (ZipArchive source = PackageArchive.OpenForRead(input, false))
            using (MemoryStream output = new MemoryStream())
            {
                long totalBytes = 0;
                // 根关系部件也计入累计解压量，模板包的每一个字节都受同一总量上限约束。
                // The root relationship part is counted too, so every byte of the template package is covered
                // by the same cumulative limit.
                string mainPart = FindMainPart(source, ref totalBytes);
                using (ZipArchive target = new ZipArchive(output, ZipArchiveMode.Create, true))
                {
                    foreach (ZipArchiveEntry entry in source.Entries)
                    {
                        string name = entry.FullName.Replace('\\', '/');
                        if (IsProcessedPart(name, mainPart))
                        {
                            byte[] data = OfficePackageLimits.ReadEntryBytes(entry, ref totalBytes);
                            XDocument document = Wml.LoadXml(data);
                            if (document != null && document.Root != null && context != null)
                            {
                                ProcessPart(document.Root, context, options);
                                byte[] updated = Wml.SerializeXml(document);
                                WriteEntry(target, name, updated);
                            }
                            else
                            {
                                WriteEntry(target, name, data);
                            }
                        }
                        else
                        {
                            WriteEntry(target, name, OfficePackageLimits.ReadEntryBytes(entry, ref totalBytes));
                        }
                    }
                }
                return output.ToArray();
            }
        }

        /// <summary>
        /// 以流为输入输出执行模板填充：模板流先读入内存处理，结果写入输出流。
        /// Fills a template from a stream into a stream: the template is read into memory first and the result
        /// is written to the output stream.
        /// </summary>
        /// <param name="templateStream">可读的模板 .docx 流。A readable stream holding the template .docx package.</param>
        /// <param name="outputStream">接收结果的 writable 流。The writable stream that receives the result.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        /// <exception cref="ArgumentNullException"><paramref name="templateStream"/> 或 <paramref name="outputStream"/> 为 null。Thrown when <paramref name="templateStream"/> or <paramref name="outputStream"/> is null.</exception>
        public static void Fill(Stream templateStream, Stream outputStream, TemplateContext context, TemplateOptions options)
        {
            if (templateStream == null) throw new ArgumentNullException("templateStream");
            if (outputStream == null) throw new ArgumentNullException("outputStream");
            using (MemoryStream memory = new MemoryStream())
            {
                templateStream.CopyTo(memory);
                byte[] result = Fill(memory.ToArray(), context, options);
                outputStream.Write(result, 0, result.Length);
            }
        }

        /// <summary>
        /// 判断 ZIP 条目是否需要做模板处理：主文档部件、word/header*.xml 与 word/footer*.xml。
        /// Decides whether a ZIP entry needs template processing: the main document part plus
        /// word/header*.xml and word/footer*.xml.
        /// </summary>
        /// <param name="name">条目路径（已规范化为 / 分隔）。The entry path, normalised to forward slashes.</param>
        /// <param name="mainPart">主文档部件路径。The path of the main document part.</param>
        /// <returns>需要处理时返回 true。True when the entry has to be processed.</returns>
        private static bool IsProcessedPart(string name, string mainPart)
        {
            if (string.Equals(name, mainPart, StringComparison.OrdinalIgnoreCase)) return true;
            if (!name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return false;
            if (name.StartsWith("word/header", StringComparison.OrdinalIgnoreCase)) return true;
            if (name.StartsWith("word/footer", StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        /// <summary>
        /// 通过包级关系 _rels/.rels 定位 officeDocument 主部件路径；解析失败时回退为 word/document.xml。
        /// Locates the officeDocument main part through the package-level _rels/.rels file and falls back to
        /// word/document.xml when it cannot be resolved.
        /// </summary>
        /// <param name="archive">模板 ZIP 包。The template ZIP package.</param>
        /// <param name="totalBytes">已累计解压的字节数，读取根关系部件时同步累加。 Running total of decompressed bytes, increased while reading the root relationships part.</param>
        /// <returns>主文档部件路径。The path of the main document part.</returns>
        private static string FindMainPart(ZipArchive archive, ref long totalBytes)
        {
            ZipArchiveEntry rels = archive.GetEntry("_rels/.rels");
            if (rels != null)
            {
                try
                {
                    XDocument document = Wml.LoadXml(OfficePackageLimits.ReadEntryBytes(rels, ref totalBytes));
                    if (document != null && document.Root != null)
                    {
                        foreach (XElement relationship in document.Root.Elements())
                        {
                            string type = (string)relationship.Attribute("Type");
                            string target = (string)relationship.Attribute("Target");
                            if (type != null && type.EndsWith("/officeDocument", StringComparison.OrdinalIgnoreCase)
                                && !string.IsNullOrEmpty(target))
                            {
                                return target.TrimStart('/');
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }
            }
            return "word/document.xml";
        }

        /// <summary>
        /// 以最佳压缩级别向目标包写入条目；部件名含路径穿越等非法字符时拒绝写入。
        /// Writes an entry into the target package at the optimal compression level and refuses names that
        /// contain path traversal or other illegal characters.
        /// </summary>
        /// <param name="archive">目标 ZIP 包。The target ZIP package.</param>
        /// <param name="name">条目路径名。The entry path name.</param>
        /// <param name="data">条目内容字节。The entry bytes.</param>
        /// <exception cref="InvalidDataException">部件名包含非法路径。Thrown when the part name contains an illegal path.</exception>
        private static void WriteEntry(ZipArchive archive, string name, byte[] data)
        {
            if (!IsSafePartName(name))
            {
                throw new InvalidDataException("文档部件名称包含非法路径: " + name);
            }
            ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
            using (Stream stream = entry.Open())
            {
                stream.Write(data, 0, data.Length);
            }
        }

        /// <summary>
        /// 校验部件路径名的安全性：拒绝空名、NUL 字符、绝对路径、盘符前缀以及 "."/".." 路径穿越段。
        /// Validates a part path name: rejects empty names, NUL characters, absolute paths, drive prefixes and
        /// "."/".." traversal segments.
        /// </summary>
        /// <param name="name">待校验的部件路径名。The part path name to validate.</param>
        /// <returns>名称安全时返回 true。True when the name is safe.</returns>
        private static bool IsSafePartName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            string normalized = name.Replace('\\', '/');
            if (normalized.IndexOf('\0') >= 0) return false;
            if (normalized.StartsWith("/", StringComparison.Ordinal)) return false;
            if (normalized.Length >= 2 && normalized[1] == ':') return false;
            string[] segments = normalized.Split('/');
            for (int i = 0; i < segments.Length; i++)
            {
                if (segments[i] == ".." || segments[i] == ".") return false;
            }
            return true;
        }

        /// <summary>
        /// 处理单个 XML 部件：依次执行表格行循环展开、段落循环展开，最后替换全部 {{token}} 占位符。
        /// Processes one XML part: expands table-row loops, then paragraph loops, and finally replaces every
        /// {{token}} placeholder.
        /// </summary>
        /// <param name="root">部件根元素。The root element of the part.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        private static void ProcessPart(XElement root, TemplateContext context, TemplateOptions options)
        {
            ProcessTableLoops(root, context, options);
            ProcessParagraphLoops(root, context, options);
            ReplaceTokensInElement(root, context, options);
        }

        /// <summary>
        /// 对部件内所有表格反复展开首个行循环（带上限保护，防止恶意模板导致死循环）。
        /// Repeatedly expands the first row loop of every table in the part, with an iteration cap that protects
        /// against malicious templates causing an endless loop.
        /// </summary>
        /// <param name="root">部件根元素。The root element of the part.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        private static void ProcessTableLoops(XElement root, TemplateContext context, TemplateOptions options)
        {
            List<XElement> tables = root.Descendants(Wml.W + "tbl").ToList();
            for (int i = 0; i < tables.Count; i++)
            {
                int guard = 0;
                while (guard++ < 1000 && ExpandFirstTableLoop(tables[i], context, options))
                {
                }
            }
        }

        /// <summary>
        /// 查找并展开表格中首个 {{#each}}..{{/each}} 行循环：
        /// 起止标记同行时按单行模板克隆展开，跨行时以中间行为模板整体复制；
        /// 每轮迭代替换 {{.}}、@index、@number、@first、@last 等上下文后插入到结束标记之前。
        /// Finds and expands the first {{#each}}..{{/each}} row loop of a table: when both markers sit in one
        /// row that row is cloned per item, otherwise the rows between them form the template. Each iteration
        /// replaces {{.}}, @index, @number, @first and @last and is inserted before the closing marker row.
        /// </summary>
        /// <param name="table">目标表格元素。The target table element.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        /// <returns>成功展开一个循环返回 true；表格中无完整循环返回 false。True when one loop was expanded; false when the table holds no complete loop.</returns>
        private static bool ExpandFirstTableLoop(XElement table, TemplateContext context, TemplateOptions options)
        {
            List<XElement> rows = table.Elements(Wml.W + "tr").ToList();
            int start = -1;
            for (int i = 0; i < rows.Count; i++)
            {
                if (LoopStartRegex.IsMatch(GetCombinedText(rows[i])))
                {
                    start = i;
                    break;
                }
            }
            if (start < 0) return false;

            int end = -1;
            for (int i = start; i < rows.Count; i++)
            {
                if (LoopEndRegex.IsMatch(GetCombinedText(rows[i])))
                {
                    end = i;
                    break;
                }
            }
            if (end < 0) return false;

            Match startMatch = LoopStartRegex.Match(GetCombinedText(rows[start]));
            string loopName = startMatch.Groups[1].Value.Trim();
            object value = context == null ? null : context.Get(loopName);
            IList<object> items = ResolveItems(value);

            if (start == end)
            {
                XElement templateRow = rows[start];
                for (int i = 0; i < items.Count; i++)
                {
                    XElement clone = new XElement(templateRow);
                    RemoveLoopMarkers(clone);
                    ReplaceTokensInElement(clone, CreateLoopContext(context, items[i], i, items.Count), options);
                    templateRow.AddBeforeSelf(clone);
                }
                templateRow.Remove();
                return true;
            }

            List<XElement> templateRows = new List<XElement>();
            for (int i = start + 1; i < end; i++) templateRows.Add(rows[i]);

            XElement endRow = rows[end];
            for (int i = 0; i < items.Count; i++)
            {
                TemplateContext rowContext = CreateLoopContext(context, items[i], i, items.Count);
                for (int r = 0; r < templateRows.Count; r++)
                {
                    XElement clone = new XElement(templateRows[r]);
                    RemoveLoopMarkers(clone);
                    ReplaceTokensInElement(clone, rowContext, options);
                    endRow.AddBeforeSelf(clone);
                }
            }

            for (int i = 0; i < templateRows.Count; i++) templateRows[i].Remove();
            rows[start].Remove();
            endRow.Remove();
            return true;
        }

        /// <summary>
        /// 将循环数据源解析为对象列表（委托 <see cref="TemplateEngine.ResolveList"/> 处理数组、集合与枚举数）。
        /// Resolves a loop data source into a list of objects, delegating arrays, collections and enumerators to
        /// <see cref="TemplateEngine.ResolveList"/>.
        /// </summary>
        /// <param name="value">数据源对象，可为 null（视为空列表）。The source object, may be null (treated as an empty list).</param>
        /// <returns>逐项装箱后的对象列表。The items boxed into an object list.</returns>
        private static IList<object> ResolveItems(object value)
        {
            IList list = TemplateEngine.ResolveList(value);
            List<object> items = new List<object>();
            for (int i = 0; i < list.Count; i++) items.Add(list[i]);
            return items;
        }

        /// <summary>
        /// 构建一次循环迭代的上下文：以当前项为数据根，注入 @index（从 0）、@number（从 1）、
        /// @first、@last 辅助变量，并与父上下文合并以便回退查找。
        /// Builds the context of a single loop iteration: the current item becomes the data root, @index
        /// (zero-based), @number (one-based), @first and @last helpers are injected, and the parent context is
        /// merged in so lookups can fall back to it.
        /// </summary>
        /// <param name="parent">外层父上下文。The enclosing parent context.</param>
        /// <param name="item">当前迭代项。The item of the current iteration.</param>
        /// <param name="index">当前迭代索引（从 0 开始）。The zero-based index of the current iteration.</param>
        /// <param name="count">迭代总项数。The total number of iterations.</param>
        /// <returns>合并后的迭代上下文。The merged iteration context.</returns>
        private static TemplateContext CreateLoopContext(TemplateContext parent, object item, int index, int count)
        {
            TemplateContext local = TemplateEngine.CreateRowContext(item);
            local.Set("@index", index);
            local.Set("@number", index + 1);
            local.Set("@first", index == 0);
            local.Set("@last", index == count - 1);
            return TemplateEngine.Merge(parent, local);
        }

        /// <summary>
        /// 对部件内所有段落容器（正文根及各表格单元格等）反复展开首个段落循环（带上限保护）。
        /// Repeatedly expands the first paragraph loop of every block container in the part (body root, table
        /// cells, ...), with an iteration cap.
        /// </summary>
        /// <param name="root">部件根元素。The root element of the part.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        private static void ProcessParagraphLoops(XElement root, TemplateContext context, TemplateOptions options)
        {
            List<XElement> containers = new List<XElement>();
            containers.Add(root);
            foreach (XElement paragraph in root.Descendants(Wml.W + "p"))
            {
                XElement parent = paragraph.Parent;
                if (parent != null && !containers.Contains(parent)) containers.Add(parent);
            }

            for (int c = 0; c < containers.Count; c++)
            {
                int guard = 0;
                while (guard++ < 1000 && ExpandFirstParagraphLoop(containers[c], context, options))
                {
                }
            }
        }

        /// <summary>
        /// 查找并展开容器内首个 {{#each}}..{{/each}} 段落循环，逻辑与行循环一致：
        /// 起止同段时按单段模板克隆，跨段时以中间段落为模板整体复制。
        /// Finds and expands the first {{#each}}..{{/each}} paragraph loop in a container using the same logic
        /// as row loops: one paragraph is cloned per item when both markers share it, otherwise the paragraphs
        /// between them form the template.
        /// </summary>
        /// <param name="container">段落所在容器元素。The container element holding the paragraphs.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        /// <returns>成功展开一个循环返回 true；无完整循环返回 false。True when one loop was expanded; false when there is no complete loop.</returns>
        private static bool ExpandFirstParagraphLoop(XElement container, TemplateContext context, TemplateOptions options)
        {
            List<XElement> paragraphs = container.Elements(Wml.W + "p").ToList();
            int start = -1;
            for (int i = 0; i < paragraphs.Count; i++)
            {
                if (LoopStartRegex.IsMatch(GetCombinedText(paragraphs[i])))
                {
                    start = i;
                    break;
                }
            }
            if (start < 0) return false;

            int end = -1;
            for (int i = start; i < paragraphs.Count; i++)
            {
                if (LoopEndRegex.IsMatch(GetCombinedText(paragraphs[i])))
                {
                    end = i;
                    break;
                }
            }
            if (end < 0) return false;

            Match startMatch = LoopStartRegex.Match(GetCombinedText(paragraphs[start]));
            string loopName = startMatch.Groups[1].Value.Trim();
            object value = context == null ? null : context.Get(loopName);
            IList<object> items = ResolveItems(value);

            if (end == start)
            {
                XElement templateParagraph = paragraphs[start];
                for (int i = 0; i < items.Count; i++)
                {
                    XElement clone = new XElement(templateParagraph);
                    RemoveLoopMarkers(clone);
                    ReplaceTokensInElement(clone, CreateLoopContext(context, items[i], i, items.Count), options);
                    templateParagraph.AddBeforeSelf(clone);
                }
                templateParagraph.Remove();
                return true;
            }

            List<XElement> templateParagraphs = new List<XElement>();
            for (int i = start + 1; i < end; i++) templateParagraphs.Add(paragraphs[i]);

            XElement endParagraph = paragraphs[end];
            for (int i = 0; i < items.Count; i++)
            {
                TemplateContext rowContext = CreateLoopContext(context, items[i], i, items.Count);
                for (int p = 0; p < templateParagraphs.Count; p++)
                {
                    XElement clone = new XElement(templateParagraphs[p]);
                    ReplaceTokensInElement(clone, rowContext, options);
                    endParagraph.AddBeforeSelf(clone);
                }
            }

            for (int i = 0; i < templateParagraphs.Count; i++) templateParagraphs[i].Remove();
            paragraphs[start].Remove();
            endParagraph.Remove();
            return true;
        }

        /// <summary>
        /// 清除元素内全部段落中的 {{#each}} 与 {{/each}} 循环标记文本。
        /// Removes the {{#each}} and {{/each}} marker text from every paragraph inside an element.
        /// </summary>
        /// <param name="element">目标元素（克隆出的模板行/段落）。The target element (a cloned template row or paragraph).</param>
        private static void RemoveLoopMarkers(XElement element)
        {
            foreach (XElement paragraph in element.DescendantsAndSelf(Wml.W + "p"))
            {
                RemoveMatches(paragraph, new[] { LoopStartRegex, LoopEndRegex }, string.Empty);
            }
        }

        /// <summary>
        /// 对元素自身及其后代中的每个段落执行占位符替换。
        /// Replaces placeholders in every paragraph of the element itself and its descendants.
        /// </summary>
        /// <param name="element">目标元素。The target element.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        private static void ReplaceTokensInElement(XElement element, TemplateContext context, TemplateOptions options)
        {
            List<XElement> paragraphs = new List<XElement>();
            if (element.Name == Wml.W + "p") paragraphs.Add(element);
            paragraphs.AddRange(element.Descendants(Wml.W + "p"));
            for (int i = 0; i < paragraphs.Count; i++)
            {
                ReplaceTokenMatch(paragraphs[i], context, options);
            }
        }

        /// <summary>
        /// 在段落的「合并文本」上匹配占位符并逐个替换（自后向前，避免偏移失效），
        /// 从而支持跨多个运行拆分的 {{token}}。
        /// Matches placeholders against the paragraph's merged text and replaces them from the end backwards so
        /// earlier offsets stay valid; this makes {{token}} markers split across several runs work.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="context">占位符数据上下文。The data context that supplies placeholder values.</param>
        /// <param name="options">模板选项。The template options.</param>
        private static void ReplaceTokenMatch(XElement paragraph, TemplateContext context, TemplateOptions options)
        {
            List<TextNode> nodes = CollectTextNodes(paragraph);
            if (nodes.Count == 0) return;
            StringBuilder combined = new StringBuilder();
            for (int i = 0; i < nodes.Count; i++) combined.Append(nodes[i].Element.Value);
            string text = combined.ToString();
            MatchCollection matches = TokenRegex.Matches(text);
            if (matches.Count == 0) return;

            for (int i = matches.Count - 1; i >= 0; i--)
            {
                Match match = matches[i];
                string replacement = EvaluateToken(match, context, options);
                if (replacement == null) continue;
                ReplaceRange(paragraph, match.Index, match.Length, replacement);
            }
        }

        /// <summary>
        /// 求值单个占位符：查上下文取值，IFormattable 值按 {{key:format}} 格式化；
        /// 未解析的标记依选项保留（返回 null）或替换为空串。
        /// Evaluates one placeholder: looks the key up in the context, formats IFormattable values with the
        /// {{key:format}} specifier, and either keeps an unresolved marker (returns null) or replaces it with an
        /// empty string according to the options.
        /// </summary>
        /// <param name="match">占位符匹配结果。The placeholder match.</param>
        /// <param name="context">数据上下文，可为 null。The data context, may be null.</param>
        /// <param name="options">模板选项。The template options.</param>
        /// <returns>替换文本；返回 null 表示保持原标记不动。The replacement text; null leaves the original marker untouched.</returns>
        private static string EvaluateToken(Match match, TemplateContext context, TemplateOptions options)
        {
            string key = match.Groups[1].Value.Trim();
            string format = match.Groups[2].Success ? match.Groups[2].Value.Trim() : null;
            object value;
            if (context == null || !context.TryGet(key, out value))
            {
                return options.LeaveUnresolvedTokens ? null : string.Empty;
            }
            if (value == null) return string.Empty;
            IFormattable formattable = value as IFormattable;
            if (formattable != null)
            {
                return formattable.ToString(format, options.Culture ?? CultureInfo.InvariantCulture);
            }
            return Convert.ToString(value, options.Culture ?? CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 在段落合并文本上按每个正则逐个匹配并替换为指定文本（自后向前处理）。
        /// Matches each regular expression against the paragraph's merged text and replaces the hits with the
        /// given text, processing from the end backwards.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="regexes">待应用的正则列表。The regular expressions to apply.</param>
        /// <param name="replacement">替换文本。The replacement text.</param>
        private static void RemoveMatches(XElement paragraph, Regex[] regexes, string replacement)
        {
            foreach (Regex regex in regexes)
            {
                List<TextNode> nodes = CollectTextNodes(paragraph);
                if (nodes.Count == 0) return;
                StringBuilder combined = new StringBuilder();
                for (int i = 0; i < nodes.Count; i++) combined.Append(nodes[i].Element.Value);
                MatchCollection matches = regex.Matches(combined.ToString());
                for (int i = matches.Count - 1; i >= 0; i--)
                {
                    ReplaceRange(paragraph, matches[i].Index, matches[i].Length, replacement);
                }
            }
        }

        /// <summary>
        /// 按文档顺序收集段落直接拥有的全部 w:t 文本节点及其合并偏移（排除嵌套段落中的节点）。
        /// Collects, in document order, every w:t node owned directly by the paragraph together with its offset
        /// in the merged text; nodes inside nested paragraphs are excluded.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <returns>文本节点位置列表。The list of text-node positions.</returns>
        private static List<TextNode> CollectTextNodes(XElement paragraph)
        {
            List<TextNode> nodes = new List<TextNode>();
            int offset = 0;
            foreach (XElement text in paragraph.Descendants(Wml.W + "t"))
            {
                XElement owner = text.Ancestors(Wml.W + "p").FirstOrDefault();
                if (owner != paragraph) continue;
                TextNode node = new TextNode();
                node.Element = text;
                node.Start = offset;
                node.Length = text.Value.Length;
                nodes.Add(node);
                offset += node.Length;
            }
            return nodes;
        }

        /// <summary>
        /// 在段落的合并文本坐标上替换指定区间：区间落在单个 w:t 时原地替换，
        /// 跨多个 w:t 时把替换文本放入首个节点、清空中间节点并从末个节点裁剪剩余部分。
        /// Replaces a range addressed in the paragraph's merged text coordinates: a range inside a single w:t is
        /// rewritten in place, otherwise the replacement goes into the first node, the middle nodes are cleared
        /// and the last node keeps only the text after the range.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="index">合并文本中的起始偏移。The start offset within the merged text.</param>
        /// <param name="length">被替换区间长度。The length of the replaced range.</param>
        /// <param name="replacement">替换文本。The replacement text.</param>
        private static void ReplaceRange(XElement paragraph, int index, int length, string replacement)
        {
            List<TextNode> nodes = CollectTextNodes(paragraph);
            if (nodes.Count == 0) return;

            int firstIndex = -1;
            int lastIndex = -1;
            for (int i = 0; i < nodes.Count; i++)
            {
                int nodeEnd = nodes[i].Start + nodes[i].Length;
                if (firstIndex < 0 && index >= nodes[i].Start && index < nodeEnd) firstIndex = i;
                if (index + length - 1 >= nodes[i].Start && index + length - 1 < nodeEnd) { lastIndex = i; break; }
            }
            if (firstIndex < 0) return;
            if (lastIndex < 0) lastIndex = nodes.Count - 1;

            TextNode first = nodes[firstIndex];
            if (firstIndex == lastIndex)
            {
                string value = first.Element.Value;
                int localStart = Math.Max(0, index - first.Start);
                int localLength = Math.Min(length, value.Length - localStart);
                first.Element.Value = value.Substring(0, localStart) + replacement
                    + value.Substring(localStart + Math.Max(0, localLength));
                PreserveSpace(first.Element);
                return;
            }

            int firstLocal = Math.Max(0, index - first.Start);
            string firstValue = first.Element.Value;
            first.Element.Value = firstValue.Substring(0, firstLocal) + replacement;
            PreserveSpace(first.Element);

            for (int i = firstIndex + 1; i < lastIndex; i++)
            {
                nodes[i].Element.Value = string.Empty;
            }

            TextNode last = nodes[lastIndex];
            int lastLocal = index + length - last.Start;
            if (lastLocal < 0) lastLocal = 0;
            if (lastLocal > last.Element.Value.Length) lastLocal = last.Element.Value.Length;
            last.Element.Value = last.Element.Value.Substring(lastLocal);
            PreserveSpace(last.Element);
        }

        /// <summary>
        /// 为 w:t 元素设置 xml:space="preserve"，防止替换后的首尾空格被 XML 规范化吞掉。
        /// Sets xml:space="preserve" on a w:t element so leading and trailing spaces introduced by a replacement
        /// survive XML whitespace normalisation.
        /// </summary>
        /// <param name="element">被修改的 w:t 元素。The w:t element being modified.</param>
        private static void PreserveSpace(XElement element)
        {
            element.SetAttributeValue(XNamespace.Xml + "space", "preserve");
        }

        /// <summary>
        /// 拼接元素内全部 w:t 文本，用于在「合并文本」上检测循环标记。
        /// Concatenates all w:t text inside an element so loop markers can be detected in the merged text.
        /// </summary>
        /// <param name="element">目标元素（行或段落）。The target element (a row or a paragraph).</param>
        /// <returns>合并后的纯文本。The merged plain text.</returns>
        private static string GetCombinedText(XElement element)
        {
            StringBuilder builder = new StringBuilder();
            foreach (XElement text in element.Descendants(Wml.W + "t"))
            {
                builder.Append(text.Value);
            }
            return builder.ToString();
        }
    }
}
