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
        /// 核心填充流程：先整包读入（累计解压量仍受同一上限约束），剥离宏与嵌入对象部件及其关系，
        /// 再逐条写出；只有主文档与 word/header*、word/footer* 部件执行循环展开与占位符替换，
        /// 但任何部件的悬空引用都会被清理。
        /// Core fill routine: the package is read in full first (the same cumulative inflated-size limit still
        /// applies), macro and embedded-object parts together with the relationships pointing at them are
        /// stripped, and the remaining entries are written out; loop expansion and placeholder replacement only
        /// happen in the main document part and the word/header*, word/footer* parts, but dangling references
        /// in any part are cleaned.
        /// </summary>
        /// <remarks>
        /// 模板来自第三方时，宏工程（vbaProject.bin、vbaData.xml）与嵌入对象（word/embeddings/*、
        /// word/activeX/*）会随填充结果一起交到最终收件人手里，一次「填表」因此变成一次带宏文档的分发；
        /// 这些部件不再随填充结果带走，无论引用它们的是主文档还是脚注、词汇表之类不展开占位符的部件。
        /// 指向本机磁盘或网络共享的外部关系（file:、smb: 等）同样剔除，只留文字。
        /// 剥离是自洽的：部件被删时它自己的 .rels 一并删除并继续下探，引用者里失效的 r:id 同步清理，
        /// [Content_Types].xml 里对应的 Override 也删掉，绝不留指向不存在部件的关系。
        /// When the template came from somebody else, the macro project (vbaProject.bin, vbaData.xml) and the
        /// embedded objects (word/embeddings/*, word/activeX/*) would travel from the fill step straight to the
        /// final recipient, turning "fill a form" into "distribute a macro-enabled document"; those parts are no
        /// longer carried over, whether the referrer is the main document or a part without placeholders such as
        /// footnotes or the glossary. External relationships that reach a local disk or a network share (file:,
        /// smb: and the like) are dropped too, keeping only their text. The strip stays self-consistent: when a
        /// part goes, its own .rels goes with it and the walk continues through it, the relationship ids that
        /// stop resolving are cleaned from every referrer, and the matching Override entries in
        /// [Content_Types].xml go with the part, so no relationship is ever left pointing at a missing part.
        /// </remarks>
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
                // 整包一次性读入：根关系部件与每一个原样保留的条目都计入累计解压量，
                // 模板包的字节数上限因此不受「是否会被改写」影响。
                // The package is read in full: the root relationships part and every carried-over entry count
                // toward the cumulative limit, so the size ceiling does not depend on whether a part is rewritten.
                List<string> names = new List<string>();
                Dictionary<string, byte[]> parts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                foreach (ZipArchiveEntry entry in source.Entries)
                {
                    string entryName = entry.FullName.Replace('\\', '/');
                    byte[] data = OfficePackageLimits.ReadEntryBytes(entry, ref totalBytes);
                    if (!parts.ContainsKey(entryName)) names.Add(entryName);
                    parts[entryName] = data;
                }

                string mainPart = FindMainPart(GetPartBytes(parts, "_rels/.rels"));
                HashSet<string> droppedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                // 每个部件登记自己失效的关系 ID，写出前据此清理引用，避免留下悬空 r:id。
                // Every part records the relationship ids that stop resolving, so the references can be cleaned
                // before writing and no dangling r:id is left behind.
                Dictionary<string, HashSet<string>> dangling =
                    new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                CollectUnsafeParts(parts, droppedParts, dangling);
                HashSet<string> droppedPartNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string dropped in droppedParts) droppedPartNames.Add("/" + dropped);

                using (ZipArchive target = new ZipArchive(output, ZipArchiveMode.Create, true))
                {
                    foreach (string name in names)
                    {
                        if (droppedParts.Contains(name)) continue;
                        byte[] data = parts[name];
                        if (IsRelationshipPart(name)) data = PruneRelationships(name, data, droppedParts, dangling);
                        else if (string.Equals(name, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                            data = PruneContentTypes(data, droppedPartNames);

                        bool expandPlaceholders = IsProcessedPart(name, mainPart);
                        HashSet<string> ids;
                        bool hasDeadIds = dangling.TryGetValue(name, out ids) && ids.Count > 0;
                        // 引用清理与占位符替换各自独立：上下文为 null 时不展开文本，但悬空 r:id 照样要清；
                        // 反过来脚注、词汇表这类不展开占位符的部件，也只在被波及时才改写。
                        // Reference cleanup and placeholder expansion are independent: with a null context no text
                        // is expanded but dangling ids are still removed, and parts without placeholders, such as
                        // footnotes or the glossary, are only rewritten when they are affected.
                        if ((expandPlaceholders || hasDeadIds)
                            && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                        {
                            XDocument document = Wml.LoadXml(data);
                            if (document != null && document.Root != null)
                            {
                                if (hasDeadIds) RemoveDanglingReferences(document.Root, ids);
                                if (expandPlaceholders && context != null) ProcessPart(document.Root, context, options);
                                data = Wml.SerializeXml(document);
                            }
                        }
                        WriteEntry(target, name, data);
                    }
                }
                return output.ToArray();
            }
        }

        /// <summary>取某个部件的字节，不存在时返回 null。Returns the bytes of a part, or null when the package has no such entry.</summary>
        /// <param name="parts">部件名到字节的映射。The map from part name to bytes.</param>
        /// <param name="name">部件名。The part name.</param>
        /// <returns>字节或 null。The bytes or null.</returns>
        private static byte[] GetPartBytes(Dictionary<string, byte[]> parts, string name)
        {
            byte[] data;
            return parts.TryGetValue(name, out data) ? data : null;
        }

        /// <summary>
        /// 找出不能随填充结果带走的部件，并登记引用它们、因此不再可解析的关系 ID。
        /// 剥离会向下递归：被删部件自己的 .rels 一并删除，其中引用的部件同样删除。
        /// Collects the parts that must not be carried over and registers the relationship ids that reference
        /// them and therefore stop resolving. The strip recurses: a dropped part loses its own .rels, and the
        /// parts referenced from there are dropped as well.
        /// </summary>
        /// <param name="parts">全部部件字节。Every part in the package.</param>
        /// <param name="droppedParts">需要剥离的部件名。The parts to strip.</param>
        /// <param name="dangling">部件名到失效关系 ID 集合。The map from part name to its dead relationship ids.</param>
        private static void CollectUnsafeParts(Dictionary<string, byte[]> parts, HashSet<string> droppedParts,
            Dictionary<string, HashSet<string>> dangling)
        {
            Stack<string> pending = new Stack<string>();
            foreach (KeyValuePair<string, byte[]> pair in parts)
            {
                if (IsRelationshipPart(pair.Key)) pending.Push(pair.Key);
            }
            HashSet<string> scanned = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (pending.Count > 0)
            {
                string relsName = pending.Pop();
                if (!scanned.Add(relsName)) continue;
                string owner = OwnerPartOf(relsName);
                XDocument document = Wml.LoadXml(GetPartBytes(parts, relsName));
                if (document == null || document.Root == null) continue;
                foreach (XElement relationship in document.Root.Elements())
                {
                    string id = (string)relationship.Attribute("Id");
                    string target = (string)relationship.Attribute("Target");
                    string mode = (string)relationship.Attribute("TargetMode");
                    if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target)) continue;
                    if (string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!WordFormatUtil.IsSafeHyperlinkTarget(target)) AddDeadId(dangling, owner, id);
                        continue;
                    }
                    string resolved = ResolvePartPath(owner, target);
                    if (resolved == null || !IsUnsafeCarryoverPart(resolved)) continue;
                    DropPartWithReferences(parts, droppedParts, resolved, pending);
                    // 主文档、页眉页脚之外的部件同样要记下失效 ID，否则正文里会留下指向已删部件的 r:id。
                    // Parts other than the main document and the headers and footnotes must record the dead id
                    // too, otherwise their text keeps an r:id pointing at a part that no longer exists.
                    AddDeadId(dangling, owner, id);
                }
            }
        }

        /// <summary>
        /// 删除部件并连同它自己的关系文件一起下探。Drops a part together with its own relationships part.
        /// </summary>
        /// <param name="parts">全部部件字节。Every part in the package.</param>
        /// <param name="droppedParts">被剥离的部件名。The stripped parts.</param>
        /// <param name="name">待删部件路径。The part to drop.</param>
        /// <param name="pending">待扫描的关系文件。The relationship parts still to scan.</param>
        private static void DropPartWithReferences(Dictionary<string, byte[]> parts, HashSet<string> droppedParts,
            string name, Stack<string> pending)
        {
            if (!droppedParts.Add(name)) return;
            string childRels = RelationshipPartOf(name);
            if (childRels == null || !parts.ContainsKey(childRels)) return;
            droppedParts.Add(childRels);
            pending.Push(childRels);
        }

        /// <summary>登记某个部件里失效的关系 ID。Records a dead relationship id inside a part.</summary>
        /// <param name="dangling">部件名到失效关系 ID 集合。The map from part name to its dead ids.</param>
        /// <param name="owner">归属部件路径。The owning part path.</param>
        /// <param name="id">失效的关系 ID。The dead relationship id.</param>
        private static void AddDeadId(Dictionary<string, HashSet<string>> dangling, string owner, string id)
        {
            if (string.IsNullOrEmpty(owner)) return;
            HashSet<string> ids;
            if (!dangling.TryGetValue(owner, out ids))
            {
                ids = new HashSet<string>(StringComparer.Ordinal);
                dangling[owner] = ids;
            }
            ids.Add(id);
        }

        /// <summary>判断部件是否为宏工程或嵌入对象。Reports whether a part is a macro project or an embedded object.</summary>
        /// <remarks>
        /// 判定不能挂在 word/ 前缀上：../ 相对目标规范化后可能落在根级的 embeddings/、activeX/，
        /// 挂前缀会让同一条关系绕开剥离。
        /// The test must not hang on the word/ prefix: a normalised relative target can land on a root-level
        /// embeddings/ or activeX/, and a prefix-bound test would let that same relationship dodge the strip.
        /// </remarks>
        /// <param name="name">部件路径名。The part path.</param>
        /// <returns>不可随填充结果带走时返回 true。True when the part must not be carried over.</returns>
        private static bool IsUnsafeCarryoverPart(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            string normalized = "/" + name.Replace('\\', '/');
            if (normalized.IndexOf("/embeddings/", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (normalized.IndexOf("/activex/", StringComparison.OrdinalIgnoreCase) >= 0) return true;
            string file = normalized.Substring(normalized.LastIndexOf('/') + 1);
            return string.Equals(file, "vbaProject.bin", StringComparison.OrdinalIgnoreCase)
                || string.Equals(file, "vbaData.xml", StringComparison.OrdinalIgnoreCase)
                || string.Equals(file, "printers.bin", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 把关系目标解析为包内绝对部件路径；外部目标或解析失败返回 null。
        /// Resolves a relationship target to an absolute part path inside the package; null for external
        /// targets or when it cannot be resolved.
        /// </summary>
        /// <remarks>
        /// 必须走 DocxPackage.ResolvePartPath 这条会消解 . 与 .. 的规范化路径：Word 与 LibreOffice 都会写出
        /// ../embeddings/oleObject1.bin 这类相对目标，若只是字符串拼接就得到 word/../embeddings/... ，
        /// 前缀判定永远不成立，宏与嵌入对象照样被带走，剥离逻辑静默失效。
        /// The target has to go through DocxPackage.ResolvePartPath, which collapses dot segments: Word and
        /// LibreOffice both emit relative targets such as ../embeddings/oleObject1.bin, and plain string
        /// concatenation would yield word/../embeddings/... where the prefix test never matches, so the macro
        /// and embedded parts would still be carried over and the strip would fail silently.
        /// </remarks>
        /// <param name="owner">引用者部件路径。The path of the part that owns the relationship.</param>
        /// <param name="target">关系目标。The relationship target.</param>
        /// <returns>包内部件路径。The part path inside the package.</returns>
        private static string ResolvePartPath(string owner, string target)
        {
            if (string.IsNullOrEmpty(target)) return null;
            string value = target.Replace('\\', '/');
            if (Uri.IsWellFormedUriString(value, UriKind.Absolute)) return null;
            return DocxPackage.ResolvePartPath(owner, value);
        }

        /// <summary>由 .rels 路径推出其归属部件。Derives the part a .rels file belongs to.</summary>
        /// <param name="relsName">.rels 条目路径。The .rels entry path.</param>
        /// <returns>归属部件路径；包级关系返回空字符串。The owning part path, empty for package-level relationships.</returns>
        private static string OwnerPartOf(string relsName)
        {
            string normalized = relsName.Replace('\\', '/');
            int index = normalized.LastIndexOf("_rels/", StringComparison.OrdinalIgnoreCase);
            if (index < 0) return string.Empty;
            // _rels/ 之前的前缀自带结尾斜杠，直接再拼 "/" 会得到 word//document.xml 这样的双斜杠路径。
            // The prefix before _rels/ already ends with a slash, so appending another one would give
            // word//document.xml and the owner would never match the part it belongs to.
            string directory = normalized.Substring(0, index).TrimEnd('/');
            string file = normalized.Substring(index + "_rels/".Length);
            if (!file.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) return string.Empty;
            string owner = file.Substring(0, file.Length - ".rels".Length);
            if (string.IsNullOrEmpty(directory) || string.IsNullOrEmpty(owner)) return owner;
            return directory + "/" + owner;
        }

        /// <summary>
        /// 由归属部件路径推出其关系文件路径。Derives the relationships part path of an owning part.
        /// </summary>
        /// <param name="part">部件路径。The part path.</param>
        /// <returns>关系文件路径，部件名为空时返回 null。The relationships part path, null when the part name has no file name.</returns>
        private static string RelationshipPartOf(string part)
        {
            if (string.IsNullOrEmpty(part)) return null;
            string normalized = part.Replace('\\', '/');
            int index = normalized.LastIndexOf('/');
            string directory = index < 0 ? string.Empty : normalized.Substring(0, index);
            string file = index < 0 ? normalized : normalized.Substring(index + 1);
            if (file.Length == 0) return null;
            return directory.Length == 0 ? "_rels/" + file + ".rels" : directory + "/_rels/" + file + ".rels";
        }

        /// <summary>判断条目是否为关系部件。Reports whether an entry is a relationships part.</summary>
        /// <param name="name">条目路径。The entry path.</param>
        /// <returns>是关系部件时返回 true。True for a .rels entry.</returns>
        private static bool IsRelationshipPart(string name)
        {
            return name.IndexOf("_rels/", StringComparison.OrdinalIgnoreCase) >= 0
                && name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 删除指向被剥离部件或危险外部目标的关系条目。Removes relationship entries whose target was stripped or whose external target is unsafe.
        /// </summary>
        /// <param name="relsName">.rels 条目路径。The .rels entry path.</param>
        /// <param name="data">关系部件字节。The relationship part bytes.</param>
        /// <param name="droppedParts">被剥离的部件名。The stripped parts.</param>
        /// <param name="dangling">部件名到失效关系 ID。The map from part to dead ids.</param>
        /// <returns>改写后的关系部件字节。The rewritten relationship bytes.</returns>
        private static byte[] PruneRelationships(string relsName, byte[] data, HashSet<string> droppedParts,
            Dictionary<string, HashSet<string>> dangling)
        {
            string owner = OwnerPartOf(relsName);
            HashSet<string> ids;
            dangling.TryGetValue(owner, out ids);
            XDocument document = Wml.LoadXml(data);
            if (document == null || document.Root == null) return data;
            List<XElement> dead = new List<XElement>();
            foreach (XElement relationship in document.Root.Elements())
            {
                string id = (string)relationship.Attribute("Id");
                string target = (string)relationship.Attribute("Target");
                string mode = (string)relationship.Attribute("TargetMode");
                if (id != null && ids != null && ids.Contains(id))
                {
                    dead.Add(relationship);
                    continue;
                }
                if (!string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase))
                {
                    string resolved = ResolvePartPath(owner, target);
                    if (resolved != null && droppedParts.Contains(resolved)) dead.Add(relationship);
                }
            }
            if (dead.Count == 0) return data;
            for (int i = 0; i < dead.Count; i++) dead[i].Remove();
            return Wml.SerializeXml(document);
        }

        /// <summary>删除 [Content_Types].xml 里指向被剥离部件的 Override。Removes the [Content_Types].xml overrides describing stripped parts.</summary>
        /// <param name="data">内容类型部件字节。The content types part bytes.</param>
        /// <param name="droppedPartNames">形如 /word/vbaProject.bin 的部件名。Part names shaped like /word/vbaProject.bin.</param>
        /// <returns>改写后的字节。The rewritten bytes.</returns>
        private static byte[] PruneContentTypes(byte[] data, HashSet<string> droppedPartNames)
        {
            if (droppedPartNames.Count == 0) return data;
            XDocument document = Wml.LoadXml(data);
            if (document == null || document.Root == null) return data;
            XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
            List<XElement> dead = new List<XElement>();
            foreach (XElement entry in document.Root.Elements())
            {
                if (entry.Name != types + "Override") continue;
                string partName = (string)entry.Attribute("PartName");
                if (partName != null && droppedPartNames.Contains(partName)) dead.Add(entry);
            }
            if (dead.Count == 0) return data;
            for (int i = 0; i < dead.Count; i++) dead[i].Remove();
            return Wml.SerializeXml(document);
        }

        /// <summary>
        /// 清理指向已删除关系的元素：超链接只去掉 r:id 保留文字，嵌入对象与控件整体移除，
        /// 其余命名空间里无法判断的元素只摘掉 r:id，避免破坏它自己的结构。
        /// Cleans up elements that pointed at a removed relationship: a hyperlink loses only its r:id and keeps
        /// its text, an embedded object or an activeX control is removed whole, and for elements outside the
        /// WordprocessingML namespace only the r:id is taken off so that their own structure stays intact.
        /// </summary>
        /// <param name="root">部件根元素。The part root element.</param>
        /// <param name="ids">失效的关系 ID。The dead relationship ids.</param>
        private static void RemoveDanglingReferences(XElement root, HashSet<string> ids)
        {
            List<XElement> objects = new List<XElement>();
            List<XElement> hyperlinks = new List<XElement>();
            foreach (XElement element in root.DescendantsAndSelf())
            {
                XAttribute attribute = element.Attribute(Wml.R + "id");
                if (attribute == null || !ids.Contains(attribute.Value)) continue;
                if (element.Name == Wml.W + "hyperlink") hyperlinks.Add(element);
                else objects.Add(element);
            }
            for (int i = 0; i < hyperlinks.Count; i++)
            {
                XAttribute reference = hyperlinks[i].Attribute(Wml.R + "id");
                if (reference != null) reference.Remove();
            }
            for (int i = 0; i < objects.Count; i++)
            {
                XElement target = objects[i];
                XElement container = target.Ancestors(Wml.W + "object").FirstOrDefault();
                if (container == null && target.Name.Namespace == Wml.W)
                {
                    container = target;
                }
                if (container == null)
                {
                    XAttribute reference = target.Attribute(Wml.R + "id");
                    if (reference != null) reference.Remove();
                    continue;
                }
                container.Remove();
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
                OfficePackageLimits.CopyBounded(templateStream, memory);
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
        /// Locates the officeDocument main part through the package-level _rels/.rels bytes and falls back to
        /// word/document.xml when it cannot be resolved.
        /// </summary>
        /// <param name="rootRelationships">包级关系部件字节，包内没有该条目时为 null。The bytes of the package-level relationships part, null when the package has no such entry.</param>
        /// <returns>主文档部件路径。The path of the main document part.</returns>
        private static string FindMainPart(byte[] rootRelationships)
        {
            if (rootRelationships != null)
            {
                try
                {
                    XDocument document = Wml.LoadXml(rootRelationships);
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
