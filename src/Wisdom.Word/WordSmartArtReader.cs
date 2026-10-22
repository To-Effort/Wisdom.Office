using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 文档中的一幅 SmartArt 图形：记录数据部件路径、节点数量与各节点的文本。
    /// One SmartArt graphic in a document: the data part path, the node count and the text of each node.
    /// </summary>
    public sealed class WordSmartArt
    {
        /// <summary>SmartArt 数据部件路径（如 word/diagrams/data1.xml）。The SmartArt data part path (for example word/diagrams/data1.xml).</summary>
        public string PartPath { get; internal set; }
        /// <summary>引用该图形的源部件路径（正文或页眉页脚）。The source part path referencing the graphic (body, header or footer).</summary>
        public string SourcePart { get; internal set; }
        /// <summary>各节点文本，按文档中的出现顺序。The text of each node, in document order.</summary>
        public IList<string> NodeTexts { get; private set; }
        /// <summary>节点数量。The number of nodes.</summary>
        public int NodeCount { get { return NodeTexts.Count; } }
        /// <summary>全部节点文本以换行拼接的结果。All node texts joined with line breaks.</summary>
        public string Text { get; private set; }

        /// <summary>内部构造：以数据部件、源部件与节点文本创建图形记录。Internal constructor: creates a graphic record from the data part, source part and node texts.</summary>
        /// <param name="partPath">数据部件路径。The data part path.</param>
        /// <param name="sourcePart">源部件路径。The source part path.</param>
        /// <param name="nodeTexts">节点文本列表。The node texts.</param>
        internal WordSmartArt(string partPath, string sourcePart, List<string> nodeTexts)
        {
            PartPath = partPath;
            SourcePart = sourcePart;
            NodeTexts = new ReadOnlyCollection<string>(nodeTexts);
            Text = string.Join("\n", nodeTexts.ToArray());
        }
    }

    /// <summary>
    /// SmartArt 检查器：读取 DOCX 中的 SmartArt 图形并提取节点文本（不重建图形结构）。
    /// SmartArt inspector that reads the SmartArt graphics of a DOCX package and extracts node text
    /// (the graphic structure itself is not rebuilt).
    /// </summary>
    /// <remarks>
    /// SmartArt 的图形定义分散在数据、布局、样式与配色四类部件中，统一文档模型不承载这些部件；
    /// 本检查器用于读取内容与做存在性判断，写回包时不会保留原 SmartArt 图形。
    ///
    /// A SmartArt definition is spread across data, layout, style and colour parts that the unified document
    /// model does not carry. This inspector reads the content and detects its presence; regenerated packages
    /// do not preserve the original SmartArt graphic.
    /// </remarks>
    public static class WordSmartArtReader
    {
        /// <summary>DrawingML 图表命名空间（dgm）。The DrawingML diagram namespace (dgm).</summary>
        private static readonly XNamespace Dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";

        /// <summary>读取指定 .docx 文件中的全部 SmartArt 图形。Reads every SmartArt graphic of the given .docx file.</summary>
        /// <param name="path">.docx 文件路径。The .docx file path.</param>
        /// <returns>按出现顺序排列的 SmartArt 列表；没有图形时为空列表。The SmartArt graphics in order of appearance; empty when none exist.</returns>
        public static IList<WordSmartArt> Read(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Read(File.ReadAllBytes(path));
        }

        /// <summary>读取 .docx 字节内容中的全部 SmartArt 图形。Reads every SmartArt graphic from .docx bytes.</summary>
        /// <param name="data">.docx 文件的完整字节内容。The full bytes of the .docx file.</param>
        /// <returns>按出现顺序排列的 SmartArt 列表；没有图形时为空列表。The SmartArt graphics in order of appearance; empty when none exist.</returns>
        /// <exception cref="InvalidDataException">数据不是有效的 ZIP 包时抛出。Thrown when the data is not a valid ZIP package.</exception>
        public static IList<WordSmartArt> Read(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length < 4 || data[0] != 0x50 || data[1] != 0x4B)
            {
                throw new InvalidDataException("The file is not a valid .docx package.");
            }
            List<WordSmartArt> result = new List<WordSmartArt>();
            Dictionary<string, bool> seenParts = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (MemoryStream memory = new MemoryStream(data, false))
                using (ZipArchive archive = PackageArchive.OpenForRead(memory, false))
                {
                    DocxPackage package = new DocxPackage(archive);
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        string partPath = entry.FullName.Replace('\\', '/');
                        if (!IsSmartArtSourcePart(partPath)) continue;
                        XDocument xml = package.GetXml(partPath);
                        if (xml == null || xml.Root == null) continue;
                        foreach (XElement reference in xml.Root.Descendants())
                        {
                            if (reference.Name != Dgm + "relIds") continue;
                            string relationshipId = Wml.AttributeValue(reference, Wml.R + "dm");
                            if (string.IsNullOrEmpty(relationshipId)) continue;
                            string dataPart = package.GetRelationshipTarget(partPath, relationshipId);
                            if (string.IsNullOrEmpty(dataPart) || seenParts.ContainsKey(dataPart)) continue;
                            seenParts[dataPart] = true;
                            result.Add(new WordSmartArt(dataPart, partPath, ReadNodeTexts(package.GetXml(dataPart))));
                        }
                    }
                }
            }
            catch (IOException ex)
            {
                throw new InvalidDataException("The .docx package is corrupt or truncated.", ex);
            }
            return result;
        }

        /// <summary>
        /// 判断部件是否可能承载 SmartArt 引用（正文、页眉、页脚三类 XML 部件）。
        /// Determines whether a part may carry SmartArt references (body, header and footer XML parts).
        /// </summary>
        /// <param name="partPath">部件路径。The part path.</param>
        /// <returns>可能承载引用时返回 true。True when the part may carry references.</returns>
        private static bool IsSmartArtSourcePart(string partPath)
        {
            if (!partPath.StartsWith("word/", StringComparison.OrdinalIgnoreCase)) return false;
            if (!partPath.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return false;
            string file = partPath.Substring(partPath.LastIndexOf('/') + 1);
            if (string.Equals(file, "document.xml", StringComparison.OrdinalIgnoreCase)) return true;
            return file.StartsWith("header", StringComparison.OrdinalIgnoreCase)
                || file.StartsWith("footer", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 从 SmartArt 数据部件读取节点文本：按 dgm:pt 节点逐个提取 dgm:t 内的纯文本，忽略空节点。
        /// Reads node text from a SmartArt data part: extracts the plain text inside dgm:t for each dgm:pt
        /// node and ignores empty nodes.
        /// </summary>
        /// <param name="dataXml">SmartArt 数据部件 XML，可为 null。The SmartArt data-part XML, may be null.</param>
        /// <returns>节点文本列表。The node text list.</returns>
        private static List<string> ReadNodeTexts(XDocument dataXml)
        {
            List<string> texts = new List<string>();
            if (dataXml == null || dataXml.Root == null) return texts;
            foreach (XElement point in dataXml.Root.Descendants(Dgm + "pt"))
            {
                XElement text = point.Element(Dgm + "t");
                if (text == null) continue;
                string value = text.Value.Trim();
                if (value.Length > 0) texts.Add(value);
            }
            return texts;
        }
    }
}
