using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Ppt
{
    /// <summary>
    /// PresentationML（.pptx）读写共用的 XML 辅助工具：OOXML 命名空间常量、
    /// 元素/属性安全访问、数值解析与安全序列化。
    /// Shared XML helpers for reading and writing PresentationML (.pptx): OOXML namespace constants,
    /// safe element/attribute access, number parsing, and secure serialization.
    /// </summary>
    internal static class Ppt
    {
        /// <summary>演示文稿主命名空间（presentationml/2006/main，前缀 p）。 Presentation main namespace (presentationml/2006/main, prefix p).</summary>
        public static readonly XNamespace Main = "http://schemas.openxmlformats.org/presentationml/2006/main";
        /// <summary>图形主命名空间（drawingml/2006/main，前缀 a）。 Drawing main namespace (drawingml/2006/main, prefix a).</summary>
        public static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        /// <summary>officeDocument 关系命名空间，用于 r:id、r:embed 等属性。 officeDocument relationships namespace, used for r:id, r:embed and similar attributes.</summary>
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        /// <summary>OPC 包关系命名空间（.rels 文件）。 OPC package relationships namespace (.rels files).</summary>
        public static readonly XNamespace PackageRels = "http://schemas.openxmlformats.org/package/2006/relationships";
        /// <summary>OPC 内容类型命名空间（[Content_Types].xml）。 OPC content types namespace ([Content_Types].xml).</summary>
        public static readonly XNamespace ContentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
        /// <summary>核心属性命名空间（docProps/core.xml）。 Core properties namespace (docProps/core.xml).</summary>
        public static readonly XNamespace CoreProperties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
        /// <summary>Dublin Core 元素命名空间。 Dublin Core elements namespace.</summary>
        public static readonly XNamespace Dc = "http://purl.org/dc/elements/1.1/";
        /// <summary>Dublin Core Terms 命名空间，用于 created/modified 时间戳。 Dublin Core Terms namespace, used for created/modified timestamps.</summary>
        public static readonly XNamespace DcTerms = "http://purl.org/dc/terms/";
        /// <summary>扩展属性命名空间（docProps/app.xml）。 Extended properties namespace (docProps/app.xml).</summary>
        public static readonly XNamespace ExtendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
        /// <summary>标记兼容性命名空间（mc:），用于 AlternateContent 选择/回退块。 Markup-compatibility namespace (mc:), used for AlternateContent choice and fallback blocks.</summary>
        public static readonly XNamespace MarkupCompat = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        /// <summary>PowerPoint 2010 扩展命名空间（p14:），切换时长按 p14:dur 存储。 PowerPoint 2010 extension namespace (p14:); the transition duration is stored as p14:dur.</summary>
        public static readonly XNamespace PowerPoint2010 = "http://schemas.microsoft.com/office/powerpoint/2010/main";

        /// <summary>关系类型片段：演示文稿。 Relationship type fragment: the presentation part.</summary>
        public const string TypePresentation = "/officeDocument";
        /// <summary>关系类型片段：幻灯片。 Relationship type fragment: a slide part.</summary>
        public const string TypeSlide = "/slide";
        /// <summary>关系类型片段：幻灯片版式。 Relationship type fragment: a slide layout part.</summary>
        public const string TypeSlideLayout = "/slideLayout";
        /// <summary>关系类型片段：幻灯片母版。 Relationship type fragment: a slide master part.</summary>
        public const string TypeSlideMaster = "/slideMaster";
        /// <summary>关系类型片段：备注页。 Relationship type fragment: notes slides.</summary>
        public const string TypeNotesMaster = "/notesMaster";
        /// <summary>关系类型片段：图片。 Relationship type fragment: an image part.</summary>
        public const string TypeImage = "/image";
        /// <summary>关系类型片段：主题。 Relationship type fragment: a theme part.</summary>
        public const string TypeTheme = "/theme";

        /// <summary>获取父元素下主命名空间的第一个子元素，父元素为 null 时返回 null。 Returns the first child element in the main namespace, or null when the parent is null.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">子元素本地名称。 Local child element name.</param>
        /// <returns>匹配的子元素或 null。 The matching child element or null.</returns>
        public static XElement Child(XElement parent, string name)
        {
            return parent == null ? null : parent.Element(Main + name);
        }

        /// <summary>父元素缺失时返回的空子元素序列（共享实例，避免逐次分配）。 Shared empty child sequence, avoiding a per-call allocation.</summary>
        private static readonly XElement[] EmptyElements = new XElement[0];

        /// <summary>获取父元素下主命名空间的同名子元素集合。 Returns the child elements with the given name in the main namespace.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">子元素本地名称。 Local child element name.</param>
        /// <returns>匹配的子元素集合，父元素为 null 时返回空集合。 The matching children, or an empty sequence when the parent is null.</returns>
        public static IEnumerable<XElement> Children(XElement parent, string name)
        {
            return parent == null ? EmptyElements : parent.Elements(Main + name);
        }

        /// <summary>获取父元素下图形命名空间（a:）的第一个子元素。 Returns the first child element in the drawing namespace (a:).</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">子元素本地名称。 Local child element name.</param>
        /// <returns>匹配的子元素或 null。 The matching child element or null.</returns>
        public static XElement DrawingChild(XElement parent, string name)
        {
            return parent == null ? null : parent.Element(Drawing + name);
        }

        /// <summary>获取父元素下图形命名空间（a:）的同名子元素集合。 Returns the child elements with the given name in the drawing namespace (a:).</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">子元素本地名称。 Local child element name.</param>
        /// <returns>匹配的子元素集合，父元素为 null 时返回空集合。 The matching children, or an empty sequence when the parent is null.</returns>
        public static IEnumerable<XElement> DrawingChildren(XElement parent, string name)
        {
            return parent == null ? EmptyElements : parent.Elements(Drawing + name);
        }

        /// <summary>获取父元素下标记兼容性命名空间（mc:）的第一个子元素。 Returns the first child element in the markup-compatibility namespace (mc:).</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">子元素本地名称。 Local child element name.</param>
        /// <returns>匹配的子元素或 null。 The matching child element or null.</returns>
        public static XElement MarkupCompatChild(XElement parent, string name)
        {
            return parent == null ? null : parent.Element(MarkupCompat + name);
        }

        /// <summary>读取元素的字符串属性值，元素或属性不存在时返回 null。 Reads a string attribute value, returning null when the element or attribute is missing.</summary>
        /// <param name="element">目标元素。 Target element.</param>
        /// <param name="name">属性本地名称。 Local attribute name.</param>
        /// <returns>属性值或 null。 The attribute value or null.</returns>
        public static string Attribute(XElement element, string name)
        {
            if (element == null) return null;
            XAttribute attribute = element.Attribute(name);
            return attribute == null ? null : attribute.Value;
        }

        /// <summary>读取元素的完整名称属性值（可带命名空间）。 Reads an attribute value using a full XName (namespace allowed).</summary>
        /// <param name="element">目标元素。 Target element.</param>
        /// <param name="name">属性完整名称。 Full attribute name including namespace.</param>
        /// <returns>属性值或 null。 The attribute value or null.</returns>
        public static string Attribute(XElement element, XName name)
        {
            if (element == null) return null;
            XAttribute attribute = element.Attribute(name);
            return attribute == null ? null : attribute.Value;
        }

        /// <summary>读取整数属性，缺失或非法时返回回退值。 Reads an integer attribute, returning the fallback when missing or invalid.</summary>
        /// <param name="element">目标元素。 Target element.</param>
        /// <param name="name">属性本地名称。 Local attribute name.</param>
        /// <param name="fallback">回退值。 Fallback value.</param>
        /// <returns>属性值或回退值。 The attribute value or the fallback.</returns>
        public static int IntAttribute(XElement element, string name, int fallback)
        {
            return ParseInt(Attribute(element, name), fallback);
        }

        /// <summary>以不变区域设置解析 32 位整数，失败时返回回退值。 Parses a 32-bit integer with the invariant culture, returning the fallback on failure.</summary>
        /// <param name="value">待解析字符串。 String to parse.</param>
        /// <param name="fallback">回退值。 Fallback value.</param>
        /// <returns>解析结果或回退值。 The parsed value or the fallback.</returns>
        public static int ParseInt(string value, int fallback)
        {
            int result;
            if (!string.IsNullOrEmpty(value) && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            return fallback;
        }

        /// <summary>以不变区域设置解析双精度浮点数，失败时返回回退值。 Parses a double with the invariant culture, returning the fallback on failure.</summary>
        /// <param name="value">待解析字符串。 String to parse.</param>
        /// <param name="fallback">回退值。 Fallback value.</param>
        /// <returns>解析结果或回退值。 The parsed value or the fallback.</returns>
        public static double ParseDouble(string value, double fallback)
        {
            double result;
            if (!string.IsNullOrEmpty(value) && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            return fallback;
        }

        /// <summary>解析 OOXML 布尔值（"1"/"true" 视为真）。 Parses an OOXML boolean ("1"/"true" count as true).</summary>
        /// <param name="value">待解析字符串。 String to parse.</param>
        /// <param name="fallback">回退值。 Fallback value.</param>
        /// <returns>解析结果或回退值。 The parsed value or the fallback.</returns>
        public static bool ParseBool(string value, bool fallback)
        {
            if (string.IsNullOrEmpty(value)) return fallback;
            return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>解析颜色字符串（RGB 或 ARGB 十六进制，ARGB 时忽略前两位 Alpha）。 Parses an RGB or ARGB hexadecimal color, ignoring the leading alpha pair of ARGB values.</summary>
        /// <param name="value">颜色文本。 Color text.</param>
        /// <param name="color">解析结果；失败时为黑色。 The parsed color; black on failure.</param>
        /// <returns>解析成功返回 true。 True when parsing succeeded.</returns>
        public static bool TryParseColor(string value, out OfficeColor color)
        {
            color = OfficeColor.Black;
            if (string.IsNullOrEmpty(value)) return false;
            string hex = value;
            if (hex.Length == 8) hex = hex.Substring(2);
            return OfficeColor.TryParse(hex, out color);
        }

        /// <summary>将浮点数格式化为不变区域设置的 G17 文本。 Formats a double as invariant-culture G17 text.</summary>
        /// <param name="value">待格式化的数值。 Value to format.</param>
        /// <returns>格式化后的字符串。 The formatted string.</returns>
        public static string ToInvariant(double value)
        {
            return value.ToString("G17", CultureInfo.InvariantCulture);
        }

        /// <summary>将磅转换为 OOXML 使用的整数 EMU。 Converts points to the integer EMU value used by OOXML.</summary>
        /// <param name="points">磅值。 Value in points.</param>
        /// <returns>EMU 整数值。 The EMU value.</returns>
        public static int Emu(double points)
        {
            return UnitConverter.RoundToInt(UnitConverter.PointsToEmu(points));
        }

        /// <summary>将 OOXML 的整数 EMU 转换为磅。 Converts an integer EMU value back to points.</summary>
        /// <param name="emu">EMU 值。 The EMU value.</param>
        /// <returns>磅值。 Value in points.</returns>
        public static double ToPoints(long emu)
        {
            return UnitConverter.EmuToPoints(emu);
        }

        /// <summary>EMU 转磅（小数 EMU，供组合坐标换算后的结果直接使用）。Converts EMU to points, accepting a fractional EMU value.</summary>
        /// <param name="emu">EMU 值。The EMU value.</param>
        /// <returns>磅值。The value in points.</returns>
        public static double ToPoints(double emu)
        {
            return UnitConverter.EmuToPoints(emu);
        }

        /// <summary>
        /// 从字节数组安全加载 XML 文档：禁止 DTD 解析且禁用外部实体解析，防止 XML 注入攻击。
        /// Loads an XML document safely from a byte array: DTD processing is prohibited and external
        /// entity resolution is disabled to prevent XML injection attacks.
        /// </summary>
        /// <param name="data">XML 字节内容。 XML byte content.</param>
        /// <returns>解析后的文档；数据为空时返回 null。 The parsed document, or null when the data is empty.</returns>
        public static XDocument LoadXml(byte[] data)
        {
            if (data == null || data.Length == 0) return null;
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Prohibit;
            settings.XmlResolver = null;
            using (MemoryStream stream = new MemoryStream(data, false))
            using (XmlReader reader = XmlReader.Create(stream, settings))
            {
                return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
            }
        }

        /// <summary>将 XML 文档序列化为无 BOM 的 UTF-8 字节数组，不缩进。 Serializes an XML document to BOM-less UTF-8 bytes without indentation.</summary>
        /// <param name="document">待序列化的文档。 Document to serialize.</param>
        /// <returns>序列化后的字节内容。 The serialized bytes.</returns>
        public static byte[] SerializeXml(XDocument document)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = document.Declaration == null;
            // 形状里的文字同样可能带着 XML 不接受的控制字符，先清掉再序列化，避免整个包写不出来。
            // Shape text can carry characters XML forbids too; they are dropped before serialising so the
            // package stays writable.
            OfficeXmlText.Sanitize(document);
            using (MemoryStream stream = new MemoryStream())
            {
                using (XmlWriter writer = XmlWriter.Create(stream, settings))
                {
                    document.Save(writer);
                }
                return stream.ToArray();
            }
        }
    }

    /// <summary>
    /// .pptx OPC 包的只读访问器：按部件路径读取 XML/原始字节，并解析、缓存 _rels 关系。
    /// Read-only accessor for a .pptx OPC package: fetches XML or raw bytes by part path and parses
    /// and caches _rels relationships.
    /// </summary>
    internal sealed class PptPackage
    {
        /// <summary>底层 zip 归档（由调用方打开并负责关闭）。 Underlying zip archive (opened and closed by the caller).</summary>
        private readonly ZipArchive _archive;
        /// <summary>关系缓存：源部件路径 → 扁平化的 Id/Type/Target/TargetMode 四元组数组。 Relationship cache mapping a source part path to flattened Id/Type/Target/TargetMode quadruples.</summary>
        private readonly Dictionary<string, string[]> _relationships =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        /// <summary>累计读取的字节数，用于包大小限额校验。 Running total of bytes read, used for the package size quota check.</summary>
        private long _totalBytes;

        /// <summary>基于已打开的 zip 归档创建包访问器。 Creates a package accessor over an already opened zip archive.</summary>
        /// <param name="archive">.pptx 包对应的 zip 归档。 The zip archive of the .pptx package.</param>
        public PptPackage(ZipArchive archive)
        {
            _archive = archive;
        }

        /// <summary>读取指定部件并解析为 XML 文档；部件不存在或 XML 非法时返回 null。 Reads a part and parses it as XML, returning null when the part is missing or invalid.</summary>
        /// <param name="partPath">包内部件路径（如 ppt/presentation.xml）。 Part path inside the package.</param>
        /// <returns>解析后的文档或 null。 The parsed document or null.</returns>
        public XDocument GetXml(string partPath)
        {
            byte[] data = GetBytes(partPath);
            if (data == null) return null;
            try
            {
                return Ppt.LoadXml(data);
            }
            catch (XmlException)
            {
                return null;
            }
        }

        /// <summary>读取部件原始字节；精确路径找不到时忽略大小写与斜杠差异逐条匹配。 Reads raw part bytes, falling back to a case- and slash-insensitive scan when the exact path is missing.</summary>
        /// <param name="partPath">包内部件路径。 Part path inside the package.</param>
        /// <returns>部件字节内容或 null。 The part bytes or null.</returns>
        public byte[] GetBytes(string partPath)
        {
            if (string.IsNullOrEmpty(partPath)) return null;
            ZipArchiveEntry entry = _archive.GetEntry(partPath);
            if (entry == null)
            {
                string normalized = partPath.Replace('\\', '/');
                foreach (ZipArchiveEntry candidate in _archive.Entries)
                {
                    if (string.Equals(candidate.FullName.Replace('\\', '/'), normalized, StringComparison.OrdinalIgnoreCase))
                    {
                        entry = candidate;
                        break;
                    }
                }
            }
            if (entry == null) return null;
            return OfficePackageLimits.ReadEntryBytes(entry, ref _totalBytes);
        }

        /// <summary>按关系类型后缀查找第一个内部关系的目标部件路径（跳过 External 关系）。 Finds the target of the first internal relationship matching a type suffix, skipping External relationships.</summary>
        /// <param name="partPath">源部件路径（根关系传空字符串）。 Source part path (empty string for package-root relationships).</param>
        /// <param name="typeSuffix">关系类型 URL 的结尾片段。 Trailing fragment of the relationship type URL.</param>
        /// <returns>解析后的目标部件路径或 null。 The resolved target path or null.</returns>
        public string GetRelationshipTargetByType(string partPath, string typeSuffix)
        {
            List<string> all = GetRelationshipTargetsByType(partPath, typeSuffix);
            return all.Count == 0 ? null : all[0];
        }

        /// <summary>按关系类型后缀查找全部内部关系的目标部件路径，按关系声明顺序返回。 Finds the targets of all internal relationships matching a type suffix, in declaration order.</summary>
        /// <param name="partPath">源部件路径。 Source part path.</param>
        /// <param name="typeSuffix">关系类型 URL 的结尾片段。 Trailing fragment of the relationship type URL.</param>
        /// <returns>解析后的目标部件路径列表。 The resolved target paths.</returns>
        public List<string> GetRelationshipTargetsByType(string partPath, string typeSuffix)
        {
            List<string> result = new List<string>();
            string[] items = LoadRelationships(partPath);
            for (int i = 0; i + 3 < items.Length; i += 4)
            {
                if (string.Equals(items[i + 3], "External", StringComparison.OrdinalIgnoreCase)) continue;
                string type = items[i + 1];
                if (type.EndsWith(typeSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(ResolvePartPath(partPath, items[i + 2]));
                }
            }
            return result;
        }

        /// <summary>按关系 Id（如 rId1）查找内部关系的目标部件路径；外部关系返回 null。 Finds the target of an internal relationship by Id; External relationships return null.</summary>
        /// <param name="partPath">源部件路径。 Source part path.</param>
        /// <param name="relationshipId">关系 Id。 Relationship Id.</param>
        /// <returns>解析后的目标部件路径或 null。 The resolved target path or null.</returns>
        public string GetRelationshipTarget(string partPath, string relationshipId)
        {
            if (string.IsNullOrEmpty(relationshipId)) return null;
            string[] items = LoadRelationships(partPath);
            for (int i = 0; i + 3 < items.Length; i += 4)
            {
                if (!string.Equals(items[i], relationshipId, StringComparison.Ordinal)) continue;
                if (string.Equals(items[i + 3], "External", StringComparison.OrdinalIgnoreCase)) return null;
                return ResolvePartPath(partPath, items[i + 2]);
            }
            return null;
        }

        /// <summary>把关系目标解析为包内绝对路径。 Resolves a relationship target into an absolute part path.</summary>
        /// <param name="partPath">源部件路径（根关系传空字符串）。 Source part path (empty string for the package root).</param>
        /// <param name="target">关系中的相对或绝对目标。 The relative or absolute target from the relationship.</param>
        /// <returns>包内绝对部件路径；目标为空时返回 null。 The absolute part path, or null when the target is empty.</returns>
        public static string ResolvePartPath(string partPath, string target)
        {
            if (string.IsNullOrEmpty(target)) return null;
            string result = target.Replace('\\', '/');
            if (result.StartsWith("/", StringComparison.Ordinal)) return result.Substring(1);
            string baseDirectory = string.Empty;
            string source = (partPath ?? string.Empty).Replace('\\', '/');
            int slash = source.LastIndexOf('/');
            if (slash >= 0) baseDirectory = source.Substring(0, slash + 1);
            List<string> segments = new List<string>();
            if (baseDirectory.Length > 0)
            {
                foreach (string head in baseDirectory.TrimEnd('/').Split('/'))
                {
                    if (head.Length > 0) segments.Add(head);
                }
            }
            foreach (string segment in result.Split('/'))
            {
                if (segment == "." || segment.Length == 0) continue;
                if (segment == "..")
                {
                    if (segments.Count > 0) segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                segments.Add(segment);
            }
            return string.Join("/", segments.ToArray());
        }

        /// <summary>解析并缓存 .rels 文件，返回扁平化的 Id/Type/Target/TargetMode 四元组数组；查询直接扫描数组，不再重建 XElement。 Parses and caches the .rels file as flattened Id/Type/Target/TargetMode quadruples; queries scan the array instead of rebuilding elements.</summary>
        /// <param name="partPath">源部件路径。 Source part path.</param>
        /// <returns>关系四元组扁平数组。 The flattened relationship array.</returns>
        private string[] LoadRelationships(string partPath)
        {
            string key = partPath ?? string.Empty;
            string[] cached;
            if (_relationships.TryGetValue(key, out cached)) return cached;

            string baseDirectory = string.Empty;
            string fileName = key;
            int slash = fileName.LastIndexOf('/');
            if (slash >= 0)
            {
                baseDirectory = fileName.Substring(0, slash + 1);
                fileName = fileName.Substring(slash + 1);
            }
            XDocument rels = GetXml(baseDirectory + "_rels/" + fileName + ".rels");
            List<string> items = new List<string>();
            if (rels != null && rels.Root != null)
            {
                foreach (XElement relationship in rels.Root.Elements(Ppt.PackageRels + "Relationship"))
                {
                    items.Add((string)relationship.Attribute("Id") ?? string.Empty);
                    items.Add((string)relationship.Attribute("Type") ?? string.Empty);
                    items.Add((string)relationship.Attribute("Target") ?? string.Empty);
                    items.Add((string)relationship.Attribute("TargetMode") ?? string.Empty);
                }
            }
            cached = items.ToArray();
            _relationships[key] = cached;
            return cached;
        }
    }
}
