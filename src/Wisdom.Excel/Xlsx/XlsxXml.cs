using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// SpreadsheetML（.xlsx）读写共用的 XML 辅助工具：OOXML 命名空间常量、
    /// 元素/属性安全访问、数值解析与 XML 序列化等。
    /// Shared XML helpers for reading and writing SpreadsheetML (.xlsx): OOXML namespace constants,
    /// safe element/attribute access, number parsing and XML serialization.
    /// </summary>
    internal static class Xlsx
    {
        /// <summary>电子表格主命名空间（spreadsheetml/2006/main）。 Spreadsheet main namespace (spreadsheetml/2006/main).</summary>
        public static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        /// <summary>officeDocument 关系命名空间，用于 r:id 等属性。 officeDocument relationships namespace, used for r:id and similar attributes.</summary>
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
        /// <summary>XML Schema 实例命名空间，用于 xsi:type 属性。 XML Schema instance namespace, used for the xsi:type attribute.</summary>
        public static readonly XNamespace Xsi = "http://www.w3.org/2001/XMLSchema-instance";

        /// <summary>获取父元素下主命名空间的第一个子元素，父元素为 null 时返回 null。 Returns the first child element in the main namespace, or null when the parent is null.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">子元素本地名称（不含命名空间前缀）。 Local child element name (without namespace prefix).</param>
        /// <returns>匹配的子元素，找不到时返回 null。 The matching child element, or null when not found.</returns>
        public static XElement Child(XElement parent, string name)
        {
            return parent == null ? null : parent.Element(Main + name);
        }

        /// <summary>父元素缺失时返回的空子元素序列（共享实例，避免逐次分配）。 Shared empty child sequence, avoiding a per-call allocation.</summary>
        private static readonly XElement[] EmptyElements = new XElement[0];

        /// <summary>获取父元素下主命名空间的同名子元素集合，父元素为 null 时返回空集合。 Returns the child elements with the given name in the main namespace, or an empty sequence when the parent is null.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">子元素本地名称。 Local child element name.</param>
        /// <returns>匹配的子元素集合。 The matching child elements.</returns>
        public static IEnumerable<XElement> Children(XElement parent, string name)
        {
            return parent == null ? EmptyElements : parent.Elements(Main + name);
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

        /// <summary>读取元素的字符串属性值（完整 XName，可带命名空间），元素或属性不存在时返回 null。 Reads a string attribute value using a full XName (namespace allowed), returning null when the element or attribute is missing.</summary>
        /// <param name="element">目标元素。 Target element.</param>
        /// <param name="name">属性完整名称（含命名空间）。 Full attribute name including namespace.</param>
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
        /// <param name="fallback">解析失败时的回退值。 Fallback value used when parsing fails.</param>
        /// <returns>属性整数值或回退值。 The integer attribute value or the fallback.</returns>
        public static int IntAttribute(XElement element, string name, int fallback)
        {
            return ParseInt(Attribute(element, name), fallback);
        }

        /// <summary>读取浮点属性，缺失或非法时返回回退值。 Reads a double attribute, returning the fallback when missing or invalid.</summary>
        /// <param name="element">目标元素。 Target element.</param>
        /// <param name="name">属性本地名称。 Local attribute name.</param>
        /// <param name="fallback">解析失败时的回退值。 Fallback value used when parsing fails.</param>
        /// <returns>属性浮点值或回退值。 The double attribute value or the fallback.</returns>
        public static double DoubleAttribute(XElement element, string name, double fallback)
        {
            return ParseDouble(Attribute(element, name), fallback);
        }

        /// <summary>以不变区域设置解析 32 位整数，失败时返回回退值。 Parses a 32-bit integer with the invariant culture, returning the fallback on failure.</summary>
        /// <param name="value">待解析字符串。 String to parse.</param>
        /// <param name="fallback">解析失败时的回退值。 Fallback value used when parsing fails.</param>
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

        /// <summary>以不变区域设置解析双精度浮点数，失败时返回回退值。 Parses a double-precision number with the invariant culture, returning the fallback on failure.</summary>
        /// <param name="value">待解析字符串。 String to parse.</param>
        /// <param name="fallback">解析失败时的回退值。 Fallback value used when parsing fails.</param>
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

        /// <summary>解析 OOXML 布尔值（"1"/"true" 视为真），缺失时返回回退值。 Parses an OOXML boolean ("1"/"true" count as true), returning the fallback when missing.</summary>
        /// <param name="value">待解析字符串。 String to parse.</param>
        /// <param name="fallback">解析失败时的回退值。 Fallback value used when parsing fails.</param>
        /// <returns>解析结果或回退值。 The parsed value or the fallback.</returns>
        public static bool ParseBool(string value, bool fallback)
        {
            if (string.IsNullOrEmpty(value)) return fallback;
            return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>将浮点数格式化为不变区域设置的 G17 文本，保证读写往返精度。 Formats a double as invariant-culture G17 text so that values survive a read/write round trip.</summary>
        /// <param name="value">待格式化的数值。 Value to format.</param>
        /// <returns>格式化后的字符串。 The formatted string.</returns>
        public static string ToInvariant(double value)
        {
            return value.ToString("G17", CultureInfo.InvariantCulture);
        }

        /// <summary>解析颜色字符串（RGB 或 ARGB 十六进制，ARGB 时忽略前两位 Alpha）。 Parses a color string (RGB or ARGB hexadecimal; the leading alpha pair of ARGB values is ignored).</summary>
        /// <param name="value">颜色文本。 Color text.</param>
        /// <param name="color">解析结果；失败时为黑色。 Parsing result; black when parsing fails.</param>
        /// <returns>解析成功返回 true。 True when parsing succeeded.</returns>
        public static bool TryParseColor(string value, out OfficeColor color)
        {
            color = OfficeColor.Black;
            if (string.IsNullOrEmpty(value)) return false;
            string hex = value;
            if (hex.Length == 8) hex = hex.Substring(2);
            return OfficeColor.TryParse(hex, out color);
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

        /// <summary>
        /// 为流式读取创建安全 XmlReader：禁止 DTD 且不解析外部实体，其余保持默认（不忽略空白，保证与 XDocument 读入一致）。
        /// Creates a safe XmlReader for streaming: DTD processing is prohibited and external entities are never resolved,
        /// everything else stays default (whitespace is not ignored, so content matches the XDocument path).
        /// </summary>
        /// <param name="input">部件输入流。 The part input stream.</param>
        /// <returns>定位到文档起点的读取器。 A reader positioned at the start of the document.</returns>
        public static XmlReader CreateSafeReader(Stream input)
        {
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.DtdProcessing = DtdProcessing.Prohibit;
            settings.XmlResolver = null;
            return XmlReader.Create(input, settings);
        }

        /// <summary>将 XML 文档序列化为无 BOM 的 UTF-8 字节数组，不缩进、按需保留 XML 声明。 Serializes an XML document to a BOM-less UTF-8 byte array without indentation, keeping the XML declaration only when the document has one.</summary>
        /// <param name="document">待序列化的文档。 Document to serialize.</param>
        /// <returns>序列化后的字节内容。 The serialized byte content.</returns>
        public static byte[] SerializeXml(XDocument document)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                using (XmlWriter writer = CreateStreamingWriter(stream, document.Declaration == null))
                {
                    document.Save(writer);
                }
                return stream.ToArray();
            }
        }

        /// <summary>创建流式 XmlWriter：无 BOM 的 UTF-8、不缩进，输出字节与 SerializeXml 一致。 Creates a streaming XmlWriter: BOM-less UTF-8, no indentation, producing bytes identical to SerializeXml.</summary>
        /// <param name="output">目标流；由调用方负责关闭。 Target stream, closed by the caller.</param>
        /// <param name="omitXmlDeclaration">是否省略 XML 声明。 Whether to omit the XML declaration.</param>
        /// <returns>写入目标流的 XmlWriter。 The XmlWriter writing into the target stream.</returns>
        public static XmlWriter CreateStreamingWriter(Stream output, bool omitXmlDeclaration)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = omitXmlDeclaration;
            settings.CloseOutput = false;
            return XmlWriter.Create(output, settings);
        }
    }

    /// <summary>
    /// .xlsx OPC 包的只读访问器：按部件路径读取 XML/原始字节，并解析、缓存 _rels 关系。
    /// 读取字节受 OfficePackageLimits 的总大小限额约束。
    /// Read-only accessor for an .xlsx OPC package: fetches XML or raw bytes by part path and parses
    /// and caches _rels relationships. Byte reads are bounded by the OfficePackageLimits total-size quota.
    /// </summary>
    internal sealed class XlsxPackage
    {
        /// <summary>底层 zip 归档（由调用方打开并负责关闭）。 Underlying zip archive (opened and closed by the caller).</summary>
        private readonly ZipArchive _archive;
        /// <summary>关系缓存：源部件路径 → 扁平化的 Id/Type/Target/TargetMode 四元组数组。 Relationship cache mapping a source part path to flattened Id/Type/Target/TargetMode quadruples.</summary>
        private readonly Dictionary<string, string[]> _relationships =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        /// <summary>累计读取的字节数，用于包大小限额校验。 Running total of bytes read, used for the package size quota check.</summary>
        private long _totalBytes;

        /// <summary>基于已打开的 zip 归档创建包访问器。 Creates a package accessor over an already opened zip archive.</summary>
        /// <param name="archive">.xlsx 包对应的 zip 归档。 The zip archive of the .xlsx package.</param>
        public XlsxPackage(ZipArchive archive)
        {
            _archive = archive;
        }

        /// <summary>读取指定部件并解析为 XML 文档；部件不存在或 XML 非法时返回 null。 Reads a part and parses it as an XML document, returning null when the part is missing or the XML is invalid.</summary>
        /// <param name="partPath">包内部件路径（如 xl/workbook.xml）。 Part path inside the package (for example xl/workbook.xml).</param>
        /// <returns>解析后的文档或 null。 The parsed document or null.</returns>
        public XDocument GetXml(string partPath)
        {
            byte[] data = GetBytes(partPath);
            if (data == null) return null;
            try
            {
                return Xlsx.LoadXml(data);
            }
            catch (XmlException)
            {
                return null;
            }
        }

        /// <summary>读取部件原始字节；路径找不到时忽略大小写与斜杠差异逐条匹配，仍不存在则返回 null。 Reads raw part bytes; when the exact path is missing it scans entries ignoring case and slash differences, and returns null if still not found.</summary>
        /// <param name="partPath">包内部件路径。 Part path inside the package.</param>
        /// <returns>部件字节内容或 null。 The part byte content or null.</returns>
        public byte[] GetBytes(string partPath)
        {
            ZipArchiveEntry entry = FindEntry(partPath);
            if (entry == null) return null;
            return OfficePackageLimits.ReadEntryBytes(entry, ref _totalBytes);
        }

        /// <summary>
        /// 以只读流方式打开部件，供大部件（如工作表）流式解析，避免整棵树驻留内存。
        /// 返回的流按 OfficePackageLimits 的上限限量，并在关闭时把已读字节计入包总量。
        /// Opens a part as a read-only stream so large parts (such as worksheets) can be parsed without keeping a
        /// whole tree in memory. The returned stream enforces the OfficePackageLimits allowance and reports the
        /// bytes it read into the package total when it is closed.
        /// </summary>
        /// <param name="partPath">包内部件路径。 Part path inside the package.</param>
        /// <returns>限量只读流；部件不存在时返回 null。 The quota-bounded read stream, or null when the part is missing.</returns>
        /// <exception cref="InvalidDataException">部件解压大小会突破安全上限。 The part would exceed the safety limits.</exception>
        public Stream OpenRead(string partPath)
        {
            ZipArchiveEntry entry = FindEntry(partPath);
            if (entry == null) return null;
            long limit = OfficePackageLimits.CheckEntry(entry, _totalBytes);
            return new LimitedEntryStream(entry.Open(), entry.FullName, limit, delta => _totalBytes += delta);
        }

        /// <summary>按部件路径查找 zip 条目：先精确匹配，再忽略大小写与斜杠差异逐条匹配。 Locates a zip entry by part path: an exact match first, then a case- and slash-insensitive scan.</summary>
        /// <param name="partPath">包内部件路径。 Part path inside the package.</param>
        /// <returns>匹配的条目，不存在时返回 null。 The matching entry, or null when absent.</returns>
        private ZipArchiveEntry FindEntry(string partPath)
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
            return entry;
        }

        /// <summary>按关系类型后缀查找第一个内部关系的目标部件路径（跳过 External 关系）。 Finds the target part path of the first internal relationship matching a type suffix, skipping External relationships.</summary>
        /// <param name="partPath">持有关系集合的源部件路径（根关系传空字符串）。 Source part path that owns the relationship collection (empty string for package-root relationships).</param>
        /// <param name="typeSuffix">关系类型 URL 的结尾片段（如 /officeDocument）。 Trailing fragment of the relationship type URL (for example /officeDocument).</param>
        /// <returns>解析后的目标部件路径，未找到时返回 null。 The resolved target part path, or null when nothing matched.</returns>
        public string GetRelationshipTargetByType(string partPath, string typeSuffix)
        {
            string[] items = LoadRelationships(partPath);
            for (int i = 0; i + 3 < items.Length; i += 4)
            {
                if (string.Equals(items[i + 3], "External", StringComparison.OrdinalIgnoreCase)) continue;
                string type = items[i + 1];
                if (type.EndsWith(typeSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    return ResolvePartPath(partPath, items[i + 2]);
                }
            }
            return null;
        }

        /// <summary>按关系 Id（如 rId1）查找内部关系的目标部件路径；外部关系返回 null。 Finds the target part path of an internal relationship by Id (for example rId1); External relationships return null.</summary>
        /// <param name="partPath">持有关系集合的源部件路径。 Source part path that owns the relationship collection.</param>
        /// <param name="relationshipId">关系 Id。 Relationship Id.</param>
        /// <returns>解析后的目标部件路径，未找到时返回 null。 The resolved target part path, or null when nothing matched.</returns>
        public string GetRelationshipTarget(string partPath, string relationshipId)
        {
            return GetRelationshipTarget(partPath, relationshipId, false);
        }

        /// <summary>按关系 Id 查找目标；includeExternal 为 true 时返回 External 关系的原始地址（供超链接使用），否则 External 关系返回 null。 Finds a relationship target by Id; when includeExternal is true the raw address of an External relationship is returned (for hyperlinks), otherwise External relationships return null.</summary>
        /// <param name="partPath">持有关系集合的源部件路径。 Source part path that owns the relationship collection.</param>
        /// <param name="relationshipId">关系 Id。 Relationship Id.</param>
        /// <param name="includeExternal">是否包含 External 关系。 Whether External relationships are included.</param>
        /// <returns>目标地址或 null。 The target address, or null.</returns>
        public string GetRelationshipTarget(string partPath, string relationshipId, bool includeExternal)
        {
            if (string.IsNullOrEmpty(relationshipId)) return null;
            string[] items = LoadRelationships(partPath);
            for (int i = 0; i + 3 < items.Length; i += 4)
            {
                if (!string.Equals(items[i], relationshipId, StringComparison.Ordinal)) continue;
                bool external = string.Equals(items[i + 3], "External", StringComparison.OrdinalIgnoreCase);
                if (external)
                {
                    return includeExternal ? items[i + 2] : null;
                }
                return ResolvePartPath(partPath, items[i + 2]);
            }
            return null;
        }

        /// <summary>
        /// 获取部件的关系条目扁平数组：首次访问时解析 .rels 文件，把
        /// Id/Type/Target/TargetMode 四元组缓存为字符串数组，后续查询直接在数组上扫描，
        /// 不再重建 XElement。
        /// Returns the flat relationship array of a part: the .rels file is parsed on first access and the
        /// Id/Type/Target/TargetMode quadruples are cached as a string array; later queries scan that array
        /// instead of rebuilding XElement objects.
        /// </summary>
        /// <param name="partPath">源部件路径。 Source part path.</param>
        /// <returns>按四元组顺序扁平化的关系数组。 The flattened relationship array.</returns>
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
                foreach (XElement relationship in rels.Root.Elements(Xlsx.PackageRels + "Relationship"))
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

        /// <summary>
        /// 将关系 Target 解析为包内规范部件路径：以 / 开头视为包根绝对路径，
        /// 否则相对源部件所在目录解析，并消解 . 与 .. 片段。
        /// Resolves a relationship Target into a canonical part path: a leading slash means a package-root
        /// absolute path, otherwise the target is resolved against the source part directory, collapsing . and .. segments.
        /// </summary>
        /// <param name="sourcePartPath">源部件路径。 Source part path.</param>
        /// <param name="target">关系 Target 原始文本。 Raw relationship Target text.</param>
        /// <returns>规范化的部件路径；target 为空时返回 null。 The normalized part path, or null when the target is empty.</returns>
        public static string ResolvePartPath(string sourcePartPath, string target)
        {
            if (string.IsNullOrEmpty(target)) return null;
            string path = target.Replace('\\', '/');
            if (path.StartsWith("/", StringComparison.Ordinal))
            {
                return path.TrimStart('/');
            }
            string baseDirectory = string.Empty;
            int slash = sourcePartPath == null ? -1 : sourcePartPath.LastIndexOf('/');
            if (slash >= 0) baseDirectory = sourcePartPath.Substring(0, slash + 1);
            string combined = baseDirectory + path;
            List<string> segments = new List<string>();
            string[] parts = combined.Split('/');
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i].Length == 0 || parts[i] == ".") continue;
                if (parts[i] == "..")
                {
                    if (segments.Count > 0) segments.RemoveAt(segments.Count - 1);
                    continue;
                }
                segments.Add(parts[i]);
            }
            return string.Join("/", segments.ToArray());
        }
    }

    /// <summary>
    /// 只读解压限量流：累计读取字节一旦超过部件上限立即抛 InvalidDataException，
    /// 并把每次读到的字节回报给包的总量计数器，使流式读取同样受解压炸弹防护约束。
    /// A read-only decompression stream with a quota: it throws InvalidDataException as soon as the bytes read
    /// exceed the part allowance, and reports every chunk to the package's running total so streaming reads stay
    /// covered by the decompression-bomb guard.
    /// </summary>
    internal sealed class LimitedEntryStream : Stream
    {
        /// <summary>被包装的原始条目流。 The wrapped entry stream.</summary>
        private readonly Stream _inner;
        /// <summary>部件路径，用于异常文本。 The part path used in the error message.</summary>
        private readonly string _name;
        /// <summary>本部件允许的解压字节上限。 The decompressed byte allowance for this part.</summary>
        private readonly long _limit;
        /// <summary>把已读字节回报给包总量计数器的回调。 Callback reporting the bytes read to the package total.</summary>
        private readonly Action<long> _report;
        /// <summary>已读取字节数。 Bytes read so far.</summary>
        private long _read;

        /// <summary>创建限量只读流。 Creates the quota-bounded read stream.</summary>
        /// <param name="inner">原始条目流。 The raw entry stream.</param>
        /// <param name="name">部件路径。 The part path.</param>
        /// <param name="limit">解压字节上限。 The decompressed byte limit.</param>
        /// <param name="report">字节消耗回报。 Reports consumed bytes.</param>
        public LimitedEntryStream(Stream inner, string name, long limit, Action<long> report)
        {
            _inner = inner;
            _name = name;
            _limit = limit;
            _report = report;
        }

        /// <summary>该流可读。 The stream is readable.</summary>
        public override bool CanRead { get { return true; } }
        /// <summary>解压流不可定位。 The decompression stream cannot be seeked.</summary>
        public override bool CanSeek { get { return false; } }
        /// <summary>解压流不可写。 The decompression stream is not writable.</summary>
        public override bool CanWrite { get { return false; } }
        /// <summary>解压流长度未知。 The decompressed length is unknown.</summary>
        public override long Length { get { throw new NotSupportedException(); } }
        /// <summary>解压流不可定位。 The decompression stream cannot be seeked.</summary>
        public override long Position { get { throw new NotSupportedException(); } set { throw new NotSupportedException(); } }

        /// <summary>读取下一块字节，超过上限立即拒绝。 Reads the next chunk, rejecting it once the allowance is exceeded.</summary>
        /// <param name="buffer">目标缓冲区。 Target buffer.</param>
        /// <param name="offset">写入起点。 Write offset.</param>
        /// <param name="count">请求字节数。 Requested byte count.</param>
        /// <returns>实际读取字节数，结束时为 0。 Bytes actually read, or 0 at the end.</returns>
        /// <exception cref="InvalidDataException">解压字节超过部件上限。 The decompressed bytes exceed the part allowance.</exception>
        public override int Read(byte[] buffer, int offset, int count)
        {
            int read = _inner.Read(buffer, offset, count);
            if (read > 0)
            {
                _read += read;
                if (_read > _limit)
                {
                    throw new InvalidDataException("文档部件解压后超过安全上限: " + _name);
                }
                if (_report != null) _report(read);
            }
            return read;
        }

        /// <summary>读取单字节，超过上限立即拒绝。 Reads one byte, rejecting it once the allowance is exceeded.</summary>
        /// <returns>字节值，结束时为 -1。 The byte value, or -1 at the end.</returns>
        public override int ReadByte()
        {
            int value = _inner.ReadByte();
            if (value >= 0)
            {
                _read++;
                if (_read > _limit)
                {
                    throw new InvalidDataException("文档部件解压后超过安全上限: " + _name);
                }
                if (_report != null) _report(1);
            }
            return value;
        }

        /// <summary>无操作：解压流不缓冲可刷新的写入。 No-op: a decompression stream buffers nothing to flush.</summary>
        public override void Flush()
        {
        }

        /// <summary>不支持定位。 Seeking is not supported.</summary>
        /// <param name="offset">偏移。 Offset.</param>
        /// <param name="origin">基准。 Origin.</param>
        /// <returns>从不返回。 Never returns.</returns>
        public override long Seek(long offset, SeekOrigin origin)
        {
            throw new NotSupportedException();
        }

        /// <summary>不支持设置长度。 Setting the length is not supported.</summary>
        /// <param name="value">长度。 Length.</param>
        public override void SetLength(long value)
        {
            throw new NotSupportedException();
        }

        /// <summary>只读流不可写。 The stream is read-only.</summary>
        /// <param name="buffer">数据。 Data.</param>
        /// <param name="offset">起点。 Offset.</param>
        /// <param name="count">长度。 Length.</param>
        public override void Write(byte[] buffer, int offset, int count)
        {
            throw new NotSupportedException();
        }

        /// <summary>关闭时释放底层条目流。 Releases the underlying entry stream on close.</summary>
        /// <param name="disposing">是否释放托管资源。 Whether managed resources are being released.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
