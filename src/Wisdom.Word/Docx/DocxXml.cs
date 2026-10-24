using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// WordprocessingML XML 辅助工具：集中定义 OOXML 命名空间，并提供元素/属性读取、数值与长度单位换算、
    /// 颜色解析以及 XML 反序列化/序列化等静态方法。
    /// Helper for WordprocessingML XML: centralises the OOXML namespaces and offers static methods for reading
    /// elements and attributes, converting measurements and colours, and loading or serialising XML.
    /// </summary>
    internal static class Wml
    {
        /// <summary>WordprocessingML 主命名空间（w:）。The main WordprocessingML namespace (w:).</summary>
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        /// <summary>关系命名空间（r:），用于图片等关系的引用 ID。The relationships namespace (r:) used for reference ids such as embedded images.</summary>
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        /// <summary>Word 绘图命名空间（wp:），承载 wp:extent 等图形尺寸信息。The word-processing drawing namespace (wp:) carrying graphic extents such as wp:extent.</summary>
        public static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        /// <summary>DrawingML 主命名空间（a:），承载 a:blip 图片引用。The DrawingML main namespace (a:) carrying image references such as a:blip.</summary>
        public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        /// <summary>图片命名空间（pic:）。The picture namespace (pic:).</summary>
        public static readonly XNamespace Pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        /// <summary>VML 命名空间（v:），用于旧式图片。The VML namespace (v:) used by legacy pictures.</summary>
        public static readonly XNamespace V = "urn:schemas-microsoft-com:vml";
        /// <summary>OPC 包关系命名空间，用于解析 *.rels 文件。The OPC package-relationships namespace used to parse *.rels files.</summary>
        public static readonly XNamespace PackageRels = "http://schemas.openxmlformats.org/package/2006/relationships";

        /// <summary>父元素缺失时返回的空子元素序列（共享实例，避免逐次分配）。Shared empty child sequence returned when a parent is missing, avoiding a per-call allocation.</summary>
        private static readonly XElement[] EmptyElements = new XElement[0];

        /// <summary>
        /// 取第一个 w: 命名空间下的指定子元素。
        /// Returns the first child element with the given local name in the w: namespace.
        /// </summary>
        /// <param name="parent">父元素，可为 null。The parent element, may be null.</param>
        /// <param name="localName">子元素的本地名称（不含前缀）。The local name of the child (without prefix).</param>
        /// <returns>找到的子元素，否则为 null。The child found, otherwise null.</returns>
        public static XElement Child(XElement parent, string localName)
        {
            return parent == null ? null : parent.Element(W + localName);
        }

        /// <summary>
        /// 取 w: 命名空间下指定本地名称的全部子元素；父元素为 null 时返回空序列。
        /// Returns all child elements with the given local name in the w: namespace; an empty sequence when the
        /// parent is null.
        /// </summary>
        /// <param name="parent">父元素，可为 null。The parent element, may be null.</param>
        /// <param name="localName">子元素的本地名称。The local name of the children.</param>
        /// <returns>匹配的子元素序列。The sequence of matching children.</returns>
        public static IEnumerable<XElement> Children(XElement parent, string localName)
        {
            if (parent == null) return EmptyElements;
            return parent.Elements(W + localName);
        }

        /// <summary>
        /// 读取 w: 命名空间下指定名称的特性值。
        /// Reads the value of an attribute in the w: namespace.
        /// </summary>
        /// <param name="element">目标元素，可为 null。The target element, may be null.</param>
        /// <param name="localName">特性本地名称。The local name of the attribute.</param>
        /// <returns>特性值，缺失时为 null。The attribute value, or null when absent.</returns>
        public static string AttributeValue(XElement element, string localName)
        {
            if (element == null) return null;
            XAttribute attribute = element.Attribute(W + localName);
            return attribute == null ? null : attribute.Value;
        }

        /// <summary>
        /// 按完整 <see cref="XName"/>（含命名空间）读取特性值。
        /// Reads an attribute value using a fully qualified <see cref="XName"/> including its namespace.
        /// </summary>
        /// <param name="element">目标元素，可为 null。The target element, may be null.</param>
        /// <param name="name">特性完整名称。The fully qualified attribute name.</param>
        /// <returns>特性值，缺失时为 null。The attribute value, or null when absent.</returns>
        public static string AttributeValue(XElement element, XName name)
        {
            if (element == null) return null;
            XAttribute attribute = element.Attribute(name);
            return attribute == null ? null : attribute.Value;
        }

        /// <summary>
        /// 读取不带命名空间前缀的特性原始值。
        /// Reads the raw value of an attribute whose name has no namespace prefix.
        /// </summary>
        /// <param name="element">目标元素，可为 null。The target element, may be null.</param>
        /// <param name="name">特性名称。The attribute name.</param>
        /// <returns>特性原始值，缺失时为 null。The raw attribute value, or null when absent.</returns>
        public static string RawAttributeValue(XElement element, string name)
        {
            if (element == null) return null;
            XAttribute attribute = element.Attribute(name);
            return attribute == null ? null : attribute.Value;
        }

        /// <summary>
        /// 读取 OOXML 的开关型元素（on/off）：元素缺失时返回回退值，特性缺失视为开启，
        /// val 为 0/false/off 时视为关闭。
        /// Reads an OOXML on/off element: the fallback is returned when the element is missing, the flag counts
        /// as on when the val attribute is absent, and 0/false/off means off.
        /// </summary>
        /// <param name="parent">包含开关元素的父元素。The parent element that contains the flag element.</param>
        /// <param name="localName">开关元素的本地名称。The local name of the flag element.</param>
        /// <param name="fallback">元素缺失时使用的回退值。The value used when the element is missing.</param>
        /// <returns>开关的最终状态。The resulting state of the flag.</returns>
        public static bool GetOnOff(XElement parent, string localName, bool fallback)
        {
            XElement element = Child(parent, localName);
            if (element == null) return fallback;
            string value = AttributeValue(element, "val");
            if (value == null) return true;
            return value != "0" && value != "false" && value != "off";
        }

        /// <summary>
        /// 读取子元素 w:val 特性中的整数值。
        /// Reads the integer stored in the w:val attribute of a child element.
        /// </summary>
        /// <param name="parent">父元素。The parent element.</param>
        /// <param name="localName">子元素本地名称。The local name of the child element.</param>
        /// <param name="fallback">解析失败时使用的回退值。The fallback used when parsing fails.</param>
        /// <returns>解析出的整数或回退值。The parsed integer or the fallback.</returns>
        public static int GetInt(XElement parent, string localName, int fallback)
        {
            XElement element = Child(parent, localName);
            return ParseInt(AttributeValue(element, "val"), fallback);
        }

        /// <summary>
        /// 读取子元素 w:val 特性中的浮点值。
        /// Reads the floating-point value stored in the w:val attribute of a child element.
        /// </summary>
        /// <param name="parent">父元素。The parent element.</param>
        /// <param name="localName">子元素本地名称。The local name of the child element.</param>
        /// <param name="fallback">解析失败时使用的回退值。The fallback used when parsing fails.</param>
        /// <returns>解析出的浮点值或回退值。The parsed value or the fallback.</returns>
        public static double GetDouble(XElement parent, string localName, double fallback)
        {
            XElement element = Child(parent, localName);
            return ParseDouble(AttributeValue(element, "val"), fallback);
        }

        /// <summary>
        /// 以不变文化区域解析整数。
        /// Parses an integer using the invariant culture.
        /// </summary>
        /// <param name="value">待解析的文本。The text to parse.</param>
        /// <param name="fallback">解析失败时使用的回退值。The fallback used when parsing fails.</param>
        /// <returns>解析结果或回退值。The parsed value or the fallback.</returns>
        public static int ParseInt(string value, int fallback)
        {
            int result;
            if (!string.IsNullOrEmpty(value) && int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            return fallback;
        }

        /// <summary>
        /// 以不变文化区域解析浮点数。
        /// Parses a floating-point number using the invariant culture.
        /// </summary>
        /// <param name="value">待解析的文本。The text to parse.</param>
        /// <param name="fallback">解析失败时使用的回退值。The fallback used when parsing fails.</param>
        /// <returns>解析结果或回退值。The parsed value or the fallback.</returns>
        public static double ParseDouble(string value, double fallback)
        {
            double result;
            if (!string.IsNullOrEmpty(value) && double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            return fallback;
        }

        /// <summary>
        /// 将缇（twip）文本换算为磅值的别名方法。
        /// Alias that converts twip text into points.
        /// </summary>
        /// <param name="value">以缇为单位的数值文本。The value expressed in twips.</param>
        /// <param name="fallback">解析失败时使用的回退值。The fallback used when parsing fails.</param>
        /// <returns>磅值或回退值。The value in points or the fallback.</returns>
        public static double Twips(string value, double fallback)
        {
            return TwipsToPoints(value, fallback);
        }

        /// <summary>
        /// 将缇（1/20 磅）换算为磅值。
        /// Converts twips (1/20 of a point) into points.
        /// </summary>
        /// <param name="value">以缇为单位的数值文本。The value expressed in twips.</param>
        /// <param name="fallback">解析失败时使用的回退值。The fallback used when parsing fails.</param>
        /// <returns>磅值或回退值。The value in points or the fallback.</returns>
        public static double TwipsToPoints(string value, double fallback)
        {
            double twips;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out twips))
            {
                return twips / UnitConverter.TwipsPerPoint;
            }
            return fallback;
        }

        /// <summary>
        /// 解析带单位的度量文本（pt、in、cm、mm、px），统一换算为磅值；无单位时按缇处理。
        /// Parses a measurement with a unit suffix (pt, in, cm, mm, px) and converts it to points; a value
        /// without a unit is treated as twips.
        /// </summary>
        /// <param name="value">度量文本。The measurement text.</param>
        /// <param name="fallback">无法解析时使用的回退值。The fallback used when the text cannot be parsed.</param>
        /// <returns>磅值或回退值。The value in points or the fallback.</returns>
        public static double MeasureToPoints(string value, double fallback)
        {
            if (string.IsNullOrEmpty(value)) return fallback;
            string text = value.Trim();
            if (text.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                return ParseDouble(text.Substring(0, text.Length - 2), fallback);
            }
            if (text.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            {
                return ParseDouble(text.Substring(0, text.Length - 2), fallback / UnitConverter.PointsPerInch) * UnitConverter.PointsPerInch;
            }
            if (text.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
            {
                return ParseDouble(text.Substring(0, text.Length - 2), fallback / 28.3464567) * 28.3464567;
            }
            if (text.EndsWith("mm", StringComparison.OrdinalIgnoreCase))
            {
                return ParseDouble(text.Substring(0, text.Length - 2), fallback / 2.83464567) * 2.83464567;
            }
            if (text.EndsWith("px", StringComparison.OrdinalIgnoreCase))
            {
                return UnitConverter.PixelsToPoints(ParseDouble(text.Substring(0, text.Length - 2), fallback));
            }
            return TwipsToPoints(text, fallback);
        }

        /// <summary>
        /// 解析 OOXML 颜色值（十六进制或 auto）；auto 与非法值返回回退颜色。
        /// Parses an OOXML colour (hexadecimal or auto); auto and invalid values yield the fallback colour.
        /// </summary>
        /// <param name="value">颜色文本。The colour text.</param>
        /// <param name="fallback">回退颜色。The fallback colour.</param>
        /// <returns>解析出的颜色或回退值。The parsed colour or the fallback.</returns>
        public static OfficeColor ParseColor(string value, OfficeColor fallback)
        {
            if (string.IsNullOrEmpty(value) || value == "auto") return fallback;
            OfficeColor color;
            if (OfficeColor.TryParse(value, out color)) return color;
            return fallback;
        }

        /// <summary>
        /// 将颜色序列化为写入 OOXML 使用的十六进制文本。
        /// Serialises a colour into the hexadecimal text written into OOXML.
        /// </summary>
        /// <param name="color">目标颜色。The colour to convert.</param>
        /// <returns>十六进制颜色文本。The hexadecimal colour text.</returns>
        public static string ColorToHex(OfficeColor color)
        {
            return color.ToHex();
        }

        /// <summary>
        /// 以安全设置（禁止 DTD、禁用外部实体解析）加载 XML，并保留空白字符；数据为空时返回 null。
        /// Loads XML with safe settings (DTD processing prohibited, external resolvers disabled) while
        /// preserving whitespace; returns null for empty data.
        /// </summary>
        /// <param name="data">XML 字节内容。The XML bytes.</param>
        /// <returns>解析出的文档，数据为空时为 null。The parsed document, or null when the data is empty.</returns>
        /// <exception cref="XmlException">XML 格式非法。Thrown when the XML is malformed.</exception>
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
        /// 将文档序列化为不带 BOM 的 UTF-8 字节，不缩进；原声明缺失时省略 XML 声明。
        /// Serialises a document into UTF-8 bytes without a BOM and without indentation; the XML declaration is
        /// omitted when the document has none.
        /// </summary>
        /// <param name="document">待序列化的文档。The document to serialise.</param>
        /// <returns>序列化后的字节内容。The serialised bytes.</returns>
        public static byte[] SerializeXml(XDocument document)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = new UTF8Encoding(false);
            settings.Indent = false;
            settings.OmitXmlDeclaration = document.Declaration == null;
            // 正文里一个 0x01 就能让 XmlWriter 抛异常，整份文档写不出来；先丢弃 XML 不接受的字符。
            // A single 0x01 in the body makes XmlWriter throw and the whole document unwritable, so the
            // characters XML forbids are dropped first.
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

        /// <summary>
        /// 将 EMU（English Metric Unit）文本换算为磅值。
        /// Converts an EMU (English Metric Unit) value expressed as text into points.
        /// </summary>
        /// <param name="value">以 EMU 为单位的数值文本。The value expressed in EMUs.</param>
        /// <param name="fallback">解析失败时使用的回退值。The fallback used when parsing fails.</param>
        /// <returns>磅值或回退值。The value in points or the fallback.</returns>
        public static double EmuToPoints(string value, double fallback)
        {
            double emu;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out emu))
            {
                return UnitConverter.EmuToPoints(emu);
            }
            return fallback;
        }

        /// <summary>
        /// 按不变文化区域把数值格式化为最多四位小数的文本，用于写入 OOXML 属性。
        /// Formats a number with the invariant culture using up to four decimals, as used for OOXML attributes.
        /// </summary>
        /// <param name="value">待格式化的数值。The value to format.</param>
        /// <returns>格式化后的文本。The formatted text.</returns>
        public static string ToInvariant(double value)
        {
            return value.ToString("0.####", CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// DOCX 包（OPC ZIP）访问器：按部件路径读取字节与 XML，并解析、缓存 OOXML 关系（*.rels），
    /// 提供关系目标查询与图片加载能力。
    /// Accessor for a DOCX (OPC ZIP) package: reads part bytes and XML by part path, parses and caches the OOXML
    /// relationships (*.rels), and resolves relationship targets and loads images.
    /// </summary>
    internal sealed class DocxPackage
    {
        /// <summary>底层 ZIP 压缩包，由调用方创建并负责释放。The underlying ZIP archive, created and owned by the caller.</summary>
        private readonly ZipArchive _archive;
        /// <summary>部件路径到关系查找表的缓存（关系只读，可直接复用同一个实例）。Cache mapping a part path to its relationship lookup; relationships are read-only so one instance is reused.</summary>
        private readonly Dictionary<string, Relationships> _relationships = new Dictionary<string, Relationships>(StringComparer.OrdinalIgnoreCase);
        /// <summary>已累计读取的字节数，用于配合 <see cref="OfficePackageLimits"/> 限制解压体积。Bytes read so far, used with <see cref="OfficePackageLimits"/> to cap the decompressed size.</summary>
        private long _totalBytes;

        /// <summary>
        /// 基于已打开的 ZIP 包创建访问器。
        /// Creates an accessor around an opened ZIP archive.
        /// </summary>
        /// <param name="archive">已打开的 DOCX ZIP 包。The opened DOCX ZIP archive.</param>
        public DocxPackage(ZipArchive archive)
        {
            _archive = archive;
        }

        /// <summary>
        /// 读取并解析指定部件的 XML；部件缺失或 XML 非法时返回 null。
        /// Reads and parses the XML of a part; returns null when the part is missing or the XML is invalid.
        /// </summary>
        /// <param name="partPath">部件路径（如 word/document.xml）。The part path (e.g. word/document.xml).</param>
        /// <returns>解析出的文档，失败时为 null。The parsed document, or null on failure.</returns>
        public XDocument GetXml(string partPath)
        {
            byte[] data = GetBytes(partPath);
            if (data == null) return null;
            try
            {
                return Wml.LoadXml(data);
            }
            catch (XmlException)
            {
                return null;
            }
        }

        /// <summary>
        /// 读取部件的原始字节：先按精确路径查找条目，找不到时对条目名做反斜杠归一化与忽略大小写的匹配。
        /// Reads the raw bytes of a part: looks the entry up by exact path first and, when missing, matches
        /// entry names after backslash normalisation using a case-insensitive comparison.
        /// </summary>
        /// <param name="partPath">部件路径，为空时返回 null。The part path; null is returned when it is empty.</param>
        /// <returns>部件字节内容，缺失时返回 null。The part bytes, or null when the part does not exist.</returns>
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

        /// <summary>
        /// 判断指定部件是否存在于包中。
        /// Determines whether the given part exists in the package.
        /// </summary>
        /// <param name="partPath">部件路径。The part path.</param>
        /// <returns>存在时返回 true。True when the part is present.</returns>
        public bool Exists(string partPath)
        {
            return GetBytes(partPath) != null;
        }

        /// <summary>
        /// 按关系 ID 查询目标并解析为相对当前部件的绝对部件路径。
        /// Looks up a relationship target by id and resolves it into a part path relative to the current part.
        /// </summary>
        /// <param name="partPath">发起关系的部件路径。The part path that owns the relationship.</param>
        /// <param name="relationshipId">关系 ID（r:id 的值）。The relationship id (the value of r:id).</param>
        /// <returns>解析后的部件路径，未找到时返回 null。The resolved part path, or null when not found.</returns>
        public string GetRelationshipTarget(string partPath, string relationshipId)
        {
            if (string.IsNullOrEmpty(relationshipId)) return null;
            Relationships relationships = GetRelationships(partPath);
            string target;
            if (!relationships.Targets.TryGetValue(relationshipId, out target)) return null;
            return ResolvePartPath(partPath, target);
        }

        /// <summary>
        /// 按关系 ID 获取完整关系信息（类型与内部/外部标记）。
        /// Gets the full relationship information (type and internal/external flag) by relationship id.
        /// </summary>
        /// <param name="partPath">发起关系的部件路径。The part path that owns the relationship.</param>
        /// <param name="relationshipId">关系 ID。The relationship id.</param>
        /// <returns>关系信息，未找到时返回 null。The relationship, or null when not found.</returns>
        public RelationshipInfo GetRelationship(string partPath, string relationshipId)
        {
            Relationships relationships = GetRelationships(partPath);
            RelationshipInfo info;
            if (relationships.Info.TryGetValue(relationshipId, out info)) return info;
            return null;
        }

        /// <summary>
        /// 查找关系类型以指定后缀结尾的首个内部关系，并返回其解析后的部件路径。
        /// Finds the first internal relationship whose type ends with the given suffix and returns its resolved
        /// part path.
        /// </summary>
        /// <param name="partPath">发起关系的部件路径。The part path that owns the relationship.</param>
        /// <param name="typeSuffix">关系类型后缀（如 /styles）。The relationship type suffix (e.g. /styles).</param>
        /// <returns>解析后的部件路径，未找到时返回 null。The resolved part path, or null when not found.</returns>
        public string GetRelationshipTargetByType(string partPath, string typeSuffix)
        {
            Relationships relationships = GetRelationships(partPath);
            foreach (KeyValuePair<string, RelationshipInfo> pair in relationships.Info)
            {
                RelationshipInfo info = pair.Value;
                if (!info.External && info.Type != null
                    && info.Type.EndsWith(typeSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    return ResolvePartPath(partPath, info.Target);
                }
            }
            return null;
        }

        /// <summary>
        /// 列出关系类型以指定后缀结尾的全部内部关系。
        /// Lists every internal relationship whose type ends with the given suffix.
        /// </summary>
        /// <param name="partPath">发起关系的部件路径。The part path that owns the relationships.</param>
        /// <param name="typeSuffix">关系类型后缀。The relationship type suffix.</param>
        /// <returns>匹配的关系信息列表。The list of matching relationships.</returns>
        public List<RelationshipInfo> GetRelationshipsByType(string partPath, string typeSuffix)
        {
            List<RelationshipInfo> result = new List<RelationshipInfo>();
            Relationships relationships = GetRelationships(partPath);
            foreach (KeyValuePair<string, RelationshipInfo> pair in relationships.Info)
            {
                RelationshipInfo info = pair.Value;
                if (!info.External && info.Type != null
                    && info.Type.EndsWith(typeSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    result.Add(info);
                }
            }
            return result;
        }

        /// <summary>
        /// 按关系 ID 加载图片：解析目标部件并构造 <see cref="OfficeImage"/>；
        /// 部件缺失、读取失败或格式不是 PNG/JPEG 时返回 null。
        /// Loads an image by relationship id: resolves the target part and builds an <see cref="OfficeImage"/>;
        /// returns null when the part is missing, cannot be read, or is neither PNG nor JPEG.
        /// </summary>
        /// <param name="partPath">引用图片的部件路径。The part path that references the image.</param>
        /// <param name="relationshipId">图片关系的 ID。The id of the image relationship.</param>
        /// <returns>加载的图片，失败时返回 null。The loaded image, or null on failure.</returns>
        public OfficeImage LoadImage(string partPath, string relationshipId)
        {
            string target = GetRelationshipTarget(partPath, relationshipId);
            if (target == null) return null;
            byte[] data = GetBytes(target);
            if (data == null || data.Length == 0) return null;
            OfficeImage image;
            try
            {
                image = new OfficeImage(data);
            }
            catch (Exception)
            {
                return null;
            }
            return image.Format == OfficeImageFormat.Png || image.Format == OfficeImageFormat.Jpeg ? image : null;
        }

        /// <summary>
        /// 把关系目标解析为包内规范路径：以 / 开头的目标视为根路径，否则相对源部件所在目录合并，
        /// 并消除 . 与 .. 段和多余分隔符。
        /// Resolves a relationship target into a canonical part path: a target starting with / is rooted at the
        /// package, otherwise it is joined with the directory of the source part, and dot segments plus redundant
        /// separators are removed.
        /// </summary>
        /// <param name="sourcePartPath">源部件路径（相对目标的基准）。The source part path used as the base for relative targets.</param>
        /// <param name="target">关系目标文本。The relationship target text.</param>
        /// <returns>规范化后的部件路径，目标为空时返回 null。The normalised part path, or null when the target is empty.</returns>
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

        /// <summary>
        /// 读取并缓存部件对应的 _rels/*.rels 关系文件，返回按关系 ID 索引的查找表；
        /// 关系表构造后只被读取，因此缓存命中时直接复用同一实例，避免每次重建字典。
        /// Reads and caches the _rels/*.rels file of a part and returns the relationships indexed by id. The table
        /// is only read after construction, so a cache hit reuses the same instance instead of rebuilding the
        /// dictionaries on every call.
        /// </summary>
        /// <param name="partPath">部件路径，可为 null（视为根部件）。The part path, may be null (treated as the root part).</param>
        /// <returns>该部件的关系集合，缺失关系文件时为空集合。The relationships of the part, empty when no rels file exists.</returns>
        private Relationships GetRelationships(string partPath)
        {
            string key = partPath ?? string.Empty;
            Relationships cached;
            if (_relationships.TryGetValue(key, out cached)) return cached;

            string relsPath;
            string baseDirectory = string.Empty;
            string fileName = key;
            int slash = fileName.LastIndexOf('/');
            if (slash >= 0)
            {
                baseDirectory = fileName.Substring(0, slash + 1);
                fileName = fileName.Substring(slash + 1);
            }
            relsPath = baseDirectory + "_rels/" + fileName + ".rels";
            XDocument rels = GetXml(relsPath);
            Relationships relationships = new Relationships();
            if (rels != null && rels.Root != null)
            {
                foreach (XElement relationship in rels.Root.Elements(Wml.PackageRels + "Relationship"))
                {
                    string id = (string)relationship.Attribute("Id");
                    string target = (string)relationship.Attribute("Target");
                    string type = (string)relationship.Attribute("Type");
                    string mode = (string)relationship.Attribute("TargetMode");
                    if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target)) continue;
                    relationships.Targets[id] = target;
                    relationships.Info[id] = new RelationshipInfo
                    {
                        Id = id,
                        Target = target,
                        Type = type ?? string.Empty,
                        External = mode == "External"
                    };
                }
            }
            _relationships[key] = relationships;
            return relationships;
        }
    }

    /// <summary>
    /// 单条 OOXML 关系的描述（Id、Target、Type 与内部/外部标记）。
    /// Description of a single OOXML relationship (id, target, type and the internal/external flag).
    /// </summary>
    internal sealed class RelationshipInfo
    {
        /// <summary>关系 ID（rId 等）。The relationship id (such as rId1).</summary>
        public string Id;
        /// <summary>关系目标（部件路径或外部 URI）。The relationship target (a part path or an external URI).</summary>
        public string Target;
        /// <summary>关系类型完整 URI。The full URI of the relationship type.</summary>
        public string Type;
        /// <summary>是否为外部关系（TargetMode="External"）。Whether the relationship is external (TargetMode="External").</summary>
        public bool External;
    }

    /// <summary>
    /// 某个部件的关系集合：ID 到目标路径的映射以及 ID 到完整关系信息的映射。
    /// The relationship set of a part: an id-to-target map and an id-to-full-relationship map.
    /// </summary>
    internal sealed class Relationships
    {
        /// <summary>关系 ID 到原始目标文本的映射。Map from relationship id to the raw target text.</summary>
        public readonly Dictionary<string, string> Targets = new Dictionary<string, string>(StringComparer.Ordinal);
        /// <summary>关系 ID 到关系详细信息对象的映射。Map from relationship id to the detailed relationship object.</summary>
        public readonly Dictionary<string, RelationshipInfo> Info = new Dictionary<string, RelationshipInfo>(StringComparer.Ordinal);
    }
}
