using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 一条 OOXML 数字签名的元数据：算法、签名引用部件与证书主体，不包含密码学验证结果。
    /// Metadata of one OOXML digital signature: algorithms, signed part references and the certificate subject,
    /// without any cryptographic verification result.
    /// </summary>
    public sealed class WordDigitalSignature
    {
        /// <summary>签名部件路径（如 _xmlsignatures/sig1.xml）。The signature part path (for example _xmlsignatures/sig1.xml).</summary>
        public string PartPath { get; internal set; }
        /// <summary>签名算法 URI。The signature algorithm URI.</summary>
        public string SignatureMethod { get; internal set; }
        /// <summary>摘要算法 URI。The digest algorithm URI.</summary>
        public string DigestMethod { get; internal set; }
        /// <summary>规范化算法 URI。The canonicalization algorithm URI.</summary>
        public string CanonicalizationMethod { get; internal set; }
        /// <summary>被签名的部件引用 URI 列表。The list of signed-part reference URIs.</summary>
        public IList<string> SignedParts { get; internal set; }
        /// <summary>签名证书的主体名称；证书缺失或无法解析时为 null。The subject name of the signing certificate; null when the certificate is missing or cannot be parsed.</summary>
        public string SignerSubject { get; internal set; }
        /// <summary>是否存在签名值节点（表示签名结构完整，不代表验证通过）。Whether a signature-value node exists (the structure is complete; this does not mean verification succeeded).</summary>
        public bool HasSignatureValue { get; internal set; }
    }

    /// <summary>
    /// 数字签名检查报告：是否签名以及每条签名的元数据。
    /// Digital-signature inspection report: whether the package is signed and the metadata of each signature.
    /// </summary>
    public sealed class WordDigitalSignatureReport
    {
        /// <summary>包内是否存在签名源部件（_xmlsignatures/origin.sigs）。Whether the package contains a signature origin part (_xmlsignatures/origin.sigs).</summary>
        public bool IsSigned { get; internal set; }
        /// <summary>解析出的签名列表；结构损坏的签名部件被跳过。The parsed signatures; malformed signature parts are skipped.</summary>
        public IList<WordDigitalSignature> Signatures { get; internal set; }

        /// <summary>默认构造：未签名、空签名列表。Default constructor: unsigned with an empty signature list.</summary>
        public WordDigitalSignatureReport()
        {
            Signatures = new List<WordDigitalSignature>();
        }
    }

    /// <summary>
    /// 数字签名检查器：检测 DOCX 包中的 OOXML 签名部件并读取其元数据。
    /// Digital-signature inspector that detects the OOXML signature parts of a DOCX package and reads their
    /// metadata.
    /// </summary>
    /// <remarks>
    /// 本检查器不验证签名的密码学有效性（验证需要 XML 规范化与公钥运算），只报告算法、引用部件与证书主体；
    /// 任何重新生成包的写出操作都会使原签名失效，写出器不会复制签名部件。
    ///
    /// The inspector does not verify cryptographic validity (that requires XML canonicalization and public-key
    /// operations); it reports algorithms, signed parts and certificate subjects. Any regenerating write
    /// invalidates the original signature, and the writer never copies signature parts.
    /// </remarks>
    public static class WordDigitalSignatureReader
    {
        /// <summary>XML 签名命名空间（ds）。The XML signature namespace (ds).</summary>
        private static readonly XNamespace Ds = "http://www.w3.org/2000/09/xmldsig#";

        /// <summary>签名源部件路径。The signature origin part path.</summary>
        private const string SignatureOriginPart = "_xmlsignatures/origin.sigs";

        /// <summary>读取指定 .docx 文件的签名信息。Reads the signature information of the given .docx file.</summary>
        /// <param name="path">.docx 文件路径。The .docx file path.</param>
        /// <returns>签名检查报告；未签名时返回空报告。The signature report; an empty report when unsigned.</returns>
        public static WordDigitalSignatureReport Read(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Read(File.ReadAllBytes(path));
        }

        /// <summary>读取 .docx 字节内容的签名信息。Reads the signature information from .docx bytes.</summary>
        /// <param name="data">.docx 文件的完整字节内容。The full bytes of the .docx file.</param>
        /// <returns>签名检查报告；未签名时返回空报告。The signature report; an empty report when unsigned.</returns>
        /// <exception cref="InvalidDataException">数据不是有效的 ZIP 包时抛出。Thrown when the data is not a valid ZIP package.</exception>
        public static WordDigitalSignatureReport Read(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length < 4 || data[0] != 0x50 || data[1] != 0x4B)
            {
                throw new InvalidDataException("The file is not a valid .docx package.");
            }
            WordDigitalSignatureReport report = new WordDigitalSignatureReport();
            try
            {
                using (MemoryStream memory = new MemoryStream(data, false))
                using (ZipArchive archive = PackageArchive.OpenForRead(memory, false))
                {
                    DocxPackage package = new DocxPackage(archive);
                    foreach (ZipArchiveEntry entry in archive.Entries)
                    {
                        string partPath = entry.FullName.Replace('\\', '/');
                        if (string.Equals(partPath, SignatureOriginPart, StringComparison.OrdinalIgnoreCase))
                        {
                            report.IsSigned = true;
                            continue;
                        }
                        if (!IsSignaturePart(partPath)) continue;
                        WordDigitalSignature signature = ReadSignature(package, partPath);
                        if (signature != null) report.Signatures.Add(signature);
                    }
                }
            }
            catch (IOException ex)
            {
                throw new InvalidDataException("The .docx package is corrupt or truncated.", ex);
            }
            return report;
        }

        /// <summary>
        /// 判断部件是否为签名 XML 部件：位于 _xmlsignatures/ 下、扩展名为 .xml 且不是关系文件或源部件。
        /// Determines whether a part is a signature XML part: under _xmlsignatures/, ending in .xml, and not a
        /// relationship file or the origin part.
        /// </summary>
        /// <param name="partPath">部件路径。The part path.</param>
        /// <returns>是签名部件时返回 true。True when the part is a signature part.</returns>
        private static bool IsSignaturePart(string partPath)
        {
            if (!partPath.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase)) return false;
            if (!partPath.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return false;
            if (partPath.IndexOf("/_rels/", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            return !string.Equals(partPath, SignatureOriginPart, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// 解析单个签名部件：读取算法、引用部件、签名值是否存在以及证书主体。
        /// Parses one signature part: algorithms, referenced parts, presence of the signature value and the
        /// certificate subject.
        /// </summary>
        /// <param name="package">DOCX 包。The DOCX package.</param>
        /// <param name="partPath">签名部件路径。The signature part path.</param>
        /// <returns>签名元数据；部件不是有效签名时返回 null。The signature metadata, or null when the part is not a valid signature.</returns>
        private static WordDigitalSignature ReadSignature(DocxPackage package, string partPath)
        {
            XDocument xml = package.GetXml(partPath);
            if (xml == null || xml.Root == null) return null;
            XElement signature = xml.Root.Name == Ds + "Signature" ? xml.Root : xml.Root.Element(Ds + "Signature");
            if (signature == null) return null;

            WordDigitalSignature result = new WordDigitalSignature();
            result.PartPath = partPath;
            result.SignatureMethod = ReadAlgorithm(signature.Element(Ds + "SignedInfo"), "SignatureMethod");
            result.CanonicalizationMethod = ReadAlgorithm(signature.Element(Ds + "SignedInfo"), "CanonicalizationMethod");
            result.DigestMethod = ReadAlgorithm(signature.Element(Ds + "SignedInfo"), "DigestMethod");
            result.HasSignatureValue = signature.Element(Ds + "SignatureValue") != null;

            List<string> references = new List<string>();
            XElement signedInfo = signature.Element(Ds + "SignedInfo");
            if (signedInfo != null)
            {
                foreach (XElement reference in signedInfo.Elements(Ds + "Reference"))
                {
                    XAttribute uri = reference.Attribute("URI");
                    if (uri != null) references.Add(uri.Value);
                }
            }
            result.SignedParts = references;
            result.SignerSubject = ReadCertificateSubject(signature);
            return result;
        }

        /// <summary>
        /// 从 SignedInfo 下的指定算法元素读取 Algorithm 特性值。
        /// Reads the Algorithm attribute of the named algorithm element under SignedInfo.
        /// </summary>
        /// <param name="signedInfo">ds:SignedInfo 元素，可为 null。The ds:SignedInfo element, may be null.</param>
        /// <param name="localName">算法元素的本地名称。The local name of the algorithm element.</param>
        /// <returns>算法 URI，缺失时返回 null。The algorithm URI, or null when missing.</returns>
        private static string ReadAlgorithm(XElement signedInfo, string localName)
        {
            if (signedInfo == null) return null;
            XElement element = signedInfo.Element(Ds + localName);
            if (element == null) element = signedInfo.Descendants(Ds + localName).FirstOrDefault();
            if (element == null) return null;
            XAttribute algorithm = element.Attribute("Algorithm");
            return algorithm == null ? null : algorithm.Value;
        }

        /// <summary>
        /// 从 ds:X509Certificate 的 Base64 内容解析签名证书主体名称，解析失败时返回 null。
        /// Parses the subject name of the signing certificate from the Base64 content of ds:X509Certificate;
        /// returns null when parsing fails.
        /// </summary>
        /// <param name="signature">ds:Signature 元素。The ds:Signature element.</param>
        /// <returns>证书主体名称，或 null。The certificate subject name, or null.</returns>
        private static string ReadCertificateSubject(XElement signature)
        {
            XElement certificate = signature.Descendants(Ds + "X509Certificate").FirstOrDefault();
            if (certificate == null || string.IsNullOrEmpty(certificate.Value)) return null;
            try
            {
                byte[] der = Convert.FromBase64String(certificate.Value.Trim());
#pragma warning disable SYSLIB0057
                using (X509Certificate2 parsed = new X509Certificate2(der))
#pragma warning restore SYSLIB0057
                {
                    return parsed.Subject;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
