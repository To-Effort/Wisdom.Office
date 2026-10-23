using System;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// Word 格式转换器：统一加载 DOC/DOCX/DOTX/RTF/HTML/ODT/纯文本，并把统一文档模型输出为
    /// DOCX/DOTX/RTF/HTML/ODT/纯文本。任意可读格式都可以转换为任意可写格式（含 30 种以上组合），
    /// 旧版二进制 .doc 仅支持读取。
    /// Word format converter: loads DOC/DOCX/DOTX/RTF/HTML/ODT/plain text into the unified document model and
    /// writes that model as DOCX/DOTX/RTF/HTML/ODT/plain text. Every readable format can be converted to every
    /// writable format (more than 30 combinations); the legacy binary .doc format is read-only.
    /// </summary>
    public static class WordFormatConverter
    {
        /// <summary>按格式从文件加载文档；格式为 Unknown 时先嗅探内容，再按扩展名兜底。Loads a document from a file in the given format; when the format is Unknown the content is sniffed first and the extension is used as a fallback.</summary>
        /// <param name="path">文件路径。The file path.</param>
        /// <param name="format">期望格式，Unknown 表示自动识别。The expected format; Unknown means auto-detection.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(string path, WordDocumentFormat format)
        {
            if (path == null) throw new ArgumentNullException("path");
            byte[] data = File.ReadAllBytes(path);
            if (format == WordDocumentFormat.Unknown)
            {
                format = WordDocumentFormats.Detect(data);
                if (format == WordDocumentFormat.Unknown)
                {
                    format = WordDocumentFormats.FromExtension(Path.GetExtension(path));
                }
            }
            return Load(data, format);
        }

        /// <summary>自动识别格式并从文件加载文档。Auto-detects the format and loads a document from a file.</summary>
        /// <param name="path">文件路径。The file path.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(string path)
        {
            return Load(path, WordDocumentFormat.Unknown);
        }

        /// <summary>按格式从流加载文档；流保持打开。Loads a document from a stream in the given format, leaving the stream open.</summary>
        /// <param name="stream">文档字节流。The document stream.</param>
        /// <param name="format">文档格式。The document format.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(Stream stream, WordDocumentFormat format)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return Load(OfficePackageLimits.ReadAllBytes(stream), format);
        }

        /// <summary>自动识别格式并从流加载文档；流保持打开。Auto-detects the format and loads a document from a stream, leaving the stream open.</summary>
        /// <param name="stream">文档字节流。The document stream.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(Stream stream)
        {
            return Load(stream, WordDocumentFormat.Unknown);
        }

        /// <summary>按格式从字节数组加载文档；格式为 Unknown 时自动嗅探。Loads a document from a byte array in the given format; Unknown triggers content sniffing.</summary>
        /// <param name="data">文档字节。The document bytes.</param>
        /// <param name="format">文档格式，Unknown 表示自动识别。The document format; Unknown means auto-detection.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        /// <exception cref="NotSupportedException">格式无法识别或不受支持时抛出。Thrown when the format is unrecognized or unsupported.</exception>
        public static OfficeDocument Load(byte[] data, WordDocumentFormat format)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (format == WordDocumentFormat.Unknown) format = WordDocumentFormats.Detect(data);
            switch (format)
            {
                case WordDocumentFormat.Docx:
                case WordDocumentFormat.Dotx:
                    return DocxReader.Load(data);
                case WordDocumentFormat.Doc:
                    return DocReader.Load(data);
                case WordDocumentFormat.Rtf:
                    return RtfReader.Load(data);
                case WordDocumentFormat.Html:
                    return HtmlReader.Load(data);
                case WordDocumentFormat.Odt:
                    return OdtReader.Load(data);
                case WordDocumentFormat.Text:
                    return OfficeDocument.FromText(DecodeText(data));
                default:
                    throw new NotSupportedException("无法识别 Word 文档格式。支持 DOC、DOCX、DOTX、RTF、HTML、ODT 和纯文本。");
            }
        }

        /// <summary>自动识别格式并从字节数组加载文档。Auto-detects the format and loads a document from a byte array.</summary>
        /// <param name="data">文档字节。The document bytes.</param>
        /// <returns>解析后的文档。The parsed document.</returns>
        public static OfficeDocument Load(byte[] data)
        {
            return Load(data, WordDocumentFormat.Unknown);
        }

        /// <summary>嗅探字节内容的 Word 格式。Sniffs the Word format of byte content.</summary>
        /// <param name="data">文件字节。The file bytes.</param>
        /// <returns>嗅探出的格式。The sniffed format.</returns>
        public static WordDocumentFormat DetectFormat(byte[] data)
        {
            return WordDocumentFormats.Detect(data);
        }

        /// <summary>识别文件路径对应的 Word 格式；内容无法判断时按扩展名兜底。Detects the Word format of a file path, falling back to the extension when the content cannot be recognized.</summary>
        /// <param name="path">文件路径。The file path.</param>
        /// <returns>识别出的格式。The detected format.</returns>
        public static WordDocumentFormat DetectFormat(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            using (FileStream stream = File.OpenRead(path))
            {
                byte[] header = new byte[8];
                int read = stream.Read(header, 0, header.Length);
                byte[] prefix = new byte[read];
                Array.Copy(header, prefix, read);
                WordDocumentFormat format = WordDocumentFormats.Detect(prefix);
                if (format != WordDocumentFormat.Unknown && format != WordDocumentFormat.Text) return format;
            }
            return WordDocumentFormats.FromExtension(Path.GetExtension(path));
        }

        /// <summary>按目标路径扩展名把源文件转换为另一种 Word 格式。Converts a source file to another Word format, chosen by the target path extension.</summary>
        /// <param name="sourcePath">源文件路径。The source file path.</param>
        /// <param name="targetPath">目标文件路径。The target file path.</param>
        public static void Convert(string sourcePath, string targetPath)
        {
            if (targetPath == null) throw new ArgumentNullException("targetPath");
            WordDocumentFormat target = WordDocumentFormats.FromExtension(Path.GetExtension(targetPath));
            if (target == WordDocumentFormat.Unknown)
            {
                throw new NotSupportedException("无法根据扩展名确定 Word 输出格式：" + targetPath);
            }
            Save(Load(sourcePath), targetPath, target);
        }

        /// <summary>把文档保存为指定格式的文件；格式为 Unknown 时按扩展名判断。Saves a document to a file in the given format; Unknown is resolved from the extension.</summary>
        /// <param name="document">待保存的文档。The document to save.</param>
        /// <param name="path">目标文件路径。The target file path.</param>
        /// <param name="format">目标格式。The target format.</param>
        public static void Save(OfficeDocument document, string path, WordDocumentFormat format)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (path == null) throw new ArgumentNullException("path");
            if (format == WordDocumentFormat.Unknown) format = WordDocumentFormats.FromExtension(Path.GetExtension(path));
            if (format == WordDocumentFormat.Unknown) throw new NotSupportedException("无法根据扩展名确定 Word 输出格式：" + path);
            if (format == WordDocumentFormat.Doc)
            {
                throw new NotSupportedException("旧版 .doc 二进制格式只支持读取，请输出为 .docx/.dotx/.rtf/.html/.odt。");
            }
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                Save(document, stream, format);
            }
        }

        /// <summary>把文档保存到流中；DOC 为只读格式，写入会抛出异常。Saves a document into a stream; DOC is read-only and writing it throws.</summary>
        /// <param name="document">待保存的文档。The document to save.</param>
        /// <param name="stream">目标流。The target stream.</param>
        /// <param name="format">目标格式。The target format.</param>
        public static void Save(OfficeDocument document, Stream stream, WordDocumentFormat format)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (stream == null) throw new ArgumentNullException("stream");
            switch (format)
            {
                case WordDocumentFormat.Docx:
                    DocxWriter.Write(document, stream);
                    break;
                case WordDocumentFormat.Dotx:
                    DocxWriter.WriteTemplate(document, stream);
                    break;
                case WordDocumentFormat.Rtf:
                    RtfWriter.Write(document, stream);
                    break;
                case WordDocumentFormat.Html:
                    HtmlWriter.Write(document, stream);
                    break;
                case WordDocumentFormat.Odt:
                    OdtWriter.Write(document, stream);
                    break;
                case WordDocumentFormat.Text:
                    {
                        byte[] bytes = new UTF8Encoding(false).GetBytes(document.GetPlainText());
                        stream.Write(bytes, 0, bytes.Length);
                        break;
                    }
                case WordDocumentFormat.Doc:
                    throw new NotSupportedException("旧版 .doc 二进制格式只支持读取，请输出为 .docx/.dotx/.rtf/.html/.odt。");
                default:
                    throw new NotSupportedException("不支持的 Word 输出格式：" + format);
            }
        }

        /// <summary>把文档序列化为指定格式的字节数组。Serialises a document into a byte array in the given format.</summary>
        /// <param name="document">待序列化的文档。The document to serialise.</param>
        /// <param name="format">目标格式。The target format.</param>
        /// <returns>目标格式的文件字节。The file bytes in the target format.</returns>
        public static byte[] ToArray(OfficeDocument document, WordDocumentFormat format)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Save(document, stream, format);
                return stream.ToArray();
            }
        }

        /// <summary>把文档字节从一种 Word 格式转换为另一种。Converts document bytes from one Word format to another.</summary>
        /// <param name="data">源文档字节。The source document bytes.</param>
        /// <param name="targetFormat">目标格式。The target format.</param>
        /// <returns>目标格式的文件字节。The file bytes in the target format.</returns>
        public static byte[] Convert(byte[] data, WordDocumentFormat targetFormat)
        {
            return ToArray(Load(data), targetFormat);
        }

        /// <summary>按 BOM 或 UTF-8 解码纯文本字节。Decodes plain-text bytes by BOM or as UTF-8.</summary>
        private static string DecodeText(byte[] data)
        {
            if (data == null || data.Length == 0) return string.Empty;
            if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF)
            {
                return new UTF8Encoding(false).GetString(data, 3, data.Length - 3);
            }
            if (data.Length >= 2 && data[0] == 0xFF && data[1] == 0xFE)
            {
                return Encoding.Unicode.GetString(data, 2, data.Length - 2);
            }
            if (data.Length >= 2 && data[0] == 0xFE && data[1] == 0xFF)
            {
                return Encoding.BigEndianUnicode.GetString(data, 2, data.Length - 2);
            }
            return new UTF8Encoding(false).GetString(data);
        }
    }
}
