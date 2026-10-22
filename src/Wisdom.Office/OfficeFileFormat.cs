using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using Wisdom.Office.Core;

namespace Wisdom.Office
{
    /// <summary>本库可识别与支持的 Office / PDF 文件格式。Office / PDF file formats that this library can detect and support.</summary>
    public enum OfficeFileFormat
    {
        /// <summary>未知或不受支持的格式。Unknown or unsupported format.</summary>
        Unknown = 0,
        /// <summary>Adobe PDF 文档（.pdf）。Adobe PDF document (.pdf).</summary>
        Pdf = 1,
        /// <summary>基于 OOXML 的 Word 文档（.docx）。OOXML-based Word document (.docx).</summary>
        Docx = 2,
        /// <summary>旧版 Word 97-2003 二进制文档（.doc），仅支持读取。Legacy Word 97-2003 binary document (.doc), read-only.</summary>
        Doc = 3,
        /// <summary>基于 OOXML 的 Excel 工作簿（.xlsx）。OOXML-based Excel workbook (.xlsx).</summary>
        Xlsx = 4,
        /// <summary>旧版 Excel 97-2003 二进制工作簿（.xls），仅支持读取。Legacy Excel 97-2003 binary workbook (.xls), read-only.</summary>
        Xls = 5,
        /// <summary>基于 OOXML 的 PowerPoint 演示文稿（.pptx）。OOXML-based PowerPoint presentation (.pptx).</summary>
        Pptx = 6,
        /// <summary>Markdown 纯文本文档（.md），按 CommonMark 与 GFM 扩展解析。Markdown plain-text document (.md), parsed with CommonMark plus GFM extensions.</summary>
        Markdown = 7,
        /// <summary>基于 OOXML 的 Word 模板（.dotx）。OOXML-based Word template (.dotx).</summary>
        Dotx = 8,
        /// <summary>RTF 富文本文档（.rtf）。Rich Text Format document (.rtf).</summary>
        Rtf = 9,
        /// <summary>HTML 网页文档（.html/.htm）。HTML document (.html/.htm).</summary>
        Html = 10,
        /// <summary>OpenDocument 文本文档（.odt）。OpenDocument Text document (.odt).</summary>
        Odt = 11,
        /// <summary>纯文本文档（.txt），仅保留文本与段落。Plain-text document (.txt), preserving text and paragraphs only.</summary>
        Text = 12,
        /// <summary>基于 OOXML 的 PowerPoint 模板（.potx）。OOXML-based PowerPoint template (.potx).</summary>
        Potx = 13
    }

    /// <summary>文件格式辅助工具：按内容或扩展名识别格式，并在格式与扩展名、显示名称之间转换。File format helper: detects formats by content or extension and converts between format, extension, and display name.</summary>
    public static class OfficeFileFormats
    {
        /// <summary>OLE 复合文档（.doc/.xls 的容器格式）的 8 字节文件头签名。The 8-byte file header signature of an OLE compound document (the container format for .doc/.xls).</summary>
        private static readonly byte[] OleSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        /// <summary>打开磁盘文件并读取文件头以识别其格式。Opens a file on disk and reads its header to detect the format.</summary>
        /// <param name="path">文件路径。File path.</param>
        /// <returns>识别出的格式；无法识别时返回 <see cref="OfficeFileFormat.Unknown"/>。The detected format; <see cref="OfficeFileFormat.Unknown"/> when it cannot be recognized.</returns>
        public static OfficeFileFormat Detect(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                return Detect(stream);
            }
        }

        /// <summary>通过检查流开头的文件头（magic bytes）识别格式，检测结束后会恢复流的原始位置。Detects the format by inspecting the file header (magic bytes) at the start of the stream; restores the stream's original position when done.</summary>
        /// <param name="stream">支持定位（CanSeek）的读取流。A readable stream that supports seeking (CanSeek).</param>
        /// <returns>识别出的格式；无法识别时返回 <see cref="OfficeFileFormat.Unknown"/>。The detected format; <see cref="OfficeFileFormat.Unknown"/> when it cannot be recognized.</returns>
        /// <remarks>识别规则：%PDF- 前缀判定为 PDF；ZIP 容器再按内部 XML 部件区分 DOCX 与 XLSX；
        /// OLE 复合文档按内部流名区分 DOC 与 XLS。
        /// Detection rules: a %PDF- prefix is treated as PDF; ZIP containers are further split into DOCX or XLSX by their internal XML parts; OLE compound documents are split into DOC or XLS by their internal stream names.</remarks>
        public static OfficeFileFormat Detect(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            if (!stream.CanSeek) throw new ArgumentException("The stream must support seeking.", "stream");

            long origin = stream.Position;
            try
            {
                byte[] header = new byte[8];
                int read = stream.Read(header, 0, header.Length);
                if (read >= 5 && header[0] == (byte)'%' && header[1] == (byte)'P'
                    && header[2] == (byte)'D' && header[3] == (byte)'F' && header[4] == (byte)'-')
                {
                    return OfficeFileFormat.Pdf;
                }

                bool zip = read >= 4 && header[0] == (byte)'P' && header[1] == (byte)'K'
                    && (header[2] == 3 || header[2] == 5) && (header[3] == 4 || header[3] == 6);
                if (zip)
                {
                    stream.Position = origin;
                    return DetectZip(stream);
                }

                bool ole = read >= 8;
                for (int i = 0; ole && i < OleSignature.Length; i++)
                {
                    if (header[i] != OleSignature[i]) ole = false;
                }
                if (ole)
                {
                    stream.Position = origin;
                    return DetectCompound(stream);
                }

                // RTF 以 "{\rtf" 开头；HTML 在开头一段文本内出现 <html 或 <!doctype html。
                // RTF starts with "{\rtf"; HTML contains <html or <!doctype html within the leading text.
                if (read >= 5 && header[0] == (byte)'{' && header[1] == (byte)'\\'
                    && (header[2] == (byte)'r' || header[2] == (byte)'R')
                    && (header[3] == (byte)'t' || header[3] == (byte)'T')
                    && (header[4] == (byte)'f' || header[4] == (byte)'F'))
                {
                    return OfficeFileFormat.Rtf;
                }

                stream.Position = origin;
                byte[] prefix = new byte[512];
                int prefixRead = stream.Read(prefix, 0, prefix.Length);
                string leading = prefixRead <= 0 ? string.Empty : System.Text.Encoding.ASCII.GetString(prefix, 0, prefixRead);
                if (leading.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0
                    || leading.IndexOf("<!doctype html", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return OfficeFileFormat.Html;
                }
                return OfficeFileFormat.Unknown;
            }
            finally
            {
                stream.Position = origin;
            }
        }

        /// <summary>识别字节数组开头部分所表示的文件格式。Detects the file format represented by the beginning of a byte array.</summary>
        /// <param name="data">文件字节内容。The byte content of the file.</param>
        /// <returns>识别出的格式；无法识别时返回 <see cref="OfficeFileFormat.Unknown"/>。The detected format; <see cref="OfficeFileFormat.Unknown"/> when it cannot be recognized.</returns>
        public static OfficeFileFormat Detect(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            using (MemoryStream stream = new MemoryStream(data, false))
            {
                return Detect(stream);
            }
        }

        /// <summary>在 ZIP 容器内定位 OOXML 部件（word/document.xml、xl/workbook.xml 等）以区分 DOCX 与 XLSX。Locates OOXML parts (word/document.xml, xl/workbook.xml, etc.) inside the ZIP container to distinguish DOCX from XLSX.</summary>
        /// <param name="stream">已定位到开头的 ZIP 流。A ZIP stream positioned at its start.</param>
        /// <returns>DOCX、XLSX 或 <see cref="OfficeFileFormat.Unknown"/>。DOCX, XLSX, or <see cref="OfficeFileFormat.Unknown"/>.</returns>
        private static OfficeFileFormat DetectZip(Stream stream)
        {
            try
            {
                using (ZipArchive archive = PackageArchive.OpenForRead(stream, true))
                {
                    ZipArchiveEntry mimetype = FindEntry(archive, "mimetype");
                    if (mimetype != null)
                    {
                        string mime = ReadEntryPrefix(mimetype, 256);
                        if (mime.IndexOf("opendocument.text", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return OfficeFileFormat.Odt;
                        }
                    }
                    if (FindEntry(archive, "content.xml") != null && FindEntry(archive, "META-INF/manifest.xml") != null)
                    {
                        return OfficeFileFormat.Odt;
                    }

                    if (FindEntry(archive, "word/document.xml") != null)
                    {
                        ZipArchiveEntry templateTypes = FindEntry(archive, "[Content_Types].xml");
                        if (templateTypes != null)
                        {
                            string templateText = ReadEntryPrefix(templateTypes, 65536);
                            if (templateText.IndexOf("wordprocessingml.template.main+xml", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return OfficeFileFormat.Dotx;
                            }
                        }
                        return OfficeFileFormat.Docx;
                    }
                    if (FindEntry(archive, "xl/workbook.xml") != null) return OfficeFileFormat.Xlsx;
                    if (FindEntry(archive, "ppt/presentation.xml") != null)
                    {
                        // pptx 与 potx 的关系类型同为 officeDocument，只有 [Content_Types].xml 的
                        // presentation Override 能区分模板。
                        // pptx and potx share the officeDocument relationship type; only the presentation
                        // Override in [Content_Types].xml distinguishes a template.
                        ZipArchiveEntry presentationTypes = FindEntry(archive, "[Content_Types].xml");
                        if (presentationTypes != null)
                        {
                            string presentationText = ReadEntryPrefix(presentationTypes, 65536);
                            if (presentationText.IndexOf("presentationml.template.main+xml", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return OfficeFileFormat.Potx;
                            }
                        }
                        return OfficeFileFormat.Pptx;
                    }

                    ZipArchiveEntry types = FindEntry(archive, "[Content_Types].xml");
                    if (types != null)
                    {
                        string content = ReadEntryPrefix(types, 65536);
                        if (content.IndexOf("wordprocessingml.template", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return OfficeFileFormat.Dotx;
                        }
                        if (content.IndexOf("wordprocessingml.document", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return OfficeFileFormat.Docx;
                        }
                        if (content.IndexOf("spreadsheetml.sheet", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return OfficeFileFormat.Xlsx;
                        }
                        if (content.IndexOf("presentationml.presentation", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return OfficeFileFormat.Pptx;
                        }
                    }
                }
            }
            catch (InvalidDataException)
            {
            }
            catch (IOException)
            {
            }
            return OfficeFileFormat.Unknown;
        }

        /// <summary>解析 OLE 复合文档，按内部流名（WordDocument / Workbook / Book）区分 DOC 与 XLS。Parses an OLE compound document and distinguishes DOC from XLS by the internal stream name (WordDocument / Workbook / Book).</summary>
        /// <param name="stream">已定位到开头的复合文档流。A compound-document stream positioned at its start.</param>
        /// <returns>DOC、XLS 或 <see cref="OfficeFileFormat.Unknown"/>。DOC, XLS, or <see cref="OfficeFileFormat.Unknown"/>.</returns>
        private static OfficeFileFormat DetectCompound(Stream stream)
        {
            try
            {
                byte[] data = ReadAll(stream);
                CompoundFile compound = new CompoundFile(data);
                if (compound.HasStream("WordDocument")) return OfficeFileFormat.Doc;
                if (compound.HasStream("Workbook") || compound.HasStream("Book")) return OfficeFileFormat.Xls;
            }
            catch (InvalidDataException)
            {
            }
            catch (IOException)
            {
            }
            return OfficeFileFormat.Unknown;
        }

        /// <summary>按名称在 ZIP 归档中查找条目（忽略大小写与路径分隔符差异）。Finds an entry by name in the ZIP archive (ignoring case and path-separator differences).</summary>
        /// <param name="archive">ZIP 归档。The ZIP archive.</param>
        /// <param name="name">条目的相对路径名。The entry's relative path name.</param>
        /// <returns>匹配的条目；未找到时返回 null。The matching entry; null when not found.</returns>
        private static ZipArchiveEntry FindEntry(ZipArchive archive, string name)
        {
            for (int i = 0; i < archive.Entries.Count; i++)
            {
                if (string.Equals(archive.Entries[i].FullName.Replace('\\', '/'), name, StringComparison.OrdinalIgnoreCase))
                {
                    return archive.Entries[i];
                }
            }
            return null;
        }

        /// <summary>读取 ZIP 条目的文本前缀，最多读取指定字符数以避免解压炸弹。Reads the text prefix of a ZIP entry, capped at the given character count to avoid decompression bombs.</summary>
        /// <param name="entry">ZIP 条目。The ZIP entry.</param>
        /// <param name="maxChars">最大字符数。The maximum number of characters.</param>
        /// <returns>读取到的文本。The text that was read.</returns>
        private static string ReadEntryPrefix(ZipArchiveEntry entry, int maxChars)
        {
            if (entry == null || maxChars <= 0) return string.Empty;
            try
            {
                using (Stream stream = entry.Open())
                using (StreamReader reader = new StreamReader(stream, System.Text.Encoding.UTF8, true))
                {
                    char[] buffer = new char[Math.Min(maxChars, 4096)];
                    System.Text.StringBuilder builder = new System.Text.StringBuilder();
                    while (builder.Length < maxChars)
                    {
                        int read = reader.Read(buffer, 0, Math.Min(buffer.Length, maxChars - builder.Length));
                        if (read <= 0) break;
                        builder.Append(buffer, 0, read);
                    }
                    return builder.ToString();
                }
            }
            catch (InvalidDataException)
            {
                return string.Empty;
            }
            catch (IOException)
            {
                return string.Empty;
            }
        }

        /// <summary>根据文件扩展名映射到对应格式。Maps a file extension to its corresponding format.</summary>
        /// <param name="extension">扩展名（如 ".docx"），大小写不敏感；为空时返回 Unknown。The extension (e.g. ".docx"), case-insensitive; returns Unknown when empty.</param>
        /// <returns>对应的 <see cref="OfficeFileFormat"/>；扩展名不认识时返回 <see cref="OfficeFileFormat.Unknown"/>。The corresponding <see cref="OfficeFileFormat"/>; <see cref="OfficeFileFormat.Unknown"/> when the extension is unrecognized.</returns>
        public static OfficeFileFormat FromExtension(string extension)
        {
            if (string.IsNullOrEmpty(extension)) return OfficeFileFormat.Unknown;
            switch (extension.ToLowerInvariant())
            {
                case ".pdf":
                    return OfficeFileFormat.Pdf;
                case ".docx":
                case ".docm":
                    return OfficeFileFormat.Docx;
                case ".dotx":
                case ".dotm":
                    return OfficeFileFormat.Dotx;
                case ".doc":
                case ".dot":
                    return OfficeFileFormat.Doc;
                case ".xlsx":
                    return OfficeFileFormat.Xlsx;
                case ".xls":
                    return OfficeFileFormat.Xls;
                case ".pptx":
                    return OfficeFileFormat.Pptx;
                case ".potx":
                case ".potm":
                    return OfficeFileFormat.Potx;
                case ".rtf":
                    return OfficeFileFormat.Rtf;
                case ".html":
                case ".htm":
                case ".xhtml":
                    return OfficeFileFormat.Html;
                case ".odt":
                    return OfficeFileFormat.Odt;
                case ".txt":
                case ".text":
                    return OfficeFileFormat.Text;
                case ".md":
                case ".markdown":
                case ".mkd":
                    return OfficeFileFormat.Markdown;
                default:
                    return OfficeFileFormat.Unknown;
            }
        }

        /// <summary>获取指定格式对应的标准文件扩展名。Gets the standard file extension for the given format.</summary>
        /// <param name="format">文件格式。The file format.</param>
        /// <returns>如 ".pdf"、".xlsx"；无对应扩展名时返回空字符串。For example ".pdf", ".xlsx"; an empty string when there is no matching extension.</returns>
        public static string GetExtension(OfficeFileFormat format)
        {
            switch (format)
            {
                case OfficeFileFormat.Pdf: return ".pdf";
                case OfficeFileFormat.Docx: return ".docx";
                case OfficeFileFormat.Dotx: return ".dotx";
                case OfficeFileFormat.Doc: return ".doc";
                case OfficeFileFormat.Xlsx: return ".xlsx";
                case OfficeFileFormat.Xls: return ".xls";
                case OfficeFileFormat.Pptx: return ".pptx";
                case OfficeFileFormat.Potx: return ".potx";
                case OfficeFileFormat.Rtf: return ".rtf";
                case OfficeFileFormat.Html: return ".html";
                case OfficeFileFormat.Odt: return ".odt";
                case OfficeFileFormat.Text: return ".txt";
                case OfficeFileFormat.Markdown: return ".md";
                default: return string.Empty;
            }
        }

        /// <summary>获取适合在界面中展示的格式中文名称。Gets the Chinese display name of the format, suitable for UI presentation.</summary>
        /// <param name="format">文件格式。The file format.</param>
        /// <returns>如 "Word 文档"、"未知格式"。For example "Word 文档", "未知格式".</returns>
        public static string GetDisplayName(OfficeFileFormat format)
        {
            switch (format)
            {
                case OfficeFileFormat.Pdf: return "PDF";
                case OfficeFileFormat.Docx: return "Word 文档";
                case OfficeFileFormat.Dotx: return "Word 模板";
                case OfficeFileFormat.Doc: return "Word 97-2003 文档";
                case OfficeFileFormat.Xlsx: return "Excel 工作簿";
                case OfficeFileFormat.Xls: return "Excel 97-2003 工作簿";
                case OfficeFileFormat.Pptx: return "PowerPoint 演示文稿";
                case OfficeFileFormat.Potx: return "PowerPoint 模板";
                case OfficeFileFormat.Rtf: return "RTF 富文本";
                case OfficeFileFormat.Html: return "HTML 网页";
                case OfficeFileFormat.Odt: return "OpenDocument 文本";
                case OfficeFileFormat.Text: return "纯文本";
                case OfficeFileFormat.Markdown: return "Markdown 文档";
                default: return "未知格式";
            }
        }

        /// <summary>将流的全部内容读取为字节数组；<see cref="MemoryStream"/> 直接调用其 ToArray 以避免中转拷贝。Reads the entire stream into a byte array; for <see cref="MemoryStream"/> it calls ToArray directly to avoid an intermediate copy.</summary>
        /// <param name="stream">源流。The source stream.</param>
        /// <returns>流的字节内容。The byte content of the stream.</returns>
        internal static byte[] ReadAll(Stream stream)
        {
            MemoryStream memory = stream as MemoryStream;
            if (memory != null) return memory.ToArray();
            using (MemoryStream buffer = new MemoryStream())
            {
                stream.CopyTo(buffer);
                return buffer.ToArray();
            }
        }
    }
}
