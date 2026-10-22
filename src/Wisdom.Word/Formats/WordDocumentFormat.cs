using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>Word 文档格式：OOXML 文档/模板、旧版二进制文档、RTF、HTML、ODT 与纯文本。Word document formats: OOXML document/template, legacy binary document, RTF, HTML, ODT and plain text.</summary>
    public enum WordDocumentFormat
    {
        /// <summary>未知格式。Unknown format.</summary>
        Unknown = 0,
        /// <summary>Word 2007+ 文档（.docx）。Word 2007+ document (.docx).</summary>
        Docx = 1,
        /// <summary>Word 2007+ 模板（.dotx）。Word 2007+ template (.dotx).</summary>
        Dotx = 2,
        /// <summary>Word 97-2003 二进制文档（.doc），仅支持读取。Word 97-2003 binary document (.doc), read-only.</summary>
        Doc = 3,
        /// <summary>Rich Text Format（.rtf）。Rich Text Format (.rtf).</summary>
        Rtf = 4,
        /// <summary>HTML 网页（.html/.htm）。HTML document (.html/.htm).</summary>
        Html = 5,
        /// <summary>OpenDocument Text（.odt）。OpenDocument Text (.odt).</summary>
        Odt = 6,
        /// <summary>纯文本（.txt），仅保留文本与段落。Plain text (.txt), preserving text and paragraphs only.</summary>
        Text = 7
    }

    /// <summary>Word 格式辅助：扩展名映射、内容嗅探与格式显示名。Word format helper: extension mapping, content sniffing and display names.</summary>
    public static class WordDocumentFormats
    {
        /// <summary>OLE 复合文档的 8 字节文件头签名。The 8-byte header signature of an OLE compound document.</summary>
        private static readonly byte[] OleSignature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

        /// <summary>从扩展名映射 Word 格式，大小写不敏感。Maps an extension to a Word format, case-insensitively.</summary>
        /// <param name="extension">扩展名，如 ".docx"。The extension, such as ".docx".</param>
        /// <returns>对应格式，未知时返回 <see cref="WordDocumentFormat.Unknown"/>。The matching format, or <see cref="WordDocumentFormat.Unknown"/>.</returns>
        public static WordDocumentFormat FromExtension(string extension)
        {
            if (string.IsNullOrEmpty(extension)) return WordDocumentFormat.Unknown;
            switch (extension.ToLowerInvariant())
            {
                case ".docx": return WordDocumentFormat.Docx;
                case ".dotx": return WordDocumentFormat.Dotx;
                case ".dot": return WordDocumentFormat.Doc;
                case ".doc": return WordDocumentFormat.Doc;
                case ".rtf": return WordDocumentFormat.Rtf;
                case ".html":
                case ".htm":
                case ".xhtml": return WordDocumentFormat.Html;
                case ".odt": return WordDocumentFormat.Odt;
                case ".txt":
                case ".text": return WordDocumentFormat.Text;
                default: return WordDocumentFormat.Unknown;
            }
        }

        /// <summary>获取格式对应的标准扩展名。Gets the standard extension for a format.</summary>
        /// <param name="format">文档格式。The document format.</param>
        /// <returns>扩展名，未知格式返回空字符串。The extension, or an empty string for unknown formats.</returns>
        public static string GetExtension(WordDocumentFormat format)
        {
            switch (format)
            {
                case WordDocumentFormat.Docx: return ".docx";
                case WordDocumentFormat.Dotx: return ".dotx";
                case WordDocumentFormat.Doc: return ".doc";
                case WordDocumentFormat.Rtf: return ".rtf";
                case WordDocumentFormat.Html: return ".html";
                case WordDocumentFormat.Odt: return ".odt";
                case WordDocumentFormat.Text: return ".txt";
                default: return string.Empty;
            }
        }

        /// <summary>获取格式的中文显示名。Gets the Chinese display name of a format.</summary>
        /// <param name="format">文档格式。The document format.</param>
        /// <returns>显示名。The display name.</returns>
        public static string GetDisplayName(WordDocumentFormat format)
        {
            switch (format)
            {
                case WordDocumentFormat.Docx: return "Word 文档";
                case WordDocumentFormat.Dotx: return "Word 模板";
                case WordDocumentFormat.Doc: return "Word 97-2003 文档";
                case WordDocumentFormat.Rtf: return "RTF 富文本";
                case WordDocumentFormat.Html: return "HTML 网页";
                case WordDocumentFormat.Odt: return "OpenDocument 文本";
                case WordDocumentFormat.Text: return "纯文本";
                default: return "未知格式";
            }
        }

        /// <summary>按文件内容嗅探 Word 格式；无法判断时返回 Unknown。Sniffs the Word format from file content; returns Unknown when it cannot decide.</summary>
        /// <param name="data">文件字节。The file bytes.</param>
        /// <returns>嗅探出的格式。The sniffed format.</returns>
        public static WordDocumentFormat Detect(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length == 0) return WordDocumentFormat.Unknown;

            int offset = 0;
            if (data.Length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) offset = 3;
            if (data.Length - offset >= 5
                && data[offset] == (byte)'{'
                && (data[offset + 1] == (byte)'\\')
                && (data[offset + 2] == (byte)'r' || data[offset + 2] == (byte)'R')
                && (data[offset + 3] == (byte)'t' || data[offset + 3] == (byte)'T')
                && (data[offset + 4] == (byte)'f' || data[offset + 4] == (byte)'F'))
            {
                return WordDocumentFormat.Rtf;
            }

            if (data.Length >= 4 && data[0] == (byte)'P' && data[1] == (byte)'K'
                && (data[2] == 3 || data[2] == 5) && (data[3] == 4 || data[3] == 6))
            {
                return DetectPackage(data);
            }

            bool ole = data.Length >= 8;
            for (int i = 0; ole && i < OleSignature.Length; i++)
            {
                if (data[i] != OleSignature[i]) ole = false;
            }
            if (ole) return WordDocumentFormat.Doc;

            char first = FirstNonWhitespace(data, offset);
            if (first == '<') return WordDocumentFormat.Html;
            return WordDocumentFormat.Text;
        }

        /// <summary>检查 ZIP 包是 OOXML 模板/文档还是 ODT。Checks whether a ZIP package is an OOXML template/document or ODT.</summary>
        /// <param name="data">ZIP 字节。The ZIP bytes.</param>
        /// <returns>Docx、Dotx、Odt 或 Unknown。Docx, Dotx, Odt or Unknown.</returns>
        private static WordDocumentFormat DetectPackage(byte[] data)
        {
            try
            {
                using (MemoryStream stream = new MemoryStream(data, false))
                using (ZipArchive archive = PackageArchive.OpenForRead(stream, true))
                {
                    ZipArchiveEntry mimetype = FindEntry(archive, "mimetype");
                    if (mimetype != null)
                    {
                        string value = ReadEntryText(mimetype, 256);
                        if (value.IndexOf("opendocument.text", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return WordDocumentFormat.Odt;
                        }
                    }
                    if (FindEntry(archive, "word/document.xml") != null)
                    {
                        ZipArchiveEntry types = FindEntry(archive, "[Content_Types].xml");
                        if (types != null)
                        {
                            string content = ReadEntryText(types, 65536);
                            if (content.IndexOf("wordprocessingml.template.main+xml", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                return WordDocumentFormat.Dotx;
                            }
                        }
                        return WordDocumentFormat.Docx;
                    }
                    if (FindEntry(archive, "content.xml") != null && FindEntry(archive, "META-INF/manifest.xml") != null)
                    {
                        return WordDocumentFormat.Odt;
                    }
                }
            }
            catch (InvalidDataException)
            {
            }
            catch (IOException)
            {
            }
            catch (ArgumentException)
            {
            }
            return WordDocumentFormat.Unknown;
        }

        /// <summary>在 ZIP 包中按名称查找条目，忽略大小写与路径分隔符差异。Finds an entry by name, ignoring case and path separator differences.</summary>
        internal static ZipArchiveEntry FindEntry(ZipArchive archive, string name)
        {
            for (int i = 0; i < archive.Entries.Count; i++)
            {
                string entryName = archive.Entries[i].FullName.Replace('\\', '/');
                if (string.Equals(entryName, name, StringComparison.OrdinalIgnoreCase)) return archive.Entries[i];
            }
            return null;
        }

        /// <summary>读取 ZIP 条目文本，最多读取指定字符数，避免解压炸弹。Reads entry text up to a character cap, avoiding decompression bombs.</summary>
        private static string ReadEntryText(ZipArchiveEntry entry, int maxChars)
        {
            using (Stream stream = entry.Open())
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8, true))
            {
                char[] buffer = new char[Math.Min(maxChars, 4096)];
                StringBuilder builder = new StringBuilder();
                while (builder.Length < maxChars)
                {
                    int read = reader.Read(buffer, 0, Math.Min(buffer.Length, maxChars - builder.Length));
                    if (read <= 0) break;
                    builder.Append(buffer, 0, read);
                }
                return builder.ToString();
            }
        }

        /// <summary>返回字节数组中第一个非空白字符，空内容返回 '\0'。Returns the first non-whitespace character, or '\0' for empty content.</summary>
        private static char FirstNonWhitespace(byte[] data, int offset)
        {
            for (int i = offset; i < data.Length && i < offset + 4096; i++)
            {
                char value = (char)data[i];
                if (!char.IsWhiteSpace(value)) return value;
            }
            return '\0';
        }
    }
}
