using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 读取旧版 Word 97-2003（.doc）二进制文档：通过 FIB piece table（CLX/PlcPcd）提取正文并重建段落与简单表格。
    /// Reads legacy Word 97-2003 (.doc) binary files. Extracts the main text using the
    /// FIB piece table (CLX/PlcPcd) and rebuilds paragraphs and simple tables.
    /// </summary>
    public static class DocReader
    {
        private const ushort WordIdentifier = 0xA5EC;
        private const int FibFcClxOffset = 0x01A2;
        private const int FibCcpTextOffset = 0x004C;

        /// <summary>
        /// 从磁盘路径读取 .doc 文档并转换为统一文档模型（若内容为 ZIP 包则按 .docx 解析）。
        /// Loads a .doc file from a disk path into the unified document model, parsing as .docx when the content is a ZIP package.
        /// </summary>
        /// <param name="path">文档文件路径。The document file path.</param>
        /// <returns>统一文档模型。The unified document model.</returns>
        public static OfficeDocument Load(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(File.ReadAllBytes(path));
        }

        /// <summary>
        /// 从流读取 .doc 文档并转换为统一文档模型（若内容为 ZIP 包则按 .docx 解析）。
        /// Loads a .doc document from a stream into the unified document model, parsing as .docx when the content is a ZIP package.
        /// </summary>
        /// <param name="stream">包含文档内容的流。The stream holding the document content.</param>
        /// <returns>统一文档模型。The unified document model.</returns>
        public static OfficeDocument Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return Load(OfficePackageLimits.ReadAllBytes(stream));
        }

        /// <summary>
        /// 核心解析入口：读取 OLE 复合文件的 WordDocument 流，校验 FIB 签名后经 piece table 解码正文并重建文档。
        /// Core parsing entry: reads the WordDocument stream of the OLE compound file, validates the FIB signature,
        /// decodes the body text through the piece table and rebuilds the document.
        /// </summary>
        /// <param name="data">文档原始字节。The raw document bytes.</param>
        /// <returns>统一文档模型。The unified document model.</returns>
        public static OfficeDocument Load(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length >= 4 && data[0] == 0x50 && data[1] == 0x4B)
            {
                return DocxReader.Load(data);
            }
            if (!CompoundFile.IsCompoundFile(data))
            {
                throw new InvalidDataException("The file is not a legacy Word document.");
            }

            CompoundFile compound = new CompoundFile(data);
            byte[] wordStream = compound.TryGetStream("WordDocument");
            if (wordStream == null || wordStream.Length < 0x200)
            {
                throw new InvalidDataException("The WordDocument stream is missing.");
            }

            ushort identifier = ByteUtils.ReadUInt16LittleEndian(wordStream, 0);
            if (identifier != WordIdentifier)
            {
                throw new InvalidDataException("The WordDocument stream has an invalid FIB signature.");
            }

            ushort flags = ByteUtils.ReadUInt16LittleEndian(wordStream, 0x0A);
            if ((flags & 0x0100) != 0)
            {
                throw new NotSupportedException("Encrypted Word documents are not supported.");
            }

            ushort nFib = ByteUtils.ReadUInt16LittleEndian(wordStream, 2);
            ushort languageId = ByteUtils.ReadUInt16LittleEndian(wordStream, 0x06);
            int codePage = GetDocumentCodePage(languageId);
            bool useTable1 = (flags & 0x0200) != 0;

            int ccpText = 0;
            if (wordStream.Length >= FibCcpTextOffset + 4)
            {
                ccpText = ByteUtils.ReadInt32LittleEndian(wordStream, FibCcpTextOffset);
            }

            string text = null;
            if (nFib >= 0x00C1 && wordStream.Length >= FibFcClxOffset + 8)
            {
                byte[] tableStream = compound.TryGetStream(useTable1 ? "1Table" : "0Table");
                if (tableStream != null)
                {
                    uint fcClx = ByteUtils.ReadUInt32LittleEndian(wordStream, FibFcClxOffset);
                    uint lcbClx = ByteUtils.ReadUInt32LittleEndian(wordStream, FibFcClxOffset + 4);
                    if (fcClx < tableStream.Length && lcbClx > 0)
                    {
                        text = ReadPieceTable(wordStream, tableStream, (int)fcClx, (int)Math.Min(lcbClx, (uint)(tableStream.Length - fcClx)), ccpText, codePage);
                    }
                }
            }

            if (text == null)
            {
                int fcMin = ByteUtils.ReadInt32LittleEndian(wordStream, 0x18);
                int fcMac = ByteUtils.ReadInt32LittleEndian(wordStream, 0x1C);
                // 两个偏移都直接来自文件，下界必须先钉住：fcMin 为负时 fcMac > fcMin 照样通过，
                // 减法随即溢出成负数或正的越界长度，把 ArgumentOutOfRangeException 直接丢给调用方，
                // 而不是文档约定的 InvalidDataException。钉住 0 <= fcMin < fcMac <= Length 之后
                // 差值天然不会溢出，也不必再补一次减法比较。
                // Both offsets come straight from the file, so the lower bound has to be pinned first: with a
                // negative fcMin the "fcMac > fcMin" gate still passes, the subtraction then wraps into a
                // negative or an out-of-range length, and ArgumentOutOfRangeException escapes to the caller
                // instead of the documented InvalidDataException. Once 0 <= fcMin < fcMac <= Length holds the
                // difference cannot overflow and no extra subtractive check is needed.
                if (fcMin >= 0 && fcMac > fcMin && fcMac <= wordStream.Length)
                {
                    text = DecodeAnsi(wordStream, fcMin, fcMac - fcMin, codePage);
                }
                else
                {
                    text = string.Empty;
                }
            }

            OfficeDocument document = BuildDocument(text);
            ReadMetadata(compound, document.Metadata);
            return document;
        }

        /// <summary>
        /// 仅提取文档正文的纯文本内容。
        /// Extracts only the plain text body of the document.
        /// </summary>
        /// <param name="path">文档文件路径。The document file path.</param>
        /// <returns>正文纯文本。The plain body text.</returns>
        public static string ExtractText(string path)
        {
            return Load(path).GetPlainText();
        }

        private static string ReadPieceTable(byte[] wordStream, byte[] tableStream, int clxOffset, int clxLength, int ccpText, int codePage)
        {
            int position = clxOffset;
            int end = clxOffset + clxLength;
            while (position < end)
            {
                byte type = tableStream[position];
                if (type == 0x01)
                {
                    if (position + 3 > end) break;
                    int size = ByteUtils.ReadUInt16LittleEndian(tableStream, position + 1);
                    position += 3 + size;
                    continue;
                }
                if (type != 0x02) break;

                if (position + 5 > end) break;
                int plcLength = ByteUtils.ReadInt32LittleEndian(tableStream, position + 1);
                int plcStart = position + 5;
                if (plcLength < 4 || (long)plcStart + plcLength > tableStream.Length) break;

                int pieceCount = (plcLength - 4) / 12;
                List<int> cps = new List<int>(pieceCount + 1);
                for (int i = 0; i <= pieceCount; i++)
                {
                    cps.Add(ByteUtils.ReadInt32LittleEndian(tableStream, plcStart + i * 4));
                }
                int pcdStart = plcStart + (pieceCount + 1) * 4;

                // 输出字符数不可能超过 WordDocument 流的字节数：ccpText 是损坏文件可直接放大的字段，
                // 容量与上限都以流长为硬上界，既避免预申请 GB 级缓冲，也避免拼接出不受限制的文本。
                // The output can never hold more characters than the WordDocument stream has bytes. ccpText is a
                // field a corrupt file can inflate, so both the capacity and the limit are capped by the stream
                // length, which prevents a gigabyte-sized pre-allocation and unbounded concatenation alike.
                int outputCap = wordStream.Length;
                StringBuilder builder = new StringBuilder(ccpText > 0 ? Math.Min(ccpText, outputCap) : 1024);
                int extracted = 0;
                int limit = ccpText > 0 ? Math.Min(ccpText, outputCap) : outputCap;
                for (int i = 0; i < pieceCount && extracted < limit; i++)
                {
                    int cpStart = cps[i];
                    int cpEnd = cps[i + 1];
                    if (cpEnd <= cpStart) continue;
                    long span = (long)cpEnd - cpStart;
                    int cpCount = span > limit - extracted ? limit - extracted : (int)span;

                    int fc = ByteUtils.ReadInt32LittleEndian(tableStream, pcdStart + i * 8 + 2);
                    bool compressed = (fc & 0x40000000) != 0;
                    int offset = fc & 0x3FFFFFFF;
                    if (compressed)
                    {
                        offset /= 2;
                        if (offset < 0 || cpCount > wordStream.Length - offset) continue;
                        builder.Append(DecodeAnsi(wordStream, offset, cpCount, codePage));
                    }
                    else
                    {
                        long bytes = (long)cpCount * 2;
                        if (offset < 0 || bytes > (long)wordStream.Length - offset) continue;
                        builder.Append(Encoding.Unicode.GetString(wordStream, offset, (int)bytes));
                    }
                    extracted += cpCount;
                }
                return builder.ToString();
            }
            return null;
        }

        private static string DecodeAnsi(byte[] data, int offset, int count, int codePage)
        {
            Encoding encoding = TryGetEncoding(codePage);
            if (encoding != null)
            {
                return encoding.GetString(data, offset, count);
            }
            return DecodeWindows1252(data, offset, count);
        }

        private static int GetDocumentCodePage(ushort languageId)
        {
            switch (languageId)
            {
                case 0x0804:
                case 0x1004:
                    return 936;
                case 0x0404:
                case 0x0C04:
                case 0x1404:
                    return 950;
                case 0x0411:
                    return 932;
                case 0x0412:
                    return 949;
                case 0x0402:
                case 0x0419:
                case 0x0422:
                case 0x0423:
                case 0x042F:
                case 0x043F:
                case 0x0440:
                case 0x0444:
                case 0x0450:
                    return 1251;
                case 0x0405:
                case 0x040E:
                case 0x0415:
                case 0x0418:
                case 0x041A:
                case 0x041B:
                case 0x041C:
                case 0x0424:
                case 0x0425:
                case 0x0426:
                case 0x0427:
                    return 1250;
                case 0x0408:
                    return 1253;
                case 0x041F:
                    return 1254;
                case 0x040D:
                    return 1255;
                case 0x0401:
                    return 1256;
                default:
                    return 1252;
            }
        }

        private static Encoding TryGetEncoding(int codePage)
        {
            if (codePage <= 0) return null;
            try
            {
                return Encoding.GetEncoding(codePage);
            }
            catch (Exception)
            {
            }
            try
            {
                Type providerType = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages", false);
                if (providerType != null)
                {
                    System.Reflection.PropertyInfo instance = providerType.GetProperty(
                        "Instance", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    System.Reflection.MethodInfo register = typeof(Encoding).GetMethod(
                        "RegisterProvider", new Type[] { typeof(EncodingProvider) });
                    if (instance != null && register != null)
                    {
                        register.Invoke(null, new object[] { instance.GetValue(null, null) });
                        return Encoding.GetEncoding(codePage);
                    }
                }
            }
            catch (Exception)
            {
            }
            return null;
        }

        private static string DecodeWindows1252(byte[] data, int offset, int count)
        {
            StringBuilder builder = new StringBuilder(count);
            int end = Math.Min(data.Length, offset + count);
            for (int i = offset; i < end; i++)
            {
                builder.Append(DecodeWindows1252Byte(data[i]));
            }
            return builder.ToString();
        }

        private static char DecodeWindows1252Byte(byte value)
        {
            if (value < 0x80 || value >= 0xA0) return (char)value;
            switch (value)
            {
                case 0x80: return '\u20AC';
                case 0x82: return '\u201A';
                case 0x83: return '\u0192';
                case 0x84: return '\u201E';
                case 0x85: return '\u2026';
                case 0x86: return '\u2020';
                case 0x87: return '\u2021';
                case 0x88: return '\u02C6';
                case 0x89: return '\u2030';
                case 0x8A: return '\u0160';
                case 0x8B: return '\u2039';
                case 0x8C: return '\u0152';
                case 0x8E: return '\u017D';
                case 0x91: return '\u2018';
                case 0x92: return '\u2019';
                case 0x93: return '\u201C';
                case 0x94: return '\u201D';
                case 0x95: return '\u2022';
                case 0x96: return '\u2013';
                case 0x97: return '\u2014';
                case 0x98: return '\u02DC';
                case 0x99: return '\u2122';
                case 0x9A: return '\u0161';
                case 0x9B: return '\u203A';
                case 0x9C: return '\u0153';
                case 0x9E: return '\u017E';
                case 0x9F: return '\u0178';
                default: return (char)value;
            }
        }

        private static OfficeDocument BuildDocument(string text)
        {
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();
            StringBuilder buffer = new StringBuilder();
            List<string> cells = new List<string>();

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\r')
                {
                    if (cells.Count > 0)
                    {
                        if (buffer.Length > 0)
                        {
                            cells.Add(buffer.ToString());
                            buffer.Length = 0;
                        }
                        FlushTable(section, cells);
                        cells.Clear();
                        continue;
                    }
                    section.Blocks.Add(new Paragraph(buffer.ToString()));
                    buffer.Length = 0;
                    continue;
                }
                if (c == '\x07')
                {
                    if (buffer.Length > 0)
                    {
                        cells.Add(buffer.ToString());
                        buffer.Length = 0;
                    }
                    else if (cells.Count > 0)
                    {
                        FlushTable(section, cells);
                        cells.Clear();
                    }
                    continue;
                }
                if (c == '\x0C')
                {
                    if (buffer.Length > 0)
                    {
                        section.Blocks.Add(new Paragraph(buffer.ToString()));
                        buffer.Length = 0;
                    }
                    section.Blocks.Add(new PageBreak());
                    continue;
                }
                if (c == '\x0B')
                {
                    buffer.Append('\n');
                    continue;
                }
                if (c == '\x1E')
                {
                    buffer.Append('-');
                    continue;
                }
                if (c == '\x13' || c == '\x14' || c == '\x15')
                {
                    continue;
                }
                if (c == '\x1F' || c == '\x01' || c == '\x02' || c == '\x08' || c == '\x00')
                {
                    continue;
                }
                if (c < 32 && c != '\t' && c != '\n')
                {
                    continue;
                }
                buffer.Append(c);
            }

            if (buffer.Length > 0) section.Blocks.Add(new Paragraph(buffer.ToString()));
            if (cells.Count > 0) FlushTable(section, cells);
            if (section.Blocks.Count == 0) section.AddParagraph(string.Empty);
            return document;
        }

        private static void FlushTable(Section section, List<string> cells)
        {
            if (cells.Count == 0) return;
            int columns = cells.Count;
            Table table = new Table();
            table.RepeatHeaderRow = false;
            double width = Math.Max(36, section.PageSetup.Width - section.PageSetup.MarginLeft - section.PageSetup.MarginRight);
            for (int i = 0; i < columns; i++) table.Columns.Add(new TableColumn(width / columns));
            TableRow row = table.AddRow();
            for (int i = 0; i < cells.Count; i++)
            {
                row.Cells.Add(new TableCell(cells[i]));
            }
            section.Blocks.Add(table);
        }

        private static void ReadMetadata(CompoundFile compound, DocumentMetadata metadata)
        {
            try
            {
                byte[] stream = compound.TryGetStream("\u0005SummaryInformation");
                if (stream == null || stream.Length < 48) return;
                if (ByteUtils.ReadUInt16LittleEndian(stream, 0) != 0xFFFE) return;
                int sections = ByteUtils.ReadInt32LittleEndian(stream, 24);
                if (sections < 1) return;
                int sectionOffset = ByteUtils.ReadInt32LittleEndian(stream, 44);
                // 减法比较：sectionOffset 来自文件，加上 8 之后可能回绕成负数而骗过上限。
                // Compare by subtraction: sectionOffset comes from the file and adding 8 can wrap into a
                // negative that then slips past the upper bound.
                if (sectionOffset < 0 || 8 > stream.Length - sectionOffset) return;
                int propertyCount = ByteUtils.ReadInt32LittleEndian(stream, sectionOffset + 4);
                if (propertyCount <= 0 || propertyCount > 1024) return;

                int codePage = ReadPropertyCodePage(stream, sectionOffset, propertyCount);

                for (int i = 0; i < propertyCount; i++)
                {
                    // i 与 propertyCount 同量级，两处偏移都用减法核对，避免相加回绕。
                    // i scales with propertyCount, so both offsets are checked subtractively instead of by
                    // adding them together.
                    int entry = sectionOffset + 8 + i * 8;
                    if (entry < 0 || 8 > stream.Length - entry) break;
                    int propertyId = ByteUtils.ReadInt32LittleEndian(stream, entry);
                    int valueOffset = ByteUtils.ReadInt32LittleEndian(stream, entry + 4);
                    int position = sectionOffset + valueOffset;
                    if (position < 0 || 4 > stream.Length - position) continue;
                    int type = ByteUtils.ReadInt32LittleEndian(stream, position);
                    switch (propertyId)
                    {
                        case 2:
                            metadata.Title = ReadPropertyString(stream, position, type, codePage) ?? metadata.Title;
                            break;
                        case 3:
                            metadata.Subject = ReadPropertyString(stream, position, type, codePage) ?? metadata.Subject;
                            break;
                        case 4:
                            metadata.Author = ReadPropertyString(stream, position, type, codePage) ?? metadata.Author;
                            break;
                        case 5:
                            metadata.Keywords = ReadPropertyString(stream, position, type, codePage) ?? metadata.Keywords;
                            break;
                        case 8:
                            if (string.IsNullOrEmpty(metadata.Author))
                            {
                                metadata.Author = ReadPropertyString(stream, position, type, codePage) ?? metadata.Author;
                            }
                            break;
                        case 12:
                            DateTime created;
                            if (ReadFileTime(stream, position, type, out created)) metadata.Created = created;
                            break;
                        case 13:
                            DateTime modified;
                            if (ReadFileTime(stream, position, type, out modified)) metadata.Modified = modified;
                            break;
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static int ReadPropertyCodePage(byte[] stream, int sectionOffset, int propertyCount)
        {
            for (int i = 0; i < propertyCount; i++)
            {
                int entry = sectionOffset + 8 + i * 8;
                if (entry + 8 > stream.Length) break;
                if (ByteUtils.ReadInt32LittleEndian(stream, entry) != 1) continue;
                int position = sectionOffset + ByteUtils.ReadInt32LittleEndian(stream, entry + 4);
                if (position < 0 || position + 6 > stream.Length) break;
                int type = ByteUtils.ReadInt32LittleEndian(stream, position);
                if (type == 0x02 || type == 0x12)
                {
                    int value = ByteUtils.ReadUInt16LittleEndian(stream, position + 4);
                    if (value > 0) return value;
                }
                break;
            }
            return 1252;
        }

        private static string ReadPropertyString(byte[] stream, int position, int type, int codePage)
        {
            if (type == 0x1E)
            {
                int length = ByteUtils.ReadInt32LittleEndian(stream, position + 4);
                if (length <= 0) return string.Empty;
                int start = position + 8;
                if (length > stream.Length - start) return null;
                int textLength = length;
                while (textLength > 0 && stream[start + textLength - 1] == 0) textLength--;
                return DecodeAnsi(stream, start, textLength, codePage);
            }
            if (type == 0x1F)
            {
                int chars = ByteUtils.ReadInt32LittleEndian(stream, position + 4);
                if (chars <= 0) return string.Empty;
                int start = position + 8;
                // UTF-16 的长度要先除后比：chars * 2 这一乘就能回绕成负数而通过上限检查。
                // Divide before comparing for the UTF-16 length: the multiply by 2 alone can wrap into a
                // negative that then passes the upper bound.
                if (chars > (stream.Length - start) / 2) return null;
                int textChars = chars;
                while (textChars > 0 && stream[start + textChars * 2 - 2] == 0 && stream[start + textChars * 2 - 1] == 0) textChars--;
                return Encoding.Unicode.GetString(stream, start, textChars * 2);
            }
            return null;
        }

        private static bool ReadFileTime(byte[] stream, int position, int type, out DateTime value)
        {
            value = DateTime.MinValue;
            if (type != 0x40 || position + 12 > stream.Length) return false;
            uint low = ByteUtils.ReadUInt32LittleEndian(stream, position + 4);
            uint high = ByteUtils.ReadUInt32LittleEndian(stream, position + 8);
            long fileTime = (long)high << 32 | low;
            if (fileTime <= 0) return false;
            try
            {
                value = DateTime.FromFileTimeUtc(fileTime);
                return true;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
        }
    }
}
