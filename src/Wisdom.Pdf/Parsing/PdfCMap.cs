using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// PDF 字符映射（CMap）表，把字体编码值映射回 Unicode 文本。
    /// A PDF character map (CMap) that maps font encoding values back to Unicode text.
    /// </summary>
    /// <remarks>仅解析 Type 0/Type 1 字体常用的 codespace、bfchar 与 bfrange 段；解析失败的字句会被跳过而不抛异常，以容忍不规范导出的 PDF。
    /// Only the codespace, bfchar and bfrange sections commonly used by Type 0/Type 1 fonts are parsed.
    /// Unparseable tokens are skipped instead of throwing, to tolerate badly exported PDFs.</remarks>
    internal sealed class PdfCMap
    {
        /// <summary>编码值到 Unicode 文本的映射。Mapping from encoding values to Unicode text.</summary>
        private readonly Dictionary<int, string> _map = new Dictionary<int, string>();

        /// <summary>
        /// 解析 CMap 流内容并构建编码到 Unicode 的映射。
        /// Parses CMap stream content and builds the encoding-to-Unicode mapping.
        /// </summary>
        /// <param name="data">CMap 流解码后的字节，null 或空时返回空表。The decoded bytes of the CMap stream; null or empty data yields an empty map.</param>
        /// <returns>解析完成的映射表。The parsed mapping table.</returns>
        public static PdfCMap Parse(byte[] data)
        {
            PdfCMap cmap = new PdfCMap();
            if (data == null || data.Length == 0) return cmap;

            PdfParser parser = new PdfParser(data, 0, null);
            List<PdfObject> pending = new List<PdfObject>();
            string mode = null;
            while (!parser.EndOfData)
            {
                parser.SkipWhitespaceAndComments();
                if (parser.EndOfData) break;
                byte next = parser.Data[parser.Position];
                if (next == (byte)'<' || next == (byte)'[' || next == (byte)'/' || next == (byte)'('
                    || (next >= (byte)'0' && next <= (byte)'9') || next == (byte)'-' || next == (byte)'+'
                    || next == (byte)'.' || next == (byte)'t' || next == (byte)'f' || next == (byte)'n')
                {
                    try
                    {
                        pending.Add(parser.ParseObject());
                    }
                    catch (InvalidDataException)
                    {
                        parser.Position++;
                    }
                    continue;
                }

                string keyword = parser.ReadKeyword();
                if (keyword == "beginbfchar")
                {
                    mode = "char";
                    pending.Clear();
                }
                else if (keyword == "beginbfrange")
                {
                    mode = "range";
                    pending.Clear();
                }
                else if (keyword == "endbfchar")
                {
                    ProcessChars(cmap, pending);
                    pending.Clear();
                    mode = null;
                }
                else if (keyword == "endbfrange")
                {
                    ProcessRanges(cmap, pending);
                    pending.Clear();
                    mode = null;
                }
                else if (keyword == "begincodespacerange" || keyword == "endcodespacerange")
                {
                    pending.Clear();
                }
                else if (mode == null)
                {
                    pending.Clear();
                }
            }
            return cmap;
        }

        /// <summary>
        /// 处理 beginbfchar 段：逐对读取源编码与目标 Unicode 串。
        /// Processes a beginbfchar section: reads source-encoding / destination-Unicode pairs in order.
        /// </summary>
        /// <param name="cmap">待填充的映射表。The map to populate.</param>
        /// <param name="pending">本段内收集到的对象，按「源、目标」成对排列。Objects collected within the section, arranged as source/destination pairs.</param>
        private static void ProcessChars(PdfCMap cmap, List<PdfObject> pending)
        {
            for (int i = 0; i + 1 < pending.Count; i += 2)
            {
                PdfString source = pending[i] as PdfString;
                PdfString destination = pending[i + 1] as PdfString;
                if (source == null || destination == null) continue;
                int code = ToCode(source.Bytes);
                string text = DecodeUtf16BigEndian(destination.Bytes);
                if (cmap._map.Count >= PdfReadLimits.MaxGlyphMapEntries) break;
                cmap._map[code] = text;
            }
        }

        /// <summary>
        /// 处理 beginbfrange 段：目标可以是单个字符串（末位码点按区间偏移递增）或字符串数组（逐一对应）。
        /// Processes a beginbfrange section: the destination may be a single string (the last code point
        /// increments with the range offset) or an array of strings (matched one by one).
        /// </summary>
        /// <param name="cmap">待填充的映射表。The map to populate.</param>
        /// <param name="pending">本段内收集到的对象，按「下界、上界、目标」三元排列。Objects collected within the section, arranged as low/high/destination triples.</param>
        /// <remarks>区间跨度超过 65536 时视为异常数据直接丢弃，跨度的比较用 long 以免 int 回绕；
        /// 整表条目数再按 MaxGlyphMapEntries 兜底，防止大量「各自合法」的区间合力把映射表堆爆。
        /// Ranges wider than 65536 codes are treated as corrupt data and dropped, comparing the span in long
        /// so the int subtraction cannot wrap; the whole map is additionally capped by MaxGlyphMapEntries,
        /// because many individually legal ranges can still exhaust memory together.</remarks>
        private static void ProcessRanges(PdfCMap cmap, List<PdfObject> pending)
        {
            for (int i = 0; i + 2 < pending.Count; i += 3)
            {
                PdfString lowString = pending[i] as PdfString;
                PdfString highString = pending[i + 1] as PdfString;
                if (lowString == null || highString == null) continue;
                int low = ToCode(lowString.Bytes);
                int high = ToCode(highString.Bytes);
                // 跨度必须按 long 计算：ToCode 用四字节拼键，最高位置 1 时键为负值，
                // 而 int 相减同样回绕，「high >= low 且差值不大」的假象会让区间达到二十亿个码位。
                // The span must be computed in long: ToCode packs four bytes, so a set high bit yields a negative
                // key, and the int subtraction wraps too — the range then looks legal while spanning two billion codes.
                long span = (long)high - (long)low;
                if (span < 0 || span > 65536) continue;
                int room = PdfReadLimits.MaxGlyphMapEntries - cmap._map.Count;
                if (room <= 0) continue;

                PdfString destination = pending[i + 2] as PdfString;
                PdfArray destinations = pending[i + 2] as PdfArray;
                long stop = Math.Min((long)high, (long)low + room - 1);
                if (destination != null)
                {
                    char[] baseChars = DecodeUtf16BigEndian(destination.Bytes).ToCharArray();
                    for (int code = low; code <= stop; code++)
                    {
                        int offset = code - low;
                        char[] chars = (char[])baseChars.Clone();
                        if (chars.Length > 0)
                        {
                            int value = chars[chars.Length - 1] + offset;
                            chars[chars.Length - 1] = (char)(value & 0xFFFF);
                        }
                        cmap._map[code] = new string(chars);
                    }
                    continue;
                }
                if (destinations != null)
                {
                    for (long j = 0; j < destinations.Count && (long)low + j <= stop; j++)
                    {
                        PdfString item = destinations[(int)j] as PdfString;
                        if (item == null) continue;
                        cmap._map[(int)((long)low + j)] = DecodeUtf16BigEndian(item.Bytes);
                    }
                }
            }
        }

        /// <summary>
        /// 把 CMap 中的编码字节按大端拼成整型键，最多取前 4 字节。
        /// Assembles CMap encoding bytes into an integer key in big-endian order, using at most the first 4 bytes.
        /// </summary>
        /// <param name="bytes">编码字节。The encoding bytes.</param>
        /// <returns>用作映射键的整数编码值。The integer encoding value used as the map key.</returns>
        private static int ToCode(byte[] bytes)
        {
            int value = 0;
            for (int i = 0; i < bytes.Length && i < 4; i++)
            {
                value = (value << 8) | bytes[i];
            }
            return value;
        }

        /// <summary>
        /// 按 UTF-16 大端解码字节（存在 BOM 时自动跳过），用于还原 CMap 中的 Unicode 目标值。
        /// Decodes bytes as UTF-16 big-endian (skipping a leading BOM when present), used to recover the
        /// Unicode destinations in a CMap.
        /// </summary>
        /// <param name="bytes">待解码字节，null 或空时返回空字符串。The bytes to decode; null or empty input returns an empty string.</param>
        /// <returns>解码后的文本。The decoded text.</returns>
        private static string DecodeUtf16BigEndian(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0) return string.Empty;
            int start = bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF ? 2 : 0;
            int count = (bytes.Length - start) / 2;
            char[] chars = new char[count];
            for (int i = 0; i < count; i++)
            {
                chars[i] = (char)((bytes[start + i * 2] << 8) | bytes[start + i * 2 + 1]);
            }
            return new string(chars);
        }

        /// <summary>
        /// 尝试按编码值取得对应文本。
        /// Looks up the text mapped to an encoding value.
        /// </summary>
        /// <param name="code">编码值。The encoding value.</param>
        /// <param name="text">命中时为映射的 Unicode 文本。On a hit, the mapped Unicode text.</param>
        /// <returns>映射存在返回 true。True when the mapping exists.</returns>
        public bool TryGet(int code, out string text)
        {
            return _map.TryGetValue(code, out text);
        }

        /// <summary>已收录的映射条目数量。The number of stored mapping entries.</summary>
        public int Count { get { return _map.Count; } }
    }
}
