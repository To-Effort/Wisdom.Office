using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// PDF 流解码器：按 /Filter 链依次还原流内容。
    /// PDF stream decoder: restores stream content by applying the /Filter chain in order.
    /// </summary>
    /// <remarks>支持 Flate、ASCIIHex、ASCII85、RunLength、LZW 及它们的缩写别名；未知过滤器会中止后续解码并原样返回已解出的数据，Crypt 则直接跳过。所有解码路径都受 <see cref="PdfReadLimits.MaxDecodedStreamBytes"/> 约束。
    /// Supports Flate, ASCIIHex, ASCII85, RunLength and LZW plus their abbreviated aliases; an unknown
    /// filter stops further decoding and returns the data decoded so far, while Crypt is skipped. Every
    /// decode path is bounded by <see cref="PdfReadLimits.MaxDecodedStreamBytes"/>.</remarks>
    internal static class PdfStreamDecoder
    {
        /// <summary>
        /// 解码单个流对象，依次应用其过滤器链。
        /// Decodes a single stream object by applying its filter chain in order.
        /// </summary>
        /// <param name="stream">待解码的流对象，null 时返回空数组。The stream to decode; null returns an empty array.</param>
        /// <param name="resolve">间接对象解析委托，可为 null（此时引用不解引用）。The indirect-object resolver; may be null, in which case references are not resolved.</param>
        /// <returns>完全解码后的字节。The fully decoded bytes.</returns>
        /// <exception cref="DecompressionLimitException">解码后体积超过安全上限时抛出，用于拦截解压炸弹。Thrown when the decoded size exceeds the safety limit, guarding against decompression bombs.</exception>
        public static byte[] Decode(PdfStream stream, Func<PdfObject, PdfObject> resolve)
        {
            if (stream == null) return new byte[0];
            byte[] data = stream.Data;
            PdfObject filterObject = stream.Dictionary.Get("Filter");
            PdfObject parmsObject = stream.Dictionary.Get("DecodeParms");
            if (parmsObject == null) parmsObject = stream.Dictionary.Get("DP");

            List<PdfObject> filters = Normalize(filterObject, resolve);
            List<PdfObject> parms = Normalize(parmsObject, resolve);
            for (int i = 0; i < filters.Count; i++)
            {
                PdfObject parm = i < parms.Count ? parms[i] : null;
                string name = GetName(filters[i], resolve);
                if (string.IsNullOrEmpty(name)) continue;

                if (name == "FlateDecode" || name == "Fl")
                {
                    data = Zlib.Decompress(data, 0, data.Length, PdfReadLimits.MaxDecodedStreamBytes);
                    data = ApplyPredictor(data, parm, resolve);
                }
                else if (name == "ASCIIHexDecode" || name == "AHx")
                {
                    data = DecodeAsciiHex(data);
                }
                else if (name == "ASCII85Decode" || name == "A85")
                {
                    data = DecodeAscii85(data);
                }
                else if (name == "RunLengthDecode" || name == "RL")
                {
                    data = DecodeRunLength(data);
                }
                else if (name == "LZWDecode" || name == "LZW")
                {
                    int earlyChange = GetInt(parm, "EarlyChange", 1, resolve);
                    data = DecodeLzw(data, earlyChange);
                    data = ApplyPredictor(data, parm, resolve);
                }
                else if (name == "Crypt")
                {
                    continue;
                }
                else
                {
                    return data;
                }
                if (data.Length > PdfReadLimits.MaxDecodedStreamBytes)
                {
                    throw new DecompressionLimitException("PDF 流解压后超过安全上限(" + PdfReadLimits.MaxDecodedStreamBytes + " 字节)。");
                }
            }
            return data;
        }

        /// <summary>
        /// 把 /Filter、/DecodeParms 统一规整为列表，以便按下标配对；单个值视为长度为 1 的链。
        /// Normalizes /Filter and /DecodeParms into lists so entries pair up by index; a single value is
        /// treated as a chain of length 1.
        /// </summary>
        /// <param name="value">原始参数对象。The raw parameter object.</param>
        /// <param name="resolve">间接对象解析委托。The indirect-object resolver.</param>
        /// <returns>解析过引用的对象列表。The list of objects with references resolved.</returns>
        private static List<PdfObject> Normalize(PdfObject value, Func<PdfObject, PdfObject> resolve)
        {
            List<PdfObject> result = new List<PdfObject>();
            if (value == null) return result;
            value = Resolve(value, resolve);
            PdfArray array = value as PdfArray;
            if (array != null)
            {
                for (int i = 0; i < array.Count; i++)
                {
                    result.Add(Resolve(array[i], resolve));
                }
                return result;
            }
            result.Add(value);
            return result;
        }

        /// <summary>
        /// 若对象是引用且提供了解析委托，则取其实际目标对象。
        /// If the object is a reference and a resolver is provided, returns its actual target object.
        /// </summary>
        /// <param name="value">可能是引用的对象。The object, which may be a reference.</param>
        /// <param name="resolve">间接对象解析委托。The indirect-object resolver.</param>
        /// <returns>解引用后的对象，无法解引用时返回原值。The dereferenced object, or the original value when it cannot be resolved.</returns>
        private static PdfObject Resolve(PdfObject value, Func<PdfObject, PdfObject> resolve)
        {
            if (value is PdfReference && resolve != null) return resolve(value);
            return value;
        }

        /// <summary>
        /// 读取名称对象的正文，非名称返回 null。
        /// Reads the body of a name object; returns null for non-names.
        /// </summary>
        /// <param name="value">可能是名称或引用的对象。The object, which may be a name or a reference.</param>
        /// <param name="resolve">间接对象解析委托。The indirect-object resolver.</param>
        /// <returns>名称正文或 null。The name body, or null.</returns>
        private static string GetName(PdfObject value, Func<PdfObject, PdfObject> resolve)
        {
            value = Resolve(value, resolve);
            PdfName name = value as PdfName;
            return name == null ? null : name.Value;
        }

        /// <summary>
        /// 从参数字典中读取整数值。
        /// Reads an integer value from a parameter dictionary.
        /// </summary>
        /// <param name="dictionaryObject">字典对象或其引用。The dictionary object or a reference to it.</param>
        /// <param name="key">要读取的键。The key to read.</param>
        /// <param name="fallback">缺失或类型不符时返回的默认值。The default returned when the key is missing or has the wrong type.</param>
        /// <param name="resolve">间接对象解析委托。The indirect-object resolver.</param>
        /// <returns>读取到的整数或默认值。The integer read, or the default value.</returns>
        internal static int GetInt(PdfObject dictionaryObject, string key, int fallback, Func<PdfObject, PdfObject> resolve)
        {
            PdfDictionary dictionary = Resolve(dictionaryObject, resolve) as PdfDictionary;
            if (dictionary == null) return fallback;
            PdfObject value = Resolve(dictionary.Get(key), resolve);
            PdfNumber number = value as PdfNumber;
            return number == null ? fallback : (int)number.Value;
        }

        /// <summary>
        /// 对 Flate/LZW 解码结果应用 PNG 或 TIFF 预测器，还原按行差值编码的图像数据。
        /// Applies the PNG or TIFF predictor to Flate/LZW output, restoring image data encoded as row
        /// differences.
        /// </summary>
        /// <param name="data">已解压的原始字节。The already-inflated raw bytes.</param>
        /// <param name="parmsObject">解码参数字典（/Predictor、/Colors、/BitsPerComponent、/Columns）。The decode parameter dictionary (/Predictor, /Colors, /BitsPerComponent, /Columns).</param>
        /// <param name="resolve">间接对象解析委托。The indirect-object resolver.</param>
        /// <returns>去除预测后的字节；预测器为 1（无预测）或参数缺失时原样返回。The predictor-removed bytes; returned unchanged when the predictor is 1 (none) or parameters are missing.</returns>
        /// <remarks>预测器 2 为 TIFF 水平差值；10–15 为 PNG 逐行滤波，每行前缀一个滤波类型字节，输出长度按整行对齐会略短于输入。
        /// Predictor 2 is TIFF horizontal differencing; 10-15 are PNG per-row filters where each row is
        /// prefixed by a filter-type byte, so the row-aligned output is slightly shorter than the input.</remarks>
        private static byte[] ApplyPredictor(byte[] data, PdfObject parmsObject, Func<PdfObject, PdfObject> resolve)
        {
            PdfDictionary parms = Resolve(parmsObject, resolve) as PdfDictionary;
            if (parms == null) return data;
            int predictor = GetInt(parms, "Predictor", 1, resolve);
            if (predictor <= 1) return data;
            int colors = Math.Max(1, GetInt(parms, "Colors", 1, resolve));
            int bits = Math.Max(1, GetInt(parms, "BitsPerComponent", 8, resolve));
            int columns = Math.Max(1, GetInt(parms, "Columns", 1, resolve));
            // 行长按 long 计算并设上限：三个参数相乘会溢出成小正值或负值，而上一行缓冲区正是按行宽分配的。
            // The row length is computed in long arithmetic and capped: the product of the three parameters
            // can overflow into a small positive or negative value, and the previous-row buffer is sized
            // from that row length.
            long rowBits = (long)colors * bits * columns;
            if (rowBits <= 0 || rowBits > (long)PdfReadLimits.MaxDecodedStreamBytes * 8) return data;
            int rowLength = (int)((rowBits + 7) / 8);
            if (rowLength <= 0 || rowLength > data.Length) return data;
            int bytesPerPixel = (int)Math.Min(rowLength, Math.Max(1, ((long)colors * bits + 7) / 8));

            if (predictor == 2)
            {
                int rows = data.Length / rowLength;
                byte[] output = new byte[data.Length];
                Buffer.BlockCopy(data, 0, output, 0, data.Length);
                for (int row = 0; row < rows; row++)
                {
                    int offset = row * rowLength;
                    for (int i = bytesPerPixel; i < rowLength; i++)
                    {
                        output[offset + i] = (byte)(output[offset + i] + output[offset + i - bytesPerPixel]);
                    }
                }
                return output;
            }

            if (predictor < 10) return data;
            int stride = rowLength + 1;
            int rowCount = data.Length / stride;
            byte[] result = new byte[rowCount * rowLength];
            byte[] previous = new byte[rowLength];
            for (int row = 0; row < rowCount; row++)
            {
                int source = row * stride;
                int filter = data[source];
                int target = row * rowLength;
                for (int i = 0; i < rowLength; i++)
                {
                    int raw = data[source + 1 + i];
                    int left = i >= bytesPerPixel ? result[target + i - bytesPerPixel] : 0;
                    int up = previous[i];
                    int upLeft = i >= bytesPerPixel ? previous[i - bytesPerPixel] : 0;
                    int value;
                    switch (filter)
                    {
                        case 1: value = raw + left; break;
                        case 2: value = raw + up; break;
                        case 3: value = raw + ((left + up) >> 1); break;
                        case 4: value = raw + Paeth(left, up, upLeft); break;
                        default: value = raw; break;
                    }
                    result[target + i] = (byte)value;
                }
                Buffer.BlockCopy(result, target, previous, 0, rowLength);
            }
            return result;
        }

        /// <summary>
        /// 在追加数据前检查累计长度是否越界，防止畸形流把内存耗尽。
        /// Checks the accumulated length before appending data, so a malformed stream cannot exhaust memory.
        /// </summary>
        /// <param name="current">已解码的字节数。The number of bytes decoded so far.</param>
        /// <param name="added">本次准备追加的字节数。The number of bytes about to be appended.</param>
        /// <exception cref="DecompressionLimitException">超出 <see cref="PdfReadLimits.MaxDecodedStreamBytes"/> 时抛出。Thrown when the limit <see cref="PdfReadLimits.MaxDecodedStreamBytes"/> would be exceeded.</exception>
        private static void EnsureCapacity(int current, int added)
        {
            if ((long)current + added > PdfReadLimits.MaxDecodedStreamBytes)
            {
                throw new DecompressionLimitException("PDF 流解压后超过安全上限(" + PdfReadLimits.MaxDecodedStreamBytes + " 字节)。");
            }
        }

        /// <summary>
        /// PNG Paeth 预测函数：在左、上、左上三个邻居中选出最接近线性预测的那个。
        /// The PNG Paeth predictor: picks whichever of the left, up and up-left neighbors is closest to the
        /// linear prediction.
        /// </summary>
        /// <param name="left">左邻居字节。The left neighbor byte.</param>
        /// <param name="up">上方邻居字节。The byte above.</param>
        /// <param name="upLeft">左上邻居字节。The up-left neighbor byte.</param>
        /// <returns>选中的预测值。The selected predictor value.</returns>
        private static int Paeth(int left, int up, int upLeft)
        {
            int estimate = left + up - upLeft;
            int distanceLeft = Math.Abs(estimate - left);
            int distanceUp = Math.Abs(estimate - up);
            int distanceUpLeft = Math.Abs(estimate - upLeft);
            if (distanceLeft <= distanceUp && distanceLeft <= distanceUpLeft) return left;
            if (distanceUp <= distanceUpLeft) return up;
            return upLeft;
        }

        /// <summary>
        /// 执行 ASCIIHexDecode：跳过非十六进制字符，遇 '>' 结束，末尾落单的半个字节按补零处理。
        /// Runs ASCIIHexDecode: non-hex characters are skipped, '>' terminates, and a trailing half byte is
        /// zero-padded.
        /// </summary>
        /// <param name="data">编码字节。The encoded bytes.</param>
        /// <returns>解码后的字节。The decoded bytes.</returns>
        private static byte[] DecodeAsciiHex(byte[] data)
        {
            List<byte> result = new List<byte>(data.Length / 2);
            int high = -1;
            for (int i = 0; i < data.Length; i++)
            {
                byte value = data[i];
                if (value == (byte)'>') break;
                int digit = HexValue(value);
                if (digit < 0) continue;
                if (high < 0) high = digit;
                else
                {
                    result.Add((byte)((high << 4) | digit));
                    high = -1;
                }
            }
            if (high >= 0) result.Add((byte)(high << 4));
            return result.ToArray();
        }

        /// <summary>
        /// 执行 ASCII85Decode：每 5 个可打印字符还原 4 字节，'z' 表示整组零，尾部不足一组以 'u' 补齐后只输出有效字节。
        /// Runs ASCII85Decode: every 5 printable characters restore 4 bytes, 'z' stands for a full group of
        /// zeros, and a short tail is padded with 'u' so only the valid bytes are emitted.
        /// </summary>
        /// <param name="data">编码字节。The encoded bytes.</param>
        /// <returns>解码后的字节。The decoded bytes.</returns>
        private static byte[] DecodeAscii85(byte[] data)
        {
            List<byte> result = new List<byte>(data.Length);
            int[] group = new int[5];
            int count = 0;
            bool start = true;
            for (int i = 0; i < data.Length; i++)
            {
                byte value = data[i];
                if (PdfParser.IsWhitespace(value)) continue;
                if (start)
                {
                    start = false;
                    if (value == (byte)'<')
                    {
                        i++;
                        if (i < data.Length && data[i] == (byte)'~') continue;
                    }
                }
                if (value == (byte)'~') break;
                if (value == (byte)'z' && count == 0)
                {
                    result.Add(0); result.Add(0); result.Add(0); result.Add(0);
                    continue;
                }
                if (value < (byte)'!' || value > (byte)'u') continue;
                group[count++] = value - (byte)'!';
                if (count == 5)
                {
                    uint block = 0;
                    for (int g = 0; g < 5; g++) block = block * 85 + (uint)group[g];
                    result.Add((byte)(block >> 24));
                    result.Add((byte)(block >> 16));
                    result.Add((byte)(block >> 8));
                    result.Add((byte)block);
                    count = 0;
                }
            }
            if (count > 0)
            {
                int original = count;
                while (count < 5) group[count++] = 84;
                uint block = 0;
                for (int g = 0; g < 5; g++) block = block * 85 + (uint)group[g];
                byte[] bytes = { (byte)(block >> 24), (byte)(block >> 16), (byte)(block >> 8), (byte)block };
                for (int i = 0; i < original - 1; i++) result.Add(bytes[i]);
            }
            return result.ToArray();
        }

        /// <summary>
        /// 执行 RunLengthDecode：长度字节 0..127 表示其后拷贝 length+1 个字节，129..255 表示把下一个字节重复 257-length 次，128 为 EOD。
        /// Runs RunLengthDecode: length bytes 0..127 copy the following length+1 bytes, 129..255 repeat the
        /// next byte 257-length times, and 128 marks end of data.
        /// </summary>
        /// <param name="data">编码字节。The encoded bytes.</param>
        /// <returns>解码后的字节。The decoded bytes.</returns>
        private static byte[] DecodeRunLength(byte[] data)
        {
            List<byte> result = new List<byte>(data.Length * 2);
            int position = 0;
            while (position < data.Length)
            {
                int length = data[position++];
                if (length == 128) break;
                if (length < 128)
                {
                    int count = length + 1;
                    EnsureCapacity(result.Count, count);
                    for (int i = 0; i < count && position < data.Length; i++) result.Add(data[position++]);
                }
                else
                {
                    if (position >= data.Length) break;
                    byte value = data[position++];
                    int count = 257 - length;
                    EnsureCapacity(result.Count, count);
                    for (int i = 0; i < count; i++) result.Add(value);
                }
            }
            return result.ToArray();
        }

        /// <summary>
        /// 执行 LZWDecode：码表初始为 256 个单字节条目加 256/257 两个控制码，码宽随表规模从 9 位增长到 12 位。
        /// Runs LZWDecode: the code table starts with 256 single-byte entries plus the 256/257 control codes,
        /// and the code width grows from 9 to 12 bits as the table fills.
        /// </summary>
        /// <param name="data">编码字节。The encoded bytes.</param>
        /// <param name="earlyChange">EarlyChange 参数，1 表示在码字用完前一位就加宽码长（PDF 默认），0 表示延迟加宽。The EarlyChange parameter; 1 widens the code one bit before the current width is exhausted (the PDF default), 0 delays the widening.</param>
        /// <returns>解码后的字节。The decoded bytes.</returns>
        /// <remarks>遇到 256 清空码表重置码宽，257 结束；对「码字恰好在下一轮才生成」的 deferral 情形按 LZW 规则用前一条目自举补全。
        /// Code 256 clears the table and resets the code width; 257 ends the stream. The classic LZW
        /// deferral case, where a code is emitted just before it is added to the table, is completed by
        /// bootstrapping from the previous entry.</remarks>
        private static byte[] DecodeLzw(byte[] data, int earlyChange)
        {
            List<byte[]> table = new List<byte[]>(4096);
            for (int i = 0; i < 256; i++) table.Add(new byte[] { (byte)i });
            table.Add(null);
            table.Add(null);
            int codeSize = 9;
            int nextCode = 258;
            int bitPosition = 0;
            byte[] previous = null;
            List<byte> output = new List<byte>(data.Length * 2);

            while (true)
            {
                int code = ReadBits(data, ref bitPosition, codeSize);
                if (code < 0) break;
                if (code == 256)
                {
                    table.RemoveRange(258, table.Count - 258);
                    codeSize = 9;
                    nextCode = 258;
                    previous = null;
                    continue;
                }
                if (code == 257) break;

                byte[] entry;
                if (code < table.Count && table[code] != null)
                {
                    entry = table[code];
                }
                else if (code == nextCode && previous != null)
                {
                    entry = new byte[previous.Length + 1];
                    Buffer.BlockCopy(previous, 0, entry, 0, previous.Length);
                    entry[previous.Length] = previous[0];
                }
                else
                {
                    break;
                }

                EnsureCapacity(output.Count, entry.Length);
                output.AddRange(entry);
                if (previous != null && nextCode < 4096)
                {
                    byte[] newEntry = new byte[previous.Length + 1];
                    Buffer.BlockCopy(previous, 0, newEntry, 0, previous.Length);
                    newEntry[previous.Length] = entry[0];
                    if (nextCode == table.Count) table.Add(newEntry);
                    else if (nextCode < table.Count) table[nextCode] = newEntry;
                    else
                    {
                        while (table.Count < nextCode) table.Add(null);
                        table.Add(newEntry);
                    }
                    nextCode++;
                    if (nextCode + earlyChange >= (1 << codeSize) && codeSize < 12) codeSize++;
                }
                previous = entry;
            }
            return output.ToArray();
        }

        /// <summary>
        /// 从字节流的指定位偏移按 MSB 优先读取若干位，并把位偏移前推。
        /// Reads the given number of bits most-significant-bit first from a bit offset in the byte stream and
        /// advances the offset.
        /// </summary>
        /// <param name="data">字节流。The byte stream.</param>
        /// <param name="bitPosition">当前位偏移，返回时指向下一码起始位。The current bit offset; on return it points at the start of the next code.</param>
        /// <param name="count">要读取的位数。The number of bits to read.</param>
        /// <returns>读到的整数值；剩余位不足时返回 -1 且不推进偏移。The integer read; -1 without advancing the offset when fewer bits remain.</returns>
        private static int ReadBits(byte[] data, ref int bitPosition, int count)
        {
            if (bitPosition + count > data.Length * 8) return -1;
            int value = 0;
            for (int i = 0; i < count; i++)
            {
                int byteIndex = bitPosition >> 3;
                int bitIndex = 7 - (bitPosition & 7);
                value = (value << 1) | ((data[byteIndex] >> bitIndex) & 1);
                bitPosition++;
            }
            return value;
        }

        /// <summary>
        /// 取单个字符的十六进制位值。
        /// Returns the hex digit value of a single character.
        /// </summary>
        /// <param name="value">ASCII 字节。The ASCII byte.</param>
        /// <returns>0–15 的位值；不是十六进制字符时返回 -1。A value from 0 to 15; -1 when the character is not a hex digit.</returns>
        private static int HexValue(byte value)
        {
            if (value >= (byte)'0' && value <= (byte)'9') return value - (byte)'0';
            if (value >= (byte)'a' && value <= (byte)'f') return value - (byte)'a' + 10;
            if (value >= (byte)'A' && value <= (byte)'F') return value - (byte)'A' + 10;
            return -1;
        }
    }
}
