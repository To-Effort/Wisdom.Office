using System;
using System.IO;
using System.IO.Compression;

namespace Wisdom.Office.Core
{
    /// <summary>解压后的数据超出安全上限时抛出的异常。Thrown when decompressed data would exceed the configured safety limit.</summary>
    public sealed class DecompressionLimitException : Exception
    {
        /// <summary>以指定消息创建解压上限异常。Creates a decompression-limit exception with the given message.</summary>
        /// <param name="message">描述超限情况的错误消息。The error message describing the limit breach.</param>
        public DecompressionLimitException(string message) : base(message)
        {
        }
    }

    /// <summary>基于 DeflateStream 的 RFC 1950 zlib 封装（零外部依赖）。An RFC 1950 zlib wrapper around DeflateStream (zero external dependencies).</summary>
    public static class Zlib
    {
        /// <summary>将数据压缩为带 2 字节 zlib 头与 Adler-32 校验尾的格式。Compresses data into the zlib format with a 2-byte header and an Adler-32 trailer.</summary>
        /// <param name="data">原始字节；null 视为空数组。The raw bytes; null is treated as an empty array.</param>
        /// <returns>zlib 封装后的字节数组。The zlib-wrapped compressed bytes.</returns>
        public static byte[] Compress(byte[] data)
        {
            if (data == null) data = new byte[0];
            using (MemoryStream output = new MemoryStream())
            {
                output.WriteByte(0x78);
                output.WriteByte(0x9C);
                using (DeflateStream deflate = new DeflateStream(output, CompressionLevel.Optimal, true))
                {
                    deflate.Write(data, 0, data.Length);
                }
                uint adler = Adler32.Compute(data, 0, data.Length);
                output.WriteByte((byte)(adler >> 24));
                output.WriteByte((byte)(adler >> 16));
                output.WriteByte((byte)(adler >> 8));
                output.WriteByte((byte)adler);
                return output.ToArray();
            }
        }

        /// <summary>解压整段 zlib 或原始 deflate 数据，不限制输出大小。Decompresses an entire zlib or raw-deflate buffer without an output size cap.</summary>
        /// <param name="data">压缩字节。The compressed bytes.</param>
        /// <returns>解压后的原始字节。The decompressed raw bytes.</returns>
        public static byte[] Decompress(byte[] data)
        {
            return Decompress(data, 0, data == null ? 0 : data.Length);
        }

        /// <summary>解压指定区段的 zlib 或原始 deflate 数据，不限制输出大小。Decompresses a zlib or raw-deflate range without an output size cap.</summary>
        /// <param name="data">压缩字节。The compressed bytes.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <param name="count">参与解压的字节数。The number of bytes to decompress.</param>
        /// <returns>解压后的原始字节。The decompressed raw bytes.</returns>
        public static byte[] Decompress(byte[] data, int offset, int count)
        {
            return Decompress(data, offset, count, long.MaxValue);
        }

        /// <summary>解压指定区段并限制输出上限；自动识别 zlib 头，失败时回退按原始 deflate 处理。Decompresses a range with a caller-supplied output cap. It sniffs the zlib header and, if inflating with the header fails, retries as raw deflate. A DecompressionLimitException is never swallowed.</summary>
        /// <param name="data">压缩字节。The compressed bytes.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <param name="count">参与解压的字节数。The number of bytes to decompress.</param>
        /// <param name="maxBytes">解压后允许的最大字节数。The maximum number of decompressed bytes allowed.</param>
        /// <returns>解压后的原始字节。The decompressed raw bytes.</returns>
        /// <exception cref="ArgumentNullException">data 为 null。data is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">offset 或 count 越界。offset or count is out of range.</exception>
        /// <exception cref="DecompressionLimitException">解压结果超过 maxBytes。The decompressed result exceeds maxBytes.</exception>
        public static byte[] Decompress(byte[] data, int offset, int count, long maxBytes)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (offset < 0 || count < 0 || offset + count > data.Length)
            {
                throw new ArgumentOutOfRangeException("offset");
            }

            bool zlibWrapped = count > 2
                && (data[offset] & 0x0F) == 8
                && (((data[offset] << 8) | data[offset + 1]) % 31) == 0;

            int skip = zlibWrapped ? 2 : 0;
            try
            {
                return Inflate(data, offset + skip, count - skip, maxBytes);
            }
            catch (DecompressionLimitException)
            {
                throw;
            }
            catch (InvalidDataException)
            {
                if (zlibWrapped)
                {
                    return Inflate(data, offset, count, maxBytes);
                }
                throw;
            }
            catch (EndOfStreamException)
            {
                if (zlibWrapped)
                {
                    return Inflate(data, offset, count, maxBytes);
                }
                throw;
            }
        }

        /// <summary>按原始 deflate 流解压指定区段，并在超出 maxBytes 时抛出上限异常。Inflates a raw deflate range through DeflateStream and throws a DecompressionLimitException as soon as the accumulated output would exceed maxBytes.</summary>
        /// <param name="data">压缩字节。The compressed bytes.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <param name="count">参与解压的字节数。The number of bytes to decompress.</param>
        /// <param name="maxBytes">解压后允许的最大字节数。The maximum number of decompressed bytes allowed.</param>
        /// <returns>解压后的原始字节。The decompressed raw bytes.</returns>
        private static byte[] Inflate(byte[] data, int offset, int count, long maxBytes)
        {
            using (MemoryStream input = new MemoryStream(data, offset, count, false))
            using (DeflateStream deflate = new DeflateStream(input, CompressionMode.Decompress))
            using (MemoryStream output = new MemoryStream())
            {
                byte[] buffer = new byte[16384];
                while (true)
                {
                    int read = deflate.Read(buffer, 0, buffer.Length);
                    if (read <= 0) break;
                    if (output.Length + read > maxBytes)
                    {
                        throw new DecompressionLimitException("解压后的数据超过安全上限(" + maxBytes + " 字节)。");
                    }
                    output.Write(buffer, 0, read);
                }
                return output.ToArray();
            }
        }
    }
}
