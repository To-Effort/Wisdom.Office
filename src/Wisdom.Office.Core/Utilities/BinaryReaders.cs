using System;
using System.IO;
using System.Text;

namespace Wisdom.Office.Core
{
    /// <summary>小端字节序的顺序读取器，基于内存字节数组并维护读写位置。A little-endian sequential reader over an in-memory byte array with a moving cursor.</summary>
    public sealed class LittleEndianReader
    {
        /// <summary>底层字节缓冲区。The underlying byte buffer.</summary>
        private readonly byte[] _data;
        /// <summary>当前读取位置（相对整个数组）。The current read position (relative to the whole array).</summary>
        private int _position;

        /// <summary>以整个数组构造读取器；data 为 null 时等价于空数组。Creates a reader over the whole array; null data behaves like an empty array.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        public LittleEndianReader(byte[] data)
            : this(data, 0, data == null ? 0 : data.Length)
        {
        }

        /// <summary>以指定起点与长度构造读取器。Creates a reader over a [offset, offset+length) window of the array. The window bounds later reads and the remaining-length calculation.</summary>
        /// <param name="data">源字节数组，不能为 null。The source byte array; must not be null.</param>
        /// <param name="offset">窗口起始偏移。The starting offset of the window.</param>
        /// <param name="length">窗口字节长度。The byte length of the window.</param>
        /// <exception cref="ArgumentNullException">data 为 null。data is null.</exception>
        public LittleEndianReader(byte[] data, int offset, int length)
        {
            if (data == null) throw new ArgumentNullException("data");
            _data = data;
            Offset = offset;
            End = offset + length;
            _position = offset;
        }

        /// <summary>窗口起始偏移。The starting offset of the readable window.</summary>
        public int Offset { get; private set; }
        /// <summary>窗口结束偏移（不含）。The exclusive end offset of the readable window.</summary>
        public int End { get; private set; }
        /// <summary>当前读取位置。The current read position.</summary>
        public int Position { get { return _position; } set { _position = value; } }
        /// <summary>窗口总长度（字节）。The total byte length of the window.</summary>
        public int Length { get { return End - Offset; } }
        /// <summary>当前位置到窗口末尾剩余的可读字节数。The number of readable bytes remaining from the current position to the window end.</summary>
        public int Remaining { get { return End - _position; } }
        /// <summary>底层字节缓冲区。The underlying byte buffer.</summary>
        public byte[] Data { get { return _data; } }

        /// <summary>判断从当前位置起是否还能读取指定数量的字节。Reports whether the given number of bytes can still be read from the current position.</summary>
        /// <param name="count">待读取的字节数。The number of bytes to read.</param>
        /// <returns>足够可读返回 true。True when enough bytes remain.</returns>
        public bool CanRead(int count)
        {
            return _position + count <= End;
        }

        /// <summary>读取一个字节并前移位置。Reads one byte and advances the position.</summary>
        /// <returns>读出的字节。The byte read.</returns>
        /// <exception cref="EndOfStreamException">已到窗口末尾。The window end was reached.</exception>
        public byte ReadByte()
        {
            if (_position >= End) throw new EndOfStreamException();
            return _data[_position++];
        }

        /// <summary>以小端序读取 2 字节无符号整数。Reads a 2-byte unsigned integer in little-endian order.</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 2 字节。Fewer than 2 bytes remain.</exception>
        public ushort ReadUInt16()
        {
            if (_position + 2 > End) throw new EndOfStreamException();
            ushort value = (ushort)(_data[_position] | (_data[_position + 1] << 8));
            _position += 2;
            return value;
        }

        /// <summary>以小端序读取 2 字节有符号整数。Reads a 2-byte signed integer in little-endian order.</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 2 字节。Fewer than 2 bytes remain.</exception>
        public short ReadInt16()
        {
            return unchecked((short)ReadUInt16());
        }

        /// <summary>以小端序读取 4 字节无符号整数。Reads a 4-byte unsigned integer in little-endian order.</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 4 字节。Fewer than 4 bytes remain.</exception>
        public uint ReadUInt32()
        {
            if (_position + 4 > End) throw new EndOfStreamException();
            uint value = (uint)(_data[_position]
                | (_data[_position + 1] << 8)
                | (_data[_position + 2] << 16)
                | (_data[_position + 3] << 24));
            _position += 4;
            return value;
        }

        /// <summary>以小端序读取 4 字节有符号整数。Reads a 4-byte signed integer in little-endian order.</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 4 字节。Fewer than 4 bytes remain.</exception>
        public int ReadInt32()
        {
            return unchecked((int)ReadUInt32());
        }

        /// <summary>以小端序读取 8 字节无符号整数（先低 4 字节后高 4 字节）。Reads an 8-byte unsigned integer in little-endian order (low word first, then high word).</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 8 字节。Fewer than 8 bytes remain.</exception>
        public ulong ReadUInt64()
        {
            ulong low = ReadUInt32();
            ulong high = ReadUInt32();
            return low | (high << 32);
        }

        /// <summary>读取指定数量的原始字节副本。Reads a copy of the given number of raw bytes.</summary>
        /// <param name="count">待读取的字节数。The number of bytes to read.</param>
        /// <returns>新分配的字节数组。A newly allocated byte array.</returns>
        /// <exception cref="EndOfStreamException">count 为负或剩余不足。count is negative or fewer bytes remain.</exception>
        public byte[] ReadBytes(int count)
        {
            if (count < 0 || _position + count > End) throw new EndOfStreamException();
            byte[] result = new byte[count];
            Buffer.BlockCopy(_data, _position, result, 0, count);
            _position += count;
            return result;
        }

        /// <summary>读取指定长度的 ASCII 字符串。Reads an ASCII string of the given length.</summary>
        /// <param name="count">字符/字节数。The number of characters (bytes).</param>
        /// <returns>解码后的字符串。The decoded string.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 count 字节。Fewer than count bytes remain.</exception>
        public string ReadAsciiString(int count)
        {
            byte[] bytes = ReadBytes(count);
            return Encoding.ASCII.GetString(bytes);
        }

        /// <summary>向前跳过指定字节数。Skips forward by the given number of bytes.</summary>
        /// <param name="count">待跳过的字节数。The number of bytes to skip.</param>
        /// <exception cref="EndOfStreamException">count 为负或超出窗口。count is negative or exceeds the window.</exception>
        public void Skip(int count)
        {
            if (count < 0 || _position + count > End) throw new EndOfStreamException();
            _position += count;
        }

        /// <summary>将读取位置直接设置为指定的绝对偏移。Sets the read position directly to the given absolute offset.</summary>
        /// <param name="position">新的绝对偏移。The new absolute offset.</param>
        public void Seek(int position)
        {
            _position = position;
        }
    }

    /// <summary>大端字节序的顺序读取器，基于内存字节数组并维护读写位置。A big-endian sequential reader over an in-memory byte array with a moving cursor.</summary>
    public sealed class BigEndianReader
    {
        /// <summary>底层字节缓冲区。The underlying byte buffer.</summary>
        private readonly byte[] _data;
        /// <summary>当前读取位置。The current read position.</summary>
        private int _position;
        /// <summary>可读窗口结束偏移（不含）。The exclusive end offset of the readable window.</summary>
        private readonly int _end;

        /// <summary>以整个数组构造读取器；data 为 null 时等价于空数组。Creates a reader over the whole array; null data behaves like an empty array.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        public BigEndianReader(byte[] data)
            : this(data, 0, data == null ? 0 : data.Length)
        {
        }

        /// <summary>以指定起点与长度构造读取器。Creates a reader over a [offset, offset+length) window of the array. The window bounds later reads and the remaining-length calculation.</summary>
        /// <param name="data">源字节数组，不能为 null。The source byte array; must not be null.</param>
        /// <param name="offset">窗口起始偏移。The starting offset of the window.</param>
        /// <param name="length">窗口字节长度。The byte length of the window.</param>
        /// <exception cref="ArgumentNullException">data 为 null。data is null.</exception>
        public BigEndianReader(byte[] data, int offset, int length)
        {
            if (data == null) throw new ArgumentNullException("data");
            _data = data;
            _position = offset;
            _end = offset + length;
        }

        /// <summary>当前读取位置。The current read position.</summary>
        public int Position { get { return _position; } set { _position = value; } }
        /// <summary>可读窗口结束偏移（不含）。The exclusive end offset of the readable window.</summary>
        public int End { get { return _end; } }
        /// <summary>当前位置到窗口末尾剩余的可读字节数。The number of readable bytes remaining from the current position to the window end.</summary>
        public int Remaining { get { return _end - _position; } }
        /// <summary>是否已到达数据末尾。Whether the end of the data has been reached.</summary>
        public bool EndOfData { get { return _position >= _end; } }

        /// <summary>读取一个字节并前移位置。Reads one byte and advances the position.</summary>
        /// <returns>读出的字节。The byte read.</returns>
        /// <exception cref="EndOfStreamException">已到窗口末尾。The window end was reached.</exception>
        public byte ReadByte()
        {
            if (_position >= _end) throw new EndOfStreamException();
            return _data[_position++];
        }

        /// <summary>以大端序读取 2 字节无符号整数。Reads a 2-byte unsigned integer in big-endian order.</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 2 字节。Fewer than 2 bytes remain.</exception>
        public ushort ReadUInt16()
        {
            if (_position + 2 > _end) throw new EndOfStreamException();
            ushort value = (ushort)((_data[_position] << 8) | _data[_position + 1]);
            _position += 2;
            return value;
        }

        /// <summary>以大端序读取 2 字节有符号整数。Reads a 2-byte signed integer in big-endian order.</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 2 字节。Fewer than 2 bytes remain.</exception>
        public short ReadInt16()
        {
            return unchecked((short)ReadUInt16());
        }

        /// <summary>以大端序读取 4 字节无符号整数。Reads a 4-byte unsigned integer in big-endian order.</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 4 字节。Fewer than 4 bytes remain.</exception>
        public uint ReadUInt32()
        {
            if (_position + 4 > _end) throw new EndOfStreamException();
            uint value = (uint)((_data[_position] << 24)
                | (_data[_position + 1] << 16)
                | (_data[_position + 2] << 8)
                | _data[_position + 3]);
            _position += 4;
            return value;
        }

        /// <summary>以大端序读取 4 字节有符号整数。Reads a 4-byte signed integer in big-endian order.</summary>
        /// <returns>读出的值。The value read.</returns>
        /// <exception cref="EndOfStreamException">剩余不足 4 字节。Fewer than 4 bytes remain.</exception>
        public int ReadInt32()
        {
            return unchecked((int)ReadUInt32());
        }

        /// <summary>读取指定数量的原始字节副本。Reads a copy of the given number of raw bytes.</summary>
        /// <param name="count">待读取的字节数。The number of bytes to read.</param>
        /// <returns>新分配的字节数组。A newly allocated byte array.</returns>
        /// <exception cref="EndOfStreamException">count 为负或剩余不足。count is negative or fewer bytes remain.</exception>
        public byte[] ReadBytes(int count)
        {
            if (count < 0 || _position + count > _end) throw new EndOfStreamException();
            byte[] result = new byte[count];
            Buffer.BlockCopy(_data, _position, result, 0, count);
            _position += count;
            return result;
        }

        /// <summary>向前跳过指定字节数。Skips forward by the given number of bytes.</summary>
        /// <param name="count">待跳过的字节数。The number of bytes to skip.</param>
        /// <exception cref="EndOfStreamException">count 为负或超出窗口。count is negative or exceeds the window.</exception>
        public void Skip(int count)
        {
            if (count < 0 || _position + count > _end) throw new EndOfStreamException();
            _position += count;
        }

        /// <summary>将读取位置直接设置为指定的绝对偏移。Sets the read position directly to the given absolute offset.</summary>
        /// <param name="position">新的绝对偏移。The new absolute offset.</param>
        public void Seek(int position)
        {
            _position = position;
        }
    }

    /// <summary>针对字节数组指定偏移的原地读写小工具。In-place helper utilities for reading and writing values at an offset within a byte array.</summary>
    public static class ByteUtils
    {
        /// <summary>从指定偏移以大端序读取 4 字节无符号整数。Reads a 4-byte unsigned integer in big-endian order at the given offset.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <returns>读出的值。The value read.</returns>
        public static uint ReadUInt32BigEndian(byte[] data, int offset)
        {
            return (uint)((data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]);
        }

        /// <summary>从指定偏移以大端序读取 2 字节无符号整数。Reads a 2-byte unsigned integer in big-endian order at the given offset.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <returns>读出的值。The value read.</returns>
        public static ushort ReadUInt16BigEndian(byte[] data, int offset)
        {
            return (ushort)((data[offset] << 8) | data[offset + 1]);
        }

        /// <summary>从指定偏移以小端序读取 2 字节无符号整数。Reads a 2-byte unsigned integer in little-endian order at the given offset.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <returns>读出的值。The value read.</returns>
        public static ushort ReadUInt16LittleEndian(byte[] data, int offset)
        {
            return (ushort)(data[offset] | (data[offset + 1] << 8));
        }

        /// <summary>从指定偏移以小端序读取 4 字节无符号整数。Reads a 4-byte unsigned integer in little-endian order at the given offset.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <returns>读出的值。The value read.</returns>
        public static uint ReadUInt32LittleEndian(byte[] data, int offset)
        {
            return (uint)(data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24));
        }

        /// <summary>从指定偏移以小端序读取 4 字节有符号整数。Reads a 4-byte signed integer in little-endian order at the given offset.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <returns>读出的值。The value read.</returns>
        public static int ReadInt32LittleEndian(byte[] data, int offset)
        {
            return unchecked((int)ReadUInt32LittleEndian(data, offset));
        }

        /// <summary>以大端序将 4 字节无符号整数写入指定偏移。Writes a 4-byte unsigned integer in big-endian order at the given offset.</summary>
        /// <param name="data">目标字节数组。The destination byte array.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <param name="value">待写入的值。The value to write.</param>
        public static void WriteUInt32BigEndian(byte[] data, int offset, uint value)
        {
            data[offset] = (byte)(value >> 24);
            data[offset + 1] = (byte)(value >> 16);
            data[offset + 2] = (byte)(value >> 8);
            data[offset + 3] = (byte)value;
        }

        /// <summary>以大端序将 2 字节无符号整数写入指定偏移。Writes a 2-byte unsigned integer in big-endian order at the given offset.</summary>
        /// <param name="data">目标字节数组。The destination byte array.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <param name="value">待写入的值。The value to write.</param>
        public static void WriteUInt16BigEndian(byte[] data, int offset, ushort value)
        {
            data[offset] = (byte)(value >> 8);
            data[offset + 1] = (byte)value;
        }

        /// <summary>将字节数组转换为小写十六进制字符串。Converts a byte array to a lowercase hexadecimal string.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        /// <returns>每字节两位的十六进制字符串。A two-digits-per-byte hexadecimal string.</returns>
        public static string ToHex(byte[] data)
        {
            StringBuilder builder = new StringBuilder(data.Length * 2);
            for (int i = 0; i < data.Length; i++)
            {
                builder.Append(data[i].ToString("x2"));
            }
            return builder.ToString();
        }
    }

    /// <summary>Adler-32 校验和算法实现，供 zlib 压缩尾使用。An Adler-32 checksum implementation used by the zlib trailer.</summary>
    public static class Adler32
    {
        /// <summary>计算指定区段的 Adler-32 校验和。Computes the Adler-32 checksum over the given range.</summary>
        /// <param name="data">源字节数组。The source byte array.</param>
        /// <param name="offset">起始偏移。The starting offset.</param>
        /// <param name="count">参与计算的字节数。The number of bytes to process.</param>
        /// <returns>32 位校验和。The 32-bit checksum.</returns>
        public static uint Compute(byte[] data, int offset, int count)
        {
            uint a = 1;
            uint b = 0;
            const uint mod = 65521;
            for (int i = 0; i < count; i++)
            {
                a = (a + data[offset + i]) % mod;
                b = (b + a) % mod;
            }
            return (b << 16) | a;
        }
    }
}
