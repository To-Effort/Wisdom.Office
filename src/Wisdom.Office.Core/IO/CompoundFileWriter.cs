using System;
using System.Collections.Generic;
using System.IO;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// 微软复合文档二进制（OLE2/CFB）写出器，与 <see cref="CompoundFile"/> 互为读写对。
    /// 采用 512 字节常规扇区与 64 字节迷你扇区：不小于 4096 字节的流直接占用常规扇区链，
    /// 更小的流集中放入根存储的迷你流，并配套写出 FAT、DIFAT 与 MiniFAT。
    /// A writer for Microsoft compound binary documents (OLE2/CFB), the counterpart of <see cref="CompoundFile"/>.
    /// It uses 512-byte regular sectors and 64-byte mini sectors: streams of at least 4096 bytes occupy chains of
    /// regular sectors, smaller ones are packed into the root storage's mini stream, with FAT, DIFAT and MiniFAT
    /// written accordingly.
    /// </summary>
    public sealed class CompoundFileWriter
    {
        /// <summary>常规扇区字节数。The number of bytes in a regular sector.</summary>
        private const int SectorSize = 512;
        /// <summary>迷你扇区字节数。The number of bytes in a mini sector.</summary>
        private const int MiniSectorSize = 64;
        /// <summary>迷你流阈值：小于该字节数的流放入迷你流。The mini-stream cutoff: streams below this byte count go into the mini stream.</summary>
        private const int MiniStreamCutoff = 4096;
        /// <summary>单条目录项字节数。The number of bytes in one directory entry.</summary>
        private const int DirectoryEntrySize = 128;
        /// <summary>头部 DIFAT 数组的槽位数。The number of DIFAT slots kept in the header.</summary>
        private const int HeaderDifatSlots = 109;
        /// <summary>每个扇区容纳的 FAT 项数。The number of FAT entries one sector holds.</summary>
        private const int FatEntriesPerSector = SectorSize / 4;
        /// <summary>每个 DIFAT 扇区容纳的 FAT 扇区号个数（末位留给下一链指针）。The number of FAT sector ids per DIFAT sector (the last slot holds the next-chain pointer).</summary>
        private const int DifatEntriesPerSector = FatEntriesPerSector - 1;
        /// <summary>每个目录扇区容纳的目录项数。The number of directory entries one sector holds.</summary>
        private const int EntriesPerSector = SectorSize / DirectoryEntrySize;
        /// <summary>目录项起始扇区字段偏移。The offset of a directory entry's starting-sector field.</summary>
        private const int EntryStartSector = 116;
        /// <summary>目录项流长度字段偏移。The offset of a directory entry's stream-size field.</summary>
        private const int EntryStreamSize = 120;
        /// <summary>链结束标记。The end-of-chain marker.</summary>
        private const uint EndOfChain = 0xFFFFFFFE;
        /// <summary>空闲扇区标记。The free-sector marker.</summary>
        private const uint FreeSector = 0xFFFFFFFF;
        /// <summary>无兄弟/无子节点标记。The no-sibling and no-child marker.</summary>
        private const uint NoStream = 0xFFFFFFFF;
        /// <summary>FAT 扇区自身在 FAT 中的取值。The FAT value describing a FAT sector itself.</summary>
        private const uint FatSectorMarker = 0xFFFFFFFD;
        /// <summary>DIFAT 扇区自身在 FAT 中的取值。The FAT value describing a DIFAT sector itself.</summary>
        private const uint DifatSectorMarker = 0xFFFFFFFC;
        /// <summary>流名称允许的最大字符数（目录项 64 字节 UTF-16 含结尾空字符）。The maximum stream-name length in characters (the 64-byte UTF-16 name field includes the terminating null).</summary>
        private const int MaxNameChars = 31;
        /// <summary>可建成二叉树的目录槽位上限（兄弟指针只有 16 位）。The directory-slot limit for building the tree (the sibling pointers are 16 bits).</summary>
        private const int MaxTreeSlots = 65535;

        /// <summary>待写入的一条流：名称与完整内容。One stream to be written: its name and full content.</summary>
        private sealed class StreamItem
        {
            /// <summary>流名称。The stream name.</summary>
            public readonly string Name;
            /// <summary>流内容。The stream content.</summary>
            public readonly byte[] Data;

            /// <summary>创建流项。Creates a stream item.</summary>
            /// <param name="name">流名称。The stream name.</param>
            /// <param name="data">流内容。The stream content.</param>
            public StreamItem(string name, byte[] data)
            {
                Name = name;
                Data = data;
            }
        }

        /// <summary>按添加顺序排列的流列表。The stream list in insertion order.</summary>
        private readonly List<StreamItem> _streams = new List<StreamItem>();

        /// <summary>添加一条流；重名（忽略大小写）或名称超长时抛异常。Adds one stream, throwing when the name is a duplicate (ignoring case) or too long.</summary>
        /// <param name="name">流名称，如 Workbook。The stream name, for example Workbook.</param>
        /// <param name="data">流内容；null 视为空流。The stream content; null counts as an empty stream.</param>
        /// <exception cref="ArgumentException">名称为空、超过 31 字符或与已有流重名。The name is empty, longer than 31 characters, or already used.</exception>
        public void AddStream(string name, byte[] data)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException("复合文档流名不能为空。", "name");
            if (name.Length > MaxNameChars) throw new ArgumentException("复合文档流名超过 " + MaxNameChars + " 字符: " + name, "name");
            for (int i = 0; i < _streams.Count; i++)
            {
                if (string.Equals(_streams[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("复合文档流名重复: " + name, "name");
                }
            }
            _streams.Add(new StreamItem(name, data ?? new byte[0]));
        }

        /// <summary>商向上取整；被除数不为正时返回 0。Divides and rounds up; returns 0 when the dividend is not positive.</summary>
        /// <param name="value">被除数。The dividend.</param>
        /// <param name="divisor">除数。The divisor.</param>
        /// <returns>向上取整的商。The rounded-up quotient.</returns>
        private static int CeilingDivide(int value, int divisor)
        {
            if (value <= 0) return 0;
            return (value + divisor - 1) / divisor;
        }

        /// <summary>
        /// 生成复合文档字节：先按阈值分流并统计各类扇区数量，再用定点迭代求出 FAT 与 DIFAT 扇区数，
        /// 然后按 FAT、DIFAT、目录、MiniFAT、常规数据、迷你流的顺序连续分配扇区号并回填链路。
        /// Produces the compound-file bytes: it splits the streams by the cutoff, counts the sectors of each kind,
        /// solves the FAT and DIFAT sector counts by fixed-point iteration, then assigns consecutive sector numbers in
        /// the order FAT, DIFAT, directory, MiniFAT, regular data, mini stream, and fills in the chains.
        /// </summary>
        /// <returns>完整的复合文档字节。The complete compound-file bytes.</returns>
        /// <exception cref="InvalidOperationException">未添加任何流。No stream was added.</exception>
        public byte[] ToBytes()
        {
            if (_streams.Count == 0) throw new InvalidOperationException("复合文档至少需要一条流。");
            int count = _streams.Count;

            // 小于阈值的流进入迷你流，各自按迷你扇区向上取整。
            // Streams below the cutoff go into the mini stream, each rounded up to a whole mini sector.
            bool[] isMini = new bool[count];
            int[] miniSectorStart = new int[count];
            int[] miniSectorCount = new int[count];
            int[] sectorCount = new int[count];
            int totalMiniSectors = 0;
            for (int i = 0; i < count; i++)
            {
                byte[] data = _streams[i].Data;
                if (data.Length < MiniStreamCutoff)
                {
                    isMini[i] = true;
                    miniSectorStart[i] = totalMiniSectors;
                    miniSectorCount[i] = CeilingDivide(data.Length, MiniSectorSize);
                    totalMiniSectors += miniSectorCount[i];
                }
                else
                {
                    sectorCount[i] = CeilingDivide(data.Length, SectorSize);
                }
            }
            int miniStreamBytes = totalMiniSectors * MiniSectorSize;

            int directorySectors = CeilingDivide(count + 1, EntriesPerSector);
            int miniFatSectors = totalMiniSectors == 0 ? 0 : CeilingDivide(totalMiniSectors, FatEntriesPerSector);
            int dataSectors = 0;
            for (int i = 0; i < count; i++)
            {
                if (!isMini[i]) dataSectors += sectorCount[i];
            }
            int miniStreamSectors = CeilingDivide(miniStreamBytes, SectorSize);

            // FAT 要覆盖含自身与 DIFAT 在内的全部扇区，而 DIFAT 数量又取决于 FAT 数量，故定点迭代。
            // The FAT must cover every sector including the FAT and DIFAT sectors themselves, while the DIFAT count
            // depends on the FAT count, so the pair is solved by fixed-point iteration.
            int contentSectors = directorySectors + miniFatSectors + dataSectors + miniStreamSectors;
            int fatSectors = 0;
            int difatSectors = 0;
            while (true)
            {
                int neededFat = CeilingDivide(contentSectors + fatSectors + difatSectors, FatEntriesPerSector);
                int neededDifat = CeilingDivide(neededFat - HeaderDifatSlots, DifatEntriesPerSector);
                if (neededFat == fatSectors && neededDifat == difatSectors) break;
                fatSectors = neededFat;
                difatSectors = neededDifat;
            }

            int next = 0;
            int fatStart = next; next += fatSectors;
            int difatStart = next; next += difatSectors;
            int directoryStart = next; next += directorySectors;
            int miniFatStart = next; next += miniFatSectors;
            int dataStart = next; next += dataSectors;
            int miniStreamStart = next; next += miniStreamSectors;

            uint[] fat = new uint[fatSectors * FatEntriesPerSector];
            Fill(fat, FreeSector);
            LinkChain(fat, fatStart, fatSectors, FatSectorMarker);
            LinkChain(fat, difatStart, difatSectors, DifatSectorMarker);
            LinkChain(fat, directoryStart, directorySectors, EndOfChain);
            LinkChain(fat, miniFatStart, miniFatSectors, EndOfChain);
            LinkChain(fat, miniStreamStart, miniStreamSectors, EndOfChain);

            uint[] starts = new uint[count];
            int cursor = dataStart;
            for (int i = 0; i < count; i++)
            {
                if (isMini[i])
                {
                    starts[i] = miniSectorCount[i] == 0 ? NoStream : (uint)miniSectorStart[i];
                    continue;
                }
                starts[i] = sectorCount[i] == 0 ? NoStream : (uint)cursor;
                LinkChain(fat, cursor, sectorCount[i], EndOfChain);
                cursor += sectorCount[i];
            }

            uint[] miniFat = new uint[miniFatSectors * FatEntriesPerSector];
            Fill(miniFat, FreeSector);
            for (int i = 0; i < count; i++)
            {
                if (isMini[i]) LinkChain(miniFat, miniSectorStart[i], miniSectorCount[i], EndOfChain);
            }

            byte[] body = new byte[next * SectorSize];
            WriteTable(body, fat, fatStart);
            WriteTable(body, miniFat, miniFatStart);

            byte[] header = new byte[SectorSize];
            WriteHeader(header, fatSectors, directoryStart, miniFatSectors,
                miniFatSectors == 0 ? EndOfChain : (uint)miniFatStart,
                difatSectors == 0 ? EndOfChain : (uint)difatStart, difatSectors);
            WriteDifat(header, body, fatStart, fatSectors, difatStart, difatSectors);

            // 流数据：常规流按各自连续的扇区区间顺序拷贝，迷你流先合并再整体落位。
            // Stream data: regular streams copy into their own contiguous sector range; mini streams are merged first
            // and then written as one block.
            byte[] miniStream = miniStreamBytes == 0 ? new byte[0] : new byte[miniStreamBytes];
            for (int i = 0; i < count; i++)
            {
                byte[] data = _streams[i].Data;
                if (data.Length == 0) continue;
                if (isMini[i])
                {
                    Buffer.BlockCopy(data, 0, miniStream, miniSectorStart[i] * MiniSectorSize, data.Length);
                }
                else
                {
                    Buffer.BlockCopy(data, 0, body, (int)starts[i] * SectorSize, data.Length);
                }
            }
            if (miniStream.Length > 0)
            {
                Buffer.BlockCopy(miniStream, 0, body, miniStreamStart * SectorSize, miniStream.Length);
            }

            WriteDirectory(body, directoryStart, starts, rootStart: miniStreamSectors == 0 ? EndOfChain : (uint)miniStreamStart,
                rootSize: (uint)miniStreamBytes);

            using (MemoryStream output = new MemoryStream(SectorSize + body.Length))
            {
                output.Write(header, 0, header.Length);
                output.Write(body, 0, body.Length);
                return output.ToArray();
            }
        }

        /// <summary>填充数组全部元素。Fills every element of the array.</summary>
        /// <param name="values">目标数组。The target array.</param>
        /// <param name="value">填充值。The value to fill.</param>
        private static void Fill(uint[] values, uint value)
        {
            for (int i = 0; i < values.Length; i++) values[i] = value;
        }

        /// <summary>把 FAT 或 MiniFAT 数组按小端序写入从某扇区起的连续空间。Writes a FAT or MiniFAT array in little-endian order into the space starting at a given sector.</summary>
        /// <param name="body">扇区缓冲区。The sector buffer.</param>
        /// <param name="table">表数组。The table array.</param>
        /// <param name="start">首扇区号。The first sector number.</param>
        private static void WriteTable(byte[] body, uint[] table, int start)
        {
            int baseOffset = start * SectorSize;
            for (int i = 0; i < table.Length; i++)
            {
                WriteUInt32(body, baseOffset + i * 4, table[i]);
            }
        }

        /// <summary>填充头部的签名、版本、扇区大小与各表首扇区/数量字段。Fills the header's signature, version, sector sizes and the first-sector and count fields of every table.</summary>
        /// <param name="header">头部缓冲区（512 字节）。The header buffer (512 bytes).</param>
        /// <param name="fatSectors">FAT 扇区数。The number of FAT sectors.</param>
        /// <param name="directoryStart">目录首扇区号。The first directory sector number.</param>
        /// <param name="miniFatSectors">MiniFAT 扇区数。The number of MiniFAT sectors.</param>
        /// <param name="miniFatStart">MiniFAT 首扇区号。The first MiniFAT sector number.</param>
        /// <param name="difatStart">DIFAT 首扇区号。The first DIFAT sector number.</param>
        /// <param name="difatSectors">DIFAT 扇区数。The number of DIFAT sectors.</param>
        private static void WriteHeader(byte[] header, int fatSectors, int directoryStart, int miniFatSectors,
            uint miniFatStart, uint difatStart, int difatSectors)
        {
            header[0] = 0xD0; header[1] = 0xCF; header[2] = 0x11; header[3] = 0xE0;
            header[4] = 0xA1; header[5] = 0xB1; header[6] = 0x1A; header[7] = 0xE1;
            WriteUInt16(header, 24, 0x003E);            // 小版本 minor version
            WriteUInt16(header, 26, 0x0003);            // 大版本 major version (V4)
            WriteUInt16(header, 28, 0xFFFE);            // 字节序 little-endian byte order
            WriteUInt16(header, 30, 9);                 // 扇区大小指数 sector-size shift
            WriteUInt16(header, 32, 6);                 // 迷你扇区指数 mini-sector shift
            WriteUInt32(header, 44, (uint)fatSectors);
            WriteUInt32(header, 48, (uint)directoryStart);
            WriteUInt32(header, 56, MiniStreamCutoff);
            WriteUInt32(header, 60, miniFatStart);
            WriteUInt32(header, 64, (uint)miniFatSectors);
            WriteUInt32(header, 68, difatStart);
            WriteUInt32(header, 72, (uint)difatSectors);
        }

        /// <summary>写出 FAT 扇区号清单：头部放前 109 项，其余按每扇区 127 项放入 DIFAT 扇区，末位为下一 DIFAT 扇区号。Writes the list of FAT sector ids: the first 109 live in the header, the rest go into DIFAT sectors at 127 ids each whose last slot holds the next DIFAT sector number.</summary>
        /// <param name="header">头部缓冲区。The header buffer.</param>
        /// <param name="body">扇区缓冲区。The sector buffer.</param>
        /// <param name="fatStart">FAT 首扇区号。The first FAT sector number.</param>
        /// <param name="fatSectors">FAT 扇区数。The number of FAT sectors.</param>
        /// <param name="difatStart">DIFAT 首扇区号。The first DIFAT sector number.</param>
        /// <param name="difatSectors">DIFAT 扇区数。The number of DIFAT sectors.</param>
        private static void WriteDifat(byte[] header, byte[] body, int fatStart, int fatSectors,
            int difatStart, int difatSectors)
        {
            for (int i = 0; i < HeaderDifatSlots; i++)
            {
                WriteUInt32(header, 76 + i * 4, i < fatSectors ? (uint)(fatStart + i) : FreeSector);
            }
            for (int k = 0; k < difatSectors; k++)
            {
                int sector = difatStart + k;
                int offset = sector * SectorSize;
                int first = HeaderDifatSlots + k * DifatEntriesPerSector;
                for (int i = 0; i < DifatEntriesPerSector; i++)
                {
                    int index = first + i;
                    WriteUInt32(body, offset + i * 4, index < fatSectors ? (uint)(fatStart + index) : FreeSector);
                }
                WriteUInt32(body, offset + DifatEntriesPerSector * 4,
                    k + 1 < difatSectors ? (uint)(sector + 1) : EndOfChain);
            }
        }

        /// <summary>
        /// 写出目录扇区：0 号槽为根存储并持有迷你流，其余槽先按名称排序，再取中点建成平衡二叉树，
        /// 使按任意顺序遍历树的读取端都能取到全部子项。
        /// Writes the directory sectors: slot 0 is the root storage which owns the mini stream, the remaining slots are
        /// sorted by name and then assembled around their midpoint into a balanced binary tree, so a reader that walks
        /// the tree reaches every entry.
        /// </summary>
        /// <param name="body">扇区缓冲区。The sector buffer.</param>
        /// <param name="directoryStart">目录首扇区号。The first directory sector number.</param>
        /// <param name="starts">各流的起始扇区号。The starting sector number of each stream.</param>
        /// <param name="rootStart">根存储起始扇区号（无迷你流时为链结束标记）。The root storage's starting sector (the end-of-chain marker when there is no mini stream).</param>
        /// <param name="rootSize">迷你流字节数。The mini-stream byte count.</param>
        private void WriteDirectory(byte[] body, int directoryStart, uint[] starts, uint rootStart, uint rootSize)
        {
            int baseOffset = directoryStart * SectorSize;
            List<int> order = new List<int>();
            for (int i = 0; i < _streams.Count; i++) order.Add(i);
            order.Sort(delegate(int a, int b)
            {
                int byName = string.Compare(_streams[a].Name, _streams[b].Name, StringComparison.OrdinalIgnoreCase);
                return byName != 0 ? byName : a.CompareTo(b);
            });

            WriteName(body, baseOffset, "Root Entry");
            WriteUInt16(body, baseOffset + 64, (ushort)(("Root Entry".Length + 1) * 2));
            body[baseOffset + 66] = 5;
            body[baseOffset + 67] = 1;
            WriteUInt32(body, baseOffset + EntryStartSector, rootStart);
            WriteUInt32(body, baseOffset + EntryStreamSize, rootSize);

            for (int position = 0; position < order.Count; position++)
            {
                int streamIndex = order[position];
                int offset = baseOffset + (position + 1) * DirectoryEntrySize;
                StreamItem item = _streams[streamIndex];
                WriteName(body, offset, item.Name);
                WriteUInt16(body, offset + 64, (ushort)((item.Name.Length + 1) * 2));
                body[offset + 66] = 2;
                body[offset + 67] = 1;
                WriteUInt32(body, offset + EntryStartSector, starts[streamIndex]);
                WriteUInt32(body, offset + EntryStreamSize, (uint)item.Data.Length);
                WriteUInt32(body, offset + 72, NoStream);
            }

            uint child = order.Count <= MaxTreeSlots
                ? LinkSubtree(body, baseOffset, 0, order.Count - 1)
                : NoStream;
            WriteUInt32(body, baseOffset + 68, NoStream);
            WriteUInt32(body, baseOffset + 72, child);
        }

        /// <summary>递归取区间中点为子树根并回填左右兄弟指针，返回该子树根的槽位号。Recursively takes the range midpoint as the sub-tree root, fills its sibling pointers and returns the root's slot number.</summary>
        /// <param name="body">扇区缓冲区。The sector buffer.</param>
        /// <param name="baseOffset">目录首字节偏移。The first directory byte offset.</param>
        /// <param name="left">区间左端。The inclusive left end of the range.</param>
        /// <param name="right">区间右端。The inclusive right end of the range.</param>
        /// <returns>子树根槽位号；区间为空返回无子节点标记。The sub-tree root slot, or the no-child marker when the range is empty.</returns>
        private static uint LinkSubtree(byte[] body, int baseOffset, int left, int right)
        {
            if (left > right) return NoStream;
            int middle = (left + right) / 2;
            int offset = baseOffset + (middle + 1) * DirectoryEntrySize;
            uint leftSibling = LinkSubtree(body, baseOffset, left, middle - 1);
            uint rightSibling = LinkSubtree(body, baseOffset, middle + 1, right);
            WriteUInt16(body, offset + 68, unchecked((ushort)leftSibling));
            WriteUInt16(body, offset + 70, unchecked((ushort)rightSibling));
            return (uint)(middle + 1);
        }

        /// <summary>写入目录项名称（UTF-16，最多 31 字符）。Writes a directory entry name (UTF-16, at most 31 characters).</summary>
        /// <param name="body">扇区缓冲区。The sector buffer.</param>
        /// <param name="offset">目录项起始偏移。The directory entry's starting offset.</param>
        /// <param name="name">名称。The name.</param>
        private static void WriteName(byte[] body, int offset, string name)
        {
            int length = Math.Min(name.Length, MaxNameChars);
            for (int i = 0; i < length; i++)
            {
                WriteUInt16(body, offset + i * 2, (ushort)name[i]);
            }
            WriteUInt16(body, offset + length * 2, 0);
        }

        /// <summary>把一段连续扇区接成链；末位写指定的结束值。Links a run of consecutive sectors, writing the given value at the last entry.</summary>
        /// <param name="table">FAT 或 MiniFAT 数组。The FAT or MiniFAT array.</param>
        /// <param name="start">首扇区号。The first sector number.</param>
        /// <param name="length">扇区个数。The number of sectors.</param>
        /// <param name="last">最后一个扇区的取值。The value stored for the last sector.</param>
        private static void LinkChain(uint[] table, int start, int length, uint last)
        {
            for (int i = 0; i < length; i++)
            {
                table[start + i] = i == length - 1 ? last : (uint)(start + i + 1);
            }
        }

        /// <summary>在小端序下写入 16 位值。Writes a 16-bit value in little-endian order.</summary>
        /// <param name="data">目标缓冲区。The target buffer.</param>
        /// <param name="offset">写入偏移。The write offset.</param>
        /// <param name="value">待写入的值。The value to write.</param>
        private static void WriteUInt16(byte[] data, int offset, ushort value)
        {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
        }

        /// <summary>在小端序下写入 32 位值。Writes a 32-bit value in little-endian order.</summary>
        /// <param name="data">目标缓冲区。The target buffer.</param>
        /// <param name="offset">写入偏移。The write offset.</param>
        /// <param name="value">待写入的值。The value to write.</param>
        private static void WriteUInt32(byte[] data, int offset, uint value)
        {
            data[offset] = (byte)value;
            data[offset + 1] = (byte)(value >> 8);
            data[offset + 2] = (byte)(value >> 16);
            data[offset + 3] = (byte)(value >> 24);
        }
    }
}
