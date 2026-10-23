using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// 只读的微软复合文档二进制（OLE2/CFB）解析器。
    /// 自行实现 FAT、DIFAT、MiniFAT 与目录解析，不依赖任何外部库。
    /// A read-only parser for Microsoft compound binary documents (OLE2/CFB).
    /// It implements FAT, DIFAT, MiniFAT and directory parsing itself, with no external dependencies.
    /// </summary>
    public sealed class CompoundFile
    {
        /// <summary>FAT 项的链结束标记值。The end-of-chain marker value for a FAT entry.</summary>
        private const uint EndOfChain = 0xFFFFFFFE;
        /// <summary>合法常规扇区编号的最大值（介于其上为特殊保留值）。The maximum valid regular sector number (values above it are special reserved markers).</summary>
        private const uint MaxRegularSector = 0xFFFFFFFA;

        /// <summary>复合文件完整字节内容。The complete byte content of the compound file.</summary>
        private readonly byte[] _data;
        /// <summary>常规扇区大小（字节），由扇区偏移指数换算。The regular sector size in bytes, derived from the sector-shift value.</summary>
        private readonly int _sectorSize;
        /// <summary>迷你扇区大小（字节）。The mini-sector size in bytes.</summary>
        private readonly int _miniSectorSize;
        /// <summary>迷你流阈值：小于该字节数的流存放于迷你流中。The mini-stream cutoff: streams smaller than this byte count are stored in the mini stream.</summary>
        private readonly uint _miniStreamCutoff;
        /// <summary>由全部 FAT 扇区拼接而成的扇区分配表。The sector allocation table assembled from all FAT sectors.</summary>
        private readonly uint[] _fat;
        /// <summary>迷你扇区分配表。The mini-sector allocation table.</summary>
        private readonly uint[] _miniFat;
        /// <summary>解析出的目录项列表。The list of parsed directory entries.</summary>
        private readonly List<DirectoryEntry> _entries;
        /// <summary>根存储扇区链拼接出的迷你流数据。The mini-stream data assembled from the root storage's sector chain.</summary>
        private readonly byte[] _miniStream;

        /// <summary>目录项：描述文件内部的一个存储或流。A directory entry: describes one storage or stream inside the file.</summary>
        private sealed class DirectoryEntry
        {
            /// <summary>存储或流的名称（UTF-16）。The name of the storage or stream (UTF-16).</summary>
            public string Name;
            /// <summary>对象类型：1 存储、2 流、5 根存储，0 表示空槽。The object type: 1 storage, 2 stream, 5 root storage, and 0 for an empty slot.</summary>
            public byte ObjectType;
            /// <summary>起始扇区号（流数据或迷你流内迷你扇区）。The starting sector number (stream data, or a mini-sector within the mini stream).</summary>
            public uint StartSector;
            /// <summary>流的字节长度。The byte length of the stream.</summary>
            public long StreamSize;
            /// <summary>在目录扇区序列中的槽位索引。The slot index within the directory sector sequence.</summary>
            public int Index;
        }

        /// <summary>解析复合文件：读取头部、FAT/DIFAT/MiniFAT、目录树与迷你流。Parses a compound file: reads the header, FAT/DIFAT/MiniFAT, the directory tree and the mini stream.</summary>
        /// <param name="data">复合文件的完整字节内容。The complete byte content of the compound file.</param>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null。Is null.</exception>
        /// <exception cref="InvalidDataException">文件过小、签名不符或扇区大小非法。The file is too small, the signature does not match, or the sector size is invalid.</exception>
        /// <remarks>损坏的 DIFAT/FAT 扇区会被静默跳过，尽最大可能恢复可读内容。Corrupted DIFAT/FAT sectors are silently skipped so that readable content is recovered as much as possible.</remarks>
        public CompoundFile(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length < 512) throw new InvalidDataException("Compound file is too small.");
            if (!IsCompoundFile(data)) throw new InvalidDataException("Not a compound file (bad signature).");

            _data = data;
            int sectorShift = ByteUtils.ReadUInt16LittleEndian(data, 30);
            int miniSectorShift = ByteUtils.ReadUInt16LittleEndian(data, 32);
            if (sectorShift < 7 || sectorShift > 14)
            {
                throw new InvalidDataException("Invalid compound file sector size.");
            }
            _sectorSize = 1 << sectorShift;
            // 小扇区指数同样要验：uint16 直接当位移量会被 & 31 截断，得到一个错误的大小。
            // The mini-sector shift needs its own check: a raw uint16 used as a shift count is truncated by
            // '& 31', yielding a wrong size, and a value above the sector size breaks every mini-chain read.
            if (miniSectorShift > 14)
            {
                throw new InvalidDataException("Invalid compound file mini-sector size.");
            }
            _miniSectorSize = 1 << miniSectorShift;
            if (_miniSectorSize <= 0 || _miniSectorSize > _sectorSize)
            {
                throw new InvalidDataException("Invalid compound file mini-sector size.");
            }
            _miniStreamCutoff = ByteUtils.ReadUInt32LittleEndian(data, 56);

            uint numFatSectors = ByteUtils.ReadUInt32LittleEndian(data, 44);
            uint firstDirSector = ByteUtils.ReadUInt32LittleEndian(data, 48);
            uint firstMiniFatSector = ByteUtils.ReadUInt32LittleEndian(data, 60);
            uint numMiniFatSectors = ByteUtils.ReadUInt32LittleEndian(data, 64);
            uint firstDifatSector = ByteUtils.ReadUInt32LittleEndian(data, 68);
            uint numDifatSectors = ByteUtils.ReadUInt32LittleEndian(data, 72);

            List<uint> fatSectorIds = new List<uint>();
            for (int i = 0; i < 109; i++)
            {
                uint sector = ByteUtils.ReadUInt32LittleEndian(data, 76 + i * 4);
                if (sector <= MaxRegularSector)
                {
                    fatSectorIds.Add(sector);
                }
            }

            HashSet<uint> visitedDifat = new HashSet<uint>();
            uint difatSector = firstDifatSector;
            uint difatGuard = 0;
            while (difatSector <= MaxRegularSector && difatGuard < numDifatSectors + 16)
            {
                difatGuard++;
                if (!visitedDifat.Add(difatSector)) break;
                int sectorOffset = SectorOffset(difatSector);
                if (sectorOffset < 0 || sectorOffset + _sectorSize > _data.Length) break;
                int entriesPerSector = _sectorSize / 4;
                for (int i = 0; i < entriesPerSector - 1; i++)
                {
                    uint sector = ByteUtils.ReadUInt32LittleEndian(_data, sectorOffset + i * 4);
                    if (sector <= MaxRegularSector)
                    {
                        fatSectorIds.Add(sector);
                    }
                }
                difatSector = ByteUtils.ReadUInt32LittleEndian(_data, sectorOffset + (entriesPerSector - 1) * 4);
            }

            if (numFatSectors > 0 && fatSectorIds.Count > (int)numFatSectors)
            {
                fatSectorIds.RemoveRange((int)numFatSectors, fatSectorIds.Count - (int)numFatSectors);
            }

            List<uint> fat = new List<uint>();
            for (int i = 0; i < fatSectorIds.Count; i++)
            {
                int offset = SectorOffset(fatSectorIds[i]);
                if (offset < 0 || offset + _sectorSize > _data.Length) continue;
                int count = _sectorSize / 4;
                for (int j = 0; j < count; j++)
                {
                    fat.Add(ByteUtils.ReadUInt32LittleEndian(_data, offset + j * 4));
                }
            }
            _fat = fat.ToArray();

            byte[] directoryBytes = ReadSectorChain(firstDirSector, long.MaxValue);
            _entries = ParseDirectory(directoryBytes);
            DirectoryEntry root = null;
            for (int i = 0; i < _entries.Count; i++)
            {
                if (_entries[i].ObjectType == 5)
                {
                    root = _entries[i];
                    break;
                }
            }
            if (root != null)
            {
                _miniStream = ReadSectorChain(root.StartSector, root.StreamSize);
            }
            else
            {
                _miniStream = new byte[0];
            }

            List<uint> miniFat = new List<uint>();
            if (firstMiniFatSector <= MaxRegularSector && numMiniFatSectors > 0)
            {
                byte[] miniFatBytes = ReadSectorChain(firstMiniFatSector, (long)numMiniFatSectors * _sectorSize);
                for (int i = 0; i + 3 < miniFatBytes.Length; i += 4)
                {
                    miniFat.Add(ByteUtils.ReadUInt32LittleEndian(miniFatBytes, i));
                }
            }
            _miniFat = miniFat.ToArray();
        }

        /// <summary>依据 8 字节魔数（D0CF11E0A1B11AE1）判断字节数组是否为复合文件。Determines whether a byte array is a compound file by its 8-byte magic number (D0CF11E0A1B11AE1).</summary>
        /// <param name="data">待检测的字节数据。The byte data to test.</param>
        /// <returns>是否具备复合文件签名。Whether it carries the compound-file signature.</returns>
        public static bool IsCompoundFile(byte[] data)
        {
            return data != null && data.Length >= 8
                && data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0
                && data[4] == 0xA1 && data[5] == 0xB1 && data[6] == 0x1A && data[7] == 0xE1;
        }

        /// <summary>惰性枚举文件内全部流（对象类型为 2）的名称。Lazily enumerates the names of all streams in the file (object type 2).</summary>
        /// <returns>流名称序列。The sequence of stream names.</returns>
        public IEnumerable<string> GetStreamNames()
        {
            for (int i = 0; i < _entries.Count; i++)
            {
                if (_entries[i].ObjectType == 2)
                {
                    yield return _entries[i].Name;
                }
            }
        }

        /// <summary>判断是否存在指定名称的流（忽略大小写）。Reports whether a stream with the given name exists (ignoring case).</summary>
        /// <param name="name">流名称。The stream name.</param>
        /// <returns>是否存在。Whether it exists.</returns>
        public bool HasStream(string name)
        {
            return FindEntry(name) != null;
        }

        /// <summary>读取指定流的全部字节。Reads all bytes of the given stream.</summary>
        /// <param name="name">流名称（忽略大小写）。The stream name (ignoring case).</param>
        /// <returns>流内容字节数组。The byte array of the stream content.</returns>
        /// <exception cref="KeyNotFoundException">流不存在。The stream does not exist.</exception>
        public byte[] GetStream(string name)
        {
            DirectoryEntry entry = FindEntry(name);
            if (entry == null || entry.ObjectType != 2)
            {
                throw new KeyNotFoundException("Stream not found: " + name);
            }
            return ReadEntryStream(entry);
        }

        /// <summary>尝试读取指定流；不存在时返回 null 而非抛异常。Attempts to read the given stream, returning null instead of throwing when it is absent.</summary>
        /// <param name="name">流名称（忽略大小写）。The stream name (ignoring case).</param>
        /// <returns>流内容，或 null。The stream content, or null.</returns>
        public byte[] TryGetStream(string name)
        {
            DirectoryEntry entry = FindEntry(name);
            if (entry == null || entry.ObjectType != 2)
            {
                return null;
            }
            return ReadEntryStream(entry);
        }

        /// <summary>按名称忽略大小写查找流目录项。Finds a stream directory entry by name, ignoring case.</summary>
        /// <param name="name">流名称。The stream name.</param>
        /// <returns>找到的目录项；不存在返回 null。The found directory entry, or null when absent.</returns>
        private DirectoryEntry FindEntry(string name)
        {
            for (int i = 0; i < _entries.Count; i++)
            {
                if (_entries[i].ObjectType == 2 && string.Equals(_entries[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return _entries[i];
                }
            }
            return null;
        }

        /// <summary>读取目录项对应的流数据：小于迷你流阈值且存在迷你流时走迷你链，否则走常规扇区链。Reads the stream data for a directory entry: uses the mini chain when the size is below the cutoff and a mini stream exists, otherwise the regular sector chain.</summary>
        /// <param name="entry">流目录项。The stream directory entry.</param>
        /// <returns>流内容字节。The stream content bytes.</returns>
        private byte[] ReadEntryStream(DirectoryEntry entry)
        {
            if (entry.StreamSize == 0)
            {
                return new byte[0];
            }
            if (entry.StreamSize < _miniStreamCutoff && _miniStream.Length > 0)
            {
                return ReadMiniChain(entry.StartSector, entry.StreamSize);
            }
            return ReadSectorChain(entry.StartSector, entry.StreamSize);
        }

        /// <summary>将扇区号换算为文件内字节偏移（头部占一个扇区）；越界返回 -1。Converts a sector number to a byte offset within the file (the header occupies one sector); returns -1 when out of range.</summary>
        /// <param name="sector">扇区号。The sector number.</param>
        /// <returns>字节偏移或 -1。The byte offset, or -1.</returns>
        private int SectorOffset(uint sector)
        {
            long offset = ((long)sector + 1) * _sectorSize;
            if (offset < 0 || offset >= _data.Length) return -1;
            return (int)offset;
        }

        /// <summary>
        /// 沿 FAT 链读取常规扇区。size 为 long.MaxValue 时读到链尾（用于目录等不定长链）；
        /// 期望长度先被压到容器实际字节数，避免损坏的流长度触发巨量预分配；以访问集合防止环路，
        /// 实际读到的字节不足时截断返回。
        /// Reads regular sectors along the FAT chain. When size is long.MaxValue it reads to the end of the chain
        /// (for variable-length chains such as the directory); the expected length is first clamped to the bytes the
        /// container actually holds, so a corrupt stream length cannot request a huge allocation. A visited set
        /// prevents cycles and the result is truncated when fewer bytes are read.
        /// </summary>
        /// <param name="startSector">起始扇区号。The starting sector number.</param>
        /// <param name="size">期望读取的字节数。The expected number of bytes to read.</param>
        /// <returns>拼接后的扇区数据。The assembled sector data.</returns>
        private byte[] ReadSectorChain(uint startSector, long size)
        {
            if (size == long.MaxValue)
            {
                using (System.IO.MemoryStream memory = new System.IO.MemoryStream())
                {
                    uint sector = startSector;
                    HashSet<uint> visited = new HashSet<uint>();
                    while (sector <= MaxRegularSector && visited.Add(sector))
                    {
                        int offset = SectorOffset(sector);
                        if (offset < 0) break;
                        int length = Math.Min(_sectorSize, _data.Length - offset);
                        if (length <= 0) break;
                        memory.Write(_data, offset, length);
                        sector = GetNextSector(sector);
                    }
                    return memory.ToArray();
                }
            }

            if (size < 0) size = 0;
            if (size > _data.Length) size = _data.Length;
            int capacity = (int)size;
            byte[] result = new byte[capacity];
            int written = 0;
            uint current = startSector;
            HashSet<uint> seen = new HashSet<uint>();
            while (written < capacity && current <= MaxRegularSector && seen.Add(current))
            {
                int offset = SectorOffset(current);
                if (offset < 0) break;
                int length = Math.Min(_sectorSize, capacity - written);
                if (offset + length > _data.Length)
                {
                    length = _data.Length - offset;
                }
                if (length <= 0) break;
                Buffer.BlockCopy(_data, offset, result, written, length);
                written += length;
                current = GetNextSector(current);
            }
            if (written < result.Length)
            {
                Array.Resize(ref result, written);
            }
            return result;
        }

        /// <summary>
        /// 沿 MiniFAT 链从迷你流中读取小流数据；期望长度先压到迷你流的实际大小（负数按 0 处理），
        /// 含环路保护，不足时截断返回。
        /// Reads small-stream data from the mini stream along the MiniFAT chain. The expected length is first
        /// clamped to the real size of the mini stream (a negative length counts as zero); cycle protection is
        /// included and the result is truncated when fewer bytes are read.
        /// </summary>
        /// <param name="startSector">起始迷你扇区号。The starting mini-sector number.</param>
        /// <param name="size">期望读取的字节数。The expected number of bytes to read.</param>
        /// <returns>拼接后的迷你流数据。The assembled mini-stream data.</returns>
        private byte[] ReadMiniChain(uint startSector, long size)
        {
            if (size < 0) size = 0;
            if (size > _miniStream.Length) size = _miniStream.Length;
            int capacity = (int)size;
            byte[] result = new byte[capacity];
            int written = 0;
            uint current = startSector;
            HashSet<uint> seen = new HashSet<uint>();
            while (written < capacity && current < _miniFat.Length && seen.Add(current))
            {
                long offset = (long)current * _miniSectorSize;
                if (offset < 0 || offset >= _miniStream.Length) break;
                int length = Math.Min(_miniSectorSize, capacity - written);
                if (offset + length > _miniStream.Length)
                {
                    length = (int)(_miniStream.Length - offset);
                }
                if (length <= 0) break;
                Buffer.BlockCopy(_miniStream, (int)offset, result, written, length);
                written += length;
                current = _miniFat[current];
            }
            if (written < result.Length)
            {
                Array.Resize(ref result, written);
            }
            return result;
        }

        /// <summary>返回 FAT 链中该扇区的后继扇区号；扇区越界时返回链结束标记。Returns the successor sector number for a sector in the FAT chain; yields the end-of-chain marker when the sector is out of range.</summary>
        /// <param name="sector">当前扇区号。The current sector number.</param>
        /// <returns>后继扇区号或链结束标记。The next sector number, or the end-of-chain marker.</returns>
        private uint GetNextSector(uint sector)
        {
            if (sector >= _fat.Length) return EndOfChain;
            return _fat[sector];
        }

        /// <summary>把目录字节流解析为目录项列表：每 128 字节一条记录，跳过空槽（类型为 0）。Parses the directory byte stream into a list of entries: each 128-byte record becomes one entry, and empty slots (type 0) are skipped.</summary>
        /// <param name="data">目录字节内容。The directory byte content.</param>
        /// <returns>解析出的目录项列表。The parsed list of directory entries.</returns>
        private List<DirectoryEntry> ParseDirectory(byte[] data)
        {
            List<DirectoryEntry> entries = new List<DirectoryEntry>();
            int count = data.Length / 128;
            for (int i = 0; i < count; i++)
            {
                int offset = i * 128;
                string name = ReadDirectoryName(data, offset);
                byte type = data[offset + 66];
                if (type == 0)
                {
                    continue;
                }
                uint startSector = ByteUtils.ReadUInt32LittleEndian(data, offset + 116);
                long streamSize = (long)ByteUtils.ReadUInt32LittleEndian(data, offset + 120)
                    | ((long)ByteUtils.ReadUInt32LittleEndian(data, offset + 124) << 32);
                DirectoryEntry entry = new DirectoryEntry();
                entry.Name = name;
                entry.ObjectType = type;
                entry.StartSector = startSector;
                entry.StreamSize = streamSize;
                entry.Index = i;
                entries.Add(entry);
            }
            return entries;
        }

        /// <summary>读取目录项名称：偏移 64 处为 UTF-16 字节长度（含结尾空字符），最多取 31 个字符。Reads a directory entry's name: the field at offset 64 is the UTF-16 byte length (including the terminating null), so at most 31 characters are decoded.</summary>
        /// <param name="data">目录字节内容。The directory byte content.</param>
        /// <param name="offset">该目录项的起始偏移。The starting offset of the directory entry.</param>
        /// <returns>解码出的名称，长度非法时为空字符串。The decoded name, or an empty string when the length is invalid.</returns>
        private static string ReadDirectoryName(byte[] data, int offset)
        {
            int length = ByteUtils.ReadUInt16LittleEndian(data, offset + 64);
            if (length < 2) return string.Empty;
            int charCount = Math.Min(32, length / 2 - 1);
            if (charCount < 0) charCount = 0;
            return Encoding.Unicode.GetString(data, offset, charCount * 2);
        }
    }

}
