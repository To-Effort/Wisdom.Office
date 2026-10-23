using System;
using System.IO;
using System.IO.Compression;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// 针对 OOXML 包中解压炸弹的进程级防护。这些上限约束的是包内部件解压后的大小，
    /// 而非磁盘上文件本身的大小。
    /// Process-wide guards against decompression bombs in OOXML packages. The limits apply to the
    /// decompressed size of package parts, not to the size of the file on disk.
    /// </summary>
    public static class OfficePackageLimits
    {
        /// <summary>单个部件解压后的最大字节数（默认 256 MB）。The maximum decompressed size of a single part (256 MB by default).</summary>
        public static long MaxEntryBytes = 256L * 1024 * 1024;
        /// <summary>整个包累计解压后的最大字节数（默认 1 GB）。The maximum cumulative decompressed size of the whole package (1 GB by default).</summary>
        public static long MaxTotalBytes = 1024L * 1024 * 1024;
        /// <summary>单个部件名的最大字符数（默认 1024）。The maximum length of one part name in characters (1024 by default).</summary>
        public const int MaxEntryNameChars = 1024;

        /// <summary>读取单个 ZIP 部件的全部字节并施加安全上限校验。Reads all bytes of a single ZIP entry while enforcing the safety limits.</summary>
        /// <param name="entry">待读取的 ZIP 部件。The ZIP entry to read.</param>
        /// <returns>部件解压后的字节数组；entry 为 null 时返回 null。The decompressed bytes of the entry; null when entry is null.</returns>
        /// <exception cref="InvalidDataException">解压大小超过单件或总量安全上限。The decompressed size exceeds the per-entry or total safety limit.</exception>
        public static byte[] ReadEntryBytes(ZipArchiveEntry entry)
        {
            long total = 0;
            return ReadEntryBytes(entry, ref total);
        }

        /// <summary>读取单个 ZIP 部件的全部字节，并把本次解压量累加到跨部件的总量计数器上。Reads all bytes of a single ZIP entry and adds the decompressed amount to a caller-supplied cumulative counter, so the total limit can be enforced across every part of one package.</summary>
        /// <param name="entry">待读取的 ZIP 部件。The ZIP entry to read.</param>
        /// <param name="totalBytes">累计已解压字节数，调用返回时加上本部件的量。The running total of decompressed bytes; incremented by this entry's size on return.</param>
        /// <returns>部件解压后的字节数组；entry 为 null 时返回 null。The decompressed bytes of the entry; null when entry is null.</returns>
        /// <exception cref="InvalidDataException">解压大小超过单件或总量安全上限。The decompressed size exceeds the per-entry or total safety limit.</exception>
        public static byte[] ReadEntryBytes(ZipArchiveEntry entry, ref long totalBytes)
        {
            if (entry == null) return null;
            long remaining = MaxTotalBytes - totalBytes;
            if (remaining <= 0)
            {
                throw new InvalidDataException("文档解压后的总大小超过安全上限(" + MaxTotalBytes + " 字节)。");
            }
            long limit = Math.Min(MaxEntryBytes, remaining);
            if (entry.Length > limit)
            {
                throw new InvalidDataException("文档部件解压后超过安全上限: " + entry.FullName);
            }

            using (Stream stream = entry.Open())
            {
                // 中央目录已声明解压尺寸时按其精确分配一次到位，省掉增长缓冲与 ToArray 的整份拷贝。
                // When the central directory declares the decompressed size, allocate exactly once and
                // skip the growing buffer plus a full ToArray copy.
                long declared = entry.Length;
                if (declared > 0 && declared <= int.MaxValue)
                {
                    byte[] exact = new byte[(int)declared];
                    int filled = ReadFully(stream, exact, 0, limit);
                    totalBytes += filled;
                    if (filled == exact.Length) return exact;
                    byte[] trimmed = new byte[filled];
                    Array.Copy(exact, trimmed, filled);
                    return trimmed;
                }

                using (MemoryStream memory = new MemoryStream())
                {
                    byte[] buffer = new byte[65536];
                    long total = 0;
                    while (true)
                    {
                        int read = stream.Read(buffer, 0, buffer.Length);
                        if (read <= 0) break;
                        total += read;
                        if (total > limit)
                        {
                            throw new InvalidDataException("文档部件解压后超过安全上限: " + entry.FullName);
                        }
                        memory.Write(buffer, 0, read);
                    }
                    totalBytes += total;
                    return memory.ToArray();
                }
            }
        }

        /// <summary>
        /// 在不读取内容的情况下校验某部件是否会突破单件与总量上限，返回该部件允许的解压字节数。
        /// 供流式读取（不整体物化部件）的调用方在打开部件前调用，读取过程需自行按返回值继续限量。
        /// Validates a part against the per-entry and total limits without reading it and returns the byte
        /// allowance for that part. Streaming callers (which never materialize the whole part) call this before
        /// opening the entry and must keep enforcing the returned limit while reading.
        /// </summary>
        /// <param name="entry">待校验的 ZIP 部件。The ZIP entry to validate.</param>
        /// <param name="usedBytes">本包已累计解压的字节数。The cumulative decompressed bytes already used by this package.</param>
        /// <returns>该部件允许的解压字节上限。The decompressed byte allowance for this entry.</returns>
        /// <exception cref="InvalidDataException">累计总量或单件上限已被突破。The total or per-entry limit is already exceeded.</exception>
        public static long CheckEntry(ZipArchiveEntry entry, long usedBytes)
        {
            if (entry == null) return 0;
            long remaining = MaxTotalBytes - usedBytes;
            if (remaining <= 0)
            {
                throw new InvalidDataException("文档解压后的总大小超过安全上限(" + MaxTotalBytes + " 字节)。");
            }
            long limit = Math.Min(MaxEntryBytes, remaining);
            if (entry.Length > limit)
            {
                throw new InvalidDataException("文档部件解压后超过安全上限: " + entry.FullName);
            }
            return limit;
        }

        /// <summary>
        /// 校验即将写入包中的部件名：拒绝空名、绝对路径、"." 与 ".." 片段、反斜杠、盘符冒号、
        /// 通配与控制字符等解包时会造成路径穿越或非法文件名的写法。
        /// Validates a part name before it is written into the package, rejecting empty names, absolute paths,
        /// "." and ".." segments, backslashes, drive colons, wildcards and control characters that would turn
        /// into path traversal or illegal file names when the package is unpacked.
        /// </summary>
        /// <param name="name">待写入的部件名。The part name about to be written.</param>
        /// <returns>校验通过的部件名本身。The same name once it is validated.</returns>
        /// <exception cref="System.ArgumentException">部件名为空、过长或包含非法片段与字符。The name is empty, too long, or contains illegal segments or characters.</exception>
        public static string ValidateEntryName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                throw new ArgumentException("文档部件名不能为空。", "name");
            }
            if (name.Length > MaxEntryNameChars)
            {
                throw new ArgumentException("文档部件名超过长度上限(" + MaxEntryNameChars + " 字符)。", "name");
            }
            if (name[0] == '/')
            {
                throw new ArgumentException("文档部件名不能以 / 开头: " + name, "name");
            }
            for (int i = 0; i < name.Length; i++)
            {
                char c = name[i];
                if (c < ' ' || c == '\u007F' || c == '\\' || c == ':' || c == '*' || c == '?'
                    || c == '"' || c == '<' || c == '>' || c == '|')
                {
                    throw new ArgumentException("文档部件名包含非法字符: " + name, "name");
                }
            }
            string[] segments = name.Split('/');
            for (int i = 0; i < segments.Length; i++)
            {
                if (segments[i] == "." || segments[i] == "..")
                {
                    throw new ArgumentException("文档部件名不能包含 . 或 .. 片段: " + name, "name");
                }
            }
            return name;
        }

        /// <summary>把流读满目标数组的指定起点，超过上限即抛错。Fills the target array from the given offset, throwing once the limit is exceeded.</summary>
        /// <param name="stream">部件流。The entry stream.</param>
        /// <param name="target">目标数组。The target array.</param>
        /// <param name="offset">写入起点。Write offset.</param>
        /// <param name="limit">单件字节上限。Per-entry byte limit.</param>
        /// <returns>实际读到的字节数。The number of bytes actually read.</returns>
        private static int ReadFully(Stream stream, byte[] target, int offset, long limit)
        {
            int filled = offset;
            while (filled < target.Length)
            {
                int read = stream.Read(target, filled, target.Length - filled);
                if (read <= 0) break;
                filled += read;
                if (filled > limit)
                {
                    throw new InvalidDataException("文档部件解压后超过安全上限。");
                }
            }
            return filled - offset;
        }

        /// <summary>
        /// 在 <see cref="MaxEntryBytes"/> 上限内把流整体读入内存。
        /// Reads a whole stream into memory within the <see cref="MaxEntryBytes"/> limit.
        /// </summary>
        /// <param name="source">待读取的流。The stream to read.</param>
        /// <returns>读到的字节数组。The bytes read.</returns>
        /// <exception cref="InvalidDataException">流的实际长度超过单件上限。The stream is longer than the per-entry limit.</exception>
        /// <remarks>读取端不能凭 <see cref="Stream.Length"/> 决定分配：管道流与网络流长度未知，声明值也可能造假。
        /// The length reported by a stream cannot decide the allocation: pipes and network streams have no
        /// known length and a declared one may be forged, so the copy is charged byte by byte.</remarks>
        public static byte[] ReadAllBytes(Stream source)
        {
            if (source == null) throw new ArgumentNullException("source");
            using (MemoryStream memory = new MemoryStream())
            {
                CopyBounded(source, memory);
                return memory.ToArray();
            }
        }

        /// <summary>
        /// 把流按块复制到目标，累计字节数超过 <see cref="MaxEntryBytes"/> 即中止。
        /// Copies a stream to a target in chunks, aborting once the running total passes <see cref="MaxEntryBytes"/>.
        /// </summary>
        /// <param name="source">数据源。The source stream.</param>
        /// <param name="target">写入目标。The destination stream.</param>
        /// <returns>复制的字节数。The number of bytes copied.</returns>
        /// <exception cref="InvalidDataException">数据超过单件上限。The data exceeds the per-entry limit.</exception>
        public static long CopyBounded(Stream source, Stream target)
        {
            if (source == null) throw new ArgumentNullException("source");
            if (target == null) throw new ArgumentNullException("target");
            byte[] buffer = new byte[81920];
            long total = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                total += read;
                if (total > MaxEntryBytes)
                {
                    throw new InvalidDataException("输入流超过安全上限(" + MaxEntryBytes + " 字节)。");
                }
                target.Write(buffer, 0, read);
            }
            return total;
        }
    }
}
