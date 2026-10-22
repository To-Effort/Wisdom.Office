using System;
using System.IO;
using System.IO.Compression;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// OOXML 压缩包（.docx/.xlsx）只读入口。.NET Framework 的 ZipArchive 在惰性枚举条目时，
    /// 会把非法条目名当作 <see cref="ArgumentException"/> 抛出，而不是在打开包时报错；
    /// 这里立即物化中央目录并统一转成 InvalidDataException，让同一份损坏文件在
    /// .NET Framework 与 .NET 上抛出相同异常类型。
    /// Read-only entry point for OOXML ZIP containers. The .NET Framework ZipArchive defers central
    /// directory parsing, so an illegal entry name surfaces later as ArgumentException instead of a
    /// corruption error at open time. This class materialises the entries immediately and translates
    /// the failure into InvalidDataException, giving identical behaviour on .NET Framework and .NET.
    /// </summary>
    public static class PackageArchive
    {
        /// <summary>以只读方式打开压缩包并立即校验中央目录。Opens the container for reading and validates the central directory up front.</summary>
        /// <param name="stream">压缩包流。The container stream.</param>
        /// <param name="leaveOpen">读取结束后是否保留流开启。Whether to leave the stream open afterwards.</param>
        /// <returns>只读压缩包实例，调用方负责释放。A read-only archive the caller must dispose.</returns>
        public static ZipArchive OpenForRead(Stream stream, bool leaveOpen)
        {
            ZipArchive archive = null;
            try
            {
                archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen);
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    // 仅为触发 .NET Framework 的中央目录解析，不消费条目。Triggers central directory parsing only; the entries are not consumed here.
                }
                return archive;
            }
            catch (ArgumentException ex)
            {
                if (archive != null) archive.Dispose();
                throw new InvalidDataException("The package contains an invalid entry name.", ex);
            }
            catch (NotSupportedException ex)
            {
                if (archive != null) archive.Dispose();
                throw new InvalidDataException("The package uses an unsupported compression method.", ex);
            }
        }
    }
}
