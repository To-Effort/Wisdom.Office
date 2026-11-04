using System;
using System.IO;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// 目标文件的原子写入：先把全部字节写进同目录的临时文件，成功后再替换目标。
    /// 直接对目标用 <c>FileMode.Create</c> 是「先截断再写」，中途任何异常（含参数校验失败）
    /// 都会把用户原文件变成 0 字节或半截包；这里把失败限制在临时文件上，目标要么保持旧内容，
    /// 要么是写完整的新内容。
    /// Atomic write to a target file: every byte goes into a temporary file in the same directory and
    /// the target is replaced only once the whole write succeeded. Opening the target with
    /// <c>FileMode.Create</c> truncates it first, so any mid-write failure (a failed argument check
    /// included) leaves the user with a zero-byte or half-written file; here a failure stays confined
    /// to the temporary file, and the target keeps either its old bytes or the complete new ones.
    /// </summary>
    public static class AtomicFileWriter
    {
        /// <summary>
        /// 用给定的写入委托原子地生成目标文件。
        /// Produces the target file atomically from the supplied write delegate.
        /// </summary>
        /// <param name="path">目标文件路径。The target file path.</param>
        /// <param name="writer">接收临时文件流的写入委托。The delegate that writes into the temporary stream.</param>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="path"/> 或 <paramref name="writer"/> 为 null。Thrown when either is null.
        /// </exception>
        public static void Write(string path, Action<Stream> writer)
        {
            if (path == null) throw new ArgumentNullException("path");
            if (writer == null) throw new ArgumentNullException("writer");

            string target = Path.GetFullPath(path);
            string directory = Path.GetDirectoryName(target);
            if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
            // 临时文件必须与目标同目录：跨卷时 File.Replace / File.Move 不是原子操作。
            // The temporary file must share the target's directory; across volumes File.Replace and
            // File.Move are no longer atomic.
            string temporary = Path.Combine(directory,
                "." + Path.GetFileName(target) + "." + Guid.NewGuid().ToString("N") + ".tmp");

            bool committed = false;
            try
            {
                using (FileStream stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    writer(stream);
                    stream.Flush();
                }

                if (File.Exists(target))
                {
                    // 同卷替换：成功即原子换上新内容，旧文件随之删除（不保留备份）。
                    // Same-volume replace: on success the new bytes are swapped in atomically and the old
                    // file is removed (no backup copy is kept).
                    File.Replace(temporary, target, null);
                }
                else
                {
                    File.Move(temporary, target);
                }
                committed = true;
            }
            finally
            {
                if (!committed && File.Exists(temporary))
                {
                    try { File.Delete(temporary); }
                    catch (IOException) { }
                    catch (UnauthorizedAccessException) { }
                }
            }
        }
    }
}
