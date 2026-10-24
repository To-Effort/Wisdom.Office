using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wisdom.Office.Core;
using Wisdom.Office.Core.Formulas;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 资源上限与解析边界的回归测试：伪造的 OLE2 FAT 扇区数、公式里声明得极大的区域、
    /// 恰好顶到缓冲区末尾的 JPEG 帧段，都只能被拒绝或按上限截断，不能把进程拖进异常或内存耗尽。
    /// Regression tests for resource ceilings and parsing boundaries: a forged OLE2 FAT sector count, a
    /// formula that declares an enormous range, and a JPEG frame segment that ends exactly at the buffer end
    /// must all be refused or clamped instead of dragging the process into an unexpected failure or an
    /// out-of-memory state.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>执行资源上限测试。Runs the resource-limit tests.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunResourceLimitTest(string outputDirectory)
        {
            RunForgedFatSectorCount(outputDirectory);
            RunFormulaDeclaredRangeCaps();
            RunJpegSizeAtBufferEnd();
            Console.WriteLine("[limits] forged sizes and edges ok");
        }

        /// <summary>
        /// 头部声明 0xFFFFFFFF 个 FAT 扇区：无符号声明值直接强转 int 会变负数，
        /// RemoveRange 随即抛 ArgumentOutOfRangeException，畸形文件应归为损坏而不是参数异常。
        /// A header claiming 0xFFFFFFFF FAT sectors: casting that unsigned declaration straight to int yields a
        /// negative number, RemoveRange then throws ArgumentOutOfRangeException, and a malformed file must read
        /// as corrupt rather than as a caller-argument failure.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunForgedFatSectorCount(string outputDirectory)
        {
            CompoundFileWriter writer = new CompoundFileWriter();
            writer.AddStream("Workbook", BuildPattern(20000, 41));
            byte[] bytes = writer.ToBytes();

            byte[] forged = (byte[])bytes.Clone();
            for (int i = 44; i < 48; i++) forged[i] = 0xFF;
            Exception failure = null;
            int streams = -1;
            try
            {
                streams = CountStreamNames(new CompoundFile(forged));
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            Require(!(failure is ArgumentOutOfRangeException),
                "forged FAT count leaked an argument exception instead of a read failure: " + failure);
            Require(failure == null || failure is InvalidDataException,
                "forged FAT count failed with " + (failure == null ? "none" : failure.GetType().Name + ": " + failure.Message));
            Require(streams <= 1, "forged FAT count still listed " + streams + " streams");

            byte[] clamped = (byte[])bytes.Clone();
            clamped[44] = 1;
            clamped[45] = 0;
            clamped[46] = 0;
            clamped[47] = 0;
            Exception clampedFailure = null;
            try
            {
                new CompoundFile(clamped);
            }
            catch (Exception ex)
            {
                clampedFailure = ex;
            }
            Require(!(clampedFailure is ArgumentOutOfRangeException),
                "understated FAT count leaked an argument exception: " + clampedFailure);

            Console.WriteLine("[limits] forged ole2 fat count handled");
        }

        /// <summary>数出流名的个数。Counts the stream names.</summary>
        /// <param name="file">已打开的 OLE2 文件。The opened OLE2 file.</param>
        /// <returns>流名个数。The number of stream names.</returns>
        private static int CountStreamNames(CompoundFile file)
        {
            int count = 0;
            foreach (string name in file.GetStreamNames()) count++;
            return count;
        }

        /// <summary>
        /// OFFSET 与 REPT 都只校验边界或次数、不校验产出规模：区域声明 100 万行、文本重复声明
        /// 三万多次乘长文本，都会把取值次数与分配量推到进程极限。
        /// OFFSET and REPT checked only their edges or their repeat count, never the produced size: a range of
        /// a million rows, or thirty thousand repeats of a long string, both push the cell-read count and the
        /// allocation up to the limit of the process.
        /// </summary>
        private static void RunFormulaDeclaredRangeCaps()
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("S1");
            sheet.GetCell(1, 1).Value = 7d;

            object offset = FormulaEngine.Evaluate(workbook, "OFFSET(A1,0,0,1000000,1000)");
            Require(ErrText(offset) == "#REF!",
                "OFFSET with an enormous declared range should be #REF! but was " + offset);

            string seed = new string('x', 2000);
            sheet.GetCell(2, 1).Value = seed;
            object rept = FormulaEngine.Evaluate(workbook, "REPT(A2,32767)");
            Require(ErrText(rept) == "#VALUE!", "REPT over 32767 characters should be #VALUE! but was " + ErrText(rept));

            object legalText = FormulaEngine.Evaluate(workbook, "LEN(REPT(\"ab\",100))");
            Require(AsDouble(legalText) == 200d, "legal REPT length should be 200 but was " + legalText);

            Console.WriteLine("[limits] formula declared sizes capped");
        }

        /// <summary>取错误文本；非错误时返回 null。Returns the error text, or null when the value is not an error.</summary>
        /// <param name="value">求值结果。The evaluated value.</param>
        /// <returns>形如 "#REF!" 的错误文本。Error text such as "#REF!".</returns>
        private static string ErrText(object value)
        {
            string text = value as string;
            if (text == null) return null;
            return text.StartsWith("#", StringComparison.Ordinal) ? text : null;
        }

        /// <summary>把求值结果按 double 取出；不是数值时返回 -1。Reads the result as a double; -1 when it is not numeric.</summary>
        /// <param name="value">求值结果。The evaluated value.</param>
        /// <returns>数值或 -1。The number or -1.</returns>
        private static double AsDouble(object value)
        {
            if (value is double) return (double)value;
            if (value is int) return (int)value;
            if (value is long) return (long)value;
            return -1d;
        }

        /// <summary>
        /// SOFn 帧段的边界读出：一段后面还有数据的合法帧必须交出尺寸，缓冲区正好停在帧头最后一个
        /// 必需字节上时属于截断，只能返回 false，不能读越界。
        /// The edges of an SOFn frame segment: a frame with data still behind it must yield its size, and a
        /// buffer that stops exactly at the last byte the header needs is truncated — it must return false
        /// instead of reading past the end.
        /// </summary>
        private static void RunJpegSizeAtBufferEnd()
        {
            byte[] jpeg = new byte[]
            {
                0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x00, 0x64, 0x00, 0x78, 0xFF, 0xDA
            };
            int width;
            int height;
            Require(OfficeImage.TryReadJpegSize(jpeg, out width, out height),
                "a frame segment followed by scan data must still yield its size");
            Require(width == 120 && height == 100, "jpeg size should be 120x100 but was " + width + "x" + height);

            // 截到宽度低字节为止：帧头声明的长度已在缓冲区外，尺寸不可信。
            // Cut at the low width byte: the length the header declares lies outside the buffer, so the size
            // is not trustworthy.
            byte[] truncated = new byte[jpeg.Length - 3];
            Buffer.BlockCopy(jpeg, 0, truncated, 0, truncated.Length);
            int badWidth = -1;
            int badHeight = -1;
            Require(!OfficeImage.TryReadJpegSize(truncated, out badWidth, out badHeight),
                "a frame segment cut short of its declared length must not report a size");
            Require(badWidth == 0 && badHeight == 0,
                "a failed size read left " + badWidth + "x" + badHeight + " behind");

            Console.WriteLine("[limits] jpeg edge handled");
        }
    }
}
