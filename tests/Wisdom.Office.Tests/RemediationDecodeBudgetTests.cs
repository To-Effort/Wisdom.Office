using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using Wisdom.Office.Core;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 4 的合批回归：<c>PD9</c> 静默空页、<c>PD5</c> 聚合与重复解码预算、
    /// <c>PD22</c> 单页累计解码预算。
    /// Remediation design episode two, phase four, batched entries: <c>PD9</c> silently empty pages,
    /// <c>PD5</c> the aggregate and repeated-decode budget, <c>PD22</c> the per-page cumulative decode budget.
    /// </summary>
    /// <remarks>
    /// 这一批共用同一条「预算与上报」通道：上限类失败原样上抛，语法损坏继续静默降级。
    /// The batch shares one budget-and-reporting route: limit failures propagate unchanged while syntax
    /// damage keeps degrading quietly.
    /// </remarks>
    internal static partial class Program
    {
        /// <summary>预算与静默丢失批次的入口。Entry point for the budget and silent-loss batch.</summary>
        private static void RunRemediationDecodeBudgetTest()
        {
            RunSilentPageLossTest();
            RunContentAggregateBudgetTest();
            RunRepeatedFormBudgetTest();
            RunFormDecodeBudgetTest();
            RunObjectStreamDecodeCacheTest();

            Console.WriteLine("[remediation] decode-budget PD9 PD5 PD22 ok");
        }

        /// <summary>
        /// PD9：内容流解压超限时不得再伪装成「这页没有文字」，上限异常要原样上抛；正文损坏仍须静默降级。
        /// PD9: an over-limit content stream may no longer masquerade as "this page has no text"; the limit
        /// exception has to propagate, while a damaged body still degrades quietly.
        /// </summary>
        private static void RunSilentPageLossTest()
        {
            long saved = PdfReadLimits.MaxDecodedStreamBytes;
            try
            {
                // 控制组：上限内的正常页照常出文，读一遍不许抛。
                // Control: an in-limit page still yields its text without throwing.
                PdfReadLimits.MaxDecodedStreamBytes = 1024L * 1024L;
                string quiet = "BT /F1 12 Tf 50 700 Td (Quiet) Tj ET";
                string quietText = PdfDocument.Load(BuildFlateContentPdf(quiet)).ExtractText();
                Require(quietText.IndexOf("Quiet", StringComparison.Ordinal) >= 0,
                    "an in-limit compressed page lost its text under a 1 MB cap, got '" + quietText + "'");

                // 炸弹页：解压后 4 MB，远超下面设的 4 KB 上限。
                // The bomb page inflates to 4 MB, far past the 4 KB cap set below.
                byte[] bomb = BuildFlateContentPacked(
                    Zlib.Compress(Encoding.ASCII.GetBytes(BuildPadding(4 * 1024 * 1024)
                        + "BT /F1 12 Tf 50 700 Td (Bomb) Tj ET")));
                PdfReadLimits.MaxDecodedStreamBytes = 4096L;
                Exception failure = null;
                try
                {
                    PdfDocument.Load(bomb).ExtractText();
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
                Require(failure is DecompressionLimitException,
                    "a content stream over the decode cap reported '"
                    + (failure == null ? "no exception, ExtractText returned normally" : failure.GetType().Name)
                    + "' instead of DecompressionLimitException");

                // 控制组：正文被改坏的压缩流属语法损坏，只能降级不能上抛。
                // Control: a corrupted compressed body is syntax damage; it must degrade, not propagate.
                byte[] packed = Zlib.Compress(Encoding.ASCII.GetBytes(quiet));
                packed[packed.Length / 2] = (byte)(packed[packed.Length / 2] ^ 0x5A);
                Exception damage = null;
                try
                {
                    PdfDocument.Load(BuildFlateContentPacked(packed)).ExtractText();
                }
                catch (Exception ex)
                {
                    damage = ex;
                }
                Require(damage == null,
                    "a corrupted content stream escaped as " + (damage == null ? "none" : damage.GetType().Name)
                    + "; syntax damage still has to degrade quietly");
            }
            finally
            {
                PdfReadLimits.MaxDecodedStreamBytes = saved;
            }
        }

        /// <summary>
        /// PD5：<c>/Contents</c> 数组把同一个流引用几十次时，聚合出口要有总预算，不得按份数无限堆进内存。
        /// PD5: when the /Contents array references one stream dozens of times, the aggregation exit needs a
        /// total budget instead of piling every copy into memory.
        /// </summary>
        private static void RunContentAggregateBudgetTest()
        {
            long saved = PdfReadLimits.MaxDecodedStreamBytes;
            try
            {
                PdfReadLimits.MaxDecodedStreamBytes = 1024L * 1024L;

                // 单个流 200 KB，低于上限，逐份解码本身合法；50 份合计 10 MB 才越界。
                // One stream is 200 KB, under the cap, so each decode is legal on its own; the fifty copies
                // only breach the limit in aggregate.
                string single = BuildPadding(200 * 1024);
                StringBuilder references = new StringBuilder();
                for (int i = 0; i < 50; i++)
                {
                    if (i > 0) references.Append(' ');
                    references.Append("4 0 R");
                }
                byte[] pdf = BuildMultiContentPdf(single, references.ToString());

                Exception failure = null;
                try
                {
                    PdfDocument.Load(pdf).ExtractText();
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
                Require(failure is DecompressionLimitException,
                    "fifty references to one 200 KB stream aggregated to 10 MB under a 1 MB cap and reported '"
                    + (failure == null ? "no exception" : failure.GetType().Name) + "'");
            }
            finally
            {
                PdfReadLimits.MaxDecodedStreamBytes = saved;
            }
        }

        /// <summary>
        /// PD5/PD22：同一个 Form 被 <c>Do</c> 五十次只算一份预算，解码结果按流实例复用，不得误伤合法的重复盖章。
        /// PD5/PD22: fifty Do invocations of one form charge the budget once, because the decoded bytes are
        /// reused per stream instance; legitimate repeated stamps must not be false-positived.
        /// </summary>
        private static void RunRepeatedFormBudgetTest()
        {
            long saved = PdfReadLimits.MaxDecodedStreamBytes;
            try
            {
                PdfReadLimits.MaxDecodedStreamBytes = 1024L * 1024L;

                byte[] pdf = BuildFormPdf(new int[] { 100 * 1024 }, 50);
                Exception failure = null;
                string text = null;
                try
                {
                    text = PdfDocument.Load(pdf).ExtractText();
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
                Require(failure == null,
                    "fifty Do invocations of one 100 KB form escaped as "
                    + (failure == null ? "none" : failure.GetType().Name + ": " + failure.Message));
                Require(text != null && text.Length > 0,
                    "the repeated form produced no text, got '" + text + "'");
            }
            finally
            {
                PdfReadLimits.MaxDecodedStreamBytes = saved;
            }
        }

        /// <summary>
        /// PD22：一层 <c>Do</c> 扇出到二十个互不相同的 Form 时，单页累计解码字节要受预算约束。
        /// PD22: when one Do fan-out reaches twenty distinct forms, the page's cumulative decoded bytes have
        /// to stay inside the budget.
        /// </summary>
        private static void RunFormDecodeBudgetTest()
        {
            long saved = PdfReadLimits.MaxDecodedStreamBytes;
            try
            {
                PdfReadLimits.MaxDecodedStreamBytes = 1024L * 1024L;

                // 二十个各 100 KB 的不同表单：单看每一个都合法，合起来 2 MB 才越界。
                // Twenty distinct 100 KB forms: each one is legal, the two megabytes together are not.
                int[] sizes = new int[20];
                for (int i = 0; i < sizes.Length; i++) sizes[i] = 100 * 1024;
                byte[] pdf = BuildFormPdf(sizes, 1);

                Exception failure = null;
                try
                {
                    PdfDocument.Load(pdf).ExtractText();
                }
                catch (Exception ex)
                {
                    failure = ex;
                }
                Require(failure is DecompressionLimitException,
                    "twenty distinct 100 KB forms decoded 2 MB under a 1 MB page cap and reported '"
                    + (failure == null ? "no exception" : failure.GetType().Name) + "'");
            }
            finally
            {
                PdfReadLimits.MaxDecodedStreamBytes = saved;
            }
        }

        /// <summary>
        /// PD5：对象流的解码结果要按流实例缓存。页树里的几百个对象同住一个对象流，修复前每取一个对象就重新
        /// 解压整条流，读取被拖成二次方；缓存生效后解压只发生一次（实测 4.6 MB 的对象流 240 个对象，
        /// 修复前 2 468 ms、修复后 84 ms）。
        /// PD5: decoded streams are cached per stream instance. Hundreds of page-tree objects share one object
        /// stream, and every object lookup used to re-inflate the whole stream, turning the read quadratic;
        /// with the cache the inflate happens once (measured on a 4.6 MB stream holding 240 objects:
        /// 2 468 ms before, 84 ms after).
        /// </summary>
        private static void RunObjectStreamDecodeCacheTest()
        {
            const int Pages = 240;
            const int Pad = 20 * 1000;

            byte[] pdf = BuildObjectStreamPdf(Pages, Pad);

            Stopwatch watch = Stopwatch.StartNew();
            int pageCount = PdfDocument.Load(pdf).PageCount;
            watch.Stop();
            long milliseconds = watch.ElapsedMilliseconds;

            Console.WriteLine("[remediation] PD5 object-stream objects=" + Pages
                + " pages=" + pageCount + " ms=" + milliseconds);

            Require(pageCount == Pages, "the object-stream page tree yielded " + pageCount
                + " pages instead of " + Pages);
            Require(milliseconds < 700, Pages + " objects of one " + (Pages * Pad / 1024.0 / 1024.0).ToString("0.0")
                + " MB object stream took " + milliseconds + " ms to read");
        }

        /// <summary>拼一份单页 PDF，正文以 Flate 压缩存储。Assembles a one-page PDF whose body is Flate-compressed.</summary>
        /// <param name="content">未压缩的正文。The uncompressed content stream.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildFlateContentPdf(string content)
        {
            return BuildFlateContentPacked(Zlib.Compress(Encoding.ASCII.GetBytes(content)));
        }

        /// <summary>拼一份单页 PDF，正文以给定的压缩字节存储。Assembles a one-page PDF whose body is the given compressed bytes.</summary>
        /// <param name="packed">已经 Flate 压缩的正文。The already Flate-compressed body.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildFlateContentPacked(byte[] packed)
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                    + "/Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Filter /FlateDecode /Length " + packed.Length + " >>\nstream\n",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            return AssembleBinaryPdf(bodies, 3, packed, "\nendstream");
        }

        /// <summary>
        /// 拼一份 <c>/Contents</c> 为数组的单页 PDF，数组内容由调用方给出（可重复引用同一个流）。
        /// Assembles a one-page PDF whose /Contents is an array supplied by the caller, so it may reference
        /// the same stream repeatedly.
        /// </summary>
        /// <param name="single">流正文（不压缩）。The uncompressed stream body.</param>
        /// <param name="references">数组内的引用串，如 <c>4 0 R 4 0 R</c>。The reference list inside the array.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildMultiContentPdf(string single, string references)
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents [" + references
                    + "] /Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Length " + Encoding.ASCII.GetByteCount(single) + " >>\nstream\n" + single + "\nendstream",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            return AssembleClassicPdf(bodies, 1, null);
        }

        /// <summary>
        /// 拼一页把若干个 Form XObject 各自 <c>Do</c> 若干次的 PDF。
        /// Assembles a page that Does each of several Form XObjects the requested number of times.
        /// </summary>
        /// <param name="sizes">每个 Form 的未压缩正文长度。The uncompressed body length of each form.</param>
        /// <param name="invocations">每个表单被 <c>Do</c> 的次数。The number of Do invocations per form.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildFormPdf(int[] sizes, int invocations)
        {
            List<string> bodies = new List<string>();
            bodies.Add("<< /Type /Catalog /Pages 2 0 R >>");
            bodies.Add("<< /Type /Pages /Kids [3 0 R] /Count 1 >>");

            StringBuilder xobjects = new StringBuilder();
            StringBuilder content = new StringBuilder();
            for (int i = 0; i < sizes.Length; i++)
            {
                string name = "F" + i;
                if (xobjects.Length > 0) xobjects.Append(' ');
                xobjects.Append('/').Append(name).Append(' ').Append(6 + i).Append(" 0 R");
                for (int k = 0; k < invocations; k++)
                {
                    content.Append('/').Append(name).Append(" Do\n");
                }
            }

            bodies.Add("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                + "/Resources << /Font << /F1 5 0 R >> /XObject << " + xobjects + " >> >> >>");
            string body = "BT /F1 12 Tf 50 700 Td (X) Tj ET\n" + content;
            bodies.Add("<< /Length " + Encoding.ASCII.GetByteCount(body) + " >>\nstream\n" + body + "\nendstream");
            bodies.Add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>");
            for (int i = 0; i < sizes.Length; i++)
            {
                string formBody = BuildPadding(sizes[i]) + "BT /F1 12 Tf 0 0 Td (R) Tj ET";
                bodies.Add("<< /Type /XObject /Subtype /Form /BBox [0 0 595 842] /Length "
                    + Encoding.ASCII.GetByteCount(formBody) + " >>\nstream\n" + formBody + "\nendstream");
            }
            return AssembleClassicPdf(bodies.ToArray(), 1, null);
        }

        /// <summary>
        /// 生成指定长度的注释正文：解释器一行就跳过，读起来几乎不花钱，只用它把解码字节数顶到预算上。
        /// Builds a comment body of the requested length: the interpreter skips it in one step, so it costs
        /// almost nothing to run and only pushes the decoded byte count onto the budget.
        /// </summary>
        /// <param name="length">正文字节数。The body length in bytes.</param>
        /// <returns>以换行收尾的注释串。The comment text ending with a newline.</returns>
        private static string BuildPadding(int length)
        {
            StringBuilder builder = new StringBuilder(length + 8);
            builder.Append("% ");
            for (int i = 0; i < length - 3; i++) builder.Append('x');
            builder.Append('\n');
            return builder.ToString();
        }

        /// <summary>
        /// 拼一份「页树全部住在对象流里」的 PDF：对象流以 Flate 压缩存储，交叉表流用 type 2 条目指向它。
        /// Assembles a PDF whose whole page tree lives inside an object stream: the stream is Flate-compressed
        /// and the cross-reference stream points into it with type 2 rows.
        /// </summary>
        /// <param name="pages">页数量。The number of pages.</param>
        /// <param name="padSize">每页字典的填充串长度，用来把对象流的解压体积顶大。The padding length of every page dictionary, used to grow the object stream's decoded size.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildObjectStreamPdf(int pages, int padSize)
        {
            Encoding ascii = Encoding.ASCII;

            List<int> numbers = new List<int>();
            List<int> offsets = new List<int>();
            StringBuilder inner = new StringBuilder();

            StringBuilder kids = new StringBuilder();
            for (int i = 0; i < pages; i++)
            {
                if (i > 0) kids.Append(' ');
                kids.Append(4 + i).Append(" 0 R");
            }

            numbers.Add(2);
            offsets.Add(inner.Length);
            inner.Append("<< /Type /Catalog /Pages 3 0 R >>");

            numbers.Add(3);
            offsets.Add(inner.Length);
            inner.Append("<< /Type /Pages /Kids [").Append(kids).Append("] /Count ").Append(pages).Append(" >>");

            // 每页的填充串各不相同：整段重复的填充会被压到几乎消失，解压耗时也就测不出来了。
            // Every page gets a different pad: one repeated pad compresses away and the inflate cost,
            // which is exactly what the cache has to remove, becomes unmeasurable.
            char[] pool = BuildIncompressibleRun(pages * padSize);
            for (int i = 0; i < pages; i++)
            {
                numbers.Add(4 + i);
                offsets.Add(inner.Length);
                inner.Append("<< /Type /Page /Parent 3 0 R /MediaBox [0 0 595 842] /Pad (")
                    .Append(pool, i * padSize, padSize)
                    .Append(") >>");
            }

            string header = string.Empty;
            for (int i = 0; i < numbers.Count; i++)
            {
                header += numbers[i] + " " + offsets[i] + " ";
            }
            byte[] packed = Zlib.Compress(ascii.GetBytes(header + inner.ToString()));

            int xrefNumber = 4 + pages;
            string objStreamHead = "<< /Type /ObjStm /N " + numbers.Count + " /First "
                + ascii.GetByteCount(header) + " /Filter /FlateDecode /Length " + packed.Length
                + " >>\nstream\n";

            // 交叉表条目要指向「N 0 obj」的开头，正文偏移与对象头偏移不能混用。
            // A cross-reference row points at the "N 0 obj" head, not at the stream body.
            long xrefAt = ascii.GetByteCount("%PDF-1.7\n") + ascii.GetByteCount("1 0 obj\n")
                + ascii.GetByteCount(objStreamHead) + packed.Length
                + ascii.GetByteCount("\nendstream\nendobj\n");

            byte[] rows = new byte[(xrefNumber + 1) * 7];
            int at = 0;
            WriteXrefRow(rows, ref at, 0, 0L, 0);              // 0：空闲号 free slot
            WriteXrefRow(rows, ref at, 1, ascii.GetByteCount("%PDF-1.7\n"), 0);   // 1：对象流本体 the object stream
            for (int i = 2; i <= 3 + pages; i++)
            {
                WriteXrefRow(rows, ref at, 2, 1, i - 2);      // 2：住在对象 1 里的第 i-2 个 index i-2 inside object 1
            }
            WriteXrefRow(rows, ref at, 1, xrefAt, 0);          // 交叉表流自身 the xref stream itself

            string xrefHead = xrefNumber + " 0 obj\n<< /Type /Xref /Size " + (xrefNumber + 1)
                + " /W [1 4 2] /Root 2 0 R /Length " + rows.Length + " >>\nstream\n";

            List<byte[]> pieces = new List<byte[]>();
            pieces.Add(ascii.GetBytes("%PDF-1.7\n"));
            pieces.Add(ascii.GetBytes("1 0 obj\n"));
            pieces.Add(ascii.GetBytes(objStreamHead));
            pieces.Add(packed);
            pieces.Add(ascii.GetBytes("\nendstream\nendobj\n"));
            pieces.Add(ascii.GetBytes(xrefHead));
            pieces.Add(rows);
            pieces.Add(ascii.GetBytes("\nendstream\nendobj\nstartxref\n" + xrefAt + "\n%%EOF\n"));

            int total = 0;
            for (int i = 0; i < pieces.Count; i++) total += pieces[i].Length;
            byte[] pdf = new byte[total];
            int cursor = 0;
            for (int i = 0; i < pieces.Count; i++)
            {
                Buffer.BlockCopy(pieces[i], 0, pdf, cursor, pieces[i].Length);
                cursor += pieces[i].Length;
            }
            return pdf;
        }

        /// <summary>写一条 <c>/W [1 4 2]</c> 的交叉表条目。Writes one /W [1 4 2] cross-reference row.</summary>
        /// <param name="rows">条目缓冲区。The row buffer.</param>
        /// <param name="at">写入位置，调用后前进 7 字节。The write cursor, advanced by seven bytes.</param>
        /// <param name="type">条目类型（0 空闲、1 直接对象、2 对象流内）。The row type (0 free, 1 plain, 2 inside an object stream).</param>
        /// <param name="field">类型 1 为偏移，类型 2 为所在对象流号。For type 1 the offset, for type 2 the containing object stream number.</param>
        /// <param name="index">类型 2 的流内序号；类型 1 传 0。The index inside the object stream for type 2; zero for type 1.</param>
        private static void WriteXrefRow(byte[] rows, ref int at, int type, long field, int index)
        {
            rows[at] = (byte)type;
            rows[at + 1] = (byte)(field >> 24);
            rows[at + 2] = (byte)(field >> 16);
            rows[at + 3] = (byte)(field >> 8);
            rows[at + 4] = (byte)field;
            rows[at + 5] = (byte)(index >> 8);
            rows[at + 6] = (byte)index;
            at += 7;
        }

        /// <summary>生成一段无规律的数字填充串；重复度高的填充会被压缩到几乎消失，解压耗时也就测不出来了。Builds an unpredictable digit run; a repetitive pad compresses away, so the inflate cost would not be measurable.</summary>
        /// <param name="length">串长度。The length.</param>
        /// <returns>填充字符数组。The padding characters.</returns>
        private static char[] BuildIncompressibleRun(int length)
        {
            char[] result = new char[length];
            int state = 12345;
            for (int i = 0; i < length; i++)
            {
                state = (state * 1103515245 + 12345) & 0x7FFFFFFF;
                result[i] = (char)('0' + state % 10);
            }
            return result;
        }
    }
}
