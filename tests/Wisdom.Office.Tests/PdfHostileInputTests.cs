using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Office.Core.Formulas;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 畸形输入回归测试：证明解析器在恶意交叉表、超长 /Length、超深页面树、越界偏移和
    /// 爆弹式内容流面前只会放弃残余对象，不会把进程拖进内存耗尽或栈溢出。
    /// Hostile-input regression tests: they prove the parser gives up on the remaining objects instead of
    /// exhausting memory or the stack when faced with a malicious cross-reference table, an oversized
    /// /Length, a very deep page tree, out-of-range offsets, or a bomb-shaped content stream.
    /// </summary>
    internal static partial class Program
    {
        private static void RunPdfHostileInputTest(string outputDirectory)
        {
            RunStreamLengthOverflow();
            RunXrefStreamSizeBomb();
            RunDeepPageTree();
            RunOutOfRangeObjectOffset();
            RunBareOperandBomb();
            RunHostilePngDimensions(outputDirectory);
            RunFormulaNestingBomb();
            RunStreamInputCeiling();
            RunUnresolvableCryptFilter();
            Console.WriteLine("[hostile] hostile-input limits ok");
        }

        /// <summary>/StmF 点名了 /CF 里不存在的过滤器时必须报错，而不是把密文当明文返回。A /StmF naming a filter that /CF never defines must raise instead of handing ciphertext back as plaintext.</summary>
        private static void RunUnresolvableCryptFilter()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << >> >>",
                "<< /Filter /Standard /V 4 /R 4 /Length 128 /P -39 /O (01234567890123456789012345678901)"
                    + " /U (01234567890123456789012345678901) /StmF /StdCF /StrF /StdCF /CF << /Other << /CFM /V2 >> >> >>"
            };
            byte[] pdf = AssembleClassicPdf(bodies, 1, null, "/Encrypt 4 0 R ");
            Exception failure = null;
            try
            {
                PdfDocument.Load(pdf, "pw");
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            Require(failure is InvalidDataException,
                "undefined crypt filter produced " + (failure == null ? "no failure" : failure.GetType().Name + ": " + failure.Message));
        }

        /// <summary>读取端按字节计费而不是相信流的声明长度。Readers charge bytes as they arrive instead of trusting the stream's declared length.</summary>
        private static void RunStreamInputCeiling()
        {
            long saved = OfficePackageLimits.MaxEntryBytes;
            OfficePackageLimits.MaxEntryBytes = 1L << 20;
            try
            {
                using (MemoryStream stream = new MemoryStream())
                {
                    byte[] block = new byte[64 * 1024];
                    for (int i = 0; i < 40; i++) stream.Write(block, 0, block.Length);
                    stream.Position = 0;
                    Exception failure = null;
                    try
                    {
                        PdfDocument.Load(stream);
                    }
                    catch (Exception ex)
                    {
                        failure = ex;
                    }
                    Require(failure is InvalidDataException,
                        "oversized input stream produced " + (failure == null ? "no failure" : failure.GetType().Name));
                }
            }
            finally
            {
                OfficePackageLimits.MaxEntryBytes = saved;
            }
            Console.WriteLine("[hostile] stream input ceiling ok");
        }

        /// <summary>五千层括号的公式必须由深度上限裁决，而不是把调用栈打爆。A formula with 5 000 nested parentheses must be settled by the depth cap instead of blowing the stack.</summary>
        private static void RunFormulaNestingBomb()
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.AddWorksheet("S1");
            string deep = new string('(', 5000) + "1" + new string(')', 5000);
            Stopwatch watch = Stopwatch.StartNew();
            object value = null;
            Exception failure = null;
            try
            {
                value = FormulaEngine.Evaluate(workbook, deep);
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            watch.Stop();
            Require(watch.ElapsedMilliseconds < 10000, "nested-paren formula took " + watch.ElapsedMilliseconds + " ms");
            Require(failure == null || failure is FormatException,
                "nested-paren formula failed with " + (failure == null ? "null" : failure.GetType().Name + ": " + failure.Message));
            Console.WriteLine("[hostile] formula nesting capped, result=" + (value == null ? failure.GetType().Name : value));
        }

        /// <summary>/Length 接近 int.MaxValue 时不得溢出为负数并据此分配巨型缓冲。/Length near int.MaxValue must not wrap and size a giant buffer.</summary>
        private static void RunStreamLengthOverflow()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << >> >>",
                "<< /Length 2147483647 >>\nstream\nBT (hostile) Tj ET\n"
            };
            byte[] pdf = AssembleClassicPdf(bodies, 1, null);
            Stopwatch watch = Stopwatch.StartNew();
            string text = null;
            Exception failure = null;
            try
            {
                text = PdfDocument.Load(pdf).ExtractText();
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            watch.Stop();
            Require(watch.ElapsedMilliseconds < 10000, "oversized /Length took " + watch.ElapsedMilliseconds + " ms");
            Require(!(failure is OutOfMemoryException) && !(failure is OverflowException),
                "oversized /Length leaked to " + (failure == null ? "null" : failure.GetType().Name));
            Require(text == null || text.Length == 0, "oversized /Length invented text: " + text);
        }

        /// <summary>交叉表流声明天文数字 /Size 与 /Index 时只能按已读到的行收敛。A cross-reference stream with an astronomical /Size or /Index must stop at the rows actually present.</summary>
        private static void RunXrefStreamSizeBomb()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << >> >>"
            };
            long[] offsets = ObjectOffsets(bodies);
            // 每行 7 字节：类型 1 字节、偏移 4 字节、代号 2 字节，只给出前两条真实条目。
            // Seven bytes per row (1 type, 4 offset, 2 generation) with only two real rows present.
            byte[] rows = new byte[]
            {
                0, 0, 0, 0, 0, 0, 0,
                1, (byte)(offsets[0] >> 24), (byte)(offsets[0] >> 16), (byte)(offsets[0] >> 8), (byte)offsets[0], 0, 0,
                1, (byte)(offsets[1] >> 24), (byte)(offsets[1] >> 16), (byte)(offsets[1] >> 8), (byte)offsets[1], 0, 0
            };
            byte[] pdf = AssembleXrefStreamPdf(bodies, rows, 1, 2000000000L, 0L, 2000000000L);
            Stopwatch watch = Stopwatch.StartNew();
            int pages = -1;
            Exception failure = null;
            try
            {
                PdfFile file = PdfFile.Load(pdf);
                pages = file.PageCount;
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            watch.Stop();
            Require(watch.ElapsedMilliseconds < 10000, "xref stream /Size bomb took " + watch.ElapsedMilliseconds + " ms");
            Require(!(failure is OutOfMemoryException), "xref stream /Size bomb exhausted memory: " + failure);
            Require(pages <= 1, "xref stream /Size bomb reported " + pages + " pages");
        }

        /// <summary>两万个单子节点的 Pages 链必须被深度上限掐断，而不是撑爆调用栈。A 20 000-deep chain of single-kid Pages nodes must be cut by the depth cap instead of the stack.</summary>
        private static void RunDeepPageTree()
        {
            int depth = 20000;
            string[] bodies = new string[depth + 2];
            bodies[0] = "<< /Type /Catalog /Pages 2 0 R >>";
            for (int i = 1; i <= depth; i++)
            {
                bodies[i] = "<< /Type /Pages /Kids [" + (i + 2) + " 0 R] /Count 1 >>";
            }
            bodies[depth + 1] = "<< /Type /Page /Parent " + (depth + 1) + " 0 R /MediaBox [0 0 595 842] /Resources << >> >>";
            byte[] pdf = AssembleClassicPdf(bodies, 1, null);

            Stopwatch watch = Stopwatch.StartNew();
            int pageCount = -1;
            Exception failure = null;
            try
            {
                pageCount = PdfFile.Load(pdf).PageCount;
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            watch.Stop();
            Require(failure == null, "deep page tree threw " + failure);
            Require(pageCount == 0, "deep page tree collected " + pageCount + " pages past the depth cap");
            Require(watch.ElapsedMilliseconds < 15000, "deep page tree took " + watch.ElapsedMilliseconds + " ms");
        }

        /// <summary>交叉表里的 64 位偏移直接截断会翻负或落到文件中部，必须按不存在处理。A 64-bit xref offset truncated straight to int wraps or lands mid-file and must read as absent.</summary>
        private static void RunOutOfRangeObjectOffset()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << >> >>",
                "<< /Length 26 >>\nstream\nBT (SECRET) Tj ET\nendstream"
            };
            long[] offsets = ObjectOffsets(bodies);
            long[] hostile = new long[bodies.Length];
            for (int i = 0; i < bodies.Length; i++) hostile[i] = offsets[i];
            // 内容流对象指向 int.MaxValue 之外：10 位十进制正好容纳 9999999999。
            // Point the content-stream object past int.MaxValue; ten decimal digits still hold 9999999999.
            hostile[3] = 9999999999L;
            byte[] pdf = AssembleClassicPdf(bodies, 1, hostile);

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
            Require(failure == null, "out-of-range object offset threw " + failure);
            Require(text == null || text.IndexOf("SECRET", StringComparison.Ordinal) < 0,
                "out-of-range object offset still decoded the stream");
        }

        /// <summary>只堆数字、永不给出算符的内容流不能无限累积操作数。A stream of bare numbers with no operator must not pile up operands forever.</summary>
        private static void RunBareOperandBomb()
        {
            StringBuilder numbers = new StringBuilder();
            for (int i = 0; i < 400000; i++) numbers.Append("1 ");
            string streamBody = numbers.ToString();
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << >> >>",
                "<< /Length " + streamBody.Length + " >>\nstream\n" + streamBody + "\nendstream"
            };
            byte[] pdf = AssembleClassicPdf(bodies, 1, null);

            Stopwatch watch = Stopwatch.StartNew();
            string text = null;
            Exception failure = null;
            try
            {
                text = PdfDocument.Load(pdf).ExtractText();
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            watch.Stop();
            Require(failure == null, "bare operand bomb threw " + failure);
            Require(watch.ElapsedMilliseconds < 30000, "bare operand bomb took " + watch.ElapsedMilliseconds + " ms");
            Require(text == null || text.Length == 0, "bare operand bomb produced text of " + text.Length + " chars");
        }

        /// <summary>IHDR 声明超大尺寸的 PNG 只能按上限拒绝，不能按乘积分配像素缓冲。A PNG whose IHDR declares enormous dimensions must be rejected by the limit, not sized by the product.</summary>
        private static void RunHostilePngDimensions(string outputDirectory)
        {
            string path = Path.Combine(outputDirectory, "hostile-dimensions.png");
            byte[] png = BuildPng(0x7F000000, 1);
            File.WriteAllBytes(path, png);

            Exception failure = null;
            try
            {
                OfficeDocument document = new OfficeDocument();
                Section section = document.AddSection(PageSetup.A4());
                section.AddParagraph("hostile png").AddImage(OfficeImage.FromFile(path), 120, 120);
                new PdfRenderer().Render(document, PrepareOutputPath(Path.Combine(outputDirectory, "hostile-dimensions.pdf")));
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            Require(failure is InvalidDataException, "hostile png produced " + (failure == null ? "no failure" : failure.GetType().Name + ": " + failure.Message));
        }

        /// <summary>拼装一份签名 + IHDR + 空 IDAT + IEND 的 PNG 外壳。Assembles a PNG shell: signature, IHDR, an empty IDAT, IEND.</summary>
        private static byte[] BuildPng(int width, int height)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                byte[] signature = new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
                stream.Write(signature, 0, signature.Length);
                byte[] ihdr = new byte[13];
                ihdr[0] = (byte)(width >> 24);
                ihdr[1] = (byte)(width >> 16);
                ihdr[2] = (byte)(width >> 8);
                ihdr[3] = (byte)width;
                ihdr[4] = (byte)(height >> 24);
                ihdr[5] = (byte)(height >> 16);
                ihdr[6] = (byte)(height >> 8);
                ihdr[7] = (byte)height;
                ihdr[8] = 8;   // 位深。Bits per component.
                ihdr[9] = 0;   // 灰度。Grayscale.
                WritePngChunk(stream, "IHDR", ihdr);
                WritePngChunk(stream, "IDAT", new byte[] { 0x78, 0x9C, 0x03, 0x00, 0x00, 0x00, 0x01 });
                WritePngChunk(stream, "IEND", new byte[0]);
                return stream.ToArray();
            }
        }

        /// <summary>写入一个长度前缀与 CRC 占位的 PNG 数据块。Writes one PNG chunk with a length prefix and placeholder CRC.</summary>
        private static void WritePngChunk(MemoryStream stream, string type, byte[] data)
        {
            int length = data.Length;
            stream.WriteByte((byte)(length >> 24));
            stream.WriteByte((byte)(length >> 16));
            stream.WriteByte((byte)(length >> 8));
            stream.WriteByte((byte)length);
            byte[] name = Encoding.ASCII.GetBytes(type);
            stream.Write(name, 0, name.Length);
            stream.Write(data, 0, data.Length);
            stream.Write(new byte[4], 0, 4);
        }

        /// <summary>计算各对象在文件中的字节偏移（下标 0 对应 1 号对象）。Computes each object's byte offset (index 0 is object 1).</summary>
        private static long[] ObjectOffsets(string[] bodies)
        {
            Encoding ascii = Encoding.ASCII;
            long[] offsets = new long[bodies.Length];
            long position = ascii.GetByteCount("%PDF-1.7\n");
            for (int i = 0; i < bodies.Length; i++)
            {
                offsets[i] = position;
                position += ascii.GetByteCount(RenderObject(i + 1, bodies[i]));
            }
            return offsets;
        }

        /// <summary>组装经典交叉表 PDF，可选覆盖条目偏移以制造畸形文件。Assembles a classic-xref PDF, optionally overriding entry offsets to forge a malformed file.</summary>
        private static byte[] AssembleClassicPdf(string[] bodies, int rootNumber, long[] overrideOffsets)
        {
            return AssembleClassicPdf(bodies, rootNumber, overrideOffsets, string.Empty);
        }

        /// <summary>组装经典交叉表 PDF，并可追加 trailer 条目（例如 /Encrypt）。Assembles a classic-xref PDF with extra trailer entries such as /Encrypt.</summary>
        private static byte[] AssembleClassicPdf(string[] bodies, int rootNumber, long[] overrideOffsets, string extraTrailer)
        {
            Encoding ascii = Encoding.ASCII;
            long[] offsets = ObjectOffsets(bodies);
            StringBuilder builder = new StringBuilder("%PDF-1.7\n");
            for (int i = 0; i < bodies.Length; i++) builder.Append(RenderObject(i + 1, bodies[i]));
            long xrefAt = ascii.GetByteCount(builder.ToString());
            builder.Append("xref\n0 ").Append(bodies.Length + 1).Append("\n0000000000 65535 f \n");
            for (int i = 0; i < bodies.Length; i++)
            {
                long offset = overrideOffsets == null ? offsets[i] : overrideOffsets[i];
                builder.Append(offset.ToString("D10")).Append(" 00000 n \n");
            }
            builder.Append("trailer\n<< /Size ").Append(bodies.Length + 1)
                .Append(" /Root ").Append(rootNumber).Append(" 0 R ").Append(extraTrailer).Append(">>\nstartxref\n").Append(xrefAt).Append("\n%%EOF\n");
            return ascii.GetBytes(builder.ToString());
        }

        /// <summary>
        /// 组装以交叉表流收尾的 PDF：表流自身不作为条目给出，只喂给读取端畸形 /Size 与 /Index。
        /// Assembles a PDF that ends with a cross-reference stream: the stream itself is not listed, so the
        /// reader only ever sees the forged /Size and /Index plus the supplied rows.
        /// </summary>
        private static byte[] AssembleXrefStreamPdf(string[] bodies, byte[] rows, int rootNumber, long size,
            long indexStart, long indexCount)
        {
            Encoding ascii = Encoding.ASCII;
            StringBuilder builder = new StringBuilder("%PDF-1.7\n");
            for (int i = 0; i < bodies.Length; i++) builder.Append(RenderObject(i + 1, bodies[i]));
            long streamAt = ascii.GetByteCount(builder.ToString());
            int streamNumber = bodies.Length + 1;
            string dictionary = "<< /Type /XRef /Size " + size + " /W [1 4 2] /Index [" + indexStart + " " + indexCount +
                "] /Root " + rootNumber + " 0 R /Length " + rows.Length + " >>";
            builder.Append(streamNumber).Append(" 0 obj\n").Append(dictionary).Append("\nstream\n");
            byte[] head = ascii.GetBytes(builder.ToString());
            byte[] tail = ascii.GetBytes("\nendstream\nendobj\nstartxref\n" + streamAt + "\n%%EOF\n");
            byte[] pdf = new byte[head.Length + rows.Length + tail.Length];
            Buffer.BlockCopy(head, 0, pdf, 0, head.Length);
            Buffer.BlockCopy(rows, 0, pdf, head.Length, rows.Length);
            Buffer.BlockCopy(tail, 0, pdf, head.Length + rows.Length, tail.Length);
            return pdf;
        }

        /// <summary>渲染单个间接对象的文本。Renders one indirect object as text.</summary>
        private static string RenderObject(int number, string body)
        {
            return number + " 0 obj\n" + body + "\nendobj\n";
        }
    }
}
