using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
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
            RunSyntaxNestingBomb();
            RunOutOfRangeObjectOffset();
            RunBareOperandBomb();
            RunHostilePngDimensions(outputDirectory);
            RunFormulaNestingBomb();
            RunStreamInputCeiling();
            RunUnresolvableCryptFilter();
            RunAesPaddingTolerance();
            RunXrefStartNumberWrap();
            RunDeletedObjectNotResurrected();
            RunObjectStreamFirstOutOfRange();
            Console.WriteLine("[hostile] hostile-input limits ok");
        }

        /// <summary>
        /// 交叉表把区段起始号报成 int 上界附近：截断成 int 之后要么回绕成真实对象号、把对象指到别处，
        /// 要么把写出端的 /Size 抬成负数。起始号必须封顶，越界的条目一律作废。
        /// A classic xref subsection whose start number sits next to the int ceiling: truncated to int it either
        /// wraps onto a real object number and redirects that object, or it lifts the writer's /Size into
        /// negative territory. Object numbers have to be capped and out-of-range entries dropped.
        /// </summary>
        private static void RunXrefStartNumberWrap()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << >> >>",
                "<< /Length 17 >>\nstream\nBT (KEEPME) Tj ET\nendstream"
            };
            // 新节起始 4294967298 截断成 int 正好是 2，也就是真实的 Pages 节点。
            // Truncating 4294967298 to int lands exactly on object 2, the real Pages node.
            byte[] pdf = AssembleTwoSectionPdf(bodies, 1, 4294967298L,
                new long[] { 12345L }, new int[] { 0 }, new string[] { "n" });

            Exception failure = null;
            int pages = -1;
            int docPages = -1;
            try
            {
                pages = PdfFile.Load(pdf).PageCount;
                docPages = PdfDocument.Load(pdf).Pages.Count;
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            Require(failure == null, "wrapped xref start number threw " + failure);
            Require(pages == 1, "wrapped xref start number left " + pages + " pages");
            Require(docPages == 1, "wrapped xref start number redirected the Pages node to " + docPages
                + " pages");

            // 起始号逼近 int 上界的第二节不能把写出端的 /Size 抬成负数。
            // A second section near the int ceiling must not lift the writer's /Size into negatives.
            byte[] high = AssembleTwoSectionPdf(bodies, 1, 2147483645L,
                new long[] { 12345L, 12346L, 12347L }, new int[] { 0, 0, 0 },
                new string[] { "n", "n", "f" });
            byte[] saved = PdfFile.Load(high).SaveToBytes();
            string savedText = Encoding.ASCII.GetString(saved);
            Require(savedText.IndexOf("/Size -") < 0, "high xref start number wrote a negative /Size");
            Require(savedText.IndexOf("-214748") < 0, "high xref start number wrote a wrapped object number");
            Require(PdfFile.Load(saved).PageCount == 1, "the saved document no longer has its page");
            Console.WriteLine("[hostile] xref start numbers capped");
        }

        /// <summary>
        /// 最后一次增量更新标记为已释放的对象，绝不能从更早的 /Prev 段里复活：
        /// 内容流一旦被复活，涂黑与删除就成了摆设。
        /// An object the newest incremental update marks free must never be resurrected from an older /Prev
        /// section: once a content stream comes back, redaction and deletion become cosmetic.
        /// </summary>
        private static void RunDeletedObjectNotResurrected()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << >> >>",
                "<< /Length 17 >>\nstream\nBT (SECRET) Tj ET\nendstream"
            };
            // 新节把页面对象标成 f，其余条目仍由基节供给。
            // The newer section frees the page object; every other entry still comes from the base.
            byte[] pdf = AssembleTwoSectionPdf(bodies, 1, 3L,
                new long[] { 0L }, new int[] { 0 }, new string[] { "f" });

            Exception failure = null;
            int filePages = -1;
            int docPages = -1;
            try
            {
                filePages = PdfFile.Load(pdf).PageCount;
                docPages = PdfDocument.Load(pdf).Pages.Count;
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            Require(failure == null, "a freed object threw " + failure);
            Require(filePages == 0, "the base section resurrected the deleted page: " + filePages + " pages");
            Require(docPages == 0, "the reading model still resolved the freed page: " + docPages + " pages");
            Console.WriteLine("[hostile] freed entries stay freed");
        }

        /// <summary>
        /// 对象流的 /First 声明到正文之外：先减后比，越界就当作没有这个对象，
        /// 不能拿 int 相加之后的位置去读，也不能按声明长度分配。
        /// An object stream whose /First points beyond its body: bounds are checked by subtracting first, so the
        /// entry is treated as absent instead of reading from a wrapped position or allocating the declared size.
        /// </summary>
        private static void RunObjectStreamFirstOutOfRange()
        {
            string[] bodies = new string[]
            {
                "<< /Type /ObjStm /N 1 /First 999999999 >>\nstream\n0 0\nendstream",
                "<< /Type /Pages /Kids [] /Count 0 >>"
            };
            long[] offsets = ObjectOffsets(bodies);
            byte[] rows = new byte[]
            {
                0, 0, 0, 0, 0, 0, 0,
                1, (byte)(offsets[0] >> 24), (byte)(offsets[0] >> 16), (byte)(offsets[0] >> 8), (byte)offsets[0], 0, 0,
                2, 0, 0, 0, 1, 0, 0
            };
            byte[] pdf = AssembleXrefStreamPdf(bodies, rows, 2, 3L, 0L, 3L);

            Stopwatch watch = Stopwatch.StartNew();
            int pages = -1;
            Exception failure = null;
            try
            {
                pages = PdfFile.Load(pdf).PageCount;
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            watch.Stop();
            Require(!(failure is OutOfMemoryException) && !(failure is OverflowException),
                "an out-of-range /First leaked to " + (failure == null ? "null" : failure.GetType().Name));
            Require(failure == null || failure is InvalidDataException,
                "an out-of-range /First failed with " + (failure == null ? "none" : failure.GetType().Name));
            Require(pages <= 0, "an out-of-range /First produced " + pages + " pages");
            Require(watch.ElapsedMilliseconds < 10000,
                "an out-of-range /First took " + watch.ElapsedMilliseconds + " ms");
            Console.WriteLine("[hostile] object stream /First range-checked");
        }

        /// <summary>
        /// 组装带两段经典交叉表的 PDF：基节按真实对象号登记，新节用给定的起始号与条目登记，
        /// 新 trailer 通过 /Prev 指向基节，于是读取端先消费新节。
        /// Assembles a PDF with two classic cross-reference sections: the base section lists the real object
        /// numbers, the newer section lists the start number and rows supplied here, and the newer trailer
        /// reaches the base through /Prev, so the reader consumes the newer section first.
        /// </summary>
        /// <param name="bodies">全部对象正文。The bodies of every object.</param>
        /// <param name="rootNumber">trailer /Root 对象号。The trailer /Root object number.</param>
        /// <param name="newerStart">新节起始对象号。The start number of the newer subsection.</param>
        /// <param name="newerOffsets">新节条目偏移。The entry offsets of the newer subsection.</param>
        /// <param name="newerGenerations">新节条目代次。The entry generations of the newer subsection.</param>
        /// <param name="newerTypes">新节条目类型（n 或 f）。The entry types of the newer subsection.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] AssembleTwoSectionPdf(string[] bodies, int rootNumber, long newerStart,
            long[] newerOffsets, int[] newerGenerations, string[] newerTypes)
        {
            Encoding ascii = Encoding.ASCII;
            long[] offsets = ObjectOffsets(bodies);
            StringBuilder builder = new StringBuilder("%PDF-1.7\n");
            for (int i = 0; i < bodies.Length; i++) builder.Append(RenderObject(i + 1, bodies[i]));

            long baseAt = ascii.GetByteCount(builder.ToString());
            builder.Append("xref\n0 ").Append(bodies.Length + 1).Append("\n0000000000 65535 f \n");
            for (int i = 0; i < bodies.Length; i++)
            {
                builder.Append(offsets[i].ToString("D10")).Append(" 00000 n \n");
            }
            builder.Append("trailer\n<< /Size ").Append(bodies.Length + 1)
                .Append(" /Root ").Append(rootNumber).Append(" 0 R >>\nstartxref\n")
                .Append(baseAt.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");

            long newerAt = ascii.GetByteCount(builder.ToString());
            builder.Append("xref\n").Append(newerStart.ToString(CultureInfo.InvariantCulture))
                .Append(' ').Append(newerOffsets.Length).Append('\n');
            for (int i = 0; i < newerOffsets.Length; i++)
            {
                builder.Append(newerOffsets[i].ToString("D10")).Append(' ')
                    .Append(newerGenerations[i].ToString("D5")).Append(' ')
                    .Append(newerTypes[i]).Append(" \n");
            }
            builder.Append("trailer\n<< /Size ").Append(bodies.Length + 1)
                .Append(" /Root ").Append(rootNumber).Append(" 0 R /Prev ")
                .Append(baseAt.ToString(CultureInfo.InvariantCulture)).Append(" >>\nstartxref\n")
                .Append(newerAt.ToString(CultureInfo.InvariantCulture)).Append("\n%%EOF\n");
            return ascii.GetBytes(builder.ToString());
        }

        /// <summary>
        /// AES-128 文档里填充不合规的串对象：写出器留下非 16 字节对齐的密文，或者填进非法填充字节时，
        /// 口令校验通过后读取端仍要交出它本来可以还原的明文。
        /// String objects whose AES-128 padding does not conform: when a writer leaves the ciphertext unaligned
        /// or fills the padding with illegal bytes, the reader still owes the user the plaintext it is able to
        /// recover once the password has been accepted.
        /// </summary>
        private static void RunAesPaddingTolerance()
        {
            byte[] id = Md5(Encoding.ASCII.GetBytes("aes-padding-probe"));
            byte[] o = Concat(Md5(Encoding.ASCII.GetBytes("owner-a")), Md5(Encoding.ASCII.GetBytes("owner-b")));
            byte[] key = ReferenceFileKey(PadPassword("pw"), o, -39, id, 4, 128, true);
            byte[] u = Concat(Slice(Rc4Rounds(key, Md5(Concat(PasswordPaddingValue, id))), 0, 16), id);

            Require(TitleOfEncryptedDocument(key, o, u, id, "PADTOLERANCE", false)
                .StartsWith("PADTOLERANCE"),
                "the whole blocks of an unaligned ciphertext must still decrypt");
            Require(TitleOfEncryptedDocument(key, o, u, id, "PADTOLERANCE", true)
                .StartsWith("PADTOLERANCE"),
                "illegal padding bytes must not cost the reader the plaintext in front of them");
            Console.WriteLine("[hostile] aes padding tolerated");
        }

        /// <summary>组装一份 AES-128 加密文档，把 /Title 写成填充有缺陷的密文串，返回读取端解密出的标题。
        /// Assembles an AES-128 document whose /Title is a string with defective padding and returns the title
        /// the reader decrypts from it.</summary>
        /// <param name="key">文件密钥。The file key.</param>
        /// <param name="o">/O 值。The /O value.</param>
        /// <param name="u">/U 值。The /U value.</param>
        /// <param name="id">trailer /ID 首元素。The first /ID element.</param>
        /// <param name="plain">标题明文。The plaintext title.</param>
        /// <param name="badPadding">true 写入非法填充字节，false 剪掉密文尾部使其不再对齐。
        /// When true the padding bytes are illegal; when false the ciphertext tail is cut so it no longer
        /// aligns to a block.</param>
        /// <returns>解密后的标题。The decrypted title.</returns>
        private static string TitleOfEncryptedDocument(byte[] key, byte[] o, byte[] u, byte[] id,
            string plain, bool badPadding)
        {
            byte[] aesKey = AesObjectKey(key, 5, 0);
            byte[] iv = Md5(Encoding.ASCII.GetBytes("initial-vector"));
            byte[] data = Encoding.ASCII.GetBytes(plain);
            // 明文至少要跨两个密集块：裁掉尾部字节后仍然留有可解密的整块。
            // The plaintext has to span at least two blocks, so cutting the tail still leaves whole blocks.
            int padded = ((data.Length + 15) / 16) * 16 + 16;
            byte[] block = new byte[padded];
            Buffer.BlockCopy(data, 0, block, 0, data.Length);
            if (badPadding)
            {
                for (int i = data.Length; i < padded; i++) block[i] = 153;
                // 只有最后一个字节像样合法：严格的 PKCS#7 校验在这里必然失败。
                // Only the final byte looks legal, so strict PKCS#7 validation fails right here.
                block[padded - 1] = 4;
            }
            byte[] cipher = Concat(iv, AesRaw(aesKey, iv, block));
            if (!badPadding) cipher = Slice(cipher, 0, cipher.Length - 3);

            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R /Info 5 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << >> >>",
                "<< /Filter /Standard /V 4 /R 4 /Length 128 /P -39 /O <" + ToHex(o) + "> /U <" + ToHex(u) + ">"
                    + " /CF << /StdCF << /CFM /AESV2 /Length 16 /AuthEvent /DocOpen >> >>"
                    + " /StmF /StdCF /StrF /StdCF >>",
                "<< /Title <" + ToHex(cipher) + "> >>"
            };
            byte[] pdf = AssembleClassicPdf(bodies, 1, null, "/Encrypt 4 0 R /ID [<" + ToHex(id)
                + "> <" + ToHex(id) + ">] ");

            PdfFile file = PdfFile.Load(pdf, "pw");
            PdfDictionary catalog = file.Resolve(file.Trailer.Get("Root")) as PdfDictionary;
            Require(catalog != null, "the crafted document must expose its catalog");
            PdfDictionary info = file.Resolve(catalog.Get("Info")) as PdfDictionary;
            Require(info != null, "the crafted document must expose its /Info dictionary through the catalog");
            PdfString title = file.Resolve(info.Get("Title")) as PdfString;
            Require(title != null, "the /Title entry must be a string");
            return Encoding.ASCII.GetString(title.Bytes);
        }

        /// <summary>按算法 1 推出某个对象的 AES-128 子密钥（追加 sAlT）。
        /// Derives the per-object AES-128 subkey through algorithm 1, appending the sAlT suffix.</summary>
        /// <param name="fileKey">文件密钥。The file key.</param>
        /// <param name="number">对象号。The object number.</param>
        /// <param name="generation">代次。The generation number.</param>
        /// <returns>16 字节子密钥。The 16-byte subkey.</returns>
        private static byte[] AesObjectKey(byte[] fileKey, int number, int generation)
        {
            return Slice(Md5(Concat(Slice(fileKey, 0, 16), new byte[]
            {
                (byte)number,
                (byte)(number >> 8),
                (byte)(number >> 16),
                (byte)generation,
                (byte)(generation >> 8),
                (byte)115,
                (byte)65,
                (byte)108,
                (byte)84
            })), 0, 16);
        }

        /// <summary>无填充的 AES-CBC 加密，用来精确控制密文末尾内容。
        /// AES-CBC encryption without padding, so the ciphertext tail is exactly what the test writes.</summary>
        /// <param name="key">子密钥。The subkey.</param>
        /// <param name="iv">初始向量。The initialization vector.</param>
        /// <param name="data">已对齐的明文。The aligned plaintext.</param>
        /// <returns>密文。The ciphertext.</returns>
        private static byte[] AesRaw(byte[] key, byte[] iv, byte[] data)
        {
            using (System.Security.Cryptography.Aes aes = System.Security.Cryptography.Aes.Create())
            {
                aes.Key = key;
                aes.Mode = System.Security.Cryptography.CipherMode.CBC;
                aes.Padding = System.Security.Cryptography.PaddingMode.None;
                aes.IV = iv;
                using (System.Security.Cryptography.ICryptoTransform transform = aes.CreateEncryptor())
                {
                    return transform.TransformFinalBlock(data, 0, data.Length);
                }
            }
        }

        /// <summary>把字节写成 PDF 十六进制串的内容。Writes bytes as the body of a PDF hex string.</summary>
        /// <param name="data">字节。The bytes.</param>
        /// <returns>大写十六进制。Upper-case hexadecimal.</returns>
        private static string ToHex(byte[] data)
        {
            StringBuilder builder = new StringBuilder(data.Length * 2);
            for (int i = 0; i < data.Length; i++) builder.Append(((int)data[i]).ToString("X2"));
            return builder.ToString();
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

        /// <summary>
        /// 十几万个左方括号构成的语法对象：数组与字典靠自递归下降，栈溢出不可捕获，会带走宿主进程，
        /// 因此深度必须封顶成可捕获的读取失败。
        /// A syntactic object built from a hundred thousand opening brackets: arrays and dictionaries descend by
        /// recursion, and a stack overflow cannot be caught — it takes the host process with it — so the depth
        /// has to be capped into a catchable read failure.
        /// </summary>
        private static void RunSyntaxNestingBomb()
        {
            int depth = 100000;
            string nested = new string('[', depth) + new string(']', depth);
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [] /Count 0 /Probe " + nested + " >>"
            };
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
            Require(watch.ElapsedMilliseconds < 15000, "nested-bracket object took " + watch.ElapsedMilliseconds + " ms");
            Require(!(failure is OutOfMemoryException), "nested-bracket object exhausted memory: " + failure);
            Require(failure == null || failure is InvalidDataException,
                "nested-bracket object failed with " + (failure == null ? "none" : failure.GetType().Name + ": " + failure.Message));
            Require(pageCount <= 0, "nested-bracket object reported " + pageCount + " pages");
            Console.WriteLine("[hostile] syntactic nesting capped, failure="
                + (failure == null ? "none, pages=" + pageCount.ToString(CultureInfo.InvariantCulture)
                                    : failure.GetType().Name));
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
