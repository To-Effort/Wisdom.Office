using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wisdom.Office.Core;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 4 回归：产出损坏文件与静默丢失。
    /// Remediation design episode two, phase four: corrupt output and silently lost content.
    /// </summary>
    /// <remarks>
    /// 需要告警通道的条目（<c>PD9</c> 与其下游 <c>PD5</c>/<c>PD22</c>）另批设计，这里只放可单独判定的条目。
    /// Entries that need the warning channel (<c>PD9</c> and its downstream <c>PD5</c>/<c>PD22</c>) are
    /// designed as their own batch; this file carries only what can be decided on its own.
    /// </remarks>
    internal static partial class Program
    {
        /// <summary>Phase 4 入口。Entry point for phase four.</summary>
        private static void RunRemediationPhaseFourTest()
        {
            RunPredictorSingleRowTest();
            RunCatalogGenerationTest();
            RunOffsetNumberMismatchTest();

            Console.WriteLine("[remediation] phase-four corrupt output and silent loss ok");
        }

        /// <summary>PD16：PNG 预测器的正文恰好等于一行长度时不得解出零字节，内容流照常可用。PD16: when a PNG-predicted body is exactly one row long the decoder must not return zero bytes; the content stream still has to work.</summary>
        private static void RunPredictorSingleRowTest()
        {
            string content = "BT /F1 12 Tf 20 700 Td (Hi) Tj ET";
            byte[] raw = Encoding.ASCII.GetBytes(content);
            byte[] packed = Zlib.Compress(raw);

            // /Columns 取成正文长度，于是「一整行」就是整段正文，行首的滤波字节永远读不到。
            // With /Columns set to the body length, a whole row spans the body, so the leading filter
            // byte can never be read.
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                    + "/Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Filter /FlateDecode /Length " + packed.Length
                    + " /DecodeParms << /Predictor 12 /Colors 1 /BitsPerComponent 8 /Columns "
                    + raw.Length + " >> >>\nstream\n",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            byte[] pdf = AssembleBinaryPdf(bodies, 3, packed, "\nendstream");

            string text = PdfDocument.Load(pdf).ExtractText();
            Require(text != null && text.Contains("Hi"),
                "a predictor body of exactly one row lost the content stream, got '" + text + "'");
        }

        /// <summary>
        /// PD17：交叉表条目指到的对象头写着别的号，这个条目即失效，不得把住在该偏移上的对象当成目标号返回。
        /// PD17: when the object header at a cross-reference row's offset declares a different number, the row
        /// is void and the object living at that offset must never be handed back under the requested number.
        /// </summary>
        private static void RunOffsetNumberMismatchTest()
        {
            string content = "BT /F1 12 Tf 20 700 Td (Hi) Tj ET";
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                    + "/Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };

            // 5 号条目改指 3 号对象的头，两条落在同一个偏移上。
            // Row 5 is redirected to object 3's header, so two rows share one offset.
            long[] offsets = ObjectOffsets(bodies);
            long[] forged = new long[bodies.Length];
            for (int i = 0; i < bodies.Length; i++) forged[i] = offsets[i];
            forged[4] = offsets[2];
            byte[] pdf = AssembleClassicPdf(bodies, 1, forged);

            PdfFile file = PdfFile.Load(pdf);
            PdfObject five = file.GetObject(5);
            Require(!IsPageDictionary(five),
                "the xref row for object 5 points at object 3's header and the reader handed that object back: "
                    + Describe(five));

            string text = PdfDocument.Load(pdf).ExtractText();
            Require(text != null && text.Contains("Hi"),
                "reading through the void row lost the page text, got '" + text + "'");
        }

        /// <summary>判断对象是否是这张 fixture 里的页节点。Tells whether the object is the fixture's page node.</summary>
        /// <param name="value">待判定的对象。The object to judge.</param>
        /// <returns>是页节点返回 true。True for the page dictionary.</returns>
        private static bool IsPageDictionary(PdfObject value)
        {
            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary == null) return false;
            return dictionary.Get("Type") != null && dictionary.Get("MediaBox") != null;
        }

        /// <summary>把对象缩成一句可诊断的描述。Condenses an object into a diagnosable one-liner.</summary>
        /// <param name="value">待描述的对象。The object to describe.</param>
        /// <returns>描述文本。The description.</returns>
        private static string Describe(PdfObject value)
        {
            if (value == null) return "null";
            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary == null) return value.GetType().Name;
            StringBuilder builder = new StringBuilder("PdfDictionary");
            foreach (KeyValuePair<string, PdfObject> pair in dictionary)
            {
                builder.Append(" /").Append(pair.Key).Append('=')
                    .Append(pair.Value == null ? "null" : pair.Value.GetType().Name);
            }
            return builder.ToString();
        }

        /// <summary>
        /// 拼装允许二进制流正文的经典交叉表 PDF：第 streamIndex 个对象的正文以原字节给出。
        /// Assembles a classic-xref PDF carrying a binary stream body: the object at streamIndex contributes
        /// its raw bytes instead of text.
        /// </summary>
        /// <param name="bodies">各对象的正文，流对象的正文以「字典 + stream 起始」结尾。The object bodies; the stream object's ends with its dictionary and the stream keyword.</param>
        /// <param name="streamIndex">流对象的下标。The index of the stream object.</param>
        /// <param name="streamBytes">流的原字节。The raw stream bytes.</param>
        /// <param name="streamTail">紧跟正文的收尾串。The terminator written right after the body.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] AssembleBinaryPdf(string[] bodies, int streamIndex, byte[] streamBytes, string streamTail)
        {
            Encoding ascii = Encoding.ASCII;
            List<byte[]> pieces = new List<byte[]>();
            List<long> offsets = new List<long>();
            long position = 0;

            Action<string> writeText = delegate(string text)
            {
                byte[] bytes = ascii.GetBytes(text);
                pieces.Add(bytes);
                position += bytes.Length;
            };

            writeText("%PDF-1.7\n");
            for (int i = 0; i < bodies.Length; i++)
            {
                offsets.Add(position);
                writeText((i + 1) + " 0 obj\n");
                writeText(bodies[i]);
                if (i == streamIndex)
                {
                    byte[] raw = streamBytes;
                    pieces.Add(raw);
                    position += raw.Length;
                    writeText(streamTail);
                }
                writeText("\nendobj\n");
            }

            long xrefAt = position;
            writeText("xref\n0 " + (bodies.Length + 1) + "\n0000000000 65535 f \n");
            for (int i = 0; i < bodies.Length; i++)
            {
                writeText(offsets[i].ToString("D10") + " 00000 n \n");
            }
            writeText("trailer\n<< /Size " + (bodies.Length + 1) + " /Root 1 0 R >>\nstartxref\n"
                + xrefAt + "\n%%EOF\n");

            using (MemoryStream memory = new MemoryStream())
            {
                for (int i = 0; i < pieces.Count; i++)
                {
                    memory.Write(pieces[i], 0, pieces[i].Length);
                }
                return memory.ToArray();
            }
        }

        /// <summary>
        /// PD10：目录对象住在非 0 世代时，保存写出的 trailer <c>/Root</c> 必须带上同一个世代，
        /// 不能只留对象号、把世代硬写成 0。
        /// PD10: when the catalog lives in a non-zero generation, the trailer /Root written on save must
        /// carry that same generation instead of hard-coding 0 next to the object number.
        /// </summary>
        private static void RunCatalogGenerationTest()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
            };
            int[] generations = new int[] { 2, 0, 0 };
            byte[] source = AssembleGenerationsPdf(bodies, generations, 1, 2);

            byte[] saved = PdfFile.Load(source).SaveToBytes();
            string text = Encoding.ASCII.GetString(saved);

            Require(text.Contains("1 2 obj"),
                "PD10 control failed: the saved file no longer writes the catalog header as '1 2 obj'");

            PdfReference root = PdfFile.Load(saved).Trailer.Get("Root") as PdfReference;
            Require(root != null, "PD10 control failed: the saved trailer carries no /Root reference");
            Require(root.Number == 1 && root.Generation == 2,
                "the trailer /Root is '" + root.Number + " " + root.Generation
                    + " R' while the object header is '1 2 obj'");
        }

        /// <summary>
        /// 拼装逐对象带世代号的经典交叉表 PDF，交叉表与 trailer 使用给定世代。
        /// Assembles a classic-xref PDF with a per-object generation; the cross-reference rows and the
        /// trailer /Root reference use the supplied generations.
        /// </summary>
        /// <param name="bodies">各对象正文，下标零起。The object bodies, zero-based.</param>
        /// <param name="generations">与正文逐位对应的世代号。The generation paired with each body.</param>
        /// <param name="rootNumber">trailer /Root 的对象号。The trailer /Root object number.</param>
        /// <param name="rootGeneration">trailer /Root 的世代号。The trailer /Root generation.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] AssembleGenerationsPdf(string[] bodies, int[] generations, int rootNumber, int rootGeneration)
        {
            Encoding ascii = Encoding.ASCII;
            List<byte[]> pieces = new List<byte[]>();
            List<long> offsets = new List<long>();
            long position = 0;

            Action<string> writeText = delegate(string text)
            {
                byte[] bytes = ascii.GetBytes(text);
                pieces.Add(bytes);
                position += bytes.Length;
            };

            writeText("%PDF-1.7\n");
            for (int i = 0; i < bodies.Length; i++)
            {
                offsets.Add(position);
                writeText((i + 1) + " " + generations[i] + " obj\n");
                writeText(bodies[i]);
                writeText("\nendobj\n");
            }

            long xrefAt = position;
            writeText("xref\n0 " + (bodies.Length + 1) + "\n0000000000 65535 f \n");
            for (int i = 0; i < bodies.Length; i++)
            {
                writeText(offsets[i].ToString("D10") + " " + generations[i].ToString("D5") + " n \n");
            }
            writeText("trailer\n<< /Size " + (bodies.Length + 1)
                + " /Root " + rootNumber + " " + rootGeneration + " R >>\nstartxref\n"
                + xrefAt + "\n%%EOF\n");

            using (MemoryStream memory = new MemoryStream())
            {
                for (int i = 0; i < pieces.Count; i++)
                {
                    memory.Write(pieces[i], 0, pieces[i].Length);
                }
                return memory.ToArray();
            }
        }
    }
}
