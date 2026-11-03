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
    }
}
