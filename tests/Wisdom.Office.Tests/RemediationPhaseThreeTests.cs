using System;
using System.Diagnostics;
using System.Text;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 3 回归：单页提取路径上的二次方热点必须退回近线性。
    /// Remediation design episode two, phase three: the quadratic hot spots on the single-page extraction
    /// path must fall back to near-linear cost.
    /// </summary>
    /// <remarks>
    /// 这一相的挂起证据仍按规范 §2 走子进程；主运行器只放「中等规模 + 时间上限」这一档，
    /// 阈值取修复后实测的数倍余量，避免机器抖动造成假红。
    /// Hang-class evidence still goes through a child process per §2 of the design; the main runner carries
    /// only the "medium size plus a time ceiling" tier, and the threshold leaves several times the measured
    /// post-fix cost so machine jitter cannot fake a red.
    /// </remarks>
    internal static partial class Program
    {
        /// <summary>Phase 3 入口。Entry point for phase three.</summary>
        private static void RunRemediationPhaseThreeTest()
        {
            RunFragmentConcatQuadraticTest();

            Console.WriteLine("[remediation] phase-three quadratic hot spots ok");
        }

        /// <summary>PD3：一行二十万字元不得按片段拼接的二次方成本提取；实测修复前 7 622 ms、修复后 116 ms。PD3: 200 000 glyphs on one line must not be extracted at the fragment-concatenation quadratic cost; measured 7 622 ms before the fix and 116 ms after.</summary>
        private static void RunFragmentConcatQuadraticTest()
        {
            const int Glyphs = 200000;

            StringBuilder body = new StringBuilder("BT /F1 12 Tf 20 700 Td ");
            for (int i = 0; i < Glyphs; i++) body.Append("(A) Tj ");
            body.Append("ET");
            string streamBody = body.ToString();

            string[] objects = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                    + "/Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Length " + streamBody.Length + " >>\nstream\n" + streamBody + "\nendstream",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            byte[] pdf = AssembleClassicPdf(objects, 1, null);

            Stopwatch watch = Stopwatch.StartNew();
            string text = PdfDocument.Load(pdf).ExtractText();
            watch.Stop();

            Console.WriteLine("[remediation] PD3 extracted " + (text == null ? 0 : text.Length)
                + " chars from " + Glyphs + " glyphs in " + watch.ElapsedMilliseconds + " ms");

            Require(text != null && text.Length >= Glyphs,
                "one-line glyph run lost text, got " + (text == null ? 0 : text.Length) + " chars");
            Require(watch.ElapsedMilliseconds < 1200,
                Glyphs + " glyphs on one line took " + watch.ElapsedMilliseconds + " ms to extract");
        }
    }
}
