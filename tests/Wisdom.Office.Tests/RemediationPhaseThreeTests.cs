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
            RunLineMergeScalingTest();

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

        /// <summary>PD4：行归并对每个片段扫全部已有行，行数翻倍的成本必须近线性。PD4: the line merge may not scan every existing line per fragment; doubling the line count must keep the cost near-linear.</summary>
        private static void RunLineMergeScalingTest()
        {
            const int Small = 20000;
            const int Large = 40000;

            byte[] smallPdf = BuildMultiLinePdf(Small);
            byte[] largePdf = BuildMultiLinePdf(Large);

            Stopwatch watch = Stopwatch.StartNew();
            int smallLines = PdfDocument.Load(smallPdf).Pages[0].GetTextLines().Count;
            watch.Stop();
            long smallMs = watch.ElapsedMilliseconds;

            watch = Stopwatch.StartNew();
            int largeLines = PdfDocument.Load(largePdf).Pages[0].GetTextLines().Count;
            watch.Stop();
            long largeMs = watch.ElapsedMilliseconds;

            Console.WriteLine("[remediation] PD4 lines=" + smallLines + "/" + largeLines
                + " ms=" + smallMs + "/" + largeMs);

            Require(smallLines == Small, Small + " separate baselines should yield " + Small
                + " lines, got " + smallLines);
            Require(largeLines == Large, Large + " separate baselines should yield " + Large
                + " lines, got " + largeLines);
            Require(largeMs < smallMs * 3 + 200, "doubling " + Small + " lines to " + Large
                + " multiplied extraction from " + smallMs + " ms to " + largeMs + " ms");
        }

        /// <summary>拼一份单页 PDF，每个片段独占一条基线：行距 30，片段高度约 14，任何两片段都不重叠。Builds a one-page PDF where every fragment owns its baseline: 30 units of leading against roughly 14 units of height, so no two fragments overlap.</summary>
        /// <param name="lines">行数，即片段数。The line count, which is also the fragment count.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildMultiLinePdf(int lines)
        {
            StringBuilder body = new StringBuilder(lines * 40);
            for (int i = 0; i < lines; i++)
            {
                body.Append("BT /F1 12 Tf 20 ").Append(20 + i * 30)
                    .Append(" Td (A) Tj ET\n");
            }
            string streamBody = body.ToString();

            string[] objects = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 " + (40 + lines * 30) + "] "
                    + "/Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Length " + streamBody.Length + " >>\nstream\n" + streamBody + "\nendstream",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            return AssembleClassicPdf(objects, 1, null);
        }
    }
}
