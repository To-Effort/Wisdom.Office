using System;
using System.IO;
using System.Text;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 2 回归：单个畸形元素不得毁掉整次读取。
    /// Remediation design episode two, phase two: one malformed element must not sink a whole read.
    /// </summary>
    /// <remarks>
    /// 崩溃类里可被捕获的那部分（空引用、越界）留在主运行器断言；<c>StackOverflowException</c> 一类
    /// 会带走进程的条目按规范 §2 走子进程取证，不在这里跑。
    /// Catchable crashers (null references, index overruns) are asserted in the main runner; entries that
    /// kill the process, such as <c>StackOverflowException</c>, are proven in a child process per §2 of the
    /// design and stay out of this file.
    /// </remarks>
    internal static partial class Program
    {
        /// <summary>Phase 2 入口。Entry point for phase two.</summary>
        /// <param name="pdfPath">样例 PDF。The sample PDF.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunRemediationPhaseTwoTest(string pdfPath, string outputDirectory)
        {
            RunPagesNodeWithoutKidsTest();
            RunObjectStreamShortHeaderTest();

            Console.WriteLine("[remediation] phase-two malformed page tree and object stream ok");
        }

        /// <summary>PD1：声明 /Type /Pages 却没有可用 /Kids 的节点不得抛空引用，按零页降级。PD1: a node declaring /Type /Pages without a usable /Kids must not throw a null reference; it degrades to zero pages.</summary>
        private static void RunPagesNodeWithoutKidsTest()
        {
            byte[] noKids = AssembleClassicPdf(new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Count 0 >>",
            }, 1, null);

            PdfDocument document = PdfDocument.Load(noKids);
            Require(document.PageCount == 0, "a Pages node without /Kids should yield zero pages, got " + document.PageCount);

            byte[] danglingKids = AssembleClassicPdf(new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Count 1 /Kids 9 0 R >>",
            }, 1, null);

            PdfDocument again = PdfDocument.Load(danglingKids);
            Require(again.PageCount == 0, "a /Kids reference to a missing object should yield zero pages, got " + again.PageCount);
        }

        /// <summary>PD2：对象流声明的 /N 大于实际配对数时，已经命中的偏移必须留下，不能整条流作废。PD2: when an object stream declares a larger /N than the pairs it actually carries, offsets already matched must survive instead of voiding the whole stream.</summary>
        private static void RunObjectStreamShortHeaderTest()
        {
            string catalog = "<< /Type /Catalog /Pages 5 0 R >>";
            string pages = "<< /Type /Pages /Kids [2 0 R] /Count 1 >>";
            string header = "4 0 5 " + catalog.Length + " ";
            string body = header + catalog + pages;
            string stream = "<< /Type /ObjStm /N 3 /First " + header.Length + " /Length " +
                Encoding.ASCII.GetByteCount(body) + " >>\nstream\n" + body + "\nendstream";
            string page = "<< /Type /Page /Parent 5 0 R /MediaBox [0 0 595 842] >>";

            string[] bodies = new string[] { stream, page };
            long[] offsets = ObjectOffsets(bodies);
            byte[] rows = new byte[]
            {
                0, 0, 0, 0, 0, 0, 0,
                1, (byte)(offsets[0] >> 24), (byte)(offsets[0] >> 16), (byte)(offsets[0] >> 8), (byte)offsets[0], 0, 0,
                1, (byte)(offsets[1] >> 24), (byte)(offsets[1] >> 16), (byte)(offsets[1] >> 8), (byte)offsets[1], 0, 0,
                0, 0, 0, 0, 0, 0, 0,
                2, 0, 0, 0, 1, 0, 0,
                2, 0, 0, 0, 1, 0, 1
            };
            byte[] pdf = AssembleXrefStreamPdf(bodies, rows, 4, 6L, 0L, 6L);

            Require(PdfDocument.Load(pdf).PageCount == 1, "a short object-stream header dropped the catalog");
            Require(PdfFile.Load(pdf).PageCount == 1, "a short object-stream header dropped the catalog in the object-level reader");
        }
    }
}
