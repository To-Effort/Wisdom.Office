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
    }
}
