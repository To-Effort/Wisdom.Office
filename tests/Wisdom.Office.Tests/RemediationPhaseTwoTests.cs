using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Text;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Pdf;
using Wisdom.Word;

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
            RunBrokenGlyphCountFontTest(outputDirectory);
            RunOdtSpanProductBudgetTest();

            Console.WriteLine("[remediation] phase-two hostile structures and fonts ok");
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

        /// <summary>PD6：maxp.numGlyphs 为 0 的畸形字体不得在保存链路上抛越界，要走既定的降级路径。PD6: a malformed font whose maxp.numGlyphs is zero must not throw an index overrun on the save path; it has to take the documented degradation route.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunBrokenGlyphCountFontTest(string outputDirectory)
        {
            byte[] font = ReadSystemTrueTypeFont();
            Require(font != null, "no glyf TrueType font found for the broken glyph count test");
            Require(EraseMaxpGlyphCount(font), "the sampled font has no maxp table");

            FontResolver.Default.RegisterFont("RemediationBroken", false, false, font);

            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();
            Paragraph paragraph = new Paragraph();
            section.Blocks.Add(paragraph);
            paragraph.AddText("AAA", new TextFormat("RemediationBroken", 12));

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-broken-glyph-count.pdf"));
            Exception failure = null;
            try
            {
                new PdfRenderer().Render(document, path);
            }
            catch (Exception ex)
            {
                failure = ex;
            }

            Require(failure == null, "a font with maxp.numGlyphs erased failed the save path with " +
                (failure == null ? "none" : failure.GetType().Name + ": " + failure.Message));
            Require(File.Exists(path) && new FileInfo(path).Length > 0, "the degraded save produced no file");
        }

        /// <summary>取一台机器上任意一个 glyf 轮廓的 TrueType 字体字节。Reads the bytes of any glyf-outline TrueType font on this machine.</summary>
        /// <returns>字体文件字节，找不到返回 null。The font bytes, or null when none is present.</returns>
        private static byte[] ReadSystemTrueTypeFont()
        {
            string directory = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory)) return null;
            string[] candidates = Directory.GetFiles(directory, "*.ttf");
            for (int i = 0; i < candidates.Length; i++)
            {
                try
                {
                    byte[] data = File.ReadAllBytes(candidates[i]);
                    if (data.Length < 1024) continue;
                    bool sfnt = (data[0] == 0 && data[1] == 1 && data[2] == 0 && data[3] == 0) ||
                        (data[0] == (byte)'t' && data[1] == (byte)'r' && data[2] == (byte)'u' && data[3] == (byte)'e');
                    bool cff = data[0] == (byte)'O' && data[1] == (byte)'T' && data[2] == (byte)'T' && data[3] == (byte)'O';
                    if (sfnt && !cff) return data;
                }
                catch (IOException)
                {
                    continue;
                }
            }
            return null;
        }

        /// <summary>把字体里 maxp 表的 numGlyphs 改写为 0，模拟被截短或被改头的字形计数。Overwrites maxp.numGlyphs with zero to mimic a truncated or hand-edited glyph count.</summary>
        /// <param name="font">字体字节，就地修改。The font bytes, modified in place.</param>
        /// <returns>找到并改写成功返回 true。True when the table was found and rewritten.</returns>
        private static bool EraseMaxpGlyphCount(byte[] font)
        {
            int numTables = (font[4] << 8) | font[5];
            for (int t = 0; t < numTables; t++)
            {
                int entry = 12 + t * 16;
                if (entry + 16 > font.Length) return false;
                if (font[entry] != (byte)'m' || font[entry + 1] != (byte)'a' ||
                    font[entry + 2] != (byte)'x' || font[entry + 3] != (byte)'p') continue;
                int offset = (font[entry + 8] << 24) | (font[entry + 9] << 16) | (font[entry + 10] << 8) | font[entry + 11];
                if (offset + 6 > font.Length) return false;
                font[offset + 4] = 0;
                font[offset + 5] = 0;
                return true;
            }
            return false;
        }

        /// <summary>WE1：在套件中调用子进程探针，确保挂起类输入不会阻塞主运行器。WE1: invoke the child-process probe from within the suite so a hang-class input cannot block the main runner.</summary>
        private static void RunOdtSpanProductBudgetTest()
        {
            string exe = Environment.ProcessPath ?? "dotnet";
            ProcessStartInfo start = new ProcessStartInfo(exe, "--crash-probe WE1");
            start.UseShellExecute = false;
            start.RedirectStandardOutput = true;
            int code = -1;
            string output = string.Empty;
            using (Process child = Process.Start(start))
            {
                output = child.StandardOutput.ReadToEnd();
                if (!child.WaitForExit(15000))
                {
                    try { child.Kill(); } catch (Exception) { }
                    Require(false, "the WE1 child probe did not finish within 15 s");
                    return;
                }
                code = child.ExitCode;
            }
            Console.WriteLine("[remediation] WE1 child probe: " + (output ?? string.Empty).Trim());
            Require(code == 0, "the WE1 child probe exited with code " + code);
        }

        /// <summary>WE1 探针：一个跨行跨列乘积畸形的 ODT 单元格不得把加载拖成假死，由调用方用超时判定。WE1 probe: one ODT cell with a forged row-by-column span product must not turn loading into a hang; the caller decides red or green by timing it out.</summary>
        /// <returns>加载在预算内完成返回 true。True when the load finishes inside the slot budget.</returns>
        private static bool ProbeOdtSpanProduct()
        {
            byte[] odt = BuildOdtWithForgedSpans(1);
            Stopwatch watch = Stopwatch.StartNew();
            OfficeDocument document = OdtReader.Load(odt);
            watch.Stop();

            Table table = FindFirstTable(document);
            Console.WriteLine("[crash-probe] WE1 loaded a 4096x4096 span cell in " + watch.ElapsedMilliseconds
                + " ms, rows=" + (table == null ? -1 : table.Rows.Count)
                + ", cells=" + (table == null || table.Rows.Count == 0 ? -1 : table.Rows[0].Cells.Count)
                + ", colSpan=" + (table == null || table.Rows.Count == 0 ? -1 : table.Rows[0].Cells[0].ColSpan));
            if (table == null) return false;
            if (table.Rows.Count != 1) return false;
            if (table.Rows[0].Cells.Count != 1) return false;
            if (table.Rows[0].Cells[0].ColSpan != 4096) return false;
            return watch.ElapsedMilliseconds < 1500;
        }

        /// <summary>拼出只有一个表体行、若干 4096×4096 跨度单元格的 ODT 包。Builds an ODT package holding one row of 4096x4096-span cells.</summary>
        /// <param name="cells">单元格数量。The number of cells.</param>
        /// <returns>ODT 字节。The ODT bytes.</returns>
        private static byte[] BuildOdtWithForgedSpans(int cells)
        {
            StringBuilder body = new StringBuilder();
            for (int i = 0; i < cells; i++)
            {
                body.Append("<table:table-cell table:number-rows-spanned=\"4096\"")
                    .Append(" table:number-columns-spanned=\"4096\"><text:p>格")
                    .Append(i)
                    .Append("</text:p></table:table-cell>");
            }

            string content = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<office:document-content" +
                " xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\"" +
                " xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\"" +
                " xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\">" +
                "<office:body><office:text><table:table><table:table-row>" + body +
                "</table:table-row></table:table></office:text></office:body></office:document-content>";

            using (MemoryStream memory = new MemoryStream())
            {
                using (ZipArchive archive = new ZipArchive(memory, ZipArchiveMode.Create, true))
                {
                    ZipArchiveEntry entry = archive.CreateEntry("content.xml");
                    using (Stream stream = entry.Open())
                    {
                        byte[] bytes = Encoding.UTF8.GetBytes(content);
                        stream.Write(bytes, 0, bytes.Length);
                    }
                }
                return memory.ToArray();
            }
        }

        /// <summary>取统一模型里的第一张表。Returns the first table of the unified model.</summary>
        /// <param name="document">统一文档。The unified document.</param>
        /// <returns>表格或 null。The table, or null.</returns>
        private static Table FindFirstTable(OfficeDocument document)
        {
            for (int s = 0; s < document.Sections.Count; s++)
            {
                List<Block> blocks = document.Sections[s].Blocks;
                for (int b = 0; b < blocks.Count; b++)
                {
                    Table table = blocks[b] as Table;
                    if (table != null) return table;
                }
            }
            return null;
        }
    }
}
