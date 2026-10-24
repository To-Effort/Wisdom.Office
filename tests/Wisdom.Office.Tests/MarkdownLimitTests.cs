using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using Wisdom.Office.Core;
using Wisdom.Markdown;
using Md = Wisdom.Markdown.Markdown;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// Markdown 读入面的上限与回退回归测试：表格的列数与单元格总数、空删除线、重复引用的图片。
    /// Regression tests for the ceilings and fallbacks on the Markdown input surface: table columns and total
    /// cells, an empty strikeout, and an image referenced over and over.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>运行 Markdown 上限与回退回归测试。Runs the Markdown ceiling and fallback regression tests.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunMarkdownLimitTest(string outputDirectory)
        {
            RunTableColumnLimitTest();
            RunTableCellLimitTest();
            RunEmptyStrikethroughTest();
            RunRepeatedImageTest(outputDirectory);
            Console.WriteLine("[mdlimit] markdown ceilings and fallbacks ok");
        }

        /// <summary>
        /// 一行两万个竖线：列数该在物化单元格之前就被拦下，而不是先拆成两万个字符串再说话。
        /// A line with twenty thousand pipes: the column count has to stop it before the cells are materialised,
        /// not after twenty thousand strings already exist.
        /// </summary>
        private static void RunTableColumnLimitTest()
        {
            string wide = new string('|', 20000) + "\n" + new string('-', 20000) + "|\n|a|b|\n";
            Exception failure = null;
            Stopwatch watch = Stopwatch.StartNew();
            try
            {
                Md.ToDocument(wide);
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            watch.Stop();
            Require(failure is InvalidDataException, "a 20 000-column table failed with " + failure);
            Require(watch.ElapsedMilliseconds < 3000, "the wide table needed " + watch.ElapsedMilliseconds + " ms");

            OfficeDocument legal = Md.ToDocument("| a | b | c |\n| --- | --- | --- |\n| 1 | 2 | 3 |\n");
            Table table = FindFirstTable(legal.Sections[0].Blocks);
            Require(table != null && table.Columns.Count == 3, "a three-column table changed shape: "
                + (table == null ? "missing" : table.Columns.Count.ToString()));

            // 上限之下的宽表照常：4094 列（4095 个竖线）仍在许可范围内。
            // A wide-but-legal table still parses: 4 094 columns (4 095 pipes) stays inside the ceiling.
            OfficeDocument edge = Md.ToDocument(new string('|', 4095) + "\n" + new string('-', 4095) + "|\n");
            Table edgeTable = FindFirstTable(edge.Sections[0].Blocks);
            Require(edgeTable != null && edgeTable.Columns.Count == 4094, "the legal wide table came out as "
                + (edgeTable == null ? "missing" : edgeTable.Columns.Count.ToString()));
        }

        /// <summary>
        /// 列数合规而行数无限：4096 列以内照样能靠行数把格子乘出来，所以总数单独封顶。
        /// Legal columns over an endless number of rows: within the column cap the product still grows, so the
        /// total carries its own ceiling.
        /// </summary>
        private static void RunTableCellLimitTest()
        {
            Exception failure = null;
            Stopwatch watch = Stopwatch.StartNew();
            try
            {
                Md.ToDocument(BuildTable(200, 600));
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            watch.Stop();
            Require(failure is InvalidDataException, "a 200x600 table failed with " + failure);
            Require(watch.ElapsedMilliseconds < 30000, "the oversized table needed "
                + watch.ElapsedMilliseconds + " ms");

            // 合规边界仍在：100 列 100 行共一万个格子照常解析。
            // The legal edge stays intact: 100 columns by 100 rows, ten thousand cells, still parses.
            Table legal = FindFirstTable(Md.ToDocument(BuildTable(100, 100)).Sections[0].Blocks);
            Require(legal != null && legal.Rows.Count == 101 && legal.Columns.Count == 100,
                "a 100x100 table changed shape: " + (legal == null ? "missing" : legal.Rows.Count + "x" + legal.Columns.Count));
        }

        /// <summary>
        /// 空的删除线不该吞掉文本：只跳过两个波浪号时 "~~~~" 的四个全部消失。
        /// An empty strikeout must not swallow text: skipping only two tildes made all four of "~~~~" disappear.
        /// </summary>
        private static void RunEmptyStrikethroughTest()
        {
            string alone = Md.ToDocument("甲 ~~~~ 乙").GetPlainText().Trim();
            Require(alone == "甲 ~~~~ 乙", "the empty strike deleted its own tildes: '" + alone + "'");
            string doubled = Md.ToDocument("前 ~~文字~~~~后").GetPlainText().Trim();
            Require(doubled.Contains("~~") && doubled.Contains("文字"),
                "a strike next to an empty one lost its tildes: '" + doubled + "'");
            Require(Md.ToDocument("~~gone~~").GetPlainText().Contains("gone"), "a real strike lost its content");

            OfficeDocument document = Md.ToDocument("~~gone~~");
            bool struck = false;
            foreach (Block block in document.Sections[0].Blocks)
            {
                Paragraph paragraph = block as Paragraph;
                if (paragraph == null) continue;
                for (int i = 0; i < paragraph.Inlines.Count; i++)
                {
                    TextRun run = paragraph.Inlines[i] as TextRun;
                    if (run != null && run.Format != null && run.Format.Strikeout) struck = true;
                }
            }
            Require(struck, "a real strike no longer marks the text struck out");
        }

        /// <summary>
        /// 同一张图引用两百次：按路径缓存后只能有一份字节，否则两百幅整图全留在模型里。
        /// Two hundred references to one picture: cached by path there is a single copy of the bytes, otherwise
        /// the model carries two hundred full images.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunRepeatedImageTest(string outputDirectory)
        {
            string directory = Path.Combine(outputDirectory, "md-images");
            Directory.CreateDirectory(directory);
            File.WriteAllBytes(Path.Combine(directory, "dot.png"), CreateTestPng(8, 8));

            MarkdownOptions options = new MarkdownOptions();
            options.ImageBaseDirectory = directory;

            StringBuilder references = new StringBuilder();
            for (int i = 0; i < 200; i++) references.Append("![图 " + i + "](dot.png)\n\n");
            OfficeDocument document = Md.ToDocument(references.ToString(), options);

            List<OfficeImage> images = new List<OfficeImage>();
            foreach (Block block in document.Sections[0].Blocks)
            {
                Paragraph paragraph = block as Paragraph;
                if (paragraph == null) continue;
                for (int i = 0; i < paragraph.Inlines.Count; i++)
                {
                    ImageRun run = paragraph.Inlines[i] as ImageRun;
                    if (run != null && run.Image != null) images.Add(run.Image);
                }
            }
            Require(images.Count == 200, "expected 200 image runs but found " + images.Count);
            for (int i = 1; i < images.Count; i++)
            {
                Require(ReferenceEquals(images[0], images[i]),
                    "reference " + i + " loaded its own copy of the picture");
            }

            Require(images[0].PixelWidth == 8 && images[0].PixelHeight == 8, "the cached picture is unreadable");
        }

        /// <summary>
        /// 造一张 columns 列、rows 行的管道表格，用于验证列数与单元格总数的两处上限。
        /// Builds a pipe table of the given columns and rows, used to check the column and total-cell ceilings.
        /// </summary>
        /// <param name="columns">列数。The column count.</param>
        /// <param name="rows">数据行数。The data row count.</param>
        /// <returns>Markdown 文本。The Markdown text.</returns>
        private static string BuildTable(int columns, int rows)
        {
            StringBuilder header = new StringBuilder();
            StringBuilder divider = new StringBuilder();
            for (int c = 0; c < columns; c++)
            {
                header.Append("| c" + c);
                divider.Append("| ---");
            }
            header.Append(" |\n");
            divider.Append(" |\n");

            StringBuilder body = new StringBuilder();
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++) body.Append("| r" + r + "_" + c);
                body.Append(" |\n");
            }
            return header.ToString() + divider + body.ToString();
        }
    }
}
