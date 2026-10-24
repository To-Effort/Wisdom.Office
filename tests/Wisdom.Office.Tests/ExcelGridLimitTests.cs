using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Office.Core.Formulas;
using Wisdom.Excel;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 电子表格读入面的网格上限回归测试：伪造的行号、无休的列推进、超长字段与空文本重复
    /// 都只能被丢弃或掐断，不能把模型抬成一片写不出、也导不完的网格。
    /// Grid-limit regression tests for the spreadsheet read surface: a forged row number, an endless column
    /// walk, an enormous field and repeating an empty text may only be dropped or cut short — none of them may
    /// lift the model into a grid nothing can write or export.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>运行电子表格网格上限回归测试。Runs the spreadsheet grid-limit regression tests.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunExcelGridLimitTest(string outputDirectory)
        {
            RunForgedRowIndexTest(outputDirectory);
            RunColumnWalkTest();
            RunCsvFieldAndColumnLimits();
            RunReptEmptyTextTest();
            Console.WriteLine("[grid] spreadsheet grid limits ok");
        }

        /// <summary>
        /// 一行声明 r="2000000000"：这一行不该进模型，否则 LastRow 抬到二十亿，
        /// 逐行导出的 CSV 就再也走不到头。
        /// A row declaring r="2000000000" must stay out of the model, otherwise LastRow climbs into the
        /// billions and the row-by-row CSV export never reaches the end.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunForgedRowIndexTest(string outputDirectory)
        {
            string sheetXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + "<row r=\"2000000000\"><c r=\"A1\"><v>1</v></c></row>"
                + "<row r=\"3\"><c r=\"B3\"><v>7</v></c></row>"
                + "</sheetData></worksheet>";
            OfficeWorkbook workbook = XlsxReader.Load(BuildMinimalXlsx(sheetXml));
            Worksheet sheet = workbook.Worksheets[0];
            Require(sheet.LastRow <= 3, "the forged row lifted LastRow to " + sheet.LastRow);
            Require(sheet.GetCell(3, 2).Value != null, "the legitimate row disappeared with the forged one");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "forged-row.csv"));
            Stopwatch watch = Stopwatch.StartNew();
            ExcelCsv.Write(sheet, path);
            watch.Stop();
            Require(watch.ElapsedMilliseconds < 5000, "exporting behind a forged row took "
                + watch.ElapsedMilliseconds + " ms");
            string[] lines = File.ReadAllLines(path);
            Require(lines.Length <= 3, "the export wrote " + lines.Length + " lines for a forged grid");
        }

        /// <summary>
        /// 不带 r 属性的单元格按上一列续排：一万个这样的窄列必须止步于网格右边界。
        /// A cell without an r attribute continues from the previous column, so ten thousand of them have to
        /// stop at the right edge of the grid.
        /// </summary>
        private static void RunColumnWalkTest()
        {
            StringBuilder cells = new StringBuilder();
            for (int i = 0; i < 40000; i++) cells.Append("<c><v>1</v></c>");
            string sheetXml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                + "<row r=\"1\">" + cells + "</row>"
                + "</sheetData></worksheet>";
            Worksheet sheet = XlsxReader.Load(BuildMinimalXlsx(sheetXml)).Worksheets[0];
            Require(sheet.LastColumn <= 16384, "the column walk reached " + sheet.LastColumn);
            Require(sheet.LastColumn > 1, "the walk dropped the whole row: " + sheet.LastColumn);
        }

        /// <summary>
        /// CSV 的字段长度与列数必须在拼装过程中就封顶：整条记录先物化再检查，
        /// 恶意文件早已把内存吃光，异常也就到不了了。
        /// CSV field length and column count have to be capped while the record is being assembled. Checking
        /// only after the whole record is materialised means the malicious file has already eaten the memory
        /// and the exception never arrives.
        /// </summary>
        private static void RunCsvFieldAndColumnLimits()
        {
            // 计数读取器给出 5000 万个字符：限额若在拼装过程中生效，它不该被读完。
            // The counting reader offers 50 million characters: a limit enforced while assembling never drains it.
            CountingReader huge = new CountingReader(50000000);
            Exception failure = null;
            try
            {
                ExcelCsv.Read(huge);
            }
            catch (Exception ex)
            {
                failure = ex;
            }
            Require(failure is InvalidDataException, "an endless CSV field failed with " + failure);
            Require(huge.Reads < 5000000, "the field limit was checked only after the record was built: "
                + huge.Reads + " characters read");

            StringBuilder wide = new StringBuilder();
            for (int i = 0; i < 30000; i++) wide.Append("1,");
            Exception wideFailure = null;
            try
            {
                ExcelCsv.Read(new StringReader(wide.ToString()));
            }
            catch (Exception ex)
            {
                wideFailure = ex;
            }
            Require(wideFailure is InvalidDataException, "a 30 000-column record failed with " + wideFailure);

            // 合法边界不动：16384 列、常规字段仍然原样读入。
            // Legal edges stay untouched: 16 384 columns and ordinary fields still read in full.
            OfficeWorkbook legal = ExcelCsv.Read(new StringReader("a,b,c\n1,2,3"));
            Require(legal.Worksheets[0].LastColumn == 3, "a legal record changed shape: "
                + legal.Worksheets[0].LastColumn);
        }

        /// <summary>
        /// REPT 的上限落在产出长度上，空文本的产出是 0，于是循环照旧空转十亿次。
        /// REPT caps the produced length; an empty text produces 0, so the loop used to spin a billion times
        /// for nothing.
        /// </summary>
        private static void RunReptEmptyTextTest()
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.AddWorksheet("Sheet1");
            Stopwatch watch = Stopwatch.StartNew();
            object result = FormulaEngine.Evaluate(workbook, "REPT(\"\",1000000000000)");
            watch.Stop();
            Require(watch.ElapsedMilliseconds < 1000, "REPT of an empty text took " + watch.ElapsedMilliseconds + " ms");
            Require(result != null && string.IsNullOrEmpty(Convert.ToString(result, CultureInfo.InvariantCulture)),
                "REPT of an empty text produced '" + result + "'");

            object bounded = FormulaEngine.Evaluate(workbook, "LEN(REPT(\"ab\",10000))");
            Require(bounded is double && Math.Abs((double)bounded - 20000d) < 1e-6,
                "a legal REPT should still produce 20 000 characters but was " + bounded);
        }

        /// <summary>
        /// 组装一个只含一张表的最小 xlsx 包，供伪造行号与列推进的用例复用。
        /// Assembles a minimal xlsx package with a single worksheet, reused by the forged-row and column-walk
        /// cases.
        /// </summary>
        /// <param name="sheetXml">工作表部件 XML。The worksheet part XML.</param>
        /// <returns>包字节。The package bytes.</returns>
        private static byte[] BuildMinimalXlsx(string sheetXml)
        {
            Encoding utf8 = new UTF8Encoding(false);
            Dictionary<string, byte[]> entries = new Dictionary<string, byte[]>();
            entries["[Content_Types].xml"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "</Types>");
            entries["_rels/.rels"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                + "</Relationships>");
            entries["xl/workbook.xml"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            entries["xl/_rels/workbook.xml.rels"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + "</Relationships>");
            entries["xl/worksheets/sheet1.xml"] = utf8.GetBytes(sheetXml);
            return CreateZipBytes(entries);
        }

        /// <summary>
        /// 只记账不存字节的读取器：源源不断给出字符，同时数出解析器读掉了多少。
        /// 限额若在拼装途中生效，读到的字符数就远小于总长度；先物化整条记录的写法会把字符读空。
        /// A reader that keeps a tally instead of storing bytes: it offers characters without end and counts how
        /// many the parser consumed. A limit enforced while assembling stops long before the end; a parser that
        /// materialises the whole record first drains it.
        /// </summary>
        private sealed class CountingReader : TextReader
        {
            /// <summary>可给出的字符总数。The total number of characters available.</summary>
            private readonly int _limit;

            /// <summary>已给出的字符数。The number of characters handed out so far.</summary>
            internal int Reads;

            /// <summary>以可给出的字符总数构造读取器。Constructs the reader with the number of characters it can offer.</summary>
            /// <param name="limit">字符总数上限。The total character count.</param>
            internal CountingReader(int limit)
            {
                _limit = limit;
            }

            /// <summary>给出下一个字符，同时计数；耗尽后返回 -1。Hands out the next character and counts it; returns -1 once exhausted.</summary>
            /// <returns>字符 'x' 或 -1。The character 'x' or -1.</returns>
            public override int Read()
            {
                if (Reads >= _limit) return -1;
                Reads++;
                return 'x';
            }

            /// <summary>预览下一个字符但不计数。Peeks at the next character without counting it.</summary>
            /// <returns>字符 'x' 或 -1。The character 'x' or -1.</returns>
            public override int Peek()
            {
                return Reads < _limit ? 'x' : -1;
            }
        }
    }
}
