using System;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using Wisdom.Excel;
using Wisdom.Office.Core;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// Excel 侧的续修条目：共享公式的表名前缀、越界日期序列、以及读表时把 <c>LastRow</c> 写进循环条件造成的平方级扫描。
    /// Excel entries of the follow-up round: sheet-qualified shared formulas, out-of-range date serials, and the
    /// quadratic scan caused by reading <c>LastRow</c> inside a loop condition.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>Excel 续修入口。Entry point for the Excel follow-up round.</summary>
        private static void RunRemediationExcelTest()
        {
            RunSharedFormulaSheetQualifiedTest();
            RunOutOfRangeDateSerialTest();
            RunCollectionLastRowSnapshotTest();

            Console.WriteLine("[remediation] excel shared formulas, date serials and row snapshot ok");
        }

        /// <summary>
        /// 共享公式里的相对引用如果带表名前缀（<c>Sheet1!A2</c>），从格必须按偏移改写它；
        /// 过去的「前一个字符是感叹号就跳过」规则让每个从格都停在主格那一行，公式看起来还在、算出的却是同一个值。
        /// A relative reference carrying a sheet prefix (<c>Sheet1!A2</c>) has to be rewritten per dependent.
        /// The old "previous character is a bang, skip" rule left every dependent on the master's row: the
        /// formula was still there, but it pointed at the master's cell.
        /// </summary>
        private static void RunSharedFormulaSheetQualifiedTest()
        {
            string sheetXml = "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetData>"
                + "<row r=\"2\">"
                + "<c r=\"A2\"><v>1</v></c>"
                + "<c r=\"B2\" t=\"str\"><f t=\"shared\" ref=\"B2:B4\" si=\"0\">Sheet1!A2+1</f><v>2</v></c>"
                + "</row>"
                + "<row r=\"3\">"
                + "<c r=\"A3\"><v>5</v></c>"
                + "<c r=\"B3\" t=\"str\"><f t=\"shared\" si=\"0\"/><v>6</v></c>"
                + "</row>"
                + "<row r=\"4\">"
                + "<c r=\"A4\"><v>9</v></c>"
                + "<c r=\"B4\" t=\"str\"><f t=\"shared\" si=\"0\"/><v>10</v></c>"
                + "</row>"
                + "</sheetData></worksheet>";
            Worksheet sheet = XlsxReader.Load(BuildFidelityPackage(sheetXml, null, null)).Worksheets[0];

            Require(sheet.GetCell(2, 2).Formula == "Sheet1!A2+1",
                "the sheet-qualified master changed: " + sheet.GetCell(2, 2).Formula);
            Require(sheet.GetCell(3, 2).Formula == "Sheet1!A3+1",
                "a sheet-qualified relative reference did not shift with the dependent: "
                + sheet.GetCell(3, 2).Formula);
            Require(sheet.GetCell(4, 2).Formula == "Sheet1!A4+1",
                "a two-row offset was not applied to the sheet-qualified reference: "
                + sheet.GetCell(4, 2).Formula);
        }

        /// <summary>
        /// 日期样式的单元格如果带一个超出 <see cref="DateTime"/> 范围的序列号，读端必须保留原始数值并继续加载；
        /// 过去 <c>epoch.AddDays(serial)</c> 直接抛 <see cref="ArgumentOutOfRangeException"/>，一个单元格废掉整本工作簿。
        /// A date-styled cell carrying a serial outside <see cref="DateTime"/>'s range must keep its raw number
        /// and let the load continue; <c>epoch.AddDays(serial)</c> used to throw
        /// <see cref="ArgumentOutOfRangeException"/> and one cell took the whole workbook down.
        /// </summary>
        private static void RunOutOfRangeDateSerialTest()
        {
            string stylesXml = "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
                + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
                + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"2\">"
                + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                + "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
                + "</cellXfs>";
            string sheetXml = "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetData><row r=\"1\">"
                + "<c r=\"A1\" s=\"1\"><v>1E+10</v></c>"
                + "<c r=\"B1\" s=\"1\"><v>45000.5</v></c>"
                + "</row></sheetData></worksheet>";

            Worksheet sheet = XlsxReader.Load(BuildFidelityPackage(sheetXml, stylesXml, null)).Worksheets[0];

            object broken = sheet.GetCell(1, 1).Value;
            Require(broken is double,
                "an out-of-range date serial should keep its raw number, got "
                + (broken == null ? "null" : broken.GetType().Name));
            Require(Math.Abs((double)broken - 1e10) < 1.0,
                "the raw serial changed while being rejected: " + broken);

            // 控制：同一份包里的正常日期序列仍要照常转成 DateTime，别把整条转换路径一起关掉。
            // Control: the ordinary serial in the same package still becomes a DateTime, so the conversion
            // path was not disabled wholesale.
            object ok = sheet.GetCell(1, 2).Value;
            Require(ok is DateTime, "an in-range date serial stopped converting: "
                + (ok == null ? "null" : ok.GetType().Name));
        }

        /// <summary>
        /// <c>Worksheet.LastRow</c> 每次访问都要遍历全表单元格，写在循环条件里就是「行数 × 单元格数」的平方级扫描；
        /// 这里用 8000 行做阈值断言，并把实测毫秒打进标记行。
        /// <c>Worksheet.LastRow</c> walks every cell on each access, so using it as a loop condition is a
        /// rows-by-cells quadratic scan. 8000 rows are bounded by a threshold, and the measured milliseconds
        /// go into the marker line.
        /// </summary>
        private static void RunCollectionLastRowSnapshotTest()
        {
            const int rows = 8000;
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Rows");
            for (int r = 1; r <= rows; r++)
            {
                sheet.GetCell(r, 1).Value = r;
                sheet.GetCell(r, 2).Value = "v" + r.ToString(CultureInfo.InvariantCulture);
            }

            Stopwatch watch = Stopwatch.StartNew();
            System.Data.DataTable table = ExcelCollection.ToDataTable(sheet);
            watch.Stop();

            Require(table.Rows.Count == rows - 1,
                "the header row should be consumed, leaving " + (rows - 1) + " rows, got " + table.Rows.Count);
            Require(watch.ElapsedMilliseconds < 600,
                "reading " + rows + " rows took " + watch.ElapsedMilliseconds
                + " ms; LastRow is still being evaluated inside the loop");
        }
    }
}
