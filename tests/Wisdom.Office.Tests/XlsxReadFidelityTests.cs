using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Wisdom.Excel;
using Wisdom.Office.Core;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>
        /// xlsx 读端保真回归：主题调色板（含索引错位的 theme 与 tint）、无 state 的窗格不再当成冻结、
        /// 纯空白文本不再当作元素间空白丢失、多区域 sqref 不再静默截断、dxf 的边框与数字格式可读回、
        /// 共享公式的从格可还原、1904 日期系统不再在写出时被改成 1900。
        /// Read-side fidelity regressions for xlsx: the theme palette (including the swizzled theme indices and
        /// tint), panes without a state that are no longer reported as frozen, whitespace-only text that is no
        /// longer dropped as inter-element whitespace, multi-range sqref that is no longer silently truncated,
        /// dxf borders and number formats that read back, shared formulas whose dependents are restored and the
        /// 1904 date system that is no longer converted to 1900 on the way out.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunXlsxReadFidelityTest(string outputDirectory)
        {
            RunThemePaletteAndTintTest();
            RunPaneStateTest(outputDirectory);
            RunWhitespaceTextTest(outputDirectory);
            RunMultiRangeSqrefTest();
            RunDxfBorderAndFormatTest(outputDirectory);
            RunSharedFormulaTest(outputDirectory);
            RunDate1904Test(outputDirectory);
            Console.WriteLine("[xlsx-fidelity] theme colours, pane state, whitespace text, multi-range sqref, "
                + "dxf border/format, shared formulas, date1904 ok");
        }

        /// <summary>
        /// 共享公式：主格带文本、从格只带 si，过去两种都按纯文本读取，从格一律读成"无公式"。
        /// 现在从格按与主格的偏移还原，绝对标记不跟着走，函数名、字符串与整列引用不受伤。
        /// Shared formulas: the master carries the text and the dependents only the si number, while both used to
        /// be read as plain text, so every dependent came back with no formula. A dependent is now translated by
        /// its offset to the master, absolute marks stay put, and function names, string literals and
        /// whole-column references survive.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunSharedFormulaTest(string outputDirectory)
        {
            string sheetXml = "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetData>"
                + "<row r=\"1\"><c r=\"A1\"><v>1</v></c></row>"
                + "<row r=\"2\">"
                + "<c r=\"B2\" t=\"str\"><f t=\"shared\" ref=\"B2:C3\" si=\"0\">A2+1</f><v>2</v></c>"
                + "<c r=\"C2\" t=\"str\"><f t=\"shared\" si=\"0\"/><v>3</v></c>"
                + "<c r=\"D2\" t=\"str\"><f t=\"shared\" ref=\"D2:E3\" si=\"1\">A2+$B$2+D$4</f><v>4</v></c>"
                + "<c r=\"E2\" t=\"str\"><f t=\"shared\" si=\"1\"/><v>5</v></c>"
                + "<c r=\"F2\" t=\"str\"><f t=\"shared\" ref=\"F2:G3\" si=\"2\">LOG10(F2)&amp;\"F2\"&amp;SUM(F:F)</f><v>6</v></c>"
                + "<c r=\"G2\" t=\"str\"><f t=\"shared\" si=\"2\"/><v>7</v></c>"
                + "<c r=\"H2\" t=\"str\"><f t=\"shared\" ref=\"H2:H3\" si=\"3\">B1048576</f><v>14</v></c>"
                + "</row>"
                + "<row r=\"3\">"
                + "<c r=\"B3\" t=\"str\"><f t=\"shared\" si=\"0\"/><v>8</v></c>"
                + "<c r=\"C3\" t=\"str\"><f t=\"shared\" si=\"0\"/><v>9</v></c>"
                + "<c r=\"E3\" t=\"str\"><f t=\"shared\" si=\"1\"/><v>10</v></c>"
                + "<c r=\"G3\" t=\"str\"><f t=\"shared\" si=\"2\"/><v>11</v></c>"
                + // 主格落在共享区域左上角，真实文件里偏移只会朝下朝右；这里取一行贴到网格底边，
                // 再加偏移就出界，必须落成 #REF! 而不是 B1048577。
                // The master sits at the top-left of the shared range, so a real file only shifts down and right.
                // This one starts on the last grid row, so the extra delta leaves it and must become #REF!
                // instead of B1048577.
                "<c r=\"H3\" t=\"str\"><f t=\"shared\" si=\"3\"/><v>12</v></c>"
                + "</row>"
                + "</sheetData></worksheet>";
            OfficeWorkbook workbook = XlsxReader.Load(BuildFidelityPackage(sheetXml, null, null));
            Worksheet sheet = workbook.Worksheets[0];
            Require(sheet.GetCell(2, 2).Formula == "A2+1", "the master formula changed: " + sheet.GetCell(2, 2).Formula);
            Require(sheet.GetCell(2, 3).Formula == "B2+1", "the same-row dependent did not shift: " + sheet.GetCell(2, 3).Formula);
            Require(sheet.GetCell(3, 2).Formula == "A3+1", "the same-column dependent did not shift: " + sheet.GetCell(3, 2).Formula);
            Require(sheet.GetCell(3, 3).Formula == "B3+1", "the diagonal dependent did not shift: " + sheet.GetCell(3, 3).Formula);
            Require(sheet.GetCell(2, 4).Formula == "A2+$B$2+D$4", "the mixed master changed: " + sheet.GetCell(2, 4).Formula);
            Require(sheet.GetCell(2, 5).Formula == "B2+$B$2+E$4",
                "absolute marks must not move with the dependent: " + sheet.GetCell(2, 5).Formula);
            Require(sheet.GetCell(2, 6).Formula == "LOG10(F2)&\"F2\"&SUM(F:F)",
                "the function-name master changed: " + sheet.GetCell(2, 6).Formula);
            Require(sheet.GetCell(2, 7).Formula == "LOG10(G2)&\"F2\"&SUM(F:F)",
                "a function name, a string or a whole column was rewritten: " + sheet.GetCell(2, 7).Formula);
            Require(sheet.GetCell(3, 8).Formula == "#REF!", "a reference pushed past the last grid row should be #REF!: "
                + sheet.GetCell(3, 8).Formula);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "xlsx-fidelity-shared-formulas.xlsx"));
            XlsxWriter.Write(workbook, path);
            string written = ReadPackagePartText(path, "xl/worksheets/sheet1.xml");
            Require(!written.Contains("t=\"shared\""),
                "the writer should expand every shared formula into a plain one: " + written);
            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Require(reloaded.Worksheets[0].GetCell(3, 3).Formula == "B3+1",
                "an expanded formula did not survive the round-trip: " + reloaded.Worksheets[0].GetCell(3, 3).Formula);
        }

        /// <summary>
        /// 1904 日期系统：读端认得 workbookPr/@date1904，写出端过去却从不写回这个标记，
        /// 于是 1904 工作簿读进来再写出去，序列号仍按 1900 基准生成，所有日期整体偏移四年。
        /// The 1904 date system: the reader honoured workbookPr/@date1904 but the writer never emitted it, so a
        /// 1904 workbook read and re-written serialised its dates against the 1900 epoch and shifted every one of
        /// them by four years.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunDate1904Test(string outputDirectory)
        {
            const double serial = 39000;
            string stylesXml = "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
                + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
                + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"2\">"
                + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                + "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/>"
                + "</cellXfs>";
            string sheetXml = "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetData><row r=\"1\"><c r=\"A1\" s=\"1\"><v>" + serial.ToString(CultureInfo.InvariantCulture)
                + "</v></c></row></sheetData></worksheet>";
            DateTime epoch1904 = new DateTime(1904, 1, 1);
            DateTime epoch1900 = new DateTime(1899, 12, 30);

            OfficeWorkbook book = XlsxReader.Load(BuildFidelityPackage(sheetXml, stylesXml, null, true));
            Require(book.Date1904, "workbookPr/@date1904 was not carried onto the model");
            object value = book.Worksheets[0].GetCell(1, 1).Value;
            Require(value is DateTime, "a date-styled number should read as DateTime but was " + value);
            Require((DateTime)value == epoch1904.AddDays(serial),
                "the 1904 epoch was not used when converting the serial: " + value);

            OfficeWorkbook plain = XlsxReader.Load(BuildFidelityPackage(sheetXml, stylesXml, null, false));
            Require(!plain.Date1904, "a workbook without the flag should stay on the 1900 system");
            Require((DateTime)plain.Worksheets[0].GetCell(1, 1).Value == epoch1900.AddDays(serial),
                "the 1900 epoch changed: " + plain.Worksheets[0].GetCell(1, 1).Value);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "xlsx-fidelity-date1904.xlsx"));
            XlsxWriter.Write(book, path);
            string writtenWorkbook = ReadPackagePartText(path, "xl/workbook.xml");
            Require(Regex.IsMatch(writtenWorkbook, "<workbookPr[^>]*date1904=\"1\""),
                "the written workbook lost the date1904 flag: " + writtenWorkbook);
            string writtenSheet = ReadPackagePartText(path, "xl/worksheets/sheet1.xml");
            Require(writtenSheet.Contains("<v>" + serial.ToString(CultureInfo.InvariantCulture) + "</v>"),
                "the date serial was rebased on the wrong epoch: " + writtenSheet);
            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Require(reloaded.Date1904, "the written 1904 workbook does not read back as 1904");
            Require(Equals(reloaded.Worksheets[0].GetCell(1, 1).Value, value),
                "the date moved by the round-trip: expected " + value + " but was "
                + reloaded.Worksheets[0].GetCell(1, 1).Value);
        }

        /// <summary>
        /// theme 索引与 a:clrScheme 的次序是错位的（0=lt1、1=dk1、2=lt2、3=dk2、4..9=accent1-6），
        /// 按字面次序取色会把深色 1 读成白色；tint 更要参与换算，否则强调色一律读成原色。
        /// The theme indices are swizzled against the a:clrScheme order (0=lt1, 1=dk1, 2=lt2, 3=dk2, 4..9=accent1-6),
        /// so reading them literally turns dark 1 into white; tint has to enter the conversion or every accent
        /// colour lands on its raw value.
        /// </summary>
        private static void RunThemePaletteAndTintTest()
        {
            string stylesXml = "<fonts count=\"6\">"
                + "<font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
                + "<font><color theme=\"1\"/></font>"
                + "<font><color theme=\"0\"/></font>"
                + "<font><color theme=\"3\"/></font>"
                + "<font><color theme=\"4\"/></font>"
                + "<font><color theme=\"4\" tint=\"0.5\"/></font>"
                + "</fonts>"
                + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
                + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"6\">"
                + "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
                + "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"2\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"3\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"4\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "<xf numFmtId=\"0\" fontId=\"5\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyFont=\"1\"/>"
                + "</cellXfs>";
            string sheetXml = "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetData>"
                + "<row r=\"1\"><c r=\"A1\" s=\"1\"/><c r=\"A2\" s=\"2\"/><c r=\"A3\" s=\"3\"/>"
                + "<c r=\"A4\" s=\"4\"/><c r=\"A5\" s=\"5\"/></row>"
                + "</sheetData></worksheet>";

            OfficeWorkbook themed = XlsxReader.Load(BuildFidelityPackage(sheetXml, stylesXml, FidThemeXml));
            Worksheet sheet = themed.Worksheets[0];
            Require(sheet.GetCell(1, 1).Style.FontColor == new OfficeColor(0x22, 0x22, 0x22),
                "theme 1 should resolve to dk1 222222 but was " + sheet.GetCell(1, 1).Style.FontColor.ToHex());
            Require(sheet.GetCell(2, 1).Style.FontColor == new OfficeColor(0xEE, 0xEE, 0xEE),
                "theme 0 should resolve to lt1 EEEEEE but was " + sheet.GetCell(2, 1).Style.FontColor.ToHex());
            Require(sheet.GetCell(3, 1).Style.FontColor == new OfficeColor(0x44, 0x44, 0x44),
                "theme 3 should resolve to dk2 444444 but was " + sheet.GetCell(3, 1).Style.FontColor.ToHex());
            OfficeColor accent = sheet.GetCell(4, 1).Style.FontColor;
            Require(accent == new OfficeColor(0x44, 0x72, 0xC4),
                "theme 4 should resolve to accent1 4472C4 but was " + accent.ToHex());

            OfficeColor tinted = sheet.GetCell(5, 1).Style.FontColor;
            Require(tinted != accent, "tint 0.5 on accent1 was ignored: " + tinted.ToHex());
            Require(tinted.R >= accent.R && tinted.G >= accent.G && tinted.B >= accent.B,
                "a positive tint should lighten, not darken: " + tinted.ToHex());
            Require(tinted.R > accent.R && tinted.B > accent.B,
                "the positive tint should move the channels away from accent1: " + tinted.ToHex());

            // 主题部件缺失时不能猜色：留在样式默认值上，也不该抛异常。
            // Without a theme part nothing may be guessed: the colour stays on the style default and no throw happens.
            OfficeWorkbook unthemed = XlsxReader.Load(BuildFidelityPackage(sheetXml, stylesXml, null));
            CellStyle fallbackStyle = new CellStyle();
            Require(unthemed.Worksheets[0].GetCell(1, 1).Style.FontColor == fallbackStyle.FontColor,
                "a missing theme should keep the style default instead of inventing a colour: "
                + unthemed.Worksheets[0].GetCell(1, 1).Style.FontColor.ToHex());
        }

        /// <summary>
        /// 只有显式 state 才是冻结：视图分割（View→Split）同样带 xSplit/ySplit 却没有 state，
        /// 缺省当冻结会把普通分割窗格报成冻结窗格。
        /// Only an explicit state means frozen: a View→Split pane carries xSplit/ySplit too but no state, so
        /// treating the absence as frozen reports an ordinary split as a frozen pane.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunPaneStateTest(string outputDirectory)
        {
            OfficeWorkbook split = XlsxReader.Load(BuildFidelityPackage(
                "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetViews><sheetView workbookViewId=\"0\">"
                + "<pane xSplit=\"2\" ySplit=\"1\" topLeftCell=\"C2\" activePane=\"bottomRight\"/>"
                + "</sheetView></sheetViews>"
                + "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>",
                null, null));
            Require(split.Worksheets[0].FrozenRows == 0 && split.Worksheets[0].FrozenColumns == 0,
                "a pane without a state is a split, not a freeze: rows=" + split.Worksheets[0].FrozenRows
                + " columns=" + split.Worksheets[0].FrozenColumns);

            OfficeWorkbook frozen = XlsxReader.Load(BuildFidelityPackage(
                "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetViews><sheetView workbookViewId=\"0\">"
                + "<pane xSplit=\"2\" ySplit=\"1\" topLeftCell=\"C2\" activePane=\"bottomRight\" state=\"frozen\"/>"
                + "</sheetView></sheetViews>"
                + "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>",
                null, null));
            Require(frozen.Worksheets[0].FrozenRows == 1 && frozen.Worksheets[0].FrozenColumns == 2,
                "state=\"frozen\" should still report 1 frozen row and 2 frozen columns: rows="
                + frozen.Worksheets[0].FrozenRows + " columns=" + frozen.Worksheets[0].FrozenColumns);

            // 写出端一贯带 state="frozen"，收紧读端不能让自家往返丢掉冻结。
            // The writer always emits state="frozen", so tightening the read must not lose our own round-trip.
            OfficeWorkbook book = new OfficeWorkbook();
            Worksheet sheet = book.AddWorksheet("Frozen");
            sheet.GetCell(1, 1).Value = "x";
            sheet.FrozenRows = 1;
            sheet.FrozenColumns = 2;
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "xlsx-fidelity-pane.xlsx"));
            XlsxWriter.Write(book, path);
            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Require(reloaded.Worksheets[0].FrozenRows == 1 && reloaded.Worksheets[0].FrozenColumns == 2,
                "the frozen pane written by this library no longer reads back frozen");
        }

        /// <summary>
        /// XmlReader 把只含空白的元素文本报成 Whitespace 而不是 Text，纯空格的缓存值因此整格消失；
        /// 写出端还要给 t="str" 的 &lt;v&gt; 带上 xml:space="preserve"，否则下一轮解析同样会丢。
        /// XmlReader reports whitespace-only element text as Whitespace, not Text, so a space-only cached value
        /// vanished cell-wide; the writer must also carry xml:space="preserve" on the t="str" &lt;v&gt; or the next
        /// parse drops it the same way.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWhitespaceTextTest(string outputDirectory)
        {
            OfficeWorkbook whitespace = XlsxReader.Load(BuildFidelityPackage(
                "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetData><row r=\"1\">"
                + "<c r=\"A1\" t=\"str\"><v xml:space=\"preserve\">   </v></c>"
                + "<c r=\"B1\" t=\"str\"><v>   </v></c>"
                + "<c r=\"C1\" t=\"str\"><v xml:space=\"preserve\"> x </v></c>"
                + "<c r=\"D1\" t=\"inlineStr\"><is><t>   </t></is></c>"
                + "</row></sheetData></worksheet>",
                null, null));
            Worksheet sheet = whitespace.Worksheets[0];
            Require(Equals(sheet.GetCell(1, 1).Value, "   "), "a preserved whitespace cache value was dropped: '"
                + sheet.GetCell(1, 1).Value + "'");
            Require(Equals(sheet.GetCell(1, 2).Value, "   "), "a whitespace-only cache value without the preserve "
                + "hint was dropped: '" + sheet.GetCell(1, 2).Value + "'");
            Require(Equals(sheet.GetCell(1, 3).Value, " x "), "a padded text changed: '" + sheet.GetCell(1, 3).Value + "'");
            Require(Equals(sheet.GetCell(1, 4).Value, "   "), "a whitespace-only inline string was dropped: '"
                + sheet.GetCell(1, 4).Value + "'");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "xlsx-fidelity-whitespace.xlsx"));
            OfficeWorkbook book = new OfficeWorkbook();
            Worksheet written = book.AddWorksheet("Space");
            written.GetCell(1, 1).Formula = "\"   \"";
            written.GetCell(1, 1).Value = "   ";
            written.GetCell(2, 1).Value = "  plain  ";
            XlsxWriter.Write(book, path);
            string sheetPart = ReadPackagePartText(path, "xl/worksheets/sheet1.xml");
            Require(Regex.IsMatch(sheetPart, "<v xml:space=\"preserve\">   </v>"),
                "the str cache value was written without the preserve hint: " + sheetPart);
            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Require(Equals(reloaded.Worksheets[0].GetCell(1, 1).Value, "   "),
                "a space-only formula result did not survive the round-trip: '"
                + reloaded.Worksheets[0].GetCell(1, 1).Value + "'");
            Require(Equals(reloaded.Worksheets[0].GetCell(2, 1).Value, "  plain  "),
                "a padded shared string did not survive the round-trip: '"
                + reloaded.Worksheets[0].GetCell(2, 1).Value + "'");
        }

        /// <summary>
        /// sqref 是按空格分隔的多区域列表，只取第一段会把 Q5:R5 的条件格式静默丢掉。
        /// sqref is a space-separated list of ranges; keeping only the first fragment silently drops the
        /// conditional formatting over Q5:R5.
        /// </summary>
        private static void RunMultiRangeSqrefTest()
        {
            OfficeWorkbook workbook = XlsxReader.Load(BuildFidelityPackage(
                "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetData><row r=\"5\"><c r=\"D5\"><v>1</v></c><c r=\"Q5\"><v>2</v></c></row></sheetData>"
                + "<conditionalFormatting sqref=\"D5:P5 Q5:R5\">"
                + "<cfRule type=\"dataBar\" priority=\"1\"><dataBar><cfvo type=\"min\"/><cfvo type=\"max\"/>"
                + "<color rgb=\"FF63C384\"/></dataBar></cfRule></conditionalFormatting>"
                + "</worksheet>",
                null, null));
            Worksheet sheet = workbook.Worksheets[0];
            Require(sheet.ConditionalFormats.Count == 1,
                "one conditionalFormatting element should stay one rule set: " + sheet.ConditionalFormats.Count);
            string range = sheet.ConditionalFormats[0].Range;
            Require(range.Contains("D5:P5") && range.Contains("Q5:R5"),
                "the second sqref fragment was dropped: " + range);
            Require(sheet.ConditionalFormats[0].Rules.Count == 1
                && sheet.ConditionalFormats[0].Rules[0].Kind == ExcelConditionalKind.DataBar,
                "the data-bar rule did not survive the multi-range read");
        }

        /// <summary>
        /// dxf 允许 numFmt 与 border 子元素，读端过去只认字体和填充；写出端的去重键同样不含边框，
        /// 于是两条只差边框的规则共用一个 dxfId，后写的边框把前一条覆盖掉。
        /// A dxf may carry numFmt and border children although the read side only knew font and fill; the write
        /// side de-duplicated without the borders too, so two rules differing only in their borders shared one
        /// dxfId and the later border overwrote the earlier one.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunDxfBorderAndFormatTest(string outputDirectory)
        {
            string stylesXml = "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
                + "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
                + "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
                + "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
                + "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>"
                + "<dxfs count=\"1\"><dxf>"
                + "<numFmt numFmtId=\"170\" formatCode=\"0.00%\"/>"
                + "<font><b/><u/><color rgb=\"FFFF0000\"/></font>"
                + "<border><bottom style=\"thin\"><color rgb=\"FFFF0000\"/></bottom></border>"
                + "<fill><patternFill><bgColor rgb=\"FFFFFFCC\"/></patternFill></fill>"
                + "</dxf></dxfs>";
            OfficeWorkbook workbook = XlsxReader.Load(BuildFidelityPackage(
                "<worksheet xmlns=\"" + FidMainNamespace + "\">"
                + "<sheetData><row r=\"1\"><c r=\"A1\"><v>9</v></c></row></sheetData>"
                + "<conditionalFormatting sqref=\"A1:A3\">"
                + "<cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"greaterThan\"><formula>5</formula></cfRule>"
                + "</conditionalFormatting></worksheet>",
                stylesXml, null));
            Worksheet sheet = workbook.Worksheets[0];
            Require(sheet.ConditionalFormats.Count == 1 && sheet.ConditionalFormats[0].Rules.Count == 1,
                "the cellIs rule was not read: " + sheet.ConditionalFormats.Count);
            CellStyle format = sheet.ConditionalFormats[0].Rules[0].Format;
            Require(format != null, "the dxf format resolved to null");
            Require(format.NumberFormat == "0.00%", "the dxf number format was lost: " + format.NumberFormat);
            Require(format.Bold && format.Underline, "the dxf font flags were lost");
            Require(format.FontColor == OfficeColor.Red, "the dxf font colour was lost: " + format.FontColor.ToHex());
            Require(format.FillColor.HasValue && format.FillColor.Value == new OfficeColor(0xFF, 0xFF, 0xCC),
                "the dxf fill was lost: " + format.FillColor);
            Require(format.BottomBorder != null && format.BottomBorder.Color == OfficeColor.Red,
                "the dxf border was lost: " + (format.BottomBorder == null ? "null" : format.BottomBorder.Color.ToHex()));

            // 写出端：两条只差下划边框的规则必须落在不同 dxfId 上，读回来边框各自保留。
            // On the write side: two rules differing only in their bottom border must land on different dxfIds and
            // read back with their own border.
            OfficeWorkbook book = new OfficeWorkbook();
            Worksheet rules = book.AddWorksheet("Rules");
            rules.GetCell(1, 1).Value = 9d;
            ExcelConditionalFormat set = new ExcelConditionalFormat("A1:A5");
            rules.ConditionalFormats.Add(set);
            CellStyle plain = new CellStyle();
            plain.FontColor = OfficeColor.Red;
            CellStyle bordered = new CellStyle();
            bordered.FontColor = OfficeColor.Red;
            bordered.BottomBorder = new BorderLine(BorderStyle.Single, 1.0, OfficeColor.Blue);
            set.AddCellIs("greaterThan", "5", plain);
            set.AddCellIs("lessThan", "1", bordered);
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "xlsx-fidelity-dxf.xlsx"));
            XlsxWriter.Write(book, path);
            string writtenStyles = ReadPackagePartText(path, "xl/styles.xml");
            Require(Regex.Matches(writtenStyles, "<dxf>").Count == 2,
                "two rules that differ only in a border shared one dxf: " + writtenStyles);
            OfficeWorkbook reloaded = XlsxReader.Load(path);
            List<ExcelConditionalRule> back = reloaded.Worksheets[0].ConditionalFormats[0].Rules;
            Require(back.Count == 2, "both cellIs rules should read back: " + back.Count);
            Require(back[0].Format != null && back[0].Format.BottomBorder == null,
                "the border-free rule picked up a border");
            Require(back[1].Format != null && back[1].Format.BottomBorder != null
                && back[1].Format.BottomBorder.Color == OfficeColor.Blue,
                "the bordered rule lost its border");
        }

        /// <summary>主命名空间文本，供手工拼装的工作表 XML 使用。Main namespace text for the hand-built worksheet XML.</summary>
        private const string FidMainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        /// <summary>一份标准主题：lt1/dk1 走 sysClr，其余走 srgbClr，用来验证 theme 索引错位。A standard theme: lt1/dk1 as sysClr and the rest as srgbClr, used to prove the swizzled theme indices.</summary>
        private const string FidThemeXml =
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Fidelity\">"
            + "<a:themeElements><a:clrScheme name=\"Colours\">"
            + "<a:lt1><a:sysClr val=\"window\" lastClr=\"EEEEEE\"/></a:lt1>"
            + "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"222222\"/></a:dk1>"
            + "<a:lt2><a:srgbClr val=\"CCCCCC\"/></a:lt2>"
            + "<a:dk2><a:srgbClr val=\"444444\"/></a:dk2>"
            + "<a:accent1><a:srgbClr val=\"4472C4\"/></a:accent1>"
            + "<a:accent2><a:srgbClr val=\"ED7D31\"/></a:accent2>"
            + "<a:accent3><a:srgbClr val=\"A9D18E\"/></a:accent3>"
            + "<a:accent4><a:srgbClr val=\"FFC000\"/></a:accent4>"
            + "<a:accent5><a:srgbClr val=\"5B9BD5\"/></a:accent5>"
            + "<a:accent6><a:srgbClr val=\"70AD47\"/></a:accent6>"
            + "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>"
            + "<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>"
            + "</a:clrScheme></a:themeElements></a:theme>";

        /// <summary>
        /// 组装一份可控的 xlsx 包：工作表、样式表与主题部件按需给出，用于喂给读端真实 Excel 会写出的形状。
        /// Assembles a controllable xlsx package: the worksheet, style sheet and theme parts are supplied on
        /// demand, feeding the reader the shapes real Excel writes.
        /// </summary>
        /// <param name="sheetXml">工作表部件 XML。The worksheet part XML.</param>
        /// <param name="stylesXml">styles.xml 根元素内容；为 null 时不带样式部件。The styles.xml root content, omitted when null.</param>
        /// <param name="themeXml">theme1.xml 内容；为 null 时不带主题部件。The theme1.xml content, omitted when null.</param>
        /// <param name="date1904">是否在 workbookPr 声明 1904 日期系统。Whether workbookPr declares the 1904 date system.</param>
        /// <returns>包字节。The package bytes.</returns>
        private static byte[] BuildFidelityPackage(string sheetXml, string stylesXml, string themeXml, bool date1904 = false)
        {
            UTF8Encoding utf8 = new UTF8Encoding(false);
            Dictionary<string, byte[]> entries = new Dictionary<string, byte[]>();
            StringBuilder types = new StringBuilder();
            types.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            types.Append("<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            if (stylesXml != null)
            {
                types.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            }
            if (themeXml != null)
            {
                types.Append("<Override PartName=\"/xl/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>");
            }
            entries["[Content_Types].xml"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + types + "</Types>");
            entries["_rels/.rels"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                + "</Relationships>");

            StringBuilder relationships = new StringBuilder();
            relationships.Append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>");
            if (stylesXml != null)
            {
                relationships.Append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            }
            if (themeXml != null)
            {
                relationships.Append("<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>");
            }
            entries["xl/_rels/workbook.xml.rels"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + relationships + "</Relationships>");
            entries["xl/workbook.xml"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                + "<workbook xmlns=\"" + FidMainNamespace + "\" "
                + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + (date1904 ? "<workbookPr date1904=\"1\"/>" : string.Empty)
                + "<sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
            entries["xl/worksheets/sheet1.xml"] = utf8.GetBytes(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + sheetXml);
            if (stylesXml != null)
            {
                entries["xl/styles.xml"] = utf8.GetBytes(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<styleSheet xmlns=\"" + FidMainNamespace + "\">" + stylesXml + "</styleSheet>");
            }
            if (themeXml != null)
            {
                entries["xl/theme/theme1.xml"] = utf8.GetBytes(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + themeXml);
            }
            return CreateZipBytes(entries);
        }
    }
}
