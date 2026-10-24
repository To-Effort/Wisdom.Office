using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using Wisdom.Excel;
using Wisdom.Office;
using Wisdom.Office.Core;
using Wisdom.Ppt;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 写出面缺陷回归：无样式的日期必须带日期格式、NaN 与无穷不能写成非法数值、
    /// XML 1.0 不接受的控制字符不能把整份文档的保存动作炸掉。
    /// Write-side regression tests: an unstyled date has to carry a date format, NaN and infinity must not be
    /// written as illegal numbers, and the control characters XML 1.0 forbids must not blow up the save of a
    /// whole document.
    /// </summary>
    internal static partial class Program
    {
        /// <summary>执行写出面回归测试。Runs the write-side regression tests.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWriteOutputTest(string outputDirectory)
        {
            RunUnstyledDateCellTest(outputDirectory);
            RunNonFiniteCellTest(outputDirectory);
            RunIllegalXmlCharacterTest(outputDirectory);
            Console.WriteLine("[write] date style, non-finite and illegal characters ok");
        }

        /// <summary>
        /// 单元格没有样式时，日期只写序列号会显示成 45064 这样的数字，必须补上日期格式；
        /// 但同一个包里的普通数字仍然应当保持无样式，不能顺手都挂上格式。
        /// A date in an unstyled cell goes out as a bare serial and shows up as digits such as 45064, so a date
        /// format has to be added — while a plain number in the same package stays unstyled instead of picking
        /// up a format by accident.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunUnstyledDateCellTest(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "write-unstyled-date.xlsx"));
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("S");
            Cell date = sheet.GetCell(1, 1);
            date.Value = new DateTime(2023, 5, 18);
            Require(date.Style == null, "probe cell should start without a style");
            sheet.GetCell(2, 1).Value = 1234d;
            XlsxWriter.Write(workbook, path);

            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Worksheet remote = reloaded.Worksheets[0];
            object back = remote.GetCell(1, 1).Value;
            Require(back is DateTime, "unstyled date read back as " + (back == null ? "null" : back.GetType().Name)
                + " value=" + back);
            Require((DateTime)back == new DateTime(2023, 5, 18), "unstyled date read back as " + back);
            Require(remote.GetCell(2, 1).Value is double, "plain number should stay a number but was "
                + DescribeValue(remote.GetCell(2, 1).Value));

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                string body = ReadZipEntryText(archive, "xl/worksheets/sheet1.xml");
                Require(body.Contains("<c r=\"A1\" s=\""), "date cell carries no style index: " + FirstFragment(body, "<c r=\"A1\""));
                Require(body.Contains("<c r=\"A2\">"), "plain number cell should stay unstyled: " + FirstFragment(body, "<c r=\"A2\""));
                string styles = ReadZipEntryText(archive, "xl/styles.xml");
                Require(styles.Contains("yyyy-mm-dd"), "no date number format was registered: " + FirstFragment(styles, "<numFmts"));
                int font = styles.IndexOf("<name val=", StringComparison.Ordinal);
                Require(font < 0 || styles.Substring(font, Math.Min(40, styles.Length - font)).Contains("Calibri"),
                    "the added date style must inherit the workbook default font, not a fresh one");
            }
        }

        /// <summary>
        /// NaN 与无穷写进 v 元素是非法数值，Excel 会把整包判为需要修复；只能落成 #NUM! 错误单元。
        /// NaN and infinity inside a v element are illegal numbers and Excel declares the whole package in need of
        /// repair, so they can only land on a #NUM! error cell.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunNonFiniteCellTest(string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "write-non-finite.xlsx"));
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("S");
            sheet.GetCell(1, 1).Value = double.NaN;
            sheet.GetCell(2, 1).Value = double.PositiveInfinity;
            sheet.GetCell(3, 1).Value = double.NegativeInfinity;
            Cell formulaCell = sheet.GetCell(4, 1);
            formulaCell.Formula = "1/0";
            formulaCell.Value = double.NaN;
            XlsxWriter.Write(workbook, path);

            using (ZipArchive archive = ZipFile.OpenRead(path))
            {
                string body = ReadZipEntryText(archive, "xl/worksheets/sheet1.xml");
                Require(!body.Contains("NaN") && !body.Contains("Infinity"),
                    "a non-finite literal reached the sheet XML: " + FirstFragment(body, "<sheetData"));
                Require(body.Contains("<c r=\"A1\" t=\"e\"><v>#NUM!</v>"), "NaN should become a #NUM! cell: "
                    + FirstFragment(body, "<c r=\"A1\""));
                Require(body.Contains("<c r=\"A3\" t=\"e\"><v>#NUM!</v>"), "negative infinity should become a #NUM! cell: "
                    + FirstFragment(body, "<c r=\"A3\""));
                Require(body.Contains("<f>1/0</f><v>#NUM!</v>"), "a formula with a non-finite cached value must keep its f and report #NUM!: "
                    + FirstFragment(body, "<c r=\"A4\""));
            }
            OfficeWorkbook reloaded = XlsxReader.Load(path);
            Worksheet remote = reloaded.Worksheets[0];
            Require(Equals(remote.GetCell(1, 1).Value, "#NUM!"), "non-finite read back as " + DescribeValue(remote.GetCell(1, 1).Value));
            Require(Equals(remote.GetCell(3, 1).Value, "#NUM!"), "negative infinity read back as " + DescribeValue(remote.GetCell(3, 1).Value));
            ValidateOoxmlPackage(path, "xl/workbook.xml");
        }

        /// <summary>
        /// 外部数据里的 0x01、0x1F 与落单的代理项是 XML 1.0 不接受的字符，System.Xml 写出时直接抛异常，
        /// 一格脏文本就能让整份工作簿、整份文档或整份演示写不出来；必须只丢字符、不丢文档。
        /// Characters such as 0x01, 0x1F and a lone surrogate coming from foreign data are not legal XML 1.0, and
        /// System.Xml throws while writing them, so one dirty cell can keep a whole workbook, document or
        /// presentation from being written; the writer has to lose the character instead of the document.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunIllegalXmlCharacterTest(string outputDirectory)
        {
            string dirty = "脏" + (char)0x01 + "数" + (char)0x1F + "据";
            // 单独的 CR 不在断言范围内：XML 的行结尾规则本来就把 #xD 归一成 #xA，与本次清洗无关。
            // A lone CR is out of scope: the XML end-of-line rule normalizes #xD into #xA regardless of this
            // clean-up.
            string withLineBreaks = "保留" + (char)9 + (char)10 + "换行";
            string emoji = "配对\U0001F600成功";
            string loneSurrogate = "落单" + (char)0xD800 + "代理项";

            string xlsxPath = PrepareOutputPath(Path.Combine(outputDirectory, "write-illegal-chars.xlsx"));
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("S" + (char)0x02);
            sheet.GetCell(1, 1).Value = dirty;
            sheet.GetCell(2, 1).Value = withLineBreaks;
            sheet.GetCell(3, 1).Value = emoji;
            sheet.GetCell(4, 1).Value = loneSurrogate;
            Cell formulaCell = sheet.GetCell(5, 1);
            formulaCell.Formula = "A1" + (char)0x03;
            formulaCell.Value = "缓" + (char)0x04 + "存";
            workbook.Metadata.Title = "标" + (char)0x05 + "题";
            XlsxWriter.Write(workbook, xlsxPath);

            OfficeWorkbook reloadedWorkbook = XlsxReader.Load(xlsxPath);
            Worksheet remote = reloadedWorkbook.Worksheets[0];
            Require(Equals(remote.GetCell(1, 1).Value, "脏数据"), "illegal characters should be dropped, got "
                + DescribeValue(remote.GetCell(1, 1).Value));
            Require(Equals(remote.GetCell(2, 1).Value, withLineBreaks),
                "tab, newline and carriage return must survive, got " + DescribeValue(remote.GetCell(2, 1).Value));
            Require(Equals(remote.GetCell(3, 1).Value, emoji), "a surrogate pair must survive");
            Require(Equals(remote.GetCell(4, 1).Value, "落单代理项"), "only the lone surrogate must be dropped, got "
                + DescribeValue(remote.GetCell(4, 1).Value));
            Require(Equals(remote.GetCell(5, 1).Value, "缓存"), "a dirty cached value must keep its clean text, got "
                + DescribeValue(remote.GetCell(5, 1).Value));
            ValidateOoxmlPackage(xlsxPath, "xl/workbook.xml");

            string docxPath = PrepareOutputPath(Path.Combine(outputDirectory, "write-illegal-chars.docx"));
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = "标" + (char)0x06 + "题";
            Section section = document.AddSection();
            section.AddParagraph("非法" + (char)0x02 + "字符");
            DocxWriter.Write(document, docxPath);
            string extracted = DocxReader.ExtractText(docxPath);
            RequireContains(extracted, "非法字符");
            Require(!ContainsControlChar(extracted), "the docx text still carries a control character: " + DescribeValue(extracted));
            ValidateOoxmlPackage(docxPath, "word/document.xml");

            string pptxPath = PrepareOutputPath(Path.Combine(outputDirectory, "write-illegal-chars.pptx"));
            OfficePresentation presentation = OfficePresentation.Widescreen();
            Slide slide = presentation.AddSlide(SlideLayoutKind.Blank);
            TextBoxShape box = new TextBoxShape(20, 20, 400, 120);
            box.AddParagraph("幻灯" + (char)0x07 + "文本");
            slide.Shapes.Add(box);
            PptxWriter.Write(presentation, pptxPath);
            OfficePresentation reloadedSlides = PptxReader.Load(pptxPath);
            string slideText = DescribeSlideText(reloadedSlides.Slides[0]);
            RequireContains(slideText, "幻灯文本");
            Require(!ContainsControlChar(slideText), "the slide text still carries a control character: " + DescribeValue(slideText));
        }

        /// <summary>取幻灯片里所有文本框的文本。Joins the text of every text box on a slide.</summary>
        /// <param name="slide">幻灯片。The slide.</param>
        /// <returns>拼接后的文本。The joined text.</returns>
        private static string DescribeSlideText(Slide slide)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                TextBoxShape box = slide.Shapes[i] as TextBoxShape;
                if (box == null) continue;
                builder.Append(box.GetText());
            }
            return builder.ToString();
        }

        /// <summary>判断文本里是否还残留 XML 不接受的字符。Reports whether the text still holds a character XML forbids.</summary>
        /// <param name="text">待检查文本。The text to check.</param>
        /// <returns>含非法字符时返回 true。True when an illegal character is present.</returns>
        private static bool ContainsControlChar(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            for (int i = 0; i < text.Length; i++)
            {
                char current = text[i];
                if (current < ' ' && current != '\t' && current != '\n' && current != '\r') return true;
            }
            return false;
        }

        /// <summary>把值描述成「类型(文本)」，文本里的控制字符写成码点。Describes a value as "type(text)" with control characters shown as code points.</summary>
        /// <param name="value">待描述的值。The value to describe.</param>
        /// <returns>描述文本。The description.</returns>
        private static string DescribeValue(object value)
        {
            if (value == null) return "null";
            string text = value as string;
            if (text == null) return value.GetType().Name + "(" + value + ")";
            StringBuilder builder = new StringBuilder();
            builder.Append(text.Length.ToString(CultureInfo.InvariantCulture)).Append('(');
            for (int i = 0; i < text.Length && i < 24; i++)
            {
                if (i > 0) builder.Append(' ');
                builder.Append(((int)text[i]).ToString(CultureInfo.InvariantCulture));
            }
            builder.Append(')');
            return "string[" + builder + "]";
        }

        /// <summary>截出包含指定标记的一段 XML，便于失败信息里看得见现场。Cuts out the XML around a marker so a failure message shows the scene.</summary>
        /// <param name="body">XML 全文。The XML text.</param>
        /// <param name="marker">标记。The marker.</param>
        /// <returns>截出的片段。The fragment.</returns>
        private static string FirstFragment(string body, string marker)
        {
            int index = body.IndexOf(marker, StringComparison.Ordinal);
            if (index < 0) return "<missing " + marker + ">";
            int end = Math.Min(body.Length, index + marker.Length + 120);
            return body.Substring(index, end - index);
        }
    }
}
