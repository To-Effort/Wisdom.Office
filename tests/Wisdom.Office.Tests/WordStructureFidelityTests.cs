using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Wisdom.Office.Core;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>执行 Word OOXML 结构保真测试：相邻表格、分节符归属、运行级字符属性与表格对齐。
        /// Runs the Word OOXML structure fidelity tests: adjacent tables, section-break ownership,
        /// run-level character properties and table alignment.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWordStructureFidelityTest(string outputDirectory)
        {
            RunAdjacentTablesStaySeparate(outputDirectory);
            RunEmptySectionKeepsOwnSectPr(outputDirectory);
            RunRunPropertiesDoNotDefeatParagraphStyles(outputDirectory);
            RunTableAlignmentUsesStandardJc(outputDirectory);
            RunDocumentDefaultColorLandsInDocDefaults(outputDirectory);
            RunListMarkersBecomeNumbering(outputDirectory);
            RunUnclosedListItemsStaySeparate();
            Console.WriteLine("[word-structure] adjacent tables, sectPr placement, run properties, table jc, docDefaults color, list numbering, unclosed li ok");
        }

        /// <summary>两张紧邻的表格之间必须插入分隔段落，否则 Word 会把它们显示成一张表。
        /// A separator paragraph must be injected between two adjacent tables, otherwise Word renders them as one table.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunAdjacentTablesStaySeparate(string outputDirectory)
        {
            var source = new OfficeDocument();
            Section section = source.AddSection();
            var first = new Table();
            first.AddRow("a", "b");
            var second = new Table();
            second.AddRow("c", "d", "e");
            section.Blocks.Add(first);
            section.Blocks.Add(second);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-structure-adjacent-tables.docx"));
            WordFormatConverter.Save(source, path, WordDocumentFormat.Docx);
            string xml = ReadPackagePartText(path, "word/document.xml");

            Require(!Regex.IsMatch(xml, "</w:tbl>\\s*<w:tbl>"), "adjacent tables were written with no separator");
            Require(Regex.Matches(xml, "<w:tbl>").Count == 2, "expected two w:tbl elements");

            OfficeDocument loaded = WordFormatConverter.Load(path);
            List<Table> tables = loaded.Sections.SelectMany(item => item.Blocks).OfType<Table>().ToList();
            Require(tables.Count == 2, "adjacent tables read back as " + tables.Count);
            Require(tables[0].Rows[0].Cells.Count == 2, "first table column count, got " + tables[0].Rows[0].Cells.Count);
            Require(tables[1].Rows[0].Cells.Count == 3, "second table column count, got " + tables[1].Rows[0].Cells.Count);
        }

        /// <summary>空节必须拥有自己的分节符：每个 pPr 至多一个 sectPr，sectPr 数量与节数一致。
        /// An empty section must carry its own section break: at most one sectPr per pPr and one sectPr per section.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunEmptySectionKeepsOwnSectPr(string outputDirectory)
        {
            var source = new OfficeDocument();
            source.AddSection().AddParagraph("第一节");
            source.AddSection();
            source.AddSection().AddParagraph("第三节");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-structure-empty-section.docx"));
            WordFormatConverter.Save(source, path, WordDocumentFormat.Docx);
            string xml = ReadPackagePartText(path, "word/document.xml");

            XDocument document = XDocument.Parse(xml);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            int total = document.Descendants(w + "sectPr").Count();
            Require(total == 3, "sectPr count, got " + total);
            Require(document.Descendants(w + "pPr").Max(item => item.Elements(w + "sectPr").Count()) == 1,
                "a single pPr carried more than one sectPr");
            Require(document.Descendants(w + "p")
                    .Count(item => item.Element(w + "pPr") != null && item.Element(w + "pPr").Element(w + "sectPr") != null) == 2,
                "only the two non-final sections may carry a sectPr inside pPr");

            OfficeDocument loaded = WordFormatConverter.Load(path);
            Require(loaded.Sections.Count == 3, "section count, got " + loaded.Sections.Count);
        }

        /// <summary>标题段落不得写出运行级字号/颜色/字体，否则样式表里的标题格式被压过；
        /// 显式设置的字符属性仍然照常写出。
        /// A heading paragraph must not emit run-level size/color/font, which would defeat the style sheet; explicit
        /// character properties must still be written.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunRunPropertiesDoNotDefeatParagraphStyles(string outputDirectory)
        {
            var source = new OfficeDocument();
            Section section = source.AddSection();
            Paragraph heading = section.AddParagraph("标题文字");
            heading.HeadingLevel = 1;
            Paragraph body = section.AddParagraph("正文");
            var highlight = new TextFormat
            {
                FontFamily = "宋体",
                FontSize = 14,
                Color = OfficeColor.Red,
                Bold = true
            };
            body.AddText("红色强调", highlight);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-structure-run-properties.docx"));
            WordFormatConverter.Save(source, path, WordDocumentFormat.Docx);
            string xml = ReadPackagePartText(path, "word/document.xml");

            string headingXml = ParagraphOf(xml, "标题文字");
            Require(headingXml.Contains("w:val=\"Heading1\""), "heading paragraph lost pStyle=Heading1");
            Require(!Regex.IsMatch(headingXml, "<w:sz "), "heading paragraph forced a run-level size");
            Require(!Regex.IsMatch(headingXml, "<w:color "), "heading paragraph forced a run-level color");
            Require(!headingXml.Contains("<w:rFonts"), "heading paragraph forced a run-level font");

            string highlightXml = ParagraphOf(xml, "红色强调");
            Require(Regex.IsMatch(highlightXml, "<w:color w:val=\"FF0000\""), "explicit red color was dropped");
            Require(Regex.IsMatch(highlightXml, "<w:sz w:val=\"28\""), "explicit 14 pt size was dropped");
            Require(highlightXml.Contains("宋体"), "explicit font family was dropped");
            Require(Regex.IsMatch(highlightXml, "<w:b[ />]"), "explicit bold was dropped");

            OfficeDocument loaded = WordFormatConverter.Load(path);
            List<Paragraph> paragraphs = loaded.Sections.SelectMany(item => item.Blocks).OfType<Paragraph>().ToList();
            Require(paragraphs.Count >= 2, "paragraph count, got " + paragraphs.Count);
            Require(paragraphs[0].HeadingLevel == 1, "heading level read back as " + paragraphs[0].HeadingLevel);
        }

        /// <summary>表格对齐使用 ECMA-376 定义的 w:jc，而不是自造的 w:tblJc。
        /// Table alignment must use the schema-defined w:jc instead of an invented w:tblJc.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunTableAlignmentUsesStandardJc(string outputDirectory)
        {
            var source = new OfficeDocument();
            Section section = source.AddSection();
            var table = new Table();
            table.AddRow("x", "y");
            table.Alignment = TableAlignment.Center;
            section.Blocks.Add(table);
            section.AddParagraph("表后");

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-structure-table-alignment.docx"));
            WordFormatConverter.Save(source, path, WordDocumentFormat.Docx);
            string xml = ReadPackagePartText(path, "word/document.xml");

            Require(!xml.Contains("tblJc"), "writer emitted the non-standard w:tblJc");
            Require(Regex.IsMatch(xml, "<w:tblPr>.*?<w:jc w:val=\"center\"", RegexOptions.Singleline),
                "centered table did not write w:jc inside w:tblPr");

            OfficeDocument loaded = WordFormatConverter.Load(path);
            Table readBack = loaded.Sections.SelectMany(item => item.Blocks).OfType<Table>().FirstOrDefault();
            Require(readBack != null, "table missing after reload");
            Require(readBack.Alignment == TableAlignment.Center, "alignment read back as " + readBack.Alignment);
        }

        /// <summary>非黑色的文档默认字色落在 docDefaults 里；与默认色一致的运行可省略颜色，
        /// 与默认色不同的运行仍显式写色。
        /// A non-black document default color belongs in docDefaults; runs matching it may omit w:color while runs
        /// differing from it still write the color explicitly.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunDocumentDefaultColorLandsInDocDefaults(string outputDirectory)
        {
            var source = new OfficeDocument();
            source.DefaultTextFormat.Color = OfficeColor.Red;
            Section section = source.AddSection();
            section.AddParagraph("默认色正文");
            var matching = new Paragraph();
            matching.AddText("同为红色", new TextFormat { Color = OfficeColor.Red });
            section.Blocks.Add(matching);

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-structure-default-color.docx"));
            WordFormatConverter.Save(source, path, WordDocumentFormat.Docx);

            string styles = ReadPackagePartText(path, "word/styles.xml");
            Require(Regex.IsMatch(styles, "rPrDefault.*?<w:color w:val=\"FF0000\".*?</w:rPrDefault>", RegexOptions.Singleline),
                "docDefaults did not carry the document default color");

            string xml = ReadPackagePartText(path, "word/document.xml");
            Require(!Regex.IsMatch(ParagraphOf(xml, "同为红色"), "<w:color"),
                "a run matching the default color still forced w:color");
            Require(Regex.IsMatch(ParagraphOf(xml, "默认色正文"), "<w:color w:val=\"000000\""),
                "a run with no explicit color did not fall back to black");
        }

        /// <summary>列表标记要写成真正的编号定义（numbering.xml + w:numPr）：有序列表逐项递增、项目符号保留符号，
        /// 被普通段落打断后重新起号，读回时标记文本与写出前一致。
        /// List markers must become real numbering (numbering.xml plus w:numPr): ordered items increment, bullets
        /// keep their symbol, an interrupted list restarts its count, and the marker text survives the round trip.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunListMarkersBecomeNumbering(string outputDirectory)
        {
            var source = new OfficeDocument();
            Section section = source.AddSection();
            for (int i = 1; i <= 3; i++)
            {
                Paragraph item = section.AddParagraph("有序第" + i + "项");
                item.ListMarker = "1.";
            }
            section.AddParagraph("中断段落");
            for (int i = 1; i <= 2; i++)
            {
                Paragraph item = section.AddParagraph("无序第" + i + "项");
                item.ListMarker = "•";
            }
            section.AddParagraph("续排第1项").ListMarker = "1.";

            string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-structure-list-numbering.docx"));
            WordFormatConverter.Save(source, path, WordDocumentFormat.Docx);

            string numbering = ReadPackagePartText(path, "word/numbering.xml");
            Require(Regex.Matches(numbering, "<w:abstractNum ").Count >= 2,
                "expected an ordered and a bullet abstractNum, got " + Regex.Matches(numbering, "<w:abstractNum ").Count);
            Require(Regex.Matches(numbering, "<w:num ").Count == 3,
                "each contiguous run (ordered, bullets, resumed ordered) must get its own num, got " + Regex.Matches(numbering, "<w:num ").Count);
            Require(numbering.Contains("w:numFmt w:val=\"bullet\""), "bullet list lost its bullet numbering format");

            string xml = ReadPackagePartText(path, "word/document.xml");
            Require(Regex.Matches(xml, "<w:numId w:val=\"1\"\\s*/>").Count == 3,
                "the three ordered items did not share one numbering sequence");
            Require(Regex.Matches(xml, "<w:numId w:val=\"2\"\\s*/>").Count == 2,
                "the bullet list did not get its own numbering sequence");
            Require(Regex.Matches(xml, "<w:numId w:val=\"3\"\\s*/>").Count == 1,
                "a list resumed after an interruption must get a fresh numbering sequence");
            Require(!Regex.IsMatch(ParagraphOf(xml, "有序第1项"), "<w:t[^>]*>1\\."),
                "the ordered marker is still written as literal text instead of numbering");

            string contentTypes = ReadPackagePartText(path, "[Content_Types].xml");
            Require(contentTypes.Contains("numbering+xml"), "numbering.xml has no content-type override");

            OfficeDocument loaded = WordFormatConverter.Load(path);
            List<Paragraph> paragraphs = loaded.Sections.SelectMany(item => item.Blocks).OfType<Paragraph>().ToList();
            Require(paragraphs.Count == 7, "paragraph count after reload, got " + paragraphs.Count);
            Require(paragraphs[0].ListMarker == "1." && paragraphs[1].ListMarker == "2." && paragraphs[2].ListMarker == "3.",
                "ordered markers read back as " + paragraphs[0].ListMarker + ", " + paragraphs[1].ListMarker + ", " + paragraphs[2].ListMarker);
            Require(paragraphs[3].ListMarker == null, "the interrupting paragraph was read as a list item");
            Require(paragraphs[4].ListMarker == "•" && paragraphs[5].ListMarker == "•",
                "bullet markers read back as " + paragraphs[4].ListMarker + " and " + paragraphs[5].ListMarker);
            Require(paragraphs[6].ListMarker == "1.",
                "a list resumed after an interruption must restart at 1., got " + paragraphs[6].ListMarker);
        }

        /// <summary>HTML 允许省略 li 的结束标签：省略时每个条目仍要单独成段，各带自己的编号/符号。
        /// HTML allows the li end tag to be omitted; even then every item must stay its own paragraph with its own
        /// marker.
        /// </summary>
        private static void RunUnclosedListItemsStaySeparate()
        {
            string html = "<html><body><ol><li>第一<li>第二<li>第三</ol><ul><li>甲<li>乙</ul></body></html>";
            OfficeDocument document = HtmlReader.Parse(html);
            List<Paragraph> paragraphs = document.Sections.SelectMany(item => item.Blocks).OfType<Paragraph>().ToList();
            Require(paragraphs.Count == 5, "unclosed li items produced " + paragraphs.Count + " paragraphs");
            Require(paragraphs[0].GetText() == "第一" && paragraphs[1].GetText() == "第二" && paragraphs[2].GetText() == "第三",
                "ordered items merged: " + paragraphs[0].GetText() + " | " + paragraphs[1].GetText() + " | " + paragraphs[2].GetText());
            Require(paragraphs[0].ListMarker == "1." && paragraphs[1].ListMarker == "2." && paragraphs[2].ListMarker == "3.",
                "ordered markers read as " + paragraphs[0].ListMarker + ", " + paragraphs[1].ListMarker + ", " + paragraphs[2].ListMarker);
            Require(paragraphs[3].GetText() == "甲" && paragraphs[4].GetText() == "乙" && paragraphs[4].ListMarker == "•",
                "bullet items merged or lost their marker");
        }

        /// <summary>截取包含指定文本的第一个段落的 XML。Returns the XML of the first paragraph containing the text.</summary>
        /// <param name="documentXml">正文部件 XML。The document part XML.</param>
        /// <param name="needle">段落内的文本。Text inside the paragraph.</param>
        /// <returns>段落 XML。The paragraph XML.</returns>
        private static string ParagraphOf(string documentXml, string needle)
        {
            int at = documentXml.IndexOf(needle, StringComparison.Ordinal);
            Require(at >= 0, "text not found in document.xml: " + needle);
            int start = Math.Max(documentXml.LastIndexOf("<w:p ", at, StringComparison.Ordinal),
                documentXml.LastIndexOf("<w:p>", at, StringComparison.Ordinal));
            int end = documentXml.IndexOf("</w:p>", at, StringComparison.Ordinal);
            Require(start >= 0 && end >= 0, "paragraph bounds not found around: " + needle);
            return documentXml.Substring(start, end - start + 6);
        }
    }
}
