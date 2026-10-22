using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wisdom.Office.Core;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>可写的 Word 格式，用于往返与转换矩阵测试。The writable Word formats used by the round-trip and conversion-matrix tests.</summary>
        private static readonly WordDocumentFormat[] WordFormatTargets =
        {
            WordDocumentFormat.Docx,
            WordDocumentFormat.Dotx,
            WordDocumentFormat.Rtf,
            WordDocumentFormat.Html,
            WordDocumentFormat.Odt,
            WordDocumentFormat.Text
        };

        /// <summary>执行 Word 全格式测试：六种可写格式往返、六源六目标转换矩阵、旧版 DOC 转换与格式嗅探。Runs the Word format tests: round trips for six writable formats, a six-source by six-target conversion matrix, legacy DOC conversion and format sniffing.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWordFormatTest(string outputDirectory)
        {
            Dictionary<WordDocumentFormat, string> paths = new Dictionary<WordDocumentFormat, string>();
            for (int i = 0; i < WordFormatTargets.Length; i++)
            {
                WordDocumentFormat format = WordFormatTargets[i];
                string path = PrepareOutputPath(Path.Combine(outputDirectory, "word-format-roundtrip" + WordDocumentFormats.GetExtension(format)));
                OfficeDocument document = CreateWordFormatSampleDocument();
                WordFormatConverter.Save(document, path, format);
                Require(File.Exists(path) && new FileInfo(path).Length > 0, WordDocumentFormats.GetDisplayName(format) + " writes a non-empty file");
                Require(WordFormatConverter.DetectFormat(path) == format,
                    "extension detection resolves " + WordDocumentFormats.GetDisplayName(format));
                Require(WordDocumentFormats.Detect(File.ReadAllBytes(path)) == format,
                    "content sniffing resolves " + WordDocumentFormats.GetDisplayName(format));
                OfficeDocument loaded = WordFormatConverter.Load(path);
                AssertWordFormatRoundTrip(loaded, format);
                paths[format] = path;
            }

            RunWordConversionMatrixTest(paths, outputDirectory);
            RunWordFormatSniffingTest(paths);
            Console.WriteLine("Word format tests finished.");
        }

        /// <summary>构造包含标题、书签、超链接、图片、合并表格、分页、两节页眉页脚与元数据的样例文档。Builds a sample document with headings, bookmarks, hyperlinks, images, merged tables, page breaks, two-section headers/footers and metadata.</summary>
        /// <returns>新的样例文档。A new sample document.</returns>
        private static OfficeDocument CreateWordFormatSampleDocument()
        {
            OfficeDocument document = new OfficeDocument();
            document.Metadata.Title = "格式转换矩阵";
            document.Metadata.Author = "Wisdom.Office";
            document.Metadata.Subject = "DOCX/DOTX/RTF/HTML/ODT/TXT 往返";
            document.Metadata.Keywords = "word;format;conversion";

            Section first = document.AddSection();
            first.Header.AddParagraph("页眉：格式转换矩阵");
            first.Footer.AddParagraph("页脚：第一节");

            Paragraph title = new Paragraph("格式转换矩阵");
            title.HeadingLevel = 1;
            title.Alignment = ParagraphAlignment.Center;
            title.BookmarkName = "FormatTop";
            title.SpaceAfter = 12;
            first.Blocks.Add(title);

            Paragraph body = new Paragraph();
            TextFormat normal = new TextFormat("Arial", 11);
            body.AddText("正文：", normal);
            body.AddText("加粗文字", new TextFormat("Arial", 11) { Bold = true, Color = OfficeColor.Red });
            body.AddText(" 链接：", normal);
            TextRun link = new TextRun("官网", normal);
            link.Link = "https://example.com/format";
            body.AddRun(link);
            body.AddImage(new OfficeImage(CreateTestPng(48, 24)), 48, 24);
            first.Blocks.Add(body);

            Table table = new Table();
            table.Columns.Add(new TableColumn(120));
            table.Columns.Add(new TableColumn(120));
            table.Columns.Add(new TableColumn(80));
            TableRow header = table.AddRow("项目", "金额", "备注");
            header.IsHeader = true;
            TableRow mergedRow = table.AddRow();
            TableCell merged = new TableCell("合并单元格");
            merged.ColSpan = 2;
            merged.BackgroundColor = OfficeColor.LightGray;
            merged.VerticalAlignment = VerticalAlignment.Middle;
            mergedRow.Cells.Add(merged);
            mergedRow.Cells.Add(new TableCell("100"));
            table.AddRow("合计", "100", "含税");
            first.Blocks.Add(table);

            first.Blocks.Add(new Paragraph("分页前内容"));
            first.Blocks.Add(new PageBreak());

            Section second = document.AddSection();
            second.PageSetup.MarginLeft = 54;
            second.Header.AddParagraph("页眉：第二节");
            second.Footer.AddParagraph("页脚：第二节");
            second.Blocks.Add(new Paragraph("第二节：格式转换矩阵"));
            return document;
        }

        /// <summary>断言某格式读回后文本、书签、链接、图片、表格合并、节与页眉页脚保持正确；纯文本只校验文本。Asserts that text, bookmarks, links, images, table merges, sections and headers/footers survive a read-back; plain text only checks text.</summary>
        /// <param name="document">读回的文档。The document read back.</param>
        /// <param name="format">文档格式。The document format.</param>
        private static void AssertWordFormatRoundTrip(OfficeDocument document, WordDocumentFormat format)
        {
            string label = WordDocumentFormats.GetDisplayName(format);
            string text = document.GetPlainText();
            RequireContains(text, "格式转换矩阵");
            RequireContains(text, "加粗文字");
            RequireContains(text, "合并单元格");
            RequireContains(text, "合计");
            RequireContains(text, "分页前内容");
            Require(document.Sections.Count > 0, label + " keeps at least one section");

            if (format == WordDocumentFormat.Text) return;

            Require(document.Sections.Count == 2, label + " keeps two sections, got " + document.Sections.Count);
            RequireContains(document.Metadata.Title, "格式转换矩阵");
            RequireContains(document.Metadata.Author, "Wisdom.Office");
            Require(BlockTextContains(document.Sections[0].Header.Blocks, "页眉：格式转换矩阵"), label + " keeps the first header");
            Require(BlockTextContains(document.Sections[1].Footer.Blocks, "页脚：第二节"), label + " keeps the second footer");

            Paragraph bookmark = FindWordParagraph(document, "格式转换矩阵");
            Require(bookmark != null && bookmark.BookmarkName == "FormatTop", label + " keeps the bookmark");

            bool hasLink = false;
            int images = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                CollectWordFormatInlines(document.Sections[s].Blocks, ref hasLink, ref images);
            }
            Require(hasLink, label + " keeps the hyperlink");
            Require(images == 1, label + " keeps one image, got " + images);

            Table table = FindWordFormatTable(document, "合并单元格");
            Require(table != null, label + " keeps the table");
            Require(table.Rows.Count == 3, label + " keeps three table rows");
            Require(table.Rows[0].IsHeader, label + " keeps the header row flag");
            Require(table.Rows[1].Cells.Count == 2, label + " collapses the merged cell, got " + table.Rows[1].Cells.Count);
            Require(table.Rows[1].Cells[0].ColSpan == 2, label + " keeps the column span");
            RequireContains(GetWordParagraphText(table.Rows[1].Cells[1].Blocks[0]), "100");
        }

        /// <summary>收集块列表中的链接与图片数量。Collects hyperlink presence and image count from a block list.</summary>
        /// <param name="blocks">块列表。The block list.</param>
        /// <param name="hasLink">是否发现链接。Whether a link was found.</param>
        /// <param name="images">图片计数。The image count.</param>
        private static void CollectWordFormatInlines(List<Block> blocks, ref bool hasLink, ref int images)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    for (int k = 0; k < paragraph.Inlines.Count; k++)
                    {
                        ImageRun image = paragraph.Inlines[k] as ImageRun;
                        if (image != null) images++;
                        TextRun run = paragraph.Inlines[k] as TextRun;
                        if (run != null && !string.IsNullOrEmpty(run.Link)) hasLink = true;
                    }
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        CollectWordFormatInlines(table.Rows[r].Cells[c].Blocks, ref hasLink, ref images);
                    }
                }
            }
        }

        /// <summary>判断块列表（含表格单元格）中是否有包含指定文本的段落。Checks whether a block list (including table cells) contains a paragraph with the given text.</summary>
        /// <param name="blocks">块列表。The block list.</param>
        /// <param name="text">文本片段。The text fragment.</param>
        /// <returns>找到时为 true。True when found.</returns>
        private static bool BlockTextContains(List<Block> blocks, string text)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    if (paragraph.GetText().IndexOf(text, StringComparison.Ordinal) >= 0) return true;
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        if (BlockTextContains(table.Rows[r].Cells[c].Blocks, text)) return true;
                    }
                }
            }
            return false;
        }

        /// <summary>查找包含指定文本的表格。Finds the table containing the given text.</summary>
        /// <param name="document">待搜索的文档。The document to search.</param>
        /// <param name="text">单元格文本片段。The cell text fragment.</param>
        /// <returns>命中的表格，未找到时为 null。The matching table, or null when nothing matches.</returns>
        private static Table FindWordFormatTable(OfficeDocument document, string text)
        {
            for (int s = 0; s < document.Sections.Count; s++)
            {
                List<Block> blocks = document.Sections[s].Blocks;
                for (int i = 0; i < blocks.Count; i++)
                {
                    Table table = blocks[i] as Table;
                    if (table != null && BlockTextContains(new List<Block> { table }, text)) return table;
                }
            }
            return null;
        }

        /// <summary>执行六源乘六目标的转换矩阵（36 种组合），并在存在旧版 DOC 夹具时追加 DOC 转换。Runs the six-source by six-target conversion matrix (36 combinations), plus legacy DOC conversions when the fixture exists.</summary>
        /// <param name="paths">各格式的源文件路径。The source path for each format.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWordConversionMatrixTest(Dictionary<WordDocumentFormat, string> paths, string outputDirectory)
        {
            string matrixDirectory = Path.Combine(outputDirectory, "word-conversion-matrix");
            Directory.CreateDirectory(matrixDirectory);
            int combinations = 0;
            for (int s = 0; s < WordFormatTargets.Length; s++)
            {
                WordDocumentFormat source = WordFormatTargets[s];
                for (int t = 0; t < WordFormatTargets.Length; t++)
                {
                    WordDocumentFormat target = WordFormatTargets[t];
                    string targetPath = PrepareOutputPath(Path.Combine(matrixDirectory,
                        "matrix-" + source + "-to-" + target + WordDocumentFormats.GetExtension(target)));
                    WordFormatConverter.Convert(paths[source], targetPath);
                    Require(File.Exists(targetPath) && new FileInfo(targetPath).Length > 0,
                        source + " to " + target + " writes a file");
                    OfficeDocument converted = WordFormatConverter.Load(targetPath);
                    RequireContains(converted.GetPlainText(), "格式转换矩阵");
                    combinations++;
                }
            }

            string docxPath = paths[WordDocumentFormat.Docx];
            string docTarget = PrepareOutputPath(Path.Combine(matrixDirectory, "matrix-docx-to-doc.doc"));
            RequireThrows(() => WordFormatConverter.Convert(docxPath, docTarget), "legacy .doc output is rejected");

            string fixturePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "fixtures", "legacy-sample.doc"));
            if (File.Exists(fixturePath))
            {
                Require(WordFormatConverter.DetectFormat(fixturePath) == WordDocumentFormat.Doc, "legacy fixture is detected as DOC");
                OfficeDocument legacy = WordFormatConverter.Load(fixturePath);
                Require(legacy.Sections.Count > 0, "legacy DOC loads at least one section");
                for (int t = 0; t < WordFormatTargets.Length; t++)
                {
                    WordDocumentFormat target = WordFormatTargets[t];
                    string targetPath = PrepareOutputPath(Path.Combine(matrixDirectory,
                        "legacy-doc-to-" + target + WordDocumentFormats.GetExtension(target)));
                    WordFormatConverter.Convert(fixturePath, targetPath);
                    Require(File.Exists(targetPath) && new FileInfo(targetPath).Length > 0,
                        "legacy DOC to " + target + " writes a file");
                    combinations++;
                }
            }

            Require(combinations >= 36, "conversion matrix covers at least 36 combinations, got " + combinations);
            Console.WriteLine("Word conversion matrix: " + combinations + " format combinations verified.");
        }

        /// <summary>校验扩展名映射、内容嗅探与门面格式枚举的一致性。Verifies extension mapping, content sniffing and facade format enum consistency.</summary>
        /// <param name="paths">各格式的样例文件路径。The sample path for each format.</param>
        private static void RunWordFormatSniffingTest(Dictionary<WordDocumentFormat, string> paths)
        {
            Require(WordDocumentFormats.FromExtension(".dotx") == WordDocumentFormat.Dotx, ".dotx maps to the template format");
            Require(WordDocumentFormats.FromExtension(".rtf") == WordDocumentFormat.Rtf, ".rtf maps to RTF");
            Require(WordDocumentFormats.FromExtension(".htm") == WordDocumentFormat.Html, ".htm maps to HTML");
            Require(WordDocumentFormats.FromExtension(".odt") == WordDocumentFormat.Odt, ".odt maps to ODT");
            Require(WordDocumentFormats.FromExtension(".txt") == WordDocumentFormat.Text, ".txt maps to plain text");
            Require(WordDocumentFormats.FromExtension(".doc") == WordDocumentFormat.Doc, ".doc maps to the legacy format");

            Require(OfficeFileFormats.FromExtension(".dotx") == OfficeFileFormat.Dotx, "facade .dotx mapping");
            Require(OfficeFileFormats.FromExtension(".rtf") == OfficeFileFormat.Rtf, "facade .rtf mapping");
            Require(OfficeFileFormats.FromExtension(".odt") == OfficeFileFormat.Odt, "facade .odt mapping");
            Require(OfficeFileFormats.FromExtension(".txt") == OfficeFileFormat.Text, "facade .txt mapping");
            Require(OfficeFileFormats.GetExtension(OfficeFileFormat.Odt) == ".odt", "facade ODT extension");
            Require(OfficeFileFormats.GetDisplayName(OfficeFileFormat.Html).Length > 0, "facade HTML display name");

            Require(OfficeFileFormats.Detect(paths[WordDocumentFormat.Docx]) == OfficeFileFormat.Docx, "facade detects DOCX");
            Require(OfficeFileFormats.Detect(paths[WordDocumentFormat.Dotx]) == OfficeFileFormat.Dotx, "facade detects DOTX");
            Require(OfficeFileFormats.Detect(paths[WordDocumentFormat.Rtf]) == OfficeFileFormat.Rtf, "facade detects RTF");
            Require(OfficeFileFormats.Detect(paths[WordDocumentFormat.Html]) == OfficeFileFormat.Html, "facade detects HTML");
            Require(OfficeFileFormats.Detect(paths[WordDocumentFormat.Odt]) == OfficeFileFormat.Odt, "facade detects ODT");

            OfficeFile file = OfficeFile.Load(paths[WordDocumentFormat.Odt]);
            Require(file.Document != null, "facade loads an ODT document");
            RequireContains(file.Document.GetPlainText(), "格式转换矩阵");
            byte[] rtfBytes = file.ToArray(OfficeFileFormat.Rtf);
            Require(OfficeFileFormats.Detect(rtfBytes) == OfficeFileFormat.Rtf, "facade saves RTF bytes");
            RequireContains(OfficeFile.Load(rtfBytes).Document.GetPlainText(), "格式转换矩阵");
        }
    }
}
