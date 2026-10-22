using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using Wisdom.Office.Core;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>执行 Word 增强能力测试：构建器、查找替换、邮件合并与文档克隆合并。Runs the Word enhancement tests: builder, find/replace, mail merge and document clone/merge.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWordEnhancementTest(string outputDirectory)
        {
            string builderPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-builder.docx"));
            RunWordBuilderTest(builderPath);
            RunWordFindReplaceTest();
            string mergePath = PrepareOutputPath(Path.Combine(outputDirectory, "word-mail-merge.docx"));
            RunWordMailMergeTest(mergePath);
            string regionPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-mail-merge-regions.docx"));
            RunWordMailMergeRegionTest(regionPath);
            string combinedPath = PrepareOutputPath(Path.Combine(outputDirectory, "word-merged.docx"));
            RunWordDocumentOperationsTest(combinedPath);
            Console.WriteLine("Word enhancement tests finished.");
        }

        /// <summary>验证 DocumentBuilder 风格 API 与超链接/书签/表格/图片/TOC 的读写往返。Verifies the DocumentBuilder-style API and the round trip of hyperlinks, bookmarks, tables, images and TOC fields.</summary>
        /// <param name="path">输出 .docx 路径。The output .docx path.</param>
        private static void RunWordBuilderTest(string path)
        {
            WordDocumentBuilder builder = new WordDocumentBuilder();
            builder.Font.FontFamily = "Microsoft YaHei";
            builder.Font.EastAsiaFontFamily = "Microsoft YaHei";
            builder.HeadingLevel = 1;
            builder.WriteLine("2026 年度报告");
            builder.HeadingLevel = 0;
            builder.InsertParagraph("正文第一段。");
            builder.InsertBookmark("Intro");
            builder.InsertParagraph("带链接：");
            builder.InsertHyperlink("官网", "https://example.com/report");
            builder.Write(" 与 ");
            builder.InsertHyperlink("目录", "#Intro");
            builder.InsertParagraph("不安全链接：");
            builder.InsertHyperlink("危险", "javascript:alert(1)");
            builder.InsertParagraph();
            builder.InsertImage(new OfficeImage(CreateTestPng(120, 40)), 120, 40);
            builder.InsertTableOfContents(3);
            builder.StartTable();
            builder.InsertCell("项目");
            builder.InsertCell("金额");
            builder.EndRow();
            builder.InsertCell("合计");
            builder.InsertCell("100");
            builder.EndTable();
            builder.InsertField("PAGE", "1");

            Require(builder.Document.Sections.Count == 1, "builder creates one section");
            Require(builder.Document.Sections[0].Blocks.Count >= 8, "builder appends blocks at the cursor");

            DocxWriter.Write(builder.Document, path);
            ValidateOoxmlPackage(path, "word/document.xml");

            OfficeDocument loaded = DocxReader.Load(path);
            Require(loaded.Sections.Count == 1, "builder round trip keeps one section");
            string text = loaded.GetPlainText();
            RequireContains(text, "2026 年度报告");
            RequireContains(text, "正文第一段。");
            RequireContains(text, "合计");
            RequireContains(text, "100");

            Paragraph intro = FindWordParagraph(loaded, "正文第一段。");
            Require(intro != null && intro.BookmarkName == "Intro", "bookmark round trips");

            Paragraph links = FindWordParagraph(loaded, "带链接：");
            string external = null;
            string anchor = null;
            int images = 0;
            if (links != null)
            {
                for (int i = 0; i < links.Inlines.Count; i++)
                {
                    TextRun run = links.Inlines[i] as TextRun;
                    if (run == null) continue;
                    if (run.Link == "https://example.com/report") external = run.Link;
                    if (run.Link == "#Intro") anchor = run.Link;
                }
            }
            Require(external != null, "external hyperlink round trips");
            Require(anchor != null, "anchor hyperlink round trips");

            bool hasToc = false;
            bool hasPageField = false;
            for (int s = 0; s < loaded.Sections.Count; s++)
            {
                List<Block> blocks = loaded.Sections[s].Blocks;
                for (int b = 0; b < blocks.Count; b++)
                {
                    Paragraph paragraph = blocks[b] as Paragraph;
                    if (paragraph != null)
                    {
                        for (int i = 0; i < paragraph.Inlines.Count; i++)
                        {
                            if (paragraph.Inlines[i] is ImageRun) images++;
                            FieldRun field = paragraph.Inlines[i] as FieldRun;
                            if (field == null) continue;
                            if (field.Instruction.StartsWith("TOC ", StringComparison.Ordinal)) hasToc = true;
                            if (field.Instruction.StartsWith("PAGE", StringComparison.Ordinal)) hasPageField = true;
                        }
                        continue;
                    }
                    Table table = blocks[b] as Table;
                    if (table != null && table.Rows.Count == 2 && table.Rows[1].Cells.Count == 2)
                    {
                        Require(table.Rows[0].Cells[0].Blocks.Count > 0, "builder table cell keeps a paragraph");
                    }
                }
            }
            Require(images == 1, "builder image round trips");
            Require(hasToc, "TOC field round trips");
            Require(hasPageField, "PAGE field round trips");

            string documentXml = ReadPackageEntryText(path, "word/document.xml");
            RequireContains(documentXml, "w:hyperlink");
            RequireContains(documentXml, "w:bookmarkStart");
            Require(!documentXml.Contains("javascript:"), "unsafe hyperlink schemes are not written");
            string relationships = ReadPackageEntryText(path, "word/_rels/document.xml.rels");
            RequireContains(relationships, "TargetMode=\"External\"");
        }

        /// <summary>验证跨运行、全词、正则、首处与表格/页眉页脚范围内的查找替换。Verifies cross-run, whole-word, regex, first-only and table/header/footer find/replace.</summary>
        private static void RunWordFindReplaceTest()
        {
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();
            Paragraph paragraph = new Paragraph();
            TextFormat bold = new TextFormat("Arial", 11) { Bold = true };
            paragraph.AddText("Hello ", bold);
            paragraph.AddText("World");
            paragraph.AddText("! Hello world.");
            section.Blocks.Add(paragraph);
            section.Header.AddParagraph("Confidential header");
            section.Footer.AddParagraph("Confidential footer");
            Table table = new Table();
            TableRow row = table.AddRow();
            TableCell cell = new TableCell();
            cell.Blocks.Add(new Paragraph("Confidential cell"));
            row.Cells.Add(cell);
            section.Blocks.Add(table);

            int count = WordFindReplace.Replace(document, "hello world", "Wisdom");
            Require(count == 2, "find/replace matches across runs and ignores case");
            Require(paragraph.GetText() == "Wisdom! Wisdom.", "cross-run replacement rewrites the paragraph");
            Require(paragraph.Inlines.Count == 2 && paragraph.Inlines[0] is TextRun, "emptied runs are removed");
            Require(((TextRun)paragraph.Inlines[0]).Format.Bold, "replacement keeps the first run formatting");

            count = WordFindReplace.Replace(document, "confidential", "内部");
            Require(count == 3, "find/replace covers header, footer and table cells");
            RequireContains(GetWordParagraphText(section.Header.Blocks[0]), "内部 header");
            RequireContains(GetWordParagraphText(section.Footer.Blocks[0]), "内部 footer");
            RequireContains(GetWordParagraphText(cell.Blocks[0]), "内部 cell");

            OfficeDocument scoped = new OfficeDocument();
            Section scopedSection = scoped.AddSection();
            scopedSection.Blocks.Add(new Paragraph("Confidential body"));
            scopedSection.Header.AddParagraph("Confidential header");
            WordFindReplaceOptions bodyOnly = new WordFindReplaceOptions();
            bodyOnly.IncludeHeadersAndFooters = false;
            count = WordFindReplace.Replace(scoped, "confidential", "内部", bodyOnly);
            Require(count == 1, "IncludeHeadersAndFooters=false restricts the scope");
            RequireContains(GetWordParagraphText(scopedSection.Header.Blocks[0]), "Confidential header");

            OfficeDocument words = new OfficeDocument();
            Paragraph wordParagraph = new Paragraph("cat catalog cat");
            words.AddSection().Blocks.Add(wordParagraph);
            WordFindReplaceOptions wholeWord = new WordFindReplaceOptions();
            wholeWord.WholeWord = true;
            count = WordFindReplace.Replace(words, "cat", "dog", wholeWord);
            Require(count == 2, "whole-word matching skips substrings");
            Require(wordParagraph.GetText() == "dog catalog dog", "whole-word replacement result");

            OfficeDocument first = new OfficeDocument();
            Paragraph firstParagraph = new Paragraph("cat cat cat");
            first.AddSection().Blocks.Add(firstParagraph);
            WordFindReplaceOptions firstWhole = new WordFindReplaceOptions();
            firstWhole.FirstOnly = true;
            firstWhole.WholeWord = true;
            count = WordFindReplace.Replace(first, "cat", "dog", firstWhole);
            Require(count == 1 && firstParagraph.GetText() == "dog cat cat", "FirstOnly replaces one occurrence");

            OfficeDocument regexDocument = new OfficeDocument();
            Paragraph codes = new Paragraph("编号 A-100、B-200");
            regexDocument.AddSection().Blocks.Add(codes);
            WordFindReplaceOptions regex = new WordFindReplaceOptions();
            regex.UseRegex = true;
            count = WordFindReplace.Replace(regexDocument, "[AB]-\\d{3}", "X", regex);
            Require(count == 2 && codes.GetText() == "编号 X、X", "regex replacement");

            OfficeDocument caseDocument = new OfficeDocument();
            Paragraph caseParagraph = new Paragraph("Word word");
            caseDocument.AddSection().Blocks.Add(caseParagraph);
            WordFindReplaceOptions matchCase = new WordFindReplaceOptions();
            matchCase.MatchCase = true;
            count = WordFindReplace.Replace(caseDocument, "word", "term", matchCase);
            Require(count == 1 && caseParagraph.GetText() == "Word term", "MatchCase distinguishes case");
        }

        /// <summary>验证字典、TemplateContext 与 DataRow 邮件合并，包括跨运行标记、格式开关、图片与未解析处理。Verifies mail merge from dictionaries, TemplateContext and DataRow, including cross-run tokens, format switches, images and unresolved handling.</summary>
        /// <param name="path">输出 .docx 路径。The output .docx path.</param>
        private static void RunWordMailMergeTest(string path)
        {
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();
            section.Header.AddParagraph("报告：{{title}}");

            Paragraph paragraph = new Paragraph();
            paragraph.AddText("客户：{{customer");
            paragraph.AddText(".name}}");
            paragraph.AddText(" 金额：{{amount:N2}} 元");
            paragraph.AddText(" 余额：");
            paragraph.Inlines.Add(new FieldRun("MERGEFIELD Balance \\# \"#,##0.00\"", string.Empty, new TextFormat("Arial", 10)));
            paragraph.AddText(" 标志：");
            paragraph.Inlines.Add(new FieldRun("MERGEFIELD Logo", string.Empty, new TextFormat("Arial", 10)));
            paragraph.AddText(" 缺失：");
            paragraph.Inlines.Add(new FieldRun("MERGEFIELD Missing", string.Empty, new TextFormat("Arial", 10)));
            section.Blocks.Add(paragraph);

            TemplateContext customer = new TemplateContext();
            customer.Set("name", "海森公司");
            TemplateContext context = new TemplateContext();
            context.Set("title", "季度报告");
            context.Set("customer", customer);
            context.Set("amount", 1234.5m);
            context.Set("Balance", 9876.5m);
            context.Set("Logo", new OfficeImage(CreateTestPng(40, 20)));

            WordMailMergeOptions options = new WordMailMergeOptions();
            options.LeaveUnresolvedFields = false;
            int merged = WordMailMerge.Execute(document, context, options);
            Require(merged == 6, "mail merge counts fields and tokens");
            RequireContains(GetWordParagraphText(section.Header.Blocks[0]), "报告：季度报告");

            string text = paragraph.GetText();
            RequireContains(text, "客户：海森公司");
            RequireContains(text, "金额：1,234.50 元");
            RequireContains(text, "余额：9,876.50");
            bool imageInserted = false;
            bool missingFieldGone = true;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                if (paragraph.Inlines[i] is ImageRun) imageInserted = true;
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field != null && field.Instruction.IndexOf("Missing", StringComparison.Ordinal) >= 0) missingFieldGone = false;
            }
            Require(imageInserted, "image value becomes an ImageRun");
            Require(missingFieldGone, "unresolved field is cleared when configured");

            DocxWriter.Write(document, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            OfficeDocument reloaded = DocxReader.Load(path);
            RequireContains(reloaded.GetPlainText(), "客户：海森公司");
            int reloadedImages = 0;
            Paragraph body = FindWordParagraph(reloaded, "客户：");
            if (body != null)
            {
                for (int i = 0; i < body.Inlines.Count; i++)
                {
                    if (body.Inlines[i] is ImageRun) reloadedImages++;
                }
            }
            Require(reloadedImages == 1, "merged image round trips");

            OfficeDocument keepDocument = new OfficeDocument();
            Paragraph keepParagraph = new Paragraph();
            keepParagraph.AddText("保留 {{unknown}} ");
            keepParagraph.Inlines.Add(new FieldRun("MERGEFIELD Unknown", string.Empty, new TextFormat()));
            keepDocument.AddSection().Blocks.Add(keepParagraph);
            int kept = WordMailMerge.Execute(keepDocument, new TemplateContext());
            Require(kept == 0, "unresolved fields are left untouched by default");
            RequireContains(keepParagraph.GetText(), "{{unknown}}");
            bool fieldKept = false;
            for (int i = 0; i < keepParagraph.Inlines.Count; i++)
            {
                if (keepParagraph.Inlines[i] is FieldRun) fieldKept = true;
            }
            Require(fieldKept, "unresolved MERGEFIELD is preserved by default");

            DataTable table = new DataTable("Report");
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Note", typeof(string));
            DataRow row = table.NewRow();
            row["Name"] = "数据行客户";
            row["Note"] = DBNull.Value;
            table.Rows.Add(row);

            OfficeDocument rowDocument = new OfficeDocument();
            Paragraph rowParagraph = new Paragraph();
            rowParagraph.Inlines.Add(new FieldRun("MERGEFIELD Name", string.Empty, new TextFormat()));
            rowParagraph.AddText(" / {{Name}} / ");
            rowParagraph.Inlines.Add(new FieldRun("MERGEFIELD Note", string.Empty, new TextFormat()));
            rowDocument.AddSection().Blocks.Add(rowParagraph);
            int rowCount = WordMailMerge.Execute(rowDocument, row);
            Require(rowCount == 3, "DataRow merge handles fields and tokens");
            Require(rowParagraph.GetText().Trim() == "数据行客户 / 数据行客户 /", "DataRow merge writes column values and maps DBNull to empty");
        }

        /// <summary>验证邮件合并区域：表格行区域、嵌套区域、段落块区域、DataTable + DataRelation，以及 DOCX 往返。Verifies mail-merge regions: table-row regions, nested regions, paragraph-block regions, DataTable + DataRelation, and the DOCX round trip.</summary>
        /// <param name="path">输出 .docx 路径。The output .docx path.</param>
        private static void RunWordMailMergeRegionTest(string path)
        {
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();
            section.Blocks.Add(new Paragraph("报价单：{{title}}"));

            Table table = new Table();
            table.Columns.Add(new TableColumn(90));
            table.Columns.Add(new TableColumn(160));

            TableRow header = table.AddRow();
            TableCell headerMarker = new TableCell();
            headerMarker.Blocks.Add(CreateRegionMarker("MERGEFIELD TableStart:orders"));
            header.Cells.Add(headerMarker);
            header.Cells.Add(new TableCell("订单明细"));
            header.IsHeader = true;

            TableRow orderRow = table.AddRow();
            orderRow.Cells.Add(new TableCell("{{orderNo}}"));
            orderRow.Cells.Add(new TableCell("{{customer}}"));

            TableRow linesStart = table.AddRow();
            TableCell linesStartMarker = new TableCell();
            linesStartMarker.Blocks.Add(CreateRegionMarker("MERGEFIELD TableStart:lines"));
            linesStart.Cells.Add(linesStartMarker);

            TableRow lineRow = table.AddRow();
            lineRow.Cells.Add(new TableCell("{{sku}}"));
            lineRow.Cells.Add(new TableCell("{{qty}}"));

            TableRow linesEnd = table.AddRow();
            TableCell linesEndMarker = new TableCell();
            linesEndMarker.Blocks.Add(CreateRegionMarker("MERGEFIELD TableEnd:lines"));
            linesEnd.Cells.Add(linesEndMarker);

            TableRow ordersEnd = table.AddRow();
            TableCell ordersEndMarker = new TableCell();
            ordersEndMarker.Blocks.Add(CreateRegionMarker("MERGEFIELD TableEnd:orders"));
            ordersEnd.Cells.Add(ordersEndMarker);

            section.Blocks.Add(table);

            TemplateContext root = new TemplateContext();
            root.Set("title", "年度报价");
            List<object> orders = new List<object>();
            orders.Add(CreateOrder("A-001", "海森公司", new string[] { "键盘", "鼠标" }));
            orders.Add(CreateOrder("A-002", "远山公司", new string[] { "显示器" }));
            root.Set("orders", orders);

            int merged = WordMailMerge.ExecuteWithRegions(document, root);
            Require(merged == 11, "region merge counts every nested scalar");

            Require(table.Rows.Count == 6, "nested table regions expand to header plus order and line rows");
            RequireContains(GetWordParagraphText(table.Rows[0].Cells[1].Blocks[0]), "订单明细");
            RequireContains(GetWordParagraphText(table.Rows[1].Cells[0].Blocks[0]), "A-001");
            RequireContains(GetWordParagraphText(table.Rows[1].Cells[1].Blocks[0]), "海森公司");
            RequireContains(GetWordParagraphText(table.Rows[2].Cells[0].Blocks[0]), "键盘");
            RequireContains(GetWordParagraphText(table.Rows[3].Cells[0].Blocks[0]), "鼠标");
            RequireContains(GetWordParagraphText(table.Rows[4].Cells[0].Blocks[0]), "A-002");
            RequireContains(GetWordParagraphText(table.Rows[5].Cells[0].Blocks[0]), "显示器");
            Require(WordDocumentContainsField(table, "TableStart:") == false, "region markers are removed from the output");

            OfficeDocument paragraphDocument = new OfficeDocument();
            Section paragraphSection = paragraphDocument.AddSection();
            paragraphSection.Blocks.Add(new Paragraph("合计 {{total}}"));
            paragraphSection.Blocks.Add(CreateRegionMarker("MERGEFIELD TableStart:groups"));
            paragraphSection.Blocks.Add(new Paragraph("组：{{name}}"));
            paragraphSection.Blocks.Add(CreateRegionMarker("MERGEFIELD TableStart:members"));
            paragraphSection.Blocks.Add(new Paragraph(" 成员：{{name}}"));
            paragraphSection.Blocks.Add(CreateRegionMarker("MERGEFIELD TableEnd:members"));
            paragraphSection.Blocks.Add(CreateRegionMarker("MERGEFIELD TableEnd:groups"));

            TemplateContext paragraphRoot = new TemplateContext();
            paragraphRoot.Set("total", 3);
            List<object> groups = new List<object>();
            groups.Add(CreateGroup("华东", new string[] { "张三", "李四" }));
            groups.Add(CreateGroup("华南", new string[] { "王五" }));
            paragraphRoot.Set("groups", groups);

            int paragraphMerged = WordMailMerge.ExecuteWithRegions(paragraphDocument, paragraphRoot);
            Require(paragraphMerged == 6, "paragraph regions merge nested tokens");
            string plain = paragraphDocument.GetPlainText();
            RequireContains(plain, "合计 3");
            RequireContains(plain, "组：华东");
            RequireContains(plain, "成员：张三");
            RequireContains(plain, "成员：王五");
            Require(!plain.Contains("TableStart") && !plain.Contains("TableEnd"), "paragraph region markers are removed");
            Require(CountParagraphs(paragraphSection.Blocks) == 6, "paragraph region repeats only the body blocks");

            DataTable orderTable = new DataTable("Orders");
            orderTable.Columns.Add("OrderNo", typeof(string));
            DataTable lineTable = new DataTable("Lines");
            lineTable.Columns.Add("OrderNo", typeof(string));
            lineTable.Columns.Add("Sku", typeof(string));
            DataRow orderA = orderTable.NewRow();
            orderA["OrderNo"] = "B-001";
            orderTable.Rows.Add(orderA);
            DataRow orderB = orderTable.NewRow();
            orderB["OrderNo"] = "B-002";
            orderTable.Rows.Add(orderB);
            lineTable.Rows.Add("B-001", "笔");
            lineTable.Rows.Add("B-001", "纸");
            lineTable.Rows.Add("B-002", "胶带");
            DataSet dataSet = new DataSet();
            dataSet.Tables.Add(orderTable);
            dataSet.Tables.Add(lineTable);
            dataSet.Relations.Add("OrderLines", orderTable.Columns["OrderNo"], lineTable.Columns["OrderNo"]);

            OfficeDocument dataTableDocument = new OfficeDocument();
            Section dataTableSection = dataTableDocument.AddSection();
            dataTableSection.Blocks.Add(CreateRelationTable());
            int relationMerged = WordMailMerge.ExecuteWithRegions(dataTableDocument, dataSet);
            Require(relationMerged == 5, "DataSet relations expand nested regions");
            Table relationResult = dataTableSection.Blocks[0] as Table;
            Require(relationResult != null && relationResult.Rows.Count == 5, "related child rows repeat inside the parent region");
            RequireContains(GetWordParagraphText(relationResult.Rows[0].Cells[0].Blocks[0]), "B-001");
            RequireContains(GetWordParagraphText(relationResult.Rows[1].Cells[1].Blocks[0]), "笔");
            RequireContains(GetWordParagraphText(relationResult.Rows[2].Cells[1].Blocks[0]), "纸");
            RequireContains(GetWordParagraphText(relationResult.Rows[3].Cells[0].Blocks[0]), "B-002");
            RequireContains(GetWordParagraphText(relationResult.Rows[4].Cells[1].Blocks[0]), "胶带");

            DocxWriter.Write(document, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            OfficeDocument reloaded = DocxReader.Load(path);
            Require(reloaded.Sections.Count == 1, "region merged document round trips");
            RequireContains(reloaded.GetPlainText(), "年度报价");
            RequireContains(reloaded.GetPlainText(), "键盘");
            RequireContains(reloaded.GetPlainText(), "显示器");
        }

        /// <summary>构建订单测试数据，行内包含明细列表。Builds an order test object carrying a line-item list.</summary>
        private static Dictionary<string, object> CreateOrder(string orderNo, string customer, string[] skus)
        {
            Dictionary<string, object> order = new Dictionary<string, object>();
            order["orderNo"] = orderNo;
            order["customer"] = customer;
            List<object> lines = new List<object>();
            for (int i = 0; i < skus.Length; i++)
            {
                Dictionary<string, object> line = new Dictionary<string, object>();
                line["sku"] = skus[i];
                line["qty"] = (i + 1) * 2;
                lines.Add(line);
            }
            order["lines"] = lines;
            return order;
        }

        /// <summary>构建分组测试数据，行内包含成员列表。Builds a group test object carrying a member list.</summary>
        private static Dictionary<string, object> CreateGroup(string name, string[] members)
        {
            Dictionary<string, object> group = new Dictionary<string, object>();
            group["name"] = name;
            List<object> items = new List<object>();
            for (int i = 0; i < members.Length; i++)
            {
                Dictionary<string, object> member = new Dictionary<string, object>();
                member["name"] = members[i];
                items.Add(member);
            }
            group["members"] = items;
            return group;
        }

        /// <summary>创建只含一个域的区域标记段落。Creates a region-marker paragraph holding a single field.</summary>
        private static Paragraph CreateRegionMarker(string instruction)
        {
            Paragraph paragraph = new Paragraph();
            paragraph.Inlines.Add(new FieldRun(instruction, string.Empty, new TextFormat("Arial", 9)));
            return paragraph;
        }

        /// <summary>构建 DataRelation 嵌套测试表格：父区域 Orders，子区域 Lines。Builds the DataRelation nesting test table: parent Orders and child Lines.</summary>
        private static Table CreateRelationTable()
        {
            Table table = new Table();
            table.Columns.Add(new TableColumn(90));
            table.Columns.Add(new TableColumn(90));
            table.Rows.Add(CreateMarkerRow("MERGEFIELD TableStart:Orders", 2));
            TableRow parentRow = table.AddRow();
            parentRow.Cells.Add(new TableCell("{{OrderNo}}"));
            parentRow.Cells.Add(new TableCell());
            TableRow childStart = table.AddRow();
            childStart.Cells.Add(new TableCell());
            childStart.Cells[0].Blocks.Add(CreateRegionMarker("MERGEFIELD TableStart:Lines"));
            childStart.Cells.Add(new TableCell());
            TableRow childRow = table.AddRow();
            childRow.Cells.Add(new TableCell());
            childRow.Cells.Add(new TableCell("{{Sku}}"));
            TableRow childEnd = table.AddRow();
            childEnd.Cells.Add(new TableCell());
            childEnd.Cells[0].Blocks.Add(CreateRegionMarker("MERGEFIELD TableEnd:Lines"));
            childEnd.Cells.Add(new TableCell());
            table.Rows.Add(CreateMarkerRow("MERGEFIELD TableEnd:Orders", 2));
            return table;
        }

        /// <summary>构建跨指定列数的标记行。Builds a marker row spanning the given number of cells.</summary>
        private static TableRow CreateMarkerRow(string instruction, int cellCount)
        {
            TableRow row = new TableRow();
            for (int i = 0; i < cellCount; i++)
            {
                TableCell cell = new TableCell();
                if (i == 0) cell.Blocks.Add(CreateRegionMarker(instruction));
                row.Cells.Add(cell);
            }
            return row;
        }

        /// <summary>检查表格内是否仍残留指定区域标记。Checks whether the table still contains the given region marker text.</summary>
        private static bool WordDocumentContainsField(Table table, string instructionPart)
        {
            for (int r = 0; r < table.Rows.Count; r++)
            {
                for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                {
                    List<Block> blocks = table.Rows[r].Cells[c].Blocks;
                    for (int b = 0; b < blocks.Count; b++)
                    {
                        Paragraph paragraph = blocks[b] as Paragraph;
                        if (paragraph == null) continue;
                        for (int i = 0; i < paragraph.Inlines.Count; i++)
                        {
                            FieldRun field = paragraph.Inlines[i] as FieldRun;
                            if (field != null && field.Instruction.IndexOf(instructionPart, StringComparison.Ordinal) >= 0) return true;
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>验证文档深拷贝、追加与合并保留节、页面设置、页眉页脚，并且互不影响。Verifies that deep clone, append and merge preserve sections, page setup and headers/footers without aliasing.</summary>
        /// <param name="path">输出 .docx 路径。The output .docx path.</param>
        private static void RunWordDocumentOperationsTest(string path)
        {
            OfficeDocument first = new OfficeDocument();
            first.Metadata.Title = "文档 A";
            PageSetup letter = new PageSetup();
            letter.Width = 612;
            letter.Height = 792;
            Section firstSection = first.AddSection(letter);
            firstSection.Header.AddParagraph("Header A");
            firstSection.Blocks.Add(new Paragraph("First body"));

            OfficeDocument second = new OfficeDocument();
            Section secondSection = second.AddSection();
            secondSection.PageSetup.IsLandscape = true;
            secondSection.Footer.AddParagraph("Footer B");
            secondSection.Blocks.Add(new Paragraph("Second body"));

            OfficeDocument merged = WordDocumentOperations.Merge(first, second);
            Require(merged.Sections.Count == 2, "Merge keeps every section");
            Require(merged.Metadata.Title == "文档 A", "Merge keeps the first document metadata");
            Require(merged.Sections[0].Header.Blocks.Count == 1, "Merge keeps the first header");
            Require(merged.Sections[1].Footer.Blocks.Count == 1, "Merge keeps the second footer");
            Require(merged.Sections[1].PageSetup.IsLandscape, "Merge keeps the second page orientation");
            RequireContains(merged.GetPlainText(), "First body");
            RequireContains(merged.GetPlainText(), "Second body");

            OfficeDocument clone = WordDocumentOperations.Clone(merged);
            Paragraph cloneParagraph = clone.Sections[0].Blocks[0] as Paragraph;
            cloneParagraph.AddText(" mutated");
            Require(!GetWordParagraphText(merged.Sections[0].Blocks[0]).Contains("mutated"), "Clone detaches content from the source");

            Paragraph sourceParagraph = second.Sections[0].Blocks[0] as Paragraph;
            sourceParagraph.AddText(" source-only");
            Require(!GetWordParagraphText(merged.Sections[1].Blocks[0]).Contains("source-only"), "Merge deep-copies the source content");

            OfficeDocument appended = WordDocumentOperations.Clone(first);
            int added = WordDocumentOperations.Append(appended, second);
            Require(added == 1 && appended.Sections.Count == 2, "Append returns the appended section count");

            DocxWriter.Write(merged, path);
            ValidateOoxmlPackage(path, "word/document.xml");
            OfficeDocument reloaded = DocxReader.Load(path);
            Require(reloaded.Sections.Count == 2, "merged document round trips both sections");
            RequireContains(reloaded.GetPlainText(), "First body");
            RequireContains(reloaded.GetPlainText(), "Second body");
            Require(reloaded.Sections[1].PageSetup.IsLandscape, "merged orientation round trips");
            RequireContains(GetWordParagraphText(reloaded.Sections[0].Header.Blocks[0]), "Header A");
            RequireContains(GetWordParagraphText(reloaded.Sections[1].Footer.Blocks[0]), "Footer B");
        }

        /// <summary>读取块的段落文本；非段落块返回空串。Reads a block's paragraph text; non-paragraph blocks return an empty string.</summary>
        /// <param name="block">待读取的块。The block to read.</param>
        /// <returns>段落文本或空串。The paragraph text or an empty string.</returns>
        private static string GetWordParagraphText(Block block)
        {
            Paragraph paragraph = block as Paragraph;
            return paragraph == null ? string.Empty : paragraph.GetText();
        }

        /// <summary>在正文、表格与页眉页脚中查找包含指定文本的段落。Finds the first paragraph containing the given text in the body, tables, headers or footers.</summary>
        /// <param name="document">待搜索的文档。The document to search.</param>
        /// <param name="text">包含的文本片段。The text fragment to look for.</param>
        /// <returns>命中的段落，未找到时为 null。The matching paragraph, or null when nothing matches.</returns>
        private static Paragraph FindWordParagraph(OfficeDocument document, string text)
        {
            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                Paragraph found = FindWordBlockParagraph(section.Blocks, text);
                if (found != null) return found;
                if (section.Header != null)
                {
                    found = FindWordBlockParagraph(section.Header.Blocks, text);
                    if (found != null) return found;
                }
                if (section.Footer != null)
                {
                    found = FindWordBlockParagraph(section.Footer.Blocks, text);
                    if (found != null) return found;
                }
            }
            return null;
        }

        /// <summary>递归查找块列表（含表格单元格）中包含指定文本的段落。Recursively finds a paragraph containing the given text in a block list, including table cells.</summary>
        /// <param name="blocks">块列表。The block list.</param>
        /// <param name="text">包含的文本片段。The text fragment to look for.</param>
        /// <returns>命中的段落，未找到时为 null。The matching paragraph, or null when nothing matches.</returns>
        private static Paragraph FindWordBlockParagraph(List<Block> blocks, string text)
        {
            for (int i = 0; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph != null)
                {
                    if (paragraph.GetText().Contains(text)) return paragraph;
                    continue;
                }
                Table table = blocks[i] as Table;
                if (table == null) continue;
                for (int r = 0; r < table.Rows.Count; r++)
                {
                    for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                    {
                        Paragraph found = FindWordBlockParagraph(table.Rows[r].Cells[c].Blocks, text);
                        if (found != null) return found;
                    }
                }
            }
            return null;
        }
    }
}
