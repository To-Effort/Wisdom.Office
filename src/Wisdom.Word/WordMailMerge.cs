using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using Wisdom.Office.Core;

namespace Wisdom.Word
{
    /// <summary>
    /// 邮件合并选项：未解析域的处理、格式化文化与搜索范围。
    /// Mail-merge options: unresolved-field handling, formatting culture and search scope.
    /// </summary>
    public sealed class WordMailMergeOptions
    {
        /// <summary>数据源中不存在的 MERGEFIELD / 标记是否原样保留，默认保留。Whether MERGEFIELD fields and tokens missing from the data source are left in place; true by default.</summary>
        public bool LeaveUnresolvedFields { get; set; } = true;
        /// <summary>格式化取值所用的区域文化，默认不变文化。The culture used to format values; the invariant culture by default.</summary>
        public CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
        /// <summary>是否合并页眉和页脚中的域与标记，默认是。Whether fields and tokens in headers and footers are merged; true by default.</summary>
        public bool IncludeHeadersAndFooters { get; set; } = true;
    }

    /// <summary>
    /// 模型级邮件合并引擎：用字典、DataRow 或 <see cref="TemplateContext"/> 填充 MERGEFIELD 域与
    /// {{token}} / {{token:format}} 标记。与 <see cref="DocxTemplate"/> 的包级实现互补——本引擎直接操作统一模型，
    /// 因此可以写入图片并保留程序化构建的文档结构。
    /// Model-level mail-merge engine: fills MERGEFIELD fields and {{token}} / {{token:format}} markers from a
    /// dictionary, a DataRow or a <see cref="TemplateContext"/>. It complements the package-level
    /// <see cref="DocxTemplate"/> by working on the unified model, which lets it insert images and preserve
    /// programmatically built documents.
    /// </summary>
    /// <remarks>
    /// 合并后的 MERGEFIELD 会转成普通文本或图文，避免 Word 更新域时又把结果清空；图片只支持
    /// <see cref="OfficeImage"/> 或图片字节数组取值，不接受文件路径。
    /// <see cref="ExecuteWithRegions(OfficeDocument, TemplateContext, WordMailMergeOptions)"/> 额外支持 Aspose 风格的 <c>MERGEFIELD TableStart:名称</c> /
    /// <c>MERGEFIELD TableEnd:名称</c> 区域：表格行或段落块按集合逐项重复，区域可以任意深度嵌套。
    /// Merged MERGEFIELD fields become plain text or images so that updating fields in Word cannot wipe the
    /// results; image values must be <see cref="OfficeImage"/> instances or image byte arrays, never file paths.
    /// <see cref="ExecuteWithRegions(OfficeDocument, TemplateContext, WordMailMergeOptions)"/> additionally supports Aspose-style
    /// <c>MERGEFIELD TableStart:Name</c> / <c>MERGEFIELD TableEnd:Name</c> regions: table rows or paragraph
    /// blocks repeat per collection item, and regions may nest to any depth.
    /// </remarks>
    public static class WordMailMerge
    {
        /// <summary>区域嵌套的最大深度，防止自引用数据导致无限展开。The maximum region nesting depth, preventing unbounded expansion for self-referencing data.</summary>
        private const int MaxRegionDepth = 32;

        /// <summary>区域标记的种类。The kind of a region marker.</summary>
        private enum RegionMarkerKind
        {
            /// <summary>非区域标记。Not a region marker.</summary>
            None,
            /// <summary>区域起点。A region start.</summary>
            Start,
            /// <summary>区域终点。A region end.</summary>
            End
        }

        /// <summary>匹配 {{key}} 与 {{key:format}} 标记，排除 #each 循环标记。Matches {{key}} and {{key:format}} tokens, excluding #each loop markers.</summary>
        private static readonly Regex TokenRegex = new Regex(
            @"\{\{\s*([^#/{}][^{}:]*?)\s*(?::\s*([^{}]+?))?\s*\}\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>解析 MERGEFIELD 的数字/日期格式开关（\# 或 \@）。Parses the numeric/date format switch (\# or \@) of a MERGEFIELD instruction.</summary>
        private static readonly Regex FieldFormatRegex = new Regex(
            @"\\\s*([#@])\s*(?:""([^""]*)""|([^\s\\]+))",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>
        /// 用字符串字典合并文档中的域与标记。
        /// Merges fields and tokens using a string dictionary.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="values">键值数据，不能为 null。The key-value data; must not be null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="values"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="values"/> is null.</exception>
        public static int Execute(OfficeDocument document, IDictionary<string, object> values)
        {
            return Execute(document, values, new WordMailMergeOptions());
        }

        /// <summary>
        /// 用字符串字典与指定选项合并文档中的域与标记。
        /// Merges fields and tokens using a string dictionary and the given options.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="values">键值数据，不能为 null。The key-value data; must not be null.</param>
        /// <param name="options">合并选项，null 时使用默认值。The options; defaults are used when null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="values"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="values"/> is null.</exception>
        public static int Execute(OfficeDocument document, IDictionary<string, object> values, WordMailMergeOptions options)
        {
            if (values == null) throw new ArgumentNullException("values");
            return Execute(document, TemplateEngine.CreateRowContext(values), options);
        }

        /// <summary>
        /// 用数据行合并文档：列名作为域/标记名，DBNull 视为空值。
        /// Merges the document with a data row: column names act as field/token names and DBNull becomes null.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="row">数据行。The data row.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 为 null。Thrown when <paramref name="document"/> is null.</exception>
        public static int Execute(OfficeDocument document, DataRow row)
        {
            return Execute(document, row, new WordMailMergeOptions());
        }

        /// <summary>
        /// 用数据行与指定选项合并文档：列名作为域/标记名，DBNull 视为空值。
        /// Merges the document with a data row and the given options: column names act as field/token names and
        /// DBNull becomes null.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="row">数据行，null 时所有域按未解析处理。The data row; null makes every field unresolved.</param>
        /// <param name="options">合并选项，null 时使用默认值。The options; defaults are used when null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 为 null。Thrown when <paramref name="document"/> is null.</exception>
        public static int Execute(OfficeDocument document, DataRow row, WordMailMergeOptions options)
        {
            return Execute(document, CreateRowContext(row), options);
        }

        /// <summary>
        /// 用模板上下文合并文档中的域与标记，支持点号路径（如 customer.name）。
        /// Merges fields and tokens using a template context, with dotted-path support (such as customer.name).
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="context">模板数据上下文，不能为 null。The template data context; must not be null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="context"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="context"/> is null.</exception>
        public static int Execute(OfficeDocument document, TemplateContext context)
        {
            return Execute(document, context, new WordMailMergeOptions());
        }

        /// <summary>
        /// 用模板上下文与指定选项合并文档中的域与标记，支持点号路径。
        /// Merges fields and tokens using a template context and the given options, with dotted-path support.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="context">模板数据上下文，不能为 null。The template data context; must not be null.</param>
        /// <param name="options">合并选项，null 时使用默认值。The options; defaults are used when null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="context"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="context"/> is null.</exception>
        public static int Execute(OfficeDocument document, TemplateContext context, WordMailMergeOptions options)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (context == null) throw new ArgumentNullException("context");
            if (options == null) options = new WordMailMergeOptions();

            int count = 0;
            WordDocumentWalker.ForEachParagraph(
                document,
                options.IncludeHeadersAndFooters,
                true,
                delegate (Paragraph paragraph)
                {
                    count += MergeFields(paragraph, context, options);
                    count += MergeTokens(paragraph, context, options);
                });
            return count;
        }

        /// <summary>
        /// 用字符串字典展开邮件合并区域。字典值为集合时对应同名区域，标量作为全局值。
        /// Expands mail-merge regions from a string dictionary. Collection values feed the region of the same
        /// name; scalar values are available globally.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="values">键值数据，不能为 null。The key-value data; must not be null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="values"/> 为 null。Thrown when <paramref name="values"/> is null.</exception>
        public static int ExecuteWithRegions(OfficeDocument document, IDictionary<string, object> values)
        {
            return ExecuteWithRegions(document, values, new WordMailMergeOptions());
        }

        /// <summary>
        /// 用字符串字典与指定选项展开邮件合并区域。
        /// Expands mail-merge regions from a string dictionary with the given options.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="values">键值数据，不能为 null。The key-value data; must not be null.</param>
        /// <param name="options">合并选项，null 时使用默认值。The options; defaults are used when null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="values"/> 为 null。Thrown when <paramref name="values"/> is null.</exception>
        public static int ExecuteWithRegions(OfficeDocument document, IDictionary<string, object> values, WordMailMergeOptions options)
        {
            if (values == null) throw new ArgumentNullException("values");
            return ExecuteWithRegions(document, TemplateEngine.CreateRowContext(values), options);
        }

        /// <summary>
        /// 用数据表展开邮件合并区域：表名对应区域名，列名对应域名，DataRelation 自动形成嵌套区域。
        /// Expands mail-merge regions from a data table: the table name is the region name, column names are
        /// field names, and <see cref="DataRelation"/> children become nested regions.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="data">数据表，不能为 null；表名为空时使用 "Table"。The data table; must not be null. An empty table name is treated as "Table".</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null。Thrown when <paramref name="data"/> is null.</exception>
        public static int ExecuteWithRegions(OfficeDocument document, DataTable data)
        {
            return ExecuteWithRegions(document, data, new WordMailMergeOptions());
        }

        /// <summary>
        /// 用数据表与指定选项展开邮件合并区域：表名对应区域名，DataRelation 自动形成嵌套区域。
        /// Expands mail-merge regions from a data table with the given options: the table name is the region
        /// name and <see cref="DataRelation"/> children become nested regions.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="data">数据表，不能为 null。The data table; must not be null.</param>
        /// <param name="options">合并选项，null 时使用默认值。The options; defaults are used when null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null。Thrown when <paramref name="data"/> is null.</exception>
        public static int ExecuteWithRegions(OfficeDocument document, DataTable data, WordMailMergeOptions options)
        {
            if (data == null) throw new ArgumentNullException("data");
            TemplateContext root = new TemplateContext();
            string regionName = string.IsNullOrEmpty(data.TableName) ? "Table" : data.TableName;
            root.Set(regionName, data);
            return ExecuteWithRegions(document, root, options);
        }

        /// <summary>
        /// 用数据集展开邮件合并区域：每个表按表名注册为顶层区域，表间 DataRelation 形成嵌套区域。
        /// Expands mail-merge regions from a data set: every table is registered as a top-level region under
        /// its table name, and <see cref="DataRelation"/> links form nested regions.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="data">数据集，不能为 null。The data set; must not be null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null。Thrown when <paramref name="data"/> is null.</exception>
        public static int ExecuteWithRegions(OfficeDocument document, DataSet data)
        {
            return ExecuteWithRegions(document, data, new WordMailMergeOptions());
        }

        /// <summary>
        /// 用数据集与指定选项展开邮件合并区域。
        /// Expands mail-merge regions from a data set with the given options.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="data">数据集，不能为 null。The data set; must not be null.</param>
        /// <param name="options">合并选项，null 时使用默认值。The options; defaults are used when null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null。Thrown when <paramref name="data"/> is null.</exception>
        public static int ExecuteWithRegions(OfficeDocument document, DataSet data, WordMailMergeOptions options)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (options == null) options = new WordMailMergeOptions();
            TemplateContext root = new TemplateContext();
            for (int i = 0; i < data.Tables.Count; i++)
            {
                DataTable table = data.Tables[i];
                if (string.IsNullOrEmpty(table.TableName)) continue;
                root.Set(table.TableName, table);
            }
            return ExecuteWithRegions(document, root, options);
        }

        /// <summary>
        /// 用模板上下文展开邮件合并区域：集合值对应同名区域，嵌套集合自动成为嵌套区域。
        /// Expands mail-merge regions from a template context: collection values feed same-named regions and
        /// nested collections become nested regions automatically.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="context">模板数据上下文，不能为 null。The template data context; must not be null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="context"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="context"/> is null.</exception>
        public static int ExecuteWithRegions(OfficeDocument document, TemplateContext context)
        {
            return ExecuteWithRegions(document, context, new WordMailMergeOptions());
        }

        /// <summary>
        /// 用模板上下文与指定选项展开邮件合并区域，支持任意深度嵌套。
        /// Expands mail-merge regions from a template context with the given options, supporting arbitrary
        /// nesting depth.
        /// </summary>
        /// <param name="document">目标文档。The target document.</param>
        /// <param name="context">模板数据上下文，不能为 null。The template data context; must not be null.</param>
        /// <param name="options">合并选项，null 时使用默认值。The options; defaults are used when null.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="document"/> 或 <paramref name="context"/> 为 null。Thrown when <paramref name="document"/> or <paramref name="context"/> is null.</exception>
        public static int ExecuteWithRegions(OfficeDocument document, TemplateContext context, WordMailMergeOptions options)
        {
            if (document == null) throw new ArgumentNullException("document");
            if (context == null) throw new ArgumentNullException("context");
            if (options == null) options = new WordMailMergeOptions();

            int count = 0;
            for (int s = 0; s < document.Sections.Count; s++)
            {
                Section section = document.Sections[s];
                count += ProcessBlocks(section.Blocks, context, options, 0);
                if (!options.IncludeHeadersAndFooters) continue;
                if (section.Header != null) count += ProcessBlocks(section.Header.Blocks, context, options, 0);
                if (section.Footer != null) count += ProcessBlocks(section.Footer.Blocks, context, options, 0);
            }
            return count;
        }

        /// <summary>
        /// 递归处理块列表：展开区域、复制区域模板，再合并区域内的域与标记。
        /// Recursively processes a block list: expands regions, clones region templates and merges the fields
        /// and tokens inside each clone.
        /// </summary>
        /// <param name="blocks">待处理的块列表，处理后就地重写。The block list, rewritten in place.</param>
        /// <param name="context">当前作用域的数据上下文。The data context of the current scope.</param>
        /// <param name="options">合并选项。The merge options.</param>
        /// <param name="depth">当前区域嵌套深度。The current region nesting depth.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        private static int ProcessBlocks(List<Block> blocks, TemplateContext context, WordMailMergeOptions options, int depth)
        {
            int count = 0;
            List<Block> output = new List<Block>(blocks.Count);
            int i = 0;
            while (i < blocks.Count)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                RegionMarkerKind kind;
                string name;
                if (paragraph != null
                    && TryGetParagraphRegionMarker(paragraph, out kind, out name)
                    && kind == RegionMarkerKind.Start
                    && depth < MaxRegionDepth)
                {
                    int end = FindRegionEndParagraph(blocks, i + 1, name);
                    if (end >= 0)
                    {
                        List<Block> template = new List<Block>();
                        for (int b = i; b <= end; b++)
                        {
                            Block clone = DocumentCloner.CloneBlock(blocks[b]);
                            Paragraph cloneParagraph = clone as Paragraph;
                            bool hadMarker = cloneParagraph != null && RemoveRegionMarkers(cloneParagraph, name);
                            if (hadMarker && !BlockHasContent(clone)) continue;
                            template.Add(clone);
                        }

                        IList items = ResolveRegionItems(context, name);
                        for (int item = 0; item < items.Count; item++)
                        {
                            TemplateContext local = TemplateEngine.Merge(context, CreateRegionRowContext(items[item]));
                            List<Block> clones = DocumentCloner.CloneBlockList(template);
                            count += ProcessBlocks(clones, local, options, depth + 1);
                            output.AddRange(clones);
                        }
                        i = end + 1;
                        continue;
                    }
                }

                if (paragraph != null)
                {
                    if (TryGetParagraphRegionMarker(paragraph, out kind, out name))
                    {
                        RemoveRegionMarkers(paragraph, null);
                    }
                    count += MergeFields(paragraph, context, options);
                    count += MergeTokens(paragraph, context, options);
                }
                else
                {
                    Table table = blocks[i] as Table;
                    if (table != null) count += ProcessTable(table, context, options, depth);
                }
                output.Add(blocks[i]);
                i++;
            }

            blocks.Clear();
            blocks.AddRange(output);
            return count;
        }

        /// <summary>
        /// 递归处理表格：展开 TableStart/TableEnd 行区域，逐个数据项克隆模板行，并继续处理单元格内的嵌套区域。
        /// Recursively processes a table: expands TableStart/TableEnd row regions, clones the template rows per
        /// data item, and continues with the regions nested inside cells.
        /// </summary>
        /// <param name="table">待处理的表格，处理后就地重写行列表。The table, whose row list is rewritten in place.</param>
        /// <param name="context">当前作用域的数据上下文。The data context of the current scope.</param>
        /// <param name="options">合并选项。The merge options.</param>
        /// <param name="depth">当前区域嵌套深度。The current region nesting depth.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        private static int ProcessTable(Table table, TemplateContext context, WordMailMergeOptions options, int depth)
        {
            int count = 0;
            List<TableRow> output = new List<TableRow>(table.Rows.Count);
            int r = 0;
            while (r < table.Rows.Count)
            {
                RegionMarkerKind kind;
                string name;
                if (TryGetRowRegionMarker(table.Rows[r], out kind, out name)
                    && kind == RegionMarkerKind.Start
                    && depth < MaxRegionDepth)
                {
                    int end = FindRegionEndRow(table, r + 1, name);
                    if (end >= 0)
                    {
                        List<TableRow> template = new List<TableRow>();
                        for (int rowIndex = r; rowIndex <= end; rowIndex++)
                        {
                            TableRow clone = DocumentCloner.CloneRow(table.Rows[rowIndex]);
                            bool hadMarker = RemoveRegionMarkersFromRow(clone, name);
                            if (hadMarker && !RowHasContent(clone)) continue;
                            template.Add(clone);
                        }

                        IList items = ResolveRegionItems(context, name);
                        if (items.Count > 0)
                        {
                            for (int t = 0; t < template.Count; t++)
                            {
                                if (!template[t].IsHeader) continue;
                                Table headerStaging = DocumentCloner.CreateTableShell(table);
                                headerStaging.Rows.Add(DocumentCloner.CloneRow(template[t]));
                                count += ProcessTable(headerStaging, context, options, depth + 1);
                                output.AddRange(headerStaging.Rows);
                            }
                            for (int item = 0; item < items.Count; item++)
                            {
                                TemplateContext local = TemplateEngine.Merge(context, CreateRegionRowContext(items[item]));
                                Table staging = DocumentCloner.CreateTableShell(table);
                                for (int t = 0; t < template.Count; t++)
                                {
                                    if (template[t].IsHeader) continue;
                                    staging.Rows.Add(DocumentCloner.CloneRow(template[t]));
                                }
                                count += ProcessTable(staging, local, options, depth + 1);
                                output.AddRange(staging.Rows);
                            }
                        }
                        r = end + 1;
                        continue;
                    }
                }

                count += ProcessRow(table.Rows[r], context, options, depth);
                output.Add(table.Rows[r]);
                r++;
            }

            table.Rows.Clear();
            table.Rows.AddRange(output);
            return count;
        }

        /// <summary>
        /// 处理表格行中的单元格内容，交给块级处理器展开单元格内的区域并合并域。
        /// Processes the cell content of a table row, delegating to the block processor to expand in-cell
        /// regions and merge fields.
        /// </summary>
        /// <param name="row">表格行。The table row.</param>
        /// <param name="context">数据上下文。The data context.</param>
        /// <param name="options">合并选项。The merge options.</param>
        /// <param name="depth">当前区域嵌套深度。The current region nesting depth.</param>
        /// <returns>已写入或清除的域/标记数量。The number of fields and tokens written or cleared.</returns>
        private static int ProcessRow(TableRow row, TemplateContext context, WordMailMergeOptions options, int depth)
        {
            int count = 0;
            for (int c = 0; c < row.Cells.Count; c++)
            {
                count += ProcessBlocks(row.Cells[c].Blocks, context, options, depth);
            }
            return count;
        }

        /// <summary>
        /// 从块列表中查找与指定区域名匹配的结束标记段落，只扫描同一层级的段落。
        /// Finds the end-marker paragraph matching the given region name in the same block list, scanning
        /// only paragraphs at that level.
        /// </summary>
        /// <param name="blocks">块列表。The block list.</param>
        /// <param name="start">开始扫描的下标。The index to start scanning from.</param>
        /// <param name="name">区域名。The region name.</param>
        /// <returns>结束段落下标，未找到为 -1。The end-paragraph index, or -1.</returns>
        private static int FindRegionEndParagraph(List<Block> blocks, int start, string name)
        {
            for (int i = start; i < blocks.Count; i++)
            {
                Paragraph paragraph = blocks[i] as Paragraph;
                if (paragraph == null) continue;
                RegionMarkerKind kind;
                string markerName;
                if (TryGetParagraphRegionMarker(paragraph, out kind, out markerName)
                    && kind == RegionMarkerKind.End
                    && string.Equals(name, markerName, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
            return -1;
        }

        /// <summary>
        /// 从表格中查找与指定区域名匹配的结束标记行，起点行之后的行才参与匹配。
        /// Finds the end-marker row matching the given region name, starting the search after the start row.
        /// </summary>
        /// <param name="table">表格。The table.</param>
        /// <param name="start">开始扫描的行下标。The row index to start scanning from.</param>
        /// <param name="name">区域名。The region name.</param>
        /// <returns>结束行下标，未找到为 -1。The end-row index, or -1.</returns>
        private static int FindRegionEndRow(Table table, int start, string name)
        {
            for (int r = start; r < table.Rows.Count; r++)
            {
                RegionMarkerKind kind;
                string markerName;
                if (TryGetRowRegionMarker(table.Rows[r], out kind, out markerName)
                    && kind == RegionMarkerKind.End
                    && string.Equals(name, markerName, StringComparison.OrdinalIgnoreCase))
                {
                    return r;
                }
            }
            return -1;
        }

        /// <summary>
        /// 读取段落中的第一个区域标记。
        /// Reads the first region marker of a paragraph.
        /// </summary>
        /// <param name="paragraph">段落。The paragraph.</param>
        /// <param name="kind">输出标记种类。Outputs the marker kind.</param>
        /// <param name="name">输出区域名。Outputs the region name.</param>
        /// <returns>存在区域标记返回 true。True when a region marker is present.</returns>
        private static bool TryGetParagraphRegionMarker(Paragraph paragraph, out RegionMarkerKind kind, out string name)
        {
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field != null && TryParseRegionMarker(field.Instruction, out kind, out name)) return true;
            }
            kind = RegionMarkerKind.None;
            name = null;
            return false;
        }

        /// <summary>
        /// 读取表格行中的第一个区域标记，只检查单元格的直接段落，不下钻嵌套表格。
        /// Reads the first region marker of a table row, checking only the direct cell paragraphs and not
        /// descending into nested tables.
        /// </summary>
        /// <param name="row">表格行。The table row.</param>
        /// <param name="kind">输出标记种类。Outputs the marker kind.</param>
        /// <param name="name">输出区域名。Outputs the region name.</param>
        /// <returns>存在区域标记返回 true。True when a region marker is present.</returns>
        private static bool TryGetRowRegionMarker(TableRow row, out RegionMarkerKind kind, out string name)
        {
            for (int c = 0; c < row.Cells.Count; c++)
            {
                List<Block> blocks = row.Cells[c].Blocks;
                for (int b = 0; b < blocks.Count; b++)
                {
                    Paragraph paragraph = blocks[b] as Paragraph;
                    if (paragraph != null && TryGetParagraphRegionMarker(paragraph, out kind, out name)) return true;
                }
            }
            kind = RegionMarkerKind.None;
            name = null;
            return false;
        }

        /// <summary>
        /// 解析区域标记指令：MERGEFIELD TableStart:名称 或 MERGEFIELD TableEnd:名称。
        /// Parses a region marker instruction: MERGEFIELD TableStart:Name or MERGEFIELD TableEnd:Name.
        /// </summary>
        /// <param name="instruction">域指令。The field instruction.</param>
        /// <param name="kind">输出标记种类。Outputs the marker kind.</param>
        /// <param name="name">输出区域名。Outputs the region name.</param>
        /// <returns>是区域标记返回 true。True when the instruction is a region marker.</returns>
        private static bool TryParseRegionMarker(string instruction, out RegionMarkerKind kind, out string name)
        {
            kind = RegionMarkerKind.None;
            name = null;
            string fieldName;
            string format;
            if (!TryParseMergeField(instruction, out fieldName, out format)) return false;

            const string startPrefix = "TableStart:";
            const string endPrefix = "TableEnd:";
            if (fieldName.StartsWith(startPrefix, StringComparison.OrdinalIgnoreCase))
            {
                name = fieldName.Substring(startPrefix.Length).Trim();
                kind = RegionMarkerKind.Start;
            }
            else if (fieldName.StartsWith(endPrefix, StringComparison.OrdinalIgnoreCase))
            {
                name = fieldName.Substring(endPrefix.Length).Trim();
                kind = RegionMarkerKind.End;
            }
            else
            {
                return false;
            }
            if (name.Length == 0)
            {
                kind = RegionMarkerKind.None;
                name = null;
                return false;
            }
            return true;
        }

        /// <summary>
        /// 移除段落中指定区域（name 为 null 时移除全部区域标记）的标记域。
        /// Removes the marker fields of the given region (all region markers when name is null) from a paragraph.
        /// </summary>
        /// <param name="paragraph">段落。The paragraph.</param>
        /// <param name="name">区域名，null 表示全部。The region name; null means all.</param>
        /// <returns>是否移除了标记。Whether any marker was removed.</returns>
        private static bool RemoveRegionMarkers(Paragraph paragraph, string name)
        {
            bool removed = false;
            for (int i = paragraph.Inlines.Count - 1; i >= 0; i--)
            {
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field == null) continue;
                RegionMarkerKind kind;
                string markerName;
                if (!TryParseRegionMarker(field.Instruction, out kind, out markerName)) continue;
                if (name != null && !string.Equals(name, markerName, StringComparison.OrdinalIgnoreCase)) continue;
                paragraph.Inlines.RemoveAt(i);
                removed = true;
            }
            return removed;
        }

        /// <summary>
        /// 移除表格行直接段落中的区域标记。
        /// Removes region markers from the direct paragraphs of a table row.
        /// </summary>
        /// <param name="row">表格行。The table row.</param>
        /// <param name="name">区域名，null 表示全部。The region name; null means all.</param>
        /// <returns>是否移除了标记。Whether any marker was removed.</returns>
        private static bool RemoveRegionMarkersFromRow(TableRow row, string name)
        {
            bool removed = false;
            for (int c = 0; c < row.Cells.Count; c++)
            {
                List<Block> blocks = row.Cells[c].Blocks;
                for (int b = 0; b < blocks.Count; b++)
                {
                    Paragraph paragraph = blocks[b] as Paragraph;
                    if (paragraph != null && RemoveRegionMarkers(paragraph, name)) removed = true;
                }
            }
            return removed;
        }

        /// <summary>
        /// 判断块是否还有可见内容：非空白文本、图片、域或非空表格。
        /// Reports whether a block still has visible content: non-blank text, an image, a field or a non-empty table.
        /// </summary>
        /// <param name="block">块。The block.</param>
        /// <returns>有内容返回 true。True when content remains.</returns>
        private static bool BlockHasContent(Block block)
        {
            Paragraph paragraph = block as Paragraph;
            if (paragraph != null)
            {
                for (int i = 0; i < paragraph.Inlines.Count; i++)
                {
                    TextRun text = paragraph.Inlines[i] as TextRun;
                    if (text != null)
                    {
                        if (!string.IsNullOrWhiteSpace(text.Text)) return true;
                        continue;
                    }
                    if (paragraph.Inlines[i] is ImageRun || paragraph.Inlines[i] is FieldRun) return true;
                }
                return false;
            }
            Table table = block as Table;
            if (table != null) return table.Rows.Count > 0;
            return true;
        }

        /// <summary>
        /// 判断表格行是否还有可见内容。
        /// Reports whether a table row still has visible content.
        /// </summary>
        /// <param name="row">表格行。The table row.</param>
        /// <returns>有内容返回 true。True when content remains.</returns>
        private static bool RowHasContent(TableRow row)
        {
            for (int c = 0; c < row.Cells.Count; c++)
            {
                List<Block> blocks = row.Cells[c].Blocks;
                for (int b = 0; b < blocks.Count; b++)
                {
                    if (BlockHasContent(blocks[b])) return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 解析区域对应的数据项列表：DataTable/DataView 取行，其它可枚举值展开，标量作为单项，缺失或 null 返回空列表。
        /// Resolves the data items of a region: DataTable/DataView yield rows, other enumerables are expanded,
        /// scalars become a single item, and missing or null values yield an empty list.
        /// </summary>
        /// <param name="context">当前作用域上下文。The current scope context.</param>
        /// <param name="name">区域名。The region name.</param>
        /// <returns>数据项列表，永不为 null。The item list, never null.</returns>
        private static IList ResolveRegionItems(TemplateContext context, string name)
        {
            object value;
            if (!context.TryGet(name, out value) || value == null) return new List<object>();

            DataTable table = value as DataTable;
            if (table != null)
            {
                List<object> rows = new List<object>(table.Rows.Count);
                for (int i = 0; i < table.Rows.Count; i++) rows.Add(table.Rows[i]);
                return rows;
            }
            DataView view = value as DataView;
            if (view != null)
            {
                List<object> rows = new List<object>(view.Count);
                for (int i = 0; i < view.Count; i++) rows.Add(view[i]);
                return rows;
            }
            return TemplateEngine.ResolveList(value);
        }

        /// <summary>
        /// 由区域数据项构建行上下文：DataRow 附带队子行以支持 DataRelation 嵌套，模板上下文复制键值，
        /// 其余对象走字典/属性反射；DBNull 转为 null。
        /// Builds the row context of a region item: DataRow carries child rows for DataRelation nesting,
        /// template contexts copy their keys, and other objects use dictionary/property reflection; DBNull
        /// becomes null.
        /// </summary>
        /// <param name="item">数据项。The data item.</param>
        /// <returns>行上下文。The row context.</returns>
        private static TemplateContext CreateRegionRowContext(object item)
        {
            DataRowView view = item as DataRowView;
            if (view != null) return CreateRegionRowContext(view.Row);

            DataRow row = item as DataRow;
            if (row != null)
            {
                TemplateContext context = CreateRowContext(row);
                if (row.Table == null || row.RowState == DataRowState.Deleted) return context;
                DataRelationCollection relations = row.Table.ChildRelations;
                for (int i = 0; i < relations.Count; i++)
                {
                    DataRelation relation = relations[i];
                    DataRow[] children;
                    try
                    {
                        children = row.GetChildRows(relation);
                    }
                    catch (InvalidOperationException)
                    {
                        continue;
                    }
                    if (children == null) continue;
                    if (!string.IsNullOrEmpty(relation.ChildTable.TableName))
                    {
                        context.Set(relation.ChildTable.TableName, children);
                    }
                    if (!string.IsNullOrEmpty(relation.RelationName)
                        && !string.Equals(relation.RelationName, relation.ChildTable.TableName, StringComparison.OrdinalIgnoreCase))
                    {
                        context.Set(relation.RelationName, children);
                    }
                }
                return context;
            }

            TemplateContext template = item as TemplateContext;
            if (template != null)
            {
                TemplateContext copy = new TemplateContext();
                foreach (KeyValuePair<string, object> pair in template.Items)
                {
                    copy.Set(pair.Key, pair.Value);
                }
                return copy;
            }

            return TemplateEngine.CreateRowContext(item);
        }

        /// <summary>
        /// 合并段落中的 MERGEFIELD 域：文本值写入普通文本、图片值写入图文，未解析时按选项保留或清空。
        /// Merges MERGEFIELD fields of one paragraph: text values become plain text, image values become
        /// pictures, and unresolved fields are kept or cleared according to the options.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="context">数据上下文。The data context.</param>
        /// <param name="options">合并选项。The merge options.</param>
        /// <returns>已处理的域数量。The number of fields handled.</returns>
        private static int MergeFields(Paragraph paragraph, TemplateContext context, WordMailMergeOptions options)
        {
            int count = 0;
            for (int i = 0; i < paragraph.Inlines.Count; i++)
            {
                FieldRun field = paragraph.Inlines[i] as FieldRun;
                if (field == null) continue;

                string name;
                string format;
                if (!TryParseMergeField(field.Instruction, out name, out format)) continue;

                object value;
                if (!context.TryGet(name, out value))
                {
                    if (options.LeaveUnresolvedFields) continue;
                    paragraph.Inlines[i] = new TextRun(string.Empty, field.Format.Clone());
                    count++;
                    continue;
                }

                OfficeImage image = ToImage(value);
                if (image != null)
                {
                    paragraph.Inlines[i] = new ImageRun(image);
                    count++;
                    continue;
                }

                paragraph.Inlines[i] = new TextRun(FormatValue(value, format, options), field.Format.Clone());
                count++;
            }
            return count;
        }

        /// <summary>
        /// 合并段落中的 {{token}} / {{token:format}} 标记，跨运行命中并保留首个命中运行的格式。
        /// Merges the {{token}} / {{token:format}} markers of one paragraph, matching across runs and keeping
        /// the formatting of the first matching run.
        /// </summary>
        /// <param name="paragraph">目标段落。The target paragraph.</param>
        /// <param name="context">数据上下文。The data context.</param>
        /// <param name="options">合并选项。The merge options.</param>
        /// <returns>已处理的标记数量。The number of tokens handled.</returns>
        private static int MergeTokens(Paragraph paragraph, TemplateContext context, WordMailMergeOptions options)
        {
            int count = 0;
            List<List<TextRun>> groups = WordTextRange.GetGroups(paragraph);
            for (int g = 0; g < groups.Count; g++)
            {
                List<TextRun> runs = groups[g];
                string text = WordTextRange.GetText(runs);
                if (text.IndexOf("{{", StringComparison.Ordinal) < 0) continue;

                MatchCollection matches = TokenRegex.Matches(text);
                if (matches.Count == 0) continue;

                for (int m = matches.Count - 1; m >= 0; m--)
                {
                    Match match = matches[m];
                    string key = match.Groups[1].Value.Trim();
                    string format = match.Groups[2].Success ? match.Groups[2].Value.Trim() : null;

                    object value;
                    if (!context.TryGet(key, out value))
                    {
                        if (options.LeaveUnresolvedFields) continue;
                        if (WordTextRange.ReplaceRange(runs, match.Index, match.Length, string.Empty)) count++;
                        continue;
                    }
                    if (ToImage(value) != null)
                    {
                        // 图片不能嵌入文本中间，图片请使用 MERGEFIELD 或 ImageRun。
                        // Images cannot be embedded inside a text span; use MERGEFIELD or ImageRun instead.
                        continue;
                    }
                    if (WordTextRange.ReplaceRange(runs, match.Index, match.Length, FormatValue(value, format, options)))
                    {
                        count++;
                    }
                }
            }
            return count;
        }

        /// <summary>
        /// 解析 MERGEFIELD 指令：取出域名（可带引号，可带 Image: 前缀）与 \#/\@ 格式开关。
        /// Parses a MERGEFIELD instruction: extracts the field name (optionally quoted, with an optional
        /// Image: prefix) and the \#/\@ format switch.
        /// </summary>
        /// <param name="instruction">域指令。The field instruction.</param>
        /// <param name="name">输出域名。Outputs the field name.</param>
        /// <param name="format">输出格式串，没有时为 null。Outputs the format string, or null when absent.</param>
        /// <returns>是 MERGEFIELD 指令返回 true。True when the instruction is a MERGEFIELD.</returns>
        private static bool TryParseMergeField(string instruction, out string name, out string format)
        {
            name = null;
            format = null;
            string text = (instruction ?? string.Empty).Trim();
            if (!text.StartsWith("MERGEFIELD", StringComparison.OrdinalIgnoreCase)) return false;
            if (text.Length > 10 && !char.IsWhiteSpace(text[10]) && text[10] != '"') return false;

            string rest = text.Length > 10 ? text.Substring(10).Trim() : string.Empty;
            if (rest.Length == 0) return false;

            if (rest[0] == '"')
            {
                int close = rest.IndexOf('"', 1);
                if (close < 1) return false;
                name = rest.Substring(1, close - 1).Trim();
                rest = rest.Substring(close + 1);
            }
            else
            {
                int end = 0;
                while (end < rest.Length && !char.IsWhiteSpace(rest[end]) && rest[end] != '\\') end++;
                name = rest.Substring(0, end).Trim();
                rest = rest.Substring(end);
            }
            if (name.Length == 0) return false;
            if (name.StartsWith("Image:", StringComparison.OrdinalIgnoreCase))
            {
                name = name.Substring("Image:".Length).Trim();
                if (name.Length == 0) return false;
            }

            Match formatMatch = FieldFormatRegex.Match(rest);
            if (formatMatch.Success)
            {
                format = formatMatch.Groups[2].Success ? formatMatch.Groups[2].Value : formatMatch.Groups[3].Value;
            }
            return true;
        }

        /// <summary>
        /// 把取值格式化为文本：字符串原样返回，IFormattable 按格式串与文化格式化，其余走 Convert。
        /// Formats a value as text: strings are returned as-is, IFormattable values use the format string and
        /// culture, and everything else goes through Convert.
        /// </summary>
        /// <param name="value">取值。The value.</param>
        /// <param name="format">格式串，可为 null。The format string; may be null.</param>
        /// <param name="options">合并选项（提供文化）。The merge options that supply the culture.</param>
        /// <returns>格式化后的文本。The formatted text.</returns>
        private static string FormatValue(object value, string format, WordMailMergeOptions options)
        {
            if (value == null) return string.Empty;
            string text = value as string;
            if (text != null) return text;
            IFormattable formattable = value as IFormattable;
            if (formattable != null) return formattable.ToString(format, options.Culture ?? CultureInfo.InvariantCulture);
            return Convert.ToString(value, options.Culture ?? CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 把取值转换为图片：OfficeImage 直接使用，字节数组按图片嗅探后构造，其它类型返回 null。
        /// Converts a value into an image: OfficeImage instances are used directly, byte arrays are sniffed and
        /// wrapped, and any other type returns null.
        /// </summary>
        /// <param name="value">取值。The value.</param>
        /// <returns>图片或 null。The image, or null.</returns>
        private static OfficeImage ToImage(object value)
        {
            OfficeImage image = value as OfficeImage;
            if (image != null) return image;
            byte[] data = value as byte[];
            if (data != null && data.Length > 0) return new OfficeImage(data);
            return null;
        }

        /// <summary>
        /// 由数据行构建上下文：列名作为键，DBNull 转为 null；行为 null 时返回空上下文。
        /// Builds a context from a data row: column names become keys and DBNull becomes null; a null row yields
        /// an empty context.
        /// </summary>
        /// <param name="row">数据行，可为 null。The data row, which may be null.</param>
        /// <returns>构建好的数据上下文。The constructed data context.</returns>
        private static TemplateContext CreateRowContext(DataRow row)
        {
            TemplateContext context = new TemplateContext();
            if (row == null || row.Table == null) return context;
            DataColumnCollection columns = row.Table.Columns;
            for (int i = 0; i < columns.Count; i++)
            {
                DataColumn column = columns[i];
                object value = row[column];
                context.Set(column.ColumnName, value == DBNull.Value ? null : value);
            }
            return context;
        }
    }
}
