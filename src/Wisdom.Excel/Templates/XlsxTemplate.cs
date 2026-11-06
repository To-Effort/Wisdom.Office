using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 用 <c>{{token}}</c> 占位值和 <c>{{#each}}</c> 行循环填充 SpreadsheetML 模板。
    /// 循环可跨单行或多行；循环区域内的合并区域会随每次迭代重复。
    /// Fills SpreadsheetML templates using <c>{{token}}</c> placeholder values and <c>{{#each}}</c> row loops.
    /// Loops may span one or more rows; merged ranges inside a loop area repeat with each iteration.
    /// </summary>
    public static class XlsxTemplate
    {
        /// <summary>匹配整格仅为单个令牌（可带 :格式 后缀）的正则，捕获键名与格式两段。</summary>
        private static readonly Regex WholeTokenRegex = new Regex(
            @"^\{\{\s*([^#/{}][^{}:]*?)\s*(?::\s*([^{}]+?))?\s*\}\}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>读取模板文件、填充上下文并以默认选项写出结果文件（使用默认区域设置）。</summary>
        /// <param name="templatePath">模板 .xlsx 文件路径。</param>
        /// <param name="outputPath">输出 .xlsx 文件路径。</param>
        /// <param name="context">令牌上下文。</param>
        public static void Fill(string templatePath, string outputPath, TemplateContext context)
        {
            Fill(templatePath, outputPath, context, new TemplateOptions());
        }

        /// <summary>读取模板文件、按指定选项填充上下文并写出结果文件。</summary>
        /// <param name="templatePath">模板 .xlsx 文件路径。</param>
        /// <param name="outputPath">输出 .xlsx 文件路径。</param>
        /// <param name="context">令牌上下文。</param>
        /// <param name="options">模板填充选项。</param>
        public static void Fill(string templatePath, string outputPath, TemplateContext context, TemplateOptions options)
        {
            if (templatePath == null) throw new ArgumentNullException("templatePath");
            if (outputPath == null) throw new ArgumentNullException("outputPath");
            OfficeWorkbook workbook = XlsxReader.Load(File.ReadAllBytes(templatePath));
            OfficeWorkbook filled = Fill(workbook, context, options);
            XlsxWriter.Write(filled, outputPath);
        }

        /// <summary>在内存中填充已加载的模板工作簿（使用默认选项），返回新工作簿。</summary>
        /// <param name="template">模板工作簿模型。</param>
        /// <param name="context">令牌上下文。</param>
        /// <returns>填充后的新工作簿。</returns>
        public static OfficeWorkbook Fill(OfficeWorkbook template, TemplateContext context)
        {
            return Fill(template, context, new TemplateOptions());
        }

        /// <summary>
        /// 在内存中填充模板工作簿：逐表调用 ProcessSheet，不修改原始模板。
        /// Fills the template workbook in memory: calls ProcessSheet for each sheet without modifying the
        /// original template.
        /// </summary>
        /// <param name="template">模板工作簿模型。</param>
        /// <param name="context">令牌上下文。</param>
        /// <param name="options">模板填充选项。</param>
        /// <returns>填充后的新工作簿。</returns>
        public static OfficeWorkbook Fill(OfficeWorkbook template, TemplateContext context, TemplateOptions options)
        {
            if (template == null) throw new ArgumentNullException("template");
            if (options == null) options = new TemplateOptions();

            OfficeWorkbook result = new OfficeWorkbook();
            result.Metadata = CloneMetadata(template.Metadata);
            result.DefaultStyle = template.DefaultStyle != null ? template.DefaultStyle.Clone() : new CellStyle();
            // 日期基准属于工作簿，不是某张表：漏掉它，1904 模板写出的每个日期都会整体偏移四年。
            // The date epoch belongs to the workbook, not one sheet: drop it and every date in a 1904 template
            // lands four years early.
            result.Date1904 = template.Date1904;
            // 定义名逐表按同一顺序生成，工作表下标不变，引用原样带走（填充不重写引用里的行号，与公式同一口径）。
            // Sheets keep their order, so local-sheet indexes stay valid; references are carried over as written,
            // the same no-rewrite rule template fill already follows for formulas.
            for (int i = 0; i < template.DefinedNames.Count; i++)
            {
                ExcelDefinedName name = template.DefinedNames[i];
                if (name == null || string.IsNullOrEmpty(name.Name)) continue;
                ExcelDefinedName copy = new ExcelDefinedName(name.Name, name.Reference);
                copy.LocalSheetId = name.LocalSheetId;
                result.DefinedNames.Add(copy);
            }
            for (int i = 0; i < template.Worksheets.Count; i++)
            {
                Worksheet filled = ProcessSheet(template.Worksheets[i], context, options);
                result.Worksheets.Add(filled);
            }
            return result;
        }

        /// <summary>
        /// 处理单个工作表：复制表级属性与列宽，检测 {{#each}} 循环块，
        /// 按行（循环行按数据条目重复）逐行拷贝到输出表，最后按行映射重建合并区域，
        /// 并把条件格式、筛选、数据验证、打印区域这些"行号写在文本里"的装饰换算到输出行。
        /// Processes one worksheet: copies sheet-level properties and column widths, detects {{#each}} loop
        /// blocks, copies rows one by one to the output sheet (loop rows repeated per data item), then rebuilds
        /// merged ranges and converts the decorations whose row numbers live inside range text — conditional
        /// formats, the filter, data validations and the print area — through the same row mapping.
        /// </summary>
        /// <param name="source">模板工作表。</param>
        /// <param name="context">外层令牌上下文。</param>
        /// <param name="options">填充选项。</param>
        /// <returns>填充后的新工作表。</returns>
        private static Worksheet ProcessSheet(Worksheet source, TemplateContext context, TemplateOptions options)
        {
            Worksheet target = new Worksheet(source.Name);
            target.ShowGridLines = source.ShowGridLines;
            target.FrozenRows = source.FrozenRows;
            target.FrozenColumns = source.FrozenColumns;
            target.FitToPage = source.FitToPage;
            target.DefaultColumnWidth = source.DefaultColumnWidth;
            target.DefaultRowHeight = source.DefaultRowHeight;
            target.DefaultStyle = source.DefaultStyle != null ? source.DefaultStyle.Clone() : new CellStyle();
            CopyPageSetup(source.PrintSetup, target.PrintSetup);
            foreach (KeyValuePair<int, double> pair in source.ColumnWidths)
            {
                target.SetColumnWidth(pair.Key, pair.Value);
            }

            Dictionary<int, List<Cell>> rows = new Dictionary<int, List<Cell>>();
            int lastRow = source.LastRow;
            foreach (Cell cell in source.Cells)
            {
                List<Cell> rowCells;
                if (!rows.TryGetValue(cell.Row, out rowCells))
                {
                    rowCells = new List<Cell>();
                    rows.Add(cell.Row, rowCells);
                }
                rowCells.Add(cell);
            }
            foreach (int row in source.RowHeights.Keys)
            {
                if (row > lastRow) lastRow = row;
            }

            List<LoopBlock> loops = DetectLoops(rows, lastRow);
            Dictionary<int, LoopBlock> loopsByRow = new Dictionary<int, LoopBlock>();
            for (int i = 0; i < loops.Count; i++)
            {
                loopsByRow[loops[i].StartRow] = loops[i];
            }

            List<Placement> placements = new List<Placement>();
            int outputRow = 1;
            for (int row = 1; row <= lastRow; row++)
            {
                LoopBlock loop;
                if (loopsByRow.TryGetValue(row, out loop))
                {
                    IList items = TemplateEngine.ResolveList(context == null ? null : context.Get(loop.Name));
                    int count = items.Count;
                    bool skipStart = loop.StartRow != loop.EndRow && IsMarkerOnlyRow(rows, loop.StartRow);
                    bool skipEnd = loop.StartRow != loop.EndRow && IsMarkerOnlyRow(rows, loop.EndRow);
                    for (int item = 0; item < count; item++)
                    {
                        TemplateContext rowContext = CreateItemContext(context, items[item], item, count);
                        for (int sourceRow = loop.StartRow; sourceRow <= loop.EndRow; sourceRow++)
                        {
                            if (sourceRow == loop.StartRow && skipStart) continue;
                            if (sourceRow == loop.EndRow && skipEnd) continue;
                            CopyRow(rows, source, target, sourceRow, outputRow, rowContext, options);
                            placements.Add(new Placement(sourceRow, outputRow, loop, item));
                            outputRow++;
                        }
                    }
                    row = loop.EndRow;
                }
                else
                {
                    CopyRow(rows, source, target, row, outputRow, context, options);
                    placements.Add(new Placement(row, outputRow, null, -1));
                    outputRow++;
                }
            }

            CopyMergedRanges(source, target, placements);
            CopyDecorations(source, target, new RowMap(lastRow, placements));
            return target;
        }

        /// <summary>
        /// 将模板行拷贝到输出行：复制行高、样式、公式（含令牌替换），
        /// 文本值经过去除循环标记与令牌解析，整格令牌可携带 :格式 后缀应用数字格式。
        /// Copies a template row to an output row: reproduces row height, styles and formulas (with token
        /// substitution), passes text values through loop-marker stripping and token resolution, and lets a
        /// whole-cell token carry a :format suffix to apply a number format.
        /// </summary>
        /// <param name="rows">按行号分组的模板单元格。</param>
        /// <param name="source">模板工作表（读取行高）。</param>
        /// <param name="target">输出工作表。</param>
        /// <param name="sourceRow">模板行号。</param>
        /// <param name="targetRow">输出行号。</param>
        /// <param name="context">当前行使用的上下文。</param>
        /// <param name="options">填充选项。</param>
        private static void CopyRow(
            Dictionary<int, List<Cell>> rows,
            Worksheet source,
            Worksheet target,
            int sourceRow,
            int targetRow,
            TemplateContext context,
            TemplateOptions options)
        {
            double height;
            if (source.RowHeights.TryGetValue(sourceRow, out height))
            {
                target.SetRowHeight(targetRow, height);
            }

            List<Cell> cells;
            if (!rows.TryGetValue(sourceRow, out cells)) return;

            for (int i = 0; i < cells.Count; i++)
            {
                Cell cell = cells[i];
                Cell copy = target.GetCell(targetRow, cell.Column);
                if (cell.Style != null) copy.Style = cell.Style.Clone();
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    copy.Formula = TemplateEngine.ContainsToken(cell.Formula)
                        ? TemplateEngine.Replace(cell.Formula, context, options)
                        : cell.Formula;
                }

                string text = cell.Value as string;
                if (text == null)
                {
                    copy.Value = cell.Value;
                    continue;
                }

                string cleaned = TemplateEngine.RemoveLoopMarkers(text);
                string format;
                copy.Value = ResolveValue(cleaned, context, options, out format);
                if (format != null && copy.Value != null && !(copy.Value is string))
                {
                    ApplyNumberFormat(copy, format);
                }
            }
        }

        /// <summary>
        /// 解析单元格文本：整格单令牌时直接返回原始类型值（数值/日期/布尔保持类型），
        /// 并输出 :格式 后缀；否则走普通令牌替换返回字符串。
        /// Resolves cell text: a single whole-cell token returns the raw typed value directly (numbers/dates/
        /// booleans keep their type) and emits the :format suffix; otherwise normal token substitution returns
        /// a string.
        /// </summary>
        /// <param name="text">去除循环标记后的单元格文本。</param>
        /// <param name="context">令牌上下文。</param>
        /// <param name="options">填充选项。</param>
        /// <param name="format">令牌携带的数字格式后缀；无则为 null。</param>
        /// <returns>解析后的单元格值。</returns>
        private static object ResolveValue(string text, TemplateContext context, TemplateOptions options, out string format)
        {
            format = null;
            if (string.IsNullOrEmpty(text)) return null;
            if (context == null) return text;

            Match match = WholeTokenRegex.Match(text);
            if (match.Success)
            {
                string key = match.Groups[1].Value.Trim();
                object value;
                if (!context.TryGet(key, out value))
                {
                    return options.LeaveUnresolvedTokens ? text : null;
                }
                format = match.Groups[2].Success ? match.Groups[2].Value.Trim() : null;
                if (value == null) return null;
                if (value is string) return value;
                if (value is bool) return value;
                if (value is DateTime || IsNumeric(value)) return value;
                return Convert.ToString(value, options.Culture);
            }

            string replaced = TemplateEngine.Replace(text, context, options);
            return replaced.Length == 0 ? null : replaced;
        }

        /// <summary>为单元格设置数字格式（样式为空时先创建）。</summary>
        /// <param name="cell">目标单元格。</param>
        /// <param name="format">数字格式代码。</param>
        private static void ApplyNumberFormat(Cell cell, string format)
        {
            if (cell.Style == null) cell.Style = new CellStyle();
            cell.Style.NumberFormat = format;
        }

        /// <summary>判断值是否为整数或浮点等数值类型。</summary>
        /// <param name="value">待判断的值。</param>
        /// <returns>数值类型返回 true。</returns>
        private static bool IsNumeric(object value)
        {
            switch (Type.GetTypeCode(value.GetType()))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return true;
                default:
                    return false;
            }
        }

        /// <summary>
        /// 检测行循环块：找到 {{#each 名称}} 起始行后，若同行含结束标记则为单行循环，
        /// 否则向后搜索 {{/each}} 所在行作为结束行。
        /// Detects row loop blocks: after finding the {{#each name}} start row, the loop is single-row when the
        /// same line also holds the end marker, otherwise later rows are searched for the {{/each}} row as the
        /// end row.
        /// </summary>
        /// <param name="rows">按行号分组的模板单元格。</param>
        /// <param name="lastRow">最大行号。</param>
        /// <returns>检测出的循环块列表。</returns>
        private static List<LoopBlock> DetectLoops(Dictionary<int, List<Cell>> rows, int lastRow)
        {
            List<LoopBlock> loops = new List<LoopBlock>();
            for (int row = 1; row <= lastRow; row++)
            {
                string text = GetRowText(rows, row);
                string name;
                if (!TemplateEngine.TryGetLoopName(text, out name)) continue;

                if (TemplateEngine.ContainsLoopEnd(text))
                {
                    loops.Add(new LoopBlock(row, row, name));
                    continue;
                }

                for (int end = row + 1; end <= lastRow; end++)
                {
                    if (!TemplateEngine.ContainsLoopEnd(GetRowText(rows, end))) continue;
                    loops.Add(new LoopBlock(row, end, name));
                    row = end;
                    break;
                }
            }
            return loops;
        }

        /// <summary>判断整行是否只包含循环标记文本（除标记外无其他内容），此类行在展开时跳过。</summary>
        /// <param name="rows">按行号分组的模板单元格。</param>
        /// <param name="row">行号。</param>
        /// <returns>仅含标记时返回 true。</returns>
        private static bool IsMarkerOnlyRow(Dictionary<int, List<Cell>> rows, int row)
        {
            List<Cell> cells;
            if (!rows.TryGetValue(row, out cells)) return false;
            bool hasText = false;
            for (int i = 0; i < cells.Count; i++)
            {
                string text = cells[i].Value as string;
                if (string.IsNullOrEmpty(text)) continue;
                hasText = true;
                if (TemplateEngine.RemoveLoopMarkers(text).Trim().Length > 0) return false;
            }
            return hasText;
        }

        /// <summary>拼接整行各单元格的文本（以换行符分隔），用于循环标记检测。</summary>
        /// <param name="rows">按行号分组的模板单元格。</param>
        /// <param name="row">行号。</param>
        /// <returns>行的合并文本；行不存在时返回空字符串。</returns>
        private static string GetRowText(Dictionary<int, List<Cell>> rows, int row)
        {
            List<Cell> cells;
            if (!rows.TryGetValue(row, out cells)) return string.Empty;
            System.Text.StringBuilder builder = new System.Text.StringBuilder();
            for (int i = 0; i < cells.Count; i++)
            {
                string text = cells[i].Value as string;
                if (text == null) continue;
                if (builder.Length > 0) builder.Append('\n');
                builder.Append(text);
            }
            return builder.ToString();
        }

        /// <summary>为循环条目创建行上下文：注入 @index/@number/@first/@last 变量并与父上下文合并。</summary>
        /// <param name="parent">外层上下文。</param>
        /// <param name="item">当前循环条目。</param>
        /// <param name="index">从 0 开始的条目序号。</param>
        /// <param name="count">条目总数。</param>
        /// <returns>合并后的行级上下文。</returns>
        private static TemplateContext CreateItemContext(TemplateContext parent, object item, int index, int count)
        {
            TemplateContext local = TemplateEngine.CreateRowContext(item);
            local.Set("@index", index);
            local.Set("@number", index + 1);
            local.Set("@first", index == 0);
            local.Set("@last", index == count - 1);
            return TemplateEngine.Merge(parent, local);
        }

        /// <summary>
        /// 重建合并区域：按源行到输出行的映射分组（基础行为一组，循环的每次迭代各为一组），
        /// 只有当某组行数与源合并区域行数相同且输出行连续时才生成新的合并区域。
        /// Rebuilds merged ranges: groups the source-row-to-output-row mapping (the base row forms one group
        /// and each loop iteration forms another); a new merged range is produced only when a group has the
        /// same number of rows as the source range and the output rows are consecutive.
        /// </summary>
        /// <param name="source">模板工作表。</param>
        /// <param name="target">输出工作表。</param>
        /// <param name="placements">源行与输出行的对应关系记录。</param>
        private static void CopyMergedRanges(Worksheet source, Worksheet target, List<Placement> placements)
        {
            for (int i = 0; i < source.MergedRanges.Count; i++)
            {
                MergedRange range = source.MergedRanges[i];
                int rowCount = range.LastRow - range.Row + 1;
                Dictionary<string, List<Placement>> groups = new Dictionary<string, List<Placement>>(StringComparer.Ordinal);
                for (int p = 0; p < placements.Count; p++)
                {
                    Placement placement = placements[p];
                    if (placement.SourceRow < range.Row || placement.SourceRow > range.LastRow) continue;
                    string key = placement.Loop == null
                        ? "base"
                        : RuntimeHelpers.GetHashCode(placement.Loop).ToString(CultureInfo.InvariantCulture)
                            + ":" + placement.ItemIndex.ToString(CultureInfo.InvariantCulture);
                    List<Placement> group;
                    if (!groups.TryGetValue(key, out group))
                    {
                        group = new List<Placement>();
                        groups.Add(key, group);
                    }
                    group.Add(placement);
                }

                foreach (KeyValuePair<string, List<Placement>> pair in groups)
                {
                    List<Placement> group = pair.Value;
                    if (group.Count != rowCount) continue;
                    group.Sort(delegate (Placement left, Placement right) { return left.OutputRow.CompareTo(right.OutputRow); });
                    bool contiguous = true;
                    for (int g = 1; g < group.Count; g++)
                    {
                        if (group[g].OutputRow != group[g - 1].OutputRow + 1)
                        {
                            contiguous = false;
                            break;
                        }
                    }
                    if (!contiguous) continue;
                    target.MergedRanges.Add(new MergedRange(group[0].OutputRow, range.Column, group.Count, range.ColumnSpan));
                }
            }
        }

        /// <summary>
        /// 复制模板里那些"不在单元格上"的设置：视图开关、隐藏行列、打印区域/重复标题、筛选、数据验证、
        /// 条件格式与图表/图片；行号按模板行到输出行的映射换算。
        /// Copies the settings that do not live on cells — view flags, hidden rows and columns, print area and
        /// repeated titles, the filter, data validations, conditional formats, charts and images — converting
        /// row numbers through the template-row to output-row mapping.
        /// </summary>
        /// <param name="source">模板工作表。</param>
        /// <param name="target">输出工作表。</param>
        /// <param name="rows">模板行到输出行的映射。</param>
        private static void CopyDecorations(Worksheet source, Worksheet target, RowMap rows)
        {
            target.IsHidden = source.IsHidden;
            target.ShowZeros = source.ShowZeros;
            target.CenterHorizontally = source.CenterHorizontally;
            target.TabColor = source.TabColor;
            foreach (int column in source.HiddenColumns) target.HiddenColumns.Add(column);
            foreach (int row in source.HiddenRows)
            {
                int mapped = rows.MappedRow(row);
                if (mapped > 0) target.HiddenRows.Add(mapped);
            }

            if (source.HasPrintArea)
            {
                int first = rows.RangeStart(source.PrintAreaFirstRow);
                int last = rows.RangeEnd(source.PrintAreaLastRow);
                if (first > 0 && last >= first)
                {
                    target.SetPrintArea(CellAddress.Format(first, source.PrintAreaFirstColumn) + ":"
                        + CellAddress.Format(last, source.PrintAreaLastColumn));
                }
            }
            if (source.PrintTitleFirstRow > 0)
            {
                int first = rows.RangeStart(source.PrintTitleFirstRow);
                int last = rows.RangeEnd(source.PrintTitleLastRow);
                if (first > 0 && last >= first)
                {
                    target.SetPrintTitleRows(first + ":" + last);
                }
            }
            if (source.PrintTitleFirstColumn > 0)
            {
                // 列方向不受行展开影响，原样带走。 Columns are untouched by row expansion, so copy as written.
                target.SetPrintTitleColumns(CellAddress.ColumnName(source.PrintTitleFirstColumn) + ":"
                    + CellAddress.ColumnName(source.PrintTitleLastColumn));
            }

            if (source.AutoFilter != null)
            {
                string range = MapRangeText(source.AutoFilter.Range, rows);
                if (range.Length > 0)
                {
                    ExcelAutoFilter filter = new ExcelAutoFilter(range);
                    for (int i = 0; i < source.AutoFilter.Columns.Count; i++)
                    {
                        ExcelFilterColumn column = source.AutoFilter.Columns[i];
                        ExcelFilterColumn copy = filter.ForColumn(column.ColumnId);
                        copy.IncludeBlanks = column.IncludeBlanks;
                        copy.Values.AddRange(column.Values);
                    }
                    target.AutoFilter = filter;
                }
            }

            for (int i = 0; i < source.DataValidations.Count; i++)
            {
                ExcelDataValidation validation = source.DataValidations[i];
                if (validation == null) continue;
                string range = MapRangeText(validation.Range, rows);
                if (range.Length == 0) continue;
                target.DataValidations.Add(new ExcelDataValidation
                {
                    Range = range,
                    Type = validation.Type,
                    Operator = validation.Operator,
                    Formula1 = validation.Formula1,
                    Formula2 = validation.Formula2,
                    AllowBlank = validation.AllowBlank,
                    ShowDropDown = validation.ShowDropDown,
                    ShowInputMessage = validation.ShowInputMessage,
                    ShowErrorMessage = validation.ShowErrorMessage,
                    PromptTitle = validation.PromptTitle,
                    Prompt = validation.Prompt,
                    ErrorTitle = validation.ErrorTitle,
                    ErrorMessage = validation.ErrorMessage,
                    ErrorStyle = validation.ErrorStyle
                });
            }

            for (int i = 0; i < source.ConditionalFormats.Count; i++)
            {
                ExcelConditionalFormat format = source.ConditionalFormats[i];
                if (format == null || format.Rules.Count == 0) continue;
                string range = MapRangeText(format.Range, rows);
                if (range.Length == 0) continue;
                ExcelConditionalFormat copy = new ExcelConditionalFormat(range);
                for (int r = 0; r < format.Rules.Count; r++) copy.Rules.Add(format.Rules[r]);
                target.ConditionalFormats.Add(copy);
            }

            for (int i = 0; i < source.Charts.Count; i++)
            {
                ExcelChart chart = source.Charts[i];
                if (chart == null) continue;
                if (rows.IsIdentity)
                {
                    target.Charts.Add(chart);
                    continue;
                }
                ExcelChart copy = CopyChart(chart, rows);
                if (chart.Series.Count > 0 && copy.Series.Count == 0) continue;
                target.Charts.Add(copy);
            }
            for (int i = 0; i < source.Images.Count; i++)
            {
                ExcelImage image = source.Images[i];
                if (image == null) continue;
                if (!rows.IsIdentity)
                {
                    int anchorRow = rows.MappedRow(image.Row);
                    if (anchorRow <= 0) continue;
                    ExcelImage copy = new ExcelImage(image.Image);
                    copy.Row = anchorRow;
                    copy.Column = image.Column;
                    copy.OffsetXPoints = image.OffsetXPoints;
                    copy.OffsetYPoints = image.OffsetYPoints;
                    copy.WidthPoints = image.WidthPoints;
                    copy.HeightPoints = image.HeightPoints;
                    target.Images.Add(copy);
                    continue;
                }
                target.Images.Add(image);
            }
        }

        /// <summary>
        /// 行号发生位移时重建图表：锚点与四条系列区域都换成输出行。字面量列表按引用共享，
        /// 与 <c>CopyWorksheet</c> 同一口径（填充结果不再回写模板）。
        /// Rebuilds a chart once rows moved: the anchor and the four series ranges become output rows. Literal
        /// caches are shared by reference, the same way <c>CopyWorksheet</c> shares them (a filled workbook is
        /// never written back into the template).
        /// </summary>
        /// <param name="source">模板图表。</param>
        /// <param name="rows">模板行到输出行的映射。</param>
        /// <returns>输出图表。</returns>
        private static ExcelChart CopyChart(ExcelChart source, RowMap rows)
        {
            ExcelChart copy = new ExcelChart();
            copy.Type = source.Type;
            copy.Title = source.Title;
            copy.Column = source.Column;
            copy.WidthPoints = source.WidthPoints;
            copy.HeightPoints = source.HeightPoints;
            copy.ShowLegend = source.ShowLegend;
            copy.LegendPosition = source.LegendPosition;
            copy.HoleSize = source.HoleSize;
            copy.ValueAxisFormat = source.ValueAxisFormat;
            int anchorRow = rows.MappedRow(source.Row);
            copy.Row = anchorRow > 0 ? anchorRow : rows.RangeStart(source.Row);
            if (copy.Row <= 0) copy.Row = 1;
            for (int i = 0; i < source.Series.Count; i++)
            {
                ExcelChartSeries series = source.Series[i];
                if (series == null) continue;
                string valueRange = MapSeriesRange(series.ValueRange, rows);
                if (string.IsNullOrEmpty(valueRange) && !string.IsNullOrEmpty(series.ValueRange)
                    && (series.Values == null || series.Values.Count == 0))
                {
                    // 数据行随空循环一起消失，这条系列已无值可画；空系列会让图表部件不合法。
                    // Its data rows vanished with an empty loop, so the series has nothing left to plot, and an
                    // empty series would make the chart part illegal.
                    continue;
                }
                // 构造函数是字面量类目/数值唯一的写入口，区域与样式属性随后补齐。
                // The constructor is the only way to hand over the literal categories and values, so the ranges
                // and style properties are filled in afterwards.
                ExcelChartSeries copySeries = new ExcelChartSeries(series.Name, series.Categories, series.Values);
                copySeries.NameRange = MapSeriesRange(series.NameRange, rows);
                copySeries.CategoryRange = MapSeriesRange(series.CategoryRange, rows);
                copySeries.ValueRange = valueRange;
                copySeries.SizeRange = MapSeriesRange(series.SizeRange, rows);
                copySeries.ChartType = series.ChartType;
                copySeries.SecondaryAxis = series.SecondaryAxis;
                copySeries.FillColor = series.FillColor;
                copySeries.ShowDataLabels = series.ShowDataLabels;
                copySeries.DataLabelFormat = series.DataLabelFormat;
                copySeries.Marker = series.Marker;
                copySeries.Smooth = series.Smooth;
                for (int c = 0; c < series.PointColors.Count; c++) copySeries.PointColors.Add(series.PointColors[c]);
                copy.Series.Add(copySeries);
            }
            return copy;
        }

        /// <summary>
        /// 把模板区域文本换算成输出区域：逐片段剥表前缀、按行映射换行号；认不出的整列引用（<c>A:A</c>）
        /// 不含行号，原样保留，其余无法换算的片段丢弃。
        /// Converts template range text into output ranges: each fragment loses its sheet prefix and has its rows
        /// remapped. Whole-column references such as <c>A:A</c> carry no row numbers and stay verbatim; other
        /// fragments that cannot be converted are dropped.
        /// </summary>
        /// <param name="text">模板区域文本。</param>
        /// <param name="rows">模板行到输出行的映射。</param>
        /// <returns>换算后的区域文本；无有效片段时为空串。</returns>
        private static string MapRangeText(string text, RowMap rows)
        {
            System.Text.StringBuilder builder = new System.Text.StringBuilder();
            if (text == null) return string.Empty;
            string[] fragments = text.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < fragments.Length; i++)
            {
                string fragment = fragments[i];
                int bang = fragment.LastIndexOf('!');
                if (bang >= 0) fragment = fragment.Substring(bang + 1);
                int row1, column1, row2, column2;
                try
                {
                    CellAddress.ParseRange(fragment, out row1, out column1, out row2, out column2);
                }
                catch (FormatException)
                {
                    if (fragment.Length == 0) continue;
                    // 整行引用（2:2）只有行号，换算后仍写成整行；整列引用（A:A）不含行号，原样带走。
                    // A whole-row reference carries row numbers only and stays whole-row after mapping; a
                    // whole-column reference has no rows at all and stays verbatim.
                    int rowFirst, rowLast;
                    if (TryReadRowPair(fragment, out rowFirst, out rowLast))
                    {
                        int mappedFirst = rows.RangeStart(rowFirst);
                        int mappedLast = rows.RangeEnd(rowLast);
                        if (mappedFirst > 0 && mappedLast >= mappedFirst)
                        {
                            AppendFragment(builder, mappedFirst.ToString(CultureInfo.InvariantCulture)
                                + ":" + mappedLast.ToString(CultureInfo.InvariantCulture));
                        }
                        continue;
                    }
                    if (HasDigit(fragment)) continue;
                    AppendFragment(builder, fragment);
                    continue;
                }
                int first = rows.RangeStart(Math.Min(row1, row2));
                int last = rows.RangeEnd(Math.Max(row1, row2));
                if (first <= 0 || last < first) continue;
                int firstColumn = Math.Min(column1, column2);
                int lastColumn = Math.Max(column1, column2);
                AppendFragment(builder, last == first && lastColumn == firstColumn
                    ? CellAddress.Format(first, firstColumn)
                    : CellAddress.Format(first, firstColumn) + ":" + CellAddress.Format(last, lastColumn));
            }
            return builder.ToString();
        }

        /// <summary>换算系列区域，模板端为 null 时保持 null，不写出空引用。</summary>
        /// <param name="text">模板区域文本。</param>
        /// <param name="rows">模板行到输出行的映射。</param>
        /// <returns>换算后的区域文本，或 null。</returns>
        private static string MapSeriesRange(string text, RowMap rows)
        {
            if (text == null) return null;
            if (text.Length == 0) return string.Empty;
            return MapRangeText(text, rows);
        }

        /// <summary>识别整行引用（<c>2:2</c>，可带 <c>$</c>）并读出起止行号。</summary>
        /// <param name="fragment">区域片段。</param>
        /// <param name="first">输出起始行号。</param>
        /// <param name="last">输出结束行号。</param>
        /// <returns>是整行引用时返回 true。</returns>
        private static bool TryReadRowPair(string fragment, out int first, out int last)
        {
            first = 0;
            last = 0;
            string[] parts = fragment.Split(':');
            if (parts.Length != 2) return false;
            for (int i = 0; i < parts.Length; i++)
            {
                string text = parts[i].Trim();
                if (text.StartsWith("$", StringComparison.Ordinal)) text = text.Substring(1);
                if (text.Length == 0 || !IsAllDigits(text)) return false;
                int value;
                if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value)) return false;
                if (value <= 0) return false;
                if (i == 0) first = value;
                else last = value;
            }
            if (last < first)
            {
                int swap = first;
                first = last;
                last = swap;
            }
            return true;
        }

        /// <summary>判断文本是否全为数字。</summary>
        /// <param name="text">待判定的文本。</param>
        /// <returns>全数字时返回 true。</returns>
        private static bool IsAllDigits(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (!char.IsDigit(text[i])) return false;
            }
            return true;
        }

        /// <summary>把一片区域并入结果文本，重复片段不再追加。</summary>
        /// <param name="builder">结果文本累加器。</param>
        /// <param name="fragment">区域片段。</param>
        private static void AppendFragment(System.Text.StringBuilder builder, string fragment)
        {
            string existing = builder.ToString();
            foreach (string token in existing.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (string.Equals(token, fragment, StringComparison.OrdinalIgnoreCase)) return;
            }
            if (builder.Length > 0) builder.Append(' ');
            builder.Append(fragment);
        }

        /// <summary>判断文本是否含数字，用于区分整列引用与坏掉的位置引用。</summary>
        /// <param name="text">区域片段。</param>
        /// <returns>含数字时返回 true。</returns>
        private static bool HasDigit(string text)
        {
            for (int i = 0; i < text.Length; i++)
            {
                if (char.IsDigit(text[i])) return true;
            }
            return false;
        }

        /// <summary>逐项复制页面尺寸与页边距设置。</summary>
        /// <param name="source">源打印设置。</param>
        /// <param name="target">目标打印设置。</param>
        private static void CopyPageSetup(PageSetup source, PageSetup target)
        {
            if (source == null) return;
            target.Width = source.Width;
            target.Height = source.Height;
            target.MarginLeft = source.MarginLeft;
            target.MarginRight = source.MarginRight;
            target.MarginTop = source.MarginTop;
            target.MarginBottom = source.MarginBottom;
            target.HeaderDistance = source.HeaderDistance;
            target.FooterDistance = source.FooterDistance;
        }

        /// <summary>深复制文档元数据的各文本与时间字段。</summary>
        /// <param name="source">源元数据；为 null 时返回空副本。</param>
        /// <returns>元数据副本。</returns>
        private static DocumentMetadata CloneMetadata(DocumentMetadata source)
        {
            DocumentMetadata copy = new DocumentMetadata();
            if (source == null) return copy;
            copy.Title = source.Title;
            copy.Author = source.Author;
            copy.Subject = source.Subject;
            copy.Keywords = source.Keywords;
            copy.Creator = source.Creator;
            copy.Created = source.Created;
            copy.Modified = source.Modified;
            return copy;
        }

        /// <summary>检测到的行循环块：起止行号与列表变量名。</summary>
        private sealed class LoopBlock
        {
            /// <summary>循环起始行（{{#each}} 标记所在行）。</summary>
            public readonly int StartRow;
            /// <summary>循环结束行（{{/each}} 标记所在行）。</summary>
            public readonly int EndRow;
            /// <summary>数据列表在上下文中的变量名。</summary>
            public readonly string Name;

            /// <summary>创建一个循环块描述。</summary>
            /// <param name="startRow">起始行号。</param>
            /// <param name="endRow">结束行号。</param>
            /// <param name="name">列表变量名。</param>
            public LoopBlock(int startRow, int endRow, string name)
            {
                StartRow = startRow;
                EndRow = endRow;
                Name = name;
            }
        }

        /// <summary>一行模板内容到输出行的落位记录，用于循环展开后重建合并区域。</summary>
        private sealed class Placement
        {
            /// <summary>模板源行号。</summary>
            public readonly int SourceRow;
            /// <summary>输出行号。</summary>
            public readonly int OutputRow;
            /// <summary>所属循环块；非循环行为 null。</summary>
            public readonly LoopBlock Loop;
            /// <summary>循环条目序号；非循环行为 -1。</summary>
            public readonly int ItemIndex;

            /// <summary>创建一条落位记录。</summary>
            /// <param name="sourceRow">模板源行号。</param>
            /// <param name="outputRow">输出行号。</param>
            /// <param name="loop">所属循环块（可为 null）。</param>
            /// <param name="itemIndex">循环条目序号（非循环为 -1）。</param>
            public Placement(int sourceRow, int outputRow, LoopBlock loop, int itemIndex)
            {
                SourceRow = sourceRow;
                OutputRow = outputRow;
                Loop = loop;
                ItemIndex = itemIndex;
            }
        }

        /// <summary>
        /// 模板行到输出行的映射：一行模板可能展开成多行输出，也可能被整块跳过（空循环、纯标记行）。
        /// 区域起点往后找最近的实存行，终点往前找，越界的行按累计位移外推。
        /// The template-row to output-row mapping: one template row may expand to several output rows or vanish
        /// entirely (an empty loop, a marker-only row). A range start looks forward for the nearest surviving row
        /// and a range end looks backward, while rows past the template's used range are extrapolated by the
        /// accumulated shift.
        /// </summary>
        private sealed class RowMap
        {
            /// <summary>每个模板行的首个输出行，0 表示该行没有落位。The first output row per template row; 0 when the row vanished.</summary>
            private readonly int[] _first;
            /// <summary>每个模板行的末个输出行。The last output row per template row.</summary>
            private readonly int[] _last;
            /// <summary>模板的最大实存行号。The largest used row of the template.</summary>
            private readonly int _lastSourceRow;
            /// <summary>输出的最后一行。The final output row.</summary>
            private readonly int _maxOutputRow;
            /// <summary>行号是否一行对一行（没有循环展开也没有行被吃掉）。Whether rows map one to one, with no expansion or dropped rows.</summary>
            private readonly bool _identity;

            /// <summary>由落位记录建立映射表。</summary>
            /// <param name="lastSourceRow">模板最大行号。</param>
            /// <param name="placements">落位记录列表。</param>
            public RowMap(int lastSourceRow, List<Placement> placements)
            {
                _lastSourceRow = lastSourceRow < 0 ? 0 : lastSourceRow;
                _first = new int[_lastSourceRow + 2];
                _last = new int[_lastSourceRow + 2];
                int maxOutputRow = 0;
                for (int i = 0; i < placements.Count; i++)
                {
                    Placement placement = placements[i];
                    if (placement.SourceRow < 1 || placement.SourceRow > _lastSourceRow) continue;
                    if (_first[placement.SourceRow] == 0) _first[placement.SourceRow] = placement.OutputRow;
                    _last[placement.SourceRow] = placement.OutputRow;
                    if (placement.OutputRow > maxOutputRow) maxOutputRow = placement.OutputRow;
                }
                _maxOutputRow = maxOutputRow;
                bool identity = true;
                for (int row = 1; row <= _lastSourceRow; row++)
                {
                    if (_first[row] != row || _last[row] != row)
                    {
                        identity = false;
                        break;
                    }
                }
                _identity = identity;
            }

            /// <summary>映射是否保持行号不变；不变时图表等装饰可直接共用实例。Whether the map keeps row numbers, letting charts and images be shared as-is.</summary>
            public bool IsIdentity { get { return _identity; } }

            /// <summary>取单个锚点行的落位；模板行被跳过时返回 0。Resolves one anchor row, returning 0 when the template row vanished.</summary>
            /// <param name="sourceRow">模板行号。</param>
            /// <returns>输出行号或 0。</returns>
            public int MappedRow(int sourceRow)
            {
                if (sourceRow < 1) return 0;
                if (sourceRow > _lastSourceRow) return Overflow(sourceRow);
                return _first[sourceRow];
            }

            /// <summary>取区域起点的落位：本行没落位时向后找最近的实存行。Resolves a range start, falling forward to the nearest surviving row.</summary>
            /// <param name="sourceRow">模板行号。</param>
            /// <returns>输出行号或 0。</returns>
            public int RangeStart(int sourceRow)
            {
                if (sourceRow > _lastSourceRow) return Overflow(sourceRow);
                for (int row = sourceRow < 1 ? 1 : sourceRow; row <= _lastSourceRow; row++)
                {
                    if (_first[row] > 0) return _first[row];
                }
                return _maxOutputRow;
            }

            /// <summary>取区域终点的落位：本行没落位时向前找最近的实存行。Resolves a range end, falling back to the nearest surviving row.</summary>
            /// <param name="sourceRow">模板行号。</param>
            /// <returns>输出行号或 0。</returns>
            public int RangeEnd(int sourceRow)
            {
                if (sourceRow > _lastSourceRow) return Overflow(sourceRow);
                for (int row = sourceRow > _lastSourceRow ? _lastSourceRow : sourceRow; row >= 1; row--)
                {
                    if (_last[row] > 0) return _last[row];
                }
                for (int row = 1; row <= _lastSourceRow; row++)
                {
                    if (_first[row] > 0) return _first[row];
                }
                return 0;
            }

            /// <summary>模板用尽范围之外的行按累计位移外推，打印区域这类"延伸到空白"的引用才不会凭空消失。Extrapolates rows past the template's used range by the accumulated shift, so print areas that run into the blank area do not vanish.</summary>
            /// <param name="sourceRow">模板行号。</param>
            /// <returns>输出行号；无任何落位时返回 0。</returns>
            private int Overflow(int sourceRow)
            {
                if (_maxOutputRow == 0) return 0;
                return _maxOutputRow + (sourceRow - _lastSourceRow);
            }
        }
    }
}
