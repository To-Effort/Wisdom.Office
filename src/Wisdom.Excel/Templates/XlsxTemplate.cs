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
            for (int i = 0; i < template.Worksheets.Count; i++)
            {
                Worksheet filled = ProcessSheet(template.Worksheets[i], context, options);
                result.Worksheets.Add(filled);
            }
            return result;
        }

        /// <summary>
        /// 处理单个工作表：复制表级属性与列宽，检测 {{#each}} 循环块，
        /// 按行（循环行按数据条目重复）逐行拷贝到输出表，最后按行映射重建合并区域。
        /// Processes one worksheet: copies sheet-level properties and column widths, detects {{#each}} loop
        /// blocks, copies rows one by one to the output sheet (loop rows repeated per data item), and finally
        /// rebuilds merged ranges from the row mapping.
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
    }
}
