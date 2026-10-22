using System;
using System.Collections.Generic;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 工作表的矩形区域视图：支持按 A1 引用或坐标创建，读写批量值、批量样式、合并与清空。
    /// A rectangular view over a worksheet: create it from an A1 reference or coordinates to read and
    /// write bulk values, bulk styles, merges and clears.
    /// </summary>
    public sealed class ExcelRange
    {
        /// <summary>区域所属工作表。The worksheet that owns the range.</summary>
        private readonly Worksheet _sheet;
        /// <summary>区域首行（从 1 开始）。The first row of the range (1-based).</summary>
        private readonly int _row1;
        /// <summary>区域首列（从 1 开始）。The first column of the range (1-based).</summary>
        private readonly int _column1;
        /// <summary>区域末行。The last row of the range.</summary>
        private readonly int _row2;
        /// <summary>区域末列。The last column of the range.</summary>
        private readonly int _column2;

        /// <summary>基于坐标创建区域；首末坐标颠倒时自动交换，超出 Excel 网格时抛异常。Creates a range from coordinates, swapping reversed corners and rejecting coordinates outside the Excel grid.</summary>
        /// <param name="sheet">所属工作表。The owning worksheet.</param>
        /// <param name="row1">首行。The first row.</param>
        /// <param name="column1">首列。The first column.</param>
        /// <param name="row2">末行。The last row.</param>
        /// <param name="column2">末列。The last column.</param>
        /// <exception cref="ArgumentNullException">sheet 为 null。sheet is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">坐标小于 1 或超出 Excel 网格。A coordinate is below 1 or outside the Excel grid.</exception>
        public ExcelRange(Worksheet sheet, int row1, int column1, int row2, int column2)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            if (row1 < 1 || row1 > CellAddress.MaxRowLiteral) throw new ArgumentOutOfRangeException("row1");
            if (row2 < 1 || row2 > CellAddress.MaxRowLiteral) throw new ArgumentOutOfRangeException("row2");
            if (column1 < 1 || column1 > CellAddress.MaxColumnLiteral) throw new ArgumentOutOfRangeException("column1");
            if (column2 < 1 || column2 > CellAddress.MaxColumnLiteral) throw new ArgumentOutOfRangeException("column2");
            _sheet = sheet;
            _row1 = Math.Min(row1, row2);
            _row2 = Math.Max(row1, row2);
            _column1 = Math.Min(column1, column2);
            _column2 = Math.Max(column1, column2);
        }

        /// <summary>区域所属工作表。The worksheet that owns the range.</summary>
        public Worksheet Worksheet { get { return _sheet; } }
        /// <summary>区域首行。The first row of the range.</summary>
        public int Row1 { get { return _row1; } }
        /// <summary>区域首列。The first column of the range.</summary>
        public int Column1 { get { return _column1; } }
        /// <summary>区域末行。The last row of the range.</summary>
        public int Row2 { get { return _row2; } }
        /// <summary>区域末列。The last column of the range.</summary>
        public int Column2 { get { return _column2; } }
        /// <summary>区域的行数。The number of rows in the range.</summary>
        public int RowCount { get { return _row2 - _row1 + 1; } }
        /// <summary>区域的列数。The number of columns in the range.</summary>
        public int ColumnCount { get { return _column2 - _column1 + 1; } }

        /// <summary>区域的 A1 引用（如 "B2:D10"）。The A1 reference of the range (such as "B2:D10").</summary>
        public string Address
        {
            get
            {
                return CellAddress.Format(_row1, _column1) + ":" + CellAddress.Format(_row2, _column2);
            }
        }

        /// <summary>区域左上角单元格。The top-left cell of the range.</summary>
        public Cell FirstCell
        {
            get { return _sheet.GetCell(_row1, _column1); }
        }

        /// <summary>按区域内 1 基偏移取得单元格，不存在时创建。Gets a cell by 1-based offsets inside the range, creating it when absent.</summary>
        /// <param name="rowOffset">行偏移（1 表示首行）。The row offset (1 is the first row).</param>
        /// <param name="columnOffset">列偏移（1 表示首列）。The column offset (1 is the first column).</param>
        /// <returns>对应单元格。The cell.</returns>
        /// <exception cref="ArgumentOutOfRangeException">偏移超出区域。The offset lies outside the range.</exception>
        public Cell this[int rowOffset, int columnOffset]
        {
            get
            {
                if (rowOffset < 1 || rowOffset > RowCount) throw new ArgumentOutOfRangeException("rowOffset");
                if (columnOffset < 1 || columnOffset > ColumnCount) throw new ArgumentOutOfRangeException("columnOffset");
                return _sheet.GetCell(_row1 + rowOffset - 1, _column1 + columnOffset - 1);
            }
        }

        /// <summary>左上角单元格的值；通过 A1 引用创建的 1x1 区域就是单格值。The top-left cell value; a 1x1 range created from an A1 reference is a single cell.</summary>
        public object Value
        {
            get { return FirstCell.Value; }
            set { FirstCell.Value = value; }
        }

        /// <summary>区域的二维值快照（行、列均从 0 开始），空单元格为 null；读取不会创建单元格。A 2-D snapshot of the range values (row and column start at 0), with null for empty cells; reading never creates cells.</summary>
        public object[,] Values
        {
            get
            {
                object[,] values = new object[RowCount, ColumnCount];
                for (int row = 0; row < RowCount; row++)
                {
                    for (int column = 0; column < ColumnCount; column++)
                    {
                        Cell cell = _sheet.FindCell(_row1 + row, _column1 + column);
                        if (cell != null) values[row, column] = cell.Value;
                    }
                }
                return values;
            }
            set
            {
                if (value == null) throw new ArgumentNullException("value");
                if (value.GetLength(0) < RowCount || value.GetLength(1) < ColumnCount)
                {
                    throw new ArgumentException("The value array is smaller than the range.", "value");
                }
                for (int row = 0; row < RowCount; row++)
                {
                    for (int column = 0; column < ColumnCount; column++)
                    {
                        SetCellValue(_row1 + row, _column1 + column, value[row, column]);
                    }
                }
            }
        }

        /// <summary>区域样式：读取返回左上角单元格样式，赋值时把同一实例应用到区域全部单元格。The range style: reading returns the top-left style, assigning applies the same instance to every cell in the range.</summary>
        public CellStyle Style
        {
            get { return FirstCell.EnsureStyle(); }
            set { ApplyStyle(value); }
        }

        /// <summary>把样式应用到区域每个单元格；为 null 时清除样式。Applies a style to every cell in the range, clearing styles when null.</summary>
        /// <param name="style">目标样式；写入前会复制一份，避免各单元格共享可变实例。The target style; it is cloned per cell so the cells do not share one mutable instance.</param>
        public void ApplyStyle(CellStyle style)
        {
            for (int row = _row1; row <= _row2; row++)
            {
                for (int column = _column1; column <= _column2; column++)
                {
                    if (style == null)
                    {
                        Cell existing = _sheet.FindCell(row, column);
                        if (existing != null) existing.Style = null;
                    }
                    else
                    {
                        _sheet.GetCell(row, column).Style = style.Clone();
                    }
                }
            }
        }

        /// <summary>对区域每个单元格配置样式（自动创建样式实例）。Configures the style of every cell in the range, creating style instances as needed.</summary>
        /// <param name="configure">样式配置委托。The style configuration delegate.</param>
        public void SetStyle(Action<CellStyle> configure)
        {
            if (configure == null) throw new ArgumentNullException("configure");
            for (int row = _row1; row <= _row2; row++)
            {
                for (int column = _column1; column <= _column2; column++)
                {
                    configure(_sheet.GetCell(row, column).EnsureStyle());
                }
            }
        }

        /// <summary>把区域值填充为同一对象（如整片底色标题或占位符）。Fills every cell in the range with the same value (such as a title or placeholder).</summary>
        /// <param name="value">填充值。The fill value.</param>
        public void SetValue(object value)
        {
            for (int row = _row1; row <= _row2; row++)
            {
                for (int column = _column1; column <= _column2; column++)
                {
                    SetCellValue(row, column, value);
                }
            }
        }

        /// <summary>合并整个区域并去重相同区域。Merges the whole range, de-duplicating an identical existing merge.</summary>
        public void Merge()
        {
            for (int i = 0; i < _sheet.MergedRanges.Count; i++)
            {
                MergedRange range = _sheet.MergedRanges[i];
                if (range.Row == _row1 && range.Column == _column1 && range.RowSpan == RowCount && range.ColumnSpan == ColumnCount)
                {
                    return;
                }
            }
            _sheet.MergedRanges.Add(new MergedRange(_row1, _column1, RowCount, ColumnCount));
        }

        /// <summary>解除完全落在区域内的合并。Removes merges that lie completely inside the range.</summary>
        public void Unmerge()
        {
            for (int i = _sheet.MergedRanges.Count - 1; i >= 0; i--)
            {
                MergedRange range = _sheet.MergedRanges[i];
                if (range.Row >= _row1 && range.Column >= _column1 && range.LastRow <= _row2 && range.LastColumn <= _column2)
                {
                    _sheet.MergedRanges.RemoveAt(i);
                }
            }
        }

        /// <summary>清空区域的值、公式与超链接，保留样式；清空后没有任何内容的单元格会被移除。Clears values, formulas and hyperlinks while keeping styles; cells left with no content are removed.</summary>
        public void Clear()
        {
            for (int row = _row1; row <= _row2; row++)
            {
                for (int column = _column1; column <= _column2; column++)
                {
                    Cell cell = _sheet.FindCell(row, column);
                    if (cell == null) continue;
                    cell.Value = null;
                    cell.Formula = null;
                    cell.Hyperlink = null;
                    cell.HyperlinkTooltip = null;
                    if (cell.Style == null) _sheet.RemoveCell(row, column);
                }
            }
        }

        /// <summary>清空区域的全部内容与样式、解除区域内合并。Clears all content and styles in the range and removes merges inside it.</summary>
        public void ClearAll()
        {
            for (int row = _row1; row <= _row2; row++)
            {
                for (int column = _column1; column <= _column2; column++)
                {
                    _sheet.RemoveCell(row, column);
                }
            }
            Unmerge();
        }

        /// <summary>把区域的值、公式、超链接与样式复制到同样大小的目标区域。Copies values, formulas, hyperlinks and styles to a target range of the same size.</summary>
        /// <param name="target">目标区域。The target range.</param>
        /// <exception cref="ArgumentNullException">target 为 null。target is null.</exception>
        /// <exception cref="ArgumentException">目标区域大小不同。The target range has a different size.</exception>
        public void CopyTo(ExcelRange target)
        {
            if (target == null) throw new ArgumentNullException("target");
            if (target.RowCount != RowCount || target.ColumnCount != ColumnCount)
            {
                throw new ArgumentException("The target range must have the same size as the source.", "target");
            }
            for (int row = 0; row < RowCount; row++)
            {
                for (int column = 0; column < ColumnCount; column++)
                {
                    Cell source = _sheet.FindCell(_row1 + row, _column1 + column);
                    int targetRow = target._row1 + row;
                    int targetColumn = target._column1 + column;
                    if (source == null)
                    {
                        target._sheet.RemoveCell(targetRow, targetColumn);
                        continue;
                    }
                    Cell destination = target._sheet.GetCell(targetRow, targetColumn);
                    destination.Value = source.Value;
                    destination.Formula = source.Formula;
                    destination.Hyperlink = source.Hyperlink;
                    destination.HyperlinkTooltip = source.HyperlinkTooltip;
                    destination.Style = source.Style == null ? null : source.Style.Clone();
                }
            }
        }

        /// <summary>按索引器语义更新单格：值为 null 时清除值，不创建无意义的新单元格。Updates one cell with indexer semantics: a null value clears it and no needless cell is created.</summary>
        /// <param name="row">行号。The row number.</param>
        /// <param name="column">列号。The column number.</param>
        /// <param name="value">新值。The new value.</param>
        private void SetCellValue(int row, int column, object value)
        {
            Cell cell = _sheet.FindCell(row, column);
            if (cell == null)
            {
                if (value == null) return;
                cell = _sheet.GetCell(row, column);
            }
            cell.Value = value;
        }

        /// <summary>按区域内单元格的显示文本自动调整列宽（中文等全角字符按两字符计）。Adjusts column widths to fit the displayed text of the cells in the range, counting full-width characters (such as CJK) as two.</summary>
        /// <param name="minWidth">最小列宽（Excel 字符单位），默认 8.43。The minimum width in Excel character units; 8.43 by default.</param>
        /// <param name="maxWidth">最大列宽，默认 60，内部上限为 255。The maximum width; 60 by default and capped at 255.</param>
        public void AutoFitColumns(double minWidth = 8.43, double maxWidth = 60)
        {
            if (double.IsNaN(minWidth) || minWidth < 0) throw new ArgumentOutOfRangeException("minWidth");
            if (double.IsNaN(maxWidth) || maxWidth < minWidth) throw new ArgumentOutOfRangeException("maxWidth");
            if (maxWidth > 255) maxWidth = 255;
            for (int column = _column1; column <= _column2; column++)
            {
                double width = 0;
                for (int row = _row1; row <= _row2; row++)
                {
                    Cell cell = _sheet.FindCell(row, column);
                    if (cell == null) continue;
                    string text = cell.DisplayText();
                    if (string.IsNullOrEmpty(text)) continue;
                    double measured = MeasureTextWidth(text);
                    if (measured > width) width = measured;
                }
                width = width == 0 ? minWidth : width + 2;
                if (width < minWidth) width = minWidth;
                if (width > maxWidth) width = maxWidth;
                _sheet.SetColumnWidth(column, width);
            }
        }

        /// <summary>测量多行文本的最大显示宽度（制表符按 4、全角字符按 2 计）。Measures the widest displayed line of a multi-line text (tabs count as 4 and full-width characters as 2).</summary>
        /// <param name="text">显示文本。The displayed text.</param>
        /// <returns>最大行宽（字符单位）。The widest line width in character units.</returns>
        private static double MeasureTextWidth(string text)
        {
            double widest = 0;
            double current = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '\r') continue;
                if (c == '\n')
                {
                    if (current > widest) widest = current;
                    current = 0;
                    continue;
                }
                current += c == '\t' ? 4 : (IsFullWidth(c) ? 2 : 1);
            }
            return current > widest ? current : widest;
        }

        /// <summary>判断字符是否按全角宽度显示（CJK、假名、谚文与全角符号）。Decides whether a character renders at full width (CJK, kana, Hangul and full-width symbols).</summary>
        /// <param name="c">待判断字符。The character to test.</param>
        /// <returns>全角返回 true。True for full-width characters.</returns>
        private static bool IsFullWidth(char c)
        {
            return (c >= 0x1100 && c <= 0x115F)
                || (c >= 0x2E80 && c <= 0xA4CF)
                || (c >= 0xAC00 && c <= 0xD7A3)
                || (c >= 0xF900 && c <= 0xFAFF)
                || (c >= 0xFE30 && c <= 0xFE6F)
                || (c >= 0xFF00 && c <= 0xFF60)
                || (c >= 0xFFE0 && c <= 0xFFE6);
        }

        /// <summary>返回区域引用。Returns the range reference.</summary>
        /// <returns>A1 引用文本。The A1 reference text.</returns>
        public override string ToString()
        {
            return Address;
        }
    }

    /// <summary>工作表区域扩展方法：按引用或坐标取得 <see cref="ExcelRange"/>。Worksheet range extensions that obtain an <see cref="ExcelRange"/> from a reference or coordinates.</summary>
    public static class ExcelRangeExtensions
    {
        /// <summary>按 A1 区域引用取得区域，如 "A1:C5"；单个引用视为 1x1 区域。Gets a range from an A1 reference such as "A1:C5"; a single reference is a 1x1 range.</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="reference">A1 区域引用。The A1 range reference.</param>
        /// <returns>区域对象。The range object.</returns>
        /// <exception cref="ArgumentNullException">sheet 为 null。sheet is null.</exception>
        /// <exception cref="FormatException">引用格式非法。The reference is malformed.</exception>
        public static ExcelRange Range(this Worksheet sheet, string reference)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            int row1, column1, row2, column2;
            CellAddress.ParseRange(reference, out row1, out column1, out row2, out column2);
            return new ExcelRange(sheet, row1, column1, row2, column2);
        }

        /// <summary>按坐标取得区域。Gets a range from coordinates.</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <param name="row1">首行。The first row.</param>
        /// <param name="column1">首列。The first column.</param>
        /// <param name="row2">末行。The last row.</param>
        /// <param name="column2">末列。The last column.</param>
        /// <returns>区域对象。The range object.</returns>
        public static ExcelRange Range(this Worksheet sheet, int row1, int column1, int row2, int column2)
        {
            return new ExcelRange(sheet, row1, column1, row2, column2);
        }

        /// <summary>取得工作表的已用区域；空表返回 A1 单元格。Gets the used range of a worksheet, returning cell A1 for an empty sheet.</summary>
        /// <param name="sheet">工作表。The worksheet.</param>
        /// <returns>已用区域。The used range.</returns>
        public static ExcelRange UsedRange(this Worksheet sheet)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            return new ExcelRange(sheet, 1, 1, Math.Max(1, sheet.LastRow), Math.Max(1, sheet.LastColumn));
        }
    }
}
