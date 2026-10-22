using System;
using System.Collections.Generic;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>PDF 页面上测得的一段水平文本运行。A horizontal run of text measured on a PDF page.</summary>
    public sealed class PdfTextFragment
    {
        /// <summary>该段的文本内容。The text content of this fragment.</summary>
        public string Text { get; internal set; }
        /// <summary>左上角横坐标（PDF 用户单位）。X coordinate of the top-left corner, in PDF user units.</summary>
        public double X { get; internal set; }
        /// <summary>左上角纵坐标。Y coordinate of the top-left corner.</summary>
        public double Y { get; internal set; }
        /// <summary>文本包围盒宽度。Width of the text bounding box.</summary>
        public double Width { get; internal set; }
        /// <summary>文本包围盒高度。Height of the text bounding box.</summary>
        public double Height { get; internal set; }
        /// <summary>基线位置（横排为设备空间纵坐标，斜排为垂直于流向的坐标），用于判定是否同一行。Baseline position (the device Y for horizontal text, the perpendicular coordinate for rotated text), used to decide whether fragments share a line.</summary>
        public double Baseline { get; internal set; }
        /// <summary>生效字号。The effective font size.</summary>
        public double FontSize { get; internal set; }
        /// <summary>文本流向相对水平方向的夹角，弧度制，0 表示正常横排，正值为顺时针倾斜。The text flow angle relative to horizontal, in radians; 0 is normal horizontal text and positive values lean clockwise.</summary>
        public double Angle { get; internal set; }

        /// <summary>包围盒右边界。Right edge of the bounding box.</summary>
        public double Right { get { return X + Width; } }
        /// <summary>包围盒下边界。Bottom edge of the bounding box.</summary>
        public double Bottom { get { return Y + Height; } }
    }

    /// <summary>共享同一条基线的一组文本片段。A group of text fragments sharing the same baseline.</summary>
    public sealed class PdfTextLine
    {
        /// <summary>按片段顺序拼接后的整行文本。The full line text, fragments concatenated in order.</summary>
        public string Text { get; internal set; }
        /// <summary>行包围盒左上角横坐标。X coordinate of the top-left corner of the line bounding box.</summary>
        public double X { get; internal set; }
        /// <summary>行包围盒左上角纵坐标。Y coordinate of the top-left corner of the line bounding box.</summary>
        public double Y { get; internal set; }
        /// <summary>行包围盒宽度。Width of the line bounding box.</summary>
        public double Width { get; internal set; }
        /// <summary>行包围盒高度。Height of the line bounding box.</summary>
        public double Height { get; internal set; }
        /// <summary>组成该行的片段，按阅读顺序排列。The fragments forming this line, in reading order.</summary>
        public IReadOnlyList<PdfTextFragment> Fragments { get; internal set; }

        /// <summary>行包围盒右边界。Right edge of the line bounding box.</summary>
        public double Right { get { return X + Width; } }
        /// <summary>行包围盒下边界。Bottom edge of the line bounding box.</summary>
        public double Bottom { get { return Y + Height; } }
    }

    /// <summary>由线条与文本间距推断出的表格及其单元格文本。A table inferred from drawn lines and text spacing, together with its cell text.</summary>
    public sealed class PdfTable
    {
        /// <summary>行数。Number of rows.</summary>
        public int RowCount { get; internal set; }
        /// <summary>列数。Number of columns.</summary>
        public int ColumnCount { get; internal set; }
        /// <summary>表格包围盒左上角横坐标。X coordinate of the top-left corner of the table bounding box.</summary>
        public double X { get; internal set; }
        /// <summary>表格包围盒左上角纵坐标。Y coordinate of the top-left corner of the table bounding box.</summary>
        public double Y { get; internal set; }
        /// <summary>表格包围盒宽度。Width of the table bounding box.</summary>
        public double Width { get; internal set; }
        /// <summary>表格包围盒高度。Height of the table bounding box.</summary>
        public double Height { get; internal set; }
        /// <summary>按行列存放的单元格文本，元素可能为 null。Cell text stored by row and column; individual elements may be null.</summary>
        public string[,] Cells { get; internal set; }

        /// <summary>
        /// 读取指定单元格的文本。
        /// Reads the text of a single cell.
        /// </summary>
        /// <param name="row">从零开始的行号。Zero-based row index.</param>
        /// <param name="column">从零开始的列号。Zero-based column index.</param>
        /// <returns>单元格文本；坐标越界、单元格为 null 或表格无数据时返回空字符串。The cell text; an empty string when the index is out of range, the cell is null, or the table has no data.</returns>
        public string GetCell(int row, int column)
        {
            if (Cells == null) return string.Empty;
            if (row < 0 || row >= RowCount || column < 0 || column >= ColumnCount) return string.Empty;
            return Cells[row, column] ?? string.Empty;
        }

        /// <summary>
        /// 以制表符分列、换行分行的形式输出表格文本。
        /// Outputs the table as plain text with tabs between columns and newlines between rows.
        /// </summary>
        /// <returns>表格的纯文本表示。The plain-text representation of the table.</returns>
        public string ToText()
        {
            return ToText("\t", "\n");
        }

        /// <summary>
        /// 以自定义分隔符输出表格文本。
        /// Outputs the table as plain text using custom separators.
        /// </summary>
        /// <param name="cellSeparator">列之间的分隔符。Separator inserted between columns.</param>
        /// <param name="rowSeparator">行之间的分隔符。Separator inserted between rows.</param>
        /// <returns>表格的纯文本表示。The plain-text representation of the table.</returns>
        public string ToText(string cellSeparator, string rowSeparator)
        {
            StringBuilder builder = new StringBuilder();
            for (int row = 0; row < RowCount; row++)
            {
                if (row > 0) builder.Append(rowSeparator);
                for (int column = 0; column < ColumnCount; column++)
                {
                    if (column > 0) builder.Append(cellSeparator);
                    builder.Append(GetCell(row, column));
                }
            }
            return builder.ToString();
        }
    }
}
