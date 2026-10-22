namespace Wisdom.Excel
{
    /// <summary>
    /// 读取 .xlsx 或 .xls 工作簿时控制导入哪些隐藏内容的选项。
    /// Options that control which hidden content is imported when reading an .xlsx or .xls workbook.
    /// </summary>
    public sealed class ExcelReadOptions
    {
        /// <summary>是否导入隐藏的工作表（hidden 或 veryHidden 状态），默认导入。 Whether hidden worksheets (hidden or veryHidden state) are imported; enabled by default.</summary>
        public bool IncludeHiddenWorksheets { get; set; }

        /// <summary>是否导入隐藏行中的单元格，默认导入。 Whether cells inside hidden rows are imported; enabled by default.</summary>
        public bool IncludeHiddenRows { get; set; }

        /// <summary>是否导入隐藏列中的单元格，默认导入。 Whether cells inside hidden columns are imported; enabled by default.</summary>
        public bool IncludeHiddenColumns { get; set; }

        /// <summary>创建选项实例，默认包含全部隐藏内容（工作表、行、列）。 Creates an options instance that includes all hidden content (worksheets, rows and columns) by default.</summary>
        public ExcelReadOptions()
        {
            IncludeHiddenWorksheets = true;
            IncludeHiddenRows = true;
            IncludeHiddenColumns = true;
        }
    }
}
