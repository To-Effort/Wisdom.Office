using System;
using System.Collections.Generic;
using System.Globalization;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>公式求值过程中的 Excel 错误类型。Excel error types produced while evaluating a formula.</summary>
    public enum FxError
    {
        /// <summary>无错误。No error.</summary>
        None = 0,
        /// <summary>#NULL!：区域交集为空。#NULL!: the intersection of two ranges is empty.</summary>
        Null,
        /// <summary>#DIV/0!：除数为零。#DIV/0!: division by zero.</summary>
        Div0,
        /// <summary>#REF!：引用无效。#REF!: an invalid reference.</summary>
        Ref,
        /// <summary>#VALUE!：值类型不匹配。#VALUE!: a value has the wrong type.</summary>
        Value,
        /// <summary>#NAME?：无法识别的名称。#NAME?: an unrecognised name.</summary>
        Name,
        /// <summary>#NUM!：数值超出范围。#NUM!: a number is out of range.</summary>
        Num,
        /// <summary>#N/A：找不到可用值。#N/A: a value is not available.</summary>
        Na,
        /// <summary>#GETTING_DATA：数据获取中（保留）。#GETTING_DATA: data is being fetched (reserved).</summary>
        GettingData
    }

    /// <summary>公式求值期间的值类别。The value category while a formula is being evaluated.</summary>
    internal enum FxKind
    {
        /// <summary>空单元格。An empty cell.</summary>
        Empty,
        /// <summary>数值。A number.</summary>
        Number,
        /// <summary>文本。A string.</summary>
        Text,
        /// <summary>布尔。A logical value.</summary>
        Logical,
        /// <summary>错误。An error.</summary>
        Error,
        /// <summary>引用区域。A reference range.</summary>
        Range
    }

    /// <summary>
    /// 公式引擎的运行时值：数值 / 文本 / 布尔 / 错误 / 引用区域。
    /// The runtime value used by the formula engine: number, text, logical, error, or a reference range.
    /// </summary>
    /// <remarks>
    /// 标量与引用统一在此类型中，函数既可接收标量也可接收区域；引用保留所属工作表与左上/右下坐标，
    /// 供 SUM/LOOKUP 展平以及 ROW/COLUMN 等函数读取。单元格惰性求值时由 <see cref="Evaluator"/> 填充 <see cref="Cells"/>。
    /// Scalars and references share one type so functions accept either; a reference keeps its worksheet plus the
    /// top-left/bottom-right coordinates for flattening (SUM, LOOKUP) and coordinate functions (ROW, COLUMN).
    /// Cell values are filled lazily by the <see cref="Evaluator"/> into <see cref="Cells"/>.
    /// </remarks>
    internal sealed class FxValue
    {
        /// <summary>值类别。The value category.</summary>
        public FxKind Kind;
        /// <summary>数值载荷（Kind==Number）。The number payload (Kind==Number).</summary>
        public double Number;
        /// <summary>文本载荷（Kind==Text）。The text payload (Kind==Text).</summary>
        public string Text;
        /// <summary>布尔载荷（Kind==Logical）。The logical payload (Kind==Logical).</summary>
        public bool Bool;
        /// <summary>错误载荷（Kind==Error）。The error payload (Kind==Error).</summary>
        public FxError Error;

        /// <summary>引用所属工作表。The worksheet a reference belongs to.</summary>
        public Worksheet Sheet;
        /// <summary>引用左上角行（1 基）。The 1-based top row of a reference.</summary>
        public int R1;
        /// <summary>引用左上角列（1 基）。The 1-based left column of a reference.</summary>
        public int C1;
        /// <summary>引用右下角行（1 基）。The 1-based bottom row of a reference.</summary>
        public int R2;
        /// <summary>引用右下角列（1 基）。The 1-based right column of a reference.</summary>
        public int C2;
        /// <summary>惰性物化的区域单元格网格。The lazily-materialised grid of cells for a range.</summary>
        public List<List<FxValue>> Cells;

        private static readonly FxValue EmptyInstance = new FxValue { Kind = FxKind.Empty };

        /// <summary>空值单例。The shared empty value.</summary>
        public static FxValue Empty { get { return EmptyInstance; } }

        /// <summary>创建数值。Creates a number.</summary>
        /// <param name="v">数值。The number.</param>
        /// <returns>数值实例。A number value.</returns>
        public static FxValue Num(double v)
        {
            return new FxValue { Kind = FxKind.Number, Number = v };
        }

        /// <summary>创建文本。Creates a text value.</summary>
        /// <param name="v">文本。The text.</param>
        /// <returns>文本实例。A text value.</returns>
        public static FxValue Txt(string v)
        {
            return new FxValue { Kind = FxKind.Text, Text = v ?? string.Empty };
        }

        /// <summary>创建布尔。Creates a logical value.</summary>
        /// <param name="v">布尔。The logical.</param>
        /// <returns>布尔实例。A logical value.</returns>
        public static FxValue Logical(bool v)
        {
            return new FxValue { Kind = FxKind.Logical, Bool = v };
        }

        /// <summary>创建错误。Creates an error value.</summary>
        /// <param name="e">错误类型。The error.</param>
        /// <returns>错误实例。An error value.</returns>
        public static FxValue Err(FxError e)
        {
            return new FxValue { Kind = FxKind.Error, Error = e };
        }

        /// <summary>从单元格值对象转换为运行时值。Converts a cell value object into a runtime value.</summary>
        /// <param name="o">单元格的原始值。The raw cell value.</param>
        /// <returns>运行时值。The runtime value.</returns>
        public static FxValue FromObject(object o)
        {
            if (o == null) return Empty;
            if (o is double) return Num((double)o);
            if (o is float) return Num(Convert.ToDouble((float)o, CultureInfo.InvariantCulture));
            if (o is decimal) return Num(Convert.ToDouble((decimal)o, CultureInfo.InvariantCulture));
            if (o is int) return Num((int)o);
            if (o is long) return Num((long)o);
            if (o is short) return Num((short)o);
            if (o is byte) return Num((byte)o);
            if (o is bool) return Logical((bool)o);
            if (o is DateTime) return Num(DateUtils.DateTimeToExcelSerial((DateTime)o, false));
            if (o is TimeSpan) return Num(((TimeSpan)o).TotalDays);
            string s = o as string;
            if (s != null) return Txt(s);
            return Txt(Convert.ToString(o, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        /// <summary>是否为错误值。Whether this value is an error.</summary>
        public bool IsError { get { return Kind == FxKind.Error; } }

        /// <summary>把错误转成可写回单元格的字符串文本；非错误返回 null。Turns an error into the string text to store back on a cell; null for non-errors.</summary>
        /// <returns>错误文本或 null。The error text, or null.</returns>
        public string ErrorText()
        {
            if (Kind != FxKind.Error) return null;
            switch (Error)
            {
                case FxError.Null: return "#NULL!";
                case FxError.Div0: return "#DIV/0!";
                case FxError.Ref: return "#REF!";
                case FxError.Value: return "#VALUE!";
                case FxError.Name: return "#NAME?";
                case FxError.Num: return "#NUM!";
                case FxError.Na: return "#N/A";
                case FxError.GettingData: return "#GETTING_DATA";
                default: return "#VALUE!";
            }
        }

        /// <summary>把结果写回单元格 <see cref="Cell.Value"/>：错误写文本，空写 null，其余保留数值/文本/布尔。Writes the result into <see cref="Cell.Value"/>: errors become text, empty becomes null, numbers/strings/booleans are kept.</summary>
        /// <returns>适合存入单元格的值对象。A value object suitable for storing on a cell.</returns>
        public object ToCellValue()
        {
            switch (Kind)
            {
                case FxKind.Empty: return null;
                case FxKind.Number: return Number;
                case FxKind.Text: return Text;
                case FxKind.Logical: return Bool;
                case FxKind.Error: return ErrorText();
                default: return null;
            }
        }
    }
}
