using System;
using System.Collections.Generic;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 共享公式（<c>&lt;f t="shared" si="0"&gt;</c>）的登记表：主格带公式文本与坐标，从格只带 si。
    /// Registration table for shared formulas (<c>&lt;f t="shared" si="0"&gt;</c>): the master cell carries the
    /// formula text and its coordinates, the dependents only carry the si number.
    /// </summary>
    internal sealed class XlsxSharedFormulas
    {
        /// <summary>si 编号到主公式的映射。Map from the si number to its master formula.</summary>
        private readonly Dictionary<int, Master> _masters = new Dictionary<int, Master>();

        /// <summary>登记一条主公式。Remembers one master formula.</summary>
        /// <param name="si">共享编号。The shared index.</param>
        /// <param name="formula">去掉前导等号的公式文本。Formula text without the leading equals sign.</param>
        /// <param name="row">主格行号。Row of the master cell.</param>
        /// <param name="column">主格列号。Column of the master cell.</param>
        internal void Remember(int si, string formula, int row, int column)
        {
            if (string.IsNullOrEmpty(formula)) return;
            _masters[si] = new Master(formula, row, column);
        }

        /// <summary>按从格与主格的偏移还原公式；si 未登记时返回 null。Restores the dependent formula by the offset to the master, returning null when the si was never registered.</summary>
        /// <param name="si">共享编号。The shared index.</param>
        /// <param name="row">从格行号。Row of the dependent cell.</param>
        /// <param name="column">从格列号。Column of the dependent cell.</param>
        /// <returns>还原的公式文本，或 null。 The translated formula, or null.</returns>
        internal string Resolve(int si, int row, int column)
        {
            Master master;
            if (!_masters.TryGetValue(si, out master)) return null;
            return XlsxSharedFormulaTranslator.Translate(master.Formula, row - master.Row, column - master.Column);
        }

        /// <summary>一条主公式及其坐标。One master formula and its coordinates.</summary>
        private sealed class Master
        {
            internal readonly string Formula;
            internal readonly int Row;
            internal readonly int Column;

            internal Master(string formula, int row, int column)
            {
                Formula = formula;
                Row = row;
                Column = column;
            }
        }
    }

    /// <summary>
    /// 把相对引用按行列偏移平移，绝对标记（<c>$</c>）保持不变，越出网格的引用落成 <c>#REF!</c>。
    /// Shifts relative references by a row/column delta, keeps the absolute <c>$</c> marks and drops references
    /// that leave the grid as <c>#REF!</c>.
    /// </summary>
    internal static class XlsxSharedFormulaTranslator
    {
        /// <summary>平移公式里的每个 A1 引用。Translates every A1 reference inside a formula.</summary>
        /// <param name="formula">主公式文本。The master formula text.</param>
        /// <param name="rowDelta">行偏移。The row delta.</param>
        /// <param name="columnDelta">列偏移。The column delta.</param>
        /// <returns>平移后的公式；输入为空或零偏移时原样返回。The translated formula; returned as-is when the input is empty or the delta is zero.</returns>
        internal static string Translate(string formula, int rowDelta, int columnDelta)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            if (rowDelta == 0 && columnDelta == 0) return formula;
            StringBuilder result = new StringBuilder(formula.Length + 8);
            int index = 0;
            while (index < formula.Length)
            {
                char current = formula[index];
                if (current == '"')
                {
                    index = CopyUntil(formula, index, current, result);
                    continue;
                }
                // 带引号的表名整段照抄：表名里的 "C1" 不是单元格引用。
                // A quoted sheet name is copied verbatim: a "C1" inside a sheet name is not a reference.
                if (current == '\'')
                {
                    index = CopyUntil(formula, index, current, result);
                    continue;
                }
                if ((current == '$' || IsLetter(current)) && IsStartOfReference(formula, index))
                {
                    int after;
                    string translated = ReadReference(formula, index, rowDelta, columnDelta, out after);
                    result.Append(translated);
                    index = after;
                    continue;
                }
                result.Append(current);
                index++;
            }
            return result.ToString();
        }

        /// <summary>从指定位置照抄到下一个同字符（含该字符），返回后续下标。Copies from the position through the next identical character (inclusive) and returns the following index.</summary>
        /// <param name="formula">公式文本。The formula text.</param>
        /// <param name="start">起始下标。The start index.</param>
        /// <param name="quote">定界字符。The delimiter character.</param>
        /// <param name="result">输出缓冲。The output builder.</param>
        /// <returns>下一个待处理下标。The next index to process.</returns>
        private static int CopyUntil(string formula, int start, char quote, StringBuilder result)
        {
            int end = start + 1;
            while (end < formula.Length && formula[end] != quote) end++;
            if (end < formula.Length) end++;
            result.Append(formula.Substring(start, end - start));
            return end;
        }

        /// <summary>判断下标处是否是一个需要平移的单元格引用起点。Whether the index starts a cell reference that must be translated.</summary>
        /// <param name="formula">公式文本。The formula text.</param>
        /// <param name="start">起始下标。The start index.</param>
        /// <returns>是引用起点返回 true。True when the index begins a reference.</returns>
        private static bool IsStartOfReference(string formula, int start)
        {
            // 前一个字符是字母、数字、下划线或点时，这里只是更长标识符的一段。
            // 感叹号**不**在此列：`Sheet1!A1` 里的 `A1` 正是需要跟着从格平移的相对引用，
            // 把它排掉会让所有带表名前缀的共享公式在每个从格里都停在同一行。
            // When the previous character is a letter, digit, underscore or dot this is only a slice of a
            // longer identifier. The bang is deliberately absent from that list: the A1 in `Sheet1!A1` is a
            // relative reference that must shift with the dependent, and excluding it left every
            // sheet-qualified shared formula pinned to its master's row.
            if (start > 0)
            {
                char previous = formula[start - 1];
                if (IsLetter(previous) || IsDigit(previous) || previous == '_' || previous == '.')
                {
                    return false;
                }
            }
            int index = start;
            if (formula[index] == '$') index++;
            int letters = 0;
            while (index < formula.Length && IsLetter(formula[index]) && letters < 4)
            {
                letters++;
                index++;
            }
            if (letters == 0 || letters >= 4) return false;
            // 行号前的 $ 只标记绝对行，不能因此判定"没有行号"。
            // The $ before the row number only marks an absolute row and must not read as "no row number".
            if (index < formula.Length && formula[index] == '$') index++;
            int digits = 0;
            while (index < formula.Length && IsDigit(formula[index]) && digits < 8)
            {
                digits++;
                index++;
            }
            // 没有行号的是整列引用（A:A）或函数名开头，不该平移。
            // Without a row number this is a whole-column reference (A:A) or the head of a function name, neither
            // of which is translated.
            if (digits == 0) return false;
            if (index >= formula.Length) return true;
            char next = formula[index];
            // 后面紧跟括号是函数名（LOG10( 里的 "G10" 不是引用），紧跟感叹号是表名前缀。
            // A following parenthesis marks a function name (the "G10" inside LOG10( is not a reference) and a
            // following bang is a sheet-name prefix.
            return next != '(' && next != '!' && !IsLetter(next) && !IsDigit(next) && next != '_' && next != '.';
        }

        /// <summary>读出并平移一个引用，返回其后的下标。Reads one reference, translates it and returns the following index.</summary>
        /// <param name="formula">公式文本。The formula text.</param>
        /// <param name="start">起始下标。The start index.</param>
        /// <param name="rowDelta">行偏移。The row delta.</param>
        /// <param name="columnDelta">列偏移。The column delta.</param>
        /// <param name="after">输出引用结束后的下标。Outputs the index just after the reference.</param>
        /// <returns>平移后的引用文本，或 <c>#REF!</c>。The translated reference, or <c>#REF!</c>.</returns>
        private static string ReadReference(string formula, int start, int rowDelta, int columnDelta, out int after)
        {
            int index = start;
            bool absoluteColumn = false;
            bool absoluteRow = false;
            if (formula[index] == '$')
            {
                absoluteColumn = true;
                index++;
            }
            int column = 0;
            while (index < formula.Length && IsLetter(formula[index]))
            {
                column = column * 26 + (char.ToUpperInvariant(formula[index]) - 'A' + 1);
                index++;
            }
            if (index < formula.Length && formula[index] == '$')
            {
                absoluteRow = true;
                index++;
            }
            int row = 0;
            while (index < formula.Length && IsDigit(formula[index]))
            {
                long wider = (long)row * 10 + (formula[index] - '0');
                // 行号越界后不再继续累加：下面的网格判定会把这种引用判成 #REF!，长数字也不会溢出 int。
                // Once the row passes the grid it stops accumulating: the range check below turns such a
                // reference into #REF! and a long digit run can no longer overflow int.
                row = wider > CellAddress.MaxRowLiteral ? CellAddress.MaxRowLiteral + 1 : (int)wider;
                index++;
            }
            after = index;
            int targetColumn = absoluteColumn ? column : column + columnDelta;
            int targetRow = absoluteRow ? row : row + rowDelta;
            if (targetRow < 1 || targetRow > CellAddress.MaxRowLiteral
                || targetColumn < 1 || targetColumn > CellAddress.MaxColumnLiteral)
            {
                return "#REF!";
            }
            return (absoluteColumn ? "$" : string.Empty) + ColumnLetters(targetColumn)
                + (absoluteRow ? "$" : string.Empty) + targetRow.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        /// <summary>列号转 A1 列字母。Converts a column number into A1 column letters.</summary>
        /// <param name="column">1 基列号。The 1-based column number.</param>
        /// <returns>列字母。The column letters.</returns>
        private static string ColumnLetters(int column)
        {
            StringBuilder builder = new StringBuilder(3);
            while (column > 0)
            {
                int remainder = (column - 1) % 26;
                builder.Insert(0, (char)('A' + remainder));
                column = (column - 1) / 26;
            }
            return builder.ToString();
        }

        /// <summary>是否为 A1 列字母（仅 ASCII）。Whether the character is an A1 column letter (ASCII only).</summary>
        /// <param name="c">待判定的字符。The character to test.</param>
        /// <returns>true 表示可作列名。True when it may name a column.</returns>
        private static bool IsLetter(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }

        /// <summary>是否为 ASCII 数字。Whether the character is an ASCII digit.</summary>
        /// <param name="c">待判定的字符。The character to test.</param>
        /// <returns>true 表示是数字。True for a digit.</returns>
        private static bool IsDigit(char c)
        {
            return c >= '0' && c <= '9';
        }
    }
}
