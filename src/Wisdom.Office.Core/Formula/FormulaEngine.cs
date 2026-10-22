using System;
using System.Collections.Generic;
using System.Globalization;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>
    /// 公式引擎对外的统一入口：解析、单式求值与整簿重算。
    /// The public entry point of the formula engine: parsing, single-formula evaluation and whole-workbook recalculation.
    /// </summary>
    /// <remarks>
    /// 重算只把结果写回 <see cref="Cell.Value"/>，保留 <see cref="Cell.Formula"/> 文本，因此默认不调用即可维持既有写出字节稳定；
    /// 需要缓存值时由调用方显式触发。Recalculation writes only into <see cref="Cell.Value"/> and keeps <see cref="Cell.Formula"/>,
    /// so it is opt-in and never changes existing write behaviour unless a caller invokes it explicitly.
    /// </remarks>
    public static class FormulaEngine
    {
        /// <summary>已支持的内置函数数量。The number of supported built-in functions.</summary>
        public static int SupportedFunctionCount { get { return Functions.Count; } }

        /// <summary>用工作簿作为求值上下文，计算一条独立公式（相对第一个工作表）。Evaluates a standalone formula using the workbook as context (relative to the first sheet).</summary>
        /// <param name="workbook">提供单元格值的工作簿。The workbook supplying cell values.</param>
        /// <param name="formula">公式文本，可带前导 =。The formula text, optionally with a leading =.</param>
        /// <returns>结果对象（double/string/bool），错误时返回错误文本字符串。The result object (double/string/bool); error text for errors.</returns>
        public static object Evaluate(OfficeWorkbook workbook, string formula)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            Worksheet sheet = workbook.FirstWorksheet;
            return Evaluate(workbook, sheet == null ? null : sheet.Name, formula);
        }

        /// <summary>在指定工作表上下文下计算一条公式。Evaluates a formula in the context of a named sheet.</summary>
        /// <param name="workbook">工作簿。The workbook.</param>
        /// <param name="sheetName">公式所在工作表名。The sheet hosting the formula.</param>
        /// <param name="formula">公式文本。The formula text.</param>
        /// <returns>结果对象。The result object.</returns>
        public static object Evaluate(OfficeWorkbook workbook, string sheetName, string formula)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            Worksheet sheet = workbook.FindWorksheet(sheetName) ?? workbook.FirstWorksheet;
            Evaluator ev = new Evaluator(workbook, null);
            FxValue result = ev.Evaluate(formula, sheet);
            return result.ToCellValue();
        }

        /// <summary>重算整个工作簿的所有公式，把结果写回单元格值。Recalculates every formula in the workbook and writes the results back into cell values.</summary>
        /// <param name="workbook">待重算的工作簿。The workbook to recalculate.</param>
        /// <returns>被计算的公式单元格数量。The number of formula cells computed.</returns>
        public static int Calculate(OfficeWorkbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            Dictionary<string, FxValue> cache = new Dictionary<string, FxValue>(StringComparer.OrdinalIgnoreCase);
            Evaluator ev = new Evaluator(workbook, cache);

            List<Cell> formulaCells = new List<Cell>();
            List<Worksheet> owners = new List<Worksheet>();
            for (int w = 0; w < workbook.Worksheets.Count; w++)
            {
                Worksheet sheet = workbook.Worksheets[w];
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        formulaCells.Add(cell);
                        owners.Add(sheet);
                    }
                }
            }

            int computed = 0;
            for (int i = 0; i < formulaCells.Count; i++)
            {
                Cell cell = formulaCells[i];
                Worksheet sheet = owners[i];
                FxValue value = ev.CellValue(sheet, cell.Row, cell.Column);
                cell.Value = value.ToCellValue();
                computed++;
            }
            return computed;
        }

        /// <summary>仅重算指定工作表的公式。Recalculates formulas on a single worksheet only.</summary>
        /// <param name="workbook">工作簿。The workbook.</param>
        /// <param name="sheetName">工作表名。The sheet name.</param>
        /// <returns>被计算的公式单元格数量。The number of formula cells computed.</returns>
        public static int Calculate(OfficeWorkbook workbook, string sheetName)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            Dictionary<string, FxValue> cache = new Dictionary<string, FxValue>(StringComparer.OrdinalIgnoreCase);
            Evaluator ev = new Evaluator(workbook, cache);
            int computed = 0;
            for (int w = 0; w < workbook.Worksheets.Count; w++)
            {
                Worksheet sheet = workbook.Worksheets[w];
                if (!string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase)) continue;
                List<Cell> formulaCells = new List<Cell>();
                foreach (Cell cell in sheet.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.Formula)) formulaCells.Add(cell);
                }
                foreach (Cell cell in formulaCells)
                {
                    FxValue value = ev.CellValue(sheet, cell.Row, cell.Column);
                    cell.Value = value.ToCellValue();
                    computed++;
                }
            }
            return computed;
        }

        /// <summary>解析公式为可读的规范化文本（调试用）。Parses a formula into normalized readable text (for debugging).</summary>
        /// <param name="formula">公式文本。The formula text.</param>
        /// <param name="tokenCount">解析成功时输出的令牌数。On success, the number of tokens produced.</param>
        /// <returns>解析成功返回 true。Returns true when parsing succeeds.</returns>
        public static bool TryParse(string formula, out int tokenCount)
        {
            tokenCount = 0;
            try
            {
                List<Token> tokens = Tokenizer.Lex(formula);
                new Parser(tokens).ParseFormula();
                tokenCount = tokens.Count;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string Format(double value)
        {
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }
    }
}
