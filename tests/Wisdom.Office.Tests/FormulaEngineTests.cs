using System;
using System.Globalization;
using Wisdom.Office.Core;
using Wisdom.Office.Core.Formulas;

namespace Wisdom.Office.Tests
{
    /// <summary>Excel 公式引擎（AST 词法/语法/求值/函数库/依赖重算）的回归测试。Regression tests for the Excel formula engine (lexer, parser, evaluator, function library and dependency recalculation).</summary>
    internal static partial class Program
    {
        private static void RunFormulaEngineTest()
        {
            OfficeWorkbook wb = BuildFormulaWorkbook();

            // 运算符与优先级。 Operators and precedence.
            Fx(wb, "1+2*3", 7.0);
            Fx(wb, "(1+2)*3", 9.0);
            Fx(wb, "2^3^2", 512.0);
            // Excel 把一元负号绑得比乘幂紧，-3^2 是 (-3)^2；数学惯例的 -9 不是 Excel 的结果。
            // Excel binds the unary sign tighter than power, so -3^2 is (-3)^2; the mathematical -9 is not Excel's answer.
            Fx(wb, "-3^2", 9.0);
            Fx(wb, "2^-3", 0.125);
            Fx(wb, "-3%+1", 0.97);
            Fx(wb, "50%", 0.5);
            Fx(wb, "\"a\"&\"b\"", "ab");
            Fx(wb, "10/0", "#DIV/0!");
            Fx(wb, "FOO(1)", "#NAME?");

            // 算术与数学函数。 Arithmetic and math functions.
            Fx(wb, "A1+A2", 30.0);
            Fx(wb, "ABS(-5)", 5.0);
            Fx(wb, "MOD(10,3)", 1.0);
            Fx(wb, "POWER(2,10)", 1024.0);
            Fx(wb, "SQRT(16)", 4.0);
            Fx(wb, "ROUND(2.5,0)", 3.0);
            Fx(wb, "INT(3.9)", 3.0);
            Fx(wb, "SUM(A1:A3)", 60.0);

            // 取整与可选参数：num_digits 为负时取整到小数点左侧，缺省参数走 Excel 的默认值。
            // Rounding and optional arguments: a negative num_digits rounds left of the decimal point and omitted
            // arguments fall back to Excel's defaults.
            Fx(wb, "ROUND(12345,-3)", 12000.0);
            Fx(wb, "ROUNDUP(1499,-2)", 1500.0);
            Fx(wb, "ROUNDDOWN(1299,-2)", 1200.0);
            Fx(wb, "TRUNC(12345.9,-3)", 12000.0);
            Fx(wb, "TRUNC(-1.9)", -1.0);
            Fx(wb, "ROUND(2.1234,1.9)", 2.1);
            Fx(wb, "CEILING(5)", 5.0);
            Fx(wb, "CEILING(5,3)", 6.0);
            Fx(wb, "FLOOR(5)", 5.0);
            Fx(wb, "LOG(100)", 2.0);
            Fx(wb, "ROUND(1,-400)", "#NUM!");

            // 单元格引用坐标：单格引用求值后仍带坐标，ROW/COLUMN/OFFSET 才认得出它是引用。
            // Single-cell references keep their coordinate after evaluation so ROW/COLUMN/OFFSET can see a reference.
            Fx(wb, "ROW(B2)", 2.0);
            Fx(wb, "COLUMN(B2)", 2.0);
            Fx(wb, "ROW(A1:A3)", 1.0);
            // 绝对引用 $A1 / $A$1：旧词法只读到列字母，整条公式落到名称分支报 #NAME?。
            // Absolute references $A1 / $A$1: the old scan stopped at the column letters, so the whole formula
            // fell through to the name branch and returned #NAME?.
            Fx(wb, "$A$1+$B1", 12.0);
            Fx(wb, "SUM($A$1:$A$2)", 30.0);
            Fx(wb, "OFFSET(A1,2,0)", 30.0);
            Fx(wb, "SUM(OFFSET(A1,0,0,3,1))", 60.0);

            // 整列/整行引用按已用范围收敛。 Whole-column and whole-row references clamp to the used range.
            Fx(wb, "SUM(A:A)", 60.0);
            Fx(wb, "SUM(1:1)", 17.0);

            // RANDBETWEEN 只给闭区间整数（早先返回带小数的值）。 RANDBETWEEN must return integers in the closed interval.
            double rb = Convert.ToDouble(FormulaEngine.Evaluate(wb, "RANDBETWEEN(1,3)"), CultureInfo.InvariantCulture);
            Require(rb >= 1.0 && rb <= 3.0 && rb == Math.Floor(rb), "RANDBETWEEN(1,3) returned " + rb);

            // 统计聚合。 Statistical aggregates.
            Fx(wb, "AVERAGE(A1:A3)", 20.0);
            Fx(wb, "MAX(A1:A3)", 30.0);
            Fx(wb, "MIN(A1:A3)", 10.0);
            Fx(wb, "COUNT(A1:A3)", 3.0);
            Fx(wb, "COUNTA(A1:B3)", 6.0);
            Fx(wb, "MEDIAN(A1:A3)", 20.0);
            Fx(wb, "STDEV(A1:A3)", 10.0);

            // 文本函数。 Text functions.
            Fx(wb, "UPPER(\"hello\")", "HELLO");
            Fx(wb, "LEFT(\"excel\",2)", "ex");
            Fx(wb, "MID(\"excel\",2,2)", "xc");
            Fx(wb, "LEN(\"hello\")", 5.0);
            Fx(wb, "SUBSTITUTE(\"aaa\",\"a\",\"b\",2)", "aba");
            Fx(wb, "TRIM(\"  a  b  \")", "a b");

            // TEXTJOIN 的分隔符在前、ignore_empty 在后；FIXED 的两个可选参数生效。
            // TEXTJOIN takes the delimiter before ignore_empty; both optional FIXED arguments now take effect.
            Fx(wb, "TEXTJOIN(\",\",TRUE,A1:A2)", "10,20");
            Fx(wb, "TEXTJOIN(\"-\",FALSE,A1:A2,\"x\")", "10-20-x");
            Fx(wb, "LEFT(\"excel\")", "e");
            Fx(wb, "FIXED(1234.5)", "1,234.50");
            Fx(wb, "FIXED(1234.5,0)", "1,235");
            Fx(wb, "FIXED(1234.5,0,FALSE)", "1235");
            Fx(wb, "FIXED(12345.678,-2)", "12,300");

            // 错误字面量整体成词：旧词法在 "/" 处断开，=#DIV/0! 直接抛格式异常。
            // Error literals lex as one run: the old tokenizer broke at "/", so =#DIV/0! threw a FormatException.
            Fx(wb, "#DIV/0!", "#DIV/0!");
            Fx(wb, "#N/A", "#N/A");
            Fx(wb, "ISERROR(#VALUE!)", true);

            // 条件聚合。 Conditional aggregation.
            Fx(wb, "SUMIF(A1:A3,\">=20\")", 50.0);
            Fx(wb, "COUNTIF(A1:A3,\">15\")", 2.0);

            // 逻辑与判断。 Logic and information.
            Fx(wb, "IF(A1>5,\"big\",\"small\")", "big");
            Fx(wb, "IFERROR(10/0,\"err\")", "err");
            Fx(wb, "AND(1,0)", false);
            Fx(wb, "OR(1,0)", true);
            Fx(wb, "NOT(0)", true);
            Fx(wb, "ISNUMBER(A1)", true);
            Fx(wb, "ISTEXT(B2)", true);

            // 跨表与查找。 Cross-sheet and lookup.
            Fx(wb, "Data!A1+Data!A2", 300.0);
            Fx(wb, "SUM('Data'!A1:A2)", 300.0);
            // 带空格的表名必须整体成词：旧词法没吃掉引号后的 "!"，'Sales Q1'!A1 退化成一个孤立表名。
            // A quoted sheet name with a space had to lex as a whole reference: the old tokenizer never consumed
            // the "!" after the closing quote, so 'Sales Q1'!A1 collapsed to a bare sheet name.
            Fx(wb, "'Sales Q1'!A1*2", 14.0);
            Fx(wb, "SUM(Data!A1:A2)", 300.0);
            Fx(wb, "VLOOKUP(2,Lookup!A1:B3,2,FALSE)", "two");
            Fx(wb, "MATCH(3,Lookup!A1:A3,0)", 3.0);
            Fx(wb, "INDEX(Lookup!A1:B3,2,2)", "two");

            // 日期函数。 Date functions.
            Fx(wb, "YEAR(44197)", 2021.0);

            // 依赖重算：D1=A1+A2 → D2=D1*2 → D3=SUM(D1:D2)。 Dependency recalculation.
            Worksheet s = wb.FindWorksheet("Sheet1");
            s.GetCell("D1").Formula = "A1+A2";
            s.GetCell("D2").Formula = "D1*2";
            s.GetCell("D3").Formula = "SUM(D1:D2)";
            int computed = FormulaEngine.Calculate(wb);
            Require(computed >= 3, "formula recalc computed too few cells: " + computed);
            FxVal("recalc D1", s.GetCell("D1").Value, 30.0);
            FxVal("recalc D2", s.GetCell("D2").Value, 60.0);
            FxVal("recalc D3", s.GetCell("D3").Value, 90.0);

            // 循环引用检测：以错误标记而非死循环。 Circular-reference detection: flagged, not a hang.
            Worksheet cyc = wb.AddWorksheet("Cyc");
            cyc.GetCell("A1").Formula = "B1+1";
            cyc.GetCell("B1").Formula = "A1+1";
            FormulaEngine.Calculate(wb);
            string flagged = Convert.ToString(cyc.GetCell("A1").Value, CultureInfo.InvariantCulture);
            Require(flagged != null && flagged.IndexOf('#') >= 0, "circular reference not flagged: " + flagged);

            Console.WriteLine("[formula] functions=" + FormulaEngine.SupportedFunctionCount + " engine regression passed");
        }

        private static OfficeWorkbook BuildFormulaWorkbook()
        {
            OfficeWorkbook wb = new OfficeWorkbook();
            Worksheet s = wb.AddWorksheet("Sheet1");
            s.GetCell("A1").Value = 10.0;
            s.GetCell("A2").Value = 20.0;
            s.GetCell("A3").Value = 30.0;
            s.GetCell("B1").Value = 2.0;
            s.GetCell("B2").Value = "apple";
            s.GetCell("B3").Value = true;
            s.GetCell("C1").Value = 5.0;

            Worksheet data = wb.AddWorksheet("Data");
            data.GetCell("A1").Value = 100.0;
            data.GetCell("A2").Value = 200.0;

            Worksheet lk = wb.AddWorksheet("Lookup");
            lk.GetCell("A1").Value = 1.0; lk.GetCell("B1").Value = "one";
            lk.GetCell("A2").Value = 2.0; lk.GetCell("B2").Value = "two";
            lk.GetCell("A3").Value = 3.0; lk.GetCell("B3").Value = "three";

            Worksheet spaced = wb.AddWorksheet("Sales Q1");
            spaced.GetCell("A1").Value = 7.0;
            return wb;
        }

        private static void Fx(OfficeWorkbook wb, string formula, double expected)
        {
            object got = FormulaEngine.Evaluate(wb, formula);
            double value;
            if (got is bool) value = (bool)got ? 1 : 0;
            else if (!double.TryParse(Convert.ToString(got, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out value))
            {
                Require(false, formula + " expected " + expected + " got " + got);
                return;
            }
            Require(Math.Abs(value - expected) < 1e-6, formula + " expected " + expected + " got " + got);
        }

        private static void Fx(OfficeWorkbook wb, string formula, string expected)
        {
            object got = FormulaEngine.Evaluate(wb, formula);
            string text = got == null ? string.Empty : Convert.ToString(got, CultureInfo.InvariantCulture);
            Require(text == expected, formula + " expected '" + expected + "' got '" + text + "'");
        }

        private static void Fx(OfficeWorkbook wb, string formula, bool expected)
        {
            object got = FormulaEngine.Evaluate(wb, formula);
            Require(got is bool && (bool)got == expected, formula + " expected " + expected + " got " + got);
        }

        private static void FxVal(string label, object got, double expected)
        {
            double value;
            bool ok = got is double && Math.Abs((double)got - expected) < 1e-6
                || double.TryParse(Convert.ToString(got, CultureInfo.InvariantCulture), NumberStyles.Float, CultureInfo.InvariantCulture, out value) && Math.Abs(value - expected) < 1e-6;
            Require(ok, label + " expected " + expected + " got " + got);
        }
    }
}
