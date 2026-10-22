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
            Fx(wb, "-3^2", -9.0);
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
