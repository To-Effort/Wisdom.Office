using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>聚合、条件求和与统计函数。Aggregation, conditional and statistical functions.</summary>
    internal static partial class Functions
    {
        private static bool CollectAllForAverage(List<FxValue> a, out List<double> nums, out FxValue error)
        {
            nums = new List<double>();
            error = null;
            foreach (FxValue v in Evaluator.Flatten(a))
            {
                if (v.IsError) { error = v; return false; }
                if (v.Kind == FxKind.Number) nums.Add(v.Number);
                else if (v.Kind == FxKind.Logical) nums.Add(v.Bool ? 1 : 0);
                else if (v.Kind == FxKind.Text) nums.Add(0);
            }
            return true;
        }

        private static FxValue CountA(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            int count = 0;
            foreach (FxValue v in Evaluator.Flatten(a))
            {
                if (v.Kind != FxKind.Empty) count++;
            }
            return FxValue.Num(count);
        }

        private static FxValue CountBlank(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            // COUNTBLANK 只在单区域时有意义。 COUNTBLANK is meaningful with a single range.
            if (a.Count == 0) return FxValue.Num(0);
            int count = 0;
            foreach (FxValue v in Evaluator.Flatten(new List<FxValue> { Evaluator.Scalar(a[0]) }))
            {
                if (v.Kind == FxKind.Empty) count++;
                else if (v.Kind == FxKind.Text && v.Text == string.Empty) count++;
            }
            return FxValue.Num(count);
        }

        private static FxValue Median(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            List<double> n;
            FxValue err;
            if (!CollectNumbers(a, out n, out err)) return err;
            if (n.Count == 0) return Err(FxError.Num);
            n.Sort();
            int m = n.Count / 2;
            double v = n.Count % 2 == 1 ? n[m] : (n[m - 1] + n[m]) / 2.0;
            return FxValue.Num(v);
        }

        private static FxValue StDev(List<FxValue> a, bool sample)
        {
            List<double> n;
            FxValue err;
            if (!CollectNumbers(a, out n, out err)) return err;
            if (n.Count == 0) return Err(FxError.Div0);
            if (sample && n.Count < 2) return Err(FxError.Div0);
            double var = Variance(n, sample);
            return FxValue.Num(Math.Sqrt(var));
        }

        private static FxValue Variance(List<FxValue> a, bool sample)
        {
            List<double> n;
            FxValue err;
            if (!CollectNumbers(a, out n, out err)) return err;
            if (n.Count == 0) return Err(FxError.Div0);
            if (sample && n.Count < 2) return Err(FxError.Div0);
            return FxValue.Num(Variance(n, sample));
        }

        private static double Variance(List<double> n, bool sample)
        {
            double mean = 0;
            foreach (double v in n) mean += v;
            mean /= n.Count;
            double sum = 0;
            foreach (double v in n) sum += (v - mean) * (v - mean);
            double denom = sample ? (n.Count - 1) : n.Count;
            return sum / denom;
        }

        private static FxValue Rank(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double needle;
            if (!NumberAt(a, 0, out needle)) return FirstError(a) ?? Err(FxError.Value);
            if (a.Count < 2) return Err(FxError.Value);
            List<double> nums = new List<double>();
            foreach (FxValue v in RangeValues(a[1]))
            {
                double d;
                if (v.Kind == FxKind.Number) nums.Add(v.Number);
                else if (TryCoerceNumericText(v, out d)) nums.Add(d);
            }
            bool ascending = a.Count > 2 && NumOr(a, 2, 0) != 0;
            int rank = 1;
            foreach (double v in nums)
            {
                bool better = ascending ? v < needle : v > needle;
                if (better) rank++;
            }
            return FxValue.Num(rank);
        }

        private static FxValue SumProduct(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            FxValue e = FirstError(a);
            if (e != null) return e;
            if (a.Count == 0) return FxValue.Num(0);
            int size = -1;
            List<List<double>> series = new List<List<double>>();
            foreach (FxValue arg in a)
            {
                List<double> vals = new List<double>();
                foreach (FxValue v in RangeValues(arg))
                {
                    double d = 0;
                    Evaluator.TryCoerceNumber(Evaluator.Scalar(v), out d);
                    vals.Add(d);
                }
                if (size == -1) size = vals.Count;
                else if (vals.Count != size) return Err(FxError.Value);
                series.Add(vals);
            }
            double total = 0;
            for (int i = 0; i < size; i++)
            {
                double product = 1;
                foreach (List<double> s in series) product *= s[i];
                total += product;
            }
            return FxValue.Num(total);
        }

        private static FxValue SumIf(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 2) return Err(FxError.Value);
            FxValue range = a[0];
            FxValue criteria = Evaluator.Scalar(a[1]);
            List<FxValue> sumValues;
            if (a.Count >= 3)
            {
                sumValues = RangeValues(a[2]);
            }
            else
            {
                sumValues = RangeValues(range);
            }
            List<FxValue> testValues = RangeValues(range);
            double total = 0;
            int count = Math.Min(sumValues.Count, testValues.Count);
            for (int i = 0; i < count; i++)
            {
                if (CriteriaMatch(testValues[i], criteria))
                {
                    double d;
                    if (Evaluator.TryCoerceNumber(Evaluator.Scalar(sumValues[i]), out d) && sumValues[i].Kind == FxKind.Number)
                        total += d;
                }
            }
            return FxValue.Num(total);
        }

        private static FxValue CountIf(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 2) return Err(FxError.Value);
            FxValue range = a[0];
            FxValue criteria = Evaluator.Scalar(a[1]);
            int count = 0;
            foreach (FxValue v in RangeValues(range))
            {
                if (CriteriaMatch(v, criteria)) count++;
            }
            return FxValue.Num(count);
        }

        private static FxValue AverageIf(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 2) return Err(FxError.Value);
            FxValue range = a[0];
            FxValue criteria = Evaluator.Scalar(a[1]);
            List<FxValue> avgValues = a.Count >= 3 ? RangeValues(a[2]) : RangeValues(range);
            List<FxValue> testValues = RangeValues(range);
            double total = 0;
            int count = 0;
            int n = Math.Min(avgValues.Count, testValues.Count);
            for (int i = 0; i < n; i++)
            {
                if (CriteriaMatch(testValues[i], criteria))
                {
                    double d;
                    if (Evaluator.TryCoerceNumber(Evaluator.Scalar(avgValues[i]), out d)) { total += d; count++; }
                }
            }
            if (count == 0) return Err(FxError.Div0);
            return FxValue.Num(total / count);
        }

        private static FxValue ErrorType(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count == 0) return Err(FxError.Na);
            FxValue v = Evaluator.Scalar(a[0]);
            if (!v.IsError) return Err(FxError.Na);
            switch (v.Error)
            {
                case FxError.Null: return FxValue.Num(1);
                case FxError.Div0: return FxValue.Num(2);
                case FxError.Value: return FxValue.Num(3);
                case FxError.Ref: return FxValue.Num(4);
                case FxError.Name: return FxValue.Num(5);
                case FxError.Num: return FxValue.Num(6);
                case FxError.Na: return FxValue.Num(7);
                default: return FxValue.Num(8);
            }
        }

        // ---- 区域/值迭代助手。 Range/value iteration helpers. ----

        /// <summary>把值展开为可迭代的单元格序列（区域逐格，标量单元素）。Expands a value into a cell sequence: each cell for a range, a single item for a scalar.</summary>
        internal static List<FxValue> RangeValues(FxValue v)
        {
            List<FxValue> list = new List<FxValue>();
            if (v.Kind == FxKind.Range && v.Cells != null)
            {
                foreach (List<FxValue> row in v.Cells)
                    foreach (FxValue c in row)
                        list.Add(c);
            }
            else
            {
                list.Add(v);
            }
            return list;
        }

        private static bool TryCoerceNumericText(FxValue v, out double d)
        {
            d = 0;
            if (v.Kind != FxKind.Text) return false;
            return double.TryParse(v.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out d);
        }

        /// <summary>按 Excel 条件语义比较单元格值与条件，支持比较运算符与 * ? 通配。Compares a cell value with a criteria using Excel semantics, supporting comparison operators and * ? wildcards.</summary>
        internal static bool CriteriaMatch(FxValue cell, FxValue criteria)
        {
            cell = Evaluator.Scalar(cell);
            criteria = Evaluator.Scalar(criteria);

            string op = "=";
            string rhs = null;

            if (criteria.Kind == FxKind.Number || criteria.Kind == FxKind.Logical)
            {
                double cd = criteria.Kind == FxKind.Number ? criteria.Number : (criteria.Bool ? 1 : 0);
                return NumericEquals(cell, cd);
            }
            if (criteria.Kind == FxKind.Empty)
            {
                return cell.Kind == FxKind.Empty || (cell.Kind == FxKind.Text && cell.Text == string.Empty);
            }
            if (criteria.Kind != FxKind.Text) return false;

            string text = criteria.Text ?? string.Empty;
            if (text.Length >= 2 && (text[0] == '>' || text[0] == '<' || text[0] == '='))
            {
                if (text.StartsWith("<>")) { op = "<>"; rhs = text.Substring(2); }
                else if (text.StartsWith(">=")) { op = ">="; rhs = text.Substring(2); }
                else if (text.StartsWith("<=")) { op = "<="; rhs = text.Substring(2); }
                else if (text.StartsWith(">")) { op = ">"; rhs = text.Substring(1); }
                else if (text.StartsWith("<")) { op = "<"; rhs = text.Substring(1); }
                else { op = "="; rhs = text.Substring(1); }
            }
            else
            {
                op = "=";
                rhs = text;
            }

            double rhsNum;
            bool rhsIsNumber = double.TryParse(rhs, NumberStyles.Float, CultureInfo.InvariantCulture, out rhsNum);

            if (rhsIsNumber)
            {
                double cellNum;
                bool cellNumeric = false;
                if (cell.Kind == FxKind.Number) { cellNum = cell.Number; cellNumeric = true; }
                else if (cell.Kind == FxKind.Logical) { cellNum = cell.Bool ? 1 : 0; cellNumeric = true; }
                else if (TryCoerceNumericText(cell, out cellNum)) cellNumeric = true;

                if (op == "=") return cellNumeric && cellNum == rhsNum;
                if (op == "<>") return !(cellNumeric && cellNum == rhsNum);
                if (!cellNumeric) return false;
                return CompareOp(op, cellNum, rhsNum);
            }

            string cellText = cell.Kind == FxKind.Text ? cell.Text : (cell.Kind == FxKind.Number ? cell.Number.ToString(CultureInfo.InvariantCulture) : string.Empty);
            bool equals = WildcardEquals(cellText, rhs ?? string.Empty);
            switch (op)
            {
                case "=": return equals;
                case "<>": return !equals;
                default:
                    int cmp = string.Compare(cellText, rhs, StringComparison.OrdinalIgnoreCase);
                    return CompareOpStr(op, cmp);
            }
        }

        private static bool NumericEquals(FxValue cell, double cd)
        {
            cell = Evaluator.Scalar(cell);
            if (cell.Kind == FxKind.Number) return cell.Number == cd;
            double d;
            if (TryCoerceNumericText(cell, out d)) return d == cd;
            return false;
        }

        private static bool CompareOp(string op, double a, double b)
        {
            switch (op)
            {
                case ">": return a > b;
                case "<": return a < b;
                case ">=": return a >= b;
                case "<=": return a <= b;
                default: return a == b;
            }
        }

        private static bool CompareOpStr(string op, int cmp)
        {
            switch (op)
            {
                case ">": return cmp > 0;
                case "<": return cmp < 0;
                case ">=": return cmp >= 0;
                case "<=": return cmp <= 0;
                default: return cmp == 0;
            }
        }

        private static bool WildcardEquals(string value, string pattern)
        {
            if (pattern.IndexOf('*') < 0 && pattern.IndexOf('?') < 0)
            {
                return string.Equals(value, pattern, StringComparison.OrdinalIgnoreCase);
            }
            return MatchWildcard(value ?? string.Empty, pattern, 0, 0);
        }

        private static bool MatchWildcard(string s, string p, int si, int pi)
        {
            while (pi < p.Length)
            {
                char pc = p[pi];
                if (pc == '*')
                {
                    if (pi == p.Length - 1) return true;
                    for (int k = si; k <= s.Length; k++)
                    {
                        if (MatchWildcard(s, p, k, pi + 1)) return true;
                    }
                    return false;
                }
                if (si >= s.Length) return false;
                if (pc == '?' || char.ToUpperInvariant(pc) == char.ToUpperInvariant(s[si]))
                {
                    si++;
                    pi++;
                    continue;
                }
                return false;
            }
            return si == s.Length;
        }
    }
}
