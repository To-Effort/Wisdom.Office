using System;
using System.Collections.Generic;
using System.Globalization;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>
    /// 内置函数库与分派入口。The built-in function library and dispatch entry point.
    /// </summary>
    internal static partial class Functions
    {
        private delegate FxValue Fn(Evaluator ev, Worksheet sheet, List<FxValue> args);

        private static readonly Dictionary<string, Fn> Registry = BuildRegistry();

        /// <summary>已注册函数数量。The number of registered functions.</summary>
        public static int Count { get { return Registry.Count; } }

        /// <summary>按名称调用函数。Invokes a function by name.</summary>
        /// <param name="name">大写函数名。The upper-cased function name.</param>
        /// <param name="args">已求值的参数。The evaluated arguments.</param>
        /// <param name="ev">求值器。The evaluator.</param>
        /// <param name="sheet">当前工作表。The current sheet.</param>
        /// <returns>函数结果。The function result.</returns>
        public static FxValue Invoke(string name, List<FxValue> args, Evaluator ev, Worksheet sheet)
        {
            Fn fn;
            if (!Registry.TryGetValue(name, out fn)) return FxValue.Err(FxError.Name);
            try
            {
                return fn(ev, sheet, args);
            }
            catch (Exception)
            {
                return FxValue.Err(FxError.Value);
            }
        }

        // ---- 参数强制转换助手。 Argument coercion helpers. ----

        private static bool NumberAt(List<FxValue> args, int index, out double d)
        {
            d = 0;
            if (index >= args.Count) { d = 0; return true; }
            FxValue v = Evaluator.Scalar(args[index]);
            if (v.IsError) { d = 0; return false; }
            return Evaluator.TryCoerceNumber(v, out d);
        }

        private static FxValue FirstError(List<FxValue> args)
        {
            foreach (FxValue a in args) if (a.IsError) return a;
            return null;
        }

        private static string TextAt(List<FxValue> args, int index)
        {
            if (index >= args.Count) return string.Empty;
            FxValue v = Evaluator.Scalar(args[index]);
            return Evaluator.ToDisplay(v);
        }

        private static bool BoolAt(List<FxValue> args, int index)
        {
            if (index >= args.Count) return false;
            FxValue v = Evaluator.Scalar(args[index]);
            if (v.Kind == FxKind.Logical) return v.Bool;
            if (v.Kind == FxKind.Number) return v.Number != 0;
            if (v.Kind == FxKind.Text)
            {
                if (string.Equals(v.Text, "TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                if (string.Equals(v.Text, "FALSE", StringComparison.OrdinalIgnoreCase)) return false;
            }
            return v.Kind != FxKind.Empty;
        }

        /// <summary>收集参数中所有数值（跳过文本与空，错误向上传播）。Collects all numeric values from the arguments, skipping text and empty; errors propagate.</summary>
        private static bool CollectNumbers(List<FxValue> args, out List<double> nums, out FxValue error)
        {
            nums = new List<double>();
            error = null;
            List<FxValue> flat = Evaluator.Flatten(args);
            foreach (FxValue v in flat)
            {
                if (v.IsError) { error = v; return false; }
                if (v.Kind == FxKind.Number) { nums.Add(v.Number); }
                else if (v.Kind == FxKind.Logical && IsDirect(args, v)) { nums.Add(v.Bool ? 1 : 0); }
            }
            return true;
        }

        // 直接传入的布尔/文本字面量会被计入；来自区域的按 Excel 规则忽略。 Direct boolean/text literals count; range-sourced ones are ignored per Excel.
        private static bool IsDirect(List<FxValue> args, FxValue v)
        {
            foreach (FxValue a in args)
            {
                if (ReferenceEquals(a, v)) return true;
            }
            return false;
        }

        private static Dictionary<string, Fn> BuildRegistry()
        {
            Dictionary<string, Fn> r = new Dictionary<string, Fn>(StringComparer.OrdinalIgnoreCase);

            // 数学。 Math.
            r["ABS"] = (e, s, a) => U(a, Math.Abs);
            r["SIGN"] = (e, s, a) => U(a, v => v > 0 ? 1 : (v < 0 ? -1 : 0));
            r["INT"] = (e, s, a) => U(a, Math.Floor);
            r["FIXED"] = (e, s, a) => { if (!NumArg(a, 0, out double d)) return Err(FxError.Value); return FxValue.Txt(d.ToString("N2", CultureInfo.InvariantCulture)); };
            r["EVEN"] = (e, s, a) => { if (!NumArg(a, 0, out double d)) return Err(FxError.Value); return FxValue.Num(RoundAwayEven(d)); };
            r["ODD"] = (e, s, a) => { if (!NumArg(a, 0, out double d)) return Err(FxError.Value); return FxValue.Num(RoundAwayOdd(d)); };
            r["TRUNC"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); double factor = Pow10(NumOr(a, 1, 0)); return FxValue.Num(TruncateTo(x, factor)); };
            r["ROUND"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); int digits = (int)NumOr(a, 1, 0); return FxValue.Num(RoundHalfUp(x, digits)); };
            r["ROUNDUP"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); int digits = (int)NumOr(a, 1, 0); return FxValue.Num(RoundMagnitude(x, digits, true)); };
            r["ROUNDDOWN"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); int digits = (int)NumOr(a, 1, 0); return FxValue.Num(RoundMagnitude(x, digits, false)); };
            r["CEILING"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); double sig = NumOr(a, 1, 1); if (sig == 0) return FxValue.Num(0); return FxValue.Num(Math.Ceiling(x / sig) * sig); };
            r["FLOOR"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); double sig = NumOr(a, 1, 1); if (sig == 0) return Err(FxError.Div0); return FxValue.Num(Math.Floor(x / sig) * sig); };
            r["MOD"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); if (!NumArg(a, 1, out double m)) return Err(FxError.Value); if (m == 0) return Err(FxError.Div0); return FxValue.Num(x - m * Math.Floor(x / m)); };
            r["POWER"] = (e, s, a) => { if (!NumArg(a, 0, out double x) || !NumArg(a, 1, out double y)) return Err(FxError.Value); double v = Math.Pow(x, y); return double.IsNaN(v) || double.IsInfinity(v) ? Err(FxError.Num) : FxValue.Num(v); };
            r["SQRT"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return x < 0 ? Err(FxError.Num) : FxValue.Num(Math.Sqrt(x)); };
            r["SQRTPI"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return FxValue.Num(Math.Sqrt(x * Math.PI)); };
            r["EXP"] = (e, s, a) => U(a, Math.Exp);
            r["LN"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return x <= 0 ? Err(FxError.Num) : FxValue.Num(Math.Log(x)); };
            r["LOG10"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return x <= 0 ? Err(FxError.Num) : FxValue.Num(Math.Log10(x)); };
            r["LOG"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); if (x <= 0) return Err(FxError.Num); double b = NumOr(a, 1, 10); return FxValue.Num(Math.Log(x) / Math.Log(b)); };
            r["PI"] = (e, s, a) => FxValue.Num(Math.PI);
            r["RADIANS"] = (e, s, a) => U(a, v => v * Math.PI / 180.0);
            r["DEGREES"] = (e, s, a) => U(a, v => v * 180.0 / Math.PI);
            r["SIN"] = (e, s, a) => U(a, Math.Sin);
            r["COS"] = (e, s, a) => U(a, Math.Cos);
            r["TAN"] = (e, s, a) => U(a, Math.Tan);
            r["ASIN"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); if (x < -1 || x > 1) return Err(FxError.Num); return FxValue.Num(Math.Asin(x)); };
            r["ACOS"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); if (x < -1 || x > 1) return Err(FxError.Num); return FxValue.Num(Math.Acos(x)); };
            r["ATAN"] = (e, s, a) => U(a, Math.Atan);
            r["ATAN2"] = (e, s, a) => { if (!NumArg(a, 0, out double x) || !NumArg(a, 1, out double y)) return Err(FxError.Value); return FxValue.Num(Math.Atan2(y, x)); };
            r["FACT"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); if (x < 0) return Err(FxError.Num); return FxValue.Num(Factorial(Math.Floor(x))); };
            r["FACTDOUBLE"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return FxValue.Num(DoubleFactorial(Math.Floor(x))); };
            r["COMBIN"] = (e, s, a) => { if (!NumArg(a, 0, out double n) || !NumArg(a, 1, out double k)) return Err(FxError.Value); return FxValue.Num(Combin(Math.Floor(n), Math.Floor(k))); };
            r["RAND"] = (e, s, a) => FxValue.Num(new Random().NextDouble());
            r["RANDBETWEEN"] = (e, s, a) => { if (!NumArg(a, 0, out double lo) || !NumArg(a, 1, out double hi)) return Err(FxError.Value); lo = Math.Ceiling(lo); hi = Math.Floor(hi); if (lo > hi) return Err(FxError.Value); return FxValue.Num(lo + new Random().NextDouble() * (hi - lo + 1 - 1e-9)); };
            r["PRODUCT"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectNumbers(a, out n, out err)) return err; double p = 1; foreach (double v in n) p *= v; return FxValue.Num(p); };
            r["SUM"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectNumbers(a, out n, out err)) return err; double t = 0; foreach (double v in n) t += v; return FxValue.Num(t); };
            r["SUMPRODUCT"] = SumProduct;
            r["SUMSQ"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectNumbers(a, out n, out err)) return err; double t = 0; foreach (double v in n) t += v * v; return FxValue.Num(t); };

            // 统计。 Statistics.
            r["COUNT"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectNumbers(a, out n, out err)) return FxValue.Num(0); return FxValue.Num(n.Count); };
            r["COUNTA"] = CountA;
            r["COUNTBLANK"] = CountBlank;
            r["AVERAGE"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectNumbers(a, out n, out err)) return err; if (n.Count == 0) return Err(FxError.Div0); double t = 0; foreach (double v in n) t += v; return FxValue.Num(t / n.Count); };
            r["AVERAGEA"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectAllForAverage(a, out n, out err)) return err; if (n.Count == 0) return Err(FxError.Div0); double t = 0; foreach (double v in n) t += v; return FxValue.Num(t / n.Count); };
            r["MAX"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectNumbers(a, out n, out err)) return err; double m = double.NegativeInfinity; foreach (double v in n) if (v > m) m = v; return FxValue.Num(n.Count == 0 ? 0 : m); };
            r["MIN"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectNumbers(a, out n, out err)) return err; double m = double.PositiveInfinity; foreach (double v in n) if (v < m) m = v; return FxValue.Num(n.Count == 0 ? 0 : m); };
            r["MAXA"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectAllForAverage(a, out n, out err)) return err; double m = double.NegativeInfinity; foreach (double v in n) if (v > m) m = v; return FxValue.Num(n.Count == 0 ? 0 : m); };
            r["MINA"] = (e, s, a) => { List<double> n; FxValue err; if (!CollectAllForAverage(a, out n, out err)) return err; double m = double.PositiveInfinity; foreach (double v in n) if (v < m) m = v; return FxValue.Num(n.Count == 0 ? 0 : m); };
            r["MEDIAN"] = Median;
            r["STDEV"] = (e, s, a) => StDev(a, true);
            r["STDEV.S"] = (e, s, a) => StDev(a, true);
            r["STDEVP"] = (e, s, a) => StDev(a, false);
            r["STDEV.P"] = (e, s, a) => StDev(a, false);
            r["VAR"] = (e, s, a) => Variance(a, true);
            r["VAR.S"] = (e, s, a) => Variance(a, true);
            r["VARP"] = (e, s, a) => Variance(a, false);
            r["VAR.P"] = (e, s, a) => Variance(a, false);
            r["SUMIF"] = SumIf;
            r["COUNTIF"] = CountIf;
            r["AVERAGEIF"] = AverageIf;
            r["RANK"] = Rank;
            r["RANK.EQ"] = Rank;

            // 逻辑。 Logical.
            r["TRUE"] = (e, s, a) => FxValue.Logical(true);
            r["FALSE"] = (e, s, a) => FxValue.Logical(false);
            r["NOT"] = (e, s, a) => { if (FirstError(a) != null) return FirstError(a); return FxValue.Logical(!BoolAt(a, 0)); };
            r["AND"] = AndOr;
            r["OR"] = OrFn;
            r["XOR"] = XorFn;
            r["IF"] = IfFn;
            r["IFERROR"] = IfError;
            r["IFNA"] = IfNa;
            r["IFS"] = IfsFn;
            r["SWITCH"] = SwitchFn;

            // 信息。 Information.
            r["ISBLANK"] = (e, s, a) => FxValue.Logical(a.Count > 0 && Evaluator.Scalar(a[0]).Kind == FxKind.Empty);
            r["ISNUMBER"] = (e, s, a) => FxValue.Logical(a.Count > 0 && Evaluator.Scalar(a[0]).Kind == FxKind.Number);
            r["ISTEXT"] = (e, s, a) => FxValue.Logical(a.Count > 0 && Evaluator.Scalar(a[0]).Kind == FxKind.Text);
            r["ISNONTEXT"] = (e, s, a) => FxValue.Logical(!(a.Count > 0 && Evaluator.Scalar(a[0]).Kind == FxKind.Text));
            r["ISLOGICAL"] = (e, s, a) => FxValue.Logical(a.Count > 0 && Evaluator.Scalar(a[0]).Kind == FxKind.Logical);
            r["ISERROR"] = (e, s, a) => FxValue.Logical(a.Count > 0 && Evaluator.Scalar(a[0]).IsError);
            r["ISERR"] = (e, s, a) =>
            {
                if (a.Count == 0) return FxValue.Logical(false);
                FxValue v = Evaluator.Scalar(a[0]);
                return FxValue.Logical(v.IsError && v.Error != FxError.Na);
            };
            r["ISNA"] = (e, s, a) =>
            {
                if (a.Count == 0) return FxValue.Logical(false);
                FxValue v = Evaluator.Scalar(a[0]);
                return FxValue.Logical(v.IsError && v.Error == FxError.Na);
            };
            r["N"] = (e, s, a) => { double d; return NumberAt(a, 0, out d) ? FxValue.Num(d) : FxValue.Num(0); };
            r["NA"] = (e, s, a) => Err(FxError.Na);
            r["ERROR.TYPE"] = ErrorType;

            RegisterExtra(r);
            return r;
        }

        private static FxValue U(List<FxValue> a, Func<double, double> f)
        {
            if (!NumberAt(a, 0, out double x)) return FirstError(a) ?? Err(FxError.Value);
            double v = f(x);
            return double.IsNaN(v) || double.IsInfinity(v) ? Err(FxError.Num) : FxValue.Num(v);
        }

        private static bool NumArg(List<FxValue> a, int i, out double d)
        {
            if (i < a.Count && a[i].IsError) { d = 0; return false; }
            return NumberAt(a, i, out d);
        }

        private static double NumOr(List<FxValue> a, int i, double fallback)
        {
            double d;
            return NumberAt(a, i, out d) ? d : fallback;
        }

        private static FxValue Err(FxError e)
        {
            return FxValue.Err(e);
        }

        private static double Pow10(int digits)
        {
            double f = 1;
            int n = digits < 0 ? -digits : digits;
            for (int i = 0; i < n; i++) f *= 10;
            return f;
        }

        private static double Pow10(double d)
        {
            return Pow10((int)d);
        }

        private static double TruncateTo(double x, double factor)
        {
            double scaled = x * factor;
            double truncated = scaled >= 0 ? Math.Floor(scaled) : Math.Ceiling(scaled);
            return truncated / factor;
        }

        private static double RoundHalfUp(double x, int digits)
        {
            double f = Pow10(digits);
            double scaled = x * f;
            double rounded = scaled >= 0 ? Math.Floor(scaled + 0.5) : Math.Ceiling(scaled - 0.5);
            return rounded / f;
        }

        private static double RoundMagnitude(double x, int digits, bool up)
        {
            double f = Pow10(digits);
            double scaled = x * f;
            double rounded = x >= 0
                ? (up ? Math.Ceiling(scaled) : Math.Floor(scaled))
                : (up ? Math.Floor(scaled) : Math.Ceiling(scaled));
            return rounded / f;
        }

        private static double RoundAwayEven(double x)
        {
            if (x == 0) return 0;
            double a = Math.Abs(x);
            double e = Math.Ceiling(a / 2.0) * 2.0;
            return x > 0 ? e : -e;
        }

        private static double RoundAwayOdd(double x)
        {
            if (x == 0) return 1;
            double a = Math.Ceiling(Math.Abs(x));
            if (a % 2 == 0) a += 1;
            return x > 0 ? a : -a;
        }

        private static double Factorial(double n)
        {
            double r = 1;
            for (double i = 2; i <= n; i++) r *= i;
            return r;
        }

        private static double DoubleFactorial(double n)
        {
            if (n <= 0) return 1;
            double r = 1;
            for (double i = n; i > 0; i -= 2) r *= i;
            return r;
        }

        private static double Combin(double n, double k)
        {
            if (k < 0 || n < 0 || k > n) return 0;
            double r = 1;
            double kk = Math.Min(k, n - k);
            for (double i = 1; i <= kk; i++)
            {
                r = r * (n - kk + i) / i;
            }
            return Math.Round(r, 6, MidpointRounding.AwayFromZero);
        }

        private static FxValue AndOr(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            FxValue e = FirstError(a);
            if (e != null) return e;
            List<FxValue> flat = Evaluator.Flatten(a);
            bool any = false;
            bool hasValue = false;
            foreach (FxValue v in flat)
            {
                if (v.Kind == FxKind.Empty) continue;
                hasValue = true;
                bool b = ToBool(v);
                if (!b) return FxValue.Logical(false);
                any = true;
            }
            return FxValue.Logical(hasValue ? any : false);
        }

        private static FxValue OrFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            FxValue e = FirstError(a);
            if (e != null) return e;
            foreach (FxValue v in Evaluator.Flatten(a))
            {
                if (v.Kind == FxKind.Empty) continue;
                if (ToBool(v)) return FxValue.Logical(true);
            }
            return FxValue.Logical(false);
        }

        private static FxValue XorFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            FxValue e = FirstError(a);
            if (e != null) return e;
            bool acc = false;
            foreach (FxValue v in Evaluator.Flatten(a))
            {
                if (v.Kind == FxKind.Empty) continue;
                if (ToBool(v)) acc = !acc;
            }
            return FxValue.Logical(acc);
        }

        private static bool ToBool(FxValue v)
        {
            if (v.Kind == FxKind.Logical) return v.Bool;
            if (v.Kind == FxKind.Number) return v.Number != 0;
            if (v.Kind == FxKind.Text)
            {
                if (string.Equals(v.Text, "TRUE", StringComparison.OrdinalIgnoreCase)) return true;
                if (string.Equals(v.Text, "FALSE", StringComparison.OrdinalIgnoreCase)) return false;
            }
            return v.Kind != FxKind.Empty;
        }

        // IF / IFERROR / IFNA / IFS / SWITCH 需要惰性求值分支，直接从 AST 之外的参数列表拿不到未求值分支，
        // 因此这些函数在 Evaluator 里特判；此处保留签名以兼容注册表查找。
        // IF/IFERROR/IFS/SWITCH need short-circuit evaluation of branches, which a pre-evaluated argument list cannot
        // provide, so they are special-cased inside the evaluator. The registrations above are placeholders never reached.
        private static FxValue IfFn(Evaluator ev, Worksheet sheet, List<FxValue> a) { return FxValue.Err(FxError.Value); }
        private static FxValue IfError(Evaluator ev, Worksheet sheet, List<FxValue> a) { return FxValue.Err(FxError.Value); }
        private static FxValue IfNa(Evaluator ev, Worksheet sheet, List<FxValue> a) { return FxValue.Err(FxError.Value); }
        private static FxValue IfsFn(Evaluator ev, Worksheet sheet, List<FxValue> a) { return FxValue.Err(FxError.Value); }
        private static FxValue SwitchFn(Evaluator ev, Worksheet sheet, List<FxValue> a) { return FxValue.Err(FxError.Value); }
    }
}
