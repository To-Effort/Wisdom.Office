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

        /// <summary>把第 index 个参数强制成数值；缺参或不可转换都返回 false。Coerces the argument at <paramref name="index"/> to a number; a missing or unconvertible argument returns false.</summary>
        /// <remarks>
        /// 缺参不能当成 0：早先返回 true 加 0，等于所有 <c>NumOr(a, i, 默认值)</c> 的默认值都被 0 顶掉，
        /// <c>CEILING(5)</c> 因此变成 0（Excel 给 #VALUE!，可选参数也该走各自的默认值）。
        /// A missing argument must not read as zero: reporting success with 0 overwrote every <c>NumOr</c>
        /// default, so CEILING(5) became 0 instead of falling back to the significance default.
        /// </remarks>
        private static bool NumberAt(List<FxValue> args, int index, out double d)
        {
            d = 0;
            if (index >= args.Count) return false;
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
            r["FIXED"] = (e, s, a) => FixedText(a);
            r["EVEN"] = (e, s, a) => { if (!NumArg(a, 0, out double d)) return Err(FxError.Value); return FxValue.Num(RoundAwayEven(d)); };
            r["ODD"] = (e, s, a) => { if (!NumArg(a, 0, out double d)) return Err(FxError.Value); return FxValue.Num(RoundAwayOdd(d)); };
            r["TRUNC"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return RoundLike(x, a, RoundTowardZero); };
            r["ROUND"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return RoundLike(x, a, RoundHalfAway); };
            r["ROUNDUP"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return RoundLike(x, a, RoundAwayFromZero); };
            r["ROUNDDOWN"] = (e, s, a) => { if (!NumArg(a, 0, out double x)) return Err(FxError.Value); return RoundLike(x, a, RoundTowardZero); };
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
            r["RAND"] = (e, s, a) => FxValue.Num(Rand.Shared.NextDouble());
            r["RANDBETWEEN"] = RandBetween;
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

        private const int RoundHalfAway = 0;
        private const int RoundAwayFromZero = 1;
        private const int RoundTowardZero = 2;

        /// <summary>
        /// ROUND / ROUNDUP / ROUNDDOWN / TRUNC 共用的缩放取整。num_digits 向零截断，为负时把取整点移到小数点
        /// 左侧（Excel 里 ROUND(12345,-3)=12000；早先的 Pow10 取了 10 的绝对次幂，负数位等于什么都不做）。
        /// 缩放因子溢出或归零时按 Excel 返回 #NUM!，参数缺失则走默认的 0 位。
        /// Shared scaled rounding for ROUND / ROUNDUP / ROUNDDOWN / TRUNC. num_digits is truncated toward zero and a
        /// negative value moves the rounding point to the left of the decimal point (Excel: ROUND(12345,-3)=12000;
        /// the old Pow10 used 10^abs(digits), so a negative count rounded nothing). An overflowing or zero factor
        /// yields #NUM! like Excel, and an omitted num_digits defaults to 0.
        /// </summary>
        private static FxValue RoundLike(double x, List<FxValue> a, int mode)
        {
            double factor = 1.0;
            if (a.Count > 1)
            {
                double digits;
                if (!NumberAt(a, 1, out digits)) return FirstError(a) ?? Err(FxError.Value);
                double whole = digits >= 0 ? Math.Floor(digits) : Math.Ceiling(digits);
                if (whole < -300 || whole > 300) return Err(FxError.Num);
                factor = Math.Pow(10.0, whole);
                if (factor <= 0 || double.IsNaN(factor) || double.IsInfinity(factor)) return Err(FxError.Num);
            }
            double scaled = x * factor;
            if (double.IsNaN(scaled) || double.IsInfinity(scaled)) return Err(FxError.Num);
            double rounded;
            if (mode == RoundHalfAway) rounded = scaled >= 0 ? Math.Floor(scaled + 0.5) : Math.Ceiling(scaled - 0.5);
            else if (mode == RoundAwayFromZero) rounded = x >= 0 ? Math.Ceiling(scaled) : Math.Floor(scaled);
            else rounded = x >= 0 ? Math.Floor(scaled) : Math.Ceiling(scaled);
            double result = rounded / factor;
            return double.IsNaN(result) || double.IsInfinity(result) ? Err(FxError.Num) : FxValue.Num(result);
        }

        /// <summary>
        /// FIXED(number, [decimals], [use_commas])：小数点固定为 "."、千分位固定为 ","，不随区域设置变化。
        /// decimals 四舍五入为整数、缺省 2，为负时先取整到小数点左侧再显示零位；use_commas 缺省 TRUE。
        /// 早先实现写死 "N2" 并丢掉后两个参数，FIXED(1234.5,0) 也返回 "1,234.50"。
        /// FIXED(number, [decimals], [use_commas]) always uses "." as the decimal point and "," as the thousands
        /// separator, independent of locale. decimals is rounded to an integer, defaults to 2, and a negative value
        /// rounds left of the decimal point while displaying zero decimals; use_commas defaults to TRUE.
        /// The old implementation hard-coded "N2" and dropped both optional arguments, so even FIXED(1234.5,0)
        /// returned "1,234.50".
        /// </summary>
        private static FxValue FixedText(List<FxValue> a)
        {
            double x;
            if (!NumberAt(a, 0, out x)) return FirstError(a) ?? Err(FxError.Value);
            int decimals = 2;
            if (a.Count > 1)
            {
                double d;
                if (!NumberAt(a, 1, out d)) return FirstError(a) ?? Err(FxError.Value);
                decimals = (int)(d >= 0 ? Math.Floor(d + 0.5) : Math.Ceiling(d - 0.5));
                if (decimals > 127) return Err(FxError.Value);
            }
            bool comma = a.Count > 2 ? BoolAt(a, 2) : true;
            double factor = decimals >= 0 ? Math.Pow(10.0, Math.Min(decimals, 15)) : Math.Pow(10.0, decimals);
            if (factor <= 0 || double.IsInfinity(factor)) return Err(FxError.Value);
            double scaled = x * factor;
            if (double.IsNaN(scaled) || double.IsInfinity(scaled)) return Err(FxError.Value);
            double rounded = (scaled >= 0 ? Math.Floor(scaled + 0.5) : Math.Ceiling(scaled - 0.5)) / factor;
            string text = rounded.ToString((comma ? "N" : "F") + Math.Max(decimals, 0), CultureInfo.InvariantCulture);
            return FxValue.Txt(text);
        }

        private static FxValue RandBetween(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double lo, hi;
            if (!NumArg(a, 0, out lo) || !NumArg(a, 1, out hi)) return FirstError(a) ?? Err(FxError.Value);
            lo = Math.Ceiling(lo);
            hi = Math.Floor(hi);
            if (lo > hi) return Err(FxError.Value);
            double count = hi - lo + 1.0;
            // 闭区间整数：先取 floor，再夹到 hi，避免随机数上界取到 1.0 时越界。
            // Inclusive integers: floor first, then clamp to hi so a random draw at the upper edge cannot overshoot.
            double pick = lo + Math.Floor(Rand.Shared.NextDouble() * count);
            if (pick > hi) pick = hi;
            if (pick < lo) pick = lo;
            return FxValue.Num(pick);
        }

        /// <summary>
        /// 每线程一个 Random 实例。RAND/RANDBETWEEN 原先每次调用都 new Random()，同一天钟刻度内的连续调用会拿到
        /// 相同种子、返回同一个"随机数"；.NET Framework 4.6.1 的 Random 也不是线程安全的。
        /// One Random per thread. RAND/RANDBETWEEN used to construct a Random per call, so consecutive calls inside
        /// the same clock tick shared a seed and returned the same "random" value; Random is also not thread-safe on
        /// .NET Framework 4.6.1.
        /// </summary>
        private static class Rand
        {
            [ThreadStatic]
            private static Random _shared;

            public static Random Shared
            {
                get
                {
                    Random r = _shared;
                    if (r == null)
                    {
                        r = new Random(unchecked(Environment.TickCount * 31) ^ System.Threading.Thread.CurrentThread.ManagedThreadId);
                        _shared = r;
                    }
                    return r;
                }
            }
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
