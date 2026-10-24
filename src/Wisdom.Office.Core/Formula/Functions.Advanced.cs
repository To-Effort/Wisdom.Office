using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>文本、日期、查找与信息类函数。Text, date, lookup and reference functions.</summary>
    internal static partial class Functions
    {
        private static FxValue Arg(List<FxValue> a, int i)
        {
            return i < a.Count ? a[i] : FxValue.Empty;
        }

        private static void RegisterExtra(Dictionary<string, Fn> r)
        {
            // ---- 文本。 Text. ----
            r["CONCATENATE"] = (e, s, a) => JoinText(a, string.Empty, false);
            r["CONCAT"] = (e, s, a) => JoinText(a, string.Empty, false);
            r["TEXTJOIN"] = TextJoin;
            r["LEN"] = (e, s, a) => FxValue.Num(TextAt(a, 0).Length);
            r["LEFT"] = (e, s, a) => { string t = TextAt(a, 0); double n = NumOr(a, 1, 1); if (n < 0) return Err(FxError.Value); int c = Math.Min((int)n, t.Length); return FxValue.Txt(t.Substring(0, c)); };
            r["RIGHT"] = (e, s, a) => { string t = TextAt(a, 0); double n = NumOr(a, 1, 1); if (n < 0) return Err(FxError.Value); int c = Math.Min((int)n, t.Length); return FxValue.Txt(t.Substring(t.Length - c, c)); };
            r["MID"] = (e, s, a) => { string t = TextAt(a, 0); double st; double ln; if (!NumberAt(a, 1, out st) || !NumberAt(a, 2, out ln)) return Err(FxError.Value); if (st < 1 || ln < 0) return Err(FxError.Num); int start = (int)st - 1; if (start >= t.Length) return FxValue.Txt(string.Empty); int len = Math.Min((int)ln, t.Length - start); return FxValue.Txt(t.Substring(start, len)); };
            r["LOWER"] = (e, s, a) => FxValue.Txt(TextAt(a, 0).ToLowerInvariant());
            r["UPPER"] = (e, s, a) => FxValue.Txt(TextAt(a, 0).ToUpperInvariant());
            r["PROPER"] = (e, s, a) => FxValue.Txt(Proper(TextAt(a, 0)));
            r["TRIM"] = (e, s, a) => FxValue.Txt(Trim(TextAt(a, 0)));
            r["CLEAN"] = (e, s, a) => { StringBuilder sb = new StringBuilder(); foreach (char ch in TextAt(a, 0)) if (!char.IsControl(ch)) sb.Append(ch); return FxValue.Txt(sb.ToString()); };
            // 上限要落在产出长度而不是重复次数上：Excel 的单元格文本上限就是 32767，
            // 只按次数封顶会让「32767 次乘以任意长文本」仍然把内存要光。
            // The cap belongs on the produced length, not the repeat count: 32767 is Excel's cell-text
            // limit, and capping only the count still lets "32767 times an arbitrarily long string" ask for
            // all the memory in the machine.
            // 空文本另算：乘积是 0，按上面的判据过得去，却依旧要空转 n 次；
            // 没有内容可重复就立刻交出空串。次数随后夹在 32767 内，(long) 转换也不会回绕。
            // An empty text is a separate case: the product is 0, so the length test passes and the loop
            // still spins n times. With nothing to repeat the result is returned at once, and the count is
            // then bounded by 32767 so the (long) conversion cannot wrap either.
            r["REPT"] = (e, s, a) => { string t = TextAt(a, 0); double n = NumOr(a, 1, 0); if (n < 0) return Err(FxError.Value); if (n < 1 || t.Length == 0) return FxValue.Txt(string.Empty); if (n > 32767) return Err(FxError.Value); long times = (long)n; long total = times * (long)t.Length; if (total > 32767) return Err(FxError.Value); StringBuilder sb = new StringBuilder((int)total); for (long i = 0; i < times; i++) sb.Append(t); return FxValue.Txt(sb.ToString()); };
            r["REPLACE"] = ReplaceFn;
            r["SUBSTITUTE"] = SubstituteFn;
            r["FIND"] = (e, s, a) => FindFn(a, true);
            r["SEARCH"] = (e, s, a) => FindFn(a, false);
            r["EXACT"] = (e, s, a) => FxValue.Logical(string.Equals(TextAt(a, 0), TextAt(a, 1), StringComparison.Ordinal));
            r["VALUE"] = ValueFn;
            r["TEXT"] = TextFormatFn;

            // ---- 日期与时间。 Date and time. ----
            r["DATE"] = DateFn;
            r["YEAR"] = (e, s, a) => SerialPart(a, d => d.Year);
            r["MONTH"] = (e, s, a) => SerialPart(a, d => d.Month);
            r["DAY"] = (e, s, a) => SerialPart(a, d => d.Day);
            r["HOUR"] = (e, s, a) => SerialPart(a, d => d.Hour);
            r["MINUTE"] = (e, s, a) => SerialPart(a, d => d.Minute);
            r["SECOND"] = (e, s, a) => SerialPart(a, d => d.Second);
            r["WEEKDAY"] = WeekdayFn;
            r["NOW"] = (e, s, a) => FxValue.Num(DateUtils.DateTimeToExcelSerial(DateTime.Now, false));
            r["TODAY"] = (e, s, a) => FxValue.Num(DateUtils.DateTimeToExcelSerial(DateTime.Today, false));
            r["DATEVALUE"] = DateValueFn;
            r["TIME"] = TimeFn;
            r["TIMEVALUE"] = TimeValueFn;
            r["DAYS"] = (e, s, a) => { double x; double y; if (!NumberAt(a, 0, out x) || !NumberAt(a, 1, out y)) return Err(FxError.Value); return FxValue.Num(Math.Truncate(x - y)); };
            r["EDATE"] = EdateFn;
            r["EOMONTH"] = EomonthFn;
            r["DATEDIF"] = DateDifFn;

            // ---- 查找与引用。 Lookup and reference. ----
            r["VLOOKUP"] = Vlookup;
            r["HLOOKUP"] = Hlookup;
            r["MATCH"] = MatchFn;
            r["INDEX"] = IndexFn;
            r["CHOOSE"] = ChooseFn;
            r["LOOKUP"] = LookupFn;
            r["COLUMN"] = ColumnFn;
            r["COLUMNS"] = (e, s, a) => WidthOf(a);
            r["ROW"] = RowFn;
            r["ROWS"] = (e, s, a) => HeightOf(a);
            r["ADDRESS"] = AddressFn;
            r["OFFSET"] = OffsetFn;

            // ---- 财务（基础）。 Finance (basic). ----
            r["FV"] = FvFn;
            r["PV"] = PvFn;
            r["PMT"] = PmtFn;
            r["NPV"] = NpvFn;
        }

        // ---- 文本实现。 Text implementations. ----

        private static FxValue JoinText(List<FxValue> a, string sep, bool ignoreEmptyFirst)
        {
            FxValue e = FirstError(a);
            if (e != null) return e;
            StringBuilder sb = new StringBuilder();
            foreach (FxValue v in Evaluator.Flatten(a))
            {
                sb.Append(Evaluator.ToDisplay(Evaluator.Scalar(v)));
            }
            return FxValue.Txt(sb.ToString());
        }

        private static FxValue TextJoin(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 2) return Err(FxError.Value);
            bool ignoreEmpty = BoolAt(a, 0);
            string delim = TextAt(a, 1);
            StringBuilder sb = new StringBuilder();
            bool first = true;
            for (int i = 2; i < a.Count; i++)
            {
                foreach (FxValue v in RangeValues(a[i]))
                {
                    FxValue sv = Evaluator.Scalar(v);
                    if (sv.IsError) return sv;
                    string t = Evaluator.ToDisplay(sv);
                    if (ignoreEmpty && t.Length == 0) continue;
                    if (!first) sb.Append(delim);
                    sb.Append(t);
                    first = false;
                }
            }
            return FxValue.Txt(sb.ToString());
        }

        private static string Proper(string t)
        {
            StringBuilder sb = new StringBuilder(t.Length);
            bool newWord = true;
            foreach (char c in t)
            {
                if (char.IsLetter(c))
                {
                    sb.Append(newWord ? char.ToUpperInvariant(c) : char.ToLowerInvariant(c));
                    newWord = false;
                }
                else
                {
                    sb.Append(c);
                    newWord = true;
                }
            }
            return sb.ToString();
        }

        private static string Trim(string t)
        {
            string s = t.Trim();
            StringBuilder sb = new StringBuilder(s.Length);
            bool lastSpace = false;
            foreach (char c in s)
            {
                bool space = char.IsWhiteSpace(c);
                if (space && lastSpace) continue;
                sb.Append(space ? ' ' : c);
                lastSpace = space;
            }
            return sb.ToString();
        }

        private static FxValue ReplaceFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            string old = TextAt(a, 0);
            double startN, countN;
            if (!NumberAt(a, 1, out startN) || !NumberAt(a, 2, out countN)) return Err(FxError.Value);
            int start = (int)startN;
            int count = (int)countN;
            string newtext = TextAt(a, 3);
            if (start < 1) return Err(FxError.Value);
            int len = old.Length;
            int head = Math.Min(start - 1, len);
            int tailStart = Math.Min(start - 1 + Math.Max(count, 0), len);
            return FxValue.Txt(old.Substring(0, head) + newtext + old.Substring(tailStart, len - tailStart));
        }

        private static FxValue SubstituteFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            string text = TextAt(a, 0);
            string old = TextAt(a, 1);
            string repl = TextAt(a, 2);
            if (old.Length == 0) return FxValue.Txt(text);
            if (a.Count >= 4)
            {
                double nth = NumOr(a, 3, 0);
                if (nth <= 0) return Err(FxError.Value);
                int target = (int)nth;
                int idx = 0;
                int occurrence = 0;
                while ((idx = text.IndexOf(old, idx, StringComparison.Ordinal)) >= 0)
                {
                    occurrence++;
                    if (occurrence == target)
                    {
                        return FxValue.Txt(text.Substring(0, idx) + repl + text.Substring(idx + old.Length));
                    }
                    idx += old.Length;
                }
                return FxValue.Txt(text);
            }
            return FxValue.Txt(text.Replace(old, repl));
        }

        private static FxValue FindFn(List<FxValue> a, bool caseSensitive)
        {
            string needle = TextAt(a, 0);
            string hay = TextAt(a, 1);
            int start = 0;
            if (a.Count >= 3)
            {
                double s;
                if (!NumberAt(a, 2, out s)) return Err(FxError.Value);
                start = (int)s - 1;
                if (start < 0) start = 0;
            }
            if (needle.Length == 0) return FxValue.Num(start + 1);
            StringComparison cmp = caseSensitive ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            if (start > hay.Length) return Err(FxError.Value);
            int idx = hay.IndexOf(needle, start, cmp);
            return idx >= 0 ? FxValue.Num(idx + 1) : Err(FxError.Value);
        }

        private static FxValue ValueFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count == 0) return FxValue.Num(0);
            FxValue v = Evaluator.Scalar(a[0]);
            if (v.IsError) return v;
            double d;
            if (Evaluator.TryCoerceNumber(v, out d)) return FxValue.Num(d);
            return Err(FxError.Value);
        }

        private static FxValue TextFormatFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 2) return Err(FxError.Value);
            FxValue value = Evaluator.Scalar(a[0]);
            string fmt = TextAt(a, 1);
            double d;
            if (Evaluator.TryCoerceNumber(value, out d) && IsDateFormat(fmt))
            {
                try
                {
                    DateTime dt = DateUtils.ExcelSerialToDateTime(d, false);
                    return FxValue.Txt(SpreadsheetFormat.FormatDate(dt, fmt));
                }
                catch (ArgumentOutOfRangeException)
                {
                    return Err(FxError.Value);
                }
            }
            if (value.Kind == FxKind.Text) return FxValue.Txt(value.Text ?? string.Empty);
            return FxValue.Txt(SpreadsheetFormat.FormatNumber(d, fmt, "G15"));
        }

        private static bool IsDateFormat(string fmt)
        {
            CellStyle probe = new CellStyle();
            probe.NumberFormat = fmt;
            return probe.IsDateFormat();
        }

        // ---- 日期实现。 Date implementations. ----

        private static FxValue DateFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double y, m, d;
            if (!NumberAt(a, 0, out y) || !NumberAt(a, 1, out m) || !NumberAt(a, 2, out d)) return FirstError(a) ?? Err(FxError.Value);
            try
            {
                // 以 1900-01-01（序列 1）为基准累加年月日，超范围自动进位，符合 Excel 语义。
                // Anchor on 1900-01-01 (serial 1) and accumulate years/months/days so out-of-range parts roll over like Excel.
                DateTime result = new DateTime(1900, 1, 1)
                    .AddYears((int)y - 1900)
                    .AddMonths((int)m - 1)
                    .AddDays((int)d - 1);
                if (result.Year < 1900) return Err(FxError.Num);
                return FxValue.Num(DateUtils.DateTimeToExcelSerial(result, false));
            }
            catch (ArgumentOutOfRangeException)
            {
                return Err(FxError.Num);
            }
        }

        private static FxValue SerialPart(List<FxValue> a, Func<DateTime, int> part)
        {
            double serial;
            if (!NumberAt(a, 0, out serial)) return FirstError(a) ?? Err(FxError.Value);
            if (serial < 0) return Err(FxError.Num);
            DateTime dt = DateUtils.ExcelSerialToDateTime(serial, false);
            return FxValue.Num(part(dt));
        }

        private static FxValue WeekdayFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double serial;
            if (!NumberAt(a, 0, out serial)) return FirstError(a) ?? Err(FxError.Value);
            int type = (int)NumOr(a, 1, 1);
            DateTime dt = DateUtils.ExcelSerialToDateTime(serial, false);
            int dow = (int)dt.DayOfWeek; // Sunday=0
            switch (type)
            {
                case 1: return FxValue.Num(dow + 1);
                case 2: return FxValue.Num(dow == 0 ? 7 : dow);
                case 3: return FxValue.Num(dow == 0 ? 6 : dow - 1);
                default: return Err(FxError.Num);
            }
        }

        private static FxValue DateValueFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            string text = TextAt(a, 0);
            DateTime dt;
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return FxValue.Num(DateUtils.DateTimeToExcelSerial(dt.Date, false));
            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt))
                return FxValue.Num(DateUtils.DateTimeToExcelSerial(dt.Date, false));
            return Err(FxError.Value);
        }

        private static FxValue TimeFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double h, m, s;
            if (!NumberAt(a, 0, out h) || !NumberAt(a, 1, out m) || !NumberAt(a, 2, out s)) return FirstError(a) ?? Err(FxError.Value);
            double total = (h * 3600 + m * 60 + s) / 86400.0;
            return FxValue.Num(total - Math.Floor(total));
        }

        private static FxValue TimeValueFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            string text = TextAt(a, 0);
            DateTime dt;
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out dt))
                return FxValue.Num((dt.Hour * 3600 + dt.Minute * 60 + dt.Second) / 86400.0);
            return Err(FxError.Value);
        }

        private static FxValue EdateFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double serial;
            double months;
            if (!NumberAt(a, 0, out serial) || !NumberAt(a, 1, out months)) return FirstError(a) ?? Err(FxError.Value);
            DateTime dt = DateUtils.ExcelSerialToDateTime(serial, false);
            return FxValue.Num(DateUtils.DateTimeToExcelSerial(dt.AddMonths((int)months), false));
        }

        private static FxValue EomonthFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double serial;
            double months;
            if (!NumberAt(a, 0, out serial) || !NumberAt(a, 1, out months)) return FirstError(a) ?? Err(FxError.Value);
            DateTime dt = DateUtils.ExcelSerialToDateTime(serial, false).AddMonths((int)months);
            DateTime end = new DateTime(dt.Year, dt.Month, DateTime.DaysInMonth(dt.Year, dt.Month));
            return FxValue.Num(DateUtils.DateTimeToExcelSerial(end, false));
        }

        private static FxValue DateDifFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double start, end;
            if (!NumberAt(a, 0, out start) || !NumberAt(a, 1, out end)) return FirstError(a) ?? Err(FxError.Value);
            if (end < start) return Err(FxError.Num);
            string unit = TextAt(a, 2).ToUpperInvariant();
            DateTime s = DateUtils.ExcelSerialToDateTime(start, false);
            DateTime e = DateUtils.ExcelSerialToDateTime(end, false);
            switch (unit)
            {
                case "D": return FxValue.Num((e - s).TotalDays);
                case "M":
                    int months = (e.Year - s.Year) * 12 + e.Month - s.Month;
                    if (e.Day < s.Day) months--;
                    return FxValue.Num(months);
                case "Y":
                    int years = e.Year - s.Year;
                    if (e.DayOfYear < s.DayOfYear) years--;
                    return FxValue.Num(years);
                default: return Err(FxError.Num);
            }
        }

        // ---- 查找实现。 Lookup implementations. ----

        private static List<List<FxValue>> Grid(FxValue v)
        {
            if (v.Kind == FxKind.Range && v.Cells != null) return v.Cells;
            List<List<FxValue>> single = new List<List<FxValue>>();
            single.Add(new List<FxValue> { v });
            return single;
        }

        private static FxValue Vlookup(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 3) return Err(FxError.Value);
            FxValue lookup = Evaluator.Scalar(a[0]);
            List<List<FxValue>> table = Grid(a[1]);
            double colD;
            if (!NumberAt(a, 2, out colD)) return Err(FxError.Value);
            int col = (int)colD;
            if (col < 1) return Err(FxError.Value);
            bool approximate = a.Count >= 4 ? BoolAt(a, 3) : true;

            FxValue best = null;
            foreach (List<FxValue> row in table)
            {
                FxValue cell = row.Count > 0 ? Evaluator.Scalar(row[0]) : FxValue.Empty;
                int cmp;
                bool comparable = TryCompareValues(lookup, cell, out cmp);
                if (!comparable) continue;
                if (cmp == 0)
                {
                    if (col > row.Count) return Err(FxError.Ref);
                    return Evaluator.Scalar(row[col - 1]);
                }
                if (approximate && cmp > 0)
                {
                    best = row.Count >= col && col >= 1 ? row[col - 1] : best;
                }
            }
            if (approximate && best != null) return Evaluator.Scalar(best);
            return Err(FxError.Na);
        }

        private static FxValue Hlookup(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 3) return Err(FxError.Value);
            FxValue lookup = Evaluator.Scalar(a[0]);
            List<List<FxValue>> table = Grid(a[1]);
            double rowD;
            if (!NumberAt(a, 2, out rowD)) return Err(FxError.Value);
            int row = (int)rowD;
            if (table.Count == 0 || row < 1 || row > table.Count) return Err(FxError.Ref);
            bool approximate = a.Count >= 4 ? BoolAt(a, 3) : true;
            int headerCols = table[0].Count;
            int chosen = -1;
            for (int c = 0; c < headerCols; c++)
            {
                FxValue cell = Evaluator.Scalar(table[0][c]);
                int cmp;
                if (!TryCompareValues(lookup, cell, out cmp)) continue;
                if (cmp == 0) { chosen = c; break; }
                if (approximate && cmp > 0) chosen = c;
            }
            if (chosen < 0) return Err(FxError.Na);
            List<FxValue> targetRow = table[row - 1];
            if (chosen >= targetRow.Count) return Err(FxError.Ref);
            return Evaluator.Scalar(targetRow[chosen]);
        }

        private static FxValue MatchFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 2) return Err(FxError.Value);
            FxValue lookup = Evaluator.Scalar(a[0]);
            double typeD;
            int type = NumberAt(a, 2, out typeD) ? (int)typeD : 1;
            List<FxValue> values = new List<FxValue>();
            foreach (FxValue v in RangeValues(a[1])) values.Add(Evaluator.Scalar(v));

            if (type == 0)
            {
                for (int i = 0; i < values.Count; i++)
                {
                    int cmp;
                    if (TryCompareValues(lookup, values[i], out cmp) && cmp == 0) return FxValue.Num(i + 1);
                }
                return Err(FxError.Na);
            }
            int bestIdx = -1;
            for (int i = 0; i < values.Count; i++)
            {
                int cmp;
                if (!TryCompareValues(lookup, values[i], out cmp)) continue;
                if (cmp == 0) return FxValue.Num(i + 1);
                if (type == 1 && cmp > 0) bestIdx = i;
                if (type == -1 && cmp < 0) bestIdx = i;
            }
            if (bestIdx >= 0) return FxValue.Num(bestIdx + 1);
            return Err(FxError.Na);
        }

        private static FxValue IndexFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 2) return Err(FxError.Value);
            List<List<FxValue>> grid = Grid(a[0]);
            double rowD = 0;
            double colD = 0;
            NumberAt(a, 1, out rowD);
            if (a.Count >= 3) NumberAt(a, 2, out colD);
            int rows = grid.Count;
            int cols = rows > 0 ? grid[0].Count : 0;

            int row = (int)rowD;
            int col = (int)colD;
            if (rows == 1 || cols == 1)
            {
                // 一维区域：第二参数是序号。 Single dimension: the second argument is the ordinal.
                if (rows == 1 && cols > 1)
                {
                    if (colD != 0) col = (int)colD;
                    if (row <= 0 || row > cols) return Err(FxError.Ref);
                    return Evaluator.Scalar(grid[0][row - 1]);
                }
                if (cols == 1)
                {
                    if (row <= 0 || row > rows) return Err(FxError.Ref);
                    return Evaluator.Scalar(grid[row - 1][0]);
                }
            }
            if (row < 1 || row > rows) return Err(FxError.Ref);
            if (col == 0) col = 1;
            if (col < 1 || col > cols) return Err(FxError.Ref);
            return Evaluator.Scalar(grid[row - 1][col - 1]);
        }

        private static FxValue ChooseFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double n;
            if (!NumberAt(a, 0, out n)) return FirstError(a) ?? Err(FxError.Value);
            int idx = (int)n;
            if (idx < 1 || idx >= a.Count) return Err(FxError.Value);
            return Evaluator.Scalar(a[idx]);
        }

        private static FxValue LookupFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 2) return Err(FxError.Value);
            FxValue lookup = Evaluator.Scalar(a[0]);
            List<FxValue> vector = new List<FxValue>();
            foreach (FxValue v in RangeValues(a[1])) vector.Add(Evaluator.Scalar(v));
            List<FxValue> result = vector;
            if (a.Count >= 3)
            {
                result = new List<FxValue>();
                foreach (FxValue v in RangeValues(a[2])) result.Add(Evaluator.Scalar(v));
            }
            int best = -1;
            for (int i = 0; i < vector.Count && i < result.Count; i++)
            {
                int cmp;
                if (TryCompareValues(lookup, vector[i], out cmp) && cmp >= 0) best = i;
            }
            return best >= 0 ? result[best] : Err(FxError.Na);
        }

        private static FxValue ColumnFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count == 0 || a[0].Kind != FxKind.Range) return Err(FxError.Value);
            return FxValue.Num(a[0].C1);
        }

        private static FxValue RowFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count == 0 || a[0].Kind != FxKind.Range) return Err(FxError.Value);
            return FxValue.Num(a[0].R1);
        }

        private static FxValue WidthOf(List<FxValue> a)
        {
            if (a.Count == 0 || a[0].Kind != FxKind.Range) return FxValue.Num(1);
            return FxValue.Num(a[0].C2 - a[0].C1 + 1);
        }

        private static FxValue HeightOf(List<FxValue> a)
        {
            if (a.Count == 0 || a[0].Kind != FxKind.Range) return FxValue.Num(1);
            return FxValue.Num(a[0].R2 - a[0].R1 + 1);
        }

        private static FxValue AddressFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double rowN, colN;
            if (!NumberAt(a, 0, out rowN) || !NumberAt(a, 1, out colN)) return FirstError(a) ?? Err(FxError.Value);
            int absMode = (int)NumOr(a, 2, 1);
            bool a1 = a.Count < 4 || BoolAt(a, 3);
            string sheetName = a.Count >= 5 && a[4].Kind == FxKind.Text ? a[4].Text + "!" : string.Empty;
            string col = Parser.ColumnToLetters((int)colN);
            string row = ((int)rowN).ToString(CultureInfo.InvariantCulture);
            string colRef = (absMode == 2 || absMode == 4) ? col : "$" + col;
            string rowRef = (absMode == 3 || absMode == 4) ? row : "$" + row;
            if (!a1) return FxValue.Txt(sheetName + rowRef + colRef);
            return FxValue.Txt(sheetName + colRef + rowRef);
        }

        private static FxValue OffsetFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            if (a.Count < 3) return Err(FxError.Value);
            FxValue reference = a[0];
            if (reference.Sheet == null) return Err(FxError.Ref);
            double rowsD, colsD;
            if (!NumberAt(a, 1, out rowsD) || !NumberAt(a, 2, out colsD)) return FirstError(a) ?? Err(FxError.Value);
            int height = a.Count >= 4 ? (int)NumOr(a, 3, 1) : 1;
            int width = a.Count >= 5 ? (int)NumOr(a, 4, 1) : 1;
            int baseRow = reference.Kind == FxKind.Range ? reference.R1 : reference.R1;
            int baseCol = reference.Kind == FxKind.Range ? reference.C1 : reference.C1;
            int r1 = baseRow + (int)rowsD;
            int c1 = baseCol + (int)colsD;
            int r2 = r1 + height - 1;
            int c2 = c1 + width - 1;
            // 只校验边界等于放任尺寸：OFFSET(A1,0,0,1000000,1000) 的行列都还在网格内，
            // 却要求十亿次取值与逐行分配，必须按普通区域的 MaxRangeCells 一并封顶。
            // Checking only the edges lets the size run wild: OFFSET(A1,0,0,1000000,1000) stays inside the
            // grid yet asks for a billion cell reads, so the ordinary range ceiling applies here too.
            if (height < 1 || width < 1) return Err(FxError.Value);
            if (r1 < 1 || c1 < 1 || r2 > CellAddress.MaxRowLiteral || c2 > CellAddress.MaxColumnLiteral) return Err(FxError.Ref);
            if ((long)height * (long)width > Evaluator.MaxRangeCells) return Err(FxError.Ref);
            FxValue result = new FxValue();
            result.Kind = FxKind.Range;
            result.Sheet = reference.Sheet;
            result.R1 = r1;
            result.C1 = c1;
            result.R2 = r2;
            result.C2 = c2;
            result.Cells = new List<List<FxValue>>(height);
            for (int r = r1; r <= r2; r++)
            {
                List<FxValue> rowVals = new List<FxValue>(width);
                for (int c = c1; c <= c2; c++) rowVals.Add(ev.CellValue(reference.Sheet, r, c));
                result.Cells.Add(rowVals);
            }
            if (height == 1 && width == 1) return result.Cells[0][0];
            return result;
        }

        // ---- 财务实现。 Finance implementations. ----

        private static FxValue FvFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double rate, nper, pmt, pv;
            if (!NumberAt(a, 0, out rate) || !NumberAt(a, 1, out nper) || !NumberAt(a, 2, out pmt)) return FirstError(a) ?? Err(FxError.Value);
            pv = NumOr(a, 3, 0);
            int type = (int)NumOr(a, 4, 0);
            double fv = -(pv * Math.Pow(1 + rate, nper) + pmt * (1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / rate);
            return FxValue.Num(fv);
        }

        private static FxValue PvFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double rate, nper, pmt;
            if (!NumberAt(a, 0, out rate) || !NumberAt(a, 1, out nper) || !NumberAt(a, 2, out pmt)) return FirstError(a) ?? Err(FxError.Value);
            double fv = NumOr(a, 3, 0);
            int type = (int)NumOr(a, 4, 0);
            double pv = -(fv * Math.Pow(1 + rate, nper) + pmt * (1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / rate);
            return FxValue.Num(pv);
        }

        private static FxValue PmtFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double rate, nper, pv;
            if (!NumberAt(a, 0, out rate) || !NumberAt(a, 1, out nper) || !NumberAt(a, 2, out pv)) return FirstError(a) ?? Err(FxError.Value);
            double fv = NumOr(a, 3, 0);
            int type = (int)NumOr(a, 4, 0);
            double pmt = -(pv * Math.Pow(1 + rate, nper) + fv) / ((1 + rate * type) * (Math.Pow(1 + rate, nper) - 1) / rate);
            return FxValue.Num(pmt);
        }

        private static FxValue NpvFn(Evaluator ev, Worksheet sheet, List<FxValue> a)
        {
            double rate;
            if (!NumberAt(a, 0, out rate)) return FirstError(a) ?? Err(FxError.Value);
            double npv = 0;
            int period = 1;
            for (int i = 1; i < a.Count; i++)
            {
                foreach (FxValue v in RangeValues(a[i]))
                {
                    double cf;
                    if (v.Kind == FxKind.Number) cf = v.Number;
                    else continue;
                    npv += cf / Math.Pow(1 + rate, period);
                    period++;
                }
            }
            return FxValue.Num(npv);
        }

        // ---- 通用值比较。 Shared value comparison. ----

        /// <summary>比较两个标量值，返回 -1/0/1 与是否可比。Compares two scalars returning -1/0/1 and whether the comparison is valid.</summary>
        internal static bool TryCompareValues(FxValue a, FxValue b, out int cmp)
        {
            cmp = 0;
            if (a.IsError || b.IsError) return false;
            if (a.Kind == FxKind.Text && b.Kind == FxKind.Text)
            {
                cmp = string.Compare(a.Text, b.Text, StringComparison.OrdinalIgnoreCase);
                return true;
            }
            double x;
            double y;
            bool ox = a.Kind == FxKind.Number || a.Kind == FxKind.Empty || a.Kind == FxKind.Logical;
            bool oy = b.Kind == FxKind.Number || b.Kind == FxKind.Empty || b.Kind == FxKind.Logical;
            if (ox && oy)
            {
                Evaluator.TryCoerceNumber(a, out x);
                Evaluator.TryCoerceNumber(b, out y);
                cmp = x.CompareTo(y);
                return true;
            }
            if (a.Kind == FxKind.Text || b.Kind == FxKind.Text)
            {
                cmp = string.Compare(Evaluator.ToDisplay(a), Evaluator.ToDisplay(b), StringComparison.OrdinalIgnoreCase);
                return true;
            }
            return false;
        }
    }
}
