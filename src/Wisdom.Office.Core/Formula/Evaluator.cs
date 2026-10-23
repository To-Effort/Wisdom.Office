using System;
using System.Collections.Generic;
using System.Globalization;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>
    /// 公式求值器：遍历 AST、解析引用、检测循环、执行运算符并分派函数。
    /// The formula evaluator: walks the AST, resolves references, detects cycles, applies operators and dispatches functions.
    /// </summary>
    /// <remarks>
    /// 引用按公式所在单元格的物理坐标解释（OOXML 存的是相对该单元格的 A1 文本，就地求值即为绝对坐标）。
    /// 整列/整行引用会被裁到已用区域，避免为 <c>A:A</c> 分配上百万个单元格。
    /// References resolve to the physical coordinates shown in A1 (correct when evaluating a cell in place).
    /// Whole-column/row references are clamped to the used range so <c>A:A</c> does not allocate over a million cells.
    /// </remarks>
    internal sealed class Evaluator
    {
        private readonly OfficeWorkbook _workbook;
        private readonly Dictionary<string, FxValue> _cache;
        private readonly HashSet<string> _path;
        private readonly Dictionary<string, Node> _astCache;

        /// <summary>整表求值的最大区域格数。The maximum number of cells materialised for a range.</summary>
        public const int MaxRangeCells = 1000000;

        /// <summary>针对一个工作簿创建求值器。Creates an evaluator bound to one workbook.</summary>
        /// <param name="workbook">工作簿。The workbook.</param>
        /// <param name="cache">跨单元格共享的求值缓存。The value cache shared across cells.</param>
        public Evaluator(OfficeWorkbook workbook, Dictionary<string, FxValue> cache)
        {
            _workbook = workbook;
            _cache = cache ?? new Dictionary<string, FxValue>(StringComparer.OrdinalIgnoreCase);
            _path = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            _astCache = new Dictionary<string, Node>(StringComparer.Ordinal);
        }

        /// <summary>工作簿。The workbook.</summary>
        public OfficeWorkbook Workbook { get { return _workbook; } }

        /// <summary>生成单元格缓存键。Builds a cache key for a cell.</summary>
        public static string CellKey(string sheet, int row, int col)
        {
            return sheet + "!" + row + "," + col;
        }

        /// <summary>求值一条公式文本。Evaluates a formula text.</summary>
        /// <param name="formula">公式文本。The formula text.</param>
        /// <param name="sheet">公式所在工作表。The sheet hosting the formula.</param>
        /// <returns>求值结果。The result.</returns>
        public FxValue Evaluate(string formula, Worksheet sheet)
        {
            Node root = ParseCached(formula);
            return Eval(root, sheet);
        }

        /// <summary>解析并缓存 AST。Parses and caches the AST.</summary>
        public Node ParseCached(string formula)
        {
            Node node;
            if (_astCache.TryGetValue(formula ?? string.Empty, out node)) return node;
            node = new Parser(Tokenizer.Lex(formula)).ParseFormula();
            _astCache[formula ?? string.Empty] = node;
            return node;
        }

        /// <summary>求值单个 AST 节点。Evaluates a single AST node.</summary>
        /// <param name="node">节点。The node.</param>
        /// <param name="sheet">当前工作表。The current sheet.</param>
        /// <returns>运行时值。The runtime value.</returns>
        public FxValue Eval(Node node, Worksheet sheet)
        {
            if (node == null) return FxValue.Empty;
            switch (node.Kind)
            {
                case NodeKind.Number: return FxValue.Num(((LiteralNode)node).Number);
                case NodeKind.Text: return FxValue.Txt(((LiteralNode)node).Text);
                case NodeKind.Bool: return FxValue.Logical(((LiteralNode)node).Bool);
                case NodeKind.Error: return FxValue.Err(((LiteralNode)node).Error);
                case NodeKind.Reference: return ResolveReference((ReferenceNode)node, sheet);
                case NodeKind.Unary: return EvalUnary((UnaryNode)node, sheet);
                case NodeKind.Binary: return EvalBinary((BinaryNode)node, sheet);
                case NodeKind.Call: return EvalCall((CallNode)node, sheet);
                default: return FxValue.Err(FxError.Name);
            }
        }

        private FxValue EvalUnary(UnaryNode node, Worksheet sheet)
        {
            FxValue operand = Eval(node.Operand, sheet);
            if (operand.IsError) return operand;
            double d;
            if (!TryCoerceNumber(operand, out d)) return FxValue.Err(FxError.Value);
            return FxValue.Num(-d);
        }

        private FxValue EvalBinary(BinaryNode node, Worksheet sheet)
        {
            if (node.Op == "%")
            {
                FxValue operand = Eval(node.Left, sheet);
                if (operand.IsError) return operand;
                double p;
                if (!TryCoerceNumber(operand, out p)) return FxValue.Err(FxError.Value);
                return FxValue.Num(p / 100.0);
            }

            FxValue left = Eval(node.Left, sheet);
            FxValue right = Eval(node.Right, sheet);
            if (left.IsError) return left;
            if (right.IsError) return right;
            FxValue ls = Scalar(left);
            FxValue rs = Scalar(right);
            if (ls.IsError) return ls;
            if (rs.IsError) return rs;

            switch (node.Op)
            {
                case "+":
                case "-":
                case "*":
                case "/":
                case "^":
                    return Arith(node.Op, ls, rs);
                case "&":
                    return FxValue.Txt(ToDisplay(ls) + ToDisplay(rs));
                case "=":
                case "<>":
                case "<":
                case ">":
                case "<=":
                case ">=":
                    return Compare(node.Op, ls, rs);
                default:
                    return FxValue.Err(FxError.Name);
            }
        }

        private static FxValue Arith(string op, FxValue a, FxValue b)
        {
            double x;
            double y;
            if (!TryCoerceNumber(a, out x)) return FxValue.Err(FxError.Value);
            if (!TryCoerceNumber(b, out y)) return FxValue.Err(FxError.Value);
            switch (op)
            {
                case "+": return FxValue.Num(x + y);
                case "-": return FxValue.Num(x - y);
                case "*": return FxValue.Num(x * y);
                case "/":
                    if (y == 0) return FxValue.Err(FxError.Div0);
                    return FxValue.Num(x / y);
                case "^":
                    double r = Math.Pow(x, y);
                    if (double.IsNaN(r) || double.IsInfinity(r)) return FxValue.Err(FxError.Num);
                    return FxValue.Num(r);
                default: return FxValue.Err(FxError.Name);
            }
        }

        private static FxValue Compare(string op, FxValue a, FxValue b)
        {
            int cmp;
            if (!TryCompare(a, b, out cmp)) return FxValue.Err(FxError.Value);
            bool result;
            switch (op)
            {
                case "=": result = cmp == 0; break;
                case "<>": result = cmp != 0; break;
                case "<": result = cmp < 0; break;
                case ">": result = cmp > 0; break;
                case "<=": result = cmp <= 0; break;
                case ">=": result = cmp >= 0; break;
                default: return FxValue.Err(FxError.Name);
            }
            return FxValue.Logical(result);
        }

        private static bool TryCompare(FxValue a, FxValue b, out int cmp)
        {
            cmp = 0;
            bool aNum = a.Kind == FxKind.Number || a.Kind == FxKind.Empty || a.Kind == FxKind.Logical;
            bool bNum = b.Kind == FxKind.Number || b.Kind == FxKind.Empty || b.Kind == FxKind.Logical;

            double an;
            double bn;
            bool okA = TryCoerceNumber(a, out an);
            bool okB = TryCoerceNumber(b, out bn);

            if (a.Kind == FxKind.Text && b.Kind == FxKind.Text)
            {
                cmp = string.Compare(a.Text, b.Text, StringComparison.OrdinalIgnoreCase);
                return true;
            }
            if (okA && okB && !(a.Kind == FxKind.Text && b.Kind != FxKind.Text) && !(b.Kind == FxKind.Text && a.Kind != FxKind.Text))
            {
                cmp = an.CompareTo(bn);
                return true;
            }
            if (a.Kind == FxKind.Text && b.Kind == FxKind.Text)
            {
                cmp = string.Compare(a.Text, b.Text, StringComparison.OrdinalIgnoreCase);
                return true;
            }
            // 跨类型比较：数值 < 文本 < 逻辑，与 Excel 的排序一致。 Cross-type ordering: number < text < logical, matching Excel.
            int rankA = KindRank(a);
            int rankB = KindRank(b);
            if (rankA != rankB) { cmp = rankA.CompareTo(rankB); return true; }
            if (rankA == 3)
            {
                cmp = string.Compare(ToDisplay(a), ToDisplay(b), StringComparison.OrdinalIgnoreCase);
                return true;
            }
            if (rankA == 4)
            {
                cmp = (a.Kind == FxKind.Logical ? (a.Bool ? 1 : 0) : 0).CompareTo(b.Kind == FxKind.Logical ? (b.Bool ? 1 : 0) : 0);
                return true;
            }
            cmp = an.CompareTo(bn);
            return true;
        }

        private static int KindRank(FxValue v)
        {
            switch (v.Kind)
            {
                case FxKind.Empty: return 1;
                case FxKind.Number: return 1;
                case FxKind.Text: return 3;
                case FxKind.Logical: return 4;
                default: return 2;
            }
        }

        private FxValue EvalCall(CallNode node, Worksheet sheet)
        {
            // 需要惰性求值分支的逻辑函数：直接从 AST 节点按需 Eval，其余参数仍即时求值。
            // Logical functions that need short-circuit branch evaluation: evaluate AST nodes on demand here.
            switch (node.Name)
            {
                case "IF": return EvalIf(node, sheet);
                case "IFERROR": return EvalIfError(node, sheet);
                case "IFNA": return EvalIfNa(node, sheet);
                case "IFS": return EvalIfs(node, sheet);
                case "SWITCH": return EvalSwitch(node, sheet);
            }

            List<FxValue> args = new List<FxValue>(node.Args.Count);
            foreach (Node arg in node.Args)
            {
                if (arg == null) args.Add(FxValue.Empty);
                else args.Add(Eval(arg, sheet));
            }
            return Functions.Invoke(node.Name, args, this, sheet);
        }

        private Node ArgAt(CallNode node, int index)
        {
            return index < node.Args.Count ? node.Args[index] : null;
        }

        private FxValue EvalIf(CallNode node, Worksheet sheet)
        {
            FxValue cond = Eval(ArgAt(node, 0), sheet);
            if (cond.IsError) return cond;
            bool take;
            FxValue c = Evaluator.Scalar(cond);
            if (c.Kind == FxKind.Number) take = c.Number != 0;
            else if (c.Kind == FxKind.Logical) take = c.Bool;
            else if (c.Kind == FxKind.Text) take = string.Equals(c.Text, "TRUE", StringComparison.OrdinalIgnoreCase);
            else take = false;
            if (take) return Eval(ArgAt(node, 1), sheet);
            if (node.Args.Count > 2) return Eval(ArgAt(node, 2), sheet);
            return FxValue.Logical(false);
        }

        private FxValue EvalIfError(CallNode node, Worksheet sheet)
        {
            FxValue v = Eval(ArgAt(node, 0), sheet);
            if (!v.IsError) return v;
            return Eval(ArgAt(node, 1), sheet);
        }

        private FxValue EvalIfNa(CallNode node, Worksheet sheet)
        {
            FxValue v = Eval(ArgAt(node, 0), sheet);
            if (v.IsError && v.Error == FxError.Na) return Eval(ArgAt(node, 1), sheet);
            return v;
        }

        private FxValue EvalIfs(CallNode node, Worksheet sheet)
        {
            for (int i = 0; i + 1 < node.Args.Count; i += 2)
            {
                FxValue cond = Eval(node.Args[i], sheet);
                if (cond.IsError) return cond;
                FxValue c = Evaluator.Scalar(cond);
                bool take;
                if (c.Kind == FxKind.Number) take = c.Number != 0;
                else if (c.Kind == FxKind.Logical) take = c.Bool;
                else take = false;
                if (take) return Eval(node.Args[i + 1], sheet);
            }
            return FxValue.Err(FxError.Na);
        }

        private FxValue EvalSwitch(CallNode node, Worksheet sheet)
        {
            if (node.Args.Count < 2) return FxValue.Err(FxError.Value);
            FxValue subject = Eval(node.Args[0], sheet);
            if (subject.IsError) return subject;
            FxValue s = Evaluator.Scalar(subject);
            for (int i = 1; i + 1 < node.Args.Count; i += 2)
            {
                FxValue test = Eval(node.Args[i], sheet);
                if (test.IsError) return test;
                int cmp;
                if (TryCompare(s, Evaluator.Scalar(test), out cmp) && cmp == 0)
                {
                    return Eval(node.Args[i + 1], sheet);
                }
            }
            if (node.Args.Count % 2 == 0) return Eval(node.Args[node.Args.Count - 1], sheet);
            return FxValue.Err(FxError.Na);
        }


        private FxValue ResolveReference(ReferenceNode node, Worksheet sheet)
        {
            Worksheet target = sheet;
            if (node.SheetPrefix != null)
            {
                target = ResolveSheet(node.SheetPrefix);
                if (target == null) return FxValue.Err(FxError.Ref);
            }

            if (!node.IsRange)
            {
                return CellValue(target, node.Row1, node.Col1);
            }

            int r1 = node.Row1;
            int c1 = node.Col1;
            int r2 = node.Row2;
            int c2 = node.Col2;
            int usedRow = target.LastRow;
            int usedCol = target.LastColumn;
            if (r2 >= CellAddress.MaxRowLiteral || r2 > usedRow) r2 = usedRow < r1 ? r1 : usedRow;
            if (c2 >= CellAddress.MaxColumnLiteral || c2 > usedCol) c2 = usedCol < c1 ? c1 : usedCol;

            long cells = (long)(r2 - r1 + 1) * (c2 - c1 + 1);
            if (cells > MaxRangeCells) return FxValue.Err(FxError.Ref);

            FxValue range = new FxValue();
            range.Kind = FxKind.Range;
            range.Sheet = target;
            range.R1 = r1;
            range.C1 = c1;
            range.R2 = r2;
            range.C2 = c2;
            range.Cells = new List<List<FxValue>>(r2 - r1 + 1);
            for (int r = r1; r <= r2; r++)
            {
                List<FxValue> rowVals = new List<FxValue>(c2 - c1 + 1);
                for (int c = c1; c <= c2; c++)
                {
                    rowVals.Add(CellValue(target, r, c));
                }
                range.Cells.Add(rowVals);
            }
            return range;
        }

        private Worksheet ResolveSheet(string prefix)
        {
            string name = prefix;
            int bang = name.LastIndexOf('!');
            if (bang >= 0) name = name.Substring(0, bang);
            name = name.Trim();
            if (name.StartsWith("'") && name.EndsWith("'") && name.Length >= 2)
            {
                name = name.Substring(1, name.Length - 2).Replace("''", "'");
            }
            else if (name.StartsWith("[") )
            {
                int close = name.IndexOf(']');
                if (close >= 0) name = name.Substring(0, close);
                if (name.StartsWith("[")) name = name.Substring(1);
            }
            return _workbook == null ? null : _workbook.FindWorksheet(name);
        }

        /// <summary>读取单元格值；公式单元格触发递归求值，检测循环。Reads a cell value; a formula cell triggers recursive evaluation with cycle detection.</summary>
        public FxValue CellValue(Worksheet sheet, int row, int col)
        {
            Cell cell = sheet.FindCell(row, col);
            if (cell == null) return FxValue.Empty;
            if (string.IsNullOrEmpty(cell.Formula))
            {
                return FxValue.FromObject(cell.Value);
            }
            string key = CellKey(sheet.Name, row, col);
            FxValue cached;
            if (_cache.TryGetValue(key, out cached)) return cached;
            if (_path.Contains(key))
            {
                // 循环引用：以 #VALUE! 报告，避免无限递归。 Circular reference: report as #VALUE! to stop infinite recursion.
                return FxValue.Err(FxError.Value);
            }
            _path.Add(key);
            FxValue result;
            try
            {
                result = Evaluate(cell.Formula, sheet);
            }
            catch (FormatException)
            {
                result = FxValue.Err(FxError.Name);
            }
            catch (Exception)
            {
                result = FxValue.Err(FxError.Value);
            }
            finally
            {
                _path.Remove(key);
            }
            _cache[key] = result;
            return result;
        }

        // ---- 强制转换工具。 Coercion helpers. ----

        /// <summary>把区域折叠为标量；多格返回 #VALUE!。Collapses a range to a scalar; more than one cell yields #VALUE!.</summary>
        public static FxValue Scalar(FxValue v)
        {
            if (v.Kind != FxKind.Range) return v;
            if (v.Cells == null || v.Cells.Count == 0) return FxValue.Empty;
            if (v.Cells.Count == 1 && v.Cells[0].Count == 1) return v.Cells[0][0];
            return FxValue.Err(FxError.Value);
        }

        /// <summary>尝试把值当作数字。Text that parses as a number, TRUE/FALSE and empty are coerced; errors are not. Returns false for non-numeric text.</summary>
        public static bool TryCoerceNumber(FxValue v, out double d)
        {
            d = 0;
            switch (v.Kind)
            {
                case FxKind.Number: d = v.Number; return true;
                case FxKind.Empty: d = 0; return true;
                case FxKind.Logical: d = v.Bool ? 1 : 0; return true;
                case FxKind.Text:
                    if (double.TryParse(v.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return true;
                    if (double.TryParse(v.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out d)) return true;
                    if (string.Equals(v.Text, "TRUE", StringComparison.OrdinalIgnoreCase)) { d = 1; return true; }
                    if (string.Equals(v.Text, "FALSE", StringComparison.OrdinalIgnoreCase)) { d = 0; return true; }
                    return false;
                default: return false;
            }
        }

        /// <summary>把值转为显示文本（用于 &amp; 与文本比较）。Converts a value to display text (for concat and text comparison).</summary>
        public static string ToDisplay(FxValue v)
        {
            switch (v.Kind)
            {
                case FxKind.Empty: return string.Empty;
                case FxKind.Number: return NumberText(v.Number);
                case FxKind.Text: return v.Text ?? string.Empty;
                case FxKind.Logical: return v.Bool ? "TRUE" : "FALSE";
                case FxKind.Error: return v.ErrorText();
                default: return string.Empty;
            }
        }

        private static string NumberText(double v)
        {
            return v.ToString("G15", CultureInfo.InvariantCulture);
        }

        /// <summary>把区域展平成一维值序列。Flattens a range into a one-dimensional sequence.</summary>
        public static List<FxValue> Flatten(IEnumerable<FxValue> args)
        {
            List<FxValue> list = new List<FxValue>();
            foreach (FxValue a in args)
            {
                if (a.Kind == FxKind.Range && a.Cells != null)
                {
                    foreach (List<FxValue> row in a.Cells)
                        foreach (FxValue cell in row)
                            list.Add(cell);
                }
                else
                {
                    list.Add(a);
                }
            }
            return list;
        }

        /// <summary>仅保留区域里的原始单元格值（含空），供 COUNTA/COUNT 等使用。Keeps raw cell values inside ranges (including empty), for COUNTA/COUNT.</summary>
        public static List<FxValue> RawCells(FxValue arg)
        {
            List<FxValue> list = new List<FxValue>();
            if (arg.Kind == FxKind.Range && arg.Cells != null)
            {
                foreach (List<FxValue> row in arg.Cells)
                    foreach (FxValue cell in row)
                        list.Add(cell);
            }
            else
            {
                list.Add(arg);
            }
            return list;
        }
    }
}
