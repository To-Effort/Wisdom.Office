using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>词法单元类别。The lexical category of a token.</summary>
    internal enum TokKind
    {
        /// <summary>数字字面量。A number literal.</summary>
        Number,
        /// <summary>文本字面量。A text literal.</summary>
        String,
        /// <summary>单元格引用或区域。A cell reference or range.</summary>
        Reference,
        /// <summary>函数名。A function name.</summary>
        Function,
        /// <summary>二元/一元运算符。A binary or unary operator.</summary>
        Operator,
        /// <summary>左括号。Open parenthesis.</summary>
        LParen,
        /// <summary>右括号。Close parenthesis.</summary>
        RParen,
        /// <summary>参数分隔符（逗号或分号）。An argument separator (comma or semicolon).</summary>
        Comma,
        /// <summary>布尔字面量。A boolean literal.</summary>
        Bool,
        /// <summary>错误字面量。An error literal.</summary>
        Error,
        /// <summary>结束。End of input.</summary>
        End
    }

    /// <summary>一个词法单元。A single token.</summary>
    internal sealed class Token
    {
        /// <summary>类别。The category.</summary>
        public TokKind Kind;
        /// <summary>运算符字符或原始文本。The operator characters or raw text.</summary>
        public string Text;
        /// <summary>数字值（Kind==Number）。The number value (Kind==Number).</summary>
        public double Number;
        /// <summary>布尔值（Kind==Bool）。The logical value (Kind==Bool).</summary>
        public bool Bool;
        /// <summary>错误值（Kind==Error）。The error (Kind==Error).</summary>
        public FxError Error;

        /// <summary>构造词法单元。Constructs a token.</summary>
        /// <param name="kind">类别。The category.</param>
        /// <param name="text">原始文本。The raw text.</param>
        public Token(TokKind kind, string text)
        {
            Kind = kind;
            Text = text;
        }
    }

    /// <summary>
    /// Excel 公式词法分析器：把公式文本切成 token 序列。
    /// The Excel formula lexer: splits formula text into a token stream.
    /// </summary>
    /// <remarks>
    /// 支持数字（含科学计数与前导小数点）、双引号文本（"" 转义）、A1 引用与 $ 绝对、区域（含跨表 Sheet!、'Sheet Name'!）、
    /// 函数名（标识符后紧跟左括号）、TRUE/FALSE、#错误字面量、以及运算符 &lt;&gt; &lt;= &gt;= = &lt; &gt; + - * / ^ &amp; %。
    /// Handles numbers (scientific notation and a leading decimal point), double-quoted text ("" escape), A1 references
    /// with $ absolute markers, ranges (including cross-sheet Sheet! and 'Sheet Name'!), function names (identifier
    /// immediately followed by an open paren), TRUE/FALSE, error literals and the operators &lt;&gt; &lt;= &gt;= = &lt; &gt; + - * / ^ &amp; %.
    /// </remarks>
    internal static class Tokenizer
    {
        // 长字面量在前：#NAME? 必须先于 #N/A、#NUM! 参与匹配，否则 #NAME? 会被切成 #NA 或 #N/A。
        // Longest literals first: #NAME? must be compared before #N/A and #NUM!, otherwise it would be
        // chopped into #NA or #N/A.
        private static readonly string[] ErrorLiterals =
        {
            "#GETTING_DATA", "#DIV/0!", "#VALUE!", "#NAME?", "#NULL!", "#REF!", "#N/A", "#NUM!"
        };

        /// <summary>把公式文本切分为 token 列表；末尾追加 End。Splits formula text into a token list and appends End at the tail.</summary>
        /// <param name="formula">公式文本（可有可无前导 =）。The formula text, with or without a leading =.</param>
        /// <returns>token 列表。The token list.</returns>
        /// <exception cref="FormatException">遇到非法字符或结构。Thrown on an illegal character or structure.</exception>
        public static List<Token> Lex(string formula)
        {
            List<Token> tokens = new List<Token>();
            if (formula == null) formula = string.Empty;
            int i = 0;
            int n = formula.Length;
            if (i < n && formula[i] == '=') i++;

            while (i < n)
            {
                char c = formula[i];
                if (c == ' ' || c == '\t' || c == '\r' || c == '\n')
                {
                    i++;
                    continue;
                }

                if (c == '"')
                {
                    StringBuilder sb = new StringBuilder();
                    i++;
                    while (i < n)
                    {
                        if (formula[i] == '"')
                        {
                            if (i + 1 < n && formula[i + 1] == '"')
                            {
                                sb.Append('"');
                                i += 2;
                                continue;
                            }
                            i++;
                            break;
                        }
                        sb.Append(formula[i]);
                        i++;
                    }
                    tokens.Add(new Token(TokKind.String, sb.ToString()));
                    continue;
                }

                if (c == '#')
                {
                    int start = i;
                    // 错误字面量含 "/" 和数字（#DIV/0!、#N/A），旧扫描在 "/" 处断开，把 =#DIV/0! 拆成
                    // "#DIV" + "0" + "!" 三段后抛 Unexpected character。按字面量表整段匹配（长的优先，
                    // #NUM! 要排在 #N/A 之前比较），未命中再退回原先只认字母的扫描，保持畸形输入的旧行为。
                    // Error literals contain "/" and digits (#DIV/0!, #N/A); the old scan broke at "/", splitting
                    // =#DIV/0! into "#DIV" + "0" + "!" and then throwing Unexpected character. Match the literal
                    // table as one run (longest first, so #NUM! is compared before #N/A) and fall back to the old
                    // letter-only scan when nothing matches, keeping the previous behaviour for malformed input.
                    string literal = MatchErrorLiteral(formula, i);
                    if (literal != null)
                    {
                        i += literal.Length;
                        tokens.Add(new Token(TokKind.Error, literal) { Error = ParseError(literal) });
                        continue;
                    }
                    i++;
                    while (i < n && (char.IsLetter(formula[i]) || formula[i] == '?' || formula[i] == '!')) i++;
                    string err = formula.Substring(start, i - start);
                    tokens.Add(new Token(TokKind.Error, err) { Error = ParseError(err) });
                    continue;
                }

                if (char.IsDigit(c) || (c == '.' && i + 1 < n && char.IsDigit(formula[i + 1])))
                {
                    int start = i;
                    while (i < n && char.IsDigit(formula[i])) i++;
                    if (i < n && formula[i] == '.')
                    {
                        i++;
                        while (i < n && char.IsDigit(formula[i])) i++;
                    }
                    if (i < n && (formula[i] == 'e' || formula[i] == 'E'))
                    {
                        int save = i;
                        i++;
                        if (i < n && (formula[i] == '+' || formula[i] == '-')) i++;
                        if (i < n && char.IsDigit(formula[i]))
                        {
                            while (i < n && char.IsDigit(formula[i])) i++;
                        }
                        else
                        {
                            i = save;
                        }
                    }
                    string text = formula.Substring(start, i - start);
                    double value;
                    if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                    {
                        throw new FormatException("Invalid number in formula: " + text);
                    }
                    // 整行引用以数字开头，会先落进数字分支：纯整数后面紧跟冒号且冒号后还是行号时才改判为引用。
                    // Whole-row references start with a digit and reach the number branch first; only a plain
                    // integer immediately followed by a colon and another row number becomes a reference.
                    bool integral = text.IndexOf('.') < 0 && text.IndexOf('e') < 0 && text.IndexOf('E') < 0;
                    if (integral && i < n && formula[i] == ':')
                    {
                        int save = i;
                        i++;
                        string secondRow = ReadWholeSpanSide(formula, ref i, false);
                        if (secondRow != null)
                        {
                            tokens.Add(new Token(TokKind.Reference, text + ":" + secondRow));
                            continue;
                        }
                        i = save;
                    }
                    tokens.Add(new Token(TokKind.Number, text) { Number = value });
                    continue;
                }

                if (c == '\'')
                {
                    // 带引号的跨表引用前缀：'My Sheet'!A1。 Quoted sheet prefix for a cross-sheet reference.
                    int start = i;
                    i++;
                    while (i < n)
                    {
                        if (formula[i] == '\'')
                        {
                            if (i + 1 < n && formula[i + 1] == '\'')
                            {
                                i += 2;
                                continue;
                            }
                            i++;
                            break;
                        }
                        i++;
                    }
                    string prefix = formula.Substring(start, i - start);
                    // 引号后紧跟的 "!" 必须吃掉：早先直接去读引用尾部，而 "!A1" 既不是 $ 也不是字母，
                    // 读出空串，'Data Set'!A1 整个退化成表名，跨表引用报 #REF!/#NAME?。
                    // The "!" after the closing quote must be consumed: the old code went straight to reading the
                    // reference tail, and because "!A1" starts with neither $ nor a letter it read an empty string,
                    // collapsing 'Data Set'!A1 to the bare sheet name and failing the cross-sheet reference.
                    string ref2 = string.Empty;
                    if (i < n && formula[i] == '!')
                    {
                        i++;
                        ref2 = "!" + ReadReferenceTail(formula, ref i);
                    }
                    tokens.Add(new Token(TokKind.Reference, prefix + ref2));
                    continue;
                }

                if (c == '$' || char.IsLetter(c) || c == '_' || c == '\\')
                {
                    int start = i;
                    string sheetPrefix = null;
                    // 先贪婪读出标识符（列字母或表名）。 Read the identifier greedily (column letters or a sheet name).
                    if (c == '$')
                    {
                        i++;
                        while (i < n && char.IsLetter(formula[i])) i++;
                        if (i < n && formula[i] == '$') i++;
                        // 绝对引用的行号也要吃掉：旧实现只读到列字母，LooksLikeCellReference("$A") 判假，
                        // 于是 $A1 与 $A$1 整个落到名称分支报 #NAME?。
                        // The row part of an absolute reference must be consumed too: the old scan stopped at the
                        // column letters, LooksLikeCellReference("$A") returned false, and both $A1 and $A$1 fell
                        // through to the name branch as #NAME?.
                        while (i < n && char.IsDigit(formula[i])) i++;
                    }
                    else
                    {
                        while (i < n && (char.IsLetter(formula[i]) || char.IsDigit(formula[i]) || formula[i] == '_' || formula[i] == '.' || formula[i] == '\\')) i++;
                        // 相对列 + 绝对行（A$1）也要认成引用：扫描在 $ 处停下时，A$1 被切成 "A" 和 "$1"
                        // 两个记号，=SUM(A$1:A$3) 直接落到名称分支报 #NAME?。只有纯列字母后面紧跟
                        // $数字 才并入，避免把 Sheet1$A2 这类表名尾巴吃错。
                        // A relative column with an absolute row (A$1) must lex as a reference too: the scan used
                        // to stop at the '$', splitting A$1 into "A" plus "$1", so =SUM(A$1:A$3) fell through to the
                        // name branch and returned #NAME?. Only pure column letters followed by $digits are merged,
                        // so a sheet prefix such as Sheet1$A2 is not mis-sliced.
                        if (i < n && formula[i] == '$' && IsColumnRun(formula, start, i))
                        {
                            int probe = i + 1;
                            while (probe < n && char.IsDigit(formula[probe])) probe++;
                            if (probe > i + 1) i = probe;
                        }
                    }
                    string ident = formula.Substring(start, i - start);

                    if (i < n && formula[i] == '!')
                    {
                        sheetPrefix = ident;
                        i++;
                        string refTail = ReadReferenceTail(formula, ref i);
                        tokens.Add(new Token(TokKind.Reference, sheetPrefix + "!" + refTail));
                        continue;
                    }

                    if (i < n && formula[i] == '(')
                    {
                        tokens.Add(new Token(TokKind.Function, ident));
                        continue;
                    }

                    if (string.Equals(ident, "TRUE", StringComparison.OrdinalIgnoreCase))
                    {
                        tokens.Add(new Token(TokKind.Bool, ident) { Bool = true });
                        continue;
                    }
                    if (string.Equals(ident, "FALSE", StringComparison.OrdinalIgnoreCase))
                    {
                        tokens.Add(new Token(TokKind.Bool, ident) { Bool = false });
                        continue;
                    }

                    if (LooksLikeCellReference(ident))
                    {
                        string tail = string.Empty;
                        if (i < n && formula[i] == ':')
                        {
                            int save = i;
                            string second = ReadReferenceTailNoColon(formula, ref i);
                            if (second == null) i = save;
                            else tail = ":" + second;
                        }
                        tokens.Add(new Token(TokKind.Reference, ident + tail));
                        continue;
                    }

                    // 整列 / 整行引用（A:A、$A:$B、1:3）：Excel 支持，早先因为 ident 里没有行号而落到名称分支，
                    // SUM(A:A) 直接报 #NAME?。这里只接两侧同kind的写法，接不上就继续按名称处理。
                    // Whole-column / whole-row references (A:A, $A:$B, 1:3): Excel supports them, but because the
                    // identifier carried no row number they used to fall through to the name branch and SUM(A:A)
                    // returned #NAME?. Only matching halves on both sides are accepted; anything else stays a name.
                    bool wholeColumn = IsWholeColumn(ident);
                    bool wholeRow = !wholeColumn && IsWholeRow(ident);
                    if ((wholeColumn || wholeRow) && i < n && formula[i] == ':')
                    {
                        int save = i;
                        i++;
                        string second = ReadWholeSpanSide(formula, ref i, wholeColumn);
                        if (second != null)
                        {
                            tokens.Add(new Token(TokKind.Reference, ident + ":" + second));
                            continue;
                        }
                        i = save;
                    }

                    // 其余当作未支持的名称，交由解析器/求值报 #NAME?。 Anything else is an unsupported name; the parser leaves it for #NAME?.
                    tokens.Add(new Token(TokKind.Function, ident));
                    continue;
                }

                if (c == '(') { tokens.Add(new Token(TokKind.LParen, "(")); i++; continue; }
                if (c == ')') { tokens.Add(new Token(TokKind.RParen, ")")); i++; continue; }
                if (c == ',' || c == ';') { tokens.Add(new Token(TokKind.Comma, c.ToString())); i++; continue; }

                if (TryReadOperator(formula, ref i, out string op))
                {
                    tokens.Add(new Token(TokKind.Operator, op));
                    continue;
                }

                throw new FormatException("Unexpected character in formula: '" + c + "'");
            }

            tokens.Add(new Token(TokKind.End, null));
            return tokens;
        }

        /// <summary>从位置读取 A1 引用尾部（$字母$数字）。Reads the A1 tail ($letters$digits) starting at the index.</summary>
        private static string ReadReferenceTail(string formula, ref int i)
        {
            int start = i;
            if (i < formula.Length && formula[i] == '$') i++;
            while (i < formula.Length && char.IsLetter(formula[i])) i++;
            if (i < formula.Length && formula[i] == '$') i++;
            while (i < formula.Length && char.IsDigit(formula[i])) i++;
            string result = formula.Substring(start, i - start);
            if (i < formula.Length && formula[i] == ':')
            {
                int save = i;
                i++;
                string second = ReadReferenceTailRaw(formula, ref i);
                result += ":" + second;
            }
            return result;
        }

        /// <summary>读取一个引用尾部原始文本。Reads a raw reference tail.</summary>
        private static string ReadReferenceTailRaw(string formula, ref int i)
        {
            int start = i;
            if (i < formula.Length && formula[i] == '$') i++;
            while (i < formula.Length && char.IsLetter(formula[i])) i++;
            if (i < formula.Length && formula[i] == '$') i++;
            while (i < formula.Length && char.IsDigit(formula[i])) i++;
            return formula.Substring(start, i - start);
        }

        /// <summary>读取冒号后的第二个引用；非法则返回 null。Reads the second reference after a colon; null when malformed.</summary>
        private static string ReadReferenceTailNoColon(string formula, ref int i)
        {
            if (i >= formula.Length || formula[i] != ':') return null;
            i++;
            int start = i;
            if (i < formula.Length && formula[i] == '$') i++;
            int letters = 0;
            while (i < formula.Length && char.IsLetter(formula[i])) { i++; letters++; }
            if (i < formula.Length && formula[i] == '$') i++;
            int digits = 0;
            while (i < formula.Length && char.IsDigit(formula[i])) { i++; digits++; }
            if (letters == 0 || digits == 0)
            {
                i = start;
                return null;
            }
            return formula.Substring(start, i - start);
        }

        /// <summary>读取标识符之后的引用尾部（不带冒号继续）。Reads the reference tail after a sheet prefix.</summary>
        private static string ReadReferenceTailAfterBang(string formula, ref int i)
        {
            return ReadReferenceTail(formula, ref i);
        }

        private static bool TryReadOperator(string formula, ref int i, out string op)
        {
            char c = formula[i];
            switch (c)
            {
                case '<':
                    if (i + 1 < formula.Length && formula[i + 1] == '>') { op = "<>"; i += 2; return true; }
                    if (i + 1 < formula.Length && formula[i + 1] == '=') { op = "<="; i += 2; return true; }
                    op = "<"; i++; return true;
                case '>':
                    if (i + 1 < formula.Length && formula[i + 1] == '=') { op = ">="; i += 2; return true; }
                    op = ">"; i++; return true;
                case '=':
                case '+':
                case '-':
                case '*':
                case '/':
                case '^':
                case '&':
                case '%':
                    op = c.ToString(); i++; return true;
                default:
                    op = null; return false;
            }
        }

        /// <summary>
        /// 判断 [start, end) 是否是一段纯列字母（A、AB 之类），即可以接 $行号 成为引用。
        /// Whether [start, end) is a pure run of column letters (A, AB, ...) that a $row suffix can turn into a
        /// reference.
        /// </summary>
        /// <param name="formula">公式全文。The full formula text.</param>
        /// <param name="start">片段起点。The inclusive start index.</param>
        /// <param name="end">片段终点。The exclusive end index.</param>
        /// <returns>是纯列字母片段返回 true。True when the slice is only column letters.</returns>
        private static bool IsColumnRun(string formula, int start, int end)
        {
            if (end <= start) return false;
            for (int idx = start; idx < end; idx++)
            {
                char c = formula[idx];
                bool upper = c >= 'A' && c <= 'Z';
                bool lower = c >= 'a' && c <= 'z';
                if (!upper && !lower) return false;
            }
            return true;
        }

        private static bool LooksLikeCellReference(string ident)
        {
            // 列字母（可含前导 $）后跟行数字。 Column letters (possibly with a leading $) followed by row digits.
            int idx = 0;
            if (idx < ident.Length && ident[idx] == '$') idx++;
            int letters = 0;
            while (idx < ident.Length && char.IsLetter(ident[idx])) { idx++; letters++; }
            if (idx < ident.Length && ident[idx] == '$') idx++;
            int digits = 0;
            while (idx < ident.Length && char.IsDigit(ident[idx])) { idx++; digits++; }
            return letters > 0 && digits > 0 && idx == ident.Length;
        }

        /// <summary>
        /// 在指定位置匹配 Excel 错误字面量，返回原文（含 "#"）或 null。
        /// Matches an Excel error literal at the given position and returns it including the "#", or null.
        /// </summary>
        private static string MatchErrorLiteral(string formula, int index)
        {
            foreach (string literal in ErrorLiterals)
            {
                if (index + literal.Length > formula.Length) continue;
                if (string.Compare(formula, index, literal, 0, literal.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return formula.Substring(index, literal.Length);
                }
            }
            return null;
        }

        /// <summary>
        /// 是否为整列引用的一段：可选 $ 后跟纯列字母。Whether one side of a whole-column reference: optional $ then column letters only.
        /// </summary>
        private static bool IsWholeColumn(string ident)
        {
            int idx = 0;
            if (idx < ident.Length && ident[idx] == '$') idx++;
            int letters = 0;
            while (idx < ident.Length && char.IsLetter(ident[idx])) { idx++; letters++; }
            return letters > 0 && idx == ident.Length;
        }

        /// <summary>
        /// 是否为整行引用的一段：可选 $ 后跟纯行号。Whether one side of a whole-row reference: optional $ then row digits only.
        /// </summary>
        private static bool IsWholeRow(string ident)
        {
            int idx = 0;
            if (idx < ident.Length && ident[idx] == '$') idx++;
            int digits = 0;
            while (idx < ident.Length && char.IsDigit(ident[idx])) { idx++; digits++; }
            return digits > 0 && idx == ident.Length;
        }

        /// <summary>
        /// 读取整列/整行引用冒号后的另一半；不匹配时回退位置并返回 null。
        /// Reads the half after the colon of a whole-column/row reference, restoring the position and returning null on a mismatch.
        /// </summary>
        private static string ReadWholeSpanSide(string formula, ref int i, bool column)
        {
            int start = i;
            if (i < formula.Length && formula[i] == '$') i++;
            int matched = 0;
            while (i < formula.Length && (column ? char.IsLetter(formula[i]) : char.IsDigit(formula[i]))) { i++; matched++; }
            if (matched == 0)
            {
                i = start;
                return null;
            }
            return formula.Substring(start, i - start);
        }

        private static FxError ParseError(string text)
        {
            string t = text.TrimStart('#').TrimEnd('!', '?').ToUpperInvariant();
            switch (t)
            {
                case "NULL": return FxError.Null;
                case "DIV/0": return FxError.Div0;
                case "REF": return FxError.Ref;
                case "VALUE": return FxError.Value;
                case "NAME": return FxError.Name;
                case "NUM": return FxError.Num;
                case "N/A": return FxError.Na;
                case "NA": return FxError.Na;
                case "GETTING_DATA": return FxError.GettingData;
                default: return FxError.Name;
            }
        }
    }
}
