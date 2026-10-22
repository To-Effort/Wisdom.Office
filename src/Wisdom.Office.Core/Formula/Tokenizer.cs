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
                    string ref2 = ReadReferenceTail(formula, ref i);
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
                    }
                    else
                    {
                        while (i < n && (char.IsLetter(formula[i]) || char.IsDigit(formula[i]) || formula[i] == '_' || formula[i] == '.' || formula[i] == '\\')) i++;
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
