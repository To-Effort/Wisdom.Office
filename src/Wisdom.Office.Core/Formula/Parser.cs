using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>
    /// 递归下降 / 优先级爬升解析器：把 token 序列解析成 AST。
    /// A recursive-descent / precedence-climbing parser that turns the token stream into an AST.
    /// </summary>
    /// <remarks>
    /// 运算符优先级（低到高）：比较（= &lt;&gt; &lt; &gt; &lt;= &gt;=）→ &amp; → + - → * / → ^ → 一元 + - → 百分号 %。
    /// Precedence, lowest to highest: comparison → concat (&amp;) → add/sub → mul/div → power → unary +,- → percent.
    /// 乘幂右结合，一元与二元比较/拼接等左结合。Power is right-associative; the rest are left-associative.
    /// </remarks>
    internal sealed class Parser
    {
        private readonly List<Token> _tokens;
        private int _pos;

        /// <summary>用 token 列表构造解析器。Constructs a parser from a token list.</summary>
        /// <param name="tokens">词法单元。The tokens.</param>
        public Parser(List<Token> tokens)
        {
            _tokens = tokens;
            _pos = 0;
        }

        /// <summary>解析整条公式并校验输入已全部消费。Parses the whole formula and checks all input was consumed.</summary>
        /// <returns>根节点。The root node.</returns>
        public Node ParseFormula()
        {
            Node node = ParseBinary(0);
            if (Current().Kind != TokKind.End)
            {
                throw new FormatException("Unexpected trailing tokens in formula.");
            }
            return node;
        }

        private Token Current()
        {
            return _tokens[_pos];
        }

        private Token Advance()
        {
            Token t = _tokens[_pos];
            if (_pos < _tokens.Count - 1) _pos++;
            return t;
        }

        private int BinaryPrecedence(string op)
        {
            switch (op)
            {
                case "=":
                case "<>":
                case "<":
                case ">":
                case "<=":
                case ">=":
                    return 1;
                case "&":
                    return 2;
                case "+":
                case "-":
                    return 3;
                case "*":
                case "/":
                    return 4;
                case "^":
                    return 5;
                default:
                    return 0;
            }
        }

        private Node ParseBinary(int minPrec)
        {
            Node left = ParseUnary();
            while (true)
            {
                Token t = Current();
                if (t.Kind != TokKind.Operator) break;
                int prec = BinaryPrecedence(t.Text);
                if (prec == 0 || prec < minPrec) break;
                Advance();
                int nextMin = t.Text == "^" ? prec : prec + 1;
                Node right = ParseBinary(nextMin);
                BinaryNode node = new BinaryNode();
                node.Kind = NodeKind.Binary;
                node.Op = t.Text;
                node.Left = left;
                node.Right = right;
                left = node;
            }
            return left;
        }

        private Node ParseUnary()
        {
            Token t = Current();
            if (t.Kind == TokKind.Operator && (t.Text == "-" || t.Text == "+"))
            {
                Advance();
                // 一元正负号优先级低于乘幂：-3^2 解析为 -(3^2)，与 Excel 一致。
                // Unary sign binds looser than power: -3^2 parses as -(3^2), matching Excel.
                Node operand = ParseBinary(5);
                if (t.Text == "+") return operand;
                UnaryNode node = new UnaryNode();
                node.Kind = NodeKind.Unary;
                node.Op = "-";
                node.Operand = operand;
                return node;
            }
            return ParsePostfix();
        }

        private Node ParsePostfix()
        {
            Node operand = ParsePrimary();
            while (Current().Kind == TokKind.Operator && Current().Text == "%")
            {
                Advance();
                BinaryNode node = new BinaryNode();
                node.Kind = NodeKind.Binary;
                node.Op = "%";
                node.Left = operand;
                node.Right = null;
                operand = node;
            }
            return operand;
        }

        private Node ParsePrimary()
        {
            Token t = Current();
            switch (t.Kind)
            {
                case TokKind.Number:
                    Advance();
                    return new LiteralNode { Kind = NodeKind.Number, Number = t.Number };
                case TokKind.String:
                    Advance();
                    return new LiteralNode { Kind = NodeKind.Text, Text = t.Text };
                case TokKind.Bool:
                    Advance();
                    return new LiteralNode { Kind = NodeKind.Bool, Bool = t.Bool };
                case TokKind.Error:
                    Advance();
                    return new LiteralNode { Kind = NodeKind.Error, Error = t.Error };
                case TokKind.LParen:
                    Advance();
                    Node inner = ParseBinary(0);
                    Expect(TokKind.RParen);
                    return inner;
                case TokKind.Function:
                    return ParseCall();
                case TokKind.Reference:
                    Advance();
                    return BuildReference(t.Text);
                default:
                    throw new FormatException("Unexpected token in formula: " + (t.Text ?? t.Kind.ToString()));
            }
        }

        private Node ParseCall()
        {
            Token name = Advance();
            Expect(TokKind.LParen);
            CallNode call = new CallNode();
            call.Kind = NodeKind.Call;
            call.Name = name.Text.ToUpperInvariant();
            if (Current().Kind == TokKind.RParen)
            {
                Advance();
                return call;
            }
            while (true)
            {
                // 允许空参数（连续逗号 / 尾随逗号）。 Allow empty arguments (consecutive or trailing commas).
                if (Current().Kind == TokKind.Comma)
                {
                    Advance();
                    call.Args.Add(null);
                    if (Current().Kind == TokKind.RParen) { Advance(); break; }
                    continue;
                }
                call.Args.Add(ParseBinary(0));
                if (Current().Kind == TokKind.Comma)
                {
                    Advance();
                    if (Current().Kind == TokKind.RParen) { Advance(); break; }
                    continue;
                }
                Expect(TokKind.RParen);
                break;
            }
            return call;
        }

        private void Expect(TokKind kind)
        {
            if (Current().Kind != kind)
            {
                throw new FormatException("Expected token " + kind + " but found " + (Current().Text ?? Current().Kind.ToString()));
            }
            if (kind != TokKind.End) Advance();
        }

        private static ReferenceNode BuildReference(string raw)
        {
            string sheet = null;
            string body = raw;
            int bang = body.LastIndexOf('!');
            if (bang >= 0)
            {
                sheet = body.Substring(0, bang);
                body = body.Substring(bang + 1);
            }

            string[] parts = body.Split(':');
            ReferenceNode node = new ReferenceNode();
            node.Kind = NodeKind.Reference;
            node.SheetPrefix = sheet;

            int row;
            int col;
            bool absRow;
            bool absCol;
            ParseCell(parts[0], out row, out col, out absRow, out absCol);
            node.Row1 = row;
            node.Col1 = col;
            node.AbsRow1 = absRow;
            node.AbsCol1 = absCol;

            if (parts.Length == 1)
            {
                node.IsRange = false;
                node.Row2 = row;
                node.Col2 = col;
            }
            else
            {
                int row2;
                int col2;
                bool ar2;
                bool ac2;
                ParseCell(parts[1], out row2, out col2, out ar2, out ac2);
                node.IsRange = true;
                node.Row2 = row2;
                node.Col2 = col2;
                if (node.Row1 > node.Row2) { int tmp = node.Row1; node.Row1 = node.Row2; node.Row2 = tmp; }
                if (node.Col1 > node.Col2) { int tmp = node.Col1; node.Col1 = node.Col2; node.Col2 = tmp; }
            }
            return node;
        }

        private static void ParseCell(string text, out int row, out int col, out bool absRow, out bool absCol)
        {
            absRow = false;
            absCol = false;
            int i = 0;
            if (i < text.Length && text[i] == '$') { absCol = true; i++; }
            long column = 0;
            int letters = 0;
            while (i < text.Length && char.IsLetter(text[i]))
            {
                column = column * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
                letters++;
                i++;
            }
            if (letters == 0 || column > CellAddress.MaxColumnLiteral) throw new FormatException("Invalid column in reference: " + text);
            if (i < text.Length && text[i] == '$') { absRow = true; i++; }
            long rowNumber = 0;
            int digits = 0;
            while (i < text.Length && char.IsDigit(text[i]))
            {
                rowNumber = rowNumber * 10 + (text[i] - '0');
                digits++;
                i++;
            }
            if (digits == 0 || rowNumber > CellAddress.MaxRowLiteral) throw new FormatException("Invalid row in reference: " + text);
            col = (int)column;
            row = (int)rowNumber;
        }

        /// <summary>把列号（1 基）转成字母文本。Converts a 1-based column number to letters.</summary>
        /// <param name="column">列号。The column.</param>
        /// <returns>列字母。The column letters.</returns>
        public static string ColumnToLetters(int column)
        {
            StringBuilder builder = new StringBuilder();
            while (column > 0)
            {
                int rem = (column - 1) % 26;
                builder.Insert(0, (char)('A' + rem));
                column = (column - 1) / 26;
            }
            return builder.ToString();
        }

        // 保留 CultureInvariant 数字转换的引用点，便于未来扩展内联数组。 Retained for future inline-array support.
        private static double InvariantParse(string text)
        {
            double value;
            double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
            return value;
        }
    }
}
