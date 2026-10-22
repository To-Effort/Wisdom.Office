using System.Collections.Generic;

namespace Wisdom.Office.Core.Formulas
{
    /// <summary>抽象语法树节点类别。The abstract-syntax-tree node category.</summary>
    internal enum NodeKind
    {
        /// <summary>数字字面量。Number literal.</summary>
        Number,
        /// <summary>文本字面量。Text literal.</summary>
        Text,
        /// <summary>布尔字面量。Logical literal.</summary>
        Bool,
        /// <summary>错误字面量。Error literal.</summary>
        Error,
        /// <summary>单元格引用。Cell reference.</summary>
        Reference,
        /// <summary>二元运算。Binary operation.</summary>
        Binary,
        /// <summary>一元运算。Unary operation.</summary>
        Unary,
        /// <summary>函数调用。Function call.</summary>
        Call,
        /// <summary>内联数组（如 {1,2;3,4}）。Inline array literal such as {1,2;3,4}.</summary>
        Array
    }

    /// <summary>AST 节点基类。Base class for AST nodes.</summary>
    internal abstract class Node
    {
        /// <summary>节点类别。The node category.</summary>
        public NodeKind Kind;
    }

    /// <summary>字面量节点（数字/文本/布尔/错误）。Literal node (number/text/logical/error).</summary>
    internal sealed class LiteralNode : Node
    {
        /// <summary>数值。Number.</summary>
        public double Number;
        /// <summary>文本。Text.</summary>
        public string Text;
        /// <summary>布尔。Logical.</summary>
        public bool Bool;
        /// <summary>错误。Error.</summary>
        public FxError Error;
    }

    /// <summary>引用节点：单个单元格或矩形区域，可带表前缀。Reference node: a single cell or a rectangular range, optionally sheet-qualified.</summary>
    internal sealed class ReferenceNode : Node
    {
        /// <summary>表前缀原文（可能带引号），null 表示本表。Raw sheet prefix (possibly quoted); null means the current sheet.</summary>
        public string SheetPrefix;
        /// <summary>是否矩形区域。Whether this is a range rather than a single cell.</summary>
        public bool IsRange;
        /// <summary>起始行（1 基）。1-based top row.</summary>
        public int Row1;
        /// <summary>起始列（1 基）。1-based left column.</summary>
        public int Col1;
        /// <summary>结束行（1 基）。1-based bottom row.</summary>
        public int Row2;
        /// <summary>结束列（1 基）。1-based right column.</summary>
        public int Col2;
        /// <summary>起始行是否为绝对引用（$）。Whether the top row is absolute ($).</summary>
        public bool AbsRow1;
        /// <summary>起始列是否绝对。Whether the left column is absolute.</summary>
        public bool AbsCol1;
    }

    /// <summary>二元运算节点。Binary operation node.</summary>
    internal sealed class BinaryNode : Node
    {
        /// <summary>运算符。Operator.</summary>
        public string Op;
        /// <summary>左操作数。Left operand.</summary>
        public Node Left;
        /// <summary>右操作数。Right operand.</summary>
        public Node Right;
    }

    /// <summary>一元运算节点。Unary operation node.</summary>
    internal sealed class UnaryNode : Node
    {
        /// <summary>运算符。Operator.</summary>
        public string Op;
        /// <summary>操作数。Operand.</summary>
        public Node Operand;
    }

    /// <summary>函数调用节点。Function call node.</summary>
    internal sealed class CallNode : Node
    {
        /// <summary>函数名（大写）。Function name, upper-cased.</summary>
        public string Name;
        /// <summary>参数节点列表。Argument nodes.</summary>
        public List<Node> Args = new List<Node>();
    }
}
