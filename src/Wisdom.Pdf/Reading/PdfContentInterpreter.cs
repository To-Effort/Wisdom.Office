using System;
using System.Collections.Generic;
using System.IO;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 一段已变换到设备空间的矢量线段，用于表格边框识别。
    /// A vector line segment already transformed into device space, used for table-border detection.
    /// </summary>
    internal sealed class PdfVectorSegment
    {
        /// <summary>起点横坐标。The start X coordinate.</summary>
        public double X1;
        /// <summary>起点纵坐标。The start Y coordinate.</summary>
        public double Y1;
        /// <summary>终点横坐标。The end X coordinate.</summary>
        public double X2;
        /// <summary>终点纵坐标。The end Y coordinate.</summary>
        public double Y2;

        /// <summary>线段长度。The length of the segment.</summary>
        public double Length
        {
            get
            {
                double dx = X2 - X1;
                double dy = Y2 - Y1;
                return Math.Sqrt(dx * dx + dy * dy);
            }
        }
    }

    /// <summary>
    /// 文本片段收集器：把解释器逐个输出的字元合并为片段，并从 PDF 左下原点翻转为以页面左上角为原点。
    /// Text-fragment collector: merges the glyphs emitted one by one by the interpreter into fragments and flips the PDF bottom-left origin to a page top-left origin.
    /// </summary>
    /// <remarks>合并判定在与文本流向对齐的旋转坐标系里完成，因此 45 度水印、竖排文字也能拼成完整片段；
    /// 包围盒取旋转矩形的轴对齐外接框，流向为 0 度时与未旋转的算法结果完全一致。
    /// The merge test runs in a coordinate frame aligned with the text flow, so a 45-degree watermark or
    /// vertical text still joins into whole fragments; the bounding box is the axis-aligned envelope of the
    /// rotated run rectangle, which reproduces the previous result exactly when the flow is horizontal.</remarks>
    internal sealed class PdfFragmentCollector
    {
        /// <summary>页面高度，用于把左下原点翻转为左上原点。The page height, used to flip the bottom-left origin to a top-left origin.</summary>
        private readonly double _pageHeight;
        /// <summary>最近一次输出的片段，用于判断能否继续并入。The most recently emitted fragment, used to decide whether a glyph can be appended to it.</summary>
        private PdfTextFragment _last;
        /// <summary>当前行流向角度的余弦。The cosine of the current line's flow angle.</summary>
        private double _cos;
        /// <summary>当前行流向角度的正弦。The sine of the current line's flow angle.</summary>
        private double _sin;
        /// <summary>当前行起点在流向轴上的坐标。The current line's start on the flow axis.</summary>
        private double _left;
        /// <summary>当前行终点在流向轴上的坐标。The current line's end on the flow axis.</summary>
        private double _right;
        /// <summary>当前行基线在法向轴上的坐标。The current line's baseline on the normal axis.</summary>
        private double _normal;
        /// <summary>当前行的流向角度，弧度制。The current line's flow angle, in radians.</summary>
        private double _angle;

        /// <summary>
        /// 创建收集器。
        /// Creates the collector.
        /// </summary>
        /// <param name="pageHeight">页面高度，用于 y 坐标翻转。The page height, used to flip the Y coordinate.</param>
        public PdfFragmentCollector(double pageHeight)
        {
            _pageHeight = pageHeight;
        }

        /// <summary>已收集的片段，按输出顺序排列。The collected fragments in output order.</summary>
        public List<PdfTextFragment> Fragments { get; } = new List<PdfTextFragment>();

        /// <summary>
        /// 记录一个字元的文本与位置；与上一片段同向、同基线、同字号且沿流向间距足够小时直接并入。
        /// Records a glyph's text and position; merges it into the previous fragment when they share flow, baseline and size and the gap along the flow is small enough.
        /// </summary>
        /// <param name="text">字元文本，空文本忽略。The glyph text; empty text is ignored.</param>
        /// <param name="x">字元起点横坐标（设备空间）。The glyph's start X coordinate (device space).</param>
        /// <param name="baseline">字元起点纵坐标（设备空间，仍为左下原点）。The glyph's start Y coordinate (device space, still bottom-left origin).</param>
        /// <param name="width">字元沿流向的位移量，负值取绝对值。The glyph's advance along the flow; negative values are made absolute.</param>
        /// <param name="fontSize">生效字号，非正值按 1 处理以避免包围盒塌陷。The effective font size; non-positive values are treated as 1 to avoid a collapsed bounding box.</param>
        /// <param name="angle">文本流向相对设备空间 x 轴的夹角，弧度制，0 表示横排。The text flow direction relative to the device X axis, in radians; 0 means horizontal text.</param>
        /// <remarks>合并间距上限随字号放缩并夹在 1.25–3.0 之间；下限允许 -1.5 倍字号，以容纳轻微重叠的定位。
        /// 上下边界取基线上方 0.78、下方 0.42 倍字号，是无字体度量时的近似值。
        /// The upper merge gap scales with the font size and is clamped between 1.25 and 3.0; the lower bound allows -1.5 times the size to tolerate slightly overlapping positions. The extents use 0.78 of the size above and 0.42 below the baseline, approximations used when no font metrics are available.</remarks>
        public void Add(string text, double x, double baseline, double width, double fontSize, double angle)
        {
            if (string.IsNullOrEmpty(text)) return;
            if (fontSize <= 0) fontSize = 1;
            if (width < 0) width = -width;

            double cos = Math.Cos(angle);
            double sin = Math.Sin(angle);
            double flow = x * cos + baseline * sin;
            double normal = baseline * cos - x * sin;

            if (_last != null && Math.Abs(_angle - angle) < 0.02
                && Math.Abs(_normal - normal) < 0.8
                && Math.Abs(_last.FontSize - fontSize) < 0.3)
            {
                double gap = flow - _right;
                double mergeGap = Math.Max(1.25, Math.Min(3.0, fontSize * 0.25));
                if (gap >= -fontSize * 1.5 && gap <= mergeGap)
                {
                    _last.Text += text;
                    _right = Math.Max(_right, flow + width);
                    Layout(_last);
                    return;
                }
            }

            PdfTextFragment fragment = new PdfTextFragment();
            fragment.Text = text;
            fragment.FontSize = fontSize;
            fragment.Baseline = normal;
            fragment.Angle = angle;
            _last = fragment;
            _angle = angle;
            _cos = cos;
            _sin = sin;
            _left = flow;
            _right = flow + width;
            _normal = normal;
            Layout(fragment);
            Fragments.Add(fragment);
        }

        /// <summary>
        /// 依据流向矩形重算片段的轴对齐包围盒。
        /// Recomputes the fragment's axis-aligned bounding box from the rotated run rectangle.
        /// </summary>
        /// <param name="fragment">待排版的片段。The fragment to lay out.</param>
        private void Layout(PdfTextFragment fragment)
        {
            double size = fragment.FontSize;
            double above = _normal + size * 0.78;
            double below = _normal - size * 0.42;

            double x1 = _left * _cos - below * _sin;
            double y1 = _left * _sin + below * _cos;
            double x2 = _left * _cos - above * _sin;
            double y2 = _left * _sin + above * _cos;
            double x3 = _right * _cos - below * _sin;
            double y3 = _right * _sin + below * _cos;
            double x4 = _right * _cos - above * _sin;
            double y4 = _right * _sin + above * _cos;

            double minX = Math.Min(Math.Min(x1, x2), Math.Min(x3, x4));
            double maxX = Math.Max(Math.Max(x1, x2), Math.Max(x3, x4));
            double minY = Math.Min(Math.Min(y1, y2), Math.Min(y3, y4));
            double maxY = Math.Max(Math.Max(y1, y2), Math.Max(y3, y4));

            fragment.X = minX;
            fragment.Y = _pageHeight - maxY;
            fragment.Width = Math.Max(0, maxX - minX);
            fragment.Height = Math.Max(0, maxY - minY);
        }
    }

    /// <summary>
    /// PDF 内容流解释器：实现文本与路径相关的图形状态算符，输出文本片段和矢量线段。
    /// PDF content-stream interpreter: implements the text- and path-related graphics-state operators, emitting text fragments and vector segments.
    /// </summary>
    /// <remarks>这是一个只求几何正确的简化解释器：颜色、线型、裁剪与变换外的效果均被忽略，
    /// 单条算符执行失败只吞掉异常继续下一条，因此对损坏的内容流具备容错性。
    /// This is a simplified interpreter concerned only with geometric correctness: color, line style, clipping and effects beyond transforms are ignored, and a failure in one operator only swallows the exception and continues, so it tolerates corrupted content streams.</remarks>
    internal sealed class PdfContentInterpreter
    {
        /// <summary>q/Q 之间保存的图形状态快照。A graphics-state snapshot saved between q/Q.</summary>
        private sealed class GraphicsState
        {
            /// <summary>当前变换矩阵副本。A copy of the current transformation matrix.</summary>
            public double[] Ctm;
            /// <summary>当前字体。The current font.</summary>
            public PdfTextFont Font;
            /// <summary>当前字号。The current font size.</summary>
            public double FontSize;
            /// <summary>字符间距 Tc。The character spacing Tc.</summary>
            public double CharSpacing;
            /// <summary>词间隔 Tw。The word spacing Tw.</summary>
            public double WordSpacing;
            /// <summary>水平缩放 Tz，百分比。The horizontal scaling Tz, in percent.</summary>
            public double HorizontalScale;
            /// <summary>行距 TL。The text leading TL.</summary>
            public double Leading;
            /// <summary>基线上移 Ts。The text rise Ts.</summary>
            public double Rise;
        }

        /// <summary>所属文档，用于解引用与解码流。The owning document, used for dereferencing and decoding streams.</summary>
        private readonly PdfDocument _document;
        /// <summary>页面高度，用于坐标翻转。The page height, used to flip coordinates.</summary>
        private readonly double _pageHeight;
        /// <summary>文本片段输出收集器。The text-fragment output collector.</summary>
        private readonly PdfFragmentCollector _collector;
        /// <summary>矢量线段输出列表。The output list of vector segments.</summary>
        private readonly List<PdfVectorSegment> _segments;
        /// <summary>字体字典到度量对象的共享缓存。A shared cache from font dictionaries to metric objects.</summary>
        private readonly Dictionary<PdfDictionary, PdfTextFont> _fonts;
        /// <summary>q/Q 状态栈。The q/Q graphics-state stack.</summary>
        private readonly Stack<GraphicsState> _stack = new Stack<GraphicsState>();
        /// <summary>当前尚未绘出的路径。The current path not yet painted.</summary>
        private readonly List<PdfVectorSegment> _path = new List<PdfVectorSegment>();

        /// <summary>当前变换矩阵 CTM，元素顺序为 a b c d e f。The current transformation matrix CTM, in a b c d e f order.</summary>
        private double[] _ctm = Identity();
        /// <summary>文本矩阵 TM。The text matrix TM.</summary>
        private double[] _tm = Identity();
        /// <summary>文本行距矩阵 TLM。The text line matrix TLM.</summary>
        private double[] _tlm = Identity();
        /// <summary>当前字体。The current font.</summary>
        private PdfTextFont _font;
        /// <summary>当前字号。The current font size.</summary>
        private double _fontSize;
        /// <summary>当前字符间距 Tc。The current character spacing Tc.</summary>
        private double _charSpacing;
        /// <summary>当前词间隔 Tw。The current word spacing Tw.</summary>
        private double _wordSpacing;
        /// <summary>当前水平缩放 Tz（百分比），默认 100。The current horizontal scaling Tz (percent), default 100.</summary>
        private double _horizontalScale = 100;
        /// <summary>当前行距 TL。The current text leading TL.</summary>
        private double _leading;
        /// <summary>当前基线上移 Ts。The current text rise Ts.</summary>
        private double _rise;
        /// <summary>路径当前点 x。The current path point X.</summary>
        private double _currentX;
        /// <summary>路径当前点 y。The current path point Y.</summary>
        private double _currentY;
        /// <summary>当前子路径起点 x，供 h 闭合命令回退。The current subpath start X, used by the h close command.</summary>
        private double _subpathX;
        /// <summary>当前子路径起点 y。The current subpath start Y.</summary>
        private double _subpathY;
        /// <summary>本轮执行使用的资源字典，Form XObject 会替换为自己的资源。The resource dictionary in use this run; a Form XObject substitutes its own resources.</summary>
        private PdfDictionary _resources;

        /// <summary>
        /// 创建解释器并绑定输出目标。
        /// Creates the interpreter and binds its output targets.
        /// </summary>
        /// <param name="document">所属文档，用于解引用与解码流。The owning document, used for dereferencing and decoding streams.</param>
        /// <param name="pageHeight">页面高度，用于坐标翻转。The page height, used to flip coordinates.</param>
        /// <param name="collector">文本片段输出收集器。The text-fragment output collector.</param>
        /// <param name="segments">矢量线段输出列表。The vector-segment output list.</param>
        /// <param name="fonts">字体字典到度量对象的共享缓存，跨页面与 Form 复用。The shared font-dictionary-to-metric cache, reused across pages and Forms.</param>
        public PdfContentInterpreter(PdfDocument document, double pageHeight, PdfFragmentCollector collector,
            List<PdfVectorSegment> segments, Dictionary<PdfDictionary, PdfTextFont> fonts)
        {
            _document = document;
            _pageHeight = pageHeight;
            _collector = collector;
            _segments = segments;
            _fonts = fonts;
        }

        /// <summary>
        /// 逐词元读取内容流并执行算符。
        /// Reads the content stream token by token and executes each operator.
        /// </summary>
        /// <param name="content">已解码的内容流字节，null 或空时不执行。The decoded content-stream bytes; nothing runs when null or empty.</param>
        /// <param name="resources">本轮使用的资源字典，null 时取空字典。The resource dictionary for this run; an empty dictionary when null.</param>
        /// <param name="initialCtm">入口 CTM，null 时取单位矩阵（Form XObject 传入叠加后的矩阵）。The entry CTM; the identity matrix when null (a Form XObject passes its composed matrix).</param>
        /// <param name="depth">当前 Form 嵌套深度，超过 12 直接返回，防止 XObject 循环引用。The current Form nesting depth; returns immediately past 12 to prevent XObject reference cycles.</param>
        /// <remarks>对象解析失败时清空已收集的操作数并前进一个字节以跳出坏词元；
        /// 线段总数超过 25 万时停止执行，避免超大路径拖垮提取。
        /// When an object fails to parse, collected operands are cleared and the cursor advances one byte to escape the bad token; execution stops once the segment count exceeds 250,000 to keep an oversized path from stalling extraction.</remarks>
        public void Run(byte[] content, PdfDictionary resources, double[] initialCtm, int depth)
        {
            if (content == null || content.Length == 0) return;
            if (depth > 12) return;
            _resources = resources ?? new PdfDictionary();
            _ctm = initialCtm == null ? Identity() : (double[])initialCtm.Clone();
            _tm = Identity();
            _tlm = Identity();

            PdfParser parser = new PdfParser(content, 0, _document.Resolve);
            List<PdfObject> operands = new List<PdfObject>();
            while (!parser.EndOfData)
            {
                parser.SkipWhitespaceAndComments();
                if (parser.EndOfData) break;
                byte value = parser.Data[parser.Position];
                if (IsObjectStart(parser, value))
                {
                    try
                    {
                        operands.Add(parser.ParseObject());
                    }
                    catch (InvalidDataException)
                    {
                        parser.Position++;
                        operands.Clear();
                    }
                    continue;
                }

                string op = parser.ReadKeyword();
                if (string.IsNullOrEmpty(op))
                {
                    parser.Position++;
                    operands.Clear();
                    continue;
                }
                if (op == "BI")
                {
                    SkipInlineImage(parser);
                    operands.Clear();
                    continue;
                }
                try
                {
                    Execute(op, operands, depth);
                }
                catch (Exception)
                {
                }
                operands.Clear();
                if (_segments.Count > 250000) break;
            }
        }

        /// <summary>
        /// 判断当前位置开始的字节是否是一个对象而非算符关键字。
        /// Determines whether the byte at the current position begins an object rather than an operator keyword.
        /// </summary>
        /// <param name="parser">当前解析器，用于确认 true/false/null 的完整拼写。The current parser, used to confirm the full spelling of true/false/null.</param>
        /// <param name="value">当前字节。The current byte.</param>
        /// <returns>应作为对象解析返回 true，否则视为算符。True when it should be parsed as an object, otherwise treated as an operator.</returns>
        private static bool IsObjectStart(PdfParser parser, byte value)
        {
            if (value == (byte)'/' || value == (byte)'(' || value == (byte)'<' || value == (byte)'['
                || value == (byte)'+' || value == (byte)'-' || value == (byte)'.'
                || (value >= (byte)'0' && value <= (byte)'9'))
            {
                return true;
            }
            if (value == (byte)'t') return parser.MatchesKeyword(parser.Position, "true");
            if (value == (byte)'f') return parser.MatchesKeyword(parser.Position, "false");
            if (value == (byte)'n') return parser.MatchesKeyword(parser.Position, "null");
            return false;
        }

        /// <summary>
        /// 分派执行单个算符。
        /// Dispatches and executes a single operator.
        /// </summary>
        /// <param name="op">算符名，如 Tj、cm、re。The operator name, such as Tj, cm, re.</param>
        /// <param name="operands">该算符累积的操作数，一律从尾部取，以容忍前面残留的多余对象。The operands accumulated for this operator, always taken from the tail to tolerate leftover objects before it.</param>
        /// <param name="depth">当前 Form 嵌套深度。The current Form nesting depth.</param>
        /// <remarks>操作数不足时静默跳过；曲线命令（c/v/y）只记录终点，控制点被忽略，因为表格识别只关心直线段；
        /// 填充与描边算符都把当前路径输出为线段，空操作算符（n/W/W*）则直接丢弃路径。
        /// Insufficient operands are silently skipped; curve commands (c/v/y) record only the endpoint and drop the control points, since table detection cares only about straight segments. Paint and stroke operators flush the path as segments, while no-op operators (n/W/W*) discard the path.</remarks>
        private void Execute(string op, List<PdfObject> operands, int depth)
        {
            switch (op)
            {
                case "q":
                    PushState();
                    return;
                case "Q":
                    PopState();
                    return;
                case "cm":
                    if (operands.Count >= 6)
                    {
                        double[] matrix = Matrix(operands, operands.Count - 6);
                        _ctm = Multiply(matrix, _ctm);
                    }
                    return;
                case "BT":
                    _tm = Identity();
                    _tlm = Identity();
                    return;
                case "ET":
                    return;
                case "Tf":
                    if (operands.Count >= 2)
                    {
                        PdfName name = _document.Resolve(operands[operands.Count - 2]) as PdfName;
                        PdfNumber size = _document.Resolve(operands[operands.Count - 1]) as PdfNumber;
                        if (name != null) _font = GetFont(name.Value);
                        if (size != null) _fontSize = size.Value;
                    }
                    return;
                case "Tc":
                    if (operands.Count >= 1) _charSpacing = GetNumber(operands[operands.Count - 1]);
                    return;
                case "Tw":
                    if (operands.Count >= 1) _wordSpacing = GetNumber(operands[operands.Count - 1]);
                    return;
                case "Tz":
                    if (operands.Count >= 1) _horizontalScale = GetNumber(operands[operands.Count - 1]);
                    return;
                case "TL":
                    if (operands.Count >= 1) _leading = GetNumber(operands[operands.Count - 1]);
                    return;
                case "Ts":
                    if (operands.Count >= 1) _rise = GetNumber(operands[operands.Count - 1]);
                    return;
                case "Td":
                case "TD":
                    if (operands.Count >= 2)
                    {
                        double tx = GetNumber(operands[operands.Count - 2]);
                        double ty = GetNumber(operands[operands.Count - 1]);
                        if (op == "TD") _leading = -ty;
                        _tlm = Multiply(Translate(tx, ty), _tlm);
                        _tm = (double[])_tlm.Clone();
                    }
                    return;
                case "Tm":
                    if (operands.Count >= 6)
                    {
                        double[] matrix = Matrix(operands, operands.Count - 6);
                        _tm = matrix;
                        _tlm = (double[])matrix.Clone();
                    }
                    return;
                case "T*":
                    _tlm = Multiply(Translate(0, -_leading), _tlm);
                    _tm = (double[])_tlm.Clone();
                    return;
                case "Tj":
                    if (operands.Count >= 1) ShowString(operands[operands.Count - 1]);
                    return;
                case "TJ":
                    if (operands.Count >= 1) ShowArray(operands[operands.Count - 1]);
                    return;
                case "'":
                    if (operands.Count >= 1)
                    {
                        _tlm = Multiply(Translate(0, -_leading), _tlm);
                        _tm = (double[])_tlm.Clone();
                        ShowString(operands[operands.Count - 1]);
                    }
                    return;
                case "\"":
                    if (operands.Count >= 3)
                    {
                        _wordSpacing = GetNumber(operands[operands.Count - 3]);
                        _charSpacing = GetNumber(operands[operands.Count - 2]);
                        _tlm = Multiply(Translate(0, -_leading), _tlm);
                        _tm = (double[])_tlm.Clone();
                        ShowString(operands[operands.Count - 1]);
                    }
                    return;
                case "Do":
                    if (operands.Count >= 1)
                    {
                        PdfName name = _document.Resolve(operands[operands.Count - 1]) as PdfName;
                        if (name != null) ExecuteXObject(name.Value, depth);
                    }
                    return;
                case "m":
                    if (operands.Count >= 2)
                    {
                        _currentX = GetNumber(operands[operands.Count - 2]);
                        _currentY = GetNumber(operands[operands.Count - 1]);
                        _subpathX = _currentX;
                        _subpathY = _currentY;
                    }
                    return;
                case "l":
                    if (operands.Count >= 2)
                    {
                        double x = GetNumber(operands[operands.Count - 2]);
                        double y = GetNumber(operands[operands.Count - 1]);
                        AddPathSegment(_currentX, _currentY, x, y);
                        _currentX = x;
                        _currentY = y;
                    }
                    return;
                case "c":
                    if (operands.Count >= 6)
                    {
                        _currentX = GetNumber(operands[operands.Count - 2]);
                        _currentY = GetNumber(operands[operands.Count - 1]);
                    }
                    return;
                case "v":
                case "y":
                    if (operands.Count >= 4)
                    {
                        _currentX = GetNumber(operands[operands.Count - 2]);
                        _currentY = GetNumber(operands[operands.Count - 1]);
                    }
                    return;
                case "re":
                    if (operands.Count >= 4)
                    {
                        double x = GetNumber(operands[operands.Count - 4]);
                        double y = GetNumber(operands[operands.Count - 3]);
                        double width = GetNumber(operands[operands.Count - 2]);
                        double height = GetNumber(operands[operands.Count - 1]);
                        AddPathSegment(x, y, x + width, y);
                        AddPathSegment(x + width, y, x + width, y + height);
                        AddPathSegment(x + width, y + height, x, y + height);
                        AddPathSegment(x, y + height, x, y);
                        _currentX = x;
                        _currentY = y;
                        _subpathX = x;
                        _subpathY = y;
                    }
                    return;
                case "h":
                    _currentX = _subpathX;
                    _currentY = _subpathY;
                    return;
                case "S":
                case "s":
                case "f":
                case "F":
                case "f*":
                case "B":
                case "B*":
                case "b":
                case "b*":
                    FlushPath();
                    return;
                case "n":
                case "W":
                case "W*":
                    _path.Clear();
                    return;
            }
        }

        /// <summary>
        /// 显示一个字符串：逐字元计算包围盒交给收集器，并推进文本矩阵。
        /// Shows a string: computes each glyph's bounding box for the collector and advances the text matrix.
        /// </summary>
        /// <param name="value">显示串对象（或其引用）。The show-string object (or its reference).</param>
        /// <remarks>字元位移量按 PDF 规范为 (宽度/1000 × 字号 + Tc + 空格时的 Tw) × Tz/100；
        /// 实际字号与宽度由合成矩阵的量级推得，因此在旋转或缩放页面下也能得到合理包围盒。
        /// The glyph displacement follows the PDF spec as (width/1000 * size + Tc + Tw for spaces) * Tz/100; the effective size and width are derived from the composed matrix magnitude, so a sensible bounding box results even on rotated or scaled pages.</remarks>
        private void ShowString(PdfObject value)
        {
            PdfString text = _document.Resolve(value) as PdfString;
            if (text == null || _font == null) return;
            List<PdfGlyph> glyphs = _font.Decode(text.Bytes);
            for (int i = 0; i < glyphs.Count; i++)
            {
                PdfGlyph glyph = glyphs[i];
                double tx = ((glyph.Width / 1000.0) * _fontSize + _charSpacing
                    + (glyph.IsWordSpace ? _wordSpacing : 0)) * _horizontalScale / 100.0;

                double[] trm = Multiply(FontMatrix(), Multiply(_tm, _ctm));
                double fontSize = Math.Sqrt(trm[2] * trm[2] + trm[3] * trm[3]);
                double flowX = _tm[0] * _ctm[0] + _tm[1] * _ctm[2];
                double flowY = _tm[0] * _ctm[1] + _tm[1] * _ctm[3];
                double angle = Math.Atan2(flowY, flowX);
                double scale = Math.Sqrt(flowX * flowX + flowY * flowY);
                double width = Math.Abs(tx * scale);
                if (width <= 0) width = Math.Abs(tx);
                _collector.Add(glyph.Text, trm[4], trm[5], width, fontSize, angle);

                _tm[4] += tx * _tm[0];
                _tm[5] += tx * _tm[1];
            }
        }

        /// <summary>
        /// 执行 TJ 数组：字符串照常显示，数字按千分之一 em 的负值调整字距。
        /// Executes a TJ array: strings are shown normally while numbers adjust spacing as negative thousandths of an em.
        /// </summary>
        /// <param name="value">数组对象（或其引用）。The array object (or its reference).</param>
        /// <remarks>较大的负数调整正是分词的依据，留给后续的行归并阶段按间距补空格。
        /// Larger negative adjustments are the basis for word breaking, deferred to the line-merging stage to insert spaces by gap.</remarks>
        private void ShowArray(PdfObject value)
        {
            PdfArray array = _document.Resolve(value) as PdfArray;
            if (array == null) return;
            for (int i = 0; i < array.Count; i++)
            {
                PdfObject item = _document.Resolve(array[i]);
                PdfString text = item as PdfString;
                if (text != null)
                {
                    ShowString(text);
                    continue;
                }
                PdfNumber number = item as PdfNumber;
                if (number != null)
                {
                    double adjustment = -number.Value / 1000.0 * _fontSize * _horizontalScale / 100.0;
                    _tm[4] += adjustment * _tm[0];
                    _tm[5] += adjustment * _tm[1];
                }
            }
        }

        /// <summary>
        /// 执行 Do：递归解释 Form 类型的 XObject 内容流。
        /// Executes Do: recursively interprets the content stream of a Form XObject.
        /// </summary>
        /// <param name="name">XObject 资源名。The XObject resource name.</param>
        /// <param name="depth">当前嵌套深度，达到 12 时不再进入，防止自引用 Form 造成无限递归。The current nesting depth; recursion stops at 12 to prevent self-referential Forms from looping forever.</param>
        /// <remarks>只有 Subtype 为 Form 的对象会被执行；表单的 /Matrix 先与外部 CTM 叠加作为入口矩阵，
        /// 资源缺失时沿用调用方的资源字典（符合规范的隐式继承）。图像类 XObject 直接忽略。
        /// Only objects whose Subtype is Form are executed; the form's /Matrix is composed with the outer CTM as the entry matrix, and when resources are absent the caller's resources are inherited (the implicit inheritance allowed by the spec). Image XObjects are ignored.</remarks>
        private void ExecuteXObject(string name, int depth)
        {
            if (depth >= 12) return;
            PdfDictionary xobjects = _document.Resolve(_resources.Get("XObject")) as PdfDictionary;
            if (xobjects == null) return;
            PdfStream form = _document.Resolve(xobjects.Get(name)) as PdfStream;
            if (form == null) return;
            PdfName subtype = _document.Resolve(form.Dictionary.Get("Subtype")) as PdfName;
            if (subtype == null || subtype.Value != "Form") return;

            byte[] content;
            try
            {
                content = _document.DecodeStream(form);
            }
            catch (Exception)
            {
                return;
            }

            double[] matrix = Identity();
            PdfArray formMatrix = _document.Resolve(form.Dictionary.Get("Matrix")) as PdfArray;
            if (formMatrix != null && formMatrix.Count >= 6)
            {
                matrix = new double[6];
                for (int i = 0; i < 6; i++)
                {
                    PdfNumber number = _document.Resolve(formMatrix[i]) as PdfNumber;
                    matrix[i] = number == null ? (i == 0 || i == 3 ? 1 : 0) : number.Value;
                }
            }

            PdfDictionary formResources = _document.Resolve(form.Dictionary.Get("Resources")) as PdfDictionary;
            if (formResources == null) formResources = _resources;
            double[] formCtm = Multiply(matrix, _ctm);
            PdfContentInterpreter interpreter = new PdfContentInterpreter(_document, _pageHeight, _collector, _segments, _fonts);
            interpreter.Run(content, formResources, formCtm, depth + 1);
        }

        /// <summary>
        /// 按资源名取得字体度量，结果按字体字典缓存。
        /// Resolves font metrics by resource name, caching the result per font dictionary.
        /// </summary>
        /// <param name="resourceName">Font 资源字典中的键。The key within the Font resource dictionary.</param>
        /// <returns>字体度量；资源缺失或字体解析失败时返回 null（失败结果同样缓存，避免重复解析）。The font metrics; null when the resource is missing or parsing fails (failures are cached too, avoiding repeated parses).</returns>
        private PdfTextFont GetFont(string resourceName)
        {
            PdfDictionary fontDictionary = _document.Resolve(_resources.Get("Font")) as PdfDictionary;
            if (fontDictionary == null) return null;
            PdfDictionary font = _document.Resolve(fontDictionary.Get(resourceName)) as PdfDictionary;
            if (font == null) return null;
            PdfTextFont result;
            if (_fonts.TryGetValue(font, out result)) return result;
            try
            {
                result = new PdfTextFont(_document, font);
            }
            catch (Exception)
            {
                result = null;
            }
            _fonts[font] = result;
            return result;
        }

        /// <summary>
        /// 执行 q：把当前图形状态压栈（CTM 存副本，避免后续 cm 改写栈内矩阵）。
        /// Executes q: pushes the current graphics state onto the stack (the CTM is copied so a later cm cannot mutate the stacked matrix).
        /// </summary>
        private void PushState()
        {
            GraphicsState state = new GraphicsState();
            state.Ctm = (double[])_ctm.Clone();
            state.Font = _font;
            state.FontSize = _fontSize;
            state.CharSpacing = _charSpacing;
            state.WordSpacing = _wordSpacing;
            state.HorizontalScale = _horizontalScale;
            state.Leading = _leading;
            state.Rise = _rise;
            _stack.Push(state);
        }

        /// <summary>
        /// 执行 Q：恢复上一个图形状态；栈为空时不做任何事（容忍多余的 Q）。
        /// Executes Q: restores the previous graphics state; does nothing when the stack is empty (tolerating stray Q operators).
        /// </summary>
        private void PopState()
        {
            if (_stack.Count == 0) return;
            GraphicsState state = _stack.Pop();
            _ctm = state.Ctm;
            _font = state.Font;
            _fontSize = state.FontSize;
            _charSpacing = state.CharSpacing;
            _wordSpacing = state.WordSpacing;
            _horizontalScale = state.HorizontalScale;
            _leading = state.Leading;
            _rise = state.Rise;
        }

        /// <summary>
        /// 以设备坐标追加一条路径线段。
        /// Appends a path segment in device coordinates.
        /// </summary>
        /// <param name="x1">起点 x（用户空间）。The start X (user space).</param>
        /// <param name="y1">起点 y（用户空间）。The start Y (user space).</param>
        /// <param name="x2">终点 x。The end X.</param>
        /// <param name="y2">终点 y。The end Y.</param>
        private void AddPathSegment(double x1, double y1, double x2, double y2)
        {
            double tx1;
            double ty1;
            double tx2;
            double ty2;
            Transform(_ctm, x1, y1, out tx1, out ty1);
            Transform(_ctm, x2, y2, out tx2, out ty2);
            PdfVectorSegment segment = new PdfVectorSegment();
            segment.X1 = tx1;
            segment.Y1 = ty1;
            segment.X2 = tx2;
            segment.Y2 = ty2;
            _path.Add(segment);
        }

        /// <summary>
        /// 描边或填充算符：把暂存路径转入输出线段并清空。
        /// For stroke or fill operators: moves the pending path into the output segments and clears it.
        /// </summary>
        private void FlushPath()
        {
            if (_path.Count > 0)
            {
                _segments.AddRange(_path);
                _path.Clear();
            }
        }

        /// <summary>
        /// 构造文本状态 Tm = [字号×Tz/100 0 0 字号 0 Rise]。
        /// Builds the text-state matrix Tm = [size*Tz/100 0 0 size 0 rise].
        /// </summary>
        /// <returns>六元素仿射矩阵。The six-element affine matrix.</returns>
        private double[] FontMatrix()
        {
            return new double[] { _fontSize * _horizontalScale / 100.0, 0, 0, _fontSize, 0, _rise };
        }

        /// <summary>
        /// 把操作数按数值读取，缺失或类型不符时取 0。
        /// Reads an operand as a number, returning 0 when it is missing or of the wrong type.
        /// </summary>
        /// <param name="value">操作数对象（可能是引用）。The operand object (possibly a reference).</param>
        /// <returns>数值。The numeric value.</returns>
        private double GetNumber(PdfObject value)
        {
            PdfNumber number = _document.Resolve(value) as PdfNumber;
            return number == null ? 0 : number.Value;
        }

        /// <summary>
        /// 从操作数列表尾部取连续 6 个数值组成仿射矩阵。
        /// Takes six consecutive numbers from the operand list to form an affine matrix.
        /// </summary>
        /// <param name="operands">操作数列表。The operand list.</param>
        /// <param name="start">起始下标，通常为 Count - 6。The start index, usually Count - 6.</param>
        /// <returns>六元素矩阵，非数值项按 0 填充。The six-element matrix, with non-numeric entries filled as 0.</returns>
        private double[] Matrix(List<PdfObject> operands, int start)
        {
            double[] matrix = new double[6];
            for (int i = 0; i < 6; i++)
            {
                PdfNumber number = _document.Resolve(operands[start + i]) as PdfNumber;
                matrix[i] = number == null ? 0 : number.Value;
            }
            return matrix;
        }

        /// <summary>
        /// 跳过内联图像（BI … ID 二进制数据 … EI），把游标移到 EI 之后。
        /// Skips an inline image (BI ... ID binary data ... EI), moving the cursor past the EI marker.
        /// </summary>
        /// <param name="parser">内容流解析器，其游标会被推进。The content-stream parser whose cursor is advanced.</param>
        /// <remarks>EI 只有前后都是空白才算结束标记，否则可能命中图像压缩数据里的字节对；找不到时直接跳到流末尾。
        /// EI counts as the end marker only when surrounded by whitespace, otherwise it can match a byte pair inside the image data; when no valid EI is found the cursor jumps to the end of the stream.</remarks>
        private static void SkipInlineImage(PdfParser parser)
        {
            byte[] data = parser.Data;
            int search = parser.Position;
            while (search + 2 <= data.Length)
            {
                int found = PdfParser.FindToken(data, "EI", search);
                if (found < 0)
                {
                    parser.Position = data.Length;
                    return;
                }
                bool before = found == 0 || PdfParser.IsWhitespace(data[found - 1]);
                bool after = found + 2 >= data.Length || PdfParser.IsWhitespace(data[found + 2]);
                if (before && after)
                {
                    parser.Position = found + 2;
                    return;
                }
                search = found + 2;
            }
            parser.Position = data.Length;
        }

        /// <summary>
        /// 单位矩阵 [1 0 0 1 0 0]。
        /// The identity matrix [1 0 0 1 0 0].
        /// </summary>
        /// <returns>新分配的六元素矩阵。A newly allocated six-element matrix.</returns>
        internal static double[] Identity()
        {
            return new double[] { 1, 0, 0, 1, 0, 0 };
        }

        /// <summary>
        /// 平移矩阵 [1 0 0 1 x y]。
        /// The translation matrix [1 0 0 1 x y].
        /// </summary>
        /// <param name="x">横向位移。The horizontal displacement.</param>
        /// <param name="y">纵向位移。The vertical displacement.</param>
        /// <returns>新分配的六元素矩阵。A newly allocated six-element matrix.</returns>
        internal static double[] Translate(double x, double y)
        {
            return new double[] { 1, 0, 0, 1, x, y };
        }

        /// <summary>
        /// 计算仿射矩阵乘积 a × b（行主序 [a b c d e f] 约定，e/f 为平移分量）。
        /// Computes the affine product a * b under the row-major [a b c d e f] convention, where e/f are the translation components.
        /// </summary>
        /// <param name="a">左矩阵。The left matrix.</param>
        /// <param name="b">右矩阵。The right matrix.</param>
        /// <returns>新分配的乘积矩阵。A newly allocated product matrix.</returns>
        internal static double[] Multiply(double[] a, double[] b)
        {
            return new double[]
            {
                a[0] * b[0] + a[1] * b[2],
                a[0] * b[1] + a[1] * b[3],
                a[2] * b[0] + a[3] * b[2],
                a[2] * b[1] + a[3] * b[3],
                a[4] * b[0] + a[5] * b[2] + b[4],
                a[4] * b[1] + a[5] * b[3] + b[5]
            };
        }

        /// <summary>
        /// 用矩阵变换一个点。
        /// Transforms a point by a matrix.
        /// </summary>
        /// <param name="matrix">六元素仿射矩阵。The six-element affine matrix.</param>
        /// <param name="x">源点横坐标。The source X coordinate.</param>
        /// <param name="y">源点纵坐标。The source Y coordinate.</param>
        /// <param name="tx">变换后的横坐标。The transformed X coordinate.</param>
        /// <param name="ty">变换后的纵坐标。The transformed Y coordinate.</param>
        private static void Transform(double[] matrix, double x, double y, out double tx, out double ty)
        {
            tx = x * matrix[0] + y * matrix[2] + matrix[4];
            ty = x * matrix[1] + y * matrix[3] + matrix[5];
        }
    }
}
