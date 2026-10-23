using System;
using System.Collections.Generic;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 页面内容提取器：驱动内容流解释器收集文本片段与矢量线段，再组装为文本行和表格。
    /// Page content extractor: drives the content stream interpreter to collect text fragments and vector
    /// segments, then assembles them into text lines and tables.
    /// </summary>
    internal static class PdfPageExtractor
    {
        /// <summary>
        /// 提取整页的文本行与表格并回写到页面对象。
        /// Extracts the page's text lines and tables and writes them back to the page object.
        /// </summary>
        /// <param name="page">目标页面。The target page.</param>
        /// <remarks>内容流语法损坏时吞掉异常，返回已收集到的部分结果，保证读取不因单页失败而中断；
        /// 字体字典到度量对象的缓存按字典引用共享，避免同一字体重复解析。
        /// Corrupted content stream syntax swallows the exception and keeps the partial results, so one bad
        /// page cannot abort the read; the font dictionary-to-metrics cache is shared per dictionary
        /// reference to avoid reparsing the same font.</remarks>
        public static void Extract(PdfPage page)
        {
            List<PdfVectorSegment> segments = new List<PdfVectorSegment>();
            PdfFragmentCollector collector = new PdfFragmentCollector(page.Height);
            Dictionary<PdfDictionary, PdfTextFont> fonts = new Dictionary<PdfDictionary, PdfTextFont>();
            PdfContentInterpreter interpreter = new PdfContentInterpreter(page.Document, page.Height, collector, segments, fonts);
            try
            {
                interpreter.Run(page.GetContentBytes(), page.Resources, null, 0);
            }
            catch (Exception)
            {
            }

            List<PdfTextLine> lines = BuildLines(collector.Fragments);
            List<PdfTable> tables = PdfTableDetector.Detect(collector.Fragments, segments, page.Height);
            page.SetExtraction(lines, tables);
        }

        /// <summary>
        /// 把文本片段按基线与垂直重叠归并成行，再按从上到下、从左到右排序。
        /// Merges text fragments into lines by baseline and vertical overlap, then sorts top-to-bottom and
        /// left-to-right.
        /// </summary>
        /// <param name="fragments">待归并的片段，null 或空文本项会被剔除。The fragments to merge; null or empty-text entries are dropped.</param>
        /// <returns>排好序的文本行。The sorted text lines.</returns>
        /// <remarks>同一行的判定条件为：基线差不超过随字号放缩的容差，且两者包围盒垂直重叠至少达到较矮者高度的一半。
        /// Two fragments share a line when their baseline difference is within a font-size-scaled tolerance
        /// and their boxes overlap vertically by at least half the shorter height.</remarks>
        internal static List<PdfTextLine> BuildLines(List<PdfTextFragment> fragments)
        {
            List<PdfTextFragment> sorted = new List<PdfTextFragment>();
            for (int i = 0; i < fragments.Count; i++)
            {
                PdfTextFragment fragment = fragments[i];
                if (fragment == null || string.IsNullOrEmpty(fragment.Text)) continue;
                sorted.Add(fragment);
            }
            sorted.Sort(delegate(PdfTextFragment a, PdfTextFragment b)
            {
                int byY = a.Y.CompareTo(b.Y);
                return byY != 0 ? byY : a.X.CompareTo(b.X);
            });

            List<LineBuilder> builders = new List<LineBuilder>();
            for (int i = 0; i < sorted.Count; i++)
            {
                PdfTextFragment fragment = sorted[i];
                double tolerance = Math.Max(2.0, Math.Max(fragment.FontSize, 6) * 0.55);
                LineBuilder target = null;
                for (int b = 0; b < builders.Count; b++)
                {
                    LineBuilder candidate = builders[b];
                    if (Math.Abs(candidate.Angle - fragment.Angle) > 0.05) continue;
                    if (Math.Abs(candidate.Baseline - fragment.Baseline) > tolerance) continue;
                    double overlap = Math.Min(candidate.Bottom, fragment.Bottom) - Math.Max(candidate.Top, fragment.Y);
                    double minimumHeight = Math.Min(candidate.Bottom - candidate.Top, fragment.Height);
                    if (minimumHeight > 0 && overlap >= minimumHeight * 0.5)
                    {
                        target = candidate;
                        break;
                    }
                }
                if (target == null)
                {
                    target = new LineBuilder();
                    builders.Add(target);
                }
                target.Add(fragment);
            }

            List<PdfTextLine> lines = new List<PdfTextLine>();
            for (int i = 0; i < builders.Count; i++)
            {
                PdfTextLine line = builders[i].Build();
                if (line != null && !string.IsNullOrEmpty(line.Text)) lines.Add(line);
            }
            lines.Sort(delegate(PdfTextLine a, PdfTextLine b)
            {
                int byY = a.Y.CompareTo(b.Y);
                return byY != 0 ? byY : a.X.CompareTo(b.X);
            });
            return lines;
        }

        /// <summary>正在组装中的一行文本，增量维护基线、上下边界。A text line under construction, maintaining its baseline and top/bottom bounds incrementally.</summary>
        private sealed class LineBuilder
        {
            /// <summary>已并入的片段集合。The fragments merged so far.</summary>
            private readonly List<PdfTextFragment> _fragments = new List<PdfTextFragment>();

            /// <summary>已纳入片段的基线均值。The running mean baseline of the included fragments.</summary>
            public double Baseline { get; private set; }
            /// <summary>行流向角度，弧度制，由首个片段确定。The line's flow angle in radians, taken from the first fragment.</summary>
            public double Angle { get; private set; }
            /// <summary>包围盒上边界。The top edge of the bounding box.</summary>
            public double Top { get; private set; }
            /// <summary>包围盒下边界。The bottom edge of the bounding box.</summary>
            public double Bottom { get; private set; }

            /// <summary>
            /// 并入一个片段，首片段决定基线初值，其后取滑动平均并扩展上下边界。
            /// Merges a fragment: the first one seeds the baseline, later ones update it by running average
            /// while expanding the top/bottom bounds.
            /// </summary>
            /// <param name="fragment">待并入的片段。The fragment to merge in.</param>
            public void Add(PdfTextFragment fragment)
            {
                if (_fragments.Count == 0)
                {
                    Baseline = fragment.Baseline;
                    Top = fragment.Y;
                    Bottom = fragment.Bottom;
                }
                else
                {
                    Baseline = (Baseline * _fragments.Count + fragment.Baseline) / (_fragments.Count + 1);
                    Top = Math.Min(Top, fragment.Y);
                    Bottom = Math.Max(Bottom, fragment.Bottom);
                }
                _fragments.Add(fragment);
            }

            /// <summary>
            /// 按横坐标排序后拼接片段文本，并依据片段间距决定是否补空格。
            /// Sorts fragments by X, concatenates their text and inserts spaces where the gap warrants.
            /// </summary>
            /// <returns>组装好的文本行；文本为空时返回 null。The assembled line; null when the text is empty.</returns>
            /// <remarks>间距阈值随字号放缩并限制在 1.2–5.0 之间，避免把字距拆成词，也避免粘连相邻词。
            /// The gap threshold scales with the font size and is clamped to 1.2-5.0, so letter spacing is
            /// not split into words and neighboring words are not glued together.</remarks>
            public PdfTextLine Build()
            {
                _fragments.Sort(delegate(PdfTextFragment a, PdfTextFragment b)
                {
                    return a.X.CompareTo(b.X);
                });

                StringBuilder text = new StringBuilder();
                double right = 0;
                for (int i = 0; i < _fragments.Count; i++)
                {
                    PdfTextFragment fragment = _fragments[i];
                    if (i > 0)
                    {
                        double gap = fragment.X - right;
                        double threshold = Math.Max(1.2, Math.Min(5.0, Math.Max(fragment.FontSize, 6) * 0.22));
                        char previous = text.Length > 0 ? text[text.Length - 1] : '\0';
                        bool needsSpace = gap > threshold
                            && !char.IsWhiteSpace(previous)
                            && fragment.Text.Length > 0
                            && !char.IsWhiteSpace(fragment.Text[0]);
                        if (needsSpace) text.Append(' ');
                    }
                    text.Append(fragment.Text);
                    right = fragment.Right;
                }

                string value = text.ToString();
                if (string.IsNullOrEmpty(value)) return null;

                PdfTextLine line = new PdfTextLine();
                line.Text = value;
                line.X = _fragments[0].X;
                line.Y = Top;
                line.Width = Math.Max(0, _fragments[_fragments.Count - 1].Right - line.X);
                line.Height = Math.Max(0, Bottom - Top);
                line.Fragments = _fragments.ToArray();
                return line;
            }
        }
    }

    /// <summary>
    /// 表格识别器：把页面上的矢量线段聚类成网格，再把文本片段按中心点分配到单元格。
    /// Table detector: clusters the page's vector segments into grids, then assigns text fragments to cells
    /// by their center points.
    /// </summary>
    /// <remarks>坐标系在此处已由 PDF 的左下原点翻转为以页面左上角为原点（y = pageHeight - y），
    /// 因此输出的表格坐标与 <see cref="PdfTextFragment"/> 一致。
    /// The coordinate system has already been flipped here from the PDF bottom-left origin to a page
    /// top-left origin (y = pageHeight - y), so table coordinates match <see cref="PdfTextFragment"/>.</remarks>
    internal static class PdfTableDetector
    {
        private const long MaxTableCells = 250000;
        private const long MaxGridIntersections = 4000000;

        /// <summary>一条网格线：Position 为垂直于自身的坐标，Start/End 为沿线方向的跨度。One grid line: Position is the perpendicular coordinate, Start/End the span along the line.</summary>
        private sealed class GridLine
        {
            /// <summary>网格线位置（横线为 y、竖线为 x）。The grid line position (y for horizontals, x for verticals).</summary>
            public double Position;
            /// <summary>线段起点（横线取 x、竖线取 y）。The start of the segment (x for horizontals, y for verticals).</summary>
            public double Start;
            /// <summary>线段终点。The end of the segment.</summary>
            public double End;
        }

        /// <summary>并查集，用于把相互交叉的横线与竖线合并为同一张表格。A union-find structure merging crossing horizontals and verticals into the same table.</summary>
        private sealed class UnionFind
        {
            /// <summary>父指针数组，根节点指向自身。The parent array; roots point to themselves.</summary>
            private readonly int[] _parent;

            /// <summary>
            /// 建立每个元素自成一组的并查集。
            /// Creates a union-find where every element starts as its own set.
            /// </summary>
            /// <param name="count">元素总数（横线数 + 竖线数）。The total element count (horizontals + verticals).</param>
            public UnionFind(int count)
            {
                _parent = new int[count];
                for (int i = 0; i < count; i++) _parent[i] = i;
            }

            /// <summary>
            /// 查找根节点，路径上做了路径减半以加速重复查询。
            /// Finds the root, applying path halving to speed up repeated queries.
            /// </summary>
            /// <param name="value">元素下标。The element index.</param>
            /// <returns>所在集合的根下标。The root index of its set.</returns>
            public int Find(int value)
            {
                while (_parent[value] != value)
                {
                    _parent[value] = _parent[_parent[value]];
                    value = _parent[value];
                }
                return value;
            }

            /// <summary>
            /// 合并两个元素所在的集合。
            /// Unions the sets containing the two elements.
            /// </summary>
            /// <param name="a">第一个元素下标。The first element index.</param>
            /// <param name="b">第二个元素下标。The second element index.</param>
            public void Union(int a, int b)
            {
                int rootA = Find(a);
                int rootB = Find(b);
                if (rootA != rootB) _parent[rootB] = rootA;
            }
        }

        /// <summary>
        /// 从线段集合中识别出所有表格。
        /// Recognizes all tables from the segment collection.
        /// </summary>
        /// <param name="fragments">页面全部文本片段，用于填充单元格。All page text fragments, used to fill the cells.</param>
        /// <param name="segments">页面矢量线段，须为非空才能形成网格。The page vector segments; a non-empty list is required to form a grid.</param>
        /// <param name="pageHeight">页面高度，用于翻转 y 坐标。The page height, used to flip the y coordinates.</param>
        /// <returns>按位置排序的表格列表；完全由内线套表的场景只保留外层表格。The tables sorted by position; when one table is fully nested inside another, only the outer one is kept.</returns>
        /// <remarks>近水平/垂直的判定阈值为 0.6 且长度至少 6；分组后还要求行列各不少于 2 条线、
        /// 面积不小于 2 格、宽高分别不少于 24/12，且至少含一个非空单元格，以避免把装饰线条误判为表格。
        /// Near-horizontal/vertical means within 0.6 and at least 6 long; after grouping, each table needs at
        /// least 2 lines in both directions, an area of at least 2 cells, size at least 24x12 and one
        /// non-empty cell, so decorative strokes are not mistaken for tables.</remarks>
        public static List<PdfTable> Detect(List<PdfTextFragment> fragments, List<PdfVectorSegment> segments, double pageHeight)
        {
            List<PdfTable> result = new List<PdfTable>();
            if (segments == null || segments.Count == 0) return result;

            List<GridLine> horizontals = new List<GridLine>();
            List<GridLine> verticals = new List<GridLine>();
            for (int i = 0; i < segments.Count; i++)
            {
                PdfVectorSegment segment = segments[i];
                double x1 = segment.X1;
                double y1 = pageHeight - segment.Y1;
                double x2 = segment.X2;
                double y2 = pageHeight - segment.Y2;
                double dx = Math.Abs(x2 - x1);
                double dy = Math.Abs(y2 - y1);
                if (dy <= 0.6 && dx >= 6)
                {
                    GridLine line = new GridLine();
                    line.Position = (y1 + y2) / 2;
                    line.Start = Math.Min(x1, x2);
                    line.End = Math.Max(x1, x2);
                    horizontals.Add(line);
                }
                else if (dx <= 0.6 && dy >= 6)
                {
                    GridLine line = new GridLine();
                    line.Position = (x1 + x2) / 2;
                    line.Start = Math.Min(y1, y2);
                    line.End = Math.Max(y1, y2);
                    verticals.Add(line);
                }
            }
            if (horizontals.Count == 0 || verticals.Count == 0) return result;

            horizontals = Merge(horizontals);
            verticals = Merge(verticals);
            if (horizontals.Count < 2 || verticals.Count < 2) return result;

            // 交叉遍历是 O(h*v)，先按上限拒掉恶意网格。
            // The intersection walk is O(h*v), so an oversized grid is rejected up front.
            if ((long)horizontals.Count * verticals.Count > MaxGridIntersections) return result;
            UnionFind union = new UnionFind(horizontals.Count + verticals.Count);
            for (int h = 0; h < horizontals.Count; h++)
            {
                GridLine horizontal = horizontals[h];
                for (int v = 0; v < verticals.Count; v++)
                {
                    GridLine vertical = verticals[v];
                    if (vertical.Position >= horizontal.Start - 1.5
                        && vertical.Position <= horizontal.End + 1.5
                        && horizontal.Position >= vertical.Start - 1.5
                        && horizontal.Position <= vertical.End + 1.5)
                    {
                        union.Union(h, horizontals.Count + v);
                    }
                }
            }

            Dictionary<int, List<int>> horizontalGroups = new Dictionary<int, List<int>>();
            Dictionary<int, List<int>> verticalGroups = new Dictionary<int, List<int>>();
            for (int h = 0; h < horizontals.Count; h++)
            {
                int root = union.Find(h);
                List<int> list;
                if (!horizontalGroups.TryGetValue(root, out list))
                {
                    list = new List<int>();
                    horizontalGroups[root] = list;
                }
                list.Add(h);
            }
            for (int v = 0; v < verticals.Count; v++)
            {
                int root = union.Find(horizontals.Count + v);
                List<int> list;
                if (!verticalGroups.TryGetValue(root, out list))
                {
                    list = new List<int>();
                    verticalGroups[root] = list;
                }
                list.Add(v);
            }

            List<PdfTable> candidates = new List<PdfTable>();
            foreach (KeyValuePair<int, List<int>> pair in horizontalGroups)
            {
                List<int> horizontalGroup;
                List<int> verticalGroup;
                if (!verticalGroups.TryGetValue(pair.Key, out verticalGroup)) continue;
                horizontalGroup = pair.Value;
                if (horizontalGroup.Count < 2 || verticalGroup.Count < 2) continue;

                List<double> rows = new List<double>();
                List<double> columns = new List<double>();
                for (int i = 0; i < horizontalGroup.Count; i++) rows.Add(horizontals[horizontalGroup[i]].Position);
                for (int i = 0; i < verticalGroup.Count; i++) columns.Add(verticals[verticalGroup[i]].Position);
                rows.Sort();
                columns.Sort();
                rows = Cluster(rows, 1.5);
                columns = Cluster(columns, 1.5);
                if (rows.Count < 2 || columns.Count < 2) continue;
                if ((rows.Count - 1) * (columns.Count - 1) < 2) continue;

                double left = columns[0];
                double right = columns[columns.Count - 1];
                double top = rows[0];
                double bottom = rows[rows.Count - 1];
                if (right - left < 24 || bottom - top < 12) continue;

                PdfTable table = BuildTable(fragments, columns, rows);
                if (table != null) candidates.Add(table);
            }

            candidates.Sort(delegate(PdfTable a, PdfTable b)
            {
                int byY = a.Y.CompareTo(b.Y);
                return byY != 0 ? byY : a.X.CompareTo(b.X);
            });

            for (int i = 0; i < candidates.Count; i++)
            {
                bool nested = false;
                for (int j = 0; j < candidates.Count; j++)
                {
                    if (i == j) continue;
                    PdfTable inner = candidates[i];
                    PdfTable outer = candidates[j];
                    if (outer.Width * outer.Height < inner.Width * inner.Height * 1.5) continue;
                    if (inner.X >= outer.X - 1 && inner.Y >= outer.Y - 1
                        && inner.X + inner.Width <= outer.X + outer.Width + 1
                        && inner.Y + inner.Height <= outer.Y + outer.Height + 1)
                    {
                        nested = true;
                        break;
                    }
                }
                if (!nested) result.Add(candidates[i]);
            }
            return result;
        }

        /// <summary>
        /// 合并位置几乎相同且沿线方向首尾相接的网格线段，把被拆成多段的同一条边还原为一条。
        /// Merges grid segments that sit at almost the same position and touch end to end, restoring one
        /// edge that was drawn as multiple segments.
        /// </summary>
        /// <param name="lines">同类（全为横线或全为竖线）的线段列表，会被就地排序。Segments of one kind (all horizontals or all verticals); sorted in place.</param>
        /// <returns>合并后的线段副本列表。The merged list of segment copies.</returns>
        private static List<GridLine> Merge(List<GridLine> lines)
        {
            lines.Sort(delegate(GridLine a, GridLine b)
            {
                int byPosition = a.Position.CompareTo(b.Position);
                return byPosition != 0 ? byPosition : a.Start.CompareTo(b.Start);
            });

            List<GridLine> merged = new List<GridLine>();
            for (int i = 0; i < lines.Count; i++)
            {
                GridLine line = lines[i];
                GridLine last = merged.Count > 0 ? merged[merged.Count - 1] : null;
                if (last != null
                    && Math.Abs(last.Position - line.Position) <= 1.2
                    && line.Start <= last.End + 3)
                {
                    last.Position = (last.Position + line.Position) / 2;
                    last.End = Math.Max(last.End, line.End);
                    continue;
                }
                GridLine copy = new GridLine();
                copy.Position = line.Position;
                copy.Start = line.Start;
                copy.End = line.End;
                merged.Add(copy);
            }
            return merged;
        }

        /// <summary>
        /// 把已排序的坐标聚为若干档，取每档均值作为网格位置，消除同一条边的抖动误差。
        /// Buckets the sorted coordinates and uses each bucket's mean as the grid position, absorbing the
        /// jitter of an edge drawn at slightly different offsets.
        /// </summary>
        /// <param name="values">升序排列的坐标序列。The ascending coordinate sequence.</param>
        /// <param name="tolerance">与当前档均值的最大允许偏差。The maximum allowed deviation from the current bucket mean.</param>
        /// <returns>每档的均值列表，数量即网格线数。The list of bucket means, one per grid line.</returns>
        private static List<double> Cluster(List<double> values, double tolerance)
        {
            List<double> result = new List<double>();
            double sum = 0;
            int count = 0;
            for (int i = 0; i < values.Count; i++)
            {
                if (count > 0 && values[i] - sum / count > tolerance)
                {
                    result.Add(sum / count);
                    sum = 0;
                    count = 0;
                }
                sum += values[i];
                count++;
            }
            if (count > 0) result.Add(sum / count);
            return result;
        }

        /// <summary>
        /// 依据网格线构建表格对象，把片段按其中心点落入的行列归入单元格。
        /// Builds a table from the grid lines, assigning each fragment to the cell containing its center.
        /// </summary>
        /// <param name="fragments">页面文本片段。The page text fragments.</param>
        /// <param name="columns">升序的竖线位置，长度为列数 + 1。Ascending vertical line positions, count equals columns + 1.</param>
        /// <param name="rows">升序的横线位置，长度为行数 + 1。Ascending horizontal line positions, count equals rows + 1.</param>
        /// <returns>表格对象；网格退化（行列数不足）或全部单元格为空时返回 null。The table; null when the grid is degenerate or all cells are empty.</returns>
        /// <remarks>单元格内可能有多行文本，复用 <see cref="PdfPageExtractor.BuildLines"/> 归行后以换行连接。
        /// A cell may hold several lines; the fragments are regrouped with
        /// <see cref="PdfPageExtractor.BuildLines"/> and joined with newlines.</remarks>
        private static PdfTable BuildTable(List<PdfTextFragment> fragments, List<double> columns, List<double> rows)
        {
            int columnCount = columns.Count - 1;
            int rowCount = rows.Count - 1;
            if (columnCount <= 0 || rowCount <= 0) return null;
            // 样本数乘积直接决定两个矩形数组的大小，必须在分配前检查。
            // The product of the two grid sample counts sizes two rectangular arrays, so it is checked
            // before anything is allocated; a hostile page can otherwise suggest a 100k x 100k matrix.
            if ((long)rowCount * columnCount > MaxTableCells) return null;

            double left = columns[0];
            double top = rows[0];
            double right = columns[columns.Count - 1];
            double bottom = rows[rows.Count - 1];

            List<PdfTextFragment>[,] cells = new List<PdfTextFragment>[rowCount, columnCount];
            for (int i = 0; i < fragments.Count; i++)
            {
                PdfTextFragment fragment = fragments[i];
                if (fragment == null || string.IsNullOrEmpty(fragment.Text)) continue;
                double centerX = fragment.X + fragment.Width / 2;
                double centerY = fragment.Y + fragment.Height / 2;
                if (centerX < left || centerX > right || centerY < top || centerY > bottom) continue;

                int column = FindIndex(columns, centerX);
                int row = FindIndex(rows, centerY);
                if (column < 0 || row < 0) continue;
                if (cells[row, column] == null) cells[row, column] = new List<PdfTextFragment>();
                cells[row, column].Add(fragment);
            }

            string[,] values = new string[rowCount, columnCount];
            bool hasText = false;
            for (int row = 0; row < rowCount; row++)
            {
                for (int column = 0; column < columnCount; column++)
                {
                    List<PdfTextFragment> cellFragments = cells[row, column];
                    string text = string.Empty;
                    if (cellFragments != null && cellFragments.Count > 0)
                    {
                        List<PdfTextLine> lines = PdfPageExtractor.BuildLines(cellFragments);
                        StringBuilder builder = new StringBuilder();
                        for (int i = 0; i < lines.Count; i++)
                        {
                            if (i > 0) builder.Append('\n');
                            builder.Append(lines[i].Text);
                        }
                        text = builder.ToString();
                        if (text.Length > 0) hasText = true;
                    }
                    values[row, column] = text;
                }
            }
            if (!hasText) return null;

            PdfTable table = new PdfTable();
            table.RowCount = rowCount;
            table.ColumnCount = columnCount;
            table.X = left;
            table.Y = top;
            table.Width = right - left;
            table.Height = bottom - top;
            table.Cells = values;
            return table;
        }

        /// <summary>
        /// 在升序网格边界中找到给定点所属的区间下标。
        /// Finds the index of the grid interval that contains the given point in ascending boundaries.
        /// </summary>
        /// <param name="boundaries">升序边界坐标，长度为区间数 + 1。The ascending boundary coordinates, count equals intervals + 1.</param>
        /// <param name="value">待定位的坐标。The coordinate to locate.</param>
        /// <returns>区间下标；点落在所有区间之外（含 1.5 的容差）时返回 -1。The interval index; -1 when the point lies outside every interval (with a 1.5 tolerance).</returns>
        private static int FindIndex(List<double> boundaries, double value)
        {
            for (int i = 0; i < boundaries.Count - 1; i++)
            {
                if (value < boundaries[i + 1] || i == boundaries.Count - 2)
                {
                    if (value < boundaries[i] - 1.5 || value > boundaries[i + 1] + 1.5) return -1;
                    return i;
                }
            }
            return -1;
        }
    }
}
