using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 把图表模型序列化为 c:chartSpace 部件：解析区域引用得到缓存值，按图表类型输出
    /// 符合 CT_BarChart / CT_LineChart / CT_PieChart / CT_DoughnutChart 子元素顺序的 XML。
    /// Serializes a chart model into a c:chartSpace part: range references are resolved into caches and
    /// the child order required by CT_BarChart, CT_LineChart, CT_PieChart and CT_DoughnutChart is emitted.
    /// </summary>
    internal static class XlsxChartParts
    {
        /// <summary>图表命名空间。 Chart namespace.</summary>
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        /// <summary>绘图命名空间。 Drawing namespace.</summary>
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        /// <summary>类目坐标轴 Id。 Category axis identifier.</summary>
        private const string CatAxisId = "111111111";
        /// <summary>数值坐标轴 Id。 Value axis identifier.</summary>
        private const string ValAxisId = "222222222";
        /// <summary>组合图次类目轴 Id。 Secondary category-axis identifier used by combo charts.</summary>
        private const string SecondaryCatAxisId = "333333333";
        /// <summary>组合图次数值轴 Id。 Secondary value-axis identifier used by combo charts.</summary>
        private const string SecondaryValAxisId = "444444444";
        /// <summary>散点/气泡图 X 轴 Id。 Identifier of the scatter and bubble X axis.</summary>
        private const string ScatterXAxisId = "555555555";
        /// <summary>散点/气泡图 Y 轴 Id。 Identifier of the scatter and bubble Y axis.</summary>
        private const string ScatterYAxisId = "666666666";

        /// <summary>一次区域或字面量解析的结果。 The result of resolving either a range or literal data.</summary>
        private sealed class ResolvedData
        {
            /// <summary>绝对引用文本；来自字面量时为 null。 Absolute reference text; null when the data came from literals.</summary>
            public string Reference;
            /// <summary>文本取值。 Text values.</summary>
            public List<string> Texts = new List<string>();
            /// <summary>数值取值。 Numeric values.</summary>
            public List<double> Numbers = new List<double>();
            /// <summary>每个数据点是否有值。 Whether each point carries a value.</summary>
            public List<bool> Present = new List<bool>();
            /// <summary>文本是否全部可解析为数字。 Whether every text parses as a number.</summary>
            public bool AllNumeric;
            /// <summary>缓存使用的数字格式代码。 Number-format code used by the cache.</summary>
            public string FormatCode = "General";
        }

        /// <summary>生成完整的 chartSpace 部件文档。 Builds the complete chartSpace part document.</summary>
        /// <param name="chart">图表模型。 Chart model.</param>
        /// <param name="workbook">所属工作簿（用于解析区域引用）。 Owning workbook, used to resolve range references.</param>
        /// <param name="host">承载图表的工作表。 Worksheet hosting the chart.</param>
        /// <returns>chartSpace 文档。 The chartSpace document.</returns>
        public static XDocument Build(ExcelChart chart, OfficeWorkbook workbook, Worksheet host)
        {
            bool hasAxes = chart.Type != ExcelChartType.Pie && chart.Type != ExcelChartType.Doughnut;
            XElement chartElement = new XElement(C + "chart");
            if (!string.IsNullOrEmpty(chart.Title))
            {
                chartElement.Add(BuildTitle(chart.Title));
                chartElement.Add(new XElement(C + "autoTitleDeleted", new XAttribute("val", "0")));
            }
            else
            {
                chartElement.Add(new XElement(C + "autoTitleDeleted", new XAttribute("val", "1")));
            }
            chartElement.Add(BuildPlotArea(chart, workbook, host, hasAxes));
            if (chart.ShowLegend && chart.Series.Count > 0 && chart.Type != ExcelChartType.Pie)
            {
                string position = string.IsNullOrEmpty(chart.LegendPosition) ? "b" : chart.LegendPosition;
                chartElement.Add(new XElement(C + "legend",
                    new XElement(C + "legendPos", new XAttribute("val", position)),
                    new XElement(C + "overlay", new XAttribute("val", "0"))));
            }
            chartElement.Add(new XElement(C + "plotVisOnly", new XAttribute("val", "1")));
            chartElement.Add(new XElement(C + "dispBlanksAs", new XAttribute("val", "gap")));

            XElement root = new XElement(C + "chartSpace",
                new XAttribute(XNamespace.Xmlns + "c", C.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "r", Xlsx.R.NamespaceName),
                chartElement);
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
        }

        /// <summary>生成图表标题元素。 Builds the chart title element.</summary>
        /// <param name="title">标题文本。 Title text.</param>
        /// <returns>title 元素。 The title element.</returns>
        private static XElement BuildTitle(string title)
        {
            return new XElement(C + "title",
                new XElement(C + "tx",
                    new XElement(C + "rich",
                        new XElement(A + "bodyPr"),
                        new XElement(A + "lstStyle"),
                        new XElement(A + "p",
                            new XElement(A + "r",
                                new XElement(A + "rPr",
                                    new XAttribute("lang", "zh-CN"),
                                    new XAttribute("sz", "1200"),
                                    new XAttribute("b", "1"),
                                    new XAttribute("dirty", "0")),
                                new XElement(A + "t", title))))),
                new XElement(C + "overlay", new XAttribute("val", "0")));
        }

        /// <summary>生成绘图区元素：布局、按类型与坐标轴分组的图表组集合与（如需）坐标轴。 Builds the plot area: layout, the chart groups bucketed by type and axis, and, when needed, the axes.</summary>
        /// <param name="chart">图表模型。 Chart model.</param>
        /// <param name="workbook">工作簿。 Workbook.</param>
        /// <param name="host">承载表。 Host worksheet.</param>
        /// <param name="hasAxes">是否含坐标轴。 Whether the chart has axes.</param>
        /// <returns>plotArea 元素。 The plotArea element.</returns>
        private static XElement BuildPlotArea(ExcelChart chart, OfficeWorkbook workbook, Worksheet host, bool hasAxes)
        {
            // 先把系列按"有效类型 + 是否次轴"分桶，组合图因此产出多个图表组，元素顺序保持系列原始次序。
            // Bucket series by effective type plus secondary-axis flag, so a combo chart yields several groups while series keep their original order.
            List<int[]> buckets = new List<int[]>();
            List<ExcelChartType> bucketTypes = new List<ExcelChartType>();
            List<bool> bucketSecondary = new List<bool>();
            for (int i = 0; i < chart.Series.Count; i++)
            {
                ExcelChartSeries series = chart.Series[i];
                ExcelChartType effective = series.ChartType ?? chart.Type;
                int found = -1;
                for (int b = 0; b < buckets.Count; b++)
                {
                    if (bucketTypes[b] == effective && bucketSecondary[b] == series.SecondaryAxis) { found = b; break; }
                }
                if (found < 0)
                {
                    buckets.Add(new int[0]);
                    bucketTypes.Add(effective);
                    bucketSecondary.Add(series.SecondaryAxis);
                    found = buckets.Count - 1;
                }
                int[] list = buckets[found];
                int[] grown = new int[list.Length + 1];
                Array.Copy(list, grown, list.Length);
                grown[list.Length] = i;
                buckets[found] = grown;
            }

            bool scatterFamily = IsScatterFamily(chart.Type);
            for (int b = 0; b < bucketTypes.Count; b++)
            {
                if (IsScatterFamily(bucketTypes[b])) { scatterFamily = true; }
            }
            bool anyScatter = false, anyCategory = false;
            for (int b = 0; b < bucketTypes.Count; b++)
            {
                if (IsScatterFamily(bucketTypes[b])) anyScatter = true; else anyCategory = true;
            }

            XElement plotArea = new XElement(C + "plotArea", new XElement(C + "layout"));
            bool needsSecondary = false;
            for (int b = 0; b < bucketTypes.Count; b++)
            {
                List<XElement> members = new List<XElement>();
                int[] indexes = buckets[b];
                for (int k = 0; k < indexes.Length; k++)
                {
                    members.Add(BuildSeries(chart.Series[indexes[k]], indexes[k], bucketTypes[b], chart, workbook, host));
                }
                if (bucketSecondary[b]) needsSecondary = true;
                plotArea.Add(BuildGroup(bucketTypes[b], members, chart, bucketSecondary[b], scatterFamily && anyScatter && !anyCategory));
            }

            if (scatterFamily && anyScatter && !anyCategory)
            {
                plotArea.Add(BuildScatterAxis(ScatterXAxisId, ScatterYAxisId, "b", chart.ValueAxisFormat));
                plotArea.Add(BuildScatterAxis(ScatterYAxisId, ScatterXAxisId, "l", chart.ValueAxisFormat));
                return plotArea;
            }
            if (!hasAxes) return plotArea;
            plotArea.Add(BuildCategoryAxis(CatAxisId, ValAxisId, "b", false));
            plotArea.Add(BuildValueAxis(ValAxisId, CatAxisId, "l", chart.ValueAxisFormat));
            if (needsSecondary)
            {
                plotArea.Add(BuildCategoryAxis(SecondaryCatAxisId, SecondaryValAxisId, "b", true));
                plotArea.Add(BuildValueAxis(SecondaryValAxisId, SecondaryCatAxisId, "r", chart.ValueAxisFormat));
            }
            return plotArea;
        }

        /// <summary>按图表类型生成一个图表组元素（barChart、lineChart、scatterChart、radarChart 等）。 Builds one chart-group element for a type (barChart, lineChart, scatterChart, radarChart and so on).</summary>
        /// <param name="type">该组的图表类型。 The group's chart type.</param>
        /// <param name="series">已构建的系列元素。 Built series elements.</param>
        /// <param name="chart">所属图表。 Owning chart.</param>
        /// <param name="secondary">是否放到次坐标轴。 Whether the group sits on the secondary axis.</param>
        /// <param name="pureScatter">整张图是否为纯散点族（使用 X/Y 双数值轴）。 Whether the whole chart is scatter-family (uses two value axes).</param>
        /// <returns>图表组元素。 The chart-group element.</returns>
        private static XElement BuildGroup(ExcelChartType type, List<XElement> series, ExcelChart chart, bool secondary, bool pureScatter)
        {
            string catId = secondary ? SecondaryCatAxisId : CatAxisId;
            string valId = secondary ? SecondaryValAxisId : ValAxisId;
            switch (type)
            {
                case ExcelChartType.Bar:
                    return Group(new XElement(C + "barChart"), "bar", "clustered", series, catId, valId);
                case ExcelChartType.StackedBar:
                    return Group(new XElement(C + "barChart"), "bar", "stacked", series, catId, valId);
                case ExcelChartType.PercentStackedBar:
                    return Group(new XElement(C + "barChart"), "bar", "percentStacked", series, catId, valId);
                case ExcelChartType.StackedColumn:
                    return Group(new XElement(C + "barChart"), "col", "stacked", series, catId, valId);
                case ExcelChartType.PercentStackedColumn:
                    return Group(new XElement(C + "barChart"), "col", "percentStacked", series, catId, valId);
                case ExcelChartType.StackedArea:
                    return PlainGroup(C + "areaChart", "stacked", series, false, catId, valId, null);
                case ExcelChartType.PercentStackedArea:
                    return PlainGroup(C + "areaChart", "percentStacked", series, false, catId, valId, null);
                case ExcelChartType.Area:
                    return PlainGroup(C + "areaChart", "standard", series, false, catId, valId, null);
                case ExcelChartType.StackedLine:
                    return PlainGroup(C + "lineChart", "stacked", series, true, catId, valId, null);
                case ExcelChartType.PercentStackedLine:
                    return PlainGroup(C + "lineChart", "percentStacked", series, true, catId, valId, null);
                case ExcelChartType.Line:
                case ExcelChartType.LineWithMarkers:
                    return PlainGroup(C + "lineChart", "standard", series, true, catId, valId, null);
                case ExcelChartType.Pie:
                {
                    XElement group = new XElement(C + "pieChart");
                    group.Add(new XElement(C + "varyColors", new XAttribute("val", "0")));
                    for (int i = 0; i < series.Count; i++) group.Add(series[i]);
                    group.Add(new XElement(C + "firstSliceAng", new XAttribute("val", "0")));
                    return group;
                }
                case ExcelChartType.Doughnut:
                {
                    XElement group = new XElement(C + "doughnutChart");
                    group.Add(new XElement(C + "grouping", new XAttribute("val", "standard")));
                    group.Add(new XElement(C + "varyColors", new XAttribute("val", "0")));
                    for (int i = 0; i < series.Count; i++) group.Add(series[i]);
                    group.Add(new XElement(C + "firstSliceAng", new XAttribute("val", "0")));
                    int hole = chart.HoleSize < 10 ? 10 : (chart.HoleSize > 90 ? 90 : chart.HoleSize);
                    group.Add(new XElement(C + "holeSize", new XAttribute("val", hole.ToString(CultureInfo.InvariantCulture))));
                    return group;
                }
                case ExcelChartType.ScatterMarkers:
                case ExcelChartType.ScatterSmoothLines:
                case ExcelChartType.ScatterLinesMarkers:
                    return ScatterGroup(type, series);
                case ExcelChartType.Bubble:
                    return BubbleGroup(series);
                case ExcelChartType.Radar:
                    return RadarGroup(series, "standard", catId, valId);
                case ExcelChartType.RadarWithMarkers:
                    return RadarGroup(series, "marker", catId, valId);
                case ExcelChartType.RadarFilled:
                    return RadarGroup(series, "filled", catId, valId);
                default:
                    return Group(new XElement(C + "barChart"), "col", "clustered", series, catId, valId);
            }
        }

        /// <summary>组装柱形/条形图表组元素。 Assembles the column or bar chart group element.</summary>
        private static XElement Group(XElement group, string direction, string grouping, List<XElement> series, string catId, string valId)
        {
            group.Add(new XElement(C + "barDir", new XAttribute("val", direction)));
            group.Add(new XElement(C + "grouping", new XAttribute("val", grouping)));
            group.Add(new XElement(C + "varyColors", new XAttribute("val", "0")));
            for (int i = 0; i < series.Count; i++) group.Add(series[i]);
            group.Add(new XElement(C + "gapWidth", new XAttribute("val", "150")));
            string overlap = grouping == "clustered" ? "-27" : "100";
            group.Add(new XElement(C + "overlap", new XAttribute("val", overlap)));
            group.Add(new XElement(C + "axId", new XAttribute("val", catId)));
            group.Add(new XElement(C + "axId", new XAttribute("val", valId)));
            return group;
        }

        /// <summary>组装折线或面积图表组元素（顺序为 grouping、varyColors、系列、线组附加、axId）。 Assembles a line or area group (grouping, varyColors, series, type extras, axIds).</summary>
        private static XElement PlainGroup(XName name, string grouping, List<XElement> series, bool isLine, string catId, string valId, string unused)
        {
            XElement group = new XElement(name);
            group.Add(new XElement(C + "grouping", new XAttribute("val", grouping)));
            group.Add(new XElement(C + "varyColors", new XAttribute("val", "0")));
            for (int i = 0; i < series.Count; i++) group.Add(series[i]);
            if (isLine) group.Add(new XElement(C + "marker", new XAttribute("val", "1")));
            group.Add(new XElement(C + "axId", new XAttribute("val", catId)));
            group.Add(new XElement(C + "axId", new XAttribute("val", valId)));
            return group;
        }

        /// <summary>组装散点图表组：按子类型确定 scatterStyle，使用 X/Y 双数值轴。 Assembles a scatter group; the sub-type selects scatterStyle and the group uses the X and Y value axes.</summary>
        private static XElement ScatterGroup(ExcelChartType type, List<XElement> series)
        {
            string style = type == ExcelChartType.ScatterSmoothLines ? "smoothMarker"
                : type == ExcelChartType.ScatterLinesMarkers ? "lineMarker" : "marker";
            XElement group = new XElement(C + "scatterChart");
            group.Add(new XElement(C + "scatterStyle", new XAttribute("val", style)));
            group.Add(new XElement(C + "varyColors", new XAttribute("val", "0")));
            for (int i = 0; i < series.Count; i++) group.Add(series[i]);
            group.Add(new XElement(C + "axId", new XAttribute("val", ScatterXAxisId)));
            group.Add(new XElement(C + "axId", new XAttribute("val", ScatterYAxisId)));
            return group;
        }

        /// <summary>组装气泡图表组：使用 X/Y 双数值轴。 Assembles a bubble group on the X and Y value axes.</summary>
        private static XElement BubbleGroup(List<XElement> series)
        {
            XElement group = new XElement(C + "bubbleChart");
            group.Add(new XElement(C + "varyColors", new XAttribute("val", "0")));
            for (int i = 0; i < series.Count; i++) group.Add(series[i]);
            group.Add(new XElement(C + "bubble3D", new XAttribute("val", "0")));
            group.Add(new XElement(C + "axId", new XAttribute("val", ScatterXAxisId)));
            group.Add(new XElement(C + "axId", new XAttribute("val", ScatterYAxisId)));
            return group;
        }

        /// <summary>组装雷达图表组：按子类型确定 radarStyle，仍用类目 + 数值轴。 Assembles a radar group; the sub-type selects radarStyle and the group uses a category plus value axis.</summary>
        private static XElement RadarGroup(List<XElement> series, string radarStyle, string catId, string valId)
        {
            XElement group = new XElement(C + "radarChart");
            group.Add(new XElement(C + "radarStyle", new XAttribute("val", radarStyle)));
            group.Add(new XElement(C + "varyColors", new XAttribute("val", "0")));
            for (int i = 0; i < series.Count; i++) group.Add(series[i]);
            group.Add(new XElement(C + "axId", new XAttribute("val", catId)));
            group.Add(new XElement(C + "axId", new XAttribute("val", valId)));
            return group;
        }

        /// <summary>生成类目轴；<paramref name="hidden"/> 为 true 时删除标签，用于组合图次类目轴。 Builds a category axis; when <paramref name="hidden"/> is true the axis is deleted, used by a combo's secondary category axis.</summary>
        private static XElement BuildCategoryAxis(string id, string crossId, string position, bool hidden)
        {
            return new XElement(C + "catAx",
                new XElement(C + "axId", new XAttribute("val", id)),
                new XElement(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax"))),
                new XElement(C + "delete", new XAttribute("val", hidden ? "1" : "0")),
                new XElement(C + "axPos", new XAttribute("val", position)),
                new XElement(C + "tickLblPos", new XAttribute("val", "nextTo")),
                new XElement(C + "crossAx", new XAttribute("val", crossId)),
                new XElement(C + "auto", new XAttribute("val", "1")),
                new XElement(C + "lblAlgn", new XAttribute("val", "ctr")),
                new XElement(C + "lblOffset", new XAttribute("val", "100")),
                new XElement(C + "noMultiLvlLbl", new XAttribute("val", "0")));
        }

        /// <summary>生成数值轴：位置、主网格线与数字格式。 Builds a value axis: position, major gridlines and number format.</summary>
        private static XElement BuildValueAxis(string id, string crossId, string position, string format)
        {
            XElement axis = new XElement(C + "valAx",
                new XElement(C + "axId", new XAttribute("val", id)),
                new XElement(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax"))),
                new XElement(C + "delete", new XAttribute("val", "0")),
                new XElement(C + "axPos", new XAttribute("val", position)));
            axis.Add(new XElement(C + "majorGridlines"));
            string code = string.IsNullOrEmpty(format) ? "General" : format;
            axis.Add(new XElement(C + "numFmt",
                new XAttribute("formatCode", code),
                new XAttribute("sourceLinked", code == "General" ? "1" : "0")));
            axis.Add(new XElement(C + "tickLblPos", new XAttribute("val", "nextTo")));
            axis.Add(new XElement(C + "crossAx", new XAttribute("val", crossId)));
            axis.Add(new XElement(C + "crosses", new XAttribute("val", "autoZero")));
            axis.Add(new XElement(C + "auto", new XAttribute("val", "1")));
            axis.Add(new XElement(C + "lblOffset", new XAttribute("val", "100")));
            return axis;
        }

        /// <summary>生成散点/气泡图的一条数值轴（X 或 Y）。 Builds one value axis (X or Y) for scatter and bubble charts.</summary>
        private static XElement BuildScatterAxis(string id, string crossId, string position, string format)
        {
            XElement axis = new XElement(C + "valAx",
                new XElement(C + "axId", new XAttribute("val", id)),
                new XElement(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax"))),
                new XElement(C + "delete", new XAttribute("val", "0")),
                new XElement(C + "axPos", new XAttribute("val", position)));
            if (position == "l") axis.Add(new XElement(C + "majorGridlines"));
            string code = string.IsNullOrEmpty(format) ? "General" : format;
            axis.Add(new XElement(C + "numFmt",
                new XAttribute("formatCode", code),
                new XAttribute("sourceLinked", code == "General" ? "1" : "0")));
            axis.Add(new XElement(C + "tickLblPos", new XAttribute("val", "nextTo")));
            axis.Add(new XElement(C + "crossAx", new XAttribute("val", crossId)));
            axis.Add(new XElement(C + "crosses", new XAttribute("val", "autoZero")));
            axis.Add(new XElement(C + "auto", new XAttribute("val", "1")));
            return axis;
        }

        /// <summary>判断图表类型是否属于散点族（散点或气泡，使用 X/Y 双数值轴）。 Whether a type belongs to the scatter family (scatter or bubble, using the two value axes).</summary>
        private static bool IsScatterFamily(ExcelChartType type)
        {
            return type == ExcelChartType.ScatterMarkers || type == ExcelChartType.ScatterSmoothLines
                || type == ExcelChartType.ScatterLinesMarkers || type == ExcelChartType.Bubble;
        }

        /// <summary>判断图表类型是否属于折线族（决定逐系列 smooth 附加）。 Whether a type belongs to the line family (controls the per-series smooth element).</summary>
        private static bool IsLineFamily(ExcelChartType type)
        {
            return type == ExcelChartType.Line || type == ExcelChartType.LineWithMarkers
                || type == ExcelChartType.StackedLine || type == ExcelChartType.PercentStackedLine;
        }

        /// <summary>判断图表类型是否属于柱/条族（决定 invertIfNegative 附加）。 Whether a type belongs to the column/bar family (controls the invertIfNegative element).</summary>
        private static bool IsBarFamily(ExcelChartType type)
        {
            return type == ExcelChartType.Column || type == ExcelChartType.Bar || type == ExcelChartType.StackedColumn
                || type == ExcelChartType.StackedBar || type == ExcelChartType.PercentStackedColumn || type == ExcelChartType.PercentStackedBar;
        }

        /// <summary>生成一个数据系列元素，元素顺序按图表类型调整。 Builds one series element, ordering children according to the chart type.</summary>
        /// <param name="series">系列模型。 Series model.</param>
        /// <param name="index">系列序号。 Series index.</param>
        /// <param name="type">该系列所属图表组的类型。 The type of the chart group this series belongs to.</param>
        /// <param name="chart">所属图表。 Owning chart.</param>
        /// <param name="workbook">工作簿。 Workbook.</param>
        /// <param name="host">承载表。 Host worksheet.</param>
        /// <returns>ser 元素。 The ser element.</returns>
        private static XElement BuildSeries(ExcelChartSeries series, int index, ExcelChartType type, ExcelChart chart,
            OfficeWorkbook workbook, Worksheet host)
        {
            bool scatter = IsScatterFamily(type);
            ResolvedData categories = ResolveText(workbook, host, series.CategoryRange, series.Categories);
            ResolvedData values = ResolveNumbers(workbook, host, series.ValueRange, series.Values);
            string valueCount = values.Numbers.Count.ToString(CultureInfo.InvariantCulture);
            string xCount = (scatter ? categories.Numbers : values.Numbers).Count.ToString(CultureInfo.InvariantCulture);

            XElement ser = new XElement(C + "ser",
                new XElement(C + "idx", new XAttribute("val", index.ToString(CultureInfo.InvariantCulture))),
                new XElement(C + "order", new XAttribute("val", index.ToString(CultureInfo.InvariantCulture))));
            XElement name = BuildSeriesName(series, workbook, host, valueCount);
            if (name != null) ser.Add(name);
            XElement style = BuildSeriesStyle(series, type);
            if (style != null) ser.Add(style);
            if (type == ExcelChartType.LineWithMarkers) ser.Add(MarkerElement());
            if (IsBarFamily(type))
            {
                ser.Add(new XElement(C + "invertIfNegative", new XAttribute("val", "0")));
            }
            for (int i = 0; i < series.PointColors.Count; i++)
            {
                ser.Add(new XElement(C + "dPt",
                    new XElement(C + "idx", new XAttribute("val", i.ToString(CultureInfo.InvariantCulture))),
                    new XElement(C + "bubble3D", new XAttribute("val", "0")),
                    new XElement(C + "spPr", SolidFill(series.PointColors[i]))));
            }
            if (series.ShowDataLabels) ser.Add(BuildDataLabels(series, type));
            if (scatter)
            {
                ser.Add(new XElement(C + "xVal", categories.Reference != null
                    ? BuildNumberReference(categories) : BuildNumberLiteral(categories, xCount)));
                ser.Add(new XElement(C + "yVal", values.Reference != null
                    ? BuildNumberReference(values) : BuildNumberLiteral(values, valueCount)));
                if (type == ExcelChartType.Bubble)
                {
                    ResolvedData sizes = ResolveNumbers(workbook, host, series.SizeRange, series.Sizes);
                    string sizeCount = sizes.Numbers.Count.ToString(CultureInfo.InvariantCulture);
                    ser.Add(new XElement(C + "bubbleSize", sizes.Reference != null
                        ? BuildNumberReference(sizes) : BuildNumberLiteral(sizes, sizeCount)));
                }
                if (type == ExcelChartType.ScatterSmoothLines)
                {
                    ser.Add(new XElement(C + "smooth", new XAttribute("val", "1")));
                }
            }
            else
            {
                ser.Add(BuildCategoryElement(categories, valueCount));
                ser.Add(BuildValueElement(values, valueCount));
                if (IsLineFamily(type))
                {
                    ser.Add(new XElement(C + "smooth", new XAttribute("val", series.Smooth ? "1" : "0")));
                }
            }
            return ser;
        }

        /// <summary>生成折线图系列的数据点标记（圆点、尺寸 7）。 Builds a line series' point marker (circle, size 7).</summary>
        /// <returns>marker 元素。 The marker element.</returns>
        private static XElement MarkerElement()
        {
            return new XElement(C + "marker",
                new XElement(C + "symbol", new XAttribute("val", "circle")),
                new XElement(C + "size", new XAttribute("val", "7")));
        }

        /// <summary>生成系列名称（tx）元素。 Builds the series name (tx) element.</summary>
        /// <param name="series">系列模型。 Series model.</param>
        /// <param name="workbook">工作簿。 Workbook.</param>
        /// <param name="host">承载表。 Host worksheet.</param>
        /// <param name="pointCount">缓存点数的文本。 Cache point count as text.</param>
        /// <returns>tx 元素。 The tx element.</returns>
        private static XElement BuildSeriesName(ExcelChartSeries series, OfficeWorkbook workbook, Worksheet host, string pointCount)
        {
            ResolvedData name = ResolveText(workbook, host, series.NameRange, null);
            if (name != null && name.Reference != null)
            {
                return new XElement(C + "tx", BuildStringReference(name));
            }
            string text = name != null && name.Texts.Count > 0 ? name.Texts[0] : series.Name;
            if (string.IsNullOrEmpty(text)) return null;
            return new XElement(C + "tx", new XElement(C + "v", text));
        }

        /// <summary>生成系列样式（spPr）元素；无颜色要求时返回 null。 Builds the series style (spPr) element, or null when no color is requested.</summary>
        /// <param name="series">系列模型。 Series model.</param>
        /// <param name="type">图表类型。 Chart type.</param>
        /// <returns>spPr 元素或 null。 The spPr element or null.</returns>
        private static XElement BuildSeriesStyle(ExcelChartSeries series, ExcelChartType type)
        {
            if (!series.FillColor.HasValue && series.PointColors.Count == 0) return null;
            if (!series.FillColor.HasValue) return null;
            OfficeColor color = series.FillColor.Value;
            XElement spPr = new XElement(C + "spPr");
            if (IsLineFamily(type) || type == ExcelChartType.ScatterLinesMarkers
                || type == ExcelChartType.ScatterSmoothLines)
            {
                spPr.Add(new XElement(A + "ln",
                    new XAttribute("w", "28575"),
                    new XElement(A + "solidFill", ColorElement(color))));
            }
            else
            {
                spPr.Add(SolidFill(color));
                spPr.Add(new XElement(A + "ln", new XAttribute("w", "12700"),
                    new XElement(A + "solidFill", ColorElement(color))));
            }
            return spPr;
        }

        /// <summary>生成实心填充元素。 Builds a solid-fill element.</summary>
        /// <param name="color">填充颜色。 Fill color.</param>
        /// <returns>solidFill 元素。 The solidFill element.</returns>
        private static XElement SolidFill(OfficeColor color)
        {
            return new XElement(A + "solidFill", ColorElement(color));
        }

        /// <summary>生成 srgbClr 颜色元素。 Builds an srgbClr color element.</summary>
        /// <param name="color">颜色。 Color.</param>
        /// <returns>srgbClr 元素。 The srgbClr element.</returns>
        private static XElement ColorElement(OfficeColor color)
        {
            return new XElement(A + "srgbClr", new XAttribute("val", color.ToHex()));
        }

        /// <summary>生成数据标签元素。 Builds the data-labels element.</summary>
        /// <param name="series">系列模型。 Series model.</param>
        /// <param name="type">图表类型。 Chart type.</param>
        /// <returns>dLbls 元素。 The dLbls element.</returns>
        private static XElement BuildDataLabels(ExcelChartSeries series, ExcelChartType type)
        {
            XElement labels = new XElement(C + "dLbls");
            if (!string.IsNullOrEmpty(series.DataLabelFormat))
            {
                labels.Add(new XElement(C + "numFmt",
                    new XAttribute("formatCode", series.DataLabelFormat),
                    new XAttribute("sourceLinked", "0")));
            }
            string position;
            if (type == ExcelChartType.Pie || type == ExcelChartType.Doughnut) position = "bestFit";
            else if (IsLineFamily(type) || type == ExcelChartType.Area || type == ExcelChartType.StackedArea
                || type == ExcelChartType.PercentStackedArea || IsScatterFamily(type)) position = "t";
            else position = "outEnd";
            labels.Add(new XElement(C + "dLblPos", new XAttribute("val", position)));
            labels.Add(new XElement(C + "showLegendKey", new XAttribute("val", "0")));
            labels.Add(new XElement(C + "showVal", new XAttribute("val", "1")));
            labels.Add(new XElement(C + "showCatName", new XAttribute("val", "0")));
            labels.Add(new XElement(C + "showSerName", new XAttribute("val", "0")));
            labels.Add(new XElement(C + "showPercent", new XAttribute("val", "0")));
            labels.Add(new XElement(C + "showBubbleSize", new XAttribute("val", "0")));
            return labels;
        }

        /// <summary>生成类目（cat）子元素，按数据来源选择引用或字面量。 Builds the cat child, choosing a reference or a literal by data source.</summary>
        /// <param name="categories">已解析的类目。 Resolved categories.</param>
        /// <param name="pointCount">点数文本。 Point count text.</param>
        /// <returns>cat 元素。 The cat element.</returns>
        private static XElement BuildCategoryElement(ResolvedData categories, string pointCount)
        {
            XElement cat = new XElement(C + "cat");
            if (categories.Reference != null && categories.AllNumeric) cat.Add(BuildNumberReference(categories));
            else if (categories.Reference != null) cat.Add(BuildStringReference(categories));
            else if (categories.AllNumeric) cat.Add(BuildNumberLiteral(categories, pointCount));
            else cat.Add(BuildStringLiteral(categories, pointCount));
            return cat;
        }

        /// <summary>生成数值（val）子元素。 Builds the val child element.</summary>
        /// <param name="values">已解析的数值。 Resolved values.</param>
        /// <param name="pointCount">点数文本。 Point count text.</param>
        /// <returns>val 元素。 The val element.</returns>
        private static XElement BuildValueElement(ResolvedData values, string pointCount)
        {
            XElement val = new XElement(C + "val");
            val.Add(values.Reference != null ? BuildNumberReference(values) : BuildNumberLiteral(values, pointCount));
            return val;
        }

        /// <summary>生成字符串引用（strRef）元素。 Builds a strRef element.</summary>
        /// <param name="data">已解析数据。 Resolved data.</param>
        /// <returns>strRef 元素。 The strRef element.</returns>
        private static XElement BuildStringReference(ResolvedData data)
        {
            XElement cache = new XElement(C + "strCache",
                new XElement(C + "ptCount", new XAttribute("val", data.Texts.Count.ToString(CultureInfo.InvariantCulture))));
            for (int i = 0; i < data.Texts.Count; i++)
            {
                cache.Add(new XElement(C + "pt",
                    new XAttribute("idx", i.ToString(CultureInfo.InvariantCulture)),
                    new XElement(C + "v", data.Texts[i])));
            }
            return new XElement(C + "strRef", new XElement(C + "f", data.Reference), cache);
        }

        /// <summary>生成数值引用（numRef）元素。 Builds a numRef element.</summary>
        /// <param name="data">已解析数据。 Resolved data.</param>
        /// <returns>numRef 元素。 The numRef element.</returns>
        private static XElement BuildNumberReference(ResolvedData data)
        {
            return new XElement(C + "numRef", new XElement(C + "f", data.Reference), BuildNumberCache(data));
        }

        /// <summary>生成字符串字面量（strLit）元素。 Builds a strLit element.</summary>
        /// <param name="data">已解析数据。 Resolved data.</param>
        /// <param name="pointCount">点数文本。 Point count text.</param>
        /// <returns>strLit 元素。 The strLit element.</returns>
        private static XElement BuildStringLiteral(ResolvedData data, string pointCount)
        {
            XElement cache = new XElement(C + "strCache",
                new XElement(C + "ptCount", new XAttribute("val", pointCount)));
            for (int i = 0; i < data.Texts.Count; i++)
            {
                cache.Add(new XElement(C + "pt",
                    new XAttribute("idx", i.ToString(CultureInfo.InvariantCulture)),
                    new XElement(C + "v", data.Texts[i])));
            }
            return new XElement(C + "strLit", cache);
        }

        /// <summary>生成数值字面量（numLit）元素。 Builds a numLit element.</summary>
        /// <param name="data">已解析数据。 Resolved data.</param>
        /// <param name="pointCount">点数文本。 Point count text.</param>
        /// <returns>numLit 元素。 The numLit element.</returns>
        private static XElement BuildNumberLiteral(ResolvedData data, string pointCount)
        {
            XElement cache = new XElement(C + "numCache",
                new XElement(C + "formatCode", data.FormatCode),
                new XElement(C + "ptCount", new XAttribute("val", pointCount)));
            for (int i = 0; i < data.Numbers.Count; i++)
            {
                if (!data.Present[i]) continue;
                cache.Add(Point(i, data.Numbers[i]));
            }
            return new XElement(C + "numLit", cache);
        }

        /// <summary>生成数值缓存（numCache）元素。 Builds a numCache element.</summary>
        /// <param name="data">已解析数据。 Resolved data.</param>
        /// <returns>numCache 元素。 The numCache element.</returns>
        private static XElement BuildNumberCache(ResolvedData data)
        {
            XElement cache = new XElement(C + "numCache",
                new XElement(C + "formatCode", data.FormatCode),
                new XElement(C + "ptCount", new XAttribute("val", data.Numbers.Count.ToString(CultureInfo.InvariantCulture))));
            for (int i = 0; i < data.Numbers.Count; i++)
            {
                if (!data.Present[i]) continue;
                cache.Add(Point(i, data.Numbers[i]));
            }
            return cache;
        }

        /// <summary>生成一个带序号的数值点。 Builds one indexed numeric point.</summary>
        /// <param name="index">点序号。 Point index.</param>
        /// <param name="value">数值。 Value.</param>
        /// <returns>pt 元素。 The pt element.</returns>
        private static XElement Point(int index, double value)
        {
            return new XElement(C + "pt",
                new XAttribute("idx", index.ToString(CultureInfo.InvariantCulture)),
                new XElement(C + "v", FormatCacheNumber(value)));
        }

        /// <summary>把数值写成紧凑的缓存文本（整数值不带小数点）。 Writes a number as compact cache text (integral values carry no decimal point).</summary>
        /// <param name="value">数值。 Value.</param>
        /// <returns>缓存文本。 The cache text.</returns>
        private static string FormatCacheNumber(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return "0";
            if (value == Math.Floor(value) && Math.Abs(value) < 1e15)
            {
                return ((long)value).ToString(CultureInfo.InvariantCulture);
            }
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        /// <summary>解析区域或字面量得到文本序列。 Resolves a range or literal list into text values.</summary>
        /// <param name="workbook">工作簿。 Workbook.</param>
        /// <param name="host">承载表（区域未写表名时使用）。 Host worksheet, used when the range omits a sheet name.</param>
        /// <param name="range">区域文本。 Range text.</param>
        /// <param name="literals">字面量列表。 Literal list.</param>
        /// <returns>解析结果。 The resolved data.</returns>
        private static ResolvedData ResolveText(OfficeWorkbook workbook, Worksheet host, string range, IList<string> literals)
        {
            ResolvedData data = new ResolvedData();
            Worksheet sheet;
            int firstRow;
            int lastRow;
            int firstColumn;
            int lastColumn;
            if (TryResolveRange(workbook, host, range, out sheet, out data.Reference, out firstRow, out lastRow, out firstColumn, out lastColumn))
            {
                bool allNumeric = true;
                foreach (Cell cell in Walk(sheet, firstRow, lastRow, firstColumn, lastColumn))
                {
                    string text = cell == null ? string.Empty : cell.DisplayText();
                    data.Texts.Add(text);
                    data.Numbers.Add(ParseDoubleOrZero(text));
                    data.Present.Add(!string.IsNullOrEmpty(text));
                    double ignored;
                    if (!string.IsNullOrEmpty(text) && !double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out ignored))
                    {
                        allNumeric = false;
                    }
                    if (data.FormatCode == "General" && cell != null && cell.Style != null
                        && !string.IsNullOrEmpty(cell.Style.NumberFormat))
                    {
                        data.FormatCode = cell.Style.NumberFormat;
                    }
                }
                data.AllNumeric = allNumeric && data.Texts.Count > 0;
                return data;
            }

            data.Reference = null;
            bool literalNumeric = true;
            if (literals != null)
            {
                for (int i = 0; i < literals.Count; i++)
                {
                    string text = literals[i] ?? string.Empty;
                    data.Texts.Add(text);
                    data.Numbers.Add(ParseDoubleOrZero(text));
                    data.Present.Add(text.Length > 0);
                    double ignored;
                    if (text.Length > 0 && !double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out ignored))
                    {
                        literalNumeric = false;
                    }
                }
            }
            data.AllNumeric = literalNumeric && data.Texts.Count > 0;
            return data;
        }

        /// <summary>解析区域或字面量得到数值序列。 Resolves a range or literal list into numeric values.</summary>
        /// <param name="workbook">工作簿。 Workbook.</param>
        /// <param name="host">承载表。 Host worksheet.</param>
        /// <param name="range">区域文本。 Range text.</param>
        /// <param name="literals">字面量数值列表。 Literal numeric list.</param>
        /// <returns>解析结果。 The resolved data.</returns>
        private static ResolvedData ResolveNumbers(OfficeWorkbook workbook, Worksheet host, string range, IList<double> literals)
        {
            ResolvedData data = new ResolvedData();
            Worksheet sheet;
            int firstRow;
            int lastRow;
            int firstColumn;
            int lastColumn;
            if (TryResolveRange(workbook, host, range, out sheet, out data.Reference, out firstRow, out lastRow, out firstColumn, out lastColumn))
            {
                foreach (Cell cell in Walk(sheet, firstRow, lastRow, firstColumn, lastColumn))
                {
                    double number = 0;
                    bool present = cell != null && TryReadNumber(cell.Value, out number);
                    data.Numbers.Add(number);
                    data.Texts.Add(cell == null ? string.Empty : cell.DisplayText());
                    data.Present.Add(present);
                    if (data.FormatCode == "General" && cell != null && cell.Style != null
                        && !string.IsNullOrEmpty(cell.Style.NumberFormat))
                    {
                        data.FormatCode = cell.Style.NumberFormat;
                    }
                }
                data.AllNumeric = true;
                return data;
            }

            data.Reference = null;
            if (literals != null)
            {
                for (int i = 0; i < literals.Count; i++)
                {
                    data.Numbers.Add(literals[i]);
                    data.Texts.Add(FormatCacheNumber(literals[i]));
                    data.Present.Add(true);
                }
            }
            data.AllNumeric = true;
            return data;
        }

        /// <summary>把区域文本解析为所属表与起止行列，并生成绝对引用文本。 Parses range text into the owning sheet plus bounds and builds the absolute reference text.</summary>
        /// <param name="workbook">工作簿。 Workbook.</param>
        /// <param name="host">未写表名时使用的承载表。 Host worksheet used when the range omits a sheet name.</param>
        /// <param name="range">区域文本。 Range text.</param>
        /// <param name="sheet">输出所属表。 Outputs the owning worksheet.</param>
        /// <param name="reference">输出绝对引用文本。 Outputs the absolute reference text.</param>
        /// <param name="firstRow">输出起始行。 Outputs the first row.</param>
        /// <param name="lastRow">输出结束行。 Outputs the last row.</param>
        /// <param name="firstColumn">输出起始列。 Outputs the first column.</param>
        /// <param name="lastColumn">输出结束列。 Outputs the last column.</param>
        /// <returns>解析成功返回 true。 True when parsing succeeded.</returns>
        private static bool TryResolveRange(OfficeWorkbook workbook, Worksheet host, string range,
            out Worksheet sheet, out string reference, out int firstRow, out int lastRow, out int firstColumn, out int lastColumn)
        {
            sheet = null;
            reference = null;
            firstRow = lastRow = firstColumn = lastColumn = 0;
            if (workbook == null || string.IsNullOrEmpty(range)) return false;

            string sheetName = null;
            string body = range;
            int bang = body.LastIndexOf('!');
            if (bang >= 0)
            {
                sheetName = body.Substring(0, bang).Trim();
                body = body.Substring(bang + 1);
                if (sheetName.Length >= 2 && sheetName[0] == '\'' && sheetName[sheetName.Length - 1] == '\'')
                {
                    sheetName = sheetName.Substring(1, sheetName.Length - 2);
                }
            }

            if (sheetName != null)
            {
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    if (string.Equals(workbook.Worksheets[i].Name, sheetName, StringComparison.Ordinal))
                    {
                        sheet = workbook.Worksheets[i];
                        break;
                    }
                }
                if (sheet == null)
                {
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        if (string.Equals(workbook.Worksheets[i].Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            sheet = workbook.Worksheets[i];
                            break;
                        }
                    }
                }
                if (sheet == null) return false;
            }
            else
            {
                sheet = host;
            }
            if (sheet == null) return false;
            if (!SpreadsheetRange.TryParse(body, out firstRow, out lastRow, out firstColumn, out lastColumn)) return false;

            reference = QuoteSheetName(sheet.Name) + "!"
                + AbsoluteAddress(firstRow, firstColumn)
                + (firstRow == lastRow && firstColumn == lastColumn ? string.Empty : ":" + AbsoluteAddress(lastRow, lastColumn));
            return true;
        }

        /// <summary>
        /// 单次区域遍历允许访问的格位上限。图表声明的引用范围可以横跨整个 Excel 网格
        /// （A1:XFD1048576 即 172 亿格），真实序列远小于此，超出上限即停止，避免把读取拖成死循环。
        /// The maximum number of positions one range walk may visit. A chart reference can span the whole grid
        /// (A1:XFD1048576 is 17.2 billion positions) while real series are far smaller, so the walk stops at the
        /// cap instead of turning the read into an endless loop.
        /// </summary>
        private const int MaxWalkPositions = 1000000;

        /// <summary>按行优先遍历区域内的单元格（空位返回 null）。 Walks the cells of a range row by row, returning null for empty positions.</summary>
        /// <param name="sheet">工作表。 Worksheet.</param>
        /// <param name="firstRow">起始行。 First row.</param>
        /// <param name="lastRow">结束行。 Last row.</param>
        /// <param name="firstColumn">起始列。 First column.</param>
        /// <param name="lastColumn">结束列。 Last column.</param>
        /// <returns>单元格序列。 The sequence of cells.</returns>
        private static IEnumerable<Cell> Walk(Worksheet sheet, int firstRow, int lastRow, int firstColumn, int lastColumn)
        {
            int visited = 0;
            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int column = firstColumn; column <= lastColumn; column++)
                {
                    if (++visited > MaxWalkPositions) yield break;
                    yield return sheet.FindCell(row, column);
                }
            }
        }

        /// <summary>把表名转换为公式引用形式（必要时加单引号）。 Converts a sheet name into formula-reference form, quoting it when required.</summary>
        /// <param name="name">表名。 Sheet name.</param>
        /// <returns>可嵌入引用的表名。 The sheet name as usable in a reference.</returns>
        private static string QuoteSheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Sheet1";
            bool plain = true;
            for (int i = 0; i < name.Length; i++)
            {
                char c = name[i];
                bool allowed = (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')
                    || c == '_' || c == '.' || c > 127;
                if (!allowed)
                {
                    plain = false;
                    break;
                }
            }
            if (plain) return name;
            return "'" + name.Replace("'", "''") + "'";
        }

        /// <summary>把行列转换为绝对地址（如 <c>$D$5</c>）。 Converts a row and column into an absolute address such as <c>$D$5</c>.</summary>
        /// <param name="row">行号。 Row number.</param>
        /// <param name="column">列号。 Column number.</param>
        /// <returns>绝对地址。 The absolute address.</returns>
        private static string AbsoluteAddress(int row, int column)
        {
            string text = CellAddress.Format(row, column);
            int letters = 0;
            while (letters < text.Length && char.IsLetter(text[letters])) letters++;
            return "$" + text.Substring(0, letters) + "$" + text.Substring(letters);
        }

        /// <summary>从单元格取值转成 double（数值直取，日期取序列值，文本尽力解析）。 Reads a cell value as a double (numbers directly, dates as serial values, text parsed when possible).</summary>
        /// <param name="value">单元格值。 Cell value.</param>
        /// <param name="number">输出数值。 Outputs the number.</param>
        /// <returns>可转成数值时返回 true。 True when the value converts to a number.</returns>
        private static bool TryReadNumber(object value, out double number)
        {
            number = 0;
            if (value == null) return false;
            if (value is double)
            {
                number = (double)value;
                return true;
            }
            if (value is float || value is decimal || value is byte || value is sbyte || value is short
                || value is ushort || value is int || value is uint || value is long || value is ulong)
            {
                number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return true;
            }
            if (value is DateTime)
            {
                number = ((DateTime)value - new DateTime(1899, 12, 30)).TotalDays;
                return true;
            }
            string text = value as string;
            if (text != null && double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) return true;
            number = 0;
            return false;
        }

        /// <summary>解析文本为 double，失败时返回 0。 Parses text into a double, returning 0 on failure.</summary>
        /// <param name="text">文本。 Text.</param>
        /// <returns>解析结果或 0。 The parsed value or 0.</returns>
        private static double ParseDoubleOrZero(string text)
        {
            double value;
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ? value : 0;
        }
    }
}
