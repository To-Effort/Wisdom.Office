using System;
using System.Collections.Generic;
using System.Globalization;

namespace Wisdom.Office.Core
{
    /// <summary>电子表格原生图表类型。The native spreadsheet chart type.</summary>
    public enum ExcelChartType
    {
        /// <summary>簇状柱形图（纵向柱）。Clustered column chart (vertical bars).</summary>
        Column = 0,
        /// <summary>簇状条形图（横向柱）。Clustered bar chart (horizontal bars).</summary>
        Bar = 1,
        /// <summary>折线图。Line chart.</summary>
        Line = 2,
        /// <summary>饼图。Pie chart.</summary>
        Pie = 3,
        /// <summary>环形图。Doughnut chart.</summary>
        Doughnut = 4,
        /// <summary>面积图。Area chart.</summary>
        Area = 5,
        /// <summary>堆积柱形图。Stacked column chart.</summary>
        StackedColumn = 6,
        /// <summary>百分比堆积柱形图。One-hundred-percent stacked column chart.</summary>
        PercentStackedColumn = 7,
        /// <summary>堆积条形图。Stacked bar chart.</summary>
        StackedBar = 8,
        /// <summary>百分比堆积条形图。One-hundred-percent stacked bar chart.</summary>
        PercentStackedBar = 9,
        /// <summary>带数据点标记的折线图。Line chart with point markers.</summary>
        LineWithMarkers = 10,
        /// <summary>堆积折线图。Stacked line chart.</summary>
        StackedLine = 11,
        /// <summary>百分比堆积折线图。One-hundred-percent stacked line chart.</summary>
        PercentStackedLine = 12,
        /// <summary>堆积面积图。Stacked area chart.</summary>
        StackedArea = 13,
        /// <summary>百分比堆积面积图。One-hundred-percent stacked area chart.</summary>
        PercentStackedArea = 14,
        /// <summary>仅带数据标记的散点图。Scatter chart showing only markers.</summary>
        ScatterMarkers = 15,
        /// <summary>带平滑线的散点图。Scatter chart with smooth lines.</summary>
        ScatterSmoothLines = 16,
        /// <summary>带直线和数据标记的散点图。Scatter chart with straight lines and markers.</summary>
        ScatterLinesMarkers = 17,
        /// <summary>气泡图（第三维决定气泡大小）。Bubble chart, a third dimension sets bubble size.</summary>
        Bubble = 18,
        /// <summary>标准雷达图。Radar chart with lines.</summary>
        Radar = 19,
        /// <summary>带数据标记的雷达图。Radar chart with markers.</summary>
        RadarWithMarkers = 20,
        /// <summary>填充雷达图。Filled radar chart.</summary>
        RadarFilled = 21
    }

    /// <summary>图表数据系列：类目标签与数值，既可用字面量，也可指向工作表单元格区域。A chart data series: category labels and values, either literal or pointing at a worksheet range.</summary>
    public sealed class ExcelChartSeries
    {
        /// <summary>系列名称文本（未设置 NameRange 时使用）。Series name text (used when NameRange is not set).</summary>
        public string Name { get; set; }
        /// <summary>系列名称所在的单元格区域，如 <c>看板!$D$4</c>。Cell range supplying the series name, such as <c>Dashboard!$D$4</c>.</summary>
        public string NameRange { get; set; }
        /// <summary>类目标签列表（未设置 CategoryRange 时使用）。Category label list (used when CategoryRange is not set).</summary>
        public IList<string> Categories { get; private set; }
        /// <summary>数值列表（未设置 ValueRange 时使用）。Value list (used when ValueRange is not set).</summary>
        public IList<double> Values { get; private set; }
        /// <summary>类目所在的单元格区域，如 <c>看板!$E$4:$P$4</c>。Cell range supplying categories, such as <c>Dashboard!$E$4:$P$4</c>.</summary>
        public string CategoryRange { get; set; }
        /// <summary>数值所在的单元格区域，如 <c>看板!$E$5:$P$5</c>。Cell range supplying values, such as <c>Dashboard!$E$5:$P$5</c>.</summary>
        public string ValueRange { get; set; }
        /// <summary>气泡大小数值列表（未设置 SizeRange 时使用，仅气泡图）。Bubble-size list (used when SizeRange is not set; bubble charts only).</summary>
        public IList<double> Sizes { get; private set; }
        /// <summary>气泡大小所在的单元格区域，仅气泡图使用。Cell range supplying bubble sizes (bubble charts only).</summary>
        public string SizeRange { get; set; }
        /// <summary>组合图中该系列改用的小节类型；为 null 时沿用图表 <see cref="ExcelChart.Type"/>。Per-series chart group used in a combo chart; inherits <see cref="ExcelChart.Type"/> when null.</summary>
        public ExcelChartType? ChartType { get; set; }
        /// <summary>是否把该系列放到组合图的次坐标轴（右侧数值轴 + 隐藏类目轴）。Whether to place this series on a combo chart's secondary axis (right value axis plus a hidden category axis).</summary>
        public bool SecondaryAxis { get; set; }
        /// <summary>系列填充色；饼图/环形图用作逐点色板的起点。Series fill color; for pie and doughnut charts it seeds the per-point color ramp.</summary>
        public OfficeColor? FillColor { get; set; }
        /// <summary>逐数据点颜色（饼图扇区、环形图分块），按序号取用，不足时回退 FillColor。Per-point colors (pie slices, doughnut segments) used by index, falling back to FillColor when exhausted.</summary>
        public List<OfficeColor> PointColors { get; private set; }
        /// <summary>是否在数据点上方显示数值标签。Whether data labels showing the value are displayed.</summary>
        public bool ShowDataLabels { get; set; }
        /// <summary>数据标签数字格式代码，如 <c>#,##0</c>。Number-format code for data labels, such as <c>#,##0</c>.</summary>
        public string DataLabelFormat { get; set; }
        /// <summary>折线图是否显示数据点标记。Whether the line chart shows markers at data points.</summary>
        public bool Marker { get; set; }
        /// <summary>折线图是否平滑曲线。Whether the line chart uses smooth curves.</summary>
        public bool Smooth { get; set; }

        /// <summary>创建一个空数据系列。Creates an empty data series.</summary>
        public ExcelChartSeries()
        {
            Categories = new List<string>();
            Values = new List<double>();
            Sizes = new List<double>();
            PointColors = new List<OfficeColor>();
        }

        /// <summary>创建带名称与字面量数据的系列。Creates a series with a name and literal data.</summary>
        /// <param name="name">系列名称。Series name.</param>
        /// <param name="categories">类目标签。Category labels.</param>
        /// <param name="values">数值。Values.</param>
        public ExcelChartSeries(string name, IList<string> categories, IList<double> values)
        {
            Name = name;
            Categories = categories ?? new List<string>();
            Values = values ?? new List<double>();
            Sizes = new List<double>();
            PointColors = new List<OfficeColor>();
        }
    }

    /// <summary>嵌入工作表的图表：类型、系列与锚定位置（以单元格 + 磅值定位）。A chart embedded in a worksheet: type, series and anchor position (given as a cell plus sizes in points).</summary>
    public sealed class ExcelChart
    {
        /// <summary>图表类型。Chart type.</summary>
        public ExcelChartType Type { get; set; }
        /// <summary>图表标题；为空时不写标题元素。Chart title; no title element is written when empty.</summary>
        public string Title { get; set; }
        /// <summary>数据系列列表。List of data series.</summary>
        public List<ExcelChartSeries> Series { get; private set; }
        /// <summary>锚定行号（从 1 开始）。Anchor row number (starting from 1).</summary>
        public int Row { get; set; }
        /// <summary>锚定列号（从 1 开始）。Anchor column number (starting from 1).</summary>
        public int Column { get; set; }
        /// <summary>图表宽度（磅）。Chart width in points.</summary>
        public double WidthPoints { get; set; }
        /// <summary>图表高度（磅）。Chart height in points.</summary>
        public double HeightPoints { get; set; }
        /// <summary>是否显示图例。Whether the legend is displayed.</summary>
        public bool ShowLegend { get; set; }
        /// <summary>图例位置：<c>b</c>、<c>t</c>、<c>l</c>、<c>r</c>，默认底部。Legend position: <c>b</c>, <c>t</c>, <c>l</c> or <c>r</c>; bottom by default.</summary>
        public string LegendPosition { get; set; }
        /// <summary>环形图中心空洞比例（0~90），仅环形图使用。Doughnut hole size (0-90), used only by doughnut charts.</summary>
        public int HoleSize { get; set; }
        /// <summary>值坐标轴数字格式代码，如 <c>#,##0</c>。Number-format code for the value axis, such as <c>#,##0</c>.</summary>
        public string ValueAxisFormat { get; set; }

        /// <summary>创建默认柱形图：锚定 A1、400x260 磅、显示底部图例。Creates a column chart anchored at A1, 400x260 points, with a bottom legend.</summary>
        public ExcelChart()
        {
            Type = ExcelChartType.Column;
            Series = new List<ExcelChartSeries>();
            Row = 1;
            Column = 1;
            WidthPoints = 400;
            HeightPoints = 260;
            ShowLegend = true;
            LegendPosition = "b";
            HoleSize = 50;
        }

        /// <summary>按单元格区域追加一个数据系列（读写两端都只记录区域，取值在写出时解析）。Adds a data series described by cell ranges (only the ranges are stored; values are resolved when writing).</summary>
        /// <param name="nameRange">系列名区域，如 <c>看板!$D$4</c>，可为 null。Series-name range such as <c>Dashboard!$D$4</c>; may be null.</param>
        /// <param name="categoryRange">类目区域，可为 null。Category range; may be null.</param>
        /// <param name="valueRange">数值区域。Value range.</param>
        /// <returns>新建并加入的系列。The created and added series.</returns>
        public ExcelChartSeries AddSeries(string nameRange, string categoryRange, string valueRange)
        {
            ExcelChartSeries series = new ExcelChartSeries();
            series.NameRange = nameRange;
            series.CategoryRange = categoryRange;
            series.ValueRange = valueRange;
            Series.Add(series);
            return series;
        }

        /// <summary>按字面量追加一个数据系列。Adds a data series from literal values.</summary>
        /// <param name="name">系列名称。Series name.</param>
        /// <param name="categories">类目标签。Category labels.</param>
        /// <param name="values">数值。Values.</param>
        /// <returns>新建并加入的系列。The created and added series.</returns>
        public ExcelChartSeries AddSeries(string name, IList<string> categories, IList<double> values)
        {
            ExcelChartSeries series = new ExcelChartSeries(name, categories, values);
            Series.Add(series);
            return series;
        }

        /// <summary>按单元格区域追加一个散点/气泡系列：X 取自 <paramref name="xRange"/>（记入 CategoryRange），Y 取自 <paramref name="yRange"/>。Adds a scatter/bubble series by cell ranges; X comes from <paramref name="xRange"/> (stored as CategoryRange) and Y from <paramref name="yRange"/>.</summary>
        /// <param name="nameRange">系列名区域。Series-name range.</param>
        /// <param name="xRange">X 值区域。X-value range.</param>
        /// <param name="yRange">Y 值区域。Y-value range.</param>
        /// <param name="sizeRange">气泡大小区域，散点图传 null。Bubble-size range; null for scatter charts.</param>
        /// <returns>新建并加入的系列。The created and added series.</returns>
        public ExcelChartSeries AddScatterSeries(string nameRange, string xRange, string yRange, string sizeRange)
        {
            ExcelChartSeries series = new ExcelChartSeries();
            series.NameRange = nameRange;
            series.CategoryRange = xRange;
            series.ValueRange = yRange;
            series.SizeRange = sizeRange;
            Series.Add(series);
            return series;
        }

        /// <summary>追加一个组合图系列，指定其在组合中的小节类型与是否放到次坐标轴。Adds a combo-chart series with the group type it uses and whether it sits on the secondary axis.</summary>
        /// <param name="nameRange">系列名区域。Series-name range.</param>
        /// <param name="categoryRange">类目区域。Category range.</param>
        /// <param name="valueRange">数值区域。Value range.</param>
        /// <param name="groupType">该系列在组合中的图表类型。The chart type this series forms within the combo.</param>
        /// <param name="secondaryAxis">是否放到次坐标轴。Whether the series uses the secondary axis.</param>
        /// <returns>新建并加入的系列。The created and added series.</returns>
        public ExcelChartSeries AddComboSeries(string nameRange, string categoryRange, string valueRange, ExcelChartType groupType, bool secondaryAxis)
        {
            ExcelChartSeries series = AddSeries(nameRange, categoryRange, valueRange);
            series.ChartType = groupType;
            series.SecondaryAxis = secondaryAxis;
            return series;
        }
    }

    /// <summary>锚定到单元格的图片：按左上单元格定位并以磅值指定显示尺寸。An image anchored to a cell: positioned by its top-left cell with the display size given in points.</summary>
    public sealed class ExcelImage
    {
        /// <summary>图片内容。Image content.</summary>
        public OfficeImage Image { get; private set; }
        /// <summary>锚定行号（从 1 开始）。Anchor row number (starting from 1).</summary>
        public int Row { get; set; }
        /// <summary>锚定列号（从 1 开始）。Anchor column number (starting from 1).</summary>
        public int Column { get; set; }
        /// <summary>相对锚定单元格的水平偏移（磅）。Horizontal offset from the anchor cell, in points.</summary>
        public double OffsetXPoints { get; set; }
        /// <summary>相对锚定单元格的垂直偏移（磅）。Vertical offset from the anchor cell, in points.</summary>
        public double OffsetYPoints { get; set; }
        /// <summary>显示宽度（磅）。Display width in points.</summary>
        public double WidthPoints { get; set; }
        /// <summary>显示高度（磅）。Display height in points.</summary>
        public double HeightPoints { get; set; }

        /// <summary>以图片内容创建锚定图片，尺寸取图片的默认显示尺寸。Creates an anchored image whose size defaults to the image's intrinsic display size.</summary>
        /// <param name="image">图片内容。Image content.</param>
        /// <exception cref="ArgumentNullException"><paramref name="image"/> 为 null。Is null.</exception>
        public ExcelImage(OfficeImage image)
        {
            if (image == null) throw new ArgumentNullException("image");
            Image = image;
            Row = 1;
            Column = 1;
            WidthPoints = image.DefaultWidthPoints;
            HeightPoints = image.DefaultHeightPoints;
        }
    }

    /// <summary>条件格式规则类型。The conditional-format rule kind.</summary>
    public enum ExcelConditionalKind
    {
        /// <summary>数据条。Data bar.</summary>
        DataBar = 0,
        /// <summary>色阶。Color scale.</summary>
        ColorScale = 1,
        /// <summary>图标集。Icon set.</summary>
        IconSet = 2,
        /// <summary>单元格数值比较高亮。Cell value comparison highlight.</summary>
        CellIs = 3
    }

    /// <summary>一条条件格式规则；按 <see cref="Kind"/> 使用对应字段。A single conditional-format rule; the fields used depend on <see cref="Kind"/>.</summary>
    public sealed class ExcelConditionalRule
    {
        /// <summary>规则类型。Rule kind.</summary>
        public ExcelConditionalKind Kind { get; set; }
        /// <summary>数据条/色阶颜色；色阶按 2 或 3 端点顺序给出。Colors for data bars and color scales, given in endpoint order (two or three) for scales.</summary>
        public List<OfficeColor> Colors { get; private set; }
        /// <summary>是否自动取最小值作为刻度起点。Whether the minimum bound is automatic.</summary>
        public bool AutoMinimum { get; set; }
        /// <summary>是否自动取最大值作为刻度终点。Whether the maximum bound is automatic.</summary>
        public bool AutoMaximum { get; set; }
        /// <summary>手动最小值（<see cref="AutoMinimum"/> 为 false 时生效）。Explicit minimum bound, used when <see cref="AutoMinimum"/> is false.</summary>
        public double Minimum { get; set; }
        /// <summary>手动最大值（<see cref="AutoMaximum"/> 为 false 时生效）。Explicit maximum bound, used when <see cref="AutoMaximum"/> is false.</summary>
        public double Maximum { get; set; }
        /// <summary>色阶中间端点取值（0~100 百分比），仅在三个颜色时生效。Color-scale midpoint as a percentage (0-100), used only when three colors are given.</summary>
        public double MiddlePercent { get; set; }
        /// <summary>图标集类型，如 <c>3TrafficLights1</c>、<c>3Arrows</c>。Icon set type, such as <c>3TrafficLights1</c> or <c>3Arrows</c>.</summary>
        public string IconSetType { get; set; }
        /// <summary>图标集阈值（百分比，升序，个数等于图标数减一）。Icon set thresholds in ascending percent order; the count is one fewer than the icon count.</summary>
        public List<double> IconThresholds { get; private set; }
        /// <summary>数值比较运算符：<c>greaterThan</c>、<c>lessThan</c>、<c>between</c>、<c>equal</c>、<c>notEqual</c>、<c>greaterThanOrEqual</c>、<c>lessThanOrEqual</c>。Comparison operator for CellIs rules.</summary>
        public string Operator { get; set; }
        /// <summary>比较公式（数字或公式文本），between 时 <see cref="SecondFormula"/> 必填。Comparison formula (a number or formula text); <see cref="SecondFormula"/> is required for <c>between</c>.</summary>
        public string Formula { get; set; }
        /// <summary>between 的第二个比较公式。Second comparison formula, used by <c>between</c>.</summary>
        public string SecondFormula { get; set; }
        /// <summary>CellIs 命中时的差量样式（字体色、粗体、背景色）。Differential style applied when a CellIs rule matches (font color, bold, background).</summary>
        public CellStyle Format { get; set; }

        /// <summary>创建自动取极值的数据条规则（绿色填充）。Creates a data-bar rule with automatic bounds and a green fill.</summary>
        public ExcelConditionalRule()
        {
            Kind = ExcelConditionalKind.DataBar;
            Colors = new List<OfficeColor>();
            IconThresholds = new List<double>();
            AutoMinimum = true;
            AutoMaximum = true;
            MiddlePercent = 50;
            IconSetType = "3TrafficLights1";
            Operator = "greaterThan";
            Colors.Add(new OfficeColor(0x63, 0xC3, 0x84));
        }
    }

    /// <summary>作用于一块区域的条件格式集合。The conditional-format set applied to one range.</summary>
    public sealed class ExcelConditionalFormat
    {
        /// <summary>区域文本，如 <c>D5:P5</c> 或 <c>看板!$D$5:$P$5</c>。Range text, such as <c>D5:P5</c> or <c>Dashboard!$D$5:$P$5</c>.</summary>
        public string Range { get; set; }
        /// <summary>该区域内的规则，按写出顺序赋予优先级。Rules inside the range; priorities are assigned in write order.</summary>
        public List<ExcelConditionalRule> Rules { get; private set; }

        /// <summary>创建作用于指定区域的条件格式集合。Creates a conditional-format set for the given range.</summary>
        /// <param name="range">区域文本。Range text.</param>
        public ExcelConditionalFormat(string range)
        {
            Range = range ?? string.Empty;
            Rules = new List<ExcelConditionalRule>();
        }

        /// <summary>追加一条数据条规则并返回它。Appends and returns a data-bar rule.</summary>
        /// <param name="fill">数据条颜色。Data-bar fill color.</param>
        /// <returns>新规则。The new rule.</returns>
        public ExcelConditionalRule AddDataBar(OfficeColor fill)
        {
            ExcelConditionalRule rule = new ExcelConditionalRule();
            rule.Kind = ExcelConditionalKind.DataBar;
            rule.Colors.Clear();
            rule.Colors.Add(fill);
            Rules.Add(rule);
            return rule;
        }

        /// <summary>追加一条双色色阶规则并返回它。Appends and returns a two-color scale rule.</summary>
        /// <param name="minimum">最小值端点颜色。Color at the minimum endpoint.</param>
        /// <param name="maximum">最大值端点颜色。Color at the maximum endpoint.</param>
        /// <returns>新规则。The new rule.</returns>
        public ExcelConditionalRule AddColorScale(OfficeColor minimum, OfficeColor maximum)
        {
            ExcelConditionalRule rule = new ExcelConditionalRule();
            rule.Kind = ExcelConditionalKind.ColorScale;
            rule.Colors.Clear();
            rule.Colors.Add(minimum);
            rule.Colors.Add(maximum);
            Rules.Add(rule);
            return rule;
        }

        /// <summary>追加一条三色色阶规则并返回它。Appends and returns a three-color scale rule.</summary>
        /// <param name="minimum">最小值端点颜色。Color at the minimum endpoint.</param>
        /// <param name="middle">中间端点颜色。Color at the midpoint.</param>
        /// <param name="maximum">最大值端点颜色。Color at the maximum endpoint.</param>
        /// <returns>新规则。The new rule.</returns>
        public ExcelConditionalRule AddColorScale(OfficeColor minimum, OfficeColor middle, OfficeColor maximum)
        {
            ExcelConditionalRule rule = AddColorScale(minimum, maximum);
            rule.Colors.Insert(1, middle);
            return rule;
        }

        /// <summary>追加一条图标集规则并返回它。Appends and returns an icon-set rule.</summary>
        /// <param name="iconSetType">图标集类型，如 <c>3TrafficLights1</c>。Icon set type, such as <c>3TrafficLights1</c>.</param>
        /// <param name="thresholds">百分比阈值（升序）。Ascending percentage thresholds.</param>
        /// <returns>新规则。The new rule.</returns>
        public ExcelConditionalRule AddIconSet(string iconSetType, IList<double> thresholds)
        {
            ExcelConditionalRule rule = new ExcelConditionalRule();
            rule.Kind = ExcelConditionalKind.IconSet;
            rule.IconSetType = string.IsNullOrEmpty(iconSetType) ? "3TrafficLights1" : iconSetType;
            rule.Colors.Clear();
            if (thresholds != null)
            {
                for (int i = 0; i < thresholds.Count; i++) rule.IconThresholds.Add(thresholds[i]);
            }
            Rules.Add(rule);
            return rule;
        }

        /// <summary>追加一条数值比较高亮规则并返回它。Appends and returns a cell-value comparison rule.</summary>
        /// <param name="operatorName">运算符，如 <c>greaterThan</c>。Operator, such as <c>greaterThan</c>.</param>
        /// <param name="formula">比较公式。Comparison formula.</param>
        /// <param name="format">命中时的差量样式。Differential style applied on match.</param>
        /// <returns>新规则。The new rule.</returns>
        public ExcelConditionalRule AddCellIs(string operatorName, string formula, CellStyle format)
        {
            ExcelConditionalRule rule = new ExcelConditionalRule();
            rule.Kind = ExcelConditionalKind.CellIs;
            rule.Operator = string.IsNullOrEmpty(operatorName) ? "greaterThan" : operatorName;
            rule.Formula = formula;
            rule.Format = format;
            rule.Colors.Clear();
            Rules.Add(rule);
            return rule;
        }
    }

    /// <summary>把 1 基行列与区域文本互转，供条件格式与图表引用共用。Converts between 1-based row/column numbers and range text, shared by conditional formatting and chart references.</summary>
    public static class SpreadsheetRange
    {
        /// <summary>解析区域文本为起止行列（支持可选表名前缀与 <c>$</c> 绝对标记）。Parses range text into start and end rows/columns (an optional sheet-name prefix and <c>$</c> absolute marks are supported).</summary>
        /// <param name="range">区域文本，如 <c>$D$5:$P$5</c>、<c>看板!D5:P5</c>。Range text, such as <c>$D$5:$P$5</c> or <c>Dashboard!D5:P5</c>.</param>
        /// <param name="firstRow">输出起始行。Outputs the first row.</param>
        /// <param name="lastRow">输出结束行。Outputs the last row.</param>
        /// <param name="firstColumn">输出起始列。Outputs the first column.</param>
        /// <param name="lastColumn">输出结束列。Outputs the last column.</param>
        /// <returns>解析成功返回 true。True when parsing succeeded.</returns>
        public static bool TryParse(string range, out int firstRow, out int lastRow, out int firstColumn, out int lastColumn)
        {
            firstRow = lastRow = firstColumn = lastColumn = 0;
            if (string.IsNullOrEmpty(range)) return false;
            string text = range;
            int bang = text.LastIndexOf('!');
            if (bang >= 0) text = text.Substring(bang + 1);
            string[] parts = text.Split(':');
            int row;
            int column;
            if (parts.Length == 1)
            {
                if (!TryParseCell(parts[0], out row, out column)) return false;
                firstRow = lastRow = row;
                firstColumn = lastColumn = column;
                return true;
            }
            if (parts.Length != 2) return false;
            if (!TryParseCell(parts[0], out firstRow, out firstColumn)) return false;
            if (!TryParseCell(parts[1], out lastRow, out lastColumn)) return false;
            if (firstRow > lastRow)
            {
                int swap = firstRow;
                firstRow = lastRow;
                lastRow = swap;
            }
            if (firstColumn > lastColumn)
            {
                int swapColumn = firstColumn;
                firstColumn = lastColumn;
                lastColumn = swapColumn;
            }
            return true;
        }

        /// <summary>解析单个单元格地址（去掉 <c>$</c>，列名按 26 进制）。Parses one cell address (strips <c>$</c> and reads the base-26 column letters).</summary>
        /// <param name="cell">单元格文本。Cell text.</param>
        /// <param name="row">输出行号。Outputs the row number.</param>
        /// <param name="column">输出列号。Outputs the column number.</param>
        /// <returns>解析成功返回 true。True when parsing succeeded.</returns>
        public static bool TryParseCell(string cell, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrEmpty(cell)) return false;
            string text = cell.Replace("$", string.Empty);
            int letters = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = char.ToUpperInvariant(text[i]);
                if (c >= 'A' && c <= 'Z') letters++;
                else break;
            }
            if (letters == 0) return false;
            for (int i = 0; i < letters; i++)
            {
                column = column * 26 + (char.ToUpperInvariant(text[i]) - 'A' + 1);
            }
            string digits = text.Substring(letters);
            if (!int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out row) || row < 1)
            {
                row = 0;
                return false;
            }
            return true;
        }

        /// <summary>格式化 1 基行列范围为 <c>A1:B2</c> 文本。Formats a 1-based row/column span as <c>A1:B2</c> text.</summary>
        /// <param name="firstRow">起始行。First row.</param>
        /// <param name="firstColumn">起始列。First column.</param>
        /// <param name="lastRow">结束行。Last row.</param>
        /// <param name="lastColumn">结束列。Last column.</param>
        /// <returns>区域文本。The range text.</returns>
        public static string Format(int firstRow, int firstColumn, int lastRow, int lastColumn)
        {
            return CellAddress.Format(firstRow, firstColumn) + ":" + CellAddress.Format(lastRow, lastColumn);
        }
    }
}
