using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 读取工作表装饰部件：条件格式（数据条/色阶/图标集/高亮）、原生图表与锚定图片，
    /// 还原为可再次写出的模型，使"读取—修改—保存"不会丢失看板元素。
    /// Reads worksheet decorations: conditional formats (data bars, color scales, icon sets, highlights),
    /// native charts and anchored pictures, restoring them into a model that can be written back so a
    /// load-modify-save cycle keeps the dashboard elements.
    /// </summary>
    internal static class XlsxDecorationReader
    {
        /// <summary>一磅对应的 EMU 数值。 EMU units per point.</summary>
        private const double EmuPerPoint = 12700.0;

        /// <summary>读取单张工作表的装饰部件并写入模型。 Reads one worksheet's decorations into the model.</summary>
        /// <param name="package">包访问器。 Package accessor.</param>
        /// <param name="sheet">目标工作表模型。 Target worksheet model.</param>
        /// <param name="sheetPart">工作表部件路径。 Worksheet part path.</param>
        /// <param name="root">工作表 XML 根元素。 Root element of the worksheet XML.</param>
        /// <param name="stylesXml">styles.xml 文档（用于解析 dxf）。 The styles.xml document, used to resolve dxf entries.</param>
        public static void Read(XlsxPackage package, Worksheet sheet, string sheetPart, XElement root, XDocument stylesXml)
        {
            ReadConditionalFormats(sheet, root, stylesXml);
            ReadDrawings(package, sheet, sheetPart, root);
        }

        /// <summary>按主命名空间取第一个子元素。 Gets the first child element in the spreadsheet main namespace.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">本地名称。 Local name.</param>
        /// <returns>子元素或 null。 The child or null.</returns>
        private static XElement MainChild(XElement parent, string name)
        {
            return parent == null ? null : parent.Element(Xlsx.Main + name);
        }

        /// <summary>按本地名称取子元素（忽略命名空间），供 chart、drawing 等异构命名空间使用。
        /// Gets children matched by local name, ignoring the namespace, for heterogeneous chart and drawing namespaces.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">本地名称。 Local name.</param>
        /// <returns>匹配的子元素集合。 The matching children.</returns>
        private static List<XElement> LocalChildren(XElement parent, string name)
        {
            List<XElement> result = new List<XElement>();
            if (parent == null) return result;
            foreach (XElement child in parent.Elements())
            {
                if (string.Equals(child.Name.LocalName, name, StringComparison.Ordinal)) result.Add(child);
            }
            return result;
        }

        /// <summary>按本地名称取第一个子元素。 Gets the first child element matching a local name.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">本地名称。 Local name.</param>
        /// <returns>子元素或 null。 The child or null.</returns>
        private static XElement LocalChild(XElement parent, string name)
        {
            List<XElement> children = LocalChildren(parent, name);
            return children.Count == 0 ? null : children[0];
        }

        /// <summary>读取 conditionalFormatting 元素并按区域聚合成模型。 Reads conditionalFormatting elements, grouping rules by range.</summary>
        /// <param name="sheet">工作表模型。 Worksheet model.</param>
        /// <param name="root">工作表根元素。 Worksheet root element.</param>
        /// <param name="stylesXml">styles.xml 文档。 The styles.xml document.</param>
        private static void ReadConditionalFormats(Worksheet sheet, XElement root, XDocument stylesXml)
        {
            foreach (XElement block in Xlsx.Children(root, "conditionalFormatting"))
            {
                string sqref = Xlsx.Attribute(block, "sqref");
                if (string.IsNullOrEmpty(sqref)) continue;
                sqref = sqref.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)[0];
                ExcelConditionalFormat format = FindOrCreateFormat(sheet, sqref);
                foreach (XElement ruleElement in Xlsx.Children(block, "cfRule"))
                {
                    ExcelConditionalRule rule = ReadRule(ruleElement, stylesXml);
                    if (rule != null) format.Rules.Add(rule);
                }
            }
        }

        /// <summary>查找同区域的既有条件格式集合，缺失时新建并加入。 Finds the format set for a range, creating and appending one when missing.</summary>
        /// <param name="sheet">工作表模型。 Worksheet model.</param>
        /// <param name="sqref">区域文本。 Range text.</param>
        /// <returns>条件格式集合。 The conditional-format set.</returns>
        private static ExcelConditionalFormat FindOrCreateFormat(Worksheet sheet, string sqref)
        {
            for (int i = 0; i < sheet.ConditionalFormats.Count; i++)
            {
                if (string.Equals(sheet.ConditionalFormats[i].Range, sqref, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet.ConditionalFormats[i];
                }
            }
            ExcelConditionalFormat format = new ExcelConditionalFormat(sqref);
            sheet.ConditionalFormats.Add(format);
            return format;
        }

        /// <summary>把一条 cfRule 还原为规则模型；类型未知时返回 null。 Restores one cfRule into a rule model, returning null for unknown types.</summary>
        /// <param name="element">cfRule 元素。 The cfRule element.</param>
        /// <param name="stylesXml">styles.xml 文档。 The styles.xml document.</param>
        /// <returns>规则模型或 null。 The rule or null.</returns>
        private static ExcelConditionalRule ReadRule(XElement element, XDocument stylesXml)
        {
            string type = Xlsx.Attribute(element, "type");
            ExcelConditionalRule rule = new ExcelConditionalRule();
            rule.Colors.Clear();
            if (type == "dataBar")
            {
                rule.Kind = ExcelConditionalKind.DataBar;
                List<XElement> bounds = new List<XElement>();
                XElement dataBar = MainChild(element, "dataBar");
                foreach (XElement cfvo in Xlsx.Children(dataBar, "cfvo")) bounds.Add(cfvo);
                if (bounds.Count >= 1) ApplyBound(rule, bounds[0], true);
                if (bounds.Count >= 2) ApplyBound(rule, bounds[1], false);
                OfficeColor color;
                if (Xlsx.TryParseColor(Xlsx.Attribute(Xlsx.Child(dataBar, "color"), "rgb"), out color))
                {
                    rule.Colors.Add(color);
                }
                else rule.Colors.Add(new OfficeColor(0x63, 0xC3, 0x84));
                return rule;
            }
            if (type == "colorScale")
            {
                rule.Kind = ExcelConditionalKind.ColorScale;
                XElement scale = MainChild(element, "colorScale");
                List<XElement> cfvos = new List<XElement>();
                foreach (XElement cfvo in Xlsx.Children(scale, "cfvo")) cfvos.Add(cfvo);
                foreach (XElement colorElement in Xlsx.Children(scale, "color"))
                {
                    OfficeColor color;
                    if (Xlsx.TryParseColor(Xlsx.Attribute(colorElement, "rgb"), out color)) rule.Colors.Add(color);
                }
                if (cfvos.Count >= 1) ApplyBound(rule, cfvos[0], true);
                if (cfvos.Count >= 3) ApplyBound(rule, cfvos[cfvos.Count - 1], false);
                if (cfvos.Count >= 3)
                {
                    rule.MiddlePercent = Xlsx.DoubleAttribute(cfvos[1], "val", 50);
                }
                if (rule.Colors.Count < 2)
                {
                    rule.Colors.Add(new OfficeColor(0xFF, 0xEB, 0x84));
                    rule.Colors.Add(new OfficeColor(0x63, 0xC3, 0x84));
                }
                return rule;
            }
            if (type == "iconSet")
            {
                rule.Kind = ExcelConditionalKind.IconSet;
                XElement iconSet = MainChild(element, "iconSet");
                rule.IconSetType = Xlsx.Attribute(iconSet, "iconSet") ?? "3TrafficLights1";
                List<XElement> cfvos = new List<XElement>();
                foreach (XElement cfvo in Xlsx.Children(iconSet, "cfvo")) cfvos.Add(cfvo);
                for (int i = 1; i < cfvos.Count; i++)
                {
                    rule.IconThresholds.Add(Xlsx.DoubleAttribute(cfvos[i], "val", 0));
                }
                return rule;
            }
            if (type == "cellIs")
            {
                rule.Kind = ExcelConditionalKind.CellIs;
                rule.Operator = Xlsx.Attribute(element, "operator") ?? "greaterThan";
                List<XElement> formulas = new List<XElement>();
                foreach (XElement formula in Xlsx.Children(element, "formula")) formulas.Add(formula);
                if (formulas.Count >= 1) rule.Formula = formulas[0].Value;
                if (formulas.Count >= 2) rule.SecondFormula = formulas[1].Value;
                rule.Format = ReadDxf(stylesXml, Xlsx.IntAttribute(element, "dxfId", -1));
                return rule;
            }
            return null;
        }

        /// <summary>把一个 cfvo 端点还原为自动刻度或显式数值。 Restores a cfvo as an automatic or explicit bound.</summary>
        /// <param name="rule">规则模型。 Rule model.</param>
        /// <param name="cfvo">cfvo 元素。 The cfvo element.</param>
        /// <param name="minimum">是否为最小端。 Whether this is the minimum bound.</param>
        private static void ApplyBound(ExcelConditionalRule rule, XElement cfvo, bool minimum)
        {
            string kind = Xlsx.Attribute(cfvo, "type");
            double value = Xlsx.DoubleAttribute(cfvo, "val", 0);
            bool automatic = kind == "min" || kind == "max" || kind == "auto";
            if (minimum)
            {
                rule.AutoMinimum = automatic;
                rule.Minimum = value;
            }
            else
            {
                rule.AutoMaximum = automatic;
                rule.Maximum = value;
            }
        }

        /// <summary>从 styles.xml 的 dxfs 列表还原差量样式；越界或缺失时返回 null。
        /// Restores a differential style from the dxfs list of styles.xml, returning null when missing or out of range.</summary>
        /// <param name="stylesXml">styles.xml 文档。 The styles.xml document.</param>
        /// <param name="index">dxfId 索引。 The dxfId index.</param>
        /// <returns>差量样式或 null。 The style or null.</returns>
        private static CellStyle ReadDxf(XDocument stylesXml, int index)
        {
            if (stylesXml == null || stylesXml.Root == null || index < 0) return null;
            XElement dxfs = MainChild(stylesXml.Root, "dxfs");
            List<XElement> list = new List<XElement>();
            foreach (XElement dxf in Xlsx.Children(dxfs, "dxf")) list.Add(dxf);
            if (index >= list.Count) return null;

            CellStyle style = new CellStyle();
            XElement font = MainChild(list[index], "font");
            if (font != null)
            {
                style.Bold = Xlsx.ParseBool(Xlsx.Attribute(MainChild(font, "b"), "val"), true);
                style.Italic = Xlsx.ParseBool(Xlsx.Attribute(MainChild(font, "i"), "val"), true);
                OfficeColor fontColor;
                if (Xlsx.TryParseColor(Xlsx.Attribute(Xlsx.Child(font, "color"), "rgb"), out fontColor))
                {
                    style.FontColor = fontColor;
                }
            }
            XElement fill = MainChild(list[index], "fill");
            XElement pattern = MainChild(fill, "patternFill");
            OfficeColor background;
            if (Xlsx.TryParseColor(Xlsx.Attribute(MainChild(pattern, "bgColor"), "rgb"), out background))
            {
                style.FillColor = background;
            }
            XElement numFmt = MainChild(list[index], "numFmt");
            string formatCode = Xlsx.Attribute(numFmt, "formatCode");
            if (!string.IsNullOrEmpty(formatCode)) style.NumberFormat = formatCode;
            return style;
        }

        /// <summary>解析工作表的 drawing 部件，还原图表与图片。 Reads the sheet's drawing part, restoring charts and pictures.</summary>
        /// <param name="package">包访问器。 Package accessor.</param>
        /// <param name="sheet">工作表模型。 Worksheet model.</param>
        /// <param name="sheetPart">工作表部件路径。 Worksheet part path.</param>
        /// <param name="root">工作表根元素。 Worksheet root element.</param>
        private static void ReadDrawings(XlsxPackage package, Worksheet sheet, string sheetPart, XElement root)
        {
            XElement drawing = MainChild(root, "drawing");
            if (drawing == null) return;
            string drawingPart = package.GetRelationshipTarget(sheetPart, Xlsx.Attribute(drawing, Xlsx.R + "id"));
            XDocument drawingXml = package.GetXml(drawingPart);
            if (drawingXml == null || drawingXml.Root == null) return;

            foreach (XElement anchor in LocalChildren(drawingXml.Root, "oneCellAnchor"))
            {
                ReadAnchor(package, sheet, drawingPart, anchor);
            }
            foreach (XElement anchor in LocalChildren(drawingXml.Root, "twoCellAnchor"))
            {
                ReadAnchor(package, sheet, drawingPart, anchor);
            }
        }

        /// <summary>还原单个锚点里的图表或图片。 Restores the chart or picture inside one anchor.</summary>
        /// <param name="package">包访问器。 Package accessor.</param>
        /// <param name="sheet">工作表模型。 Worksheet model.</param>
        /// <param name="drawingPart">绘图部件路径。 Drawing part path.</param>
        /// <param name="anchor">锚点元素。 The anchor element.</param>
        private static void ReadAnchor(XlsxPackage package, Worksheet sheet, string drawingPart, XElement anchor)
        {
            XElement from = LocalChild(anchor, "from");
            int row = Xlsx.ParseInt(Value(LocalChild(from, "row")), 0) + 1;
            int column = Xlsx.ParseInt(Value(LocalChild(from, "col")), 0) + 1;
            double offsetX = Xlsx.ParseInt(Value(LocalChild(from, "colOff")), 0) / EmuPerPoint;
            double offsetY = Xlsx.ParseInt(Value(LocalChild(from, "rowOff")), 0) / EmuPerPoint;
            XElement ext = LocalChild(anchor, "ext");
            double width = Xlsx.ParseInt(Xlsx.Attribute(ext, "cx"), 0) / EmuPerPoint;
            double height = Xlsx.ParseInt(Xlsx.Attribute(ext, "cy"), 0) / EmuPerPoint;

            XElement graphic = LocalChild(anchor, "graphicFrame");
            if (graphic != null)
            {
                string id = Xlsx.Attribute(LocalChild(LocalChild(LocalChild(graphic, "graphic"), "graphicData"), "chart"),
                    Xlsx.R + "id");
                string chartPart = package.GetRelationshipTarget(drawingPart, id);
                ExcelChart chart = ReadChart(package.GetXml(chartPart));
                if (chart == null) return;
                chart.Row = row;
                chart.Column = column;
                if (width > 0) chart.WidthPoints = width;
                if (height > 0) chart.HeightPoints = height;
                sheet.Charts.Add(chart);
                return;
            }

            XElement pic = LocalChild(anchor, "pic");
            if (pic == null) return;
            string embed = Xlsx.Attribute(LocalChild(LocalChild(pic, "blipFill"), "blip"), Xlsx.R + "embed");
            string mediaPart = package.GetRelationshipTarget(drawingPart, embed);
            byte[] data = package.GetBytes(mediaPart);
            if (data == null || data.Length == 0) return;
            ExcelImage image = new ExcelImage(new OfficeImage(data));
            image.Row = row;
            image.Column = column;
            image.OffsetXPoints = offsetX;
            image.OffsetYPoints = offsetY;
            if (width > 0) image.WidthPoints = width;
            if (height > 0) image.HeightPoints = height;
            sheet.Images.Add(image);
        }

        /// <summary>把 chartSpace 部件还原为图表模型；非图表命名空间或结构缺失时返回 null。
        /// Restores a chartSpace part into the chart model, returning null when the part is missing or malformed.</summary>
        /// <param name="document">chartSpace 文档。 The chartSpace document.</param>
        /// <returns>图表模型或 null。 The chart or null.</returns>
        private static ExcelChart ReadChart(XDocument document)
        {
            if (document == null || document.Root == null) return null;
            XElement chart = LocalChild(document.Root, "chart");
            XElement plotArea = LocalChild(chart, "plotArea");
            if (plotArea == null) return null;

            ExcelChart model = new ExcelChart();
            model.Title = ChartTitle(LocalChild(chart, "title"));
            XElement legend = LocalChild(chart, "legend");
            model.ShowLegend = legend != null;
            model.LegendPosition = Xlsx.Attribute(LocalChild(legend, "legendPos"), "val") ?? "b";

            List<XElement> groups = ChartGroups(plotArea);
            if (groups.Count == 0) return null;
            model.Type = ChartType(groups[0], LocalChild(groups[0], "ser"));
            XElement doughnut = groups[groups.Count - 1];
            model.HoleSize = Xlsx.IntAttribute(LocalChild(doughnut, "holeSize"), "val", 50);
            XElement valAx = LocalChild(plotArea, "valAx");
            model.ValueAxisFormat = Xlsx.Attribute(LocalChild(valAx, "numFmt"), "formatCode");

            for (int g = 0; g < groups.Count; g++)
            {
                XElement group = groups[g];
                ExcelChartType groupType = ChartType(group, LocalChild(group, "ser"));
                bool secondary = IsSecondaryGroup(group);
                bool overrideNeeded = groupType != model.Type || secondary;
                foreach (XElement ser in LocalChildren(group, "ser"))
                {
                    ExcelChartSeries series = ReadSeries(ser, groupType);
                    if (series == null) continue;
                    if (overrideNeeded)
                    {
                        series.ChartType = groupType;
                        series.SecondaryAxis = secondary;
                    }
                    model.Series.Add(series);
                }
            }
            if (model.Series.Count == 0) return null;
            return model;
        }

        /// <summary>判断图表组是否使用次坐标轴（含次轴 Id 之一）。 Whether a chart group points at the secondary axis (carrying one of its identifiers).</summary>
        /// <param name="group">图表组元素。 The chart group.</param>
        /// <returns>次轴返回 true。 True for a secondary group.</returns>
        private static bool IsSecondaryGroup(XElement group)
        {
            foreach (XElement axId in LocalChildren(group, "axId"))
            {
                string val = Xlsx.Attribute(axId, "val");
                if (val == "333333333" || val == "444444444") return true;
            }
            return false;
        }

        /// <summary>取绘图区里出现的全部图表组元素（按文档顺序），供组合图使用。 Returns every chart-group element inside the plot area in document order, for combo charts.</summary>
        /// <param name="plotArea">绘图区元素。 The plot area.</param>
        /// <returns>图表组元素列表。 The groups.</returns>
        private static List<XElement> ChartGroups(XElement plotArea)
        {
            List<XElement> result = new List<XElement>();
            string[] names = { "barChart", "lineChart", "pieChart", "doughnutChart", "areaChart",
                "scatterChart", "bubbleChart", "radarChart" };
            foreach (XElement child in plotArea.Elements())
            {
                foreach (string name in names)
                {
                    if (string.Equals(child.Name.LocalName, name, StringComparison.Ordinal)) { result.Add(child); break; }
                }
            }
            return result;
        }

        /// <summary>取绘图区里出现的第一个图表组元素。 Returns the first chart-group element inside the plot area.</summary>
        /// <param name="plotArea">绘图区元素。 The plot area.</param>
        /// <returns>图表组元素或 null。 The group or null.</returns>
        private static XElement ChartGroup(XElement plotArea)
        {
            List<XElement> groups = ChartGroups(plotArea);
            return groups.Count == 0 ? null : groups[0];
        }

        /// <summary>由图表组元素、柱体方向与分组方式推断图表类型。 Infers the chart type from the group element, bar direction and grouping.</summary>
        /// <param name="group">图表组元素。 The chart group.</param>
        /// <param name="firstSeries">该组第一个系列元素（用于识别带标记折线），可为 null。 The group's first series (used to detect a marker line), may be null.</param>
        /// <returns>图表类型。 The chart type.</returns>
        private static ExcelChartType ChartType(XElement group, XElement firstSeries)
        {
            string name = group.Name.LocalName;
            string grouping = Xlsx.Attribute(LocalChild(group, "grouping"), "val");
            switch (name)
            {
                case "pieChart": return ExcelChartType.Pie;
                case "doughnutChart": return ExcelChartType.Doughnut;
                case "bubbleChart": return ExcelChartType.Bubble;
                case "scatterChart":
                {
                    string style = Xlsx.Attribute(LocalChild(group, "scatterStyle"), "val");
                    if (style == "lineMarker") return ExcelChartType.ScatterLinesMarkers;
                    if (style == "smoothMarker" || style == "smooth") return ExcelChartType.ScatterSmoothLines;
                    return ExcelChartType.ScatterMarkers;
                }
                case "radarChart":
                {
                    string style = Xlsx.Attribute(LocalChild(group, "radarStyle"), "val");
                    if (style == "filled") return ExcelChartType.RadarFilled;
                    if (style == "marker") return ExcelChartType.RadarWithMarkers;
                    return ExcelChartType.Radar;
                }
                case "lineChart":
                {
                    if (grouping == "stacked") return ExcelChartType.StackedLine;
                    if (grouping == "percentStacked") return ExcelChartType.PercentStackedLine;
                    bool marker = firstSeries != null && LocalChild(firstSeries, "marker") != null;
                    return marker ? ExcelChartType.LineWithMarkers : ExcelChartType.Line;
                }
                case "areaChart":
                {
                    if (grouping == "stacked") return ExcelChartType.StackedArea;
                    if (grouping == "percentStacked") return ExcelChartType.PercentStackedArea;
                    return ExcelChartType.Area;
                }
                case "barChart":
                {
                    bool horizontal = Xlsx.Attribute(LocalChild(group, "barDir"), "val") == "bar";
                    if (grouping == "percentStacked") return horizontal ? ExcelChartType.PercentStackedBar : ExcelChartType.PercentStackedColumn;
                    if (grouping == "stacked") return horizontal ? ExcelChartType.StackedBar : ExcelChartType.StackedColumn;
                    return horizontal ? ExcelChartType.Bar : ExcelChartType.Column;
                }
                default: return ExcelChartType.Column;
            }
        }

        /// <summary>读取图表标题文本。 Reads the chart title text.</summary>
        /// <param name="title">title 元素。 The title element.</param>
        /// <returns>标题文本，缺失时为 null。 The title or null.</returns>
        private static string ChartTitle(XElement title)
        {
            if (title == null) return null;
            foreach (XElement text in AllDescendants(title, "t"))
            {
                if (!string.IsNullOrEmpty(text.Value)) return text.Value;
            }
            return null;
        }

        /// <summary>还原一个数据系列：引用区域或字面量缓存、颜色、数据标签与平滑标记。
        /// Restores one series: range references or literal caches, colors, data labels and the smooth flag.</summary>
        /// <param name="ser">ser 元素。 The ser element.</param>
        /// <param name="type">所属图表类型。 Owning chart type.</param>
        /// <returns>系列模型或 null。 The series or null.</returns>
        private static ExcelChartSeries ReadSeries(XElement ser, ExcelChartType type)
        {
            ExcelChartSeries series = new ExcelChartSeries();
            XElement tx = LocalChild(ser, "tx");
            string nameReference = ReferenceOf(tx);
            if (nameReference != null) series.NameRange = nameReference;
            else
            {
                XElement direct = LocalChild(tx, "v");
                series.Name = direct == null ? null : direct.Value;
            }

            XElement cat = LocalChild(ser, "cat");
            XElement val = LocalChild(ser, "val");
            XElement xEl = LocalChild(ser, "xVal");
            XElement yEl = LocalChild(ser, "yVal");
            XElement sizeEl = LocalChild(ser, "bubbleSize");

            if (xEl != null || yEl != null)
            {
                // 散点/气泡族：X 记入 CategoryRange，Y 记入 ValueRange，气泡大小记入 SizeRange。
                // Scatter family: X goes to CategoryRange, Y to ValueRange, bubble size to SizeRange.
                series.CategoryRange = ReferenceOf(xEl);
                series.ValueRange = ReferenceOf(yEl);
                if (series.CategoryRange == null)
                {
                    foreach (double number in ReadLiteralNumbers(xEl)) series.Categories.Add(number.ToString(CultureInfo.InvariantCulture));
                }
                if (series.ValueRange == null)
                {
                    foreach (double number in ReadLiteralNumbers(yEl)) series.Values.Add(number);
                }
                series.SizeRange = ReferenceOf(sizeEl);
                if (series.SizeRange != null || sizeEl != null)
                {
                    foreach (double number in ReadLiteralNumbers(sizeEl)) series.Sizes.Add(number);
                }
            }
            else
            {
                series.CategoryRange = ReferenceOf(cat);
                series.ValueRange = ReferenceOf(val);
                List<string> literalCategories = ReadLiteralTexts(cat);
                List<double> literalValues = ReadLiteralNumbers(val);
                if (series.CategoryRange == null && literalCategories.Count > 0)
                {
                    foreach (string text in literalCategories) series.Categories.Add(text);
                }
                if (series.ValueRange == null && literalValues.Count > 0)
                {
                    foreach (double number in literalValues) series.Values.Add(number);
                }
            }
            if (series.ValueRange == null && series.Values.Count == 0) return null;

            XElement spPr = LocalChild(ser, "spPr");
            OfficeColor fill;
            if (type == ExcelChartType.Line)
            {
                XElement line = LocalChild(spPr, "ln");
                if (Xlsx.TryParseColor(Xlsx.Attribute(LocalChild(LocalChild(line, "solidFill"), "srgbClr"), "val"), out fill))
                {
                    series.FillColor = fill;
                }
            }
            else if (Xlsx.TryParseColor(Xlsx.Attribute(LocalChild(LocalChild(spPr, "solidFill"), "srgbClr"), "val"), out fill))
            {
                series.FillColor = fill;
            }

            foreach (XElement point in LocalChildren(ser, "dPt"))
            {
                OfficeColor color;
                if (Xlsx.TryParseColor(Xlsx.Attribute(
                    LocalChild(LocalChild(LocalChild(point, "spPr"), "solidFill"), "srgbClr"), "val"), out color))
                {
                    series.PointColors.Add(color);
                }
            }

            XElement labels = LocalChild(ser, "dLbls");
            series.ShowDataLabels = labels != null
                && Xlsx.ParseBool(Xlsx.Attribute(LocalChild(labels, "showVal"), "val"), false);
            series.DataLabelFormat = Xlsx.Attribute(LocalChild(labels, "numFmt"), "formatCode");
            series.Smooth = Xlsx.ParseBool(Xlsx.Attribute(LocalChild(ser, "smooth"), "val"), false);
            series.Marker = Xlsx.ParseBool(Xlsx.Attribute(LocalChild(ser, "marker"), "val"), false);
            return series;
        }

        /// <summary>取 strRef 或 numRef 里的 c:f 引用文本。 Reads the c:f reference text of a strRef or numRef element.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <returns>引用文本，缺失时为 null。 The reference or null.</returns>
        private static string ReferenceOf(XElement parent)
        {
            if (parent == null) return null;
            XElement strRef = LocalChild(parent, "strRef");
            XElement numRef = LocalChild(parent, "numRef");
            XElement formula = LocalChild(strRef != null ? strRef : numRef, "f");
            if (formula == null) return null;
            string text = formula.Value == null ? null : formula.Value.Trim();
            return string.IsNullOrEmpty(text) ? null : text;
        }

        /// <summary>读取 strLit/strCache 里的文本点序列。 Reads the text points of a strLit or strCache cache.</summary>
        /// <param name="parent">cat 或 tx 元素。 The cat or tx element.</param>
        /// <returns>文本列表，缺失时返回按点数预填的列表。 The texts, prefilled by point count when missing.</returns>
        private static List<string> ReadLiteralTexts(XElement parent)
        {
            List<string> result = new List<string>();
            if (parent == null) return result;
            XElement cache = LocalChild(parent, "strLit");
            if (cache == null) cache = LocalChild(parent, "strCache");
            if (cache == null) return result;
            XElement inner = LocalChild(cache, "strCache");
            if (inner != null) cache = inner;
            int count = Xlsx.IntAttribute(LocalChild(cache, "ptCount"), "val", 0);
            for (int i = 0; i < count; i++) result.Add(string.Empty);
            foreach (XElement point in LocalChildren(cache, "pt"))
            {
                int index = Xlsx.IntAttribute(point, "idx", -1);
                XElement value = LocalChild(point, "v");
                if (index < 0 || index >= result.Count) continue;
                result[index] = value == null ? string.Empty : value.Value;
            }
            return result;
        }

        /// <summary>读取 numLit/numCache 里的数值点序列。 Reads the numeric points of a numLit or numCache cache.</summary>
        /// <param name="parent">val 元素。 The val element.</param>
        /// <returns>数值列表，缺失时返回按点数预填的列表。 The values, prefilled by point count when missing.</returns>
        private static List<double> ReadLiteralNumbers(XElement parent)
        {
            List<double> result = new List<double>();
            if (parent == null) return result;
            XElement cache = LocalChild(parent, "numLit");
            if (cache == null) cache = LocalChild(parent, "numCache");
            if (cache == null) return result;
            XElement inner = LocalChild(cache, "numCache");
            if (inner != null) cache = inner;
            int count = Xlsx.IntAttribute(LocalChild(cache, "ptCount"), "val", 0);
            for (int i = 0; i < count; i++) result.Add(0);
            foreach (XElement point in LocalChildren(cache, "pt"))
            {
                int index = Xlsx.IntAttribute(point, "idx", -1);
                XElement value = LocalChild(point, "v");
                if (index < 0 || index >= result.Count || value == null) continue;
                result[index] = Xlsx.ParseDouble(value.Value, 0);
            }
            return result;
        }

        /// <summary>深度遍历全部同名后代元素。 Walks every descendant element with the given local name.</summary>
        /// <param name="parent">父元素。 Parent element.</param>
        /// <param name="name">本地名称。 Local name.</param>
        /// <returns>后代元素列表。 The descendants.</returns>
        private static List<XElement> AllDescendants(XElement parent, string name)
        {
            List<XElement> result = new List<XElement>();
            if (parent == null) return result;
            foreach (XElement descendant in parent.Descendants())
            {
                if (string.Equals(descendant.Name.LocalName, name, StringComparison.Ordinal)) result.Add(descendant);
            }
            return result;
        }

        /// <summary>取元素的文本内容，元素为 null 时返回空串。 Reads an element's text, returning an empty string when it is null.</summary>
        /// <param name="element">目标元素。 The element.</param>
        /// <returns>文本内容。 The text.</returns>
        private static string Value(XElement element)
        {
            return element == null || element.Value == null ? string.Empty : element.Value.Trim();
        }
    }
}
