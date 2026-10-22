using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 生成 worksheet 的 conditionalFormatting 元素：数据条、色阶、图标集与数值比较高亮。
    /// Builds the worksheet's conditionalFormatting elements: data bars, color scales, icon sets and
    /// cell-value comparisons.
    /// </summary>
    internal sealed class XlsxConditionalFormatBuilder
    {
        /// <summary>全局优先级计数器（Excel 要求规则优先级唯一）。 Global priority counter; Excel requires unique rule priorities.</summary>
        private int _priority;

        /// <summary>为一张工作表生成全部 conditionalFormatting 元素。 Builds every conditionalFormatting element of one worksheet.</summary>
        /// <param name="sheet">工作表。 Worksheet.</param>
        /// <param name="registerDxf">登记差量样式并返回 dxfId 的回调。 Callback registering a differential style and returning its dxfId.</param>
        /// <returns>按区域顺序排列的元素列表。 Elements ordered by range.</returns>
        public List<XElement> Build(Worksheet sheet, Func<CellStyle, int> registerDxf)
        {
            List<XElement> result = new List<XElement>();
            if (sheet == null || sheet.ConditionalFormats == null) return result;
            for (int i = 0; i < sheet.ConditionalFormats.Count; i++)
            {
                ExcelConditionalFormat format = sheet.ConditionalFormats[i];
                if (format == null || format.Range == null || format.Range.Length == 0) continue;
                if (format.Rules.Count == 0) continue;
                XElement container = new XElement(Xlsx.Main + "conditionalFormatting",
                    new XAttribute("sqref", StripSheetPrefix(format.Range)));
                for (int r = 0; r < format.Rules.Count; r++)
                {
                    XElement rule = BuildRule(format.Rules[r], registerDxf);
                    if (rule != null) container.Add(rule);
                }
                if (container.HasElements) result.Add(container);
            }
            return result;
        }

        /// <summary>生成单条 cfRule 元素；规则不可写出时返回 null。 Builds one cfRule element, or null when the rule cannot be written.</summary>
        /// <param name="rule">规则模型。 Rule model.</param>
        /// <param name="registerDxf">差量样式登记回调。 Differential-style registration callback.</param>
        /// <returns>cfRule 元素或 null。 The cfRule element or null.</returns>
        private XElement BuildRule(ExcelConditionalRule rule, Func<CellStyle, int> registerDxf)
        {
            if (rule == null) return null;
            _priority++;
            string priority = _priority.ToString(CultureInfo.InvariantCulture);
            switch (rule.Kind)
            {
                case ExcelConditionalKind.DataBar:
                {
                    XElement dataBar = new XElement(Xlsx.Main + "dataBar",
                        BoundElements(rule),
                        new XElement(Xlsx.Main + "color",
                            new XAttribute("rgb", ColorOf(rule, 0).ToArgbHex())));
                    return new XElement(Xlsx.Main + "cfRule",
                        new XAttribute("type", "dataBar"),
                        new XAttribute("priority", priority),
                        dataBar);
                }
                case ExcelConditionalKind.ColorScale:
                {
                    XElement scale = new XElement(Xlsx.Main + "colorScale");
                    int stops = rule.Colors.Count < 2 ? 2 : (rule.Colors.Count > 3 ? 3 : rule.Colors.Count);
                    if (stops == 2)
                    {
                        scale.Add(BoundElements(rule));
                    }
                    else
                    {
                        scale.Add(new XElement(Xlsx.Main + "cfvo", new XAttribute("type", "min")));
                        scale.Add(new XElement(Xlsx.Main + "cfvo",
                            new XAttribute("type", "percentile"),
                            new XAttribute("val", Percent(rule.MiddlePercent))));
                        scale.Add(new XElement(Xlsx.Main + "cfvo", new XAttribute("type", "max")));
                    }
                    for (int i = 0; i < stops; i++)
                    {
                        scale.Add(new XElement(Xlsx.Main + "color", new XAttribute("rgb", ColorOf(rule, i).ToArgbHex())));
                    }
                    return new XElement(Xlsx.Main + "cfRule",
                        new XAttribute("type", "colorScale"),
                        new XAttribute("priority", priority),
                        scale);
                }
                case ExcelConditionalKind.IconSet:
                {
                    int icons = IconCount(rule.IconSetType);
                    XElement iconSet = new XElement(Xlsx.Main + "iconSet",
                        new XAttribute("iconSet", string.IsNullOrEmpty(rule.IconSetType) ? "3TrafficLights1" : rule.IconSetType));
                    iconSet.Add(new XElement(Xlsx.Main + "cfvo",
                        new XAttribute("type", "percent"),
                        new XAttribute("val", "0")));
                    for (int i = 1; i < icons; i++)
                    {
                        double threshold = rule.IconThresholds != null && rule.IconThresholds.Count >= i
                            ? rule.IconThresholds[i - 1]
                            : i * 100.0 / icons;
                        iconSet.Add(new XElement(Xlsx.Main + "cfvo",
                            new XAttribute("type", "percent"),
                            new XAttribute("val", Percent(threshold))));
                    }
                    return new XElement(Xlsx.Main + "cfRule",
                        new XAttribute("type", "iconSet"),
                        new XAttribute("priority", priority),
                        iconSet);
                }
                default:
                {
                    if (string.IsNullOrEmpty(rule.Formula)) return null;
                    XElement cellIs = new XElement(Xlsx.Main + "cfRule",
                        new XAttribute("type", "cellIs"),
                        new XAttribute("operator", string.IsNullOrEmpty(rule.Operator) ? "greaterThan" : rule.Operator),
                        new XAttribute("dxfId", registerDxf != null && rule.Format != null
                            ? registerDxf(rule.Format).ToString(CultureInfo.InvariantCulture)
                            : "0"),
                        new XAttribute("priority", priority),
                        new XElement(Xlsx.Main + "formula", rule.Formula));
                    if (string.Equals(rule.Operator, "between", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(rule.Operator, "notBetween", StringComparison.OrdinalIgnoreCase))
                    {
                        cellIs.Add(new XElement(Xlsx.Main + "formula", rule.SecondFormula ?? rule.Formula));
                    }
                    return cellIs;
                }
            }
        }

        /// <summary>生成数据条/双色阶的起止 cfvo 元素。 Builds the min/max cfvo elements of a data bar or two-color scale.</summary>
        /// <param name="rule">规则模型。 Rule model.</param>
        /// <returns>两个 cfvo 元素。 The two cfvo elements.</returns>
        private static XElement[] BoundElements(ExcelConditionalRule rule)
        {
            return new XElement[]
            {
                Bound(rule, true),
                Bound(rule, false)
            };
        }

        /// <summary>生成单端 cfvo：自动取极值时用 min/max 类型，否则用 num 类型。 Builds one cfvo, using min/max for automatic bounds and num for explicit bounds.</summary>
        /// <param name="rule">规则模型。 Rule model.</param>
        /// <param name="minimum">是否为最小端。 Whether this is the minimum bound.</param>
        /// <returns>cfvo 元素。 The cfvo element.</returns>
        private static XElement Bound(ExcelConditionalRule rule, bool minimum)
        {
            bool automatic = minimum ? rule.AutoMinimum : rule.AutoMaximum;
            if (automatic)
            {
                return new XElement(Xlsx.Main + "cfvo", new XAttribute("type", minimum ? "min" : "max"));
            }
            return new XElement(Xlsx.Main + "cfvo",
                new XAttribute("type", "num"),
                new XAttribute("val", FormatNumber(minimum ? rule.Minimum : rule.Maximum)));
        }

        /// <summary>取规则的第 n 个颜色，缺失时回退绿色。 Reads the nth color of a rule, falling back to green when missing.</summary>
        /// <param name="rule">规则模型。 Rule model.</param>
        /// <param name="index">颜色序号。 Color index.</param>
        /// <returns>颜色。 The color.</returns>
        private static OfficeColor ColorOf(ExcelConditionalRule rule, int index)
        {
            if (rule.Colors != null && index < rule.Colors.Count) return rule.Colors[index];
            if (index == 2) return new OfficeColor(0xF8, 0x69, 0x6A);
            if (index == 1) return new OfficeColor(0xFF, 0xEB, 0x84);
            return new OfficeColor(0x63, 0xC3, 0x84);
        }

        /// <summary>按图标集名称前缀推断图标数量（3/4/5 个）。 Infers the icon count from the name prefix (3, 4 or 5 icons).</summary>
        /// <param name="iconSetType">图标集类型。 Icon set type.</param>
        /// <returns>图标数量。 The icon count.</returns>
        private static int IconCount(string iconSetType)
        {
            if (string.IsNullOrEmpty(iconSetType)) return 3;
            char first = iconSetType[0];
            if (first == '4') return 4;
            if (first == '5') return 5;
            return 3;
        }

        /// <summary>把百分比阈值规整到 0~100 并转为文本。 Clamps a percentage threshold to 0-100 and converts it to text.</summary>
        /// <param name="value">百分比。 Percentage.</param>
        /// <returns>文本值。 The text value.</returns>
        private static string Percent(double value)
        {
            if (value < 0) value = 0;
            if (value > 100) value = 100;
            return FormatNumber(value);
        }

        /// <summary>以不变区域设置格式化数值（整数值不带小数点）。 Formats a number with the invariant culture (integral values carry no decimal point).</summary>
        /// <param name="value">数值。 Value.</param>
        /// <returns>文本值。 The text value.</returns>
        private static string FormatNumber(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return "0";
            if (value == Math.Floor(value) && Math.Abs(value) < 1e15)
            {
                return ((long)value).ToString(CultureInfo.InvariantCulture);
            }
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        /// <summary>去掉区域文本里的表名前缀，sqref 只接受本表坐标。 Strips the sheet-name prefix because sqref accepts local coordinates only.</summary>
        /// <param name="range">区域文本。 Range text.</param>
        /// <returns>本表坐标文本。 The local coordinate text.</returns>
        private static string StripSheetPrefix(string range)
        {
            int bang = range.LastIndexOf('!');
            return bang < 0 ? range : range.Substring(bang + 1);
        }
    }
}
