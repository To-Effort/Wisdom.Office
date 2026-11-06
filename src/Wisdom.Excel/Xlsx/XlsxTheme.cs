using System;
using System.Globalization;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// xl/theme/theme1.xml 的颜色方案：把 theme 索引解析成具体 RGB，供 styles.xml 里
    /// <c>&lt;color theme="4" tint="0.5"/&gt;</c> 这类引用取色。
    /// The colour scheme of xl/theme/theme1.xml: turns a theme index into a concrete RGB so style entries such as
    /// <c>&lt;color theme="4" tint="0.5"/&gt;</c> can be resolved instead of falling back to black.
    /// </summary>
    /// <remarks>
    /// SpreadsheetML 的 theme 序号与 theme 部件里的排列顺序不同：0 是 lt1（背景 1）、1 是 dk1（文本 1）、
    /// 2 是 lt2、3 是 dk2，之后才是 accent1-6、hlink 与 folHlink。早先读取端完全不查主题，
    /// 实测 <c>&lt;color theme="4" tint="0.5"/&gt;</c> 读成纯黑，随主题变化的强调色全部丢失。
    /// SpreadsheetML theme indices are swizzled relative to the part order: 0 is lt1, 1 is dk1, 2 is lt2, 3 is dk2,
    /// then accent1-6, hlink and folHlink. The reader used to ignore themes entirely, so a measured
    /// <c>&lt;color theme="4" tint="0.5"/&gt;</c> came back as pure black and every themed accent colour was lost.
    /// </remarks>
    internal sealed class XlsxThemeColors
    {
        /// <summary>SpreadsheetML theme 索引对应的 clrScheme 子元素名。 clrScheme child names in SpreadsheetML theme-index order.</summary>
        private static readonly string[] SchemeOrder =
        {
            "lt1", "dk1", "lt2", "dk2",
            "accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
            "hlink", "folHlink"
        };

        /// <summary>按 theme 索引排列的调色板；缺失主题部件时为 null。 The palette in theme-index order; null when no theme part was available.</summary>
        private readonly OfficeColor[] _palette;

        /// <summary>无主题时的共享实例。 The shared instance used when there is no theme.</summary>
        public static readonly XlsxThemeColors None = new XlsxThemeColors(null);

        /// <summary>私有构造，仅由 <see cref="Load"/> 与 <see cref="None"/> 使用。 Private constructor used only by <see cref="Load"/> and <see cref="None"/>.</summary>
        /// <param name="palette">调色板或 null。 The palette, or null.</param>
        private XlsxThemeColors(OfficeColor[] palette)
        {
            _palette = palette;
        }

        /// <summary>解析 theme 部件；部件缺失或没有可用颜色方案时返回 <see cref="None"/>。
        /// Parses a theme part, returning <see cref="None"/> when the part or its colour scheme is missing.</summary>
        /// <param name="themeXml">xl/theme/theme1.xml 文档。 The theme part document.</param>
        /// <returns>调色板实例。 The palette instance.</returns>
        public static XlsxThemeColors Load(XDocument themeXml)
        {
            if (themeXml == null || themeXml.Root == null) return None;
            XElement scheme = FindByLocalName(themeXml.Root, "clrScheme");
            if (scheme == null) return None;
            OfficeColor[] palette = new OfficeColor[SchemeOrder.Length];
            bool any = false;
            for (int i = 0; i < SchemeOrder.Length; i++)
            {
                XElement entry = FindByLocalName(scheme, SchemeOrder[i]);
                if (entry == null) continue;
                OfficeColor color;
                if (TryReadSchemeColor(entry, out color))
                {
                    palette[i] = color;
                    any = true;
                }
            }
            return any ? new XlsxThemeColors(palette) : None;
        }

        /// <summary>按 theme 索引与 tint 取色。Colors the theme index is unknown or absent are refused.
        /// Resolves a theme index and tint, refusing unknown or missing entries.</summary>
        /// <param name="themeIndex">SpreadsheetML theme 索引。 The SpreadsheetML theme index.</param>
        /// <param name="tint">明暗偏移（-1 到 1）。 Lightness shift from -1 to 1.</param>
        /// <param name="color">解析出的颜色。 The resolved colour.</param>
        /// <returns>成功返回 true。 True when a colour was resolved.</returns>
        public bool TryGet(int themeIndex, double tint, out OfficeColor color)
        {
            color = default(OfficeColor);
            if (_palette == null || themeIndex < 0 || themeIndex >= _palette.Length) return false;
            OfficeColor baseColor = _palette[themeIndex];
            // 未填到的槽位保持 default（Alpha 0）；调色板里不存在合法的透明主题色。
            // Unfilled slots stay default (alpha 0); no legitimate theme colour in the palette is transparent.
            if (baseColor.A == 0) return false;
            color = ApplyTint(baseColor, tint);
            return true;
        }

        /// <summary>把 tint 作用到 RGB 上：正数向白靠拢、负数向黑靠拢，与 Excel 的线性近似一致。
        /// Applies a tint to an RGB value: positive moves toward white, negative toward black, matching Excel's
        /// linear approximation.</summary>
        /// <param name="color">基准颜色。 Base colour.</param>
        /// <param name="tint">明暗偏移。 Lightness shift.</param>
        /// <returns>偏移后的颜色。 The shifted colour.</returns>
        private static OfficeColor ApplyTint(OfficeColor color, double tint)
        {
            if (tint == 0.0 || double.IsNaN(tint)) return color;
            if (tint > 1.0) tint = 1.0;
            if (tint < -1.0) tint = -1.0;
            return new OfficeColor(
                Shift(color.R, tint),
                Shift(color.G, tint),
                Shift(color.B, tint));
        }

        /// <summary>按 tint 平移单个通道。 Shifts one channel by the tint.</summary>
        /// <param name="channel">通道值。 Channel value.</param>
        /// <param name="tint">明暗偏移。 Lightness shift.</param>
        /// <returns>平移后的通道值。 The shifted channel value.</returns>
        private static byte Shift(byte channel, double tint)
        {
            double value = tint > 0
                ? channel * (1.0 - tint) + 255.0 * tint
                : channel * (1.0 + tint);
            if (value < 0.0) value = 0.0;
            if (value > 255.0) value = 255.0;
            return (byte)Math.Round(value, MidpointRounding.AwayFromZero);
        }

        /// <summary>读取 clrScheme 条目内的颜色：srgbClr 的 val，其次 sysClr 的 lastClr。
        /// Reads the colour inside a clrScheme entry: the srgbClr val first, then the sysClr lastClr.</summary>
        /// <param name="entry">dk1/lt1/accentN 等条目。 A dk1/lt1/accentN entry.</param>
        /// <param name="color">解析出的颜色。 The resolved colour.</param>
        /// <returns>成功返回 true。 True when a colour was parsed.</returns>
        private static bool TryReadSchemeColor(XElement entry, out OfficeColor color)
        {
            color = default(OfficeColor);
            foreach (XElement child in entry.Elements())
            {
                string name = child.Name.LocalName;
                if (name == "srgbClr")
                {
                    string value = child.Attribute("val") == null ? null : child.Attribute("val").Value;
                    OfficeColor parsed;
                    if (Xlsx.TryParseColor(value, out parsed))
                    {
                        color = parsed;
                        return true;
                    }
                }
                else if (name == "sysClr")
                {
                    XAttribute last = child.Attribute("lastClr");
                    OfficeColor parsed;
                    if (last != null && Xlsx.TryParseColor(last.Value, out parsed))
                    {
                        color = parsed;
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>按本地名称深度查找第一个元素，忽略命名空间差异。
        /// Finds the first descendant matching a local name, ignoring namespace differences.</summary>
        /// <param name="root">搜索起点。 Search root.</param>
        /// <param name="localName">本地名称。 Local name.</param>
        /// <returns>匹配元素或 null。 The match, or null.</returns>
        private static XElement FindByLocalName(XElement root, string localName)
        {
            if (root == null) return null;
            foreach (XElement element in root.Descendants())
            {
                if (string.Equals(element.Name.LocalName, localName, StringComparison.Ordinal)) return element;
            }
            return null;
        }
    }

    /// <summary>
    /// OOXML color 元素统一取色：rgb 优先，其次 theme（含 tint），再次 indexed，最后回退。
    /// Shared resolution of an OOXML color element: rgb first, then theme (with tint), then indexed, then fallback.
    /// </summary>
    internal static class XlsxColors
    {
        /// <summary>BIFF/旧式索引颜色表（按 indexed 属性取值）。 BIFF/legacy indexed palette addressed by the indexed attribute.</summary>
        private static readonly byte[][] IndexedColors =
        {
            new byte[] { 0x00, 0x00, 0x00 }, new byte[] { 0xFF, 0xFF, 0xFF },
            new byte[] { 0xFF, 0x00, 0x00 }, new byte[] { 0x00, 0xFF, 0x00 },
            new byte[] { 0x00, 0x00, 0xFF }, new byte[] { 0xFF, 0xFF, 0x00 },
            new byte[] { 0xFF, 0x00, 0xFF }, new byte[] { 0x00, 0xFF, 0xFF },
            new byte[] { 0x00, 0x00, 0x00 }, new byte[] { 0xFF, 0xFF, 0xFF },
            new byte[] { 0xFF, 0x00, 0x00 }, new byte[] { 0x00, 0xFF, 0x00 },
            new byte[] { 0x00, 0x00, 0xFF }, new byte[] { 0xFF, 0xFF, 0x00 },
            new byte[] { 0xFF, 0x00, 0xFF }, new byte[] { 0x00, 0xFF, 0xFF },
            new byte[] { 0x80, 0x00, 0x00 }, new byte[] { 0x00, 0x80, 0x00 },
            new byte[] { 0x00, 0x00, 0x80 }, new byte[] { 0x80, 0x80, 0x00 },
            new byte[] { 0x80, 0x00, 0x80 }, new byte[] { 0x00, 0x80, 0x80 },
            new byte[] { 0xC0, 0xC0, 0xC0 }, new byte[] { 0x80, 0x80, 0x80 },
            new byte[] { 0x99, 0x99, 0xFF }, new byte[] { 0x99, 0x33, 0x66 },
            new byte[] { 0xFF, 0xFF, 0xCC }, new byte[] { 0xCC, 0xFF, 0xFF },
            new byte[] { 0x66, 0x00, 0x66 }, new byte[] { 0xFF, 0x80, 0x80 },
            new byte[] { 0x00, 0x66, 0xCC }, new byte[] { 0xCC, 0xCC, 0xFF },
            new byte[] { 0x00, 0x00, 0x80 }, new byte[] { 0xFF, 0x00, 0xFF },
            new byte[] { 0xFF, 0xFF, 0x00 }, new byte[] { 0x00, 0xFF, 0xFF },
            new byte[] { 0x80, 0x00, 0x80 }, new byte[] { 0x80, 0x00, 0x00 },
            new byte[] { 0x00, 0x80, 0x80 }, new byte[] { 0x00, 0x00, 0xFF },
            new byte[] { 0x00, 0xCC, 0xFF }, new byte[] { 0xCC, 0xFF, 0xFF },
            new byte[] { 0xCC, 0xFF, 0xCC }, new byte[] { 0xFF, 0xFF, 0x99 },
            new byte[] { 0x99, 0xCC, 0xFF }, new byte[] { 0xFF, 0x99, 0xCC },
            new byte[] { 0xCC, 0x99, 0xFF }, new byte[] { 0xFF, 0xCC, 0x99 },
            new byte[] { 0x33, 0x66, 0xFF }, new byte[] { 0x33, 0xCC, 0xCC },
            new byte[] { 0x99, 0xCC, 0x00 }, new byte[] { 0xFF, 0xCC, 0x00 },
            new byte[] { 0xFF, 0x99, 0x00 }, new byte[] { 0xFF, 0x66, 0x00 },
            new byte[] { 0x66, 0x66, 0x99 }, new byte[] { 0x96, 0x96, 0x96 },
            new byte[] { 0x00, 0x33, 0x66 }, new byte[] { 0x33, 0x99, 0x66 },
            new byte[] { 0x00, 0x33, 0x00 }, new byte[] { 0x33, 0x33, 0x00 },
            new byte[] { 0x99, 0x33, 0x33 }, new byte[] { 0x33, 0x33, 0x99 },
            new byte[] { 0x99, 0x99, 0x33 }
        };

        /// <summary>解析 color 元素的属性并返回具体 RGB。 Reads a color element's attributes into a concrete RGB value.</summary>
        /// <param name="element">color/bgColor/fgColor 元素。 The color, bgColor or fgColor element.</param>
        /// <param name="theme">主题调色板；可为 <see cref="XlsxThemeColors.None"/>。 Theme palette; may be <see cref="XlsxThemeColors.None"/>.</param>
        /// <param name="fallback">无法解析时的回退颜色。 Fallback colour when nothing resolves.</param>
        /// <returns>解析出的颜色。 The resolved colour.</returns>
        public static OfficeColor Read(XElement element, XlsxThemeColors theme, OfficeColor fallback)
        {
            OfficeColor parsed;
            return TryRead(element, theme, out parsed) ? parsed : fallback;
        }

        /// <summary>该 color 元素是否表示"自动色"（indexed 64 或 auto=1）。 Whether the color element means "automatic" (indexed 64 or auto=1).</summary>
        /// <param name="element">color 元素。 The color element.</param>
        /// <returns>自动色返回 true。 True for an automatic colour.</returns>
        public static bool IsAutomatic(XElement element)
        {
            if (element == null) return false;
            return Xlsx.Attribute(element, "indexed") == "64" || Xlsx.Attribute(element, "auto") == "1";
        }

        /// <summary>按 rgb → theme(含 tint) → indexed 的顺序解析颜色；元素只有 auto 或属性全不可用时返回 false，
        /// 调用方据此保留原值。
        /// Resolves the colour in rgb, then theme (with tint), then indexed order; returns false when the element only
        /// carries auto or nothing parses, so the caller keeps its previous value.</summary>
        /// <param name="element">color/bgColor/fgColor 元素。 The color, bgColor or fgColor element.</param>
        /// <param name="theme">主题调色板。 Theme palette.</param>
        /// <param name="color">解析出的颜色。 The resolved colour.</param>
        /// <returns>可解析返回 true。 True when a colour was resolved.</returns>
        public static bool TryRead(XElement element, XlsxThemeColors theme, out OfficeColor color)
        {
            color = default(OfficeColor);
            if (element == null) return false;
            OfficeColor parsed;

            string rgb = Xlsx.Attribute(element, "rgb");
            if (!string.IsNullOrEmpty(rgb) && Xlsx.TryParseColor(rgb, out parsed))
            {
                color = parsed;
                return true;
            }

            // 主题色必须在 indexed 之前判定：Excel 写主题色时 indexed 常带 0，先查索引会把强调色读成黑色。
            // Theme colours are tested before indexed: Excel usually carries indexed="0" alongside a themed colour,
            // so resolving the index first would turn every accent into black.
            string themeText = Xlsx.Attribute(element, "theme");
            int themeIndex;
            if (!string.IsNullOrEmpty(themeText)
                && int.TryParse(themeText, NumberStyles.Integer, CultureInfo.InvariantCulture, out themeIndex)
                && theme != null && theme.TryGet(themeIndex, Xlsx.DoubleAttribute(element, "tint", 0.0), out parsed))
            {
                color = parsed;
                return true;
            }

            string indexed = Xlsx.Attribute(element, "indexed");
            int index;
            if (!string.IsNullOrEmpty(indexed)
                && int.TryParse(indexed, NumberStyles.Integer, CultureInfo.InvariantCulture, out index))
            {
                if (index >= 0 && index < IndexedColors.Length)
                {
                    byte[] value = IndexedColors[index];
                    color = new OfficeColor(value[0], value[1], value[2]);
                    return true;
                }
                if (index == 64 || index == 65)
                {
                    color = index == 64 ? OfficeColor.Black : OfficeColor.White;
                    return true;
                }
            }

            return false;
        }
    }
}
