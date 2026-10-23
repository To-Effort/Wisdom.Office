using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Ppt
{
    /// <summary>
    /// PPTX 读取器：将 PresentationML（.pptx）压缩包解析为统一演示文稿模型。
    /// 占位符在真实文件里常常不带自身几何，因此几何按「形状自身 → 版式同名占位符 → 母版同名占位符 →
    /// 按幻灯片尺寸推算的默认矩形」四级回落；单个部件解析失败只跳过该部件，不影响其余内容。
    /// PPTX reader that parses a PresentationML (.pptx) package into the unified presentation model.
    /// Real files often omit the geometry of placeholders, so geometry falls back through four levels:
    /// the shape's own transform, the matching layout placeholder, the matching master placeholder, and a
    /// default rectangle derived from the slide size. A failure in one part only skips that part.
    /// </summary>
    public static class PptxReader
    {
        /// <summary>备注页关系类型后缀 / Relationship type suffix of a notes slide part.</summary>
        private const string NotesSlideSuffix = "/notesSlide";

        /// <summary>
        /// 表格单元格跨行/跨列上限（真实幻灯片表格远小于此），防御恶意 gridSpan、rowSpan 撑爆下游按跨度分配的输出。
        /// The cap on a table cell's row or column span (real slide tables are far below it); it guards against a hostile
        /// gridSpan/rowSpan blowing up downstream output that allocates per span.
        /// </summary>
        private const int MaxCellSpan = 1024;

        /// <summary>EMU 矩形：形状外接框，单位与 PresentationML 一致。An EMU rectangle: a shape bounding box in PresentationML units.</summary>
        private sealed class Box
        {
            /// <summary>左边距（EMU）。Left offset in EMU.</summary>
            public readonly long X;
            /// <summary>上边距（EMU）。Top offset in EMU.</summary>
            public readonly long Y;
            /// <summary>宽度（EMU）。Width in EMU.</summary>
            public readonly long CX;
            /// <summary>高度（EMU）。Height in EMU.</summary>
            public readonly long CY;

            /// <summary>按 EMU 数值创建矩形。Creates a rectangle from EMU values.</summary>
            /// <param name="x">左边距（EMU）。Left offset in EMU.</param>
            /// <param name="y">上边距（EMU）。Top offset in EMU.</param>
            /// <param name="cx">宽度（EMU）。Width in EMU.</param>
            /// <param name="cy">高度（EMU）。Height in EMU.</param>
            public Box(long x, long y, long cx, long cy)
            {
                X = x;
                Y = y;
                CX = cx;
                CY = cy;
            }
        }

        /// <summary>
        /// 组合形状的累计坐标变换（平移 + 各轴缩放），把子形状坐标空间的 EMU 映射到幻灯片坐标空间。
        /// The accumulated transform of grouped shapes (translation plus per-axis scaling) that maps a
        /// child coordinate space into slide coordinates.
        /// </summary>
        private sealed class ShapeTransform
        {
            /// <summary>横向缩放系数。The horizontal scale factor.</summary>
            public readonly double ScaleX;
            /// <summary>纵向缩放系数。The vertical scale factor.</summary>
            public readonly double ScaleY;
            /// <summary>横向平移量（EMU）。The horizontal translation in EMU.</summary>
            public readonly double OffsetX;
            /// <summary>纵向平移量（EMU）。The vertical translation in EMU.</summary>
            public readonly double OffsetY;

            /// <summary>创建恒等变换。Creates the identity transform.</summary>
            public ShapeTransform()
            {
                ScaleX = 1.0;
                ScaleY = 1.0;
                OffsetX = 0.0;
                OffsetY = 0.0;
            }

            /// <summary>创建指定缩放与平移的变换。Creates a transform with the given scaling and translation.</summary>
            /// <param name="scaleX">横向缩放系数。Horizontal scale factor.</param>
            /// <param name="scaleY">纵向缩放系数。Vertical scale factor.</param>
            /// <param name="offsetX">横向平移量（EMU）。Horizontal translation in EMU.</param>
            /// <param name="offsetY">纵向平移量（EMU）。Vertical translation in EMU.</param>
            public ShapeTransform(double scaleX, double scaleY, double offsetX, double offsetY)
            {
                ScaleX = scaleX;
                ScaleY = scaleY;
                OffsetX = offsetX;
                OffsetY = offsetY;
            }

            /// <summary>是否为恒等变换（无需换算坐标）。Whether the transform is the identity (no coordinate conversion needed).</summary>
            public bool IsIdentity
            {
                get { return ScaleX == 1.0 && ScaleY == 1.0 && OffsetX == 0.0 && OffsetY == 0.0; }
            }

            /// <summary>
            /// 在现有变换外再套一层分组的 off/ext/chOff/chExt 映射，得到子形状到幻灯片空间的复合变换。
            /// Wraps the group's off/ext/chOff/chExt mapping around the current transform, producing the
            /// composite child-space-to-slide-space transform.
            /// </summary>
            /// <param name="offX">分组目标框左边距（EMU）。Group offset X in EMU.</param>
            /// <param name="offY">分组目标框上边距（EMU）。Group offset Y in EMU.</param>
            /// <param name="extX">分组目标框宽度（EMU）。Group extent width in EMU.</param>
            /// <param name="extY">分组目标框高度（EMU）。Group extent height in EMU.</param>
            /// <param name="chOffX">子坐标空间原点左边距（EMU）。Child coordinate space origin X in EMU.</param>
            /// <param name="chOffY">子坐标空间原点上边距（EMU）。Child coordinate space origin Y in EMU.</param>
            /// <param name="chExtX">子坐标空间宽度（EMU）。Child coordinate space width in EMU.</param>
            /// <param name="chExtY">子坐标空间高度（EMU）。Child coordinate space height in EMU.</param>
            /// <returns>复合后的变换；子空间尺寸非法时保持原变换（按子形状原始 EMU 直出）。The composed transform; unchanged when the child extent is invalid, so raw EMU values are used.</returns>
            public ShapeTransform WithGroup(double offX, double offY, double extX, double extY, double chOffX, double chOffY, double chExtX, double chExtY)
            {
                if (chExtX <= 0.0 || chExtY <= 0.0) return this;
                double scaleX = extX > 0.0 ? extX / chExtX : 1.0;
                double scaleY = extY > 0.0 ? extY / chExtY : 1.0;
                return new ShapeTransform(
                    ScaleX * scaleX,
                    ScaleY * scaleY,
                    offX + (OffsetX - chOffX) * scaleX,
                    offY + (OffsetY - chOffY) * scaleY);
            }
        }

        /// <summary>
        /// 包级读取会话：持有包访问器、目标模型与幻灯片尺寸，并惰性解析、缓存主题色表。
        /// Package-level read session: holds the package accessor, the target model and the slide size,
        /// and lazily parses and caches the theme colour table.
        /// </summary>
        private sealed class ReadSession
        {
            /// <summary>包访问器。The package accessor.</summary>
            public readonly PptPackage Package;
            /// <summary>演示文稿主部件路径。The presentation part path.</summary>
            public readonly string PresentationPart;
            /// <summary>正在填充的演示文稿模型。The presentation model being populated.</summary>
            public readonly OfficePresentation Presentation;
            /// <summary>主题 XML 根元素缓存（a:theme）；null 表示已解析且无可用主题。Cached a:theme root; null once parsed when no theme part resolves.</summary>
            private XElement _themeRoot;
            /// <summary>主题是否已解析（区分"未解析"与"解析后为 null"）。Whether the theme root was already resolved (separates unparsed from parsed-but-null).</summary>
            private bool _themeResolved;
            /// <summary>主题色缓存：clrScheme 子元素名 → 实际颜色；null 表示尚未解析。Cached theme colours keyed by clrScheme child name; null until parsed.</summary>
            private Dictionary<string, OfficeColor> _theme;

            /// <summary>创建读取会话。Creates a read session.</summary>
            /// <param name="package">包访问器。The package accessor.</param>
            /// <param name="presentationPart">演示文稿主部件路径。The presentation part path.</param>
            /// <param name="presentation">目标模型。The target model.</param>
            public ReadSession(PptPackage package, string presentationPart, OfficePresentation presentation)
            {
                Package = package;
                PresentationPart = presentationPart;
                Presentation = presentation;
            }

            /// <summary>
            /// 读取填充元素宿主（spPr / tcPr / bg / ln）里的颜色：solidFill 优先，其次 fillRef，
            /// noFill 与渐变、图案填充一律返回 null（统一模型无法表达继承）。
            /// Reads the colour from a fill host (spPr / tcPr / bg / ln): solidFill first, then fillRef;
            /// noFill plus gradient and pattern fills all return null, since the unified model cannot express inheritance.
            /// </summary>
            /// <param name="host">含填充子元素的父元素。The parent element that may carry fill children.</param>
            /// <returns>解析出的颜色，或 null（无填充/未知）。The resolved colour, or null for no fill or an unknown fill.</returns>
            public OfficeColor? ReadFill(XElement host)
            {
                XElement solid = Ppt.DrawingChild(host, "solidFill");
                if (solid != null) return ReadColorElement(FirstChildElement(solid));
                XElement fillRef = Ppt.DrawingChild(host, "fillRef");
                if (fillRef != null) return ReadColorElement(FirstChildElement(fillRef));
                return null;
            }

            /// <summary>
            /// 解析单个颜色元素（a:srgbClr / a:schemeClr / a:sysClr），无法解析时返回 null。
            /// 颜色上的 tint/shade/alpha 等子级调整不参与计算：统一模型只有一个 RGB 值。
            /// Parses one colour element (a:srgbClr / a:schemeClr / a:sysClr), returning null when unresolvable.
            /// Child adjustments such as tint, shade and alpha are not applied, because the unified model stores
            /// a single RGB value only.
            /// </summary>
            /// <param name="color">颜色元素。The colour element.</param>
            /// <returns>解析出的颜色或 null。The resolved colour or null.</returns>
            public OfficeColor? ReadColorElement(XElement color)
            {
                if (color == null || color.Name.Namespace != Ppt.Drawing) return null;
                string name = color.Name.LocalName;
                OfficeColor parsed;
                if (name == "srgbClr")
                {
                    if (Ppt.TryParseColor(Ppt.Attribute(color, "val"), out parsed)) return parsed;
                    return null;
                }
                if (name == "sysClr")
                {
                    if (Ppt.TryParseColor(Ppt.Attribute(color, "lastClr"), out parsed)) return parsed;
                    return null;
                }
                if (name == "schemeClr") return ResolveSchemeColor(Ppt.Attribute(color, "val"));
                return null;
            }

            /// <summary>
            /// 读取 a:rPr 运行格式：b/i/u/strike 开关、sz 字号（百分之一磅）、latin/ea 字体名与 solidFill 颜色。
            /// 元素不存在或没有任何格式属性时返回 null，使文本运行沿用占位符样式而不是被写成默认值。
            /// Reads an a:rPr run format: the b/i/u/strike toggles, sz in hundredths of a point, the latin/ea
            /// font names and the solidFill colour. Returns null when the element is absent or carries no
            /// formatting, so the run inherits the placeholder style instead of being pinned to defaults.
            /// </summary>
            /// <param name="rPr">a:rPr 元素，可为 null。The a:rPr element; may be null.</param>
            /// <returns>解析出的文本格式，或 null。The parsed text format, or null.</returns>
            public TextFormat ReadRunFormat(XElement rPr)
            {
                if (rPr == null) return null;
                TextFormat format = new TextFormat();
                bool any = false;
                if (Ppt.ParseBool(Ppt.Attribute(rPr, "b"), false)) { format.Bold = true; any = true; }
                if (Ppt.ParseBool(Ppt.Attribute(rPr, "i"), false)) { format.Italic = true; any = true; }
                string underline = Ppt.Attribute(rPr, "u");
                if (underline != null && underline != "none") { format.Underline = true; any = true; }
                string strike = Ppt.Attribute(rPr, "strike");
                if (strike == "sngStrike" || strike == "dblStrike" || strike == "heavyStrike") { format.Strikeout = true; any = true; }
                string baseline = Ppt.Attribute(rPr, "baseline");
                int offset = Ppt.ParseInt(baseline, 0);
                if (offset > 0) { format.Superscript = true; any = true; }
                else if (offset < 0) { format.Subscript = true; any = true; }
                double size = Ppt.ParseDouble(Ppt.Attribute(rPr, "sz"), 0.0);
                if (size > 0.0) { format.FontSize = size / 100.0; any = true; }
                XElement latin = Ppt.DrawingChild(rPr, "latin");
                if (latin != null)
                {
                    string typeface = Ppt.Attribute(latin, "typeface");
                    if (!string.IsNullOrEmpty(typeface)) { format.AsciiFontFamily = typeface; format.FontFamily = typeface; any = true; }
                }
                XElement eastAsia = Ppt.DrawingChild(rPr, "ea");
                if (eastAsia != null)
                {
                    string typeface = Ppt.Attribute(eastAsia, "typeface");
                    if (!string.IsNullOrEmpty(typeface)) { format.EastAsiaFontFamily = typeface; any = true; }
                }
                OfficeColor? color = ReadFill(rPr);
                if (color.HasValue) { format.Color = color.Value; any = true; }
                return any ? format : null;
            }

            /// <summary>
            /// 把 schemeClr 的语义值映射为主题实际颜色。accent1~6 等直接取同名色，
            /// tx1/bg1/tx2/bg2 是 dk1/lt1/dk2/lt2 的别名；phClr 依赖形状样式引用，留空不猜。
            /// Maps a schemeClr semantic value to the actual theme colour. accent1 to accent6 are read as-is,
            /// tx1/bg1/tx2/bg2 are aliases of dk1/lt1/dk2/lt2, and phClr, which depends on the shape style
            /// reference, is left unresolved rather than guessed.
            /// </summary>
            /// <param name="value">schemeClr 的 val 属性。The val attribute of a schemeClr.</param>
            /// <returns>解析出的颜色或 null。The resolved colour or null.</returns>
            private OfficeColor? ResolveSchemeColor(string value)
            {
                if (string.IsNullOrEmpty(value)) return null;
                Dictionary<string, OfficeColor> theme = Theme();
                if (theme.Count == 0) return null;
                string key = value;
                if (key == "phClr") return null;
                OfficeColor color;
                if (theme.TryGetValue(key, out color)) return color;
                return null;
            }

            /// <summary>
            /// 惰性解析主题色表：母版部件经 presentation 的 slideMaster 关系取得（取不到时回退
            /// sldMasterIdLst 的第一个 r:id），再由母版 rels 的 theme 关系定位 theme1.xml，解析一次后缓存。
            /// Lazily resolves the theme colour table: the master part comes from the presentation's
            /// slideMaster relationship (falling back to the first r:id of sldMasterIdLst), the theme part
            /// from the master's theme relationship; the table is parsed once and then cached.
            /// </summary>
            /// <returns>主题色表（可能为空字典）。The theme colour table, possibly empty.</returns>
            private Dictionary<string, OfficeColor> Theme()
            {
                if (_theme != null) return _theme;
                Dictionary<string, OfficeColor> colors = new Dictionary<string, OfficeColor>(StringComparer.OrdinalIgnoreCase);
                XElement themeRoot = ThemeRoot();
                if (themeRoot != null)
                {
                    XElement themeElements = Ppt.DrawingChild(themeRoot, "themeElements");
                    XElement colorScheme = Ppt.DrawingChild(themeElements, "clrScheme");
                    if (colorScheme != null)
                    {
                        // clrScheme 的子元素名即 schemeClr 的语义值，故按声明顺序整体登记，
                        // hlink/folHlink 这类不常列举的名字也因此可解析。
                        // The child names of clrScheme are exactly the schemeClr semantic values, so every
                        // child is registered in document order; rarely listed names such as hlink and
                        // folHlink therefore stay resolvable.
                        foreach (XElement entry in colorScheme.Elements())
                        {
                            AddThemeColor(colors, entry);
                        }
                    }
                    // tx1/bg1/tx2/bg2 是 dk1/lt1/dk2/lt2 的别名，只有被别名键解析成功才登记。
                    // tx1/bg1/tx2/bg2 alias dk1/lt1/dk2/lt2 and are registered only when the aliased key resolved.
                    AddThemeAlias(colors, "tx1", "dk1");
                    AddThemeAlias(colors, "bg1", "lt1");
                    AddThemeAlias(colors, "tx2", "dk2");
                    AddThemeAlias(colors, "bg2", "lt2");
                }
                _theme = colors;
                return colors;
            }

            /// <summary>
            /// 惰性定位并缓存 a:theme 根元素：母版部件经 presentation 的 slideMaster 关系取得
            /// （取不到时回退 sldMasterIdLst 的第一个 r:id），再由母版 rels 的 theme 关系定位 theme1.xml。
            /// Lazily locates and caches the a:theme root: the master part comes from the presentation's
            /// slideMaster relationship (falling back to the first r:id of sldMasterIdLst), the theme part
            /// from the master's theme relationship.
            /// </summary>
            /// <returns>主题根元素或 null。The theme root or null.</returns>
            private XElement ThemeRoot()
            {
                if (_themeResolved) return _themeRoot;
                _themeResolved = true;
                string masterPart = Package.GetRelationshipTargetByType(PresentationPart, Ppt.TypeSlideMaster);
                if (string.IsNullOrEmpty(masterPart)) masterPart = FirstMasterFromXml();
                if (!string.IsNullOrEmpty(masterPart))
                {
                    string themePart = Package.GetRelationshipTargetByType(masterPart, Ppt.TypeTheme);
                    XDocument themeXml = Package.GetXml(themePart);
                    if (themeXml != null && themeXml.Root != null) _themeRoot = themeXml.Root;
                }
                return _themeRoot;
            }

            /// <summary>
            /// 读取主题深改信息到模型：clrScheme 的六个可覆盖阶段（dk1/lt1/dk2/lt2/hlink/folHlink）、
            /// majorFont/minorFont 的 latin 字体名与主题名。读回的模型字段全部显式赋值，
            /// 即使与内置默认相同；写出器把 sysClr 与等效 srgbClr 常量渲染为同一颜色，重存结果稳定。
            /// Reads deep-edit theme data into the model: the six overridable clrScheme phases,
            /// the major/minor latin font names and the theme name. Every read field is assigned
            /// explicitly even when it equals the built-in default; the writer renders sysClr and its
            /// equivalent srgbClr constant as the same colour, so resaving stays stable.
            /// </summary>
            public void ReadTheme()
            {
                XElement themeRoot = ThemeRoot();
                if (themeRoot == null) return;
                XElement themeElements = Ppt.DrawingChild(themeRoot, "themeElements");
                XElement colorScheme = Ppt.DrawingChild(themeElements, "clrScheme");
                XElement fontScheme = Ppt.DrawingChild(themeElements, "fontScheme");
                PresentationTheme theme = new PresentationTheme();
                if (colorScheme != null)
                {
                    OfficeColor? text = PhaseColor(colorScheme, "dk1");
                    OfficeColor? background = PhaseColor(colorScheme, "lt1");
                    OfficeColor? textSecondary = PhaseColor(colorScheme, "dk2");
                    OfficeColor? backgroundSecondary = PhaseColor(colorScheme, "lt2");
                    OfficeColor? hyperlink = PhaseColor(colorScheme, "hlink");
                    OfficeColor? followedHyperlink = PhaseColor(colorScheme, "folHlink");
                    theme.Text = text;
                    theme.Background = background;
                    theme.TextSecondary = textSecondary;
                    theme.BackgroundSecondary = backgroundSecondary;
                    theme.Hyperlink = hyperlink;
                    theme.FollowedHyperlink = followedHyperlink;
                    theme.Name = Ppt.Attribute(colorScheme, "name");
                }
                if (fontScheme != null)
                {
                    theme.FontMajor = FontTypeface(Ppt.DrawingChild(fontScheme, "majorFont"));
                    theme.FontMinor = FontTypeface(Ppt.DrawingChild(fontScheme, "minorFont"));
                }
                Presentation.Theme = theme;
            }

            /// <summary>取 clrScheme 指定阶段的颜色值。Reads the colour of one clrScheme phase.</summary>
            /// <param name="colorScheme">clrScheme 元素。The clrScheme element.</param>
            /// <param name="phase">阶段本地名。The phase local name.</param>
            /// <returns>颜色或 null。The colour or null.</returns>
            private OfficeColor? PhaseColor(XElement colorScheme, string phase)
            {
                XElement entry = Ppt.DrawingChild(colorScheme, phase);
                if (entry == null) return null;
                return ReadColorElement(FirstChildElement(entry));
            }

            /// <summary>取字体方案的 latin 字体名。Reads the latin typeface name of a font scheme section.</summary>
            /// <param name="section">majorFont 或 minorFont 元素。The majorFont or minorFont element.</param>
            /// <returns>字体名或 null。The typeface or null.</returns>
            private static string FontTypeface(XElement section)
            {
                XElement latin = Ppt.DrawingChild(section, "latin");
                string typeface = Ppt.Attribute(latin, "typeface");
                return string.IsNullOrEmpty(typeface) ? null : typeface;
            }

            /// <summary>按 sldMasterIdLst 的第一个关系 Id 取母版部件路径。Reads the first master part path from sldMasterIdLst by its relationship Id.</summary>
            /// <returns>母版部件路径或 null。The master part path or null.</returns>
            private string FirstMasterFromXml()
            {
                XDocument presentationXml = Package.GetXml(PresentationPart);
                if (presentationXml == null || presentationXml.Root == null) return null;
                XElement masters = Ppt.Child(presentationXml.Root, "sldMasterIdLst");
                foreach (XElement master in Ppt.Children(masters, "sldMasterId"))
                {
                    string part = Package.GetRelationshipTarget(PresentationPart, Ppt.Attribute(master, Ppt.R + "id"));
                    if (!string.IsNullOrEmpty(part)) return part;
                }
                return null;
            }

            /// <summary>登记 clrScheme 的一个颜色条目（取第一个可解析的颜色子元素）。Registers one clrScheme entry, using its first parsable colour child.</summary>
            /// <param name="colors">目标色表。The colour table.</param>
            /// <param name="entry">clrScheme 的子元素。The clrScheme child element.</param>
            private void AddThemeColor(Dictionary<string, OfficeColor> colors, XElement entry)
            {
                if (entry == null) return;
                OfficeColor? color = ReadColorElement(FirstChildElement(entry));
                if (color != null) colors[entry.Name.LocalName] = color.Value;
            }

            /// <summary>为一个已解析的主题色登记别名键；被别名的键不存在时不写入。Registers an alias key for a resolved theme colour, writing nothing when the aliased key is absent.</summary>
            /// <param name="colors">色表。The colour table.</param>
            /// <param name="alias">别名键。The alias key.</param>
            /// <param name="key">被引用的键。The referenced key.</param>
            private static void AddThemeAlias(Dictionary<string, OfficeColor> colors, string alias, string key)
            {
                OfficeColor color;
                if (colors.TryGetValue(key, out color)) colors[alias] = color;
            }

            /// <summary>读取幻灯片尺寸部件并写入模型；缺件或数值非法时保留 960x540 磅默认值。Reads the slide size into the model, keeping the 960x540 point defaults when the element is missing or invalid.</summary>
            /// <param name="presentationXml">presentation.xml 文档。The presentation.xml document.</param>
            public void ReadSlideSize(XDocument presentationXml)
            {
                if (presentationXml == null || presentationXml.Root == null) return;
                XElement size = Ppt.Child(presentationXml.Root, "sldSz");
                long cx = ParseLong(Ppt.Attribute(size, "cx"), 0);
                long cy = ParseLong(Ppt.Attribute(size, "cy"), 0);
                if (cx > 0) Presentation.SlideWidth = Ppt.ToPoints(cx);
                if (cy > 0) Presentation.SlideHeight = Ppt.ToPoints(cy);
            }

            /// <summary>取得父元素的第一个子元素。Returns the first child element of a parent.</summary>
            /// <param name="parent">父元素。The parent element.</param>
            /// <returns>第一个子元素或 null。The first child element or null.</returns>
            private static XElement FirstChildElement(XElement parent)
            {
                if (parent == null) return null;
                foreach (XElement child in parent.Elements()) return child;
                return null;
            }
        }

        /// <summary>
        /// 单张幻灯片的读取器：负责该幻灯片对应的版式/母版占位符几何缓存、形状映射、备注与切换效果。
        /// Reader for one slide: caches the layout and master placeholder geometry used as fallbacks and
        /// performs shape mapping, notes and transition reading.
        /// </summary>
        private sealed class SlideReader
        {
            /// <summary>包级读取会话。The package-level read session.</summary>
            private readonly ReadSession _session;
            /// <summary>当前幻灯片部件路径。The current slide part path.</summary>
            private readonly string _slidePart;
            /// <summary>版式占位符几何表：占位符键 → EMU 矩形。Layout placeholder geometry keyed by placeholder key.</summary>
            private readonly Dictionary<string, Box> _layoutBoxes;
            /// <summary>母版占位符几何表：占位符键 → EMU 矩形。Master placeholder geometry keyed by placeholder key.</summary>
            private readonly Dictionary<string, Box> _masterBoxes;
            /// <summary>cNvPr id 到形状映射，供 p:timing 的动画目标定位。Maps cNvPr ids to shapes so that p:timing animation targets can be resolved.</summary>
            private readonly Dictionary<int, SlideShape> _shapeById = new Dictionary<int, SlideShape>();

            /// <summary>为一张幻灯片创建读取器，同时预读其版式与母版的占位符几何。Creates the reader for one slide, pre-reading placeholder geometry from its layout and master.</summary>
            /// <param name="session">包级读取会话。The package-level read session.</param>
            /// <param name="slidePart">幻灯片部件路径。The slide part path.</param>
            public SlideReader(ReadSession session, string slidePart)
            {
                _session = session;
                _slidePart = slidePart;
                string layoutPart = session.Package.GetRelationshipTargetByType(slidePart, Ppt.TypeSlideLayout);
                string masterPart = session.Package.GetRelationshipTargetByType(layoutPart, Ppt.TypeSlideMaster);
                _layoutBoxes = CollectPlaceholderBoxes(session.Package.GetXml(layoutPart));
                _masterBoxes = CollectPlaceholderBoxes(session.Package.GetXml(masterPart));
            }

            /// <summary>
            /// 读取一张幻灯片：背景、形状树（含组合递归）、版式类别推断、切换效果与备注。
            /// Reads one slide: background, shape tree (groups expanded recursively), inferred layout kind,
            /// transition and speaker notes.
            /// </summary>
            /// <param name="slideXml">幻灯片部件 XML，可为 null。The slide part XML, possibly null.</param>
            /// <returns>读取出的幻灯片；XML 不可用时返回空白幻灯片。The slide, or a blank one when the XML is unavailable.</returns>
            public Slide Read(XDocument slideXml)
            {
                Slide slide = new Slide();
                if (slideXml == null || slideXml.Root == null) return slide;
                XElement root = slideXml.Root;
                XElement cSld = Ppt.Child(root, "cSld");
                ReadBackground(cSld, slide);
                ReadShapes(Ppt.Child(cSld, "spTree"), slide, new ShapeTransform(), 0);
                slide.Layout = InferLayoutKind(slide);
                XElement transition = FindTransition(root);
                slide.Transition = ReadTransition(transition);
                slide.TransitionDuration = ReadTransitionDuration(transition);
                slide.AdvanceAfterSeconds = ReadAdvanceAfter(transition);
                slide.NotesText = ReadNotes(root);
                ReadTiming(root);
                // p:timing 只解析入场动画的类别、目标形状、延迟与时长，以及主序列的触发方式；
                // 强调/退出动画、媒体与命令节点在本统一模型没有对应表达，按需求跳过。
                // p:timing is parsed only for entrance kinds, their target shape, delay and duration, and the main
                // sequence trigger; emphasis and exit animations plus media and command nodes have no
                // representation in the unified model and are skipped as specified.
                return slide;
            }

            /// <summary>读取 p:transition 的 advTm 自动换片时间（毫秒）并换算为秒。Reads the advTm auto-advance time of p:transition in milliseconds and converts it to seconds.</summary>
            /// <param name="transition">p:transition 元素。The p:transition element.</param>
            /// <returns>停留秒数，无自动换片为 0。Seconds to stay, 0 without auto-advance.</returns>
            private static double ReadAdvanceAfter(XElement transition)
            {
                int milliseconds = Ppt.ParseInt(Ppt.Attribute(transition, "advTm"), 0);
                return milliseconds <= 0 ? 0 : milliseconds / 1000.0;
            }

            /// <summary>解析 p:timing：凡带 presetClass="entr" 的 cTn 视为入场行为，行为子树内的
            /// p:spid/@val（或 p:bldP/@spid）定位目标形状，效果子元素与 filter 判定类别，
            /// cTn/@dur 给出时长，stCondLst 的数值 delay 给出触发后延迟。
            /// Parses p:timing: every cTn with presetClass="entr" is an entrance behavior; the p:spid/@val (or
            /// p:bldP/@spid) inside the behavior resolves the target shape, the effect child and its filter map
            /// the kind, cTn/@dur gives the duration and a numeric stCondLst delay the after-start trigger delay.</summary>
            /// <param name="slideRoot">p:sld 根元素。The p:sld root.</param>
            private void ReadTiming(XElement slideRoot)
            {
                XElement timing = Ppt.Child(slideRoot, "timing");
                if (timing == null || _shapeById.Count == 0) return;
                foreach (XElement ctn in timing.Descendants(Ppt.Main + "cTn"))
                {
                    if (Ppt.Attribute(ctn, "presetClass") != "entr") continue;
                    SlideShape shape = ResolveTargetShape(ctn);
                    if (shape == null || shape.Animation != SlideAnimationKind.None) continue;
                    ApplyAnimation(shape, ctn);
                }
            }

            /// <summary>在入场行为子树中解析目标形状：p:bldP/@spid 优先，其次第一个 p:spid/@val。
            /// Resolves the behavior's target shape: p:bldP/@spid first, otherwise the first p:spid/@val.</summary>
            /// <param name="effect">入场行为 cTn。The entrance behavior cTn.</param>
            /// <returns>映射到的模型形状或 null。The mapped shape or null.</returns>
            private SlideShape ResolveTargetShape(XElement effect)
            {
                int id = -1;
                foreach (XElement descendant in effect.Descendants())
                {
                    if (descendant.Name.Namespace != Ppt.Main) continue;
                    if (descendant.Name.LocalName == "bldP")
                    {
                        id = Ppt.ParseInt(Ppt.Attribute(descendant, "spid"), -1);
                        if (id > 0) break;
                    }
                    if (descendant.Name.LocalName == "spid")
                    {
                        id = Ppt.ParseInt(Ppt.Attribute(descendant, "val"), -1);
                        if (id > 0) break;
                    }
                }
                SlideShape shape;
                return id > 0 && _shapeById.TryGetValue(id, out shape) ? shape : null;
            }

            /// <summary>把入场行为 cTn 写入形状模型：类别由效果子元素与 filter 判定，时长取 cTn/@dur，
            /// 触发延迟自行为节点向上层 par/clk 的 stCondLst 查找首个数值 delay（PowerPoint 把「之后延迟」
            /// 记在外层单击节点上）；indef 表示「单击时」，保持 0。
            /// Applies the entrance behavior to the model: the kind comes from the effect child and its filter,
            /// the duration from cTn/@dur, and the delay is the first numeric stCondLst delay found walking up the
            /// ancestor par/clk nodes (PowerPoint records "delay after start" on the outer click node);
            /// "indef" means an on-click trigger and stays 0.</summary>
            /// <param name="shape">目标形状。The target shape.</param>
            /// <param name="effect">入场行为 cTn。The entrance behavior cTn.</param>
            private static void ApplyAnimation(SlideShape shape, XElement effect)
            {
                XElement childList = Ppt.Child(effect, "childTnLst");
                if (childList == null) return;
                SlideAnimationKind kind = SlideAnimationKind.None;
                foreach (XElement candidate in childList.Elements())
                {
                    kind = MapAnimationKind(candidate);
                    if (kind != SlideAnimationKind.None) break;
                }
                if (kind == SlideAnimationKind.None) return;
                shape.Animation = kind;
                int durationMs = Ppt.ParseInt(Ppt.Attribute(effect, "dur"), 0);
                if (durationMs > 0) shape.AnimationDurationSeconds = durationMs / 1000.0;
                XElement node = effect;
                while (node != null && node.Name != Ppt.Main + "seq")
                {
                    // 触发节点既可能是带 stCondLst 直接子元素的 cTn，也可能是仅包着一个 cTn 的 par/clk。
                    // A trigger node is either a cTn with a direct stCondLst or a par/clk wrapping a single cTn.
                    XElement condition = node;
                    if (condition.Name.LocalName != "cTn")
                    {
                        XElement only = FirstChild(condition, "cTn");
                        if (only == null) { node = node.Parent == null ? null : node.Parent.Parent; continue; }
                        condition = only;
                    }
                    XElement cond = FirstChild(Ppt.Child(condition, "stCondLst"), "cond");
                    if (cond != null)
                    {
                        int delayMs;
                        if (int.TryParse(Ppt.Attribute(cond, "delay"), NumberStyles.Integer, CultureInfo.InvariantCulture, out delayMs) && delayMs > 0)
                        {
                            shape.AnimationDelaySeconds = delayMs / 1000.0;
                            break;
                        }
                    }
                    node = node.Parent == null ? null : node.Parent.Parent;
                }
            }

            /// <summary>把效果子元素映射为动画类别；非效果元素返回 None。Maps an effect child onto the animation kind; non-effect elements return None.</summary>
            /// <param name="candidate">childTnLst 的子元素。The child of childTnLst.</param>
            /// <returns>动画类别。The animation kind.</returns>
            private static SlideAnimationKind MapAnimationKind(XElement candidate)
            {
                if (candidate == null || candidate.Name.Namespace != Ppt.Main) return SlideAnimationKind.None;
                switch (candidate.Name.LocalName)
                {
                    case "bldP": return SlideAnimationKind.Split;
                    case "animMotion": return SlideAnimationKind.FlyIn;
                    case "animScale": return SlideAnimationKind.Zoom;
                    case "animEffect":
                        {
                            string filter = Ppt.Attribute(candidate, "filter");
                            if (filter != null && filter.StartsWith("wipe", StringComparison.OrdinalIgnoreCase)) return SlideAnimationKind.Wipe;
                            return SlideAnimationKind.Fade;
                        }
                    default: return SlideAnimationKind.None;
                }
            }

            /// <summary>登记形状元素的 cNvPr id，供 p:timing 目标解析。Registers the element's cNvPr id for p:timing target resolution.</summary>
            /// <param name="element">p:sp / p:pic / p:graphicFrame 元素。The shape element.</param>
            /// <param name="shape">已创建的模型形状。The model shape just created.</param>
            private void RegisterShapeId(XElement element, SlideShape shape)
            {
                XElement idElement = FirstNestedChild(element, "cNvPr");
                if (idElement == null) return;
                int id = Ppt.ParseInt(Ppt.Attribute(idElement, "id"), -1);
                if (id > 0 && shape != null) _shapeById[id] = shape;
            }

            /// <summary>在元素子树中按本地名称查找第一个后代（nvSpPr 等包装层命名固定但层级不同）。Finds the first descendant by local name, since wrappers such as nvSpPr vary in depth.</summary>
            /// <param name="parent">父元素。The parent element.</param>
            /// <param name="localName">子元素本地名。The child local name.</param>
            /// <returns>匹配的后代或 null。The match or null.</returns>
            private static XElement FirstNestedChild(XElement parent, string localName)
            {
                if (parent == null) return null;
                foreach (XElement descendant in parent.Descendants())
                {
                    if (descendant.Name.Namespace == Ppt.Main && descendant.Name.LocalName == localName) return descendant;
                }
                return null;
            }

            /// <summary>取得指定命名空间的第一个子元素。Returns the first child element in the given namespace.</summary>
            /// <param name="parent">父元素。The parent element.</param>
            /// <param name="localName">子元素本地名。The child local name.</param>
            /// <returns>匹配的子元素或 null。The match or null.</returns>
            private static XElement FirstChild(XElement parent, string localName)
            {
                return Ppt.Child(parent, localName);
            }

            /// <summary>读取形状树中的受支持元素；组合递归展开，深度上限防御畸形文件。Reads supported children of a shape tree; groups expand recursively with a depth guard against malformed files.</summary>
            /// <param name="parent">形状树或组合元素。The shape tree or group element.</param>
            /// <param name="slide">目标幻灯片。The target slide.</param>
            /// <param name="transform">累计坐标变换。The accumulated coordinate transform.</param>
            /// <param name="depth">当前递归深度。The current recursion depth.</param>
            private void ReadShapes(XElement parent, Slide slide, ShapeTransform transform, int depth)
            {
                if (parent == null || depth > 32) return;
                foreach (XElement element in parent.Elements())
                {
                    if (element.Name.Namespace != Ppt.Main) continue;
                    switch (element.Name.LocalName)
                    {
                        case "sp":
                            ReadAutoOrTextBox(slide, element, transform);
                            break;
                        case "pic":
                            ReadPicture(slide, element, transform);
                            break;
                        case "graphicFrame":
                            ReadGraphicFrame(slide, element, transform);
                            break;
                        case "grpSp":
                            ReadGroup(slide, element, transform, depth);
                            break;
                        default:
                            // p:cxnSp 连接线、p:contentPart 与 p:sp 之外的结构在本库统一模型里没有对应表达，整体跳过。
                            // Connector (p:cxnSp), p:contentPart and any other construct have no representation in the
                            // unified model and are skipped as a whole.
                            break;
                    }
                }
            }

            /// <summary>展开组合形状：按 grpSpPr 的 off/ext 与 chOff/chExt 复合坐标变换后递归读取子形状。Expands a group: composes the child transform from grpSpPr off/ext against chOff/chExt, then reads children recursively.</summary>
            /// <param name="slide">目标幻灯片。The target slide.</param>
            /// <param name="group">p:grpSp 元素。The p:grpSp element.</param>
            /// <param name="transform">上层累计变换。The transform inherited from the parent.</param>
            /// <param name="depth">当前递归深度。The current recursion depth.</param>
            private void ReadGroup(Slide slide, XElement group, ShapeTransform transform, int depth)
            {
                XElement grpPr = Ppt.Child(group, "grpSpPr");
                ShapeTransform childTransform = transform;
                XElement xfrm = Ppt.DrawingChild(grpPr, "xfrm");
                if (xfrm != null)
                {
                    XElement off = Ppt.DrawingChild(xfrm, "off");
                    XElement ext = Ppt.DrawingChild(xfrm, "ext");
                    XElement chOff = Ppt.DrawingChild(xfrm, "chOff");
                    XElement chExt = Ppt.DrawingChild(xfrm, "chExt");
                    childTransform = transform.WithGroup(
                        ParseLong(Ppt.Attribute(off, "x"), 0),
                        ParseLong(Ppt.Attribute(off, "y"), 0),
                        ParseLong(Ppt.Attribute(ext, "cx"), 0),
                        ParseLong(Ppt.Attribute(ext, "cy"), 0),
                        ParseLong(Ppt.Attribute(chOff, "x"), 0),
                        ParseLong(Ppt.Attribute(chOff, "y"), 0),
                        ParseLong(Ppt.Attribute(chExt, "cx"), 0),
                        ParseLong(Ppt.Attribute(chExt, "cy"), 0));
                }
                // rot/hflip/vflip 不参与换算：统一模型的形状不含旋转与翻转，缩放与平移已覆盖主要场景。
                // Rotation and flipping are not converted: the unified model has no rotation or flip on shapes,
                // so scaling plus translation covers the practical cases.
                ReadShapes(group, slide, childTransform, depth + 1);
            }

            /// <summary>读取幻灯片背景色：仅 solidFill 有值，bgRef 与渐变返回 null（继承母版）。Reads the slide background colour: only solidFill yields a value, while bgRef and gradients stay null so the master is inherited.</summary>
            /// <param name="cSld">p:cSld 元素。The p:cSld element.</param>
            /// <param name="slide">目标幻灯片。The target slide.</param>
            private void ReadBackground(XElement cSld, Slide slide)
            {
                XElement background = Ppt.Child(Ppt.Child(cSld, "bg"), "bgPr");
                slide.BackgroundColor = _session.ReadFill(background != null ? background : Ppt.Child(cSld, "bg"));
            }

            /// <summary>
            /// 读取 p:sp：带 p:ph 或 prstGeom 为 rect 的产出 TextBoxShape，其余产出 AutoShape（继承 TextBoxShape，
            /// 未知 prst 降级为 Rectangle 但保留文本）。
            /// Reads a p:sp: shapes carrying p:ph or a rect prstGeom become TextBoxShape, everything else becomes
            /// AutoShape (which derives from TextBoxShape, so polymorphism stays safe); unknown presets degrade to
            /// Rectangle while keeping their text.
            /// </summary>
            /// <param name="slide">目标幻灯片。The target slide.</param>
            /// <param name="sp">p:sp 元素。The p:sp element.</param>
            /// <param name="transform">累计坐标变换。The accumulated coordinate transform.</param>
            private void ReadAutoOrTextBox(Slide slide, XElement sp, ShapeTransform transform)
            {
                XElement spPr = Ppt.Child(sp, "spPr");
                XElement ph = FindPlaceholder(sp);
                bool isPlaceholder = ph != null;
                string preset = Ppt.Attribute(Ppt.DrawingChild(spPr, "prstGeom"), "prst");
                bool hasCustomGeometry = Ppt.DrawingChild(spPr, "custGeom") != null;
                bool isAutoShape = !isPlaceholder && (hasCustomGeometry || (!string.IsNullOrEmpty(preset) && preset != "rect"));

                TextBoxShape box = isAutoShape ? new AutoShape() : new TextBoxShape();
                box.Name = ReadShapeName(sp, "nvSpPr");
                box.Placeholder = MapPlaceholderKind(ph);
                box.PlaceholderIndex = ph == null ? -1 : Ppt.IntAttribute(ph, "idx", -1);
                ApplyGeometry(box, spPr, PlaceholderKey(ph, "title"), transform);
                box.FillColor = _session.ReadFill(spPr);
                box.Outline = ReadOutline(spPr);
                XElement txBody = Ppt.Child(sp, "txBody");
                if (txBody != null)
                {
                    ReadBodyProperties(txBody, box);
                    foreach (Paragraph paragraph in ReadParagraphs(txBody)) box.Paragraphs.Add(paragraph);
                }
                if (isAutoShape)
                {
                    AutoShape auto = (AutoShape)box;
                    auto.Kind = MapAutoShapeKind(preset);
                    auto.CornerRadius = ReadRoundRectRadius(spPr);
                }
                RegisterShapeId(sp, box);
                AddShape(slide, box);
            }

            /// <summary>读取 p:pic：blip 的 r:embed 关系指向图片部件；外部关系或图片缺件时整形跳过不中断。Reads a p:pic: the blip r:embed relationship points at the image part; pictures behind external relationships or missing parts are skipped without aborting the read.</summary>
            /// <param name="slide">目标幻灯片。The target slide.</param>
            /// <param name="pic">p:pic 元素。The p:pic element.</param>
            /// <param name="transform">累计坐标变换。The accumulated coordinate transform.</param>
            private void ReadPicture(Slide slide, XElement pic, ShapeTransform transform)
            {
                XElement blipFill = Ppt.Child(pic, "blipFill");
                XElement blip = Ppt.DrawingChild(blipFill, "blip");
                string embed = Ppt.Attribute(blip, Ppt.R + "embed");
                string imagePart = _session.Package.GetRelationshipTarget(_slidePart, embed);
                if (string.IsNullOrEmpty(imagePart))
                {
                    // External 关系或关系缺失：无字节可还原则丢弃该形状，写出器不会输出空图片。
                    // External relationship or missing target: without bytes the shape is dropped, since the
                    // writer cannot emit an empty picture.
                    return;
                }
                byte[] bytes = _session.Package.GetBytes(imagePart);
                if (bytes == null || bytes.Length == 0) return;

                PictureShape picture = new PictureShape();
                XElement spPr = Ppt.Child(pic, "spPr");
                picture.Name = ReadShapeName(pic, "nvPicPr");
                XElement ph = FindPlaceholder(pic);
                picture.Placeholder = MapPlaceholderKind(ph, PlaceholderKind.Picture);
                picture.PlaceholderIndex = ph == null ? -1 : Ppt.IntAttribute(ph, "idx", -1);
                picture.Image = new OfficeImage(bytes);
                picture.AlternativeText = Ppt.Attribute(Ppt.Child(Ppt.Child(pic, "nvPicPr"), "cNvPr"), "descr");
                ApplyGeometry(picture, spPr, PlaceholderKey(ph, "pic"), transform);
                RegisterShapeId(pic, picture);
                AddShape(slide, picture);
            }

            /// <summary>
            /// 读取 p:graphicFrame：仅 a:graphicData 的 uri 以 /table 结尾时产出 TableShape；
            /// 图表、SmartArt 与 OLE 对象整体跳过，因为统一模型只有表格这一种嵌入对象表达，
            /// 半解析反而会产出语义错误的内容。
            /// Reads a p:graphicFrame: a TableShape is produced only when the a:graphicData uri ends with /table.
            /// Charts, SmartArt and OLE objects are skipped entirely because the unified model expresses no other
            /// embedded object and a partial read would yield semantically wrong content.
            /// </summary>
            /// <param name="slide">目标幻灯片。The target slide.</param>
            /// <param name="graphicFrame">p:graphicFrame 元素。The p:graphicFrame element.</param>
            /// <param name="transform">累计坐标变换。The accumulated coordinate transform.</param>
            private void ReadGraphicFrame(Slide slide, XElement graphicFrame, ShapeTransform transform)
            {
                XElement graphic = Ppt.DrawingChild(graphicFrame, "graphic");
                XElement graphicData = Ppt.DrawingChild(graphic, "graphicData");
                string uri = Ppt.Attribute(graphicData, "uri");
                if (uri == null || !uri.EndsWith("/table", StringComparison.OrdinalIgnoreCase)) return;

                TableShape table = new TableShape();
                XElement xfrm = Ppt.Child(graphicFrame, "xfrm");
                table.Name = ReadShapeName(graphicFrame, "nvGraphicFramePr");
                XElement ph = FindPlaceholder(graphicFrame);
                table.Placeholder = MapPlaceholderKind(ph, PlaceholderKind.Table);
                table.PlaceholderIndex = ph == null ? -1 : Ppt.IntAttribute(ph, "idx", -1);
                ApplyFrameGeometry(table, xfrm, PlaceholderKey(ph, "tbl"), transform);

                XElement tbl = Ppt.DrawingChild(graphicData, "tbl");
                ReadTableColumns(tbl, table);
                foreach (XElement rowElement in Ppt.DrawingChildren(tbl, "tr"))
                {
                    table.Table.Rows.Add(ReadTableRow(rowElement));
                }
                RegisterShapeId(graphicFrame, table);
                AddShape(slide, table);
            }

            /// <summary>读取表格一行：a:tr@h 转磅，hMerge/vMerge 续接单元格被丢弃，因为统一模型只保留合并起点加 ColSpan/RowSpan。Reads one table row: a:tr@h becomes points, while hMerge/vMerge continuation cells are dropped because the unified model keeps only the merge origin plus ColSpan/RowSpan.</summary>
            /// <param name="rowElement">a:tr 元素。The a:tr element.</param>
            /// <returns>读取出的表格行。The parsed row.</returns>
            private TableRow ReadTableRow(XElement rowElement)
            {
                TableRow row = new TableRow();
                row.Height = Ppt.ToPoints(ParseLong(Ppt.Attribute(rowElement, "h"), 0));
                foreach (XElement cellElement in Ppt.DrawingChildren(rowElement, "tc"))
                {
                    if (Ppt.ParseBool(Ppt.Attribute(cellElement, "hMerge"), false)) continue;
                    if (Ppt.ParseBool(Ppt.Attribute(cellElement, "vMerge"), false)) continue;
                    row.Cells.Add(ReadTableCell(cellElement));
                }
                return row;
            }

            /// <summary>读取单个表格单元格：段落进 Blocks，tcPr 提供底色、合并跨度与垂直对齐。Reads one cell: paragraphs go into Blocks while tcPr supplies shading, merge spans and vertical alignment.</summary>
            /// <param name="cellElement">a:tc 元素。The a:tc element.</param>
            /// <returns>读取出的单元格。The parsed cell.</returns>
            private TableCell ReadTableCell(XElement cellElement)
            {
                TableCell cell = new TableCell();
                cell.ColSpan = Math.Min(MaxCellSpan, Math.Max(1, Ppt.IntAttribute(cellElement, "gridSpan", 1)));
                cell.RowSpan = Math.Min(MaxCellSpan, Math.Max(1, Ppt.IntAttribute(cellElement, "rowSpan", 1)));
                XElement txBody = Ppt.DrawingChild(cellElement, "txBody");
                if (txBody != null)
                {
                    foreach (Paragraph paragraph in ReadParagraphs(txBody)) cell.Blocks.Add(paragraph);
                }
                XElement tcPr = Ppt.DrawingChild(cellElement, "tcPr");
                cell.BackgroundColor = _session.ReadFill(tcPr);
                cell.VerticalAlignment = MapCellAnchor(Ppt.Attribute(tcPr, "anchor"));
                // tcPr 的 a:lnL/a:lnR/a:lnT/a:lnB 与 marT/marB/marR 不读入：统一模型的单元格边框与
                // 内边距是表级设置，逐格还原会让写出器重复输出四周边框。首个单元格的 marL 已在
                // ReadTableColumns 里作为表级内边距的近似值使用。
                // The per-edge borders (a:lnL/lnR/lnT/lnB) and the remaining marT/marB/marR insets are not
                // read: the unified model keeps borders and padding at table level, so restoring them cell by
                // cell would make the writer repeat the surrounding borders. The first cell's marL is already
                // used in ReadTableColumns as an approximation of the table padding.
                return cell;
            }

            /// <summary>
            /// 把 a:tblGrid 的列宽按形状宽度分摊为磅写入 Table.Columns；
            /// 统一模型只有表级 CellPadding，故用首个单元格的 marL 近似还原内边距。
            /// Converts a:tblGrid column widths into point widths spread over the shape width and stores them in
            /// Table.Columns; the unified model only has table-level CellPadding, so the first cell's marL
            /// approximates the cell padding.
            /// </summary>
            /// <param name="tbl">a:tbl 元素。The a:tbl element.</param>
            /// <param name="table">目标表格形状。The target table shape.</param>
            private void ReadTableColumns(XElement tbl, TableShape table)
            {
                XElement grid = Ppt.DrawingChild(tbl, "tblGrid");
                List<long> widths = new List<long>();
                long total = 0;
                foreach (XElement column in Ppt.DrawingChildren(grid, "gridCol"))
                {
                    long width = ParseLong(Ppt.Attribute(column, "w"), 0);
                    if (width < 0) width = 0;
                    widths.Add(width);
                    total += width;
                }
                if (widths.Count == 0) return;
                double available = table.Width > 0 ? table.Width : Ppt.ToPoints(total);
                for (int i = 0; i < widths.Count; i++)
                {
                    double width = total > 0 ? available * widths[i] / total : available / widths.Count;
                    table.Table.Columns.Add(new TableColumn(width));
                }

                XElement firstRow = Ppt.DrawingChild(tbl, "tr");
                XElement firstCell = Ppt.DrawingChild(firstRow, "tc");
                long margin = ParseLong(Ppt.Attribute(Ppt.DrawingChild(firstCell, "tcPr"), "marL"), 0);
                if (margin > 0) table.Table.CellPadding = Ppt.ToPoints(margin);
            }

            /// <summary>读取 a:bodyPr 的垂直对齐、内边距与自动缩放。Reads the vertical anchor, insets and autofit from a:bodyPr.</summary>
            /// <param name="txBody">p:txBody 元素。The p:txBody element.</param>
            /// <param name="box">目标文本框。The target text box.</param>
            private void ReadBodyProperties(XElement txBody, TextBoxShape box)
            {
                // bodyPr 在图形命名空间（a:bodyPr），按主命名空间查找会永远取空。
                // bodyPr lives in the drawing namespace (a:bodyPr); a main-namespace lookup would always miss it.
                XElement bodyPr = Ppt.DrawingChild(txBody, "bodyPr");
                box.VerticalAlignment = MapVerticalAnchor(Ppt.Attribute(bodyPr, "anchor"));
                long left = ParseLong(Ppt.Attribute(bodyPr, "lIns"), -1);
                long right = ParseLong(Ppt.Attribute(bodyPr, "rIns"), -1);
                long top = ParseLong(Ppt.Attribute(bodyPr, "tIns"), -1);
                long bottom = ParseLong(Ppt.Attribute(bodyPr, "bIns"), -1);
                if (left >= 0) box.MarginLeft = Ppt.ToPoints(left);
                if (right >= 0) box.MarginRight = Ppt.ToPoints(right);
                if (top >= 0) box.MarginTop = Ppt.ToPoints(top);
                if (bottom >= 0) box.MarginBottom = Ppt.ToPoints(bottom);
                if (Ppt.DrawingChild(bodyPr, "normAutofit") != null) box.ShrinkToFit = true;
            }

            /// <summary>读取文本部件的全部段落。Reads every paragraph of a text body element.</summary>
            /// <param name="txBody">p:txBody 或 a:txBody 元素，可为 null。The p:txBody or a:txBody element; may be null.</param>
            /// <returns>段落列表；部件不存在时为空列表。The paragraphs, or an empty list when the element is absent.</returns>
            private List<Paragraph> ReadParagraphs(XElement txBody)
            {
                List<Paragraph> paragraphs = new List<Paragraph>();
                if (txBody == null) return paragraphs;
                // 段落默认对齐与缩进本可由 a:lstStyle 决定，本库不解析样式继承链，故只读显式的 a:pPr。
                // Paragraph defaults could come from a:lstStyle; the style inheritance chain is not resolved
                // here, so only the explicit a:pPr of each paragraph is read.
                foreach (XElement paragraphElement in Ppt.DrawingChildren(txBody, "p"))
                {
                    paragraphs.Add(ReadParagraph(paragraphElement));
                }
                return paragraphs;
            }

            /// <summary>读取单个 a:p 段落：pPr 提供对齐、缩进与列表标记，子元素提供文本运行与换行。Reads one a:p paragraph: pPr gives alignment, indentation and the list marker while children supply runs and line breaks.</summary>
            /// <param name="paragraphElement">a:p 元素。The a:p element.</param>
            /// <returns>读取出的段落。The parsed paragraph.</returns>
            private Paragraph ReadParagraph(XElement paragraphElement)
            {
                Paragraph paragraph = new Paragraph();
                XElement pPr = Ppt.DrawingChild(paragraphElement, "pPr");
                if (pPr != null)
                {
                    paragraph.Alignment = MapAlignment(Ppt.Attribute(pPr, "algn"));
                    paragraph.LeftIndent = Ppt.ToPoints(ParseLong(Ppt.Attribute(pPr, "marL"), 0));
                    paragraph.FirstLineIndent = Ppt.ToPoints(ParseLong(Ppt.Attribute(pPr, "indent"), 0));
                    paragraph.ListMarker = ReadListMarker(pPr);
                }
                foreach (XElement child in paragraphElement.Elements())
                {
                    if (child.Name.Namespace != Ppt.Drawing) continue;
                    switch (child.Name.LocalName)
                    {
                        case "r":
                            paragraph.AddRun(ReadRun(child));
                            break;
                        case "br":
                            paragraph.AddLineBreak();
                            break;
                        case "fld":
                            // 字段（页码/日期）按其缓存结果还原为普通文本运行，读取端不保留字段指令。
                            // Fields (slide number, date) are restored as plain text from their cached result;
                            // the field instruction itself is not carried into the model.
                            paragraph.AddText(child.Value, _session.ReadRunFormat(Ppt.DrawingChild(child, "rPr")));
                            break;
                        default:
                            break;
                    }
                }
                return paragraph;
            }

            /// <summary>读取单个 a:r 文本运行，保留原始文本（含 xml:space 保留的空格与空文本）。Reads one a:r text run, preserving the raw text including retained spaces and empty runs.</summary>
            /// <param name="runElement">a:r 元素。The a:r element.</param>
            /// <returns>读取出的文本运行。The parsed text run.</returns>
            private TextRun ReadRun(XElement runElement)
            {
                XElement textElement = Ppt.DrawingChild(runElement, "t");
                string text = textElement == null ? string.Empty : textElement.Value;
                return new TextRun(text, _session.ReadRunFormat(Ppt.DrawingChild(runElement, "rPr")));
            }

            /// <summary>解析 a:pPr 的列表标记：buNone 表示无标记，buChar 取字符，buAutoNum 输出该编号格式的示例标记。Resolves the list marker from a:pPr: buNone means no marker, buChar yields its character and buAutoNum yields a representative marker of that numbering format.</summary>
            /// <param name="pPr">a:pPr 元素。The a:pPr element.</param>
            /// <returns>标记文本，或 null 表示非列表段落。The marker text, or null for a non-list paragraph.</returns>
            private static string ReadListMarker(XElement pPr)
            {
                if (Ppt.DrawingChild(pPr, "buNone") != null) return null;
                XElement buChar = Ppt.DrawingChild(pPr, "buChar");
                if (buChar != null) return Ppt.Attribute(buChar, "char");
                XElement buAutoNum = Ppt.DrawingChild(pPr, "buAutoNum");
                if (buAutoNum != null) return MapAutoNumberMarker(Ppt.Attribute(buAutoNum, "type"));
                return null;
            }

            /// <summary>
            /// 把 buAutoNum 类型映射为代表性标记文本。统一模型没有计数器，故只还原编号样式而不重算实际序号。
            /// Maps a buAutoNum type to a representative marker string. The unified model has no counters, so the
            /// numbering style is restored while the real ordinal is not recomputed.
            /// </summary>
            /// <param name="type">ST_TextAutoNumbering 类型值。The ST_TextAutoNumbering value.</param>
            /// <returns>标记文本，例如 "1."。The marker text, such as "1.".</returns>
            private static string MapAutoNumberMarker(string type)
            {
                if (string.IsNullOrEmpty(type) || type == "noAutoNum" || type == "none") return null;
                if (type == "arabicPeriod" || type == "arabicPlain") return "1.";
                if (type == "arabicZeroPeriod" || type == "arabicZeroPlain") return "01.";
                if (type == "alphaLcPeriod" || type == "lAlphaPeriod" || type == "alphaLcParenBoth" || type == "lAlphaParenBoth") return "a.";
                if (type == "alphaUpcPeriod" || type == "uAlphaPeriod") return "A.";
                if (type == "romanLcPeriod" || type == "lRomanPeriod") return "i.";
                if (type == "romanUpcPeriod" || type == "uRomanPeriod") return "I.";
                if (type.EndsWith("Both", StringComparison.OrdinalIgnoreCase)) return "(1)";
                if (type.EndsWith("Dash", StringComparison.OrdinalIgnoreCase)) return "1-";
                if (type.EndsWith("ParenR", StringComparison.OrdinalIgnoreCase)) return "1)";
                return "1.";
            }

            /// <summary>读取 a:ln 描边：w 转磅、solidFill 取色、prstDash 映射线型；noFill 或缺失时返回 null。Reads the a:ln outline: w becomes points, solidFill provides the colour and prstDash maps the style; noFill or a missing element yields null.</summary>
            /// <param name="spPr">p:spPr 元素。The p:spPr element.</param>
            /// <returns>边框线描述或 null。The border line, or null when there is no outline.</returns>
            private BorderLine ReadOutline(XElement spPr)
            {
                XElement ln = Ppt.DrawingChild(spPr, "ln");
                if (ln == null || Ppt.DrawingChild(ln, "noFill") != null) return null;
                // a:ln 缺省线宽为 1 磅。The implicit a:ln width is one point.
                double width = Ppt.ToPoints(ParseLong(Ppt.Attribute(ln, "w"), 12700));
                OfficeColor? color = _session.ReadFill(ln);
                return new BorderLine(MapDashStyle(Ppt.Attribute(Ppt.DrawingChild(ln, "prstDash"), "val")), width, color ?? OfficeColor.Black);
            }

            /// <summary>把预设线型映射到统一模型线型（预置名不含下划线）。Maps a preset dash name onto the unified border style (preset names contain no underscores).</summary>
            /// <param name="value">prstDash 的 val 属性。The val attribute of prstDash.</param>
            /// <returns>统一模型线型。The unified border style.</returns>
            private static BorderStyle MapDashStyle(string value)
            {
                if (string.IsNullOrEmpty(value) || value == "solid") return BorderStyle.Single;
                if (value == "sysDot" || value == "lnDot" || value == "systDot" || value == "dot" || value == "roundDot") return BorderStyle.Dotted;
                if (value == "sysDash" || value == "dash" || value == "dashDot" || value == "lgDash" || value == "lgDashDot" || value == "lgDashDotDot") return BorderStyle.Dashed;
                if (value == "longDashDash") return BorderStyle.Dashed;
                return BorderStyle.Single;
            }

            /// <summary>
            /// 定位幻灯片上的 p:transition：优先直挂根元素的写法，其次 mc:AlternateContent 的 mc:Choice，最后 mc:Fallback。
            /// Locates the slide's p:transition: a direct child of the root first, then the mc:Choice of an
            /// mc:AlternateContent block, and finally its mc:Fallback.
            /// </summary>
            /// <param name="slideRoot">p:sld 根元素。The p:sld root element.</param>
            /// <returns>transition 元素，无切换时 null。The transition element, or null when there is none.</returns>
            private static XElement FindTransition(XElement slideRoot)
            {
                if (slideRoot == null) return null;
                XElement direct = Ppt.Child(slideRoot, "transition");
                if (direct != null) return direct;
                XElement alternate = Ppt.MarkupCompatChild(slideRoot, "AlternateContent");
                if (alternate == null) return null;
                XElement choice = Ppt.MarkupCompatChild(alternate, "Choice");
                XElement transition = Ppt.Child(choice, "transition");
                if (transition != null) return transition;
                return Ppt.Child(Ppt.MarkupCompatChild(alternate, "Fallback"), "transition");
            }

            /// <summary>
            /// 读取切换时长（秒）：ISO/IEC 29500 的 CT_SlideTransition 不带时长，写在 p14:dur 上；
            /// 部分第三方工具直接写无命名空间的 dur，两者都按毫秒解释。
            /// Reads the transition duration in seconds: CT_SlideTransition carries no duration in ISO/IEC 29500,
            /// so it lives in p14:dur; some third-party writers use a plain dur attribute. Both are milliseconds.
            /// </summary>
            /// <param name="transition">p:transition 元素。The p:transition element.</param>
            /// <returns>时长秒数，缺失或非法为 0。The duration in seconds, 0 when missing or invalid.</returns>
            private static double ReadTransitionDuration(XElement transition)
            {
                if (transition == null) return 0;
                int milliseconds = Ppt.ParseInt(Ppt.Attribute(transition, Ppt.PowerPoint2010 + "dur"), 0);
                if (milliseconds <= 0)
                {
                    milliseconds = Ppt.ParseInt(Ppt.Attribute(transition, "dur"), 0);
                }
                return milliseconds <= 0 ? 0 : milliseconds / 1000.0;
            }

            /// <summary>读取 p:transition 的切换类别；p:audioCd 与自动换片时间 advTm 按需求忽略。Reads the transition kind from p:transition; p:audioCd and the auto-advance time are ignored as specified.</summary>
            /// <param name="transition">p:transition 元素。The p:transition element.</param>
            /// <returns>切换效果类别。The transition kind.</returns>
            private static SlideTransitionKind ReadTransition(XElement transition)
            {
                if (transition == null) return SlideTransitionKind.None;
                if (Ppt.Child(transition, "fade") != null) return SlideTransitionKind.Fade;
                if (Ppt.Child(transition, "push") != null) return SlideTransitionKind.Push;
                if (Ppt.Child(transition, "wipe") != null) return SlideTransitionKind.Wipe;
                if (Ppt.Child(transition, "split") != null) return SlideTransitionKind.Split;
                return SlideTransitionKind.None;
            }

            /// <summary>
            /// 读取备注：slide rels 里 /notesSlide 部件中 ph type 为 body（含默认 body）的文本部件，多段落以换行连接。
            /// Reads speaker notes: the text body of the notesSlide part referenced from the slide relationships whose
            /// placeholder type is body (the implicit default included), joining paragraphs with line feeds.
            /// </summary>
            /// <param name="slideRoot">p:sld 根元素。The p:sld root element.</param>
            /// <returns>备注纯文本，无备注时为空字符串。The notes text, empty when there are none.</returns>
            private string ReadNotes(XElement slideRoot)
            {
                string notesPart = _session.Package.GetRelationshipTargetByType(_slidePart, NotesSlideSuffix);
                if (string.IsNullOrEmpty(notesPart)) return string.Empty;
                XDocument notesXml = _session.Package.GetXml(notesPart);
                if (notesXml == null || notesXml.Root == null) return string.Empty;
                XElement spTree = Ppt.Child(Ppt.Child(notesXml.Root, "cSld"), "spTree");
                StringBuilder builder = new StringBuilder();
                if (spTree != null)
                {
                    foreach (XElement sp in Ppt.Children(spTree, "sp"))
                    {
                        XElement ph = FindPlaceholder(sp);
                        if (!IsNotesBody(ph)) continue;
                        foreach (Paragraph paragraph in ReadParagraphs(Ppt.Child(sp, "txBody")))
                        {
                            if (builder.Length > 0) builder.Append('\n');
                            builder.Append(paragraph.GetText());
                        }
                    }
                }
                return builder.ToString();
            }

            /// <summary>判断备注页占位符是否为正文占位符（type 省略时即 body）。Tests whether a notes placeholder is the body placeholder (an omitted type means body).</summary>
            /// <param name="ph">p:ph 元素。The p:ph element.</param>
            /// <returns>是正文占位符时返回 true。True for the body placeholder.</returns>
            private static bool IsNotesBody(XElement ph)
            {
                if (ph == null) return false;
                string type = Ppt.Attribute(ph, "type");
                return string.IsNullOrEmpty(type) || type == "body" || type == "obj";
            }

            /// <summary>按幻灯片尺寸比例推算占位符默认外接框（版式与母版都没有几何时的最后回落）。Derives the default placeholder box proportionally to the slide size, used only when neither layout nor master has geometry.</summary>
            /// <param name="phKey">占位符键（可为 null）。The placeholder key, possibly null.</param>
            /// <returns>EMU 矩形。The EMU rectangle.</returns>
            private Box DefaultBox(string phKey)
            {
                double width = _session.Presentation.SlideWidth;
                double height = _session.Presentation.SlideHeight;
                string type = phKey == null ? string.Empty : phKey;
                int bar = type.IndexOf('|');
                if (bar >= 0) type = type.Substring(0, bar);
                bool isTitle = type == "title" || type == "ctrTitle";
                double top = isTitle ? height * 0.06 : height * 0.30;
                double shapeHeight = isTitle ? height * 0.20 : height * 0.58;
                return new Box(Ppt.Emu(width * 0.08), Ppt.Emu(top), Ppt.Emu(width * 0.84), Ppt.Emu(shapeHeight));
            }

            /// <summary>把形状几何写入模型（磅）。own 几何参与组合坐标换算，继承/默认几何已是幻灯片空间故不再换算。Writes the shape geometry into the model in points; own geometry goes through the group transform, while inherited or default geometry is already in slide space.</summary>
            /// <param name="target">目标形状。The target shape.</param>
            /// <param name="spPr">p:spPr 元素。The p:spPr element.</param>
            /// <param name="phKey">占位符键（无 ph 时为 null）。The placeholder key, null without p:ph.</param>
            /// <param name="transform">累计坐标变换。The accumulated coordinate transform.</param>
            private void ApplyGeometry(SlideShape target, XElement spPr, string phKey, ShapeTransform transform)
            {
                Box own = ReadOwnBox(spPr);
                Box box = own ?? ResolveInheritedOrDefaultBox(phKey);
                ApplyBox(target, box, own != null ? transform : new ShapeTransform());
            }

            /// <summary>把 graphicFrame 的 p:xfrm（或继承几何）写入形状。Writes a graphicFrame's p:xfrm, or inherited geometry, onto the shape.</summary>
            /// <param name="target">目标形状。The target shape.</param>
            /// <param name="xfrm">graphicFrame 的直接子元素 p:xfrm。The graphicFrame child p:xfrm element.</param>
            /// <param name="phKey">占位符键（无 ph 时为 null）。The placeholder key, null without p:ph.</param>
            /// <param name="transform">累计坐标变换。The accumulated coordinate transform.</param>
            private void ApplyFrameGeometry(SlideShape target, XElement xfrm, string phKey, ShapeTransform transform)
            {
                Box own = ReadBoxFromTransform(xfrm);
                ApplyBox(target, own ?? ResolveInheritedOrDefaultBox(phKey), xfrm != null ? transform : new ShapeTransform());
            }

            /// <summary>EMU 矩形换算为磅并落到形状上。Converts an EMU rectangle to points and assigns it to the shape.</summary>
            /// <param name="target">目标形状。The target shape.</param>
            /// <param name="box">EMU 矩形。The EMU rectangle.</param>
            /// <param name="transform">累计坐标变换。The accumulated coordinate transform.</param>
            private static void ApplyBox(SlideShape target, Box box, ShapeTransform transform)
            {
                if (box == null || target == null) return;
                if (!transform.IsIdentity)
                {
                    target.Left = Ppt.ToPoints(Math.Round(box.X * transform.ScaleX + transform.OffsetX));
                    target.Top = Ppt.ToPoints(Math.Round(box.Y * transform.ScaleY + transform.OffsetY));
                    target.Width = Ppt.ToPoints(Math.Round(box.CX * transform.ScaleX));
                    target.Height = Ppt.ToPoints(Math.Round(box.CY * transform.ScaleY));
                    return;
                }
                target.Left = Ppt.ToPoints(box.X);
                target.Top = Ppt.ToPoints(box.Y);
                target.Width = Ppt.ToPoints(box.CX);
                target.Height = Ppt.ToPoints(box.CY);
            }

            /// <summary>读取形状自身的 a:xfrm 外接框，缺失时返回 null。Reads the shape's own a:xfrm box, returning null when absent.</summary>
            /// <param name="spPr">p:spPr 或 a:spPr 元素。The p:spPr or a:spPr element.</param>
            /// <returns>EMU 矩形或 null。The EMU box or null.</returns>
            private static Box ReadOwnBox(XElement spPr)
            {
                return ReadBoxFromTransform(Ppt.DrawingChild(spPr, "xfrm"));
            }

            /// <summary>从 xfrm 元素的 off/ext 读取矩形；off/ext 不完整时返回 null。Reads a rectangle from an xfrm element's off/ext children, returning null when they are incomplete.</summary>
            /// <param name="xfrm">xfrm 元素（命名空间由调用方选定）。The xfrm element (namespace chosen by the caller).</param>
            /// <returns>EMU 矩形或 null。The EMU box or null.</returns>
            private static Box ReadBoxFromTransform(XElement xfrm)
            {
                if (xfrm == null) return null;
                XElement off = Ppt.DrawingChild(xfrm, "off");
                XElement ext = Ppt.DrawingChild(xfrm, "ext");
                if (off == null || ext == null) return null;
                long cx = ParseLong(Ppt.Attribute(ext, "cx"), 0);
                long cy = ParseLong(Ppt.Attribute(ext, "cy"), 0);
                if (cx <= 0 || cy <= 0) return null;
                return new Box(ParseLong(Ppt.Attribute(off, "x"), 0), ParseLong(Ppt.Attribute(off, "y"), 0), cx, cy);
            }

            /// <summary>
            /// 占位符几何回落链的第二、三级：先查版式同名占位符（键含 idx → 仅 type → title/ctrTitle 互换），
            /// 再查母版同名占位符，全部落空时用按幻灯片尺寸推算的默认矩形。
            /// Levels two and three of the geometry fallback chain: the layout placeholder is looked up first
            /// (key with idx, then type only, then the title/ctrTitle alias), then the master placeholder, and
            /// when both miss, the proportional default rectangle is used.
            /// </summary>
            /// <param name="phKey">占位符键；null 表示非占位符形状。The placeholder key; null for a non-placeholder shape.</param>
            /// <returns>EMU 矩形。The EMU box.</returns>
            private Box ResolveInheritedOrDefaultBox(string phKey)
            {
                if (!string.IsNullOrEmpty(phKey))
                {
                    Box fromLayout = LookupBox(_layoutBoxes, phKey);
                    if (fromLayout != null) return fromLayout;
                    Box fromMaster = LookupBox(_masterBoxes, phKey);
                    if (fromMaster != null) return fromMaster;
                }
                return DefaultBox(phKey);
            }

            /// <summary>按候选键顺序查表：完整键 → 仅 type → title/ctrTitle 别名互换。Tries candidate keys in order: the full key, the type only, then the title/ctrTitle alias swap.</summary>
            /// <param name="boxes">几何表。The geometry table.</param>
            /// <param name="phKey">占位符键。The placeholder key.</param>
            /// <returns>匹配矩形或 null。The matching box or null.</returns>
            private static Box LookupBox(Dictionary<string, Box> boxes, string phKey)
            {
                if (boxes == null) return null;
                string typeOnly = phKey;
                int bar = typeOnly.IndexOf('|');
                if (bar >= 0) typeOnly = typeOnly.Substring(0, bar);
                Box box;
                if (boxes.TryGetValue(phKey, out box)) return box;
                if (boxes.TryGetValue(typeOnly, out box)) return box;
                string alias = typeOnly == "title" ? "ctrTitle" : typeOnly == "ctrTitle" ? "title" : null;
                if (alias != null && boxes.TryGetValue(alias, out box)) return box;
                return null;
            }

            /// <summary>收集版式/母版部件里带几何的占位符矩形（同键首次出现优先）。Collects placeholder rectangles that carry geometry from a layout or master part, keeping the first occurrence per key.</summary>
            /// <param name="document">版式或母版 XML 文档。The layout or master document.</param>
            /// <returns>占位符键到矩形的表。A table from placeholder key to box.</returns>
            private static Dictionary<string, Box> CollectPlaceholderBoxes(XDocument document)
            {
                Dictionary<string, Box> boxes = new Dictionary<string, Box>(StringComparer.OrdinalIgnoreCase);
                if (document == null || document.Root == null) return boxes;
                XElement spTree = Ppt.Child(Ppt.Child(document.Root, "cSld"), "spTree");
                if (spTree == null) return boxes;
                foreach (XElement element in spTree.Elements())
                {
                    if (element.Name.Namespace != Ppt.Main) continue;
                    if (element.Name.LocalName != "sp" && element.Name.LocalName != "pic") continue;
                    XElement ph = FindPlaceholder(element);
                    // 版式/母版里 type 与 idx 同时省略的形状按其元素名补默认类型，与幻灯片侧解析规则一致。
                    // When both type and idx are omitted in a layout or master, the element name supplies the
                    // default type, matching how the slide side builds the key.
                    string key = PlaceholderKey(ph, element.Name.LocalName == "pic" ? "pic" : "title");
                    if (key == null) continue;
                    Box box = ReadOwnBox(Ppt.Child(element, "spPr"));
                    if (box == null) continue;
                    string typeOnly = key;
                    int bar = typeOnly.IndexOf('|');
                    if (bar >= 0) typeOnly = typeOnly.Substring(0, bar);
                    if (!boxes.ContainsKey(key)) boxes[key] = box;
                    if (!boxes.ContainsKey(typeOnly)) boxes[typeOnly] = box;
                }
                return boxes;
            }

            /// <summary>定位形状的 p:nvPr/p:ph 占位符声明。Finds the p:nvPr/p:ph placeholder declaration of a shape.</summary>
            /// <param name="shape">形状元素。The shape element.</param>
            /// <returns>p:ph 元素或 null。The p:ph element or null.</returns>
            private static XElement FindPlaceholder(XElement shape)
            {
                foreach (XElement child in shape.Elements())
                {
                    if (child.Name.Namespace != Ppt.Main) continue;
                    if (!child.Name.LocalName.EndsWith("Pr", StringComparison.Ordinal)) continue;
                    if (!child.Name.LocalName.StartsWith("nv", StringComparison.Ordinal)) continue;
                    XElement ph = Ppt.Child(Ppt.Child(child, "nvPr"), "ph");
                    if (ph != null) return ph;
                }
                return null;
            }

            /// <summary>
            /// 生成占位符几何匹配键：省略的 type 与 <see cref="MapPlaceholderKind(XElement, PlaceholderKind)"/>
            /// 同规则解析（无 idx 时用调用方给出的默认类型，有 idx 时按 body），obj 归一为 body；
            /// 有 idx 时追加 "|idx"。
            /// Builds the placeholder geometry key: an omitted type is resolved exactly as
            /// <see cref="MapPlaceholderKind(XElement, PlaceholderKind)"/> does (the caller's default when no idx
            /// is present, otherwise body), obj normalises to body, and an idx is appended as "|idx".
            /// </summary>
            /// <param name="ph">p:ph 元素。The p:ph element.</param>
            /// <param name="missingType">type 与 idx 都省略时的类型。The type used when both type and idx are absent.</param>
            /// <returns>匹配键；非占位符返回 null。The key, or null for a non-placeholder shape.</returns>
            private static string PlaceholderKey(XElement ph, string missingType)
            {
                if (ph == null) return null;
                string idx = Ppt.Attribute(ph, "idx");
                string type = Ppt.Attribute(ph, "type");
                if (string.IsNullOrEmpty(type))
                {
                    type = string.IsNullOrEmpty(idx) ? missingType : "body";
                }
                else if (type == "obj")
                {
                    type = "body";
                }
                return string.IsNullOrEmpty(idx) ? type : type + "|" + idx;
            }

            /// <summary>读取形状名称（nv*/cNvPr 的 name 属性）。Reads the shape name, the name attribute of the nv*/cNvPr element.</summary>
            /// <param name="shape">形状元素。The shape element.</param>
            /// <param name="name">非可视属性元素名（nvSpPr 等）。The non-visual property element name (nvSpPr and so on).</param>
            /// <returns>形状名称或空字符串。The shape name or an empty string.</returns>
            private static string ReadShapeName(XElement shape, string name)
            {
                foreach (XElement child in shape.Elements())
                {
                    if (child.Name.Namespace != Ppt.Main) continue;
                    if (child.Name.LocalName != name) continue;
                    string value = Ppt.Attribute(Ppt.Child(child, "cNvPr"), "name");
                    if (!string.IsNullOrEmpty(value)) return value;
                }
                return string.Empty;
            }

            /// <summary>把预设图形名映射为 AutoShapeKind；未知预设降级为 Rectangle。Maps a preset geometry name to AutoShapeKind; unknown presets degrade to Rectangle.</summary>
            /// <param name="preset">prstGeom 的 prst 属性。The prst attribute of prstGeom.</param>
            /// <returns>自选图形类别。The auto shape kind.</returns>
            private static AutoShapeKind MapAutoShapeKind(string preset)
            {
                if (string.IsNullOrEmpty(preset)) return AutoShapeKind.Rectangle;
                if (preset == "roundRect") return AutoShapeKind.RoundedRectangle;
                if (preset == "ellipse") return AutoShapeKind.Ellipse;
                if (preset == "triangle") return AutoShapeKind.Triangle;
                if (preset == "diamond") return AutoShapeKind.Diamond;
                if (preset == "pentagon") return AutoShapeKind.Pentagon;
                if (preset == "hexagon") return AutoShapeKind.Hexagon;
                if (preset == "rightArrow") return AutoShapeKind.RightArrow;
                if (preset == "chevron") return AutoShapeKind.Chevron;
                if (preset == "line") return AutoShapeKind.Line;
                return AutoShapeKind.Rectangle;
            }

            /// <summary>读取 roundRect 的调整参数还原圆角比例（adj 的 val 以十万分之一为单位）。Reads the roundRect adjustment to restore the corner radius ratio, where adj val is measured in hundred-thousandths.</summary>
            /// <param name="spPr">p:spPr 元素。The p:spPr element.</param>
            /// <returns>圆角比例；缺省沿用模型默认值 0.15。The ratio, or the model default of 0.15 when absent.</returns>
            private static double ReadRoundRectRadius(XElement spPr)
            {
                XElement avLst = Ppt.DrawingChild(Ppt.DrawingChild(spPr, "prstGeom"), "avLst");
                foreach (XElement guide in Ppt.DrawingChildren(avLst, "gd"))
                {
                    if (Ppt.Attribute(guide, "name") != "adj") continue;
                    // PresentationML 的调整参数写成 fmla="val N"，val 属性只是少数生成器的简写。
                    // PresentationML writes adjustments as fmla="val N"; a bare val attribute is a rare shorthand.
                    long value = ParseAdjustmentFormula(Ppt.Attribute(guide, "fmla"));
                    if (value < 0) value = ParseLong(Ppt.Attribute(guide, "val"), -1);
                    if (value >= 0) return Math.Min(0.5, value / 100000.0);
                }
                return 0.15;
            }

            /// <summary>解析 <c>fmla="val N"</c> 形式的调整公式，其余形式返回 -1。Parses an adjustment formula of the form <c>fmla="val N"</c>; any other form yields -1.</summary>
            /// <param name="fmla">公式文本。The formula text.</param>
            /// <returns>数值，无法解析时为 -1。The value, or -1 when it cannot be parsed.</returns>
            private static long ParseAdjustmentFormula(string fmla)
            {
                if (string.IsNullOrEmpty(fmla)) return -1;
                string text = fmla.Trim();
                if (!text.StartsWith("val ", StringComparison.OrdinalIgnoreCase)) return -1;
                return ParseLong(text.Substring(4).Trim(), -1);
            }

            /// <summary>
            /// 读取占位符类别：<c>type</c> 省略时按规范默认是 title，但实践中「省略 type 而写 idx」
            /// 是版式里的正文/内容占位符，故这种情况按 body 处理。
            /// Reads the placeholder kind: an omitted <c>type</c> defaults to title per the schema, but in real
            /// files a type omitted while an idx is written is the body or content placeholder of a layout, so
            /// that case is treated as body.
            /// </summary>
            /// <param name="ph">p:ph 元素，可为 null。The p:ph element; may be null.</param>
            /// <returns>占位符类别；非占位符返回 None。The placeholder kind, or None for a non-placeholder shape.</returns>
            private static PlaceholderKind MapPlaceholderKind(XElement ph)
            {
                return MapPlaceholderKind(ph, PlaceholderKind.Title);
            }

            /// <summary>与 <see cref="MapPlaceholderKind(XElement)"/> 相同，但由调用方给出「无 type 且无 idx」时的类别。Same as <see cref="MapPlaceholderKind(XElement)"/> with the caller supplying the kind used when both type and idx are absent.</summary>
            /// <param name="ph">p:ph 元素，可为 null。The p:ph element; may be null.</param>
            /// <param name="missingType">无 type 且无 idx 时采用的类别。The kind to use when both type and idx are absent.</param>
            /// <returns>占位符类别。The placeholder kind.</returns>
            private static PlaceholderKind MapPlaceholderKind(XElement ph, PlaceholderKind missingType)
            {
                if (ph == null) return PlaceholderKind.None;
                string type = Ppt.Attribute(ph, "type");
                if (!string.IsNullOrEmpty(type)) return MapPlaceholderType(type);
                return string.IsNullOrEmpty(Ppt.Attribute(ph, "idx")) ? missingType : PlaceholderKind.Body;
            }

            /// <summary>把 p:ph 的 type 映射为占位符类别；未列举的类型返回 None 但仍视为占位符。Maps a p:ph type onto the placeholder kind; unlisted types return None while still counting as placeholders.</summary>
            /// <param name="type">p:ph 的 type 属性。The type attribute of p:ph.</param>
            /// <returns>占位符类别。The placeholder kind.</returns>
            private static PlaceholderKind MapPlaceholderType(string type)
            {
                if (type == "body") return PlaceholderKind.Body;
                if (type == "obj") return PlaceholderKind.Content;
                if (type == "title" || type == "ctrTitle") return PlaceholderKind.Title;
                if (type == "subTitle") return PlaceholderKind.Subtitle;
                if (type == "pic") return PlaceholderKind.Picture;
                if (type == "tbl") return PlaceholderKind.Table;
                if (type == "chart" || type == "dgm") return PlaceholderKind.Chart;
                if (type == "dt" || type == "date") return PlaceholderKind.Date;
                if (type == "ftr" || type == "footer") return PlaceholderKind.Footer;
                if (type == "sldNum" || type == "slideNumber") return PlaceholderKind.SlideNumber;
                return PlaceholderKind.None;
            }

            /// <summary>把 a:pPr 的 algn 映射到段落对齐；其它预设值按左对齐处理。Maps a:pPr algn onto a paragraph alignment; remaining presets are treated as left alignment.</summary>
            /// <param name="value">algn 属性值。The algn attribute value.</param>
            /// <returns>段落对齐方式。The paragraph alignment.</returns>
            private static ParagraphAlignment MapAlignment(string value)
            {
                if (value == "ctr" || value == "dist") return ParagraphAlignment.Center;
                if (value == "just" || value == "justLow" || value == "thaiDist") return ParagraphAlignment.Justify;
                if (value == "r") return ParagraphAlignment.Right;
                return ParagraphAlignment.Left;
            }

            /// <summary>把 anchor 属性映射到垂直对齐（t/ctr/b）。Maps the anchor attribute to vertical alignment (t/ctr/b).</summary>
            /// <param name="value">anchor 属性值。The anchor attribute value.</param>
            /// <returns>垂直对齐方式。The vertical alignment.</returns>
            private static SlideTextAnchor MapVerticalAnchor(string value)
            {
                if (value == "ctr") return SlideTextAnchor.Middle;
                if (value == "b") return SlideTextAnchor.Bottom;
                return SlideTextAnchor.Top;
            }

            /// <summary>把 a:anchor 映射为文档模型的单元格垂直对齐。Maps a:anchor onto the document model's cell vertical alignment.</summary>
            /// <param name="value">anchor 属性值。The anchor attribute value.</param>
            /// <returns>单元格垂直对齐。The cell's vertical alignment.</returns>
            private static VerticalAlignment MapCellAnchor(string value)
            {
                if (value == "ctr") return VerticalAlignment.Middle;
                if (value == "b") return VerticalAlignment.Bottom;
                return VerticalAlignment.Top;
            }

            /// <summary>
            /// 依据幻灯片上的占位符与形状推断版式类别。pptx 不保存统一模型的 SlideLayoutKind，只能按占位符组合启发式还原；
            /// 含图片占位符或图片形状判为 TitleAndPicture，表格、图表占位符或表格形状判为 TitleAndTable，多个内容占位符判为 TwoContent。
            /// Infers the layout kind from the placeholders and shapes present on the slide. pptx does not store the unified
            /// SlideLayoutKind, so a placeholder-combination heuristic is used: a picture placeholder or picture shape yields
            /// TitleAndPicture, a table/chart placeholder or table shape yields TitleAndTable, and several content placeholders yield TwoContent.
            /// </summary>
            /// <param name="slide">已填充形状的幻灯片。The slide whose shapes were filled.</param>
            /// <returns>推断出的版式类别。The inferred layout kind.</returns>
            private static SlideLayoutKind InferLayoutKind(Slide slide)
            {
                bool hasTitle = false;
                bool hasPicture = false;
                bool hasTable = false;
                int content = 0;
                for (int i = 0; i < slide.Shapes.Count; i++)
                {
                    SlideShape shape = slide.Shapes[i];
                    PlaceholderKind kind = shape.Placeholder;
                    if (kind == PlaceholderKind.Title) hasTitle = true;
                    // 无占位符的表格与图片形状同样算证据：写出器的 graphicFrame 不带 p:ph。
                    // A table or picture without a placeholder counts too: the writer's graphicFrame carries no p:ph.
                    else if (kind == PlaceholderKind.Picture || shape is PictureShape) hasPicture = true;
                    else if (kind == PlaceholderKind.Table || kind == PlaceholderKind.Chart || shape is TableShape) hasTable = true;
                    else if (kind == PlaceholderKind.Body || kind == PlaceholderKind.Content) content++;
                }
                // 图片/表格占位符本身不足以判定「有标题」，只有 title 或 ctrTitle 才算标题。
                // A picture or table placeholder does not imply a title; only title or ctrTitle counts.
                if (hasPicture) return hasTitle ? SlideLayoutKind.TitleAndPicture : SlideLayoutKind.Blank;
                if (hasTable) return hasTitle ? SlideLayoutKind.TitleAndTable : SlideLayoutKind.Blank;
                if (hasTitle)
                {
                    if (content >= 2) return SlideLayoutKind.TwoContent;
                    if (content == 1) return SlideLayoutKind.TitleAndContent;
                    return SlideLayoutKind.Title;
                }
                // 只有内容占位符时归入 TitleAndContent：模型没有「仅内容」这一类别。
                // With content placeholders only, TitleAndContent is the closest kind, since the model has no
                // content-only category.
                return content > 0 ? SlideLayoutKind.TitleAndContent : SlideLayoutKind.Blank;
            }

            /// <summary>把形状加入幻灯片并按添加顺序分配标识（统一模型的 Id 供写出器使用）。Adds a shape to the slide and assigns its identifier from the insertion order, which the writer uses as the shape id.</summary>
            /// <param name="slide">目标幻灯片。The target slide.</param>
            /// <param name="shape">待加入的形状。The shape to add.</param>
            private static void AddShape(Slide slide, SlideShape shape)
            {
                if (shape == null) return;
                slide.Shapes.Add(shape);
                shape.Id = slide.Shapes.Count;
            }
        }

        /// <summary>
        /// 从文件路径加载 .pptx 演示文稿。
        /// Loads a .pptx presentation from a file path.
        /// </summary>
        /// <param name="path">.pptx 文件路径。The path of the .pptx file.</param>
        /// <returns>解析得到的统一演示文稿模型。The parsed unified presentation model.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> 为 null。Thrown when <paramref name="path"/> is null.</exception>
        /// <exception cref="InvalidDataException">文件不是有效的 .pptx 包，或实为旧版 .ppt。Thrown when the file is not a valid .pptx package, or is actually a legacy .ppt.</exception>
        public static OfficePresentation Load(string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(File.ReadAllBytes(path));
        }

        /// <summary>
        /// 从可读流加载 .pptx 演示文稿（流先完整复制到内存再解析）。
        /// Loads a .pptx presentation from a readable stream, copying it fully into memory before parsing.
        /// </summary>
        /// <param name="stream">包含 .pptx 字节内容的流。A stream containing the .pptx bytes.</param>
        /// <returns>解析得到的统一演示文稿模型。The parsed unified presentation model.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="stream"/> 为 null。Thrown when <paramref name="stream"/> is null.</exception>
        /// <exception cref="InvalidDataException">数据不是有效的 .pptx 包。Thrown when the data is not a valid .pptx package.</exception>
        public static OfficePresentation Load(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return Load(OfficePackageLimits.ReadAllBytes(stream));
        }

        /// <summary>
        /// 从字节数组加载 .pptx 演示文稿：先校验文件头（OLE 复合文档即旧版 .ppt 被拒绝，ZIP 签名必需），
        /// 再定位 presentation.xml 并逐张读取幻灯片。
        /// Loads a .pptx presentation from a byte array: the header is validated first (OLE compound files, i.e.
        /// legacy .ppt, are rejected and a ZIP signature is required), then presentation.xml is located and every
        /// slide is read in order.
        /// </summary>
        /// <param name="data">.pptx 文件的完整字节内容。The full bytes of the .pptx file.</param>
        /// <returns>解析得到的统一演示文稿模型。The parsed unified presentation model.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null。Thrown when <paramref name="data"/> is null.</exception>
        /// <exception cref="InvalidDataException">数据不是 .pptx 包、实为旧版 .ppt，或包已损坏/截断。Thrown when the data is not a .pptx package, is a legacy .ppt, or the package is corrupt or truncated.</exception>
        public static OfficePresentation Load(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length >= 4 && data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0)
            {
                throw new InvalidDataException("The file uses the legacy binary .ppt format, which this library does not support; save it as .pptx first.");
            }
            if (data.Length < 4 || data[0] != 0x50 || data[1] != 0x4B)
            {
                throw new InvalidDataException("The file is not a valid .pptx package.");
            }

            try
            {
                using (MemoryStream memory = new MemoryStream(data, false))
                using (ZipArchive archive = PackageArchive.OpenForRead(memory, false))
                {
                    return ReadPackage(new PptPackage(archive));
                }
            }
            catch (InvalidDataException)
            {
                throw;
            }
            catch (IOException ex)
            {
                throw new InvalidDataException("The .pptx package is corrupt or truncated.", ex);
            }
        }

        /// <summary>
        /// 提取演示文稿全部文本：逐张输出标题行与各形状文本，表格行以制表符分隔，备注以 [备注] 前缀跟在对应幻灯片后。
        /// Extracts all presentation text: one title line plus the text of every shape per slide, table rows
        /// separated by tabs, and speaker notes prefixed with [备注] after their slide.
        /// </summary>
        /// <param name="path">.pptx 文件路径。The path of the .pptx file.</param>
        /// <returns>纯文本内容。The plain-text content.</returns>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> 为 null。Thrown when <paramref name="path"/> is null.</exception>
        /// <exception cref="InvalidDataException">文件不是有效的 .pptx 包。Thrown when the file is not a valid .pptx package.</exception>
        public static string ExtractText(string path)
        {
            OfficePresentation presentation = Load(path);
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                Slide slide = presentation.Slides[i];
                if (builder.Length > 0) builder.AppendLine();
                string title = slide.GetTitle();
                if (!string.IsNullOrEmpty(title)) builder.AppendLine(title);
                for (int s = 0; s < slide.Shapes.Count; s++)
                {
                    TextBoxShape text = slide.Shapes[s] as TextBoxShape;
                    if (text != null)
                    {
                        // 标题已由 GetTitle 单独成行输出，此处跳过避免重复；空文本框不产生空行。
                        // The title is already emitted once by GetTitle, so it is skipped here; and a text box
                        // without text contributes no blank line.
                        if (text.Placeholder == PlaceholderKind.Title) continue;
                        if (string.IsNullOrEmpty(text.GetText())) continue;
                        foreach (Paragraph paragraph in text.Paragraphs)
                        {
                            builder.AppendLine(paragraph.GetText());
                        }
                        continue;
                    }
                    TableShape table = slide.Shapes[s] as TableShape;
                    if (table != null)
                    {
                        for (int r = 0; r < table.Table.Rows.Count; r++)
                        {
                            List<TableCell> cells = table.Table.Rows[r].Cells;
                            for (int c = 0; c < cells.Count; c++)
                            {
                                if (c > 0) builder.Append('\t');
                                for (int b = 0; b < cells[c].Blocks.Count; b++)
                                {
                                    Paragraph cellParagraph = cells[c].Blocks[b] as Paragraph;
                                    if (cellParagraph != null) builder.Append(cellParagraph.GetText());
                                }
                            }
                            if (cells.Count > 0) builder.AppendLine();
                        }
                    }
                }
                if (!string.IsNullOrEmpty(slide.NotesText))
                {
                    builder.Append("[备注] ").AppendLine(slide.NotesText);
                }
            }
            return builder.ToString();
        }

        /// <summary>
        /// 读取包的核心流程：经 _rels/.rels 定位 presentation.xml，读取幻灯片尺寸与元数据，
        /// 再按 p:sldIdLst 声明顺序逐张解析幻灯片；单个幻灯片部件失败只会跳过该幻灯片。
        /// Core package read flow: presentation.xml is located through _rels/.rels, the slide size and metadata are
        /// read, then each slide is parsed in p:sldIdLst order; a failure in one slide part only skips that slide.
        /// </summary>
        /// <param name="package">包访问器。The package accessor.</param>
        /// <returns>填充完成的演示文稿模型。The populated presentation model.</returns>
        private static OfficePresentation ReadPackage(PptPackage package)
        {
            string presentationPart = package.GetRelationshipTargetByType(string.Empty, Ppt.TypePresentation);
            if (string.IsNullOrEmpty(presentationPart)) presentationPart = "ppt/presentation.xml";
            XDocument presentationXml = package.GetXml(presentationPart);
            if (presentationXml == null || presentationXml.Root == null)
            {
                throw new InvalidDataException("The package does not contain a presentation part and is not a valid .pptx file.");
            }

            OfficePresentation presentation = new OfficePresentation();
            ReadSession session = new ReadSession(package, presentationPart, presentation);
            session.ReadSlideSize(presentationXml);
            session.ReadTheme();
            ReadMetadata(package, presentation.Metadata);

            // presentation.xml 的其余子元素按需求不读：p:notesSz 是备注页尺寸（模型只有幻灯片尺寸）、
            // p:clrMapOvr 与 p:defaultTextStyle 属于样式继承，p:sldMasterIdLst/p:sldLayoutIdLst 只在
            // 解析占位符几何与主题时按需经关系取用，不单独建列表。
            // The remaining presentation children are not read as specified: p:notesSz is the notes-page size
            // (the model only carries the slide size), p:clrMapOvr and p:defaultTextStyle belong to style
            // inheritance, and p:sldMasterIdLst plus p:sldLayoutIdLst are used through relationships only when
            // placeholder geometry or the theme is resolved.
            XElement slideIdList = Ppt.Child(presentationXml.Root, "sldIdLst");
            foreach (XElement slideId in Ppt.Children(slideIdList, "sldId"))
            {
                string slidePart = package.GetRelationshipTarget(presentationPart, Ppt.Attribute(slideId, Ppt.R + "id"));
                if (string.IsNullOrEmpty(slidePart)) continue;
                try
                {
                    XDocument slideXml = package.GetXml(slidePart);
                    if (slideXml == null || slideXml.Root == null) continue;
                    presentation.Slides.Add(new SlideReader(session, slidePart).Read(slideXml));
                }
                catch (XmlException)
                {
                    // 幻灯片 XML 结构异常：跳过该页继续读取其余内容。Malformed slide XML: skip this slide and keep reading.
                }
            }
            return presentation;
        }

        /// <summary>
        /// 读取 docProps/core.xml 元数据：dc:title/dc:subject/dc:creator/cp:keywords/dcterms 时间戳；
        /// dc:creator 与 Word/Excel 读取器一致地映射到 Author，部件缺失时保持模型默认值。
        /// Reads docProps/core.xml metadata: dc:title, dc:subject, dc:creator, cp:keywords and the dcterms
        /// timestamps. dc:creator maps to Author exactly as the Word and Excel readers do, and defaults are kept
        /// when the part is missing.
        /// </summary>
        /// <param name="package">包访问器。The package accessor.</param>
        /// <param name="metadata">待填充的元数据。The metadata to populate.</param>
        private static void ReadMetadata(PptPackage package, DocumentMetadata metadata)
        {
            XDocument core = package.GetXml("docProps/core.xml");
            if (core == null || core.Root == null) return;

            XElement title = core.Root.Element(Ppt.Dc + "title");
            if (title != null) metadata.Title = title.Value;
            XElement subject = core.Root.Element(Ppt.Dc + "subject");
            if (subject != null) metadata.Subject = subject.Value;
            XElement creator = core.Root.Element(Ppt.Dc + "creator");
            if (creator != null) metadata.Author = creator.Value;
            XElement keywords = core.Root.Element(Ppt.CoreProperties + "keywords");
            if (keywords != null) metadata.Keywords = keywords.Value;

            DateTime created;
            if (TryParseTimestamp(core.Root.Element(Ppt.DcTerms + "created"), out created)) metadata.Created = created;
            DateTime modified;
            if (TryParseTimestamp(core.Root.Element(Ppt.DcTerms + "modified"), out modified)) metadata.Modified = modified;
        }

        /// <summary>按不变区域设置解析 dcterms 时间戳（按 UTC 解释）。Parses a dcterms timestamp with the invariant culture, interpreting it as UTC.</summary>
        /// <param name="element">时间戳元素。The timestamp element.</param>
        /// <param name="value">解析结果；失败时为 DateTime.MinValue。The parsed value; DateTime.MinValue on failure.</param>
        /// <returns>解析成功返回 true。True when parsing succeeded.</returns>
        private static bool TryParseTimestamp(XElement element, out DateTime value)
        {
            value = DateTime.MinValue;
            if (element == null || string.IsNullOrEmpty(element.Value)) return false;
            return DateTime.TryParse(
                element.Value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                out value);
        }

        /// <summary>按不变区域设置解析 64 位整数（EMU 可超出 int 范围）。Parses a 64-bit integer with the invariant culture, since EMU values may exceed int range.</summary>
        /// <param name="value">待解析文本。The text to parse.</param>
        /// <param name="fallback">失败时的回退值。The fallback on failure.</param>
        /// <returns>解析结果或回退值。The parsed value or the fallback.</returns>
        private static long ParseLong(string value, long fallback)
        {
            long result;
            if (!string.IsNullOrEmpty(value) && long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out result))
            {
                return result;
            }
            return fallback;
        }
    }
}
