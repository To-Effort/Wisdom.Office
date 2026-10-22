using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>PDF 输出所用字体的抽象基类，统一度量、编码与写出接口。The abstract base class for fonts used in PDF output, unifying metrics, encoding and writing.</summary>
    /// <remarks>
    /// 典型流程：MarkUsed 收集用到的码点，Prepare 写出字体对象并登记资源，EncodeText 生成内容流字符串。
    /// The typical flow is: MarkUsed collects the code points used, Prepare writes the font object and
    /// registers it in the resources, and EncodeText produces the string for the content stream.
    /// </remarks>
    public abstract class PdfFont
    {
        /// <summary>资源字典中的字体名（如 F1）。The font name in the resource dictionary (e.g. F1).</summary>
        public string ResourceName { get; internal set; }
        /// <summary>调试用的字体名称。The font name for debugging.</summary>
        public abstract string DebugName { get; }
        /// <summary>
        /// 判断字体是否含指定码点的字形。
        /// Determines whether the font has a glyph for the code point.
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <returns>含该字形返回 true。True when the glyph exists.</returns>
        public abstract bool HasGlyph(int codePoint);
        /// <summary>
        /// 计算指定码点在给定字号下的前进宽度。
        /// Computes the advance width of the code point at the given font size.
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>前进宽度（点）。The advance width in points.</returns>
        public abstract double GetAdvance(int codePoint, double fontSize);
        /// <summary>
        /// 基线以上的 ascent 高度。
        /// The ascent height above the baseline.
        /// </summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>ascent 高度（点）。The ascent in points.</returns>
        public abstract double GetAscent(double fontSize);
        /// <summary>
        /// 基线以下的 descent 深度（正值）。
        /// The descent depth below the baseline (positive).
        /// </summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>descent 深度（点）。The descent in points.</returns>
        public abstract double GetDescent(double fontSize);
        /// <summary>
        /// 字体自然行高。
        /// The natural line height of the font.
        /// </summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>行高（点）。The line height in points.</returns>
        public abstract double GetLineHeight(double fontSize);
        /// <summary>
        /// 登记文本中实际用到的字符，供子集化收集码点。
        /// Records the characters actually used by the text, so subsetting can collect code points.
        /// </summary>
        /// <param name="text">用到的文本。The text being used.</param>
        public abstract void MarkUsed(string text);

        /// <summary>登记单个码点被使用，避免排版时先拼出临时字符串。
        /// Registers one code point as used without building a temporary string during layout.</summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        public virtual void MarkUsed(int codePoint)
        {
            MarkUsed(UnicodeUtil.ToUtf16String(codePoint));
        }
        /// <summary>
        /// 把字体写出为 PDF 对象并登记到页面资源字典。
        /// Writes the font as a PDF object and registers it in the page resource dictionary.
        /// </summary>
        /// <param name="writer">对象写出器。The object writer.</param>
        /// <param name="resourceFonts">页面 Font 资源字典。The page Font resource dictionary.</param>
        public abstract void Prepare(PdfObjectWriter writer, PdfDictionary resourceFonts);
        /// <summary>
        /// 把文本编码为内容流 Tj 使用的十六进制串。
        /// Encodes text into the hexadecimal string used by the Tj operator in a content stream.
        /// </summary>
        /// <param name="text">待编码文本。The text to encode.</param>
        /// <returns>十六进制编码结果。The hexadecimal encoding.</returns>
        public abstract string EncodeText(string text);
    }

    /// <summary>基于 TrueType 字体的嵌入子集字体（Type0/CIDFontType2）。An embedded TrueType subset font using the Type0/CIDFontType2 structure.</summary>
    /// <remarks>
    /// Prepare 时对 MarkUsed 收集的码点做子集化，重排 CID 并生成 CIDToGIDMap 与 ToUnicode；子集前缀为
    /// 六字母标签加「+」。编码按 Identity-H 双字节 CID 输出，因此必须先 Prepare 再 EncodeText。
    /// During Prepare the used code points are subset, CIDs are renumbered and the CIDToGIDMap and ToUnicode
    /// streams are built; the subset name is a six-letter tag plus "+". Encoding emits two-byte CIDs under
    /// Identity-H, so Prepare must run before EncodeText.
    /// </remarks>
    public sealed class TrueTypePdfFont : PdfFont
    {
        /// <summary>底层 TrueType 字面。The underlying TrueType face.</summary>
        private readonly TrueTypeFace _face;
        /// <summary>MarkUsed 收集的码点集合。The code points collected by MarkUsed.</summary>
        private readonly HashSet<int> _usedCodePoints = new HashSet<int>();
        /// <summary>六字母子集标签。The six-letter subset tag.</summary>
        private readonly string _subsetTag;
        /// <summary>Prepare 后生成的码点到 CID 映射。The code-point-to-CID map built by Prepare.</summary>
        private Dictionary<int, int> _cidByCodePoint;
        /// <summary>CID 到子集内新字形号的映射。The map from CID to the renumbered glyph id inside the subset.</summary>
        private int[] _gidByCid;
        /// <summary>CID 到原始文本的映射，用于生成 ToUnicode。The map from CID to original text, used to build ToUnicode.</summary>
        private string[] _textByCid;
        /// <summary>是否已完成写出。Whether the font has been prepared.</summary>
        private bool _prepared;

        /// <summary>
        /// 由字面与资源名创建嵌入字体，通常经 PdfFontSet 获得。
        /// Creates an embedded font from a face and resource name; normally obtained through PdfFontSet.
        /// </summary>
        /// <param name="face">TrueType 字面。The TrueType face.</param>
        /// <param name="resourceName">资源字典中的字体名。The font name in the resource dictionary.</param>
        /// <param name="subsetTag">六字母子集标签。The six-letter subset tag.</param>
        internal TrueTypePdfFont(TrueTypeFace face, string resourceName, string subsetTag)
        {
            _face = face;
            ResourceName = resourceName;
            _subsetTag = subsetTag;
        }

        /// <summary>底层 TrueType 字面。The underlying TrueType face.</summary>
        public TrueTypeFace Face { get { return _face; } }

        /// <summary>调试名，取家族名与子族名拼接。Debug name, built from the family and subfamily names.</summary>
        public override string DebugName
        {
            get { return _face.FamilyName + " " + _face.SubfamilyName; }
        }

        /// <summary>
        /// 委托字面判断是否含该码点的字形。
        /// Delegates to the face to decide whether the code point has a glyph.
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <returns>含该字形返回 true。True when the glyph exists.</returns>
        public override bool HasGlyph(int codePoint)
        {
            return _face.HasGlyph(codePoint);
        }

        /// <summary>
        /// 按字面 advance 与设计单位换算前进宽度。
        /// Converts the face advance width through the design-unit scale.
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>前进宽度（点）。The advance width in points.</returns>
        /// <remarks>码点无对应字形（映射到 0 且非 .notdef 命中）时返回 0 宽度。Returns 0 width when the code point has no glyph (maps to 0 without a .notdef hit).</remarks>
        public override double GetAdvance(int codePoint, double fontSize)
        {
            int gid = _face.GetGlyphId(codePoint);
            if (gid == 0)
            {
                return 0;
            }
            return _face.GetAdvanceWidth(gid) * fontSize / _face.UnitsPerEm;
        }

        /// <summary>
        /// 按字面 Ascender 换算 ascent。
        /// Scales the face Ascender into points.
        /// </summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>ascent 高度（点）。The ascent in points.</returns>
        public override double GetAscent(double fontSize)
        {
            return _face.Ascender * fontSize / _face.UnitsPerEm;
        }

        /// <summary>
        /// 按字面 Descender 换算 descent。
        /// Scales the face Descender into points.
        /// </summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>descent 深度（点，正值）。The descent in points, positive.</returns>
        /// <remarks>字体的 Descender 通常为负值，此处取反保证返回深度为正。The face's Descender is usually negative; it is negated here so the returned depth is positive.</remarks>
        public override double GetDescent(double fontSize)
        {
            return -_face.Descender * fontSize / _face.UnitsPerEm;
        }

        /// <summary>
        /// 行高取 Ascender−Descender+LineGap 的换算值。
        /// Line height is the scaled Ascender minus Descender plus LineGap.
        /// </summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>自然行高（点）。The natural line height in points.</returns>
        public override double GetLineHeight(double fontSize)
        {
            return (_face.Ascender - _face.Descender + _face.LineGap) * fontSize / _face.UnitsPerEm;
        }

        /// <summary>
        /// 把文本的每个码点加入子集候选集合。
        /// Adds every code point of the text to the subsetting candidate set.
        /// </summary>
        /// <param name="text">用到的文本，null 或空时忽略。The text being used; null or empty is ignored.</param>
        public override void MarkUsed(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            foreach (int codePoint in UnicodeUtil.EnumerateCodePoints(text))
            {
                _usedCodePoints.Add(codePoint);
            }
        }

        /// <summary>把码点直接加入子集候选集合。Adds the code point straight to the subsetting candidate set.</summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        public override void MarkUsed(int codePoint)
        {
            _usedCodePoints.Add(codePoint);
        }

        /// <summary>
        /// 按子集 CID 表把文本编码为双字节十六进制串。
        /// Encodes text into a two-byte-per-glyph hexadecimal string using the subset CID table.
        /// </summary>
        /// <param name="text">待编码文本。The text to encode.</param>
        /// <returns>4 位十六进制 CID 串联。The concatenated 4-digit hex CIDs.</returns>
        /// <exception cref="InvalidOperationException">尚未 Prepare 时抛出。Thrown when Prepare has not run yet.</exception>
        /// <remarks>CID 以 4 位十六进制大写输出（Identity-H 双字节）；未登记的码点回退到 CID 0。CIDs are emitted as 4 uppercase hex digits (two-byte Identity-H); unregistered code points fall back to CID 0.</remarks>
        public override string EncodeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            if (!_prepared) throw new InvalidOperationException("Font must be prepared before encoding text.");
            StringBuilder builder = new StringBuilder(text.Length * 4);
            foreach (int codePoint in UnicodeUtil.EnumerateCodePoints(text))
            {
                int cid;
                if (!_cidByCodePoint.TryGetValue(codePoint, out cid))
                {
                    cid = 0;
                }
                builder.Append(HexDigit(cid >> 12));
                builder.Append(HexDigit(cid >> 8));
                builder.Append(HexDigit(cid >> 4));
                builder.Append(HexDigit(cid));
            }
            return builder.ToString();
        }

        /// <summary>取一个十六进制位的大写字符。Takes the uppercase character of one hexadecimal nibble.</summary>
        /// <param name="value">含 4 位有效信息的数值。The value carrying 4 significant bits.</param>
        /// <returns>十六进制字符。The hex character.</returns>
        private static char HexDigit(int value)
        {
            int nibble = value & 15;
            return (char)(nibble < 10 ? '0' + nibble : 'A' + nibble - 10);
        }

        /// <summary>
        /// 执行子集化并写出完整字体链：FontFile2 → FontDescriptor → CIDFontType2 → Type0，同时登记 ToUnicode。
        /// Performs subsetting and writes the full font chain FontFile2 -> FontDescriptor -> CIDFontType2 ->
        /// Type0, along with the ToUnicode stream.
        /// </summary>
        /// <param name="writer">对象写出器。The object writer.</param>
        /// <param name="resourceFonts">页面 Font 资源字典。The page Font resource dictionary.</param>
        /// <exception cref="ArgumentNullException">writer 或 resourceFonts 为 null 时抛出。Thrown when writer or resourceFonts is null.</exception>
        /// <remarks>重复调用无效果；CID 从 1 起按排序后的码点编号，0 留给 .notdef。Repeated calls do nothing; CIDs start at 1 in sorted code-point order, with 0 reserved for .notdef.</remarks>
        public override void Prepare(PdfObjectWriter writer, PdfDictionary resourceFonts)
        {
            if (_prepared) return;
            if (writer == null) throw new ArgumentNullException("writer");
            if (resourceFonts == null) throw new ArgumentNullException("resourceFonts");

            List<int> codePoints = new List<int>(_usedCodePoints);
            codePoints.Sort();

            HashSet<int> glyphs = new HashSet<int>();
            glyphs.Add(0);
            for (int i = 0; i < codePoints.Count; i++)
            {
                int gid = _face.GetGlyphId(codePoints[i]);
                if (gid > 0) glyphs.Add(gid);
            }

            TrueTypeFace.SubsetResult subset = _face.BuildSubset(glyphs);
            _cidByCodePoint = new Dictionary<int, int>(codePoints.Count);
            _gidByCid = new int[codePoints.Count + 1];
            _textByCid = new string[codePoints.Count + 1];
            for (int i = 0; i < codePoints.Count; i++)
            {
                int oldGid = _face.GetGlyphId(codePoints[i]);
                int newGid;
                if (!subset.OldToNewGlyphId.TryGetValue(oldGid, out newGid)) newGid = 0;
                int cid = i + 1;
                _cidByCodePoint[codePoints[i]] = cid;
                _gidByCid[cid] = newGid;
                _textByCid[cid] = UnicodeUtil.ToUtf16String(codePoints[i]);
            }

            byte[] cidToGid = new byte[(codePoints.Count + 1) * 2];
            for (int cid = 0; cid < _gidByCid.Length; cid++)
            {
                cidToGid[cid * 2] = (byte)(_gidByCid[cid] >> 8);
                cidToGid[cid * 2 + 1] = (byte)_gidByCid[cid];
            }

            byte[] toUnicode = BuildToUnicodeCMap();
            string subsetFontName = _subsetTag + "+" + _face.PostScriptName;
            int scale = 1000;

            PdfDictionary fontFileDict = new PdfDictionary();
            fontFileDict.SetNumber("Length1", subset.Data.Length);
            PdfReference fontFileRef = writer.AddStream(fontFileDict, subset.Data, true);

            PdfDictionary descriptor = new PdfDictionary();
            descriptor.SetName("Type", "FontDescriptor");
            descriptor.SetName("FontName", subsetFontName);
            int flags = 4 | (_face.IsItalic ? 64 : 0);
            descriptor.SetNumber("Flags", flags);
            descriptor.Set("FontBBox", new PdfArray()
                .AddNumber(ScaleMetric(_face.XMin, scale))
                .AddNumber(ScaleMetric(_face.YMin, scale))
                .AddNumber(ScaleMetric(_face.XMax, scale))
                .AddNumber(ScaleMetric(_face.YMax, scale)));
            descriptor.SetNumber("ItalicAngle", Math.Round(_face.ItalicAngle, 4));
            descriptor.SetNumber("Ascent", ScaleMetric(_face.Ascender, scale));
            descriptor.SetNumber("Descent", ScaleMetric(_face.Descender, scale));
            descriptor.SetNumber("CapHeight", ScaleMetric(_face.CapHeight, scale));
            descriptor.SetNumber("XHeight", ScaleMetric(_face.XHeight, scale));
            descriptor.SetNumber("StemV", _face.IsBold ? 120 : 80);
            descriptor.Set("FontFile2", fontFileRef);
            if (!string.IsNullOrEmpty(_face.FamilyName))
            {
                descriptor.SetString("FontFamily", _face.FamilyName);
            }
            PdfReference descriptorRef = writer.Add(descriptor);

            PdfDictionary cidFont = new PdfDictionary();
            cidFont.SetName("Type", "Font");
            cidFont.SetName("Subtype", "CIDFontType2");
            cidFont.SetName("BaseFont", subsetFontName);
            PdfDictionary cidSystemInfo = new PdfDictionary();
            cidSystemInfo.Set("Registry", PdfString.Literal("Adobe"));
            cidSystemInfo.Set("Ordering", PdfString.Literal("Identity"));
            cidSystemInfo.SetNumber("Supplement", 0);
            cidFont.Set("CIDSystemInfo", cidSystemInfo);
            cidFont.Set("FontDescriptor", descriptorRef);

            PdfArray widths = new PdfArray();
            for (int cid = 1; cid < _gidByCid.Length; cid++)
            {
                int codePoint = codePoints[cid - 1];
                int advance = _face.GetAdvanceWidth(_face.GetGlyphId(codePoint));
                int width = (int)Math.Round(advance * 1000.0 / _face.UnitsPerEm);
                widths.AddNumber(cid);
                widths.Add(new PdfArray().AddNumber(width));
            }
            cidFont.Set("W", widths);
            cidFont.SetNumber("DW", 1000);
            PdfReference cidToGidRef = writer.AddStream(cidToGid, true);
            cidFont.Set("CIDToGIDMap", cidToGidRef);
            PdfReference cidFontRef = writer.Add(cidFont);

            PdfDictionary type0 = new PdfDictionary();
            type0.SetName("Type", "Font");
            type0.SetName("Subtype", "Type0");
            type0.SetName("BaseFont", subsetFontName);
            type0.SetName("Encoding", "Identity-H");
            type0.Set("DescendantFonts", new PdfArray().Add(cidFontRef));
            type0.Set("ToUnicode", writer.AddStream(toUnicode, true));
            PdfReference type0Ref = writer.Add(type0);
            resourceFonts.Set(ResourceName, type0Ref);

            _prepared = true;
        }

        /// <summary>
        /// 把字体设计单位度量缩放到给定 em（1000）空间。
        /// Scales a font design-unit metric into the given em (1000) space.
        /// </summary>
        /// <param name="value">设计单位值。The value in design units.</param>
        /// <param name="units">目标 em 空间。The target em space.</param>
        /// <returns>缩放并取整后的度量。The rounded scaled metric.</returns>
        private int ScaleMetric(int value, int units)
        {
            return (int)Math.Round(value * (double)units / _face.UnitsPerEm);
        }

        /// <summary>
        /// 生成 ToUnicode CMap：把子集 CID 映射回原始 UTF-16 文本，保证复制/检索得到正确 Unicode。
        /// Builds the ToUnicode CMap: maps subset CIDs back to the original UTF-16 text so copy and search
        /// yield correct Unicode.
        /// </summary>
        /// <returns>CMap 流字节（ASCII）。The CMap stream bytes (ASCII).</returns>
        /// <remarks>bfchar 段按每 100 条分块，符合规范对单段条目数的建议。The bfchar sections are chunked at 100 entries each, following the specification's guidance on section size.</remarks>
        private byte[] BuildToUnicodeCMap()
        {
            StringBuilder builder = new StringBuilder();
            builder.Append("/CIDInit /ProcSet findresource begin\n");
            builder.Append("12 dict begin\nbegincmap\n");
            builder.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
            builder.Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n");
            builder.Append("1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n");

            int total = _textByCid == null ? 0 : _textByCid.Length - 1;
            int index = 1;
            while (index <= total)
            {
                int count = Math.Min(100, total - index + 1);
                builder.Append(count.ToString(CultureInfo.InvariantCulture) + " beginbfchar\n");
                for (int i = 0; i < count; i++)
                {
                    int cid = index + i;
                    string text = _textByCid[cid];
                    builder.Append("<" + cid.ToString("X4", CultureInfo.InvariantCulture) + "> <");
                    builder.Append(UnicodeUtil.ToUtf16Hex(text));
                    builder.Append(">\n");
                }
                builder.Append("endbfchar\n");
                index += count;
            }

            builder.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
            return Encoding.ASCII.GetBytes(builder.ToString());
        }
    }

    /// <summary>PDF 标准 14 字体之一（Type1，WinAnsiEncoding），不嵌入字体文件。One of the PDF standard 14 fonts (Type1 with WinAnsiEncoding), embedded without a font file.</summary>
    /// <remarks>
    /// 度量按 Helvetica 近似值估算，只覆盖 Latin/常见标点，CJK 一律回退为 '?'；用于找不到 TrueType 字体时的兜底。
    /// Metrics use Helvetica approximations and cover only Latin/common punctuation; CJK falls back to '?'.
    /// It serves as the last resort when no TrueType font is available.
    /// </remarks>
    public sealed class StandardPdfFont : PdfFont
    {
        /// <summary>共享的 Helvetica 实例（资源名 F0）。The shared Helvetica instance (resource name F0).</summary>
        private static readonly StandardPdfFont _helvetica = new StandardPdfFont("Helvetica", "F0");
        /// <summary>BaseFont 名称。The BaseFont name.</summary>
        private readonly string _baseFont;
        /// <summary>是否已登记到资源字典。Whether the font has been registered in resources.</summary>
        private bool _prepared;

        /// <summary>共享的 Helvetica 标准字体。The shared Helvetica standard font.</summary>
        public static StandardPdfFont Helvetica { get { return _helvetica; } }

        /// <summary>
        /// 创建指定 BaseFont 与资源名的标准字体。
        /// Creates a standard font with the given BaseFont and resource name.
        /// </summary>
        /// <param name="baseFont">BaseFont 名称。The BaseFont name.</param>
        /// <param name="resourceName">资源字典中的字体名。The font name in the resource dictionary.</param>
        public StandardPdfFont(string baseFont, string resourceName)
        {
            _baseFont = baseFont;
            ResourceName = resourceName;
        }

        /// <summary>调试名即 BaseFont 名。The debug name, equal to the BaseFont name.</summary>
        public override string DebugName { get { return _baseFont; } }

        /// <summary>
        /// 码点可映射到 WinAnsi 字节即视为有字形。
        /// The code point has a glyph when it maps to a WinAnsi byte.
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <returns>可映射返回 true。True when mappable.</returns>
        public override bool HasGlyph(int codePoint)
        {
            return ToWinAnsi(codePoint) >= 0;
        }

        /// <summary>
        /// 按 Helvetica 近似宽度表估算前进宽度。
        /// Estimates the advance width from the Helvetica approximation table.
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>前进宽度（点）。The advance width in points.</returns>
        /// <remarks>宽度取 Helvetica 近似常数表，0x2E80 以上（CJK 区）按全角 1.0 计。Widths come from a Helvetica approximation table; code points at or above 0x2E80 (CJK) count as full-width 1.0.</remarks>
        public override double GetAdvance(int codePoint, double fontSize)
        {
            double width;
            if (codePoint == ' ') width = 0.278;
            else if (codePoint == '.' || codePoint == ',' || codePoint == ':' || codePoint == ';' || codePoint == '\'') width = 0.278;
            else if (codePoint == 'i' || codePoint == 'l' || codePoint == 'j' || codePoint == '!' || codePoint == '|') width = 0.278;
            else if (codePoint == 'f' || codePoint == 't' || codePoint == 'r') width = 0.333;
            else if (codePoint == 'm' || codePoint == 'w') width = 0.833;
            else if (codePoint == 'M' || codePoint == 'W') width = 0.889;
            else if (codePoint >= '0' && codePoint <= '9') width = 0.556;
            else if (codePoint >= 'A' && codePoint <= 'Z') width = 0.667;
            else if (codePoint >= 'a' && codePoint <= 'z') width = 0.5;
            else if (codePoint >= 0x2E80) width = 1.0;
            else width = 0.5;
            return width * fontSize;
        }

        /// <summary>标准字体近似 ascent（0.718×字号）。The standard-font approximation, 0.718 times the font size.</summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>ascent 高度（点）。The ascent in points.</returns>
        public override double GetAscent(double fontSize) { return 0.718 * fontSize; }
        /// <summary>标准字体近似 descent（0.207×字号）。The standard-font approximation, 0.207 times the font size.</summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>descent 深度（点）。The descent in points.</returns>
        public override double GetDescent(double fontSize) { return 0.207 * fontSize; }
        /// <summary>标准字体近似行高（1.15×字号）。The standard-font approximation, 1.15 times the font size.</summary>
        /// <param name="fontSize">字号（点）。The font size in points.</param>
        /// <returns>行高（点）。The line height in points.</returns>
        public override double GetLineHeight(double fontSize) { return 1.15 * fontSize; }
        /// <summary>
        /// 标准字体不嵌入，无需登记用量。
        /// Standard fonts are not embedded, so usage tracking is unnecessary.
        /// </summary>
        /// <param name="text">用到的文本。The text being used.</param>
        /// <remarks>标准字体不嵌入，无需收集码点，此实现为空。Standard fonts are not embedded, so no code-point collection is needed and this is a no-op.</remarks>
        public override void MarkUsed(string text) { }

        /// <summary>
        /// 写出轻量 Type1 字典并登记到资源。
        /// Writes the lightweight Type1 dictionary and registers it in the resources.
        /// </summary>
        /// <param name="writer">对象写出器。The object writer.</param>
        /// <param name="resourceFonts">页面 Font 资源字典。The page Font resource dictionary.</param>
        /// <remarks>只写出轻量 Type1 字典并登记资源，重复调用无效果。Only a lightweight Type1 dictionary is written and registered; repeated calls do nothing.</remarks>
        public override void Prepare(PdfObjectWriter writer, PdfDictionary resourceFonts)
        {
            if (_prepared) return;
            PdfDictionary font = new PdfDictionary();
            font.SetName("Type", "Font");
            font.SetName("Subtype", "Type1");
            font.SetName("BaseFont", _baseFont);
            font.SetName("Encoding", "WinAnsiEncoding");
            resourceFonts.Set(ResourceName, writer.Add(font));
            _prepared = true;
        }

        /// <summary>
        /// 把文本编码为单字节 WinAnsi 十六进制串。
        /// Encodes text as a one-byte-per-glyph WinAnsi hexadecimal string.
        /// </summary>
        /// <param name="text">待编码文本。The text to encode.</param>
        /// <returns>两位十六进制串联。The concatenated two-digit hex bytes.</returns>
        /// <remarks>逐码点转成单字节 WinAnsi（两位十六进制）；不可映射的字符写作 '?'。Each code point becomes one WinAnsi byte (two hex digits); unmappable characters are written as '?'.</remarks>
        public override string EncodeText(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            StringBuilder builder = new StringBuilder(text.Length * 2);
            foreach (int codePoint in UnicodeUtil.EnumerateCodePoints(text))
            {
                int value = ToWinAnsi(codePoint);
                if (value < 0) value = '?';
                builder.Append(HexDigit(value >> 4));
                builder.Append(HexDigit(value));
            }
            return builder.ToString();
        }

        /// <summary>取一个十六进制位的大写字符。Takes the uppercase character of one hexadecimal nibble.</summary>
        /// <param name="value">含 4 位有效信息的数值。The value carrying 4 significant bits.</param>
        /// <returns>十六进制字符。The hex character.</returns>
        private static char HexDigit(int value)
        {
            int nibble = value & 15;
            return (char)(nibble < 10 ? '0' + nibble : 'A' + nibble - 10);
        }

        /// <summary>
        /// 把 Unicode 码点转换为 WinAnsiEncoding 字节值。
        /// Converts a Unicode code point to its WinAnsiEncoding byte value.
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <returns>WinAnsi 字节值；不可映射返回 -1（空白类按空格处理）。The WinAnsi byte; -1 when unmappable (whitespace maps to space).</returns>
        private static int ToWinAnsi(int codePoint)
        {
            if (codePoint == 0x20AC) return 0x80;
            if (codePoint == 0x201A) return 0x82;
            if (codePoint == 0x0192) return 0x83;
            if (codePoint == 0x201E) return 0x84;
            if (codePoint == 0x2026) return 0x85;
            if (codePoint == 0x2020) return 0x86;
            if (codePoint == 0x2021) return 0x87;
            if (codePoint == 0x02C6) return 0x88;
            if (codePoint == 0x2030) return 0x89;
            if (codePoint == 0x0160) return 0x8A;
            if (codePoint == 0x2039) return 0x8B;
            if (codePoint == 0x0152) return 0x8C;
            if (codePoint == 0x017D) return 0x8E;
            if (codePoint == 0x2018) return 0x91;
            if (codePoint == 0x2019) return 0x92;
            if (codePoint == 0x201C) return 0x93;
            if (codePoint == 0x201D) return 0x94;
            if (codePoint == 0x2022) return 0x95;
            if (codePoint == 0x2013) return 0x96;
            if (codePoint == 0x2014) return 0x97;
            if (codePoint == 0x02DC) return 0x98;
            if (codePoint == 0x2122) return 0x99;
            if (codePoint == 0x0161) return 0x9A;
            if (codePoint == 0x203A) return 0x9B;
            if (codePoint == 0x0153) return 0x9C;
            if (codePoint == 0x017E) return 0x9E;
            if (codePoint == 0x0178) return 0x9F;

            if ((codePoint >= 0x20 && codePoint <= 0x7E) || (codePoint >= 0xA0 && codePoint <= 0xFF))
            {
                return codePoint;
            }
            if (codePoint == '\n' || codePoint == '\r' || codePoint == '\t') return ' ';
            return -1;
        }
    }

    /// <summary>一次输出的字体集合：按格式解析主字体、按码点做回退并统一分配资源名与子集标签。The font set for one output run: resolves primary fonts by format, falls back by code point and allocates resource names and subset tags.</summary>
    /// <remarks>
    /// 同一字面在集合内只对应一个 TrueTypePdfFont，主字体与回退字体各有缓存，避免重复子集化。
    /// A face maps to a single TrueTypePdfFont within the set; primary and fallback lookups are cached to
    /// avoid repeated subsetting.
    /// </remarks>
    public sealed class PdfFontSet
    {
        /// <summary>底层字体解析器。The underlying font resolver.</summary>
        private readonly FontResolver _resolver;
        /// <summary>「家族|粗|斜」到主字体的缓存。Cache from the family/bold/italic key to the primary font.</summary>
        private readonly Dictionary<PrimaryKey, PdfFont> _primaryCache = new Dictionary<PrimaryKey, PdfFont>();
        /// <summary>字面到嵌入字体的映射。The map from face to embedded font.</summary>
        private readonly Dictionary<TrueTypeFace, TrueTypePdfFont> _faceFonts = new Dictionary<TrueTypeFace, TrueTypePdfFont>();
        /// <summary>「类别|粗|斜」到回退字体的缓存。Cache from the category/bold/italic key to the fallback font.</summary>
        private readonly Dictionary<FallbackKey, PdfFont> _fallbackCache = new Dictionary<FallbackKey, PdfFont>();
        /// <summary>本集合累计创建的全部字体，PrepareAll 按此顺序写出。All fonts created by the set; PrepareAll writes them in this order.</summary>
        private readonly List<PdfFont> _fonts = new List<PdfFont>();
        /// <summary>惰性创建的 Helvetica 兜底标准字体。The lazily created Helvetica fallback standard font.</summary>
        private StandardPdfFont _standard;
        /// <summary>资源名/子集标签计数器。The resource-name and subset-tag counter.</summary>
        private int _fontCounter;

        /// <summary>
        /// 用进程级默认解析器创建字体集合。
        /// Creates a font set backed by the process-wide default resolver.
        /// </summary>
        public PdfFontSet()
            : this(null)
        {
        }

        /// <summary>
        /// 用指定解析器创建字体集合。
        /// Creates a font set with the given resolver.
        /// </summary>
        /// <param name="resolver">字体解析器，null 时使用 FontResolver.Default。The font resolver; null falls back to FontResolver.Default.</param>
        public PdfFontSet(FontResolver resolver)
        {
            _resolver = resolver ?? FontResolver.Default;
        }

        /// <summary>底层字体解析器。The underlying font resolver.</summary>
        public FontResolver Resolver { get { return _resolver; } }
        /// <summary>已创建的全部字体。All fonts created so far.</summary>
        public IReadOnlyList<PdfFont> Fonts { get { return _fonts; } }

        /// <summary>
        /// 取得指定文字格式的主字体，含未命中时按常规字重回退及 Arial/DejaVu/Liberation 兜底。
        /// Gets the primary font for the format, retrying at regular weight on a miss and finally falling
        /// back to Arial/DejaVu/Liberation or the standard Helvetica.
        /// </summary>
        /// <param name="format">文字格式，null 时使用默认格式。The text format; null uses the default format.</param>
        /// <returns>主字体，永不为 null。The primary font, never null.</returns>
        public PdfFont GetPrimary(TextFormat format)
        {
            if (format == null) format = new TextFormat();
            PrimaryKey key = new PrimaryKey(format.FontFamily, format.Bold, format.Italic);
            PdfFont cached;
            if (_primaryCache.TryGetValue(key, out cached))
            {
                return cached;
            }

            TrueTypeFace face = _resolver.Resolve(format.FontFamily, format.Bold, format.Italic);
            if (face == null && (format.Bold || format.Italic))
            {
                face = _resolver.Resolve(format.FontFamily, false, false);
            }
            if (face == null)
            {
                face = _resolver.Resolve("Arial", false, false)
                    ?? _resolver.Resolve("DejaVu Sans", false, false)
                    ?? _resolver.Resolve("Liberation Sans", false, false);
            }

            PdfFont font;
            if (face != null)
            {
                font = GetFontForFace(face);
            }
            else
            {
                if (_standard == null)
                {
                    _standard = new StandardPdfFont("Helvetica", NextResourceName());
                    _fonts.Add(_standard);
                }
                font = _standard;
            }

            _primaryCache[key] = font;
            return font;
        }

        /// <summary>
        /// 为码点选取可显示它的字体：先试主字体，未命中再按 CJK/emoji/符号/拉丁分类缓存回退字体。
        /// Chooses a font able to render the code point: tries the primary first, then the cached fallback by
        /// CJK/emoji/symbol/Latin category.
        /// </summary>
        /// <param name="format">文字格式。The text format.</param>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <returns>可用的字体；全部失败时仍返回主字体（呈现为 .notdef/?）。A usable font; when everything fails the primary is returned anyway (rendering .notdef/?).</returns>
        public PdfFont GetFontForCodePoint(TextFormat format, int codePoint)
        {
            PdfFont primary = GetPrimary(format);
            if (primary.HasGlyph(codePoint)) return primary;

            string category = FontResolver.IsCjkCodePoint(codePoint) ? "cjk"
                : (codePoint >= 0x1F000 && codePoint <= 0x1FAFF) ? "emoji"
                : (codePoint > 0xFF) ? "symbol" : "latin";
            FallbackKey key = new FallbackKey(category, format.Bold, format.Italic);
            PdfFont fallback;
            if (_fallbackCache.TryGetValue(key, out fallback) && fallback.HasGlyph(codePoint))
            {
                return fallback;
            }

            TrueTypeFace face = _resolver.ResolveFallback(codePoint, format.Bold, format.Italic);
            if (face != null)
            {
                PdfFont font = GetFontForFace(face);
                _fallbackCache[key] = font;
                return font;
            }

            return primary;
        }

        /// <summary>
        /// 取得字面对应的嵌入字体，同字面复用同一实例。
        /// Gets the embedded font for a face, reusing one instance per face.
        /// </summary>
        /// <param name="face">TrueType 字面。The TrueType face.</param>
        /// <returns>对应的嵌入字体。The matching embedded font.</returns>
        private TrueTypePdfFont GetFontForFace(TrueTypeFace face)
        {
            TrueTypePdfFont font;
            if (_faceFonts.TryGetValue(face, out font))
            {
                return font;
            }
            font = new TrueTypePdfFont(face, NextResourceName(), NextSubsetTag());
            _faceFonts[face] = font;
            _fonts.Add(font);
            return font;
        }

        /// <summary>
        /// 生成下一个字体资源名（F1、F2…）。
        /// Generates the next font resource name (F1, F2, ...).
        /// </summary>
        /// <returns>资源名字符串。The resource name string.</returns>
        private string NextResourceName()
        {
            _fontCounter++;
            return "F" + _fontCounter.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// 依据计数器生成六位 A–Z 子集标签（如 ABCAAA）。
        /// Derives a six-letter A-Z subset tag from the counter (e.g. ABCAAA).
        /// </summary>
        /// <returns>子集标签。The subset tag.</returns>
        private string NextSubsetTag()
        {
            int value = _fontCounter;
            char[] tag = new char[6];
            for (int i = 5; i >= 0; i--)
            {
                tag[i] = (char)('A' + (value % 26));
                value /= 26;
            }
            return new string(tag);
        }

        /// <summary>
        /// 依次写出集合内全部字体到资源字典。
        /// Writes every font in the set into the resource dictionary.
        /// </summary>
        /// <param name="writer">对象写出器。The object writer.</param>
        /// <param name="fontResources">页面 Font 资源字典。The page Font resource dictionary.</param>
        public void PrepareAll(PdfObjectWriter writer, PdfDictionary fontResources)
        {
            for (int i = 0; i < _fonts.Count; i++)
            {
                _fonts[i].Prepare(writer, fontResources);
            }
        }

        /// <summary>主字体缓存键：家族名（忽略大小写与首尾空白）+ 粗体 + 斜体。
        /// Primary font cache key: family (case- and trim-insensitive) plus bold and italic.</summary>
        private struct PrimaryKey : IEquatable<PrimaryKey>
        {
            private readonly string _family;
            private readonly bool _bold;
            private readonly bool _italic;

            public PrimaryKey(string family, bool bold, bool italic)
            {
                _family = string.IsNullOrEmpty(family) ? "Arial" : family;
                _bold = bold;
                _italic = italic;
            }

            public bool Equals(PrimaryKey other)
            {
                return _bold == other._bold
                    && _italic == other._italic
                    && string.Equals(_family, other._family, StringComparison.OrdinalIgnoreCase);
            }

            public override bool Equals(object obj)
            {
                return obj is PrimaryKey && Equals((PrimaryKey)obj);
            }

            public override int GetHashCode()
            {
                return StringComparer.OrdinalIgnoreCase.GetHashCode(_family) ^ (_bold ? 1 : 0) ^ (_italic ? 2 : 0);
            }
        }

        /// <summary>回退字体缓存键：字符类别 + 粗体 + 斜体。Fallback cache key: character category plus bold and italic.</summary>
        private struct FallbackKey : IEquatable<FallbackKey>
        {
            private readonly string _category;
            private readonly bool _bold;
            private readonly bool _italic;

            public FallbackKey(string category, bool bold, bool italic)
            {
                _category = category;
                _bold = bold;
                _italic = italic;
            }

            public bool Equals(FallbackKey other)
            {
                return _bold == other._bold
                    && _italic == other._italic
                    && string.Equals(_category, other._category, StringComparison.OrdinalIgnoreCase);
            }

            public override bool Equals(object obj)
            {
                return obj is FallbackKey && Equals((FallbackKey)obj);
            }

            public override int GetHashCode()
            {
                return StringComparer.OrdinalIgnoreCase.GetHashCode(_category) ^ (_bold ? 1 : 0) ^ (_italic ? 2 : 0);
            }
        }
    }

    /// <summary>UTF 码点枚举与转换小工具集。A small utility set for enumerating and converting UTF code points.</summary>
    public static class UnicodeUtil
    {
        /// <summary>
        /// 按码点遍历字符串，正确配对 UTF-16 代理对。
        /// Iterates the string by code point, pairing UTF-16 surrogates correctly.
        /// </summary>
        /// <param name="text">待遍历文本，null 或空时不产出元素。The text to walk; null or empty yields no elements.</param>
        /// <returns>码点序列。The sequence of code points.</returns>
        public static IEnumerable<int> EnumerateCodePoints(string text)
        {
            if (string.IsNullOrEmpty(text)) yield break;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                {
                    yield return char.ConvertToUtf32(c, text[i + 1]);
                    i++;
                }
                else
                {
                    yield return c;
                }
            }
        }

        /// <summary>
        /// 把码点转成 UTF-16 字符串，非法码点退化为 "?"。
        /// Converts a code point to a UTF-16 string, degrading invalid values to "?".
        /// </summary>
        /// <param name="codePoint">Unicode 码点。The Unicode code point.</param>
        /// <returns>转换结果或 "?"。The converted text, or "?".</returns>
        public static string ToUtf16String(int codePoint)
        {
            try
            {
                return char.ConvertFromUtf32(codePoint);
            }
            catch (Exception)
            {
                return "?";
            }
        }

        /// <summary>
        /// 把 UTF-16 字符串转为四 digits 一码元的十六进制文本，用于 ToUnicode CMap。
        /// Converts a UTF-16 string to hex text of four digits per unit, used by the ToUnicode CMap.
        /// </summary>
        /// <param name="text">待转换文本。The text to convert.</param>
        /// <returns>大写十六进制串。The uppercase hexadecimal string.</returns>
        public static string ToUtf16Hex(string text)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                builder.Append(((int)text[i]).ToString("X4", CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }
    }
}
