using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 解码得到的一个字元：编码值、对应文本与宽度。
    /// One decoded glyph: its encoding value, restored text and width.
    /// </summary>
    internal sealed class PdfGlyph
    {
        /// <summary>还原出的文本，无法还原时为空字符串。The restored text; an empty string when it cannot be recovered.</summary>
        public string Text;
        /// <summary>字体内的编码值（单字节码或 CID）。The encoding value inside the font (single-byte code or CID).</summary>
        public int Code;
        /// <summary>字形宽度，单位为千分之一 em。The glyph width in thousandths of an em.</summary>
        public int Width;
        /// <summary>是否为词间空格，供分词与断行使用。Whether this is an inter-word space, used for tokenizing and line breaking.</summary>
        public bool IsWordSpace;
    }

    /// <summary>
    /// PDF 字体度量与编码解码器，负责把内容流中的字节串翻译为 Unicode 文本并给出宽度。
    /// PDF font metrics and encoding decoder: translates byte strings from content streams into Unicode
    /// text and provides widths.
    /// </summary>
    /// <remarks>文本还原优先级为：ToUnicode CMap、内置差异表（单字节）、WinAnsi 回退；两字节 CID 字体若无 ToUnicode 则只能得到空文本，但仍能取得宽度用于排版。
    /// Text restoration tries, in order, the ToUnicode CMap, the built Differences table (single-byte fonts)
    /// and the WinAnsi fallback; a two-byte CID font without ToUnicode yields empty text but still provides
    /// widths for layout.</remarks>
    internal sealed class PdfTextFont
    {
        /// <summary>常用 AGL 字形名到 Unicode 的静态映射表。Static map from common AGL glyph names to Unicode.</summary>
        private static readonly Dictionary<string, string> GlyphNames = BuildGlyphNames();

        /// <summary>所属文档，用于解引用与解码流。The owning document, used for dereferencing and stream decoding.</summary>
        private readonly PdfDocument _document;
        /// <summary>是否为两字节编码（Type0 / CID 字体）。Whether the encoding is two-byte (Type0 / CID fonts).</summary>
        private readonly bool _twoByte;
        /// <summary>编码是否近似 Identity（当前与两字节标记同值）。Whether the encoding is Identity-like (currently the same as the two-byte flag).</summary>
        private readonly bool _identityLike;
        /// <summary>ToUnicode CMap，解析失败时为 null。The ToUnicode CMap; null when parsing failed.</summary>
        private readonly PdfCMap _toUnicode;
        /// <summary>码位到宽度的表（千分之一 em）。Map from code to width in thousandths of an em.</summary>
        private readonly Dictionary<int, int> _widths = new Dictionary<int, int>();
        /// <summary>自定义差异表：码位到 Unicode 文本。Custom Differences table: code to Unicode text.</summary>
        private readonly Dictionary<int, string> _differences = new Dictionary<int, string>();
        /// <summary>查不到宽度时使用的默认宽度（千分之一 em）。The default width used when a lookup misses, in thousandths of an em.</summary>
        private readonly int _defaultWidth;
        /// <summary>字体的 BaseFont 名称。The BaseFont name of the font.</summary>
        public string BaseFontName { get; private set; }

        /// <summary>
        /// 读取字体字典及其后代字体、编码信息，建立宽度表与差异表。
        /// Reads the font dictionary plus its descendant font and encoding, building the width and
        /// differences tables.
        /// </summary>
        /// <param name="document">所属文档，用于解引用和解码流。The owning document, used for dereferencing and stream decoding.</param>
        /// <param name="dictionary">字体字典（Type0、CIDFontType0/2 或简单字体）。The font dictionary (Type0, CIDFontType0/2 or a simple font).</param>
        /// <exception cref="ArgumentNullException"><paramref name="dictionary"/> 为 null 时抛出。Thrown when <paramref name="dictionary"/> is null.</exception>
        /// <remarks>ToUnicode 解析过程中的任何异常都会被吞掉并退回内置编码表，以保证读取整篇文档不因个别坏字体而失败。
        /// Any exception raised while parsing ToUnicode is swallowed and the built-in encoding tables are
        /// used instead, so one broken font cannot fail the whole document read.</remarks>
        public PdfTextFont(PdfDocument document, PdfDictionary dictionary)
        {
            _document = document;
            if (dictionary == null) throw new ArgumentNullException("dictionary");

            PdfName baseFont = Resolve(dictionary.Get("BaseFont")) as PdfName;
            BaseFontName = baseFont == null ? string.Empty : baseFont.Value;

            PdfName subtype = Resolve(dictionary.Get("Subtype")) as PdfName;
            string subtypeName = subtype == null ? string.Empty : subtype.Value;
            _twoByte = subtypeName == "Type0" || subtypeName == "CIDFontType0" || subtypeName == "CIDFontType2";
            _identityLike = _twoByte;

            PdfObject toUnicodeObject = Resolve(dictionary.Get("ToUnicode"));
            PdfStream toUnicodeStream = toUnicodeObject as PdfStream;
            if (toUnicodeStream != null)
            {
                try
                {
                    _toUnicode = PdfCMap.Parse(_document.DecodeStream(toUnicodeStream));
                }
                catch (Exception)
                {
                    _toUnicode = null;
                }
            }

            if (_twoByte)
            {
                _defaultWidth = 1000;
                PdfArray descendants = Resolve(dictionary.Get("DescendantFonts")) as PdfArray;
                PdfDictionary descendant = null;
                if (descendants != null && descendants.Count > 0)
                {
                    descendant = Resolve(descendants[0]) as PdfDictionary;
                }
                if (descendant != null)
                {
                    PdfNumber defaultWidth = Resolve(descendant.Get("DW")) as PdfNumber;
                    if (defaultWidth != null) _defaultWidth = (int)defaultWidth.Value;
                    ReadCidWidths(Resolve(descendant.Get("W")) as PdfArray);
                }
                return;
            }

            _defaultWidth = 500;
            PdfNumber missing = null;
            PdfDictionary descriptor = Resolve(dictionary.Get("FontDescriptor")) as PdfDictionary;
            if (descriptor != null) missing = Resolve(descriptor.Get("MissingWidth")) as PdfNumber;
            if (missing != null) _defaultWidth = (int)missing.Value;

            PdfNumber firstChar = Resolve(dictionary.Get("FirstChar")) as PdfNumber;
            PdfArray widths = Resolve(dictionary.Get("Widths")) as PdfArray;
            if (firstChar != null && widths != null)
            {
                int start = (int)firstChar.Value;
                for (int i = 0; i < widths.Count; i++)
                {
                    PdfNumber width = Resolve(widths[i]) as PdfNumber;
                    if (width != null) _widths[start + i] = (int)width.Value;
                }
            }

            PdfObject encodingObject = Resolve(dictionary.Get("Encoding"));
            PdfDictionary encoding = encodingObject as PdfDictionary;
            if (encoding != null)
            {
                PdfArray differences = Resolve(encoding.Get("Differences")) as PdfArray;
                if (differences != null) ReadDifferences(differences);
            }
        }

        /// <summary>字体是否使用两字节编码，决定字符串读取时的步长。Whether the font uses two-byte encoding, deciding the stride when reading strings.</summary>
        public bool IsTwoByte { get { return _twoByte; } }

        /// <summary>
        /// 解析引用对象为其目标值。
        /// Dereferences an object to its target value.
        /// </summary>
        /// <param name="value">可能是引用的对象。The object, which may be a reference.</param>
        /// <returns>实际对象或原值。The actual object, or the original value.</returns>
        private PdfObject Resolve(PdfObject value)
        {
            return _document.Resolve(value);
        }

        /// <summary>
        /// 读取 CID 字体的 /W 宽度数组，支持「起始码 [宽度…]」与「起始码 结束码 统一宽度」两种形式。
        /// Reads a CID font's /W width array, supporting both "start [widths...]" and
        /// "start end unified-width" entry forms.
        /// </summary>
        /// <param name="widths">/W 数组，null 时直接返回。The /W array; returns immediately when null.</param>
        /// <remarks>单项区间跨度超过 65536 视为异常数据跳过，避免畸形宽度表撑爆内存。
        /// Ranges wider than 65536 are treated as corrupt and skipped, keeping a malformed width table from
        /// exhausting memory.</remarks>
        private void ReadCidWidths(PdfArray widths)
        {
            if (widths == null) return;
            int index = 0;
            while (index < widths.Count)
            {
                PdfNumber first = Resolve(widths[index]) as PdfNumber;
                if (first == null) break;
                index++;
                int startCode = (int)first.Value;
                if (index >= widths.Count) break;
                PdfArray list = Resolve(widths[index]) as PdfArray;
                if (list != null)
                {
                    index++;
                    for (int i = 0; i < list.Count; i++)
                    {
                        PdfNumber width = Resolve(list[i]) as PdfNumber;
                        if (width != null) _widths[startCode + i] = (int)width.Value;
                    }
                    continue;
                }
                PdfNumber last = Resolve(widths[index]) as PdfNumber;
                index++;
                if (last == null || index >= widths.Count) break;
                PdfNumber value = Resolve(widths[index]) as PdfNumber;
                index++;
                if (value == null) continue;
                int endCode = (int)last.Value;
                // 同 CMap：码位来自声明值，int 相减会回绕，跨度比较必须用 long。
                // As in the CMap: codes come from declared values and the int subtraction wraps, so the span
                // is compared in long.
                if (endCode < startCode || (long)endCode - (long)startCode > 65536) continue;
                if (_widths.Count >= PdfReadLimits.MaxGlyphMapEntries) break;
                for (int code = startCode; code <= endCode; code++) _widths[code] = (int)value.Value;
            }
        }

        /// <summary>
        /// 读取自定义编码的 /Differences 数组，把字形名转成 Unicode 并按码位记录。
        /// Reads a custom encoding's /Differences array, converting glyph names to Unicode keyed by code.
        /// </summary>
        /// <param name="differences">差异数组，元素为码位数字或字形名。The differences array, holding code numbers or glyph names.</param>
        private void ReadDifferences(PdfArray differences)
        {
            int code = 0;
            for (int i = 0; i < differences.Count; i++)
            {
                PdfNumber number = Resolve(differences[i]) as PdfNumber;
                if (number != null)
                {
                    code = (int)number.Value;
                    continue;
                }
                PdfName name = Resolve(differences[i]) as PdfName;
                if (name == null) continue;
                string text = GlyphNameToUnicode(name.Value);
                if (!string.IsNullOrEmpty(text)) _differences[code] = text;
                code++;
            }
        }

        /// <summary>
        /// 把显示串（Tj/TJ 的字节内容）解码为字元序列。
        /// Decodes a display string (the byte content of Tj/TJ) into a glyph sequence.
        /// </summary>
        /// <param name="bytes">编码字节，null 或空时返回空列表。The encoded bytes; null or empty yields an empty list.</param>
        /// <returns>逐码位的字元列表。The glyph list, one per code.</returns>
        /// <remarks>两字节字体若末尾落单（数据被截断），该残字节按单字节码位处理而不丢弃。
        /// When a two-byte font's data ends with an odd trailing byte (truncated input), that byte is
        /// treated as a single-byte code instead of being dropped.</remarks>
        public List<PdfGlyph> Decode(byte[] bytes)
        {
            List<PdfGlyph> result = new List<PdfGlyph>();
            if (bytes == null || bytes.Length == 0) return result;

            if (_twoByte)
            {
                for (int i = 0; i + 1 < bytes.Length; i += 2)
                {
                    int code = (bytes[i] << 8) | bytes[i + 1];
                    AddGlyph(result, code);
                }
                if ((bytes.Length & 1) != 0) AddGlyph(result, bytes[bytes.Length - 1]);
                return result;
            }

            for (int i = 0; i < bytes.Length; i++) AddGlyph(result, bytes[i]);
            return result;
        }

        /// <summary>
        /// 为单个码位生成字元，按 ToUnicode、差异表、WinAnsi 的顺序尝试还原文本。
        /// Builds one glyph for a code, restoring text by trying ToUnicode, the Differences table and
        /// WinAnsi in turn.
        /// </summary>
        /// <param name="result">输出列表，字元追加到末尾。The output list; the glyph is appended at the end.</param>
        /// <param name="code">码位。The code.</param>
        private void AddGlyph(List<PdfGlyph> result, int code)
        {
            PdfGlyph glyph = new PdfGlyph();
            glyph.Code = code;
            glyph.Width = DefaultWidthFor(code);
            glyph.IsWordSpace = code == 32;
            string text;
            if (_toUnicode != null && _toUnicode.TryGet(code, out text))
            {
                glyph.Text = text;
            }
            else if (_twoByte)
            {
                glyph.Text = string.Empty;
            }
            else if (_differences.TryGetValue(code, out text))
            {
                glyph.Text = text;
            }
            else
            {
                glyph.Text = WinAnsiToUnicode(code);
            }
            result.Add(glyph);
        }

        /// <summary>
        /// 取码位宽度，缺失时回退到字体默认宽度。
        /// Gets the width of a code, falling back to the font's default width when absent.
        /// </summary>
        /// <param name="code">码位。The code.</param>
        /// <returns>千分之一 em 为单位的宽度。The width in thousandths of an em.</returns>
        private int DefaultWidthFor(int code)
        {
            int width;
            if (_widths.TryGetValue(code, out width)) return width;
            return _defaultWidth;
        }

        /// <summary>
        /// 把 WinAnsiEncoding 的码位翻译为 Unicode 字符：0x20–0x7E 与 0xA0–0xFF 等同 Latin-1，0x80–0x9F 为标点符号区。
        /// Translates a WinAnsiEncoding code into a Unicode character: 0x20-0x7E and 0xA0-0xFF equal
        /// Latin-1, while 0x80-0x9F is the punctuation block.
        /// </summary>
        /// <param name="code">0–255 的码位。A code from 0 to 255.</param>
        /// <returns>对应字符；0x81、0x8D 等未定义码位返回空字符串。The matching character; undefined codes such as 0x81 and 0x8D return an empty string.</returns>
        public static string WinAnsiToUnicode(int code)
        {
            if ((code >= 0x20 && code <= 0x7E) || (code >= 0xA0 && code <= 0xFF))
            {
                return ((char)code).ToString();
            }
            if (code == 0x80) return "\u20AC";
            if (code == 0x82) return "\u201A";
            if (code == 0x83) return "\u0192";
            if (code == 0x84) return "\u201E";
            if (code == 0x85) return "\u2026";
            if (code == 0x86) return "\u2020";
            if (code == 0x87) return "\u2021";
            if (code == 0x88) return "\u02C6";
            if (code == 0x89) return "\u2030";
            if (code == 0x8A) return "\u0160";
            if (code == 0x8B) return "\u2039";
            if (code == 0x8C) return "\u0152";
            if (code == 0x8E) return "\u017D";
            if (code == 0x91) return "\u2018";
            if (code == 0x92) return "\u2019";
            if (code == 0x93) return "\u201C";
            if (code == 0x94) return "\u201D";
            if (code == 0x95) return "\u2022";
            if (code == 0x96) return "\u2013";
            if (code == 0x97) return "\u2014";
            if (code == 0x98) return "\u02DC";
            if (code == 0x99) return "\u2122";
            if (code == 0x9A) return "\u0161";
            if (code == 0x9B) return "\u203A";
            if (code == 0x9C) return "\u0153";
            if (code == 0x9E) return "\u017E";
            if (code == 0x9F) return "\u0178";
            return string.Empty;
        }

        /// <summary>
        /// 把字形名翻译为 Unicode 文本。
        /// Translates a glyph name into Unicode text.
        /// </summary>
        /// <param name="name">字形名，如 comma、uni4F60、u4F60、Aacute 形式的组合名。The glyph name, e.g. comma, uni4F60, u4F60 or composed names like Aacute.</param>
        /// <returns>对应文本；无法识别时返回空字符串。The matching text; an empty string when unrecognized.</returns>
        /// <remarks>依次尝试内置 AGL 表、uniXXXX 系列（每 4 位十六进制一个字符）、uXXXX 系列（可超出基本平面，按 UTF-32 转 UTF-16），
        /// 最后处理「单字母 + 抑扬符后缀」这类组合名，先拼出组合序列再做 NFC 合成。
        /// It tries the built-in AGL table, the uniXXXX family (one character per 4 hex digits), the uXXXX
        /// family (may exceed the BMP, converted from UTF-32 to UTF-16) and finally composed names of
        /// "single letter + diacritic suffix", which are decomposed and then normalized to NFC.</remarks>
        public static string GlyphNameToUnicode(string name)
        {
            if (string.IsNullOrEmpty(name)) return string.Empty;
            string direct;
            if (GlyphNames.TryGetValue(name, out direct)) return direct;

            if (name.Length > 3 && name.StartsWith("uni", StringComparison.Ordinal))
            {
                StringBuilder builder = new StringBuilder();
                for (int i = 3; i + 4 <= name.Length; i += 4)
                {
                    int value;
                    if (!int.TryParse(name.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value))
                    {
                        return string.Empty;
                    }
                    builder.Append((char)value);
                }
                if (builder.Length > 0) return builder.ToString();
            }
            if (name.Length >= 5 && name[0] == 'u')
            {
                int value;
                if (int.TryParse(name.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out value) && value <= 0x10FFFF)
                {
                    try
                    {
                        return char.ConvertFromUtf32(value);
                    }
                    catch (ArgumentOutOfRangeException)
                    {
                        return string.Empty;
                    }
                }
            }

            string[] accentSuffixes = { "grave", "acute", "circumflex", "tilde", "dieresis", "ring", "cedilla", "caron", "breve", "ogonek", "dotaccent" };
            string[] combining = { "\u0300", "\u0301", "\u0302", "\u0303", "\u0308", "\u030A", "\u0327", "\u030C", "\u0306", "\u0328", "\u0307" };
            for (int i = 0; i < accentSuffixes.Length; i++)
            {
                if (name.Length == 1 + accentSuffixes[i].Length && name.EndsWith(accentSuffixes[i], StringComparison.Ordinal))
                {
                    char baseChar = name[0];
                    if ((baseChar >= 'A' && baseChar <= 'Z') || (baseChar >= 'a' && baseChar <= 'z'))
                    {
                        string composed = (baseChar + combining[i]).Normalize(NormalizationForm.FormC);
                        return composed;
                    }
                }
            }
            return string.Empty;
        }

        /// <summary>
        /// 构建字形名到 Unicode 的静态映射表（字母、数字、常用标点、排版符号与连字）。
        /// Builds the static glyph-name-to-Unicode table (letters, digits, common punctuation, typographic
        /// marks and ligatures).
        /// </summary>
        /// <returns>以序号比较键名的映射表。The table keyed by ordinal-compared names.</returns>
        private static Dictionary<string, string> BuildGlyphNames()
        {
            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.Ordinal);
            for (char c = 'A'; c <= 'Z'; c++) map[c.ToString()] = c.ToString();
            for (char c = 'a'; c <= 'z'; c++) map[c.ToString()] = c.ToString();
            string[] digitNames = { "zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine" };
            for (int i = 0; i < digitNames.Length; i++) map[digitNames[i]] = i.ToString(CultureInfo.InvariantCulture);

            Add(map, "space", " ");
            Add(map, "exclam", "!");
            Add(map, "quotedbl", "\"");
            Add(map, "numbersign", "#");
            Add(map, "dollar", "$");
            Add(map, "percent", "%");
            Add(map, "ampersand", "&");
            Add(map, "quotesingle", "'");
            Add(map, "quoteright", "\u2019");
            Add(map, "quoteleft", "\u2018");
            Add(map, "parenleft", "(");
            Add(map, "parenright", ")");
            Add(map, "asterisk", "*");
            Add(map, "plus", "+");
            Add(map, "comma", ",");
            Add(map, "hyphen", "-");
            Add(map, "period", ".");
            Add(map, "slash", "/");
            Add(map, "colon", ":");
            Add(map, "semicolon", ";");
            Add(map, "less", "<");
            Add(map, "equal", "=");
            Add(map, "greater", ">");
            Add(map, "question", "?");
            Add(map, "at", "@");
            Add(map, "bracketleft", "[");
            Add(map, "backslash", "\\");
            Add(map, "bracketright", "]");
            Add(map, "asciicircum", "^");
            Add(map, "underscore", "_");
            Add(map, "grave", "`");
            Add(map, "braceleft", "{");
            Add(map, "bar", "|");
            Add(map, "braceright", "}");
            Add(map, "asciitilde", "~");
            Add(map, "quotesinglbase", "\u201A");
            Add(map, "quotedblleft", "\u201C");
            Add(map, "quotedblright", "\u201D");
            Add(map, "quotedblbase", "\u201E");
            Add(map, "bullet", "\u2022");
            Add(map, "endash", "\u2013");
            Add(map, "emdash", "\u2014");
            Add(map, "ellipsis", "\u2026");
            Add(map, "dagger", "\u2020");
            Add(map, "daggerdbl", "\u2021");
            Add(map, "perthousand", "\u2030");
            Add(map, "guilsinglleft", "\u2039");
            Add(map, "guilsinglright", "\u203A");
            Add(map, "guillemotleft", "\u00AB");
            Add(map, "guillemotright", "\u00BB");
            Add(map, "trademark", "\u2122");
            Add(map, "copyright", "\u00A9");
            Add(map, "registered", "\u00AE");
            Add(map, "degree", "\u00B0");
            Add(map, "plusminus", "\u00B1");
            Add(map, "multiply", "\u00D7");
            Add(map, "divide", "\u00F7");
            Add(map, "fi", "\uFB01");
            Add(map, "fl", "\uFB02");
            Add(map, "ff", "\uFB00");
            Add(map, "ffi", "\uFB03");
            Add(map, "ffl", "\uFB04");
            Add(map, "germandbls", "\u00DF");
            Add(map, "AE", "\u00C6");
            Add(map, "ae", "\u00E6");
            Add(map, "OE", "\u0152");
            Add(map, "oe", "\u0153");
            Add(map, "Oslash", "\u00D8");
            Add(map, "oslash", "\u00F8");
            Add(map, "Lslash", "\u0141");
            Add(map, "lslash", "\u0142");
            Add(map, "dotlessi", "\u0131");
            Add(map, "dotlessj", "\u0237");
            Add(map, "Euro", "\u20AC");
            Add(map, "currency", "\u00A4");
            Add(map, "brokenbar", "\u00A6");
            Add(map, "section", "\u00A7");
            Add(map, "paragraph", "\u00B6");
            Add(map, "periodcentered", "\u00B7");
            Add(map, "onesuperior", "\u00B9");
            Add(map, "twosuperior", "\u00B2");
            Add(map, "threesuperior", "\u00B3");
            Add(map, "onequarter", "\u00BC");
            Add(map, "onehalf", "\u00BD");
            Add(map, "threequarters", "\u00BE");
            Add(map, "mu", "\u00B5");
            Add(map, "nbspace", "\u00A0");
            Add(map, "sfthyphen", "\u00AD");
            Add(map, "apple", "\uF8FF");
            return map;
        }

        /// <summary>
        /// 向映射表写入一条字形名（存在则覆盖）。
        /// Writes one glyph name into the table, overwriting an existing entry.
        /// </summary>
        /// <param name="map">目标映射表。The target table.</param>
        /// <param name="name">字形名。The glyph name.</param>
        /// <param name="value">对应 Unicode 文本。The matching Unicode text.</param>
        private static void Add(Dictionary<string, string> map, string name, string value)
        {
            map[name] = value;
        }
    }
}
