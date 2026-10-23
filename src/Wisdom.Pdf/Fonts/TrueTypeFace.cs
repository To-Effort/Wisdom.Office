using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// 零依赖的 TrueType / TrueType 集合（TTC）解析器，支持字形子集化。仅支持 glyf 轮廓字体（不支持 CFF/OTF）。
    /// Zero-dependency TrueType / TrueType Collection parser with subsetting support. Only glyf-outline fonts are supported (not CFF/OTF).
    /// </summary>
    public sealed class TrueTypeFace
    {
        /// <summary>单个 sfnt 表的定位与校验信息。Location and checksum information for one sfnt table.</summary>
        private sealed class TableInfo
        {
            /// <summary>四字符表标签。The four-character table tag.</summary>
            public string Tag;
            /// <summary>表数据在文件中的字节偏移。The byte offset of the table data within the file.</summary>
            public int Offset;
            /// <summary>表数据的字节长度。The byte length of the table data.</summary>
            public int Length;
            /// <summary>表校验和。The table checksum.</summary>
            public uint Checksum;
        }

        /// <summary>字体文件原始字节。The raw font file bytes.</summary>
        private readonly byte[] _file;
        /// <summary>按标签索引的表目录。The table directory keyed by tag.</summary>
        private readonly Dictionary<string, TableInfo> _tables;
        /// <summary>解析后的 loca 偏移数组。The parsed loca offset array.</summary>
        private int[] _locaOffsets;
        /// <summary>字符到字形映射表。The character-to-glyph mapping table.</summary>
        private readonly CmapTable _cmap;

        /// <summary>
        /// 解析指定文件字节中给定面偏移处的字体面。
        /// Parses the font face at the given face offset within the supplied file bytes.
        /// </summary>
        /// <param name="file">字体文件或 TTC 的完整字节。The complete bytes of the font file or TTC.</param>
        /// <param name="faceOffset">该面在文件中的起始偏移。The start offset of this face within the file.</param>
        /// <exception cref="NotSupportedException">遇到 CFF/OTF 或不含 glyf 轮廓的字体时抛出。Thrown for CFF/OTF fonts or fonts lacking glyf outlines.</exception>
        /// <exception cref="InvalidDataException">sfnt 版本头不受支持时抛出。Thrown when the sfnt version header is unsupported.</exception>
        internal TrueTypeFace(byte[] file, int faceOffset)
        {
            _file = file;
            _tables = new Dictionary<string, TableInfo>(StringComparer.Ordinal);

            uint sfntVersion = ReadU32(faceOffset);
            if (sfntVersion == 0x4F54544F) // OTTO (CFF)
            {
                throw new NotSupportedException("CFF/OpenType fonts (OTTO) are not supported for PDF embedding.");
            }
            if (sfntVersion != 0x00010000 && sfntVersion != 0x74727565 && sfntVersion != 0x74797031)
            {
                throw new InvalidDataException("Unsupported TrueType sfnt version.");
            }

            int numTables = ReadU16(faceOffset + 4);
            for (int i = 0; i < numTables; i++)
            {
                int record = faceOffset + 12 + i * 16;
                if (record + 16 > file.Length) break;
                TableInfo info = new TableInfo();
                info.Tag = ReadTag(file, record);
                info.Checksum = ReadU32(record + 4);
                info.Offset = (int)ReadU32(record + 8);
                info.Length = (int)ReadU32(record + 12);
                // 用 long 相加：offset 与 length 都是无符号 32 位读入，int 相加会翻负而绕过校验。
                // Add in long: offset and length come from unsigned 32-bit reads, so int addition can wrap
                // negative and sail past the bound, later allocating or reading out of range.
                if (info.Offset < 0 || info.Length < 0 || (long)info.Offset + info.Length > file.Length)
                {
                    continue;
                }
                if (!_tables.ContainsKey(info.Tag))
                {
                    _tables.Add(info.Tag, info);
                }
            }

            if (!_tables.ContainsKey("glyf") || !_tables.ContainsKey("loca") || !_tables.ContainsKey("head")
                || !_tables.ContainsKey("hhea") || !_tables.ContainsKey("hmtx") || !_tables.ContainsKey("maxp"))
            {
                throw new NotSupportedException("Font does not contain TrueType glyf outlines.");
            }

            TableInfo head = _tables["head"];
            UnitsPerEm = ReadU16(head.Offset + 18);
            if (UnitsPerEm <= 0 || UnitsPerEm > 16384) UnitsPerEm = 2048;
            XMin = ReadS16(head.Offset + 36);
            YMin = ReadS16(head.Offset + 38);
            XMax = ReadS16(head.Offset + 40);
            YMax = ReadS16(head.Offset + 42);
            MacStyle = ReadU16(head.Offset + 44);
            IndexToLocFormat = ReadS16(head.Offset + 50);

            TableInfo hhea = _tables["hhea"];
            Ascender = ReadS16(hhea.Offset + 4);
            Descender = ReadS16(hhea.Offset + 6);
            LineGap = ReadS16(hhea.Offset + 8);
            NumberOfHMetrics = ReadU16(hhea.Offset + 34);

            TableInfo maxp = _tables["maxp"];
            GlyphCount = ReadU16(maxp.Offset + 4);

            TableInfo os2;
            if (_tables.TryGetValue("OS/2", out os2))
            {
                ushort os2Version = ReadU16(os2.Offset);
                ushort fsSelection = ReadU16(os2.Offset + 62);
                IsBold = (fsSelection & 0x20) != 0;
                IsItalic = (fsSelection & 0x01) != 0;
                int typoAscender = ReadS16(os2.Offset + 68);
                int typoDescender = ReadS16(os2.Offset + 70);
                int typoLineGap = ReadS16(os2.Offset + 72);
                if (typoAscender != 0 || typoDescender != 0)
                {
                    Ascender = typoAscender;
                    Descender = typoDescender;
                    LineGap = typoLineGap;
                }
                CapHeight = 0;
                XHeight = 0;
                if (os2Version >= 2 && os2.Offset + 90 <= os2.Offset + os2.Length)
                {
                    XHeight = ReadS16(os2.Offset + 86);
                    CapHeight = ReadS16(os2.Offset + 88);
                }
                try
                {
                    ushort fsType = ReadU16(os2.Offset + 8);
                    EmbeddingRestricted = (fsType & 0x0002) != 0;
                }
                catch (Exception)
                {
                }
            }

            if (!IsBold && (MacStyle & 0x01) != 0) IsBold = true;
            if (!IsItalic && (MacStyle & 0x02) != 0) IsItalic = true;
            if (Ascender == 0) Ascender = (int)(UnitsPerEm * 0.8);
            if (Descender == 0) Descender = -(int)(UnitsPerEm * 0.2);
            if (CapHeight == 0) CapHeight = (int)(Ascender * 0.7);
            if (XHeight == 0) XHeight = (int)(Ascender * 0.5);

            TableInfo post;
            if (_tables.TryGetValue("post", out post) && post.Length >= 8)
            {
                int rawAngle = ReadS16(post.Offset + 4);
                int fraction = ReadU16(post.Offset + 6);
                ItalicAngle = rawAngle + fraction / 65536.0;
                if (!IsItalic && ItalicAngle != 0) IsItalic = true;
            }

            ParseNames(file);
            ParseLoca();
            _cmap = CmapTable.Load(file, _tables);
        }

        /// <summary>家族名称。The family name.</summary>
        public string FamilyName { get; private set; }
        /// <summary>子族名称。The subfamily name.</summary>
        public string SubfamilyName { get; private set; }
        /// <summary>完整名称。The full name.</summary>
        public string FullName { get; private set; }
        /// <summary>PostScript 名称。The PostScript name.</summary>
        public string PostScriptName { get; private set; }
        /// <summary>是否为粗体。Whether the face is bold.</summary>
        public bool IsBold { get; private set; }
        /// <summary>是否为斜体。Whether the face is italic.</summary>
        public bool IsItalic { get; private set; }
        /// <summary>是否受嵌入限制（OS/2 fsType 禁止嵌入）。Whether embedding is restricted (OS/2 fsType forbids it).</summary>
        public bool EmbeddingRestricted { get; private set; }
        /// <summary>每 em 单位数。The units-per-em value.</summary>
        public int UnitsPerEm { get; private set; }
        /// <summary>上升高度。The ascent metric.</summary>
        public int Ascender { get; private set; }
        /// <summary>下降高度。The descent metric.</summary>
        public int Descender { get; private set; }
        /// <summary>行距间隙。The line gap metric.</summary>
        public int LineGap { get; private set; }
        /// <summary>大写字母高度。The cap-height metric.</summary>
        public int CapHeight { get; private set; }
        /// <summary>x 高度。The x-height metric.</summary>
        public int XHeight { get; private set; }
        /// <summary>边界框左下 X。The bounding-box minimum X.</summary>
        public int XMin { get; private set; }
        /// <summary>边界框左下 Y。The bounding-box minimum Y.</summary>
        public int YMin { get; private set; }
        /// <summary>边界框右上 X。The bounding-box maximum X.</summary>
        public int XMax { get; private set; }
        /// <summary>边界框右上 Y。The bounding-box maximum Y.</summary>
        public int YMax { get; private set; }
        /// <summary>字形总数。The total glyph count.</summary>
        public int GlyphCount { get; private set; }
        /// <summary>水平度量条目数。The number of horizontal-metric entries.</summary>
        public int NumberOfHMetrics { get; private set; }
        /// <summary>loca 索引格式（0 短、1 长）。The loca index format (0 short, 1 long).</summary>
        public int IndexToLocFormat { get; private set; }
        /// <summary>斜体角度。The italic angle.</summary>
        public double ItalicAngle { get; private set; }
        /// <summary>head 表中的 MacStyle 位标志。The MacStyle bit flags from the head table.</summary>
        private ushort MacStyle { get; set; }

        /// <summary>
        /// 返回文件中包含的字体面数量（TTC）或 1（单字体）。
        /// Returns the number of faces in the file (TTC) or 1 for a single font.
        /// </summary>
        /// <param name="file">字体或 TTC 字节。The font or TTC bytes.</param>
        /// <returns>面数量；无法识别时返回 0。The face count; 0 when unrecognized.</returns>
        public static int GetFaceCount(byte[] file)
        {
            if (file == null || file.Length < 12) return 0;
            if (ReadTag(file, 0) != "ttcf") return 1;
            uint count = ReadU32(file, 8);
            if (count > 1000) return 0;
            return (int)count;
        }

        /// <summary>
        /// 返回指定面索引在 TTC 文件中的起始偏移（单字体返回 0）。
        /// Returns the start offset of the given face index in a TTC file (0 for a single font).
        /// </summary>
        /// <param name="file">字体或 TTC 字节。The font or TTC bytes.</param>
        /// <param name="faceIndex">从 0 开始的面索引。The zero-based face index.</param>
        /// <returns>该面的字节偏移。The byte offset of the face.</returns>
        /// <exception cref="InvalidDataException">文件过小无法解析头部时抛出。Thrown when the file is too small to parse the header.</exception>
        /// <exception cref="ArgumentOutOfRangeException">面索引越界时抛出。Thrown when the face index is out of range.</exception>
        public static int GetFaceOffset(byte[] file, int faceIndex)
        {
            if (file == null || file.Length < 12) throw new InvalidDataException("Font file is too small.");
            if (ReadTag(file, 0) != "ttcf") return 0;
            uint count = ReadU32(file, 8);
            if (faceIndex < 0 || faceIndex >= count) throw new ArgumentOutOfRangeException("faceIndex");
            return (int)ReadU32(file, 12 + faceIndex * 4);
        }

        /// <summary>
        /// 通过 cmap 将码位映射为字形编号。
        /// Maps a code point to a glyph id through the cmap.
        /// </summary>
        /// <param name="codePoint">Unicode 码位。The Unicode code point.</param>
        /// <returns>字形编号；未命中返回 0。The glyph id; 0 when unmapped.</returns>
        public int GetGlyphId(int codePoint)
        {
            return _cmap.GetGlyphId(codePoint);
        }

        /// <summary>
        /// 读取指定字形的推进宽度（em 单位）。
        /// Reads the advance width (in em units) for the given glyph.
        /// </summary>
        /// <param name="glyphId">字形编号。The glyph id.</param>
        /// <returns>推进宽度；越界返回 0，超出度量范围返回默认半宽。The advance width; 0 when out of range, or a default half-width beyond the metrics.</returns>
        public int GetAdvanceWidth(int glyphId)
        {
            if (glyphId < 0 || glyphId >= GlyphCount) return 0;
            TableInfo hmtx = _tables["hmtx"];
            int metrics = NumberOfHMetrics > 0 ? NumberOfHMetrics : 1;
            int index = glyphId < metrics ? glyphId : metrics - 1;
            int offset = hmtx.Offset + index * 4;
            if (offset + 2 > hmtx.Offset + hmtx.Length) return UnitsPerEm / 2;
            return ReadU16(offset);
        }

        /// <summary>
        /// 读取指定字形的左侧游标（LSB）。
        /// Reads the left side bearing (LSB) for the given glyph.
        /// </summary>
        /// <param name="glyphId">字形编号。The glyph id.</param>
        /// <returns>左侧游标值；越界返回 0。The LSB value; 0 when out of range.</returns>
        public int GetLeftSideBearing(int glyphId)
        {
            if (glyphId < 0 || glyphId >= GlyphCount) return 0;
            TableInfo hmtx = _tables["hmtx"];
            int metrics = NumberOfHMetrics > 0 ? NumberOfHMetrics : 1;
            if (glyphId < metrics)
            {
                return ReadS16(hmtx.Offset + glyphId * 4 + 2);
            }
            int offset = hmtx.Offset + metrics * 4 + (glyphId - metrics) * 2;
            if (offset + 2 > hmtx.Offset + hmtx.Length) return 0;
            return ReadS16(offset);
        }

        /// <summary>
        /// 判断该面是否覆盖指定码位（能映射到非零字形）。
        /// Determines whether the face covers the given code point (maps to a non-zero glyph).
        /// </summary>
        /// <param name="codePoint">Unicode 码位。The Unicode code point.</param>
        /// <returns>有字形返回 true。True when a glyph exists.</returns>
        public bool HasGlyph(int codePoint)
        {
            return GetGlyphId(codePoint) != 0;
        }

        /// <summary>子集化构建结果。The result of building a subset.</summary>
        public sealed class SubsetResult
        {
            /// <summary>子集字体文件的完整字节。The complete bytes of the subset font file.</summary>
            public byte[] Data;
            /// <summary>子集中按新顺序排列的原始字形编号。The old glyph ids in their new subset order.</summary>
            public List<int> OldGlyphIds;
            /// <summary>原始字形编号到子集字形编号的映射。The mapping from old glyph ids to subset glyph ids.</summary>
            public Dictionary<int, int> OldToNewGlyphId;
        }

        /// <summary>
        /// 依据请求的字形集合构建仅含所需字形的子集字体。
        /// Builds a subset font containing only the glyphs reachable from the requested set.
        /// </summary>
        /// <param name="requestedGlyphs">请求包含的原始字形编号集合，可为 null。The set of requested old glyph ids; may be null.</param>
        /// <returns>含子集字节及字形重映射编号的结果对象。The result carrying the subset bytes and glyph remapping ids.</returns>
        /// <remarks>先把 glyph .notdef（0 号）与每个请求字形入队，再沿复合字形的组件闭包展开，
        /// 重写组件引用为新编号，并重建 head/hhea/maxp/hmtx/loca/glyf 及可选表，最后打包为合法 sfnt。
        /// It seeds glyph .notdef (id 0) and each requested glyph, expands the composite-glyph component closure, rewrites component references to new ids, rebuilds head/hhea/maxp/hmtx/loca/glyf plus optional tables, and packs a valid sfnt.</remarks>
        public SubsetResult BuildSubset(ICollection<int> requestedGlyphs)
        {
            bool[] include = new bool[GlyphCount];
            include[0] = true;
            Queue<int> queue = new Queue<int>();
            if (requestedGlyphs != null)
            {
                foreach (int gid in requestedGlyphs)
                {
                    if (gid > 0 && gid < GlyphCount && !include[gid])
                    {
                        include[gid] = true;
                        queue.Enqueue(gid);
                    }
                }
            }

            while (queue.Count > 0)
            {
                int gid = queue.Dequeue();
                byte[] glyph = GetGlyphData(gid);
                if (glyph.Length == 0) continue;
                List<int> components = GetComponentGlyphIds(glyph);
                for (int i = 0; i < components.Count; i++)
                {
                    int component = components[i];
                    if (component > 0 && component < GlyphCount && !include[component])
                    {
                        include[component] = true;
                        queue.Enqueue(component);
                    }
                }
            }

            List<int> newOrder = new List<int>();
            for (int gid = 0; gid < GlyphCount; gid++)
            {
                if (include[gid])
                {
                    newOrder.Add(gid);
                }
            }

            Dictionary<int, int> oldToNew = new Dictionary<int, int>();
            for (int i = 0; i < newOrder.Count; i++)
            {
                oldToNew[newOrder[i]] = i;
            }

            MemoryStream glyf = new MemoryStream();
            List<uint> loca = new List<uint>();
            for (int i = 0; i < newOrder.Count; i++)
            {
                loca.Add((uint)glyf.Length);
                byte[] glyph = GetGlyphData(newOrder[i]);
                if (glyph.Length == 0) continue;

                if (ReadS16(glyph, 0) < 0)
                {
                    glyph = (byte[])glyph.Clone();
                    List<int> positions = GetComponentGlyphIdPositions(glyph);
                    for (int p = 0; p < positions.Count; p++)
                    {
                        int position = positions[p];
                        if (position + 2 > glyph.Length) continue;
                        int componentOld = (glyph[position] << 8) | glyph[position + 1];
                        int componentNew;
                        if (!oldToNew.TryGetValue(componentOld, out componentNew))
                        {
                            componentNew = 0;
                        }
                        glyph[position] = (byte)(componentNew >> 8);
                        glyph[position + 1] = (byte)componentNew;
                    }
                }

                glyf.Write(glyph, 0, glyph.Length);
                while ((glyf.Length & 3) != 0) glyf.WriteByte(0);
            }
            loca.Add((uint)glyf.Length);

            byte[] head = GetTableCopy("head");
            WriteU32(head, 8, 0);
            WriteU16(head, 50, 1); // long loca offsets

            byte[] hhea = GetTableCopy("hhea");
            WriteU16(hhea, 34, (ushort)newOrder.Count);

            byte[] maxp = GetTableCopy("maxp");
            WriteU16(maxp, 4, (ushort)newOrder.Count);

            byte[] hmtx = new byte[newOrder.Count * 4];
            for (int i = 0; i < newOrder.Count; i++)
            {
                int oldGid = newOrder[i];
                WriteU16(hmtx, i * 4, (ushort)GetAdvanceWidth(oldGid));
                WriteS16(hmtx, i * 4 + 2, (short)GetLeftSideBearing(oldGid));
            }

            byte[] locaBytes = new byte[loca.Count * 4];
            for (int i = 0; i < loca.Count; i++)
            {
                WriteU32(locaBytes, i * 4, loca[i]);
            }

            Dictionary<string, byte[]> outputTables = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            outputTables["head"] = head;
            outputTables["hhea"] = hhea;
            outputTables["maxp"] = maxp;
            outputTables["hmtx"] = hmtx;
            outputTables["loca"] = locaBytes;
            outputTables["glyf"] = glyf.ToArray();

            string[] optional = new string[] { "cvt ", "fpgm", "prep", "gasp", "OS/2", "name", "post", "VDMX" };
            for (int i = 0; i < optional.Length; i++)
            {
                byte[] table = GetTableCopy(optional[i]);
                if (table != null)
                {
                    outputTables[optional[i]] = table;
                }
            }

            byte[] subsetFont = BuildSfnt(outputTables);
            SubsetResult result = new SubsetResult();
            result.Data = subsetFont;
            result.OldGlyphIds = newOrder;
            result.OldToNewGlyphId = oldToNew;
            return result;
        }

        /// <summary>
        /// 解析 name 表填充家族、子族、全称与 PostScript 名称，并据子族文本补判粗体/斜体。
        /// Parses the name table to fill family, subfamily, full and PostScript names, and infers bold/italic from subfamily text.
        /// </summary>
        /// <param name="file">字体文件字节。The font file bytes.</param>
        private void ParseNames(byte[] file)
        {
            FamilyName = "Unknown";
            SubfamilyName = "Regular";
            FullName = "Unknown";
            PostScriptName = "UnknownFont";
            TableInfo name;
            if (!_tables.TryGetValue("name", out name))
            {
                return;
            }

            int count = ReadU16(name.Offset + 2);
            int stringOffset = ReadU16(name.Offset + 4);
            string family = null;
            string subfamily = null;
            string full = null;
            string postScript = null;
            string typographicFamily = null;
            string typographicSubfamily = null;

            for (int i = 0; i < count; i++)
            {
                int record = name.Offset + 6 + i * 12;
                if (record + 12 > name.Offset + name.Length) break;
                int platform = ReadU16(record);
                int encoding = ReadU16(record + 2);
                int language = ReadU16(record + 4);
                int nameId = ReadU16(record + 6);
                int length = ReadU16(record + 8);
                int offset = ReadU16(record + 10);
                int start = name.Offset + stringOffset + offset;
                if (start < 0 || start + length > name.Offset + name.Length) continue;

                bool preferred = platform == 3 && (encoding == 1 || encoding == 10);
                bool fallback = platform == 0 || platform == 1;
                if (!preferred && !fallback) continue;
                if (preferred && language != 0x0409 && language != 0) continue;

                string value = preferred || platform == 0
                    ? ReadUtf16BigEndian(file, start, length)
                    : ReadLatin1(file, start, length);
                if (string.IsNullOrEmpty(value)) continue;

                if (nameId == 16 && (typographicFamily == null || preferred)) typographicFamily = value;
                else if (nameId == 17 && (typographicSubfamily == null || preferred)) typographicSubfamily = value;
                else if (nameId == 1 && (family == null || preferred)) family = value;
                else if (nameId == 2 && (subfamily == null || preferred)) subfamily = value;
                else if (nameId == 4 && (full == null || preferred)) full = value;
                else if (nameId == 6 && (postScript == null || preferred)) postScript = value;
            }

            if (!string.IsNullOrEmpty(typographicFamily)) family = typographicFamily;
            if (!string.IsNullOrEmpty(typographicSubfamily)) subfamily = typographicSubfamily;
            if (!string.IsNullOrEmpty(family)) FamilyName = family.Trim();
            if (!string.IsNullOrEmpty(subfamily)) SubfamilyName = subfamily.Trim();
            if (!string.IsNullOrEmpty(full)) FullName = full.Trim(); else FullName = FamilyName + " " + SubfamilyName;
            if (!string.IsNullOrEmpty(postScript)) PostScriptName = SanitizePostScriptName(postScript);
            else PostScriptName = SanitizePostScriptName(FamilyName.Replace(" ", string.Empty));

            if (SubfamilyName.IndexOf("Bold", StringComparison.OrdinalIgnoreCase) >= 0) IsBold = true;
            if (SubfamilyName.IndexOf("Italic", StringComparison.OrdinalIgnoreCase) >= 0
                || SubfamilyName.IndexOf("Oblique", StringComparison.OrdinalIgnoreCase) >= 0) IsItalic = true;
        }

        /// <summary>
        /// 把名称清洗为合法的 PostScript 名（仅字母数字、连字符、下划线，最长 63）。
        /// Sanitizes a name into a valid PostScript name (letters, digits, hyphen, underscore, at most 63 characters).
        /// </summary>
        /// <param name="value">原始名称。The raw name.</param>
        /// <returns>清洗后的名称；为空时返回 WisdomFont。The sanitized name, or WisdomFont when empty.</returns>
        private static string SanitizePostScriptName(string value)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < value.Length && builder.Length < 63; i++)
            {
                char c = value[i];
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '-' || c == '_')
                {
                    builder.Append(c);
                }
            }
            return builder.Length == 0 ? "WisdomFont" : builder.ToString();
        }

        /// <summary>
        /// 依据 head 的索引格式解析 loca 表为字形偏移数组。
        /// Parses the loca table into a glyph-offset array using the head index format.
        /// </summary>
        private void ParseLoca()
        {
            TableInfo loca = _tables["loca"];
            int count = GlyphCount + 1;
            _locaOffsets = new int[count];
            if (IndexToLocFormat == 0)
            {
                int max = Math.Min(count, loca.Length / 2);
                for (int i = 0; i < max; i++)
                {
                    _locaOffsets[i] = ReadU16(loca.Offset + i * 2) * 2;
                }
            }
            else
            {
                int max = Math.Min(count, loca.Length / 4);
                for (int i = 0; i < max; i++)
                {
                    _locaOffsets[i] = (int)ReadU32(loca.Offset + i * 4);
                }
            }
        }

        /// <summary>
        /// 从 glyf 表取出指定字形的原始字节。
        /// Extracts the raw bytes of the given glyph from the glyf table.
        /// </summary>
        /// <param name="glyphId">字形编号。The glyph id.</param>
        /// <returns>字形数据；越界或空字形返回长度 0 的数组。The glyph data; a zero-length array when out of range or empty.</returns>
        public byte[] GetGlyphData(int glyphId)
        {
            if (glyphId < 0 || glyphId + 1 >= _locaOffsets.Length) return new byte[0];
            int start = _locaOffsets[glyphId];
            int end = _locaOffsets[glyphId + 1];
            TableInfo glyf = _tables["glyf"];
            if (start < 0 || end <= start || start >= glyf.Length) return new byte[0];
            if (end > glyf.Length) end = glyf.Length;
            byte[] result = new byte[end - start];
            Buffer.BlockCopy(_file, glyf.Offset + start, result, 0, result.Length);
            return result;
        }

        /// <summary>
        /// 解析复合字形引用的组件字形编号列表。
        /// Parses the list of component glyph ids referenced by a composite glyph.
        /// </summary>
        /// <param name="glyph">字形数据。The glyph data.</param>
        /// <returns>组件字形编号；简单字形返回空列表。The component ids; empty for a simple glyph.</returns>
        private static List<int> GetComponentGlyphIds(byte[] glyph)
        {
            List<int> positions = GetComponentGlyphIdPositions(glyph);
            List<int> ids = new List<int>(positions.Count);
            for (int i = 0; i < positions.Count; i++)
            {
                int p = positions[i];
                if (p + 2 <= glyph.Length)
                {
                    ids.Add((glyph[p] << 8) | glyph[p + 1]);
                }
            }
            return ids;
        }

        /// <summary>
        /// 按复合字形组件标志逐段扫描，返回每个组件字形编号字段的字节位置。
        /// Walks the composite-glyph component flags and returns the byte position of each component glyph-id field.
        /// </summary>
        /// <param name="glyph">字形数据。The glyph data.</param>
        /// <returns>组件编号字段位置列表；非复合字形返回空列表。The list of positions; empty for a non-composite glyph.</returns>
        private static List<int> GetComponentGlyphIdPositions(byte[] glyph)
        {
            List<int> positions = new List<int>();
            if (glyph.Length < 10 || ReadS16(glyph, 0) >= 0) return positions;

            int position = 10;
            int guard = 0;
            while (position + 4 <= glyph.Length && guard++ < 256)
            {
                ushort flags = (ushort)((glyph[position] << 8) | glyph[position + 1]);
                positions.Add(position + 2);
                position += 4;
                position += (flags & 0x0001) != 0 ? 4 : 2;
                if ((flags & 0x0008) != 0) position += 2;
                else if ((flags & 0x0040) != 0) position += 4;
                else if ((flags & 0x0080) != 0) position += 8;
                if ((flags & 0x0020) == 0)
                {
                    break;
                }
            }
            return positions;
        }

        /// <summary>
        /// 复制指定标签表的原始字节。
        /// Copies the raw bytes of the table with the given tag.
        /// </summary>
        /// <param name="tag">四字符表标签。The four-character table tag.</param>
        /// <returns>表数据副本；标签不存在返回 null。A copy of the table data; null when the tag is absent.</returns>
        private byte[] GetTableCopy(string tag)
        {
            TableInfo info;
            if (!_tables.TryGetValue(tag, out info)) return null;
            byte[] result = new byte[info.Length];
            Buffer.BlockCopy(_file, info.Offset, result, 0, info.Length);
            return result;
        }

        /// <summary>
        /// 将若干表打包成合法的 sfnt 容器（含排序的表目录与 4 字节对齐）。
        /// Packs the given tables into a valid sfnt container with a sorted, 4-byte-aligned table directory.
        /// </summary>
        /// <param name="tables">标签到表数据的映射。The map from tag to table data.</param>
        /// <returns>完整的 sfnt 字体字节。The complete sfnt font bytes.</returns>
        private static byte[] BuildSfnt(Dictionary<string, byte[]> tables)
        {
            List<string> tags = new List<string>(tables.Keys);
            tags.Sort(StringComparer.Ordinal);
            int tableCount = tags.Count;
            int headerSize = 12 + tableCount * 16;
            int totalSize = headerSize;
            for (int i = 0; i < tableCount; i++)
            {
                totalSize += (tables[tags[i]].Length + 3) & ~3;
            }

            byte[] output = new byte[totalSize];
            WriteU32(output, 0, 0x00010000);
            WriteU16(output, 4, (ushort)tableCount);
            int searchRange = 1;
            int entrySelector = 0;
            while ((searchRange * 2) <= tableCount)
            {
                searchRange *= 2;
                entrySelector++;
            }
            searchRange *= 16;
            WriteU16(output, 6, (ushort)searchRange);
            WriteU16(output, 8, (ushort)entrySelector);
            WriteU16(output, 10, (ushort)(tableCount * 16 - searchRange));

            int offset = headerSize;
            for (int i = 0; i < tableCount; i++)
            {
                byte[] data = tables[tags[i]];
                int record = 12 + i * 16;
                WriteTag(output, record, tags[i]);
                WriteU32(output, record + 4, ComputeChecksum(data));
                WriteU32(output, record + 8, (uint)offset);
                WriteU32(output, record + 12, (uint)data.Length);
                Buffer.BlockCopy(data, 0, output, offset, data.Length);
                offset += (data.Length + 3) & ~3;
            }
            return output;
        }

        /// <summary>
        /// 计算 sfnt 表的大端 32 位校验和（末尾不足 4 字节按补零处理）。
        /// Computes the big-endian 32-bit sfnt table checksum (a trailing partial word is zero-padded).
        /// </summary>
        /// <param name="data">表数据。The table data.</param>
        /// <returns>校验和。The checksum.</returns>
        private static uint ComputeChecksum(byte[] data)
        {
            uint sum = 0;
            int i = 0;
            for (; i + 3 < data.Length; i += 4)
            {
                sum = unchecked(sum + ((uint)data[i] << 24 | (uint)data[i + 1] << 16 | (uint)data[i + 2] << 8 | data[i + 3]));
            }
            if (i < data.Length)
            {
                uint last = (uint)data[i] << 24;
                if (i + 1 < data.Length) last |= (uint)data[i + 1] << 16;
                if (i + 2 < data.Length) last |= (uint)data[i + 2] << 8;
                sum = unchecked(sum + last);
            }
            return sum;
        }

        /// <summary>
        /// 读取大端 UTF-16 字符串。
        /// Reads a big-endian UTF-16 string.
        /// </summary>
        private static string ReadUtf16BigEndian(byte[] data, int offset, int length)
        {
            if (length < 2) return string.Empty;
            return Encoding.BigEndianUnicode.GetString(data, offset, length);
        }

        /// <summary>
        /// 读取 Latin-1（逐字节当作字符）字符串。
        /// Reads a Latin-1 string (each byte mapped directly to a char).
        /// </summary>
        private static string ReadLatin1(byte[] data, int offset, int length)
        {
            char[] chars = new char[length];
            for (int i = 0; i < length; i++)
            {
                chars[i] = (char)data[offset + i];
            }
            return new string(chars);
        }

        /// <summary>
        /// 读取 4 字节标签字符串。
        /// Reads a 4-byte tag string.
        /// </summary>
        private static string ReadTag(byte[] data, int offset)
        {
            if (offset + 4 > data.Length) return string.Empty;
            return new string(new[]
            {
                (char)data[offset], (char)data[offset + 1], (char)data[offset + 2], (char)data[offset + 3]
            });
        }

        /// <summary>
        /// 写入 4 字节标签，短于 4 字符用空格补齐。
        /// Writes a 4-byte tag, space-padding tags shorter than four characters.
        /// </summary>
        private static void WriteTag(byte[] data, int offset, string tag)
        {
            for (int i = 0; i < 4; i++)
            {
                data[offset + i] = i < tag.Length ? (byte)tag[i] : (byte)0x20;
            }
        }

        /// <summary>
        /// 从指定缓冲的大端偏移读取 16 位无符号整数。
        /// Reads a big-endian 16-bit unsigned integer at the offset in the given buffer.
        /// </summary>
        private static ushort ReadU16(byte[] data, int offset)
        {
            if (offset < 0 || offset + 2 > data.Length) return 0;
            return (ushort)((data[offset] << 8) | data[offset + 1]);
        }

        /// <summary>
        /// 从内部文件缓冲的大端偏移读取 16 位无符号整数。
        /// Reads a big-endian 16-bit unsigned integer at the offset in the internal file buffer.
        /// </summary>
        private ushort ReadU16(int offset)
        {
            if (offset < 0 || offset + 2 > _file.Length) return 0;
            return (ushort)((_file[offset] << 8) | _file[offset + 1]);
        }

        /// <summary>
        /// 从内部文件缓冲读取 16 位有符号整数。
        /// Reads a big-endian 16-bit signed integer from the internal file buffer.
        /// </summary>
        private short ReadS16(int offset)
        {
            return unchecked((short)ReadU16(offset));
        }

        /// <summary>
        /// 从指定缓冲读取 16 位有符号整数。
        /// Reads a big-endian 16-bit signed integer from the given buffer.
        /// </summary>
        private static short ReadS16(byte[] data, int offset)
        {
            return unchecked((short)ReadU16(data, offset));
        }

        /// <summary>
        /// 从内部文件缓冲读取 32 位无符号整数。
        /// Reads a big-endian 32-bit unsigned integer from the internal file buffer.
        /// </summary>
        private uint ReadU32(int offset)
        {
            if (offset < 0 || offset + 4 > _file.Length) return 0;
            return (uint)((_file[offset] << 24) | (_file[offset + 1] << 16) | (_file[offset + 2] << 8) | _file[offset + 3]);
        }

        /// <summary>
        /// 从指定缓冲读取 32 位无符号整数。
        /// Reads a big-endian 32-bit unsigned integer from the given buffer.
        /// </summary>
        private static uint ReadU32(byte[] data, int offset)
        {
            if (offset < 0 || offset + 4 > data.Length) return 0;
            return (uint)((data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]);
        }

        /// <summary>
        /// 以大端序写入 16 位无符号整数。
        /// Writes a 16-bit unsigned integer in big-endian order.
        /// </summary>
        private static void WriteU16(byte[] data, int offset, ushort value)
        {
            if (offset < 0 || offset + 2 > data.Length) return;
            data[offset] = (byte)(value >> 8);
            data[offset + 1] = (byte)value;
        }

        /// <summary>
        /// 以大端序写入 16 位有符号整数。
        /// Writes a 16-bit signed integer in big-endian order.
        /// </summary>
        private static void WriteS16(byte[] data, int offset, short value)
        {
            WriteU16(data, offset, unchecked((ushort)value));
        }

        /// <summary>
        /// 以大端序写入 32 位无符号整数。
        /// Writes a 32-bit unsigned integer in big-endian order.
        /// </summary>
        private static void WriteU32(byte[] data, int offset, uint value)
        {
            if (offset < 0 || offset + 4 > data.Length) return;
            data[offset] = (byte)(value >> 24);
            data[offset + 1] = (byte)(value >> 16);
            data[offset + 2] = (byte)(value >> 8);
            data[offset + 3] = (byte)value;
        }

        /// <summary>
        /// cmap 表的惰性解析缓存，支持格式 0/4/6/12 的码位到字形映射。
        /// Lazy-parsed cmap cache supporting code-point-to-glyph mapping for formats 0, 4, 6 and 12.
        /// </summary>
        private sealed class CmapTable
        {
            /// <summary>选中的子表格式，-1 表示无可用表。The chosen subtable format; -1 when no usable table exists.</summary>
            private int _format;
            /// <summary>格式 4 的每段结束码。Format 4 per-segment end codes.</summary>
            private int[] _endCodes;
            /// <summary>格式 4 的每段起始码。Format 4 per-segment start codes.</summary>
            private int[] _startCodes;
            /// <summary>格式 4 的每段 idDelta。Format 4 per-segment idDelta values.</summary>
            private int[] _idDelta;
            /// <summary>格式 4 的每段 idRangeOffset。Format 4 per-segment idRangeOffset values.</summary>
            private int[] _idRangeOffset;
            /// <summary>格式 4 的 idRangeOffset 基址。The idRangeOffset base address for format 4.</summary>
            private int _idRangeOffsetBase;
            /// <summary>格式 12 的每段起始码。Format 12 per-group start codes.</summary>
            private int[] _groupsStarts;
            /// <summary>格式 12 的每段结束码。Format 12 per-group end codes.</summary>
            private int[] _groupsEnds;
            /// <summary>格式 12 的每段起始字形。Format 12 per-group start glyph ids.</summary>
            private int[] _groupsGlyphs;
            /// <summary>格式 6 的起始码。The start code for format 6.</summary>
            private int _firstCode;
            /// <summary>格式 0/6 的字形数组。The glyph array for formats 0 and 6.</summary>
            private int[] _glyphArray;
            /// <summary>所依附的文件字节。The file bytes this table reads from.</summary>
            private byte[] _file;
            /// <summary>格式 0 子表的数据偏移。The data offset of the format 0 subtable.</summary>
            private int _format0Offset;

            /// <summary>
            /// 加载 cmap 并为各子表评分，选出覆盖最广、优先级最高者解析。
            /// Loads the cmap, scores each subtable, and parses the best-covered, highest-priority one.
            /// </summary>
            /// <param name="file">字体文件字节。The font file bytes.</param>
            /// <param name="tables">表目录。The table directory.</param>
            /// <returns>解析好的 cmap；缺失或无合适子表时返回格式为 -1 的空表。The parsed cmap; an empty table with format -1 when the cmap or a suitable subtable is missing.</returns>
            public static CmapTable Load(byte[] file, Dictionary<string, TableInfo> tables)
            {
                TableInfo cmap;
                if (!tables.TryGetValue("cmap", out cmap))
                {
                    return new CmapTable { _format = -1 };
                }

                int bestScore = -1;
                int bestSubtable = -1;
                int bestFormat = -1;
                int count = ReadU16Static(file, cmap.Offset + 2);
                for (int i = 0; i < count; i++)
                {
                    int record = cmap.Offset + 4 + i * 8;
                    if (record + 8 > cmap.Offset + cmap.Length) break;
                    int platform = ReadU16Static(file, record);
                    int encoding = ReadU16Static(file, record + 2);
                    int offset = cmap.Offset + (int)ReadU32Static(file, record + 4);
                    if (offset + 2 > file.Length) continue;
                    int format = ReadU16Static(file, offset);
                    int score = 0;
                    if (format == 12) score += 100;
                    else if (format == 4) score += 60;
                    else if (format == 6) score += 20;
                    else if (format == 0) score += 10;
                    else continue;

                    if (platform == 3 && encoding == 10) score += 50;
                    else if (platform == 3 && encoding == 1) score += 40;
                    else if (platform == 0 && (encoding == 4 || encoding == 6)) score += 45;
                    else if (platform == 0) score += 30;
                    else if (platform == 3) score += 20;

                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestSubtable = offset;
                        bestFormat = format;
                    }
                }

                if (bestSubtable < 0)
                {
                    return new CmapTable { _format = -1 };
                }

                CmapTable table = new CmapTable();
                table._file = file;
                table._format = bestFormat;
                if (bestFormat == 4)
                {
                    table.ParseFormat4(bestSubtable);
                }
                else if (bestFormat == 12)
                {
                    table.ParseFormat12(bestSubtable);
                }
                else if (bestFormat == 6)
                {
                    table._firstCode = ReadU16Static(file, bestSubtable + 6);
                    int entryCount = ReadU16Static(file, bestSubtable + 8);
                    table._glyphArray = new int[entryCount];
                    for (int i = 0; i < entryCount; i++)
                    {
                        table._glyphArray[i] = ReadU16Static(file, bestSubtable + 10 + i * 2);
                    }
                }
                else if (bestFormat == 0)
                {
                    table._format0Offset = bestSubtable + 6;
                    table._glyphArray = new int[256];
                    for (int i = 0; i < 256; i++)
                    {
                        table._glyphArray[i] = file[bestSubtable + 6 + i];
                    }
                }
                return table;
            }

            /// <summary>
            /// 解析格式 4 分段映射表（含符号化 idDelta 与 idRangeOffset 基址）。
            /// Parses the format 4 segmented mapping (including the signed idDelta values and the idRangeOffset base).
            /// </summary>
            /// <param name="offset">子表起始偏移。The subtable start offset.</param>
            private void ParseFormat4(int offset)
            {
                int segCountX2 = ReadU16Static(_file, offset + 6);
                int segCount = segCountX2 / 2;
                _endCodes = new int[segCount];
                _startCodes = new int[segCount];
                _idDelta = new int[segCount];
                _idRangeOffset = new int[segCount];
                int endBase = offset + 14;
                int startBase = endBase + segCountX2 + 2;
                int deltaBase = startBase + segCountX2;
                int rangeBase = deltaBase + segCountX2;
                _idRangeOffsetBase = rangeBase;
                for (int i = 0; i < segCount; i++)
                {
                    _endCodes[i] = ReadU16Static(_file, endBase + i * 2);
                    _startCodes[i] = ReadU16Static(_file, startBase + i * 2);
                    int delta = ReadU16Static(_file, deltaBase + i * 2);
                    _idDelta[i] = delta >= 0x8000 ? delta - 0x10000 : delta;
                    _idRangeOffset[i] = ReadU16Static(_file, rangeBase + i * 2);
                }
            }

            /// <summary>
            /// 解析格式 12 连续映射段（组数上限 20 万以防越界）。
            /// Parses the format 12 sequential-map groups (group count capped at 200,000 to guard against overruns).
            /// </summary>
            /// <param name="offset">子表起始偏移。The subtable start offset.</param>
            private void ParseFormat12(int offset)
            {
                int groupCount = (int)ReadU32Static(_file, offset + 12);
                if (groupCount > 200000) groupCount = 200000;
                _groupsStarts = new int[groupCount];
                _groupsEnds = new int[groupCount];
                _groupsGlyphs = new int[groupCount];
                for (int i = 0; i < groupCount; i++)
                {
                    int record = offset + 16 + i * 12;
                    _groupsStarts[i] = (int)ReadU32Static(_file, record);
                    _groupsEnds[i] = (int)ReadU32Static(_file, record + 4);
                    _groupsGlyphs[i] = (int)ReadU32Static(_file, record + 8);
                }
            }

            /// <summary>
            /// 将码位映射为字形编号，按已选格式走对应的二分或数组查找。
            /// Maps a code point to a glyph id, using the binary search or array lookup for the selected format.
            /// </summary>
            /// <param name="codePoint">Unicode 码位。The Unicode code point.</param>
            /// <returns>字形编号；未映射返回 0。The glyph id; 0 when unmapped.</returns>
            public int GetGlyphId(int codePoint)
            {
                if (codePoint < 0) return 0;
                if (_format == 4)
                {
                    if (codePoint > 0xFFFF) return 0;
                    int low = 0;
                    int high = _endCodes.Length - 1;
                    int segment = -1;
                    while (low <= high)
                    {
                        int mid = (low + high) / 2;
                        if (_endCodes[mid] >= codePoint)
                        {
                            segment = mid;
                            high = mid - 1;
                        }
                        else
                        {
                            low = mid + 1;
                        }
                    }
                    if (segment < 0 || codePoint < _startCodes[segment]) return 0;
                    if (_idRangeOffset[segment] == 0)
                    {
                        return (_idDelta[segment] + codePoint) & 0xFFFF;
                    }
                    int index = _idRangeOffsetBase + segment * 2 + _idRangeOffset[segment]
                        + (codePoint - _startCodes[segment]) * 2;
                    int glyph = ReadU16Static(_file, index);
                    if (glyph == 0) return 0;
                    return (glyph + _idDelta[segment]) & 0xFFFF;
                }
                if (_format == 12)
                {
                    int low = 0;
                    int high = _groupsStarts.Length - 1;
                    while (low <= high)
                    {
                        int mid = (low + high) / 2;
                        if (_groupsEnds[mid] < codePoint) low = mid + 1;
                        else if (_groupsStarts[mid] > codePoint) high = mid - 1;
                        else return _groupsGlyphs[mid] + (codePoint - _groupsStarts[mid]);
                    }
                    return 0;
                }
                if (_format == 6)
                {
                    int index = codePoint - _firstCode;
                    if (index < 0 || index >= _glyphArray.Length) return 0;
                    return _glyphArray[index];
                }
                if (_format == 0)
                {
                    if (codePoint < 0 || codePoint > 255) return 0;
                    return _glyphArray[codePoint];
                }
                return 0;
            }

            /// <summary>
            /// 从任意缓冲读取大端 16 位无符号整数。
            /// Reads a big-endian 16-bit unsigned integer from any buffer.
            /// </summary>
            private static ushort ReadU16Static(byte[] data, int offset)
            {
                if (offset < 0 || offset + 2 > data.Length) return 0;
                return (ushort)((data[offset] << 8) | data[offset + 1]);
            }

            /// <summary>
            /// 从任意缓冲读取大端 32 位无符号整数。
            /// Reads a big-endian 32-bit unsigned integer from any buffer.
            /// </summary>
            private static uint ReadU32Static(byte[] data, int offset)
            {
                if (offset < 0 || offset + 4 > data.Length) return 0;
                return (uint)((data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3]);
            }
        }
    }
}
