using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 读取旧版 Excel 97-2003（.xls）工作簿：BIFF8 记录封装在 OLE2 复合文件中。
    /// 单元格值、日期、数字格式、合并区域、行/列尺寸与打印设置会被导入统一工作簿模型。
    /// Reads legacy Excel 97-2003 (.xls) workbooks, where BIFF8 records are wrapped in an OLE2 compound file.
    /// Cell values, dates, number formats, merged ranges, row/column sizes and print settings are imported
    /// into the unified workbook model.
    /// </summary>
    public static class XlsReader
    {
        /// <summary>从文件路径加载 .xls 工作簿（包含全部隐藏内容）。</summary>
        /// <param name="path">.xls 文件路径。</param>
        /// <returns>解析后的工作簿模型。</returns>
        public static OfficeWorkbook Load(string path)
        {
            return Load(path, null);
        }

        /// <summary>从文件路径按指定选项加载 .xls 工作簿。</summary>
        /// <param name="path">.xls 文件路径。</param>
        /// <param name="options">读取选项；null 时使用默认值。</param>
        /// <returns>解析后的工作簿模型。</returns>
        public static OfficeWorkbook Load(string path, ExcelReadOptions options)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(File.ReadAllBytes(path), options);
        }

        /// <summary>从流加载 .xls 工作簿（包含全部隐藏内容）；流会被完整读入内存。</summary>
        /// <param name="stream">包含 .xls 内容的流。</param>
        /// <returns>解析后的工作簿模型。</returns>
        public static OfficeWorkbook Load(Stream stream)
        {
            return Load(stream, null);
        }

        /// <summary>按指定选项从流加载 .xls 工作簿；流会被完整读入内存。</summary>
        /// <param name="stream">包含 .xls 内容的流。</param>
        /// <param name="options">读取选项；null 时使用默认值。</param>
        /// <returns>解析后的工作簿模型。</returns>
        public static OfficeWorkbook Load(Stream stream, ExcelReadOptions options)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return Load(OfficePackageLimits.ReadAllBytes(stream), options);
        }

        /// <summary>从字节数组加载 .xls 工作簿（包含全部隐藏内容）。</summary>
        /// <param name="data">.xls 文件完整字节。</param>
        /// <returns>解析后的工作簿模型。</returns>
        public static OfficeWorkbook Load(byte[] data)
        {
            return Load(data, null);
        }

        /// <summary>
        /// 按指定选项从字节数组加载 .xls 工作簿。检测到 zip 头（.xlsx）或非 OLE2 文件时抛出 InvalidDataException；
        /// 加密工作簿抛出 NotSupportedException。
        /// Loads a .xls workbook from a byte array using the given options. Throws InvalidDataException when a
        /// zip header (.xlsx) or a non-OLE2 file is detected, and NotSupportedException for encrypted workbooks.
        /// </summary>
        /// <param name="data">.xls 文件完整字节。</param>
        /// <param name="options">读取选项；null 时使用默认值。</param>
        /// <returns>解析后的工作簿模型。</returns>
        public static OfficeWorkbook Load(byte[] data, ExcelReadOptions options)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length >= 4 && data[0] == 0x50 && data[1] == 0x4B)
            {
                throw new InvalidDataException("The file is an .xlsx package. Use XlsxReader instead.");
            }
            if (!CompoundFile.IsCompoundFile(data))
            {
                throw new InvalidDataException("The file is not a legacy Excel workbook.");
            }

            CompoundFile compound = new CompoundFile(data);
            byte[] stream = compound.TryGetStream("Workbook");
            if (stream == null) stream = compound.TryGetStream("Book");
            if (stream == null || stream.Length < 4)
            {
                throw new InvalidDataException("The Workbook stream is missing.");
            }

            Reader reader = new Reader(stream, options ?? new ExcelReadOptions());
            OfficeWorkbook workbook = reader.Read();
            ReadMetadata(compound, workbook.Metadata);
            return workbook;
        }

        /// <summary>提取工作簿全部文本：逐表输出 [表名] 标题行，每行单元格以制表符分隔；
        /// 与 XLSX 一致地沿稀疏单元格序列输出，避免按行列全网格扫描。</summary>
        /// <param name="path">.xls 文件路径。</param>
        /// <returns>纯文本内容。</returns>
        public static string ExtractText(string path)
        {
            OfficeWorkbook workbook = Load(path);
            StringBuilder builder = new StringBuilder();
            for (int s = 0; s < workbook.Worksheets.Count; s++)
            {
                Worksheet sheet = workbook.Worksheets[s];
                if (builder.Length > 0) builder.AppendLine();
                builder.Append('[').Append(sheet.Name).Append(']').AppendLine();
                XlsxReader.AppendSheetText(builder, sheet);
            }
            return builder.ToString();
        }

        /// <summary>
        /// Workbook 流解析器：先扫描全局记录（BOUNDSHEET 表目录、SST 共享字符串、
        /// FORMAT/XF/FONT 样式表、DATE1904），再按表偏移逐个解析工作表记录。
        /// 下列常量为本解析器支持的 BIFF8 记录编号。
        /// Workbook stream parser: first scans the global records (BOUNDSHEET sheet directory, SST shared strings,
        /// FORMAT/XF/FONT style tables, DATE1904), then parses each worksheet's records by sheet offset.
        /// The constants below are the BIFF8 record ids supported by this parser.
        /// </summary>
        private sealed class Reader
        {
            /// <summary>BOF：记录块起始标志（文件/工作表级）。</summary>
            private const ushort RecordBof = 0x0809;
            /// <summary>EOF：记录块结束标志。</summary>
            private const ushort RecordEof = 0x000A;
            /// <summary>CONTINUE：超长记录（如 SST）的续块。</summary>
            private const ushort RecordContinue = 0x003C;
            /// <summary>FILEPASS：工作簿加密标记（不支持读取）。</summary>
            private const ushort RecordFilePass = 0x002F;
            /// <summary>BOUNDSHEET：工作表目录项（名称、隐藏状态、BOF 偏移）。</summary>
            private const ushort RecordBoundSheet = 0x0085;
            /// <summary>SST：共享字符串表。</summary>
            private const ushort RecordSst = 0x00FC;
            /// <summary>FORMAT：自定义数字格式定义。</summary>
            private const ushort RecordFormat = 0x041E;
            /// <summary>XF：单元格格式记录（引用字体与格式编号）。</summary>
            private const ushort RecordXf = 0x00E0;
            /// <summary>FONT：字体记录。</summary>
            private const ushort RecordFont = 0x0031;
            /// <summary>DATE1904：日期系统选择（1904 或 1900）。</summary>
            private const ushort RecordDateFormat = 0x0022;
            /// <summary>LABELsst：引用共享字符串的文本单元格。</summary>
            private const ushort RecordLabelSst = 0x00FD;
            /// <summary>LABEL：直接存储 Unicode 文本的单元格。</summary>
            private const ushort RecordLabel = 0x0204;
            /// <summary>RSTRING：带富文本 runs 的字符串单元格。</summary>
            private const ushort RecordRString = 0x00D6;
            /// <summary>NUMBER：IEEE 双精度数值单元格。</summary>
            private const ushort RecordNumber = 0x0203;
            /// <summary>RK：4 字节压缩数值单元格。</summary>
            private const ushort RecordRk = 0x027E;
            /// <summary>MULRK：同一行多个连续 RK 数值单元格。</summary>
            private const ushort RecordMulRk = 0x00BD;
            /// <summary>FORMULA：公式单元格及其缓存结果。</summary>
            private const ushort RecordFormula = 0x0006;
            /// <summary>STRING：公式字符串型缓存结果（跟在 FORMULA 之后）。</summary>
            private const ushort RecordString = 0x0207;
            /// <summary>BOOLERR：布尔或错误值单元格。</summary>
            private const ushort RecordBoolErr = 0x0205;
            /// <summary>BLANK：带样式的空单元格。</summary>
            private const ushort RecordBlank = 0x0201;
            /// <summary>MULBLANK：同一行多个连续带样式空单元格。</summary>
            private const ushort RecordMulBlank = 0x00BE;
            /// <summary>ROW：行记录（行高、隐藏标记）。</summary>
            private const ushort RecordRow = 0x0208;
            /// <summary>COLINFO：列区间信息（宽度、隐藏标记）。</summary>
            private const ushort RecordColInfo = 0x007D;
            /// <summary>DEFCOLWIDTH：默认列宽。</summary>
            private const ushort RecordDefColWidth = 0x0055;
            /// <summary>STANDARDWIDTH：标准列宽（1/256 字符单位）。</summary>
            private const ushort RecordStandardWidth = 0x0099;
            /// <summary>DEFAULTROWHEIGHT：默认行高（twip 单位）。</summary>
            private const ushort RecordDefaultRowHeight = 0x0225;
            /// <summary>MERGEDCELLS：合并单元格区域列表。</summary>
            private const ushort RecordMergedCells = 0x00E5;
            /// <summary>WINDOW2：工作表窗口选项（网格线、冻结标记等）。</summary>
            private const ushort RecordWindow2 = 0x023E;
            /// <summary>PANE：冻结窗格位置。</summary>
            private const ushort RecordPane = 0x0041;
            /// <summary>SETUP：打印页面设置（纸张、方向、缩放适配）。</summary>
            private const ushort RecordSetup = 0x00A1;
            /// <summary>LEFTMARGIN：左边距（英寸）。</summary>
            private const ushort RecordLeftMargin = 0x0026;
            /// <summary>RIGHTMARGIN：右边距（英寸）。</summary>
            private const ushort RecordRightMargin = 0x0027;
            /// <summary>TOPMARGIN：上边距（英寸）。</summary>
            private const ushort RecordTopMargin = 0x0028;
            /// <summary>BOTTOMMARGIN：下边距（英寸）。</summary>
            private const ushort RecordBottomMargin = 0x0029;
            /// <summary>HEADERMARGIN：页眉距离（英寸）。</summary>
            private const ushort RecordHeaderMargin = 0x0014;
            /// <summary>FOOTERMARGIN：页脚距离（英寸）。</summary>
            private const ushort RecordFooterMargin = 0x0015;

            /// <summary>BOUNDSHEET 目录项上限，超出部分视为损坏文件并忽略。 Cap on BOUNDSHEET directory entries; extras are ignored as a corrupt file.</summary>
            private const int MaxSheets = 1024;

            private readonly byte[] _stream;
            private readonly ExcelReadOptions _options;
            private List<BiffRecord> _records;
            private readonly List<SheetEntry> _sheets = new List<SheetEntry>();
            private Dictionary<int, int> _offsetIndex;
            private readonly List<string> _sharedStrings = new List<string>();
            private readonly Dictionary<int, string> _formats = new Dictionary<int, string>();
            private readonly List<CellStyle> _styles = new List<CellStyle>();
            private readonly List<CellStyle> _fonts = new List<CellStyle>();
            private bool _date1904;
            private bool _sawFilePass;
            private CellStyle _defaultStyle = new CellStyle();

            /// <summary>创建解析器。</summary>
            /// <param name="stream">Workbook 流的完整字节。</param>
            /// <param name="options">读取选项。</param>
            public Reader(byte[] stream, ExcelReadOptions options)
            {
                _stream = stream;
                _options = options;
            }

            /// <summary>
            /// 执行完整读取：拆分 BIFF 记录、校验 BOF 版本为 BIFF8，解析全局记录后
            /// 按 BOUNDSHEET 偏移（找不到时顺序回退到工作表 BOF）逐表读取内容。
            /// Performs the full read: splits the BIFF records, validates the BOF version as BIFF8, parses the
            /// global records, then reads each sheet from its BOUNDSHEET offset (falling back sequentially to
            /// worksheet BOFs when the offset is not found).
            /// </summary>
            /// <returns>填充完成的工作簿模型。</returns>
            public OfficeWorkbook Read()
            {
                _records = ReadRecords(_stream);
                if (_records.Count == 0)
                {
                    throw new InvalidDataException("The Workbook stream contains no BIFF records.");
                }

                BiffRecord first = _records[0];
                if (first.Id != RecordBof || first.Length < 4)
                {
                    throw new InvalidDataException("The Workbook stream does not start with a BOF record.");
                }
                ushort version = ByteUtils.ReadUInt16LittleEndian(first.Data, 0);
                if (version != 0x0600)
                {
                    throw new NotSupportedException("Only BIFF8 (Excel 97-2003) workbooks are supported.");
                }

                int globalEnd = FindFirstWorksheetRecord();
                ParseGlobals(globalEnd);
                if (_sawFilePass)
                {
                    throw new NotSupportedException("Encrypted Excel workbooks are not supported.");
                }

                OfficeWorkbook workbook = new OfficeWorkbook();
                _defaultStyle = _styles.Count > 0 ? _styles[0].Clone() : new CellStyle();
                workbook.DefaultStyle = _defaultStyle.Clone();

                int fallbackIndex = globalEnd;
                HashSet<int> parsed = new HashSet<int>();
                for (int i = 0; i < _sheets.Count; i++)
                {
                    SheetEntry entry = _sheets[i];
                    if (entry.Hidden && !_options.IncludeHiddenWorksheets) continue;
                    Worksheet sheet = workbook.AddWorksheet(entry.Name);
                    sheet.IsHidden = entry.Hidden;
                    int start = FindRecordIndex(entry.Offset);
                    if (start < 0)
                    {
                        while (fallbackIndex < _records.Count)
                        {
                            BiffRecord record = _records[fallbackIndex];
                            if (record.Id == RecordBof && record.Length >= 4
                                && ByteUtils.ReadUInt16LittleEndian(record.Data, 2) == 0x0010)
                            {
                                start = fallbackIndex;
                                fallbackIndex++;
                                break;
                            }
                            fallbackIndex++;
                        }
                    }
                    if (start >= 0 && parsed.Add(start)) ReadWorksheet(sheet, start);
                    if (!_options.IncludeHiddenRows || !_options.IncludeHiddenColumns)
                    {
                        RemoveHiddenCells(sheet);
                    }
                }

                if (workbook.Worksheets.Count == 0)
                {
                    throw new InvalidDataException("The workbook does not contain any worksheets.");
                }
                return workbook;
            }

            /// <summary>查找第一个工作表级 BOF（substream=0x0010）的下标，作为全局记录区的终点。</summary>
            /// <returns>工作表 BOF 下标；不存在时返回记录总数。</returns>
            private int FindFirstWorksheetRecord()
            {
                for (int i = 1; i < _records.Count; i++)
                {
                    BiffRecord record = _records[i];
                    if (record.Id != RecordBof || record.Length < 4) continue;
                    if (ByteUtils.ReadUInt16LittleEndian(record.Data, 2) == 0x0010) return i;
                }
                return _records.Count;
            }

            /// <summary>
            /// 按流内字节偏移查找对应记录的下标。首次调用时一次性建立偏移索引，
            /// 之后为 O(1)，避免每个 BOUNDSHEET 都线性扫描全部记录。
            /// Finds the record index for a byte offset in the stream. The offset index is built once on first
            /// use so later lookups are O(1) instead of a linear scan per BOUNDSHEET entry.
            /// </summary>
            /// <param name="offset">BOUNDSHEET 记录的工作表 BOF 偏移。 The worksheet BOF offset from the BOUNDSHEET record.</param>
            /// <returns>记录下标；未找到返回 -1。 The record index, or -1 when not found.</returns>
            private int FindRecordIndex(int offset)
            {
                if (offset <= 0) return -1;
                Dictionary<int, int> index = _offsetIndex;
                if (index == null)
                {
                    index = new Dictionary<int, int>();
                    for (int i = 0; i < _records.Count; i++)
                    {
                        int key = _records[i].Offset;
                        if (!index.ContainsKey(key)) index.Add(key, i);
                    }
                    _offsetIndex = index;
                }
                int found;
                return index.TryGetValue(offset, out found) ? found : -1;
            }

            /// <summary>
            /// 解析全局记录区：收集 BOUNDSHEET 目录、SST（含 CONTINUE 续块拼接）、
            /// FORMAT 自定义格式、FONT 字体、XF 样式与 DATE1904 设置；FILEPASS 仅标记加密。
            /// Parses the global record area: collects the BOUNDSHEET directory, SST (with CONTINUE chunk
            /// stitching), FORMAT custom number formats, FONT fonts, XF styles and the DATE1904 setting;
            /// FILEPASS only marks the file as encrypted.
            /// </summary>
            /// <param name="globalEnd">全局记录区的记录数终点。</param>
            private void ParseGlobals(int globalEnd)
            {
                for (int i = 0; i < globalEnd; i++)
                {
                    BiffRecord record = _records[i];
                    byte[] data = record.Data;
                    switch (record.Id)
                    {
                        case RecordFilePass:
                            _sawFilePass = true;
                            break;
                        case RecordBoundSheet:
                            ReadBoundSheet(data);
                            break;
                        case RecordSst:
                            {
                                List<byte[]> chunks = new List<byte[]>();
                                chunks.Add(data);
                                while (i + 1 < globalEnd && _records[i + 1].Id == RecordContinue)
                                {
                                    i++;
                                    chunks.Add(_records[i].Data);
                                }
                                ReadSharedStrings(chunks);
                                break;
                            }
                        case RecordFormat:
                            if (data.Length >= 5)
                            {
                                int id = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int position = 2;
                                string code = ReadUnicodeString(data, ref position);
                                if (code != null) _formats[id] = code;
                            }
                            break;
                        case RecordFont:
                            _fonts.Add(ReadFont(data));
                            break;
                        case RecordXf:
                            _styles.Add(ReadXf(data));
                            break;
                        case RecordDateFormat:
                            if (data.Length >= 2) _date1904 = ByteUtils.ReadUInt16LittleEndian(data, 0) == 1;
                            break;
                    }
                }
            }

            /// <summary>解析 BOUNDSHEET 记录：取工作表 BOF 偏移、隐藏标志与 8 位长度的表名（仅保留 worksheet 类型）。</summary>
            /// <param name="data">记录数据。</param>
            private void ReadBoundSheet(byte[] data)
            {
                if (data.Length < 8 || _sheets.Count >= MaxSheets) return;
                int offset = ByteUtils.ReadInt32LittleEndian(data, 0);
                int flags = ByteUtils.ReadUInt16LittleEndian(data, 4);
                int sheetType = (flags >> 8) & 0xFF;
                if (sheetType != 0) return;
                bool hidden = (flags & 0x03) != 0;

                int charCount = data[6];
                int optionFlags = data[7];
                int position = 8;
                bool wide = (optionFlags & 0x01) != 0;
                int bytes = charCount * (wide ? 2 : 1);
                if (position + bytes > data.Length) bytes = data.Length - position;
                string name = wide
                    ? (bytes >= 2 ? Encoding.Unicode.GetString(data, position, bytes - (bytes % 2)) : string.Empty)
                    : DecodeLatin1(data, position, bytes);
                if (name.Length == 0) name = "Sheet" + (_sheets.Count + 1).ToString(CultureInfo.InvariantCulture);
                _sheets.Add(new SheetEntry(name, offset, hidden));
            }

            /// <summary>解析 SST 记录体：读取总数与唯一串数后逐个提取字符串（唯一串数量超 200 万视为损坏）。</summary>
            /// <param name="chunks">SST 及其 CONTINUE 续块的数据列表。</param>
            private void ReadSharedStrings(List<byte[]> chunks)
            {
                SstReader reader = new SstReader(chunks);
                uint total = reader.ReadUInt32();
                uint unique = reader.ReadUInt32();
                if (unique > 2000000) throw new InvalidDataException("The shared string table is too large.");
                for (uint i = 0; i < unique; i++)
                {
                    _sharedStrings.Add(reader.ReadString());
                }
            }

            /// <summary>解析 FONT 记录：提取字号（twip 转磅）、粗体/斜体/下划线标志、颜色索引与字体名。</summary>
            /// <param name="data">记录数据。</param>
            /// <returns>仅含字体属性的样式对象。</returns>
            private CellStyle ReadFont(byte[] data)
            {
                CellStyle style = new CellStyle();
                if (data.Length < 16) return style;

                int height = ByteUtils.ReadUInt16LittleEndian(data, 0);
                if (height > 0) style.FontSize = height / 20.0;
                ushort flags = ByteUtils.ReadUInt16LittleEndian(data, 2);
                style.Bold = (flags & 0x0001) != 0;
                style.Italic = (flags & 0x0002) != 0;
                ushort colorIndex = ByteUtils.ReadUInt16LittleEndian(data, 4);
                style.FontColor = MapColorIndex(colorIndex);
                int underline = data[10];
                style.Underline = underline != 0;

                int charCount = data[14];
                bool wide = (data[15] & 0x01) != 0;
                int position = 16;
                int bytes = charCount * (wide ? 2 : 1);
                if (position + bytes > data.Length) bytes = data.Length - position;
                if (bytes > 0)
                {
                    style.FontFamily = wide
                        ? Encoding.Unicode.GetString(data, position, bytes - (bytes % 2))
                        : DecodeLatin1(data, position, bytes);
                }
                if (string.IsNullOrEmpty(style.FontFamily)) style.FontFamily = "Calibri";
                return style;
            }

            /// <summary>解析 XF 记录：以引用的 FONT 属性为基础，附加 FORMAT 编号对应的数字格式。</summary>
            /// <param name="data">记录数据。</param>
            /// <returns>该 XF 索引对应的单元格样式。</returns>
            private CellStyle ReadXf(byte[] data)
            {
                CellStyle style = new CellStyle();
                if (data.Length >= 2)
                {
                    int fontIndex = ByteUtils.ReadUInt16LittleEndian(data, 0);
                    if (fontIndex >= 0 && fontIndex < _fonts.Count)
                    {
                        style = _fonts[fontIndex].Clone();
                    }
                }
                if (data.Length >= 4)
                {
                    int formatId = ByteUtils.ReadUInt16LittleEndian(data, 2);
                    style.NumberFormat = GetFormat(formatId);
                }
                return style;
            }

            /// <summary>解析格式编号：优先自定义 FORMAT，其次内置格式，最终回退 General。</summary>
            /// <param name="id">格式编号。</param>
            /// <returns>格式代码文本。</returns>
            private string GetFormat(int id)
            {
                string custom;
                if (_formats.TryGetValue(id, out custom)) return custom;
                return XlsxNumberFormats.GetBuiltIn(id) ?? "General";
            }

            /// <summary>按选项移除位于隐藏行或隐藏列中的单元格（先收集后删除，避免遍历中修改集合）。</summary>
            /// <param name="sheet">要清理的工作表。</param>
            private void RemoveHiddenCells(Worksheet sheet)
            {
                List<Cell> removed = new List<Cell>();
                foreach (Cell cell in sheet.Cells)
                {
                    if (!_options.IncludeHiddenRows && sheet.HiddenRows.Contains(cell.Row)) removed.Add(cell);
                    else if (!_options.IncludeHiddenColumns && sheet.HiddenColumns.Contains(cell.Column)) removed.Add(cell);
                }
                for (int i = 0; i < removed.Count; i++)
                {
                    sheet.RemoveCell(removed[i].Row, removed[i].Column);
                }
            }

            /// <summary>
            /// 从给定 BOF 下标开始解析单个工作表子流直到匹配的 EOF（支持嵌套 BOF 计数），
            /// 处理各类单元格记录、行/列尺寸、合并区域、窗口冻结与打印设置记录。
            /// FORMULA 结果为字符串时挂起单元格等待紧随其后的 STRING 记录。
            /// Parses one worksheet substream from the given BOF index up to the matching EOF (counting nested
            /// BOFs), handling cell records of all kinds, row/column sizes, merged ranges, window freeze and
            /// print-setup records. When a FORMULA result is a string, the cell is deferred until the STRING
            /// record that immediately follows.
            /// </summary>
            /// <param name="sheet">目标工作表模型。</param>
            /// <param name="startIndex">工作表级 BOF 记录下标。</param>
            private void ReadWorksheet(Worksheet sheet, int startIndex)
            {
                bool frozen = false;
                int paneRows = 0;
                int paneColumns = 0;
                int nested = 0;
                Cell pendingStringCell = null;

                for (int i = startIndex; i < _records.Count; i++)
                {
                    BiffRecord record = _records[i];
                    byte[] data = record.Data;
                    ushort id = record.Id;

                    if (id == RecordBof)
                    {
                        if (i > startIndex) nested++;
                        continue;
                    }
                    if (id == RecordEof)
                    {
                        if (nested > 0)
                        {
                            nested--;
                            continue;
                        }
                        break;
                    }
                    if (nested > 0) continue;

                    switch (id)
                    {
                        case RecordLabelSst:
                            if (data.Length >= 10)
                            {
                                int index = ByteUtils.ReadInt32LittleEndian(data, 6);
                                string value = index >= 0 && index < _sharedStrings.Count ? _sharedStrings[index] : null;
                                SetCell(sheet, data, 0, 4, value);
                            }
                            break;

                        case RecordLabel:
                        case RecordRString:
                            if (data.Length >= 8)
                            {
                                int position = 6;
                                string value = ReadUnicodeString(data, ref position);
                                SetCell(sheet, data, 0, 4, value);
                            }
                            break;

                        case RecordNumber:
                            if (data.Length >= 14)
                            {
                                double number = ReadDouble(data, 6);
                                SetCell(sheet, data, 0, 4, ConvertNumber(sheet, number, ByteUtils.ReadUInt16LittleEndian(data, 4)));
                            }
                            break;

                        case RecordRk:
                            if (data.Length >= 10)
                            {
                                double number = DecodeRk(ByteUtils.ReadUInt32LittleEndian(data, 6));
                                SetCell(sheet, data, 0, 4, ConvertNumber(sheet, number, ByteUtils.ReadUInt16LittleEndian(data, 4)));
                            }
                            break;

                        case RecordMulRk:
                            if (data.Length >= 6)
                            {
                                int row = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int firstColumn = ByteUtils.ReadUInt16LittleEndian(data, 2);
                                int count = (data.Length - 6) / 6;
                                for (int c = 0; c < count; c++)
                                {
                                    int offset = 4 + c * 6;
                                    int xf = ByteUtils.ReadUInt16LittleEndian(data, offset);
                                    double number = DecodeRk(ByteUtils.ReadUInt32LittleEndian(data, offset + 2));
                                    SetCell(sheet, row, firstColumn + c, xf, ConvertNumber(sheet, number, xf));
                                }
                            }
                            break;

                        case RecordFormula:
                            if (data.Length >= 22)
                            {
                                int row = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int column = ByteUtils.ReadUInt16LittleEndian(data, 2);
                                int xf = ByteUtils.ReadUInt16LittleEndian(data, 4);
                                pendingStringCell = null;
                                if (data[12] == 0xFF && data[13] == 0xFF)
                                {
                                    int kind = data[6];
                                    if (kind == 0)
                                    {
                                        Cell cell = SetCell(sheet, row, column, xf, null);
                                        pendingStringCell = cell;
                                    }
                                    else if (kind == 1)
                                    {
                                        SetCell(sheet, row, column, xf, data[8] != 0);
                                    }
                                    else if (kind == 2)
                                    {
                                        SetCell(sheet, row, column, xf, GetErrorText(data[8]));
                                    }
                                    else
                                    {
                                        SetCell(sheet, row, column, xf, null);
                                    }
                                }
                                else
                                {
                                    SetCell(sheet, row, column, xf, ConvertNumber(sheet, ReadDouble(data, 6), xf));
                                }
                            }
                            break;

                        case RecordString:
                            if (pendingStringCell != null)
                            {
                                int position = 0;
                                pendingStringCell.Value = ReadUnicodeString(data, ref position);
                                pendingStringCell = null;
                            }
                            break;

                        case RecordBoolErr:
                            if (data.Length >= 8)
                            {
                                int row = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int column = ByteUtils.ReadUInt16LittleEndian(data, 2);
                                int xf = ByteUtils.ReadUInt16LittleEndian(data, 4);
                                object value = data[7] == 0 ? (object)(data[6] != 0) : GetErrorText(data[6]);
                                SetCell(sheet, row, column, xf, value);
                            }
                            break;

                        case RecordBlank:
                            if (data.Length >= 6)
                            {
                                int row = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int column = ByteUtils.ReadUInt16LittleEndian(data, 2);
                                int xf = ByteUtils.ReadUInt16LittleEndian(data, 4);
                                SetCell(sheet, row, column, xf, null);
                            }
                            break;

                        case RecordMulBlank:
                            if (data.Length >= 6)
                            {
                                int row = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int firstColumn = ByteUtils.ReadUInt16LittleEndian(data, 2);
                                int count = (data.Length - 6) / 2;
                                for (int c = 0; c < count; c++)
                                {
                                    int xf = ByteUtils.ReadUInt16LittleEndian(data, 4 + c * 2);
                                    SetCell(sheet, row, firstColumn + c, xf, null);
                                }
                            }
                            break;

                        case RecordRow:
                            if (data.Length >= 8)
                            {
                                int row = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int height = ByteUtils.ReadUInt16LittleEndian(data, 6);
                                if (height > 0)
                                {
                                    sheet.SetRowHeight(row + 1, height / 20.0);
                                }
                                if (data.Length >= 14 && (ByteUtils.ReadUInt16LittleEndian(data, 12) & 0x0020) != 0)
                                {
                                    sheet.SetRowHidden(row + 1, true);
                                }
                            }
                            break;

                        case RecordColInfo:
                            if (data.Length >= 8)
                            {
                                int first = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int last = ByteUtils.ReadUInt16LittleEndian(data, 2);
                                int width = ByteUtils.ReadUInt16LittleEndian(data, 4);
                                bool hidden = data.Length >= 10 && (ByteUtils.ReadUInt16LittleEndian(data, 8) & 0x0001) != 0;
                                if (last > 16383) last = 16383;
                                for (int c = first; c <= last; c++)
                                {
                                    if (width > 0) sheet.SetColumnWidth(c + 1, width / 256.0);
                                    if (hidden) sheet.SetColumnHidden(c + 1, true);
                                }
                            }
                            break;

                        case RecordDefColWidth:
                            if (data.Length >= 2)
                            {
                                int width = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                if (width > 0) sheet.DefaultColumnWidth = width;
                            }
                            break;

                        case RecordStandardWidth:
                            if (data.Length >= 2)
                            {
                                int width = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                if (width > 0) sheet.DefaultColumnWidth = width / 256.0;
                            }
                            break;

                        case RecordDefaultRowHeight:
                            if (data.Length >= 4)
                            {
                                int height = ByteUtils.ReadUInt16LittleEndian(data, 2);
                                if (height > 0) sheet.DefaultRowHeight = height / 20.0;
                            }
                            break;

                        case RecordMergedCells:
                            if (data.Length >= 2)
                            {
                                int count = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int position = 2;
                                for (int m = 0; m < count && position + 8 <= data.Length; m++, position += 8)
                                {
                                    int firstRow = ByteUtils.ReadUInt16LittleEndian(data, position);
                                    int lastRow = ByteUtils.ReadUInt16LittleEndian(data, position + 2);
                                    int firstColumn = ByteUtils.ReadUInt16LittleEndian(data, position + 4);
                                    int lastColumn = ByteUtils.ReadUInt16LittleEndian(data, position + 6);
                                    if (lastRow < firstRow || lastColumn < firstColumn) continue;
                                    sheet.MergedRanges.Add(new MergedRange(firstRow + 1, firstColumn + 1,
                                        lastRow - firstRow + 1, lastColumn - firstColumn + 1));
                                }
                            }
                            break;

                        case RecordWindow2:
                            if (data.Length >= 2)
                            {
                                int options = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                sheet.ShowGridLines = (options & 0x0002) != 0;
                                frozen = (options & 0x0008) != 0;
                                if (frozen)
                                {
                                    sheet.FrozenRows = paneRows;
                                    sheet.FrozenColumns = paneColumns;
                                }
                            }
                            break;

                        case RecordPane:
                            if (data.Length >= 9)
                            {
                                paneRows = ByteUtils.ReadUInt16LittleEndian(data, 4);
                                paneColumns = ByteUtils.ReadUInt16LittleEndian(data, 6);
                                if (frozen)
                                {
                                    sheet.FrozenRows = paneRows;
                                    sheet.FrozenColumns = paneColumns;
                                }
                            }
                            break;

                        case RecordSetup:
                            if (data.Length >= 12)
                            {
                                int paperSize = ByteUtils.ReadUInt16LittleEndian(data, 0);
                                int fitWidth = ByteUtils.ReadUInt16LittleEndian(data, 6);
                                int fitHeight = ByteUtils.ReadUInt16LittleEndian(data, 8);
                                int options = ByteUtils.ReadUInt16LittleEndian(data, 10);
                                ApplyPaperSize(sheet.PrintSetup, paperSize);
                                if ((options & 0x0002) != 0)
                                {
                                    double swap = sheet.PrintSetup.Width;
                                    sheet.PrintSetup.Width = sheet.PrintSetup.Height;
                                    sheet.PrintSetup.Height = swap;
                                }
                                if (fitWidth != 0 || fitHeight != 0) sheet.FitToPage = true;
                            }
                            break;

                        case RecordLeftMargin:
                            if (data.Length >= 8) sheet.PrintSetup.MarginLeft = ReadDouble(data, 0) * 72.0;
                            break;
                        case RecordRightMargin:
                            if (data.Length >= 8) sheet.PrintSetup.MarginRight = ReadDouble(data, 0) * 72.0;
                            break;
                        case RecordTopMargin:
                            if (data.Length >= 8) sheet.PrintSetup.MarginTop = ReadDouble(data, 0) * 72.0;
                            break;
                        case RecordBottomMargin:
                            if (data.Length >= 8) sheet.PrintSetup.MarginBottom = ReadDouble(data, 0) * 72.0;
                            break;
                        case RecordHeaderMargin:
                            if (data.Length >= 8) sheet.PrintSetup.HeaderDistance = ReadDouble(data, 0) * 72.0;
                            break;
                        case RecordFooterMargin:
                            if (data.Length >= 8) sheet.PrintSetup.FooterDistance = ReadDouble(data, 0) * 72.0;
                            break;
                    }
                }
            }

            /// <summary>从记录字节中按偏移读取行/列索引与 XF 索引并写入单元格值。</summary>
            /// <param name="sheet">目标工作表。</param>
            /// <param name="data">记录数据。</param>
            /// <param name="rowOffset">行号（16 位）的字节偏移，列号在其后 2 字节。</param>
            /// <param name="xfOffset">XF 索引的字节偏移。</param>
            /// <param name="value">单元格值。</param>
            /// <returns>写入的单元格。</returns>
            private Cell SetCell(Worksheet sheet, byte[] data, int rowOffset, int xfOffset, object value)
            {
                int row = ByteUtils.ReadUInt16LittleEndian(data, rowOffset);
                int column = ByteUtils.ReadUInt16LittleEndian(data, rowOffset + 2);
                int xf = ByteUtils.ReadUInt16LittleEndian(data, xfOffset);
                return SetCell(sheet, row, column, xf, value);
            }

            /// <summary>按 0 基行列索引写入单元格值并应用 XF 样式（模型行列按 1 基存储）。</summary>
            /// <param name="sheet">目标工作表。</param>
            /// <param name="row">0 基行号。</param>
            /// <param name="column">0 基列号。</param>
            /// <param name="xf">XF 样式索引。</param>
            /// <param name="value">单元格值。</param>
            /// <returns>写入的单元格。</returns>
            private Cell SetCell(Worksheet sheet, int row, int column, int xf, object value)
            {
                Cell cell = sheet.GetCell(row + 1, column + 1);
                cell.Value = value;
                cell.Style = GetStyle(xf);
                return cell;
            }

            /// <summary>数值单元格转换：样式为日期格式时把序列号转为 DateTime，否则原样返回数值。</summary>
            /// <param name="sheet">所在工作表（未使用日期基准由 _date1904 决定）。</param>
            /// <param name="value">数值（日期序列号）。</param>
            /// <param name="xf">XF 样式索引。</param>
            /// <returns>double 或 DateTime 值。</returns>
            private object ConvertNumber(Worksheet sheet, double value, int xf)
            {
                if (!IsDateFormat(xf)) return value;
                DateTime epoch = _date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
                if (value < 0) value = 0;
                return epoch.AddDays(value);
            }

            /// <summary>按 XF 索引返回样式副本；越界时返回默认样式副本。</summary>
            /// <param name="xf">XF 样式索引。</param>
            /// <returns>样式对象副本。</returns>
            private CellStyle GetStyle(int xf)
            {
                if (xf >= 0 && xf < _styles.Count) return _styles[xf].Clone();
                return _defaultStyle.Clone();
            }

            /// <summary>判断指定 XF 索引的样式是否为日期格式。</summary>
            /// <param name="xf">XF 样式索引。</param>
            /// <returns>日期格式返回 true。</returns>
            private bool IsDateFormat(int xf)
            {
                return xf >= 0 && xf < _styles.Count && _styles[xf].IsDateFormat();
            }
        }

        /// <summary>按纸张编号（1 Letter、5 Legal、8 A4，其余默认 A4）设置页面宽高（磅），横向由调用方交换。</summary>
        /// <param name="setup">目标页面设置。</param>
        /// <param name="paperSize">BIFF SETUP 记录中的纸张编号。</param>
        private static void ApplyPaperSize(PageSetup setup, int paperSize)
        {
            if (paperSize == 1)
            {
                setup.Width = 612;
                setup.Height = 792;
            }
            else if (paperSize == 5)
            {
                setup.Width = 612;
                setup.Height = 1008;
            }
            else if (paperSize == 8)
            {
                setup.Width = 841.89;
                setup.Height = 1190.55;
            }
            else
            {
                setup.Width = 595.28;
                setup.Height = 841.89;
            }
        }

        /// <summary>将 BIFF 错误码映射为 Excel 错误文本（如 #DIV/0!），未知码返回 #ERROR。</summary>
        /// <param name="code">错误码字节。</param>
        /// <returns>错误文本。</returns>
        private static string GetErrorText(int code)
        {
            switch (code)
            {
                case 0x00: return "#NULL!";
                case 0x07: return "#DIV/0!";
                case 0x0F: return "#VALUE!";
                case 0x17: return "#REF!";
                case 0x1D: return "#NAME?";
                case 0x24: return "#NUM!";
                case 0x2A: return "#N/A";
                default: return "#ERROR";
            }
        }

        /// <summary>
        /// 解码 RK 压缩数值：bit1 为 1 时取 30 位有符号整数，否则高 30 位是 double 尾数高位；
        /// bit0 为 1 时结果再除以 100。
        /// Decodes an RK compressed number: when bit1 is set the value is a 30-bit signed integer, otherwise the
        /// high 30 bits hold the top mantissa bits of a double; when bit0 is set the result is divided by 100.
        /// </summary>
        /// <param name="rk">RK 记录的 4 字节原始值。</param>
        /// <returns>解码后的数值。</returns>
        private static double DecodeRk(uint rk)
        {
            double value;
            if ((rk & 0x02) != 0)
            {
                value = (int)rk >> 2;
            }
            else
            {
                long bits = (long)(rk & 0xFFFFFFFC) << 32;
                value = BitConverter.Int64BitsToDouble(bits);
            }
            if ((rk & 0x01) != 0) value /= 100.0;
            return value;
        }

        /// <summary>从小端字节流指定偏移读取 IEEE 754 双精度浮点数。</summary>
        /// <param name="data">字节数据。</param>
        /// <param name="offset">起始偏移。</param>
        /// <returns>读取的浮点值。</returns>
        private static double ReadDouble(byte[] data, int offset)
        {
            ulong bits = (ulong)ByteUtils.ReadUInt32LittleEndian(data, offset)
                | ((ulong)ByteUtils.ReadUInt32LittleEndian(data, offset + 4) << 32);
            return BitConverter.Int64BitsToDouble((long)bits);
        }

        /// <summary>
        /// 读取 BIFF8 Unicode 字符串（2 字节长度 + 选项标志 + 窄/宽字符数据），
        /// 跳过富文本 runs 与扩展信息，并把 position 推进到字符串末尾之后。
        /// Reads a BIFF8 Unicode string (2-byte length + option flags + narrow/wide character data), skipping
        /// rich-text runs and additional information, and advances position past the end of the string.
        /// </summary>
        /// <param name="data">字节数据。</param>
        /// <param name="position">字符串起始偏移；返回时指向字符串之后。</param>
        /// <returns>解码后的文本；数据不足时返回 null。</returns>
        private static string ReadUnicodeString(byte[] data, ref int position)
        {
            if (position + 3 > data.Length) return null;
            int charCount = ByteUtils.ReadUInt16LittleEndian(data, position);
            position += 2;
            int flags = data[position++];
            bool wide = (flags & 0x01) != 0;
            bool extended = (flags & 0x04) != 0;
            bool rich = (flags & 0x08) != 0;

            int runs = 0;
            int extendedLength = 0;
            if (rich)
            {
                if (position + 2 > data.Length) return null;
                runs = ByteUtils.ReadUInt16LittleEndian(data, position);
                position += 2;
            }
            if (extended)
            {
                if (position + 4 > data.Length) return null;
                extendedLength = ByteUtils.ReadInt32LittleEndian(data, position);
                position += 4;
                if (extendedLength < 0) extendedLength = 0;
            }

            int bytes = charCount * (wide ? 2 : 1);
            if (position + bytes > data.Length) bytes = data.Length - position;
            string value = wide
                ? Encoding.Unicode.GetString(data, position, bytes - (bytes % 2))
                : DecodeLatin1(data, position, bytes);
            position += bytes;

            int skip = runs * 4 + extendedLength;
            if (skip > data.Length - position) skip = data.Length - position;
            position += skip;
            return value;
        }

        /// <summary>
        /// SST 跨块读取器：把 SST 及其 CONTINUE 续块当作连续字节流读取，
        /// 自动处理续块边界（每块首字节为新段选项标志）与窄/宽字符切换。
        /// Cross-chunk SST reader: reads the SST and its CONTINUE chunks as one continuous byte stream,
        /// automatically handling chunk boundaries (the first byte of each chunk is the option flags of the new
        /// segment) and narrow/wide character switching.
        /// </summary>
        private sealed class SstReader
        {
            private readonly List<byte[]> _chunks;
            private int _chunk;
            private int _position;

            /// <summary>创建读取器。</summary>
            /// <param name="chunks">SST 记录与各 CONTINUE 续块的数据列表。</param>
            public SstReader(List<byte[]> chunks)
            {
                _chunks = chunks;
            }

            /// <summary>跨块读取 4 字节小端无符号整数。</summary>
            /// <returns>读取的数值。</returns>
            public uint ReadUInt32()
            {
                int low = ReadUInt16();
                int high = ReadUInt16();
                return (uint)(low | (high << 16));
            }

            /// <summary>跨块读取 2 字节小端无符号整数。</summary>
            /// <returns>读取的数值。</returns>
            public int ReadUInt16()
            {
                int low = ReadByte();
                int high = ReadByte();
                return low | (high << 8);
            }

            /// <summary>读取单字节，当前块耗尽时前进到下一块；数据结束后返回 0。</summary>
            /// <returns>字节值或 0。</returns>
            private int ReadByte()
            {
                while (_chunk < _chunks.Count && _position >= _chunks[_chunk].Length)
                {
                    _chunk++;
                    _position = 0;
                }
                if (_chunk >= _chunks.Count) return 0;
                return _chunks[_chunk][_position++];
            }

            /// <summary>
            /// 读取一个 Unicode 字符串：先读长度与选项标志（宽/窄、富文本 runs、扩展信息），
            /// 逐块解码字符；跨块时以每块首字节重新判断该段的宽窄编码。
            /// Reads one Unicode string: first reads the length and option flags (wide/narrow, rich-text runs,
            /// additional information), then decodes characters chunk by chunk; across chunks the first byte of
            /// each chunk re-decides the narrow/wide encoding of that segment.
            /// </summary>
            /// <returns>解码后的字符串。</returns>
            public string ReadString()
            {
                int charCount = ReadUInt16();
                int flags = ReadByte();
                bool wide = (flags & 0x01) != 0;
                bool extended = (flags & 0x04) != 0;
                bool rich = (flags & 0x08) != 0;

                int runs = 0;
                int extendedLength = 0;
                if (rich)
                {
                    runs = ReadUInt16();
                }
                if (extended)
                {
                    extendedLength = (int)ReadUInt32();
                    if (extendedLength < 0) extendedLength = 0;
                }

                StringBuilder builder = new StringBuilder(charCount);
                int remaining = charCount;
                int consumed = 0;
                while (remaining > 0)
                {
                    while (_chunk < _chunks.Count && _position >= _chunks[_chunk].Length)
                    {
                        _chunk++;
                        _position = 0;
                        if (_chunk >= _chunks.Count) break;
                        if (consumed > 0)
                        {
                            wide = (_chunks[_chunk][_position++] & 0x01) != 0;
                        }
                    }
                    if (_chunk >= _chunks.Count) break;

                    byte[] chunk = _chunks[_chunk];
                    int bytesPerChar = wide ? 2 : 1;
                    int available = (chunk.Length - _position) / bytesPerChar;
                    if (available <= 0)
                    {
                        _position = chunk.Length;
                        continue;
                    }
                    int take = Math.Min(remaining, available);
                    if (wide)
                    {
                        builder.Append(Encoding.Unicode.GetString(chunk, _position, take * 2));
                    }
                    else
                    {
                        builder.Append(DecodeLatin1(chunk, _position, take));
                    }
                    _position += take * bytesPerChar;
                    remaining -= take;
                    consumed += take;
                }

                SkipRaw(runs * 4 + extendedLength);
                return builder.ToString();
            }

            /// <summary>跳过指定数量的原始字节（用于富文本 runs 与扩展信息区）。</summary>
            /// <param name="count">要跳过的字节数。</param>
            private void SkipRaw(int count)
            {
                while (count > 0)
                {
                    while (_chunk < _chunks.Count && _position >= _chunks[_chunk].Length)
                    {
                        _chunk++;
                        _position = 0;
                    }
                    if (_chunk >= _chunks.Count) return;
                    int available = _chunks[_chunk].Length - _position;
                    int take = Math.Min(count, available);
                    _position += take;
                    count -= take;
                }
            }
        }

        /// <summary>拆分后的单条 BIFF 记录：编号、数据与在流中的字节偏移。</summary>
        private sealed class BiffRecord
        {
            /// <summary>记录编号。</summary>
            public ushort Id;
            /// <summary>记录数据（截断到流末尾内）。</summary>
            public byte[] Data;
            /// <summary>记录头在流中的字节偏移。</summary>
            public int Offset;
            /// <summary>数据长度。</summary>
            public int Length { get { return Data == null ? 0 : Data.Length; } }
        }

        /// <summary>BOUNDSHEET 目录项：表名、工作表 BOF 偏移与隐藏标记。</summary>
        private sealed class SheetEntry
        {
            /// <summary>工作表名称。</summary>
            public readonly string Name;
            /// <summary>工作表 BOF 记录在流中的字节偏移。</summary>
            public readonly int Offset;
            /// <summary>是否隐藏。</summary>
            public readonly bool Hidden;

            /// <summary>创建目录项。</summary>
            /// <param name="name">表名。</param>
            /// <param name="offset">BOF 偏移。</param>
            /// <param name="hidden">隐藏标记。</param>
            public SheetEntry(string name, int offset, bool hidden)
            {
                Name = name;
                Offset = offset;
                Hidden = hidden;
            }
        }

        /// <summary>把字节流拆分为 BIFF 记录序列（4 字节头 + 数据），记录数上限 200 万防损坏死循环。</summary>
        /// <param name="data">Workbook 流字节。</param>
        /// <returns>按流顺序排列的记录列表。</returns>
        private static List<BiffRecord> ReadRecords(byte[] data)
        {
            List<BiffRecord> records = new List<BiffRecord>();
            int position = 0;
            while (position + 4 <= data.Length && records.Count < 2000000)
            {
                ushort id = ByteUtils.ReadUInt16LittleEndian(data, position);
                ushort length = ByteUtils.ReadUInt16LittleEndian(data, position + 2);
                int dataStart = position + 4;
                if (dataStart + length > data.Length) length = (ushort)Math.Max(0, data.Length - dataStart);
                byte[] recordData = new byte[length];
                if (length > 0) Buffer.BlockCopy(data, dataStart, recordData, 0, length);
                records.Add(new BiffRecord { Id = id, Data = recordData, Offset = position });
                position = dataStart + length;
            }
            return records;
        }

        /// <summary>将 BIFF 调色板颜色索引映射为 RGB 颜色（0-15 基本色、16-28 扩展色，未知回退黑色）。</summary>
        /// <param name="index">颜色索引。</param>
        /// <returns>对应的颜色。</returns>
        private static OfficeColor MapColorIndex(int index)
        {
            switch (index)
            {
                case 0:
                case 8: return OfficeColor.Black;
                case 1:
                case 9: return OfficeColor.White;
                case 2:
                case 10: return OfficeColor.Red;
                case 3:
                case 11: return new OfficeColor(0, 255, 0);
                case 4:
                case 12: return OfficeColor.Blue;
                case 5:
                case 13: return new OfficeColor(255, 255, 0);
                case 6:
                case 14: return new OfficeColor(255, 0, 255);
                case 7:
                case 15: return new OfficeColor(0, 255, 255);
                case 16: return new OfficeColor(0x80, 0, 0);
                case 17: return new OfficeColor(0, 0x80, 0);
                case 18: return new OfficeColor(0, 0, 0x80);
                case 19: return new OfficeColor(0x80, 0x80, 0);
                case 20: return new OfficeColor(0x80, 0, 0x80);
                case 21: return new OfficeColor(0, 0x80, 0x80);
                case 22: return new OfficeColor(0xC0, 0xC0, 0xC0);
                case 23: return new OfficeColor(0x80, 0x80, 0x80);
                case 24: return new OfficeColor(0x99, 0x99, 0xFF);
                case 25: return new OfficeColor(0x99, 0x33, 0x66);
                case 26: return new OfficeColor(0xFF, 0xFF, 0xCC);
                case 27: return new OfficeColor(0xCC, 0xFF, 0xFF);
                case 28: return new OfficeColor(0x66, 0, 0x66);
                default: return OfficeColor.Black;
            }
        }

        /// <summary>按 Latin-1（字节值即码点）解码窄字符区段，越界自动截断。</summary>
        /// <param name="data">字节数据。</param>
        /// <param name="offset">起始偏移。</param>
        /// <param name="count">字节数。</param>
        /// <returns>解码后的文本。</returns>
        private static string DecodeLatin1(byte[] data, int offset, int count)
        {
            int end = Math.Min(data.Length, offset + count);
            char[] chars = new char[Math.Max(0, end - offset)];
            for (int i = 0; i < chars.Length; i++)
            {
                chars[i] = (char)data[offset + i];
            }
            return new string(chars);
        }

        /// <summary>
        /// 从 OLE2 的 SummaryInformation 属性流读取标题、主题、作者、关键词与创建/修改时间；
        /// 任何解析失败都静默忽略，元数据保持默认值。
        /// Reads title, subject, author, keywords and created/modified times from the OLE2 SummaryInformation
        /// property stream; any parsing failure is silently ignored and the metadata keeps its default values.
        /// </summary>
        /// <param name="compound">OLE2 复合文件。</param>
        /// <param name="metadata">要填充的元数据对象。</param>
        private static void ReadMetadata(CompoundFile compound, DocumentMetadata metadata)
        {
            try
            {
                byte[] stream = compound.TryGetStream("\u0005SummaryInformation");
                if (stream == null || stream.Length < 48) return;
                if (ByteUtils.ReadUInt16LittleEndian(stream, 0) != 0xFFFE) return;
                int sections = ByteUtils.ReadInt32LittleEndian(stream, 24);
                if (sections < 1) return;
                int sectionOffset = ByteUtils.ReadInt32LittleEndian(stream, 44);
                if (sectionOffset < 0 || sectionOffset + 8 > stream.Length) return;
                int propertyCount = ByteUtils.ReadInt32LittleEndian(stream, sectionOffset + 4);
                if (propertyCount <= 0 || propertyCount > 1024) return;
                int codePage = ReadPropertyCodePage(stream, sectionOffset, propertyCount);

                for (int i = 0; i < propertyCount; i++)
                {
                    int entry = sectionOffset + 8 + i * 8;
                    if (entry + 8 > stream.Length) break;
                    int propertyId = ByteUtils.ReadInt32LittleEndian(stream, entry);
                    int valueOffset = ByteUtils.ReadInt32LittleEndian(stream, entry + 4);
                    int position = sectionOffset + valueOffset;
                    if (position < 0 || position + 4 > stream.Length) continue;
                    int type = ByteUtils.ReadInt32LittleEndian(stream, position);
                    switch (propertyId)
                    {
                        case 2:
                            metadata.Title = ReadPropertyString(stream, position, type, codePage) ?? metadata.Title;
                            break;
                        case 3:
                            metadata.Subject = ReadPropertyString(stream, position, type, codePage) ?? metadata.Subject;
                            break;
                        case 4:
                            metadata.Author = ReadPropertyString(stream, position, type, codePage) ?? metadata.Author;
                            break;
                        case 5:
                            metadata.Keywords = ReadPropertyString(stream, position, type, codePage) ?? metadata.Keywords;
                            break;
                        case 12:
                            DateTime created;
                            if (ReadFileTime(stream, position, type, out created)) metadata.Created = created;
                            break;
                        case 13:
                            DateTime modified;
                            if (ReadFileTime(stream, position, type, out modified)) metadata.Modified = modified;
                            break;
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        /// <summary>读取属性集首条 CodePage 属性（Id=1）的代码页编号；缺失时回退 1252。</summary>
        /// <param name="stream">属性流字节。</param>
        /// <param name="sectionOffset">属性段起始偏移。</param>
        /// <param name="propertyCount">属性数量。</param>
        /// <returns>代码页编号。</returns>
        private static int ReadPropertyCodePage(byte[] stream, int sectionOffset, int propertyCount)
        {
            for (int i = 0; i < propertyCount; i++)
            {
                int entry = sectionOffset + 8 + i * 8;
                if (entry + 8 > stream.Length) break;
                if (ByteUtils.ReadInt32LittleEndian(stream, entry) != 1) continue;
                int position = sectionOffset + ByteUtils.ReadInt32LittleEndian(stream, entry + 4);
                if (position < 0 || position + 6 > stream.Length) break;
                int type = ByteUtils.ReadInt32LittleEndian(stream, position);
                if (type == 0x02 || type == 0x12)
                {
                    int value = ByteUtils.ReadUInt16LittleEndian(stream, position + 4);
                    if (value > 0) return value;
                }
                break;
            }
            return 1252;
        }

        /// <summary>读取字符串属性值：0x1E 按声明代码页解码 ANSI 串，0x1F 解码 UTF-16 串，其余类型返回 null。</summary>
        /// <param name="stream">属性流字节。</param>
        /// <param name="position">属性值起始偏移（指向 4 字节类型）。</param>
        /// <param name="type">属性类型编号。</param>
        /// <param name="codePage">ANSI 字符串使用的代码页。</param>
        /// <returns>属性文本或 null。</returns>
        private static string ReadPropertyString(byte[] stream, int position, int type, int codePage)
        {
            if (type == 0x1E)
            {
                int length = ByteUtils.ReadInt32LittleEndian(stream, position + 4);
                if (length <= 0) return string.Empty;
                int start = position + 8;
                if (start + length > stream.Length) return null;
                int textLength = length;
                while (textLength > 0 && stream[start + textLength - 1] == 0) textLength--;
                return DecodeAnsi(stream, start, textLength, codePage);
            }
            if (type == 0x1F)
            {
                int chars = ByteUtils.ReadInt32LittleEndian(stream, position + 4);
                if (chars <= 0) return string.Empty;
                int start = position + 8;
                if (start + chars * 2 > stream.Length) return null;
                int textChars = chars;
                while (textChars > 0 && stream[start + textChars * 2 - 2] == 0 && stream[start + textChars * 2 - 1] == 0) textChars--;
                return Encoding.Unicode.GetString(stream, start, textChars * 2);
            }
            return null;
        }

        /// <summary>读取 FILETIME 属性（类型 0x40，1601 起计的 100 纳秒值）并转为 UTC DateTime。</summary>
        /// <param name="stream">属性流字节。</param>
        /// <param name="position">属性值起始偏移。</param>
        /// <param name="type">属性类型编号。</param>
        /// <param name="value">解析结果；失败时为 DateTime.MinValue。</param>
        /// <returns>解析成功返回 true。</returns>
        private static bool ReadFileTime(byte[] stream, int position, int type, out DateTime value)
        {
            value = DateTime.MinValue;
            if (type != 0x40 || position + 12 > stream.Length) return false;
            uint low = ByteUtils.ReadUInt32LittleEndian(stream, position + 4);
            uint high = ByteUtils.ReadUInt32LittleEndian(stream, position + 8);
            long fileTime = (long)high << 32 | low;
            if (fileTime <= 0) return false;
            try
            {
                value = DateTime.FromFileTimeUtc(fileTime);
                return true;
            }
            catch (ArgumentOutOfRangeException)
            {
                return false;
            }
        }

        /// <summary>按指定代码页解码字节区段；编码不可用时回退 Latin-1 解码。</summary>
        /// <param name="data">字节数据。</param>
        /// <param name="offset">起始偏移。</param>
        /// <param name="count">字节数。</param>
        /// <param name="codePage">代码页编号。</param>
        /// <returns>解码后的文本。</returns>
        private static string DecodeAnsi(byte[] data, int offset, int count, int codePage)
        {
            Encoding encoding = TryGetEncoding(codePage);
            if (encoding != null)
            {
                try
                {
                    return encoding.GetString(data, offset, count);
                }
                catch (ArgumentException)
                {
                }
            }
            return DecodeLatin1(data, offset, count);
        }

        /// <summary>
        /// 获取指定代码页的 Encoding：netstandard2.0 下基础库不含 ANSI 代码页，
        /// 失败时反射尝试注册 CodePagesEncodingProvider 后重试，仍不可用返回 null。
        /// Gets the Encoding for the given code page: the netstandard2.0 base library has no ANSI code pages,
        /// so on failure it reflectively tries to register CodePagesEncodingProvider and retries, returning
        /// null when still unavailable.
        /// </summary>
        /// <param name="codePage">代码页编号。</param>
        /// <returns>编码对象或 null。</returns>
        private static Encoding TryGetEncoding(int codePage)
        {
            if (codePage <= 0) return null;
            try
            {
                return Encoding.GetEncoding(codePage);
            }
            catch (Exception)
            {
            }
            try
            {
                Type providerType = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages", false);
                if (providerType != null)
                {
                    System.Reflection.PropertyInfo instance = providerType.GetProperty(
                        "Instance", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                    System.Reflection.MethodInfo register = typeof(Encoding).GetMethod(
                        "RegisterProvider", new Type[] { typeof(EncodingProvider) });
                    if (instance != null && register != null)
                    {
                        register.Invoke(null, new object[] { instance.GetValue(null, null) });
                        return Encoding.GetEncoding(codePage);
                    }
                }
            }
            catch (Exception)
            {
            }
            return null;
        }
    }
}
