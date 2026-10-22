using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 流式读取得到的一个单元格：列号、值、样式索引与公式。
    /// One cell produced by streaming query: column, value, style index and formula.
    /// </summary>
    public struct ExcelCellData
    {
        /// <summary>创建单元格数据。Creates cell data.</summary>
        /// <param name="column">1 基列号。1-based column index.</param>
        /// <param name="value">单元格值（可能为 null）。Cell value (may be null).</param>
        /// <param name="styleIndex">cellXfs 样式索引；无样式时为 -1。cellXfs style index; -1 when the cell has no style.</param>
        /// <param name="formula">公式文本；无公式时为 null。Formula text, or null when the cell has no formula.</param>
        public ExcelCellData(int column, object value, int styleIndex, string formula)
        {
            Column = column;
            Value = value;
            StyleIndex = styleIndex;
            Formula = formula;
        }

        /// <summary>1 基列号。1-based column index.</summary>
        public int Column { get; private set; }
        /// <summary>单元格值：数字为 double、布尔为 bool、日期为 DateTime、文本为 string。Cell value: double for numbers, bool for booleans, DateTime for dates and string for text.</summary>
        public object Value { get; private set; }
        /// <summary>cellXfs 样式索引；-1 表示默认样式。cellXfs style index; -1 means the default style.</summary>
        public int StyleIndex { get; private set; }
        /// <summary>公式文本；null 表示普通值单元格。Formula text; null means a plain value cell.</summary>
        public string Formula { get; private set; }

        /// <summary>是否带有公式。Whether the cell carries a formula.</summary>
        public bool HasFormula { get { return !string.IsNullOrEmpty(Formula); } }
    }

    /// <summary>
    /// 流式读取得到的一行：行号、行高、隐藏标记与按列号升序排列的非空单元格。
    /// One row produced by streaming query: row index, height, hidden flag and non-empty cells sorted by column.
    /// </summary>
    public sealed class ExcelRowData
    {
        /// <summary>行内单元格列表（按列号升序）。Cells of this row, sorted by column.</summary>
        private readonly List<ExcelCellData> _cells;

        /// <summary>创建一行数据，仅供查询内部使用。Creates a row; intended for internal query use.</summary>
        /// <param name="rowIndex">1 基行号。1-based row index.</param>
        /// <param name="cells">按列号升序的单元格列表。Cells sorted by column.</param>
        /// <param name="hidden">该行是否隐藏。Whether the row is hidden.</param>
        /// <param name="height">行高（磅）；0 表示默认。Row height in points; 0 means default.</param>
        internal ExcelRowData(int rowIndex, List<ExcelCellData> cells, bool hidden, double height)
        {
            RowIndex = rowIndex;
            _cells = cells;
            Hidden = hidden;
            Height = height;
        }

        /// <summary>1 基行号。1-based row index.</summary>
        public int RowIndex { get; private set; }
        /// <summary>该行是否隐藏。Whether the row is hidden.</summary>
        public bool Hidden { get; private set; }
        /// <summary>行高（磅）；0 表示使用默认行高。Row height in points; 0 means the default height.</summary>
        public double Height { get; private set; }
        /// <summary>行内非空单元格（按列号升序）。Non-empty cells of the row, sorted by column.</summary>
        public IReadOnlyList<ExcelCellData> Cells { get { return _cells; } }
        /// <summary>非空单元格数量。Number of non-empty cells.</summary>
        public int Count { get { return _cells.Count; } }
        /// <summary>最小列号；空行返回 0。Smallest column index, or 0 for an empty row.</summary>
        public int FirstColumn { get { return _cells.Count == 0 ? 0 : _cells[0].Column; } }
        /// <summary>最大列号；空行返回 0。Largest column index, or 0 for an empty row.</summary>
        public int LastColumn { get { return _cells.Count == 0 ? 0 : _cells[_cells.Count - 1].Column; } }

        /// <summary>按列号取单元格值；不存在或为空时返回 null。Gets the value at a column, returning null when the cell is absent or empty.</summary>
        /// <param name="column">1 基列号。1-based column index.</param>
        public object this[int column]
        {
            get
            {
                object value;
                return TryGetValue(column, out value) ? value : null;
            }
        }

        /// <summary>按列号查单元格值（二分查找）。Looks up a cell value by column (binary search).</summary>
        /// <param name="column">1 基列号。1-based column index.</param>
        /// <param name="value">找到的单元格值；未找到时为 null。The value when found; null otherwise.</param>
        /// <returns>该列存在单元格时返回 true。True when a cell exists at the column.</returns>
        public bool TryGetValue(int column, out object value)
        {
            int low = 0;
            int high = _cells.Count - 1;
            while (low <= high)
            {
                int middle = low + ((high - low) >> 1);
                int current = _cells[middle].Column;
                if (current == column)
                {
                    value = _cells[middle].Value;
                    return true;
                }
                if (current < column) low = middle + 1;
                else high = middle - 1;
            }
            value = null;
            return false;
        }
    }

    /// <summary>
    /// MiniExcel 风格的低内存流式查询：不构建整本工作簿模型，逐行读取工作表并即时转换，
    /// 适合百万行级数据导出/校验；共享字符串表仍会整体加载（与 MiniExcel 语义一致）。
    /// MiniExcel-style low-memory streaming query: no whole-workbook model is built; the worksheet is read
    /// row by row and values are converted on the fly, which suits million-row export/validation. The shared
    /// string table is still loaded in full, matching MiniExcel semantics.
    /// </summary>
    public static class ExcelQuery
    {
        /// <summary>流式读取时逐节点比较用的主命名空间 URI。Main namespace URI compared per node while streaming.</summary>
        private static readonly string MainNamespace = Xlsx.Main.NamespaceName;

        /// <summary>从 .xlsx 文件逐行读取指定工作表（默认第一个非隐藏表）。Streams the specified worksheet (the first non-hidden sheet by default) from an .xlsx file.</summary>
        /// <param name="path">.xlsx 文件路径。Path of the .xlsx file.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <returns>按行号升序的数据行；只在枚举期间持有文件句柄。Data rows in ascending row order; the file stays open only while enumerating.</returns>
        public static IEnumerable<ExcelRowData> ReadRows(string path, string sheetName = null)
        {
            if (path == null) throw new ArgumentNullException("path");
            return ReadRowsFromPath(path, sheetName);
        }

        /// <summary>从流逐行读取指定工作表；不可定位的流会被复制到内存。Streams the specified worksheet from a stream; non-seekable streams are copied into memory.</summary>
        /// <param name="stream">包含 .xlsx 内容的流；调用方负责关闭。Stream containing the .xlsx content; the caller owns it.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <returns>按行号升序的数据行。Data rows in ascending row order.</returns>
        public static IEnumerable<ExcelRowData> ReadRows(Stream stream, string sheetName = null)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            return ReadRowsFromStream(stream, sheetName);
        }

        /// <summary>从字节数组逐行读取指定工作表。Streams the specified worksheet from a byte array.</summary>
        /// <param name="data">.xlsx 包的完整字节。Complete bytes of the .xlsx package.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <returns>按行号升序的数据行。Data rows in ascending row order.</returns>
        public static IEnumerable<ExcelRowData> ReadRows(byte[] data, string sheetName = null)
        {
            if (data == null) throw new ArgumentNullException("data");
            return ReadRowsFromBytes(data, sheetName);
        }

        /// <summary>流式查询为对象集合：默认用首行表头按名称匹配成员，空行与空值被跳过。Streams a worksheet into an object collection: the first row is matched by header name by default, and empty rows/values are skipped.</summary>
        /// <typeparam name="T">元素类型，需有无参构造函数。The element type; it needs a parameterless constructor.</typeparam>
        /// <param name="path">.xlsx 文件路径。Path of the .xlsx file.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <param name="options">行列位置选项；null 使用默认值。Row/column position options; defaults are used when null.</param>
        /// <returns>逐行反序列化的对象序列。Objects deserialized row by row.</returns>
        public static IEnumerable<T> Query<T>(string path, string sheetName = null, ExcelCollectionOptions options = null) where T : new()
        {
            if (path == null) throw new ArgumentNullException("path");
            ValidateOptions(options);
            return QueryRows<T>(ReadRowsFromPath(path, sheetName), options ?? new ExcelCollectionOptions());
        }

        /// <summary>从流式查询为对象集合；不可定位的流会被复制到内存。Streams a worksheet from a stream into an object collection; non-seekable streams are copied into memory.</summary>
        /// <typeparam name="T">元素类型，需有无参构造函数。The element type; it needs a parameterless constructor.</typeparam>
        /// <param name="stream">包含 .xlsx 内容的流；调用方负责关闭。Stream containing the .xlsx content; the caller owns it.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <param name="options">行列位置选项；null 使用默认值。Row/column position options; defaults are used when null.</param>
        /// <returns>逐行反序列化的对象序列。Objects deserialized row by row.</returns>
        public static IEnumerable<T> Query<T>(Stream stream, string sheetName = null, ExcelCollectionOptions options = null) where T : new()
        {
            if (stream == null) throw new ArgumentNullException("stream");
            ValidateOptions(options);
            return QueryRows<T>(ReadRowsFromStream(stream, sheetName), options ?? new ExcelCollectionOptions());
        }

        /// <summary>从字节数组流式查询为对象集合。Streams a worksheet from a byte array into an object collection.</summary>
        /// <typeparam name="T">元素类型，需有无参构造函数。The element type; it needs a parameterless constructor.</typeparam>
        /// <param name="data">.xlsx 包的完整字节。Complete bytes of the .xlsx package.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <param name="options">行列位置选项；null 使用默认值。Row/column position options; defaults are used when null.</param>
        /// <returns>逐行反序列化的对象序列。Objects deserialized row by row.</returns>
        public static IEnumerable<T> Query<T>(byte[] data, string sheetName = null, ExcelCollectionOptions options = null) where T : new()
        {
            if (data == null) throw new ArgumentNullException("data");
            ValidateOptions(options);
            return QueryRows<T>(ReadRowsFromBytes(data, sheetName), options ?? new ExcelCollectionOptions());
        }

        /// <summary>流式读取为字典序列：键取自首行表头，重复表头自动加序号，空值不放入字典。Streams a worksheet into dictionaries: keys come from the header row, duplicate headers get a suffix and empty values are omitted.</summary>
        /// <param name="path">.xlsx 文件路径。Path of the .xlsx file.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <param name="options">行列位置选项；null 使用默认值。Row/column position options; defaults are used when null.</param>
        /// <returns>每行一个字典的序列（键不区分大小写）。One dictionary per row (case-insensitive keys).</returns>
        public static IEnumerable<Dictionary<string, object>> ReadDictionaries(string path, string sheetName = null, ExcelCollectionOptions options = null)
        {
            if (path == null) throw new ArgumentNullException("path");
            ValidateOptions(options);
            return DictionaryRows(ReadRowsFromPath(path, sheetName), options ?? new ExcelCollectionOptions());
        }

        /// <summary>从流式读取为字典序列；不可定位的流会被复制到内存。Streams a worksheet from a stream into dictionaries; non-seekable streams are copied into memory.</summary>
        /// <param name="stream">包含 .xlsx 内容的流；调用方负责关闭。Stream containing the .xlsx content; the caller owns it.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <param name="options">行列位置选项；null 使用默认值。Row/column position options; defaults are used when null.</param>
        /// <returns>每行一个字典的序列（键不区分大小写）。One dictionary per row (case-insensitive keys).</returns>
        public static IEnumerable<Dictionary<string, object>> ReadDictionaries(Stream stream, string sheetName = null, ExcelCollectionOptions options = null)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            ValidateOptions(options);
            return DictionaryRows(ReadRowsFromStream(stream, sheetName), options ?? new ExcelCollectionOptions());
        }

        /// <summary>从字节数组流式读取为字典序列。Streams a worksheet from a byte array into dictionaries.</summary>
        /// <param name="data">.xlsx 包的完整字节。Complete bytes of the .xlsx package.</param>
        /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
        /// <param name="options">行列位置选项；null 使用默认值。Row/column position options; defaults are used when null.</param>
        /// <returns>每行一个字典的序列（键不区分大小写）。One dictionary per row (case-insensitive keys).</returns>
        public static IEnumerable<Dictionary<string, object>> ReadDictionaries(byte[] data, string sheetName = null, ExcelCollectionOptions options = null)
        {
            if (data == null) throw new ArgumentNullException("data");
            ValidateOptions(options);
            return DictionaryRows(ReadRowsFromBytes(data, sheetName), options ?? new ExcelCollectionOptions());
        }

        /// <summary>校验行列位置选项。Validates the row/column position options.</summary>
        /// <param name="options">待校验的选项；null 视为默认值。Options to validate; null is treated as defaults.</param>
        private static void ValidateOptions(ExcelCollectionOptions options)
        {
            if (options == null) return;
            if (options.StartRow < 1) throw new ArgumentOutOfRangeException("options");
            if (options.StartColumn < 1) throw new ArgumentOutOfRangeException("options");
        }

        /// <summary>打开文件并委托核心读取；迭代器结束时关闭文件。Opens the file and delegates to the core reader; the file is closed when enumeration ends.</summary>
        /// <param name="path">文件路径。File path.</param>
        /// <param name="sheetName">工作表名称。Worksheet name.</param>
        /// <returns>数据行序列。The data-row sequence.</returns>
        private static IEnumerable<ExcelRowData> ReadRowsFromPath(string path, string sheetName)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                foreach (ExcelRowData row in ReadRowsCore(stream, sheetName)) yield return row;
            }
        }

        /// <summary>保证流可定位后委托核心读取；需要复制时由内部分配并释放内存流。Ensures the stream is seekable and delegates to the core reader, copying into a memory stream when needed.</summary>
        /// <param name="stream">调用方提供的流。The caller-provided stream.</param>
        /// <param name="sheetName">工作表名称。Worksheet name.</param>
        /// <returns>数据行序列。The data-row sequence.</returns>
        private static IEnumerable<ExcelRowData> ReadRowsFromStream(Stream stream, string sheetName)
        {
            if (stream.CanSeek)
            {
                foreach (ExcelRowData row in ReadRowsCore(stream, sheetName)) yield return row;
                yield break;
            }
            MemoryStream memory = new MemoryStream();
            try
            {
                stream.CopyTo(memory);
                memory.Position = 0;
                foreach (ExcelRowData row in ReadRowsCore(memory, sheetName)) yield return row;
            }
            finally
            {
                memory.Dispose();
            }
        }

        /// <summary>用字节数组创建只读内存流并委托核心读取。Wraps the byte array in a read-only memory stream and delegates to the core reader.</summary>
        /// <param name="data">.xlsx 包字节。The .xlsx package bytes.</param>
        /// <param name="sheetName">工作表名称。Worksheet name.</param>
        /// <returns>数据行序列。The data-row sequence.</returns>
        private static IEnumerable<ExcelRowData> ReadRowsFromBytes(byte[] data, string sheetName)
        {
            using (MemoryStream memory = new MemoryStream(data, false))
            {
                foreach (ExcelRowData row in ReadRowsCore(memory, sheetName)) yield return row;
            }
        }

        /// <summary>
        /// 查询核心：打开 zip 包、加载样式/共享字符串并定位工作表部件，然后逐行流式解析 sheetData；
        /// 枚举结束（或提前释放枚举器）时关闭归档与部件流。
        /// Query core: opens the zip package, loads styles and shared strings, locates the worksheet part and then
        /// streams sheetData row by row; the archive and part stream are closed when enumeration ends or the
        /// enumerator is disposed early.
        /// </summary>
        /// <param name="stream">可定位的包流。The seekable package stream.</param>
        /// <param name="sheetName">工作表名称。Worksheet name.</param>
        /// <returns>数据行序列。The data-row sequence.</returns>
        private static IEnumerable<ExcelRowData> ReadRowsCore(Stream stream, string sheetName)
        {
            using (ZipArchive archive = PackageArchive.OpenForRead(stream, true))
            {
                XlsxPackage package = new XlsxPackage(archive);
                QuerySource source = QuerySource.Open(package, sheetName);
                Stream sheetStream = package.OpenRead(source.SheetPart);
                if (sheetStream == null) yield break;
                using (sheetStream)
                using (XmlReader reader = Xlsx.CreateSafeReader(sheetStream))
                {
                    foreach (ExcelRowData row in ReadSheetRows(reader, source)) yield return row;
                }
            }
        }

        /// <summary>在 worksheet XML 中定位 sheetData，并逐行产出非空数据行。Locates sheetData in the worksheet XML and yields each non-empty row.</summary>
        /// <param name="reader">工作表部件读取器。The worksheet-part reader.</param>
        /// <param name="source">查询上下文。The query context.</param>
        /// <returns>数据行序列。The data-row sequence.</returns>
        private static IEnumerable<ExcelRowData> ReadSheetRows(XmlReader reader, QuerySource source)
        {
            reader.MoveToContent();
            if (reader.NodeType != XmlNodeType.Element || reader.IsEmptyElement) yield break;
            int rootDepth = reader.Depth;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == rootDepth) yield break;
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.Depth != rootDepth + 1 || reader.LocalName != "sheetData") continue;
                if (reader.IsEmptyElement) yield break;
                int dataDepth = reader.Depth;
                int lastRowIndex = 0;
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == dataDepth) yield break;
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (reader.Depth != dataDepth + 1 || reader.LocalName != "row") continue;
                    ExcelRowData row = ReadRow(reader, source, ref lastRowIndex);
                    if (row != null) yield return row;
                }
                yield break;
            }
        }

        /// <summary>读取一行：行号、行高、隐藏标记与非空单元格；没有非空单元格时返回 null。Reads one row (index, height, hidden flag, non-empty cells), returning null when it has no non-empty cells.</summary>
        /// <param name="reader">定位在 row 起始标签的读取器。The reader positioned on the row start tag.</param>
        /// <param name="source">查询上下文。The query context.</param>
        /// <param name="lastRowIndex">上一行号，用于缺失 r 属性时续行。The previous row index, used when the r attribute is missing.</param>
        /// <returns>数据行或 null。The data row, or null.</returns>
        private static ExcelRowData ReadRow(XmlReader reader, QuerySource source, ref int lastRowIndex)
        {
            int rowIndex = Xlsx.ParseInt(reader.GetAttribute("r"), lastRowIndex + 1);
            if (rowIndex < 1 || rowIndex > CellAddress.MaxRowLiteral) rowIndex = lastRowIndex + 1;
            lastRowIndex = rowIndex;
            bool hidden = Xlsx.ParseBool(reader.GetAttribute("hidden"), false);
            double height = Xlsx.ParseDouble(reader.GetAttribute("ht"), 0);
            if (reader.IsEmptyElement) return null;

            List<ExcelCellData> cells = new List<ExcelCellData>();
            int rowDepth = reader.Depth;
            int lastColumnIndex = 0;
            bool sorted = true;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == rowDepth) break;
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.Depth != rowDepth + 1 || reader.LocalName != "c") continue;
                ExcelCellData cell;
                if (!TryReadCell(reader, source, ref lastColumnIndex, out cell)) continue;
                if (cells.Count > 0 && cells[cells.Count - 1].Column >= cell.Column) sorted = false;
                cells.Add(cell);
            }
            if (cells.Count == 0) return null;
            if (!sorted) cells.Sort(CompareCells);
            return new ExcelRowData(rowIndex, cells, hidden, height);
        }

        /// <summary>读取一个 c 元素；没有值也没有公式时返回 false。Reads one c element, returning false when it has neither a value nor a formula.</summary>
        /// <param name="reader">定位在 c 起始标签的读取器。The reader positioned on the cell start tag.</param>
        /// <param name="source">查询上下文。The query context.</param>
        /// <param name="lastColumnIndex">上一列号，用于缺失 r 属性时续列。The previous column index, used when the r attribute is missing.</param>
        /// <param name="cell">解析出的单元格。The parsed cell.</param>
        /// <returns>单元格应被纳入结果时返回 true。True when the cell belongs in the result.</returns>
        private static bool TryReadCell(XmlReader reader, QuerySource source, ref int lastColumnIndex, out ExcelCellData cell)
        {
            cell = default(ExcelCellData);
            string reference = reader.GetAttribute("r");
            string type = reader.GetAttribute("t");
            int styleIndex = Xlsx.ParseInt(reader.GetAttribute("s"), -1);
            int column = lastColumnIndex + 1;
            if (!string.IsNullOrEmpty(reference))
            {
                int row;
                int parsedColumn;
                try
                {
                    CellAddress.Parse(reference, out row, out parsedColumn);
                    column = parsedColumn;
                }
                catch (FormatException)
                {
                    column = lastColumnIndex + 1;
                }
            }
            if (column < 1) column = lastColumnIndex + 1;
            if (column > CellAddress.MaxColumnLiteral)
            {
                if (column > lastColumnIndex) lastColumnIndex = column;
                return false;
            }
            lastColumnIndex = column;

            string formula = null;
            string valueText = null;
            bool hasValue = false;
            string inlineText = null;
            if (!reader.IsEmptyElement)
            {
                int cellDepth = reader.Depth;
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == cellDepth) break;
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (reader.NamespaceURI != MainNamespace)
                    {
                        XlsxReader.SkipElementStreaming(reader);
                        continue;
                    }
                    switch (reader.LocalName)
                    {
                        case "f":
                            if (formula == null) formula = XlsxReader.ReadElementTextStreaming(reader);
                            else XlsxReader.SkipElementStreaming(reader);
                            break;
                        case "v":
                            if (!hasValue)
                            {
                                valueText = XlsxReader.ReadElementTextStreaming(reader);
                                hasValue = true;
                            }
                            else XlsxReader.SkipElementStreaming(reader);
                            break;
                        case "is":
                            if (inlineText == null) inlineText = XlsxReader.ReadRichTextStreaming(reader, MainNamespace);
                            else XlsxReader.SkipElementStreaming(reader);
                            break;
                        default:
                            XlsxReader.SkipElementStreaming(reader);
                            break;
                    }
                }
            }

            object value = null;
            if (type == "inlineStr" && inlineText != null)
            {
                value = inlineText;
            }
            else if (type == "s" && hasValue)
            {
                int index = Xlsx.ParseInt(valueText, -1);
                if (index >= 0 && index < source.SharedStrings.Count) value = source.SharedStrings[index];
            }
            else if (hasValue)
            {
                value = XlsxReader.ConvertValue(type, valueText, styleIndex, source.Styles, source.Date1904);
            }

            if (value == null && string.IsNullOrEmpty(formula)) return false;
            cell = new ExcelCellData(column, value, styleIndex, formula);
            return true;
        }

        /// <summary>按列号比较两个单元格，用于输入乱序时的一次性排序。Compares two cells by column, used for the one-time sort when the input order is broken.</summary>
        /// <param name="left">左单元格。Left cell.</param>
        /// <param name="right">右单元格。Right cell.</param>
        /// <returns>比较结果。The comparison result.</returns>
        private static int CompareCells(ExcelCellData left, ExcelCellData right)
        {
            return left.Column.CompareTo(right.Column);
        }

        /// <summary>把流式行序列转为对象序列：表头按名称匹配，找不到的成员保留默认值。Turns the row stream into objects: headers match by name and unmatched members keep their default values.</summary>
        /// <typeparam name="T">元素类型。The element type.</typeparam>
        /// <param name="rows">流式行序列。The streamed row sequence.</param>
        /// <param name="options">已校验的位置选项。The validated position options.</param>
        /// <returns>对象序列。The object sequence.</returns>
        private static IEnumerable<T> QueryRows<T>(IEnumerable<ExcelRowData> rows, ExcelCollectionOptions options) where T : new()
        {
            ExcelColumnMap[] maps = ExcelCollection.GetMaps(typeof(T));
            int[] columns = new int[maps.Length];
            bool mapped = !options.FirstRowIsHeader;
            FillPositionalColumns(columns, options.StartColumn);
            foreach (ExcelRowData row in rows)
            {
                if (row.RowIndex < options.StartRow) continue;
                if (!mapped)
                {
                    if (row.RowIndex == options.StartRow)
                    {
                        FillHeaderColumns(row, maps, options.StartColumn, columns);
                        mapped = true;
                        continue;
                    }
                    mapped = true;
                }
                T item = new T();
                bool any = false;
                for (int i = 0; i < maps.Length; i++)
                {
                    if (columns[i] < 1) continue;
                    object cellValue = row[columns[i]];
                    if (cellValue == null) continue;
                    object converted;
                    if (!ExcelCollection.TryConvertValue(cellValue, maps[i].ValueType, out converted)) continue;
                    maps[i].SetValue(item, converted);
                    any = true;
                }
                if (any) yield return item;
            }
        }

        /// <summary>填充按位置对应的列号（起始列 + 成员序号），超出网格的成员标记为 -1。Fills position-based columns (start column + member ordinal), marking members beyond the grid as -1.</summary>
        /// <param name="columns">待填充的列号数组。The column array to fill.</param>
        /// <param name="startColumn">起始列号。The start column.</param>
        private static void FillPositionalColumns(int[] columns, int startColumn)
        {
            for (int i = 0; i < columns.Length; i++)
            {
                int column = startColumn + i;
                columns[i] = column <= CellAddress.MaxColumnLiteral ? column : -1;
            }
        }

        /// <summary>按首行表头名称填充列号；找不到的成员标记为 -1，重复表头以首个为准。Fills columns from header names; unmatched members become -1 and the first duplicate header wins.</summary>
        /// <param name="header">表头行。The header row.</param>
        /// <param name="maps">成员的列映射。The member column mappings.</param>
        /// <param name="startColumn">起始列号。The start column.</param>
        /// <param name="columns">待填充的列号数组。The column array to fill.</param>
        private static void FillHeaderColumns(ExcelRowData header, ExcelColumnMap[] maps, int startColumn, int[] columns)
        {
            Dictionary<string, int> byName = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < header.Cells.Count; i++)
            {
                ExcelCellData cell = header.Cells[i];
                if (cell.Column < startColumn || cell.Value == null) continue;
                string name = Convert.ToString(cell.Value, CultureInfo.InvariantCulture);
                if (name == null) continue;
                name = name.Trim();
                if (name.Length == 0 || byName.ContainsKey(name)) continue;
                byName[name] = cell.Column;
            }
            for (int i = 0; i < maps.Length; i++)
            {
                int column;
                columns[i] = byName.TryGetValue(maps[i].Name, out column) ? column : -1;
            }
        }

        /// <summary>把流式行序列转为字典序列。Turns the row stream into dictionaries.</summary>
        /// <param name="rows">流式行序列。The streamed row sequence.</param>
        /// <param name="options">已校验的位置选项。The validated position options.</param>
        /// <returns>字典序列。The dictionary sequence.</returns>
        private static IEnumerable<Dictionary<string, object>> DictionaryRows(IEnumerable<ExcelRowData> rows, ExcelCollectionOptions options)
        {
            string[] names = null;
            int[] columns = null;
            foreach (ExcelRowData row in rows)
            {
                if (row.RowIndex < options.StartRow) continue;
                if (names == null)
                {
                    if (options.FirstRowIsHeader && row.RowIndex == options.StartRow)
                    {
                        names = BuildHeaderNames(row, options.StartColumn, out columns);
                        continue;
                    }
                    // 表头行缺失或按位置读取：使用 ColumnN 键，绝不把首条数据当成表头吞掉。
                    // The header row is missing or positional mode is active: use ColumnN keys and never
                    // consume the first data row as a header.
                    names = new string[0];
                    columns = new int[0];
                }

                Dictionary<string, object> result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                if (names.Length > 0)
                {
                    for (int i = 0; i < names.Length; i++)
                    {
                        object value = row[columns[i]];
                        if (value != null) result[names[i]] = value;
                    }
                }
                else
                {
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        ExcelCellData cell = row.Cells[i];
                        if (cell.Column < options.StartColumn || cell.Value == null) continue;
                        result["Column" + (cell.Column - options.StartColumn + 1).ToString(CultureInfo.InvariantCulture)] = cell.Value;
                    }
                }
                if (result.Count > 0) yield return result;
            }
        }

        /// <summary>用表头行生成不重复的列名与对应列号；空表头按相对位置命名为 ColumnN。Builds unique column names and their columns from the header row; empty headers are named ColumnN by relative position.</summary>
        /// <param name="header">表头行。The header row.</param>
        /// <param name="startColumn">起始列号。The start column.</param>
        /// <param name="columns">生成的列号数组。The generated column array.</param>
        /// <returns>生成的列名数组。The generated name array.</returns>
        private static string[] BuildHeaderNames(ExcelRowData header, int startColumn, out int[] columns)
        {
            List<string> names = new List<string>();
            List<int> positions = new List<int>();
            HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < header.Cells.Count; i++)
            {
                ExcelCellData cell = header.Cells[i];
                if (cell.Column < startColumn) continue;
                string name = cell.Value == null ? null : Convert.ToString(cell.Value, CultureInfo.InvariantCulture);
                if (name != null) name = name.Trim();
                if (string.IsNullOrEmpty(name))
                {
                    name = "Column" + (cell.Column - startColumn + 1).ToString(CultureInfo.InvariantCulture);
                }
                string unique = name;
                int suffix = 2;
                while (!used.Add(unique))
                {
                    unique = name + suffix.ToString(CultureInfo.InvariantCulture);
                    suffix++;
                }
                names.Add(unique);
                positions.Add(cell.Column);
            }
            columns = positions.ToArray();
            return names.ToArray();
        }

        /// <summary>
        /// 一次查询所需的包上下文：工作表部件路径、样式表、共享字符串表与日期系统。
        /// Package context of one query: worksheet part path, style table, shared-string table and date system.
        /// </summary>
        private sealed class QuerySource
        {
            /// <summary>工作表部件路径。Worksheet part path.</summary>
            public string SheetPart;
            /// <summary>样式表。Style table.</summary>
            public XlsxStyles Styles;
            /// <summary>共享字符串表。Shared-string table.</summary>
            public List<string> SharedStrings;
            /// <summary>是否使用 1904 日期系统。Whether the 1904 date system is used.</summary>
            public bool Date1904;

            /// <summary>打开工作簿元数据，定位目标工作表并加载样式与共享字符串。Opens the workbook metadata, locates the target sheet and loads styles and shared strings.</summary>
            /// <param name="package">包访问器。The package accessor.</param>
            /// <param name="sheetName">工作表名称；null 表示第一个非隐藏表。Worksheet name; null selects the first non-hidden sheet.</param>
            /// <returns>查询上下文。The query context.</returns>
            public static QuerySource Open(XlsxPackage package, string sheetName)
            {
                string workbookPart = package.GetRelationshipTargetByType(string.Empty, "/officeDocument");
                if (string.IsNullOrEmpty(workbookPart)) workbookPart = "xl/workbook.xml";
                XDocument workbookXml = package.GetXml(workbookPart);
                if (workbookXml == null || workbookXml.Root == null)
                {
                    throw new InvalidDataException("The .xlsx package does not contain a workbook part.");
                }

                QuerySource source = new QuerySource();
                XElement properties = Xlsx.Child(workbookXml.Root, "workbookPr");
                source.Date1904 = Xlsx.ParseBool(Xlsx.Attribute(properties, "date1904"), false);
                source.Styles = new XlsxStyles(package.GetXml(package.GetRelationshipTargetByType(workbookPart, "/styles")));
                source.SharedStrings = LoadSharedStrings(package,
                    package.GetRelationshipTargetByType(workbookPart, "/sharedStrings"));
                source.SheetPart = ResolveSheetPart(package, workbookPart, workbookXml, sheetName);
                return source;
            }

            /// <summary>按名称选择工作表；名称为空时选择第一个非隐藏表，找不到时抛出 ArgumentException。Selects a sheet by name; an empty name selects the first non-hidden sheet and an unknown name throws ArgumentException.</summary>
            /// <param name="package">包访问器。The package accessor.</param>
            /// <param name="workbookPart">workbook 部件路径。The workbook part path.</param>
            /// <param name="workbookXml">workbook 文档。The workbook document.</param>
            /// <param name="sheetName">目标表名；null 表示第一个非隐藏表。Target sheet name; null means the first non-hidden sheet.</param>
            /// <returns>工作表部件路径。The worksheet part path.</returns>
            private static string ResolveSheetPart(XlsxPackage package, string workbookPart, XDocument workbookXml, string sheetName)
            {
                XElement sheets = Xlsx.Child(workbookXml.Root, "sheets");
                string firstVisible = null;
                string firstName = null;
                List<string> available = new List<string>();
                foreach (XElement sheet in Xlsx.Children(sheets, "sheet"))
                {
                    string name = Xlsx.Attribute(sheet, "name") ?? string.Empty;
                    string relationshipId = Xlsx.Attribute(sheet, Xlsx.R + "id");
                    string state = Xlsx.Attribute(sheet, "state");
                    bool hidden = state == "hidden" || state == "veryHidden";
                    string part = package.GetRelationshipTarget(workbookPart, relationshipId);
                    if (string.IsNullOrEmpty(part)) continue;
                    available.Add(name);
                    if (firstName == null) firstName = part;
                    if (!hidden && firstVisible == null) firstVisible = part;
                    if (sheetName != null && string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase)) return part;
                }
                if (sheetName == null)
                {
                    return firstVisible ?? firstName;
                }
                throw new ArgumentException("Worksheet '" + sheetName + "' was not found. Available: "
                    + string.Join(", ", available.ToArray()), "sheetName");
            }
        }

        /// <summary>
        /// 流式加载共享字符串表：只保留 si 的直接 t 与 r/t 文本（与 DOM 读取语义一致），
        /// 不为整张表构建 XML 树，避免大表在查询期间占用双份内存。
        /// Streams the shared-string table: only the direct t and r/t text of each si is kept (matching the DOM
        /// reader), so no XML tree is materialized and a large table does not double the memory held by a query.
        /// </summary>
        /// <param name="package">包访问器。The package accessor.</param>
        /// <param name="partPath">sharedStrings 部件路径；可为 null。The sharedStrings part path; may be null.</param>
        /// <returns>共享字符串列表。The shared-string list.</returns>
        private static List<string> LoadSharedStrings(XlsxPackage package, string partPath)
        {
            List<string> result = new List<string>();
            if (string.IsNullOrEmpty(partPath)) return result;
            Stream stream = package.OpenRead(partPath);
            if (stream == null) return result;
            using (stream)
            using (XmlReader reader = Xlsx.CreateSafeReader(stream))
            {
                reader.MoveToContent();
                if (reader.NodeType != XmlNodeType.Element || reader.IsEmptyElement) return result;
                int rootDepth = reader.Depth;
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == rootDepth) break;
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (reader.Depth != rootDepth + 1 || reader.LocalName != "si") continue;
                    result.Add(ReadSharedItem(reader));
                }
            }
            return result;
        }

        /// <summary>读取一个 si 元素：拼接直接 t 与 r 运行块中的 t，忽略注音等其它子元素。Reads one si element by concatenating its direct t and the t of each r run, ignoring phonetic and other children.</summary>
        /// <param name="reader">定位在 si 起始标签的读取器。The reader positioned on the si start tag.</param>
        /// <returns>拼接后的文本。The concatenated text.</returns>
        private static string ReadSharedItem(XmlReader reader)
        {
            if (reader.IsEmptyElement) return string.Empty;
            int depth = reader.Depth;
            StringBuilder builder = new StringBuilder();
            string container = null;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement)
                {
                    if (reader.Depth == depth) break;
                    if (reader.Depth == depth + 1) container = null;
                    continue;
                }
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.Depth == depth + 1)
                {
                    container = reader.LocalName;
                    if (container == "t") builder.Append(XlsxReader.ReadElementTextStreaming(reader));
                    continue;
                }
                if (reader.Depth == depth + 2 && container == "r" && reader.LocalName == "t")
                {
                    builder.Append(XlsxReader.ReadElementTextStreaming(reader));
                }
            }
            return builder.ToString();
        }
    }
}
