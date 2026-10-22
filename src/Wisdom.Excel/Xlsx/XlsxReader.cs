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
    /// <summary>将 SpreadsheetML（.xlsx）包读取为统一工作簿模型。 Reads a SpreadsheetML (.xlsx) package into the unified workbook model.</summary>
    public static class XlsxReader
    {
        /// <summary>流式读取时逐节点比较用的主命名空间 URI。 The main namespace URI compared per node while streaming.</summary>
        private static readonly string MainNamespace = Xlsx.Main.NamespaceName;

        /// <summary>从文件路径加载工作簿（包含全部隐藏内容）。 Loads a workbook from a file path, including all hidden content.</summary>
        /// <param name="path">.xlsx 文件路径。 .xlsx file path.</param>
        /// <returns>解析后的工作簿模型。 The parsed workbook model.</returns>
        public static OfficeWorkbook Load(string path)
        {
            return Load(path, null);
        }

        /// <summary>从文件路径按指定选项加载工作簿。 Loads a workbook from a file path using the given options.</summary>
        /// <param name="path">.xlsx 文件路径。 .xlsx file path.</param>
        /// <param name="options">读取选项；null 时使用默认值。 Read options; defaults are used when null.</param>
        /// <returns>解析后的工作簿模型。 The parsed workbook model.</returns>
        public static OfficeWorkbook Load(string path, ExcelReadOptions options)
        {
            if (path == null) throw new ArgumentNullException("path");
            return Load(File.ReadAllBytes(path), options);
        }

        /// <summary>从流加载工作簿（包含全部隐藏内容）；流会被完整读入内存。 Loads a workbook from a stream, including all hidden content; the stream is read fully into memory.</summary>
        /// <param name="stream">包含 .xlsx 内容的流。 Stream containing the .xlsx content.</param>
        /// <returns>解析后的工作簿模型。 The parsed workbook model.</returns>
        public static OfficeWorkbook Load(Stream stream)
        {
            return Load(stream, null);
        }

        /// <summary>按指定选项从流加载工作簿；流会被完整读入内存。 Loads a workbook from a stream using the given options; the stream is read fully into memory.</summary>
        /// <param name="stream">包含 .xlsx 内容的流。 Stream containing the .xlsx content.</param>
        /// <param name="options">读取选项；null 时使用默认值。 Read options; defaults are used when null.</param>
        /// <returns>解析后的工作簿模型。 The parsed workbook model.</returns>
        public static OfficeWorkbook Load(Stream stream, ExcelReadOptions options)
        {
            if (stream == null) throw new ArgumentNullException("stream");
            using (MemoryStream memory = new MemoryStream())
            {
                stream.CopyTo(memory);
                return Load(memory.ToArray(), options);
            }
        }

        /// <summary>从字节数组加载工作簿（包含全部隐藏内容）。 Loads a workbook from a byte array, including all hidden content.</summary>
        /// <param name="data">.xlsx 包的完整字节。 Complete bytes of the .xlsx package.</param>
        /// <returns>解析后的工作簿模型。 The parsed workbook model.</returns>
        public static OfficeWorkbook Load(byte[] data)
        {
            return Load(data, null);
        }

        /// <summary>
        /// 按指定选项从字节数组加载工作簿。检测到 OLE2 头（.xls）或非法 zip 头时抛出 InvalidDataException。
        /// Loads a workbook from a byte array using the given options. Throws InvalidDataException when an
        /// OLE2 header (.xls) or an invalid zip header is detected.
        /// </summary>
        /// <param name="data">.xlsx 包的完整字节。 Complete bytes of the .xlsx package.</param>
        /// <param name="options">读取选项；null 时使用默认值。 Read options; defaults are used when null.</param>
        /// <returns>解析后的工作簿模型。 The parsed workbook model.</returns>
        /// <exception cref="System.IO.InvalidDataException">字节不是合法 .xlsx 包，或实为旧版 .xls 文件。 The bytes are not a valid .xlsx package, or the file is actually a legacy .xls workbook.</exception>
        public static OfficeWorkbook Load(byte[] data, ExcelReadOptions options)
        {
            if (data == null) throw new ArgumentNullException("data");
            if (data.Length >= 8 && data[0] == 0xD0 && data[1] == 0xCF && data[2] == 0x11 && data[3] == 0xE0)
            {
                throw new InvalidDataException("The file is a legacy .xls workbook. Use XlsReader instead.");
            }
            if (data.Length < 4 || data[0] != 0x50 || data[1] != 0x4B)
            {
                throw new InvalidDataException("The file is not a valid .xlsx package.");
            }

            try
            {
                using (MemoryStream memory = new MemoryStream(data, false))
                using (ZipArchive archive = PackageArchive.OpenForRead(memory, false))
                {
                    XlsxPackage package = new XlsxPackage(archive);
                    return ReadPackage(package, options ?? new ExcelReadOptions());
                }
            }
            catch (IOException ex)
            {
                throw new InvalidDataException("The .xlsx package is corrupt or truncated.", ex);
            }
        }

        /// <summary>
        /// 提取工作簿全部文本：逐表输出 [表名] 标题行，每行单元格以制表符分隔。
        /// 单元格按稀疏存储，这里沿有序单元格序列逐行输出，空列只按列差补制表符，
        /// 因此一个写到远端坐标的单元格不会把提取拖成行列全网格扫描。
        /// Extracts all workbook text: each sheet is emitted with a [name] header line and tab-separated cell
        /// values per row. Cells are stored sparsely, so this walks the ordered cell sequence and pads empty
        /// columns by the column gap; a single cell written to a distant coordinate cannot turn extraction into
        /// a full row-by-column grid scan.
        /// </summary>
        /// <param name="path">.xlsx 文件路径。 .xlsx file path.</param>
        /// <returns>纯文本内容。 The plain-text content.</returns>
        public static string ExtractText(string path)
        {
            OfficeWorkbook workbook = Load(path);
            StringBuilder builder = new StringBuilder();
            for (int s = 0; s < workbook.Worksheets.Count; s++)
            {
                Worksheet sheet = workbook.Worksheets[s];
                if (builder.Length > 0) builder.AppendLine();
                builder.Append('[').Append(sheet.Name).Append(']').AppendLine();
                AppendSheetText(builder, sheet);
            }
            return builder.ToString();
        }

        /// <summary>
        /// 只遍历已存在的单元格来拼接一张表的文本：无单元格的空行不产生输出行，行内每空一列补一个制表符，
        /// 行尾多余制表符被裁掉。相比逐行逐列扫描整张网格，这避免了稀疏表按 行数×列数 的空转。
        /// Appends the text of one sheet by walking only the cells that exist: rows with no cells emit no line,
        /// each skipped column inside a row emits one tab, and a row's trailing tabs are trimmed. Unlike sweeping
        /// the whole grid this avoids spinning over rows x columns on sparse sheets.
        /// </summary>
        /// <param name="builder">接收文本的输出缓冲。The buffer receiving the text.</param>
        /// <param name="sheet">待输出的工作表。The sheet to emit.</param>
        public static void AppendSheetText(StringBuilder builder, Worksheet sheet)
        {
            int currentRow = 0;
            int previousColumn = 1;
            StringBuilder line = new StringBuilder();
            foreach (Cell cell in sheet.Cells)
            {
                if (cell.Row != currentRow)
                {
                    if (line.Length > 0) builder.AppendLine(line.ToString().TrimEnd('\t'));
                    line.Clear();
                    currentRow = cell.Row;
                    previousColumn = 1;
                }
                int gap = cell.Column > previousColumn ? cell.Column - previousColumn : 0;
                for (int i = 0; i < gap; i++) line.Append('\t');
                previousColumn = cell.Column;
                string text = cell.DisplayText();
                if (!string.IsNullOrEmpty(text)) line.Append(text);
            }
            if (line.Length > 0) builder.AppendLine(line.ToString().TrimEnd('\t'));
        }

        /// <summary>
        /// 读取包的核心流程：定位 workbook 部件，加载样式表与共享字符串表、文档元数据和
        /// date1904 设置，再按 sheet 关系逐个读取工作表；按选项跳过隐藏表并剔除隐藏行/列单元格。
        /// Core package read flow: locates the workbook part, loads the style table, shared string table,
        /// document metadata and the date1904 setting, then reads each worksheet through its relationship;
        /// hidden sheets are skipped and hidden row/column cells removed according to the options.
        /// </summary>
        /// <param name="package">.xlsx 包访问器。 .xlsx package accessor.</param>
        /// <param name="options">读取选项。 Read options.</param>
        /// <returns>填充完成的工作簿模型。 The populated workbook model.</returns>
        private static OfficeWorkbook ReadPackage(XlsxPackage package, ExcelReadOptions options)
        {
            string workbookPart = package.GetRelationshipTargetByType(string.Empty, "/officeDocument");
            if (string.IsNullOrEmpty(workbookPart)) workbookPart = "xl/workbook.xml";
            XDocument workbookXml = package.GetXml(workbookPart);
            if (workbookXml == null || workbookXml.Root == null)
            {
                throw new InvalidDataException("The .xlsx package does not contain a workbook part.");
            }

            string stylesPart = package.GetRelationshipTargetByType(workbookPart, "/styles");
            string sharedStringsPart = package.GetRelationshipTargetByType(workbookPart, "/sharedStrings");
            XDocument stylesXml = package.GetXml(stylesPart);
            XlsxStyles styles = new XlsxStyles(stylesXml);
            List<string> sharedStrings = ReadSharedStrings(package.GetXml(sharedStringsPart));

            OfficeWorkbook workbook = new OfficeWorkbook();
            workbook.DefaultStyle = styles.DefaultStyle.Clone();
            ReadMetadata(package, workbook.Metadata);

            XElement workbookProperties = Xlsx.Child(workbookXml.Root, "workbookPr");
            bool date1904 = Xlsx.ParseBool(Xlsx.Attribute(workbookProperties, "date1904"), false);

            XElement sheetsElement = Xlsx.Child(workbookXml.Root, "sheets");
            XElement definedNames = Xlsx.Child(workbookXml.Root, "definedNames");
            Dictionary<int, string> printAreas = ReadPrintAreas(definedNames);
            Dictionary<int, Worksheet> loadedSheets = new Dictionary<int, Worksheet>();
            int workbookIndex = 0;
            foreach (XElement sheetElement in Xlsx.Children(sheetsElement, "sheet"))
            {
                string name = Xlsx.Attribute(sheetElement, "name");
                string relationshipId = Xlsx.Attribute(sheetElement, Xlsx.R + "id");
                string state = Xlsx.Attribute(sheetElement, "state");
                bool hidden = state == "hidden" || state == "veryHidden";
                int sheetIndex = workbookIndex;
                workbookIndex++;
                if (hidden && !options.IncludeHiddenWorksheets) continue;
                string sheetPart = package.GetRelationshipTarget(workbookPart, relationshipId);
                if (string.IsNullOrEmpty(sheetPart)) continue;

                Worksheet sheet = workbook.AddWorksheet(string.IsNullOrEmpty(name) ? "Sheet" + (workbook.Worksheets.Count + 1) : name);
                sheet.IsHidden = hidden;

                // 工作表部件按流解析：sheetData 逐格直接落到模型，其余顶层元素保留为骨架供装饰读取。
                // The worksheet part is parsed as a stream: sheetData lands cell-by-cell in the model while the
                // remaining top-level elements are kept as a skeleton for the decoration reader.
                XElement sheetRoot;
                try
                {
                    sheetRoot = ReadWorksheetPart(package, sheetPart, sheet, styles, sharedStrings, date1904);
                }
                catch (XmlException)
                {
                    workbook.Worksheets.Remove(sheet);
                    continue;
                }
                if (sheetRoot == null)
                {
                    workbook.Worksheets.Remove(sheet);
                    continue;
                }

                string printArea;
                if (printAreas.TryGetValue(sheetIndex, out printArea))
                {
                    try { sheet.SetPrintArea(printArea); }
                    catch (FormatException) { }
                    catch (ArgumentOutOfRangeException) { }
                }
                ReadWorksheet(sheet, sheetRoot, styles, sharedStrings, date1904, package, sheetPart);
                XlsxDecorationReader.Read(package, sheet, sheetPart, sheetRoot, stylesXml);
                loadedSheets[sheetIndex] = sheet;
                if (!options.IncludeHiddenRows || !options.IncludeHiddenColumns)
                {
                    RemoveHiddenCells(sheet, options);
                }
            }

            ReadDefinedNames(definedNames, loadedSheets, workbook);
            return workbook;
        }

        /// <summary>读取 autoFilter：区域引用与各列的取值条件。Reads the autoFilter: the range reference and the per-column value conditions.</summary>
        /// <param name="sheet">目标工作表。Target worksheet.</param>
        /// <param name="root">worksheet 根元素。The worksheet root element.</param>
        private static void ReadAutoFilter(Worksheet sheet, XElement root)
        {
            XElement autoFilter = Xlsx.Child(root, "autoFilter");
            if (autoFilter == null) return;
            string reference = Xlsx.Attribute(autoFilter, "ref");
            if (string.IsNullOrEmpty(reference)) return;
            ExcelAutoFilter filter = new ExcelAutoFilter(reference);
            foreach (XElement columnElement in Xlsx.Children(autoFilter, "filterColumn"))
            {
                int columnId = Xlsx.IntAttribute(columnElement, "colId", -1);
                if (columnId < 0) continue;
                XElement filters = Xlsx.Child(columnElement, "filters");
                if (filters == null) continue;
                ExcelFilterColumn column = filter.ForColumn(columnId);
                column.IncludeBlanks = Xlsx.ParseBool(Xlsx.Attribute(filters, "blank"), false);
                foreach (XElement valueElement in Xlsx.Children(filters, "filter"))
                {
                    string value = Xlsx.Attribute(valueElement, "val");
                    if (value != null) column.Values.Add(value);
                }
            }
            sheet.AutoFilter = filter;
        }

        /// <summary>读取 dataValidations：类型、运算符、公式、提示与错误文本，多区域引用只保留第一段。Reads dataValidations: type, operator, formulas, prompts and error texts; only the first area of a multi-area reference is kept.</summary>
        /// <param name="sheet">目标工作表。Target worksheet.</param>
        /// <param name="root">worksheet 根元素。The worksheet root element.</param>
        private static void ReadDataValidations(Worksheet sheet, XElement root)
        {
            XElement container = Xlsx.Child(root, "dataValidations");
            foreach (XElement element in Xlsx.Children(container, "dataValidation"))
            {
                DataValidationType type;
                if (!TryParseValidationType(Xlsx.Attribute(element, "type"), out type)) continue;
                string reference = Xlsx.Attribute(element, "sqref");
                if (string.IsNullOrEmpty(reference)) continue;
                int separator = reference.IndexOf(' ');
                if (separator > 0) reference = reference.Substring(0, separator);
                ExcelDataValidation validation = new ExcelDataValidation();
                validation.Range = reference;
                validation.Type = type;
                validation.Operator = ParseValidationOperator(Xlsx.Attribute(element, "operator"));
                validation.AllowBlank = Xlsx.ParseBool(Xlsx.Attribute(element, "allowBlank"), false);
                validation.ShowDropDown = !Xlsx.ParseBool(Xlsx.Attribute(element, "showDropDown"), false);
                validation.ShowInputMessage = Xlsx.ParseBool(Xlsx.Attribute(element, "showInputMessage"), false);
                validation.ShowErrorMessage = Xlsx.ParseBool(Xlsx.Attribute(element, "showErrorMessage"), false);
                string errorStyle = Xlsx.Attribute(element, "errorStyle");
                if (errorStyle == "warning") validation.ErrorStyle = DataValidationErrorStyle.Warning;
                else if (errorStyle == "information") validation.ErrorStyle = DataValidationErrorStyle.Information;
                validation.PromptTitle = Xlsx.Attribute(element, "promptTitle") ?? string.Empty;
                validation.Prompt = Xlsx.Attribute(element, "prompt") ?? string.Empty;
                validation.ErrorTitle = Xlsx.Attribute(element, "errorTitle") ?? string.Empty;
                validation.ErrorMessage = Xlsx.Attribute(element, "error") ?? string.Empty;
                XElement formula1 = Xlsx.Child(element, "formula1");
                if (formula1 != null) validation.Formula1 = formula1.Value;
                XElement formula2 = Xlsx.Child(element, "formula2");
                if (formula2 != null) validation.Formula2 = formula2.Value;
                sheet.DataValidations.Add(validation);
            }
        }

        /// <summary>解析数据验证类型名称；未知类型返回 false。Parses a validation type name, returning false for unknown types.</summary>
        /// <param name="value">type 属性值。The type attribute value.</param>
        /// <param name="type">输出类型。The parsed type.</param>
        /// <returns>识别成功返回 true。True when the name is recognized.</returns>
        private static bool TryParseValidationType(string value, out DataValidationType type)
        {
            type = DataValidationType.None;
            if (string.IsNullOrEmpty(value)) return false;
            switch (value)
            {
                case "whole": type = DataValidationType.Whole; return true;
                case "decimal": type = DataValidationType.Decimal; return true;
                case "list": type = DataValidationType.List; return true;
                case "date": type = DataValidationType.Date; return true;
                case "time": type = DataValidationType.Time; return true;
                case "textLength": type = DataValidationType.TextLength; return true;
                case "custom": type = DataValidationType.Custom; return true;
                default: return false;
            }
        }

        /// <summary>解析数据验证运算符名称，未知或缺失时按 between 处理。Parses a validation operator name, falling back to between when unknown or missing.</summary>
        /// <param name="value">operator 属性值。The operator attribute value.</param>
        /// <returns>运算符。The operator.</returns>
        private static DataValidationOperator ParseValidationOperator(string value)
        {
            switch (value)
            {
                case "notBetween": return DataValidationOperator.NotBetween;
                case "equal": return DataValidationOperator.Equal;
                case "notEqual": return DataValidationOperator.NotEqual;
                case "lessThan": return DataValidationOperator.LessThan;
                case "lessThanOrEqual": return DataValidationOperator.LessThanOrEqual;
                case "greaterThan": return DataValidationOperator.GreaterThan;
                case "greaterThanOrEqual": return DataValidationOperator.GreaterThanOrEqual;
                default: return DataValidationOperator.Between;
            }
        }

        /// <summary>读取超链接：外部地址来自 External 关系，表内位置写成 "#" 前缀；超大区域只还原左上角单元格。Reads hyperlinks: external addresses come from External relationships, in-workbook locations get a "#" prefix, and an oversized range is restored on its top-left cell only.</summary>
        /// <param name="sheet">目标工作表。Target worksheet.</param>
        /// <param name="root">worksheet 根元素。The worksheet root element.</param>
        /// <param name="package">包访问器。The package accessor.</param>
        /// <param name="sheetPart">工作表部件路径。The worksheet part path.</param>
        private static void ReadHyperlinks(Worksheet sheet, XElement root, XlsxPackage package, string sheetPart)
        {
            XElement container = Xlsx.Child(root, "hyperlinks");
            if (container == null) return;
            const int MaxCellsPerLink = 65536;
            foreach (XElement element in Xlsx.Children(container, "hyperlink"))
            {
                string reference = Xlsx.Attribute(element, "ref");
                if (string.IsNullOrEmpty(reference)) continue;
                string target = null;
                string relationshipId = Xlsx.Attribute(element, Xlsx.R + "id");
                if (!string.IsNullOrEmpty(relationshipId))
                {
                    target = package == null ? null : package.GetRelationshipTarget(sheetPart, relationshipId, true);
                }
                else
                {
                    string location = Xlsx.Attribute(element, "location");
                    if (!string.IsNullOrEmpty(location)) target = "#" + location;
                }
                if (string.IsNullOrEmpty(target)) continue;
                string tooltip = Xlsx.Attribute(element, "tooltip");

                int row1, column1, row2, column2;
                try
                {
                    CellAddress.ParseRange(reference, out row1, out column1, out row2, out column2);
                }
                catch (FormatException)
                {
                    continue;
                }
                catch (ArgumentOutOfRangeException)
                {
                    continue;
                }
                if (row1 > row2) { int swap = row1; row1 = row2; row2 = swap; }
                if (column1 > column2) { int swap = column1; column1 = column2; column2 = swap; }
                if ((long)(row2 - row1 + 1) * (column2 - column1 + 1) > MaxCellsPerLink)
                {
                    row2 = row1;
                    column2 = column1;
                }
                for (int row = row1; row <= row2; row++)
                {
                    for (int column = column1; column <= column2; column++)
                    {
                        Cell cell = sheet.GetCell(row, column);
                        if (!string.IsNullOrEmpty(cell.Hyperlink)) continue;
                        cell.Hyperlink = target;
                        cell.HyperlinkTooltip = tooltip;
                    }
                }
            }
        }

        /// <summary>读取工作簿定义名：跳过内置的 _xlnm 名称，局部名称按已加载工作表重新定位。Reads workbook defined names: built-in _xlnm names are skipped and local scopes are re-mapped onto the loaded worksheets.</summary>
        /// <param name="definedNames">workbook.xml 的 definedNames 元素。The definedNames element of workbook.xml.</param>
        /// <param name="loadedSheets">原始表索引到已加载工作表的映射。Map from original sheet index to loaded worksheet.</param>
        /// <param name="workbook">目标工作簿。Target workbook.</param>
        private static void ReadDefinedNames(XElement definedNames, Dictionary<int, Worksheet> loadedSheets, OfficeWorkbook workbook)
        {
            if (definedNames == null) return;
            foreach (XElement defined in Xlsx.Children(definedNames, "definedName"))
            {
                string name = Xlsx.Attribute(defined, "name");
                if (string.IsNullOrEmpty(name)) continue;
                if (name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) continue;
                string reference = defined.Value == null ? string.Empty : defined.Value.Trim();
                if (reference.Length == 0) continue;
                ExcelDefinedName definedName = new ExcelDefinedName(name, reference);
                int localSheetId = Xlsx.IntAttribute(defined, "localSheetId", -1);
                if (localSheetId >= 0)
                {
                    Worksheet sheet;
                    if (!loadedSheets.TryGetValue(localSheetId, out sheet)) continue;
                    definedName.LocalSheetId = workbook.Worksheets.IndexOf(sheet);
                }
                workbook.DefinedNames.Add(definedName);
            }
        }

        /// <summary>从 definedNames 中收集各表打印区域，键为工作表索引，多区域引用只取第一段。 Collects per-sheet print areas from definedNames keyed by worksheet index; for multi-area references only the first segment is kept.</summary>
        /// <param name="definedNames">workbook.xml 的 definedNames 元素。 The definedNames element of workbook.xml.</param>
        /// <returns>工作表索引到打印区域引用的映射。 Map from worksheet index to print-area reference.</returns>
        private static Dictionary<int, string> ReadPrintAreas(XElement definedNames)
        {
            Dictionary<int, string> result = new Dictionary<int, string>();
            if (definedNames == null) return result;
            foreach (XElement defined in Xlsx.Children(definedNames, "definedName"))
            {
                string name = Xlsx.Attribute(defined, "name");
                if (name == null) continue;
                if (!string.Equals(name, "Print_Area", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(name, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase)) continue;
                int sheetIndex = Xlsx.IntAttribute(defined, "localSheetId", -1);
                if (sheetIndex < 0 || result.ContainsKey(sheetIndex)) continue;
                string reference = defined.Value == null ? string.Empty : defined.Value.Trim();
                int comma = reference.IndexOf(',');
                if (comma >= 0) reference = reference.Substring(0, comma).Trim();
                if (reference.Length > 0) result[sheetIndex] = reference;
            }
            return result;
        }

        /// <summary>读取共享字符串表（sharedStrings.xml），按 si 顺序返回列表。 Reads the shared string table (sharedStrings.xml) and returns the strings in si order.</summary>
        /// <param name="document">sharedStrings.xml 文档；为 null 时返回空列表。 The sharedStrings.xml document; an empty list is returned when null.</param>
        /// <returns>共享字符串列表。 The list of shared strings.</returns>
        private static List<string> ReadSharedStrings(XDocument document)
        {
            List<string> result = new List<string>();
            if (document == null || document.Root == null) return result;
            foreach (XElement item in Xlsx.Children(document.Root, "si"))
            {
                result.Add(ReadRichText(item));
            }
            return result;
        }

        /// <summary>读取富文本元素：拼接直接 t 子元素与全部 r 运行块中的文本。 Reads a rich-text element by concatenating the direct t child with the text of every r run.</summary>
        /// <param name="element">si 或 is 元素。 An si or is element.</param>
        /// <returns>合并后的纯文本。 The concatenated plain text.</returns>
        private static string ReadRichText(XElement element)
        {
            StringBuilder builder = new StringBuilder();
            XElement text = Xlsx.Child(element, "t");
            if (text != null) builder.Append(text.Value);
            XElement runs = element;
            foreach (XElement run in runs.Elements(Xlsx.Main + "r"))
            {
                XElement runText = Xlsx.Child(run, "t");
                if (runText != null) builder.Append(runText.Value);
            }
            return builder.ToString();
        }

        /// <summary>
        /// 读取单个工作表部件的表级设置：视图/分页设置、列宽与隐藏列、
        /// 合并区域以及打印设置（单元格与行高已由流式读取完成）。
        /// Reads the sheet-level settings of one worksheet part: views and page-break settings, column widths and
        /// hidden columns, merged ranges and print setup (cells and row heights are handled by the streaming read).
        /// </summary>
        /// <param name="sheet">目标工作表模型。 Target worksheet model.</param>
        /// <param name="root">worksheet XML 根元素。 Root element of the worksheet XML.</param>
        /// <param name="styles">已解析的样式表。 The parsed style table.</param>
        /// <param name="sharedStrings">共享字符串表。 The shared string table.</param>
        /// <param name="date1904">是否使用 1904 日期系统。 Whether the 1904 date system is used.</param>
        /// <param name="package">包访问器（解析超链接关系）。 The package accessor (resolves hyperlink relationships).</param>
        /// <param name="sheetPart">工作表部件路径。 The worksheet part path.</param>
        private static void ReadWorksheet(
            Worksheet sheet,
            XElement root,
            XlsxStyles styles,
            List<string> sharedStrings,
            bool date1904,
            XlsxPackage package,
            string sheetPart)
        {
            XElement properties = Xlsx.Child(root, "sheetPr");
            if (properties != null)
            {
                XElement pageSetupProperties = Xlsx.Child(properties, "pageSetUpPr");
                sheet.FitToPage = Xlsx.ParseBool(Xlsx.Attribute(pageSetupProperties, "fitToPage"), false);
                XElement tabColor = Xlsx.Child(properties, "tabColor");
                OfficeColor color;
                if (tabColor != null && Xlsx.TryParseColor(Xlsx.Attribute(tabColor, "rgb"), out color))
                {
                    sheet.TabColor = color;
                }
            }

            XElement printOptions = Xlsx.Child(root, "printOptions");
            if (printOptions != null)
            {
                sheet.CenterHorizontally = Xlsx.ParseBool(Xlsx.Attribute(printOptions, "horizontalCentered"), false);
            }

            XElement formatProperties = Xlsx.Child(root, "sheetFormatPr");            if (formatProperties != null)
            {
                sheet.DefaultRowHeight = Xlsx.DoubleAttribute(formatProperties, "defaultRowHeight", sheet.DefaultRowHeight);
                sheet.DefaultColumnWidth = Xlsx.DoubleAttribute(formatProperties, "defaultColWidth", sheet.DefaultColumnWidth);
            }

            XElement views = Xlsx.Child(root, "sheetViews");
            XElement view = Xlsx.Child(views, "sheetView");
            if (view != null)
            {
                sheet.ShowGridLines = Xlsx.ParseBool(Xlsx.Attribute(view, "showGridLines"), true);
                sheet.ShowZeros = Xlsx.ParseBool(Xlsx.Attribute(view, "showZeros"), true);
                XElement pane = Xlsx.Child(view, "pane");
                if (pane != null)
                {
                    string state = Xlsx.Attribute(pane, "state");
                    if (state == null || state == "frozen" || state == "frozenSplit")
                    {
                        sheet.FrozenRows = Math.Max(0, Xlsx.IntAttribute(pane, "ySplit", 0));
                        sheet.FrozenColumns = Math.Max(0, Xlsx.IntAttribute(pane, "xSplit", 0));
                    }
                }
            }

            XElement columns = Xlsx.Child(root, "cols");
            foreach (XElement column in Xlsx.Children(columns, "col"))
            {
                int min = Xlsx.IntAttribute(column, "min", 1);
                int max = Xlsx.IntAttribute(column, "max", min);
                double width = Xlsx.DoubleAttribute(column, "width", sheet.DefaultColumnWidth);
                bool hidden = Xlsx.ParseBool(Xlsx.Attribute(column, "hidden"), false);
                if (min < 1) min = 1;
                if (max < min) max = min;
                if (max > 16384) max = 16384;
                for (int c = min; c <= max; c++)
                {
                    sheet.SetColumnWidth(c, width);
                    if (hidden) sheet.SetColumnHidden(c, true);
                }
            }

            // sheetData 的单元格与行高已由 ReadWorksheetPart 流式读取，这里只处理其余表级设置。
            // Cell values and row heights from sheetData were streamed by ReadWorksheetPart; only the remaining
            // sheet-level settings are handled here.
            XElement mergeCells = Xlsx.Child(root, "mergeCells");
            foreach (XElement mergeCell in Xlsx.Children(mergeCells, "mergeCell"))
            {
                string reference = Xlsx.Attribute(mergeCell, "ref");
                if (string.IsNullOrEmpty(reference)) continue;
                try
                {
                    sheet.Merge(reference);
                }
                catch (FormatException)
                {
                }
            }

            ReadAutoFilter(sheet, root);
            ReadDataValidations(sheet, root);
            ReadHyperlinks(sheet, root, package, sheetPart);
            ReadPrintSetup(sheet, root);
        }

        /// <summary>
        /// 流式读取工作表部件：sheetData 的行列直接落到模型，其余顶层元素读入骨架根返回，
        /// 使十万级单元格的工作表不再物化整棵 XML 树。
        /// Reads a worksheet part as a stream: rows and cells from sheetData land directly in the model while the
        /// remaining top-level elements are captured into a skeleton root, so a sheet with hundreds of thousands of
        /// cells no longer materializes a whole XML tree.
        /// </summary>
        /// <param name="package">包访问器。 The package accessor.</param>
        /// <param name="sheetPart">工作表部件路径。 The worksheet part path.</param>
        /// <param name="sheet">要填充的工作表。 The worksheet to populate.</param>
        /// <param name="styles">样式表。 The style table.</param>
        /// <param name="sharedStrings">共享字符串表。 The shared string table.</param>
        /// <param name="date1904">是否使用 1904 日期系统。 Whether the 1904 date system is used.</param>
        /// <returns>不含 sheetData 的骨架根元素；部件缺失时返回 null（XmlException 由调用方处理）。 The skeleton root without sheetData, or null when the part is missing (XmlException propagates to the caller).</returns>
        private static XElement ReadWorksheetPart(XlsxPackage package, string sheetPart, Worksheet sheet,
            XlsxStyles styles, List<string> sharedStrings, bool date1904)
        {
            Stream input = package.OpenRead(sheetPart);
            if (input == null) return null;
            XElement root = new XElement(Xlsx.Main + "worksheet");
            using (input)
            using (XmlReader reader = Xlsx.CreateSafeReader(input))
            {
                reader.MoveToContent();
                if (reader.NodeType != XmlNodeType.Element) return null;
                int rootDepth = reader.Depth;
                if (reader.IsEmptyElement) return root;
                // ReadFrom 之后各框架停留的节点不同（有的停在结束标签，有的已越到下一个节点），
                // 因此读取一棵子树后先检查当前节点，再决定是否 Read。
                // Different frameworks park the reader differently after ReadFrom (on the end tag, or already on
                // the following node), so a subtree is followed by inspecting the current node before reading again.
                bool pending = false;
                while (pending || reader.Read())
                {
                    pending = false;
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (reader.Depth != rootDepth + 1) continue;
                    if (reader.LocalName == "sheetData")
                    {
                        ReadSheetDataStreaming(reader, sheet, styles, sharedStrings, date1904);
                        continue;
                    }
                    XElement captured = XElement.ReadFrom(reader) as XElement;
                    if (captured != null) root.Add(captured);
                    pending = true;
                }
            }
            return root;
        }

        /// <summary>逐行消费 sheetData，直到其结束标签。 Consumes sheetData row by row until its end tag.</summary>
        /// <param name="reader">定位在 sheetData 起始标签的读取器。 The reader positioned on the sheetData start tag.</param>
        /// <param name="sheet">要填充的工作表。 The worksheet to populate.</param>
        /// <param name="styles">样式表。 The style table.</param>
        /// <param name="sharedStrings">共享字符串表。 The shared string table.</param>
        /// <param name="date1904">是否使用 1904 日期系统。 Whether the 1904 date system is used.</param>
        private static void ReadSheetDataStreaming(XmlReader reader, Worksheet sheet, XlsxStyles styles,
            List<string> sharedStrings, bool date1904)
        {
            if (reader.IsEmptyElement) return;
            int dataDepth = reader.Depth;
            int lastRowIndex = 0;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == dataDepth) return;
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.Depth != dataDepth + 1 || reader.LocalName != "row") continue;
                ReadRowStreaming(reader, sheet, styles, sharedStrings, date1904, ref lastRowIndex);
            }
        }

        /// <summary>读取一行：行号、行高、隐藏标记与其中的单元格。 Reads one row: index, height, the hidden flag and its cells.</summary>
        /// <param name="reader">定位在 row 起始标签的读取器。 The reader positioned on the row start tag.</param>
        /// <param name="sheet">要填充的工作表。 The worksheet to populate.</param>
        /// <param name="styles">样式表。 The style table.</param>
        /// <param name="sharedStrings">共享字符串表。 The shared string table.</param>
        /// <param name="date1904">是否使用 1904 日期系统。 Whether the 1904 date system is used.</param>
        /// <param name="lastRowIndex">上一行号，用于缺失 r 属性时续行。 The previous row index, used to continue when the r attribute is missing.</param>
        private static void ReadRowStreaming(XmlReader reader, Worksheet sheet, XlsxStyles styles,
            List<string> sharedStrings, bool date1904, ref int lastRowIndex)
        {
            int rowIndex = Xlsx.ParseInt(reader.GetAttribute("r"), lastRowIndex + 1);
            if (rowIndex < 1) rowIndex = lastRowIndex + 1;
            lastRowIndex = rowIndex;
            double height = Xlsx.ParseDouble(reader.GetAttribute("ht"), 0);
            if (height > 0) sheet.SetRowHeight(rowIndex, height);
            if (Xlsx.ParseBool(reader.GetAttribute("hidden"), false)) sheet.SetRowHidden(rowIndex, true);

            if (reader.IsEmptyElement) return;
            int rowDepth = reader.Depth;
            int lastColumnIndex = 0;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == rowDepth) return;
                if (reader.NodeType != XmlNodeType.Element) continue;
                if (reader.Depth != rowDepth + 1 || reader.LocalName != "c") continue;
                ReadCellStreaming(reader, sheet, styles, sharedStrings, date1904, rowIndex, ref lastColumnIndex);
            }
        }

        /// <summary>
        /// 读取一个单元格元素：解析 r/s/t 属性与 f、v、is 子元素，再按原有语义写入模型。
        /// Reads one cell element: the r/s/t attributes plus its f, v and is children, then writes into the model
        /// with the same semantics as the document-based path.
        /// </summary>
        /// <param name="reader">定位在 c 起始标签的读取器。 The reader positioned on the cell start tag.</param>
        /// <param name="sheet">要填充的工作表。 The worksheet to populate.</param>
        /// <param name="styles">样式表。 The style table.</param>
        /// <param name="sharedStrings">共享字符串表。 The shared string table.</param>
        /// <param name="date1904">是否使用 1904 日期系统。 Whether the 1904 date system is used.</param>
        /// <param name="rowIndex">所属行号。 The owning row index.</param>
        /// <param name="lastColumnIndex">上一列号，用于缺失 r 属性时续列。 The previous column index, used to continue when the r attribute is missing.</param>
        private static void ReadCellStreaming(XmlReader reader, Worksheet sheet, XlsxStyles styles,
            List<string> sharedStrings, bool date1904, int rowIndex, ref int lastColumnIndex)
        {
            string reference = reader.GetAttribute("r");
            string type = reader.GetAttribute("t");
            int styleIndex = Xlsx.ParseInt(reader.GetAttribute("s"), -1);
            int row = rowIndex;
            int column = lastColumnIndex + 1;
            if (!string.IsNullOrEmpty(reference))
            {
                try
                {
                    CellAddress.Parse(reference, out row, out column);
                }
                catch (FormatException)
                {
                    row = rowIndex;
                    column = lastColumnIndex + 1;
                }
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
                        SkipElementStreaming(reader);
                        continue;
                    }
                    switch (reader.LocalName)
                    {
                        case "f":
                            if (formula == null) formula = ReadElementTextStreaming(reader);
                            else SkipElementStreaming(reader);
                            break;
                        case "v":
                            if (!hasValue)
                            {
                                valueText = ReadElementTextStreaming(reader);
                                hasValue = true;
                            }
                            else SkipElementStreaming(reader);
                            break;
                        case "is":
                            if (inlineText == null) inlineText = ReadRichTextStreaming(reader, MainNamespace);
                            else SkipElementStreaming(reader);
                            break;
                        default:
                            SkipElementStreaming(reader);
                            break;
                    }
                }
            }

            object cellValue = null;
            if (type == "inlineStr" && inlineText != null)
            {
                cellValue = inlineText;
            }
            else if (type == "s" && hasValue)
            {
                int index = Xlsx.ParseInt(valueText, -1);
                if (index >= 0 && index < sharedStrings.Count) cellValue = sharedStrings[index];
            }
            else if (hasValue)
            {
                cellValue = ConvertValue(type, valueText, styleIndex, styles, date1904);
            }

            bool hasFormula = !string.IsNullOrEmpty(formula);
            bool hasStyle = styleIndex >= 0;
            if (cellValue == null && !hasFormula && !hasStyle)
            {
                return;
            }

            Cell cell = sheet.GetCell(row, column);
            cell.Value = cellValue;
            if (hasFormula) cell.Formula = formula;
            if (hasStyle) cell.Style = styles.GetStyle(styleIndex);
        }

        /// <summary>消费 is 元素内的全部 t 文本（与基于 XDocument 的富文本读取等价）。 Consumes every t text inside an is element, equivalent to the XDocument-based rich-text read.</summary>
        /// <param name="reader">定位在 is 起始标签的读取器。 The reader positioned on the is start tag.</param>
        /// <param name="mainNamespace">主命名空间 URI。 The main namespace URI.</param>
        /// <returns>拼接后的文本。 The concatenated text.</returns>
        internal static string ReadRichTextStreaming(XmlReader reader, string mainNamespace)
        {
            StringBuilder builder = new StringBuilder();
            if (reader.IsEmptyElement) return builder.ToString();
            int depth = reader.Depth;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) return builder.ToString();
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t" && reader.NamespaceURI == mainNamespace)
                {
                    builder.Append(ReadElementTextStreaming(reader));
                    continue;
                }
            }
            return builder.ToString();
        }

        /// <summary>
        /// 读取纯文本元素的内容，并把读取器停在元素自身的结束标签上。
        /// 不使用 ReadElementContentAsString：它在 .NET Framework 与 .NET Core 上停留的位置不同，
        /// 会让外层按深度判断边界的循环漏掉结束标签。
        /// Reads the text of an element and parks the reader on that element's own end tag.
        /// ReadElementContentAsString is avoided on purpose: it leaves the reader at different positions on
        /// .NET Framework and .NET Core, which makes the depth-based boundary checks of the outer loops miss
        /// their end tags.
        /// </summary>
        /// <param name="reader">定位在元素起始标签的读取器。 The reader positioned on the element start tag.</param>
        /// <returns>元素文本内容。 The element text.</returns>
        internal static string ReadElementTextStreaming(XmlReader reader)
        {
            if (reader.IsEmptyElement) return string.Empty;
            int depth = reader.Depth;
            StringBuilder builder = new StringBuilder();
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) return builder.ToString();
                if (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.CDATA
                    || reader.NodeType == XmlNodeType.SignificantWhitespace)
                {
                    builder.Append(reader.Value);
                }
            }
            return builder.ToString();
        }

        /// <summary>跳过当前元素的剩余内容，读取器停在其结束标签上。 Skips the rest of the current element, leaving the reader on its end tag.</summary>
        /// <param name="reader">读取器。 The reader.</param>
        internal static void SkipElementStreaming(XmlReader reader)
        {
            if (reader.IsEmptyElement) return;
            int depth = reader.Depth;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) return;
            }
        }

        /// <summary>按选项移除位于隐藏行或隐藏列中的单元格（先收集后删除，避免遍历中修改集合）。 Removes cells located in hidden rows or columns according to the options (collected first to avoid mutating the collection while iterating).</summary>
        /// <param name="sheet">要清理的工作表。 Worksheet to clean up.</param>
        /// <param name="options">读取选项。 Read options.</param>
        private static void RemoveHiddenCells(Worksheet sheet, ExcelReadOptions options)
        {
            List<Cell> removed = new List<Cell>();
            foreach (Cell cell in sheet.Cells)
            {
                if (!options.IncludeHiddenRows && sheet.HiddenRows.Contains(cell.Row)) removed.Add(cell);
                else if (!options.IncludeHiddenColumns && sheet.HiddenColumns.Contains(cell.Column)) removed.Add(cell);
            }
            for (int i = 0; i < removed.Count; i++)
            {
                sheet.RemoveCell(removed[i].Row, removed[i].Column);
            }
        }

        /// <summary>
        /// 按单元格类型转换 v 元素文本：b 布尔、e/str 原样文本、d ISO 日期；
        /// 其余按数字解析，若样式为日期格式则把序列号转为 DateTime。
        /// Converts the v element text by cell type: b to boolean, e/str to verbatim text, d to an ISO date;
        /// everything else is parsed as a number, and when the style is a date format the serial
        /// is converted into a DateTime.
        /// </summary>
        /// <param name="type">单元格 t 属性值（可为 null）。 The cell t attribute value (may be null).</param>
        /// <param name="text">v 元素文本。 Text of the v element.</param>
        /// <param name="styleIndex">样式索引。 Style index.</param>
        /// <param name="styles">样式表。 Style table.</param>
        /// <param name="date1904">是否使用 1904 日期系统。 Whether the 1904 date system is used.</param>
        /// <returns>转换后的单元格值。 The converted cell value.</returns>
        internal static object ConvertValue(string type, string text, int styleIndex, XlsxStyles styles, bool date1904)
        {
            if (text == null) return null;
            if (type == "b")
            {
                return text == "1" || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
            }
            if (type == "e") return text;
            if (type == "str") return text;
            if (type == "d")
            {
                DateTime date;
                if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) return date;
                return text;
            }

            double number;
            if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
            {
                return text;
            }
            if (styles.IsDateFormat(styleIndex))
            {
                return FromSerial(number, date1904);
            }
            return number;
        }

        /// <summary>将 Excel 日期序列号转换为 DateTime（1904 系统基准 1904-01-01，否则 1899-12-30）。 Converts an Excel date serial into a DateTime (epoch 1904-01-01 for the 1904 system, otherwise 1899-12-30).</summary>
        /// <param name="serial">日期序列号；负数按 0 处理。 Date serial number; negative values are treated as 0.</param>
        /// <param name="date1904">是否使用 1904 日期系统。 Whether the 1904 date system is used.</param>
        /// <returns>对应的日期时间。 The corresponding date and time.</returns>
        private static DateTime FromSerial(double serial, bool date1904)
        {
            DateTime epoch = date1904 ? new DateTime(1904, 1, 1) : new DateTime(1899, 12, 30);
            if (serial < 0) serial = 0;
            return epoch.AddDays(serial);
        }

        /// <summary>读取 pageMargins 与 pageSetup：页边距按英寸换算为磅，按纸张编号近似页面尺寸并处理横向。 Reads pageMargins and pageSetup: margins are converted from inches to points and the page size is approximated from the paper number, honoring landscape orientation.</summary>
        /// <param name="sheet">目标工作表。 Target worksheet.</param>
        /// <param name="root">worksheet XML 根元素。 Root element of the worksheet XML.</param>
        private static void ReadPrintSetup(Worksheet sheet, XElement root)
        {
            PageSetup setup = sheet.PrintSetup;
            XElement margins = Xlsx.Child(root, "pageMargins");
            if (margins != null)
            {
                setup.MarginLeft = Xlsx.DoubleAttribute(margins, "left", setup.MarginLeft / 72.0) * 72.0;
                setup.MarginRight = Xlsx.DoubleAttribute(margins, "right", setup.MarginRight / 72.0) * 72.0;
                setup.MarginTop = Xlsx.DoubleAttribute(margins, "top", setup.MarginTop / 72.0) * 72.0;
                setup.MarginBottom = Xlsx.DoubleAttribute(margins, "bottom", setup.MarginBottom / 72.0) * 72.0;
                setup.HeaderDistance = Xlsx.DoubleAttribute(margins, "header", setup.HeaderDistance / 72.0) * 72.0;
                setup.FooterDistance = Xlsx.DoubleAttribute(margins, "footer", setup.FooterDistance / 72.0) * 72.0;
            }

            XElement pageSetup = Xlsx.Child(root, "pageSetup");
            if (pageSetup == null) return;
            int paperSize = Xlsx.IntAttribute(pageSetup, "paperSize", 9);
            double width = 595.28;
            double height = 841.89;
            if (paperSize == 1)
            {
                width = 612;
                height = 792;
            }
            else if (paperSize == 5)
            {
                width = 612;
                height = 1008;
            }
            else if (paperSize == 8)
            {
                width = 841.89;
                height = 1190.55;
            }
            string orientation = Xlsx.Attribute(pageSetup, "orientation");
            if (orientation == "landscape")
            {
                double swap = width;
                width = height;
                height = swap;
            }
            setup.Width = width;
            setup.Height = height;
        }

        /// <summary>从 docProps/core.xml 读取标题、作者、主题、关键词与创建/修改时间。 Reads title, author, subject, keywords and the created/modified timestamps from docProps/core.xml.</summary>
        /// <param name="package">.xlsx 包访问器。 .xlsx package accessor.</param>
        /// <param name="metadata">要填充的元数据对象。 Metadata object to populate.</param>
        private static void ReadMetadata(XlsxPackage package, DocumentMetadata metadata)
        {
            XDocument core = package.GetXml("docProps/core.xml");
            if (core == null || core.Root == null) return;

            XElement title = core.Root.Element(Xlsx.Dc + "title");
            if (title != null) metadata.Title = title.Value;
            XElement creator = core.Root.Element(Xlsx.Dc + "creator");
            if (creator != null) metadata.Author = creator.Value;
            XElement subject = core.Root.Element(Xlsx.Dc + "subject");
            if (subject != null) metadata.Subject = subject.Value;
            XElement keywords = core.Root.Element(Xlsx.CoreProperties + "keywords");
            if (keywords != null) metadata.Keywords = keywords.Value;

            DateTime created;
            if (TryParseTimestamp(core.Root.Element(Xlsx.DcTerms + "created"), out created)) metadata.Created = created;
            DateTime modified;
            if (TryParseTimestamp(core.Root.Element(Xlsx.DcTerms + "modified"), out modified)) metadata.Modified = modified;
        }

        /// <summary>解析 dcterms 时间戳元素（按 UTC 解释）。 Parses a dcterms timestamp element (interpreted as UTC).</summary>
        /// <param name="element">时间戳元素。 Timestamp element.</param>
        /// <param name="value">解析结果；失败时为 DateTime.MinValue。 Parsing result; DateTime.MinValue on failure.</param>
        /// <returns>解析成功返回 true。 True when parsing succeeded.</returns>
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
    }
}
