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
    /// <summary>将统一工作簿模型写出为 SpreadsheetML（.xlsx）包。 Writes the unified workbook model as a SpreadsheetML (.xlsx) package.</summary>
    public static class XlsxWriter
    {
        /// <summary>OPC 关系部件的内容类型。 Content type of OPC relationship parts.</summary>
        private const string ContentTypeRels = "application/vnd.openxmlformats-package.relationships+xml";
        /// <summary>普通 XML 部件的默认内容类型。 Default content type for plain XML parts.</summary>
        private const string ContentTypeXml = "application/xml";
        /// <summary>workbook 主部件的内容类型。 Content type of the workbook main part.</summary>
        private const string ContentTypeWorkbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        /// <summary>worksheet 部件的内容类型。 Content type of worksheet parts.</summary>
        private const string ContentTypeWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
        /// <summary>styles 部件的内容类型。 Content type of the styles part.</summary>
        private const string ContentTypeStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
        /// <summary>sharedStrings 部件的内容类型。 Content type of the sharedStrings part.</summary>
        private const string ContentTypeSharedStrings = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
        /// <summary>核心属性部件（docProps/core.xml）的内容类型。 Content type of the core properties part (docProps/core.xml).</summary>
        private const string ContentTypeCore = "application/vnd.openxmlformats-package.core-properties+xml";
        /// <summary>扩展属性部件（docProps/app.xml）的内容类型。 Content type of the extended properties part (docProps/app.xml).</summary>
        private const string ContentTypeApp = "application/vnd.openxmlformats-officedocument.extended-properties+xml";

        /// <summary>指向包主文档的关系类型。 Relationship type pointing at the package main document.</summary>
        private const string OfficeDocumentRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        /// <summary>核心属性关系类型。 Relationship type for core properties.</summary>
        private const string CorePropertiesRelationship = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        /// <summary>扩展属性关系类型。 Relationship type for extended properties.</summary>
        private const string ExtendedPropertiesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        /// <summary>工作表关系类型。 Relationship type for worksheets.</summary>
        private const string WorksheetRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
        /// <summary>样式表关系类型。 Relationship type for the style table.</summary>
        private const string StylesRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        /// <summary>共享字符串表关系类型。 Relationship type for the shared string table.</summary>
        private const string SharedStringsRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

        /// <summary>将工作簿写入指定路径的文件（覆盖已存在文件）。 Writes the workbook to a file at the given path, overwriting an existing file.</summary>
        /// <param name="workbook">要写出的工作簿模型。 Workbook model to write.</param>
        /// <param name="path">输出 .xlsx 文件路径。 Output .xlsx file path.</param>
        public static void Write(OfficeWorkbook workbook, string path)
        {
            if (path == null) throw new ArgumentNullException("path");
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                Write(workbook, stream);
            }
        }

        /// <summary>将工作簿写出的字节数组形式（内存 .xlsx 包）。 Writes the workbook as a byte array (an in-memory .xlsx package).</summary>
        /// <param name="workbook">要写出的工作簿模型。 Workbook model to write.</param>
        /// <returns>.xlsx 包的完整字节。 The complete bytes of the .xlsx package.</returns>
        public static byte[] Write(OfficeWorkbook workbook)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Write(workbook, stream);
                return stream.ToArray();
            }
        }

        /// <summary>将工作簿写入流：先准备表名/绘图部件，再把各部件逐个流式写入 zip 归档。 Writes the workbook to a stream: sheet names and decoration parts are prepared first, then every part is streamed into the zip archive.</summary>
        /// <param name="workbook">要写出的工作簿模型。 Workbook model to write.</param>
        /// <param name="stream">目标流；调用方负责关闭。 Target stream; the caller is responsible for closing it.</param>
        public static void Write(OfficeWorkbook workbook, Stream stream)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            if (stream == null) throw new ArgumentNullException("stream");
            WriterState state = new WriterState(workbook);
            state.Prepare();
            state.Save(stream);
        }

        /// <summary>
        /// 一次写出过程的内部状态：集中登记样式与共享字符串，并把所有部件写入 zip 包；
        /// 工作表部件用 XmlWriter 直接流式写入归档条目，不再构建 XML DOM。
        /// Internal state of one write operation: registers styles and shared strings with de-duplication
        /// and writes all parts into the zip package. Worksheet parts are streamed straight into the archive
        /// entries with an XmlWriter instead of being built as an XML DOM.
        /// </summary>
        private sealed class WriterState
        {
            /// <summary>待写出的工作簿模型。 Workbook model being written.</summary>
            private readonly OfficeWorkbook _workbook;
            /// <summary>样式登记器（fonts/fills/borders/numFmts/cellXfs 去重）。 Style registry de-duplicating fonts, fills, borders, number formats and cellXfs.</summary>
            private readonly StyleRegistry _styles;
            /// <summary>已登记的唯一共享字符串列表，下标即索引。 List of registered unique shared strings whose position is the index.</summary>
            private readonly List<string> _sharedStrings = new List<string>();
            /// <summary>字符串到共享字符串索引的查找表（区分大小写）。 Case-sensitive lookup from string value to shared-string index.</summary>
            private readonly Dictionary<string, int> _sharedStringIds = new Dictionary<string, int>(StringComparer.Ordinal);
            /// <summary>共享字符串引用总次数，写入 sst 的 count 属性。 Total shared-string references, written as the sst count attribute.</summary>
            private int _sharedStringReferences;
            /// <summary>样式对象引用到 cellXfs 索引的缓存（免重复拼接去重键）。 Cache mapping a CellStyle reference to its cellXfs index, avoiding repeated de-duplication key building.</summary>
            private readonly Dictionary<CellStyle, int> _styleIndexes = new Dictionary<CellStyle, int>();
            /// <summary>日期补格式场景下样式对象引用到索引的缓存。 Same cache for the date-format case where a cloned style with a default date format is registered.</summary>
            private readonly Dictionary<CellStyle, int> _dateStyleIndexes = new Dictionary<CellStyle, int>();
            /// <summary>空样式日期单元所用补格式索引，首次需要时登记；-1 表示尚未登记。 CellXfs index of the date format used by an unstyled date cell, registered on first need; -1 until then.</summary>
            private int _plainDateStyleIndex = -1;
            /// <summary>复用的单元格地址字符缓冲区，避免每个地址分配中间字符串。 Reusable character buffer for cell addresses, avoiding intermediate string allocations per address.</summary>
            private readonly char[] _addressBuffer = new char[24];
            /// <summary>索引值到其文本的懒分配缓存。 Lazily filled cache mapping an index value to its text.</summary>
            private string[] _indexTexts = new string[64];
            /// <summary>每个表的名称与 sheetId 二元组（按表顺序）。 Name and sheetId pair of each sheet, in sheet order.</summary>
            private readonly List<string[]> _sheetNames = new List<string[]>();
            /// <summary>活动表下标（第一个非隐藏表）。 Index of the active sheet (the first non-hidden one).</summary>
            private int _activeSheetIndex;
            /// <summary>图表与图片生成的附加部件。 Extra parts produced for charts and images.</summary>
            private XlsxDrawingParts _decorations;
            /// <summary>各表超链接与工作表关系部件。 Per-sheet hyperlinks and sheet relationship parts.</summary>
            private XlsxHyperlinkParts _hyperlinks;
            /// <summary>条件格式生成器（维护全局规则优先级）。 Conditional-format builder (it owns the global rule priorities).</summary>
            private readonly XlsxConditionalFormatBuilder _conditionalFormats = new XlsxConditionalFormatBuilder();

            /// <summary>基于待写出的工作簿创建写出状态。 Creates the write state for a workbook model.</summary>
            /// <param name="workbook">工作簿模型。 Workbook model.</param>
            public WriterState(OfficeWorkbook workbook)
            {
                _workbook = workbook;
                _styles = new StyleRegistry(workbook.DefaultStyle ?? new CellStyle());
            }

            /// <summary>准备阶段：规范化表名、确定首个可见表为活动表并构建绘图部件。 Prepare phase: normalizes sheet names, picks the first visible sheet as active and builds the decoration parts.</summary>
            public void Prepare()
            {
                List<string> names = CreateSheetNames();
                _activeSheetIndex = 0;
                for (int i = 0; i < _workbook.Worksheets.Count; i++)
                {
                    if (!_workbook.Worksheets[i].IsHidden)
                    {
                        _activeSheetIndex = i;
                        break;
                    }
                }
                _decorations = XlsxDrawingParts.Build(_workbook);
                _hyperlinks = XlsxHyperlinkParts.Build(_workbook, _decorations);
                for (int i = 0; i < names.Count; i++)
                {
                    _sheetNames.Add(new string[] { names[i], (i + 1).ToString(CultureInfo.InvariantCulture) });
                }
            }

            /// <summary>
            /// 保存阶段：按 OPC 结构写入各内容类型、关系与属性部件，随后把每个工作表流式写入归档，
            /// 最后写出在流式过程中才登记完整的样式表与共享字符串表。
            /// Save phase: writes the content types, relationships and properties parts following the OPC structure,
            /// streams each worksheet into the archive, and finally writes the style and shared-string tables,
            /// which are only complete after the streaming pass.
            /// </summary>
            /// <param name="stream">目标流。 Target stream.</param>
            public void Save(Stream stream)
            {
                using (ZipArchive archive = new ZipArchive(stream, ZipArchiveMode.Create, true))
                {
                    WriteEntry(archive, "[Content_Types].xml", Xlsx.SerializeXml(CreateContentTypes()));
                    WriteEntry(archive, "_rels/.rels", Xlsx.SerializeXml(CreateRootRelationships()));
                    WriteEntry(archive, "docProps/core.xml", Xlsx.SerializeXml(CreateCoreProperties()));
                    WriteEntry(archive, "docProps/app.xml", Xlsx.SerializeXml(CreateAppProperties()));
                    WriteEntry(archive, "xl/workbook.xml", Xlsx.SerializeXml(CreateWorkbook()));
                    WriteEntry(archive, "xl/_rels/workbook.xml.rels", Xlsx.SerializeXml(CreateWorkbookRelationships()));
                    for (int i = 0; i < _workbook.Worksheets.Count; i++)
                    {
                        string number = (i + 1).ToString(CultureInfo.InvariantCulture);
                        WriteWorksheet(archive, "xl/worksheets/sheet" + number + ".xml",
                            _workbook.Worksheets[i], i == _activeSheetIndex, i);
                        byte[] sheetRelationships = CreateSheetRelationships(i);
                        if (sheetRelationships != null)
                        {
                            WriteEntry(archive, "xl/worksheets/_rels/sheet" + number + ".xml.rels", sheetRelationships);
                        }
                    }
                    WriteEntry(archive, "xl/styles.xml", _styles.ToXml());
                    WriteEntry(archive, "xl/sharedStrings.xml", Xlsx.SerializeXml(CreateSharedStrings()));
                    if (_decorations != null)
                    {
                        for (int i = 0; i < _decorations.Entries.Count; i++)
                        {
                            WriteEntry(archive, _decorations.Entries[i].Key, _decorations.Entries[i].Value);
                        }
                    }
                }
            }

            /// <summary>向归档写入一个部件条目（最优压缩）。 Writes one part entry into the archive using optimal compression.</summary>
            /// <param name="archive">目标 zip 归档。 Target zip archive.</param>
            /// <param name="name">部件路径名。 Part path name.</param>
            /// <param name="data">部件字节内容。 Part byte content.</param>
            private static void WriteEntry(ZipArchive archive, string name, byte[] data)
            {
                ZipArchiveEntry entry = archive.CreateEntry(OfficePackageLimits.ValidateEntryName(name), CompressionLevel.Optimal);
                using (Stream stream = entry.Open())
                {
                    stream.Write(data, 0, data.Length);
                }
            }

            /// <summary>生成工作表的关系部件：绘图固定在 rId1，外部超链接依次使用后续 Id；两者皆无时返回 null。 Builds a sheet relationship part: the drawing keeps rId1 and external hyperlinks take the following ids; returns null when the sheet has neither.</summary>
            /// <param name="sheetIndex">表下标。 Sheet index.</param>
            /// <returns>关系部件字节或 null。 Relationship part bytes or null.</returns>
            private byte[] CreateSheetRelationships(int sheetIndex)
            {
                string drawingTarget = _decorations != null ? _decorations.GetDrawingTarget(sheetIndex) : null;
                XElement relationships = _hyperlinks != null
                    ? _hyperlinks.CreateRelationships(sheetIndex, drawingTarget)
                    : null;
                if (relationships == null) return null;
                return Xlsx.SerializeXml(new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), relationships));
            }

            /// <summary>为所有工作表生成合法且唯一的表名（去非法字符、限 31 字符、重名加数字后缀）。 Generates legal and unique names for all worksheets (illegal characters removed, capped at 31 characters, numeric suffix on duplicates).</summary>
            /// <returns>按原顺序排列的表名列表。 List of sheet names in the original order.</returns>
            private List<string> CreateSheetNames()
            {
                List<string> names = new List<string>();
                HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < _workbook.Worksheets.Count; i++)
                {
                    string name = SanitizeSheetName(_workbook.Worksheets[i].Name);
                    if (name.Length == 0) name = "Sheet" + (i + 1).ToString(CultureInfo.InvariantCulture);
                    string unique = name;
                    int suffix = 2;
                    while (!used.Add(unique))
                    {
                        string suffixText = suffix.ToString(CultureInfo.InvariantCulture);
                        int keep = Math.Min(name.Length, 31 - suffixText.Length);
                        unique = name.Substring(0, keep) + suffixText;
                        suffix++;
                    }
                    names.Add(unique);
                }
                return names;
            }

            /// <summary>移除表名中的非法字符（方括号、冒号、星号、问号、斜杠及控制符）并截断到 31 字符。 Removes illegal characters (brackets, colon, asterisk, question mark, slashes and control characters) from a sheet name and truncates it to 31 characters.</summary>
            /// <param name="name">原始表名。 Original sheet name.</param>
            /// <returns>清洗后的表名。 The sanitized sheet name.</returns>
            private static string SanitizeSheetName(string name)
            {
                if (string.IsNullOrEmpty(name)) return string.Empty;
                StringBuilder builder = new StringBuilder(name.Length);
                for (int i = 0; i < name.Length && builder.Length < 31; i++)
                {
                    char c = name[i];
                    if (c == '[' || c == ']' || c == ':' || c == '*' || c == '?' || c == '/' || c == '\\' || c < 32)
                    {
                        continue;
                    }
                    builder.Append(c);
                }
                return builder.ToString().Trim();
            }

            /// <summary>生成 xl/workbook.xml：表列表、隐藏状态、活动视图、打印区域定义名与强制重算设置。 Generates xl/workbook.xml: the sheet list, hidden states, active view, print-area defined names and the force-recalc setting.</summary>
            /// <returns>workbook 部件文档。 The workbook part document.</returns>
            private XDocument CreateWorkbook()
            {
                bool anyVisible = false;
                for (int i = 0; i < _workbook.Worksheets.Count; i++)
                {
                    if (!_workbook.Worksheets[i].IsHidden) anyVisible = true;
                }
                XElement sheets = new XElement(Xlsx.Main + "sheets");
                for (int i = 0; i < _sheetNames.Count; i++)
                {
                    XElement sheetElement = new XElement(Xlsx.Main + "sheet",
                        new XAttribute("name", _sheetNames[i][0]),
                        new XAttribute("sheetId", _sheetNames[i][1]),
                        new XAttribute(Xlsx.R + "id", "rId" + (i + 1).ToString(CultureInfo.InvariantCulture)));
                    if (_workbook.Worksheets[i].IsHidden && (anyVisible || i != 0))
                    {
                        sheetElement.Add(new XAttribute("state", "hidden"));
                    }
                    sheets.Add(sheetElement);
                }

                XElement root = new XElement(Xlsx.Main + "workbook",
                    new XAttribute(XNamespace.Xmlns + "r", Xlsx.R.NamespaceName),
                    new XElement(Xlsx.Main + "fileVersion", new XAttribute("appName", "xl")),
                    new XElement(Xlsx.Main + "workbookPr", new XAttribute("defaultThemeVersion", "164011")),
                    new XElement(Xlsx.Main + "bookViews",
                        new XElement(Xlsx.Main + "workbookView",
                            new XAttribute("xWindow", "0"),
                            new XAttribute("yWindow", "0"),
                            new XAttribute("windowWidth", "24000"),
                            new XAttribute("windowHeight", "12000"))),
                    sheets,
                    CreateDefinedNames(),
                    new XElement(Xlsx.Main + "calcPr",
                        new XAttribute("calcId", "191029"),
                        new XAttribute("fullCalcOnLoad", "1")));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>为设置了打印区域的工作表生成 _xlnm.Print_Area 定义名，并追加工作簿的全局/局部定义名。 Generates _xlnm.Print_Area defined names for worksheets that have a print area, then appends the workbook's global and local defined names.</summary>
            /// <returns>definedNames 元素或 null。 The definedNames element or null.</returns>
            private XElement CreateDefinedNames()
            {
                XElement container = null;
                for (int i = 0; i < _workbook.Worksheets.Count; i++)
                {
                    Worksheet sheet = _workbook.Worksheets[i];
                    if (!sheet.HasPrintArea) continue;
                    if (container == null) container = new XElement(Xlsx.Main + "definedNames");
                    string reference = QuoteSheetName(_sheetNames[i][0]) + "!"
                        + CellAddress.Format(sheet.PrintAreaFirstRow, sheet.PrintAreaFirstColumn) + ":"
                        + CellAddress.Format(sheet.PrintAreaLastRow, sheet.PrintAreaLastColumn);
                    container.Add(new XElement(Xlsx.Main + "definedName",
                        new XAttribute("name", "_xlnm.Print_Area"),
                        new XAttribute("localSheetId", i.ToString(CultureInfo.InvariantCulture)),
                        reference));
                }
                for (int i = 0; i < _workbook.DefinedNames.Count; i++)
                {
                    ExcelDefinedName name = _workbook.DefinedNames[i];
                    if (name == null || !IsValidDefinedName(name.Name) || string.IsNullOrEmpty(name.Reference)) continue;
                    if (container == null) container = new XElement(Xlsx.Main + "definedNames");
                    XElement element = new XElement(Xlsx.Main + "definedName",
                        new XAttribute("name", name.Name),
                        name.Reference);
                    if (name.LocalSheetId >= 0 && name.LocalSheetId < _workbook.Worksheets.Count)
                    {
                        element.Add(new XAttribute("localSheetId", name.LocalSheetId.ToString(CultureInfo.InvariantCulture)));
                    }
                    container.Add(element);
                }
                return container;
            }

            /// <summary>判断定义名是否合法：首字符为字母、下划线或反斜杠，后续为字母数字或 . _ \，且不是单元格引用。 Checks whether a defined name is legal: it starts with a letter, underscore or backslash, continues with letters, digits or . _ \, and is not a cell reference.</summary>
            /// <param name="name">定义名。 The defined name.</param>
            /// <returns>合法返回 true。 True when the name is legal.</returns>
            private static bool IsValidDefinedName(string name)
            {
                if (string.IsNullOrEmpty(name) || name.Length > 255) return false;
                char first = name[0];
                if (!char.IsLetter(first) && first != '_' && first != '\\') return false;
                for (int i = 1; i < name.Length; i++)
                {
                    char c = name[i];
                    if (!char.IsLetterOrDigit(c) && c != '.' && c != '_' && c != '\\') return false;
                }
                int row;
                int column;
                try
                {
                    CellAddress.Parse(name, out row, out column);
                    return false;
                }
                catch (FormatException)
                {
                    return true;
                }
            }

            /// <summary>按公式引用规则给表名加单引号（内部单引号翻倍转义）。 Quotes a sheet name for formula references, doubling embedded single quotes.</summary>
            /// <param name="name">表名。 Sheet name.</param>
            /// <returns>带引号的表名。 The quoted sheet name.</returns>
            private static string QuoteSheetName(string name)
            {
                return "'" + name.Replace("'", "''") + "'";
            }

            /// <summary>生成 workbook.xml.rels：各工作表部件加 styles 与 sharedStrings 的关系。 Generates workbook.xml.rels: relationships for each worksheet part plus styles and sharedStrings.</summary>
            /// <returns>关系部件文档。 The relationships part document.</returns>
            private XDocument CreateWorkbookRelationships()
            {
                XElement root = new XElement(Xlsx.PackageRels + "Relationships");
                int sheetCount = _sheetNames.Count;
                for (int i = 0; i < sheetCount; i++)
                {
                    root.Add(new XElement(Xlsx.PackageRels + "Relationship",
                        new XAttribute("Id", "rId" + (i + 1).ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("Type", WorksheetRelationship),
                        new XAttribute("Target", "worksheets/sheet" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml")));
                }
                root.Add(new XElement(Xlsx.PackageRels + "Relationship",
                    new XAttribute("Id", "rId" + (sheetCount + 1).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", StylesRelationship),
                    new XAttribute("Target", "styles.xml")));
                root.Add(new XElement(Xlsx.PackageRels + "Relationship",
                    new XAttribute("Id", "rId" + (sheetCount + 2).ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("Type", SharedStringsRelationship),
                    new XAttribute("Target", "sharedStrings.xml")));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>生成包级 _rels/.rels：指向 workbook、核心属性与扩展属性的三条关系。 Generates the package-level _rels/.rels with the three relationships pointing at the workbook, core properties and extended properties.</summary>
            /// <returns>关系部件文档。 The relationships part document.</returns>
            private XDocument CreateRootRelationships()
            {
                XElement root = new XElement(Xlsx.PackageRels + "Relationships",
                    new XElement(Xlsx.PackageRels + "Relationship",
                        new XAttribute("Id", "rId1"),
                        new XAttribute("Type", OfficeDocumentRelationship),
                        new XAttribute("Target", "xl/workbook.xml")),
                    new XElement(Xlsx.PackageRels + "Relationship",
                        new XAttribute("Id", "rId2"),
                        new XAttribute("Type", CorePropertiesRelationship),
                        new XAttribute("Target", "docProps/core.xml")),
                    new XElement(Xlsx.PackageRels + "Relationship",
                        new XAttribute("Id", "rId3"),
                        new XAttribute("Type", ExtendedPropertiesRelationship),
                        new XAttribute("Target", "docProps/app.xml")));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>生成 [Content_Types].xml：rels/xml 默认类型与各已知部件及工作表的覆盖类型。 Generates [Content_Types].xml: default rels/xml types plus override entries for each known part and worksheet.</summary>
            /// <returns>内容类型部件文档。 The content types part document.</returns>
            private XDocument CreateContentTypes()
            {
                XElement root = new XElement(Xlsx.ContentTypes + "Types",
                    new XElement(Xlsx.ContentTypes + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", ContentTypeRels)),
                    new XElement(Xlsx.ContentTypes + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", ContentTypeXml)),
                    new XElement(Xlsx.ContentTypes + "Override", new XAttribute("PartName", "/xl/workbook.xml"), new XAttribute("ContentType", ContentTypeWorkbook)),
                    new XElement(Xlsx.ContentTypes + "Override", new XAttribute("PartName", "/xl/styles.xml"), new XAttribute("ContentType", ContentTypeStyles)),
                    new XElement(Xlsx.ContentTypes + "Override", new XAttribute("PartName", "/xl/sharedStrings.xml"), new XAttribute("ContentType", ContentTypeSharedStrings)),
                    new XElement(Xlsx.ContentTypes + "Override", new XAttribute("PartName", "/docProps/core.xml"), new XAttribute("ContentType", ContentTypeCore)),
                    new XElement(Xlsx.ContentTypes + "Override", new XAttribute("PartName", "/docProps/app.xml"), new XAttribute("ContentType", ContentTypeApp)));
                for (int i = 0; i < _sheetNames.Count; i++)
                {
                    root.Add(new XElement(Xlsx.ContentTypes + "Override",
                        new XAttribute("PartName", "/xl/worksheets/sheet" + (i + 1).ToString(CultureInfo.InvariantCulture) + ".xml"),
                        new XAttribute("ContentType", ContentTypeWorksheet)));
                }
                if (_decorations != null) _decorations.AddContentTypes(root);
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>生成 xl/sharedStrings.xml：每个字符串一个 si/t 元素并保留空白。 Generates xl/sharedStrings.xml: one si/t element per string with whitespace preserved.</summary>
            /// <returns>共享字符串部件文档。 The shared strings part document.</returns>
            private XDocument CreateSharedStrings()
            {
                XElement root = new XElement(Xlsx.Main + "sst",
                    new XAttribute("count", _sharedStringReferences.ToString(CultureInfo.InvariantCulture)),
                    new XAttribute("uniqueCount", _sharedStrings.Count.ToString(CultureInfo.InvariantCulture)));
                for (int i = 0; i < _sharedStrings.Count; i++)
                {
                    root.Add(new XElement(Xlsx.Main + "si",
                        new XElement(Xlsx.Main + "t",
                            new XAttribute(XNamespace.Xml + "space", "preserve"),
                            _sharedStrings[i])));
                }
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>生成 docProps/core.xml：标题、主题、作者、关键词与创建/修改时间（W3CDTF 格式）。 Generates docProps/core.xml: title, subject, author, keywords and the created/modified timestamps in W3CDTF form.</summary>
            /// <returns>核心属性部件文档。 The core properties part document.</returns>
            private XDocument CreateCoreProperties()
            {
                DocumentMetadata metadata = _workbook.Metadata ?? new DocumentMetadata();
                XElement root = new XElement(Xlsx.CoreProperties + "coreProperties",
                    new XAttribute(XNamespace.Xmlns + "cp", Xlsx.CoreProperties.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "dc", Xlsx.Dc.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "dcterms", Xlsx.DcTerms.NamespaceName),
                    new XAttribute(XNamespace.Xmlns + "xsi", Xlsx.Xsi.NamespaceName),
                    new XElement(Xlsx.Dc + "title", metadata.Title ?? string.Empty),
                    new XElement(Xlsx.Dc + "subject", metadata.Subject ?? string.Empty),
                    new XElement(Xlsx.Dc + "creator", metadata.Author ?? string.Empty),
                    new XElement(Xlsx.CoreProperties + "keywords", metadata.Keywords ?? string.Empty),
                    new XElement(Xlsx.CoreProperties + "lastModifiedBy", metadata.Author ?? string.Empty),
                    new XElement(Xlsx.DcTerms + "created",
                        new XAttribute(Xlsx.Xsi + "type", "dcterms:W3CDTF"),
                        FormatTimestamp(metadata.Created)),
                    new XElement(Xlsx.DcTerms + "modified",
                        new XAttribute(Xlsx.Xsi + "type", "dcterms:W3CDTF"),
                        FormatTimestamp(metadata.Modified)));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>将时间转换为 UTC 的 W3CDTF 文本（yyyy-MM-ddTHH:mm:ssZ）。 Converts a timestamp to UTC W3CDTF text (yyyy-MM-ddTHH:mm:ssZ).</summary>
            /// <param name="value">待格式化的时间。 Timestamp to format.</param>
            /// <returns>ISO 8601 格式文本。 The ISO 8601 formatted text.</returns>
            private static string FormatTimestamp(DateTime value)
            {
                DateTime utc = value.Kind == DateTimeKind.Utc ? value : value.ToUniversalTime();
                return utc.ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
            }

            /// <summary>生成 docProps/app.xml：写入应用程序名称等最小扩展属性集。 Generates docProps/app.xml with a minimal extended property set such as the application name.</summary>
            /// <returns>扩展属性部件文档。 The extended properties part document.</returns>
            private XDocument CreateAppProperties()
            {
                XElement root = new XElement(Xlsx.ExtendedProperties + "Properties",
                    new XElement(Xlsx.ExtendedProperties + "Application", "Wisdom.Office"),
                    new XElement(Xlsx.ExtendedProperties + "DocSecurity", "0"),
                    new XElement(Xlsx.ExtendedProperties + "ScaleCrop", "false"),
                    new XElement(Xlsx.ExtendedProperties + "Company", string.Empty),
                    new XElement(Xlsx.ExtendedProperties + "AppVersion", "1.0"));
                return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
            }

            /// <summary>
            /// 把单个工作表部件流式写入归档条目：dimension、视图与冻结窗格、格式默认值、cols、
            /// sheetData（含行高与隐藏行）、mergeCells 及页边距/页面设置。
            /// Streams one worksheet part into an archive entry: dimension, views and frozen panes, format
            /// defaults, cols, sheetData (with row heights and hidden rows), mergeCells and the margins/page setup.
            /// </summary>
            /// <param name="archive">目标 zip 归档。 Target zip archive.</param>
            /// <param name="name">部件路径名。 Part path name.</param>
            /// <param name="sheet">要写出的工作表。 Worksheet to write.</param>
            /// <param name="first">是否为活动表（决定 tabSelected 标记）。 Whether this is the active sheet (drives the tabSelected flag).</param>
            /// <param name="sheetIndex">表下标（用于定位绘图部件）。 Sheet index, used to locate the drawing part.</param>
            private void WriteWorksheet(ZipArchive archive, string name, Worksheet sheet, bool first, int sheetIndex)
            {
                ZipArchiveEntry entry = archive.CreateEntry(OfficePackageLimits.ValidateEntryName(name), CompressionLevel.Optimal);
                using (Stream stream = entry.Open())
                {
                    using (XmlWriter writer = Xlsx.CreateStreamingWriter(stream, false))
                    {
                        WriteWorksheetXml(writer, sheet, first, sheetIndex);
                    }
                }
            }

            /// <summary>按 worksheet 部件的元素顺序写出全部 XML。 Writes the whole part following the worksheet element order.</summary>
            /// <param name="writer">目标写出器。 Target writer.</param>
            /// <param name="sheet">工作表。 Worksheet.</param>
            /// <param name="first">是否为活动表。 Whether this is the active sheet.</param>
            /// <param name="sheetIndex">表下标。 Sheet index.</param>
            private void WriteWorksheetXml(XmlWriter writer, Worksheet sheet, bool first, int sheetIndex)
            {
                string main = Xlsx.Main.NamespaceName;
                writer.WriteStartDocument(true);
                writer.WriteStartElement("worksheet", main);
                writer.WriteAttributeString("xmlns", "r", null, Xlsx.R.NamespaceName);

                if (sheet.FitToPage || sheet.TabColor.HasValue)
                {
                    writer.WriteStartElement("sheetPr", main);
                    if (sheet.TabColor.HasValue)
                    {
                        writer.WriteStartElement("tabColor", main);
                        writer.WriteAttributeString("rgb", sheet.TabColor.Value.ToArgbHex());
                        writer.WriteEndElement();
                    }
                    if (sheet.FitToPage)
                    {
                        writer.WriteStartElement("pageSetUpPr", main);
                        writer.WriteAttributeString("fitToPage", "1");
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }

                int lastRow;
                int lastColumn;
                MeasureUsedRange(sheet, out lastRow, out lastColumn);
                writer.WriteStartElement("dimension", main);
                writer.WriteAttributeString("ref", "A1:" + CellAddress.Format(lastRow, lastColumn));
                writer.WriteEndElement();

                writer.WriteStartElement("sheetViews", main);
                writer.WriteStartElement("sheetView", main);
                writer.WriteAttributeString("workbookViewId", "0");
                writer.WriteAttributeString("showGridLines", sheet.ShowGridLines ? "1" : "0");
                if (first) writer.WriteAttributeString("tabSelected", "1");
                if (!sheet.ShowZeros) writer.WriteAttributeString("showZeros", "0");
                if (sheet.FrozenRows > 0 || sheet.FrozenColumns > 0)
                {
                    string pane = PaneName(sheet.FrozenRows, sheet.FrozenColumns);
                    string anchor = CellAddress.Format(sheet.FrozenRows + 1, sheet.FrozenColumns + 1);
                    writer.WriteStartElement("pane", main);
                    writer.WriteAttributeString("state", "frozen");
                    writer.WriteAttributeString("topLeftCell", anchor);
                    writer.WriteAttributeString("activePane", pane);
                    if (sheet.FrozenColumns > 0) writer.WriteAttributeString("xSplit", sheet.FrozenColumns.ToString(CultureInfo.InvariantCulture));
                    if (sheet.FrozenRows > 0) writer.WriteAttributeString("ySplit", sheet.FrozenRows.ToString(CultureInfo.InvariantCulture));
                    writer.WriteEndElement();
                    writer.WriteStartElement("selection", main);
                    writer.WriteAttributeString("pane", pane);
                    writer.WriteAttributeString("activeCell", anchor);
                    writer.WriteAttributeString("sqref", anchor);
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                writer.WriteEndElement();

                writer.WriteStartElement("sheetFormatPr", main);
                writer.WriteAttributeString("defaultRowHeight", Xlsx.ToInvariant(sheet.DefaultRowHeight));
                writer.WriteAttributeString("defaultColWidth", Xlsx.ToInvariant(sheet.DefaultColumnWidth));
                writer.WriteAttributeString("baseColWidth", "8");
                writer.WriteEndElement();

                WriteColumns(writer, sheet);
                WriteSheetData(writer, sheet);
                WriteAutoFilter(writer, sheet);

                if (sheet.MergedRanges.Count > 0)
                {
                    writer.WriteStartElement("mergeCells", main);
                    writer.WriteAttributeString("count", sheet.MergedRanges.Count.ToString(CultureInfo.InvariantCulture));
                    for (int i = 0; i < sheet.MergedRanges.Count; i++)
                    {
                        MergedRange range = sheet.MergedRanges[i];
                        writer.WriteStartElement("mergeCell", main);
                        writer.WriteAttributeString("ref", CellAddress.Format(range.Row, range.Column) + ":"
                            + CellAddress.Format(range.LastRow, range.LastColumn));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }

                List<XElement> formats = _conditionalFormats.Build(sheet, _styles.RegisterDxf);
                for (int i = 0; i < formats.Count; i++) formats[i].WriteTo(writer);
                WriteDataValidations(writer, sheet);
                _hyperlinks.Write(writer, sheetIndex);

                PageSetup setup = sheet.PrintSetup ?? new PageSetup();
                if (sheet.CenterHorizontally)
                {
                    writer.WriteStartElement("printOptions", main);
                    writer.WriteAttributeString("horizontalCentered", "1");
                    writer.WriteEndElement();
                }
                writer.WriteStartElement("pageMargins", main);
                writer.WriteAttributeString("left", Xlsx.ToInvariant(setup.MarginLeft / 72.0));
                writer.WriteAttributeString("right", Xlsx.ToInvariant(setup.MarginRight / 72.0));
                writer.WriteAttributeString("top", Xlsx.ToInvariant(setup.MarginTop / 72.0));
                writer.WriteAttributeString("bottom", Xlsx.ToInvariant(setup.MarginBottom / 72.0));
                writer.WriteAttributeString("header", Xlsx.ToInvariant(setup.HeaderDistance / 72.0));
                writer.WriteAttributeString("footer", Xlsx.ToInvariant(setup.FooterDistance / 72.0));
                writer.WriteEndElement();

                int paperSize = IsLetter(setup) ? 1 : 9;
                writer.WriteStartElement("pageSetup", main);
                writer.WriteAttributeString("paperSize", paperSize.ToString(CultureInfo.InvariantCulture));
                writer.WriteAttributeString("orientation", setup.Width > setup.Height ? "landscape" : "portrait");
                if (sheet.FitToPage)
                {
                    writer.WriteAttributeString("fitToWidth", "1");
                    writer.WriteAttributeString("fitToHeight", "0");
                }
                writer.WriteEndElement();

                if (_decorations != null && _decorations.GetDrawingTarget(sheetIndex) != null)
                {
                    writer.WriteStartElement("drawing", main);
                    writer.WriteAttributeString("r", "id", Xlsx.R.NamespaceName, "rId1");
                    writer.WriteEndElement();
                }

                writer.WriteEndElement();
                writer.WriteEndDocument();
            }

            /// <summary>判断页面尺寸（允许 2 磅误差）是否为 Letter 纸张（612x792 磅）。 Checks whether the page size is Letter (612x792 points) allowing a 2-point tolerance.</summary>
            /// <param name="setup">页面设置。 Page setup.</param>
            /// <returns>是 Letter 返回 true。 True when the paper is Letter.</returns>
            private static bool IsLetter(PageSetup setup)
            {
                double shortSide = Math.Min(setup.Width, setup.Height);
                double longSide = Math.Max(setup.Width, setup.Height);
                return Math.Abs(shortSide - 612) < 2 && Math.Abs(longSide - 792) < 2;
            }

            /// <summary>按冻结行列数返回 OOXML pane 名称（bottomRight/bottomLeft/topRight）。 Returns the OOXML pane name (bottomRight/bottomLeft/topRight) for the frozen row and column counts.</summary>
            /// <param name="rows">冻结行数。 Frozen row count.</param>
            /// <param name="columns">冻结列数。 Frozen column count.</param>
            /// <returns>窗格名称。 The pane name.</returns>
            private static string PaneName(int rows, int columns)
            {
                if (rows > 0 && columns > 0) return "bottomRight";
                if (rows > 0) return "bottomLeft";
                return "topRight";
            }

            /// <summary>写出 cols 元素：将列宽/隐藏列按相邻且属性相同的列合并为区间。 Writes the cols element, merging adjacent columns with equal width/hidden attributes into ranges.</summary>
            /// <param name="writer">目标写出器。 Target writer.</param>
            /// <param name="sheet">工作表。 Worksheet.</param>
            private void WriteColumns(XmlWriter writer, Worksheet sheet)
            {
                string main = Xlsx.Main.NamespaceName;
                SortedDictionary<int, ColumnEntry> entries = new SortedDictionary<int, ColumnEntry>();
                foreach (KeyValuePair<int, double> pair in sheet.ColumnWidths)
                {
                    entries[pair.Key] = new ColumnEntry(pair.Value, sheet.HiddenColumns.Contains(pair.Key));
                }
                foreach (int column in sheet.HiddenColumns)
                {
                    if (!entries.ContainsKey(column)) entries[column] = new ColumnEntry(double.NaN, true);
                }
                writer.WriteStartElement("cols", main);
                List<int> keys = new List<int>(entries.Keys);
                int index = 0;
                while (index < keys.Count)
                {
                    int start = keys[index];
                    ColumnEntry entry = entries[start];
                    int end = start;
                    int next = index + 1;
                    while (next < keys.Count && keys[next] == end + 1 && entries[keys[next]].Matches(entry))
                    {
                        end = keys[next];
                        next++;
                    }
                    writer.WriteStartElement("col", main);
                    writer.WriteAttributeString("min", start.ToString(CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("max", end.ToString(CultureInfo.InvariantCulture));
                    if (!double.IsNaN(entry.Width))
                    {
                        writer.WriteAttributeString("width", Xlsx.ToInvariant(entry.Width));
                        writer.WriteAttributeString("customWidth", "1");
                    }
                    if (entry.Hidden) writer.WriteAttributeString("hidden", "1");
                    writer.WriteEndElement();
                    index = next;
                }
                writer.WriteEndElement();
            }

            /// <summary>列写出用的中间值：宽度（NaN 表示未显式设置）与隐藏标记。 Intermediate value for column output: width (NaN means not explicitly set) and the hidden flag.</summary>
            private sealed class ColumnEntry
            {
                /// <summary>创建列条目。 Creates a column entry.</summary>
                /// <param name="width">列宽；NaN 表示无自定义宽度。 Column width; NaN means no custom width.</param>
                /// <param name="hidden">是否隐藏。 Whether the column is hidden.</param>
                public ColumnEntry(double width, bool hidden)
                {
                    Width = width;
                    Hidden = hidden;
                }

                /// <summary>列宽（NaN 表示未显式设置）。 Column width (NaN means it was not set explicitly).</summary>
                public double Width { get; private set; }
                /// <summary>是否隐藏列。 Whether the column is hidden.</summary>
                public bool Hidden { get; private set; }

                /// <summary>判断另一列条目可否与本条目合并为同一 col 区间（隐藏状态相同且宽度近似相等）。 Checks whether another column entry can be merged into the same col range: identical hidden state and approximately equal widths.</summary>
                /// <param name="other">相邻列条目。 Adjacent column entry.</param>
                /// <returns>可合并时返回 true。 True when the entries can be merged.</returns>
                public bool Matches(ColumnEntry other)
                {
                    if (Hidden != other.Hidden) return false;
                    if (double.IsNaN(Width) || double.IsNaN(other.Width)) return double.IsNaN(Width) && double.IsNaN(other.Width);
                    return Math.Abs(Width - other.Width) < 0.0001;
                }
            }

            /// <summary>写出 autoFilter 元素：区域引用与各列允许的取值。 Writes the autoFilter element: the range reference and the allowed values of each column.</summary>
            /// <param name="writer">目标写出器。 Target writer.</param>
            /// <param name="sheet">工作表。 Worksheet.</param>
            private static void WriteAutoFilter(XmlWriter writer, Worksheet sheet)
            {
                ExcelAutoFilter filter = sheet.AutoFilter;
                if (filter == null || string.IsNullOrEmpty(filter.Range)) return;
                string reference;
                int width;
                if (!TryNormalizeRange(filter.Range, out reference, out width)) return;
                string main = Xlsx.Main.NamespaceName;
                writer.WriteStartElement("autoFilter", main);
                writer.WriteAttributeString("ref", reference);
                for (int i = 0; i < filter.Columns.Count; i++)
                {
                    ExcelFilterColumn column = filter.Columns[i];
                    if (column == null || column.ColumnId < 0 || column.ColumnId >= width) continue;
                    if (column.Values.Count == 0 && !column.IncludeBlanks) continue;
                    writer.WriteStartElement("filterColumn", main);
                    writer.WriteAttributeString("colId", column.ColumnId.ToString(CultureInfo.InvariantCulture));
                    writer.WriteStartElement("filters", main);
                    if (column.IncludeBlanks) writer.WriteAttributeString("blank", "1");
                    for (int v = 0; v < column.Values.Count; v++)
                    {
                        writer.WriteStartElement("filter", main);
                        writer.WriteAttributeString("val", column.Values[v] ?? string.Empty);
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }

            /// <summary>写出 dataValidations 元素：类型、运算符、公式、提示与错误文本。 Writes the dataValidations element: types, operators, formulas, prompts and error texts.</summary>
            /// <param name="writer">目标写出器。 Target writer.</param>
            /// <param name="sheet">工作表。 Worksheet.</param>
            private static void WriteDataValidations(XmlWriter writer, Worksheet sheet)
            {
                List<ExcelDataValidation> items = new List<ExcelDataValidation>();
                for (int i = 0; i < sheet.DataValidations.Count; i++)
                {
                    ExcelDataValidation validation = sheet.DataValidations[i];
                    if (validation == null || validation.Type == DataValidationType.None) continue;
                    string reference;
                    int width;
                    if (!TryNormalizeRange(validation.Range, out reference, out width)) continue;
                    items.Add(validation);
                }
                if (items.Count == 0) return;
                string main = Xlsx.Main.NamespaceName;
                writer.WriteStartElement("dataValidations", main);
                writer.WriteAttributeString("count", items.Count.ToString(CultureInfo.InvariantCulture));
                for (int i = 0; i < items.Count; i++)
                {
                    ExcelDataValidation validation = items[i];
                    string reference;
                    int width;
                    TryNormalizeRange(validation.Range, out reference, out width);
                    writer.WriteStartElement("dataValidation", main);
                    writer.WriteAttributeString("type", ValidationTypeName(validation.Type));
                    if (validation.Type != DataValidationType.List && validation.Type != DataValidationType.Custom)
                    {
                        writer.WriteAttributeString("operator", ValidationOperatorName(validation.Operator));
                    }
                    if (validation.AllowBlank) writer.WriteAttributeString("allowBlank", "1");
                    if (!validation.ShowDropDown) writer.WriteAttributeString("showDropDown", "1");
                    if (validation.ShowInputMessage) writer.WriteAttributeString("showInputMessage", "1");
                    if (validation.ShowErrorMessage) writer.WriteAttributeString("showErrorMessage", "1");
                    if (validation.ErrorStyle != DataValidationErrorStyle.Stop)
                    {
                        writer.WriteAttributeString("errorStyle", validation.ErrorStyle == DataValidationErrorStyle.Warning ? "warning" : "information");
                    }
                    if (!string.IsNullOrEmpty(validation.PromptTitle)) writer.WriteAttributeString("promptTitle", OfficeXmlText.Strip(validation.PromptTitle));
                    if (!string.IsNullOrEmpty(validation.Prompt)) writer.WriteAttributeString("prompt", OfficeXmlText.Strip(validation.Prompt));
                    if (!string.IsNullOrEmpty(validation.ErrorTitle)) writer.WriteAttributeString("errorTitle", OfficeXmlText.Strip(validation.ErrorTitle));
                    if (!string.IsNullOrEmpty(validation.ErrorMessage)) writer.WriteAttributeString("error", OfficeXmlText.Strip(validation.ErrorMessage));
                    writer.WriteAttributeString("sqref", reference);
                    string formula1 = NormalizeListFormula(validation);
                    if (!string.IsNullOrEmpty(formula1))
                    {
                        writer.WriteStartElement("formula1", main);
                        writer.WriteString(OfficeXmlText.Strip(formula1));
                        writer.WriteEndElement();
                    }
                    if (!string.IsNullOrEmpty(validation.Formula2)
                        && (validation.Operator == DataValidationOperator.Between || validation.Operator == DataValidationOperator.NotBetween))
                    {
                        writer.WriteStartElement("formula2", main);
                        writer.WriteString(OfficeXmlText.Strip(validation.Formula2));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }

            /// <summary>列表验证的公式规范化：未加引号且含逗号的字面量列表补上外围引号，区域引用保持原样。 Normalizes a list formula: an unquoted comma-separated literal list gets surrounding quotes while range references are kept as-is.</summary>
            /// <param name="validation">数据验证规则。 The validation rule.</param>
            /// <returns>写入 formula1 的文本。 The text written into formula1.</returns>
            private static string NormalizeListFormula(ExcelDataValidation validation)
            {
                string formula = validation.Formula1;
                if (string.IsNullOrEmpty(formula) || validation.Type != DataValidationType.List) return formula;
                string trimmed = formula.Trim();
                if (trimmed.Length == 0) return formula;
                if (trimmed[0] == '"' || trimmed[0] == '=') return trimmed;
                if (trimmed.IndexOf('!') >= 0 || trimmed.IndexOf(':') >= 0 || trimmed.IndexOf('$') >= 0) return trimmed;
                if (trimmed.IndexOf(',') < 0 && trimmed.IndexOf(';') < 0) return trimmed;
                return "\"" + trimmed + "\"";
            }

            /// <summary>数据验证类型对应的 OOXML 名称。 The OOXML name of a validation type.</summary>
            /// <param name="type">验证类型。 Validation type.</param>
            /// <returns>OOXML 类型名。 The OOXML type name.</returns>
            private static string ValidationTypeName(DataValidationType type)
            {
                switch (type)
                {
                    case DataValidationType.Whole: return "whole";
                    case DataValidationType.Decimal: return "decimal";
                    case DataValidationType.List: return "list";
                    case DataValidationType.Date: return "date";
                    case DataValidationType.Time: return "time";
                    case DataValidationType.TextLength: return "textLength";
                    default: return "custom";
                }
            }

            /// <summary>数据验证运算符对应的 OOXML 名称。 The OOXML name of a validation operator.</summary>
            /// <param name="value">运算符。 The operator.</param>
            /// <returns>OOXML 运算符名。 The OOXML operator name.</returns>
            private static string ValidationOperatorName(DataValidationOperator value)
            {
                switch (value)
                {
                    case DataValidationOperator.NotBetween: return "notBetween";
                    case DataValidationOperator.Equal: return "equal";
                    case DataValidationOperator.NotEqual: return "notEqual";
                    case DataValidationOperator.LessThan: return "lessThan";
                    case DataValidationOperator.LessThanOrEqual: return "lessThanOrEqual";
                    case DataValidationOperator.GreaterThan: return "greaterThan";
                    case DataValidationOperator.GreaterThanOrEqual: return "greaterThanOrEqual";
                    default: return "between";
                }
            }

            /// <summary>把区域引用规范化为 "A1:B2" 形式并返回列数；非法引用返回 false。 Normalizes a range reference to "A1:B2" form and returns the column count; invalid references return false.</summary>
            /// <param name="range">原始引用。 The raw reference.</param>
            /// <param name="reference">规范化后的引用。 The normalized reference.</param>
            /// <param name="width">区域的列数。 The number of columns in the range.</param>
            /// <returns>解析成功返回 true。 True when parsing succeeded.</returns>
            private static bool TryNormalizeRange(string range, out string reference, out int width)
            {
                reference = null;
                width = 0;
                if (string.IsNullOrEmpty(range)) return false;
                int row1, column1, row2, column2;
                try
                {
                    CellAddress.ParseRange(range, out row1, out column1, out row2, out column2);
                }
                catch (FormatException)
                {
                    return false;
                }
                if (row1 < 1 || column1 < 1 || row2 < 1 || column2 < 1) return false;
                if (row1 > row2) { int swap = row1; row1 = row2; row2 = swap; }
                if (column1 > column2) { int swap = column1; column1 = column2; column2 = swap; }
                reference = CellAddress.Format(row1, column1) + ":" + CellAddress.Format(row2, column2);
                width = column2 - column1 + 1;
                return true;
            }

            /// <summary>判断单元格是否需要写出（有值、有公式、带样式或带超链接）。 Decides whether a cell must be written (it has a value, a formula, a style or a hyperlink).</summary>
            /// <param name="cell">目标单元格。 Target cell.</param>
            /// <returns>需要写出返回 true。 True when the cell is written.</returns>
            private static bool IsWritableCell(Cell cell)
            {
                return cell.Value != null || !string.IsNullOrEmpty(cell.Formula) || cell.Style != null
                    || !string.IsNullOrEmpty(cell.Hyperlink);
            }

            /// <summary>
            /// 统计已用区域右下角：取需写出单元格、合并区域、自定义行高与隐藏行/列的最大行列，至少 A1。
            /// Measures the used-range bottom-right corner from the writable cells, merged ranges, custom row
            /// heights and hidden rows/columns, never smaller than A1.
            /// </summary>
            /// <param name="sheet">工作表。 Worksheet.</param>
            /// <param name="lastRow">输出末行。 Outputs the last row.</param>
            /// <param name="lastColumn">输出末列。 Outputs the last column.</param>
            private static void MeasureUsedRange(Worksheet sheet, out int lastRow, out int lastColumn)
            {
                lastRow = 0;
                lastColumn = 0;
                foreach (Cell cell in sheet.Cells)
                {
                    if (!IsWritableCell(cell)) continue;
                    if (cell.Row > lastRow) lastRow = cell.Row;
                    if (cell.Column > lastColumn) lastColumn = cell.Column;
                }
                for (int i = 0; i < sheet.MergedRanges.Count; i++)
                {
                    if (sheet.MergedRanges[i].LastRow > lastRow) lastRow = sheet.MergedRanges[i].LastRow;
                    if (sheet.MergedRanges[i].LastColumn > lastColumn) lastColumn = sheet.MergedRanges[i].LastColumn;
                }
                foreach (int row in sheet.RowHeights.Keys)
                {
                    if (row > lastRow) lastRow = row;
                }
                foreach (int row in sheet.HiddenRows)
                {
                    if (row > lastRow) lastRow = row;
                }
                foreach (int column in sheet.HiddenColumns)
                {
                    if (column > lastColumn) lastColumn = column;
                }
                if (lastRow < 1) lastRow = 1;
                if (lastColumn < 1) lastColumn = 1;
            }

            /// <summary>
            /// 写出 sheetData：单元格按 Worksheet.Cells 的行优先顺序单遍流式输出，
            /// 仅有行高或隐藏标记而无单元格的行按行号归并补写为空行元素。
            /// Writes sheetData in a single streaming pass over Worksheet.Cells, which are already row-major,
            /// merging in empty row elements for rows that only carry a custom height or the hidden flag.
            /// </summary>
            /// <param name="writer">目标写出器。 Target writer.</param>
            /// <param name="sheet">工作表。 Worksheet.</param>
            private void WriteSheetData(XmlWriter writer, Worksheet sheet)
            {
                string main = Xlsx.Main.NamespaceName;
                int[] extraRows = CollectExtraRows(sheet);
                writer.WriteStartElement("sheetData", main);
                int extra = 0;
                int openRow = -1;
                foreach (Cell cell in sheet.Cells)
                {
                    if (!IsWritableCell(cell)) continue;
                    int row = cell.Row;
                    while (extra < extraRows.Length && extraRows[extra] < row)
                    {
                        if (openRow >= 0)
                        {
                            writer.WriteEndElement();
                            openRow = -1;
                        }
                        WriteRowStart(writer, main, sheet, extraRows[extra]);
                        writer.WriteEndElement();
                        extra++;
                    }
                    if (extra < extraRows.Length && extraRows[extra] == row) extra++;
                    if (openRow != row)
                    {
                        if (openRow >= 0) writer.WriteEndElement();
                        openRow = row;
                        WriteRowStart(writer, main, sheet, row);
                    }
                    WriteCell(writer, main, cell);
                }
                if (openRow >= 0) writer.WriteEndElement();
                while (extra < extraRows.Length)
                {
                    WriteRowStart(writer, main, sheet, extraRows[extra]);
                    writer.WriteEndElement();
                    extra++;
                }
                writer.WriteEndElement();
            }

            /// <summary>收集没有单元格但仍需生成 row 元素的行号（自定义行高或隐藏行），升序去重。 Collects the row numbers that have no cells but still need a row element (custom height or hidden), sorted ascending and de-duplicated.</summary>
            /// <param name="sheet">工作表。 Worksheet.</param>
            /// <returns>行号数组。 The array of row numbers.</returns>
            private static int[] CollectExtraRows(Worksheet sheet)
            {
                List<int> values = new List<int>(sheet.RowHeights.Count + sheet.HiddenRows.Count);
                foreach (int row in sheet.RowHeights.Keys) values.Add(row);
                foreach (int row in sheet.HiddenRows) values.Add(row);
                if (values.Count == 0) return new int[0];
                values.Sort();
                List<int> unique = new List<int>(values.Count);
                for (int i = 0; i < values.Count; i++)
                {
                    if (i == 0 || values[i] != values[i - 1]) unique.Add(values[i]);
                }
                return unique.ToArray();
            }

            /// <summary>写出 row 开始标签及其行高/隐藏属性（不闭合）。 Writes the row start tag with its height and hidden attributes, left open.</summary>
            /// <param name="writer">目标写出器。 Target writer.</param>
            /// <param name="main">主命名空间。 Main namespace.</param>
            /// <param name="sheet">工作表。 Worksheet.</param>
            /// <param name="row">行号。 Row number.</param>
            private static void WriteRowStart(XmlWriter writer, string main, Worksheet sheet, int row)
            {
                writer.WriteStartElement("row", main);
                writer.WriteAttributeString("r", row.ToString(CultureInfo.InvariantCulture));
                double height;
                if (sheet.RowHeights.TryGetValue(row, out height) && height > 0)
                {
                    writer.WriteAttributeString("ht", Xlsx.ToInvariant(height));
                    writer.WriteAttributeString("customHeight", "1");
                }
                if (sheet.HiddenRows.Contains(row)) writer.WriteAttributeString("hidden", "1");
            }

            /// <summary>
            /// 单元格数值文本：整值走 long.ToString，比 G17 快约一个数量级且分配更少；
            /// 非整值仍用 G17，保证读回时精度可往返。
            /// Cell number text: integral values go through long.ToString, roughly an order of magnitude faster
            /// than G17 with fewer allocations; fractional values keep G17 so the round trip stays exact.
            /// </summary>
            /// <param name="number">数值。 The number.</param>
            /// <returns>不变体格式文本。 The invariant-format text.</returns>
            private static string NumberText(double number)
            {
                if (number >= -9007199254740992.0 && number <= 9007199254740992.0
                    && number == Math.Floor(number))
                {
                    return ((long)number).ToString(CultureInfo.InvariantCulture);
                }
                return Xlsx.ToInvariant(number);
            }

            /// <summary>
            /// 直接在可复用字符缓冲区上拼出"行字母+行数字"形式的单元格地址，每个地址只分配一个字符串；
            /// 列号非法时回退到 CellAddress.Format 以抛出一致的异常。
            /// Builds the "column letters + row digits" cell address in a reusable character buffer so each address
            /// allocates a single string; invalid column numbers fall back to CellAddress.Format to keep the same
            /// exception behavior.
            /// </summary>
            /// <param name="row">行号（从 1 开始）。 Row number (1-based).</param>
            /// <param name="column">列号（从 1 开始）。 Column number (1-based).</param>
            /// <returns>单元格地址文本。 The cell address text.</returns>
            private string CellAddressText(int row, int column)
            {
                if (column < 1 || row < 0) return CellAddress.Format(row, column);
                char[] buffer = _addressBuffer;
                int length = 0;
                int current = column;
                while (current > 0)
                {
                    current--;
                    buffer[length++] = (char)('A' + (current % 26));
                    current /= 26;
                }
                Reverse(buffer, 0, length);
                int digitsStart = length;
                int rowNumber = row;
                if (rowNumber == 0) buffer[length++] = '0';
                else
                {
                    while (rowNumber > 0)
                    {
                        buffer[length++] = (char)('0' + (rowNumber % 10));
                        rowNumber /= 10;
                    }
                }
                Reverse(buffer, digitsStart, length);
                return new string(buffer, 0, length);
            }

            /// <summary>就地反转缓冲区中 [start, end) 区间。 Reverses buffer[start, end) in place.</summary>
            /// <param name="buffer">字符缓冲区。 Character buffer.</param>
            /// <param name="start">区间起点。 Start index.</param>
            /// <param name="end">区间终点（不含）。 End index, exclusive.</param>
            private static void Reverse(char[] buffer, int start, int end)
            {
                for (int left = start, right = end - 1; left < right; left++, right--)
                {
                    char swap = buffer[left];
                    buffer[left] = buffer[right];
                    buffer[right] = swap;
                }
            }

            /// <summary>取小整数的不变文化字符串形式，按值缓存复用，避免同一索引反复分配。 Returns the invariant text of a small index, cached per value so the same index is not allocated repeatedly.</summary>
            /// <param name="value">非负索引。 Non-negative index.</param>
            /// <returns>缓存的字符串文本。 The cached string text.</returns>
            private string IndexText(int value)
            {
                if (value < 0) return value.ToString(CultureInfo.InvariantCulture);
                string[] cache = _indexTexts;
                if (value >= cache.Length)
                {
                    int length = cache.Length;
                    while (length <= value) length *= 2;
                    Array.Resize(ref _indexTexts, length);
                    cache = _indexTexts;
                }
                string text = cache[value];
                if (text == null)
                {
                    text = value.ToString(CultureInfo.InvariantCulture);
                    cache[value] = text;
                }
                return text;
            }

            /// <summary>注册单元格样式并返回索引；DateTime 值若样式非日期格式则补默认日期格式。 Registers the cell style and returns its index; DateTime values whose style is not a date format get the default date format.</summary>
            /// <param name="cell">目标单元格。 Target cell.</param>
            /// <returns>cellXfs 中的样式索引。 The style index inside cellXfs.</returns>
            private int GetStyleIndex(Cell cell)
            {
                CellStyle style = cell.Style;
                bool isDate = cell.Value is DateTime;
                if (style == null)
                {
                    // 空样式不等于「这一格不需要日期格式」：不补 numFmt，日期就以 General 写成 45064 这样的
                    // 序列号，收件人看到的是一串数字而不是日期。补的格式必须从 xf 0 的默认样式派生，
                    // 直接 new CellStyle() 会顺带把 Calibri 预设写进包里。
                    // A null style does not mean "this cell needs no date format": without a numFmt the date goes
                    // out as a General-format serial such as 45064 and the recipient sees digits. The format has to
                    // be derived from the default style behind xf 0, because a plain new CellStyle() would carry the
                    // preset Calibri font into the package.
                    if (!isDate) return 0;
                    if (_plainDateStyleIndex < 0)
                    {
                        CellStyle adjusted = _styles.DefaultStyle.Clone();
                        adjusted.NumberFormat = "yyyy-mm-dd";
                        _plainDateStyleIndex = _styles.Register(adjusted);
                    }
                    return _plainDateStyleIndex;
                }
                bool dateFallback = isDate && !style.IsDateFormat();
                Dictionary<CellStyle, int> cache = dateFallback ? _dateStyleIndexes : _styleIndexes;
                int cached;
                if (cache.TryGetValue(style, out cached)) return cached;
                int index;
                if (dateFallback)
                {
                    CellStyle adjusted = style.Clone();
                    adjusted.NumberFormat = "yyyy-mm-dd";
                    index = _styles.Register(adjusted);
                }
                else
                {
                    index = _styles.Register(style);
                }
                cache[style] = index;
                return index;
            }

            /// <summary>登记共享字符串：引用计数加一，首次出现的字符串入表并返回其索引。 Registers a shared string: increments the reference count and appends first-time values, returning their index.</summary>
            /// <param name="value">字符串值。 String value.</param>
            /// <returns>共享字符串索引。 The shared-string index.</returns>
            private int GetSharedString(string value)
            {
                _sharedStringReferences++;
                int index;
                if (_sharedStringIds.TryGetValue(value, out index)) return index;
                index = _sharedStrings.Count;
                _sharedStrings.Add(value);
                _sharedStringIds[value] = index;
                return index;
            }

            /// <summary>
            /// 写出 c 元素：先算出样式索引与值文本，再按 r/s/t 顺序写属性，最后写可选 f 公式
            /// （去前导 =）与 v 值（字符串走共享表或 str 缓存值，布尔写 0/1，日期写序列号，其余数值直接写数字）。
            /// Writes the c element: the style index and the value text are resolved first so the r/s/t attributes can
            /// follow in that order, then the optional f formula (leading = removed) and the v value are written
            /// (strings use the shared table or a str cached value, booleans write 0/1, dates write the serial
            /// number, other numbers are written directly).
            /// </summary>
            /// <param name="writer">目标写出器。 Target writer.</param>
            /// <param name="main">主命名空间。 Main namespace.</param>
            /// <param name="cell">单元格模型。 Cell model.</param>
            private void WriteCell(XmlWriter writer, string main, Cell cell)
            {
                string formula = cell.Formula;
                bool hasFormula = !string.IsNullOrEmpty(formula);
                if (hasFormula && formula.StartsWith("=", StringComparison.Ordinal)) formula = formula.Substring(1);

                int styleIndex = GetStyleIndex(cell);
                object value = cell.Value;
                string type = null;
                string text = null;
                if (value != null)
                {
                    if (value is string)
                    {
                        string content = (string)value;
                        if (hasFormula)
                        {
                            type = "str";
                            text = content;
                        }
                        else
                        {
                            type = "s";
                            text = IndexText(GetSharedString(content));
                        }
                    }
                    else if (value is bool)
                    {
                        type = "b";
                        text = (bool)value ? "1" : "0";
                    }
                    else if (value is DateTime)
                    {
                        text = NumberText(ToSerial((DateTime)value));
                    }
                    else
                    {
                        double number;
                        if (TryGetNumber(value, out number))
                        {
                            if (double.IsNaN(number) || double.IsInfinity(number))
                            {
                                // 数字单元存不下 NaN 与无穷：写出 <v>NaN</v> 是非法数值，Excel 会把整包判为
                                // 需要修复；按 Excel 自己的表示法落成 #NUM! 错误单元。
                                // A numeric cell cannot hold NaN or infinity: <v>NaN</v> is not a legal number and
                                // Excel declares the whole package in need of repair, so the value lands on
                                // Excel's own representation of an unrepresentable result, the #NUM! error.
                                type = "e";
                                text = "#NUM!";
                            }
                            else
                            {
                                text = NumberText(number);
                            }
                        }
                        else
                        {
                            string fallback = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                            if (hasFormula)
                            {
                                type = "str";
                                text = fallback;
                            }
                            else
                            {
                                type = "s";
                                text = IndexText(GetSharedString(fallback));
                            }
                        }
                    }
                }

                writer.WriteStartElement("c", main);
                writer.WriteAttributeString("r", CellAddressText(cell.Row, cell.Column));
                if (styleIndex > 0) writer.WriteAttributeString("s", IndexText(styleIndex));
                if (type != null) writer.WriteAttributeString("t", type);
                if (hasFormula)
                {
                    writer.WriteStartElement("f", main);
                    // 流式写法绕开了 SerializeXml 的统一清洗，公式与缓存值自己把非法字符去掉，
                    // 否则一格脏文本就让 XmlWriter 抛异常、整份工作簿写坏。
                    // This streaming path bypasses the clean-up inside SerializeXml, so the formula and the cached
                    // value drop the illegal characters themselves — otherwise one dirty cell throws in XmlWriter
                    // and the whole workbook fails to be written.
                    writer.WriteString(OfficeXmlText.Strip(formula));
                    writer.WriteEndElement();
                }
                if (value != null)
                {
                    writer.WriteStartElement("v", main);
                    writer.WriteString(OfficeXmlText.Strip(text));
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
            }

            /// <summary>
            /// 尝试将整数、浮点或 Decimal 等数值类型统一转为 double；
            /// 常见类型先用类型测试短路，其余数值类型再走 TypeCode 分支。
            /// Tries to convert numeric types (integers, floating point or Decimal) into a double; the common
            /// types are short-circuited with type tests and the remaining numeric types fall back to TypeCode.
            /// </summary>
            /// <param name="value">待转换的值。 Value to convert.</param>
            /// <param name="number">转换结果；失败时为 0。 Conversion result; 0 when it fails.</param>
            /// <returns>转换成功返回 true。 True when the conversion succeeded.</returns>
            private static bool TryGetNumber(object value, out double number)
            {
                if (value is int)
                {
                    number = (int)value;
                    return true;
                }
                if (value is double)
                {
                    number = (double)value;
                    return true;
                }
                if (value is long)
                {
                    number = (long)value;
                    return true;
                }
                if (value is float)
                {
                    number = (float)value;
                    return true;
                }
                if (value is decimal)
                {
                    number = (double)(decimal)value;
                    return true;
                }
                switch (Type.GetTypeCode(value.GetType()))
                {
                    case TypeCode.Byte:
                    case TypeCode.SByte:
                    case TypeCode.Int16:
                    case TypeCode.UInt16:
                    case TypeCode.Int32:
                    case TypeCode.UInt32:
                    case TypeCode.Int64:
                    case TypeCode.UInt64:
                    case TypeCode.Single:
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                        number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                        return true;
                    default:
                        number = 0;
                        return false;
                }
            }

            /// <summary>将 DateTime 转为 Excel 日期序列号（1900 日期系统，基准 1899-12-30）。 Converts a DateTime into an Excel date serial number (1900 date system, epoch 1899-12-30).</summary>
            /// <param name="value">日期时间值。 Date-time value.</param>
            /// <returns>序列号（可含小数部分表示时间）。 The serial number; the fractional part represents the time.</returns>
            private static double ToSerial(DateTime value)
            {
                DateTime epoch = new DateTime(1899, 12, 30);
                return (value - epoch).TotalDays;
            }


            /// <summary>
            /// styles.xml 生成器：对字体、填充、边框、数字格式和 xf 组合分别按内容键去重登记，
            /// 最终序列化为完整的样式表部件字节。
            /// styles.xml generator: registers fonts, fills, borders, number formats and xf combinations
            /// de-duplicated by content key, and finally serializes the complete style part bytes.
            /// </summary>
            private sealed class StyleRegistry
            {
                /// <summary>已登记的字体条目列表。 Registered font entries.</summary>
                private readonly List<FontEntry> _fonts = new List<FontEntry>();
                /// <summary>字体内容键到 fonts 索引的映射。 Font content key to fonts index map.</summary>
                private readonly Dictionary<string, int> _fontIds = new Dictionary<string, int>(StringComparer.Ordinal);
                /// <summary>已登记的填充条目列表。 Registered fill entries.</summary>
                private readonly List<FillEntry> _fills = new List<FillEntry>();
                /// <summary>填充键（颜色十六进制或 none/gray125）到 fills 索引的映射。 Fill key (color hex or none/gray125) to fills index map.</summary>
                private readonly Dictionary<string, int> _fillIds = new Dictionary<string, int>(StringComparer.Ordinal);
                /// <summary>已登记的边框条目列表。 Registered border entries.</summary>
                private readonly List<BorderEntry> _borders = new List<BorderEntry>();
                /// <summary>四边线组合键到 borders 索引的映射。 Four-edge combination key to borders index map.</summary>
                private readonly Dictionary<string, int> _borderIds = new Dictionary<string, int>(StringComparer.Ordinal);
                /// <summary>格式代码到 numFmtId 的映射。 Format code to numFmtId map.</summary>
                private readonly Dictionary<string, int> _formatIds = new Dictionary<string, int>(StringComparer.Ordinal);
                /// <summary>待输出的自定义格式列表（编号与代码）。 Custom formats to output, as id and code pairs.</summary>
                private readonly List<KeyValuePair<int, string>> _customFormats = new List<KeyValuePair<int, string>>();
                /// <summary>已登记的 cellXfs 条目列表。 Registered cellXfs entries.</summary>
                private readonly List<XfEntry> _xfs = new List<XfEntry>();
                /// <summary>样式组合键到 cellXfs 索引的映射。 Style combination key to cellXfs index map.</summary>
                private readonly Dictionary<string, int> _xfIds = new Dictionary<string, int>(StringComparer.Ordinal);
                /// <summary>已登记的差量样式（dxfs）列表。 Registered differential styles (dxfs).</summary>
                private readonly List<CellStyle> _dxfs = new List<CellStyle>();
                /// <summary>差量样式键到 dxfs 索引的映射。 Differential-style key to dxfs index map.</summary>
                private readonly Dictionary<string, int> _dxfIds = new Dictionary<string, int>(StringComparer.Ordinal);
                /// <summary>下一个可用的自定义格式编号（从 164 起）。 Next available custom format id (starting at 164).</summary>
                private int _nextFormatId = 164;

                /// <summary>登记为 xf 0 的工作簿默认样式，供空样式的单元格派生使用。The workbook default style registered as xf 0, the base for cells that carry no style.</summary>
                public readonly CellStyle DefaultStyle;

                /// <summary>初始化必需占位项：fills 前两位（none、gray125）、空边框，并登记默认样式作为 xf 0。 Seeds the mandatory placeholders: the first two fills (none, gray125), an empty border, and registers the default style as xf 0.</summary>
                /// <param name="defaultStyle">工作簿默认样式。 Workbook default style.</param>
                public StyleRegistry(CellStyle defaultStyle)
                {
                    DefaultStyle = defaultStyle;
                    _fills.Add(new FillEntry(null));
                    _fills.Add(new FillEntry(null));
                    _fillIds["none"] = 0;
                    _fillIds["gray125"] = 1;
                    _borders.Add(new BorderEntry(null, null, null, null));
                    _borderIds[""] = 0;
                    Register(defaultStyle);
                }

                /// <summary>登记条件格式用到的差量样式（dxfs），按键去重并返回 dxfId。 Registers a differential style used by conditional formatting (dxfs), de-duplicated by key, and returns its dxfId.</summary>
                /// <param name="style">差量样式。 The differential style.</param>
                /// <returns>dxfs 集合索引。 Index into the dxfs collection.</returns>
                public int RegisterDxf(CellStyle style)
                {
                    if (style == null) style = new CellStyle();
                    string key = (style.Bold ? "b" : "-") + (style.Italic ? "i" : "-") + "|"
                        + style.FontColor.ToArgbHex() + "|"
                        + (style.FillColor.HasValue ? style.FillColor.Value.ToArgbHex() : "none") + "|"
                        + (style.NumberFormat ?? string.Empty);
                    int index;
                    if (_dxfIds.TryGetValue(key, out index)) return index;
                    index = _dxfs.Count;
                    _dxfs.Add(style);
                    _dxfIds[key] = index;
                    return index;
                }

                /// <summary>生成差量样式 dxf 元素：仅写出字体与填充两段。 Builds a dxf element, emitting only the font and fill sections.</summary>
                /// <param name="style">差量样式。 The differential style.</param>
                /// <returns>dxf 元素。 The dxf element.</returns>
                public static XElement CreateDxf(CellStyle style)
                {
                    XElement dxf = new XElement(Xlsx.Main + "dxf");
                    XElement font = new XElement(Xlsx.Main + "font");
                    if (style.Bold) font.Add(new XElement(Xlsx.Main + "b", new XAttribute("val", "1")));
                    if (style.Italic) font.Add(new XElement(Xlsx.Main + "i", new XAttribute("val", "1")));
                    if (style.FontColor != OfficeColor.Black)
                    {
                        font.Add(new XElement(Xlsx.Main + "color", new XAttribute("rgb", style.FontColor.ToArgbHex())));
                    }
                    if (font.HasElements) dxf.Add(font);
                    if (style.FillColor.HasValue)
                    {
                        dxf.Add(new XElement(Xlsx.Main + "fill",
                            new XElement(Xlsx.Main + "patternFill",
                                new XElement(Xlsx.Main + "bgColor", new XAttribute("rgb", style.FillColor.Value.ToArgbHex())))));
                    }
                    return dxf;
                }

                /// <summary>按字体/填充/边框/格式/对齐组合键登记样式并返回 cellXfs 索引，重复样式复用同一索引。 Registers a style by its font/fill/border/format/alignment combination key and returns the cellXfs index; duplicate styles reuse the same index.</summary>
                /// <param name="style">要登记的单元格样式。 Cell style to register.</param>
                /// <returns>样式索引。 The style index.</returns>
                public int Register(CellStyle style)
                {
                    int fontId = GetFontId(style);
                    int fillId = GetFillId(style);
                    int borderId = GetBorderId(style);
                    int formatId = GetFormatId(style.NumberFormat);
                    string alignment = AlignmentKey(style);
                    string key = fontId.ToString(CultureInfo.InvariantCulture) + "|"
                        + fillId.ToString(CultureInfo.InvariantCulture) + "|"
                        + borderId.ToString(CultureInfo.InvariantCulture) + "|"
                        + formatId.ToString(CultureInfo.InvariantCulture) + "|"
                        + alignment;
                    int index;
                    if (_xfIds.TryGetValue(key, out index)) return index;
                    index = _xfs.Count;
                    XfEntry entry = new XfEntry();
                    entry.FontId = fontId;
                    entry.FillId = fillId;
                    entry.BorderId = borderId;
                    entry.FormatId = formatId;
                    entry.Horizontal = style.HorizontalAlignment;
                    entry.Vertical = style.VerticalAlignment;
                    entry.WrapText = style.WrapText;
                    _xfs.Add(entry);
                    _xfIds[key] = index;
                    return index;
                }

                /// <summary>按字体内容键（族名/字号/粗斜下划线/颜色）去重登记字体并返回索引。 Registers a font de-duplicated by its content key (family, size, bold/italic/underline, color) and returns the index.</summary>
                /// <param name="style">单元格样式。 Cell style.</param>
                /// <returns>fonts 集合索引。 Index into the fonts collection.</returns>
                private int GetFontId(CellStyle style)
                {
                    string key = (style.FontFamily ?? "Calibri") + "|"
                        + Xlsx.ToInvariant(style.FontSize <= 0 ? 11 : style.FontSize) + "|"
                        + (style.Bold ? "1" : "0") + (style.Italic ? "1" : "0") + (style.Underline ? "1" : "0") + "|"
                        + style.FontColor.ToArgbHex();
                    int id;
                    if (_fontIds.TryGetValue(key, out id)) return id;
                    id = _fonts.Count;
                    _fonts.Add(new FontEntry(style.FontFamily, style.FontSize <= 0 ? 11 : style.FontSize,
                        style.Bold, style.Italic, style.Underline, style.FontColor));
                    _fontIds[key] = id;
                    return id;
                }

                /// <summary>按填充色十六进制键去重登记填充并返回索引；无填充色返回 0。 Registers a fill de-duplicated by the hexadecimal color key and returns the index; returns 0 when there is no fill color.</summary>
                /// <param name="style">单元格样式。 Cell style.</param>
                /// <returns>fills 集合索引。 Index into the fills collection.</returns>
                private int GetFillId(CellStyle style)
                {
                    if (!style.FillColor.HasValue) return 0;
                    string key = style.FillColor.Value.ToArgbHex();
                    int id;
                    if (_fillIds.TryGetValue(key, out id)) return id;
                    id = _fills.Count;
                    _fills.Add(new FillEntry(style.FillColor.Value));
                    _fillIds[key] = id;
                    return id;
                }

                /// <summary>按四边线组合键去重登记边框并返回索引。 Registers a border de-duplicated by the four-edge combination key and returns the index.</summary>
                /// <param name="style">单元格样式。 Cell style.</param>
                /// <returns>borders 集合索引。 Index into the borders collection.</returns>
                private int GetBorderId(CellStyle style)
                {
                    string key = BorderKey(style.LeftBorder) + "|" + BorderKey(style.RightBorder) + "|"
                        + BorderKey(style.TopBorder) + "|" + BorderKey(style.BottomBorder);
                    int id;
                    if (_borderIds.TryGetValue(key, out id)) return id;
                    id = _borders.Count;
                    _borders.Add(new BorderEntry(style.LeftBorder, style.RightBorder, style.TopBorder, style.BottomBorder));
                    _borderIds[key] = id;
                    return id;
                }

                /// <summary>生成单边线的去重键；无线或线型为 None 时返回空字符串。 Builds the de-duplication key of one border edge; returns an empty string when there is no line or the style is None.</summary>
                /// <param name="line">边线描述。 Border line description.</param>
                /// <returns>边线键文本。 The border key text.</returns>
                private static string BorderKey(BorderLine line)
                {
                    if (line == null || line.Style == BorderStyle.None) return string.Empty;
                    return line.Style.ToString() + "|" + Xlsx.ToInvariant(line.Width) + "|" + line.Color.ToArgbHex();
                }

                /// <summary>解析格式代码对应编号：General 为 0，命中内置格式用其编号，否则分配自定义编号（164 起）。 Resolves the id of a format code: General maps to 0, built-in formats reuse their id, otherwise a custom id is allocated starting at 164.</summary>
                /// <param name="format">格式代码。 Format code.</param>
                /// <returns>numFmtId 编号。 The numFmtId value.</returns>
                private int GetFormatId(string format)
                {
                    if (string.IsNullOrEmpty(format) || format == "General") return 0;
                    int id;
                    if (_formatIds.TryGetValue(format, out id)) return id;
                    int builtIn = XlsxNumberFormats.FindBuiltIn(format);
                    if (builtIn >= 0)
                    {
                        _formatIds[format] = builtIn;
                        return builtIn;
                    }
                    id = _nextFormatId++;
                    _formatIds[format] = id;
                    _customFormats.Add(new KeyValuePair<int, string>(id, format));
                    return id;
                }

                /// <summary>生成对齐方式（水平/垂直/自动换行）的去重键。 Builds the de-duplication key for alignment (horizontal/vertical/wrap text).</summary>
                /// <param name="style">单元格样式。 Cell style.</param>
                /// <returns>对齐键文本。 The alignment key text.</returns>
                private static string AlignmentKey(CellStyle style)
                {
                    return style.HorizontalAlignment.ToString() + "|" + style.VerticalAlignment.ToString() + "|" + (style.WrapText ? "1" : "0");
                }

                /// <summary>序列化整个样式表：numFmts/fonts/fills/borders/cellXfs 等集合按登记顺序输出为部件字节。 Serializes the whole style sheet: the numFmts/fonts/fills/borders/cellXfs collections are emitted in registration order as part bytes.</summary>
                /// <returns>styles.xml 的 UTF-8 字节内容。 The UTF-8 byte content of styles.xml.</returns>
                public byte[] ToXml()
                {
                    XElement root = new XElement(Xlsx.Main + "styleSheet");
                    if (_customFormats.Count > 0)
                    {
                        XElement formats = new XElement(Xlsx.Main + "numFmts",
                            new XAttribute("count", _customFormats.Count.ToString(CultureInfo.InvariantCulture)));
                        for (int i = 0; i < _customFormats.Count; i++)
                        {
                            formats.Add(new XElement(Xlsx.Main + "numFmt",
                                new XAttribute("numFmtId", _customFormats[i].Key.ToString(CultureInfo.InvariantCulture)),
                                new XAttribute("formatCode", _customFormats[i].Value)));
                        }
                        root.Add(formats);
                    }

                    XElement fonts = new XElement(Xlsx.Main + "fonts",
                        new XAttribute("count", _fonts.Count.ToString(CultureInfo.InvariantCulture)));
                    for (int i = 0; i < _fonts.Count; i++) fonts.Add(_fonts[i].ToXml());
                    root.Add(fonts);

                    XElement fills = new XElement(Xlsx.Main + "fills",
                        new XAttribute("count", _fills.Count.ToString(CultureInfo.InvariantCulture)));
                    for (int i = 0; i < _fills.Count; i++) fills.Add(_fills[i].ToXml(i == 0, i == 1));
                    root.Add(fills);

                    XElement borders = new XElement(Xlsx.Main + "borders",
                        new XAttribute("count", _borders.Count.ToString(CultureInfo.InvariantCulture)));
                    for (int i = 0; i < _borders.Count; i++) borders.Add(_borders[i].ToXml());
                    root.Add(borders);

                    root.Add(new XElement(Xlsx.Main + "cellStyleXfs",
                        new XAttribute("count", "1"),
                        new XElement(Xlsx.Main + "xf",
                            new XAttribute("numFmtId", "0"),
                            new XAttribute("fontId", "0"),
                            new XAttribute("fillId", "0"),
                            new XAttribute("borderId", "0"))));

                    XElement cellFormats = new XElement(Xlsx.Main + "cellXfs",
                        new XAttribute("count", _xfs.Count.ToString(CultureInfo.InvariantCulture)));
                    for (int i = 0; i < _xfs.Count; i++) cellFormats.Add(_xfs[i].ToXml());
                    root.Add(cellFormats);

                    root.Add(new XElement(Xlsx.Main + "cellStyles",
                        new XAttribute("count", "1"),
                        new XElement(Xlsx.Main + "cellStyle",
                            new XAttribute("name", "Normal"),
                            new XAttribute("xfId", "0"),
                            new XAttribute("builtinId", "0"))));
                    XElement differential = new XElement(Xlsx.Main + "dxfs",
                        new XAttribute("count", _dxfs.Count.ToString(CultureInfo.InvariantCulture)));
                    for (int i = 0; i < _dxfs.Count; i++) differential.Add(CreateDxf(_dxfs[i]));
                    root.Add(differential);
                    root.Add(new XElement(Xlsx.Main + "tableStyles",
                        new XAttribute("count", "0"),
                        new XAttribute("defaultTableStyle", "TableStyleMedium2"),
                        new XAttribute("defaultPivotStyle", "PivotStyleLight16")));

                    XDocument document = new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root);
                    return Xlsx.SerializeXml(document);
                }

                /// <summary>字体登记条目：保存去重后的字体属性并按 OOXML 顺序生成 font 元素。 Font registration entry: stores de-duplicated font attributes and builds the font element in OOXML order.</summary>
                private sealed class FontEntry
                {
                    /// <summary>字体族名。 Font family name.</summary>
                    private readonly string _name;
                    /// <summary>字号（磅）。 Font size in points.</summary>
                    private readonly double _size;
                    /// <summary>是否加粗。 Bold flag.</summary>
                    private readonly bool _bold;
                    /// <summary>是否斜体。 Italic flag.</summary>
                    private readonly bool _italic;
                    /// <summary>是否下划线。 Underline flag.</summary>
                    private readonly bool _underline;
                    /// <summary>字体颜色。 Font color.</summary>
                    private readonly OfficeColor _color;

                    /// <summary>创建字体条目；名称为空时回退 Calibri。 Creates a font entry; an empty name falls back to Calibri.</summary>
                    /// <param name="name">字体族名。 Font family name.</param>
                    /// <param name="size">字号（磅）。 Font size in points.</param>
                    /// <param name="bold">是否加粗。 Whether the font is bold.</param>
                    /// <param name="italic">是否斜体。 Whether the font is italic.</param>
                    /// <param name="underline">是否下划线。 Whether the font is underlined.</param>
                    /// <param name="color">字体颜色。 Font color.</param>
                    public FontEntry(string name, double size, bool bold, bool italic, bool underline, OfficeColor color)
                    {
                        _name = string.IsNullOrEmpty(name) ? "Calibri" : name;
                        _size = size;
                        _bold = bold;
                        _italic = italic;
                        _underline = underline;
                        _color = color;
                    }

                    /// <summary>生成 font 元素（sz/color/b/i/u/name/family）。 Builds the font element (sz/color/b/i/u/name/family).</summary>
                    /// <returns>font XML 元素。 The font XML element.</returns>
                    public XElement ToXml()
                    {
                        XElement font = new XElement(Xlsx.Main + "font",
                            new XElement(Xlsx.Main + "sz", new XAttribute("val", Xlsx.ToInvariant(_size))),
                            new XElement(Xlsx.Main + "color", new XAttribute("rgb", "FF" + _color.ToHex())));
                        if (_bold) font.Add(new XElement(Xlsx.Main + "b"));
                        if (_italic) font.Add(new XElement(Xlsx.Main + "i"));
                        if (_underline) font.Add(new XElement(Xlsx.Main + "u"));
                        font.Add(new XElement(Xlsx.Main + "name", new XAttribute("val", _name)));
                        font.Add(new XElement(Xlsx.Main + "family", new XAttribute("val", "2")));
                        return font;
                    }
                }

                /// <summary>填充登记条目；null 颜色表示占位填充。 Fill registration entry; a null color marks a placeholder fill.</summary>
                private sealed class FillEntry
                {
                    /// <summary>填充色；null 表示占位项。 Fill color; null means a placeholder entry.</summary>
                    private readonly OfficeColor? _color;

                    /// <summary>创建填充条目。 Creates a fill entry.</summary>
                    /// <param name="color">填充色；null 表示 none/gray125 占位项。 Fill color; null means the none/gray125 placeholder.</param>
                    public FillEntry(OfficeColor? color)
                    {
                        _color = color;
                    }

                    /// <summary>生成 fill 元素：前两位固定输出 none 与 gray125（Excel 兼容性要求），其余输出 solid 图案。 Builds the fill element: the first two entries always emit none and gray125 (required for Excel compatibility), the rest emit a solid pattern.</summary>
                    /// <param name="first">是否为 fills 集合第一项。 Whether this is the first entry of the fills collection.</param>
                    /// <param name="second">是否为 fills 集合第二项。 Whether this is the second entry of the fills collection.</param>
                    /// <returns>fill XML 元素。 The fill XML element.</returns>
                    public XElement ToXml(bool first, bool second)
                    {
                        if (first)
                        {
                            return new XElement(Xlsx.Main + "fill",
                                new XElement(Xlsx.Main + "patternFill", new XAttribute("patternType", "none")));
                        }
                        if (second)
                        {
                            return new XElement(Xlsx.Main + "fill",
                                new XElement(Xlsx.Main + "patternFill", new XAttribute("patternType", "gray125")));
                        }
                        return new XElement(Xlsx.Main + "fill",
                            new XElement(Xlsx.Main + "patternFill",
                                new XAttribute("patternType", "solid"),
                                new XElement(Xlsx.Main + "fgColor", new XAttribute("rgb", "FF" + _color.Value.ToHex())),
                                new XElement(Xlsx.Main + "bgColor", new XAttribute("indexed", "64"))));
                    }
                }

                /// <summary>边框登记条目：保存四边线并生成 border 元素。 Border registration entry: stores the four edges and builds the border element.</summary>
                private sealed class BorderEntry
                {
                    /// <summary>左边线。 Left edge.</summary>
                    private readonly BorderLine _left;
                    /// <summary>右边线。 Right edge.</summary>
                    private readonly BorderLine _right;
                    /// <summary>上边线。 Top edge.</summary>
                    private readonly BorderLine _top;
                    /// <summary>下边线。 Bottom edge.</summary>
                    private readonly BorderLine _bottom;

                    /// <summary>创建边框条目。 Creates a border entry.</summary>
                    /// <param name="left">左边线。 Left edge.</param>
                    /// <param name="right">右边线。 Right edge.</param>
                    /// <param name="top">上边线。 Top edge.</param>
                    /// <param name="bottom">下边线。 Bottom edge.</param>
                    public BorderEntry(BorderLine left, BorderLine right, BorderLine top, BorderLine bottom)
                    {
                        _left = left;
                        _right = right;
                        _top = top;
                        _bottom = bottom;
                    }

                    /// <summary>生成包含左右上下四边与对角线占位的 border 元素。 Builds the border element containing the four edges plus a diagonal placeholder.</summary>
                    /// <returns>border XML 元素。 The border XML element.</returns>
                    public XElement ToXml()
                    {
                        return new XElement(Xlsx.Main + "border",
                            CreateEdge("left", _left),
                            CreateEdge("right", _right),
                            CreateEdge("top", _top),
                            CreateEdge("bottom", _bottom),
                            new XElement(Xlsx.Main + "diagonal"));
                    }

                    /// <summary>生成单边元素：虚线/点线/双线按线型输出，实线按宽度映射 thin/medium/thick；无线时输出空元素。 Builds one edge element: dashed/dotted/double lines map by style, solid lines map thin/medium/thick by width; an empty element is emitted when there is no line.</summary>
                    /// <param name="name">边元素名（left/right/top/bottom）。 Edge element name (left/right/top/bottom).</param>
                    /// <param name="line">边线描述。 Border line description.</param>
                    /// <returns>边 XML 元素。 The edge XML element.</returns>
                    private static XElement CreateEdge(string name, BorderLine line)
                    {
                        XElement edge = new XElement(Xlsx.Main + name);
                        if (line == null || line.Style == BorderStyle.None) return edge;
                        string style;
                        if (line.Style == BorderStyle.Dashed) style = "dashed";
                        else if (line.Style == BorderStyle.Dotted) style = "dotted";
                        else if (line.Style == BorderStyle.Double) style = "double";
                        else if (line.Width >= 2) style = "thick";
                        else if (line.Width >= 1) style = "medium";
                        else style = "thin";
                        edge.Add(new XAttribute("style", style));
                        edge.Add(new XElement(Xlsx.Main + "color", new XAttribute("rgb", "FF" + line.Color.ToHex())));
                        return edge;
                    }
                }

                /// <summary>cellXfs 登记条目：组合字体/填充/边框/格式索引与对齐设置，生成 xf 元素。 cellXfs registration entry: combines font/fill/border/format indexes with alignment settings to build the xf element.</summary>
                private sealed class XfEntry
                {
                    /// <summary>字体索引。 Font index.</summary>
                    public int FontId;
                    /// <summary>填充索引。 Fill index.</summary>
                    public int FillId;
                    /// <summary>边框索引。 Border index.</summary>
                    public int BorderId;
                    /// <summary>数字格式编号。 Number-format id.</summary>
                    public int FormatId;
                    /// <summary>水平对齐。 Horizontal alignment.</summary>
                    public HorizontalAlignment Horizontal;
                    /// <summary>垂直对齐。 Vertical alignment.</summary>
                    public VerticalAlignment Vertical;
                    /// <summary>是否自动换行。 Wrap-text flag.</summary>
                    public bool WrapText;

                    /// <summary>生成 xf 元素：仅对非默认项输出 applyXxx 标记，非默认对齐时附加 alignment 子元素。 Builds the xf element: applyXxx flags are emitted only for non-default members, and an alignment child element is added for non-default alignment.</summary>
                    /// <returns>xf XML 元素。 The xf XML element.</returns>
                    public XElement ToXml()
                    {
                        XElement xf = new XElement(Xlsx.Main + "xf",
                            new XAttribute("numFmtId", FormatId.ToString(CultureInfo.InvariantCulture)),
                            new XAttribute("fontId", FontId.ToString(CultureInfo.InvariantCulture)),
                            new XAttribute("fillId", FillId.ToString(CultureInfo.InvariantCulture)),
                            new XAttribute("borderId", BorderId.ToString(CultureInfo.InvariantCulture)),
                            new XAttribute("xfId", "0"));
                        if (FormatId != 0) xf.Add(new XAttribute("applyNumberFormat", "1"));
                        if (FontId != 0) xf.Add(new XAttribute("applyFont", "1"));
                        if (FillId != 0 && FillId != 1) xf.Add(new XAttribute("applyFill", "1"));
                        if (BorderId != 0) xf.Add(new XAttribute("applyBorder", "1"));
                        bool alignment = Horizontal != HorizontalAlignment.General
                            || Vertical != VerticalAlignment.Bottom
                            || WrapText;
                        if (alignment)
                        {
                            xf.Add(new XAttribute("applyAlignment", "1"));
                            XElement element = new XElement(Xlsx.Main + "alignment");
                            if (Horizontal != HorizontalAlignment.General)
                            {
                                string horizontal = Horizontal == HorizontalAlignment.Center ? "center"
                                    : Horizontal == HorizontalAlignment.Right ? "right"
                                    : Horizontal == HorizontalAlignment.Justify ? "justify" : "left";
                                element.Add(new XAttribute("horizontal", horizontal));
                            }
                            if (Vertical != VerticalAlignment.Bottom)
                            {
                                string vertical = Vertical == VerticalAlignment.Middle ? "center" : "top";
                                element.Add(new XAttribute("vertical", vertical));
                            }
                            if (WrapText) element.Add(new XAttribute("wrapText", "1"));
                            xf.Add(element);
                        }
                        return xf;
                    }
                }
            }
        }
    }
}
