using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>CSV 导入导出的分隔符、编码与 BOM 选项。Separator, encoding and BOM options for CSV import and export.</summary>
    public sealed class ExcelCsvOptions
    {
        /// <summary>字段分隔符，默认逗号；不能是双引号或换行符。The field separator, a comma by default; it cannot be a quote or a line break.</summary>
        public char Separator { get; set; }
        /// <summary>写文件时是否输出 UTF-8 BOM，默认输出（便于 Excel 正确识别中文）。Whether a UTF-8 BOM is written, enabled by default so Excel detects Chinese text correctly.</summary>
        public bool WriteBom { get; set; }
        /// <summary>显式编码；为 null 时使用 UTF-8。读取时仍会自动识别 BOM。An explicit encoding; UTF-8 is used when null. Reading still detects a BOM.</summary>
        public Encoding Encoding { get; set; }
        /// <summary>读取时的工作表名称，默认 "Sheet1"。The worksheet name used when reading, "Sheet1" by default.</summary>
        public string SheetName { get; set; }

        /// <summary>创建默认选项：逗号分隔、输出 BOM、UTF-8、"Sheet1"。Creates default options: comma-separated, BOM written, UTF-8 and "Sheet1".</summary>
        public ExcelCsvOptions()
        {
            Separator = ',';
            WriteBom = true;
            SheetName = "Sheet1";
        }
    }

    /// <summary>
    /// 轻量 CSV 读写：按 RFC 4180 处理引号与转义（双引号翻倍、字段内可含分隔符与换行），
    /// 写出采用流式按行输出，读取逐条记录解析并限制列数/字段长度/行数。
    /// Lightweight CSV support: handles RFC 4180 quoting (doubled quotes, separators and newlines inside
    /// fields), streams the worksheet out row by row, and parses records with column, field-length and row limits.
    /// </summary>
    public static class ExcelCsv
    {
        /// <summary>CSV 允许的最大列数（与 Excel 网格一致）。The maximum number of CSV columns (matching the Excel grid).</summary>
        private const int MaxColumns = 16384;
        /// <summary>单个字段的最大字符数。The maximum number of characters in one field.</summary>
        private const int MaxFieldChars = 4 * 1024 * 1024;
        /// <summary>CSV 允许的最大行数（与 Excel 网格一致）。The maximum number of CSV rows (matching the Excel grid).</summary>
        private const int MaxRows = 1048576;

        /// <summary>把工作簿的第一个工作表写为 CSV 文件。Writes the first worksheet of a workbook to a CSV file.</summary>
        /// <param name="workbook">源工作簿。The source workbook.</param>
        /// <param name="path">目标文件路径。The target file path.</param>
        /// <param name="options">分隔符与编码选项。The separator and encoding options.</param>
        /// <exception cref="ArgumentNullException">workbook 或 path 为 null。workbook or path is null.</exception>
        /// <exception cref="ArgumentException">工作簿没有工作表。The workbook has no worksheet.</exception>
        public static void Write(OfficeWorkbook workbook, string path, ExcelCsvOptions options = null)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            if (workbook.Worksheets.Count == 0) throw new ArgumentException("The workbook has no worksheet.", "workbook");
            Write(workbook.Worksheets[0], path, options);
        }

        /// <summary>把工作表写为 CSV 文件。Writes a worksheet to a CSV file.</summary>
        /// <param name="sheet">源工作表。The source worksheet.</param>
        /// <param name="path">目标文件路径。The target file path.</param>
        /// <param name="options">分隔符与编码选项。The separator and encoding options.</param>
        public static void Write(Worksheet sheet, string path, ExcelCsvOptions options = null)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            if (path == null) throw new ArgumentNullException("path");
            ExcelCsvOptions effective = options ?? new ExcelCsvOptions();
            Encoding encoding = effective.Encoding ?? new UTF8Encoding(effective.WriteBom);
            using (StreamWriter writer = new StreamWriter(path, false, encoding))
            {
                Write(sheet, writer, effective);
            }
        }

        /// <summary>把工作表逐行写入文本写出器（不写 BOM）。Writes a worksheet row by row into a text writer (no BOM is written).</summary>
        /// <param name="sheet">源工作表。The source worksheet.</param>
        /// <param name="writer">目标文本写出器。The target text writer.</param>
        /// <param name="options">分隔符选项。The separator options.</param>
        public static void Write(Worksheet sheet, TextWriter writer, ExcelCsvOptions options = null)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            if (writer == null) throw new ArgumentNullException("writer");
            char separator = ValidateSeparator(options == null ? ',' : options.Separator);
            int lastRow = sheet.LastRow;
            IEnumerator<Cell> cells = sheet.Cells.GetEnumerator();
            bool hasCell = cells.MoveNext();
            for (int row = 1; row <= lastRow; row++)
            {
                StringBuilder line = new StringBuilder();
                bool first = true;
                int nextColumn = 1;
                while (hasCell && cells.Current.Row == row)
                {
                    Cell cell = cells.Current;
                    if (cell.Value != null)
                    {
                        while (nextColumn < cell.Column)
                        {
                            WriteField(line, separator, string.Empty, ref first);
                            nextColumn++;
                        }
                        WriteField(line, separator, FormatValue(cell.Value), ref first);
                        nextColumn = cell.Column + 1;
                    }
                    hasCell = cells.MoveNext();
                }
                writer.Write(line.ToString());
                writer.Write("\r\n");
            }
            cells.Dispose();
        }

        /// <summary>读取 CSV 文件为单工作表工作簿（数值与布尔自动转换，其余保持文本）。Reads a CSV file into a single-sheet workbook (numbers and booleans are converted, everything else stays text).</summary>
        /// <param name="path">CSV 文件路径。The CSV file path.</param>
        /// <param name="options">分隔符、编码与表名选项。The separator, encoding and sheet-name options.</param>
        /// <returns>只含一个工作表的工作簿。A workbook with one worksheet.</returns>
        public static OfficeWorkbook Read(string path, ExcelCsvOptions options = null)
        {
            if (path == null) throw new ArgumentNullException("path");
            ExcelCsvOptions effective = options ?? new ExcelCsvOptions();
            using (StreamReader reader = new StreamReader(path, effective.Encoding ?? Encoding.UTF8, true))
            {
                return Read(reader, effective);
            }
        }

        /// <summary>从文本读取器解析 CSV；支持引号字段内换行，空字段不创建单元格。Parses CSV from a text reader; newlines inside quoted fields are supported and empty fields create no cell.</summary>
        /// <param name="reader">源文本读取器。The source text reader.</param>
        /// <param name="options">分隔符与表名选项。The separator and sheet-name options.</param>
        /// <returns>只含一个工作表的工作簿。A workbook with one worksheet.</returns>
        public static OfficeWorkbook Read(TextReader reader, ExcelCsvOptions options = null)
        {
            if (reader == null) throw new ArgumentNullException("reader");
            ExcelCsvOptions effective = options ?? new ExcelCsvOptions();
            char separator = ValidateSeparator(effective.Separator);
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet(string.IsNullOrEmpty(effective.SheetName) ? "Sheet1" : effective.SheetName);
            int row = 1;
            foreach (List<CsvField> record in ReadRecords(reader, separator))
            {
                if (row > MaxRows) throw new InvalidDataException("The CSV file exceeds the Excel row limit.");
                if (record.Count > MaxColumns) throw new InvalidDataException("The CSV file exceeds the Excel column limit.");
                for (int i = 0; i < record.Count; i++)
                {
                    CsvField field = record[i];
                    if (field.Text.Length > MaxFieldChars) throw new InvalidDataException("A CSV field exceeds the maximum length.");
                    if (!field.Quoted && field.Text.Length == 0) continue;
                    sheet.GetCell(row, i + 1).Value = field.Quoted ? field.Text : ParseValue(field.Text);
                }
                row++;
            }
            return workbook;
        }

        /// <summary>逐条解析 CSV 记录：双引号内的分隔符与换行原样保留，双写引号还原为一个引号。Parses CSV records one at a time: separators and newlines inside quotes are kept and doubled quotes collapse into one.</summary>
        /// <param name="reader">源文本读取器。The source text reader.</param>
        /// <param name="separator">字段分隔符。The field separator.</param>
        /// <returns>逐条记录的字段列表。The field list of each record.</returns>
        private static IEnumerable<List<CsvField>> ReadRecords(TextReader reader, char separator)
        {
            List<CsvField> fields = new List<CsvField>();
            StringBuilder field = new StringBuilder();
            bool quoted = false;
            bool inQuotes = false;
            int value;
            while ((value = reader.Read()) >= 0)
            {
                char c = (char)value;
                if (inQuotes)
                {
                    if (c == '"')
                    {
                        if (reader.Peek() == '"')
                        {
                            reader.Read();
                            field.Append('"');
                        }
                        else
                        {
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        field.Append(c);
                    }
                    continue;
                }
                if (c == '"')
                {
                    if (field.Length == 0)
                    {
                        inQuotes = true;
                        quoted = true;
                    }
                    else
                    {
                        field.Append(c);
                    }
                    continue;
                }
                if (c == separator)
                {
                    fields.Add(new CsvField(field.ToString(), quoted));
                    field.Length = 0;
                    quoted = false;
                    continue;
                }
                if (c == '\r' || c == '\n')
                {
                    if (c == '\r' && reader.Peek() == '\n') reader.Read();
                    fields.Add(new CsvField(field.ToString(), quoted));
                    yield return fields;
                    fields = new List<CsvField>();
                    field.Length = 0;
                    quoted = false;
                    continue;
                }
                field.Append(c);
            }
            if (field.Length > 0 || quoted || fields.Count > 0)
            {
                fields.Add(new CsvField(field.ToString(), quoted));
                yield return fields;
            }
        }

        /// <summary>校验分隔符不能是引号或换行符。Validates that the separator is neither a quote nor a line break.</summary>
        /// <param name="separator">候选分隔符。The candidate separator.</param>
        /// <returns>原分隔符。The separator.</returns>
        /// <exception cref="ArgumentException">分隔符非法。The separator is invalid.</exception>
        private static char ValidateSeparator(char separator)
        {
            if (separator == '"' || separator == '\r' || separator == '\n')
            {
                throw new ArgumentException("The CSV separator cannot be a quote or a line break.", "separator");
            }
            return separator;
        }

        /// <summary>按 CSV 规则写出一个字段：含分隔符、引号、换行或首尾空白时用引号包裹，内部引号翻倍。Writes one CSV field, quoting it when it contains the separator, a quote, a line break or surrounding whitespace, and doubling inner quotes.</summary>
        /// <param name="line">目标行缓冲。The target line buffer.</param>
        /// <param name="separator">字段分隔符。The field separator.</param>
        /// <param name="value">字段文本。The field text.</param>
        /// <param name="first">是否为本行第一个字段（负责补分隔符）。Whether this is the first field of the row (it writes the leading separator otherwise).</param>
        private static void WriteField(StringBuilder line, char separator, string value, ref bool first)
        {
            if (!first) line.Append(separator);
            first = false;
            if (value == null) return;
            bool quote = false;
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c == separator || c == '"' || c == '\r' || c == '\n')
                {
                    quote = true;
                    break;
                }
            }
            if (!quote && value.Length > 0)
            {
                quote = char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]);
            }
            if (!quote)
            {
                line.Append(value);
                return;
            }
            line.Append('"');
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                if (c == '"') line.Append('"');
                line.Append(c);
            }
            line.Append('"');
        }

        /// <summary>把单元格值格式化为 CSV 文本：日期用 ISO 时间、布尔用 TRUE/FALSE、数值用不变区域格式。Formats a cell value as CSV text: dates use ISO timestamps, booleans use TRUE/FALSE and numbers use the invariant culture.</summary>
        /// <param name="value">单元格值。The cell value.</param>
        /// <returns>字段文本。The field text.</returns>
        private static string FormatValue(object value)
        {
            if (value is DateTime)
            {
                DateTime date = (DateTime)value;
                return date.TimeOfDay == TimeSpan.Zero
                    ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
                    : date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            }
            if (value is bool) return (bool)value ? "TRUE" : "FALSE";
            if (value is double) return ((double)value).ToString("G17", CultureInfo.InvariantCulture);
            if (value is float) return ((float)value).ToString("G9", CultureInfo.InvariantCulture);
            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        /// <summary>解析未加引号的字段：整数、浮点与布尔（TRUE/FALSE）自动转换，其余保持字符串。Parses an unquoted field: integers, floating-point values and booleans (TRUE/FALSE) are converted, everything else stays a string.</summary>
        /// <param name="text">字段文本。The field text.</param>
        /// <returns>解析后的值。The parsed value.</returns>
        private static object ParseValue(string text)
        {
            if (string.IsNullOrEmpty(text)) return null;
            if (text.Equals("TRUE", StringComparison.OrdinalIgnoreCase)) return true;
            if (text.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) return false;
            long integer;
            if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out integer))
            {
                return integer >= int.MinValue && integer <= int.MaxValue ? (object)(int)integer : integer;
            }
            double number;
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
            {
                return number;
            }
            return text;
        }

        /// <summary>一个 CSV 字段的文本与是否带引号。The text of one CSV field and whether it was quoted.</summary>
        private struct CsvField
        {
            /// <summary>字段文本。The field text.</summary>
            public readonly string Text;
            /// <summary>是否由引号包裹。Whether the field was quoted.</summary>
            public readonly bool Quoted;

            /// <summary>创建字段。Creates a field.</summary>
            /// <param name="text">字段文本。The field text.</param>
            /// <param name="quoted">是否带引号。Whether the field was quoted.</param>
            public CsvField(string text, bool quoted)
            {
                Text = text;
                Quoted = quoted;
            }
        }
    }
}
