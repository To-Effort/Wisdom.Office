using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Excel
{
    /// <summary>
    /// 标在属性或字段上的列映射特性：自定义表头、顺序与数字格式，或排除该成员。
    /// A column-mapping attribute for a property or field: custom header, order and number format, or exclusion.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property | AttributeTargets.Field, AllowMultiple = false)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        /// <summary>创建不带表头名称的映射（使用成员名）。Creates a mapping without a header name (the member name is used).</summary>
        public ExcelColumnAttribute()
        {
        }

        /// <summary>以指定表头名称创建映射。Creates a mapping with the given header name.</summary>
        /// <param name="name">表头名称。The header name.</param>
        public ExcelColumnAttribute(string name)
        {
            Name = name;
        }

        /// <summary>表头名称；为空时使用成员名。The header name; the member name is used when empty.</summary>
        public string Name { get; set; }
        /// <summary>列顺序，从 1 开始从小到大排列；未指定（0）时排在最后。The column order from 1 upward; unspecified (0) columns come last.</summary>
        public int Order { get; set; }
        /// <summary>写出时的 Excel 数字/日期格式代码。The Excel number/date format code used when writing.</summary>
        public string Format { get; set; }
        /// <summary>是否忽略该成员。Whether the member is ignored.</summary>
        public bool Ignore { get; set; }
    }

    /// <summary>集合映射与 CSV 共用的行列位置选项。Row and column position options shared by collection mapping and CSV operations.</summary>
    public sealed class ExcelCollectionOptions
    {
        /// <summary>写出时是否包含表头行，默认包含。Whether a header row is written; enabled by default.</summary>
        public bool IncludeHeader { get; set; }
        /// <summary>读取时是否把首行当作表头，默认是；关闭时按成员顺序逐列映射。Whether the first row is treated as the header when reading; enabled by default. When disabled, columns are mapped by member order.</summary>
        public bool FirstRowIsHeader { get; set; }
        /// <summary>起始行号（从 1 开始），默认 1。The starting row number (1-based), 1 by default.</summary>
        public int StartRow { get; set; }
        /// <summary>起始列号（从 1 开始），默认 1。The starting column number (1-based), 1 by default.</summary>
        public int StartColumn { get; set; }

        /// <summary>创建默认选项：包含表头、首行是表头、从 A1 开始。Creates default options: headers included, first row treated as header, starting at A1.</summary>
        public ExcelCollectionOptions()
        {
            IncludeHeader = true;
            FirstRowIsHeader = true;
            StartRow = 1;
            StartColumn = 1;
        }
    }

    /// <summary>一个成员的列映射：表头、格式、取值与赋值方式。The column mapping of one member: header, format and accessors.</summary>
    internal sealed class ExcelColumnMap
    {
        /// <summary>表头名称。The header name.</summary>
        public string Name;
        /// <summary>写出时的数字格式代码。The number format used when writing.</summary>
        public string Format;
        /// <summary>成员类型。The member type.</summary>
        public Type ValueType;
        /// <summary>属性访问器；字段映射时为 null。The property accessor; null for field mappings.</summary>
        public PropertyInfo Property;
        /// <summary>字段访问器；属性映射时为 null。The field accessor; null for property mappings.</summary>
        public FieldInfo Field;

        /// <summary>读取成员值。Reads the member value.</summary>
        /// <param name="target">目标对象。The target object.</param>
        /// <returns>成员值。The member value.</returns>
        public object GetValue(object target)
        {
            return Property != null ? Property.GetValue(target, null) : Field.GetValue(target);
        }

        /// <summary>写入成员值。Writes the member value.</summary>
        /// <param name="target">目标对象。The target object.</param>
        /// <param name="value">新值。The new value.</param>
        public void SetValue(object target, object value)
        {
            if (Property != null) Property.SetValue(target, value, null);
            else Field.SetValue(target, value);
        }
    }

    /// <summary>
    /// MiniExcel/EPPlus 风格的集合映射：把对象集合或 DataTable 写入工作表，或把工作表反序列化为对象列表/DataTable。
    /// MiniExcel/EPPlus-style collection mapping: writes an object collection or DataTable into a worksheet,
    /// or deserializes a worksheet into object lists or a DataTable.
    /// </summary>
    public static class ExcelCollection
    {
        /// <summary>类型到列映射的缓存。Cache from a type to its column mappings.</summary>
        private static readonly ConcurrentDictionary<Type, ExcelColumnMap[]> MapCache =
            new ConcurrentDictionary<Type, ExcelColumnMap[]>();

        /// <summary>取得类型的列映射（按特性顺序排序并缓存）。Gets the column mappings of a type, ordered by the attribute and cached.</summary>
        /// <param name="type">目标类型。The target type.</param>
        /// <returns>列映射数组。The column mappings.</returns>
        internal static ExcelColumnMap[] GetMaps(Type type)
        {
            return MapCache.GetOrAdd(type, BuildMaps);
        }

        /// <summary>通过反射构建列映射：公开可读属性与非只读字段，支持 ExcelColumn 的名称/顺序/格式/忽略。Builds column mappings through reflection: public readable properties and writable fields, honoring ExcelColumn name, order, format and ignore.</summary>
        /// <param name="type">目标类型。The target type.</param>
        /// <returns>列映射数组。The column mappings.</returns>
        private static ExcelColumnMap[] BuildMaps(Type type)
        {
            List<MapCandidate> candidates = new List<MapCandidate>();
            int sequence = 0;
            PropertyInfo[] properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < properties.Length; i++)
            {
                PropertyInfo property = properties[i];
                if (!property.CanRead || !property.CanWrite || property.GetIndexParameters().Length > 0) continue;
                ExcelColumnAttribute attribute = (ExcelColumnAttribute)Attribute.GetCustomAttribute(property, typeof(ExcelColumnAttribute), true);
                if (attribute != null && attribute.Ignore) continue;
                ExcelColumnMap map = new ExcelColumnMap();
                map.Name = attribute == null || string.IsNullOrEmpty(attribute.Name) ? property.Name : attribute.Name;
                map.Format = attribute == null ? null : attribute.Format;
                map.ValueType = property.PropertyType;
                map.Property = property;
                candidates.Add(new MapCandidate(map, attribute == null || attribute.Order == 0 ? int.MaxValue : attribute.Order, sequence++));
            }
            FieldInfo[] fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < fields.Length; i++)
            {
                FieldInfo field = fields[i];
                if (field.IsInitOnly || field.IsLiteral) continue;
                ExcelColumnAttribute attribute = (ExcelColumnAttribute)Attribute.GetCustomAttribute(field, typeof(ExcelColumnAttribute), true);
                if (attribute != null && attribute.Ignore) continue;
                ExcelColumnMap map = new ExcelColumnMap();
                map.Name = attribute == null || string.IsNullOrEmpty(attribute.Name) ? field.Name : attribute.Name;
                map.Format = attribute == null ? null : attribute.Format;
                map.ValueType = field.FieldType;
                map.Field = field;
                candidates.Add(new MapCandidate(map, attribute == null || attribute.Order == 0 ? int.MaxValue : attribute.Order, sequence++));
            }
            candidates.Sort(CompareCandidates);
            ExcelColumnMap[] maps = new ExcelColumnMap[candidates.Count];
            for (int i = 0; i < candidates.Count; i++) maps[i] = candidates[i].Map;
            return maps;
        }

        /// <summary>按 Order 升序、再按发现顺序比较两个候选映射。Compares two candidate mappings by Order, then by discovery sequence.</summary>
        /// <param name="left">左候选。The left candidate.</param>
        /// <param name="right">右候选。The right candidate.</param>
        /// <returns>比较结果。The comparison result.</returns>
        private static int CompareCandidates(MapCandidate left, MapCandidate right)
        {
            int order = left.Order.CompareTo(right.Order);
            return order != 0 ? order : left.Sequence.CompareTo(right.Sequence);
        }

        /// <summary>构建列映射时的排序候选。A sort candidate used while building the column mappings.</summary>
        private sealed class MapCandidate
        {
            /// <summary>列映射。The column mapping.</summary>
            public readonly ExcelColumnMap Map;
            /// <summary>显式顺序。The explicit order.</summary>
            public readonly int Order;
            /// <summary>发现顺序，用于稳定排序。The discovery sequence used for a stable order.</summary>
            public readonly int Sequence;

            /// <summary>创建候选。Creates a candidate.</summary>
            /// <param name="map">列映射。The column mapping.</param>
            /// <param name="order">显式顺序。The explicit order.</param>
            /// <param name="sequence">发现顺序。The discovery sequence.</param>
            public MapCandidate(ExcelColumnMap map, int order, int sequence)
            {
                Map = map;
                Order = order;
                Sequence = sequence;
            }
        }

        /// <summary>把对象集合写入新工作表并返回该表。Writes an object collection into a new worksheet and returns it.</summary>
        /// <typeparam name="T">元素类型。The element type.</typeparam>
        /// <param name="workbook">目标工作簿。The target workbook.</param>
        /// <param name="items">对象集合。The object collection.</param>
        /// <param name="sheetName">工作表名称；为空时使用类型名。The sheet name; the type name is used when empty.</param>
        /// <param name="options">行列位置选项。The position options.</param>
        /// <returns>写入的工作表。The written worksheet.</returns>
        /// <exception cref="ArgumentNullException">workbook 或 items 为 null。workbook or items is null.</exception>
        public static Worksheet Write<T>(OfficeWorkbook workbook, IEnumerable<T> items, string sheetName = null, ExcelCollectionOptions options = null)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            if (items == null) throw new ArgumentNullException("items");
            Worksheet sheet = workbook.AddWorksheet(string.IsNullOrEmpty(sheetName) ? typeof(T).Name : sheetName);
            LoadFromCollection(sheet, items, options);
            return sheet;
        }

        /// <summary>把对象集合写入现有工作表，可控制表头与起始位置。Writes an object collection into an existing worksheet with configurable header and start position.</summary>
        /// <typeparam name="T">元素类型。The element type.</typeparam>
        /// <param name="sheet">目标工作表。The target worksheet.</param>
        /// <param name="items">对象集合。The object collection.</param>
        /// <param name="options">行列位置选项。The position options.</param>
        public static void LoadFromCollection<T>(Worksheet sheet, IEnumerable<T> items, ExcelCollectionOptions options = null)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            if (items == null) throw new ArgumentNullException("items");
            ExcelCollectionOptions effective = options ?? new ExcelCollectionOptions();
            if (effective.StartRow < 1) throw new ArgumentOutOfRangeException("options");
            if (effective.StartColumn < 1 || effective.StartColumn > CellAddress.MaxColumnLiteral) throw new ArgumentOutOfRangeException("options");
            ExcelColumnMap[] maps = GetMaps(typeof(T));
            int row = effective.StartRow;
            int column = effective.StartColumn;
            if (effective.IncludeHeader)
            {
                for (int i = 0; i < maps.Length; i++)
                {
                    sheet.GetCell(row, column + i).Value = maps[i].Name;
                }
                row++;
            }
            if ((long)row > CellAddress.MaxRowLiteral) throw new ArgumentOutOfRangeException("options");
            foreach (T item in items)
            {
                if (item != null)
                {
                    for (int i = 0; i < maps.Length; i++)
                    {
                        object value = maps[i].GetValue(item);
                        if (value == null) continue;
                        Cell cell = sheet.GetCell(row, column + i);
                        cell.Value = ToCellValue(value);
                        ApplyFormat(cell, maps[i]);
                    }
                }
                row++;
                if ((long)row > CellAddress.MaxRowLiteral) throw new InvalidOperationException("The collection exceeds the Excel row limit.");
            }
        }

        /// <summary>把 DataTable 写入新工作表并返回该表。Writes a DataTable into a new worksheet and returns it.</summary>
        /// <param name="workbook">目标工作簿。The target workbook.</param>
        /// <param name="table">数据表。The data table.</param>
        /// <param name="sheetName">工作表名称；为空时使用 DataTable 名称。The sheet name; the DataTable name is used when empty.</param>
        /// <param name="options">行列位置选项。The position options.</param>
        /// <returns>写入的工作表。The written worksheet.</returns>
        public static Worksheet Write(OfficeWorkbook workbook, DataTable table, string sheetName = null, ExcelCollectionOptions options = null)
        {
            if (workbook == null) throw new ArgumentNullException("workbook");
            if (table == null) throw new ArgumentNullException("table");
            string name = string.IsNullOrEmpty(sheetName) ? table.TableName : sheetName;
            Worksheet sheet = workbook.AddWorksheet(string.IsNullOrEmpty(name) ? "Sheet1" : name);
            LoadFromDataTable(sheet, table, options);
            return sheet;
        }

        /// <summary>把 DataTable 写入现有工作表。Writes a DataTable into an existing worksheet.</summary>
        /// <param name="sheet">目标工作表。The target worksheet.</param>
        /// <param name="table">数据表。The data table.</param>
        /// <param name="options">行列位置选项。The position options.</param>
        public static void LoadFromDataTable(Worksheet sheet, DataTable table, ExcelCollectionOptions options = null)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            if (table == null) throw new ArgumentNullException("table");
            ExcelCollectionOptions effective = options ?? new ExcelCollectionOptions();
            if (effective.StartRow < 1) throw new ArgumentOutOfRangeException("options");
            if (effective.StartColumn < 1 || effective.StartColumn > CellAddress.MaxColumnLiteral) throw new ArgumentOutOfRangeException("options");
            int row = effective.StartRow;
            int column = effective.StartColumn;
            if (effective.IncludeHeader)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    sheet.GetCell(row, column + i).Value = table.Columns[i].ColumnName;
                }
                row++;
            }
            foreach (DataRow dataRow in table.Rows)
            {
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    object value = dataRow[i];
                    if (value == null || value == DBNull.Value) continue;
                    Cell cell = sheet.GetCell(row, column + i);
                    cell.Value = ToCellValue(value);
                    ApplyFormat(cell, table.Columns[i].DataType, null);
                }
                row++;
                if ((long)row > CellAddress.MaxRowLiteral) throw new InvalidOperationException("The table exceeds the Excel row limit.");
            }
        }

        /// <summary>读取工作表为对象列表：默认用首行表头按名称匹配，找不到的成员留默认值。Reads a worksheet into an object list: the first row is matched by header name by default and unmatched members keep their default values.</summary>
        /// <typeparam name="T">元素类型，需有无参构造函数。The element type; it needs a parameterless constructor.</typeparam>
        /// <param name="sheet">源工作表。The source worksheet.</param>
        /// <param name="options">行列位置选项。The position options.</param>
        /// <returns>对象列表。The object list.</returns>
        public static List<T> ToList<T>(Worksheet sheet, ExcelCollectionOptions options = null) where T : new()
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            ExcelCollectionOptions effective = options ?? new ExcelCollectionOptions();
            if (effective.StartRow < 1) throw new ArgumentOutOfRangeException("options");
            if (effective.StartColumn < 1) throw new ArgumentOutOfRangeException("options");
            ExcelColumnMap[] maps = GetMaps(typeof(T));
            int[] columns = ResolveColumns(sheet, maps, effective);
            int firstDataRow = effective.StartRow + (effective.FirstRowIsHeader ? 1 : 0);
            List<T> result = new List<T>();
            for (int row = firstDataRow; row <= sheet.LastRow; row++)
            {
                T item = new T();
                bool any = false;
                for (int i = 0; i < maps.Length; i++)
                {
                    if (columns[i] < 1) continue;
                    Cell cell = sheet.FindCell(row, columns[i]);
                    if (cell == null || cell.Value == null) continue;
                    object converted;
                    if (!TryConvertValue(cell.Value, maps[i].ValueType, out converted)) continue;
                    maps[i].SetValue(item, converted);
                    any = true;
                }
                if (any) result.Add(item);
            }
            return result;
        }

        /// <summary>读取工作表为 DataTable：默认首行做列名，空表返回空表。Reads a worksheet into a DataTable; the first row supplies column names by default and an empty sheet yields an empty table.</summary>
        /// <param name="sheet">源工作表。The source worksheet.</param>
        /// <param name="options">行列位置选项。The position options.</param>
        /// <returns>数据表。The data table.</returns>
        public static DataTable ToDataTable(Worksheet sheet, ExcelCollectionOptions options = null)
        {
            if (sheet == null) throw new ArgumentNullException("sheet");
            ExcelCollectionOptions effective = options ?? new ExcelCollectionOptions();
            if (effective.StartRow < 1) throw new ArgumentOutOfRangeException("options");
            if (effective.StartColumn < 1) throw new ArgumentOutOfRangeException("options");
            DataTable table = new DataTable(sheet.Name);
            int firstDataRow = effective.StartRow + (effective.FirstRowIsHeader ? 1 : 0);
            int lastColumn = Math.Max(effective.StartColumn, sheet.LastColumn);
            List<string> names = new List<string>();
            HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int column = effective.StartColumn; column <= lastColumn; column++)
            {
                string name = null;
                if (effective.FirstRowIsHeader)
                {
                    Cell header = sheet.FindCell(effective.StartRow, column);
                    if (header != null && header.Value != null)
                    {
                        name = Convert.ToString(header.Value, CultureInfo.InvariantCulture);
                        if (name != null) name = name.Trim();
                    }
                }
                if (string.IsNullOrEmpty(name)) name = "Column" + (column - effective.StartColumn + 1).ToString(CultureInfo.InvariantCulture);
                string unique = name;
                int suffix = 2;
                while (!used.Add(unique))
                {
                    unique = name + suffix.ToString(CultureInfo.InvariantCulture);
                    suffix++;
                }
                names.Add(unique);
                table.Columns.Add(unique, typeof(object));
            }
            for (int row = firstDataRow; row <= sheet.LastRow; row++)
            {
                DataRow dataRow = table.NewRow();
                bool any = false;
                for (int i = 0; i < names.Count; i++)
                {
                    Cell cell = sheet.FindCell(row, effective.StartColumn + i);
                    if (cell == null || cell.Value == null) continue;
                    dataRow[i] = cell.Value;
                    any = true;
                }
                if (any) table.Rows.Add(dataRow);
            }
            return table;
        }

        /// <summary>把对象值转换为可写入工作表的类型：枚举写名称、Guid/TimeSpan 写字符串，其余保持原类型。Converts an object value into a worksheet-friendly type: enums become names, Guid and TimeSpan become strings, others keep their type.</summary>
        /// <param name="value">对象值。The object value.</param>
        /// <returns>单元格值。The cell value.</returns>
        internal static object ToCellValue(object value)
        {
            if (value == null) return null;
            Type type = value.GetType();
            if (type.IsEnum) return value.ToString();
            if (value is Guid || value is TimeSpan) return Convert.ToString(value, CultureInfo.InvariantCulture);
            return value;
        }

        /// <summary>按列映射的格式代码设置数字/日期格式；日期值未指定格式时补默认日期格式。Applies the mapping's format code as the number/date format, adding a default date format for DateTime values without one.</summary>
        /// <param name="cell">目标单元格。The target cell.</param>
        /// <param name="map">列映射。The column mapping.</param>
        private static void ApplyFormat(Cell cell, ExcelColumnMap map)
        {
            ApplyFormat(cell, map.ValueType, map.Format);
        }

        /// <summary>按值类型与格式代码设置数字/日期格式。Applies a number/date format from the value type and format code.</summary>
        /// <param name="cell">目标单元格。The target cell.</param>
        /// <param name="valueType">值类型。The value type.</param>
        /// <param name="format">显式格式代码。The explicit format code.</param>
        private static void ApplyFormat(Cell cell, Type valueType, string format)
        {
            if (!string.IsNullOrEmpty(format))
            {
                cell.EnsureStyle().NumberFormat = format;
                return;
            }
            Type underlying = Nullable.GetUnderlyingType(valueType) ?? valueType;
            if (underlying == typeof(DateTime) || underlying == typeof(DateTimeOffset))
            {
                cell.EnsureStyle().NumberFormat = "yyyy-mm-dd hh:mm:ss";
            }
        }

        /// <summary>解析读取时的列位置：首行做表头时按名称匹配，否则按成员顺序连续排列。Resolves read columns: header-name matching when the first row is a header, otherwise consecutive columns in member order.</summary>
        /// <param name="sheet">源工作表。The source worksheet.</param>
        /// <param name="maps">列映射。The column mappings.</param>
        /// <param name="options">选项。The options.</param>
        /// <returns>每个映射对应的列号（0 表示未映射）。The column number of each mapping (0 means unmapped).</returns>
        private static int[] ResolveColumns(Worksheet sheet, ExcelColumnMap[] maps, ExcelCollectionOptions options)
        {
            int[] columns = new int[maps.Length];
            if (!options.FirstRowIsHeader)
            {
                for (int i = 0; i < maps.Length; i++) columns[i] = options.StartColumn + i;
                return columns;
            }
            Dictionary<string, int> headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int lastColumn = Math.Max(options.StartColumn, sheet.LastColumn);
            for (int column = options.StartColumn; column <= lastColumn; column++)
            {
                Cell cell = sheet.FindCell(options.StartRow, column);
                if (cell == null || cell.Value == null) continue;
                string name = Convert.ToString(cell.Value, CultureInfo.InvariantCulture);
                if (string.IsNullOrEmpty(name)) continue;
                name = name.Trim();
                if (!headers.ContainsKey(name)) headers[name] = column;
            }
            for (int i = 0; i < maps.Length; i++)
            {
                int column;
                if (headers.TryGetValue(maps[i].Name, out column)) columns[i] = column;
            }
            return columns;
        }

        /// <summary>把单元格值转换为目标成员类型；支持可空类型、枚举、Guid、TimeSpan 与常见数值/日期/布尔文本。Converts a cell value into the target member type, supporting nullable types, enums, Guid, TimeSpan and common numeric/date/boolean texts.</summary>
        /// <param name="value">单元格值。The cell value.</param>
        /// <param name="targetType">目标类型。The target type.</param>
        /// <param name="result">转换结果。The converted value.</param>
        /// <returns>转换成功返回 true。True when conversion succeeded.</returns>
        internal static bool TryConvertValue(object value, Type targetType, out object result)
        {
            result = null;
            Type underlying = Nullable.GetUnderlyingType(targetType) ?? targetType;
            if (value == null)
            {
                return !underlying.IsValueType;
            }
            if (underlying == typeof(object))
            {
                result = value;
                return true;
            }
            if (underlying.IsInstanceOfType(value))
            {
                result = value;
                return true;
            }
            string text = value as string;
            try
            {
                if (underlying == typeof(string))
                {
                    result = Convert.ToString(value, CultureInfo.InvariantCulture);
                    return true;
                }
                if (underlying == typeof(bool))
                {
                    if (value is bool)
                    {
                        result = value;
                        return true;
                    }
                    if (text != null)
                    {
                        string trimmed = text.Trim();
                        if (trimmed.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || trimmed == "1"
                            || trimmed.Equals("YES", StringComparison.OrdinalIgnoreCase))
                        {
                            result = true;
                            return true;
                        }
                        if (trimmed.Equals("FALSE", StringComparison.OrdinalIgnoreCase) || trimmed == "0"
                            || trimmed.Equals("NO", StringComparison.OrdinalIgnoreCase))
                        {
                            result = false;
                            return true;
                        }
                        return false;
                    }
                    result = Convert.ToDouble(value, CultureInfo.InvariantCulture) != 0;
                    return true;
                }
                if (underlying == typeof(DateTime))
                {
                    if (value is DateTime)
                    {
                        result = value;
                        return true;
                    }
                    if (value is double)
                    {
                        double serial = (double)value;
                        if (serial >= 0 && serial < 2958466)
                        {
                            result = DateTime.FromOADate(serial);
                            return true;
                        }
                        return false;
                    }
                    if (text != null)
                    {
                        DateTime date;
                        if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)
                            || DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date))
                        {
                            result = date;
                            return true;
                        }
                    }
                    return false;
                }
                if (underlying.IsEnum)
                {
                    if (text != null)
                    {
                        result = Enum.Parse(underlying, text.Trim(), true);
                    }
                    else
                    {
                        result = Enum.ToObject(underlying, Convert.ChangeType(value, Enum.GetUnderlyingType(underlying), CultureInfo.InvariantCulture));
                    }
                    return true;
                }
                if (underlying == typeof(Guid))
                {
                    Guid guid;
                    if (text != null && Guid.TryParse(text.Trim(), out guid))
                    {
                        result = guid;
                        return true;
                    }
                    return false;
                }
                if (underlying == typeof(TimeSpan))
                {
                    TimeSpan span;
                    if (value is TimeSpan)
                    {
                        result = value;
                        return true;
                    }
                    if (text != null && TimeSpan.TryParse(text.Trim(), CultureInfo.InvariantCulture, out span))
                    {
                        result = span;
                        return true;
                    }
                    return false;
                }
                if (underlying == typeof(char))
                {
                    if (value is char)
                    {
                        result = value;
                        return true;
                    }
                    if (text != null && text.Length > 0)
                    {
                        result = text[0];
                        return true;
                    }
                    return false;
                }
                result = Convert.ChangeType(value, underlying, CultureInfo.InvariantCulture);
                return true;
            }
            catch (FormatException)
            {
                return false;
            }
            catch (InvalidCastException)
            {
                return false;
            }
            catch (OverflowException)
            {
                return false;
            }
            catch (ArgumentException)
            {
                return false;
            }
        }
    }
}
