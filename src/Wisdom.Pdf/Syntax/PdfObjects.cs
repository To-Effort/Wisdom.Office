using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Wisdom.Pdf
{
    /// <summary>
    /// PDF 文档对象模型的基类，对应 PDF 规范中的直接对象与间接对象引用。
    /// Base class of the PDF document object model, covering the direct objects and indirect references of
    /// the PDF specification.
    /// </summary>
    public abstract class PdfObject
    {
    }

    /// <summary>
    /// PDF 空对象（null），表示一个不存在的值。
    /// The PDF null object, representing an absent value.
    /// </summary>
    public sealed class PdfNull : PdfObject
    {
        /// <summary>全局唯一的 null 实例，避免重复分配。The single global null instance, avoiding repeated allocation.</summary>
        public static readonly PdfNull Instance = new PdfNull();

        /// <summary>仅允许通过 <see cref="Instance"/> 获取实例。Instances are only obtainable through <see cref="Instance"/>.</summary>
        private PdfNull() { }
    }

    /// <summary>
    /// PDF 布尔对象（true / false）。
    /// The PDF boolean object (true / false).
    /// </summary>
    public sealed class PdfBoolean : PdfObject
    {
        /// <summary>共享的 true 实例。The shared true instance.</summary>
        public static readonly PdfBoolean True = new PdfBoolean(true);
        /// <summary>共享的 false 实例。The shared false instance.</summary>
        public static readonly PdfBoolean False = new PdfBoolean(false);

        /// <summary>布尔值本身。The boolean value itself.</summary>
        public bool Value { get; private set; }

        /// <summary>
        /// 创建一个指定布尔值的 PDF 布尔对象。
        /// Creates a PDF boolean object with the given value.
        /// </summary>
        /// <param name="value">布尔值。The boolean value.</param>
        public PdfBoolean(bool value)
        {
            Value = value;
        }
    }

    /// <summary>
    /// PDF 数值对象，整数与实数统一以 double 承载。
    /// The PDF number object; integers and reals are both carried as double.
    /// </summary>
    public sealed class PdfNumber : PdfObject
    {
        /// <summary>数值本身。The numeric value itself.</summary>
        public double Value { get; private set; }

        /// <summary>
        /// 创建一个指定数值的 PDF 数值对象。
        /// Creates a PDF number object with the given value.
        /// </summary>
        /// <param name="value">数值。The numeric value.</param>
        public PdfNumber(double value)
        {
            Value = value;
        }

        /// <summary>共享的零值实例，常用于数组与矩阵的默认填充。The shared zero instance, commonly used to default-fill arrays and matrices.</summary>
        public static PdfNumber Zero = new PdfNumber(0);
    }

    /// <summary>
    /// PDF 名称对象（以 / 开头的标识符，如 /Type、/Page）。
    /// The PDF name object: an identifier starting with /, such as /Type or /Page.
    /// </summary>
    public sealed class PdfName : PdfObject
    {
        /// <summary>不含前导斜杠的名称正文。The name body without the leading slash.</summary>
        public string Value { get; private set; }

        /// <summary>
        /// 创建一个 PDF 名称对象。
        /// Creates a PDF name object.
        /// </summary>
        /// <param name="value">名称正文，不能为空。The name body; must not be empty.</param>
        /// <exception cref="ArgumentException">当 <paramref name="value"/> 为 null 或空字符串时抛出。Thrown when <paramref name="value"/> is null or an empty string.</exception>
        public PdfName(string value)
        {
            if (string.IsNullOrEmpty(value)) throw new ArgumentException("PDF name cannot be empty.", "value");
            Value = value;
        }

        /// <summary>
        /// 返回带前导斜杠的名称文本表示。
        /// Returns the textual representation of the name with its leading slash.
        /// </summary>
        /// <returns>形如 /Type 的字符串。A string such as /Type.</returns>
        public override string ToString()
        {
            return "/" + Value;
        }
    }

    /// <summary>
    /// PDF 字符串对象，内部以原始字节保存，可为字面串或十六进制串。
    /// The PDF string object, stored as raw bytes; either a literal or a hexadecimal string.
    /// </summary>
    public sealed class PdfString : PdfObject
    {
        /// <summary>字符串的原始字节内容（可能含编码前缀，如 UTF-16LE BOM）。The raw byte content (may include an encoding prefix such as a UTF-16LE BOM).</summary>
        public byte[] Bytes { get; private set; }
        /// <summary>是否以十六进制形式写出；false 表示字面串（括号形式）。Whether the string is written in hex form; false means the literal (parenthesized) form.</summary>
        public bool IsHex { get; private set; }

        /// <summary>
        /// 用给定字节与写出形式创建 PDF 字符串对象。
        /// Creates a PDF string object from the given bytes and writing form.
        /// </summary>
        /// <param name="bytes">字符串内容字节，null 视为空数组。The string content bytes; null is treated as an empty array.</param>
        /// <param name="hex">true 表示以十六进制形式写出。True writes the string in hexadecimal form.</param>
        public PdfString(byte[] bytes, bool hex)
        {
            Bytes = bytes ?? new byte[0];
            IsHex = hex;
        }

        /// <summary>
        /// 替换字符串的原始字节，供解密与重写时原地更新对象。
        /// Replaces the raw bytes of the string, used by decryption and rewriting to update objects in place.
        /// </summary>
        /// <param name="bytes">新的字节内容。The new byte content.</param>
        internal void ReplaceBytes(byte[] bytes)
        {
            Bytes = bytes ?? new byte[0];
        }

        /// <summary>
        /// 由文本创建字面串（括号形式）对象。
        /// Creates a literal (parenthesized) string object from text.
        /// </summary>
        /// <param name="text">文本内容。The text content.</param>
        /// <returns>非十六进制的 PDF 字符串。A non-hex PDF string.</returns>
        public static PdfString Literal(string text)
        {
            return new PdfString(EncodePdfDocString(text), false);
        }

        /// <summary>
        /// 由文本创建带 UTF-16LE BOM 的十六进制字符串，用于保证中文等多语言元数据可被阅读器正确显示。
        /// Creates a hexadecimal string with a UTF-16LE BOM from text, so multilingual metadata such as
        /// Chinese is displayed correctly by PDF readers.
        /// </summary>
        /// <param name="text">文本内容，null 视为空字符串。The text content; null is treated as an empty string.</param>
        /// <returns>UTF-16LE 编码的十六进制 PDF 字符串。A UTF-16LE encoded hexadecimal PDF string.</returns>
        public static PdfString Unicode(string text)
        {
            if (text == null) text = string.Empty;
            byte[] bytes = new byte[text.Length * 2 + 2];
            bytes[0] = 0xFE;
            bytes[1] = 0xFF;
            for (int i = 0; i < text.Length; i++)
            {
                bytes[2 + i * 2] = (byte)(text[i] >> 8);
                bytes[3 + i * 2] = (byte)text[i];
            }
            return new PdfString(bytes, true);
        }

        /// <summary>
        /// 由任意字节创建十六进制字符串对象，用于写入已编码好的二进制内容。
        /// Creates a hexadecimal string object from arbitrary bytes, used for pre-encoded binary content.
        /// </summary>
        /// <param name="bytes">字节内容。The byte content.</param>
        /// <returns>十六进制的 PDF 字符串。A hexadecimal PDF string.</returns>
        public static PdfString Hex(byte[] bytes)
        {
            return new PdfString(bytes, true);
        }

        /// <summary>
        /// 将文本编码为 PDF 文档字符串字节。
        /// Encodes text into PDF doc-string bytes.
        /// </summary>
        /// <param name="text">待编码文本，null 或空时返回空数组。The text to encode; null or empty input yields an empty array.</param>
        /// <returns>每个字符占一个字节的编码结果。The encoded result with one byte per character.</returns>
        private static byte[] EncodePdfDocString(string text)
        {
            if (string.IsNullOrEmpty(text)) return new byte[0];
            // 采用 ASCII 加基础 Latin-1 回退，保证元数据在所有阅读器中都能读出；超出范围的字符以 '?' 代替。
            byte[] result = new byte[text.Length];
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                result[i] = c <= 0xFF ? (byte)c : (byte)'?';
            }
            return result;
        }
    }

    /// <summary>
    /// PDF 间接对象引用（形如 "n g R"），指向交叉表中的某个对象号。
    /// A PDF indirect object reference (of the form "n g R") pointing at an object number in the
    /// cross-reference table.
    /// </summary>
    public sealed class PdfReference : PdfObject
    {
        /// <summary>对象号。The object number.</summary>
        public int Number { get; private set; }
        /// <summary>代号（generation），常规文档为 0。The generation number, 0 for regular documents.</summary>
        public int Generation { get; private set; }

        /// <summary>
        /// 创建一个指向指定对象号与代号的引用。
        /// Creates a reference to the given object number and generation.
        /// </summary>
        /// <param name="number">对象号。The object number.</param>
        /// <param name="generation">代号。The generation number.</param>
        public PdfReference(int number, int generation)
        {
            Number = number;
            Generation = generation;
        }

        /// <summary>
        /// 返回 PDF 语法形式的引用文本。
        /// Returns the reference in PDF syntax form.
        /// </summary>
        /// <returns>形如 "12 0 R" 的字符串。A string such as "12 0 R".</returns>
        public override string ToString()
        {
            return Number.ToString(CultureInfo.InvariantCulture) + " " + Generation.ToString(CultureInfo.InvariantCulture) + " R";
        }
    }

    /// <summary>
    /// PDF 数组对象，元素顺序即写出顺序，允许混合任意对象类型。
    /// The PDF array object; element order equals write order and mixed object types are allowed.
    /// </summary>
    public sealed class PdfArray : PdfObject, IEnumerable<PdfObject>
    {
        /// <summary>按写出顺序存放的元素。The elements in write order.</summary>
        private readonly List<PdfObject> _items = new List<PdfObject>();

        /// <summary>元素个数。The number of elements.</summary>
        public int Count { get { return _items.Count; } }
        /// <summary>按下标访问元素。Accesses an element by index.</summary>
        /// <param name="index">从零开始的下标。The zero-based index.</param>
        /// <returns>该位置的 PDF 对象。The PDF object at that position.</returns>
        public PdfObject this[int index] { get { return _items[index]; } }

        /// <summary>
        /// 追加一个元素；传入 null 时以 PDF null 占位，避免写出非法数组。
        /// Appends an element; a null argument is stored as the PDF null so no invalid array is written.
        /// </summary>
        /// <param name="value">待追加的对象。The object to append.</param>
        /// <returns>当前数组自身，便于链式调用。This array itself, for chaining.</returns>
        public PdfArray Add(PdfObject value)
        {
            _items.Add(value ?? PdfNull.Instance);
            return this;
        }

        /// <summary>
        /// 追加一个数值元素。
        /// Appends a numeric element.
        /// </summary>
        /// <param name="value">数值。The numeric value.</param>
        /// <returns>当前数组自身，便于链式调用。This array itself, for chaining.</returns>
        public PdfArray AddNumber(double value)
        {
            return Add(new PdfNumber(value));
        }

        /// <summary>
        /// 追加一个名称元素。
        /// Appends a name element.
        /// </summary>
        /// <param name="value">名称正文（不含前导斜杠）。The name body (without the leading slash).</param>
        /// <returns>当前数组自身，便于链式调用。This array itself, for chaining.</returns>
        public PdfArray AddName(string value)
        {
            return Add(new PdfName(value));
        }

        /// <summary>
        /// 移除指定下标的元素，供页面树与注释数组的精确编辑使用。
        /// Removes the element at the given index, used for precise edits of page trees and annotation arrays.
        /// </summary>
        /// <param name="index">从零开始的下标。The zero-based index.</param>
        public void RemoveAt(int index)
        {
            _items.RemoveAt(index);
        }

        /// <summary>
        /// 顺序遍历数组元素。
        /// Iterates the array elements in order.
        /// </summary>
        /// <returns>元素的枚举器。The element enumerator.</returns>
        public IEnumerator<PdfObject> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        /// <summary>
        /// 非泛型枚举入口，供 <see cref="IEnumerable"/> 调用方使用。
        /// The non-generic enumeration entry point for <see cref="IEnumerable"/> callers.
        /// </summary>
        /// <returns>元素的枚举器。The element enumerator.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }
    }

    /// <summary>
    /// PDF 字典对象，键到对象的映射。
    /// The PDF dictionary object, a mapping from keys to objects.
    /// </summary>
    /// <remarks>内部以列表按插入顺序保存键值对，因此枚举与写出顺序稳定，检索为线性查找；字典按键区分大小写。
    /// Key/value pairs are kept in a list in insertion order, so enumeration and writing are stable while
    /// lookup is linear; keys are matched case-sensitively.</remarks>
    public sealed class PdfDictionary : PdfObject, IEnumerable<KeyValuePair<string, PdfObject>>
    {
        /// <summary>按插入顺序存放的键值对。The key/value pairs in insertion order.</summary>
        private readonly List<KeyValuePair<string, PdfObject>> _items = new List<KeyValuePair<string, PdfObject>>();

        /// <summary>键值对个数。The number of key/value pairs.</summary>
        public int Count { get { return _items.Count; } }

        /// <summary>
        /// 按键读取值对象。
        /// Reads the value object for a key.
        /// </summary>
        /// <param name="key">字典键（不含前导斜杠），区分大小写。The dictionary key (without the leading slash), case-sensitive.</param>
        /// <returns>对应的对象；键不存在时返回 null。The matching object, or null when the key is absent.</returns>
        public PdfObject Get(string key)
        {
            for (int i = 0; i < _items.Count; i++)
            {
                if (string.Equals(_items[i].Key, key, StringComparison.Ordinal))
                {
                    return _items[i].Value;
                }
            }
            return null;
        }

        /// <summary>
        /// 判断字典是否包含指定键。
        /// Determines whether the dictionary contains the given key.
        /// </summary>
        /// <param name="key">字典键。The dictionary key.</param>
        /// <returns>存在返回 true。True when the key exists.</returns>
        public bool ContainsKey(string key)
        {
            return Get(key) != null;
        }

        /// <summary>
        /// 设置键值：键已存在则原地覆盖并保持原有位置，否则追加到末尾。
        /// Sets a key/value pair: an existing key is overwritten in place keeping its position, otherwise the
        /// pair is appended at the end.
        /// </summary>
        /// <param name="key">字典键。The dictionary key.</param>
        /// <param name="value">值对象，null 时写入 PDF null。The value object; null writes the PDF null.</param>
        /// <returns>当前字典自身，便于链式调用。This dictionary itself, for chaining.</returns>
        public PdfDictionary Set(string key, PdfObject value)
        {
            for (int i = 0; i < _items.Count; i++)
            {
                if (string.Equals(_items[i].Key, key, StringComparison.Ordinal))
                {
                    _items[i] = new KeyValuePair<string, PdfObject>(key, value ?? PdfNull.Instance);
                    return this;
                }
            }
            _items.Add(new KeyValuePair<string, PdfObject>(key, value ?? PdfNull.Instance));
            return this;
        }

        /// <summary>
        /// 设置一个数值类型的键值。
        /// Sets a numeric key/value pair.
        /// </summary>
        /// <param name="key">字典键。The dictionary key.</param>
        /// <param name="value">数值。The numeric value.</param>
        /// <returns>当前字典自身，便于链式调用。This dictionary itself, for chaining.</returns>
        public PdfDictionary SetNumber(string key, double value)
        {
            return Set(key, new PdfNumber(value));
        }

        /// <summary>
        /// 设置一个名称类型的键值。
        /// Sets a name-type key/value pair.
        /// </summary>
        /// <param name="key">字典键。The dictionary key.</param>
        /// <param name="value">名称正文（不含前导斜杠）。The name body (without the leading slash).</param>
        /// <returns>当前字典自身，便于链式调用。This dictionary itself, for chaining.</returns>
        public PdfDictionary SetName(string key, string value)
        {
            return Set(key, new PdfName(value));
        }

        /// <summary>
        /// 设置一个 UTF-16LE 文本键值，适用于标题、作者等可能含中文的元数据。
        /// Sets a UTF-16LE text key/value pair, suited to metadata such as titles and authors that may contain
        /// Chinese characters.
        /// </summary>
        /// <param name="key">字典键。The dictionary key.</param>
        /// <param name="value">文本内容。The text content.</param>
        /// <returns>当前字典自身，便于链式调用。This dictionary itself, for chaining.</returns>
        public PdfDictionary SetString(string key, string value)
        {
            return Set(key, PdfString.Unicode(value));
        }

        /// <summary>
        /// 设置一个布尔类型的键值。
        /// Sets a boolean key/value pair.
        /// </summary>
        /// <param name="key">字典键。The dictionary key.</param>
        /// <param name="value">布尔值。The boolean value.</param>
        /// <returns>当前字典自身，便于链式调用。This dictionary itself, for chaining.</returns>
        public PdfDictionary SetBoolean(string key, bool value)
        {
            return Set(key, value ? PdfBoolean.True : PdfBoolean.False);
        }

        /// <summary>
        /// 移除指定键及其值。
        /// Removes the given key and its value.
        /// </summary>
        /// <param name="key">字典键。The dictionary key.</param>
        /// <returns>确实删除了条目返回 true，键不存在返回 false。True when an entry was actually removed, false when the key was absent.</returns>
        public bool Remove(string key)
        {
            for (int i = 0; i < _items.Count; i++)
            {
                if (string.Equals(_items[i].Key, key, StringComparison.Ordinal))
                {
                    _items.RemoveAt(i);
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 浅拷贝字典：复制全部键值对的引用，值对象本身不递归复制。
        /// Shallow-copies the dictionary: all key/value references are copied, value objects are not cloned
        /// recursively.
        /// </summary>
        /// <returns>新的字典实例。The new dictionary instance.</returns>
        public PdfDictionary Clone()
        {
            PdfDictionary copy = new PdfDictionary();
            for (int i = 0; i < _items.Count; i++)
            {
                copy.Set(_items[i].Key, _items[i].Value);
            }
            return copy;
        }

        /// <summary>
        /// 按插入顺序遍历字典的键值对。
        /// Iterates the dictionary's key/value pairs in insertion order.
        /// </summary>
        /// <returns>键值对的枚举器。The key/value enumerator.</returns>
        public IEnumerator<KeyValuePair<string, PdfObject>> GetEnumerator()
        {
            return _items.GetEnumerator();
        }

        /// <summary>
        /// 非泛型枚举入口，供 <see cref="IEnumerable"/> 调用方使用。
        /// The non-generic enumeration entry point for <see cref="IEnumerable"/> callers.
        /// </summary>
        /// <returns>键值对的枚举器。The key/value enumerator.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return _items.GetEnumerator();
        }
    }

    /// <summary>
    /// PDF 流对象：字典描述的元数据加上与之关联的二进制内容。
    /// The PDF stream object: dictionary-described metadata plus the associated binary content.
    /// </summary>
    public sealed class PdfStream : PdfObject
    {
        /// <summary>流的字典，存放 /Length、/Filter 等键。The stream dictionary holding keys such as /Length and /Filter.</summary>
        public PdfDictionary Dictionary { get; private set; }
        /// <summary>流的原始字节（写出时可能是已压缩的，取决于 <see cref="Compress"/>）。The raw stream bytes (possibly compressed at write time, depending on <see cref="Compress"/>).</summary>
        public byte[] Data { get; private set; }
        /// <summary>是否在写出时对该流做 Flate 压缩。Whether the stream is Flate-compressed when written.</summary>
        public bool Compress { get; set; }

        /// <summary>
        /// 创建一个尚无字典的流对象。
        /// Creates a stream object without a preset dictionary.
        /// </summary>
        /// <param name="data">流内容，null 视为空数组。The stream content; null is treated as an empty array.</param>
        /// <param name="compress">是否压缩写出。Whether to compress on write.</param>
        public PdfStream(byte[] data, bool compress)
        {
            Dictionary = new PdfDictionary();
            Data = data ?? new byte[0];
            Compress = compress;
        }

        /// <summary>
        /// 创建一个带指定字典的流对象。
        /// Creates a stream object with the given dictionary.
        /// </summary>
        /// <param name="dictionary">流字典，null 时新建空字典。The stream dictionary; a new empty dictionary is created when null.</param>
        /// <param name="data">流内容，null 视为空数组。The stream content; null is treated as an empty array.</param>
        /// <param name="compress">是否压缩写出。Whether to compress on write.</param>
        public PdfStream(PdfDictionary dictionary, byte[] data, bool compress)
        {
            Dictionary = dictionary ?? new PdfDictionary();
            Data = data ?? new byte[0];
            Compress = compress;
        }

        /// <summary>
        /// 替换流的原始字节，供解密与重写时原地更新对象。
        /// Replaces the raw stream bytes, used by decryption and rewriting to update objects in place.
        /// </summary>
        /// <param name="data">新的流内容。The new stream content.</param>
        internal void ReplaceData(byte[] data)
        {
            Data = data ?? new byte[0];
        }
    }

    /// <summary>
    /// PDF 数值的文本格式化助手，负责把 double 写成合法且紧凑的 PDF 数字。
    /// Text formatting helper for PDF numbers, writing doubles as valid and compact PDF numbers.
    /// </summary>
    internal static class PdfNumberFormat
    {
        /// <summary>
        /// 按 PDF 语法格式化数值：整数值省略小数点，小数最多保留 6 位，NaN 与无穷大退化为 0。
        /// Formats a number per PDF syntax: integral values drop the decimal point, fractions keep at most 6
        /// digits, and NaN or infinity degrade to 0.
        /// </summary>
        /// <param name="value">待格式化的数值。The value to format.</param>
        /// <returns>使用不变文化的数字字符串。The culture-invariant numeric string.</returns>
        public static string Format(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                return "0";
            }
            if (Math.Abs(value - Math.Round(value)) < 0.0000001)
            {
                return ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
            }
            return value.ToString("0.######", CultureInfo.InvariantCulture);
        }
    }

    /// <summary>
    /// PDF 名称的转义助手，保证名称可被任何阅读器解析。
    /// Escaping helper for PDF names, ensuring any reader can parse them.
    /// </summary>
    internal static class PdfNameEscape
    {
        /// <summary>
        /// 将名称中的空白、控制字符、DEL 以及 PDF 定界符替换为 #HH 形式（非 ASCII 字符先按 UTF-8 展开）。
        /// Replaces whitespace, control characters, DEL and PDF delimiters in a name with #HH escapes
        /// (non-ASCII characters are first expanded as UTF-8).
        /// </summary>
        /// <param name="name">待转义的名称正文。The name body to escape.</param>
        /// <returns>转义后的名称字符串。The escaped name string.</returns>
        public static string Escape(string name)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < name.Length; i++)
            {
                char c = name[i];
                if (c <= 32 || c >= 127 || c == '#' || c == '/' || c == '(' || c == ')' || c == '<' || c == '>' || c == '['
                    || c == ']' || c == '{' || c == '}' || c == '%')
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(c.ToString());
                    for (int b = 0; b < bytes.Length; b++)
                    {
                        builder.Append('#');
                        builder.Append(bytes[b].ToString("x2", CultureInfo.InvariantCulture));
                    }
                }
                else
                {
                    builder.Append(c);
                }
            }
            return builder.ToString();
        }
    }
}
