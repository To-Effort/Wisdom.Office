using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;

namespace Wisdom.Office.Core
{
    /// <summary>
    /// 自定义文档属性集合：保存名称与强类型值，名称不区分大小写，枚举时保持插入顺序。
    /// A collection of custom document properties: keeps names and strongly typed values, compares names
    /// case-insensitively, and preserves insertion order while enumerating.
    /// </summary>
    /// <remarks>
    /// 支持的值类型为 <see cref="string"/>、<see cref="bool"/>、<see cref="int"/>、<see cref="long"/>、
    /// <see cref="double"/>、<see cref="decimal"/> 与 <see cref="DateTime"/>，对应 OOXML 自定义属性
    /// (docProps/custom.xml) 的 lpwstr / bool / i4 / i8 / r8 / decimal / filetime 类型。
    /// Supported value types are string, bool, int, long, double, decimal and DateTime, matching the
    /// lpwstr / bool / i4 / i8 / r8 / decimal / filetime types of OOXML custom properties.
    /// </remarks>
    public sealed class CustomPropertyCollection : IEnumerable<KeyValuePair<string, object>>
    {
        /// <summary>不区分大小写的属性值字典。The case-insensitive property value dictionary.</summary>
        private readonly Dictionary<string, object> _values =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        /// <summary>按插入顺序保存的属性名。The property names kept in insertion order.</summary>
        private readonly List<string> _order = new List<string>();

        /// <summary>属性数量。The number of properties.</summary>
        public int Count
        {
            get { return _values.Count; }
        }

        /// <summary>按插入顺序返回全部属性名。Returns every property name in insertion order.</summary>
        public IList<string> Names
        {
            get { return _order.AsReadOnly(); }
        }

        /// <summary>
        /// 按键读取或写入属性；读取不存在的键返回 null，写入 null 等价于删除该属性。
        /// Reads or writes a property by name; reading a missing key returns null, and writing null removes
        /// the property.
        /// </summary>
        /// <param name="name">属性名，不能为空。The property name; must not be empty.</param>
        /// <returns>属性值，未找到时为 null。The property value, or null when missing.</returns>
        public object this[string name]
        {
            get
            {
                object value;
                return TryGetValue(name, out value) ? value : null;
            }
            set
            {
                if (value == null)
                {
                    Remove(name);
                    return;
                }
                Set(name, value);
            }
        }

        /// <summary>设置文本属性。Sets a text property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">文本值，null 按空字符串处理。The text value; null becomes an empty string.</param>
        public void Set(string name, string value)
        {
            SetValue(name, value ?? string.Empty);
        }

        /// <summary>设置布尔属性。Sets a Boolean property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">布尔值。The Boolean value.</param>
        public void Set(string name, bool value)
        {
            SetValue(name, value);
        }

        /// <summary>设置 32 位整数属性。Sets a 32-bit integer property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">整数值。The integer value.</param>
        public void Set(string name, int value)
        {
            SetValue(name, value);
        }

        /// <summary>设置 64 位整数属性。Sets a 64-bit integer property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">长整数值。The long value.</param>
        public void Set(string name, long value)
        {
            SetValue(name, value);
        }

        /// <summary>设置双精度浮点属性。Sets a double-precision property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">浮点值。The floating-point value.</param>
        public void Set(string name, double value)
        {
            SetValue(name, value);
        }

        /// <summary>设置十进制小数属性。Sets a decimal property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">十进制值。The decimal value.</param>
        public void Set(string name, decimal value)
        {
            SetValue(name, value);
        }

        /// <summary>设置日期时间属性（按 UTC 写出 filetime）。Sets a date/time property (written as a UTC filetime).</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">日期时间值。The date/time value.</param>
        public void Set(string name, DateTime value)
        {
            SetValue(name, value);
        }

        /// <summary>
        /// 设置任意受支持类型的属性；不支持的类型抛出 <see cref="ArgumentException"/>。
        /// Sets a property of any supported type; unsupported types throw <see cref="ArgumentException"/>.
        /// </summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">属性值，不能为 null。The property value; must not be null.</param>
        public void Set(string name, object value)
        {
            if (value == null) throw new ArgumentNullException("value");
            SetValue(name, value);
        }

        /// <summary>判断是否存在指定属性。Determines whether the given property exists.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <returns>存在时返回 true。True when present.</returns>
        public bool Contains(string name)
        {
            return !string.IsNullOrEmpty(name) && _values.ContainsKey(name);
        }

        /// <summary>尝试读取属性值。Tries to read a property value.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">成功时输出属性值，否则为 null。Outputs the value on success, otherwise null.</param>
        /// <returns>是否找到。Whether the property exists.</returns>
        public bool TryGetValue(string name, out object value)
        {
            if (string.IsNullOrEmpty(name))
            {
                value = null;
                return false;
            }
            return _values.TryGetValue(name, out value);
        }

        /// <summary>读取文本属性，缺失或为空时返回回退值。Reads a text property, returning the fallback when missing or null.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="defaultValue">回退值。The fallback value.</param>
        /// <returns>属性文本。The property text.</returns>
        public string GetString(string name, string defaultValue = "")
        {
            object value;
            if (!TryGetValue(name, out value) || value == null) return defaultValue;
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        /// <summary>读取 32 位整数属性。Reads a 32-bit integer property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="defaultValue">缺失或无法转换时的回退值。The fallback when missing or not convertible.</param>
        /// <returns>整数值。The integer value.</returns>
        public int GetInt32(string name, int defaultValue = 0)
        {
            object value;
            if (!TryGetValue(name, out value) || value == null) return defaultValue;
            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return defaultValue;
            }
        }

        /// <summary>读取 64 位整数属性。Reads a 64-bit integer property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="defaultValue">缺失或无法转换时的回退值。The fallback when missing or not convertible.</param>
        /// <returns>长整数值。The long value.</returns>
        public long GetInt64(string name, long defaultValue = 0)
        {
            object value;
            if (!TryGetValue(name, out value) || value == null) return defaultValue;
            try
            {
                return Convert.ToInt64(value, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return defaultValue;
            }
        }

        /// <summary>读取双精度属性。Reads a double property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="defaultValue">缺失或无法转换时的回退值。The fallback when missing or not convertible.</param>
        /// <returns>浮点值。The floating-point value.</returns>
        public double GetDouble(string name, double defaultValue = 0)
        {
            object value;
            if (!TryGetValue(name, out value) || value == null) return defaultValue;
            try
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return defaultValue;
            }
        }

        /// <summary>读取布尔属性。Reads a Boolean property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="defaultValue">缺失或无法转换时的回退值。The fallback when missing or not convertible.</param>
        /// <returns>布尔值。The Boolean value.</returns>
        public bool GetBoolean(string name, bool defaultValue = false)
        {
            object value;
            if (!TryGetValue(name, out value) || value == null) return defaultValue;
            try
            {
                return Convert.ToBoolean(value, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return defaultValue;
            }
        }

        /// <summary>读取日期时间属性。Reads a date/time property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="defaultValue">缺失或无法转换时的回退值。The fallback when missing or not convertible.</param>
        /// <returns>日期时间值。The date/time value.</returns>
        public DateTime GetDateTime(string name, DateTime defaultValue = default(DateTime))
        {
            object value;
            if (!TryGetValue(name, out value) || value == null) return defaultValue;
            try
            {
                return Convert.ToDateTime(value, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return defaultValue;
            }
        }

        /// <summary>删除指定属性。Removes the given property.</summary>
        /// <param name="name">属性名。The property name.</param>
        /// <returns>确实删除时返回 true。True when the property was removed.</returns>
        public bool Remove(string name)
        {
            if (string.IsNullOrEmpty(name) || !_values.Remove(name)) return false;
            for (int i = _order.Count - 1; i >= 0; i--)
            {
                if (string.Equals(_order[i], name, StringComparison.OrdinalIgnoreCase))
                {
                    _order.RemoveAt(i);
                }
            }
            return true;
        }

        /// <summary>清空全部自定义属性。Removes every custom property.</summary>
        public void Clear()
        {
            _values.Clear();
            _order.Clear();
        }

        /// <summary>返回一个独立的深拷贝。Returns an independent deep copy.</summary>
        /// <returns>属性相同的新集合。A new collection with the same properties.</returns>
        public CustomPropertyCollection Clone()
        {
            CustomPropertyCollection copy = new CustomPropertyCollection();
            for (int i = 0; i < _order.Count; i++)
            {
                copy._order.Add(_order[i]);
                copy._values[_order[i]] = _values[_order[i]];
            }
            return copy;
        }

        /// <summary>按插入顺序枚举属性。Enumerates properties in insertion order.</summary>
        /// <returns>属性枚举器。The property enumerator.</returns>
        public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
        {
            for (int i = 0; i < _order.Count; i++)
            {
                string name = _order[i];
                yield return new KeyValuePair<string, object>(name, _values[name]);
            }
        }

        /// <summary>按插入顺序枚举属性（非泛型入口）。Enumerates properties in insertion order (non-generic entry point).</summary>
        /// <returns>属性枚举器。The property enumerator.</returns>
        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// 校验名称与值类型后写入字典；新名称追加到顺序表末尾，已有名称保留原位置。
        /// Validates the name and value type, then writes the dictionary; new names are appended to the
        /// order list while existing names keep their position.
        /// </summary>
        /// <param name="name">属性名。The property name.</param>
        /// <param name="value">属性值。The property value.</param>
        private void SetValue(string name, object value)
        {
            if (string.IsNullOrEmpty(name)) throw new ArgumentException("Custom property name must not be empty.", "name");
            if (name.Length > 255) throw new ArgumentException("Custom property name must not exceed 255 characters.", "name");
            ValidateXmlText(name, "name");
            if (!IsSupportedValue(value))
            {
                throw new ArgumentException(
                    "Unsupported custom property type: " + (value == null ? "null" : value.GetType().FullName),
                    "value");
            }
            string text = value as string;
            if (text != null) ValidateXmlText(text, "value");
            if (value is double && (double.IsNaN((double)value) || double.IsInfinity((double)value)))
            {
                throw new ArgumentException("Custom property double values must be finite.", "value");
            }
            if (!_values.ContainsKey(name)) _order.Add(name);
            _values[name] = value;
        }

        /// <summary>
        /// 校验文本可安全写入 XML 1.0：拒绝未配对的代理项、C0 控制字符与 U+FFFE/U+FFFF。
        /// Validates that text can be written safely into XML 1.0: rejects unpaired surrogates, C0 control
        /// characters and U+FFFE/U+FFFF.
        /// </summary>
        /// <param name="text">待校验文本。The text to validate.</param>
        /// <param name="parameterName">异常参数名。The parameter name used by the exception.</param>
        /// <exception cref="ArgumentException">文本含 XML 不允许的字符。Thrown when the text contains characters XML forbids.</exception>
        private static void ValidateXmlText(string text, string parameterName)
        {
            for (int i = 0; i < text.Length; i++)
            {
                char character = text[i];
                if (char.IsHighSurrogate(character))
                {
                    if (i + 1 >= text.Length || !char.IsLowSurrogate(text[i + 1]))
                    {
                        throw new ArgumentException("Custom property text contains an unpaired surrogate.", parameterName);
                    }
                    i++;
                    continue;
                }
                if (char.IsLowSurrogate(character))
                {
                    throw new ArgumentException("Custom property text contains an unpaired surrogate.", parameterName);
                }
                if ((character < 0x20 && character != '\t' && character != '\n' && character != '\r')
                    || character == '\uFFFE'
                    || character == '\uFFFF')
                {
                    throw new ArgumentException("Custom property text contains a character that XML does not allow.", parameterName);
                }
            }
        }

        /// <summary>判断值类型是否可写入 OOXML 自定义属性。Checks whether the value type can be written as an OOXML custom property.</summary>
        /// <param name="value">待检查的值。The value to inspect.</param>
        /// <returns>受支持时返回 true。True when supported.</returns>
        private static bool IsSupportedValue(object value)
        {
            return value is string
                || value is bool
                || value is int
                || value is long
                || value is double
                || value is decimal
                || value is DateTime;
        }
    }
}
