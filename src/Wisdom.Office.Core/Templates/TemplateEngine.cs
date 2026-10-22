using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Wisdom.Office.Core
{
    /// <summary>忽略大小写、支持点号路径解析的模板数据上下文。A case-insensitive template data context with dotted-path resolution.</summary>
    public sealed class TemplateContext
    {
        /// <summary>以忽略大小写比较器存储的键值字典。The key-value dictionary stored with a case-insensitive comparer.</summary>
        private readonly Dictionary<string, object> _values =
            new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

        /// <summary>按键读写值的索引器；读取走 <see cref="Get"/> 的点号路径解析，写入直接落到顶层字典。Indexer to read or write a value by key; reads use the dotted-path resolution of <see cref="Get"/>, while writes land in the top-level dictionary.</summary>
        /// <param name="key">键或点号路径。The key or dotted path.</param>
        /// <returns>解析出的值，未找到时为 null。The resolved value, or null when not found.</returns>
        public object this[string key]
        {
            get { return Get(key); }
            set { _values[key] = value; }
        }

        /// <summary>在顶层字典设置一个键值。Sets a key-value pair in the top-level dictionary.</summary>
        /// <param name="key">键名。The key.</param>
        /// <param name="value">任意值。The value.</param>
        public void Set(string key, object value)
        {
            _values[key] = value;
        }

        /// <summary>判断顶层字典是否直接包含指定键。Reports whether the top-level dictionary directly contains the given key.</summary>
        /// <param name="key">键名。The key.</param>
        /// <returns>是否包含。Whether the key is present.</returns>
        public bool Contains(string key)
        {
            return _values.ContainsKey(key);
        }

        /// <summary>顶层全部键值对的枚举。An enumeration of all top-level key-value pairs.</summary>
        public IEnumerable<KeyValuePair<string, object>> Items
        {
            get { return _values; }
        }

        /// <summary>按点号路径解析并返回值，未找到时返回 null。Resolves a dotted path and returns its value, or null when not found.</summary>
        /// <param name="path">点号分隔的路径。The dot-separated path.</param>
        /// <returns>解析出的值或 null。The resolved value or null.</returns>
        public object Get(string path)
        {
            object value;
            if (TryGet(path, out value))
            {
                return value;
            }
            return null;
        }

        /// <summary>尝试解析点号路径：首段取自顶层字典，其后各段逐级通过成员反射/字典查找。Attempts to resolve a dotted path. The first segment is read from the top-level dictionary; each following segment is resolved against the current object via <see cref="TryGetMember"/>. Returns false as soon as any segment cannot be resolved.</summary>
        /// <param name="path">点号分隔的路径；空或 null 直接失败。The dot-separated path; empty or null fails immediately.</param>
        /// <param name="value">解析成功时输出最终值，否则为 null。Outputs the resolved value on success, otherwise null.</param>
        /// <returns>是否成功解析到值。Whether a value was successfully resolved.</returns>
        public bool TryGet(string path, out object value)
        {
            value = null;
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            string[] parts = path.Split('.');
            object current = null;
            bool found = false;
            for (int i = 0; i < parts.Length; i++)
            {
                if (i == 0)
                {
                    if (!_values.TryGetValue(parts[0], out current))
                    {
                        return false;
                    }
                    found = true;
                    continue;
                }
                if (!TryGetMember(current, parts[i], out current))
                {
                    return false;
                }
            }
            value = current;
            return found;
        }

        /// <summary>从单个对象解析一个成员：依次尝试嵌套模板上下文、字符串键字典、非泛型字典、公共实例属性与公共实例字段。Resolves one member against a source object, trying in order: a nested <see cref="TemplateContext"/>, an <c>IDictionary&lt;string, object&gt;</c>, a non-generic <c>IDictionary</c> (case-insensitive key match), a public instance property, then a public instance field. Indexer properties are skipped.</summary>
        /// <param name="source">被读取的对象，null 直接失败。The object to read from; null fails immediately.</param>
        /// <param name="member">成员名或键。The member name or key.</param>
        /// <param name="value">解析成功时输出成员值，否则为 null。Outputs the member value on success, otherwise null.</param>
        /// <returns>是否找到该成员。Whether the member was found.</returns>
        internal static bool TryGetMember(object source, string member, out object value)
        {
            value = null;
            if (source == null)
            {
                return false;
            }

            TemplateContext nestedContext = source as TemplateContext;
            if (nestedContext != null)
            {
                return nestedContext.TryGet(member, out value);
            }

            IDictionary<string, object> stringDictionary = source as IDictionary<string, object>;
            if (stringDictionary != null)
            {
                return stringDictionary.TryGetValue(member, out value);
            }

            IDictionary dictionary = source as IDictionary;
            if (dictionary != null)
            {
                foreach (object key in dictionary.Keys)
                {
                    if (string.Equals(Convert.ToString(key, CultureInfo.InvariantCulture), member, StringComparison.OrdinalIgnoreCase))
                    {
                        value = dictionary[key];
                        return true;
                    }
                }
            }

            Type type = source.GetType();
            PropertyInfo property = type.GetProperty(member, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (property != null && property.GetIndexParameters().Length == 0)
            {
                value = property.GetValue(source, null);
                return true;
            }

            FieldInfo field = type.GetField(member, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (field != null)
            {
                value = field.GetValue(source);
                return true;
            }

            return false;
        }
    }

    /// <summary>模板渲染选项：未解析标记的处理、HTML 编码与区域文化。Rendering options for templates: unresolved-token handling, HTML encoding and culture.</summary>
    public sealed class TemplateOptions
    {
        /// <summary>无法解析的标记是否原样保留，true 保留、false 移除。Whether unresolved tokens are left in place; true keeps them, false removes them.</summary>
        public bool LeaveUnresolvedTokens { get; set; }
        /// <summary>替换值前是否对其进行 HTML 转义。Whether substituted values are HTML-escaped before insertion.</summary>
        public bool HtmlEncodeValues { get; set; }
        /// <summary>格式化取值所用的区域文化，默认不变文化。The culture used when formatting values; the invariant culture by default.</summary>
        public CultureInfo Culture { get; set; }

        /// <summary>默认构造：保留未解析标记、不 HTML 编码、使用不变文化。Default constructor: leaves unresolved tokens, skips HTML encoding, and uses the invariant culture.</summary>
        public TemplateOptions()
        {
            LeaveUnresolvedTokens = true;
            Culture = CultureInfo.InvariantCulture;
        }
    }

    /// <summary>基于双花括号标记与 #each 循环的轻量模板引擎。A lightweight template engine based on double-brace tokens and #each loops.</summary>
    public static class TemplateEngine
    {
        /// <summary>匹配 <c>{{ key }}</c> 或 <c>{{ key : format }}</c> 形式的值标记。Matches value tokens of the form <c>{{ key }}</c> or <c>{{ key : format }}</c>.</summary>
        private static readonly Regex TokenRegex = new Regex(
            @"\{\{\s*([^#/{}][^{}:]*?)\s*(?::\s*([^{}]+?))?\s*\}\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        /// <summary>匹配 <c>{{#each name}}</c> 循环起始标记。Matches the <c>{{#each name}}</c> loop-start token.</summary>
        private static readonly Regex LoopStartRegex = new Regex(
            @"\{\{\s*#each\s+([^{}]+?)\s*\}\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// <summary>匹配 <c>{{/each}}</c> 循环结束标记。Matches the <c>{{/each}}</c> loop-end token.</summary>
        private static readonly Regex LoopEndRegex = new Regex(
            @"\{\{\s*/each\s*\}\}",
            RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

        /// <summary>判断文本中是否含至少一个值标记。Reports whether the text contains at least one value token.</summary>
        /// <param name="text">待检测文本。The text to inspect.</param>
        /// <returns>是否含标记。Whether a token is present.</returns>
        public static bool ContainsToken(string text)
        {
            return !string.IsNullOrEmpty(text) && TokenRegex.IsMatch(text);
        }

        /// <summary>以默认选项将文本中的标记替换为上下文取值。Substitutes tokens in the text with context values using default options.</summary>
        /// <param name="text">模板文本。The template text.</param>
        /// <param name="context">数据上下文。The data context.</param>
        /// <returns>替换后的文本。The substituted text.</returns>
        public static string Replace(string text, TemplateContext context)
        {
            return Replace(text, context, new TemplateOptions());
        }

        /// <summary>将文本中的值标记替换为上下文取值，支持可选格式串与 HTML 编码。Replaces value tokens with context values, applying an optional format string, the option culture, and HTML encoding when enabled. Unresolved keys honor <see cref="TemplateOptions.LeaveUnresolvedTokens"/>; a resolved null renders as an empty string. Text or context that is null/empty is returned unchanged.</summary>
        /// <param name="text">模板文本。The template text.</param>
        /// <param name="context">数据上下文，null 时原样返回。The data context; returned unchanged when null.</param>
        /// <param name="options">渲染选项，null 时回退默认。The rendering options; default options are used when null.</param>
        /// <returns>替换后的文本。The substituted text.</returns>
        public static string Replace(string text, TemplateContext context, TemplateOptions options)
        {
            if (string.IsNullOrEmpty(text) || context == null)
            {
                return text;
            }
            if (options == null) options = new TemplateOptions();

            return TokenRegex.Replace(text, delegate (Match match)
            {
                string key = match.Groups[1].Value.Trim();
                string format = match.Groups[2].Success ? match.Groups[2].Value.Trim() : null;
                object value;
                if (!context.TryGet(key, out value))
                {
                    return options.LeaveUnresolvedTokens ? match.Value : string.Empty;
                }
                if (value == null)
                {
                    return string.Empty;
                }
                string result;
                IFormattable formattable = value as IFormattable;
                if (formattable != null)
                {
                    result = formattable.ToString(format, options.Culture);
                }
                else
                {
                    result = Convert.ToString(value, options.Culture);
                }
                if (options.HtmlEncodeValues)
                {
                    result = System.Security.SecurityElement.Escape(result);
                }
                return result;
            });
        }

        /// <summary>从单元格文本解析 #each 循环名，成功时输出该名称。Parses the #each loop name from a cell text and outputs it on success.</summary>
        /// <param name="cellText">单元格文本。The cell text.</param>
        /// <param name="name">解析出的循环名，失败时为 null。The parsed loop name, or null on failure.</param>
        /// <returns>是否存在循环起始标记。Whether a loop-start token is present.</returns>
        public static bool TryGetLoopName(string cellText, out string name)
        {
            name = null;
            if (string.IsNullOrEmpty(cellText)) return false;
            Match match = LoopStartRegex.Match(cellText);
            if (!match.Success) return false;
            name = match.Groups[1].Value.Trim();
            return true;
        }

        /// <summary>判断单元格文本是否为 #each 循环结束标记。Reports whether the cell text is a #each loop-end token.</summary>
        /// <param name="cellText">单元格文本。The cell text.</param>
        /// <returns>是否为循环结束标记。Whether it is a loop-end token.</returns>
        public static bool IsLoopEnd(string cellText)
        {
            return !string.IsNullOrEmpty(cellText) && LoopEndRegex.IsMatch(cellText);
        }

        /// <summary>判断文本中是否含 #each 循环起始标记。Reports whether the text contains a #each loop-start token.</summary>
        /// <param name="text">待检测文本。The text to inspect.</param>
        /// <returns>是否含起始标记。Whether a loop-start token is present.</returns>
        public static bool ContainsLoopStart(string text)
        {
            return !string.IsNullOrEmpty(text) && LoopStartRegex.IsMatch(text);
        }

        /// <summary>判断文本中是否含 #each 循环结束标记。Reports whether the text contains a #each loop-end token.</summary>
        /// <param name="text">待检测文本。The text to inspect.</param>
        /// <returns>是否含结束标记。Whether a loop-end token is present.</returns>
        public static bool ContainsLoopEnd(string text)
        {
            return !string.IsNullOrEmpty(text) && LoopEndRegex.IsMatch(text);
        }

        /// <summary>移除文本中的全部 #each 循环起始与结束标记。Removes all #each loop-start and loop-end tokens from the text.</summary>
        /// <param name="text">待处理文本。The text to process.</param>
        /// <returns>清除循环标记后的文本。The text with loop markers stripped.</returns>
        public static string RemoveLoopMarkers(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;
            return LoopEndRegex.Replace(LoopStartRegex.Replace(text, string.Empty), string.Empty);
        }

        /// <summary>将取值规整为列表：已是 IList 直接返回，其他可枚举（字符串除外）展开，标量包装为单元素列表。Normalizes a value into an <c>IList</c>. A value that is already an IList is returned as-is; other non-string enumerables are materialized into a list; a scalar (or null) becomes a single-element (or empty) list.</summary>
        /// <param name="value">待规整的取值。The value to normalize.</param>
        /// <returns>规整后的列表。The normalized list.</returns>
        public static IList ResolveList(object value)
        {
            if (value == null) return new List<object>();
            IList list = value as IList;
            if (list != null) return list;

            IEnumerable enumerable = value as IEnumerable;
            if (enumerable != null && !(value is string))
            {
                List<object> result = new List<object>();
                foreach (object item in enumerable)
                {
                    result.Add(item);
                }
                return result;
            }
            return new List<object> { value };
        }

        /// <summary>由单行对象构建局部上下文：字符串字典/普通字典逐项拷贝，否则反射读取公共可读属性。Builds a local row context. Dictionary-backed rows are copied key by key; otherwise each public readable, non-indexer property is reflected in. Properties that throw while reading are ignored. A null row yields an empty context.</summary>
        /// <param name="row">行数据对象。The row data object.</param>
        /// <returns>构建好的局部上下文。The constructed local context.</returns>
        public static TemplateContext CreateRowContext(object row)
        {
            TemplateContext context = new TemplateContext();
            if (row == null) return context;

            IDictionary<string, object> stringDictionary = row as IDictionary<string, object>;
            if (stringDictionary != null)
            {
                foreach (KeyValuePair<string, object> pair in stringDictionary)
                {
                    context.Set(pair.Key, pair.Value);
                }
                return context;
            }

            IDictionary dictionary = row as IDictionary;
            if (dictionary != null)
            {
                foreach (object key in dictionary.Keys)
                {
                    context.Set(Convert.ToString(key, CultureInfo.InvariantCulture), dictionary[key]);
                }
                return context;
            }

            Type type = row.GetType();
            PropertyInfo[] properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            for (int i = 0; i < properties.Length; i++)
            {
                if (properties[i].CanRead && properties[i].GetIndexParameters().Length == 0)
                {
                    try
                    {
                        context.Set(properties[i].Name, properties[i].GetValue(row, null));
                    }
                    catch (Exception)
                    {
                        // Ignore properties that throw while being read.
                    }
                }
            }
            return context;
        }

        /// <summary>合并父子上下文：局部值优先，父上下文仅填充局部缺失的键。Merges parent and child contexts: local values win, and the parent only fills keys missing locally.</summary>
        /// <param name="parent">父级上下文。The parent context.</param>
        /// <param name="local">局部上下文，优先级更高。The local context, which takes precedence.</param>
        /// <returns>合并后的新上下文。A new merged context.</returns>
        public static TemplateContext Merge(TemplateContext parent, TemplateContext local)
        {
            TemplateContext merged = new TemplateContext();
            if (local != null)
            {
                foreach (KeyValuePair<string, object> pair in local.Items)
                {
                    merged.Set(pair.Key, pair.Value);
                }
            }
            if (parent != null)
            {
                foreach (KeyValuePair<string, object> pair in parent.Items)
                {
                    if (!merged.Contains(pair.Key))
                    {
                        merged.Set(pair.Key, pair.Value);
                    }
                }
            }
            return merged;
        }

        /// <summary>以不变文化将取值转为文本，null 转为空字符串。Converts a value to text using the invariant culture; null becomes an empty string.</summary>
        /// <param name="value">待转换的取值。The value to convert.</param>
        /// <returns>转换后的文本。The converted text.</returns>
        public static string ConvertToText(object value)
        {
            if (value == null) return string.Empty;
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
    }
}
