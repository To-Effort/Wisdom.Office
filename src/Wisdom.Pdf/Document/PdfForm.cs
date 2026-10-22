using System;
using System.Collections.Generic;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// AcroForm 表单域的语义类型。The semantic type of an AcroForm field.
    /// </summary>
    public enum PdfFormFieldKind
    {
        /// <summary>文本域（<c>/FT /Tx</c>），含多行、密码、组合框等由 <see cref="PdfFormField.Flags"/> 细分的变体。
        /// A text field (<c>/FT /Tx</c>); multiline, password and comb variants are distinguished by <see cref="PdfFormField.Flags"/>.</summary>
        Text = 0,
        /// <summary>复选框。A check box.</summary>
        CheckBox = 1,
        /// <summary>单选按钮组，子域是各选项。A radio button group whose children are the alternatives.</summary>
        RadioButton = 2,
        /// <summary>普通按钮（无值，扁平化时跳过）。A push button (valueless; skipped when flattening).</summary>
        PushButton = 3,
        /// <summary>下拉组合框。A combo box.</summary>
        ComboBox = 4,
        /// <summary>列表框。A list box.</summary>
        ListBox = 5,
        /// <summary>数字签名域，本库只读不填。A digital signature field; readable but never filled by this library.</summary>
        Signature = 6,
        /// <summary>无法归类的域。A field this library cannot classify.</summary>
        Unknown = 7
    }

    /// <summary>
    /// AcroForm 表单：字段树、默认外观与 XFA 存在性的读取入口，配合
    /// <see cref="PdfFile.SetFormFieldValue"/>、<see cref="PdfFile.FillForm(System.Collections.Generic.IDictionary{string,string})"/>
    /// 与 <see cref="PdfFile.FlattenForm"/> 完成填表。
    /// The AcroForm: the read entry point for the field tree, default appearance and XFA presence, paired with
    /// <see cref="PdfFile.SetFormFieldValue"/>, <see cref="PdfFile.FillForm(System.Collections.Generic.IDictionary{string,string})"/>
    /// and <see cref="PdfFile.FlattenForm"/> for form filling.
    /// </summary>
    public sealed class PdfForm
    {
        /// <summary>宿主文件。The owning file.</summary>
        private readonly PdfFile _file;
        /// <summary>目录里的 AcroForm 字典。The AcroForm dictionary from the catalog.</summary>
        private readonly PdfDictionary _dictionary;
        /// <summary>全部字段节点，含中间层父域。Every field node, intermediate parents included.</summary>
        private readonly List<PdfFormField> _fields;
        /// <summary>完全限定名到字段的索引（大小写不敏感）。Index from fully qualified names to fields (case-insensitive).</summary>
        private readonly Dictionary<string, PdfFormField> _byName;

        /// <summary>
        /// 由 <see cref="PdfFile.GetForm"/> 构造表单模型。
        /// Builds the form model; called from <see cref="PdfFile.GetForm"/>.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">AcroForm 字典。The AcroForm dictionary.</param>
        /// <param name="fields">字段节点列表。The field nodes.</param>
        internal PdfForm(PdfFile file, PdfDictionary dictionary, List<PdfFormField> fields)
        {
            _file = file;
            _dictionary = dictionary;
            _fields = fields ?? new List<PdfFormField>();
            _byName = new Dictionary<string, PdfFormField>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < _fields.Count; i++)
            {
                string name = _fields[i].FullyQualifiedName;
                if (!string.IsNullOrEmpty(name) && !_byName.ContainsKey(name)) _byName[name] = _fields[i];
            }
        }

        /// <summary>全部字段节点，按文档中的顺序排列。All field nodes in document order.</summary>
        public IReadOnlyList<PdfFormField> Fields { get { return _fields; } }

        /// <summary>可填写的叶子字段（不含父域与签名/按钮）。Fillable terminal fields, excluding parents, signature and push buttons.</summary>
        public IEnumerable<PdfFormField> FillableFields
        {
            get
            {
                for (int i = 0; i < _fields.Count; i++)
                {
                    PdfFormField item = _fields[i];
                    if (item.IsTerminal && !item.IsReadOnly && item.Kind != PdfFormFieldKind.PushButton) yield return item;
                }
            }
        }

        /// <summary>
        /// 是否要求阅读器自行生成外观（<c>/NeedAppearances</c>）。
        /// Whether viewers must build appearances themselves (<c>/NeedAppearances</c>).
        /// </summary>
        /// <remarks>写入字段值时会自动置为 true，调用 <see cref="PdfFile.GenerateFormAppearances"/> 后重新置为 false，
        /// 这样"只填值不生成外观"的中间状态在阅读器里也不会显示旧值。
        /// Setting field values flips this to true and <see cref="PdfFile.GenerateFormAppearances"/> flips it back to
        /// false, so the intermediate "value set, appearance stale" state never shows the old value in a viewer.</remarks>
        public bool NeedAppearances
        {
            get { return PdfFormReader.GetBoolean(_file, _dictionary, "NeedAppearances"); }
            set
            {
                if (value) _dictionary.SetBoolean("NeedAppearances", true);
                else _dictionary.Remove("NeedAppearances");
                _file.MarkCatalogModified();
            }
        }

        /// <summary>文档是否带 XFA 表单流。Whether the form carries XFA streams.</summary>
        public bool HasXfa { get { return _dictionary.ContainsKey("XFA"); } }

        /// <summary>AcroForm 字典所在对象号。The object number of the AcroForm dictionary.</summary>
        public int ObjectNumber { get; internal set; }

        /// <summary>
        /// 按完全限定名取字段，找不到返回 null；层级分隔符是点号，大小写不敏感。
        /// Looks a field up by its fully qualified name, returning null when absent; hierarchy levels are joined
        /// by dots and the lookup is case-insensitive.
        /// </summary>
        /// <param name="fullyQualifiedName">完全限定名，例如 <c>person.name</c>。The fully qualified name, e.g. <c>person.name</c>.</param>
        /// <returns>字段或 null。The field or null.</returns>
        public PdfFormField GetField(string fullyQualifiedName)
        {
            if (string.IsNullOrEmpty(fullyQualifiedName)) return null;
            PdfFormField field;
            return _byName.TryGetValue(fullyQualifiedName, out field) ? field : null;
        }

        /// <summary>
        /// 生成表单摘要文本，便于日志与调试输出。
        /// Builds a human readable form summary for logs and debugging.
        /// </summary>
        /// <returns>每行一个字段：名称、类型、当前值。One line per field: name, kind, current value.</returns>
        public override string ToString()
        {
            return "PdfForm(" + _fields.Count + " fields, XFA=" + HasXfa + ")";
        }
    }

    /// <summary>
    /// 单个表单域：既读得出名称、类型、选项与当前值，也能直接写回值。
    /// A single form field: exposes name, kind, options and current value, and accepts a new value.
    /// </summary>
    public sealed class PdfFormField
    {
        /// <summary>宿主文件。The owning file.</summary>
        private readonly PdfFile _file;
        /// <summary>字段字典。The field dictionary.</summary>
        private readonly PdfDictionary _dictionary;
        /// <summary>继承而来的域类型（<c>/FT</c>）。The inherited field type (<c>/FT</c>).</summary>
        private readonly string _fieldType;
        /// <summary>继承而来的标志位（<c>/Ff</c>）。The inherited flags (<c>/Ff</c>).</summary>
        private readonly int _flags;
        /// <summary>继承而来的默认外观串（<c>/DA</c>）。The inherited default appearance string (<c>/DA</c>).</summary>
        private readonly string _defaultAppearance;
        /// <summary>继承而来的对齐方式（<c>/Q</c>）。The inherited alignment (<c>/Q</c>).</summary>
        private readonly int _quadding;
        /// <summary>本字段的注释对象号（自身即控件，或子控件）。The annotation object numbers backing this field (itself when it is a widget, or its widget kids).</summary>
        private readonly List<int> _annotations = new List<int>();
        /// <summary>子字段。Child fields.</summary>
        private readonly List<PdfFormField> _children = new List<PdfFormField>();
        /// <summary>选项列表，元素为导出值。The option list, export values.</summary>
        private readonly List<string> _options = new List<string>();
        /// <summary>勾选态名称，复选框与单选按钮用。The on-state name of check boxes and radio buttons.</summary>
        private string _onState;
        /// <summary>控件矩形，缺失时由父域或子控件回填。The widget rectangle, back-filled from the parent or a widget kid.</summary>
        private PdfFormFieldRect _rect;
        /// <summary>所在页下标，未挂到页面上时为 -1。The zero-based page index, -1 when the field is not on a page.</summary>
        private int _pageIndex = -1;

        /// <summary>
        /// 构造字段模型；呈现信息由 <c>PdfFile.BuildForm</c> 事后回填。
        /// Builds the field model; presentation data is back-filled afterwards by <c>PdfFile.BuildForm</c>.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">字段字典。The field dictionary.</param>
        /// <param name="objectNumber">字段对象号。The field object number.</param>
        /// <param name="parent">父字段，根字段为 null。The parent field; null for a root field.</param>
        /// <param name="name">部分名 <c>/T</c>。The partial name <c>/T</c>.</param>
        /// <param name="fieldType">生效域类型。The effective field type.</param>
        /// <param name="flags">生效标志位。The effective flags.</param>
        /// <param name="defaultAppearance">生效默认外观串。The effective default appearance string.</param>
        /// <param name="quadding">生效对齐方式。The effective alignment.</param>
        internal PdfFormField(
            PdfFile file,
            PdfDictionary dictionary,
            int objectNumber,
            PdfFormField parent,
            string name,
            string fieldType,
            int flags,
            string defaultAppearance,
            int quadding)
        {
            _file = file;
            _dictionary = dictionary;
            ObjectNumber = objectNumber;
            Parent = parent;
            Name = name ?? string.Empty;
            _fieldType = fieldType ?? string.Empty;
            _flags = flags;
            _defaultAppearance = defaultAppearance ?? string.Empty;
            _quadding = quadding;
            // 没有部分名的子节点（如单选选项、纯控件）不进入名称索引，完全限定名留空。
            // Kids without a partial name (radio options, plain widgets) stay out of the name index, so the
            // fully qualified name is left empty.
            if (string.IsNullOrEmpty(Name)) FullyQualifiedName = string.Empty;
            else if (parent == null || string.IsNullOrEmpty(parent.FullyQualifiedName)) FullyQualifiedName = Name;
            else FullyQualifiedName = parent.FullyQualifiedName + "." + Name;
        }

        /// <summary>字段对象号，便于直接按对象级 API 检视或修改底层条目。The field object number, for object-level inspection and edits.</summary>
        public int ObjectNumber { get; internal set; }

        /// <summary>字段字典，供外观生成读写 <c>/AP</c>、<c>/V</c> 等条目；对外暴露以便检视原始条目。The field dictionary; appearance generation reads and writes /AP, /V and friends here, and callers can inspect the raw entries.</summary>
        public PdfDictionary Dictionary { get { return _dictionary; } }

        /// <summary>部分名（<c>/T</c>）。The partial name (<c>/T</c>).</summary>
        public string Name { get; private set; }

        /// <summary>完全限定名，层级以点号连接。The fully qualified name, hierarchy joined by dots.</summary>
        public string FullyQualifiedName { get; private set; }

        /// <summary>父字段；根字段为 null。The parent field; null for root fields.</summary>
        public PdfFormField Parent { get; private set; }

        /// <summary>子字段列表。The child fields.</summary>
        public IReadOnlyList<PdfFormField> Children { get { return _children; } }

        /// <summary>语义类型。The semantic kind.</summary>
        public PdfFormFieldKind Kind
        {
            get
            {
                if (_fieldType == "Tx") return PdfFormFieldKind.Text;
                if (_fieldType == "Sig") return PdfFormFieldKind.Signature;
                if (_fieldType == "Btn")
                {
                    if ((_flags & PdfFieldFlags.ButtonPush) != 0) return PdfFormFieldKind.PushButton;
                    if ((_flags & PdfFieldFlags.ButtonRadio) != 0 || _children.Count > 0) return PdfFormFieldKind.RadioButton;
                    if ((_flags & PdfFieldFlags.ButtonCheck) != 0 || _children.Count == 0) return PdfFormFieldKind.CheckBox;
                    return PdfFormFieldKind.CheckBox;
                }
                if (_fieldType == "Ch")
                {
                    return (_flags & PdfFieldFlags.ChoiceCombo) != 0 ? PdfFormFieldKind.ComboBox : PdfFormFieldKind.ListBox;
                }
                return PdfFormFieldKind.Unknown;
            }
        }

        /// <summary>原始域类型名（<c>/FT</c>），已按父域继承。The raw <c>/FT</c> name, resolved through inheritance.</summary>
        public string FieldType { get { return _fieldType; } }

        /// <summary>原始标志位（<c>/Ff</c>），已按父域继承。The raw <c>/Ff</c> flags, resolved through inheritance.</summary>
        public int Flags { get { return _flags; } }

        /// <summary>是否多行文本域。Whether the text field accepts multiple lines.</summary>
        public bool IsMultiline { get { return (_flags & PdfFieldFlags.TextMultiline) != 0; } }

        /// <summary>是否密码域（值只写入 <c>/V</c>，外观留空）。Whether the field is a password field (value written, appearance left blank).</summary>
        public bool IsPassword { get { return (_flags & PdfFieldFlags.TextPassword) != 0; } }

        /// <summary>是否只读。Whether the field is read-only.</summary>
        public bool IsReadOnly { get { return (_flags & PdfFieldFlags.ReadOnly) != 0; } }

        /// <summary>是否叶子字段（自身携带值与控件）。Whether this node is a terminal field carrying a value and a widget.</summary>
        public bool IsTerminal { get { return _children.Count == 0; } }

        /// <summary>最大长度（<c>/MaxLen</c>），0 表示不限制。The <c>/MaxLen</c> limit; 0 means unlimited.</summary>
        public int MaximumLength
        {
            get { return PdfFormReader.GetInt(_file, _dictionary, "MaxLen", 0); }
        }

        /// <summary>选项导出值列表（<c>/Opt</c>）。The option export values (<c>/Opt</c>).</summary>
        public IReadOnlyList<string> Options { get { return _options; } }

        /// <summary>勾选态名称；复选框与单选按钮的 <c>/V</c> 取值。The on-state name used as <c>/V</c> for check and radio fields.</summary>
        public string OnState { get { return _onState; } }

        /// <summary>所在页下标，未挂到页面上时为 -1。The page index, -1 when the field is not placed on a page.</summary>
        public int PageIndex { get { return _pageIndex; } }

        /// <summary>控件矩形；无控件时为 null。The widget rectangle; null when the field has no widget.</summary>
        public PdfFormFieldRect? Rect { get { return _rect.IsEmpty ? (PdfFormFieldRect?)null : _rect; } }

        /// <summary>外观生成使用的原始矩形。The raw rectangle used by appearance generation.</summary>
        internal PdfFormFieldRect RectValue
        {
            get { return _rect; }
            set { _rect = value; }
        }

        /// <summary>可写的所在页下标。The writable page index.</summary>
        internal int PageIndexValue
        {
            get { return _pageIndex; }
            set { _pageIndex = value; }
        }

        /// <summary>可写的勾选态名称。The writable on-state name.</summary>
        internal string OnStateValue
        {
            get { return _onState; }
            set { _onState = value; }
        }

        /// <summary>生效默认外观串（<c>/DA</c>）。The effective default appearance string (<c>/DA</c>).</summary>
        public string DefaultAppearance { get { return _defaultAppearance; } }

        /// <summary>生效对齐方式（<c>/Q</c>）：0 左、1 中、2 右。The effective alignment (<c>/Q</c>): 0 left, 1 centre, 2 right.</summary>
        public int Alignment { get { return _quadding; } }

        /// <summary>本字段关联的注释对象号。The annotation object numbers associated with this field.</summary>
        internal IReadOnlyList<int> Annotations { get { return _annotations; } }

        /// <summary>
        /// 当前值：文本域返回文本，复选框/单选/选项域返回状态名；未填写时返回 null。
        /// The current value: text for text fields, the state name for check, radio and choice fields, null when unset.
        /// </summary>
        /// <remarks>读取时按 <c>/V</c> → <c>/DV</c> 的顺序取值，与阅读器显示默认值的行为一致。
        /// Values fall back from <c>/V</c> to <c>/DV</c>, matching how viewers display the default value.</remarks>
        public string Value
        {
            get
            {
                PdfObject raw = _dictionary.Get("V");
                if (raw == null) raw = _dictionary.Get("DV");
                return PdfFormReader.ReadValue(_file, raw, _fieldType);
            }
            set
            {
                if (_fieldType == "Sig") throw new InvalidOperationException("Signature fields cannot be filled.");
                _file.WriteFieldValue(this, value);
            }
        }

        /// <summary>多选域（列表框/组合框）的取值集合，单选域返回单个值的长度为 1 的列表。
        /// The selected values of a multi-select field; single-select fields return a one-item list.</summary>
        public IReadOnlyList<string> SelectedValues
        {
            get
            {
                List<string> result = new List<string>();
                PdfObject raw = _dictionary.Get("V");
                if (raw == null) raw = _dictionary.Get("DV");
                PdfObject resolved = _file.Resolve(raw);
                if (resolved is PdfArray)
                {
                    PdfArray array = (PdfArray)resolved;
                    for (int i = 0; i < array.Count; i++)
                    {
                        string item = PdfFormReader.ReadValue(_file, array[i], _fieldType);
                        if (!string.IsNullOrEmpty(item)) result.Add(item);
                    }
                }
                else
                {
                    string single = PdfFormReader.ReadValue(_file, raw, _fieldType);
                    if (!string.IsNullOrEmpty(single)) result.Add(single);
                }
                return result;
            }
        }

        /// <summary>
        /// 以完全限定名显示字段，便于调试输出。
        /// Renders the field by fully qualified name for debugging.
        /// </summary>
        /// <returns>形如 <c>name [Text]="value"</c> 的文本。Text such as <c>name [Text]="value"</c>.</returns>
        public override string ToString()
        {
            return FullyQualifiedName + " [" + Kind + "]=\"" + (Value ?? string.Empty) + "\"";
        }

        /// <summary>子字段登记入口。Registers a child field.</summary>
        /// <param name="child">子字段。The child.</param>
        internal void AddChild(PdfFormField child) { _children.Add(child); }

        /// <summary>选项登记入口。Registers an option value.</summary>
        /// <param name="option">选项导出值。The option export value.</param>
        internal void AddOption(string option) { if (!string.IsNullOrEmpty(option)) _options.Add(option); }

        /// <summary>把字段字典的 <c>/Opt</c> 读进选项列表。Loads the field dictionary's <c>/Opt</c> into the option list.</summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">字段字典。The field dictionary.</param>
        internal void AddOptions(PdfFile file, PdfDictionary dictionary) { PdfFormReader.ReadOptions(file, dictionary, _options); }

        /// <summary>注释对象号登记入口。Registers an annotation object number.</summary>
        /// <param name="number">注释对象号。The annotation object number.</param>
        internal void AddAnnotation(int number) { if (!_annotations.Contains(number)) _annotations.Add(number); }
    }

    /// <summary>
    /// 表单域字段标志位常量（PDF 32000-1 表 221/226/227/228）。
    /// Form field flag constants (PDF 32000-1 Tables 221/226/227/228).
    /// </summary>
    internal static class PdfFieldFlags
    {
        /// <summary>必填（所有类型，bit 1）。Required (all types, bit 1).</summary>
        public const int Required = 0x1;
        /// <summary>不导出（所有类型，bit 2）。NoExport (all types, bit 2).</summary>
        public const int NoExport = 0x2;
        /// <summary>只读（所有类型，bit 3）。ReadOnly (all types, bit 3).</summary>
        public const int ReadOnly = 0x4;
        /// <summary>多行文本（Tx，bit 12）。Multiline text (Tx, bit 12).</summary>
        public const int TextMultiline = 0x800;
        /// <summary>密码文本（Tx，bit 13）。Password text (Tx, bit 13).</summary>
        public const int TextPassword = 0x1000;
        /// <summary>等宽字符格子（Tx，bit 18）。Comb of equally spaced cells (Tx, bit 18).</summary>
        public const int TextComb = 0x20000;
        /// <summary>富文本（Tx，bit 23）。Rich text (Tx, bit 23).</summary>
        public const int TextRichText = 0x400000;
        /// <summary>普通按钮（Btn，bit 17）。Push button (Btn, bit 17).</summary>
        public const int ButtonPush = 0x10000;
        /// <summary>单选组（Btn，bit 18）。Radio group (Btn, bit 18).</summary>
        public const int ButtonRadio = 0x20000;
        /// <summary>复选框（Btn，bit 19）。Check box (Btn, bit 19).</summary>
        public const int ButtonCheck = 0x40000;
        /// <summary>取消勾选时保留一个选中项（Btn，bit 20）。At least one radio stays on (Btn, bit 20).</summary>
        public const int ButtonNoToggleToOff = 0x80000;
        /// <summary>组合框（Ch，bit 18）。Combo box (Ch, bit 18).</summary>
        public const int ChoiceCombo = 0x20000;
        /// <summary>可编辑组合框（Ch，bit 19）。Editable combo box (Ch, bit 19).</summary>
        public const int ChoiceEdit = 0x40000;
        /// <summary>选项排序（Ch，bit 20）。Sorted options (Ch, bit 20).</summary>
        public const int ChoiceSort = 0x80000;
        /// <summary>多选（Ch，bit 22）。Multi-select (Ch, bit 22).</summary>
        public const int ChoiceMultiSelect = 0x200000;
    }

    /// <summary>
    /// 表单域控件矩形（页面坐标，原点左下）。A widget rectangle in page coordinates, origin bottom-left.
    /// </summary>
    public struct PdfFormFieldRect
    {
        /// <summary>左边界。The left edge.</summary>
        public double Left;
        /// <summary>下边界。The bottom edge.</summary>
        public double Bottom;
        /// <summary>右边界。The right edge.</summary>
        public double Right;
        /// <summary>上边界。The top edge.</summary>
        public double Top;

        /// <summary>宽度（点）。Width in points.</summary>
        public double Width { get { return Right - Left; } }

        /// <summary>高度（点）。Height in points.</summary>
        public double Height { get { return Top - Bottom; } }

        /// <summary>矩形是否退化（未给出 <c>/Rect</c>）。Whether the rectangle is degenerate (no <c>/Rect</c> given).</summary>
        public bool IsEmpty { get { return Right <= Left || Top <= Bottom; } }
    }

    /// <summary>
    /// 表单对象树的读取工具：解析继承条目、解码字段值、把注释映射到页面。
    /// Reading helpers for the form object tree: inherited entries, value decoding and annotation-to-page mapping.
    /// </summary>
    internal static class PdfFormReader
    {
        /// <summary>
        /// 读取名称条目，沿父域链回溯。
        /// Reads a name entry, walking up the parent chain.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">起始字典。The starting dictionary.</param>
        /// <param name="key">条目名（不含斜杠）。The entry name without the slash.</param>
        /// <returns>名称文本，缺失时为 null。The name, null when absent.</returns>
        public static string GetNameChain(PdfFile file, PdfDictionary dictionary, string key)
        {
            PdfDictionary current = dictionary;
            int guard = 0;
            while (current != null && guard++ < 32)
            {
                PdfName name = file.Resolve(current.Get(key)) as PdfName;
                if (name != null) return name.Value;
                current = GetParent(file, current);
            }
            return null;
        }

        /// <summary>
        /// 读取整数条目，沿父域链回溯。
        /// Reads an integer entry, walking up the parent chain.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">起始字典。The starting dictionary.</param>
        /// <param name="key">条目名。The entry name.</param>
        /// <param name="fallback">缺失时的返回值。The value to return when absent.</param>
        /// <returns>整数值。The integer value.</returns>
        public static int GetInheritedInt(PdfFile file, PdfDictionary dictionary, string key, int fallback)
        {
            PdfDictionary current = dictionary;
            int guard = 0;
            while (current != null && guard++ < 32)
            {
                PdfNumber number = file.Resolve(current.Get(key)) as PdfNumber;
                if (number != null) return (int)number.Value;
                current = GetParent(file, current);
            }
            return fallback;
        }

        /// <summary>
        /// 读取本层整数条目（不继承）。
        /// Reads an integer entry from this level only (no inheritance).
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">字典。The dictionary.</param>
        /// <param name="key">条目名。The entry name.</param>
        /// <param name="fallback">缺失时的返回值。The value to return when absent.</param>
        /// <returns>整数值。The integer value.</returns>
        public static int GetInt(PdfFile file, PdfDictionary dictionary, string key, int fallback)
        {
            PdfNumber number = file.Resolve(dictionary == null ? null : dictionary.Get(key)) as PdfNumber;
            return number == null ? fallback : (int)number.Value;
        }

        /// <summary>
        /// 读取本层布尔条目，缺失时返回 false。
        /// Reads a boolean entry from this level; false when absent.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">字典。The dictionary.</param>
        /// <param name="key">条目名。The entry name.</param>
        /// <returns>布尔值。The boolean value.</returns>
        public static bool GetBoolean(PdfFile file, PdfDictionary dictionary, string key)
        {
            PdfObject value = file.Resolve(dictionary == null ? null : dictionary.Get(key));
            PdfBoolean flag = value as PdfBoolean;
            if (flag != null) return flag.Value;
            PdfName name = value as PdfName;
            return name != null && name.Value == "true";
        }

        /// <summary>
        /// 读取字符串条目，沿父域链回溯。
        /// Reads a string entry, walking up the parent chain.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">起始字典。The starting dictionary.</param>
        /// <param name="key">条目名。The entry name.</param>
        /// <returns>文本，缺失时为 null。The text, null when absent.</returns>
        public static string GetStringChain(PdfFile file, PdfDictionary dictionary, string key)
        {
            PdfDictionary current = dictionary;
            int guard = 0;
            while (current != null && guard++ < 32)
            {
                string text = ReadString(file, current.Get(key));
                if (text != null) return text;
                current = GetParent(file, current);
            }
            return null;
        }

        /// <summary>
        /// 读取本层文本条目（字符串对象）。
        /// Reads a text entry (a string object) from this level.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">字典。The dictionary.</param>
        /// <param name="key">条目名。The entry name.</param>
        /// <returns>文本，缺失时为 null。The text, null when absent.</returns>
        public static string GetString(PdfFile file, PdfDictionary dictionary, string key)
        {
            return ReadString(file, dictionary == null ? null : dictionary.Get(key));
        }

        /// <summary>
        /// 取父域字典。
        /// Gets the parent field dictionary.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">当前字典。The current dictionary.</param>
        /// <returns>父字典或 null。The parent dictionary or null.</returns>
        public static PdfDictionary GetParent(PdfFile file, PdfDictionary dictionary)
        {
            PdfObject parent = dictionary == null ? null : dictionary.Get("Parent");
            return file.Resolve(parent) as PdfDictionary;
        }

        /// <summary>
        /// 读取字符串对象，自动识别 UTF-16BE（FE FF）与 PDFDocEncoding；个别生成器把 <c>/T</c> 写成名字，这里一并接受。
        /// Reads a string object, recognising UTF-16BE (FE FF) and PDFDocEncoding; names are accepted too because
        /// some producers write <c>/T</c> as a name.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="value">原始对象，可为引用。The raw object, possibly a reference.</param>
        /// <returns>解码文本，既非字符串也非名字时返回 null。The decoded text; null when the object is neither a string nor a name.</returns>
        public static string ReadString(PdfFile file, PdfObject value)
        {
            PdfObject resolved = file.Resolve(value);
            PdfString text = resolved as PdfString;
            if (text != null) return PdfDocumentInfo.DecodePdfString(text.Bytes);
            PdfName shorthand = resolved as PdfName;
            return shorthand == null ? null : shorthand.Value;
        }

        /// <summary>
        /// 按域类型读取 <c>/V</c> 或 <c>/DV</c>：文本域取字符串，按钮与选项域取名。
        /// Reads a <c>/V</c> or <c>/DV</c> entry by field type: strings for text fields, names for buttons and choices.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="raw">原始对象。The raw object.</param>
        /// <param name="fieldType">域类型名。The field type name.</param>
        /// <returns>值文本；未填写或退化为 Off 时返回 null。The value; null when unset or reduced to /Off.</returns>
        public static string ReadValue(PdfFile file, PdfObject raw, string fieldType)
        {
            PdfObject value = file.Resolve(raw);
            if (value == null) return null;
            PdfName name = value as PdfName;
            if (name != null) return name.Value == "Off" ? null : name.Value;
            PdfString text = value as PdfString;
            if (text != null)
            {
                string decoded = PdfDocumentInfo.DecodePdfString(text.Bytes);
                return string.IsNullOrEmpty(decoded) ? null : decoded;
            }
            return null;
        }

        /// <summary>
        /// 解析 <c>/Opt</c> 为导出值列表；每项可能是"导出值/显示值"二元数组。
        /// Resolves <c>/Opt</c> into export values, where each item may be a two-element export/display pair.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">字段字典。The field dictionary.</param>
        /// <param name="target">输出列表。The list to fill.</param>
        public static void ReadOptions(PdfFile file, PdfDictionary dictionary, List<string> target)
        {
            PdfArray options = file.Resolve(dictionary.Get("Opt")) as PdfArray;
            if (options == null) return;
            for (int i = 0; i < options.Count; i++)
            {
                PdfObject item = file.Resolve(options[i]);
                PdfArray pair = item as PdfArray;
                if (pair != null)
                {
                    if (pair.Count == 0) continue;
                    item = file.Resolve(pair[0]);
                }
                string text = ReadValue(file, item, "Ch");
                if (!string.IsNullOrEmpty(text)) target.Add(text);
            }
        }

        /// <summary>
        /// 读取 <c>/Rect</c> 数组并归一化为左下原点矩形。
        /// Reads the <c>/Rect</c> array and normalises it into a bottom-left origin rectangle.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">字段或控件字典。The field or widget dictionary.</param>
        /// <returns>矩形；缺失或非法时返回空矩形。The rectangle; empty when absent or malformed.</returns>
        public static PdfFormFieldRect ReadRect(PdfFile file, PdfDictionary dictionary)
        {
            PdfFormFieldRect rect = new PdfFormFieldRect();
            PdfArray array = file.Resolve(dictionary == null ? null : dictionary.Get("Rect")) as PdfArray;
            if (array == null || array.Count < 4) return rect;
            double[] values = new double[4];
            for (int i = 0; i < 4; i++)
            {
                PdfNumber number = file.Resolve(array[i]) as PdfNumber;
                values[i] = number == null ? 0 : number.Value;
            }
            rect.Left = Math.Min(values[0], values[2]);
            rect.Right = Math.Max(values[0], values[2]);
            rect.Bottom = Math.Min(values[1], values[3]);
            rect.Top = Math.Max(values[1], values[3]);
            return rect;
        }

        /// <summary>
        /// 收集复选框或单选项的勾选态名称：取 <c>/AP /N</c> 里除 <c>/Off</c> 之外的第一个键。
        /// Collects the on-state name of a check box or radio alternative: the first <c>/AP /N</c> key other than <c>/Off</c>.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">控件字典。The widget dictionary.</param>
        /// <returns>勾选态名称，无法判定时返回 null。The on-state name, null when undetectable.</returns>
        public static string ReadOnState(PdfFile file, PdfDictionary dictionary)
        {
            PdfDictionary appearance = file.Resolve(dictionary == null ? null : dictionary.Get("AP")) as PdfDictionary;
            if (appearance == null) return null;
            PdfDictionary normal = file.Resolve(appearance.Get("N")) as PdfDictionary;
            if (normal == null) return null;
            foreach (KeyValuePair<string, PdfObject> pair in normal)
            {
                if (pair.Key != "Off") return pair.Key;
            }
            return null;
        }

        /// <summary>
        /// 收集单选项从 <c>/MK /AS</c>（即 <c>/AP</c> 状态字典）读出的勾选态名称，兼容 <c>/MK</c> 只提供首选态的情况。
        /// Collects an on-state name from the appearance states, tolerating widgets that only advertise a preferred state.
        /// </summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">控件字典。The widget dictionary.</param>
        /// <param name="fallback">回退名称（通常是部分名 <c>/T</c>）。The fallback name, usually the partial name.</param>
        /// <returns>勾选态名称。The on-state name.</returns>
        public static string ReadOnStateOrFallback(PdfFile file, PdfDictionary dictionary, string fallback)
        {
            string state = ReadOnState(file, dictionary);
            if (!string.IsNullOrEmpty(state)) return state;
            return string.IsNullOrEmpty(fallback) ? null : fallback;
        }
    }
}
