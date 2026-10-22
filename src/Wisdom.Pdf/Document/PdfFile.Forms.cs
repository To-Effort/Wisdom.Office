using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>
    /// AcroForm 表单能力：读取字段树、填值、生成本库自有的外观流、扁平化以及剥离 XFA。
    /// AcroForm support: reading the field tree, filling values, building this library's own appearance
    /// streams, flattening and stripping XFA.
    /// </summary>
    /// <remarks>
    /// 设计参考 Aspose.PDF 的 <c>Document.Form.Fields</c> 与表单扁平化、iText 7 的 <c>PdfAcroForm</c>
    /// （含 <c>/NeedAppearances</c> 语义）。与它们不同的是外观流由本库自己绘制：
    /// 文本走 TrueType 子集嵌入，因此中文、日文值在表单里可以直接显示，不依赖阅读器装了 CJK 字体；
    /// 勾选与单选标记用矢量线条绘制，不依赖 ZapfDingbats。
    /// Modelled on Aspose.PDF's <c>Document.Form.Fields</c> and form flattening plus iText 7's
    /// <c>PdfAcroForm</c> (including the <c>/NeedAppearances</c> semantics). Appearances are drawn here instead of
    /// delegated to the viewer: text goes through TrueType subset embedding, so Chinese and Japanese values render
    /// without a CJK font in the viewer, and check and radio marks are vector strokes rather than ZapfDingbats glyphs.
    /// </remarks>
    public sealed partial class PdfFile
    {
        /// <summary>表单控件内的左右留白（点）。Horizontal padding inside a widget.</summary>
        private const double FormPaddingX = 1.0;
        /// <summary>表单外观的缺省字号（点）。The default widget font size in points.</summary>
        private const double FormDefaultFontSize = 11.0;
        /// <summary>表单外观允许的最小字号（点）。The smallest widget font size allowed.</summary>
        private const double FormMinFontSize = 4.0;
        /// <summary>自动缩放时字号的递减步长（点）。The size step used when auto-shrinking.</summary>
        private const double FormFontSizeStep = 0.5;

        /// <summary>已构建的表单模型缓存。The cached form model.</summary>
        private PdfForm _form;
        /// <summary>表单模型是否已构建（区分"无表单"与"尚未构建"）。Whether the model was built, separating "no form" from "not built yet".</summary>
        private bool _formBuilt;

        /// <summary>
        /// 文档是否带有 AcroForm。
        /// Whether the document carries an AcroForm.
        /// </summary>
        public bool HasForm
        {
            get { return Resolve(Catalog.Get("AcroForm")) is PdfDictionary; }
        }

        /// <summary>
        /// 取得表单模型；文档没有表单时返回 null。
        /// Gets the form model, or null when the document has no form.
        /// </summary>
        /// <remarks>模型是字段字典的视图，填值、生成外观之后无需重新获取；同一实例内复用。
        /// The model views the live field dictionaries, so filling values or regenerating appearances does not
        /// require rebuilding it; the same instance is reused.</remarks>
        /// <returns>表单模型或 null。The form or null.</returns>
        public PdfForm GetForm()
        {
            if (_formBuilt) return _form;
            _formBuilt = true;
            PdfDictionary acroForm = Resolve(Catalog.Get("AcroForm")) as PdfDictionary;
            if (acroForm == null) return null;
            _form = BuildForm(acroForm);
            return _form;
        }

        /// <summary>
        /// 按完全限定名填写一个字段。
        /// Fills one field by its fully qualified name.
        /// </summary>
        /// <param name="fullyQualifiedName">字段名，层级以点号连接。The field name; hierarchy joined by dots.</param>
        /// <param name="value">值文本；复选框与单选举荐直接传状态名，<c>null</c>、空串、<c>Off</c>、<c>false</c>、<c>0</c> 视为取消勾选。
        /// The value; for check and radio fields a state name is preferred, while <c>null</c>, empty, <c>Off</c>,
        /// <c>false</c> and <c>0</c> clear the field.</param>
        /// <returns>字段存在并完成写入时返回 true。True when the field exists and was written.</returns>
        /// <remarks>写入后会把 AcroForm 的 <c>/NeedAppearances</c> 置为 true，阅读器因此不会显示旧外观；
        /// 想让 PDF 在任意阅读器（含只认外观流的阅读器）里都显示新值，再调用
        /// <see cref="GenerateFormAppearances"/>，或改用 <see cref="FillForm(IDictionary{string,string})"/>。
        /// Writing flips <c>/NeedAppearances</c> to true so viewers stop showing the stale appearance; to render the
        /// new value everywhere (including viewers that only read appearance streams) call
        /// <see cref="GenerateFormAppearances"/>, or use <see cref="FillForm(IDictionary{string,string})"/> instead.</remarks>
        public bool SetFormFieldValue(string fullyQualifiedName, string value)
        {
            PdfForm form = GetForm();
            if (form == null) return false;
            PdfFormField field = form.GetField(fullyQualifiedName);
            if (field == null) return false;
            WriteFieldValue(field, value);
            return true;
        }

        /// <summary>
        /// 批量填写字段并生成外观流。
        /// Fills fields in bulk and generates appearance streams.
        /// </summary>
        /// <param name="values">完全限定名到值的映射，未知字段名会被跳过。Map from fully qualified names to values; unknown names are skipped.</param>
        /// <returns>实际写入的字段数。The number of fields actually written.</returns>
        public int FillForm(IDictionary<string, string> values)
        {
            return FillForm(values, true);
        }

        /// <summary>
        /// 批量填写字段，可选择是否顺带生成外观流。
        /// Fills fields in bulk, optionally regenerating appearance streams.
        /// </summary>
        /// <param name="values">完全限定名到值的映射。The map of names to values.</param>
        /// <param name="generateAppearances">true 表示写入完成后统一生成外观。True regenerates appearances once after all values are set.</param>
        /// <returns>实际写入的字段数。The number of fields actually written.</returns>
        /// <remarks>批量生成比逐字段生成更省：字体子集在整批里只嵌一份。
        /// Batching is cheaper than per-field regeneration because the font subset is embedded once for the whole batch.</remarks>
        public int FillForm(IDictionary<string, string> values, bool generateAppearances)
        {
            if (values == null) throw new ArgumentNullException("values");
            PdfForm form = GetForm();
            if (form == null) return 0;

            int filled = 0;
            foreach (KeyValuePair<string, string> pair in values)
            {
                PdfFormField field = form.GetField(pair.Key);
                if (field == null) continue;
                WriteFieldValue(field, pair.Value);
                filled++;
            }
            if (filled > 0 && generateAppearances) GenerateFormAppearances();
            return filled;
        }

        /// <summary>
        /// 为所有需要显示的字段重新生成 <c>/AP /N</c> 外观流。
        /// Rebuilds the <c>/AP /N</c> appearance stream of every field that needs one.
        /// </summary>
        /// <remarks>文本域、组合框、列表框按字段矩形与 <c>/DA</c> 字号排版：字号为 0 时自动缩放，多行域会折行；
        /// 复选框与单选按钮画出边框与两种状态（含 <c>/Off</c>）。密码域跳过，值仍保留在 <c>/V</c> 里。
        /// 完成后 <c>/NeedAppearances</c> 复位为 false。
        /// Text fields, combo boxes and list boxes are laid out inside the widget rectangle at the <c>/DA</c> size
        /// (auto-shrinking, wrapping for multiline fields); check boxes and radio buttons get both states including
        /// <c>/Off</c>. Password fields are skipped, their value staying in <c>/V</c>. <c>/NeedAppearances</c> is
        /// reset to false afterwards.</remarks>
        public void GenerateFormAppearances()
        {
            PdfForm form = GetForm();
            if (form == null) return;

            List<FormAppearanceJob> jobs = new List<FormAppearanceJob>();
            for (int i = 0; i < form.Fields.Count; i++) CollectAppearanceJobs(form.Fields[i], jobs);
            if (jobs.Count == 0)
            {
                form.NeedAppearances = false;
                return;
            }

            PdfFileObjectWriter writer = new PdfFileObjectWriter(this);
            PdfDictionary resources = new PdfDictionary();
            resources.Set("ProcSet", new PdfArray().AddName("PDF").AddName("Text").AddName("ImageB").AddName("ImageC").AddName("ImageI"));
            PdfDictionary fontResources = new PdfDictionary();
            resources.Set("Font", fontResources);
            PdfFontSet fonts = new PdfFontSet(FontResolver.Default);

            for (int i = 0; i < jobs.Count; i++) PlanAppearanceJob(jobs[i], fonts);

            // 子集必须等全部字形登记完才能落盘，所以字体准备与资源对象化都排在这次批量排版之后。
            // A subset can only be written once every glyph has been registered, so font preparation and the
            // resource object itself come after the whole batch has been planned.
            for (int i = 0; i < jobs.Count; i++)
            {
                List<FormTextLine> lines = jobs[i].Lines;
                if (lines == null) continue;
                for (int l = 0; l < lines.Count; l++)
                {
                    List<FormTextRun> runs = lines[l].Runs;
                    for (int r = 0; r < runs.Count; r++) runs[r].Font.Prepare(writer, fontResources);
                }
            }
            PdfReference resourcesReference = writer.Add(resources);

            for (int i = 0; i < jobs.Count; i++) WriteAppearance(jobs[i], writer, resourcesReference);

            form.NeedAppearances = false;
        }

        /// <summary>
        /// 把已填写的表单扁平化：外观流并入页面内容，移除对应控件注释与字段节点。
        /// Flattens the filled form: appearance streams are merged into the page content and the matching widget
        /// annotations and field nodes are removed.
        /// </summary>
        /// <remarks>
        /// 扁平化前先执行一次 <see cref="GenerateFormAppearances"/>，中文值因此同样会被烧进页面；
        /// 画不出外观的字段（签名、按钮、未挂到页面上的字段）保持交互态留在表单里。
        /// 全部字段都消失时目录里的 <c>/AcroForm</c> 一并移除，XFA 流随之失效。
        /// <see cref="GenerateFormAppearances"/> runs first, so Chinese values are burned into the page too.
        /// Fields with no drawable appearance (signatures, push buttons, unplaced fields) stay interactive.
        /// When nothing is left the catalog <c>/AcroForm</c> is dropped, discarding XFA along with it.</remarks>
        /// <returns>被扁平化的字段数。The number of flattened fields.</returns>
        public int FlattenForm()
        {
            PdfForm form = GetForm();
            if (form == null) return 0;
            GenerateFormAppearances();

            Dictionary<int, StringBuilder> byPage = new Dictionary<int, StringBuilder>();
            List<PdfFormField> flattened = new List<PdfFormField>();
            int serial = 0;
            for (int i = 0; i < form.Fields.Count; i++)
            {
                serial++;
                if (FlattenField(form.Fields[i], serial, byPage)) flattened.Add(form.Fields[i]);
            }
            if (flattened.Count == 0) return 0;

            foreach (KeyValuePair<int, StringBuilder> pair in byPage)
            {
                AppendPageContent(pair.Key, Encoding.ASCII.GetBytes(pair.Value.ToString()), true);
            }
            for (int i = 0; i < flattened.Count; i++) RemoveFieldFromTree(flattened[i]);

            // 子域全部落地后，空的中间父域（如单选组）也要一并摘掉，否则 /AcroForm 会留在目录里。
            // Once every child has been painted, emptied parents (a radio group among them) go too, otherwise
            // /AcroForm would survive in the catalog.
            for (int i = 0; i < form.Fields.Count; i++)
            {
                PdfFormField item = form.Fields[i];
                if (item.IsTerminal || flattened.Contains(item)) continue;
                PdfArray remaining = Resolve(item.Dictionary.Get("Kids")) as PdfArray;
                if (remaining == null || remaining.Count == 0) RemoveFieldFromTree(item);
            }

            PdfDictionary acroForm = Resolve(Catalog.Get("AcroForm")) as PdfDictionary;
            if (acroForm == null) return flattened.Count;
            PdfArray roots = Resolve(acroForm.Get("Fields")) as PdfArray;
            bool anyLeft = roots != null && roots.Count > 0;
            if (!anyLeft)
            {
                Catalog.Remove("AcroForm");
                MarkCatalogModified();
                _form = null;
                _formBuilt = false;
            }
            return flattened.Count;
        }

        /// <summary>
        /// 剥离 <c>/XFA</c> 流，使填写的 AcroForm 值成为唯一可见内容。
        /// Strips the <c>/XFA</c> streams so the filled AcroForm values are the only visible content.
        /// </summary>
        /// <returns>确实剥离了 XFA 时返回 true。True when XFA was present and removed.</returns>
        /// <remarks>Acrobat 渲染混合表单时会用 XFA 覆盖 AcroForm 外观，填值后必须剥离或扁平化才能所见即所得。
        /// Acrobat renders hybrid forms from XFA and ignores the AcroForm appearance, so a filled hybrid form must be
        /// stripped or flattened to be what-you-see-is-what-you-get.</remarks>
        public bool RemoveFormXfa()
        {
            PdfDictionary acroForm = Resolve(Catalog.Get("AcroForm")) as PdfDictionary;
            if (acroForm == null || !acroForm.ContainsKey("XFA")) return false;
            acroForm.Remove("XFA");
            MarkCatalogModified();
            return true;
        }

        /// <summary>标记目录已改动。Marks the catalog as modified.</summary>
        internal void MarkCatalogModified()
        {
            if (CatalogNumber >= 0) MarkModified(CatalogNumber);
        }

        /// <summary>
        /// 建立表单模型：先映射控件注释到页面，再递归解析字段树并回填矩形、页码与状态名。
        /// Builds the form model: annotations are mapped to pages first, the field tree is parsed recursively, then
        /// rectangles, page indexes and state names are filled in.
        /// </summary>
        /// <param name="acroForm">AcroForm 字典。The AcroForm dictionary.</param>
        /// <returns>表单模型。The form.</returns>
        private PdfForm BuildForm(PdfDictionary acroForm)
        {
            Dictionary<int, int> annotationPages = BuildAnnotationPageMap();
            List<PdfFormField> fields = new List<PdfFormField>();
            HashSet<int> visited = new HashSet<int>();

            PdfArray roots = Resolve(acroForm.Get("Fields")) as PdfArray;
            if (roots != null)
            {
                for (int i = 0; i < roots.Count; i++)
                {
                    PdfReference reference = roots[i] as PdfReference;
                    PdfDictionary dictionary = Resolve(roots[i]) as PdfDictionary;
                    if (dictionary == null || reference == null) continue;
                    BuildFormField(dictionary, reference.Number, null, visited, fields);
                }
            }

            for (int i = fields.Count - 1; i >= 0; i--) BackfillPresentation(fields[i], annotationPages);
            return new PdfForm(this, acroForm, fields);
        }

        /// <summary>
        /// 递归构造字段节点。
        /// Builds field nodes recursively.
        /// </summary>
        /// <param name="dictionary">字段字典。The field dictionary.</param>
        /// <param name="number">字段对象号。The field object number.</param>
        /// <param name="parent">父字段。The parent field.</param>
        /// <param name="visited">已访问对象号，防环。Visited object numbers, guarding against loops.</param>
        /// <param name="all">输出字段列表。The output field list.</param>
        /// <returns>新建的字段节点；重复访问时返回 null。The created node, or null when the object was already visited.</returns>
        private PdfFormField BuildFormField(
            PdfDictionary dictionary,
            int number,
            PdfFormField parent,
            HashSet<int> visited,
            List<PdfFormField> all)
        {
            if (!visited.Add(number)) return null;

            string name = PdfFormReader.ReadString(this, dictionary.Get("T"));
            string fieldType = PdfFormReader.GetNameChain(this, dictionary, "FT");
            int flags = PdfFormReader.GetInheritedInt(this, dictionary, "Ff", 0);
            string defaultAppearance = PdfFormReader.GetStringChain(this, dictionary, "DA");
            int quadding = PdfFormReader.GetInheritedInt(this, dictionary, "Q", 0);

            PdfFormField field = new PdfFormField(this, dictionary, number, parent, name, fieldType, flags,
                defaultAppearance, quadding);
            field.AddOptions(this, dictionary);
            field.RectValue = PdfFormReader.ReadRect(this, dictionary);
            all.Add(field);

            PdfArray kids = Resolve(dictionary.Get("Kids")) as PdfArray;
            if (kids != null)
            {
                // 单选组的子节点是"选项"，各自有勾选态；其余子域没有 /FT 与 /T 时只是同一字段的多个控件外观。
                // In a radio group every kid is an option with its own on-state; elsewhere a kid without /FT and
                // /T is just another widget of the same field.
                bool optionGroup = fieldType == "Btn" && (flags & PdfFieldFlags.ButtonRadio) != 0;
                for (int i = 0; i < kids.Count; i++)
                {
                    PdfReference reference = kids[i] as PdfReference;
                    PdfDictionary kid = Resolve(kids[i]) as PdfDictionary;
                    if (kid == null || reference == null) continue;
                    bool widgetOnly = !optionGroup && !kid.ContainsKey("FT") && !kid.ContainsKey("T") && !kid.ContainsKey("V");
                    if (widgetOnly)
                    {
                        field.AddAnnotation(reference.Number);
                        if (field.RectValue.IsEmpty) field.RectValue = PdfFormReader.ReadRect(this, kid);
                        continue;
                    }
                    PdfFormField child = BuildFormField(kid, reference.Number, field, visited, all);
                    if (child != null) field.AddChild(child);
                }
            }

            if (fieldType == "Btn")
            {
                string state = PdfFormReader.ReadOnState(this, dictionary);
                if (string.IsNullOrEmpty(state) && field.Annotations.Count > 0)
                {
                    PdfDictionary widget = GetObject<PdfDictionary>(field.Annotations[0]);
                    state = PdfFormReader.ReadOnState(this, widget);
                }
                field.OnStateValue = PdfFormReader.ReadOnStateOrFallback(this, dictionary,
                    string.IsNullOrEmpty(name) ? "Yes" : name);
                if (!string.IsNullOrEmpty(state)) field.OnStateValue = state;
            }
            return field;
        }

        /// <summary>
        /// 回填字段的所在页，并在矩形缺失时向子控件或父域借。
        /// Back-fills the page of a field and borrows a rectangle from its widget kids or its parent when absent.
        /// </summary>
        /// <param name="field">目标字段（子域已处理完）。The target field, whose children are already resolved.</param>
        /// <param name="annotationPages">注释对象号到页面下标的映射。Map from annotation object number to page index.</param>
        private void BackfillPresentation(PdfFormField field, Dictionary<int, int> annotationPages)
        {
            int page;
            if (annotationPages.TryGetValue(field.ObjectNumber, out page)) field.PageIndexValue = page;
            else
            {
                for (int i = 0; i < field.Annotations.Count; i++)
                {
                    if (annotationPages.TryGetValue(field.Annotations[i], out page))
                    {
                        field.PageIndexValue = page;
                        break;
                    }
                }
            }

            if (!field.RectValue.IsEmpty) return;
            for (int i = 0; i < field.Children.Count; i++)
            {
                if (field.Children[i].RectValue.IsEmpty) continue;
                field.RectValue = field.Children[i].RectValue;
                return;
            }
            if (field.Parent != null) field.RectValue = field.Parent.RectValue;
        }

        /// <summary>
        /// 扫描全部页面的 <c>/Annots</c>，得到注释对象号到页面下标的映射。
        /// Scans every page's <c>/Annots</c> to map annotation object numbers to page indexes.
        /// </summary>
        /// <returns>注释对象号到页面下标的映射。The map of annotation objects to pages.</returns>
        /// <remarks>内联在 <c>/Annots</c> 里的注释没有对象号，无法参与映射，这类控件不会被扁平化。
        /// Annotations inlined into <c>/Annots</c> have no object number, cannot join the map, and are not flattened.</remarks>
        private Dictionary<int, int> BuildAnnotationPageMap()
        {
            Dictionary<int, int> result = new Dictionary<int, int>();
            int count = PageCount;
            for (int index = 0; index < count; index++)
            {
                PdfDictionary page = GetPageDictionary(index);
                PdfArray annots = Resolve(page.Get("Annots")) as PdfArray;
                if (annots == null) continue;
                for (int i = 0; i < annots.Count; i++)
                {
                    PdfReference reference = annots[i] as PdfReference;
                    if (reference == null) continue;
                    if (!result.ContainsKey(reference.Number)) result[reference.Number] = index;
                }
            }
            return result;
        }

        /// <summary>
        /// 收集一个字段（及其子域）需要生成的外观任务。
        /// Collects the appearance jobs of a field and its children.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <param name="jobs">任务列表。The job list.</param>
        private void CollectAppearanceJobs(PdfFormField field, List<FormAppearanceJob> jobs)
        {
            if (field.IsPassword) return;

            if (field.Kind == PdfFormFieldKind.RadioButton)
            {
                for (int i = 0; i < field.Children.Count; i++)
                {
                    PdfFormField kid = field.Children[i];
                    if (!kid.RectValue.IsEmpty) jobs.Add(FormAppearanceJob.ForMark(kid, true));
                }
                return;
            }

            if (!field.IsTerminal)
            {
                for (int i = 0; i < field.Children.Count; i++) CollectAppearanceJobs(field.Children[i], jobs);
                return;
            }

            switch (field.Kind)
            {
                case PdfFormFieldKind.Text:
                case PdfFormFieldKind.ComboBox:
                case PdfFormFieldKind.ListBox:
                    jobs.Add(FormAppearanceJob.ForText(field));
                    break;
                case PdfFormFieldKind.CheckBox:
                    if (!field.RectValue.IsEmpty) jobs.Add(FormAppearanceJob.ForMark(field, false));
                    break;
                default:
                    break;
            }
        }

        /// <summary>
        /// 排版一个外观任务：定字号、折行并登记字形；标记类任务只量尺寸。
        /// Lays out one appearance job: picks the size, wraps the text and registers glyphs; mark jobs only measure.
        /// </summary>
        /// <param name="job">任务。The job.</param>
        /// <param name="fonts">共享字体集。The shared font set.</param>
        private void PlanAppearanceJob(FormAppearanceJob job, PdfFontSet fonts)
        {
            if (job.IsMark) return;

            PdfFormField field = job.Field;
            job.Width = field.RectValue.Width;
            job.Height = field.RectValue.Height;
            if (job.Width <= 0 || job.Height <= 0) return;

            string value = field.Value ?? string.Empty;
            double declared;
            OfficeColor color;
            ParseDefaultAppearance(field.DefaultAppearance, out declared, out color);
            job.Color = color;

            TextFormat format = new TextFormat();
            format.FontFamily = "Arial";
            bool multiline = field.IsMultiline;
            double available = Math.Max(1, job.Width - 2 * FormPaddingX);
            double usableHeight = Math.Max(1, job.Height - 2 * FormPaddingX);
            double maximum = declared > 0 ? declared : Math.Min(FormDefaultFontSize, Math.Max(6, usableHeight));

            for (double size = maximum; size >= FormMinFontSize; size -= FormFontSizeStep)
            {
                format.FontSize = size;
                List<FormTextLine> lines = WrapText(value, size, available, multiline, fonts, format);
                if (lines.Count == 0)
                {
                    job.Lines = lines;
                    job.FontSize = size;
                    return;
                }
                double needed = lines.Count * size * 1.2;
                if (multiline)
                {
                    if (needed <= usableHeight || size <= FormMinFontSize)
                    {
                        job.Lines = lines;
                        job.FontSize = size;
                        return;
                    }
                    continue;
                }
                if (lines.Count == 1 || size <= FormMinFontSize)
                {
                    job.Lines = new List<FormTextLine> { MergeSingleLine(lines) };
                    job.FontSize = size;
                    return;
                }
            }
        }

        /// <summary>
        /// 把折出的多行合回一行（单行域装不下时靠裁剪，而不是换行）。
        /// Joins the wrapped lines back into one (a single-line field clips instead of wrapping).
        /// </summary>
        /// <param name="lines">折行结果。The wrapped lines.</param>
        /// <returns>合并后的单行。The merged single line.</returns>
        private static FormTextLine MergeSingleLine(List<FormTextLine> lines)
        {
            FormTextLine merged = new FormTextLine();
            double width = 0;
            for (int i = 0; i < lines.Count; i++)
            {
                for (int r = 0; r < lines[i].Runs.Count; r++)
                {
                    FormTextRun run = lines[i].Runs[r];
                    merged.Append(run.Font, run.Text, run.Width);
                    width += run.Width;
                }
            }
            merged.Width = width;
            return merged;
        }

        /// <summary>
        /// 按字号把值折行；换行只在多行域发生，同时把用到的码点登记进子集。
        /// Wraps the value at the given size; only multiline fields break lines, and used code points are
        /// registered for subsetting.
        /// </summary>
        /// <param name="value">值文本。The value text.</param>
        /// <param name="size">字号（点）。The font size in points.</param>
        /// <param name="maxWidth">可用宽度。The available width.</param>
        /// <param name="multiline">是否允许多行。Whether multiple lines are allowed.</param>
        /// <param name="fonts">共享字体集。The shared font set.</param>
        /// <param name="format">文本格式（字号已设为 <paramref name="size"/>）。The text format, already carrying the size.</param>
        /// <returns>折行结果，空值时为空列表。The wrapped lines; empty when the value is empty.</returns>
        private List<FormTextLine> WrapText(
            string value,
            double size,
            double maxWidth,
            bool multiline,
            PdfFontSet fonts,
            TextFormat format)
        {
            List<FormTextLine> lines = new List<FormTextLine>();
            if (string.IsNullOrEmpty(value)) return lines;

            FormTextLine current = new FormTextLine();
            double width = 0;
            foreach (int codePoint in UnicodeUtil.EnumerateCodePoints(value))
            {
                if (codePoint == '\r') continue;
                if (codePoint == '\n')
                {
                    if (!multiline) continue;
                    lines.Add(current);
                    current = new FormTextLine();
                    width = 0;
                    continue;
                }

                PdfFont font = fonts.GetFontForCodePoint(format, codePoint);
                font.MarkUsed(codePoint);
                string glyph = UnicodeUtil.ToUtf16String(codePoint);
                double advance = font.GetAdvance(codePoint, size);

                if (multiline && width + advance > maxWidth && current.Runs.Count > 0)
                {
                    lines.Add(current);
                    current = new FormTextLine();
                    width = 0;
                    if (codePoint == ' ') continue;
                }
                current.Append(font, glyph, advance);
                width += advance;
                current.Width = width;
            }
            if (current.Runs.Count > 0 || lines.Count == 0) lines.Add(current);
            return lines;
        }

        /// <summary>
        /// 写出一个外观并挂到字段与其控件注释上。
        /// Writes one appearance and attaches it to the field and its widget annotations.
        /// </summary>
        /// <param name="job">任务。The job.</param>
        /// <param name="writer">对象写入器。The object writer.</param>
        /// <param name="resources">共享资源引用。The shared resources reference.</param>
        private void WriteAppearance(FormAppearanceJob job, PdfFileObjectWriter writer, PdfReference resources)
        {
            if (job.IsMark) { WriteMarkAppearance(job, writer, resources); return; }
            if (job.Width <= 0 || job.Height <= 0) return;

            string body = job.Lines == null || job.Lines.Count == 0
                ? "q\nQ\n"
                : BuildTextAppearanceBody(job);
            PdfReference appearance = AddObjectReference(new PdfStream(
                NewFormXObject(job.Width, job.Height, resources), Encoding.ASCII.GetBytes(body), true));

            AttachAppearance(job.Field, appearance, null);
        }

        /// <summary>
        /// 写出复选框与单选项的 <c>/Off</c> 与选中两种状态。
        /// Writes both the <c>/Off</c> and on states of a check box or radio alternative.
        /// </summary>
        /// <param name="job">任务。The job.</param>
        /// <param name="writer">对象写入器。The object writer.</param>
        /// <param name="resources">共享资源引用。The shared resources reference.</param>
        private void WriteMarkAppearance(FormAppearanceJob job, PdfFileObjectWriter writer, PdfReference resources)
        {
            PdfFormField field = job.Field;
            double width = field.RectValue.Width;
            double height = field.RectValue.Height;
            string onState = string.IsNullOrEmpty(field.OnState) ? (string.IsNullOrEmpty(field.Name) ? "Yes" : field.Name) : field.OnState;

            PdfReference off = AddObjectReference(new PdfStream(NewFormXObject(width, height, resources),
                Encoding.ASCII.GetBytes(BuildMarkBody(width, height, false, job.Radio)), true));
            PdfReference on = AddObjectReference(new PdfStream(NewFormXObject(width, height, resources),
                Encoding.ASCII.GetBytes(BuildMarkBody(width, height, true, job.Radio)), true));

            PdfDictionary states = new PdfDictionary();
            states.Set("Off", off);
            states.Set(onState, on);
            AttachAppearance(field, null, states);

            // 单选组的值记在父域 /V 上，选项自身只有 /AS；因此选中态要回看父域。
            // A radio group keeps its value in the parent /V while each option only records /AS, so the
            // selected state has to be read back from the parent.
            bool selected;
            if (job.Radio)
            {
                PdfFormField group = field.Parent;
                string groupValue = group == null ? null : group.Value;
                selected = !string.IsNullOrEmpty(groupValue)
                    && string.Equals(groupValue, onState, StringComparison.OrdinalIgnoreCase);
            }
            else selected = !string.IsNullOrEmpty(field.Value);
            string state = selected ? onState : "Off";
            field.Dictionary.SetName("AS", state);
            if (!job.Radio)
            {
                if (selected) field.Dictionary.SetName("V", onState);
                else field.Dictionary.Remove("V");
            }
            MarkModified(field.ObjectNumber);

            IReadOnlyList<int> annotations = field.Annotations;
            for (int i = 0; i < annotations.Count; i++)
            {
                PdfDictionary widget = GetObject<PdfDictionary>(annotations[i]);
                if (widget == null) continue;
                widget.SetName("AS", state);
                MarkModified(annotations[i]);
            }
        }

        /// <summary>
        /// 生成外观 Form XObject 字典。
        /// Builds the form XObject dictionary of an appearance.
        /// </summary>
        /// <param name="width">宽度（点）。The width in points.</param>
        /// <param name="height">高度（点）。The height in points.</param>
        /// <param name="resources">共享资源引用。The shared resources reference.</param>
        /// <returns>已设好 <c>/BBox</c> 与 <c>/Resources</c> 的字典。A dictionary with /BBox and /Resources set.</returns>
        private PdfDictionary NewFormXObject(double width, double height, PdfReference resources)
        {
            PdfDictionary form = new PdfDictionary();
            form.SetName("Type", "XObject");
            form.SetName("Subtype", "Form");
            form.SetNumber("FormType", 1);
            form.Set("BBox", new PdfArray().AddNumber(0).AddNumber(0).AddNumber(width).AddNumber(height));
            form.Set("Resources", resources);
            return form;
        }

        /// <summary>
        /// 画文本外观内容流：裁剪到控件框，逐行逐片段写 <c>Tj</c>。
        /// Draws the text appearance: clipped to the widget box, one <c>Tj</c> per run of every line.
        /// </summary>
        /// <param name="job">任务。The job.</param>
        /// <returns>内容流文本。The content stream.</returns>
        private string BuildTextAppearanceBody(FormAppearanceJob job)
        {
            StringBuilder body = new StringBuilder();
            body.Append("q\n");
            body.Append(Num(0) + " " + Num(0) + " " + Num(job.Width) + " " + Num(job.Height) + " re W n\n");
            body.Append(ColorOperator(job.Color));

            double lineHeight = job.FontSize * 1.2;
            double ascent = job.FontSize * 0.75;
            double descent = job.FontSize * 0.25;
            for (int l = 0; l < job.Lines.Count; l++)
            {
                List<FormTextRun> runs = job.Lines[l].Runs;
                for (int r = 0; r < runs.Count; r++)
                {
                    double fontAscent = runs[r].Font.GetAscent(job.FontSize);
                    double fontDescent = runs[r].Font.GetDescent(job.FontSize);
                    if (fontAscent > ascent) ascent = fontAscent;
                    if (fontDescent > descent) descent = fontDescent;
                }
            }

            double blockHeight = (job.Lines.Count - 1) * lineHeight + ascent + descent;
            double baseline = (job.Height - blockHeight) / 2 + ascent;

            for (int l = 0; l < job.Lines.Count; l++)
            {
                FormTextLine line = job.Lines[l];
                double x = FormPaddingX;
                if (job.Field.Alignment == 1) x = Math.Max(FormPaddingX, (job.Width - line.Width) / 2);
                else if (job.Field.Alignment == 2) x = Math.Max(FormPaddingX, job.Width - line.Width - FormPaddingX);

                for (int r = 0; r < line.Runs.Count; r++)
                {
                    FormTextRun run = line.Runs[r];
                    body.Append("BT\n");
                    body.Append("/" + run.Font.ResourceName + " " + Num(job.FontSize) + " Tf\n");
                    body.Append("1 0 0 1 " + Num(x) + " " + Num(baseline) + " Tm\n");
                    body.Append("<" + run.Font.EncodeText(run.Text) + "> Tj\n");
                    body.Append("ET\n");
                    x += run.Width;
                }
                baseline -= lineHeight;
            }
            body.Append("Q\n");
            return body.ToString();
        }

        /// <summary>
        /// 画复选框（方框加对勾）或单选项（圆圈加实心点）的外观内容流。
        /// Draws the appearance of a check box (square plus tick) or a radio alternative (circle plus dot).
        /// </summary>
        /// <param name="width">宽度（点）。The width in points.</param>
        /// <param name="height">高度（点）。The height in points.</param>
        /// <param name="marked">是否画选中标记。Whether to draw the on-state mark.</param>
        /// <param name="radio">true 用圆圈，false 用方框。True draws a circle, false a square.</param>
        /// <returns>内容流文本。The content stream.</returns>
        private static string BuildMarkBody(double width, double height, bool marked, bool radio)
        {
            double size = Math.Min(width, height);
            double inset = Math.Max(0.5, size * 0.08);
            double stroke = Math.Max(0.4, size * 0.06);
            StringBuilder body = new StringBuilder();
            body.Append("q\n");
            body.Append("0 0 0 RG " + Num(stroke) + " w\n");
            if (radio)
            {
                body.Append(CirclePath(width / 2, height / 2, Math.Max(0.5, size / 2 - inset)));
                body.Append("S\n");
                if (marked)
                {
                    body.Append("0 0 0 rg ");
                    body.Append(CirclePath(width / 2, height / 2, Math.Max(0.4, size * 0.2)));
                    body.Append("f\n");
                }
            }
            else
            {
                body.Append(Num(inset) + " " + Num(inset) + " " + Num(width - 2 * inset) + " " + Num(height - 2 * inset) + " re S\n");
                if (marked)
                {
                    double tick = Math.Max(1.0, size * 0.16);
                    body.Append(Num(tick) + " w 0 j 0 J\n");
                    body.Append(Num(width * 0.20) + " " + Num(height * 0.50) + " m\n");
                    body.Append(Num(width * 0.42) + " " + Num(height * 0.30) + " l\n");
                    body.Append(Num(width * 0.82) + " " + Num(height * 0.74) + " l\n");
                    body.Append("S\n");
                }
            }
            body.Append("Q\n");
            return body.ToString();
        }

        /// <summary>
        /// 生成近似圆的四段贝塞尔路径。
        /// Builds a four-segment bezier approximation of a circle.
        /// </summary>
        /// <param name="centerX">圆心横坐标。The centre X.</param>
        /// <param name="centerY">圆心纵坐标。The centre Y.</param>
        /// <param name="radius">半径。The radius.</param>
        /// <returns>路径文本（不含绘制算子）。The path text without a painting operator.</returns>
        private static string CirclePath(double centerX, double centerY, double radius)
        {
            const double k = 0.5522847498;
            double c = radius * k;
            StringBuilder path = new StringBuilder();
            path.Append(Num(centerX + radius) + " " + Num(centerY) + " m\n");
            path.Append(Num(centerX + radius) + " " + Num(centerY + c) + " " + Num(centerX + c) + " " + Num(centerY + radius) + " " + Num(centerX) + " " + Num(centerY + radius) + " c\n");
            path.Append(Num(centerX - c) + " " + Num(centerY + radius) + " " + Num(centerX - radius) + " " + Num(centerY + c) + " " + Num(centerX - radius) + " " + Num(centerY) + " c\n");
            path.Append(Num(centerX - radius) + " " + Num(centerY - c) + " " + Num(centerX - c) + " " + Num(centerY - radius) + " " + Num(centerX) + " " + Num(centerY - radius) + " c\n");
            path.Append(Num(centerX + c) + " " + Num(centerY - radius) + " " + Num(centerX + radius) + " " + Num(centerY - c) + " " + Num(centerX + radius) + " " + Num(centerY) + " c\n");
            return path.ToString();
        }

        /// <summary>
        /// 把外观挂到字段字典与其全部控件注释的 <c>/AP /N</c>。
        /// Attaches the appearance to the <c>/AP /N</c> of the field dictionary and every widget annotation.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <param name="stream">外观流引用（文本域用）。The appearance stream reference (text fields).</param>
        /// <param name="states">状态字典（按钮用）。The state dictionary (buttons).</param>
        private void AttachAppearance(PdfFormField field, PdfReference stream, PdfDictionary states)
        {
            List<PdfDictionary> targets = new List<PdfDictionary>();
            targets.Add(field.Dictionary);
            IReadOnlyList<int> annotations = field.Annotations;
            for (int i = 0; i < annotations.Count; i++)
            {
                PdfDictionary widget = GetObject<PdfDictionary>(annotations[i]);
                if (widget != null && !targets.Contains(widget)) targets.Add(widget);
            }

            for (int i = 0; i < targets.Count; i++)
            {
                PdfDictionary dictionary = targets[i];
                PdfDictionary appearance = dictionary.Get("AP") as PdfDictionary;
                if (appearance == null)
                {
                    appearance = new PdfDictionary();
                    dictionary.Set("AP", appearance);
                }
                appearance.Set("N", states != null ? (PdfObject)states : stream);
                MarkModified(field.ObjectNumber);
            }
        }

        /// <summary>
        /// 解析 <c>/DA</c> 默认外观串里的字号与填充色。
        /// Parses the font size and fill colour out of a <c>/DA</c> default appearance string.
        /// </summary>
        /// <param name="da">默认外观串。The default appearance string.</param>
        /// <param name="fontSize">字号，缺失或非法时为 0。The size; 0 when absent or malformed.</param>
        /// <param name="color">填充色，缺失时为黑色。The fill colour; black when absent.</param>
        private static void ParseDefaultAppearance(string da, out double fontSize, out OfficeColor color)
        {
            fontSize = 0;
            color = new OfficeColor(0, 0, 0);
            if (string.IsNullOrEmpty(da)) return;

            string[] tokens = da.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            double last = 0;
            List<double> numbers = new List<double>();
            for (int i = 0; i < tokens.Length; i++)
            {
                string token = tokens[i];
                double value;
                if (double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                {
                    last = value;
                    numbers.Add(value);
                    continue;
                }
                string operatorToken = token.TrimStart('/');
                if (operatorToken == "Tf") fontSize = last;
                else if (operatorToken == "rg" && numbers.Count >= 3)
                {
                    color = new OfficeColor(ToByte(numbers[numbers.Count - 3]), ToByte(numbers[numbers.Count - 2]), ToByte(numbers[numbers.Count - 1]));
                }
                else if (operatorToken == "g" && numbers.Count >= 1)
                {
                    byte gray = ToByte(numbers[numbers.Count - 1]);
                    color = new OfficeColor(gray, gray, gray);
                }
                numbers.Clear();
            }
        }

        /// <summary>把 0~1 的通道值转成字节。Converts a 0-1 channel value into a byte.</summary>
        /// <param name="value">通道值。The channel value.</param>
        /// <returns>0~255。A value from 0 to 255.</returns>
        private static byte ToByte(double value)
        {
            if (value < 0) value = 0;
            if (value > 1) value = 1;
            return (byte)Math.Round(value * 255);
        }

        /// <summary>
        /// 写入字段值：文本域写字符串，按钮与选项域写状态名，并同步各控件的 <c>/AS</c>。
        /// Writes a field value: strings for text fields, state names for buttons and choices, keeping every
        /// widget's <c>/AS</c> in sync.
        /// </summary>
        /// <param name="field">目标字段。The target field.</param>
        /// <param name="value">新值。The new value.</param>
        internal void WriteFieldValue(PdfFormField field, string value)
        {
            PdfDictionary dictionary = field.Dictionary;
            string fieldType = field.FieldType;

            if (fieldType == "Tx")
            {
                if (string.IsNullOrEmpty(value)) dictionary.Remove("V");
                else dictionary.Set("V", PdfString.Unicode(value));
                MarkModified(field.ObjectNumber);
                MarkFormNeedsAppearances();
                return;
            }
            if (fieldType == "Ch")
            {
                WriteChoiceValue(field, dictionary, value);
                MarkFormNeedsAppearances();
                return;
            }
            if (fieldType == "Btn")
            {
                WriteButtonValue(field, dictionary, value);
                MarkFormNeedsAppearances();
                return;
            }

            dictionary.Set("V", PdfString.Unicode(value ?? string.Empty));
            MarkModified(field.ObjectNumber);
        }

        /// <summary>
        /// 写入选项域值：命中 <c>/Opt</c> 时写名称，可编辑组合框允许自由文本。
        /// Writes a choice value: a name when it matches <c>/Opt</c>, free text for editable combo boxes.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <param name="dictionary">字段字典。The field dictionary.</param>
        /// <param name="value">新值。The new value.</param>
        private void WriteChoiceValue(PdfFormField field, PdfDictionary dictionary, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                dictionary.Remove("V");
                MarkModified(field.ObjectNumber);
                return;
            }
            bool known = field.Options.Count == 0;
            for (int i = 0; i < field.Options.Count; i++)
            {
                if (string.Equals(field.Options[i], value, StringComparison.Ordinal))
                {
                    known = true;
                    break;
                }
            }
            bool editable = (field.Flags & PdfFieldFlags.ChoiceEdit) != 0;
            if (known && !editable) dictionary.SetName("V", value);
            else dictionary.Set("V", PdfString.Unicode(value));
            MarkModified(field.ObjectNumber);
        }

        /// <summary>
        /// 写入按钮域值：复选框写勾选态，单选组在各选项之间分配 <c>/AS</c>。
        /// Writes a button value: the on-state for a check box, and per-alternative <c>/AS</c> for a radio group.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <param name="dictionary">字段字典。The field dictionary.</param>
        /// <param name="value">新值；<c>Off</c>、空串、<c>false</c>、<c>0</c> 表示取消。The new value; <c>Off</c>, empty, <c>false</c> and <c>0</c> clear it.</param>
        private void WriteButtonValue(PdfFormField field, PdfDictionary dictionary, string value)
        {
            bool clears = string.IsNullOrEmpty(value)
                || value == "Off"
                || value.Equals("false", StringComparison.OrdinalIgnoreCase)
                || value == "0";

            if (field.Kind == PdfFormFieldKind.RadioButton)
            {
                string selected = clears ? null : ResolveRadioState(field, value);
                if (selected == null) dictionary.Remove("V");
                else dictionary.SetName("V", selected);
                MarkModified(field.ObjectNumber);

                for (int i = 0; i < field.Children.Count; i++)
                {
                    PdfFormField kid = field.Children[i];
                    string state = string.IsNullOrEmpty(kid.OnState) ? kid.Name : kid.OnState;
                    bool on = selected != null && string.Equals(selected, state, StringComparison.Ordinal);
                    if (kid.Dictionary.ContainsKey("AS") || kid.IsTerminal) kid.Dictionary.SetName("AS", on ? (state ?? "On") : "Off");
                    MarkModified(kid.ObjectNumber);
                }
                return;
            }

            string onState = clears ? null : (IsStateName(field, value) ? value : field.OnState ?? "Yes");
            if (onState == null) dictionary.Remove("V");
            else dictionary.SetName("V", onState);
            MarkModified(field.ObjectNumber);
            ApplyStateToWidgets(field, onState);
        }

        /// <summary>
        /// 把给定值解释为单选组的选项名：状态名优先，其次选项的部分名，都不中时按序号取选项。
        /// Interprets a value as a radio alternative: the state name first, then the partial name, then the
        /// alternative by index.
        /// </summary>
        /// <param name="field">单选组字段。The radio group.</param>
        /// <param name="value">给定值。The supplied value.</param>
        /// <returns>选项名。The alternative name.</returns>
        private string ResolveRadioState(PdfFormField field, string value)
        {
            if (IsStateName(field, value)) return value;
            for (int i = 0; i < field.Children.Count; i++)
            {
                PdfFormField kid = field.Children[i];
                string state = string.IsNullOrEmpty(kid.OnState) ? kid.Name : kid.OnState;
                if (string.Equals(state, value, StringComparison.OrdinalIgnoreCase)) return state;
            }
            int index;
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out index)
                && index >= 1 && index <= field.Children.Count)
            {
                PdfFormField kid = field.Children[index - 1];
                string state = string.IsNullOrEmpty(kid.OnState) ? kid.Name : kid.OnState;
                if (!string.IsNullOrEmpty(state)) return state;
            }
            return value;
        }

        /// <summary>
        /// 判断给定值是否是该字段（或其子项）已知的勾选态名称。
        /// Whether the supplied value is a known on-state name of the field or one of its children.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <param name="value">给定值。The supplied value.</param>
        /// <returns>已知时返回 true。True when known.</returns>
        private bool IsStateName(PdfFormField field, string value)
        {
            if (string.IsNullOrEmpty(value)) return false;
            if (string.Equals(field.OnState, value, StringComparison.Ordinal)) return true;
            for (int i = 0; i < field.Children.Count; i++)
            {
                if (string.Equals(field.Children[i].OnState, value, StringComparison.Ordinal)) return true;
            }
            PdfDictionary widget = ResolveWidgetDictionary(field);
            PdfDictionary appearance = widget == null ? null : Resolve(widget.Get("AP")) as PdfDictionary;
            PdfDictionary normal = appearance == null ? null : Resolve(appearance.Get("N")) as PdfDictionary;
            return normal != null && normal.ContainsKey(value);
        }

        /// <summary>
        /// 把状态名同步到该字段所有控件注释的 <c>/AS</c>。
        /// Syncs the state name into the <c>/AS</c> of every widget annotation of the field.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <param name="state">状态名，null 表示 <c>/Off</c>。The state name; null means <c>/Off</c>.</param>
        private void ApplyStateToWidgets(PdfFormField field, string state)
        {
            IReadOnlyList<int> annotations = field.Annotations;
            for (int i = 0; i < annotations.Count; i++)
            {
                PdfDictionary widget = GetObject<PdfDictionary>(annotations[i]);
                if (widget == null) continue;
                widget.SetName("AS", state ?? "Off");
                MarkModified(annotations[i]);
            }
        }

        /// <summary>
        /// 取字段的控件字典：优先控件子项，其次字段自身。
        /// Gets the widget dictionary of a field: a widget kid first, otherwise the field itself.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <returns>控件字典或 null。The widget dictionary or null.</returns>
        private PdfDictionary ResolveWidgetDictionary(PdfFormField field)
        {
            IReadOnlyList<int> annotations = field.Annotations;
            if (annotations.Count > 0)
            {
                PdfDictionary widget = GetObject<PdfDictionary>(annotations[0]);
                if (widget != null) return widget;
            }
            return field.Dictionary;
        }

        /// <summary>把 AcroForm 标记为需要阅读器重建外观。Flags the form so viewers rebuild appearances.</summary>
        private void MarkFormNeedsAppearances()
        {
            PdfForm form = GetForm();
            if (form != null) form.NeedAppearances = true;
        }

        /// <summary>
        /// 扁平化单个字段：把 <c>/AP /N</c> 画进页面内容并从 <c>/Annots</c> 移除。
        /// Flattens one field: paints <c>/AP /N</c> into the page content and removes it from <c>/Annots</c>.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <param name="serial">序号，用于生成资源名。The serial used to build the resource name.</param>
        /// <param name="byPage">按页收集的内容流。The per-page content being collected.</param>
        /// <returns>确实处理完（画进页面或清除控件）时返回 true。True when handled (painted or cleared).</returns>
        private bool FlattenField(PdfFormField field, int serial, Dictionary<int, StringBuilder> byPage)
        {
            if (!field.IsTerminal) return false;
            if (field.Kind == PdfFormFieldKind.PushButton || field.Kind == PdfFormFieldKind.Signature) return false;

            // 画不出外观（密码域、无 /AP 的字段）或找不到宿主页面时原样保留，绝不把值丢掉。
            // When nothing can be painted (password fields, fields without /AP) or the page cannot be found the
            // field stays interactive; a value is never dropped silently.
            PdfReference appearance = ResolveNormalAppearance(field);
            if (appearance == null) return false;
            int pageIndex = ResolveWidgetPage(field);
            if (pageIndex < 0) return false;

            PdfStream stream = GetObject<PdfStream>(appearance.Number);
            PdfDictionary streamDictionary = stream == null ? null : stream.Dictionary;
            PdfFormFieldRect rect = field.RectValue;
            double[] box = AppearanceBox(streamDictionary, rect);
            double scaleX = box[2] - box[0] > 0 ? rect.Width / (box[2] - box[0]) : 1;
            double scaleY = box[3] - box[1] > 0 ? rect.Height / (box[3] - box[1]) : 1;
            double translateX = rect.Left - box[0] * scaleX;
            double translateY = rect.Bottom - box[1] * scaleY;

            string resourceName = "Fm" + serial.ToString(CultureInfo.InvariantCulture);
            SetPageResource(pageIndex, "XObject", resourceName, appearance);

            StringBuilder builder;
            if (!byPage.TryGetValue(pageIndex, out builder))
            {
                builder = new StringBuilder();
                byPage[pageIndex] = builder;
            }
            builder.Append("q\n" + Num(scaleX) + " 0 0 " + Num(scaleY) + " " + Num(translateX) + " " + Num(translateY) + " cm\n");
            builder.Append("/" + resourceName + " Do\nQ\n");
            return RemoveWidgetAnnotations(field);
        }

        /// <summary>
        /// 取字段应显示的 <c>/AP /N</c> 引用，复选与单选按 <c>/AS</c> 选状态。
        /// Gets the <c>/AP /N</c> reference to display, following <c>/AS</c> for check and radio fields.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <returns>外观流引用；无法定位时返回 null。The appearance reference, or null.</returns>
        private PdfReference ResolveNormalAppearance(PdfFormField field)
        {
            PdfDictionary widget = ResolveWidgetDictionary(field);
            PdfDictionary appearance = widget == null ? null : Resolve(widget.Get("AP")) as PdfDictionary;
            if (appearance == null) appearance = Resolve(field.Dictionary.Get("AP")) as PdfDictionary;
            if (appearance == null) return null;

            PdfObject rawNormal = appearance.Get("N");
            PdfReference direct = rawNormal as PdfReference;
            if (direct != null) return direct;
            PdfObject normal = Resolve(rawNormal);
            PdfStream inline = normal as PdfStream;
            if (inline != null) return AddObjectReference(inline);

            PdfDictionary states = normal as PdfDictionary;
            if (states == null) return null;
            string state = PdfFormReader.GetNameChain(this, widget, "AS");
            if (string.IsNullOrEmpty(state)) state = field.Value;
            if (string.IsNullOrEmpty(state)) state = "Off";
            return states.Get(state) as PdfReference;
        }

        /// <summary>
        /// 取外观流的 <c>/BBox</c>，缺失时退化为控件矩形。
        /// Gets the appearance <c>/BBox</c>, falling back to the widget rectangle.
        /// </summary>
        /// <param name="dictionary">外观流字典。The appearance stream dictionary.</param>
        /// <param name="rect">控件矩形。The widget rectangle.</param>
        /// <returns>左、下、右、上四条边。The left, bottom, right and top edges.</returns>
        private double[] AppearanceBox(PdfDictionary dictionary, PdfFormFieldRect rect)
        {
            PdfArray box = dictionary == null ? null : Resolve(dictionary.Get("BBox")) as PdfArray;
            if (box == null || box.Count < 4)
            {
                return new double[] { 0, 0, Math.Max(1, rect.Width), Math.Max(1, rect.Height) };
            }
            double[] values = new double[4];
            for (int i = 0; i < 4; i++)
            {
                PdfNumber number = Resolve(box[i]) as PdfNumber;
                values[i] = number == null ? 0 : number.Value;
            }
            return new double[]
            {
                Math.Min(values[0], values[2]),
                Math.Min(values[1], values[3]),
                Math.Max(values[0], values[2]),
                Math.Max(values[1], values[3])
            };
        }

        /// <summary>
        /// 定位字段控件所在页：先看已回填的页码，再查控件注释。
        /// Finds the page hosting the field's widget: the back-filled index first, then its widget annotations.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <returns>页面下标，未找到时 -1。The page index, -1 when not found.</returns>
        private int ResolveWidgetPage(PdfFormField field)
        {
            if (field.PageIndex >= 0) return field.PageIndex;
            IReadOnlyList<int> annotations = field.Annotations;
            if (annotations.Count == 0) return -1;
            Dictionary<int, int> map = BuildAnnotationPageMap();
            for (int i = 0; i < annotations.Count; i++)
            {
                int page;
                if (map.TryGetValue(annotations[i], out page)) return page;
            }
            return -1;
        }

        /// <summary>
        /// 从各页 <c>/Annots</c> 中移除该字段的控件注释。
        /// Removes the field's widget annotations from the pages' <c>/Annots</c>.
        /// </summary>
        /// <param name="field">字段。The field.</param>
        /// <returns>字段在树上至少一条注释（或本无控件）时返回 true。True when the field is detachable, i.e. no inline annotation blocks the removal.</returns>
        private bool RemoveWidgetAnnotations(PdfFormField field)
        {
            List<int> numbers = new List<int>();
            numbers.Add(field.ObjectNumber);
            numbers.AddRange(field.Annotations);

            Dictionary<int, int> map = BuildAnnotationPageMap();
            HashSet<int> touchedPages = new HashSet<int>();
            for (int i = 0; i < numbers.Count; i++)
            {
                int page;
                if (!map.TryGetValue(numbers[i], out page)) continue;
                PdfDictionary pageDictionary = GetPageDictionary(page);
                PdfArray annots = Resolve(pageDictionary.Get("Annots")) as PdfArray;
                if (annots != null)
                {
                    for (int a = annots.Count - 1; a >= 0; a--)
                    {
                        PdfReference reference = annots[a] as PdfReference;
                        if (reference != null && reference.Number == numbers[i]) annots.RemoveAt(a);
                    }
                }
                touchedPages.Add(page);
            }
            foreach (int page in touchedPages)
            {
                PdfDictionary pageDictionary = GetPageDictionary(page);
                MarkModified(page);
                PdfReference annotsReference = pageDictionary.Get("Annots") as PdfReference;
                if (annotsReference != null) MarkModified(annotsReference.Number);
            }
            return true;
        }

        /// <summary>
        /// 把字段从表单树里摘掉（父域 <c>/Kids</c> 或根 <c>/Fields</c>）。
        /// Detaches a field from the form tree (its parent's <c>/Kids</c> or the root <c>/Fields</c>).
        /// </summary>
        /// <param name="field">字段。The field.</param>
        private void RemoveFieldFromTree(PdfFormField field)
        {
            PdfDictionary parent = field.Parent == null ? null : field.Parent.Dictionary;
            PdfArray array;
            int ownerNumber;
            if (parent != null)
            {
                array = Resolve(parent.Get("Kids")) as PdfArray;
                ownerNumber = field.Parent.ObjectNumber;
            }
            else
            {
                PdfDictionary acroForm = Resolve(Catalog.Get("AcroForm")) as PdfDictionary;
                array = acroForm == null ? null : Resolve(acroForm.Get("Fields")) as PdfArray;
                ownerNumber = -1;
            }
            if (array == null) return;
            for (int i = array.Count - 1; i >= 0; i--)
            {
                PdfReference reference = array[i] as PdfReference;
                if (reference != null && reference.Number == field.ObjectNumber) array.RemoveAt(i);
            }
            if (parent != null) MarkModified(ownerNumber);
            else
            {
                PdfDictionary acroForm = Resolve(Catalog.Get("AcroForm")) as PdfDictionary;
                PdfReference fieldsReference = acroForm == null ? null : acroForm.Get("Fields") as PdfReference;
                if (fieldsReference != null) MarkModified(fieldsReference.Number);
            }
        }

        /// <summary>
        /// 一个字段外观的排版任务。One layout job for a field appearance.
        /// </summary>
        private sealed class FormAppearanceJob
        {
            /// <summary>目标字段。The target field.</summary>
            public PdfFormField Field;
            /// <summary>是否为勾选/单选标记外观。Whether this is a check or radio mark appearance.</summary>
            public bool IsMark;
            /// <summary>标记形状是否为单选圆圈。Whether the mark is a radio circle.</summary>
            public bool Radio;
            /// <summary>外观宽度（点）。The appearance width in points.</summary>
            public double Width;
            /// <summary>外观高度（点）。The appearance height in points.</summary>
            public double Height;
            /// <summary>排版后的行；空值时为空列表。The laid-out lines; empty when the value is empty.</summary>
            public List<FormTextLine> Lines;
            /// <summary>生效字号（点）。The effective font size.</summary>
            public double FontSize;
            /// <summary>文字颜色。The text colour.</summary>
            public OfficeColor Color;

            /// <summary>创建文本外观任务。Creates a text appearance job.</summary>
            /// <param name="field">字段。The field.</param>
            /// <returns>任务。The job.</returns>
            public static FormAppearanceJob ForText(PdfFormField field)
            {
                FormAppearanceJob job = new FormAppearanceJob();
                job.Field = field;
                return job;
            }

            /// <summary>创建勾选标记任务。Creates a mark appearance job.</summary>
            /// <param name="field">字段。The field.</param>
            /// <param name="radio">true 画圆圈，false 画方框。True draws a circle, false a square.</param>
            /// <returns>任务。The job.</returns>
            public static FormAppearanceJob ForMark(PdfFormField field, bool radio)
            {
                FormAppearanceJob job = new FormAppearanceJob();
                job.Field = field;
                job.IsMark = true;
                job.Radio = radio;
                return job;
            }
        }

        /// <summary>
        /// 外观里的一行文本。One line of appearance text.
        /// </summary>
        private sealed class FormTextLine
        {
            /// <summary>行宽（点）。The line width in points.</summary>
            public double Width;
            /// <summary>文本片段。The text runs.</summary>
            public readonly List<FormTextRun> Runs = new List<FormTextRun>();

            /// <summary>
            /// 追加一个字符：同字体就并入上一个片段，否则新开片段。
            /// Appends a character, extending the previous run when the font is unchanged.
            /// </summary>
            /// <param name="font">字体。The font.</param>
            /// <param name="text">字符文本。The character text.</param>
            /// <param name="width">字符宽度。The character width.</param>
            public void Append(PdfFont font, string text, double width)
            {
                if (Runs.Count > 0 && Runs[Runs.Count - 1].Font == font)
                {
                    Runs[Runs.Count - 1].Append(text, width);
                    return;
                }
                Runs.Add(new FormTextRun(font, text, width));
            }
        }

        /// <summary>
        /// 外观里同一字体的连续文本片段。A run of appearance text sharing one font.
        /// </summary>
        private sealed class FormTextRun
        {
            /// <summary>字体。The font.</summary>
            public readonly PdfFont Font;
            /// <summary>文本。The text.</summary>
            private readonly StringBuilder _text = new StringBuilder();
            /// <summary>宽度（点）。The width in points.</summary>
            public double Width;

            /// <summary>创建片段。Creates a run.</summary>
            /// <param name="font">字体。The font.</param>
            /// <param name="text">文本。The text.</param>
            /// <param name="width">宽度。The width.</param>
            public FormTextRun(PdfFont font, string text, double width)
            {
                Font = font;
                _text.Append(text);
                Width = width;
            }

            /// <summary>片段文本。The run text.</summary>
            public string Text { get { return _text.ToString(); } }

            /// <summary>追加文本与宽度。Appends text and width.</summary>
            /// <param name="text">文本。The text.</param>
            /// <param name="width">宽度。The width.</param>
            public void Append(string text, double width)
            {
                _text.Append(text);
                Width += width;
            }
        }
    }
}
