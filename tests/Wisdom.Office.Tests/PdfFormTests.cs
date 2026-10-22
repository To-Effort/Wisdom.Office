using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// AcroForm 填表、外观生成与扁平化回归测试。Regression tests for AcroForm filling, appearance
    /// generation and flattening.
    /// </summary>
    /// <remarks>测试用源 PDF 没有表单，因此先用对象级 API 现场拼出一份含文本域、多行域、复选框、
    /// 单选组与组合框的 AcroForm（含一份假的 <c>/XFA</c>），再走读取→填值→生成外观→扁平化→加密回读五步；
    /// 结构断言直接读 <see cref="PdfFile"/> 对象，值断言走 <see cref="PdfFormField.Value"/>，
    /// 可见性断言走 <see cref="PdfDocument"/> 的完整提取链。
    /// The source PDF carries no form, so an AcroForm with a text field, a multiline field, a check box, a radio
    /// group and a combo box (plus a fake <c>/XFA</c>) is assembled first through the object-level API. The suite
    /// then runs read → fill → generate appearances → flatten → encrypted round trip; structure assertions read
    /// <see cref="PdfFile"/> objects directly, value assertions use <see cref="PdfFormField.Value"/> and
    /// visibility assertions run the whole <see cref="PdfDocument"/> extraction chain.</remarks>
    internal static partial class Program
    {
        /// <summary>多行文本标志位（bit 12）。The multiline text flag (bit 12).</summary>
        private const int MultilineFlag = 0x800;
        /// <summary>复选框标志位（bit 19）。The check box flag (bit 19).</summary>
        private const int CheckFlag = 0x40000;
        /// <summary>单选组标志位（bit 18）。The radio group flag (bit 18).</summary>
        private const int RadioFlag = 0x20000;
        /// <summary>组合框标志位（bit 18）。The combo box flag (bit 18).</summary>
        private const int ComboFlag = 0x20000;
        /// <summary>密码文本标志位（bit 13）。The password text flag (bit 13).</summary>
        private const int PasswordFlag = 0x1000;

        private static void RunPdfFormTest(string sourcePdfPath, string outputDirectory)
        {
            byte[] source = File.ReadAllBytes(sourcePdfPath);

            RunFormWithoutAcroForm(source);
            RunFormRead(BuildFormFixture(source, false), outputDirectory);
            RunFormFillAndAppearances(source, outputDirectory);
            RunFormFlattenVisible(source, outputDirectory);
            RunFormFlattenKeepsUnpainted(source);
            RunFormEncryptedRoundTrip(source, outputDirectory);
            RunFormXfaAndHybrid(source);

            Console.WriteLine("PDF form tests passed.");
        }

        /// <summary>无表单文档的读取入口不抛异常且返回空结果。The read entry points of a formless document return empty results without throwing.</summary>
        /// <param name="source">无表单的源 PDF 字节。The formless source PDF bytes.</param>
        private static void RunFormWithoutAcroForm(byte[] source)
        {
            PdfFile file = PdfFile.Load(source);
            Require(!file.HasForm, "a plain PDF must not report a form");
            Require(file.GetForm() == null, "GetForm must return null without an AcroForm");
            Require(!file.SetFormFieldValue("whatever", "x"), "filling a missing field must report false");

            Dictionary<string, string> values = new Dictionary<string, string>();
            values["whatever"] = "x";
            Require(file.FillForm(values) == 0, "FillForm must fill nothing on a formless document");
        }

        /// <summary>字段树读取：名称、类型、选项、矩形与所在页。Field-tree reading: names, kinds, options, rectangles and pages.</summary>
        /// <param name="fixture">带 AcroForm 的 PDF 字节。The PDF bytes carrying the fixture AcroForm.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunFormRead(byte[] fixture, string outputDirectory)
        {
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-form-fixture.pdf"));
            File.WriteAllBytes(path, fixture);

            PdfFile file = PdfFile.Load(fixture);
            Require(file.HasForm, "the fixture must expose an AcroForm");
            PdfForm form = file.GetForm();
            Require(form != null, "GetForm must return the fixture form");
            Require(form.Fields.Count == 8, "expected 8 field nodes, got " + form.Fields.Count);

            PdfFormField name = form.GetField("customer.name");
            Require(name != null && name.Kind == PdfFormFieldKind.Text, "customer.name must read back as a text field");
            Require(!name.IsMultiline, "customer.name is single-line");
            Require(name.PageIndex == 0, "customer.name must resolve to page 0, got " + name.PageIndex);
            Require(name.Rect.HasValue && Math.Abs(name.Rect.Value.Width - 208) < 0.01,
                "customer.name width expected 208, got " + (name.Rect.HasValue ? name.Rect.Value.Width : -1));

            PdfFormField address = form.GetField("customer.address");
            Require(address != null && address.IsMultiline, "customer.address must inherit the multiline flag");

            // 层级名要能透过父域回退命中：/T 只写在父域上，子域按 "a.b" 寻址。
            PdfFormField plan = form.GetField("order.plan");
            Require(plan != null && plan.Kind == PdfFormFieldKind.RadioButton, "order.plan must be a radio group");
            Require(plan.Children.Count == 3, "order.plan should have 3 alternatives, got " + plan.Children.Count);
            Require(plan.Value == null, "an unfilled radio group reports no value");

            PdfFormField agree = form.GetField("order.agree");
            Require(agree != null && agree.Kind == PdfFormFieldKind.CheckBox, "order.agree must be a check box");
            Require(agree.OnState == "Yes", "the check box on-state should be parsed from /AP, got " + agree.OnState);

            PdfFormField region = form.GetField("order.region");
            Require(region != null && region.Kind == PdfFormFieldKind.ComboBox, "order.region must be a combo box");
            Require(region.Options.Count == 3, "order.region should expose 3 options, got " + region.Options.Count);
            Require(region.Options[1] == "south", "the second option should be the export value 'south'");

            Require(form.NeedAppearances, "the fixture ships a stale appearance and asks viewers to rebuild");
        }

        /// <summary>填值写回 <c>/V</c> 与状态名，并在生成外观后复位 <c>/NeedAppearances</c>。Values are written into /V and state names, and /NeedAppearances resets once appearances are generated.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunFormFillAndAppearances(byte[] source, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(BuildFormFixture(source, false));
            Dictionary<string, string> values = new Dictionary<string, string>();
            values["customer.name"] = "张三 Zhang San";
            values["customer.address"] = "北京市朝阳区望京街道 12 号\nBuilding 12, Wangjing, Chaoyang, Beijing";
            values["order.agree"] = "true";
            values["order.plan"] = "pro";
            values["order.region"] = "south";
            int filled = file.FillForm(values);
            Require(filled == 5, "FillForm reported " + filled + " of 5 fields");

            PdfForm form = file.GetForm();
            Require(!form.NeedAppearances, "appearance generation must reset /NeedAppearances");
            Require(form.GetField("customer.name").Value == "张三 Zhang San", "the Chinese value must round trip in memory");
            Require(form.GetField("order.agree").Value == "Yes", "true should map onto the check box on-state");
            Require(form.GetField("order.plan").Value == "pro", "the radio group must hold the chosen state");
            Require(form.GetField("order.region").Value == "south", "a matching option should be written as a name");

            // 每个可见控件都要有自己的 /AP /N，且内容流里出现文本绘制算子。
            PdfDictionary nameDict = form.GetField("customer.name").Dictionary;
            PdfDictionary nameAppearance = Resolve(file, nameDict.Get("AP")) as PdfDictionary;
            Require(nameAppearance != null && nameAppearance.ContainsKey("N"), "customer.name gained no /AP /N");
            PdfStream stream = Resolve(file, nameAppearance.Get("N")) as PdfStream;
            Require(stream != null, "customer.name appearance is not a stream");
            string body = Encoding.ASCII.GetString(file.DecodeStream(stream));
            Require(body.Contains("Tj") && body.Contains("Tf"), "the appearance draws no text: " + body);
            Require(body.Contains(" re W n"), "the appearance is not clipped to the widget box");

            PdfDictionary agreeDict = form.GetField("order.agree").Dictionary;
            PdfDictionary agreeNormal = Resolve(file, Resolve(file, agreeDict.Get("AP")) as PdfDictionary, "N") as PdfDictionary;
            Require(agreeNormal != null && agreeNormal.ContainsKey("Off") && agreeNormal.ContainsKey("Yes"),
                "the check box must ship both /Off and /Yes states");

            // 单选组只有被选中的那一项状态名写进 /AS，其余为 /Off。
            PdfFormField plan = form.GetField("order.plan");
            int on = 0;
            for (int i = 0; i < plan.Children.Count; i++)
            {
                PdfDictionary kid = plan.Children[i].Dictionary;
                string state = kid.Get("AS") as PdfName == null ? null : ((PdfName)kid.Get("AS")).Value;
                if (state == plan.Children[i].OnState) on++;
            }
            Require(on == 1, "exactly one radio alternative must be /AS on, got " + on);

            byte[] saved = file.SaveToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-form-filled.pdf"));
            File.WriteAllBytes(path, saved);

            PdfFile reloaded = PdfFile.Load(path);
            PdfForm back = reloaded.GetForm();
            Require(back.GetField("customer.name").Value == "张三 Zhang San", "the text value did not survive the save");
            Require(back.GetField("order.plan").Value == "pro", "the radio value did not survive the save");
            Require(!back.NeedAppearances, "/NeedAppearances reset did not survive the save");
        }

        /// <summary>扁平化后值必须出现在页面内容里，且控件注释与 AcroForm 一并消失。After flattening the values must show up in the page content, with widget annotations and the AcroForm gone.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunFormFlattenVisible(byte[] source, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(BuildFormFixture(source, false));
            int pages = file.PageCount;
            Dictionary<string, string> values = new Dictionary<string, string>();
            values["customer.name"] = "李四 Li Si";
            values["customer.address"] = "望京 12 号\nWangjing 12";
            values["order.agree"] = "Yes";
            values["order.plan"] = "max";
            values["order.region"] = "north";
            file.FillForm(values);

            int flattened = file.FlattenForm();
            Require(flattened == 7, "expected 7 flattened fields, got " + flattened);
            byte[] saved = file.SaveToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-form-flattened.pdf"));
            File.WriteAllBytes(path, saved);

            PdfDocument document = PdfDocument.Load(path);
            Require(document.PageCount == pages, "flattening changed the page count");
            string text = document.ExtractText();
            RequireContains(text, "李四");
            RequireContains(text, "Li Si");
            RequireContains(text, "望京");
            RequireContains(text, "Wangjing");
            RequireContains(text, "north");

            PdfFile reloaded = PdfFile.Load(path);
            Require(!reloaded.HasForm, "a fully flattened document must drop the AcroForm");
            Require(!reloaded.RemoveFormXfa(), "nothing must be left to strip after flattening");
            for (int i = 0; i < reloaded.PageCount; i++)
            {
                PdfDictionary page = reloaded.GetPageDictionary(i);
                PdfArray annots = Resolve(reloaded, page.Get("Annots")) as PdfArray;
                Require(annots == null || annots.Count == 0, "page " + i + " still carries " + (annots == null ? 0 : annots.Count) + " annotations");
            }
        }

        /// <summary>扁平化不能吞掉画不出外观的字段：密码域与没挂到页面上的字段必须原样保留。Flattening must not swallow fields it cannot paint: password fields and unplaced fields stay untouched.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        private static void RunFormFlattenKeepsUnpainted(byte[] source)
        {
            PdfFile file = PdfFile.Load(source);
            PdfReference page = file.GetPageReference(0);
            PdfArray fields = new PdfArray();
            PdfArray annots = new PdfArray();
            fields.Add(AddWidget(file, page, "login.secret", "Tx", PasswordFlag, Rect(72, 600, 260, 616), annots));
            // 故意不放进 /Annots：这个字段有值、有矩形，但页面上找不到它的控件。
            // Deliberately kept out of /Annots: the field has a value and a rectangle but no widget on any page.
            fields.Add(AddWidget(file, page, "login.ghost", "Tx", 0, Rect(72, 560, 260, 576), new PdfArray()));

            PdfDictionary acroForm = new PdfDictionary();
            acroForm.Set("Fields", fields);
            acroForm.SetString("DA", "/Helv 12 Tf 0 g");
            file.GetPageDictionary(0).Set("Annots", annots);
            file.MarkModified(page.Number);
            file.Catalog.Set("AcroForm", file.AddObjectReference(acroForm));
            file.MarkModified(file.CatalogNumber);

            Require(file.SetFormFieldValue("login.secret", "s3cr3t"), "a password field must accept a value");
            Require(file.SetFormFieldValue("login.ghost", "unplaced"), "an unplaced field must accept a value");
            int painted = file.FlattenForm();
            Require(painted == 0, "nothing may flatten when no appearance can be painted: " + painted);
            Require(file.HasForm, "the AcroForm must survive a flatten that painted nothing");

            byte[] saved = file.SaveToBytes();
            PdfForm back = PdfFile.Load(saved).GetForm();
            Require(back.GetField("login.secret").Value == "s3cr3t", "flattening must not drop the password value");
            Require(back.GetField("login.ghost").Value == "unplaced", "flattening must not drop the unplaced value");
            Require(back.GetField("login.ghost").PageIndex == -1, "a field outside /Annots must report no page");
            Require(back.GetField("login.secret").IsPassword, "the password flag must be readable");
        }

        /// <summary>加密文档上的填表与回读：值与外观都要能在口令下重新解出。Form filling on an encrypted document: values and appearances must decrypt again under the password.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunFormEncryptedRoundTrip(byte[] source, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(BuildFormFixture(source, false));
            file.SetFormFieldValue("customer.name", "加密王 Encrypted Wang");
            file.SetFormFieldValue("order.plan", "basic");
            file.GenerateFormAppearances();

            PdfEncryptionOptions options = new PdfEncryptionOptions();
            options.Algorithm = PdfEncryptionAlgorithm.Aes128;
            options.UserPassword = "form-user";
            options.OwnerPassword = "form-owner";
            file.SetEncryption(options);

            byte[] saved = file.SaveToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-form-encrypted.pdf"));
            File.WriteAllBytes(path, saved);

            PdfFile reloaded = PdfFile.Load(path, "form-user");
            PdfForm form = reloaded.GetForm();
            Require(form != null, "the encrypted form must be readable");
            Require(form.GetField("customer.name").Value == "加密王 Encrypted Wang", "the encrypted text value did not round trip");
            Require(form.GetField("order.plan").Value == "basic", "the encrypted radio value did not round trip");

            reloaded.FlattenForm();
            byte[] flattened = reloaded.SaveToBytes();
            string flattenedPath = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-form-flattened-encrypted.pdf"));
            File.WriteAllBytes(flattenedPath, flattened);
            PdfDocument document = PdfDocument.Load(flattened, "form-user");
            RequireContains(document.ExtractText(), "Encrypted Wang");
        }

        /// <summary>XFA 剥离：混合表单去掉 <c>/XFA</c> 后仍以 AcroForm 为准。XFA stripping: a hybrid form falls back to the AcroForm once /XFA is gone.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        private static void RunFormXfaAndHybrid(byte[] source)
        {
            PdfFile file = PdfFile.Load(BuildFormFixture(source, true));
            Require(file.GetForm().HasXfa, "the hybrid fixture must expose /XFA");
            Require(file.RemoveFormXfa(), "RemoveFormXfa must report the removal");
            Require(!file.GetForm().HasXfa, "/XFA must be gone after stripping");
            Require(file.GetForm().GetField("customer.name") != null, "stripping must not break the field tree");

            byte[] saved = file.SaveToBytes();
            PdfFile reloaded = PdfFile.Load(saved);
            Require(!reloaded.GetForm().HasXfa, "/XFA must stay gone after the save");
            Require(reloaded.SetFormFieldValue("customer.name", "剥离后 after-strip"), "filling must still work after stripping");
        }

        /// <summary>
        /// 用对象级 API 现场拼出一份 AcroForm 夹具：两个文本域、一个复选框、一个三选一单选组、一个组合框。
        /// Assembles a fixture AcroForm through the object-level API: two text fields, a check box, a
        /// three-way radio group and a combo box.
        /// </summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="withXfa">true 时附带一份假的 <c>/XFA</c> 流以模拟混合表单。True attaches a fake <c>/XFA</c> stream to simulate a hybrid form.</param>
        /// <returns>带表单的 PDF 字节。The PDF bytes carrying the fixture form.</returns>
        private static byte[] BuildFormFixture(byte[] source, bool withXfa)
        {
            PdfFile file = PdfFile.Load(source);
            PdfReference page = file.GetPageReference(0);
            PdfDictionary pageDictionary = file.GetPageDictionary(0);

            PdfArray fields = new PdfArray();
            PdfArray annots = new PdfArray();

            fields.Add(AddWidget(file, page, "customer.name", "Tx", 0,
                Rect(72, 700, 280, 716), annots));
            fields.Add(AddWidget(file, page, "customer.address", "Tx", MultilineFlag,
                Rect(72, 660, 300, 700), annots));

            PdfReference agree = AddWidget(file, page, "order.agree", "Btn", CheckFlag,
                Rect(72, 620, 84, 632), annots);
            AttachButtonStates(file, agree, "Yes");
            fields.Add(agree);

            PdfDictionary plan = new PdfDictionary();
            plan.SetName("FT", "Btn");
            plan.SetString("T", "order.plan");
            plan.SetString("DA", "/Helv 0 Tf 0 g");
            plan.SetNumber("Ff", RadioFlag);
            PdfReference planReference = file.AddObjectReference(plan);
            PdfArray kids = new PdfArray();
            string[] states = { "basic", "pro", "max" };
            for (int i = 0; i < states.Length; i++)
            {
                double left = 72 + i * 30;
                PdfReference kid = AddWidget(file, page, null, null, 0, Rect(left, 580, left + 12, 592), annots);
                PdfDictionary kidDictionary = file.GetObject<PdfDictionary>(kid.Number);
                kidDictionary.Set("Parent", planReference);
                file.MarkModified(kid.Number);
                AttachButtonStates(file, kid, states[i]);
                kids.Add(kid);
            }
            plan.Set("Kids", kids);
            fields.Add(planReference);

            PdfReference region = AddWidget(file, page, "order.region", "Ch", ComboFlag,
                Rect(320, 700, 440, 716), annots);
            PdfDictionary regionDictionary = file.GetObject<PdfDictionary>(region.Number);
            regionDictionary.Set("Opt", new PdfArray()
                .Add(new PdfArray().AddName("north").Add(PdfString.Unicode("北区")))
                .Add(new PdfArray().AddName("south").Add(PdfString.Unicode("南区")))
                .Add(PdfString.Unicode("海外 Overseas")));
            file.MarkModified(region.Number);
            fields.Add(region);

            PdfDictionary acroForm = new PdfDictionary();
            acroForm.Set("Fields", fields);
            acroForm.SetString("DA", "/Helv 12 Tf 0 g");
            acroForm.SetBoolean("NeedAppearances", true);
            if (withXfa)
            {
                PdfDictionary xfa = new PdfDictionary();
                xfa.SetName("Subtype", "XML");
                acroForm.Set("XFA", file.AddObjectReference(new PdfStream(xfa, Encoding.ASCII.GetBytes("<xml/>"), false)));
            }
            PdfReference acroFormReference = file.AddObjectReference(acroForm);

            pageDictionary.Set("Annots", annots);
            file.MarkModified(page.Number);
            file.Catalog.Set("AcroForm", acroFormReference);
            file.MarkModified(file.CatalogNumber);
            return file.SaveToBytes();
        }

        /// <summary>
        /// 新建一个控件型字段并登记到页面的 <c>/Annots</c>。
        /// Creates a widget-style field and registers it in the page's <c>/Annots</c>.
        /// </summary>
        /// <param name="file">目标文件。The target file.</param>
        /// <param name="page">所在页引用。The page reference.</param>
        /// <param name="name">字段名，null 表示只作控件。The field name; null means a plain widget.</param>
        /// <param name="fieldType">域类型名，null 表示由父域决定。The field type; null inherits it from the parent.</param>
        /// <param name="flags">标志位。The flags.</param>
        /// <param name="rect">控件矩形。The widget rectangle.</param>
        /// <param name="annots">页面注释数组。The page annotation array.</param>
        /// <returns>新字段的引用。The reference of the new field.</returns>
        private static PdfReference AddWidget(
            PdfFile file,
            PdfReference page,
            string name,
            string fieldType,
            int flags,
            PdfArray rect,
            PdfArray annots)
        {
            PdfDictionary field = new PdfDictionary();
            field.SetName("Type", "Annot");
            field.SetName("Subtype", "Widget");
            if (!string.IsNullOrEmpty(fieldType)) field.SetName("FT", fieldType);
            if (!string.IsNullOrEmpty(name)) field.SetString("T", name);
            if (flags != 0) field.SetNumber("Ff", flags);
            field.Set("Rect", rect);
            field.Set("P", page);
            field.SetString("DA", "/Helv 0 Tf 0 g");
            PdfReference reference = file.AddObjectReference(field);
            annots.Add(reference);
            return reference;
        }

        /// <summary>
        /// 给按钮型控件补一份 <c>/AP /N</c> 状态字典，使勾选态名称可读。
        /// Attaches an <c>/AP /N</c> state dictionary to a button widget so its on-state name is discoverable.
        /// </summary>
        /// <param name="file">目标文件。The target file.</param>
        /// <param name="widget">控件引用。The widget reference.</param>
        /// <param name="onState">勾选态名称。The on-state name.</param>
        private static void AttachButtonStates(PdfFile file, PdfReference widget, string onState)
        {
            PdfDictionary blank = new PdfDictionary();
            blank.SetName("Type", "XObject");
            blank.SetName("Subtype", "Form");
            blank.Set("BBox", new PdfArray().AddNumber(0).AddNumber(0).AddNumber(12).AddNumber(12));

            PdfReference off = file.AddObjectReference(new PdfStream(blank.Clone(), Encoding.ASCII.GetBytes("q\nQ\n"), false));
            PdfReference on = file.AddObjectReference(new PdfStream(blank.Clone(), Encoding.ASCII.GetBytes("q\nQ\n"), false));

            PdfDictionary states = new PdfDictionary();
            states.Set("Off", off);
            states.Set(onState, on);
            PdfDictionary appearance = new PdfDictionary();
            appearance.Set("N", states);

            PdfDictionary dictionary = file.GetObject<PdfDictionary>(widget.Number);
            dictionary.Set("AP", appearance);
            dictionary.SetName("AS", "Off");
            file.MarkModified(widget.Number);
        }

        /// <summary>构造 <c>/Rect</c> 数组。Builds a <c>/Rect</c> array.</summary>
        /// <param name="left">左边界。The left edge.</param>
        /// <param name="bottom">下边界。The bottom edge.</param>
        /// <param name="right">右边界。The right edge.</param>
        /// <param name="top">上边界。The top edge.</param>
        /// <returns>四个数构成的数组。The four-element array.</returns>
        private static PdfArray Rect(double left, double bottom, double right, double top)
        {
            return new PdfArray().AddNumber(left).AddNumber(bottom).AddNumber(right).AddNumber(top);
        }

        /// <summary>解析对象引用。Resolves an object reference.</summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="value">原始对象。The raw object.</param>
        /// <returns>解析后的对象。The resolved object.</returns>
        private static PdfObject Resolve(PdfFile file, PdfObject value)
        {
            return file.Resolve(value);
        }

        /// <summary>从字典里按键解析对象。Resolves an entry of a dictionary.</summary>
        /// <param name="file">宿主文件。The owning file.</param>
        /// <param name="dictionary">字典。The dictionary.</param>
        /// <param name="key">键。The key.</param>
        /// <returns>解析后的对象或 null。The resolved object or null.</returns>
        private static PdfObject Resolve(PdfFile file, PdfDictionary dictionary, string key)
        {
            return dictionary == null ? null : file.Resolve(dictionary.Get(key));
        }
    }
}
