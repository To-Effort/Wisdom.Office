using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Wisdom.Office.Core;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 修复规范第二集 Phase 4 回归：产出损坏文件与静默丢失。
    /// Remediation design episode two, phase four: corrupt output and silently lost content.
    /// </summary>
    /// <remarks>
    /// 预算与上报批（<c>PD9</c> 与其下游 <c>PD5</c>/<c>PD22</c>）放在 <c>RemediationDecodeBudgetTests.cs</c>，
    /// Word 侧的契约批（<c>WE2</c>、<c>WE6</c>）放在 <c>RemediationWordPhaseFourTests.cs</c>，
    /// PPT 侧的形状批（<c>PT6</c> 起）放在 <c>RemediationPptPhaseFourTests.cs</c>，三者都由下面的入口末尾调用；
    /// 本文件只放 PDF 侧可单独判定的条目。
    /// The budget and reporting batch (<c>PD9</c> and its downstream <c>PD5</c>/<c>PD22</c>) lives in
    /// <c>RemediationDecodeBudgetTests.cs</c>, the Word contract batch (<c>WE2</c>, <c>WE6</c>) in
    /// <c>RemediationWordPhaseFourTests.cs</c>, and the PPT shape batch (from <c>PT6</c>) in
    /// <c>RemediationPptPhaseFourTests.cs</c>; all three are called at the bottom of the entry point below.
    /// This file carries only the PDF entries that can be decided on their own.
    /// </remarks>
    internal static partial class Program
    {
        /// <summary>Phase 4 入口。Entry point for phase four.</summary>
        private static void RunRemediationPhaseFourTest()
        {
            RunPredictorSingleRowTest();
            RunCatalogGenerationTest();
            RunOffsetNumberMismatchTest();
            RunRebuiltXrefIncrementalTest();
            RunWatermarkBBoxTest();
            RunFlattenResourceNameTest();
            RunAppearanceWriteBackTest();
            RunNamedCmapTextTest();
            RunInheritedResourceAliasTest();
            RunAbbreviatedFilterKeyTest();
            RunNameEscapeRoundTripTest();
            RunSignatureFieldGuardTest();
            RunFormDefaultAppearanceInheritanceTest();
            RunDefaultAppearanceCmykTest();
            RunPdfDateDesignatorTest();
            RunRemediationDecodeBudgetTest();
            RunRemediationWordPhaseFourTest();
            RunRemediationPptPhaseFourTest();

            Console.WriteLine("[remediation] phase-four corrupt output and silent loss ok");
        }

        /// <summary>PD16：PNG 预测器的正文恰好等于一行长度时不得解出零字节，内容流照常可用。PD16: when a PNG-predicted body is exactly one row long the decoder must not return zero bytes; the content stream still has to work.</summary>
        private static void RunPredictorSingleRowTest()
        {
            string content = "BT /F1 12 Tf 20 700 Td (Hi) Tj ET";
            byte[] raw = Encoding.ASCII.GetBytes(content);
            byte[] packed = Zlib.Compress(raw);

            // /Columns 取成正文长度，于是「一整行」就是整段正文，行首的滤波字节永远读不到。
            // With /Columns set to the body length, a whole row spans the body, so the leading filter
            // byte can never be read.
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                    + "/Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Filter /FlateDecode /Length " + packed.Length
                    + " /DecodeParms << /Predictor 12 /Colors 1 /BitsPerComponent 8 /Columns "
                    + raw.Length + " >> >>\nstream\n",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            byte[] pdf = AssembleBinaryPdf(bodies, 3, packed, "\nendstream");

            string text = PdfDocument.Load(pdf).ExtractText();
            Require(text != null && text.Contains("Hi"),
                "a predictor body of exactly one row lost the content stream, got '" + text + "'");
        }

        /// <summary>
        /// PD17：交叉表条目指到的对象头写着别的号，这个条目即失效，不得把住在该偏移上的对象当成目标号返回。
        /// PD17: when the object header at a cross-reference row's offset declares a different number, the row
        /// is void and the object living at that offset must never be handed back under the requested number.
        /// </summary>
        private static void RunOffsetNumberMismatchTest()
        {
            string content = "BT /F1 12 Tf 20 700 Td (Hi) Tj ET";
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                    + "/Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };

            // 5 号条目改指 3 号对象的头，两条落在同一个偏移上。
            // Row 5 is redirected to object 3's header, so two rows share one offset.
            long[] offsets = ObjectOffsets(bodies);
            long[] forged = new long[bodies.Length];
            for (int i = 0; i < bodies.Length; i++) forged[i] = offsets[i];
            forged[4] = offsets[2];
            byte[] pdf = AssembleClassicPdf(bodies, 1, forged);

            PdfFile file = PdfFile.Load(pdf);
            PdfObject five = file.GetObject(5);
            Require(!IsPageDictionary(five),
                "the xref row for object 5 points at object 3's header and the reader handed that object back: "
                    + Describe(five));

            string text = PdfDocument.Load(pdf).ExtractText();
            Require(text != null && text.Contains("Hi"),
                "reading through the void row lost the page text, got '" + text + "'");
        }

        /// <summary>判断对象是否是这张 fixture 里的页节点。Tells whether the object is the fixture's page node.</summary>
        /// <param name="value">待判定的对象。The object to judge.</param>
        /// <returns>是页节点返回 true。True for the page dictionary.</returns>
        private static bool IsPageDictionary(PdfObject value)
        {
            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary == null) return false;
            return dictionary.Get("Type") != null && dictionary.Get("MediaBox") != null;
        }

        /// <summary>把对象缩成一句可诊断的描述。Condenses an object into a diagnosable one-liner.</summary>
        /// <param name="value">待描述的对象。The object to describe.</param>
        /// <returns>描述文本。The description.</returns>
        private static string Describe(PdfObject value)
        {
            if (value == null) return "null";
            PdfDictionary dictionary = value as PdfDictionary;
            if (dictionary == null) return value.GetType().Name;
            StringBuilder builder = new StringBuilder("PdfDictionary");
            foreach (KeyValuePair<string, PdfObject> pair in dictionary)
            {
                builder.Append(" /").Append(pair.Key).Append('=')
                    .Append(pair.Value == null ? "null" : pair.Value.GetType().Name);
            }
            return builder.ToString();
        }

        /// <summary>
        /// PD8：交叉表靠全文扫描救回来的文档再做增量保存，追加段既没有 <c>/Prev</c> 也不含未改动的对象，
        /// 从最后一个 <c>startxref</c> 读起的阅读器只能看见这次改动的几个对象。这样的保存必须改成全量重写。
        /// PD8: an incremental save on a document whose cross-reference came from the scan fallback appends a
        /// section with neither /Prev nor the unchanged objects, so a reader starting at the last startxref
        /// sees only what this edit touched; that save has to become a full rewrite.
        /// </summary>
        private static void RunRebuiltXrefIncrementalTest()
        {
            string first = "BT /F1 12 Tf 20 700 Td (First) Tj ET";
            string second = "BT /F1 12 Tf 20 700 Td (Second) Tj ET";
            string font = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>";
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 5 0 R "
                    + "/Resources << /Font << /F1 7 0 R >> >> >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 6 0 R "
                    + "/Resources << /Font << /F1 7 0 R >> >> >>",
                "<< /Length " + first.Length + " >>\nstream\n" + first + "\nendstream",
                "<< /Length " + second.Length + " >>\nstream\n" + second + "\nendstream",
                font,
            };
            byte[] broken = BreakStartxref(AssembleClassicPdf(bodies, 1, null), 9);

            PdfFile file = PdfFile.Load(broken);
            Require(file.PageCount == 2,
                "PD8 control failed: the scan-recovered document must open with 2 pages, got " + file.PageCount);
            Require(PdfDocument.Load(broken).ExtractText().Contains("Second"),
                "PD8 control failed: the scan-recovered document itself already lost the second page text");

            file.RemovePage(0);
            byte[] saved = file.SaveIncrementalToBytes();

            PdfDocument reopened = PdfDocument.Load(saved);
            string text = reopened.ExtractText();
            Require(reopened.PageCount == 1, "the incremental update kept " + reopened.PageCount
                + " page(s), expected 1");
            Require(text != null && text.Contains("Second"),
                "the incremental update on a scan-recovered document dropped the untouched objects: "
                    + "second page text reads '" + text + "'");
        }

        /// <summary>
        /// PD11：水印表单的 <c>/BBox</c> 必须框住正文真正画到的坐标。用页面 <c>MediaBox</c> 的原点与边界时，
        /// 旋转页的阅读方向比 <c>MediaBox</c> 宽，非零原点页的正文整段落在框外，两种都在 <c>Do</c> 的裁剪里丢掉。
        /// PD11: the watermark form's /BBox has to cover the coordinates its body actually draws. Taken from the
        /// page MediaBox it clips the wider reading direction of a rotated page and the whole body of a page
        /// whose MediaBox does not start at the origin.
        /// </summary>
        private static void RunWatermarkBBoxTest()
        {
            // 用例一：612x792 的页转到 90 度后按 792x612 阅读，正文中心摆在 x 约 735 处。
            // Case one: a 612x792 page turned to 90 reads as 792x612, and the anchor lands near x=735.
            CheckWatermarkBBox("[0 0 612 792]", 90, 700, 300, 792, 612, "a page rotated to 90");
            // 用例二：原点非零的竖版页阅读方向仍是 612x792，正文中心摆在 y 约 170 处。
            // Case two: a portrait page with a raised origin still reads as 612x792, with the anchor near y=170.
            CheckWatermarkBBox("[0 200 612 992]", 0, 100, 600, 612, 792, "a page with a raised MediaBox origin");
        }

        /// <summary>给一页加自由摆放的水印，再断言表单 <c>/BBox</c> 罩得住正文的摆放点。Adds a freely placed watermark to one page and asserts its form /BBox covers the drawn anchor.</summary>
        /// <param name="mediaBox">页面 MediaBox 字面量。The literal page MediaBox.</param>
        /// <param name="rotate">页面 /Rotate，0 表示不写这个键。The page /Rotate; 0 omits the key.</param>
        /// <param name="left">水印距阅读方向左边的距离。The watermark distance from the reading-direction left edge.</param>
        /// <param name="top">水印距阅读方向上边的距离。The watermark distance from the reading-direction top edge.</param>
        /// <param name="expectWidth">表单应有的宽度。The width the form box has to carry.</param>
        /// <param name="expectHeight">表单应有的高度。The height the form box has to carry.</param>
        /// <param name="label">失败信息里的用例名。The case name used in failure messages.</param>
        private static void CheckWatermarkBBox(string mediaBox, int rotate, double left, double top,
            double expectWidth, double expectHeight, string label)
        {
            PdfFile file = PdfFile.Load(BuildWatermarkSource(mediaBox, rotate));
            WatermarkOptions options = WatermarkOptions.ForText("WD");
            options.Placement = WatermarkPlacement.Free;
            options.Angle = 0;
            options.AutoFit = false;
            options.Left = left;
            options.Top = top;
            options.Layer = WatermarkLayer.Foreground;
            file.AddWatermark(options);

            PdfFile reload = PdfFile.Load(file.SaveToBytes());
            PdfStream form = WatermarkFormStream(reload, 0);
            Require(form != null, label + ": the saved page carries no watermark form xobject");

            double[] box = BoxNumbers(form);
            Require(box != null, label + ": the watermark form was written without a /BBox");
            double[] anchor = FirstCmTranslate(Encoding.ASCII.GetString(reload.DecodeStream(form)));
            Require(anchor != null, label + ": the watermark form body has no cm to read the anchor from");

            Require(box[2] - box[0] == expectWidth && box[3] - box[1] == expectHeight,
                label + ": the form /BBox is [" + BoxNum(box[0]) + " " + BoxNum(box[1]) + " " + BoxNum(box[2]) + " "
                    + BoxNum(box[3]) + "] but the watermark draws in a " + BoxNum(expectWidth) + " x "
                    + BoxNum(expectHeight) + " reading direction");
            Require(anchor[0] >= box[0] && anchor[0] <= box[2] && anchor[1] >= box[1] && anchor[1] <= box[3],
                label + ": the watermark anchor (" + BoxNum(anchor[0]) + ", " + BoxNum(anchor[1])
                    + ") sits outside the form /BBox [" + BoxNum(box[0]) + " " + BoxNum(box[1]) + " "
                    + BoxNum(box[2]) + " " + BoxNum(box[3]) + "], so Do clips it away");
        }

        /// <summary>组装一页带正文与 Type1 字体的经典 PDF，MediaBox 与 /Rotate 由调用方给定。Assembles a one-page classic PDF with body text and a Type1 font, taking the MediaBox and /Rotate from the caller.</summary>
        /// <param name="mediaBox">页面 MediaBox 字面量。The literal page MediaBox.</param>
        /// <param name="rotate">页面 /Rotate，0 表示不写这个键。The page /Rotate; 0 omits the key.</param>
        /// <returns>完整的 PDF 字节。The complete PDF bytes.</returns>
        private static byte[] BuildWatermarkSource(string mediaBox, int rotate)
        {
            string content = "BT /F1 12 Tf 60 700 Td (Body) Tj ET";
            string rotation = rotate == 0 ? string.Empty : " /Rotate " + rotate;
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox " + mediaBox + rotation
                    + " /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            return AssembleClassicPdf(bodies, 1, null);
        }

        /// <summary>从页面资源里取出水印表单对象流。Picks the watermark form stream out of the page resources.</summary>
        /// <param name="file">已重开的文档。The reopened document.</param>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>表单流；找不到时 null。The form stream, or null when absent.</returns>
        private static PdfStream WatermarkFormStream(PdfFile file, int index)
        {
            string outer = WatermarkOuterStream(file, index);
            if (outer == null) return null;
            int slash = outer.IndexOf('/');
            if (slash < 0) return null;
            int end = slash + 1;
            while (end < outer.Length && !char.IsWhiteSpace(outer[end])) end++;
            string name = outer.Substring(slash + 1, end - slash - 1);

            PdfDictionary page = file.GetPageDictionary(index);
            PdfDictionary resources = file.Resolve(page.Get("Resources")) as PdfDictionary;
            PdfDictionary xobjects = resources == null ? null : file.Resolve(resources.Get("XObject")) as PdfDictionary;
            if (xobjects == null) return null;
            return file.Resolve(xobjects.Get(name)) as PdfStream;
        }

        /// <summary>读出表单字典里 <c>/BBox</c> 的四个数。Reads the four numbers of the form dictionary's /BBox.</summary>
        /// <param name="form">水印表单流。The watermark form stream.</param>
        /// <returns>四个数；缺失或不足四个时 null。The four numbers, or null when absent or short.</returns>
        private static double[] BoxNumbers(PdfStream form)
        {
            PdfArray box = form.Dictionary.Get("BBox") as PdfArray;
            if (box == null || box.Count < 4) return null;
            double[] values = new double[4];
            for (int i = 0; i < 4; i++)
            {
                PdfNumber number = i < box.Count ? box[i] as PdfNumber : null;
                if (number == null) return null;
                values[i] = number.Value;
            }
            return values;
        }

        /// <summary>取正文里第一条 <c>cm</c> 的平移分量，即水印的摆放点。Reads the translation of the body's first cm, which is where the watermark is anchored.</summary>
        /// <param name="body">解码后的表单正文。The decoded form body.</param>
        /// <returns>平移分量；读不到六个数时 null。The translation, or null when six operands cannot be read.</returns>
        private static double[] FirstCmTranslate(string body)
        {
            string[] tokens = body.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 6; i < tokens.Length; i++)
            {
                if (tokens[i] != "cm") continue;
                double[] values = new double[6];
                bool ok = true;
                for (int k = 0; k < 6; k++)
                {
                    double parsed;
                    if (!double.TryParse(tokens[i - 6 + k], NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
                    {
                        ok = false;
                        break;
                    }
                    values[k] = parsed;
                }
                if (ok) return new double[] { values[4], values[5] };
            }
            return null;
        }

        /// <summary>把坐标写进失败信息里，最多四位小数。Formats a coordinate for failure messages with up to four decimals.</summary>
        /// <param name="value">坐标值。The coordinate.</param>
        /// <returns>不变文化的数字串。The culture-invariant number text.</returns>
        private static string BoxNum(double value)
        {
            return value.ToString("0.####", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// PD14：扁平化给外观流挑资源名时用的是「Fm + 本轮序号」，序号每轮都从 1 起，于是第二轮的 <c>/Fm1</c>
        /// 直接顶掉第一轮留在页面资源里的那一条，而第一轮追加的 <c>q…/Fm1 Do Q</c> 还在，画出来就成了另一个字段。
        /// PD14: flattening names each appearance "Fm&lt;serial&gt;" and the serial restarts at one every round, so
        /// the second round's /Fm1 overwrites the first round's page resource while the content that round
        /// appended still does /Fm1 Do - and paints the other field's appearance.
        /// </summary>
        private static void RunFlattenResourceNameTest()
        {
            PdfFile file = PdfFile.Load(BuildWatermarkSource("[0 0 612 792]", 0));
            PdfReference page = file.GetPageReference(0);
            PdfArray fields = new PdfArray();
            PdfArray annots = new PdfArray();
            fields.Add(AddWidget(file, page, "keep.now", "Tx", 0, Rect(72, 600, 260, 616), annots));
            // 第二个域故意不挂进 /Annots：这一轮画不出它，它得留到下一轮，序号于是又从 1 数起。
            // The second field stays out of /Annots on purpose: it cannot be painted this round, so it waits for
            // the next one and the serial restarts at one there.
            fields.Add(AddWidget(file, page, "keep.later", "Tx", 0, Rect(72, 560, 260, 576), new PdfArray()));

            PdfDictionary acroForm = new PdfDictionary();
            acroForm.Set("Fields", fields);
            acroForm.SetString("DA", "/Helv 12 Tf 0 g");
            file.GetPageDictionary(0).Set("Annots", annots);
            file.MarkModified(page.Number);
            file.Catalog.Set("AcroForm", file.AddObjectReference(acroForm));
            file.MarkModified(file.CatalogNumber);

            Require(file.SetFormFieldValue("keep.now", "Now"), "the placed field must accept a value");
            Require(file.SetFormFieldValue("keep.later", "Later"), "the deferred field must accept a value");
            int firstRoundPainted = file.FlattenForm();
            Require(firstRoundPainted == 1, "only the placed field may flatten in the first round, got " + firstRoundPainted);

            PdfFile again = PdfFile.Load(file.SaveToBytes());
            PdfFormField deferred = again.GetForm().GetField("keep.later");
            Require(deferred != null, "the unplaced field must survive a flatten that could not paint it");
            PdfDictionary deferredPage = again.GetPageDictionary(0);
            PdfArray placed = new PdfArray();
            PdfArray firstRound = again.Resolve(deferredPage.Get("Annots")) as PdfArray;
            if (firstRound != null)
            {
                for (int i = 0; i < firstRound.Count; i++) placed.Add(firstRound[i]);
            }
            placed.Add(new PdfReference(deferred.ObjectNumber, 0));
            deferredPage.Set("Annots", placed);
            again.MarkModified(again.GetPageReference(0).Number);

            // 补挂之后必须重开一份：表单模型是按 /Annots 建立的，同一份内存里的旧模型还不知道它有了控件。
            // The document is reopened after the surgery: the form model is built from /Annots, and the one
            // already in memory has no idea the widget is placed now.
            PdfFile third = PdfFile.Load(again.SaveToBytes());
            int secondRound = third.FlattenForm();
            Require(secondRound == 1, "the deferred field must flatten once it is placed on the page, got " + secondRound);
            PdfFile back = PdfFile.Load(third.SaveToBytes());

            int[] targets = DistinctAppearanceTargets(back, 0);
            Require(targets.Length == 2, "two flatten rounds left only " + targets.Length
                + " appearance(s) reachable from the page: the second round reused the first round's resource name");
        }

        /// <summary>数页面内容里每个 <c>Do</c> 实际指到的不同对象号。Counts the distinct object numbers the page content's Do operators resolve to.</summary>
        /// <param name="file">已重开的文档。The reopened document.</param>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>去重后的对象号。The object numbers without duplicates.</returns>
        private static int[] DistinctAppearanceTargets(PdfFile file, int index)
        {
            string content = Encoding.ASCII.GetString(file.GetPageContentBytes(index));
            string[] tokens = content.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            PdfDictionary page = file.GetPageDictionary(index);
            PdfDictionary resources = file.Resolve(page.Get("Resources")) as PdfDictionary;
            PdfDictionary group = resources == null ? null : file.Resolve(resources.Get("XObject")) as PdfDictionary;
            List<int> numbers = new List<int>();
            for (int i = 1; i < tokens.Length; i++)
            {
                if (tokens[i] != "Do" || tokens[i - 1].Length < 2 || tokens[i - 1][0] != '/') continue;
                PdfReference reference = group == null ? null : group.Get(tokens[i - 1].Substring(1)) as PdfReference;
                if (reference != null && !numbers.Contains(reference.Number)) numbers.Add(reference.Number);
            }
            return numbers.ToArray();
        }

        /// <summary>
        /// PD15：外观要同时挂到字段字典和它的每一个控件注释上，可登记改动时只报了字段号，增量保存于是原样写回控件的旧字节；
        /// 读 <c>/AP</c> 又漏了 <c>Resolve</c>，住在间接对象里的 <c>/AP</c> 会被整块换成内联字典，<c>/R</c> 与状态映射一起丢。
        /// PD15: an appearance is attached to the field dictionary and to each of its widget annotations, yet only
        /// the field number was marked modified, so an incremental save re-emitted the widgets' old bytes; and
        /// reading /AP without Resolve replaced an indirect /AP wholesale, dropping /R along with the state maps.
        /// </summary>
        private static void RunAppearanceWriteBackTest()
        {
            // 用例一：字段本体与控件分住两个对象，两个都是从原交叉表里读出来的旧对象。
            // 增量保存必须把控件的新 /AP 一起写出去。
            // Case one: field and widget live in two objects, both loaded from the original cross-reference,
            // and the incremental update has to carry the widget's new /AP as well.
            PdfFile file = PdfFile.Load(BuildSplitFieldSource(false));
            Require(file.SetFormFieldValue("pd15.value", "Typed"), "the text field must accept a value");
            file.GenerateFormAppearances();

            PdfReference expected = AppearanceReference(file, file.GetObject<PdfDictionary>(WidgetNumber));
            Require(expected != null && expected.Number != OldAppearanceNumber,
                "appearance generation never reached the widget annotation");
            PdfFile incremental = PdfFile.Load(file.SaveIncrementalToBytes());
            PdfReference written = AppearanceReference(incremental, incremental.GetObject<PdfDictionary>(WidgetNumber));
            Require(written != null && written.Number == expected.Number,
                "the incremental save left the widget's old appearance in the file: /AP /N should be "
                    + expected.Number + " 0 R but the saved widget carries "
                    + (written == null ? "no appearance at all" : written.Number.ToString(CultureInfo.InvariantCulture) + " 0 R"));

            // 用例二：控件的 /AP 住在间接对象里，重写只能改那一个字典，不能整块换成内联的新字典。
            // Case two: the widget's /AP is an indirect dictionary; rewriting it must edit that dictionary,
            // not swap in a fresh inline one.
            PdfFile second = PdfFile.Load(BuildSplitFieldSource(true));
            Require(second.SetFormFieldValue("pd15.value", "Later"), "the second text field must accept a value");
            second.GenerateFormAppearances();
            PdfFile back = PdfFile.Load(second.SaveToBytes());
            PdfDictionary savedAp = back.Resolve(back.GetObject<PdfDictionary>(WidgetNumber).Get("AP")) as PdfDictionary;
            Require(savedAp != null, "the widget lost its /AP dictionary entirely after regenerating appearances");
            Require(savedAp.Get("R") != null,
                "regenerating the appearance threw away the indirect /AP together with its /R entry");
        }

        /// <summary>样例里控件注释的对象号。The widget annotation object number of the sample file.</summary>
        private const int WidgetNumber = 6;

        /// <summary>样例里那份旧外观流的对象号。The old appearance stream object number of the sample file.</summary>
        private const int OldAppearanceNumber = 10;

        /// <summary>
        /// 拼一个「字段本体 7 号 + 控件注释 6 号」的文本域 PDF。
        /// 控件与字段都是文件里原有的间接对象，不是本次编辑新建的，改动登记漏没漏才看得出来。
        /// Assembles a PDF whose text field lives in object 7 and whose widget annotation lives in object 6.
        /// Both are pre-existing indirect objects rather than ones added by this edit, which is what makes a
        /// missing modification mark observable.
        /// </summary>
        /// <param name="indirectAppearance">控件的 <c>/AP</c> 是否住在间接对象 9 号里。Whether the widget's /AP is the indirect object 9.</param>
        /// <returns>完整的 PDF 字节。The complete PDF bytes.</returns>
        private static byte[] BuildSplitFieldSource(bool indirectAppearance)
        {
            string content = "BT /Helv 12 Tf 60 700 Td (Body) Tj ET";
            string oldBody = "q\n1 0 0 1 0 0 cm\nQ\n";
            string appearance = indirectAppearance
                ? "/AP 9 0 R"
                : "/AP << /N " + OldAppearanceNumber + " 0 R >>";
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R /AcroForm 8 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R"
                    + " /Resources << /Font << /Helv 5 0 R >> >> /Annots [6 0 R] >>",
                "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
                "<< /Type /Annot /Subtype /Widget /Rect [72 600 260 616] /P 3 0 R /Parent 7 0 R " + appearance + " >>",
                "<< /FT /Tx /T (pd15.value) /DA (/Helv 12 Tf 0 g) /Kids [6 0 R] >>",
                "<< /Fields [7 0 R] /DA (/Helv 12 Tf 0 g) /DR << /Font << /Helv 5 0 R >> >> >>",
                "<< /N " + OldAppearanceNumber + " 0 R /R " + OldAppearanceNumber + " 0 R >>",
                "<< /Type /XObject /Subtype /Form /BBox [0 0 188 16] /Length " + oldBody.Length
                    + " >>\nstream\n" + oldBody + "\nendstream",
            };
            return AssembleClassicPdf(bodies, 1, null);
        }

        /// <summary>取出字典上 <c>/AP /N</c> 的外观引用，<c>/AP</c> 是间接对象也一样。Reads the /AP /N appearance reference off a dictionary, following /AP when it is indirect.</summary>
        /// <param name="file">所在文档。The owning document.</param>
        /// <param name="holder">持有 /AP 的字典。The dictionary carrying /AP.</param>
        /// <returns>外观引用；没有时 null。The reference, or null.</returns>
        private static PdfReference AppearanceReference(PdfFile file, PdfDictionary holder)
        {
            if (holder == null) return null;
            PdfDictionary appearance = file.Resolve(holder.Get("AP")) as PdfDictionary;
            if (appearance == null) return null;
            return appearance.Get("N") as PdfReference;
        }

        /// <summary>
        /// PD18：<c>/Encoding</c> 名字从没被读过，于是「码位本身就是 UCS-2」的命名 CMap（部分国产驱动产物）
        /// 在没有 <c>/ToUnicode</c> 时整页抽出为空。
        /// PD18: the /Encoding name was never read, so a named CMap whose codes are already UCS-2 units
        /// (what some region-specific drivers emit) extracted as an empty page once /ToUnicode was absent.
        /// </summary>
        private static void RunNamedCmapTextTest()
        {
            // 用例一：/UniGB-UCS2-H，两个字节直接就是一个字符。
            // Case one: with /UniGB...-H... each byte pair is already one character.
            string extracted = PdfDocument.Load(BuildNamedCmapSource("UniGB-UCS2-H", "4E2D65876D4B8BD5")).ExtractText();
            Require(extracted.IndexOf("中文测试", StringComparison.Ordinal) >= 0,
                "a /UniGB-UCS2-H page without /ToUnicode extracted '" + extracted + "' instead of 中文测试");

            // 用例二：UTF-16 命名 CMap 同理，两两大端拼起来就是文本。
            // Case two: a UTF-16 named CMap is the same story, big-endian pairs spell the text.
            Require(PdfDocument.Load(BuildNamedCmapSource("UniGB-UTF16-H", "4E2D65876D4B8BD5")).ExtractText()
                .IndexOf("中文测试", StringComparison.Ordinal) >= 0,
                "a /UniGB-UTF16-H page without /ToUnicode lost its text");

            // 用例三：正文落单一个字节时，成对的部分仍要抽出来，不能整串作废。
            // Case three: when the body ends with an odd byte, the paired part still has to come out.
            string truncated = PdfDocument.Load(BuildNamedCmapSource("UniGB-UCS2-H", "4E2D65")).ExtractText();
            Require(truncated.IndexOf('中') >= 0,
                "a truncated /UniGB...-H body lost the paired character, got '" + truncated + "'");

            // 控制用例：Identity 与 UTF-8 的码位不是字符编号，不能当成大端码元伪造出文字。
            // Control: Identity and UTF-8 codes are not character numbers, so they must not be faked as
            // big-endian code units.
            Require(PdfDocument.Load(BuildNamedCmapSource("Identity-H", "4E2D65876D4B8BD5")).ExtractText()
                .IndexOf("中文测试", StringComparison.Ordinal) < 0,
                "an /Identity-H page without /ToUnicode was turned into invented text");
            Require(PdfDocument.Load(BuildNamedCmapSource("UniGB-UTF8-H", "4E2D65876D4B8BD5")).ExtractText()
                .IndexOf("中文测试", StringComparison.Ordinal) < 0,
                "a /UniGB...-H... body is variable length; reading it as fixed pairs must not claim 中文测试");
        }

        /// <summary>
        /// 拼一页 Type0 字体文档：编码名由参数给定，正文是一段十六进制的双字节串，且刻意不写 <c>/ToUnicode</c>。
        /// Assembles a one-page document with a Type0 font: the encoding name comes from the argument, the
        /// body is a hex string of two-byte codes, and /ToUnicode is deliberately left out.
        /// </summary>
        /// <param name="encodingName">编码名，如 <c>UniGB-UCS2-H</c>。The encoding name, such as UniGB-UCS2-H.</param>
        /// <param name="hexBody">十六进制显示串（不含尖括号）。The display string in hex, without the angle brackets.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildNamedCmapSource(string encodingName, string hexBody)
        {
            string content = "BT /F1 12 Tf 50 700 Td <" + hexBody + "> Tj ET";
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                    + "/Resources << /Font << /F1 5 0 R >> >> >>",
                "<< /Length " + content.Length + " >>\nstream\n" + content + "\nendstream",
                "<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /" + encodingName
                    + " /DescendantFonts [6 0 R] >>",
                "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /STSong-Light "
                    + "/CIDSystemInfo << /Registry (Adobe) /Ordering (GB18030) /Supplement 5 >> /DW 1000 >>",
            };
            return AssembleClassicPdf(bodies, 1, null);
        }

        /// <summary>
        /// PD19：页面自己不带 <c>/Resources</c> 时，<c>SetPageResource</c> 把继承来的那一份直接钉在本页上再落笔，
        /// 一次登记同时写进了 <c>/Pages</c> 与所有兄弟页；同名条目被后一页顶掉前一页，前一页画的是后一页的对象。
        /// PD19: when a page carries no /Resources of its own, SetPageResource pins the inherited dictionary
        /// onto that page and writes straight into it, so one registration lands on /Pages and on every
        /// sibling page at once; the later page overwrites the earlier one under the same name, leaving the
        /// earlier page to draw the later page's object.
        /// </summary>
        private static void RunInheritedResourceAliasTest()
        {
            PdfFile file = PdfFile.Load(BuildInheritedResourcePdf());

            int firstNumber = file.AddObject(BuildLogoForm());
            int secondNumber = file.AddObject(BuildLogoForm());
            file.SetPageResource(0, "XObject", "Logo", new PdfReference(firstNumber, 0));
            file.SetPageResource(1, "XObject", "Logo", new PdfReference(secondNumber, 0));

            PdfDictionary group0 = XObjectGroup(file, 0);
            PdfDictionary group1 = XObjectGroup(file, 1);
            Require(ReferenceNumber(group0 == null ? null : group0.Get("Logo")) == firstNumber,
                "page 0 registered /Logo as " + firstNumber + " 0 R but the entry now reads '"
                    + ReferenceText(group0 == null ? null : group0.Get("Logo"))
                    + "': the sibling page wrote through the shared inherited /Resources");
            Require(!ReferenceEquals(group0, group1),
                "both pages resolve /XObject to one and the same group, so a per-page resource name is "
                    + "a document-wide name");

            // 第三页什么都没登记，却也跟着多了这个分组——继承字典被就地改写的直接证据。
            // The third page registered nothing yet inherits the group as well, which is the smoking gun
            // that the inherited dictionary itself was rewritten in place.
            Require(file.GetPageResources(2).Get("XObject") == null,
                "page 2 never registered a resource, but the inherited /Pages dictionary gained an "
                    + "/XObject group holding " + Describe(file.GetPageResources(2).Get("XObject")) + " of its own");

            // 存一趟再读回来：同名条目在文件里也得各归各页，不能整篇只剩最后一次登记。
            // Round-tripped through a save, the same name still has to mean a different object per page
            // instead of collapsing onto the last registration.
            PdfFile back = PdfFile.Load(file.SaveToBytes());
            PdfDictionary saved0 = XObjectGroup(back, 0);
            PdfDictionary saved1 = XObjectGroup(back, 1);
            Require(ReferenceNumber(saved0 == null ? null : saved0.Get("Logo")) == firstNumber,
                "after saving, page 0's /Logo is '" + ReferenceText(saved0 == null ? null : saved0.Get("Logo"))
                    + "' instead of " + firstNumber + " 0 R");
            Require(ReferenceNumber(saved1 == null ? null : saved1.Get("Logo")) == secondNumber,
                "after saving, page 1's /Logo is '" + ReferenceText(saved1 == null ? null : saved1.Get("Logo"))
                    + "' instead of " + secondNumber + " 0 R");

            // 控制用例：/Font 分组本来就在继承字典里，页面登记新字体时不能把兄弟页的字体也换掉。
            // Control: the /Font group already lives in the inherited dictionary, so registering a font on
            // one page must not swap the sibling page's font.
            int helvetica = FontNumber(file, 0);
            file.SetPageResource(0, "Font", "F1", new PdfReference(secondNumber, 0));
            Require(FontNumber(file, 1) == helvetica,
                "page 0 re-registered /F1, but page 1's /F1 now resolves to the same object");
        }

        /// <summary>
        /// PD24：流字典按规范认 <c>/DP</c> 缩写，却认不出同族的 <c>/F</c> 缩写——写成 <c>/F /Fl</c> 的正文
        /// 当成未过滤的原字节交回，整页文本静默变空。
        /// PD24: the stream decoder honours the abbreviated /DP key but not its sibling /F, so content
        /// written as "/F /Fl" is handed back as unfiltered raw bytes and the page's text silently empties.
        /// </summary>
        private static void RunAbbreviatedFilterKeyTest()
        {
            string content = "BT /F1 12 Tf 20 700 Td (Hi) Tj ET";
            byte[] raw = Encoding.ASCII.GetBytes(content);
            byte[] packed = Zlib.Compress(raw);

            // 用例一：/Filter 写成缩写 /F，压缩过的正文照常可读。
            // Case one: with /Filter written as the abbreviated /F, the compressed body still has to read.
            string text = PdfDocument.Load(BuildFilterKeyPdf("/F /Fl")).ExtractText();
            Require(text != null && text.IndexOf("Hi", StringComparison.Ordinal) >= 0,
                "a content stream filtered with the abbreviated /F key decoded as raw bytes, got '" + text + "'");

            // 控制用例：加密转换的流字典里 /F 是一个标志整数，不能当成过滤器名认。
            // Control: in a crypt-converted stream dictionary /F is a flags integer and must not be read as
            // a filter name.
            string flagsText = PdfDocument.Load(BuildFilterKeyPdf("/F 3")).ExtractText();
            Require(flagsText != null && flagsText.IndexOf("Hi", StringComparison.Ordinal) >= 0,
                "a plain stream whose /F holds a flags integer lost its text, got '" + flagsText + "'");
        }

        /// <summary>
        /// 拼一页正文流，流字典里过滤器那一项由参数给定。Assembles a one-page content stream whose filter entry comes from the argument.
        /// </summary>
        /// <param name="filterEntry">流字典里的过滤器语句，如 <c>/F /Fl</c>。The filter clause, such as "/F /Fl".</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildFilterKeyPdf(string filterEntry)
        {
            string content = "BT /F1 12 Tf 20 700 Td (Hi) Tj ET";
            byte[] packed = Zlib.Compress(Encoding.ASCII.GetBytes(content));
            bool compressed = filterEntry.IndexOf("/Fl", StringComparison.Ordinal) >= 0;
            byte[] body = compressed ? packed : Encoding.ASCII.GetBytes(content);
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R "
                    + "/Resources << /Font << /F1 5 0 R >> >> >>",
                "<< " + filterEntry + " /Length " + body.Length + " >>\nstream\n",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            return AssembleBinaryPdf(bodies, 3, body, "\nendstream");
        }

        /// <summary>
        /// PD25：名称的 <c>#HH</c> 转义在两端不成对——写出端把非 ASCII 按 UTF-8 展开，读回端一个字节还原一个字符，
        /// 载进来再存一次盘，名称逐次变形；代理对更惨，两半各退化成一个 <c>?</c>。
        /// PD25: the #HH escaping of names is unpaired across the two ends — the writer expands non-ASCII as
        /// UTF-8 while the reader turns every byte back into one character, so a load-and-save cycle rewrites
        /// the name each time; a surrogate pair is worse, its halves each collapse into a "?".
        /// </summary>
        private static void RunNameEscapeRoundTripTest()
        {
            PdfFile file = PdfFile.Load(BuildCatalogNameFixture());
            file.Catalog.SetName("Label", "中A文");
            byte[] saved = file.SaveToBytes();
            string written = Encoding.ASCII.GetString(saved);
            Require(written.IndexOf("/Label /#e4#b8#adA#e6#96#87", StringComparison.Ordinal) >= 0,
                "PD25 control failed: the writer no longer expands a name as UTF-8, file holds '"
                    + NameClause(written, "/Label") + "'");

            PdfName label = PdfFile.Load(saved).Catalog.Get("Label") as PdfName;
            Require(label != null && label.Value == "中A文",
                "a name written as /#e4#b8#adA#e6#96#87 read back as '"
                    + (label == null ? "<missing>" : label.Value) + "' instead of 中A文");

            // 成对代理必须按整个码位展开；一半一半地编码会得到两个问号。
            // A surrogate pair has to expand as the whole code point; encoding each half alone yields a pair
            // of question marks.
            PdfFile starFile = PdfFile.Load(BuildCatalogNameFixture());
            starFile.Catalog.SetName("Star", "★😀x");
            byte[] starSaved = starFile.SaveToBytes();
            string starWritten = Encoding.ASCII.GetString(starSaved);
            Require(starWritten.IndexOf("/Star /#e2#98#85#f0#9f#98#80x", StringComparison.Ordinal) >= 0,
                "the astral character did not survive as its four UTF-8 bytes, file holds '"
                    + NameClause(starWritten, "/Star") + "'");
            PdfName star = PdfFile.Load(starSaved).Catalog.Get("Star") as PdfName;
            Require(star != null && star.Value == "★😀x",
                "the astral name read back as '" + (star == null ? "<missing>" : star.Value) + "' instead of ★😀x");

            // 控制用例：单个高位字节不是合法的 UTF-8 起始串，只能按历史口径逐字节还原，
            // 不能凭空造出替换符——旧稿的名称本来就是这个样子。
            // Control: a lone high byte is not a valid UTF-8 sequence, so it must still come back byte per
            // char the legacy way instead of being invented as a replacement character.
            byte[] legacy = AssembleClassicPdf(new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R /Marked /A#80B >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
            }, 1, null);
            PdfName legacyName = PdfFile.Load(legacy).Catalog.Get("Marked") as PdfName;
            Require(legacyName != null && legacyName.Value.Length == 3
                && legacyName.Value[0] == 0x41 && legacyName.Value[1] == 0x80 && legacyName.Value[2] == 0x42,
                "a legacy one-byte name came back as '"
                    + (legacyName == null ? "<missing>" : EscapeForMessage(legacyName.Value))
                    + "' instead of A\\u0080B byte for byte");
        }

        /// <summary>
        /// PD26：签名域的 <c>/V</c> 是一份 <c>/Sig</c> 字典，<see cref="PdfFile.SetFormFieldValue"/> 与
        /// <see cref="PdfFile.FillForm(IDictionary{string,string})"/> 却按普通文本域直接覆写它。
        /// PD26: a signature field's /V is a /Sig dictionary, yet SetFormFieldValue and FillForm overwrite it
        /// the way they would a text field.
        /// </summary>
        private static void RunSignatureFieldGuardTest()
        {
            PdfFile file = PdfFile.Load(BuildWatermarkSource("[0 0 612 792]", 0));
            PdfReference page = file.GetPageReference(0);
            PdfArray fields = new PdfArray();
            PdfArray annots = new PdfArray();
            fields.Add(AddWidget(file, page, "form.name", "Tx", 0, Rect(72, 600, 260, 616), annots));

            // 签名域照真稿的形状搭：/V 指向一份 /Sig 字典，域上另挂 /ByteRange 与 /Contents。
            // The signature field is built the way a real one looks: /V points at a /Sig dictionary while the
            // field itself carries /ByteRange and /Contents.
            PdfDictionary signature = new PdfDictionary();
            signature.SetName("Type", "Sig");
            signature.SetString("Name", "YF");
            signature.SetString("M", "D:20261105090000+08'00'");
            int signatureNumber = file.AddObject(signature);

            PdfDictionary sigField = new PdfDictionary();
            sigField.SetName("Type", "Annot");
            sigField.SetName("Subtype", "Widget");
            sigField.SetName("FT", "Sig");
            sigField.SetString("T", "form.sig");
            sigField.Set("V", new PdfReference(signatureNumber, 0));
            sigField.Set("ByteRange", new PdfArray().AddNumber(0).AddNumber(8192).AddNumber(9216).AddNumber(1024));
            sigField.Set("Rect", Rect(72, 520, 260, 560));
            sigField.Set("P", page);
            int sigNumber = file.AddObject(sigField);
            fields.Add(new PdfReference(sigNumber, 0));
            annots.Add(new PdfReference(sigNumber, 0));

            PdfDictionary acroForm = new PdfDictionary();
            acroForm.Set("Fields", fields);
            acroForm.SetString("DA", "/Helv 12 Tf 0 g");
            file.GetPageDictionary(0).Set("Annots", annots);
            file.MarkModified(page.Number);
            file.Catalog.Set("AcroForm", file.AddObjectReference(acroForm));
            file.MarkModified(file.CatalogNumber);

            PdfFormField sig = file.GetForm().GetField("form.sig");
            Require(sig != null && sig.Kind == PdfFormFieldKind.Signature,
                "the fixture's /Sig field is not seen as a signature field");

            // 控制用例：文本域照常填得进去，签名域在模型属性上本来就是拒绝的。
            // Control: a text field still fills, and the model property already refuses signature fields.
            Require(file.FillForm(new Dictionary<string, string> { { "form.name", "Zhang" } }, false) == 1,
                "the plain text field next to the signature must still take a value");
            Require(file.GetForm().GetField("form.name").Value == "Zhang",
                "the text field reads back '" + file.GetForm().GetField("form.name").Value + "' instead of Zhang");
            bool modelRefused = false;
            try { sig.Value = "hacked"; }
            catch (InvalidOperationException) { modelRefused = true; }
            Require(modelRefused, "PdfFormField.Value accepted a write on a signature field");

            bool fileRefused = false;
            try { file.SetFormFieldValue("form.sig", "hacked"); }
            catch (InvalidOperationException) { fileRefused = true; }
            Require(fileRefused,
                "PdfFile.SetFormFieldValue wrote straight through to a /Sig field instead of refusing the way "
                    + "PdfFormField.Value does: the library promises signature fields are never filled");

            bool bulkRefused = false;
            try { file.FillForm(new Dictionary<string, string> { { "form.sig", "hacked" } }, false); }
            catch (InvalidOperationException) { bulkRefused = true; }
            Require(bulkRefused,
                "PdfFile.FillForm counted a /Sig field among the fields it filled instead of refusing it");

            // 可填写清单的说明自己写着「不含父域与签名/按钮」，枚举里却留着签名域：照着说明批量填的调用者
            // 会得到一份「保证填得进去」的名单，名单上第一条就撞进拒绝写入。
            // FillableFields documents itself as excluding parents, signature and push-button fields yet still
            // yields the signature field, so a caller that fills exactly the list it promises throws on it.
            bool listed = false;
            foreach (PdfFormField item in file.GetForm().FillableFields)
            {
                if (item.Kind == PdfFormFieldKind.Signature) listed = true;
            }
            Require(!listed,
                "PdfForm.FillableFields yields a signature field although its own summary says it excludes them");

            PdfFile back = PdfFile.Load(file.SaveToBytes());
            PdfDictionary saved = back.GetObject<PdfDictionary>(sig.ObjectNumber);
            PdfObject rawValue = saved == null ? null : saved.Get("V");
            PdfDictionary savedSignature = rawValue == null ? null : back.Resolve(rawValue) as PdfDictionary;
            PdfName savedType = savedSignature == null ? null : savedSignature.Get("Type") as PdfName;
            Require(savedType != null && savedType.Value == "Sig",
                "the signature's /V no longer resolves to a /Sig dictionary: it now reads '"
                    + (savedType == null ? Describe(rawValue) : savedType.Value) + "'");
        }

        /// <summary>
        /// PD27：<c>/DA</c> 是表单元里登记全文档默认外观的那一层，读端只沿父域回溯，走到顶就停。
        /// PD27: /DA also lives on the AcroForm dictionary as the document-wide default, yet the reader walks
        /// only up the parent chain and stops there.
        /// </summary>
        private static void RunFormDefaultAppearanceInheritanceTest()
        {
            PdfFile file = PdfFile.Load(BuildWatermarkSource("[0 0 612 792]", 0));
            PdfReference page = file.GetPageReference(0);

            // 表单级默认外观写 14 号纯蓝；一个控件不带自己的 /DA，另一个带 9 号纯红。
            // The form-level default says 14pt pure blue; one widget carries no /DA of its own, the other 9pt red.
            PdfDictionary acroForm = new PdfDictionary();
            acroForm.SetString("DA", "/Helv 14 Tf 0 0 1 rg");
            PdfArray fields = new PdfArray();
            PdfArray annots = new PdfArray();
            fields.Add(AddFormWidget(file, page, "da.inherited", null, annots));
            fields.Add(AddFormWidget(file, page, "da.own", "/Helv 9 Tf 1 0 0 rg", annots));
            acroForm.Set("Fields", fields);
            file.GetPageDictionary(0).Set("Annots", annots);
            file.MarkModified(page.Number);
            file.Catalog.Set("AcroForm", file.AddObjectReference(acroForm));
            file.MarkModified(file.CatalogNumber);

            PdfForm form = file.GetForm();
            Require(form.GetField("da.inherited") != null && form.GetField("da.own") != null,
                "both text fields must be read back from the fixture");
            Require(form.GetField("da.inherited").DefaultAppearance == "/Helv 14 Tf 0 0 1 rg",
                "the field without its own /DA reads the default appearance as '"
                    + form.GetField("da.inherited").DefaultAppearance
                    + "' instead of the AcroForm document default /Helv 14 Tf 0 0 1 rg");
            Require(form.GetField("da.own").DefaultAppearance == "/Helv 9 Tf 1 0 0 rg",
                "a field's own /DA must still win over the form-level default, got '"
                    + form.GetField("da.own").DefaultAppearance + "'");

            file.SetFormFieldValue("da.inherited", "Hi");
            file.SetFormFieldValue("da.own", "Hi");
            file.GenerateFormAppearances();

            string inherited = AppearanceBodyOf(file, "da.inherited");
            Require(inherited.IndexOf(" 14 Tf", StringComparison.Ordinal) >= 0,
                "the inherited size is missing from the generated appearance: " + Shorten(inherited));
            Require(inherited.IndexOf("0 0 1 rg", StringComparison.Ordinal) >= 0,
                "the inherited colour is missing from the generated appearance: " + Shorten(inherited));

            string own = AppearanceBodyOf(file, "da.own");
            Require(own.IndexOf(" 9 Tf", StringComparison.Ordinal) >= 0 && own.IndexOf("1 0 0 rg", StringComparison.Ordinal) >= 0,
                "the field's own size and colour are missing from its appearance: " + Shorten(own));
        }

        /// <summary>
        /// 新建一个文本控件，<c>/DA</c> 可为 null 以模拟「只靠表单级默认外观」的第三方稿。
        /// Creates a text widget whose /DA may be null, imitating a third-party form that relies on the
        /// document-wide default appearance.
        /// </summary>
        /// <param name="file">目标文件。The target file.</param>
        /// <param name="page">所在页引用。The page reference.</param>
        /// <param name="name">字段名。The field name.</param>
        /// <param name="da">本域默认外观串，null 表示不写。The field's own default appearance; null writes nothing.</param>
        /// <param name="annots">页面注释数组。The page annotation array.</param>
        /// <returns>新字段的引用。The reference of the new field.</returns>
        private static PdfReference AddFormWidget(PdfFile file, PdfReference page, string name, string da, PdfArray annots)
        {
            PdfDictionary field = new PdfDictionary();
            field.SetName("Type", "Annot");
            field.SetName("Subtype", "Widget");
            field.SetName("FT", "Tx");
            field.SetString("T", name);
            field.Set("Rect", Rect(72, 600, 260, 616));
            field.Set("P", page);
            if (da != null) field.SetString("DA", da);
            PdfReference reference = file.AddObjectReference(field);
            annots.Add(reference);
            return reference;
        }

        /// <summary>取一个字段当前 <c>/AP /N</c> 正文的 ASCII 文本。Reads the ASCII body of a field's current /AP /N stream.</summary>
        /// <param name="file">目标文件。The target file.</param>
        /// <param name="name">字段名。The field name.</param>
        /// <returns>正文文本，缺外观时给占位串。The body, or a placeholder when no appearance exists.</returns>
        private static string AppearanceBodyOf(PdfFile file, string name)
        {
            PdfFormField field = file.GetForm().GetField(name);
            PdfDictionary dictionary = field == null ? null : file.GetObject<PdfDictionary>(field.ObjectNumber);
            PdfObject apObject = dictionary == null ? null : dictionary.Get("AP");
            PdfDictionary appearance = apObject as PdfDictionary;
            if (appearance == null && apObject != null) appearance = file.Resolve(apObject) as PdfDictionary;
            PdfObject normal = appearance == null ? null : appearance.Get("N");
            PdfStream stream = normal as PdfStream;
            if (stream == null && normal != null) stream = file.Resolve(normal) as PdfStream;
            return stream == null ? "<no appearance>" : Encoding.ASCII.GetString(stream.Data);
        }

        /// <summary>把外观正文压成一行、截到 160 字，供失败信息引用。Flattens an appearance body to one line, cut to 160 chars, for failure messages.</summary>
        /// <param name="body">外观正文。The appearance body.</param>
        /// <returns>单行摘录。The single-line excerpt.</returns>
        private static string Shorten(string body)
        {
            string one = body.Replace("\n", " ");
            if (one.Length > 160) one = one.Substring(0, 160) + "...";
            return one;
        }

        /// <summary>
        /// PD28：<c>/DA</c> 用 CMYK 上色算子 <c>k</c>／<c>c</c> 时读端解不出颜色，生成外观一律画成黑字。
        /// PD28: when /DA sets the colour with the CMYK operators k or c the reader finds no colour at all, so a
        /// generated appearance paints black whatever the form asked for.
        /// </summary>
        private static void RunDefaultAppearanceCmykTest()
        {
            PdfFile file = PdfFile.Load(BuildWatermarkSource("[0 0 612 792]", 0));
            PdfReference page = file.GetPageReference(0);

            // 印刷向的表单常把 /DA 写成 CMYK：0 0.75 0.75 0.25 k 是一块砖红，对应 RGB 约 191/48/48。
            // Print-oriented forms routinely write /DA in CMYK: 0 0.75 0.75 0.25 k is a brick red, about 191/48/48.
            PdfDictionary acroForm = new PdfDictionary();
            PdfArray fields = new PdfArray();
            PdfArray annots = new PdfArray();
            fields.Add(AddFormWidget(file, page, "da.cmyk", "/Helv 12 Tf 0 0.75 0.75 0.25 k", annots));
            fields.Add(AddFormWidget(file, page, "da.rgb", "/Helv 12 Tf 0.2 0.6 0.9 rg", annots));
            acroForm.Set("Fields", fields);
            file.GetPageDictionary(0).Set("Annots", annots);
            file.MarkModified(page.Number);
            file.Catalog.Set("AcroForm", file.AddObjectReference(acroForm));
            file.MarkModified(file.CatalogNumber);

            file.SetFormFieldValue("da.cmyk", "Hi");
            file.SetFormFieldValue("da.rgb", "Hi");
            file.GenerateFormAppearances();

            // 控制用例：DeviceRGB 的 /DA 一直是解得开的，颜色应当原样落进外观。
            // Control: a DeviceRGB /DA always parsed, so its colour must reach the appearance unchanged.
            double[] rgb = FirstFillTriple(AppearanceBodyOf(file, "da.rgb"));
            Require(rgb != null && Near(rgb[0], 0.2) && Near(rgb[1], 0.6) && Near(rgb[2], 0.9),
                "the DeviceRGB control reads " + DescribeTriple(rgb) + " instead of 0.2 0.6 0.9");

            double[] cmyk = FirstFillTriple(AppearanceBodyOf(file, "da.cmyk"));
            Require(cmyk != null && Near(cmyk[0], 191.0 / 255.0) && Near(cmyk[1], 48.0 / 255.0) && Near(cmyk[2], 48.0 / 255.0),
                "a CMYK default appearance of 0 0.75 0.75 0.25 k painted " + DescribeTriple(cmyk)
                    + " instead of the brick red it describes");
        }

        /// <summary>
        /// PD29：<c>/Info</c> 日期串的时区标注（<c>Z</c> 与 <c>±hh'mm'</c>）在两端都不参与换算——写出端恒按本地
        /// 时间加本地偏移，读回端却把标注整段丢掉，于是别人按 UTC 记的稿子读回来墙上时间原样不动，比真实时刻差几个小时。
        /// PD29: the time-zone designator of an /Info date string — <c>Z</c> or <c>±hh'mm'</c> — takes part in
        /// neither end. The writer always stamps local time plus the local offset, while the reader drops the
        /// designator, so a document stamped in UTC comes back with the wall clock unchanged, hours away from
        /// the instant it records.
        /// </summary>
        private static void RunPdfDateDesignatorTest()
        {
            // 控制用例：没有任何时区标注时规范就按本地时间理解，墙上时间一读一写都不许动。
            // Control: with no designator the spec means local time, so the wall clock must stay put.
            DateTime plain = ReadCreationDate("D:20260115120000");
            Require(plain.Ticks == new DateTime(2026, 1, 15, 12, 0, 0).Ticks,
                "a date without a designator reads " + plain.ToString("O", CultureInfo.InvariantCulture)
                    + " instead of 2026-01-15 12:00:00 local");

            DateTime utc = ReadCreationDate("D:20260115120000Z");
            DateTime expectedUtc = ToLocalWallTime(new DateTime(2026, 1, 15, 12, 0, 0, DateTimeKind.Utc));
            Require(utc.Ticks == expectedUtc.Ticks,
                "the UTC designator is dropped: D:20260115120000Z reads "
                    + utc.ToString("O", CultureInfo.InvariantCulture) + ", the instant it records is "
                    + expectedUtc.ToString("O", CultureInfo.InvariantCulture));

            DateTime zero = ReadCreationDate("D:20260115120000+00'00'");
            Require(zero.Ticks == expectedUtc.Ticks,
                "an explicit +00'00' offset reads " + zero.ToString("O", CultureInfo.InvariantCulture));

            // +05'30'：12:00 是该偏移下的墙上时间，落到本地就是 06:30 UTC 那一刻。
            // 12:00 is the wall clock at +05:30, so the instant is 06:30 UTC.
            DateTime half = ReadCreationDate("D:20260115120000+05'30'");
            DateTime expectedHalf = ToLocalWallTime(new DateTime(2026, 1, 15, 6, 30, 0, DateTimeKind.Utc));
            Require(half.Ticks == expectedHalf.Ticks,
                "a +05'30' date reads " + half.ToString("O", CultureInfo.InvariantCulture) + " instead of "
                    + expectedHalf.ToString("O", CultureInfo.InvariantCulture));

            DateTime minus = ReadCreationDate("D:20260115120000-08'00'");
            DateTime expectedMinus = ToLocalWallTime(new DateTime(2026, 1, 15, 20, 0, 0, DateTimeKind.Utc));
            Require(minus.Ticks == expectedMinus.Ticks,
                "a -08'00' date reads " + minus.ToString("O", CultureInfo.InvariantCulture) + " instead of "
                    + expectedMinus.ToString("O", CultureInfo.InvariantCulture));

            // 只有年份的简写仍然要能用，且不带任何时分秒。
            // The year-only shorthand still has to work, with the clock left at zero.
            DateTime year = ReadCreationDate("D:2026");
            Require(year.Ticks == new DateTime(2026, 1, 1, 0, 0, 0).Ticks,
                "D:2026 reads " + year.ToString("O", CultureInfo.InvariantCulture));
        }

        /// <summary>把一份只带 <c>/Info</c> 日期的 PDF 交给两条读取路径，要求读出的创建时间一致。Feeds a PDF whose /Info carries only a date to both readers and requires the same creation date back.</summary>
        /// <param name="literal">日期串正文。The date literal.</param>
        /// <returns>读出的创建时间。The creation date as read.</returns>
        private static DateTime ReadCreationDate(string literal)
        {
            byte[] pdf = BuildInfoDatePdf(literal);
            DateTime? fromFile = PdfFile.Load(pdf).Info.CreationDate;
            DateTime? fromDocument = PdfDocument.Load(pdf).Info.CreationDate;
            Require(fromFile.HasValue, "PdfFile read no creation date from " + literal);
            Require(fromDocument.HasValue, "PdfDocument read no creation date from " + literal);
            Require(fromFile.Value.Ticks == fromDocument.Value.Ticks,
                "the two readers disagree on " + literal + ": " + fromFile.Value.ToString("O", CultureInfo.InvariantCulture)
                    + " against " + fromDocument.Value.ToString("O", CultureInfo.InvariantCulture));
            return fromFile.Value;
        }

        /// <summary>把 UTC 时刻换成本地墙上时间（Kind 抹平，便于逐刻比较）。Converts a UTC instant into the local wall clock with Kind stripped so the ticks compare.</summary>
        /// <param name="instant">带 UTC 标记的时刻。The UTC-marked instant.</param>
        /// <returns>本地墙上时间。The local wall time.</returns>
        private static DateTime ToLocalWallTime(DateTime instant)
        {
            return DateTime.SpecifyKind(TimeZoneInfo.ConvertTimeFromUtc(instant, TimeZoneInfo.Local), DateTimeKind.Unspecified);
        }

        /// <summary>拼一份单页 PDF，<c>/Info</c> 里只有那一条创建与修改日期。Assembles a one-page PDF whose /Info holds only the given creation and modification date.</summary>
        /// <param name="literal">日期串正文。The date literal.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildInfoDatePdf(string literal)
        {
            return AssembleClassicPdf(new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
                "<< /CreationDate (" + literal + ") /ModDate (" + literal + ") >>",
            }, 1, null, "/Info 4 0 R ");
        }

        /// <summary>取外观正文里第一条 <c>rg</c> 填充色的三个分量。Reads the three channels of the first rg fill operator in an appearance body.</summary>
        /// <param name="body">外观正文。The appearance body.</param>
        /// <returns>三个 0~1 的分量，没有 <c>rg</c> 时返回 null。The three 0-1 channels, null when no rg operator exists.</returns>
        private static double[] FirstFillTriple(string body)
        {
            if (body == null) return null;
            int at = body.IndexOf(" rg", StringComparison.Ordinal);
            if (at < 0) return null;
            int start = at;
            while (start > 0 && (body[start - 1] == '.' || body[start - 1] == ' ' || (body[start - 1] >= '0' && body[start - 1] <= '9'))) start--;
            string[] tokens = body.Substring(start, at - start).Trim()
                .Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length < 3) return null;
            double[] result = new double[3];
            for (int i = 0; i < 3; i++)
            {
                double value;
                if (!double.TryParse(tokens[tokens.Length - 3 + i], NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return null;
                result[i] = value;
            }
            return result;
        }

        /// <summary>判断 0~1 的分量是否落在容差内。Whether a 0-1 channel sits within tolerance of the expected value.</summary>
        /// <param name="value">实际分量。The actual channel.</param>
        /// <param name="expected">期望分量。The expected channel.</param>
        /// <returns>相差不到 0.01 时 true。True when they differ by less than 0.01.</returns>
        private static bool Near(double value, double expected)
        {
            return Math.Abs(value - expected) < 0.01;
        }

        /// <summary>把一组分量写成可读文本。Writes a channel triple as readable text.</summary>
        /// <param name="value">分量，可为 null。The channels, possibly null.</param>
        /// <returns>可读描述。A printable description.</returns>
        private static string DescribeTriple(double[] value)
        {
            if (value == null) return "<no rg operator>";
            return value[0].ToString("0.###", CultureInfo.InvariantCulture) + " "
                + value[1].ToString("0.###", CultureInfo.InvariantCulture) + " "
                + value[2].ToString("0.###", CultureInfo.InvariantCulture);
        }

        /// <summary>拼一份目录里只挂着页面树的夹具，供名称读写。Assembles a fixture whose catalog only carries the page tree, for name round-trips.</summary>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildCatalogNameFixture()
        {
            return AssembleClassicPdf(new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
            }, 1, null);
        }

        /// <summary>取出文件里某个名称键后面那一句，便于失败信息点名。Pulls the clause written after one name key out of the file for the failure message.</summary>
        /// <param name="fileText">保存结果的 ASCII 文本。The ASCII text of the saved file.</param>
        /// <param name="key">带斜杠的名称键。The name key with its slash.</param>
        /// <returns>该句文本，找不到时给占位串。The clause, or a placeholder when absent.</returns>
        private static string NameClause(string fileText, string key)
        {
            int at = fileText.IndexOf(key, StringComparison.Ordinal);
            if (at < 0) return "<" + key + " missing>";
            int end = at;
            while (end < fileText.Length && fileText[end] != '\n' && fileText[end] != '>') end++;
            return fileText.Substring(at, end - at);
        }

        /// <summary>把名称里的不可见字符显形。Makes the invisible characters of a name visible.</summary>
        /// <param name="value">名称正文。The name body.</param>
        /// <returns>可打印的描述。A printable description.</returns>
        private static string EscapeForMessage(string value)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] >= 32 && value[i] < 127) builder.Append(value[i]);
                else builder.Append("\\u").Append(((int)value[i]).ToString("x4", CultureInfo.InvariantCulture));
            }
            return builder.ToString();
        }

        /// <summary>
        /// 拼一份三页文档：资源字典只挂在 <c>/Pages</c> 上，三页都靠继承取用，这是第三方 PDF 的常见写法。
        /// Assembles a three-page document whose resource dictionary lives only on /Pages, so all three
        /// pages read it by inheritance — a layout real-world generators actually emit.
        /// </summary>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] BuildInheritedResourcePdf()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R 4 0 R 5 0 R] /Count 3 "
                    + "/Resources << /Font << /F1 6 0 R >> >> >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
                "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            };
            return AssembleClassicPdf(bodies, 1, null);
        }

        /// <summary>造一枚最简表单 XObject，供两页各自登记。Builds a minimal form XObject for the two pages to register.</summary>
        /// <returns>表单字典。The form dictionary.</returns>
        private static PdfDictionary BuildLogoForm()
        {
            PdfDictionary form = new PdfDictionary();
            form.SetName("Type", "XObject");
            form.SetName("Subtype", "Form");
            form.Set("BBox", new PdfArray().AddNumber(0).AddNumber(0).AddNumber(40).AddNumber(40));
            return form;
        }

        /// <summary>取某页 <c>/XObject</c> 分组，取不到返回 null。Returns one page's /XObject group, or null when it has none.</summary>
        /// <param name="file">待检查的文档。The document to inspect.</param>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>分组字典或 null。The group dictionary, or null.</returns>
        private static PdfDictionary XObjectGroup(PdfFile file, int index)
        {
            PdfDictionary resources = file.GetPageResources(index);
            return resources == null ? null : resources.Get("XObject") as PdfDictionary;
        }

        /// <summary>取某页 <c>/Font /F1</c> 指向的对象号，取不到返回 -1。Returns the object number behind one page's /Font /F1, or -1.</summary>
        /// <param name="file">待检查的文档。The document to inspect.</param>
        /// <param name="index">页面下标。The page index.</param>
        /// <returns>对象号。The object number.</returns>
        private static int FontNumber(PdfFile file, int index)
        {
            PdfDictionary resources = file.GetPageResources(index);
            PdfDictionary fonts = resources == null ? null : resources.Get("Font") as PdfDictionary;
            PdfObject value = fonts == null ? null : fonts.Get("F1");
            return ReferenceNumber(value);
        }

        /// <summary>引用对象号，非引用返回 -1。The object number of a reference, or -1 for any other object.</summary>
        /// <param name="value">待判定的对象。The object to inspect.</param>
        /// <returns>对象号。The object number.</returns>
        private static int ReferenceNumber(PdfObject value)
        {
            PdfReference reference = value as PdfReference;
            return reference == null ? -1 : reference.Number;
        }

        /// <summary>把资源值写成可读文本，引用直接给「n g R」，便于失败信息点名。
        /// Renders a resource value readably — references print as "n g R" — so the failure message names the culprit.</summary>
        /// <param name="value">待描述的对象。The object to describe.</param>
        /// <returns>描述串。The description.</returns>
        private static string ReferenceText(PdfObject value)
        {
            PdfReference reference = value as PdfReference;
            return reference == null ? Describe(value) : reference.ToString();
        }

        /// <summary>把末尾 startxref 的数字换成一个落在文件内的错号。Replaces the trailing startxref number with an in-range wrong offset.</summary>
        /// <param name="pdf">完整的 PDF 字节。The complete PDF bytes.</param>
        /// <param name="bogusOffset">要写进去的错误偏移。The forged offset to write.</param>
        /// <returns>改坏后的 PDF 字节。The damaged PDF bytes.</returns>
        private static byte[] BreakStartxref(byte[] pdf, int bogusOffset)
        {
            string text = Encoding.ASCII.GetString(pdf);
            int at = text.LastIndexOf("startxref", StringComparison.Ordinal);
            Require(at >= 0, "PD8 fixture has no startxref to break");
            int digits = at + "startxref".Length;
            while (digits < text.Length && (text[digits] == '\n' || text[digits] == '\r' || text[digits] == ' '))
            {
                digits++;
            }
            int end = digits;
            while (end < text.Length && text[end] >= '0' && text[end] <= '9') end++;
            string rebuilt = text.Substring(0, digits) + bogusOffset.ToString() + text.Substring(end);
            return Encoding.ASCII.GetBytes(rebuilt);
        }

        /// <summary>
        /// 拼装允许二进制流正文的经典交叉表 PDF：第 streamIndex 个对象的正文以原字节给出。
        /// Assembles a classic-xref PDF carrying a binary stream body: the object at streamIndex contributes
        /// its raw bytes instead of text.
        /// </summary>
        /// <param name="bodies">各对象的正文，流对象的正文以「字典 + stream 起始」结尾。The object bodies; the stream object's ends with its dictionary and the stream keyword.</param>
        /// <param name="streamIndex">流对象的下标。The index of the stream object.</param>
        /// <param name="streamBytes">流的原字节。The raw stream bytes.</param>
        /// <param name="streamTail">紧跟正文的收尾串。The terminator written right after the body.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] AssembleBinaryPdf(string[] bodies, int streamIndex, byte[] streamBytes, string streamTail)
        {
            Encoding ascii = Encoding.ASCII;
            List<byte[]> pieces = new List<byte[]>();
            List<long> offsets = new List<long>();
            long position = 0;

            Action<string> writeText = delegate(string text)
            {
                byte[] bytes = ascii.GetBytes(text);
                pieces.Add(bytes);
                position += bytes.Length;
            };

            writeText("%PDF-1.7\n");
            for (int i = 0; i < bodies.Length; i++)
            {
                offsets.Add(position);
                writeText((i + 1) + " 0 obj\n");
                writeText(bodies[i]);
                if (i == streamIndex)
                {
                    byte[] raw = streamBytes;
                    pieces.Add(raw);
                    position += raw.Length;
                    writeText(streamTail);
                }
                writeText("\nendobj\n");
            }

            long xrefAt = position;
            writeText("xref\n0 " + (bodies.Length + 1) + "\n0000000000 65535 f \n");
            for (int i = 0; i < bodies.Length; i++)
            {
                writeText(offsets[i].ToString("D10") + " 00000 n \n");
            }
            writeText("trailer\n<< /Size " + (bodies.Length + 1) + " /Root 1 0 R >>\nstartxref\n"
                + xrefAt + "\n%%EOF\n");

            using (MemoryStream memory = new MemoryStream())
            {
                for (int i = 0; i < pieces.Count; i++)
                {
                    memory.Write(pieces[i], 0, pieces[i].Length);
                }
                return memory.ToArray();
            }
        }

        /// <summary>
        /// PD10：目录对象住在非 0 世代时，保存写出的 trailer <c>/Root</c> 必须带上同一个世代，
        /// 不能只留对象号、把世代硬写成 0。
        /// PD10: when the catalog lives in a non-zero generation, the trailer /Root written on save must
        /// carry that same generation instead of hard-coding 0 next to the object number.
        /// </summary>
        private static void RunCatalogGenerationTest()
        {
            string[] bodies = new string[]
            {
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] >>",
            };
            int[] generations = new int[] { 2, 0, 0 };
            byte[] source = AssembleGenerationsPdf(bodies, generations, 1, 2);

            byte[] saved = PdfFile.Load(source).SaveToBytes();
            string text = Encoding.ASCII.GetString(saved);

            Require(text.Contains("1 2 obj"),
                "PD10 control failed: the saved file no longer writes the catalog header as '1 2 obj'");

            PdfReference root = PdfFile.Load(saved).Trailer.Get("Root") as PdfReference;
            Require(root != null, "PD10 control failed: the saved trailer carries no /Root reference");
            Require(root.Number == 1 && root.Generation == 2,
                "the trailer /Root is '" + root.Number + " " + root.Generation
                    + " R' while the object header is '1 2 obj'");
        }

        /// <summary>
        /// 拼装逐对象带世代号的经典交叉表 PDF，交叉表与 trailer 使用给定世代。
        /// Assembles a classic-xref PDF with a per-object generation; the cross-reference rows and the
        /// trailer /Root reference use the supplied generations.
        /// </summary>
        /// <param name="bodies">各对象正文，下标零起。The object bodies, zero-based.</param>
        /// <param name="generations">与正文逐位对应的世代号。The generation paired with each body.</param>
        /// <param name="rootNumber">trailer /Root 的对象号。The trailer /Root object number.</param>
        /// <param name="rootGeneration">trailer /Root 的世代号。The trailer /Root generation.</param>
        /// <returns>PDF 字节。The PDF bytes.</returns>
        private static byte[] AssembleGenerationsPdf(string[] bodies, int[] generations, int rootNumber, int rootGeneration)
        {
            Encoding ascii = Encoding.ASCII;
            List<byte[]> pieces = new List<byte[]>();
            List<long> offsets = new List<long>();
            long position = 0;

            Action<string> writeText = delegate(string text)
            {
                byte[] bytes = ascii.GetBytes(text);
                pieces.Add(bytes);
                position += bytes.Length;
            };

            writeText("%PDF-1.7\n");
            for (int i = 0; i < bodies.Length; i++)
            {
                offsets.Add(position);
                writeText((i + 1) + " " + generations[i] + " obj\n");
                writeText(bodies[i]);
                writeText("\nendobj\n");
            }

            long xrefAt = position;
            writeText("xref\n0 " + (bodies.Length + 1) + "\n0000000000 65535 f \n");
            for (int i = 0; i < bodies.Length; i++)
            {
                writeText(offsets[i].ToString("D10") + " " + generations[i].ToString("D5") + " n \n");
            }
            writeText("trailer\n<< /Size " + (bodies.Length + 1)
                + " /Root " + rootNumber + " " + rootGeneration + " R >>\nstartxref\n"
                + xrefAt + "\n%%EOF\n");

            using (MemoryStream memory = new MemoryStream())
            {
                for (int i = 0; i < pieces.Count; i++)
                {
                    memory.Write(pieces[i], 0, pieces[i].Length);
                }
                return memory.ToArray();
            }
        }
    }
}
