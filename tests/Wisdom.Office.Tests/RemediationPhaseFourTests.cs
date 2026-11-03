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
    /// 需要告警通道的条目（<c>PD9</c> 与其下游 <c>PD5</c>/<c>PD22</c>）另批设计，这里只放可单独判定的条目。
    /// Entries that need the warning channel (<c>PD9</c> and its downstream <c>PD5</c>/<c>PD22</c>) are
    /// designed as their own batch; this file carries only what can be decided on its own.
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
