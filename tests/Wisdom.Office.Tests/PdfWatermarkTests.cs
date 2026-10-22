using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Wisdom.Pdf;

namespace Wisdom.Office.Tests
{
    /// <summary>
    /// 文档水印（文字与图像）回归测试。Regression tests for text and image document watermarks.
    /// </summary>
    /// <remarks>覆盖背景/前景层次、平铺、指定页、页面旋转补偿、图像水印、增量保存与加密文档六条主路径；
    /// 结构断言直接读取对象级 <see cref="PdfFile"/>，文本断言走 <see cref="PdfDocument"/> 的完整提取链，
    /// 因此写入端与读取端同时受验。
    /// The suite covers the six main paths: background/foreground layer, tiling, selected pages, page
    /// rotation compensation, image watermarks, incremental save and encrypted documents. Structural
    /// assertions read the object-level <see cref="PdfFile"/> directly while text assertions run the whole
    /// <see cref="PdfDocument"/> extraction chain, so the writer and the reader are verified together.</remarks>
    internal static partial class Program
    {
        private static void RunPdfWatermarkTest(string sourcePdfPath, string outputDirectory)
        {
            byte[] source = File.ReadAllBytes(sourcePdfPath);
            int sourcePages = PdfFile.Load(source).PageCount;
            Require(sourcePages >= 2, "watermark test needs at least two source pages: " + sourcePages);

            RunWatermarkTextBackground(source, sourcePages, outputDirectory);
            RunWatermarkTextTile(source, outputDirectory);
            RunWatermarkSelectedPages(source, sourcePages, outputDirectory);
            RunWatermarkRotatedPage(source, outputDirectory);
            RunWatermarkImage(source, sourcePages, outputDirectory);
            RunWatermarkIncremental(source, sourcePages, outputDirectory);
            RunWatermarkEncrypted(source, outputDirectory);
            RunWatermarkOptionValidation(source);

            Console.WriteLine("PDF watermark tests passed.");
        }

        /// <summary>背景文字水印：默认 45 度对角线、浅灰、30% 不透明度，内容流排在原页面内容之前。Background text watermark: the default 45-degree diagonal, light gray, 30% opacity, drawn before the page content.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="sourcePages">源页数。The source page count.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWatermarkTextBackground(byte[] source, int sourcePages, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(source);
            file.AddWatermark("机密文件 CONFIDENTIAL");
            byte[] saved = file.SaveToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-watermark-text.pdf"));
            File.WriteAllBytes(path, saved);

            PdfDocument document = PdfDocument.Load(path);
            string text = document.ExtractText();
            Require(document.PageCount == sourcePages, "background watermark changed the page count");
            RequireContains(text, "机密文件");
            RequireContains(text, "CONFIDENTIAL");
            Require(SlotHasWatermark(PdfFile.Load(saved), 0, true), "background watermark is not the first content stream");
            Require(CountOccurrences(text, "机密文件") == sourcePages, "one watermark per page expected: " + CountOccurrences(text, "机密文件"));
        }

        /// <summary>前景平铺水印：自定义角度、加粗、低不透明度，内容流排在原页面内容之后。Foreground tiled watermark: custom angle, bold, low opacity, drawn after the page content.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWatermarkTextTile(byte[] source, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(source);
            WatermarkOptions options = WatermarkOptions.ForText("内部资料 INTERNAL USE");
            options.Layer = WatermarkLayer.Foreground;
            options.Placement = WatermarkPlacement.Tile;
            options.Opacity = 0.15;
            options.FontSize = 28;
            options.Bold = true;
            options.Angle = 30;
            file.AddWatermark(options);
            byte[] saved = file.SaveToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-watermark-tile.pdf"));
            File.WriteAllBytes(path, saved);

            PdfDocument document = PdfDocument.Load(path);
            string text = document.ExtractText();
            Require(CountOccurrences(text, "内部资料") >= 8, "tiling produced too few repeats: " + CountOccurrences(text, "内部资料"));
            RequireContains(text, "INTERNAL USE");
            Require(SlotHasWatermark(PdfFile.Load(saved), 0, false), "foreground watermark is not the last content stream");
        }

        /// <summary>指定页水印：只作用于给定页码，其余页面保持原样。Selected-page watermark: only the given page indexes are touched, the rest stay clean.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="sourcePages">源页数。The source page count.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWatermarkSelectedPages(byte[] source, int sourcePages, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(source);
            file.AddWatermark(WatermarkOptions.ForText("仅第一页 ONLY-FIRST"), 0);
            byte[] saved = file.SaveToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-watermark-selected.pdf"));
            File.WriteAllBytes(path, saved);

            PdfDocument document = PdfDocument.Load(path);
            Require(document.PageCount == sourcePages, "selected-page watermark changed the page count");
            RequireContains(document.Pages[0].ExtractText(), "仅第一页");
            for (int i = 1; i < document.PageCount; i++)
            {
                Require(!document.Pages[i].ExtractText().Contains("仅第一页"), "page " + i + " should not carry the watermark");
            }
        }

        /// <summary>旋转页水印：/Rotate 为 90 的页面仍按阅读方向水平居中，写出端负责补回旋转矩阵。Rotated-page watermark: a /Rotate 90 page still gets a reading-direction horizontal watermark because the writer compensates.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWatermarkRotatedPage(byte[] source, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(source);
            file.SetPageRotation(0, 90);
            WatermarkOptions options = WatermarkOptions.ForText("旋后水平 HORIZONTAL");
            options.Placement = WatermarkPlacement.Center;
            options.Angle = 0;
            file.AddWatermark(options, 0);
            byte[] saved = file.SaveToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-watermark-rotated.pdf"));
            File.WriteAllBytes(path, saved);

            PdfFile reload = PdfFile.Load(saved);
            Require(reload.GetPageRotation(0) == 90, "page rotation lost after watermark");
            string[] pair = WatermarkPair(reload, 0);
            Require(pair[0].Contains("0 1 -1 0"), "rotation compensation matrix missing in page stream");
            Require(BalancedGraphics(pair[0]) && BalancedGraphics(pair[1]), "graphics state unbalanced on a rotated page");
            RequireContains(PdfDocument.Load(path).Pages[0].ExtractText(), "旋后水平");
        }

        /// <summary>图像水印：右下角嵌入 PNG，附带 SMask 透明通道与 ExtGState 不透明度。Image watermark: a PNG placed at the bottom right with an SMask alpha channel and an ExtGState opacity.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="sourcePages">源页数。The source page count.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWatermarkImage(byte[] source, int sourcePages, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(source);
            WatermarkOptions options = WatermarkOptions.ForImage(CreateTestPng(64, 64));
            options.ImageWidth = 180;
            options.Layer = WatermarkLayer.Foreground;
            options.Placement = WatermarkPlacement.BottomRight;
            options.Opacity = 0.5;
            file.AddWatermark(options);
            byte[] saved = file.SaveToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-watermark-image.pdf"));
            File.WriteAllBytes(path, saved);

            Require(PdfDocument.Load(path).PageCount == sourcePages, "image watermark changed the page count");
            PdfFile reload = PdfFile.Load(saved);
            Require(CountImageXObjects(reload) >= 1, "image XObject not embedded");
            string[] pair = WatermarkPair(reload, 0);
            Require(pair[1].Contains("/WmImage Do"), "form does not draw the watermark image");
            Require(pair[1].Contains(" gs"), "alpha ExtGState not applied");
            Require(BalancedGraphics(pair[0]) && BalancedGraphics(pair[1]), "graphics state unbalanced on the image watermark");
        }

        /// <summary>增量水印：整量保存后再次加载并增量追加第二轮水印，两轮文字都可读。Incremental watermark: a full save is reloaded and a second pass is appended incrementally; both watermark texts stay readable.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="sourcePages">源页数。The source page count.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWatermarkIncremental(byte[] source, int sourcePages, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(source);
            file.AddWatermark("增量 INCREMENTAL");
            byte[] full = file.SaveToBytes();

            PdfFile again = PdfFile.Load(full);
            again.AddWatermark(WatermarkOptions.ForText("第二轮 SECOND-PASS"), 0);
            byte[] incremental = again.SaveIncrementalToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-watermark-incremental.pdf"));
            File.WriteAllBytes(path, incremental);

            PdfDocument document = PdfDocument.Load(path);
            string first = document.Pages[0].ExtractText();
            RequireContains(first, "增量");
            RequireContains(first, "第二轮");
            Require(document.PageCount == sourcePages, "incremental watermark changed the page count");
            Require(!document.Pages[1].ExtractText().Contains("第二轮"), "second pass leaked to another page");
        }

        /// <summary>加密文档水印：AES-128 加密后仍可追加水印并增量保存，解密读取文本正常。Encrypted document watermark: an AES-128 document still accepts a watermark and an incremental save, and the text decrypts back.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunWatermarkEncrypted(byte[] source, string outputDirectory)
        {
            PdfFile file = PdfFile.Load(source);
            PdfEncryptionOptions encryption = new PdfEncryptionOptions();
            encryption.Algorithm = PdfEncryptionAlgorithm.Aes128;
            encryption.UserPassword = "wm-user";
            encryption.OwnerPassword = "wm-owner";
            file.SetEncryption(encryption);
            byte[] encrypted = file.SaveToBytes();

            PdfFile opened = PdfFile.Load(encrypted, "wm-user");
            opened.AddWatermark("加密水印 SEALED");
            byte[] sealedBytes = opened.SaveIncrementalToBytes();
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "pdf-watermark-encrypted.pdf"));
            File.WriteAllBytes(path, sealedBytes);

            RequireContains(PdfDocument.Load(path, "wm-user").ExtractText(), "加密水印");
            RequireContains(PdfDocument.Load(path, "wm-user").ExtractText(), "SEALED");
            Require(PdfFile.Load(sealedBytes, "wm-owner").IsEncrypted, "watermark lost the encryption");
        }

        /// <summary>参数校验：空文字、空图像、非法字号/不透明度/边距都会被拒绝，越界页码则被忽略。Option validation: empty text, empty image and illegal size/opacity/margin are rejected, while an out-of-range page index is ignored.</summary>
        /// <param name="source">源 PDF 字节。The source PDF bytes.</param>
        private static void RunWatermarkOptionValidation(byte[] source)
        {
            RequireThrows(delegate()
            {
                PdfFile file = PdfFile.Load(source);
                file.AddWatermark(string.Empty);
            }, "empty watermark text should be rejected");

            RequireThrows(delegate()
            {
                PdfFile file = PdfFile.Load(source);
                file.AddWatermark(WatermarkOptions.ForImage(new byte[0]));
            }, "empty watermark image should be rejected");

            RequireThrows(delegate()
            {
                PdfFile file = PdfFile.Load(source);
                file.AddWatermark((WatermarkOptions)null);
            }, "null watermark options should be rejected");

            RequireThrows(delegate()
            {
                PdfFile file = PdfFile.Load(source);
                WatermarkOptions options = WatermarkOptions.ForText("字号");
                options.FontSize = 0;
                file.AddWatermark(options);
            }, "non-positive font size should be rejected");

            RequireThrows(delegate()
            {
                PdfFile file = PdfFile.Load(source);
                WatermarkOptions options = WatermarkOptions.ForText("透明度");
                options.Opacity = 1.5;
                file.AddWatermark(options);
            }, "opacity above one should be rejected");

            RequireThrows(delegate()
            {
                PdfFile file = PdfFile.Load(source);
                WatermarkOptions options = WatermarkOptions.ForText("边距");
                options.Margin = -1;
                file.AddWatermark(options);
            }, "negative margin should be rejected");

            PdfFile ignored = PdfFile.Load(source);
            WatermarkOptions outOfRange = WatermarkOptions.ForText("越页 OUT-OF-RANGE");
            ignored.AddWatermark(outOfRange, 9999);
            string text = PdfDocument.Load(ignored.SaveToBytes()).ExtractText();
            Require(!text.Contains("OUT-OF-RANGE"), "out-of-range page index should draw nothing");
        }

        /// <summary>
        /// 判断指定页的内容流数组首段或末段是否为水印。Determines whether the first or last content stream of a page is the watermark.
        /// </summary>
        /// <param name="file">对象级文档。The object-level document.</param>
        /// <param name="index">零基页码。The zero-based page index.</param>
        /// <param name="first">true 检查首段，false 检查末段。True checks the first stream, false the last.</param>
        /// <returns>目标段包含 XObject 绘制返回 true。True when the target stream draws an XObject.</returns>
        private static bool SlotHasWatermark(PdfFile file, int index, bool first)
        {
            List<PdfStream> streams = PageContentStreams(file, index);
            if (streams.Count == 0) return false;
            PdfStream target = first ? streams[0] : streams[streams.Count - 1];
            return Encoding.ASCII.GetString(file.DecodeStream(target)).Contains(" Do");
        }

        /// <summary>
        /// 取出指定页的全部内容流对象。Returns every content stream object of the given page.
        /// </summary>
        /// <param name="file">对象级文档。The object-level document.</param>
        /// <param name="index">零基页码。The zero-based page index.</param>
        /// <returns>内容流列表，Contents 为数组时按原顺序返回。The stream list, in order when Contents is an array.</returns>
        private static List<PdfStream> PageContentStreams(PdfFile file, int index)
        {
            PdfDictionary page = file.GetPageDictionary(index);
            PdfObject contents = file.Resolve(page.Get("Contents"));
            List<PdfStream> streams = new List<PdfStream>();

            PdfArray array = contents as PdfArray;
            if (array != null)
            {
                for (int i = 0; i < array.Count; i++)
                {
                    PdfStream item = file.Resolve(array[i]) as PdfStream;
                    if (item != null) streams.Add(item);
                }
                return streams;
            }

            PdfStream single = contents as PdfStream;
            if (single != null) streams.Add(single);
            return streams;
        }

        /// <summary>
        /// 读取指定页的水印外层内容流与内层 Form 内容流。Reads a page's outer watermark stream and its inner form stream.
        /// </summary>
        /// <param name="file">对象级文档。The object-level document.</param>
        /// <param name="index">零基页码。The zero-based page index.</param>
        /// <returns>长度为 2 的字符串数组：[0] 外层、[1] 内层，缺失时为空串。A two-item array: [0] the outer stream, [1] the form, empty when absent.</returns>
        private static string[] WatermarkPair(PdfFile file, int index)
        {
            string outer = WatermarkOuterStream(file, index);
            if (outer == null) return new string[] { string.Empty, string.Empty };

            string name = null;
            int slash = outer.IndexOf('/');
            if (slash >= 0)
            {
                int end = slash + 1;
                while (end < outer.Length && !char.IsWhiteSpace(outer[end])) end++;
                name = outer.Substring(slash + 1, end - slash - 1);
            }

            PdfDictionary page = file.GetPageDictionary(index);
            PdfDictionary resources = file.Resolve(page.Get("Resources")) as PdfDictionary;
            PdfDictionary xobjects = resources == null ? null : file.Resolve(resources.Get("XObject")) as PdfDictionary;
            PdfObject value = xobjects == null || name == null ? null : xobjects.Get(name);
            PdfStream form = file.Resolve(value) as PdfStream;
            string inner = form == null ? string.Empty : Encoding.ASCII.GetString(file.DecodeStream(form));
            return new string[] { outer, inner };
        }

        /// <summary>
        /// 找到以保存图形状态开头并绘制 XObject 的那段页面内容流。Finds the page content stream that opens a graphics state and draws an XObject.
        /// </summary>
        /// <param name="file">对象级文档。The object-level document.</param>
        /// <param name="index">零基页码。The zero-based page index.</param>
        /// <returns>内容流文本，未找到返回 null。The stream text, or null when none matches.</returns>
        private static string WatermarkOuterStream(PdfFile file, int index)
        {
            foreach (PdfStream stream in PageContentStreams(file, index))
            {
                string text = Encoding.ASCII.GetString(file.DecodeStream(stream));
                if (text.TrimStart().StartsWith("q", StringComparison.Ordinal) && text.Contains(" Do")) return text;
            }
            return null;
        }

        /// <summary>
        /// 校验内容流中的 q/Q 成对且不出现负深度。Verifies that q and Q pair up in a content stream without a negative depth.
        /// </summary>
        /// <param name="content">内容流文本。The content stream text.</param>
        /// <returns>配对完整返回 true。True when balanced.</returns>
        private static bool BalancedGraphics(string content)
        {
            string[] tokens = content.Split(new char[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            int depth = 0;
            foreach (string token in tokens)
            {
                if (token == "q") depth++;
                else if (token == "Q")
                {
                    depth--;
                    if (depth < 0) return false;
                }
            }
            return depth == 0;
        }

        /// <summary>
        /// 统计文档中的图像 XObject 数量。Counts the image XObjects in the document.
        /// </summary>
        /// <param name="file">对象级文档。The object-level document.</param>
        /// <returns>图像 XObject 数。The image XObject count.</returns>
        private static int CountImageXObjects(PdfFile file)
        {
            int count = 0;
            foreach (int number in file.ObjectNumbers)
            {
                PdfStream stream = file.GetObject<PdfStream>(number);
                if (stream == null) continue;
                PdfName subtype = stream.Dictionary.Get("Subtype") as PdfName;
                if (subtype != null && subtype.Value == "Image") count++;
            }
            return count;
        }
    }
}
