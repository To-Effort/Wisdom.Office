using System;
using System.Collections.Generic;
using System.IO;
using Wisdom.Excel;
using Wisdom.Office.Core;
using Wisdom.Word;

namespace Wisdom.Office.Tests
{
    internal static partial class Program
    {
        /// <summary>
        /// 图片格式保真：GIF/BMP/TIFF/EMF/WMF 要认得出、带得走，PNG/JPEG 之外的字节不再在读端凭空消失，
        /// 同时 PDF 仍然只嵌入它能栅格化的 PNG/JPEG。
        /// Image format fidelity: GIF, BMP, TIFF, EMF and WMF are detected and carried, so bytes that are not
        /// PNG or JPEG no longer vanish on the read side, while PDF still embeds only the PNG and JPEG it can
        /// actually rasterise.
        /// </summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunImageFidelityTest(string outputDirectory)
        {
            RunImageSniffingTest();
            RunDocxKeepsGifImageTest(outputDirectory);
            RunXlsxKeepsGifImageTest(outputDirectory);
            RunHtmlImageFallbackTest();
            Console.WriteLine("[image] gif/bmp/emf/tiff/wmf sniffing, docx and xlsx carry-through, html fallback ok");
        }

        /// <summary>嗅探：魔数决定格式、MIME 与扩展名，GIF/BMP 还要读出像素尺寸。 Sniffing: the magic bytes decide format, MIME type and extension, and GIF/BMP also report their pixel size.</summary>
        private static void RunImageSniffingTest()
        {
            OfficeImage gif = new OfficeImage(CreateTestGif(3, 2));
            Require(gif.Format == OfficeImageFormat.Gif, "a GIF89a header was not sniffed as gif: " + gif.Format);
            Require(gif.ContentType == "image/gif", "wrong gif mime: " + gif.ContentType);
            Require(gif.FileExtension == ".gif", "wrong gif extension: " + gif.FileExtension);
            Require(gif.PixelWidth == 3 && gif.PixelHeight == 2,
                "the gif logical screen size was lost: " + gif.PixelWidth + "x" + gif.PixelHeight);

            OfficeImage bmp = new OfficeImage(CreateTestBmp(4, 5));
            Require(bmp.Format == OfficeImageFormat.Bmp, "a BM header was not sniffed as bmp: " + bmp.Format);
            Require(bmp.ContentType == "image/bmp", "wrong bmp mime: " + bmp.ContentType);
            Require(bmp.PixelWidth == 4 && bmp.PixelHeight == 5,
                "the bmp info header size was lost: " + bmp.PixelWidth + "x" + bmp.PixelHeight);

            OfficeImage bottomUp = new OfficeImage(CreateTestBmp(4, -5));
            Require(bottomUp.PixelHeight == 5,
                "a bottom-up BMP height must be reported positive: " + bottomUp.PixelHeight);

            OfficeImage emf = new OfficeImage(CreateTestEmf());
            Require(emf.Format == OfficeImageFormat.Emf, "an EMF record header was not sniffed: " + emf.Format);
            Require(emf.ContentType == "image/x-emf", "wrong emf mime: " + emf.ContentType);

            OfficeImage wmf = new OfficeImage(CreateTestWmf());
            Require(wmf.Format == OfficeImageFormat.Wmf, "a placeable WMF header was not sniffed: " + wmf.Format);
            Require(wmf.ContentType == "image/x-wmf", "wrong wmf mime: " + wmf.ContentType);

            OfficeImage tiff = new OfficeImage(CreateTestTiff());
            Require(tiff.Format == OfficeImageFormat.Tiff, "a little-endian TIFF header was not sniffed: " + tiff.Format);
            Require(tiff.FileExtension == ".tiff", "wrong tiff extension: " + tiff.FileExtension);

            OfficeImage junk = new OfficeImage(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14 });
            Require(junk.Format == OfficeImageFormat.Unknown, "random bytes were sniffed as an image: " + junk.Format);

            // 短于头部长度的字节不能越界读。
            // Bytes shorter than a header must not be read out of bounds.
            OfficeImage truncated = new OfficeImage(new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39 });
            Require(truncated.Format == OfficeImageFormat.Unknown,
                "a truncated GIF header was taken as a real image: " + truncated.Format);
        }

        /// <summary>docx 里的一张 GIF 必须原样带回来，并写成 .gif 媒体部件；过去只认 PNG/JPEG，插图直接被吞掉。
        /// A GIF inside a docx has to come back unchanged and be written as a .gif media part; accepting only
        /// PNG and JPEG used to swallow the illustration outright.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunDocxKeepsGifImageTest(string outputDirectory)
        {
            byte[] gif = CreateTestGif(8, 6);
            OfficeDocument document = new OfficeDocument();
            Section section = document.AddSection();
            Paragraph picture = new Paragraph();
            section.Blocks.Add(picture);
            picture.AddImage(new OfficeImage(gif), 60, 45);
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "image-fidelity-gif.docx"));
            DocxWriter.Write(document, path);

            string contentTypes = ReadPackagePartText(path, "[Content_Types].xml");
            Require(contentTypes.Contains("Extension=\"gif\""),
                "the gif media part was written without its content type: " + contentTypes);

            OfficeDocument loaded = DocxReader.Load(path);
            ImageRun run = FindFirstImage(loaded);
            Require(run != null, "the gif image disappeared from the docx round trip");
            Require(run.Image.Format == OfficeImageFormat.Gif,
                "the reloaded image is no longer a gif: " + run.Image.Format);
            Require(run.Image.PixelWidth == 8 && run.Image.PixelHeight == 6,
                "the gif size was lost on the way back: " + run.Image.PixelWidth + "x" + run.Image.PixelHeight);
            Require(run.Image.Data.Length == gif.Length,
                "the gif bytes changed: " + run.Image.Data.Length + " vs " + gif.Length);

            // 再写一次：媒体部件仍是 .gif，Word 打开不会因为 image.bin 而报文档损坏。
            // Written again, the media part is still .gif; Word would call an image.bin a corrupt document.
            string second = PrepareOutputPath(Path.Combine(outputDirectory, "image-fidelity-gif-again.docx"));
            DocxWriter.Write(loaded, second);
            string again = ReadPackagePartText(second, "[Content_Types].xml");
            Require(again.Contains("Extension=\"gif\"") && !again.Contains("Extension=\"bin\""),
                "the second write lost the gif media type: " + again);
        }

        /// <summary>工作簿里的 GIF 同样要带得走：媒体部件名与 [Content_Types].xml 都得按真实格式写。
        /// A GIF inside a workbook has to travel too: both the media part name and [Content_Types].xml follow
        /// the real format.</summary>
        /// <param name="outputDirectory">输出目录。The output directory.</param>
        private static void RunXlsxKeepsGifImageTest(string outputDirectory)
        {
            OfficeWorkbook workbook = new OfficeWorkbook();
            Worksheet sheet = workbook.AddWorksheet("Logo");
            sheet.GetCell(1, 1).Value = "x";
            ExcelImage logo = new ExcelImage(new OfficeImage(CreateTestGif(16, 9)));
            logo.Row = 1;
            logo.Column = 3;
            logo.WidthPoints = 40;
            logo.HeightPoints = 24;
            sheet.Images.Add(logo);
            string path = PrepareOutputPath(Path.Combine(outputDirectory, "image-fidelity-gif.xlsx"));
            XlsxWriter.Write(workbook, path);

            string contentTypes = ReadPackagePartText(path, "[Content_Types].xml");
            Require(contentTypes.Contains("Extension=\"gif\" ContentType=\"image/gif\""),
                "the workbook guessed a mime type for the gif: " + contentTypes);
            OfficeWorkbook loaded = XlsxReader.Load(path);
            Require(loaded.Worksheets[0].Images.Count == 1, "the workbook image disappeared: "
                + loaded.Worksheets[0].Images.Count);
            Require(loaded.Worksheets[0].Images[0].Image.Format == OfficeImageFormat.Gif,
                "the reloaded workbook image is no longer a gif: " + loaded.Worksheets[0].Images[0].Image.Format);
            Require(loaded.Worksheets[0].Images[0].Image.PixelWidth == 16,
                "the gif size was lost through the workbook: "
                + loaded.Worksheets[0].Images[0].Image.PixelWidth);
        }

        /// <summary>HTML 的内嵌图片：认得出的格式按各自 MIME 走，魔数对不上就退回 alt 文本，绝不能塞成 .bin。
        /// Inline HTML images: a recognised format keeps its own mime type, bytes whose magic does not match fall
        /// back to the alt text and must never be stored as .bin.</summary>
        private static void RunHtmlImageFallbackTest()
        {
            string gifHtml = "<html><body><p><img src=\"data:image/gif;base64,"
                + Convert.ToBase64String(CreateTestGif(4, 4)) + "\" alt=\"g\" /></p></body></html>";
            OfficeDocument document = HtmlReader.Parse(gifHtml);
            Require(FindFirstImage(document) != null, "a base64 gif in html was dropped");
            Require(FindFirstImage(document).Image.Format == OfficeImageFormat.Gif,
                "the html gif kept another format: " + FindFirstImage(document).Image.Format);

            string hostile = "<html><body><p><img src=\"data:image/png;base64,"
                + Convert.ToBase64String(new byte[] { 9, 8, 7, 6, 5, 4, 3, 2, 1, 0, 9, 8, 7, 6, 5, 4 })
                + "\" alt=\"missing\" /></p></body></html>";
            OfficeDocument refused = HtmlReader.Parse(hostile);
            Require(FindFirstImage(refused) == null,
                "bytes with no image magic were accepted just because the mime said image/png");
            RequireContains(refused.GetPlainText(), "missing");

            string emfHtml = "<html><body><p><img src=\"data:image/x-emf;base64,"
                + Convert.ToBase64String(CreateTestEmf()) + "\" alt=\"e\" /></p></body></html>";
            Require(FindFirstImage(HtmlReader.Parse(emfHtml)) == null,
                "a vector metafile was inlined into html output where no browser can paint it");
        }

        /// <summary>遍历文档正文找出第一张图片。 Walks the document body for the first image.</summary>
        /// <param name="document">待检查的文档。The document to inspect.</param>
        /// <returns>首个图片行内元素；没有时返回 null。The first image inline, or null.</returns>
        private static ImageRun FindFirstImage(OfficeDocument document)
        {
            for (int s = 0; s < document.Sections.Count; s++)
            {
                foreach (Block block in document.Sections[s].Blocks)
                {
                    Paragraph paragraph = block as Paragraph;
                    if (paragraph == null) continue;
                    for (int i = 0; i < paragraph.Inlines.Count; i++)
                    {
                        ImageRun run = paragraph.Inlines[i] as ImageRun;
                        if (run != null && run.Image != null) return run;
                    }
                }
            }
            return null;
        }

        /// <summary>构造一枚最小 GIF89a：只有逻辑屏幕描述符与结束符，够嗅探与取尺寸。 Builds a minimal GIF89a: a logical screen descriptor plus the trailer, enough to sniff and read the size.</summary>
        /// <param name="width">像素宽度。 Pixel width.</param>
        /// <param name="height">像素高度。 Pixel height.</param>
        /// <returns>图片字节。 The image bytes.</returns>
        private static byte[] CreateTestGif(int width, int height)
        {
            byte[] data = new byte[14];
            data[0] = 0x47;
            data[1] = 0x49;
            data[2] = 0x46;
            data[3] = 0x38;
            data[4] = 0x39;
            data[5] = 0x61;
            data[6] = (byte)(width & 0xFF);
            data[7] = (byte)((width >> 8) & 0xFF);
            data[8] = (byte)(height & 0xFF);
            data[9] = (byte)((height >> 8) & 0xFF);
            data[13] = 0x3B;
            return data;
        }

        /// <summary>构造一枚最小 BMP（BITMAPINFOHEADER，40 字节头）。 Builds a minimal BMP with a 40-byte BITMAPINFOHEADER.</summary>
        /// <param name="width">像素宽度。 Pixel width.</param>
        /// <param name="height">像素高度；负数表示自顶向下的位图。 Pixel height; negative means a top-down bitmap.</param>
        /// <returns>图片字节。 The image bytes.</returns>
        private static byte[] CreateTestBmp(int width, int height)
        {
            byte[] data = new byte[54];
            data[0] = 0x42;
            data[1] = 0x4D;
            WriteInt32LittleEndian(data, 2, data.Length);
            WriteInt32LittleEndian(data, 14, 40);
            WriteInt32LittleEndian(data, 18, width);
            WriteInt32LittleEndian(data, 22, height);
            data[26] = 1;
            data[28] = 24;
            return data;
        }

        /// <summary>构造一枚最小 EMF：首记录类型 1 与 " EMF" 签名。 Builds a minimal EMF: the first record type 1 plus the " EMF" signature.</summary>
        /// <returns>图片字节。 The image bytes.</returns>
        private static byte[] CreateTestEmf()
        {
            byte[] data = new byte[48];
            WriteInt32LittleEndian(data, 0, 1);
            WriteInt32LittleEndian(data, 4, 48);
            data[40] = 0x20;
            data[41] = 0x45;
            data[42] = 0x4D;
            data[43] = 0x46;
            return data;
        }

        /// <summary>构造一枚可放置 WMF（D7 CDC6 9A 魔数）。 Builds a placeable WMF with the D7 CDC6 9A magic.</summary>
        /// <returns>图片字节。 The image bytes.</returns>
        private static byte[] CreateTestWmf()
        {
            byte[] data = new byte[22];
            data[0] = 0xD7;
            data[1] = 0xCD;
            data[2] = 0xC6;
            data[3] = 0x9A;
            return data;
        }

        /// <summary>构造一枚最小 TIFF 头（小端 II 加魔数 42）。 Builds a minimal little-endian TIFF header: II plus the 42 magic.</summary>
        /// <returns>图片字节。 The image bytes.</returns>
        private static byte[] CreateTestTiff()
        {
            return new byte[] { 0x49, 0x49, 0x2A, 0x00, 0x08, 0x00, 0x00, 0x00 };
        }

        /// <summary>写入小端 32 位整数。 Writes a little-endian 32-bit integer.</summary>
        /// <param name="data">字节缓冲。 The byte buffer.</param>
        /// <param name="offset">起始偏移。 The offset.</param>
        /// <param name="value">整数值。 The value.</param>
        private static void WriteInt32LittleEndian(byte[] data, int offset, int value)
        {
            data[offset] = (byte)(value & 0xFF);
            data[offset + 1] = (byte)((value >> 8) & 0xFF);
            data[offset + 2] = (byte)((value >> 16) & 0xFF);
            data[offset + 3] = (byte)((value >> 24) & 0xFF);
        }
    }
}
