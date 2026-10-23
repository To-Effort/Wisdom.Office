using System;
using System.Collections.Generic;
using System.IO;
using Wisdom.Office.Core;

namespace Wisdom.Pdf
{
    /// <summary>嵌入 PDF 输出的一张图像，负责写出图像 XObject。An image embedded in PDF output, responsible for writing the image XObject.</summary>
    /// <remarks>
    /// JPEG 原样透传（DCTDecode），其余格式解码为位图并以 Flate 压缩，透明通道写为 SMask。
    /// JPEG data is passed through unchanged (DCTDecode); other formats are decoded to bitmaps and
    /// Flate-compressed, with any alpha channel emitted as an SMask.
    /// </remarks>
    internal sealed class PdfImage
    {
        /// <summary>源图像数据。The source image data.</summary>
        private readonly OfficeImage _source;
        /// <summary>惰性解码后的图像位图数据。The lazily decoded image bitmap data.</summary>
        private PdfImageData _decoded;
        /// <summary>是否已写出 XObject，避免重复登记。Whether the XObject was already written, preventing duplicate registration.</summary>
        private bool _prepared;

        /// <summary>资源字典中的名称（如 Im1）。The name in the resource dictionary (e.g. Im1).</summary>
        public string ResourceName { get; set; }
        /// <summary>图像像素宽度。The pixel width of the image.</summary>
        public double Width { get { return _source.PixelWidth; } }
        /// <summary>图像像素高度。The pixel height of the image.</summary>
        public double Height { get { return _source.PixelHeight; } }
        /// <summary>源图像对象。The source image object.</summary>
        public OfficeImage Source { get { return _source; } }

        /// <summary>
        /// 创建一个待嵌入的 PDF 图像。
        /// Creates a PDF image pending embedding.
        /// </summary>
        /// <param name="source">源图像。The source image.</param>
        public PdfImage(OfficeImage source)
        {
            _source = source;
        }

        /// <summary>
        /// 把图像写成 XObject 并登记到页面资源；重复调用无效果。
        /// Writes the image as an XObject and registers it in the page resources; repeated calls have no effect.
        /// </summary>
        /// <param name="writer">对象写出器。The object writer.</param>
        /// <param name="imageResources">页面 XObject 资源字典。The page XObject resource dictionary.</param>
        public void Prepare(PdfObjectWriter writer, PdfDictionary imageResources)
        {
            if (_prepared) return;
            if (_decoded == null) _decoded = PdfImageData.Decode(_source);

            PdfDictionary dictionary = new PdfDictionary();
            dictionary.SetName("Type", "XObject");
            dictionary.SetName("Subtype", "Image");
            dictionary.SetNumber("Width", _decoded.Width);
            dictionary.SetNumber("Height", _decoded.Height);

            if (_decoded.IsJpeg)
            {
                dictionary.SetName("ColorSpace", _decoded.ColorSpace);
                dictionary.SetNumber("BitsPerComponent", _decoded.BitsPerComponent);
                dictionary.SetName("Filter", "DCTDecode");
                if (_decoded.ColorSpace == "DeviceCMYK")
                {
                    dictionary.Set("Decode", new PdfArray()
                        .AddNumber(1).AddNumber(0)
                        .AddNumber(1).AddNumber(0)
                        .AddNumber(1).AddNumber(0)
                        .AddNumber(1).AddNumber(0));
                }
                imageResources.Set(ResourceName, writer.AddStream(dictionary, _source.Data, false));
                _prepared = true;
                return;
            }

            dictionary.SetName("ColorSpace", _decoded.IsGray ? "DeviceGray" : "DeviceRGB");
            dictionary.SetNumber("BitsPerComponent", 8);
            dictionary.SetBoolean("Interpolate", true);

            if (_decoded.Alpha != null)
            {
                PdfDictionary alphaDictionary = new PdfDictionary();
                alphaDictionary.SetNumber("Width", _decoded.Width);
                alphaDictionary.SetNumber("Height", _decoded.Height);
                alphaDictionary.SetName("ColorSpace", "DeviceGray");
                alphaDictionary.SetNumber("BitsPerComponent", 8);
                alphaDictionary.SetName("Filter", "FlateDecode");
                PdfReference alphaRef = writer.AddStream(alphaDictionary, Zlib.Compress(_decoded.Alpha), false);
                dictionary.Set("SMask", alphaRef);
            }

            dictionary.SetName("Filter", "FlateDecode");
            imageResources.Set(ResourceName, writer.AddStream(dictionary, Zlib.Compress(_decoded.Pixels), false));
            _prepared = true;
        }

        /// <summary>
        /// 计算图像内容的 FNV-1a 哈希，用于跨页去重。
        /// Computes the FNV-1a hash of the image bytes, used for deduplication across pages.
        /// </summary>
        /// <param name="data">图像原始字节。The raw image bytes.</param>
        /// <returns>64 位哈希值。The 64-bit hash value.</returns>
        public static ulong ComputeHash(byte[] data)
        {
            ulong hash = 14695981039346656037UL;
            for (int i = 0; i < data.Length; i++)
            {
                hash ^= data[i];
                hash *= 1099511628211UL;
            }
            return hash;
        }
    }

    /// <summary>解码后的图像位图数据，供写出图像 XObject 使用。Decoded image bitmap data used when writing the image XObject.</summary>
    internal sealed class PdfImageData
    {
        private const long MaxPixelBytes = 256L * 1024 * 1024;

        /// <summary>像素宽度。The pixel width.</summary>
        public int Width;
        /// <summary>像素高度。The pixel height.</summary>
        public int Height;
        /// <summary>是否为可直接透传的 JPEG。Whether this is a JPEG that can be passed through directly.</summary>
        public bool IsJpeg;
        /// <summary>是否为灰度图像。Whether the image is grayscale.</summary>
        public bool IsGray;
        /// <summary>每分量位数。The bits per component.</summary>
        public int BitsPerComponent;
        /// <summary>JPEG 的色彩空间名（DeviceGray/DeviceRGB/DeviceCMYK）。The JPEG color space name (DeviceGray/DeviceRGB/DeviceCMYK).</summary>
        public string ColorSpace;
        /// <summary>展开后的像素字节（灰度 1 通道、彩色 3 通道）。The expanded pixel bytes (1 channel for gray, 3 for color).</summary>
        public byte[] Pixels;
        /// <summary>单独的 Alpha 通道，全不透明时为 null。The separate alpha channel; null when fully opaque.</summary>
        public byte[] Alpha;

        /// <summary>
        /// 按图像格式分派解码，仅支持 PNG 与 JPEG。
        /// Dispatches decoding by image format; only PNG and JPEG are supported.
        /// </summary>
        /// <param name="image">源图像。The source image.</param>
        /// <returns>解码后的位图数据。The decoded bitmap data.</returns>
        /// <exception cref="ArgumentNullException">源图像为 null 时抛出。Thrown when the source image is null.</exception>
        /// <exception cref="NotSupportedException">格式不是 PNG 或 JPEG 时抛出。Thrown when the format is neither PNG nor JPEG.</exception>
        public static PdfImageData Decode(OfficeImage image)
        {
            if (image == null) throw new ArgumentNullException("image");
            if (image.Format == OfficeImageFormat.Jpeg)
            {
                return DecodeJpeg(image);
            }
            if (image.Format == OfficeImageFormat.Png)
            {
                return DecodePng(image);
            }
            throw new NotSupportedException("Only PNG and JPEG images can be embedded in PDF output.");
        }

        /// <summary>
        /// 解析 JPEG 头部信息：不解码熵数据，只读取 SOF 段的尺寸、分量数与精度。
        /// Parses the JPEG header: entropy data is never decoded, only the SOF segment's dimensions,
        /// component count and precision are read.
        /// </summary>
        /// <param name="image">源图像，其字节将原样嵌入。The source image whose bytes are embedded verbatim.</param>
        /// <returns>标记为 JPEG 的位图数据。The bitmap data flagged as JPEG.</returns>
        /// <exception cref="InvalidDataException">无法定位有效 SOF 段时抛出。Thrown when no valid SOF segment can be located.</exception>
        /// <exception cref="NotSupportedException">分量数不是 1/3/4 时抛出。Thrown when the component count is not 1/3/4.</exception>
        private static PdfImageData DecodeJpeg(OfficeImage image)
        {
            int width;
            int height;
            int components;
            int precision;
            if (!ReadJpegInfo(image.Data, out width, out height, out components, out precision))
            {
                throw new InvalidDataException("Unable to parse JPEG image.");
            }
            PdfImageData data = new PdfImageData();
            data.Width = width;
            data.Height = height;
            data.IsJpeg = true;
            data.BitsPerComponent = precision;
            if (components == 1)
            {
                data.ColorSpace = "DeviceGray";
                data.IsGray = true;
            }
            else if (components == 3)
            {
                data.ColorSpace = "DeviceRGB";
            }
            else if (components == 4)
            {
                data.ColorSpace = "DeviceCMYK";
            }
            else
            {
                throw new NotSupportedException("Unsupported JPEG component count: " + components);
            }
            return data;
        }

        /// <summary>
        /// 扫描 JPEG 标记段定位 SOF0–SOF15 帧头，读出精度、高、宽与分量数。
        /// Scans JPEG marker segments to locate an SOF0-SOF15 frame header and reads precision, height, width
        /// and component count.
        /// </summary>
        /// <param name="data">JPEG 字节。The JPEG bytes.</param>
        /// <param name="width">输出图像宽度。Outputs the image width.</param>
        /// <param name="height">输出图像高度。Outputs the image height.</param>
        /// <param name="components">输出分量数。Outputs the component count.</param>
        /// <param name="precision">输出每分量位数。Outputs the bits per component.</param>
        /// <returns>成功解析到有效帧头返回 true。True when a valid frame header was parsed.</returns>
        private static bool ReadJpegInfo(byte[] data, out int width, out int height, out int components, out int precision)
        {
            width = 0;
            height = 0;
            components = 0;
            precision = 8;
            int position = 2;
            while (position + 3 < data.Length)
            {
                if (data[position] != 0xFF)
                {
                    position++;
                    continue;
                }
                int marker = data[position + 1];
                position += 2;
                if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7) || marker == 0x01)
                {
                    continue;
                }
                if (position + 1 >= data.Length) break;
                int length = (data[position] << 8) | data[position + 1];
                if (length < 2) break;
                if ((marker >= 0xC0 && marker <= 0xC3)
                    || (marker >= 0xC5 && marker <= 0xC7)
                    || (marker >= 0xC9 && marker <= 0xCB)
                    || (marker >= 0xCD && marker <= 0xCF))
                {
                    if (position + 7 >= data.Length) break;
                    precision = data[position + 2];
                    height = (data[position + 3] << 8) | data[position + 4];
                    width = (data[position + 5] << 8) | data[position + 6];
                    components = data[position + 7];
                    return width > 0 && height > 0 && components > 0;
                }
                position += length;
            }
            return false;
        }

        /// <summary>
        /// 完整解码 PNG：遍历块、解压 IDAT、逐行撤销 PNG 滤波器，再按色彩类型展开为 8 位 RGB/灰度加可选 Alpha。
        /// Fully decodes a PNG: walks the chunks, inflates IDAT, undoes the per-row PNG filters and expands
        /// the color type into 8-bit RGB/gray plus optional alpha.
        /// </summary>
        /// <param name="image">源图像。The source image.</param>
        /// <returns>展开后的位图数据，全不透明时 Alpha 为 null。The expanded bitmap data; Alpha is null when fully opaque.</returns>
        /// <exception cref="NotSupportedException">遇到隔行扫描、不支持的位深/色彩类型/滤波类型时抛出。Thrown for interlacing or unsupported bit depth, color type or filter type.</exception>
        /// <exception cref="InvalidDataException">尺寸非法或调色板缺失时抛出。Thrown for invalid dimensions or a missing palette.</exception>
        private static PdfImageData DecodePng(OfficeImage image)
        {
            byte[] png = image.Data;
            int width = 0;
            int height = 0;
            int bitDepth = 8;
            int colorType = 6;
            int interlace = 0;
            byte[] palette = null;
            byte[] transparency = null;
            List<byte[]> idatChunks = new List<byte[]>();

            int position = 8;
            while (position + 12 <= png.Length)
            {
                int length = (png[position] << 24) | (png[position + 1] << 16) | (png[position + 2] << 8) | png[position + 3];
                string type = new string(new[]
                {
                    (char)png[position + 4], (char)png[position + 5], (char)png[position + 6], (char)png[position + 7]
                });
                int dataOffset = position + 8;
                if (length < 0 || dataOffset + length > png.Length) break;
                if (type == "IHDR")
                {
                    width = (png[dataOffset] << 24) | (png[dataOffset + 1] << 16) | (png[dataOffset + 2] << 8) | png[dataOffset + 3];
                    height = (png[dataOffset + 4] << 24) | (png[dataOffset + 5] << 16) | (png[dataOffset + 6] << 8) | png[dataOffset + 7];
                    bitDepth = png[dataOffset + 8];
                    colorType = png[dataOffset + 9];
                    int compression = png[dataOffset + 10];
                    int filter = png[dataOffset + 11];
                    interlace = png[dataOffset + 12];
                    if (compression != 0 || filter != 0)
                    {
                        throw new NotSupportedException("Unsupported PNG compression or filter method.");
                    }
                }
                else if (type == "PLTE")
                {
                    palette = new byte[length];
                    Buffer.BlockCopy(png, dataOffset, palette, 0, length);
                }
                else if (type == "tRNS")
                {
                    transparency = new byte[length];
                    Buffer.BlockCopy(png, dataOffset, transparency, 0, length);
                }
                else if (type == "IDAT")
                {
                    byte[] chunk = new byte[length];
                    Buffer.BlockCopy(png, dataOffset, chunk, 0, length);
                    idatChunks.Add(chunk);
                }
                else if (type == "IEND")
                {
                    break;
                }
                position = dataOffset + length + 4;
            }

            if (width <= 0 || height <= 0) throw new InvalidDataException("Invalid PNG dimensions.");
            // 单边与像素总量双重上限：int 相乘会翻负或落到很小的正值，像素缓冲区正是按它分配的。
            // Both a single side and the total area are capped: int multiplication can wrap negative or
            // land on a small positive value, and the pixel buffer is sized from that product.
            if (width > 65535 || height > 65535) throw new InvalidDataException("PNG dimensions exceed the safety limit.");
            if (interlace != 0) throw new NotSupportedException("Interlaced PNG images are not supported.");
            if (bitDepth != 8 && bitDepth != 16 && bitDepth != 4 && bitDepth != 2 && bitDepth != 1)
            {
                throw new NotSupportedException("Unsupported PNG bit depth: " + bitDepth);
            }

            int channels;
            switch (colorType)
            {
                case 0: channels = 1; break;
                case 2: channels = 3; break;
                case 3: channels = 1; break;
                case 4: channels = 2; break;
                case 6: channels = 4; break;
                default: throw new NotSupportedException("Unsupported PNG color type: " + colorType);
            }

            int rowBytes = (width * channels * bitDepth + 7) / 8;
            if ((long)rowBytes * height > MaxPixelBytes) throw new InvalidDataException("PNG pixel buffer exceeds the safety limit.");
            int bpp = Math.Max(1, channels * bitDepth / 8);
            byte[] decompressed;
            using (MemoryStream idat = new MemoryStream())
            {
                for (int i = 0; i < idatChunks.Count; i++)
                {
                    idat.Write(idatChunks[i], 0, idatChunks[i].Length);
                }
                byte[] idatBytes = idat.ToArray();
                // IDAT 拼起来可能是压缩爆弹，解压后字节数必须受上限约束。
                // The concatenated IDAT stream may be a zip bomb, so the inflated size stays bounded.
                decompressed = Zlib.Decompress(idatBytes, 0, idatBytes.Length, PdfReadLimits.MaxDecodedStreamBytes);
            }

            int expected = rowBytes * height + height;
            if (decompressed.Length < expected)
            {
                byte[] padded = new byte[expected];
                Buffer.BlockCopy(decompressed, 0, padded, 0, decompressed.Length);
                decompressed = padded;
            }

            byte[] pixels = new byte[rowBytes * height];
            int input = 0;
            int output = 0;
            for (int y = 0; y < height; y++)
            {
                int filter = decompressed[input++];
                for (int x = 0; x < rowBytes; x++)
                {
                    int raw = decompressed[input++];
                    int left = x >= bpp ? pixels[output + x - bpp] : 0;
                    int up = y > 0 ? pixels[output - rowBytes + x] : 0;
                    int upLeft = (y > 0 && x >= bpp) ? pixels[output - rowBytes + x - bpp] : 0;
                    int value;
                    switch (filter)
                    {
                        case 0: value = raw; break;
                        case 1: value = raw + left; break;
                        case 2: value = raw + up; break;
                        case 3: value = raw + ((left + up) >> 1); break;
                        case 4: value = raw + Paeth(left, up, upLeft); break;
                        default: throw new NotSupportedException("Unsupported PNG filter type: " + filter);
                    }
                    pixels[output + x] = (byte)value;
                }
                output += rowBytes;
            }

            PdfImageData result = new PdfImageData();
            result.Width = width;
            result.Height = height;
            result.IsGray = colorType == 0 || colorType == 4;
            int pixelCount = width * height;
            result.Pixels = new byte[pixelCount * (result.IsGray ? 1 : 3)];
            bool hasAlpha = colorType == 4 || colorType == 6 || transparency != null;
            if (hasAlpha)
            {
                result.Alpha = new byte[pixelCount];
                for (int i = 0; i < result.Alpha.Length; i++) result.Alpha[i] = 255;
            }

            int maxSample = (1 << bitDepth) - 1;
            for (int y = 0; y < height; y++)
            {
                int rowStart = y * rowBytes;
                for (int x = 0; x < width; x++)
                {
                    int pixelIndex = y * width + x;
                    int r = 0, g = 0, b = 0, a = 255;
                    if (colorType == 0)
                    {
                        int sample = ReadPngSample(pixels, rowStart, x, 1, 0, bitDepth);
                        int gray = ScaleSample(sample, bitDepth, maxSample);
                        r = g = b = gray;
                        if (transparency != null && transparency.Length >= 2)
                        {
                            int transparentGray = (transparency[0] << 8) | transparency[1];
                            if (sample == transparentGray) a = 0;
                        }
                    }
                    else if (colorType == 2)
                    {
                        int sr = ReadPngSample(pixels, rowStart, x, 3, 0, bitDepth);
                        int sg = ReadPngSample(pixels, rowStart, x, 3, 1, bitDepth);
                        int sb = ReadPngSample(pixels, rowStart, x, 3, 2, bitDepth);
                        r = ScaleSample(sr, bitDepth, maxSample);
                        g = ScaleSample(sg, bitDepth, maxSample);
                        b = ScaleSample(sb, bitDepth, maxSample);
                        if (transparency != null && transparency.Length >= 6)
                        {
                            int tr = (transparency[0] << 8) | transparency[1];
                            int tg = (transparency[2] << 8) | transparency[3];
                            int tb = (transparency[4] << 8) | transparency[5];
                            if (sr == tr && sg == tg && sb == tb) a = 0;
                        }
                    }
                    else if (colorType == 3)
                    {
                        int index = ReadPngSample(pixels, rowStart, x, 1, 0, bitDepth);
                        if (palette == null || index * 3 + 2 >= palette.Length)
                        {
                            throw new InvalidDataException("PNG palette is missing or invalid.");
                        }
                        r = palette[index * 3];
                        g = palette[index * 3 + 1];
                        b = palette[index * 3 + 2];
                        if (transparency != null && index < transparency.Length)
                        {
                            a = transparency[index];
                        }
                    }
                    else if (colorType == 4)
                    {
                        int gray = ReadPngSample(pixels, rowStart, x, 2, 0, bitDepth);
                        a = ReadPngSample(pixels, rowStart, x, 2, 1, bitDepth);
                        r = g = b = ScaleSample(gray, bitDepth, maxSample);
                        a = ScaleSample(a, bitDepth, maxSample);
                    }
                    else if (colorType == 6)
                    {
                        int sr = ReadPngSample(pixels, rowStart, x, 4, 0, bitDepth);
                        int sg = ReadPngSample(pixels, rowStart, x, 4, 1, bitDepth);
                        int sb = ReadPngSample(pixels, rowStart, x, 4, 2, bitDepth);
                        int sa = ReadPngSample(pixels, rowStart, x, 4, 3, bitDepth);
                        r = ScaleSample(sr, bitDepth, maxSample);
                        g = ScaleSample(sg, bitDepth, maxSample);
                        b = ScaleSample(sb, bitDepth, maxSample);
                        a = ScaleSample(sa, bitDepth, maxSample);
                    }

                    if (result.IsGray)
                    {
                        result.Pixels[pixelIndex] = (byte)r;
                    }
                    else
                    {
                        result.Pixels[pixelIndex * 3] = (byte)r;
                        result.Pixels[pixelIndex * 3 + 1] = (byte)g;
                        result.Pixels[pixelIndex * 3 + 2] = (byte)b;
                    }
                    if (result.Alpha != null)
                    {
                        result.Alpha[pixelIndex] = (byte)a;
                    }
                }
            }

            if (result.Alpha != null)
            {
                bool opaque = true;
                for (int i = 0; i < result.Alpha.Length; i++)
                {
                    if (result.Alpha[i] != 255)
                    {
                        opaque = false;
                        break;
                    }
                }
                if (opaque) result.Alpha = null;
            }

            return result;
        }

        /// <summary>
        /// 从反滤波后的行字节中读取指定位深的像素分量，支持 1/2/4/8/16 位与子字节打包。
        /// Reads one pixel component of the given bit depth from the unfiltered row bytes, handling
        /// 1/2/4/8/16-bit and sub-byte packing.
        /// </summary>
        /// <param name="data">反滤波后的完整像素字节。The fully unfiltered pixel bytes.</param>
        /// <param name="rowStart">该行的起始偏移。The start offset of the row.</param>
        /// <param name="x">像素列号。The pixel column index.</param>
        /// <param name="channels">每像素通道数。The channels per pixel.</param>
        /// <param name="channel">通道序号。The channel index.</param>
        /// <param name="bitDepth">每分量位数。The bits per component.</param>
        /// <returns>原始样本值（未缩放到 8 位）。The raw sample value (not yet scaled to 8 bits).</returns>
        private static int ReadPngSample(byte[] data, int rowStart, int x, int channels, int channel, int bitDepth)
        {
            if (bitDepth == 8)
            {
                return data[rowStart + x * channels + channel];
            }
            if (bitDepth == 16)
            {
                int offset = rowStart + (x * channels + channel) * 2;
                return (data[offset] << 8) | data[offset + 1];
            }
            int perByte = 8 / bitDepth;
            int byteIndex = x / perByte;
            int shift = 8 - bitDepth * ((x % perByte) + 1);
            int mask = (1 << bitDepth) - 1;
            return (data[rowStart + byteIndex] >> shift) & mask;
        }

        /// <summary>
        /// 把任意位深的样本线性缩放到 8 位。
        /// Linearly scales a sample of any bit depth down to 8 bits.
        /// </summary>
        /// <param name="sample">原始样本值。The raw sample value.</param>
        /// <param name="bitDepth">源位深。The source bit depth.</param>
        /// <param name="maxSample">源最大值。The source maximum value.</param>
        /// <returns>0–255 的 8 位样本。An 8-bit sample from 0 to 255.</returns>
        private static int ScaleSample(int sample, int bitDepth, int maxSample)
        {
            if (bitDepth == 8) return sample;
            if (bitDepth == 16) return sample >> 8;
            return (sample * 255) / maxSample;
        }

        /// <summary>
        /// PNG Paeth 预测函数，选出最接近线性预测的邻居值。
        /// The PNG Paeth predictor, selecting whichever neighbor is closest to the linear prediction.
        /// </summary>
        /// <param name="a">左邻居值。The left neighbor value.</param>
        /// <param name="b">上方邻居值。The value above.</param>
        /// <param name="c">左上邻居值。The up-left neighbor value.</param>
        /// <returns>选中的预测值。The selected predictor value.</returns>
        private static int Paeth(int a, int b, int c)
        {
            int p = a + b - c;
            int pa = Math.Abs(p - a);
            int pb = Math.Abs(p - b);
            int pc = Math.Abs(p - c);
            if (pa <= pb && pa <= pc) return a;
            if (pb <= pc) return b;
            return c;
        }
    }
}
