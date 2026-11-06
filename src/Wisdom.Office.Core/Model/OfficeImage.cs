using System;
using System.IO;

namespace Wisdom.Office.Core
{
    /// <summary>可自动识别的图片编码格式。An image encoding format that can be auto-detected.</summary>
    public enum OfficeImageFormat
    {
        /// <summary>无法识别的格式。An unrecognized format.</summary>
        Unknown = 0,
        /// <summary>PNG 格式。The PNG format.</summary>
        Png = 1,
        /// <summary>JPEG 格式。The JPEG format.</summary>
        Jpeg = 2,
        /// <summary>GIF 格式。The GIF format.</summary>
        Gif = 3,
        /// <summary>BMP（位图）格式。The BMP (device-independent bitmap) format.</summary>
        Bmp = 4,
        /// <summary>TIFF 格式。The TIFF format.</summary>
        Tiff = 5,
        /// <summary>Windows 图元文件。The Windows metafile format.</summary>
        Emf = 6,
        /// <summary>Windows 位图图元文件。The Windows bitmap-metafile format.</summary>
        Wmf = 7
    }

    /// <summary>内存图片：构造时自动嗅探编码格式与扩展名并尽量读出像素尺寸；位图栅格化只支持 PNG 与 JPEG。
    /// An in-memory image: sniffs the encoding, extension and pixel dimensions at construction; only PNG and JPEG
    /// can be rasterised, the other formats are carried through documents unchanged.</summary>
    public sealed class OfficeImage
    {
        /// <summary>原始图片字节。The raw image bytes.</summary>
        public byte[] Data { get; private set; }
        /// <summary>识别出的编码格式。The detected encoding format.</summary>
        public OfficeImageFormat Format { get; private set; }
        /// <summary>MIME 内容类型。The MIME content type.</summary>
        public string ContentType { get; private set; }
        /// <summary>像素宽度，0 表示未能读出。The pixel width; 0 means it could not be read.</summary>
        public int PixelWidth { get; private set; }
        /// <summary>像素高度，0 表示未能读出。The pixel height; 0 means it could not be read.</summary>
        public int PixelHeight { get; private set; }
        /// <summary>建议的文件扩展名（含点），如 ".png"。The suggested file extension (with the dot), such as ".png".</summary>
        public string FileExtension { get; private set; }

        /// <summary>从字节数组构造图片并自动识别格式与尺寸。Builds an image from a byte array and auto-detects its format and dimensions.</summary>
        /// <param name="data">图片字节，不能为 null。The image bytes; must not be null.</param>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> 为 null。Is null.</exception>
        public OfficeImage(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            Data = data;
            Detect();
        }

        /// <summary>从磁盘文件读取字节并构造图片。Reads bytes from a disk file and builds an image.</summary>
        /// <param name="path">图片文件路径。The image file path.</param>
        /// <returns>识别后的图片实例。The detected image instance.</returns>
        public static OfficeImage FromFile(string path)
        {
            return new OfficeImage(File.ReadAllBytes(path));
        }

        /// <summary>按文件头魔数嗅探图片格式，设置 MIME 类型与扩展名，并尝试读取像素尺寸。Sniffs the image format from the file-header magic bytes, sets the MIME type and extension, and tries to read pixel dimensions.</summary>
        private void Detect()
        {
            if (Data.Length > 8
                && Data[0] == 0x89 && Data[1] == 0x50 && Data[2] == 0x4E && Data[3] == 0x47
                && Data[4] == 0x0D && Data[5] == 0x0A && Data[6] == 0x1A && Data[7] == 0x0A)
            {
                Format = OfficeImageFormat.Png;
                ContentType = "image/png";
                FileExtension = ".png";
                int width;
                int height;
                if (TryReadPngSize(Data, out width, out height))
                {
                    PixelWidth = width;
                    PixelHeight = height;
                }
                return;
            }

            if (Data.Length > 3 && Data[0] == 0xFF && Data[1] == 0xD8)
            {
                Format = OfficeImageFormat.Jpeg;
                ContentType = "image/jpeg";
                FileExtension = ".jpg";
                int width;
                int height;
                if (TryReadJpegSize(Data, out width, out height))
                {
                    PixelWidth = width;
                    PixelHeight = height;
                }
                return;
            }

            if (IsGif(Data))
            {
                Format = OfficeImageFormat.Gif;
                ContentType = "image/gif";
                FileExtension = ".gif";
                PixelWidth = ReadUInt16LittleEndian(Data, 6);
                PixelHeight = ReadUInt16LittleEndian(Data, 8);
                return;
            }

            if (IsBmp(Data))
            {
                Format = OfficeImageFormat.Bmp;
                ContentType = "image/bmp";
                FileExtension = ".bmp";
                int headerSize = ReadInt32LittleEndian(Data, 14);
                if (headerSize == 12)
                {
                    PixelWidth = ReadUInt16LittleEndian(Data, 18);
                    PixelHeight = ReadUInt16LittleEndian(Data, 20);
                }
                else
                {
                    PixelWidth = ReadInt32LittleEndian(Data, 18);
                    PixelHeight = Math.Abs(ReadInt32LittleEndian(Data, 22));
                }
                if (PixelWidth < 0 || PixelHeight < 0 || PixelWidth > 65535 || PixelHeight > 65535)
                {
                    PixelWidth = 0;
                    PixelHeight = 0;
                }
                return;
            }

            if (IsTiff(Data))
            {
                // TIFF 尺寸要遍历 IFD 标签，这里只登记格式：字节原样带出文档，栅格化交给 Word 与 Excel。
                // Reading a TIFF size means walking the IFD tags, so only the format is recorded here: the bytes
                // travel through the documents unchanged and Word or Excel rasterises them.
                Format = OfficeImageFormat.Tiff;
                ContentType = "image/tiff";
                FileExtension = ".tiff";
                return;
            }

            if (IsEmf(Data))
            {
                Format = OfficeImageFormat.Emf;
                ContentType = "image/x-emf";
                FileExtension = ".emf";
                return;
            }

            if (IsWmf(Data))
            {
                Format = OfficeImageFormat.Wmf;
                ContentType = "image/x-wmf";
                FileExtension = ".wmf";
                return;
            }

            Format = OfficeImageFormat.Unknown;
            ContentType = "application/octet-stream";
            FileExtension = ".bin";
        }

        /// <summary>嗅探 GIF：魔数 GIF87a/GIF89a，尺寸取自逻辑屏幕描述符（偏移 6~9，小端）。 Sniffs GIF: the GIF87a/GIF89a magic, with the size taken from the logical screen descriptor (offsets 6-9, little-endian).</summary>
        /// <param name="data">图片字节。The image bytes.</param>
        /// <returns>识别成功时返回 true。True when the bytes are a GIF.</returns>
        private static bool IsGif(byte[] data)
        {
            return data.Length > 9 && data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x38;
        }

        /// <summary>嗅探 BMP：BM 魔数加 BITMAPINFOHEADER/BITMAPCOREHEADER 的头长度。 Sniffs BMP: the BM magic plus a BITMAPINFOHEADER or BITMAPCOREHEADER header length.</summary>
        /// <param name="data">图片字节。The image bytes.</param>
        /// <returns>识别成功时返回 true。True when the bytes are a BMP.</returns>
        private static bool IsBmp(byte[] data)
        {
            if (data.Length < 26 || data[0] != 0x42 || data[1] != 0x4D) return false;
            int headerSize = ReadInt32LittleEndian(data, 14);
            return headerSize == 12 || headerSize >= 40;
        }

        /// <summary>嗅探 TIFF：II/MM 字节序标记加魔数 42。 Sniffs TIFF: the II/MM byte-order mark plus the 42 magic.</summary>
        /// <param name="data">图片字节。The image bytes.</param>
        /// <returns>识别成功时返回 true。True when the bytes are a TIFF.</returns>
        private static bool IsTiff(byte[] data)
        {
            if (data.Length < 8) return false;
            if (data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00) return true;
            return data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A;
        }

        /// <summary>嗅探 EMF：首记录类型 1 且偏移 40 处是 " EMF" 签名。 Sniffs EMF: the first record type is 1 and the " EMF" signature sits at offset 40.</summary>
        /// <param name="data">图片字节。The image bytes.</param>
        /// <returns>识别成功时返回 true。True when the bytes are an EMF.</returns>
        private static bool IsEmf(byte[] data)
        {
            if (data.Length < 44) return false;
            if (ReadInt32LittleEndian(data, 0) != 1) return false;
            return data[40] == 0x20 && data[41] == 0x45 && data[42] == 0x4D && data[43] == 0x46;
        }

        /// <summary>嗅探 WMF：可放置头（D7 CDC6 9A）或磁盘元文件的 "PM" 魔数。 Sniffs WMF: the placeable header (D7 CDC6 9A) or the "PM" magic of a disk metafile.</summary>
        /// <param name="data">图片字节。The image bytes.</param>
        /// <returns>识别成功时返回 true。True when the bytes are a WMF.</returns>
        private static bool IsWmf(byte[] data)
        {
            if (data.Length < 20) return false;
            if (ReadInt32LittleEndian(data, 0) == unchecked((int)0x9AC6CDD7)) return true;
            return data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x01 && data[3] == 0x00
                && data[18] == 0x50 && data[19] == 0x4D;
        }

        /// <summary>读取小端 32 位整数。Reads a little-endian 32-bit integer.</summary>
        /// <param name="data">字节缓冲。The byte buffer.</param>
        /// <param name="offset">起始偏移（调用方保证可读）。The offset; the caller guarantees it is readable.</param>
        /// <returns>整数值。The value.</returns>
        private static int ReadInt32LittleEndian(byte[] data, int offset)
        {
            return data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24);
        }

        /// <summary>读取小端 16 位无符号整数。Reads a little-endian 16-bit unsigned integer.</summary>
        /// <param name="data">字节缓冲。The byte buffer.</param>
        /// <param name="offset">起始偏移（调用方保证可读）。The offset; the caller guarantees it is readable.</param>
        /// <returns>整数值。The value.</returns>
        private static int ReadUInt16LittleEndian(byte[] data, int offset)
        {
            return data[offset] | (data[offset + 1] << 8);
        }

        /// <summary>按 96 DPI 换算的默认显示宽度（像素×72/96 磅）；无尺寸信息时回退 200 磅。The default display width converted at 96 DPI (pixels × 72/96 points); falls back to 200 points when no size is known.</summary>
        public double DefaultWidthPoints
        {
            get { return PixelWidth > 0 ? PixelWidth * 72.0 / 96.0 : 200; }
        }

        /// <summary>按 96 DPI 换算的默认显示高度（像素×72/96 磅）；无尺寸信息时回退 150 磅。The default display height converted at 96 DPI (pixels × 72/96 points); falls back to 150 points when no size is known.</summary>
        public double DefaultHeightPoints
        {
            get { return PixelHeight > 0 ? PixelHeight * 72.0 / 96.0 : 150; }
        }

        /// <summary>从 PNG 文件头 IHDR 块（偏移 16~23，大端）读取像素宽高。Reads pixel width and height from the PNG IHDR block (offset 16-23, big-endian).</summary>
        /// <param name="data">PNG 字节。The PNG bytes.</param>
        /// <param name="width">输出宽度，失败时为 0。Outputs the width; 0 on failure.</param>
        /// <param name="height">输出高度，失败时为 0。Outputs the height; 0 on failure.</param>
        /// <returns>成功读出且尺寸有效时为 true。True when the size was read and is valid.</returns>
        public static bool TryReadPngSize(byte[] data, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (data == null || data.Length < 24)
            {
                return false;
            }
            if (data[0] != 0x89 || data[1] != 0x50 || data[2] != 0x4E || data[3] != 0x47)
            {
                return false;
            }
            width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
            height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
            return width > 0 && height > 0;
        }

        /// <summary>扫描 JPEG 段标记，从 SOFn 帧段（0xC0~0xCF 中排除 0xC4/0xC8/0xCC）读取像素宽高。Scans JPEG segment markers and reads pixel width/height from an SOFn frame segment (0xC0-0xCF excluding 0xC4/0xC8/0xCC).</summary>
        /// <param name="data">JPEG 字节。The JPEG bytes.</param>
        /// <param name="width">输出宽度，失败时为 0。Outputs the width; 0 on failure.</param>
        /// <param name="height">输出高度，失败时为 0。Outputs the height; 0 on failure.</param>
        /// <returns>成功读出且尺寸有效时为 true。True when the size was read and is valid.</returns>
        public static bool TryReadJpegSize(byte[] data, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (data == null || data.Length < 4 || data[0] != 0xFF || data[1] != 0xD8)
            {
                return false;
            }

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
                if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7))
                {
                    continue;
                }
                if (position + 1 >= data.Length)
                {
                    break;
                }
                int length = (data[position] << 8) | data[position + 1];
                if (length < 2)
                {
                    break;
                }
                if ((marker >= 0xC0 && marker <= 0xC3)
                    || (marker >= 0xC5 && marker <= 0xC7)
                    || (marker >= 0xC9 && marker <= 0xCB)
                    || (marker >= 0xCD && marker <= 0xCF))
                {
                    if (position + 7 >= data.Length)
                    {
                        break;
                    }
                    height = (data[position + 3] << 8) | data[position + 4];
                    width = (data[position + 5] << 8) | data[position + 6];
                    return width > 0 && height > 0;
                }
                position += length;
            }
            return false;
        }
    }
}
