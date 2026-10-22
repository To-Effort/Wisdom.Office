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
        Jpeg = 2
    }

    /// <summary>内存图片：构造时自动嗅探 PNG/JPEG 编码并读出像素尺寸。An in-memory image: sniffs PNG/JPEG encoding and reads pixel dimensions at construction.</summary>
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

            Format = OfficeImageFormat.Unknown;
            ContentType = "application/octet-stream";
            FileExtension = ".bin";
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
