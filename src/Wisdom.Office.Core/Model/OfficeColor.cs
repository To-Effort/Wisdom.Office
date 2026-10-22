using System;
using System.Globalization;

namespace Wisdom.Office.Core
{
    /// <summary>统一文档模型使用的 RGBA 颜色值类型。An RGBA color value type used by the unified document model.</summary>
    public readonly struct OfficeColor : IEquatable<OfficeColor>
    {
        /// <summary>Alpha 通道（0 全透明 ~ 255 不透明）。The alpha channel (0 fully transparent to 255 opaque).</summary>
        public readonly byte A;
        /// <summary>红色通道。The red channel.</summary>
        public readonly byte R;
        /// <summary>绿色通道。The green channel.</summary>
        public readonly byte G;
        /// <summary>蓝色通道。The blue channel.</summary>
        public readonly byte B;

        /// <summary>创建不透明（Alpha=255）的颜色。Creates an opaque color (alpha 255).</summary>
        /// <param name="r">红色通道。The red channel.</param>
        /// <param name="g">绿色通道。The green channel.</param>
        /// <param name="b">蓝色通道。The blue channel.</param>
        public OfficeColor(byte r, byte g, byte b)
            : this(255, r, g, b)
        {
        }

        /// <summary>创建指定 Alpha 通道的颜色。Creates a color with the specified alpha channel.</summary>
        /// <param name="a">Alpha 通道。The alpha channel.</param>
        /// <param name="r">红色通道。The red channel.</param>
        /// <param name="g">绿色通道。The green channel.</param>
        /// <param name="b">蓝色通道。The blue channel.</param>
        public OfficeColor(byte a, byte r, byte g, byte b)
        {
            A = a;
            R = r;
            G = g;
            B = b;
        }

        /// <summary>纯黑。Pure black.</summary>
        public static readonly OfficeColor Black = new OfficeColor(0, 0, 0);
        /// <summary>纯白。Pure white.</summary>
        public static readonly OfficeColor White = new OfficeColor(255, 255, 255);
        /// <summary>全透明。Fully transparent.</summary>
        public static readonly OfficeColor Transparent = new OfficeColor(0, 0, 0, 0);
        /// <summary>红色。Red.</summary>
        public static readonly OfficeColor Red = new OfficeColor(255, 0, 0);
        /// <summary>绿色（0,128,0）。Green (0,128,0).</summary>
        public static readonly OfficeColor Green = new OfficeColor(0, 128, 0);
        /// <summary>纯蓝。Pure blue.</summary>
        public static readonly OfficeColor Blue = new OfficeColor(0, 0, 255);
        /// <summary>中灰。Medium gray.</summary>
        public static readonly OfficeColor Gray = new OfficeColor(128, 128, 128);
        /// <summary>浅灰（表格底纹常用）。Light gray (commonly used for table shading).</summary>
        public static readonly OfficeColor LightGray = new OfficeColor(217, 217, 217);

        /// <summary>输出 RGB 十六进制字符串（不含 Alpha），如 "1F3A5F"。Outputs the RGB hexadecimal string (without alpha), such as "1F3A5F".</summary>
        /// <returns>六位大写十六进制字符串。A six-character uppercase hexadecimal string.</returns>
        public string ToHex()
        {
            return R.ToString("X2", CultureInfo.InvariantCulture)
                 + G.ToString("X2", CultureInfo.InvariantCulture)
                 + B.ToString("X2", CultureInfo.InvariantCulture);
        }

        /// <summary>输出 ARGB 十六进制字符串（AARRGGBB）。Outputs the ARGB hexadecimal string (AARRGGBB).</summary>
        /// <returns>八位大写十六进制字符串。An eight-character uppercase hexadecimal string.</returns>
        public string ToArgbHex()
        {
            return A.ToString("X2", CultureInfo.InvariantCulture) + ToHex();
        }

        /// <summary>从 32 位 ARGB 整数（高位为 Alpha）构造颜色。Builds a color from a 32-bit ARGB integer (high byte is alpha).</summary>
        /// <param name="argb">ARGB 整数值。The ARGB integer value.</param>
        /// <returns>对应的颜色。The corresponding color.</returns>
        public static OfficeColor FromArgb(int argb)
        {
            return new OfficeColor(
                (byte)((argb >> 24) & 0xFF),
                (byte)((argb >> 16) & 0xFF),
                (byte)((argb >> 8) & 0xFF),
                (byte)(argb & 0xFF));
        }

        /// <summary>解析 HTML 风格颜色字符串；格式非法时抛出 <see cref="FormatException"/>。Parses an HTML-style color string; throws <see cref="FormatException"/> when the format is invalid.</summary>
        /// <param name="value">颜色字符串，如 "#RGB"、"RRGGBB"、"0xAARRGGBB"。The color string, such as "#RGB", "RRGGBB" or "0xAARRGGBB".</param>
        /// <returns>解析得到的颜色。The parsed color.</returns>
        public static OfficeColor FromHtml(string value)
        {
            OfficeColor color;
            if (!TryParse(value, out color))
            {
                throw new FormatException("Invalid color value: " + value);
            }
            return color;
        }

        /// <summary>尝试解析颜色字符串：支持 3/6/8 位十六进制，可带 "#" 或 "0x" 前缀；6 位视为 RGB、8 位视为 ARGB。Tries to parse a color string, accepting 3-, 6- or 8-digit hex with an optional "#" or "0x" prefix; 6 digits are read as RGB and 8 digits as ARGB.</summary>
        /// <param name="value">待解析的字符串。The string to parse.</param>
        /// <param name="color">解析成功时输出颜色，失败时输出黑色。Outputs the parsed color on success, or black on failure.</param>
        /// <returns>解析是否成功。Whether parsing succeeded.</returns>
        public static bool TryParse(string value, out OfficeColor color)
        {
            color = Black;
            if (string.IsNullOrEmpty(value))
            {
                return false;
            }

            string hex = value.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal))
            {
                hex = hex.Substring(1);
            }
            else if (hex.StartsWith("0x", StringComparison.OrdinalIgnoreCase))
            {
                hex = hex.Substring(2);
            }

            if (hex.Length == 3)
            {
                hex = new string(new[]
                {
                    hex[0], hex[0], hex[1], hex[1], hex[2], hex[2]
                });
            }

            uint number;
            if (hex.Length == 6 && uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out number))
            {
                color = new OfficeColor((byte)(number >> 16), (byte)(number >> 8), (byte)number);
                return true;
            }

            if (hex.Length == 8 && uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out number))
            {
                color = new OfficeColor((byte)(number >> 24), (byte)(number >> 16), (byte)(number >> 8), (byte)number);
                return true;
            }

            return false;
        }

        /// <summary>逐通道比较两个颜色是否相等。Compares two colors channel by channel for equality.</summary>
        /// <param name="other">待比较的颜色。The color to compare with.</param>
        /// <returns>四个通道全部相同时返回 true。True when all four channels are equal.</returns>
        public bool Equals(OfficeColor other)
        {
            return A == other.A && R == other.R && G == other.G && B == other.B;
        }

        /// <summary>对象相等比较：仅当 <paramref name="obj"/> 为 <see cref="OfficeColor"/> 时逐通道比较。Object equality comparison; channels are compared only when <paramref name="obj"/> is an <see cref="OfficeColor"/>.</summary>
        /// <param name="obj">待比较对象。The object to compare with.</param>
        /// <returns>是否相等。Whether the two are equal.</returns>
        public override bool Equals(object obj)
        {
            return obj is OfficeColor && Equals((OfficeColor)obj);
        }

        /// <summary>组合 A/R/G/B 计算稳定哈希值。Combines A/R/G/B into a stable hash value.</summary>
        /// <returns>哈希码。The hash code.</returns>
        public override int GetHashCode()
        {
            unchecked
            {
                int hash = A;
                hash = (hash * 397) ^ R;
                hash = (hash * 397) ^ G;
                hash = (hash * 397) ^ B;
                return hash;
            }
        }

        /// <summary>颜色值相等运算符。The color-value equality operator.</summary>
        /// <param name="left">左侧颜色。The left color.</param>
        /// <param name="right">右侧颜色。The right color.</param>
        /// <returns>是否相等。Whether the two are equal.</returns>
        public static bool operator ==(OfficeColor left, OfficeColor right)
        {
            return left.Equals(right);
        }

        /// <summary>颜色值不等运算符。The color-value inequality operator.</summary>
        /// <param name="left">左侧颜色。The left color.</param>
        /// <param name="right">右侧颜色。The right color.</param>
        /// <returns>是否不相等。Whether the two differ.</returns>
        public static bool operator !=(OfficeColor left, OfficeColor right)
        {
            return !left.Equals(right);
        }

        /// <summary>输出 "#RRGGBB" 形式的字符串。Outputs a string in "#RRGGBB" form.</summary>
        /// <returns>带井号前缀的十六进制颜色。The hexadecimal color with a leading hash sign.</returns>
        public override string ToString()
        {
            return "#" + ToHex();
        }
    }
}
