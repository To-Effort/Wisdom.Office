using System;

namespace Wisdom.Office.Core
{
    /// <summary>Office 度量单位换算工具。Utility for converting between Office measurement units.</summary>
    public static class UnitConverter
    {
        /// <summary>每英寸的磅数（72）。The number of points per inch (72).</summary>
        public const double PointsPerInch = 72.0;
        /// <summary>每磅的缇数（20）。The number of twips per point (20).</summary>
        public const double TwipsPerPoint = 20.0;
        /// <summary>每磅的 EMU 数（12700）。The number of English Metric Units per point (12700).</summary>
        public const double EmusPerPoint = 12700.0;
        /// <summary>每英寸的像素数（96，即 96 DPI）。The number of pixels per inch (96, i.e. 96 DPI).</summary>
        public const double PixelsPerInch = 96.0;

        /// <summary>将缇（twip）换算为磅。Converts twips to points.</summary>
        /// <param name="twips">缇数值。The twip value.</param>
        /// <returns>等值的磅数。The equivalent value in points.</returns>
        public static double TwipsToPoints(double twips)
        {
            return twips / TwipsPerPoint;
        }

        /// <summary>将磅换算为缇（twip）。Converts points to twips.</summary>
        /// <param name="points">磅数值。The point value.</param>
        /// <returns>等值的缇数。The equivalent value in twips.</returns>
        public static double PointsToTwips(double points)
        {
            return points * TwipsPerPoint;
        }

        /// <summary>将 EMU（English Metric Unit）换算为磅。Converts English Metric Units to points.</summary>
        /// <param name="emu">EMU 数值。The EMU value.</param>
        /// <returns>等值的磅数。The equivalent value in points.</returns>
        public static double EmuToPoints(double emu)
        {
            return emu / EmusPerPoint;
        }

        /// <summary>将磅换算为 EMU（English Metric Unit）。Converts points to English Metric Units.</summary>
        /// <param name="points">磅数值。The point value.</param>
        /// <returns>等值的 EMU 数。The equivalent value in EMUs.</returns>
        public static double PointsToEmu(double points)
        {
            return points * EmusPerPoint;
        }

        /// <summary>按 96 DPI 将像素换算为磅。Converts pixels to points assuming 96 DPI.</summary>
        /// <param name="pixels">像素数值。The pixel value.</param>
        /// <returns>等值的磅数。The equivalent value in points.</returns>
        public static double PixelsToPoints(double pixels)
        {
            return pixels * PointsPerInch / PixelsPerInch;
        }

        /// <summary>按 96 DPI 将磅换算为像素。Converts points to pixels assuming 96 DPI.</summary>
        /// <param name="points">磅数值。The point value.</param>
        /// <returns>等值的像素数。The equivalent value in pixels.</returns>
        public static double PointsToPixels(double points)
        {
            return points * PixelsPerInch / PointsPerInch;
        }

        /// <summary>将英寸换算为磅。Converts inches to points.</summary>
        /// <param name="inches">英寸数值。The inch value.</param>
        /// <returns>等值的磅数。The equivalent value in points.</returns>
        public static double InchesToPoints(double inches)
        {
            return inches * PointsPerInch;
        }

        /// <summary>将磅换算为英寸。Converts points to inches.</summary>
        /// <param name="points">磅数值。The point value.</param>
        /// <returns>等值的英寸数。The equivalent value in inches.</returns>
        public static double PointsToInches(double points)
        {
            return points / PointsPerInch;
        }

        /// <summary>四舍五入到最接近的整数，中点远离零取整。Rounds to the nearest integer with midpoints rounded away from zero.</summary>
        /// <param name="value">待取整的值。The value to round.</param>
        /// <returns>取整后的整数。The rounded integer.</returns>
        public static int RoundToInt(double value)
        {
            return (int)Math.Round(value, MidpointRounding.AwayFromZero);
        }

        /// <summary>将值限制到 [min, max] 区间内。Clamps a value into the inclusive range from min to max.</summary>
        /// <param name="value">待限制的值。The value to clamp.</param>
        /// <param name="min">允许的最小值。The minimum allowed value.</param>
        /// <param name="max">允许的最大值。The maximum allowed value.</param>
        /// <returns>限制后的值。The clamped value.</returns>
        public static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
    }

    /// <summary>Excel 序列日期与 .NET DateTime 的互转工具。Utility for converting between Excel serial dates and .NET DateTime.</summary>
    public static class DateUtils
    {
        /// <summary>将 Excel 序列日期数转换为 DateTime；1904 日期系统需叠加 1462 天的纪元偏移。Converts an Excel serial date number to a DateTime; the 1904 date system requires a 1462-day epoch offset.</summary>
        /// <param name="serial">Excel 序列日期数。The Excel serial date number.</param>
        /// <param name="date1904">是否使用 1904 日期系统。Whether the 1904 date system is in use.</param>
        /// <returns>对应的日期时间。The corresponding date and time.</returns>
        public static DateTime ExcelSerialToDateTime(double serial, bool date1904)
        {
            if (date1904)
            {
                return DateTime.FromOADate(serial + 1462.0);
            }
            if (serial < 1)
            {
                return DateTime.FromOADate(serial);
            }
            return DateTime.FromOADate(serial);
        }

        /// <summary>将 DateTime 转换为 Excel 序列日期数；1904 日期系统需减去 1462 天的纪元偏移。Converts a DateTime to an Excel serial date number; the 1904 date system subtracts a 1462-day epoch offset.</summary>
        /// <param name="value">日期时间值。The date and time value.</param>
        /// <param name="date1904">是否使用 1904 日期系统。Whether the 1904 date system is in use.</param>
        /// <returns>Excel 序列日期数。The Excel serial date number.</returns>
        public static double DateTimeToExcelSerial(DateTime value, bool date1904)
        {
            double serial = value.ToOADate();
            if (date1904)
            {
                serial -= 1462.0;
            }
            return serial;
        }

        /// <summary>从 Excel 序列日期数的小数部分提取当天秒数（0~86400）。Extracts the seconds-of-day (0 to 86400) from the fractional part of an Excel serial date number.</summary>
        /// <param name="serial">Excel 序列日期数。The Excel serial date number.</param>
        /// <returns>当天的秒数。The number of seconds within the day.</returns>
        public static int ExcelSerialToTimeSeconds(double serial)
        {
            double fraction = serial - Math.Floor(serial);
            return (int)Math.Round(fraction * 86400.0);
        }
    }
}
