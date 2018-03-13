using System;

namespace ExcelHelper
{
    public static class Extentions
    {
        /// <summary>
        /// 文字轉換日期
        /// </summary>
        /// <param name="str"></param>
        /// <param name="timeFormat"></param>
        /// <returns></returns>
        public static DateTime ToDateTime(this string str , string timeFormat = "yyyy/MM/dd HH:mm:ss")
        {
            IFormatProvider culture = new System.Globalization.CultureInfo("zh-TW", true);
            DateTime time = DateTime.ParseExact(str, timeFormat, culture);
            return time;
        }

        /// <summary>
        /// 轉換日期
        /// </summary>
        /// <param name="unixTime"></param>
        /// <returns></returns>
        public static DateTime ToDateTime(this double unixTime)
        {
            return DateTime.FromOADate(unixTime);
        }

    }
}