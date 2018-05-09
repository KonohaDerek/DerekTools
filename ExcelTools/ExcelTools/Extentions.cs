using System;
using System.Collections.Generic;

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
        /// <summary>
        /// 取得字典中的資料
        /// </summary>
        /// <param name="dic">字典</param>
        /// <param name="key">主鍵值</param>
        /// <param name="defaultValue">預設值(空字串)</param>
        /// <returns>資料內容</returns>
        public static string Get(this Dictionary<string, string> dic, string key, string defaultValue = "")
        {
            return dic.Get<string>(key, defaultValue);
        }

        /// <summary>
        /// 取得字典中的資料
        /// </summary>
        /// <typeparam name="T">資料類別</typeparam>
        /// <param name="dic">字典</param>
        /// <param name="key">主鍵值</param>
        /// <param name="defaultValue">預設值</param>
        /// <returns>資料內容</returns>
        public static T Get<T>(this Dictionary<string, T> dic, string key, T defaultValue = default(T))
        {
            if (dic.ContainsKey(key) == false) return defaultValue;
            var value = dic[key];
            return value;
        }

        /// <summary>
        /// 更新或新增字典中的資料
        /// </summary>
        /// <param name="dic">字典</param>
        /// <param name="key">主鍵值</param>
        /// <param name="value">資料內容</param>
        /// <param name="isEmptyNotSet">是否不保存空值</param>
        /// <returns>資料內容</returns>
        public static void Set(this Dictionary<string, string> dic, string key, string value, bool isEmptyNotSet = true)
        {
            // 指定不保存空值
            if (isEmptyNotSet && string.IsNullOrWhiteSpace(value)) return;

            // 檢查字典有值則取代，否則新增參數
            if (dic.ContainsKey(key))
            {
                dic[key] = value;
            }
            else
            {
                dic.Add(key, value);
            }
        }
    }
}