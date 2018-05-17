using System;
using System.ComponentModel;
using System.Reflection;

namespace ApiHelper.Models
{
    /// <summary>
    ///叫用 Web 服務(API)的參數傳遞方式。
    /// </summary>
    public enum RequestMode
    {
        ByURL,
        ByJSON,
        ByXML
    }

    public enum RequestMethodMode
    {
        GET,
        POST
    }

    /// <summary>
    /// 叫用 Web 服務(API)的回覆內容格式。
    /// </summary>
    public enum ResponseMode
    {
        ByStream,
        ByText,
        ByJSON,
        ByXML,
    }

    /// 簽名類型。
    /// </summary>
    public enum SignTypes
    {
        /// <summary>
        /// 無簽章。
        /// </summary>
        None = 0,
        MD5,
        SHA256,
    }

    public enum ResultCodeType : int
    {
        /// <summary>
        /// 成功
        /// </summary>
        [Description("成功")]
        gSuccess = 0,

        /// <summary>
        /// 失敗
        /// </summary>
        [Description("失敗")]
        gFailed = 1,

        /// <summary>
        /// 連線失敗
        /// </summary>
        [Description("連線失敗")]
        ConnectFail = 2000,

        /// <summary>
        /// 連線功能不存在或未合法 
        /// </summary>
        [Description("連線功能不存在或不允許")]
        ConnectNotFoundOrNotAllowed = 2001,

        /// <summary>
        /// 未預期的系統錯誤
        /// </summary>
        [Description("未預期的系統錯誤")]
        gSystemFail = 9999,
    }

    public static class Extensioin
    {
        /// <summary>
        /// 取得結果代碼完整長度(4碼數字) 
        /// </summary>
        /// <param name="code">成果代碼</param>
        /// <returns>完整長度</returns>
        public static string ToCodeFormat(this ResultCodeType code)
        {
            return Convert.ToInt32(code).ToString().PadLeft(4, '0');
        }

        /// <summary>
        /// 取得Enum描述
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static string GetEnumDescription(this Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(
                typeof(DescriptionAttribute),
                false);

            if (attributes != null &&
                attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }
    }

}