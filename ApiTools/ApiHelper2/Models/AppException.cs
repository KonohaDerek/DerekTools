using ApiHelper.Models;
using System;

namespace ApiHelper2.Models
{
    public class AppException : Exception
    {
        /// <summary>
        /// 系統應用程式的錯誤訊息。
        /// </summary>
        /// <param name="code">錯誤代碼</param>
        public AppException(ResultCodeType code) : this(code, null, null, false) { }

        /// <summary>
        /// 系統應用程式的錯誤訊息。
        /// </summary>
        /// <param name="code">錯誤代碼</param>
        /// <param name="ex">內部錯誤</param>
        public AppException(ResultCodeType code, Exception ex) : this(code, null, ex, false) { }

        /// <summary>
        /// 系統應用程式的錯誤訊息。
        /// </summary>
        /// <param name="code">錯誤代碼</param>
        /// <param name="mesg">附加的錯誤訊息</param>
        public AppException(ResultCodeType code, string mesg) : this(code, mesg, null, false) { }

        /// <summary>
        /// 系統應用程式的錯誤訊息。
        /// </summary>
        /// <param name="code">錯誤代碼</param>
        /// <param name="mesg">附加的錯誤訊息</param>
        /// <param name="ex">內部錯誤內容</param>
        public AppException(ResultCodeType code, string mesg, Exception ex) : this(code, mesg, ex, false) { }

        /// <summary>
        /// 系統應用程式的錯誤訊息。
        /// </summary>
        /// <remarks>
        /// 當返回的錯誤訊息(ResultDesc)已包含錯誤代碼，則適用這個多型來建立錯誤訊息避免代碼會多重顯示。
        /// </remarks>
        /// <param name="code">錯誤代碼</param>
        /// <param name="mesg">附加的錯誤訊息</param>
        /// <param name="ex">內部錯誤內容</param>
        /// <param name="isDirectOut">是否直接輸出訊息(不做代碼合併的處理)</param>
        public AppException(ResultCodeType code, string mesg, Exception ex, bool isDirectOut) : base(isDirectOut ? mesg : GetResultDesc(code, mesg, ex, true), ex)
        {
            InnerResultCode = code;                                                 // 錯誤代碼(實體)
            ResultCode = code.ToCodeFormat();                                       // 錯誤代碼(數值字串)
            ResultDesc = isDirectOut ? mesg : GetResultDesc(code, mesg, ex, true);  // 錯誤描述訊息(含代碼)
        }

        /// <summary>
        /// 取得錯誤代碼(實體)
        /// </summary>
        /// <returns></returns>

        public ResultCodeType GetResultCode()
        {
            return InnerResultCode;
        }

        /// <summary>
        /// 錯誤代碼(實體)
        /// </summary>
        protected virtual ResultCodeType InnerResultCode { get; private set; }

        /// <summary>
        /// 錯誤代碼
        /// </summary>
        public virtual string ResultCode { get; private set; }

        /// <summary>
        /// 結果描述訊息
        /// </summary>
        public virtual string ResultDesc { get; private set; }

        /// <summary>
        /// 額外資訊
        /// </summary>
        public virtual string Additional { get; set; }

        /// <summary>
        /// 取得系統錯誤代碼說明文
        ///     摘要: 預設格式 "交易錯誤(FE0001)"
        /// </summary>
        /// <param name="code">代碼值，同時也會帶出代碼訊息做為錯誤訊息</param>
        /// <param name="mesg">補充的錯誤訊息</param>
        /// <param name="ex">繼承的錯誤內容</param>
        /// <returns></returns>
        public static string GetResultDesc(ResultCodeType code, string mesg = null, Exception ex = null)
        {
            return GetResultDesc(code, mesg, ex, true);
        }

        /// <summary>
        /// 取得系統錯誤代碼說明文(不含代碼)
        ///     摘要: 預設格式 "交易錯誤"
        /// </summary>
        /// <param name="code">代碼值，同時也會帶出代碼訊息做為錯誤訊息</param>
        /// <param name="mesg">補充的錯誤訊息</param>
        /// <param name="ex">繼承的錯誤內容</param>
        /// <returns></returns>
        public static string GetResultDescNotCode(ResultCodeType code, string mesg = null, Exception ex = null)
        {
            return GetResultDesc(code, mesg, ex, false);
        }

        /// <summary>
        /// 取得系統錯誤代碼說明文
        ///     摘要: 預設格式 "交易錯誤(FE0001)"
        /// </summary>
        /// <param name="code">代碼值，同時也會帶出代碼訊息做為錯誤訊息</param>
        /// <param name="mesg">補充的錯誤訊息</param>
        /// <param name="ex">繼承的錯誤內容</param>
        /// <param name="hasCode">格式包含代碼</param>
        /// <returns></returns>
        public static string GetResultDesc(ResultCodeType code, string mesg, Exception ex, bool hasCode)
        {
            var desc = code.GetEnumDescription();
            var sb = new System.Text.StringBuilder();

            // 代碼描述加入到錯誤訊息
            sb.Append(desc);

            // 系統未預期錯誤(FE9999)，會帶入 InnerException 的錯誤編號
            if (code == ResultCodeType.gSystemFail && ex != null)
            {
                int errHResult = (ex is AppException) ? (int)((AppException)ex).InnerResultCode : ex.HResult;
                sb.Append(string.Format("(0x{0:X})", errHResult));
            }

            // 有額外的附加訊息
            if (string.IsNullOrWhiteSpace(mesg) == false)
            {
                sb.AppendFormat(":{0}", mesg);
            }

            // 訊息是否帶入代碼值
            if (hasCode)
            {
                sb.AppendFormat("(FE{0})", code.ToCodeFormat());
            }

            return sb.ToString();
        }
    }
}
