using System.Net;

namespace ApiHelper2.Models
{
    /// <summary>
    /// API 回傳的結果資訊
    /// </summary>
    public class APIResult<T>
    {
        /// <summary>
        /// 回傳的Response 標頭
        /// </summary>
        public WebHeaderCollection ResponseHeaders;

        /// <summary>
        /// 回傳的 Response 字串
        /// </summary>
        public T Response { get; internal set; }
    }
}