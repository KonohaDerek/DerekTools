using ApiHelper.Models;
using ApiHelper2;
using ApiHelper2.Models;
using System;
using System.IO;
using System.Net;
using System.Text;

namespace ApiHelper
{
    public abstract class ApiBase
    {
        /// <summary>
        /// 建構基底
        /// </summary>
        public ApiBase()
        {

        }

        #region Properties
        /// <summary>
        /// 呼叫支付機構的服務位址。
        /// </summary>
        public virtual string ServiceUrl { get; set; }


        /// <summary>
        /// 連線逾時(秒)，可由各錢包再重新定義。
        /// </summary>
        protected virtual int ConnectTimeOut
        {
            get
            {
                return 10;
            }
        }


        #endregion

        #region 交易發送 POST 方法

        /// <summary>
        /// 發送請求至錢包商的服務。
        /// </summary>
        /// <param name="walletRequest">錢包請求服務的參數集</param>
        /// <param name="postURL">POST的位址，若無指定則會發送到ServiceUrl的位置</param>
        /// <param name="bankPayID">銀行付款id</param>
        /// <param name="bankRefundID">銀行退款id</param>
        /// <param name="traceLog">交易log</param>
        /// <returns></returns>
        protected APIResult<T> PostToAPI<T>(RequestBase apiRequest, string postURL = null, long? bankPayID = null, long? bankRefundID = null)
        {
            var result = new APIResult<T>();
            var mode = apiRequest.Mode;
            var requestUrl = string.Empty;
            string responseData = null; // POST 無回傳值應該回寫 NULL
            try
            {
                // 取得傳送的參數值(字串)。
                var requestContent = apiRequest.ToString();
                requestUrl = (string.IsNullOrWhiteSpace(postURL)) ? ServiceUrl : postURL;

                // 忽略SSL檢核憑證
                if (requestUrl.StartsWith("https"))
                {
                    if (ServicePointManager.ServerCertificateValidationCallback == null)
                    {
                        ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                    }
                }

                // GET 模式，將會把參數加到 URL 中。
                if (apiRequest.Method == RequestMethodMode.GET)
                {
                    requestUrl = string.Format("{0}?{1}", requestUrl, apiRequest);
                    requestContent = requestUrl;
                }

                // 建立交易請求連線
                var request = WebRequest.CreateHttp(requestUrl);
                request.Method = Convert.ToString(apiRequest.Method);    // 叫用的模式由錢包決定
                request.Timeout = ConnectTimeOut * 1000;                    // 逾時時間，以毫秒為單位

                apiRequest.SetHeader(request.Headers);                   // 加入Header資訊，當傳送的資料有需要追加Header值則需要實作此方法   

                #region 建立交易請求，(GET/POST)填入的參數不同

                // GET 模式，固定為URL的 Content-Type
                if (apiRequest.Method == RequestMethodMode.GET)
                {
                    request.ContentType = "application/x-www-form-urlencoded";
                    requestContent = WebUtility.UrlDecode(requestContent);
                }
                // POST 模式，設定Content-type & 寫入Body資料流 & 內文長度的資訊。
                if (apiRequest.Method == RequestMethodMode.POST)
                {
                    switch (mode)
                    {
                        case RequestMode.ByURL:
                            request.ContentType = "application/x-www-form-urlencoded";
                            requestContent = WebUtility.UrlDecode(requestContent);
                            break;
                        case RequestMode.ByJSON:
                            request.ContentType = "application/json";
                            break;
                        case RequestMode.ByXML:
                            request.ContentType = "text/xml; encoding='utf-8'";
                            break;
                    }
                }

                #endregion

                var httpRequest = default(HttpWebRequest);

                    #region 交易直接傳送

                    httpRequest = request;

                    // POST模式，發送資料前建立連繫通道並且將內文寫入資料流
                    if (apiRequest.Method == RequestMethodMode.POST)
                    {
                        var data = Encoding.UTF8.GetBytes(requestContent);
                        request.ContentLength = data.Length;

                        using (var requestStream = request.GetRequestStream())
                        {
                            requestStream.Write(data, 0, data.Length);
                        }
                    }
                    #endregion
                #region 發送交易並接收結果
                using (var response = (HttpWebResponse)httpRequest.GetResponse())
                {
                    // 如果 StatusCode 在範圍 200-299 中，才表示 HTTP 回應成功的值。
                    if (response.StatusCode < HttpStatusCode.OK || response.StatusCode >= HttpStatusCode.MultipleChoices)
                    {
                        // 錢包端連線回應錯誤。
                        responseData = response.StatusDescription;
                        result.Response = GetResponseContent<T>(apiRequest.ResponseMode, responseData);
                        throw new Exception("連線錯誤");
                    }
                    // 成功回傳資料
                    using (var responseStream = new StreamReader(response.GetResponseStream()))
                    {
                        responseData = responseStream.ReadToEnd();

                        // 填入到回覆內文中
                        result.Response = GetResponseContent<T>(apiRequest.ResponseMode, responseData);
                        // 回應Header
                        result.ResponseHeaders = response.Headers;
                    }
                }

                #endregion
            }
            catch (Exception ex)
            {
                // 處理 Post 過程的錯誤
                throw OnSetException(ex);
            }
            return result;
        }

        /// <summary>
        /// 返回錢包的交易結果
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="mode"></param>
        /// <param name="walletResponse"></param>
        /// <returns></returns>
        protected T GetResponseContent<T>(ResponseMode mode, string Response)
        {
            // 將回應的 JSON 轉為ESUN PaymentResponse 格式內容
            var result = default(T);
            switch (mode)
            {
                case ResponseMode.ByJSON:
                    // 指定返回字串格式，不處理JSON的反序列化
                    if (typeof(T) == typeof(string))
                    {
                        result = (T)(Object)Response;
                    }
                    // 處理JSON的反序列化                    
                    else
                    {
                        result = JsonUtil.DeserializeObject<T>(Response);
                    }
                    break;
                case ResponseMode.ByText:
                case ResponseMode.ByStream:
                case ResponseMode.ByXML:
                    result = (T)(Object)Response;
                    break;
                default:
                    throw new NotSupportedException();
            }
            return result;
        }

        /// <summary>
        /// 設定回傳的 Exception 
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="walletHistory"></param>
        /// <returns></returns>
        private Exception OnSetException(Exception ex)
        {
            // 回傳錯誤
            var outEx = ex;
            // 錢包遠端連線失敗
            if (ex is WebException)
            {
                var webEx = (ex as WebException);
                var webHttpCode = ((webEx)?.Response as HttpWebResponse)?.StatusCode;
                // 判斷屬於 404 類型的錯誤，返回歸屬NotFound的錯誤     
                switch (webEx.Status)
                {
                    case WebExceptionStatus.NameResolutionFailure:
                    case WebExceptionStatus.ConnectFailure:
                    case WebExceptionStatus.ProxyNameResolutionFailure:
                    case WebExceptionStatus.TrustFailure:
                    case WebExceptionStatus.SecureChannelFailure:
                    case WebExceptionStatus.RequestProhibitedByProxy:
                        outEx = new AppException(ResultCodeType.ConnectNotFoundOrNotAllowed, ex);
                        break;
                    case WebExceptionStatus.Timeout:
                        outEx = new AppException(ResultCodeType.ConnectFail, "連線逾時", ex);
                        break;
                    default:
                        outEx = new AppException(ResultCodeType.ConnectFail, ex);
                        break;
                }
            }
            
            return outEx;
        }
        #endregion


    }
}
