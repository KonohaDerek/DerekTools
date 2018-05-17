using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace ApiHelper.Models
{
    public abstract class RequestBase
    {
        /// <summary>
        /// 叫用 Web 服務(API)的參數傳遞方式。
        /// </summary>
        [JsonIgnore]
        public abstract RequestMode Mode { get; }

        /// <summary>
        /// 叫用 Web 服務(API)的參數傳遞方式。
        /// </summary>
        [JsonIgnore]
        public abstract ResponseMode ResponseMode { get; }

        /// <summary>
        /// 叫用請求服務的模式，預設POST。
        /// </summary>
        [JsonIgnore]
        public virtual RequestMethodMode Method { get { return RequestMethodMode.POST; } }

        /// <summary>
        /// 簽名類型。
        /// </summary>
        [JsonIgnore]
        public abstract SignTypes SignType { get; }

        /// <summary>
        /// 簽名金鑰。
        /// </summary>
        [JsonIgnore]
        public virtual object MacKey { get; set; }

        /// <summary>
        /// 未加密 Request 訊息。
        /// </summary>
        [JsonIgnore]
        public virtual string DecryptRequestData { get; set; }

        /// <summary>
        /// 建立 Rquest 參數。
        /// </summary>
        /// <returns></returns>
        protected virtual RquestParams CreateRquestParams()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 轉為 URL 參數格式。
        /// </summary>
        /// <returns></returns>
        public string ToQueryString()
        {
            var reqParams = CreateRquestParams();

            if (Method == RequestMethodMode.GET)
            {
                // 要先以參數名稱排序。
                // VIP : According to RFC1866 section 8.2.1 both names and values should be encoded.
                var queryValues = reqParams.OrderBy(p => p.Key).Select(k => string.Format("{0}={1}", HttpUtility.UrlEncode(k.Key), HttpUtility.UrlEncode(k.Value.ToString()))).ToList();
                var queryString = string.Join("&", queryValues);
                return queryString;
            }
            else
            {
                var queryValues = reqParams.Select(k => string.Format("{0}={1}", k.Key, k.Value.ToString())).ToList();
                var queryString = string.Join("\n", queryValues);
                return queryString;
            }
        }

        /// <summary>
        /// 轉為 Json 格式字串。
        /// </summary>
        /// <param name="setting">自訂輸出設定。</param>
        /// <returns></returns>
        public virtual string ToJSON(JsonSerializerSettings setting = null)
        {
            if (setting == null)
            {
                setting = new JsonSerializerSettings()
                {
                    Formatting = Formatting.None,
                    NullValueHandling = NullValueHandling.Ignore,       // Note : 可以略掉 null 屬性。
                    //DateFormatString = "yyyy-MM-dd HH:mm:ss",
                    Error = (s, e) => { Console.WriteLine(e.ErrorContext); },
                    // 還有很多功能...
                };
            }
            var jsonStr = JsonConvert.SerializeObject(this, setting);
            return jsonStr;
        }



        /// <summary>
        /// 以自訂的轉換方法，將此物件轉為 Json 格式字串。
        /// </summary>
        /// <param name="converter"></param>
        /// <returns></returns>
        public string ToJSON(Func<RequestBase, string> converter)
        {
            if (converter == null) throw new ArgumentNullException("converter");
            var jsonStr = converter(this);
            return jsonStr;
        }


        /// <summary>
        /// 轉為 Xml 格式字串。
        /// </summary>
        /// <returns></returns>
        public virtual string ToXML()
        {
            var xmlStr = XmlUtil.Serialize(this);
            return xmlStr;
        }

        /// <summary>
        /// 轉為 XML 格式
        /// </summary>
        /// <returns></returns>
        public string ToXML(Func<RequestBase, string> converter)
        {
            if (converter == null) throw new ArgumentNullException("converter");

            var xmlStr = converter(this);
            return xmlStr;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string ToFormData()
        {
            throw new NotSupportedException();
        }

        /// <summary>
        /// 以未簽名的 Rquest 參數字串，表示 <see cref="RequestBase"/> 物件內容。
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            var s = string.Empty;

            var mode = this.Mode;
            // GET只能是URL型式
            if (Method == RequestMethodMode.GET)
            {
                mode = RequestMode.ByURL;
            }

            switch (mode)
            {
                case RequestMode.ByURL: s = ToQueryString(); break;
                case RequestMode.ByJSON: s = ToJSON(); break;
                case RequestMode.ByXML: s = ToXML(); break;
                    //case RequestMode.ByForm: ToFormData(); break;
            }
            return s;
        }

        /// <summary>
        /// 設定額外所需的Header
        /// 
        /// 預設不加入任何標頭
        /// </summary>
        /// <param name="headers"></param>
        public virtual void SetHeader(WebHeaderCollection headers) { }

#if DEBUG
        /// <summary>
        /// 識別此請求是否為測試資料
        ///     摘要: 配合錢包 ITestService 實作才有此功能
        ///     當有實作此功能時，在Post的部分會優先識別此筆交易是否適用測試的規則。
        ///     例如(BuyerID = 2800280028002800 開頭)
        ///     再執行 Test 的方法 並接收回傳值。
        ///     流程只會包裝於 IF DEBUG 
        /// </summary>
        [JsonIgnore]
        public bool IsTest { get; set; }

#endif
    }







}
