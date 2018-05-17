using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ApiHelper2
{
    public static class JsonUtil
    {
        /// <summary>
        /// 預設序列化的設定
        /// </summary>
        private static readonly JsonSerializerSettings _JsonSerializerSetting = new JsonSerializerSettings()
        {
            MissingMemberHandling = MissingMemberHandling.Ignore,
            NullValueHandling = NullValueHandling.Ignore
        };

        /// <summary>
        /// 定義屬於 物件、陣列 的屬性。
        /// </summary>
        private static string[] _ObjectArrayFeilds = new string[] { "RefundItems" };

        /// <summary>
        /// 物件序列化為Json格式
        /// </summary>
        /// <param name="obj">資料物件</param>
        /// <returns>Json字串</returns>
        public static string SerializeObject(object obj)
        {
            return JsonConvert.SerializeObject(obj, _JsonSerializerSetting);
        }
        /// <summary>
        /// 物件序列化為Json格式(有縮排)
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string SerializeObjectIndented(object obj)
        {
            var sets = new JsonSerializerSettings()
            {
                MissingMemberHandling = MissingMemberHandling.Ignore,
                NullValueHandling = NullValueHandling.Ignore,
                Formatting = Formatting.Indented
            };
            return JsonConvert.SerializeObject(obj, sets);
        }
        /// <summary>
        /// 物件序列化為字典格式
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> SerializeObjectToDictionary(object obj)
        {
            // null返回空集合
            if (obj == null) return new Dictionary<string, string>();
            // 序列化為字典
            return DeserializeToDictionary(JsonConvert.SerializeObject(obj, _JsonSerializerSetting));
        }
        /// <summary>
        /// 物件序列化為指定的另一個物件json序列化，用來只轉出繼承物件的屬性
        /// </summary>
        /// <param name="obj">資料物件</param>
        /// <returns>Json字串</returns>
        public static string SerializeObjectNext<T>(object obj)
        {
            return JsonConvert.SerializeObject(JsonConvert.DeserializeObject<T>(JsonConvert.SerializeObject(obj, _JsonSerializerSetting), _JsonSerializerSetting), _JsonSerializerSetting);
        }
        ///// <summary>
        ///// Json格式反序列化
        /////     摘要: 當反序列化失敗會返回 AppException 的錯誤
        ///// </summary>
        ///// <param name="json">Json字串</param>
        ///// <returns>反序列的物件</returns>
        //public static object DeserializeObject(string json) {
        //    try {
        //        return JsonConvert.DeserializeObject(json, JsonSerializerSetting);
        //    }
        //    catch (JsonException ex) {
        //        throw new AppException(ResultCodeType.gJsonDeserializeFail, ex);
        //    }
        //}
        /// <summary>
        ///  Json格式反序列化並轉為指定物件
        ///     摘要: 當反序列化失敗會返回 AppException 的錯誤
        /// </summary>
        /// <typeparam name="T">指定物件Type</typeparam>
        /// <param name="json">Json字串</param>
        /// <returns>反序列的物件</returns>
        public static T DeserializeObject<T>(string json)
        {
                return DeserializeNonDuplicates<T>(json);
        }
        /// <summary>
        /// 取得Json格式的資料字典
        ///     摘要: 此方法的字典會提供不取分大小寫的Key
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> JsonGet(this string json)
        {
            // 空字串返回空的字典
            if (string.IsNullOrWhiteSpace(json)) return new Dictionary<string, string>();

            // 序列化為字典
            Dictionary<string, string> dic = null;
         
            var dataDic = DeserializeToDictionary(json);
            dic = new Dictionary<string, string>(dataDic, StringComparer.OrdinalIgnoreCase); // 不區分大小寫
           
            return dic;
        }
        /// <summary>
        /// 取得指定Json格式中的某項特定屬性。
        ///     摘要: 當json格式 無法正常反序列化 會發生 Newtonsoft.Json.JsonReaderException
        /// </summary>
        /// <param name="json">Json字串</param>
        /// <param name="property">屬性名稱</param>
        /// <returns>屬性值，若無此屬性則為null</returns>
        public static string JsonGet(this string json, string property)
        {
            var dic = JsonGet(json);
            if (dic.ContainsKey(property) == false) return null;
            return dic[property];
        }
        /// <summary>
        /// 取得指定Json格式中的某項特定屬性並轉為指定類型的物件。
        /// </summary>
        /// <param name="json">Json字串</param>
        /// <param name="property">屬性名稱</param>
        /// <returns>屬性值，若無此屬性則為null</returns>
        public static T JsonGetData<T>(this string json, string property)
        {
            var data = JsonGet(json, property);
            if (data == null) return default(T);
            var result = JsonConvert.DeserializeObject<T>(data, _JsonSerializerSetting);
            return result;
        }
        /// <summary>
        /// 取得指定Json格式中屬性集
        ///     摘要: 當json格式 無法正常反序列化 會發生 Newtonsoft.Json.JsonReaderException
        /// </summary>
        /// <param name="json">Json字串</param>
        /// <param name="propertys">屬性陣列</param>
        /// <returns>屬性值陣列</returns>
        public static string[] JsonGetArray(this string json, params string[] propertys)
        {
            var dic = JsonGet(json);
            var result = propertys.Select<string, string>(o => Convert.ToString(dic[o])).ToArray();
            return result;
        }
        /// <summary>
        /// 檢查字串是否符合Json格式
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static bool IsValidJson(string input)
        {
            //如果傳入字串為空，則直接判定為非JSON
            if (string.IsNullOrWhiteSpace(input))
            {
                return false;
            }

            input = input.Trim();
            if ((input.StartsWith("{") && input.EndsWith("}")) || //For object
                (input.StartsWith("[") && input.EndsWith("]"))) //For array
            {
                try
                {
                    var obj = JToken.Parse(input);
                    return true;
                }
                catch (JsonReaderException jex)
                {
                    Console.WriteLine("JsonReaderException:" + jex.Message);
                    return false;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Json 安全性反序列化
        ///     摘要: 依據弱點掃描問題，排除Json資訊中重覆使用的屬性資訊做反序列化的處理。
        /// </summary>
        /// <typeparam name="T">指定物件Type</typeparam>
        /// <param name="json">Json字串</param>
        /// <returns>反序列的物件</returns>
        private static T DeserializeNonDuplicates<T>(string json)
        {
            JToken jObj = default(JToken);
            byte[] jsonBytes = Encoding.UTF8.GetBytes(json);
            using (var sr = new System.IO.StreamReader(new System.IO.MemoryStream(jsonBytes)))
            {
                var jsr = new JsonTextReader(sr);
                jObj = DeserializeReader(jsr);
            }
            return jObj.ToObject<T>();
        }
        /// <summary>
        /// Json 反序列化物件(讀取分析)
        /// </summary>
        /// <param name="reader"></param>
        /// <returns></returns>
        private static JToken DeserializeReader(JsonTextReader reader)
        {
            if (reader.TokenType == JsonToken.None)
            {
                reader.Read();
            }
            // 處理物件格式
            if (reader.TokenType == JsonToken.StartObject)
            {
                reader.Read();
                JObject obj = new JObject();
                while (reader.TokenType != JsonToken.EndObject)
                {
                    string propName = (string)reader.Value;
                    reader.Read();
                    JToken newValue = DeserializeReader(reader);

                    JToken existingValue = obj[propName];
                    // 加入新值
                    if (existingValue == null)
                    {
                        obj.Add(new JProperty(propName, newValue));
                    }
                    // 重覆的屬性會發出JsonException
                    else
                    {
                        throw new JsonException(string.Format(@"傳入的JSON格式有重覆的屬性""{0}""", propName));
                    }
                    //// 處理集合
                    //// PS: 重覆的集合屬性會合併成一個集合值
                    //else if (existingValue.Type == JTokenType.Array) {
                    //    CombineWithArray((JArray)existingValue, newValue);
                    //}
                    //// 處理重覆屬性
                    //// PS: 重覆的屬性會合併成一個集合值
                    //else {
                    //    JProperty prop = (JProperty)existingValue.Parent;
                    //    JArray array = new JArray();
                    //    prop.Value = array;
                    //    array.Add(existingValue);
                    //    CombineWithArray(array, newValue);
                    //}
                    reader.Read();
                }
                return obj;
            }
            // 處理集合格式
            if (reader.TokenType == JsonToken.StartArray)
            {
                reader.Read();
                JArray array = new JArray();
                while (reader.TokenType != JsonToken.EndArray)
                {
                    array.Add(DeserializeReader(reader));
                    reader.Read();
                }
                return array;
            }
            return new JValue(reader.Value);
        }
        /// <summary>
        /// 合併集合
        /// </summary>
        /// <param name="array"></param>
        /// <param name="value"></param>
        private static void CombineWithArray(JArray array, JToken value)
        {
            if (value.Type == JTokenType.Array)
            {
                foreach (JToken child in value.Children())
                    array.Add(child);
            }
            else
            {
                array.Add(value);
            }
        }

        /// <summary>
        /// 排序JSON OBJECT
        /// </summary>
        /// <param name="jObj"></param>
        public static void Sort(JObject jObj)
        {
            var props = jObj.Properties().ToList();
            foreach (var prop in props)
            {
                prop.Remove();
            }

            foreach (var prop in props.OrderBy(p => p.Name))
            {
                jObj.Add(prop);
                if (prop.Value is JObject)
                    Sort((JObject)prop.Value);
            }
        }

        /// <summary>
        /// 轉換CSV內容為Json
        /// 內容為須包含Header之完整內容
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="content"></param>
        /// <returns>Json</returns>
        public static IList<T> CsvToJson<T>(this string text)
        {
            // var lines = text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
            // 換行符號可能因為程式語言而有所不同
            var lines = text.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);
            return lines.CsvToJson<T>();
        }

        /// <summary>
        /// 轉換CSV內容為Json
        /// 內容為須包含Header之完整內容
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="content"></param>
        /// <returns>Json</returns>
        public static IList<T> CsvToJson<T>(this string[] lines)
        {
            string[] header = null;
            JArray jArray = new JArray();
            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }
                var content = line.CsvLineToArray();
                if (header == null)
                {
                    header = content;
                }
                else
                {
                    JObject jobj = new JObject();
                    for (int i = 0; i < header.Length; i++)
                    {
                        jobj.Add(header[i].Trim(), content[i]);
                    }
                    jArray.Add(jobj);
                }
            }
            return DeserializeObject<IList<T>>(jArray.ToString());
        }

        /// <summary>
        /// 解析CSV 
        /// 將單行字串轉為字串陣列
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        public static string[] CsvLineToArray(this string line)
        {
            string pattern = ",(?=(?:[^\"]*\"[^\"]*\")*(?![^\"]*\"))";
            Regex r = new Regex(pattern);
            return r.Split(line);
        }

        /// <summary>
        /// Json 格式反序列化成字典
        /// </summary>
        /// <param name="json">Json字串</param>
        /// <returns>資料字典</returns>
        public static Dictionary<string, string> DeserializeToDictionary(string json)
        {
            // Json DeserializeObject
            JToken jObj = (JToken)JsonConvert.DeserializeObject(json);
            // 陣列物件不處理字典化
            if (jObj is JArray) return new Dictionary<string, string>();

            // 檢查 Json 元素中包含 物件、陣列 的屬性，需要額外處理。
            var isHasObject = jObj.Any(o => o.HasValues && ((JProperty)o).Value.Type == JTokenType.Array || ((JProperty)o).Value.Type == JTokenType.Object);
            // 一般物件
            if (isHasObject == false)
            {
                return JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
            }

            // 依 JProperty 取值個別填入字典，若為物件、陣列元素則填入 Json 序列化的字串。
            var settings = new JsonSerializerSettings()
            {
                MissingMemberHandling = MissingMemberHandling.Ignore,
                NullValueHandling = NullValueHandling.Ignore
            };
            var dic = new Dictionary<string, string>();
            foreach (JProperty jProp in jObj)
            {
                if (jProp.HasValues == false) continue;

                string jValue = string.Empty;
                switch (jProp.Value.Type)
                {
                    // 陣列、物件元素
                    case JTokenType.Array:
                    case JTokenType.Object:
                        jValue = JsonConvert.SerializeObject(jProp.Value, settings);
                        break;
                    // 一般屬性值
                    default:
                        jValue = jProp.Value.ToString();
                        break;
                }

                dic.Add(jProp.Name, jValue);
            }
            return dic;
        }
        /// <summary>
        /// 字典序列化為 Json 格式
        /// </summary>
        /// <param name="dic">資料字典</param>
        /// <param name="objectFields">定義輸出為物件的欄位</param>
        /// <returns>Json字串</returns>
        public static string SerializeFromDictionary(Dictionary<string, string> dic)
        {
            // 檢查資料字典沒有定義為 物件、陣列 的屬性
            if (dic.Keys.Any(o => _ObjectArrayFeilds.Contains(o)) == false)
            {
                return JsonConvert.SerializeObject(dic);
            }

            // 處理包含 陣列、物件 字典序列化的問題            
            var sb = new StringBuilder();
            using (var sw = new System.IO.StringWriter(sb))
            {
                var jsw = new JsonTextWriter(sw);
                jsw.WriteStartObject();
                foreach (var dicValue in dic)
                {
                    // 寫入屬性名稱
                    jsw.WritePropertyName(dicValue.Key);
                    // 檢查有定義參數值物件或集合，則額外處理
                    bool isRawValue = false;
                    if (_ObjectArrayFeilds.Contains(dicValue.Key))
                    {
                        isRawValue = (dicValue.Value.StartsWith("{") && dicValue.Value.EndsWith("}")) ||
                                     (dicValue.Value.StartsWith("[{") && dicValue.Value.EndsWith("}]"));
                    }

                    // 物件、集合格式直接輸出
                    // isRawValue == true 例: "A":{"Name"="Hello"}
                    // else 例: "A":"{"Name"="Hello"}"
                    if (isRawValue)
                    {
                        jsw.WriteRawValue(dicValue.Value);
                    }
                    // 資料值為一般值
                    else
                    {
                        jsw.WriteValue(dicValue.Value);
                    }
                }
                jsw.WriteEndObject();
            }
            return sb.ToString();
        }


    }
}
