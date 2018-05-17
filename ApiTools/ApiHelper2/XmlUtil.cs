using System;
using System.IO;
using System.Text;
using System.Xml.Serialization;

namespace ApiHelper
{
    /// <summary>
    /// Xml服務功能
    /// </summary>
    public static class XmlUtil
    {
        /// <summary>
        /// Xml檔案 反序列化
        /// </summary>
        /// <typeparam name="T">物件類別</typeparam>
        /// <param name="path">Xml檔案路徑</param>
        /// <param name="root">Root標籤</param>
        /// <returns>轉換的物件</returns>
        public static T DeserializeFile<T>(string path, string root = null)
        {
            var xmlSer = GetXmlSerializer(typeof(T), root);
            var desObject = default(T);
            using (var sr = new StreamReader(path))
            {
                desObject = (T)xmlSer.Deserialize(sr);
            }
            return desObject;
        }
        /// <summary>
        /// Xml反序列化
        /// </summary>
        /// <typeparam name="T">物件類別</typeparam>
        /// <param name="xml">Xml字串</param>
        /// <param name="root">Root標籤</param>
        /// <returns>轉換的物件</returns>
        public static T Deserialize<T>(string xml, string root = null)
        {
            var xmlSer = GetXmlSerializer(typeof(T), root);
            var desObject = (T)xmlSer.Deserialize(new StringReader(xml));
            return desObject;
        }

        /// <summary>
        /// 取得指定型別的序列化結構
        /// </summary>
        /// <param name="type">指定型別</param>
        /// <param name="root">Root標籤</param>
        /// <returns>序列化結構</returns>
        public static XmlSerializer GetXmlSerializer(Type type, string root = null)
        {
            var xmlSer = default(XmlSerializer);
            if (string.IsNullOrWhiteSpace(root) == false)
            {
                xmlSer = new XmlSerializer(type, new XmlRootAttribute(root));
            }
            else
            {
                xmlSer = new XmlSerializer(type);
            }
            return xmlSer;
        }

        /// <summary>
        /// Xml序列化
        /// </summary>
        /// <param name="obj">序列化物件</param>
        /// <param name="root">Root標籤</param>
        /// <returns>Xml字串</returns>
        public static string Serialize(object obj, string root = null)
        {
            var xmlSer = GetXmlSerializer(obj.GetType(), root);

            // 輸出utf-8格式的XML字串
            var xml = string.Empty;
            using (var ms = new MemoryStream())
            {
                var sw = new StreamWriter(ms, Encoding.UTF8);
                xmlSer.Serialize(sw, obj);
                ms.Seek(0, SeekOrigin.Begin);
                var sr = new StreamReader(ms, Encoding.UTF8);
                xml = sr.ReadToEnd();
            }

            return xml;
        }
    }
}
