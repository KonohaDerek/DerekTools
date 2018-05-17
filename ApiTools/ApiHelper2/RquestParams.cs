using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ApiHelper
{
    /// <summary>
    /// 串接URL參數物件。
    /// </summary>
    public class RquestParams : IEnumerable<KeyValuePair<string, object>>
    {
        readonly Dictionary<string, object> _Items = new Dictionary<string, object>();

        /// <summary>
        /// 建立 URL 參數物件
        /// </summary>
        public RquestParams() { }

        /// <summary>
        /// 建立 URL 參數物件
        /// </summary>
        /// <param name="item">參考資料</param>
        /// <param name="flag">挑選的屬性Flag</param>
        public RquestParams(object item, BindingFlags flag = (BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.DeclaredOnly))
        {
            var props = item.GetType().GetProperties(flag);
            foreach (var prop in props)
            {
                AddOption(prop.Name, prop.GetValue(item));
            }
        }

        /// <summary>
        /// 增加 URL 參數
        /// </summary>
        /// <param name="paramName">參數名稱。</param>
        /// <param name="value">參數值</param>
        /// <exception cref="ArgumentNullException">參數名稱為 null 或空字串。</exception>
        /// <exception cref="ArgumentException">已存在具有相同名稱的參數。</exception>
        public void Add(string paramName, object value)
        {
            if (string.IsNullOrEmpty(paramName)) throw new ArgumentNullException("paramName");
            if (_Items.ContainsKey(paramName)) throw new ArgumentException("paramName");         // 不重覆加入
            if (value == null || string.IsNullOrEmpty(value.ToString()))
            {              // 非選擇性參數，不可為空值。
                throw new ArgumentNullException("value");
            }

            _Items.Add(paramName, value);
        }

        /// <summary>
        /// 增加選擇性的 URL 參數
        /// </summary>
        /// <param name="paramName">參數名稱。</param>
        /// <param name="value">參數值</param>
        /// <exception cref="ArgumentNullException">參數名稱為 null 或空字串。</exception>
        public void AddOption(string paramName, object value)
        {
            if (string.IsNullOrEmpty(paramName)) throw new ArgumentNullException("paramName");
            if (value == null || string.IsNullOrEmpty(value.ToString())) return;        // 選擇性參數，若為空值則不加入。
            if (_Items.ContainsKey(paramName)) return;                                  // 不重覆加入

            _Items.Add(paramName, value);
        }

        #region IEnumerable<KeyValuePair<string,object>> 成員

        /// <summary>
        /// 取得參數的集合列舉
        /// </summary>
        /// <returns></returns>
        public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
        {
            return _Items.GetEnumerator();
        }

        #endregion

        #region IEnumerable 成員

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _Items.GetEnumerator();
        }

        #endregion
    }
}
