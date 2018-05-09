using System;
using System.Runtime.InteropServices;

namespace ExcelHelper.Attribute
{
    /// <summary>
    /// 定義欄位是否可以匯出
    /// </summary>
    [AttributeUsage(AttributeTargets.All, Inherited = false)]
    [ComVisible(true)]
    public class ExportAttribute : System.Attribute
    {
        /// <summary>
        /// 欄位標題
        /// </summary>
        private string ColumnHeader { get; set; }

        /// <summary>
        /// 格式化
        /// </summary>
        private string DataFormatString { get; set; }

        /// <summary>
        /// 排序
        /// </summary>
        private int Order { get; set; }

        /// <summary>
        /// 資料類型
        /// </summary>
        private Type DataType { get; set; }

        /// <summary>
        /// Excel數值格式化
        /// </summary>
        private string Numberformat { get; set; }

        /// <summary>
        /// Excel 函數公式
        /// </summary>
        private string Formula { get; set; }

        /// <summary>
        /// 屬性對應名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 是否為多欄位的屬性
        /// </summary>
        public bool IsMultipleFields { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public ExportAttribute()
        {
        }

        /// <summary>
        /// 匯出Excel屬性
        /// 
        /// Formula 公式目前只支援SUM 、 AVG 、 COUNT，範圍為整欄內容
        /// </summary>
        /// <param name="ColumnHeader">匯出欄位名稱</param>
        /// <param name="DataFormatString">格式化</param>
        /// <param name="Order">排序</param>
        /// <param name="DataType">資料類型</param>
        /// <param name="Numberformat">數字格式</param>
        /// <param name="Formula">使用公式: Formula 公式目前只支援SUM 、 AVG 、 COUNT，範圍為整欄內容。
        ///                       使用範例 : 欄位公式輸入SUM ， 則輸出的Excel 公式為 SUM(M:N)</param>
        public ExportAttribute(string ColumnHeader, int Order = 0, string DataFormatString = "{0}", Type DataType = null, string Numberformat = "", string Formula = "")
        {
            this.ColumnHeader = ColumnHeader;

            this.DataFormatString = DataFormatString;

            this.Order = Order;

            this.DataType = DataType;

            this.Numberformat = Numberformat;

            this.Formula = Formula;
        }

        /// <summary>
        /// 匯出Excel屬性
        /// </summary>
        /// <param name="ColumnHeader">匯出欄位名稱</param>
        /// <param name="DataType">資料類型</param>
        /// <param name="Numberformat">數字格式</param>
        /// <param name="Formula">使用公式: Formula 公式目前只支援SUM 、 AVG 、 COUNT，範圍為整欄內容。
        ///                       使用範例 : 欄位公式輸入SUM ， 則輸出的Excel 公式為 SUM(M:N)</param>
        public ExportAttribute(string ColumnHeader, Type DataType, string Numberformat = "", string Formula = "")
        {
            this.ColumnHeader = ColumnHeader;

            DataFormatString = "{0}";

            Order = 0;

            this.DataType = DataType;

            this.Numberformat = Numberformat;

            this.Formula = Formula;
        }

        /// <summary>
        /// 輸出欄位標題
        /// </summary>
        /// <returns></returns>
        public string GetColumnHeader()
        {
            return this.ColumnHeader;
        }

        /// <summary>
        /// 取得格式化字串
        /// </summary>
        /// <returns></returns>
        public string GetFormat()
        {
            return DataFormatString;
        }


        /// <summary>
        /// 取得排序編號
        /// </summary>
        /// <returns></returns>
        public int GetOrder()
        {
            return Order;
        }

        public Type GetDataType()
        {
            return DataType;

        }

        /// <summary>
        /// 取得Excel Format
        /// </summary>
        /// <returns></returns>
        public string GetNumberformat()
        {
            return Numberformat;
        }

        /// <summary>
        /// 取得Excel 公式
        /// </summary>
        /// <returns></returns>
        public string GetFormula()
        {
            return Formula;
        }


        /// <summary>
        /// 輸出格式化資料
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public object GetFormatValue(object value, Type propertyType = null)
        {
            if (DataType != null)
                propertyType = DataType;

            var valueType = value.GetType();
            //判斷輸出並依類型給值
            //一般字串
            if (propertyType == typeof(string))
            {
                return string.Format(DataFormatString, value);
            }
            //10進制數值
            else if (propertyType == typeof(decimal) || propertyType == typeof(decimal?))
            {
                decimal amt = 0.0M;

                if (valueType == typeof(decimal) || valueType == typeof(decimal?))
                {
                    return value;
                }
                if (decimal.TryParse(string.Format("{0}", value), out amt))
                {
                    return amt;
                }
                else
                {
                    return value;
                }
            }
            //數值
            else if (propertyType == typeof(int) || propertyType == typeof(int?))
            {
                int count = 0;
                if (valueType == typeof(int) || valueType == typeof(int?))
                {
                    return value;
                }

                if (int.TryParse(string.Format("{0}", value), out count))
                {
                    return count;
                }
                else
                {
                    return value;
                }
            }
            //其他
            else
            {
                return string.Format(DataFormatString, value);
            }
        }
    }
}
