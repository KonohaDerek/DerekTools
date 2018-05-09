namespace ExcelHelper.Models
{
    /// <summary>
    /// 匯出Excel多重欄位的定義值
    /// </summary>
    public class ExcelColumnValue
    {
        /// <summary>
        /// 標題文字
        /// </summary>
        public string ColumnText { get; set; }

        /// <summary>
        /// 欄位名稱
        /// </summary>
        public string Column { get; set; }

        /// <summary>
        /// 欄位排序參數 (例:10A、10B)
        /// </summary>
        public string Order { get; set; }

        /// <summary>
        /// 欄位值
        /// </summary>
        public object Value { get; set; }
    }
}
