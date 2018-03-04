using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper.Models
{
    public class FileDataInfo
    {
        /// <summary>
        /// 檔案名稱
        /// </summary>
        public string FileName
        {
            get; set;
        }

        /// <summary>
        /// 檔案內容
        /// </summary>
        public byte[] FileData
        {
            get; set;
        }
    }
}
