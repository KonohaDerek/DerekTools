using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelHelper.Models;
using OfficeOpenXml;

namespace ExcelHelper
{
    public static class ExcelHelper
    {
        /// <summary>
        /// 匯入Excel資料轉置為List<T>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileBytes"></param>
        /// <returns></returns>
        public static IEnumerable<T> ImportExcelAsync<T>(byte[] fileBytes , string workSheetsName="")
            where T : class, new()
        {
            using (MemoryStream fs = new MemoryStream(fileBytes))
            {

                //載入Excel檔案
                using (ExcelPackage ep = new ExcelPackage(fs))
                {
                    ExcelWorksheet sheet =string.IsNullOrWhiteSpace(workSheetsName) ? ep.Workbook.Worksheets[workSheetsName] : ep.Workbook.Worksheets[1];//取得Sheet
                    if (sheet == null)
                    {
                        throw new ArgumentNullException(string.Format("指定的工作表:{0}不存在。",workSheetsName));
                    }
                    List<T> RowData = new List<T>();

                    bool isLastRow = false;
                    int RowId = 2;   // 因為有標題列，所以從第2列開始讀起

                    do  // 讀取資料，直到讀到空白列為止
                    {
                        string cellValue = sheet.Cells[RowId, 1].Text;
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            isLastRow = true;
                        }
                        else
                        {
                            
                        }
                    } while (!isLastRow);
                }
            }
            return null;
        }

        /// <summary>
        /// 由範本產生Excel檔案
        /// </summary>
        /// <param name="templatePath"></param>
        /// <param name="rowIndex"></param>
        /// <param name="paramsList"></param>
        /// <returns></returns>
        public static async Task<FileDataInfo> ExportExcelByTemplateAsync(string templatePath, string exportFileName, int rowIndex = 0, params IEnumerable[] paramsList)
        {
            return await Task.Run(() =>
            {
                /*
               * 由事先定義好的範本檔案將要匯出的資料填入
               * 在Excel中欄位的定義名稱是寫入在Workbook之中，所以由Workbook去找到對應的欄位名稱
               * 在依rowIndex去填入位置，如果rowIndex起始為0，表示資料要寫在對應的欄位上
               * ****/
                FileDataInfo RET = new FileDataInfo();
                FileInfo FileInfoXLSTemplate = new FileInfo(templatePath);

                using (ExcelPackage package = new ExcelPackage(null, FileInfoXLSTemplate))
                {
                    var wb = package.Workbook;
                    ExcelDataProcess(rowIndex, wb, paramsList);
                    //設定檔案
                    RET.FileData = package.GetAsByteArray();
                    RET.FileName = string.Format("{0}.xlsx", exportFileName);
                }
                return RET;
            }).ConfigureAwait(false);
        }

        /// <summary>
        /// 處理Excel資料對應
        /// </summary>
        /// <param name="rowIndex"></param>
        /// <param name="wb"></param>
        /// <param name="paramsList"></param>
        private static void ExcelDataProcess(int rowIndex, ExcelWorkbook wb, params IEnumerable[] paramsList)
        {
            /*
            * 1.先行判斷傳入的params有幾組，依個別組別去篩選並填入資料
            * 2.由GetProperties查詢出目前物件的屬性資料
            * 3.檢查Excel Workbook中是否有相對應的資料欄位，有則瑱入，無則略過不做處理
            * 4.todo : 加入Attribute 去判定是否要格式化資料或是套用公式
            * *********/
            foreach (var datas in paramsList)
            {
                var index = rowIndex;
                foreach (var item in datas)
                {
                    //取出操作對象的屬性
                    var Properties = item.GetType().GetProperties();
                    foreach (var propertity in Properties)
                    {
                        if (wb.Names.ContainsKey(propertity.Name))
                        {
                            var cell = wb.Names[propertity.Name];

                            cell.Worksheet.Cells[cell.End.Row + index, cell.End.Column].Value = propertity.GetValue(item);
                        }
                    }
                    index++;
                }
            }
        }


    }
}

