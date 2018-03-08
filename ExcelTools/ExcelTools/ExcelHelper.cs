using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelHelper.Models;
using Jil;
using OfficeOpenXml;
using OfficeOpenXml.Table;

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
        public static List<T> ImportExcelAsync<T>(this byte[] fileBytes , string workSheetsName="")
            where T : class, new()
        {
            var result = default(List<T>);
            using (MemoryStream fs = new MemoryStream(fileBytes))
            {

                //載入Excel檔案
                using (ExcelPackage ep = new ExcelPackage(fs))
                {
                    ExcelWorksheet sheet =string.IsNullOrWhiteSpace(workSheetsName) ? ep.Workbook.Worksheets[1] : ep.Workbook.Worksheets[workSheetsName] ;//取得Sheet
                    if (sheet == null)
                    {
                        throw new ArgumentNullException(string.Format("指定的工作表:{0}不存在。",workSheetsName));
                    }
                    List<T> RowData = new List<T>();
                   
                    //處理Excel資料
                    result = sheet.ConvertTableToObjects<T>();
                }
            }
            return result;
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

        /// <summary>
        /// 檢查資料
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        private static void CheckData<T>(T data) where T : class, new()
        {
            var Validator = new DataAnnotationValidator();
            if (!Validator.TryValidate(data))
            {
                throw new Exception(string.Format("資料驗證失敗 : {0}", JSON.Serialize(Validator.ValidationResults.Select(o => o.ErrorMessage))));
            }
        }


        private static List<T> ConvertTableToObjects<T>(this ExcelWorksheet sheet) where T : new()
        {
            //有無定義欄位的Flag
            var definedFlag = sheet.Workbook.Names.Any();
            //Get the properties of T
            var tprops = typeof(T)
                .GetProperties()
                .ToList();

            //取得總行數資料
            var groups = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column]
                .GroupBy(cell => cell.Start.Row)
                .ToList();

            //由資料列的第一行取得每一欄的資料類型
            var types = groups
                .Skip(1)
                .First()
                .Select(rcell => rcell.Value.GetType())
                .ToList();

            //Assume first row has the column names
            var colnames = definedFlag ? 
                //如果是使用定義表，則先行排序在處理
                sheet.Workbook.Names.OrderBy(hcell=>hcell.Start.Column).Select((hcell, idx) => new {hcell.Name,index = idx}).Where(o => tprops.Select(p => p.Name).Contains(o.Name)).ToList() :
                //如果不是則直接使用Group結果
                groups.First().Select((hcell, idx) => new {Name = hcell.IsName? hcell.Text: hcell.Value.ToString(),index = idx}).Where(o => tprops.Select(p => p.Name).Contains(o.Name)).ToList();

            //Everything after the header is data
            var rowvalues = groups
                .Skip(1) //Exclude header
                .Select(cg => cg.Select(c => c.Value).ToList());


            //Create the collection container
            var collection = rowvalues
                .Select(row =>
                {
                    var tnew = new T();
                    colnames.ForEach(colname =>
                    {
                        //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                        var val = row[colname.index];
                        var type = types[colname.index];
                        var prop = tprops.First(p => p.Name == colname.Name);

                        //If it is numeric it is a double since that is how excel stores all numbers
                        if (type == typeof(double))
                        {
                            //Unbox it
                            var unboxedVal = (double)val;

                            //FAR FROM A COMPLETE LIST!!!
                            if (prop.PropertyType == typeof(Int32))
                                prop.SetValue(tnew, (int)unboxedVal);
                            else if (prop.PropertyType == typeof(double))
                                prop.SetValue(tnew, unboxedVal);
                            else if (prop.PropertyType == typeof(DateTime))
                                prop.SetValue(tnew, convertDateTime(unboxedVal));
                            else if (prop.PropertyType == typeof(string))
                                prop.SetValue(tnew, Convert.ToString(unboxedVal));
                            else
                                throw new NotImplementedException(String.Format("Type '{0}' not implemented yet!", prop.PropertyType.Name));
                        }
                        //如果是日期格式則轉換
                        else if (type == typeof(DateTime))
                        {
                            if (prop.PropertyType == typeof(DateTime))
                                prop.SetValue(tnew, val);
                            else
                                prop.SetValue(tnew, ((DateTime)val).ToString());
                        }
                        else
                        {
                            //Its a string
                            prop.SetValue(tnew, val);
                        }
                    });

                    return tnew;
                });


            //Send it back
            return collection.ToList();
        }

        /// <summary>
        /// 轉換日期
        /// </summary>
        /// <param name="unixTime"></param>
        /// <returns></returns>
        private static DateTime convertDateTime(double unixTime)
        {
           
            return   DateTime.FromOADate(unixTime);
        }

    }
}

