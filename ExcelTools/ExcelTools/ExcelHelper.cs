using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ExcelHelper.Attribute;
using ExcelHelper.Models;
using Jil;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace ExcelHelper
{
    public static class ExcelHelper
    {
        /// <summary>
        /// 產生Execl byte array
        /// </summary>
        /// <typeparam name="TModel">匯出資料型別</typeparam>
        /// <param name="datas">資料集合</param>
        /// <param name="sheetName">工作表名稱</param>
        /// <returns></returns>
        public static byte[] ExportExcel<TModel>(IEnumerable<TModel> datas, string sheetName, string ExcelTitle = "")
        {
            if (datas == null) return null;
            byte[] RET = null;
            var execlColumes = new List<ExportAttribute>();             // 欄位資訊

            // 取出操作對象的屬性
            var Properties = typeof(TModel).GetProperties();
            // 屬性字典
            var temp_Properties = Properties.ToDictionary(o => o.Name);
            // 資料內容字典
            var temp = new Dictionary<string, List<object>>();
            // 排序加值，當輸出欄位有安插"多重欄位"的屬性值則會用上此值做加值
            int orderBonus = 0;

            // 整理物件屬性提供的額外資訊
            var propAttrs = (from expProp in Properties
                             from expAttr in expProp.GetCustomAttributes(typeof(ExportAttribute), false).OfType<ExportAttribute>()
                             select new
                             {
                                 Property = expProp,
                                 Attribute = expAttr,
                                 Order = expAttr.GetOrder()
                             }).OrderBy(o => o.Order).ToList();


            // 取得屬性類別上的 ExportAttribute 來分析，並建立資料字典
            orderBonus = 0;
            foreach (var propAttr in propAttrs)
            {
                var prop = propAttr.Property;
                var attr = propAttr.Attribute;

                // 有多重欄位時，展開欄位中的資訊
                if (propAttr.Attribute.IsMultipleFields == true && datas?.Count() > 0)
                {
                    // 取得首筆資料提供的欄位陣列，datas中的欄位數前提必須相同才不會有輸出上的問題。
                    var mainData = datas.First();
                    var mulipleFields = (IEnumerable<ExcelColumnValue>)prop.GetValue(mainData);
                    // 展開欄位時做 ExcelColumnValue 的內置排序
                    mulipleFields = mulipleFields.OrderBy(o => o.Order);
                    // 展開多重欄位的資訊，並且填入建立Excel的資料字典
                    foreach (var mulipleField in mulipleFields)
                    {
                        orderBonus += 1;
                        int orderField = attr.GetOrder() + orderBonus;
                        var attrSub = new ExportAttribute(mulipleField.ColumnText, orderField, attr.GetFormat(), attr.GetDataType(), attr.GetNumberformat(), attr.GetFormula());
                        attrSub.IsMultipleFields = false;
                        attrSub.Name = mulipleField.Column;
                        execlColumes.Add(attrSub);
                        temp.Add(attrSub.Name, datas.Select(o => ((IEnumerable<ExcelColumnValue>)prop.GetValue(o)).FirstOrDefault(p => p.Column == attrSub.Name)?.Value).ToList());
                    }
                }
                else
                {
                    // 判斷欄位輸出有插入值，則更新 ExportAttribute.Order 的順位。
                    var attSub = (orderBonus == 0) ? attr :
                        new ExportAttribute(attr.GetColumnHeader(), attr.GetOrder() + orderBonus, attr.GetFormat(), attr.GetDataType(), attr.GetNumberformat(), attr.GetFormula());
                    attSub.Name = prop.Name;
                    // 填入操作字典
                    execlColumes.Add(attSub);
                    temp.Add(prop.Name, datas.Select(o => prop.GetValue(o)).ToList());
                }
            }

            // 在記憶體中建立一個Excel物件
            using (ExcelPackage ep = new ExcelPackage())
            {
                ExcelWorksheet sheet = ep.Workbook.Worksheets.Add(sheetName);
                #region 標頭
                // 分拆設定好的欄位名稱/欄位值的索引資訊。
                var excelTitleWithIndex = execlColumes.OrderBy(o => o.GetOrder())
                                            .Select((title, index) => new {
                                                value = title,
                                                index,
                                                hasFormula = !string.IsNullOrWhiteSpace(title.GetFormula())
                                            }).ToList();

                int startContentRowIndex = 1;
                int HeaderCount = excelTitleWithIndex.Count();
                if (!string.IsNullOrWhiteSpace(ExcelTitle))
                {
                    sheet.Cells["A1"].Value = ExcelTitle;
                    sheet.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;//水平置中對齊
                    sheet.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;//垂直置中對齊
                    sheet.Cells["A1"].Style.Font.Size = 16;//字體大小
                    sheet.Cells[1, 1, 1, HeaderCount].Merge = true;
                    sheet.Cells["A1"].AutoFitColumns();    //自動調整欄位大小
                    //添加標題，所以內容下移
                    startContentRowIndex += 1;
                }
                // 填入標頭資料
                foreach (var itme in excelTitleWithIndex)
                {
                    sheet.Cells[startContentRowIndex, 1 + itme.index].Value = itme.value.GetColumnHeader();
                    sheet.Cells[startContentRowIndex, 1 + itme.index].AutoFitColumns();//自動調整欄位大小
                }
                //加完Header 將內容下移
                startContentRowIndex += 1;
                #endregion
                // 寫入明細
                object sheetValue = null;
                Debug.WriteLine("開始:" + DateTime.Now.ToString("HH:mm:ss:ffff"));
                foreach (var item in excelTitleWithIndex)
                {
                    int rowIndex = 0; // 用Index的原因是，資料量可能很大，Project會造成效率不彰。
                    var propertity = temp_Properties.Get(item.value.Name);
                    var temp_datas = temp.Get(item.value.Name);
                    int celNum = 1 + item.index;
                    var PropertyType = item.value.GetDataType() ?? propertity.PropertyType;
                    var Numberformat = item.value.GetNumberformat();
                    foreach (var dataRow in temp_datas)
                    {
                        int rowNum = startContentRowIndex + rowIndex;
                        try
                        {
                            sheetValue = dataRow ?? "";
                            //填入格式化數值
                            sheetValue = item.value.GetFormatValue(sheetValue, PropertyType);
                        }
                        catch (Exception)
                        {
                            sheetValue = "";
                        }
                        //填入Excel 格式
                        sheet.Cells[rowNum, celNum].Style.Numberformat.Format = Numberformat;
                        sheet.Cells[rowNum, celNum].Value = sheetValue;
                        rowIndex++;
                    }
                    if (item.hasFormula)
                    {
                        //修改Formula 只附加於整欄最下方
                        //範圍為整欄資料
                        sheet.Cells[startContentRowIndex + rowIndex, celNum].Formula = string.Format("{0}({1}:{2})", item.value.GetFormula(), sheet.Cells[startContentRowIndex, celNum].Address, sheet.Cells[startContentRowIndex + rowIndex - 1, celNum].Address);
                        //填入Excel 格式
                        //依原本欄位設置格式添加
                        sheet.Cells[startContentRowIndex + rowIndex, celNum].Style.Numberformat.Format = Numberformat;
                    }
                }
                sheet.Cells.AutoFitColumns();
                Debug.WriteLine("結束:" + DateTime.Now.ToString("HH:mm:ss:ffff"));
                RET = ep.GetAsByteArray();
            }

            return RET;
        }

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

        /// <summary>
        /// 轉換資料表內資料至物件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="sheet"></param>
        /// <returns></returns>
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

            //取得資料起始欄
            var rowIndex = definedFlag ? GetRowIndex(sheet, tprops) : GetRowIndex(groups, tprops) ;

            //Assume first row has the column names
            var colnames = definedFlag ? 
                                 GetColnames(sheet, tprops)
                                : GetColnames(groups.Skip(rowIndex-1).First(), tprops);

            //由資料列的第一行取得每一欄的資料類型
            var types = groups
                .Skip(rowIndex)
                .First()
                .Select(rcell => rcell.Value.GetType())
                .ToList();

            //Everything after the header is data
            var rowvalues = groups
                .Skip(rowIndex) //Exclude header
                .Select(cg => cg.Select(c => c.Value).ToList());


            //Create the collection container
            var collection = rowvalues
                .Select(row =>
                {
                    var tnew = new T();
                    colnames.ForEach(colname =>
                    {
                        //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                        var val = row[colname.Item2];
                        var type = types[colname.Item2];
                        var prop = tprops.First(p => p.Name == colname.Item1);

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
                                prop.SetValue(tnew, unboxedVal.ToDateTime());
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
        /// 取得ColumnIndex
        /// </summary>
        /// <param name="groups"></param>
        /// <param name="sheet"></param>
        /// <param name="tprops"></param>
        /// <returns></returns>
        private static int GetRowIndex(List<IGrouping<int, ExcelRangeBase>> groups, List<PropertyInfo> tprops)
        {
            foreach (var group in groups)
            {
               var rowIndexs =group.Select(o => new
                {
                    Name =  o.Text,
                    rowIndex = o.Start.Row
                }).Where(o => tprops.Select(p => p.Name).Contains(o.Name)).Select(o => o.rowIndex).Distinct();
                var rowIndex = rowIndexs.Where(o => o > 0).Distinct().FirstOrDefault();
                if (rowIndex > 0)
                {
                    return rowIndex;
                }
            }
            //如果都沒有找到則預設跳過首列
            return 1;
        }

        /// <summary>
        /// 取得ColumnIndex
        /// </summary>
        /// <param name="groups"></param>
        /// <param name="sheet"></param>
        /// <param name="tprops"></param>
        /// <returns></returns>
        private static int GetRowIndex( ExcelWorksheet sheet, List<PropertyInfo> tprops)
        {
            return sheet.Workbook.Names.Select(hcell => new
            {
                hcell.Name,
                rowIndex = hcell.Start.Row
            }).Where(o => tprops.Select(p => p.Name).Contains(o.Name)).Select(o=>o.rowIndex).Distinct().First();
        }

        /// <summary>
        /// 由群組取得對應的Colname
        /// </summary>
        /// <param name="groups"></param>
        /// <param name="tprops"></param>
        /// <returns></returns>
        private static List<Tuple<string, int>> GetColnames(IGrouping<int, ExcelRangeBase> group, List<PropertyInfo> tprops)
        {
            return group.Select((hcell, idx) => new {
                Name =  hcell.Text,
                index = idx
            }).Where(o => tprops.Select(p => p.Name).Contains(o.Name))
              .AsEnumerable()
            .Select(c => new Tuple<string, int>(c.Name, c.index)).ToList();
        }

     

        /// <summary>
        /// 由定義對應取得對應的Colnames
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="tprops"></param>
        /// <returns></returns>
        private static List<Tuple<string,int>> GetColnames(ExcelWorksheet sheet, List<PropertyInfo> tprops)
        {
            //如果是使用定義表，則先行排序在處理
            return sheet.Workbook.Names.OrderBy(hcell => hcell.Start.Column).Select((hcell, idx) => new
            {
                hcell.Name,
                index = idx
            }).Where(o => tprops.Select(p => p.Name).Contains(o.Name))
            .AsEnumerable()
            .Select(c => new Tuple<string, int>(c.Name, c.index)).ToList();
        }

    }
}

