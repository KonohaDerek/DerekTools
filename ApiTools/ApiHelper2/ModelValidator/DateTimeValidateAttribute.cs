using System;
using System.ComponentModel.DataAnnotations;

namespace ApiHelper.ModelValidator
{
    /// <summary>
    /// 檢查字串日期格式
    /// </summary>
    public class DateTimeValidateAttribute : RequiredAttribute
    {
        /// <summary>
        /// 指定的日期時間格式
        /// </summary>
        public string SupportFormat { get; set; }

        /// <summary>
        /// 檢查日期必須為今日之前
        /// </summary>
        public bool ToDayBefore { get; set; }

        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {

            // 檢查必填項目
            var result = base.IsValid(value, validationContext);
            if (result != null) return result;

            // 檢查Format格式
            var dateDT = Convert.ToString(value + "");
            var validateDT = default(DateTime);
            var supportFMT = new string[] { "yyyyMMddHHmmss" };
            if (string.IsNullOrWhiteSpace(SupportFormat) == false)
            {
                supportFMT = new string[] { SupportFormat };
            }

            if (DateTime.TryParseExact(dateDT, supportFMT, null, System.Globalization.DateTimeStyles.None, out validateDT) == false)
            {
                if (DateTime.TryParse(dateDT, out validateDT) == false)
                {
                    var mesg = string.Format("{0} 時間格式錯誤。", validationContext.DisplayName);
                    return new ValidationResult(mesg, new string[] { validationContext.MemberName });
                }
            }

            // 檢查日期必須為今日之前
            if (ToDayBefore)
            {
                if ((validateDT.Date > DateTime.Now.Date))
                {
                    var mesg = string.Format("{0} 必須於今日之前。", validationContext.DisplayName);
                    return new ValidationResult(mesg, new string[] { validationContext.MemberName });
                }
            }
            return ValidationResult.Success;
        }

    }
}
