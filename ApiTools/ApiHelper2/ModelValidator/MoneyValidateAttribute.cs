using System;
using System.ComponentModel.DataAnnotations;

namespace ApiHelper.ModelValidator
{
    /// <summary>
    /// 金額檢核
    /// </summary>
    public class MoneyValidateAttribute : RequiredAttribute
    {
        /// <summary>
        /// 金額檢核
        /// </summary>
        public MoneyValidateAttribute() : base()
        {
            MaxValue = 999999999M;   // 預設上限 N[9,0]
        }

        /// <summary>
        /// 金額檢核
        /// </summary>
        /// <param name="max">指定金額上限</param>
        public MoneyValidateAttribute(string max) : base()
        {
            MaxValue = decimal.Parse(max);
        }

        protected override ValidationResult IsValid(object value, ValidationContext validationContext)
        {

            // 檢查必填項目
            var result = base.IsValid(value, validationContext);
            if (result != null) return result;

            // 檢查金額格式
            string moneyStr = Convert.ToString(value);
            decimal money = 0.0M;
            string errorMessage = string.Empty;

            if (decimal.TryParse(moneyStr, out money) == false)
            {
                errorMessage = "格式錯誤。";
            }
            else if (money < 1)
            {
                errorMessage = "不可小於1元。";
            }
            else if ((money % 1) > 0)
            {
                errorMessage = "不可有小數位。";
            }
            else if (money > MaxValue)
            {
                errorMessage = string.Format("超過{0:N0}上限。", MaxValue);
            }

            if (string.IsNullOrWhiteSpace(errorMessage) == false)
            {
                var mesg = string.Format(@"{0} {1}", validationContext.DisplayName, errorMessage);
                return new ValidationResult(mesg, new string[] { validationContext.MemberName });
            }

            return ValidationResult.Success;
        }

        /// <summary>
        /// 金額上限
        /// </summary>
        public decimal MaxValue { get; private set; }

    }
}
