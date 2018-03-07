using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ExcelHelper.Models
{
    /// <summary>
    /// 資料註釋驗證器
    /// </summary>
    public class DataAnnotationValidator
    {
        /// <summary>
        /// 驗證結果清單
        /// </summary>
        public List<ValidationResult> ValidationResults
        {
            get;
            private set;
        }

        /// <summary>
        /// 檢驗
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        public bool TryValidate(object model)
        {
            var context = new ValidationContext(model, serviceProvider: null, items: null);
            var validationResults = new List<ValidationResult>();

            var result = Validator.TryValidateObject(model, context, validationResults, validateAllProperties: true);

            this.ValidationResults = new List<ValidationResult>();

            this.ValidationResults.AddRange(validationResults.OfType<ValidationResult>().Where(x => !(x is CompositeValidationResult)));

            var customValidationResults = validationResults.OfType<CompositeValidationResult>();

            //將各個驗證結果加入驗證結果清單
            foreach (var customValidationResult in customValidationResults)
            {
                if (customValidationResult.Results.Count() == 0 && !string.IsNullOrWhiteSpace(customValidationResult.ErrorMessage))
                {
                    this.ValidationResults.Add(new ValidationResult(customValidationResult.ErrorMessage));
                }
                this.ValidationResults.AddRange(customValidationResult.Results);
            }

            return result;
        }
    }
}