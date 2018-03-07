using System.Collections;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using ExcelHelper.Models;

namespace ExcelHelper.Attribute
{
    public class NestedValidationAttribute : ValidationAttribute
    {
        /// <summary>
        /// Validates the specified entity with respect to the current validation attribute.
        /// </summary>
        /// <param name="entity">The entity to validate.</param>
        /// <param name="validationContext">The context information about the validation operation.</param>
        /// <returns>
        /// An instance of the <see cref="T:System.ComponentModel.DataAnnotations.ValidationResult" /> class.
        /// </returns>
        protected override ValidationResult IsValid(object entity, ValidationContext validationContext)
        {
            var displayName = validationContext.DisplayName;
            //紀錄錯誤屬性名稱
            var compositeResults = new CompositeValidationResult(string.Format("Validation for {0} failed!", displayName));

            //如果屬性是IEnumerable 則迴圈處理檢查到的錯誤
            if (entity is IEnumerable)
            {
                IEnumerable items = (IEnumerable)entity;

                var index = 0;
                foreach (var item in items)
                {
                    var validator = new DataAnnotationValidator();

                    validator.TryValidate(item);
                    var results = validator.ValidationResults;

                    if (results.Count != 0)
                    {
                        results.ForEach(x => compositeResults.AddResult(x, displayName, index));
                        index++;
                    }
                }
                //如果index 不為0表示有錯誤
                if (index > 0)
                    return compositeResults;
                else
                    return ValidationResult.Success;
            }
            else
            {
                //不為 IEnumerable 則做一般屬性檢查
                var validator = new DataAnnotationValidator();

                validator.TryValidate(entity);
                var results = validator.ValidationResults;
                //回應錯誤不為空則處理錯誤訊息
                if (results.Count != 0)
                {
                    results.ForEach(x => compositeResults.AddResult(x, displayName));
                    return compositeResults;
                }
            }

            return ValidationResult.Success;
        }
    }
}