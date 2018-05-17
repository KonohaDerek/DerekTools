using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ApiHelper.ModelValidator
{
    public class DataAnnotationValidator
    {
        public List<ValidationResult> ValidationResults
        {
            get;
            private set;
        }

        public bool TryValidate(object model)
        {
            var context = new ValidationContext(model, serviceProvider: null, items: null);
            var validationResults = new List<ValidationResult>();

            var result = Validator.TryValidateObject(model, context, validationResults, validateAllProperties: true);

            this.ValidationResults = new List<ValidationResult>();

            this.ValidationResults.AddRange(validationResults.OfType<ValidationResult>().Where(x => !(x is CompositeValidationResult)));

            var customValidationResults = validationResults.OfType<CompositeValidationResult>();

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