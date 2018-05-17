using System.Collections;
using System.ComponentModel.DataAnnotations;

namespace ApiHelper.ModelValidator
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
            var compositeResults = new CompositeValidationResult(string.Format("Validation for {0} failed!", displayName));

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
                if (index > 0)
                    return compositeResults;
                else
                    return ValidationResult.Success;
            }
            else
            {
                var validator = new DataAnnotationValidator();

                validator.TryValidate(entity);
                var results = validator.ValidationResults;

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