using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace ApiHelper.ModelValidator
{
    public class CompositeValidationResult : ValidationResult
    {
        private readonly List<ValidationResult> _results = new List<ValidationResult>();

        public IEnumerable<ValidationResult> Results
        {
            get
            {
                return _results;
            }
        }

        public CompositeValidationResult(string errorMessage)
            : base(errorMessage)
        {
        }

        public void AddResult(ValidationResult validationResult, string displayName)
        {
            var fieldName = validationResult.MemberNames.FirstOrDefault();
            if (fieldName != null)
            {
                var propertyName = "";
#if DEBUG
                propertyName = string.Format("{0}.{1}", displayName, fieldName);
#else
                propertyName =   string.Format("{0}", displayName);
#endif

                var errorMessage = validationResult.ErrorMessage.Replace(fieldName, propertyName);

                var memberNames = validationResult.MemberNames.Select(x => propertyName).ToList();
                var result = new ValidationResult(errorMessage, memberNames);

                _results.Add(result);
            }
        }

        public void AddResult(ValidationResult validationResult, string displayName, int index)
        {
            var fieldName = validationResult.MemberNames.FirstOrDefault();
            if (fieldName != null)
            {
                var propertyName = "";
#if DEBUG
                propertyName = string.Format("{0}[{2}].{1}", displayName, fieldName, index.ToString());
#else
                propertyName =   string.Format("{0}", displayName);
#endif
                var errorMessage = validationResult.ErrorMessage.Replace(fieldName, propertyName);

                var memberNames = validationResult.MemberNames.Select(x => propertyName).ToList();
                var result = new ValidationResult(errorMessage, memberNames);

                _results.Add(result);
            }
        }
    }
}