using System;
using System.Linq.Expressions;
using ReportsAPISite.Extensions;

namespace ReportsAPISite.Services.Validators.Validations
{
    public static class IsRequiredValidation
    {
        public static IValidator<T> IsRequired<T, TProperty>(this IValidator<T> validator, Expression<Func<T, TProperty>> property)
        {
            var candidate = property.GetCandidateProperties(validator.Candidate);
            
            var isNullOrEmpty = candidate.Value is string
                ? string.IsNullOrEmpty(candidate.Value as string)
                : candidate.Value == null;

            if (isNullOrEmpty)
            {
                validator.AddError($"{candidate.Name} is required.", candidate.Name);
            }

            return validator;
        }
    }
}