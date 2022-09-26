using System.Linq;
using ReportsAPISite.Exceptions.WebApi;
using ReportsAPISite.Services.Validators;

namespace ReportsAPISite.Services.Validators
{
    public class Validator<T> : IValidator<T>
    {
        private readonly WebApiExceptionBuilder _exceptionBuilder;

        public Validator(T candidate)
        {
            Candidate = candidate;
            _exceptionBuilder = new WebApiExceptionBuilder();
        }

        public T Candidate { get; }

        public void AddError(string message, string property)
        {
            _exceptionBuilder.WithError(message, property);
        }

        public void ThrowIfInvalid()
        {
            var webApiException = _exceptionBuilder.Build();

            if (webApiException.Errors.Any())
            {
                throw webApiException;
            }
        }
    }
}