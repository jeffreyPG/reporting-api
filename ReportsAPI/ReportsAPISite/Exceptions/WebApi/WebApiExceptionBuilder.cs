namespace ReportsAPISite.Exceptions.WebApi
{
    public class WebApiExceptionBuilder : IWebApiExceptionBuilder
    {
        private readonly WebApiException _exception = new WebApiException();

        public IWebApiExceptionBuilder WithError(string message)
        {
            return WithError(message, null);
        }

        public IWebApiExceptionBuilder WithError(string message, string property)
        {
            var error = new Error
            {
                Message = message,
                Property = property
            };

            _exception.Errors.Add(error);

            return this;
        }

        public IWebApiExceptionBuilder WithStatusCode(int statusCode)
        {
            _exception.StatusCode = statusCode;

            return this;
        }

        public IWebApiExceptionBuilder WithRequestId(string requestId)
        {
            _exception.RequestId = requestId;

            return this;
        }

        public WebApiException Build()
        {
            return _exception;
        }
    }
}
