namespace ReportsAPISite.Exceptions.WebApi
{
    public static class WebApiExceptionHelpers
    {
        public static WebApiExceptionResult ToResult(this WebApiException exception)
        {
            var result = new WebApiExceptionResult
            {
                RequestId = exception.RequestId,
                Errors = exception.Errors,
                Message = exception.Message,
                StatusCode = exception.StatusCode
            };

            return result;
        }

        public static WebApiException ToException(this WebApiExceptionResult result)
        {
            var exception = new WebApiException
            {
                RequestId = result.RequestId,
                Errors = result.Errors,
                StatusCode = result.StatusCode
            };

            return exception;
        }
    }
}