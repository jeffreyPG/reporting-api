namespace ReportsAPISite.Exceptions.WebApi
{
    public interface IWebApiExceptionBuilder
    {
        IWebApiExceptionBuilder WithError(string message);
        IWebApiExceptionBuilder WithError(string message, string property);
        IWebApiExceptionBuilder WithStatusCode(int statusCode);
        IWebApiExceptionBuilder WithRequestId(string requestId);
        WebApiException Build();
    }
}