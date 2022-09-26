using System;
using System.Threading.Tasks;
using ReportsAPISite.Services;
using Microsoft.Owin;
using ReportsAPISite.Services.Logging;
using ReportsAPISite.Exceptions.WebApi;
using Newtonsoft.Json;

namespace ReportsAPISite.Middleware
{
    public class GlobalExceptionMiddleware : OwinMiddleware
    {
        private readonly OwinMiddleware _next;
        private readonly SimuwattContext _ezContext;

        public GlobalExceptionMiddleware(OwinMiddleware next, SimuwattContext ezContext) : base(next)
        {
            _next = next;
            _ezContext = ezContext;
        }

        public override async Task Invoke(IOwinContext context)
        {
            try
            {
                await _next.Invoke(context);
            }
            catch (OperationCanceledException exception)
            {
                WebApiExceptionResult result;
                Logger.LogWarning("the operation was canceled");
                result = new WebApiExceptionBuilder()
                    .WithStatusCode(500)
                    .WithRequestId(_ezContext.RequestId)
                    .Build()
                    .ToResult();
                DisplayError(context, result);
            }
            catch (Exception exception)
            {
                WebApiExceptionResult result;

                if (exception is WebApiException webApiExcpetion)
                {
                    webApiExcpetion.RequestId = _ezContext.RequestId;

                    result = webApiExcpetion.ToResult();
                }
                else
                {
                    Logger.LogException(exception);

                    result = new WebApiExceptionBuilder()
                        .WithStatusCode(500)
                        .WithRequestId(_ezContext.RequestId)
                        .Build()
                        .ToResult();
                }
                DisplayError(context, result);
            }
        }

        private static void DisplayError(IOwinContext context, WebApiExceptionResult result)
        {
            var settings = new JsonSerializerSettings
            {
                Formatting = Formatting.Indented,
                NullValueHandling = NullValueHandling.Ignore
            };

            var json = JsonConvert.SerializeObject(result, settings);

            context.Response.Headers.Add("EZ-WebApiException", new[] { "true" });
            context.Response.StatusCode = result.StatusCode;
            context.Response.Write(json);
        }
    }
}