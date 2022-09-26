using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Owin;
using ReportsAPISite.Services.ConfigProvider;
using ReportsAPISite.Exceptions.WebApi;

namespace ReportsAPISite.Middleware
{
    public class APIKeyAuthenticationMiddleware : OwinMiddleware
    {
        private readonly IConfigProvider _config;
        private readonly OwinMiddleware _next;

        public APIKeyAuthenticationMiddleware(OwinMiddleware next, IConfigProvider config) : base(next)
        {
            _next = next;
            _config = config;
        }

        public override async Task Invoke(IOwinContext context)
        {
            var apiKey = _config.ReportsAPISiteKey;
            var authenticationNotNeeded = string.IsNullOrEmpty(apiKey) ||
                                          context.Request.Path.Value == "/" ||
                                          context.Request.Path.Value.StartsWith("/hangfire") ||
                                          context.Request.Path.Value.StartsWith("/swagger");

            if (authenticationNotNeeded)
            {
                await _next.Invoke(context);

                return;
            }

            var isAuthenticated = context.Request.Headers["X-APIKey"] == apiKey ||
                                  context.Request.Uri.ParseQueryString()["api_key"] == apiKey;

            if (isAuthenticated)
            {
                await _next.Invoke(context);

                return;
            }

            throw new WebApiExceptionBuilder()
                .WithError("Unauthorized")
                .WithStatusCode(401)
                .Build();
        }
    }
}