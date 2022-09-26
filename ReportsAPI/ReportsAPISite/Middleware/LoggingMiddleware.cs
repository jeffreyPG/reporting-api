using System.Diagnostics;
using System.Threading.Tasks;
using ReportsAPISite.Services;
using Microsoft.Owin;
using ReportsAPISite.Services.Logging;
using StructureMap;

namespace ReportsAPISite.Middleware
{
    public class LoggingMiddleware : OwinMiddleware
    {
        private readonly OwinMiddleware _next;
        private readonly IContainer _container;

        public LoggingMiddleware(OwinMiddleware next, IContainer container) : base(next)
        {
            _next = next;
            _container = container;
        }

        public override async Task Invoke(IOwinContext context)
        {
            var isSwagger = context.Request.Path.Value.StartsWith("/swagger") ||
                            context.Request.Path.Value == "/";

            if (isSwagger)
            {
                await _next.Invoke(context);

                return;
            }

            var stopwatch = new Stopwatch();
            var ezContext = _container.GetInstance<SimuwattContext>();

            Logger.LogRequestStarted(context, stopwatch, ezContext);

            await _next.Invoke(context);

            Logger.LogRequestCompleted(context, stopwatch);
        }
    }
}