using System;
using System.Threading;
using Microsoft.Owin.BuilderProperties;
using Owin;

namespace ReportsAPISite.Services
{
    public static class AppBuilderExtensions
    {
        public static void OnAppDisposing(this IAppBuilder app, Action callback)
        {
            var properties = new AppProperties(app.Properties);
            var token = properties.OnAppDisposing;

            if (token != CancellationToken.None)
            {
                token.Register(callback);
            }
        }
    }
}