using ReportsAPISite.Services;
using ReportsAPISite.DependencyResolution;
using ReportsAPISite.Middleware;
using ReportsAPISite.Services.ConfigProvider;
using ReportsAPISite.Services.Hangfire;
using ReportsAPISite.Services.Logging;
using ReportsAPISite.Services.ResourceProvider;
using Hangfire;
using Hangfire.PostgreSql;
using Owin;
using Swashbuckle.Application;
using System.Web.Http;
using System.Web.Http.ExceptionHandling;

namespace ReportsAPISite
{
    public class Startup
    {
        public void Configuration(IAppBuilder appBuilder)
        {

            var container = IoC.Initialize();
            var dependencyResolver = new StructureMapWebApiDependencyResolver(container);
            var httpConfig = new HttpConfiguration { DependencyResolver = dependencyResolver };
            var config = container.GetInstance<IConfigProvider>();
            var ezContext = container.GetInstance<SimuwattContext>();
            var dbUp = new DbUpProcessor(config, container.GetInstance<IResourceProvider>(), true);

            Logger.Initialize(config);
            Logger.LogApplicationStarted();

            appBuilder.OnAppDisposing(Logger.LogApplicationStopped);

            dbUp.EnsureDbVersion();

            httpConfig.Filters.Add(new TrimStringsFilterAttribute());
            httpConfig.Services.Replace(typeof(IExceptionHandler), new PassthroughExceptionHandler());
            httpConfig.MapHttpAttributeRoutes();

            httpConfig
                .EnableSwagger(x =>
                {
                    x.DescribeAllEnumsAsStrings();
                    x.SingleApiVersion("v1", config.SiteName);
                    x.ApiKey("apiKey")
                        .Description("API Key Authentication")
                        .Name("X-APIKey")
                        .In("header");
                })
                .EnableSwaggerUi(x => x.DocumentTitle(config.SiteName));

            appBuilder.Use<LoggingMiddleware>(container);
            appBuilder.Use<GlobalExceptionMiddleware>(ezContext);
            appBuilder.Use<APIKeyAuthenticationMiddleware>(config);

            var hasDatabase = !string.IsNullOrEmpty(config.DbConnectionString);
            if (hasDatabase)
            {
                GlobalConfiguration.Configuration.UsePostgreSqlStorage(config.DbConnectionString);
                GlobalConfiguration.Configuration.UseActivator(new HangfireJobActivator(container));
                appBuilder.UseHangfireDashboard("/hangfire", new DashboardOptions { });
                appBuilder.UseHangfireServer();
            }
            appBuilder.UseWebApi(httpConfig);
            
        }
    }
}