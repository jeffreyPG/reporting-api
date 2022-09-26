using System;
using System.Diagnostics;
using System.Net;
using ReportsAPISite.Services.ConfigProvider;
using Microsoft.Owin;
using Serilog;
using Serilog.Context;
using Serilog.Events;

namespace ReportsAPISite.Services.Logging
{
    public static class Logger
    {
        private static readonly Stopwatch ApplicationSw = new Stopwatch();

        public static void Initialize(IConfigProvider config)
        {

            var minimumLevel = (LogEventLevel)Enum.Parse(typeof(LogEventLevel), config.LoggingMinimumLevel);

            Log.Logger = new LoggerConfiguration()
                .Enrich.FromLogContext()
                .EnrichWithCommonProperties(config.EnvironmentName)
                .IncludeFlurl()
                .MinimumLevel.Is(minimumLevel)
                .MinimumLevel.Override("Hangfire.Server.ServerHeartbeatProcess", LogEventLevel.Information)
                .MinimumLevel.Override("Hangfire.Server.ServerJobCancellationWatcher", LogEventLevel.Information)
                .MinimumLevel.Override("Hangfire.PostgreSql.ExpirationManager", LogEventLevel.Information)
                .MinimumLevel.Override("Hangfire.Server.BackgroundServerProcess", LogEventLevel.Information)
                .MinimumLevel.Override("Hangfire.Processing.BackgroundExecution", LogEventLevel.Information)
                .DestructureOwinModels()
                .ConsiderWritingToLoggly(config.LogglyToken)
                .CreateLogger();
        }

        public static void LogApplicationStarted()
        {
            Log.Information("Application started: \"{Application:l} {Version:l}\" [{Environment:l}:{Machine:l}]");

            ApplicationSw.Start();
        }

        public static void LogApplicationStopped()
        {
            ApplicationSw.Stop();

            var durationInMinutes = ApplicationSw.ElapsedMilliseconds / 60000;
            Log.ForContext("DurationInMinutes", durationInMinutes)
                .Information("Application stopped: \"{Application:l} {Version:l}\" [{Environment:l}:{Machine:l}] ({DurationInMinutes:n0} m)");

            Log.CloseAndFlush();
        }

        public static void LogRequestStarted(IOwinContext context, Stopwatch stopwatch, SimuwattContext ezContext)
        {
            LogContext.PushProperty("RequestID", ezContext.RequestId);
            LogContext.PushProperty("CorrelationID", ezContext.CorrelationId);

            var request = context.Request.ToLogModel();

            Log.ForContext("Request", request, true)
                .Verbose("Request started: {Method:l} {path}", context.Request.Method, context.Request.Path);

            stopwatch?.Start();
        }

        public static void LogRequestCompleted(IOwinContext context, Stopwatch stopwatch)
        {
            stopwatch?.Stop();

            var durationInMilliseconds = stopwatch?.ElapsedMilliseconds;
            var response = context.Response.ToLogModel();

            Log.ForContext("Response", response, true)
                .Verbose("Request completed: {Method:l} {path} [{StatusCode} {StatusDescription:l}] ({DurationInMilliseconds:n0} ms)",
                    context.Request.Method,
                    context.Request.Path,
                    context.Response.StatusCode,
                    ((HttpStatusCode)context.Response.StatusCode).ToString(),
                    durationInMilliseconds);
        }

        public static void LogException(Exception exception)
        {
            Log.ForContext("Exception", exception, true)
                .Error("{Message}", exception.Message);
        }

        public static void LogWarning(string message)
        {
            Log.Warning(message);
        }
    }
}