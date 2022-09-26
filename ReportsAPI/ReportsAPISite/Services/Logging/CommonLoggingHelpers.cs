using System;
using System.Diagnostics;
using System.Reflection;
using Serilog;

namespace ReportsAPISite.Services.Logging
{
    public static class CommonLoggingHelpers
    {
        public static LoggerConfiguration EnrichWithCommonProperties(this LoggerConfiguration config, string environmentName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var application = assembly.GetName().Name;
            var version = FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion;

            config.Enrich.WithProperty("Application", application)
                .Enrich.WithProperty("Version", version)
                .Enrich.WithProperty("Machine", Environment.MachineName)
                .Enrich.WithProperty("Environment", environmentName);

            return config;
        }
    }
}