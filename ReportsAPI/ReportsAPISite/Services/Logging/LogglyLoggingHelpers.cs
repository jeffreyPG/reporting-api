using Loggly.Config;
using Serilog;

namespace ReportsAPISite.Services.Logging
{
    public static class LogglyLoggingHelpers
    {
        public static LoggerConfiguration ConsiderWritingToLoggly(this LoggerConfiguration loggerConfig, string logglyToken)
        {
            if (string.IsNullOrEmpty(logglyToken))
            {
                return loggerConfig;
            }

            var logglyConfig = LogglyConfig.Instance;

            logglyConfig.CustomerToken = logglyToken;
            logglyConfig.Transport.EndpointHostname = "logs-01.loggly.com";
            logglyConfig.Transport.EndpointPort = 443;
            logglyConfig.Transport.LogTransport = LogTransport.Https;

            loggerConfig.WriteTo.Loggly();

            return loggerConfig;
        }
    }
}