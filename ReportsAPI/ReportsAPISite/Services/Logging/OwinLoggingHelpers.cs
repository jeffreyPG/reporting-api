using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using Microsoft.Owin;
using Serilog;

namespace ReportsAPISite.Services.Logging
{
    public static class OwinLoggingHelpers
    {
        public static LoggerConfiguration DestructureOwinModels(this LoggerConfiguration config)
        {
            config
                .DestructureKeyValuePairs();

            return config;
        }

        public static object ToLogModel(this IOwinRequest request)
        {
            var content = request.Body.ReadAsString();
            var result = new
            {
                Path = request.Path.Value,
                request.Method,
                request.Headers,
                QueryParams = request.Query,
                QueryString = request.QueryString.Value,
                request.ContentType,
                ContentLength = content?.Length,
                request.Scheme,
                Host = request.Host.Value,
                Content = content
            };

            return result;
        }

        public static object ToLogModel(this IOwinResponse response)
        {
            var result = new
            {
                response.StatusCode,
                response.Headers,
                response.ContentType,
                StatusDescription = ((HttpStatusCode)response.StatusCode).ToString()
            };

            return result;
        }

        private static LoggerConfiguration DestructureKeyValuePairs(this LoggerConfiguration loggerConfiguration)
        {
            loggerConfiguration.Destructure.ByTransforming<KeyValuePair<string, string[]>>(x =>
            {
                var value = String.Join(",", x.Value);

                return $"{x.Key}: {value}";
            });

            return loggerConfiguration;
        }

        private static string ReadAsString(this Stream stream)
        {
            if (stream.Length <= 0)
            {
                return string.Empty;
            }

            stream.Position = 0;

            var result = new StreamReader(stream).ReadToEnd();

            stream.Position = 0;

            return result;
        }
    }
}