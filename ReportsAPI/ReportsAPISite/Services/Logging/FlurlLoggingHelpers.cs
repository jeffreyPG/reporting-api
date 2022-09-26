using System.Collections.Generic;
using System.Net.Http;
using Flurl;
using Flurl.Http;
using Serilog;

namespace ReportsAPISite.Services.Logging
{
    public static class FlurlLoggingHelpers
    {
        public static LoggerConfiguration IncludeFlurl(this LoggerConfiguration config)
        {
            FlurlHttp.Configure(c =>
            {
            });

            config
                .DestructureHttpRequestMessage()
                .DestructureHttpResponseMessage()
                .DestructureKeyValuePairs()
                .DestructureQueryParameter();

            return config;
        }

        private static LoggerConfiguration DestructureQueryParameter(this LoggerConfiguration config)
        {
            config.Destructure.ByTransforming<QueryParameter>(x =>
            {
                var value = string.Join(",", x.Value);

                return $"{x.Name}: {value}";
            });

            return config;
        }

        private static LoggerConfiguration DestructureKeyValuePairs(this LoggerConfiguration config)
        {
            config.Destructure.ByTransforming<KeyValuePair<string, IEnumerable<string>>>(x =>
            {
                var value = string.Join(",", x.Value);

                return $"{x.Key}: {value}";
            });

            return config;
        }

        private static LoggerConfiguration DestructureHttpRequestMessage(this LoggerConfiguration config)
        {
            config.Destructure.ByTransforming<HttpRequestMessage>(x => new
            {
                CompleteURL = x.RequestUri.OriginalString,
                Method = x.Method.ToString(),
                x.RequestUri.Host,
                x.RequestUri.Scheme,
                x.Headers,
                x.RequestUri.Query
            });

            return config;
        }

        private static LoggerConfiguration DestructureHttpResponseMessage(this LoggerConfiguration config)
        {
            config.Destructure.ByTransforming<HttpResponseMessage>(x => new
            {
                StatusCode = (int)x.StatusCode,
                x.Headers
            });

            return config;
        }

    }
}