using System;
using System.Web;

namespace ReportsAPISite.Services
{
    public class SimuwattContext
    {
        public SimuwattContext()
        {
            RequestId = Guid.NewGuid().ToString();
        }

        public string RequestId { get; }

        public string CorrelationId
        {
            get
            {
                var correlationId = string.IsNullOrEmpty(HttpContext.Current.Request.Headers["X-Correlation-ID"])
                    ? RequestId
                    : HttpContext.Current.Request.Headers["X-Correlation-ID"];

                return correlationId;
            }
        }
    }
}