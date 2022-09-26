using System.Collections.Generic;

namespace ReportsAPISite.Exceptions.WebApi
{
    public class WebApiExceptionResult
    {
        public string RequestId { get; set; }

        public List<Error> Errors { get; set; }

        public int StatusCode { get; set; }

        public string Message { get; set; }
    }
}