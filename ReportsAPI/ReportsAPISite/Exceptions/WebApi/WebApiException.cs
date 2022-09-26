using System;
using System.Collections.Generic;
using System.Linq;

namespace ReportsAPISite.Exceptions.WebApi
{
    public class WebApiException : Exception
    {
        public WebApiException()
        {
            StatusCode = 400;
            Errors = new List<Error>();
        }

        public string RequestId { get; set; }

        public List<Error> Errors { get; set; }

        public int StatusCode { get; set; }

        public override string Message
        {
            get
            {
                string message;

                switch (Errors.Count)
                {
                    case 0:
                        var contactSupport = string.IsNullOrEmpty(RequestId)
                            ? string.Empty
                            : $" Please contact support and give the following Request Id: '{RequestId}'";

                        message = $"An unexpected problem has occurred.{contactSupport}";

                        break;
                    case 1:
                        message = Errors.First().Message;

                        break;
                    default:
                        var errorMessages = Errors.Select(x => $"'{x.Message}'");
                        var joinedErrorMessages = string.Join(", ", errorMessages);

                        message = $"Errors occured: {joinedErrorMessages}";

                        break;
                }

                return message;
            }
        }
    }
}