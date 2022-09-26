using Hangfire.Dashboard;
using Microsoft.Owin;

namespace ReportsAPISite.Services.Hangfire
{
    public class SecureIpAuthorizationFilter : IDashboardAuthorizationFilter
    {
        private readonly string _secureIpAddress;

        public SecureIpAuthorizationFilter(string secureIpAddress)
        {
            _secureIpAddress = secureIpAddress;
        }

        public bool Authorize(DashboardContext context)
        {
            var owinContext = new OwinContext(context.GetOwinEnvironment());
            var userIpAddress = owinContext.Request.Headers["X-Forwarded-For"] ?? owinContext.Request.RemoteIpAddress ?? "";

            var userIsLocal = userIpAddress == "127.0.0.1" || userIpAddress == "::1";
            var userIsSecure = userIpAddress == _secureIpAddress;

            if (userIsLocal || userIsSecure)
            {
                return true;
            }
            return false;
        }
    }
}