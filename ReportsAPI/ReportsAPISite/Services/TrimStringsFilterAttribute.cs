using System.Web.Http.Controllers;
using System.Web.Http.Filters;
using ReportsAPISite.Extensions;

namespace ReportsAPISite.Services
{
    public class TrimStringsFilterAttribute : ActionFilterAttribute
    {
        public override void OnActionExecuting(HttpActionContext actionContext)
        {
            actionContext.ActionArguments.TrimStrings();
        }
    }
}