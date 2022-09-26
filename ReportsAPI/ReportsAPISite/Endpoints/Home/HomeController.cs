using System.Web.Http;
using System.Web.Http.Description;

namespace ReportsAPISite.Endpoints.Home
{
    public class HomeController : ApiController
    {
        [HttpGet]
        [Route("")]
        [ApiExplorerSettings(IgnoreApi = true)]
        public IHttpActionResult Get()
        {
            var location = $"{Request.RequestUri.Scheme}://{Request.RequestUri.Authority}/swagger/ui/index";

            return Redirect(location);
        }
    }
}