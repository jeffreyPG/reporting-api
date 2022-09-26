using System.Web.Http;

namespace ReportsAPISite.Endpoints.Samples
{
    public partial class SamplesController
    {
        [HttpPost]
        [Route("Samples/Add/")]
        public void Add(string sampleName)
        {
            // notice this method name is the same as the filename

            // and notice that the class name is NOT the filename, but we are using partial files for this

            // notice we use HttpPost since this is adding data

            // notice the route is ControllerName/MethodName
        }
    }
}