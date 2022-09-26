using System.Web.Http;

namespace ReportsAPISite.Endpoints.Samples
{
    public partial class SamplesController
    {
        [HttpGet]
        [Route("Samples/Get/")]
        public SampleModel Get(int id)
        {
            // notice this method name is the same as the filename

            // and notice that the class name is NOT the filename, but we are using partial files for this

            // notice we use HttpGet since this is not modifying data

            // notice the route is ControllerName/MethodName

            return new SampleModel{Name = id.ToString()};
        }
    }

    public class SampleModel
    {
        public string Name { get; set; }
    }
}