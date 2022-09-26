using System.Web.Http;

namespace ReportsAPISite.Endpoints.Samples
{
    public partial class SamplesController
    {
        [HttpGet]
        [Route("Samples/GetAll/")]
        public void GetAll()
        {
            //var args = new PageArgs(); // the orderBy feature is for paged result sets
            //var orderBy = args.OrderBys.ToOrderByClause("name"); // defaults to what you put here, ascending

            //var sql = $@"
            //    select id,
            //        name,
            //        description
            //    from roles
            //    {orderBy}";
        }
    }
}