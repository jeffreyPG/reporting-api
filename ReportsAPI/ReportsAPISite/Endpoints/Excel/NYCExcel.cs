using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ReportsAPISite.Services.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace ReportsAPISite.Endpoints.Excel
{
    public partial class ExcelController
    {

        [HttpPost]
        [Route("api/Excel/nycExcel")]
        public HttpResponseMessage NYCExcel(JObject jsonResult)
        {
            NYCExcel excel = new NYCExcel();
            NYCData NYCObj = JsonConvert.DeserializeObject<NYCData>(jsonResult.ToString());

            HttpResponseMessage result = new HttpResponseMessage();

            result = new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(excel.PopulateNYCData(NYCObj))
            };

            result.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.xlsx");
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            return result;
        }

    }
}