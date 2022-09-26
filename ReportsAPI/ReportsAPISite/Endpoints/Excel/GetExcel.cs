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

        /// <summary>
        /// Populates an Excel file with data
        /// </summary>
        /// <param name="jsonResult">The JSON object with data that will go in the Excel file.</param>
        /// <returns>Returns an Excel document</returns>
        [HttpPost]
        [Route("api/excel")]
        public HttpResponseMessage Excel(JObject jsonResult)
        {
            ProjectExcel excel = new ProjectExcel();
            ProjectData projectObj = JsonConvert.DeserializeObject<ProjectData>(jsonResult.ToString());

            // if the type is project
            // if the orientation is vertical
            // else if the orientation is horizontal

            // else if the type is data
            // if the orientation is vertical
            // else if the orientation is horizontal

            // else if the type is BOM
            // if the orientation is vertical
            // else if the orientation is horizontal

            var result = new HttpResponseMessage();

            if (projectObj.type == "project")
            {

                if (projectObj.orientation == "horizontal")
                {
                    result = new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ByteArrayContent(excel.GenerateHorizontalProject(projectObj))
                    };
                }

                if (projectObj.orientation == "vertical")
                {
                    result = new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new ByteArrayContent(excel.GenerateVerticalProject(projectObj))
                    };
                }
            }

            result.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.xlsx");
            result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

            return result;
        }


    }
}