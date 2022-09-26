using Newtonsoft.Json;
using ReportsAPISite.Models.Excel;
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
        /// Generates the Building And Project Report
        /// </summary>
        /// <param name="model"> The input along with the data </param>
        /// <param name="type"> This parameter indicates whether the report is for Building or for Project </param>
        /// <returns></returns>
        [HttpPost]
        [Route("api/report/{type}")]
        public HttpResponseMessage GetReport(SpreadSheetReport model, string type)
        {
            log.Info($"GetReport : Started, Report type - {type}");
            log.Info($"GetReport : Building Id - {model?.BuildingId}");
            var result = new HttpResponseMessage();

            try
            {
                if (type == Utils.Constants.Building)
                {
                    ModelState.Remove("model.ProjectReportData.Layout");
                    ModelState.Remove("model.ProjectReportData.ProjectData");
                }
                else if (type == Utils.Constants.Project)
                {
                    ModelState.Remove("model.BuildingReportData.ReportData");
                }

                if (!ModelState.IsValid)
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ModelState);
                }

                log.Info($"Input Model Recieved : {JsonConvert.SerializeObject(model)}");
                log.Info($"Generating Report : Started");
                var reportResult = createExcelDocument.GetSpreadsheetReport(model, type);
                result = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(reportResult.Content)
                };

                log.Info($"Generating Report : Completed");
                result.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse($"attachment; filename={type}Report.xlsx");
                result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            }
            catch (Exception ex)
            {
                result = new HttpResponseMessage(HttpStatusCode.InternalServerError)
                {
                    Content = new StringContent(ex.ToString())
                };
            }

            return result;
        }
    }
}