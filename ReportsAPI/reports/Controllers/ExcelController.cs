 using reports.Models;
using System.Collections.Generic;
using System.Web.Mvc;
using reports.Excel;
using System.Web.Http;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net;
using System.Net.Http.Headers;
using System;
using System.IO;
using OfficeOpenXml;
using log4net;
using System.Linq;

namespace reports.Controllers
{
    /// <summary>
    /// Exposes all excel realted API
    /// </summary>
    public class ExcelController : ApiController
    {
        /// <summary>
        /// Log4net object
        /// </summary>
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        /// <summary>
        /// project excel property
        /// </summary>
        private readonly IProjectExcel projectExcel;

        /// <summary>
        /// Initalises the members of Excel controller
        /// </summary>
        /// <param name="projectExcel"></param>
        public ExcelController(IProjectExcel projectExcel)
        {
            this.projectExcel = projectExcel;
        }

        /// <summary>
        /// Populates an Excel template from NYC with data from a JSON Object
        /// </summary>
        /// <param name="jsonResult">The JSON object with data that will go in the Excel file.</param>
        /// <returns>Returns an Excel document</returns>
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/nycExcel")]
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

        /// <summary>
        /// Populates an Excel file with data
        /// </summary>
        /// <param name="jsonResult">The JSON object with data that will go in the Excel file.</param>
        /// <returns>Returns an Excel document</returns>
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/excel")]
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

                if(projectObj.orientation == "horizontal")
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
        
        /// <summary>
        /// Generates the Building And Project Report
        /// </summary>
        /// <param name="model"> The input along with the data </param>
        /// <param name="type"> This parameter indicates whether the report is for Building or for Project </param>
        /// <returns></returns>
        [System.Web.Http.HttpPost]
        [System.Web.Http.Route("api/report/{type}")]
        public HttpResponseMessage GetReport(SpreadSheetReport model, string type)
        {
            log.Info($"GetReport : Started, Report type - {type}");
            log.Info($"GetReport : Building Id - {model?.BuildingId}");
            var result = new HttpResponseMessage();

            try
            {
                if(type == Utils.Constants.Building)
                {
                    ModelState.Remove("model.ProjectReportData.Layout");
                    ModelState.Remove("model.ProjectReportData.ProjectData");
                } 
                else if(type == Utils.Constants.Project)
                {
                    ModelState.Remove("model.BuildingReportData.ReportData");
                }

                if (!ModelState.IsValid)
                {
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ModelState);
                }

                log.Info($"Input Model Recieved : {JsonConvert.SerializeObject(model)}");
                log.Info($"Generating Report : Started");
                var reportResult = this.projectExcel.GetSpreadsheetReport(model, type);
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