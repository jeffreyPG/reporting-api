using log4net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using reports.Endpoints.CreateWordDocumentController.Models;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace reports.Endpoints.CreateWordDocumentController
{
    public partial class CreateWordDocumentController
    {

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        
        [HttpPost]
        [Route("v2/CreateWordDocument/Create")]
        public HttpResponseMessage Create([FromBody] CreateArgs data)
        {

            data.Validate();
            log.Info("CreateWordDocument : Started");
            HttpResponseMessage result = new HttpResponseMessage();

            try
            {

                log.Info("CreateWordDocument : Started");
                log.Info($"Incoming args: {data.ToString()}");

                var uploadPath = HttpContext.Current.Server.MapPath("~/GeneratedDocs");
                if (!Directory.Exists(uploadPath))
                {
                    // if directory doesn't exist, create the directory
                    Directory.CreateDirectory(uploadPath);
                }

                var fileNameFullPath = $"{uploadPath }\\Report_{DateTime.Now.ToString("MMddyyyyHHmmss")}.docx"; //temp file
                var fileName = $"\\Report_{DateTime.Now.ToString("MMddyyyyHHmmss")}.docx"; //temp file
                fileDocumentStorage.Delete(fileName);

                createDocumentService.CreateWordDocument(data, fileNameFullPath);
                
                var applyStyle = !string.IsNullOrEmpty(data.ReportStyleS3Id);

                if(applyStyle)
                {

                    var fileToDownload = createDocumentService.ReplaceDocumentStyles(data, fileNameFullPath);

                    result = new HttpResponseMessage(HttpStatusCode.OK);
                    // Read the file into a Byte Array
                    byte[] bytes = File.ReadAllBytes(fileToDownload);
                    result.Content = new ByteArrayContent(bytes);
                    result.Content.Headers.ContentLength = bytes.LongLength;

                    // Set the Content Disposition Header Value and File name
                    result.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.docx");
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

                    fileDocumentStorage.Delete(fileName);
                }
                else
                {

                    result = new HttpResponseMessage(HttpStatusCode.OK);
                    // Read the file into a Byte Array
                    byte[] bytes = File.ReadAllBytes(fileNameFullPath);
                    result.Content = new ByteArrayContent(bytes);
                    result.Content.Headers.ContentLength = bytes.LongLength;

                    // Set the Content Disposition Header Value and File name
                    result.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.docx");
                    result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

                    fileDocumentStorage.Delete(fileName);

                    log.Info("CreateWordDocument : Completed");
                }

            }
            catch (Exception ex)
            {
                log.Error("CreateWordDocument : failed, error-" + ex.ToString());
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex);
            }

            return result;

        }

    }

}