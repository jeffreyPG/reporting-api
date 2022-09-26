using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace ReportsAPISite.Endpoints.Word
{
    public partial class WordController
    {

        [HttpPost]
        [Route("v2/Word/Create")]
        public HttpResponseMessage Create(CreateArgs args)
        {

            args.Validate();
            log.Info("CreateWordDocument : Started");
            var result = new HttpResponseMessage();

            try
            {

                log.Info("CreateWordDocument : Started");
                log.Info($"Incoming args: {args.ToString()}");

                var uploadPath = HttpContext.Current.Server.MapPath("~/GeneratedDocs");
                if (!Directory.Exists(uploadPath))
                {
                    // if directory doesn't exist, create the directory
                    Directory.CreateDirectory(uploadPath);
                }

                var fileNameFullPath = $"{uploadPath }\\Report_{DateTime.Now.ToString("MMddyyyyHHmmss")}.docx"; //temp file
                var fileName = $"\\Report_{DateTime.Now.ToString("MMddyyyyHHmmss")}.docx"; //temp file
                fileDocumentStorage.Delete(fileName);

                createWordDocument.CreateWordDocument(args, fileNameFullPath);

                var applyStyle = !string.IsNullOrEmpty(args.ReportStyleS3Id);

                if (applyStyle)
                {

                    var fileToDownload = createWordDocument.ReplaceDocumentStyles(args, fileNameFullPath);

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

    public class CreateArgs
    {

        public string HtmlString { get; set; }
        public bool IncludeToC { get; set; }
        public bool ShowPageNumber { get; set; }
        public string ToCDepth { get; set; }
        public bool PageNumberDisplayOnHeader { get; set; }
        public string PageNumberPosition { get; set; }

        public string HeaderImage { get; set; }
        public string HeaderContent { get; set; }
        public string FooterContent { get; set; }

        // if null or empty, dont apply
        public string ReportStyleS3Id { get; set; }

        public string OrganizationId { get; set; }
        public string BuildingId { get; set; }
        public string ReportId { get; set; }

        public string ToString()
        {
            // TODO: implement

            return HtmlString;
        }

        // TODO: implement
        public void Validate()
        {

        }
    }

}