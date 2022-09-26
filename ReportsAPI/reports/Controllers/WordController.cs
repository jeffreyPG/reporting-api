using System;
using System.IO;
using System.Web.Http;
using System.Net.Http;
using System.Web;
using System.Net;
using reports.Word;
using System.Net.Http.Headers;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using DocumentFormat.OpenXml;
using OpenXmlPowerTools;
using reports.Models;
using log4net;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using log4net.Config;
using System.Drawing;
using System.Configuration;
using reports.Services.Word;
using Microsoft.Office.Interop.Word;
using System.Web.Hosting;
using System.Collections.Generic;
using iTextSharp.text.pdf;

[assembly: log4net.Config.XmlConfigurator(ConfigFile = "log4net.config", Watch = true)]
namespace reports.Controllers
{
    public class WordController : ApiController
    {
        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        /// <summary>
        /// Replaces one Word document styles with styles from another Word Document
        /// </summary>
        /// <param name="generatedReport">The Word document where you want to change the styles.</param>
        /// <param name="storedStyledReport">The Word document where you want to take the styles from.</param>
        /// IsPdf a boolean parameter which will be set to true if pdf output is required and false if doc output is required
        /// <returns>Returns a Word document. The generatedReport .doc file with new styles.</returns>
        [HttpPost]
        [Route("api/word/replaceStyles")]
        public async Task<HttpResponseMessage> WordStylesAsync(bool IsPdf = true)
        {

            log.Info("WordStylesAsync : Started");
            var exMessage = string.Empty;
            var uploadPath = HttpContext.Current.Server.MapPath("~/generatedDocs");

            log.Info($"uploadPath: {uploadPath}");
            HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

            if (!Request.Content.IsMimeMultipartContent())
            {
                return Request.CreateResponse(HttpStatusCode.UnsupportedMediaType);
            }

            if (!Directory.Exists(uploadPath))
            {
                // if directory doesn't exist, create the directory
                Directory.CreateDirectory(uploadPath);
            }

            try
            {

                var createWordDocumentService = new CreateDocumentService();

                var provider = new MultipartFormDataStreamProvider(uploadPath);
                string generatedReportPath = null;
                string storedS3ReportPath = null;
                string storedS3Key = null;
                var options = string.Empty;
                bool reportStyles = false;
                // saves on local hard drive
                await Request.Content.ReadAsMultipartAsync(provider);
                // for file file
                foreach (var file in provider.FileData)
                {
                    generatedReportPath = file.LocalFileName;
                }
                // for the form data
                foreach (var key in provider.FormData.AllKeys)
                {
                    foreach (var val in provider.FormData.GetValues(key))
                    {

                        log.Info("WordStylesAsync : key=" + key);
                        log.Info("WordStylesAsync : val=" + val);
                        if (key == "storedStyledReport")
                        {
                            storedS3Key = val;
                        }

                        if (key == "reportStyles")
                        {
                            reportStyles = val.ToString() == "1" ? true : false;
                        }

                        if (key == "options")
                        {
                            options = val;
                        }

                        log.Info($"key:value={key}:{val}");
                    }
                }

                HtmlData data = JsonConvert.DeserializeObject<HtmlData>(options);

                // TODO: encapsulate all of this s3 thing in a new service
                var isStaticStyle = ConfigurationManager.AppSettings["IsStaticStyle"];
                var s3object = new GetS3Object();

                if (isStaticStyle == "true")
                {
                    storedS3ReportPath = HttpContext.Current.Server.MapPath("~/StylesDoc/DocStyles.docx");
                }
                else
                {
                    // get document from s3
                    //storedS3ReportPath = storedS3Key;
                    storedS3ReportPath = s3object.ReadObjectData(storedS3Key);
                }

                log.Info("WordStylesAsync : storedS3ReportPath:" + storedS3ReportPath);
                log.Info("WordStylesAsync : generatedReportPath:" + generatedReportPath);
                log.Info("WordStylesAsync : reportStyles:" + reportStyles);

                // replace styles now that you have the file paths for each
                ReplaceDocStyles.ReplaceStyles(storedS3ReportPath, generatedReportPath);

                // replace theme content, too (handles color, fonts, etc.)
                ReplaceDocStyles.CopyThemeContent(storedS3ReportPath, generatedReportPath);

                if (reportStyles)
                {
                    //Replace document margin
                    ReplaceDocStyles.ReplaceMargin(storedS3ReportPath, generatedReportPath);
                    //replace other styles
                    ReplaceDocStyles.ReplaceOtherStyles(storedS3ReportPath, generatedReportPath, data);
                }



                // remove temporary files
                if (!(isStaticStyle == "true"))
                {
                    if (File.Exists(storedS3ReportPath))
                    {
                        File.Delete(storedS3ReportPath);
                    }
                }

                log.Info("WordStylesAsync : completed");
                byte[] bytes = File.ReadAllBytes(generatedReportPath);
                // byte[] fileBytes = getFileBytesFromDB();
                var tmpFile = Path.GetTempFileName();
                File.WriteAllBytes(tmpFile, bytes);

                Application app = new Application();
                Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(generatedReportPath);

                // Save DOCX into a PDF
                string pdfPath = uploadPath + "\\Report_" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".pdf";
                //var pdfPath = @"D:\Report_"+ DateTime.Now.ToString("MMddyyyyHHmmss") + ".pdf";
                doc.SaveAs2(pdfPath, WdSaveFormat.wdFormatPDF);
                string pdfPathDoc = uploadPath + "\\Report_" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".docx";


                if (IsPdf)
                {// Save DOCX into a PDF
                    doc.SaveAs2(pdfPath, WdSaveFormat.wdFormatPDF);
                    doc.Close();
                    // Read the file into a Byte Array
                    byte[] bytesPdf = File.ReadAllBytes(pdfPath);
                    response.Content = new ByteArrayContent(bytesPdf);
                    response.Content.Headers.ContentLength = bytesPdf.LongLength;

                    // Set the Content Disposition Header Value and File name
                    response.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.pdf");
                    response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
                }
                else
                {// Save DOCX into a RTF
                    doc.SaveAs2(pdfPathDoc, WdSaveFormat.wdFormatDocumentDefault);
                    doc.Close();
                    // Read the file into a Byte Array
                    byte[] bytesDocx = File.ReadAllBytes(pdfPathDoc);
                    response.Content = new ByteArrayContent(bytesDocx);
                    response.Content.Headers.ContentLength = bytesDocx.LongLength;

                    // Set the Content Disposition Header Value and File name
                    response.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.docx");
                    response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/msword");
                }

                app.Quit(); // VERY IMPORTANT: do this to close the MS Word instance
                //byte[] pdfFileBytes = File.ReadAllBytes(pdfPath);
                File.Delete(tmpFile);

                if (File.Exists(generatedReportPath))
                {
                    File.Delete(generatedReportPath);
                }


                //File.WriteAllBytes(@"D:\hello.pdf", bytes);
                //return generated report with new styles
                return response;
            }
            catch (Exception e)
            {
                log.Error("WordStylesAsync : failed, error:" + e.ToString());
                return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e);
            }
        }

        /// <summary>
        /// create a word document using openXML from html snippet
        /// </summary>
        /// <param name="jsonResult">The Html  where you want to change the styles.</param>
        /// <returns>Returns a Word document. The generatedReport .doc file with new styles.</returns>
        [System.Web.Http.HttpPost]
        [Route("api/word/createDocument")]
        public HttpResponseMessage CreateWordDocument(JObject jsonResult)
        {

            // TODO: use IoC
            var createWordDocumentService = new CreateDocumentService();
            log.Info("CreateWordDocument : Started");

            HttpResponseMessage result = new HttpResponseMessage();
            try
            {

                log.Info("CreateWordDocument : incomingArgs=" + jsonResult.ToString());
                HtmlData data = JsonConvert.DeserializeObject<HtmlData>(jsonResult.ToString());
                if (data != null)
                {

                    var strHTML = data.HtmlString;
                    log.Info("CreateWordDocument : html body=" + strHTML);
                    log.Info("CreateWordDocument : isIncludeTOC=" + data.isIncludeTOC);
                    log.Info("CreateWordDocument : isShowPageNumber=" + data.isShowPageNumber);
                    log.Info("CreateWordDocument : isPageNumberDisplayOnHeader=" + data.isPageNumberDisplayOnHeader);
                    log.Info("CreateWordDocument : TOCDept=" + data.TOCDept);
                    log.Info("CreateWordDocument : pageNumberPosition=" + data.pageNumberPosition);
                    log.Info("CreateWordDocument : ReportStyles=" + data.ReportStyles);
                    var uploadPath = HttpContext.Current.Server.MapPath("~/GeneratedDocs");

                    if (!Directory.Exists(uploadPath))
                    {
                        // if directory doesn't exist, create the directory
                        Directory.CreateDirectory(uploadPath);
                    }

                    string filename = uploadPath + "\\Report_" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".docx"; //temp file
                    if (File.Exists(filename))
                    {
                        File.Delete(filename);
                    }

                    using (MemoryStream generatedDocument = new MemoryStream())
                    {

                        using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                        {

                            MainDocumentPart mainPart = package.MainDocumentPart;
                            if (mainPart == null)
                            {
                                mainPart = package.AddMainDocumentPart();
                                new DocumentFormat.OpenXml.Wordprocessing.Document(new Body()).Save(mainPart);
                            }

                            //convert html to word open xml
                            HtmlToOpenXml.HtmlConverter converter = new HtmlToOpenXml.HtmlConverter(mainPart);
                            Body body = mainPart.Document.Body;
                            var paragraphs = converter.Parse(strHTML);
                            for (int i = 0; i < paragraphs.Count; i++)
                            {
                                body.Append(paragraphs[i]);
                            }

                            mainPart.Document.Save();

                            createWordDocumentService.SetHeaderContent(package, data.HeaderContent);
                            createWordDocumentService.SetFooterContent(package, data.FooterContent);

                        }

                        File.WriteAllBytes(filename, generatedDocument.ToArray());
                        if (data.isIncludeTOC)
                        {
                            createWordDocumentService.CreateTOC(filename, data.TOCDept);
                        }
                        result = new HttpResponseMessage(HttpStatusCode.OK);
                        // Read the file into a Byte Array
                        byte[] bytes = File.ReadAllBytes(filename);
                        result.Content = new ByteArrayContent(bytes);
                        result.Content.Headers.ContentLength = bytes.LongLength;

                        // Set the Content Disposition Header Value and File name
                        result.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.docx");
                        result.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
                        if (File.Exists(filename))
                        {
                            File.Delete(filename);
                        }
                        log.Info("CreateWordDocument : Completed");
                    }
                }
                else
                {
                    log.Info("CreateWordDocument : No data specified in input request");
                    return Request.CreateErrorResponse(HttpStatusCode.BadRequest, "No data specified in input request");
                }
            }
            catch (Exception ex)
            {
                log.Error("CreateWordDocument : failed, error-" + ex.ToString());
                return Request.CreateErrorResponse(HttpStatusCode.BadRequest, ex);
            }
            return result;
        }


        [HttpPost]
        [Route("api/word/mergePdf")]
        public HttpResponseMessage MergePdf(string attachment_list)
        {
            var httpContext = HttpContext.Current;
            HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.OK);

            // Check for any uploaded file  
            if (httpContext.Request.Files.Count > 0)
            {
                byte[] pdfFileBytes = null;
                List<byte[]> vs = new List<byte[]>();

                //Loop through uploaded files  
                for (int i = 0; i < httpContext.Request.Files.Count; i++)
                {
                    HttpPostedFile httpPostedFile = httpContext.Request.Files[i];
                    if (httpPostedFile != null)
                    {
                        // Construct file save path  
                        var fileSavePath = Path.Combine(HostingEnvironment.MapPath("~/generatedDocs"), httpPostedFile.FileName);
                        // Save the uploaded file  
                        httpPostedFile.SaveAs(fileSavePath);

                        pdfFileBytes = File.ReadAllBytes(fileSavePath);
                        vs.Add(pdfFileBytes);
                    }
                    pdfFileBytes = concatAndAddContent(vs);
                }



                string[] filesList = attachment_list.Split(',');

                foreach (var file in filesList)
                {
                    try
                    {
                        var s3object = new GetS3Object();

                        string storedS3ReportPath = s3object.ReadObjectData(file);
                        pdfFileBytes = File.ReadAllBytes(storedS3ReportPath);
                        vs.Add(pdfFileBytes);
                        pdfFileBytes = concatAndAddContent(vs);
                    }
                    catch
                    {

                    }
                }

                //Path where mereged pdf will save.
                var uploadPath = HttpContext.Current.Server.MapPath("~/generatedDocs");
                string pdfPath = uploadPath + "\\Report_" + DateTime.Now.ToString("MMddyyyyHHmmss") + ".pdf";
                File.WriteAllBytes(pdfPath, pdfFileBytes);

                // Read the file into a Byte Array
                byte[] bytesPdf = File.ReadAllBytes(pdfPath);
                response.Content = new ByteArrayContent(bytesPdf);
                response.Content.Headers.ContentLength = bytesPdf.LongLength;

                // Set the Content Disposition Header Value and File name
                response.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.pdf");
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

            }

            // Return status code

            return response;
        }

        public static byte[] concatAndAddContent(List<byte[]> pdfByteContent)
        {

            using (var ms = new MemoryStream())
            {
                // Need to install iTextSharp
                //Install-Package iTextSharp -Version 5.5.13.2
                using (var doc = new iTextSharp.text.Document())
                {
                    using (var copy = new PdfSmartCopy(doc, ms))
                    {
                        doc.Open();

                        //Loop through each byte array
                        foreach (var p in pdfByteContent)
                        {

                            //Create a PdfReader bound to that byte array
                            using (var reader = new PdfReader(p))
                            {

                                //Add the entire document instead of page-by-page
                                copy.AddDocument(reader);
                            }
                        }

                        doc.Close();
                    }
                }

                //Return just before disposing
                return ms.ToArray();
            }
        }

    }
}
