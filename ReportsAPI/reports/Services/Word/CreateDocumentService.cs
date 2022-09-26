using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using reports.Endpoints.CreateWordDocumentController.Models;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using System.Drawing;
using System.Net;
using OpenXmlPowerTools;
using log4net;
using System.Net.Http;
using System.Configuration;
using reports.Word;
using HeaderContent = reports.Models.HeaderContent;
using FooterContent = reports.Models.FooterContent;
using reports.Services.Word.HeaderFunctions;
using reports.Services.Word.FooterFunctions;

namespace reports.Services.Word
{
    public class CreateDocumentService : ICreateWordDocument
    {

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public string CreateWordDocument(CreateArgs args, string fileName)
        {

            try
            {
                using (MemoryStream generatedDocument = new MemoryStream())
                {

                    using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                    {

                        MainDocumentPart mainPart = package.MainDocumentPart;
                        if (mainPart == null)
                        {
                            mainPart = package.AddMainDocumentPart();
                            new Document(new Body()).Save(mainPart);
                        }

                        //convert html to word open xml
                        HtmlToOpenXml.HtmlConverter converter = new HtmlToOpenXml.HtmlConverter(mainPart);
                        Body body = mainPart.Document.Body;
                        var paragraphs = converter.Parse(args.HtmlString);

                        for (int i = 0; i < paragraphs.Count; i++)
                        {
                            body.Append(paragraphs[i]);
                        }

                        mainPart.Document.Save();

                        //set header content
                        if (args.HeaderContent != null || (args.ShowPageNumber && args.PageNumberDisplayOnHeader))
                        {
                            bool isHeader = false;
                            string pos = args.PageNumberDisplayOnHeader ? args.PageNumberPosition : string.Empty;
                            string imagePath = string.Empty;
                            string headerCon = string.Empty;
                            if (args.HeaderContent != null)
                            {
                                imagePath = args.HeaderImage;
                                headerCon = args.HeaderContent;
                            }
                            if (!string.IsNullOrEmpty(imagePath))
                            {
                                isHeader = true;
                            }
                            if (!string.IsNullOrEmpty(headerCon))
                            {
                                isHeader = true;
                            }
                            if (!string.IsNullOrEmpty(pos))
                            {
                                isHeader = true;
                            }
                            if (isHeader)
                            {
                                SetHeaderContent(package, pos, imagePath, headerCon);
                            }
                        }

                        //set footer content
                        if (args.FooterContent != null || (args.ShowPageNumber && !args.PageNumberDisplayOnHeader))
                        {
                            string pos = args.ShowPageNumber && !args.PageNumberDisplayOnHeader ? args.PageNumberPosition : string.Empty;
                            string footerC = string.Empty;
                            if (args.FooterContent != null)
                            {
                                footerC = args.FooterContent;
                            }
                            SetFooterContent(package, footerC, pos);
                        }
                    }

                    File.WriteAllBytes(fileName, generatedDocument.ToArray());

                    if (args.IncludeToC)
                    {
                        CreateTOC(fileName, args.ToCDepth);
                    }

                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            
            return fileName;

        }
        
        public string ReplaceDocumentStyles(CreateArgs args, string fileName)
        {

            log.Info($"WordStylesAsync : Started with {fileName}");            
            var exMessage = string.Empty;
            var generatedReportPath = fileName;

            try
            {

                //var provider = new MultipartFormDataStreamProvider(uploadPath);
                
                //string storedS3ReportPath = null;
                //string storedS3Key = null;
                //bool reportStyles = false;
                // saves on local hard drive
                //await Request.Content.ReadAsMultipartAsync(provider);
                // for file file
                //foreach (var file in provider.FileData)
                //{
                //    generatedReportPath = file.LocalFileName;
                //}
                //// for the form data
                //foreach (var key in provider.FormData.AllKeys)
                //{

                //    foreach (var val in provider.FormData.GetValues(key))
                //    {
                //        log.Info("WordStylesAsync : key=" + key);
                //        log.Info("WordStylesAsync : val=" + val);
                //        if (key == "storedStyledReport")
                //        {
                //            storedS3Key = val;
                //        }
                //        if (key == "reportStyles")
                //        {
                //            reportStyles = val.ToString() == "1" ? true : false;
                //        }
                //    }
                //}

                //// TODO: encapsulate all of this s3 thing in a new service
                //var isStaticStyle = ConfigurationManager.AppSettings["IsStaticStyle"];
                var s3object = new GetS3Object();
                var storedS3ReportPath = s3object.ReadObjectData(args.ReportStyleS3Id);

                //if (isStaticStyle == "true")
                //{
                //    storedS3ReportPath = HttpContext.Current.Server.MapPath("~/StylesDoc/DocStyles.docx");
                //}
                //else
                //{
                //    // get document from s3
                //    storedS3ReportPath = s3object.ReadObjectData(storedS3Key);
                //}

                log.Info("WordStylesAsync : storedS3ReportPath:" + storedS3ReportPath);
                log.Info("WordStylesAsync : generatedReportPath:" + generatedReportPath);
                log.Info("WordStylesAsync : reportStyles:" + args.ReportStyleS3Id);

                // replace styles now that you have the file paths for each
                ReplaceDocStyles.ReplaceStyles(storedS3ReportPath, generatedReportPath);

                // replace theme content, too (handles color, fonts, etc.)
                ReplaceDocStyles.CopyThemeContent(storedS3ReportPath, generatedReportPath);

                if (!string.IsNullOrEmpty(args.ReportStyleS3Id))
                {
                    //Replace document margin
                    ReplaceDocStyles.ReplaceMargin(storedS3ReportPath, generatedReportPath);
                    
                    // TODO: figure this out
                    ReplaceDocStyles.ReplaceOtherStyles(storedS3ReportPath, generatedReportPath, null);
                }

                return generatedReportPath;

                // Read the file into a Byte Array
                //byte[] bytes = File.ReadAllBytes(generatedReportPath);
                //response.Content = new ByteArrayContent(bytes);
                //response.Content.Headers.ContentLength = bytes.LongLength;

                //// Set the Content Disposition Header Value and File name
                //response.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse("attachment; filename=Report.docx");
                //response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.wordprocessingml.document");

                //// remove temporary files
                //if (!(isStaticStyle == "true"))
                //{
                //    if (File.Exists(storedS3ReportPath))
                //    {
                //        File.Delete(storedS3ReportPath);
                //    }
                //}

                //if (File.Exists(generatedReportPath))
                //{
                //    File.Delete(generatedReportPath);
                //}

                //log.Info("WordStylesAsync : completed");
                ////return generated report with new styles
                //return response;

            }
            catch (Exception e)
            {
                log.Error("WordStylesAsync : failed, error:" + e.ToString());
                //return Request.CreateErrorResponse(HttpStatusCode.InternalServerError, e);
                return generatedReportPath;
            }

        }

        #region "General Creating Word Document Methods"

        public void AddElementIfMissing(XDocument partXDoc, XElement existing, string newElement)
        {
            if (existing != null)
            {
                return;
            }
            XElement newXElement = XElement.Parse(newElement);
            newXElement.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
            partXDoc.Root.Add(newXElement);
        }

        public void UpdateAStylePartForToc(XDocument partXDoc)
        {

            log.Info("UpdateAStylePartForToc - inside");
            AddElementIfMissing(
                partXDoc,
                partXDoc.Root.Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "TOCHeading")
                    .FirstOrDefault(),
                @"<w:style w:type='paragraph' w:styleId='TOCHeading' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                    <w:name w:val='TOC Heading'/>
                    <w:basedOn w:val='Heading1'/>
                    <w:next w:val='Normal'/>
                    <w:uiPriority w:val='39'/>
                    <w:semiHidden/>
                    <w:unhideWhenUsed/>
                    <w:qFormat/>
                    <w:pPr>
                      <w:outlineLvl w:val='9'/>
                    </w:pPr>
                    <w:rPr>
                      <w:lang w:eastAsia='ja-JP'/>
                    </w:rPr>
                  </w:style>");

            for (int i = 1; i <= 6; ++i)
            {
                AddElementIfMissing(
                    partXDoc,
                    partXDoc.Root.Elements(W.style)
                        .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == ("TOC" + i.ToString()))
                        .FirstOrDefault(),
                    String.Format(
                        @"<w:style w:type='paragraph' w:styleId='TOC{0}' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                            <w:name w:val='toc {0}'/>
                            <w:basedOn w:val='Normal'/>
                            <w:next w:val='Normal'/>
                            <w:autoRedefine/>
                            <w:uiPriority w:val='39'/>
                            <w:unhideWhenUsed/>
                            <w:pPr>
                              <w:spacing w:after='100'/>
                            </w:pPr>
                          </w:style>", i));
            }

            AddElementIfMissing(
                partXDoc,
                partXDoc.Root.Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" && (string)e.Attribute(W.styleId) == "Hyperlink")
                    .FirstOrDefault(),
                @"<w:style w:type='character' w:styleId='Hyperlink' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                     <w:name w:val='Hyperlink'/>
                     <w:basedOn w:val='DefaultParagraphFont'/>
                     <w:uiPriority w:val='99'/>
                     <w:unhideWhenUsed/>
                     <w:rPr>
                       <w:color w:val='0000FF' w:themeColor='hyperlink'/>
                       <w:u w:val='single'/>
                     </w:rPr>
                   </w:style>");

            AddElementIfMissing(
                partXDoc,
                partXDoc.Root.Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "paragraph" && (string)e.Attribute(W.styleId) == "BalloonText")
                    .FirstOrDefault(),
                @"<w:style w:type='paragraph' w:styleId='BalloonText' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                    <w:name w:val='Balloon Text'/>
                    <w:basedOn w:val='Normal'/>
                    <w:link w:val='BalloonTextChar'/>
                    <w:uiPriority w:val='99'/>
                    <w:semiHidden/>
                    <w:unhideWhenUsed/>
                    <w:pPr>
                      <w:spacing w:after='0' w:line='240' w:lineRule='auto'/>
                    </w:pPr>
                    <w:rPr>
                      <w:rFonts w:ascii='Tahoma' w:hAnsi='Tahoma' w:cs='Tahoma'/>
                      <w:sz w:val='16'/>
                      <w:szCs w:val='16'/>
                    </w:rPr>
                  </w:style>");

            AddElementIfMissing(
                partXDoc,
                partXDoc.Root.Elements(W.style)
                    .Where(e => (string)e.Attribute(W.type) == "character" &&
                        (bool?)e.Attribute(W.customStyle) == true && (string)e.Attribute(W.styleId) == "BalloonTextChar")
                    .FirstOrDefault(),
                @"<w:style w:type='character' w:customStyle='1' w:styleId='BalloonTextChar' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                    <w:name w:val='Balloon Text Char'/>
                    <w:basedOn w:val='DefaultParagraphFont'/>
                    <w:link w:val='BalloonText'/>
                    <w:uiPriority w:val='99'/>
                    <w:semiHidden/>
                    <w:rPr>
                      <w:rFonts w:ascii='Tahoma' w:hAnsi='Tahoma' w:cs='Tahoma'/>
                      <w:sz w:val='16'/>
                      <w:szCs w:val='16'/>
                    </w:rPr>
                  </w:style>");
        }

        public void UpdateFontTablePart(WordprocessingDocument doc)
        {
            FontTablePart fontTablePart = doc.MainDocumentPart.FontTablePart;
            if (fontTablePart != null)
            {

                XDocument fontTableXDoc = fontTablePart.GetXDocument();

                AddElementIfMissing(fontTableXDoc,
                    fontTableXDoc
                        .Root
                        .Elements(W.font)
                        .Where(e => (string)e.Attribute(W.name) == "Tahoma")
                        .FirstOrDefault(),
                    @"<w:font w:name='Tahoma' xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                     <w:panose1 w:val='020B0604030504040204'/>
                     <w:charset w:val='00'/>
                     <w:family w:val='swiss'/>
                     <w:pitch w:val='variable'/>
                     <w:sig w:usb0='E1002EFF' w:usb1='C000605B' w:usb2='00000029' w:usb3='00000000' w:csb0='000101FF' w:csb1='00000000'/>
                   </w:font>");

                fontTablePart.PutXDocument();
            }
        }

        public void UpdateStylesPartForToc(WordprocessingDocument doc)
        {
            StylesPart stylesPart = doc.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart != null)
            {
                log.Info("UpdateStylesPartForToc - inside");
                XDocument stylesXDoc = stylesPart.GetXDocument();
                UpdateAStylePartForToc(stylesXDoc);
                stylesPart.PutXDocument();
            }
        }

        public void UpdateStylesWithEffectsPartForToc(WordprocessingDocument doc)
        {
            StylesWithEffectsPart stylesWithEffectsPart = doc.MainDocumentPart.StylesWithEffectsPart;
            if (stylesWithEffectsPart != null)
            {
                // throw new OpenXmlPowerToolsException("Document does not contain styles with effects part");
                XDocument stylesWithEffectsXDoc = stylesWithEffectsPart.GetXDocument();
                UpdateAStylePartForToc(stylesWithEffectsXDoc);
                stylesWithEffectsPart.PutXDocument();
            }
        }

        public void CreateTOC(string fileName, string depth)
        {
            log.Info("CreateTOC : Started");
            try
            {
                using (WordprocessingDocument wdoc = WordprocessingDocument.Open(fileName, true))
                {

                    try
                    {
                        UpdateFontTablePart(wdoc);
                        UpdateStylesPartForToc(wdoc);
                        UpdateStylesWithEffectsPartForToc(wdoc);
                    }
                    catch (Exception exception)
                    {
                        log.Error($"CreateTOC : failed while applying toc heading styles, error-{exception.ToString()}");
                    }

                    depth = string.IsNullOrEmpty(depth) ? "6" : depth;

                    var title = "Table of Contents";
                    var switches = $"TOC \\o '1-{depth}' \\h \\z \\u";

                    var xmlString = @"<w:sdt xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>
                          <w:sdtPr>
                            <w:docPartObj>
                              <w:docPartGallery w:val='Table of Contents'/>
                              <w:docPartUnique/>
                            </w:docPartObj>
                          </w:sdtPr>
                          <w:sdtEndPr>
                            <w:rPr>
                             <w:rFonts w:asciiTheme='minorHAnsi' w:cstheme='minorBidi' w:eastAsiaTheme='minorHAnsi' w:hAnsiTheme='minorHAnsi'/>
                             <w:color w:val='auto'/>
                             <w:sz w:val='22'/>
                             <w:szCs w:val='22'/>
                             <w:lang w:eastAsia='en-US'/>
                            </w:rPr>
                          </w:sdtEndPr>
                          <w:sdtContent>
                            <w:p>
                              <w:pPr>
                                <w:pStyle w:val='TOCHeading'/>
                              </w:pPr>
                              <w:r>
                                <w:t>{0}</w:t>
                              </w:r>
                            </w:p>
                            <w:p>
                              <w:pPr>
                                 <w:pStyle w:val='TOC1'/>
                                 <w:tabs>
                                      <w:tab w:val='right' w:leader='dot' w:pos=''/>
                                 </w:tabs>
                                <w:rPr>
                                  <w:noProof/>
                                </w:rPr>
                              </w:pPr>
                              <w:r>
                                <w:fldChar w:fldCharType='begin' w:dirty='true'/>
                              </w:r>
                              <w:r>
                                <w:instrText xml:space='preserve'> {1} </w:instrText>
                              </w:r>
                              <w:r>
                                <w:fldChar w:fldCharType='separate'/>
                              </w:r>
                            </w:p>
                            <w:p>
                              <w:r>
                                <w:rPr>
                                  <w:b/>
                                  <w:bCs/>
                                  <w:noProof/>
                                </w:rPr>
                                <w:fldChar w:fldCharType='end'/>
                              </w:r>
                            </w:p>
                          </w:sdtContent>
                        </w:sdt>";

                    XElement sdt = XElement.Parse(String.Format(xmlString, title, switches));
                    XDocument mainXDoc = wdoc.MainDocumentPart.GetXDocument();

                    var hasToCLocation = wdoc.MainDocumentPart.GetXDocument().Descendants(W.p).Where(foo => foo.HasAttributes).FirstOrDefault();

                    if (hasToCLocation != null)
                    {
                        var hasAttribute = hasToCLocation.FirstAttribute.Name == "id" &&
                                            hasToCLocation.FirstAttribute.Value.ToLower().Equals("tableofcontents");

                        if (hasAttribute)
                        {
                            hasToCLocation.AddAfterSelf(sdt);
                        }

                    }
                    else
                    {
                        XElement firstParagraph = wdoc.MainDocumentPart.GetXDocument().Descendants(W.p).FirstOrDefault();
                        firstParagraph.AddBeforeSelf(sdt);
                    }

                    wdoc.MainDocumentPart.PutXDocument();
                    log.Info("CreateTOC : Completed");
                }
            }
            catch (Exception ex)
            {
                log.Error("CreateTOC : failed, error-" + ex.ToString());
            }
        }

        [Obsolete]
        public void SetFooterContent(WordprocessingDocument doc, string content, string position)
        {
            try
            {
                log.Info("SetFooterContent : started");
                log.Info("SetFooterContent : content:" + content);
                log.Info("SetFooterContent : position:" + position);

                // Get the main document part.
                MainDocumentPart mainDocPart = doc.MainDocumentPart;
                FooterPart footerPart = mainDocPart.AddNewPart<FooterPart>();
                var rId = mainDocPart.GetIdOfPart(footerPart);
                var footerRef = new FooterReference { Id = rId };
                var sectionProps = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();

                if (sectionProps == null)
                {
                    sectionProps = new SectionProperties();
                    doc.MainDocumentPart.Document.Body.Append(sectionProps);
                }

                sectionProps.RemoveAllChildren<FooterReference>();
                sectionProps.Append(footerRef);
                var footer = new Footer();

                if (!string.IsNullOrEmpty(content))
                {
                    Paragraph paragraph1 = new Paragraph() { };
                    Run run1 = new Run();
                    Text text1 = new Text();
                    text1.Text = content;
                    run1.Append(text1);
                    paragraph1.Append(run1);
                    footer.Append(paragraph1);
                    log.Info("SetFooterContent : footer added content =" + content);
                }

                log.Info("SetFooterContent : footer pagenumber position =" + position);
                //set page number in footer

                if (!string.IsNullOrEmpty(position))
                {

                    Paragraph paragraph2 = new Paragraph();
                    ParagraphProperties paragraphProperties = new ParagraphProperties();
                    ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "Footer" };
                    paragraphProperties.Append(paragraphStyleId);
                    JustificationValues pos = (JustificationValues)Enum.Parse(typeof(JustificationValues), position, true);
                    Justification justification = new Justification() { Val = pos };
                    paragraphProperties.Append(justification);
                    paragraph2.Append(paragraphProperties);

                    Run run1 = new Run();
                    Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
                    text1.Text = "";
                    run1.Append(text1);
                    paragraph2.Append(run1);

                    Run run2 = new Run();
                    FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                    run2.Append(fieldChar1);
                    paragraph2.Append(run2);

                    Run run3 = new Run();
                    FieldCode fieldCode1 = new FieldCode();
                    fieldCode1.Text = "PAGE";
                    run3.Append(fieldCode1);
                    paragraph2.Append(run3);

                    Run run4 = new Run();
                    FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };
                    run4.Append(fieldChar3);
                    paragraph2.Append(run4);
                    footer.Append(paragraph2);
                    log.Info("SetFooterContent : footer pagenumber added");
                }

                footerPart.Footer = footer;
                footerPart.Footer.Save();
                log.Info("SetFooterContent : Completed");

            }
            catch (Exception ex)
            {
                log.Error("SetFooterContent : failed, error " + ex.ToString());
            }
        }

        [Obsolete]
        public void SetHeaderContent(WordprocessingDocument doc, string position, string image, string content)
        {

            try
            {

                log.Info("SetHeaderContent : Started");
                log.Info("SetHeaderContent : image:" + image);
                log.Info("SetHeaderContent : content:" + content);
                log.Info("SetHeaderContent : position:" + position);
                var mainDocPart = doc.MainDocumentPart;
                long iWidth = 0;
                long iHeight = 0;

                if (!mainDocPart.HeaderParts.Any())
                {
                    mainDocPart.DeleteParts(mainDocPart.HeaderParts);
                    var newHeaderPart = mainDocPart.AddNewPart<HeaderPart>();

                    // try this instead
                    string imagePartID = string.Empty;
                    if (!string.IsNullOrEmpty(image))
                    {

                        var imgPart = newHeaderPart.AddImagePart(ImagePartType.Jpeg, "rId999");
                        imagePartID = newHeaderPart.GetIdOfPart(imgPart);
                        WebClient wc = new WebClient();
                        using (Stream fs = wc.OpenRead(image))
                        {
                            using (Image sourceImage = Image.FromStream(fs, true, true))
                            {
                                iWidth = sourceImage.Width;
                                iHeight = sourceImage.Height;
                            }
                        }
                        using (Stream fs = wc.OpenRead(image))
                        {
                            imgPart.FeedData(fs);
                        }
                        log.Info("SetHeaderContent : iWidth:" + iWidth);
                        log.Info("SetHeaderContent : iHeight:" + iHeight);
                    }
                    var rId = mainDocPart.GetIdOfPart(newHeaderPart);
                    var headerRef = new HeaderReference { Id = rId };
                    var sectionProps = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();

                    if (sectionProps == null)
                    {
                        sectionProps = new SectionProperties();
                        doc.MainDocumentPart.Document.Body.Append(sectionProps);
                    }

                    sectionProps.RemoveAllChildren<HeaderReference>();
                    sectionProps.Append(headerRef);
                    Header head = new Header();

                    if (!string.IsNullOrEmpty(imagePartID))
                    {
                        head = GeneratePicture(imagePartID, content, iWidth, iHeight);
                    }
                    else
                    {
                        head = new Header();
                        var paragraph = new Paragraph();
                        var run = new Run();
                        Text text = new Text();
                        text.Text = content;
                        run.Append(text);
                        paragraph.Append(run);
                        head.Append(paragraph);
                    }
                    //page number
                    if (!string.IsNullOrEmpty(position))
                    {
                        Paragraph paragraph_1 = new Paragraph();
                        ParagraphProperties paragraphProperties = new ParagraphProperties();
                        ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "Header" };
                        paragraphProperties.Append(paragraphStyleId);
                        JustificationValues pos = (JustificationValues)Enum.Parse(typeof(JustificationValues), position, true);
                        Justification justification = new Justification() { Val = pos };
                        paragraphProperties.Append(justification);
                        paragraph_1.Append(paragraphProperties);

                        Run run_2 = new Run();
                        FieldChar fieldChar_1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
                        run_2.Append(fieldChar_1);
                        paragraph_1.Append(run_2);

                        Run run_3 = new Run();
                        FieldCode fieldCode_1 = new FieldCode();
                        fieldCode_1.Text = "PAGE";
                        run_3.Append(fieldCode_1);
                        paragraph_1.Append(run_3);

                        Run run_4 = new Run();
                        FieldChar fieldChar_3 = new FieldChar() { FieldCharType = FieldCharValues.End };
                        run_4.Append(fieldChar_3);
                        paragraph_1.Append(run_4);
                        head.Append(paragraph_1);
                    }
                    newHeaderPart.Header = head;
                    newHeaderPart.Header.Save();
                }
                log.Info("SetHeaderContent : Completed");
            }
            catch (Exception ex)
            {
                log.Error("SetHeaderContent : failed, error=" + ex.ToString());
            }
        }

        [Obsolete]
        private static Header GeneratePicture(string relationshipId, string content, long iWidth, long iHeight)
        {
            iWidth = (long)Math.Round((decimal)iWidth * 4127);
            iHeight = (long)Math.Round((decimal)iHeight * 4127);
            log.Error("SetHeaderContent : iWidth=" + iWidth);
            log.Error("SetHeaderContent : iWidth=" + iHeight);
            var element =
                new Drawing(
                    new DW.Anchor(
                        new DW.SimplePosition() { X = 0L, Y = 0L },
                        new DW.HorizontalPosition(new DW.PositionOffset() { Text = "0" }) { RelativeFrom = DW.HorizontalRelativePositionValues.Column },
                        new DW.VerticalPosition(new DW.PositionOffset() { Text = "0" }) { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph },
                        new DW.Extent() { Cx = iWidth, Cy = iHeight },
                        new DW.EffectExtent()
                        {
                            LeftEdge = 0L,
                            TopEdge = 0L,
                            RightEdge = 0L,
                            BottomEdge = 0L
                        },
                        new DW.WrapThrough(
                        new DW.WrapPolygon(
                            new DW.StartPoint() { X = 0L, Y = 0L },
                            new DW.LineTo() { X = 0L, Y = 20935L },
                            new DW.LineTo() { X = 20935L, Y = 20935L },
                            new DW.LineTo() { X = 20935L, Y = 0L },
                            new DW.LineTo() { X = 0L, Y = 0L }
                        )
                        { Edited = false }
                        )
                        { WrapText = DW.WrapTextValues.BothSides },
                        new DW.DocProperties()
                        {
                            Id = (UInt32Value)1U,
                            Name = "NIS Logo"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(
                            new A.GraphicFrameLocks() { NoChangeAspect = true }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)0U,
                                            Name = "nis.png"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()),
                                    new PIC.BlipFill(
                                        new A.Blip(
                                            new A.BlipExtensionList(
                                                new A.BlipExtension(
                                                    new A14.UseLocalDpi() { Val = false }
                                                    )
                                                {
                                                    Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                })
                                        )
                                        {
                                            Embed = relationshipId,
                                            CompressionState =
                                                A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset() { X = 0L, Y = 0L },
                                            new A.Extents() { Cx = iWidth, Cy = iHeight }),
                                        new A.PresetGeometry(
                                            new A.AdjustValueList()
                                        )
                                        { Preset = A.ShapeTypeValues.Rectangle }))
                            )
                            { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }),
                        new Wp14.RelativeWidth(
                            new Wp14.PercentageWidth() { Text = "0" }
                            )
                        { ObjectId = Wp14.SizeRelativeHorizontallyValues.Margin },
                        new Wp14.RelativeHeight(
                            new Wp14.PercentageHeight() { Text = "0" }
                            )
                        { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Margin }
                    )
                    {
                        DistanceFromTop = (UInt32Value)0U,
                        DistanceFromBottom = (UInt32Value)0U,
                        DistanceFromLeft = (UInt32Value)114300U,
                        DistanceFromRight = (UInt32Value)114300U,
                        SimplePos = false,
                        RelativeHeight = (UInt32Value)251658240U,
                        BehindDoc = false,
                        Locked = false,
                        LayoutInCell = true,
                        AllowOverlap = true,
                        EditId = "5186AF0D",
                        AnchorId = "62A11D4C"
                    });

            var header = new Header();
            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "008757A6", RsidRunAdditionDefault = "00A63940", ParagraphId = "500CD76E", TextId = "77777777" };

            Run run2 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            run2.Append(runProperties1);
            run2.Append(element);

            Run run3 = new Run();
            Text text2 = new Text();
            text2.Text = content;

            run3.Append(text2);
            paragraph3.Append(run2);
            paragraph3.Append(run3);

            header.Append(paragraph3);

            return header;
        }

        public void SetHeaderContent(WordprocessingDocument doc, HeaderContent headerContent)
        {

            log.Info("SetHeaderContent : Started");

            if (headerContent == null)
            {
                log.Info("SetHeaderContent : headerContent was null, no header added.");
                return;
            }

            try
            {
                                
                log.Info($"SetHeaderContent : {headerContent.ToString()}");

                HeaderTypeDecider decider;

                HeaderTypes headerType;
                if (!Enum.TryParse(headerContent.Position, out headerType))
                {
                    log.Error($"SetHeaderContent : Could not parse {headerContent.Position} into the enum, using none.");
                    headerType = HeaderTypes.none;
                }

                switch (headerType)
                {
                    case HeaderTypes.image_and_pagenumber:
                        decider = new HeaderImageAndPageNumber();
                        break;
                    case HeaderTypes.image_and_text:
                        decider = new HeaderImageAndText();
                        break;
                    case HeaderTypes.image_center:
                        decider = new HeaderImageCenter();
                        break;
                    case HeaderTypes.text_and_image:
                        decider = new HeaderTextAndImage();
                        break;
                    case HeaderTypes.text_and_pagenumber:
                        decider = new HeaderTextAndPageNumber();
                        break;
                    case HeaderTypes.text_center:
                        decider = new HeaderTextCenter();
                        break;
                    case HeaderTypes.none:
                        decider = new HeaderNone();
                        break;
                    default:
                        decider = new HeaderNone();
                        break;
                }

                decider.TemplateMethod(doc, headerContent);

                log.Info("SetHeaderContent : Completed");

            }
            catch (Exception ex)
            {
                log.Error("SetHeaderContent : failed, error=" + ex.ToString());
            }
        }

        public void SetFooterContent(WordprocessingDocument doc, FooterContent footerContent)
        {

            try
            {

                log.Info($"SetFooterContent : {footerContent.ToString()}");

                if(footerContent == null)
                {
                    return;
                }

                FooterTypeDecider decider;

                FooterTypes footerType;
                if (!Enum.TryParse(footerContent.Position, out footerType))
                {
                    log.Error($"SetFooterContent : Could not parse {footerContent.Position} into the enum, using none.");
                    footerType = FooterTypes.none;
                }

                switch (footerType)
                {
                    case FooterTypes.pagenumber_center:
                        decider = new FooterPageNumberCenter();
                        break;
                    case FooterTypes.pagenumber_left:
                        decider = new FooterPageNumberLeft();
                        break;
                    case FooterTypes.pagenumber_right:
                        decider = new FooterPageNumberRight();
                        break;
                    case FooterTypes.image_and_pagenumber:
                        decider = new FooterImageAndPageNumber();
                        break;
                    case FooterTypes.text_and_pagenumber:
                        decider = new FooterTextAndPageNumber();
                        break;
                    case FooterTypes.none:
                        decider = new FooterNone();
                        break;
                    default:
                        decider = new FooterNone();
                        break;
                }

                decider.TemplateMethod(doc, footerContent);

                log.Info("SetFooterContent : Completed");

            }
            catch (Exception ex)
            {
                log.Error("SetFooterContent : failed, error " + ex.ToString());
            }

        }

        #endregion
    }
}