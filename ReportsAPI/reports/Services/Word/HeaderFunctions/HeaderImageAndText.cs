using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using reports.Models;

namespace reports.Services.Word.HeaderFunctions
{

    public class HeaderImageAndText : HeaderTypeDecider
    {
        protected override void Run(WordprocessingDocument doc, HeaderContent headerContent, HeaderPart newHeaderPart, Header header)
        {
            AddImageAndText(headerContent, header, doc, newHeaderPart);
        }
    }
    
}