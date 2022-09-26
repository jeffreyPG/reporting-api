using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using reports.Models;

namespace reports.Services.Word.FooterFunctions
{
    public class FooterImageAndPageNumber : FooterTypeDecider
    {
        protected override void Run(WordprocessingDocument doc, FooterContent footerContent, FooterPart footerPart, Footer footer)
        {
            AddImageAndPageNumber(doc, footerContent, footerPart, footer);
        }
    }
}