using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using reports.Models;

namespace reports.Services.Word.FooterFunctions
{
    public class FooterPageNumberLeft : FooterTypeDecider
    {
        protected override void Run(WordprocessingDocument doc, FooterContent footerContent, FooterPart footerPart, Footer footer)
        {
            AddPageNumber(footerContent, footer, JustificationValues.Left);
        }
    }
}