using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using reports.Models;

namespace reports.Services.Word.HeaderFunctions
{

    public class HeaderNone : HeaderTypeDecider
    {
        protected override void Run(WordprocessingDocument doc, HeaderContent headerContent, HeaderPart newHeaderPart, Header header)
        {
            headerContent.Content = string.Empty;
            AddText(headerContent, header, AbsolutePositionTabAlignmentValues.Left);
        }
    }
    
}