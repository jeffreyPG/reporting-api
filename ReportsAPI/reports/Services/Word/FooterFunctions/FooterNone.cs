using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using reports.Models;

namespace reports.Services.Word.FooterFunctions
{
    public class FooterNone : FooterTypeDecider
    {
        protected override void Run(WordprocessingDocument doc, FooterContent footerContent, FooterPart footerPart, Footer footer)
        {
            footerContent.Content = string.Empty;
            AddText(footerContent, footer);
        }
    }
}