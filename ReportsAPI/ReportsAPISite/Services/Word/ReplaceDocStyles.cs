using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using MSInterop = Microsoft.Office.Interop.Word;

namespace ReportsAPISite.Services.Word
{
    // TODO: move this into its own c# project library
    public class ReplaceDocStyles
    {

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        // To copy contents of one package part.
        public static void CopyThemeContent(string fromDocument, string toDocument)
        {
            log.Info("CopyThemeContent : started");
            using (WordprocessingDocument wordDoc1 = WordprocessingDocument.Open(fromDocument, false))
            {
                using (WordprocessingDocument wordDoc2 = WordprocessingDocument.Open(toDocument, true))
                {
                    ThemePart themePart1 = wordDoc1.MainDocumentPart.ThemePart;
                    ThemePart themePart2 = wordDoc2.MainDocumentPart.ThemePart;
                    if (themePart2 == null)
                    {
                        themePart2 = wordDoc2.MainDocumentPart.AddNewPart<ThemePart>();
                    }
                    using (StreamReader streamReader = new StreamReader(themePart1.GetStream()))
                    {
                        using (StreamWriter streamWriter = new StreamWriter(themePart2.GetStream(FileMode.Create)))
                        {
                            streamWriter.Write(streamReader.ReadToEnd());
                        }
                    }
                }
            }
            log.Info("CopyThemeContent : completed");
        }

        // Replace the styles in the "to" document with the styles in
        // the "from" document.
        public static void ReplaceStyles(string fromDoc, string toDoc)
        {
            log.Info("ReplaceStyles : started");
            // Extract and replace the styles part.
            var node = ExtractStylesPart(fromDoc, false);
            if (node != null)
            {
                ReplaceStylesPart(toDoc, node, false);
            }

            // Extract and replace the stylesWithEffects part. To fully support 
            // round-tripping from Word 2010 to Word 2007, you should 
            // replace this part, as well.
            node = ExtractStylesPart(fromDoc);
            if (node != null)
            {
                ReplaceStylesPart(toDoc, node);
            }

            log.Info("ReplaceStyles : completed");
            return;
        }

        // Given a file and an XDocument instance that contains the content of 
        // a styles or stylesWithEffects part, replace the styles in the file 
        // with the styles in the XDocument.
        public static void ReplaceStylesPart(string fileName, XDocument newStyles, bool setStylesWithEffectsPart = true)
        {
            // Open the document for write access and get a reference.
            using (var document = WordprocessingDocument.Open(fileName, true))
            {
                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the
                // stylesPart variable.
                StylesPart stylesPart = null;
                if (setStylesWithEffectsPart)
                {
                    stylesPart = docPart.StylesWithEffectsPart;
                }
                else
                {
                    stylesPart = docPart.StyleDefinitionsPart;
                }

                // If the part exists, populate it with the new styles.
                if (stylesPart != null)
                {
                    newStyles.Save(new StreamWriter(stylesPart.GetStream(FileMode.Create, FileAccess.Write)));
                }
            }
        }

        // Extract the styles or stylesWithEffects part from a 
        // word processing document as an XDocument instance.
        public static XDocument ExtractStylesPart(string fileName, bool getStylesWithEffectsPart = true)
        {
            // Declare a variable to hold the XDocument.
            XDocument styles = null;

            // Open the document for read access and get a reference.
            using (var document = WordprocessingDocument.Open(fileName, false))
            {
                // Get a reference to the main document part.
                var docPart = document.MainDocumentPart;

                // Assign a reference to the appropriate part to the
                // stylesPart variable.
                StylesPart stylesPart = null;
                if (getStylesWithEffectsPart)
                {
                    stylesPart = docPart.StylesWithEffectsPart;
                }
                else
                {
                    stylesPart = docPart.StyleDefinitionsPart;
                }

                // If the part exists, read it into the XDocument.
                if (stylesPart != null)
                {
                    using (var reader = XmlNodeReader.Create(stylesPart.GetStream(FileMode.Open, FileAccess.Read)))
                    {
                        // Create the XDocument.
                        styles = XDocument.Load(reader);
                    }
                }
            }
            // Return the XDocument instance.
            return styles;
        }

        //replace the margin 
        public static void ReplaceMargin(string fromDoc, string toDoc)
        {
            log.Info("ReplaceMargin : started");
            var top = 0;
            var bottom = 0;
            UInt32Value left = 0;
            UInt32Value right = 0;
            UInt32Value header = 0;
            UInt32Value footer = 0;
            UInt32Value gutter = 0;
            try
            {
                using (WordprocessingDocument wdoc = WordprocessingDocument.Open(fromDoc, true))
                {
                    MainDocumentPart docPart = wdoc.MainDocumentPart;
                    var sections = docPart.Document.Descendants<SectionProperties>();

                    foreach (SectionProperties sectPr in sections)
                    {
                        PageMargin pgMar = sectPr.Descendants<PageMargin>().FirstOrDefault();
                        if (pgMar != null)
                        {
                            top = pgMar.Top.Value;
                            bottom = pgMar.Bottom.Value;
                            left = pgMar.Left.Value;
                            right = pgMar.Right.Value;
                            header = pgMar.Header.Value;
                            footer = pgMar.Footer.Value;
                            gutter = pgMar.Gutter.Value;
                        }
                    }
                }
                using (WordprocessingDocument wdoc = WordprocessingDocument.Open(toDoc, true))
                {
                    MainDocumentPart mainPart = wdoc.MainDocumentPart;
                    SectionProperties sectionProps = new SectionProperties();
                    PageMargin pageMargin = new PageMargin() { Top = top, Right = right, Bottom = bottom, Left = left, Header = header, Footer = footer, Gutter = gutter };
                    sectionProps.Append(pageMargin);
                    mainPart.Document.Body.Append(sectionProps);
                }
                log.Info("ReplaceMargin : started");
            }
            catch (Exception ex)
            {
                log.Info("ReplaceMargin : failed, error:" + ex.ToString());
            }

        }

        private static bool IsEmpty(string text)
        {
            return text.Replace("\r", "").Replace("\a", "").Replace("\t", "").Trim() == "";
        }

        public static void ReplaceOtherStyles(string fromDocument, string toDocument)
        {

            log.Info("ReplaceOtherStyles : Started");
            MSInterop.Document styleDocucument = null;
            MSInterop.Document destinationDocument = null;
            // TODO: we need to try and move away from Interop whenever we can

            MSInterop.Application word = new MSInterop.Application();
            try
            {

                log.Info("ReplaceOtherStyles : reading source doc=" + fromDocument);
                styleDocucument = word.Documents.Open(fromDocument, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                log.Info("ReplaceOtherStyles : reading dest doc=" + toDocument);
                destinationDocument = word.Documents.Open(toDocument, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                log.Info("ReplaceOtherStyles : replacing styles from source to destination");
                var isAnyDocumentNull = styleDocucument == null || destinationDocument == null;
                if (isAnyDocumentNull)
                {
                    throw new Exception("either source or destination document were null");
                }

                MSInterop.Style style = null;
                //this code is to change Header Style
                foreach (MSInterop.Section sec in styleDocucument.Sections)
                {
                    foreach (MSInterop.Section sec2 in destinationDocument.Sections)
                    {
                        if (sec.Headers.Count > 0 && sec.Headers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary] != null && sec.Headers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range != null)
                        {
                            if (sec.Headers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font != null)
                            {
                                if (!IsEmpty(sec.Headers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Text))
                                {
                                    sec2.Headers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font = sec.Headers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Font;
                                }
                            }
                            style = sec.Headers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.get_Style() as MSInterop.Style;
                            if (style != null)
                            {
                                sec2.Headers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.set_Style(style);
                            }
                        }
                    }
                }
                log.Info("ReplaceOtherStyles : applied header style");

                //this code is to change Footer Style
                // TODO: applying styles is overriding the footer options
                var doesSourceHaveAFooter = styleDocucument.Sections.Count > 0 && styleDocucument.Sections[1].Footers.Count > 0;
                var isTheFooterNotNull = styleDocucument.Sections[1].Footers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary] != null
                                                && styleDocucument.Sections[1].Footers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range != null;

                if (doesSourceHaveAFooter && isTheFooterNotNull)
                {

                    var destinationPrimaryFooter = destinationDocument.Sections[1].Footers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    var sourcePrimaryFooter = styleDocucument.Sections[1].Footers[MSInterop.WdHeaderFooterIndex.wdHeaderFooterPrimary];

                    style = sourcePrimaryFooter.Range.get_Style();
                    if (style != null)
                    {
                        destinationPrimaryFooter.Range.set_Style(style);
                    }

                    var doesFooterHaveAFontAndText = sourcePrimaryFooter.Range.Font != null && !IsEmpty(sourcePrimaryFooter.Range.Text);
                    if (doesFooterHaveAFontAndText)
                    {
                        destinationPrimaryFooter.Range.Font.Color = sourcePrimaryFooter.Range.Font.Color;
                        //destinationPrimaryFooter.Range.Text = sourcePrimaryFooter.Range.Text;

                        var doesSourceHavePrimaryFontAlignment = sourcePrimaryFooter.Range.Paragraphs != null &&
                                                                        sourcePrimaryFooter.Range.Paragraphs.First != null;

                        if (doesSourceHavePrimaryFontAlignment)
                        {
                            destinationPrimaryFooter.Range.Paragraphs.Alignment = sourcePrimaryFooter.Range.Paragraphs.First.Alignment;
                        }
                    }
                }
                log.Info("ReplaceOtherStyles : applied footer style");

                foreach (MSInterop.Table sourceTable in styleDocucument.Tables)
                {
                    foreach (MSInterop.Table destinationTable in destinationDocument.Tables)
                    {

                        var title = destinationTable.Title;
                        //For preventing the styles to be applied for Two column layout functionality

                        if (title == "XCelHeader")
                        {
                            //table2.Range.Font.Reset();
                            destinationTable.ApplyStyleFirstColumn = false;
                            destinationTable.ApplyStyleHeadingRows = false;
                            destinationTable.ApplyStyleLastColumn = false;
                            destinationTable.ApplyStyleLastRow = false;
                            destinationTable.ApplyStyleColumnBands = false;
                        }
                        else if (title == "TwoColumn")
                        {
                            //table2.Range.Font.Reset();
                            destinationTable.ApplyStyleFirstColumn = false;
                            destinationTable.ApplyStyleHeadingRows = false;
                            destinationTable.ApplyStyleLastColumn = false;
                            destinationTable.ApplyStyleLastRow = false;
                            destinationTable.Range.Paragraphs.Alignment = MSInterop.WdParagraphAlignment.wdAlignParagraphLeft;
                        }
                        else
                        {
                            style = sourceTable.get_Style();
                            if (style != null)
                            {
                                destinationTable.set_Style(style);
                            }

                            if (sourceTable.Range != null && sourceTable.Range.Font != null)
                            {
                                // TODO: this is not replacing the font at all
                                //table2.Range.Font.res = table.Range.Font;

                                // Reset() did the trick
                                destinationTable.Range.Font.Reset();
                            }
                        }

                    }
                }

                log.Info("ReplaceOtherStyles : applied table style");
                if (styleDocucument.Lists != null && styleDocucument.Lists.Count > 0)
                {
                    foreach (MSInterop.List list1 in destinationDocument.Lists)
                    {
                        if (styleDocucument.Lists.Count > 0 && styleDocucument.Lists[1].Range != null)
                        {
                            if (styleDocucument.Lists[1].Range.ParagraphFormat != null)
                            {
                                try
                                {
                                    list1.Range.ParagraphFormat.SpaceBefore = styleDocucument.Lists[1].Range.ParagraphFormat.SpaceBefore;
                                    list1.Range.ParagraphFormat.SpaceAfter = styleDocucument.Lists[1].Range.ParagraphFormat.SpaceAfter;
                                    list1.Range.ParagraphFormat.SpaceBeforeAuto = styleDocucument.Lists[1].Range.ParagraphFormat.SpaceBeforeAuto;
                                    list1.Range.ParagraphFormat.SpaceAfterAuto = styleDocucument.Lists[1].Range.ParagraphFormat.SpaceAfterAuto;
                                    list1.Range.ParagraphFormat.LineSpacing = styleDocucument.Lists[1].Range.ParagraphFormat.LineSpacing;
                                    list1.Range.ParagraphFormat.LineSpacingRule = styleDocucument.Lists[1].Range.ParagraphFormat.LineSpacingRule;
                                    list1.Range.ParagraphFormat.KeepTogether = styleDocucument.Lists[1].Range.ParagraphFormat.KeepTogether;
                                    ((MSInterop.Style)list1.Range.ParagraphFormat.get_Style()).NoSpaceBetweenParagraphsOfSameStyle = ((MSInterop.Style)styleDocucument.Lists[1].Range.ParagraphFormat.get_Style()).NoSpaceBetweenParagraphsOfSameStyle;
                                    log.Info("ReplaceOtherStyles : applied ParagraphFormat style");
                                }
                                catch (Exception ex) { log.Info("ReplaceOtherStyles : failed while applying ParagraphFormat-" + ex.ToString()); }
                            }
                            if (styleDocucument.Lists[1].Range.Font != null)
                            {
                                if (!IsEmpty(styleDocucument.Lists[1].Range.Text))
                                {
                                    try
                                    {
                                        list1.Range.Font = styleDocucument.Lists[1].Range.Font;
                                        log.Info("ReplaceOtherStyles : applied list Font style");
                                    }
                                    catch (Exception ex) { log.Info("ReplaceOtherStyles : failed while applying font-" + ex.ToString()); }
                                }
                            }
                            if (styleDocucument.Lists[1].Range.ListFormat != null && styleDocucument.Lists[1].Range.ListFormat.ListTemplate != null)
                            {
                                try
                                {
                                    list1.Range.ListFormat.ApplyListTemplate(styleDocucument.Lists[1].Range.ListFormat.ListTemplate);
                                    log.Info("ReplaceOtherStyles : applied ListTemplate style");
                                }
                                catch (Exception ex) { log.Info("ReplaceOtherStyles : failed while applying ListFormat-" + ex.ToString()); }
                            }

                        }
                    }
                }

                if (styleDocucument.TablesOfContents.Count > 0)
                {
                    try
                    {
                        destinationDocument.TablesOfContents.Format = styleDocucument.TablesOfContents.Format;
                        destinationDocument.TablesOfContents[1].Update();
                        log.Info("ReplaceOtherStyles : applied TablesOfContents style");
                    }
                    catch (Exception ex) { log.Info("ReplaceOtherStyles : failed while applying TablesOfContents -" + ex.ToString()); }
                }
                log.Info("ReplaceOtherStyles : applied list style");

                log.Info("ReplaceOtherStyles : completed");

            }
            catch (Exception ex)
            {
                log.Error("ReplaceOtherStyles : failed, error=" + ex.ToString());
            }
            finally
            {
                if (styleDocucument != null)
                {
                    styleDocucument.Save();
                    styleDocucument.Close();
                }
                if (destinationDocument != null)
                {
                    destinationDocument.Save();
                    destinationDocument.Close();
                }
                word.Quit();
            }
        }

    }
}