using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

using HeaderContent = reports.Models.HeaderContent;
using System.Net;
using System.IO;
using System.Drawing;
using log4net;
using System.Linq;
using HtmlToOpenXml.Extensions;
using reports.Services.Word.Models;

namespace reports.Services.Word.HeaderFunctions
{

    public enum HeaderTypes
    {
        none = 0,
        image_and_text,
        image_and_pagenumber,
        text_and_image,
        text_and_pagenumber,
        text_center,
        image_center
    }

    public abstract class HeaderTypeDecider
    {

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        protected abstract void Run(WordprocessingDocument doc, HeaderContent headerContent, HeaderPart newHeaderPart, Header header);

        public void TemplateMethod(WordprocessingDocument doc, HeaderContent headerContent)
        {

            var mainDocumentPart = doc.MainDocumentPart;

            if (!mainDocumentPart.HeaderParts.Any())
            {

                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
                var newHeaderPart = mainDocumentPart.AddNewPart<HeaderPart>();

                var referenceId = mainDocumentPart.GetIdOfPart(newHeaderPart);
                var headerRef = new HeaderReference { Id = referenceId };
                var sectionProps = doc.MainDocumentPart.Document.Body.Elements<SectionProperties>().LastOrDefault();

                if (sectionProps == null)
                {
                    sectionProps = new SectionProperties();
                    doc.MainDocumentPart.Document.Body.Append(sectionProps);
                }

                sectionProps.RemoveAllChildren<HeaderReference>();
                sectionProps.Append(headerRef);

                var header = new Header();

                // template method
                Run(doc, headerContent, newHeaderPart, header);
                
                AddDivider(headerContent, header);

                newHeaderPart.Header = header;
                newHeaderPart.Header.Save();

            }

        }

        public void GeneratePicture(HeaderContent headerContent, HeaderPart newHeaderPart, WordprocessingDocument doc, Header header, JustificationValues justification)
        {
            
            long iWidth = 0, iHeight = 0;
            var mainDocumentPart = doc.MainDocumentPart;

            var imagePartId = string.Empty;
            if (!string.IsNullOrEmpty(headerContent.Image))
            {

                var imgPart = newHeaderPart.AddImagePart(ImagePartType.Jpeg, "rId999");
                imagePartId = newHeaderPart.GetIdOfPart(imgPart);
                WebClient wc = new WebClient();
                using (Stream fs = wc.OpenRead(headerContent.Image))
                {
                    using (Image sourceImage = Image.FromStream(fs, true, true))
                    {
                        iWidth = sourceImage.Width;
                        iHeight = sourceImage.Height;
                    }
                }

                using (Stream fs = wc.OpenRead(headerContent.Image))
                {
                    imgPart.FeedData(fs);
                }

                log.Info("SetHeaderContent : iWidth:" + iWidth);
                log.Info("SetHeaderContent : iHeight:" + iHeight);

            }

            iWidth = (long)Math.Round((decimal)iWidth * 9525);
            iHeight = (long)Math.Round((decimal)iHeight * 9525);

            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.SimplePosition() { X = 0L, Y = 0L },
                        new DW.HorizontalPosition(new DW.PositionOffset() { Text = "0" })
                            { RelativeFrom = DW.HorizontalRelativePositionValues.Margin },
                        new DW.VerticalPosition(new DW.PositionOffset() { Text = "0" })
                            { RelativeFrom = DW.VerticalRelativePositionValues.Paragraph },
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
                                new Pic.Picture(
                                    new Pic.NonVisualPictureProperties(
                                        new Pic.NonVisualDrawingProperties()
                                        {
                                            Id = (UInt32Value)0U,
                                            Name = "nis.png"
                                        },
                                        new Pic.NonVisualPictureDrawingProperties()),
                                    new Pic.BlipFill(
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
                                            Embed = imagePartId,
                                            CompressionState =
                                                A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(
                                            new A.FillRectangle())),
                                    new Pic.ShapeProperties(
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
                        DistanceFromLeft = (UInt32Value)0U,
                        DistanceFromRight = (UInt32Value)0U,
                        EditId = "5186AF0D",
                        AnchorId = "62A11D4C"
                    });


            var paragraph = new Paragraph() { RsidParagraphAddition = "008757A6", RsidRunAdditionDefault = "00A63940", ParagraphId = "500CD76E", TextId = "77777777" };
            
            var paragraphProperties = new ParagraphProperties();
            var justification1 = new Justification() { Val = justification };
            paragraphProperties.Append(justification1);

            var run = new Run();

            var runProperties1 = new RunProperties();
            var noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            run.Append(runProperties1);
            run.Append(element);

            paragraph.Append(paragraphProperties);
            paragraph.Append(run);
            header.Append(paragraph);

        }

        public void AddText(HeaderContent headerContent, Header header, AbsolutePositionTabAlignmentValues alignment)
        {

            var positionalTab = new PositionalTab() { Alignment = alignment, RelativeTo = AbsolutePositionTabPositioningBaseValues.Margin, Leader = AbsolutePositionTabLeaderCharValues.None };

            var paragraph = new Paragraph();
            var run = new Run();
            run.Append(positionalTab);

            var text = new Text();
            text.Text = headerContent.Content;

            run.Append(text);
            paragraph.Append(run);

            header.Append(paragraph);

        }

        public void AddDivider(HeaderContent headerContent, Header header)
        {

            if (string.IsNullOrEmpty(headerContent.Divider))
            {
                return;
            }

            //AddEmptyParagraphToHeader(header);

            var fillColor = headerContent.Divider.FillColorOrDefault();
            var thickness = headerContent.Divider.ThicknessOrDefault();

            var paragraph1 = new Paragraph() { RsidParagraphAddition = "001D70B7", RsidParagraphProperties = "001D70B7", RsidRunAdditionDefault = "00457449", ParagraphId = "6F622EFE", TextId = "673E6002" };

            var paragraphProperties1 = new ParagraphProperties();
            var paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            var run = new Run();

            Picture picture1 = new Picture() { AnchorId = "11C91873" };
            V.Rectangle rectangle1 = new V.Rectangle() { Id = "_x0000_i1025", Style = $"height:{thickness};mso-position-horizontal:absolute", Horizontal = true, HorizontalStandard = true, HorizontalNoShade = true, HorizontalAlignment = Ovml.HorizontalRuleAlignmentValues.Center, FillColor = fillColor, Stroked = false };

            picture1.Append(rectangle1);

            run.Append(picture1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run);

            header.Append(paragraph1);

        }

        private static void AddEmptyParagraphToHeader(Header header)
        {
            var emptyRun = new Run();
            var emptyParagraph = new Paragraph();
            emptyParagraph.Append(emptyRun);
            header.Append(emptyParagraph);
        }

        public void AddPageNumberToHeader(HeaderContent headerContent, Header header, string position)
        {

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
                header.Append(paragraph_1);
            }

        }

        [Obsolete]
        public void AddTextAndPicture(HeaderContent headerContent, Header header, WordprocessingDocument doc, HeaderPart headerPart, JustificationValues firstJustification, JustificationValues secondJustification)
        {


            #region table

            var table = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = -5, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "5747" };
            GridColumn gridColumn2 = new GridColumn() { Width = "3096" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00FA14B1", RsidTableRowAddition = "008B6C0B", RsidTableRowProperties = "008B6C0B", ParagraphId = "5FEC0FF4", TextId = "77777777" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)990U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "6667", Type = TableWidthUnitValues.Dxa };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(shading1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00F1402E", RsidParagraphProperties = "00F1402E", RsidRunAdditionDefault = "008B6C0B", ParagraphId = "347747B3", TextId = "77777777" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            Tabs tabs1 = new Tabs();
            TabStop tabStop1 = new TabStop() { Val = TabStopValues.Right, Position = 9360 };

            tabs1.Append(tabStop1);
            Justification justification1 = new Justification() { Val = firstJustification };

            ParagraphMarkRunProperties paragraphMarkRunProperties1 = new ParagraphMarkRunProperties();
            RunFonts runFonts1 = new RunFonts() { Ascii = "Univers LT Std 57 Condensed", HighAnsi = "Univers LT Std 57 Condensed" };
            FontSize fontSize1 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript1 = new FontSizeComplexScript() { Val = "32" };

            paragraphMarkRunProperties1.Append(runFonts1);
            paragraphMarkRunProperties1.Append(fontSize1);
            paragraphMarkRunProperties1.Append(fontSizeComplexScript1);

            paragraphProperties1.Append(tabs1);
            paragraphProperties1.Append(justification1);
            paragraphProperties1.Append(paragraphMarkRunProperties1);

            Run run1 = new Run() { RsidRunProperties = "005B0830" };

            RunProperties runProperties1 = new RunProperties();
            RunFonts runFonts2 = new RunFonts() { Ascii = "Univers LT Std 57 Condensed", HighAnsi = "Univers LT Std 57 Condensed" };
            FontSize fontSize2 = new FontSize() { Val = "32" };
            FontSizeComplexScript fontSizeComplexScript2 = new FontSizeComplexScript() { Val = "32" };

            runProperties1.Append(runFonts2);
            runProperties1.Append(fontSize2);
            runProperties1.Append(fontSizeComplexScript2);
            Text text1 = new Text();
            text1.Text = headerContent.Content;

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "006F4615", RsidParagraphAddition = "009A0695", RsidParagraphProperties = "00F1402E", RsidRunAdditionDefault = "00660FA3", ParagraphId = "1A24ED6D", TextId = "34080537" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();

            Tabs tabs2 = new Tabs();
            TabStop tabStop2 = new TabStop() { Val = TabStopValues.Right, Position = 9360 };

            tabs2.Append(tabStop2);
            Justification justification2 = new Justification() { Val = firstJustification };

            ParagraphMarkRunProperties paragraphMarkRunProperties2 = new ParagraphMarkRunProperties();
            RunFonts runFonts3 = new RunFonts() { Ascii = "Univers LT Std 57 Condensed", HighAnsi = "Univers LT Std 57 Condensed" };
            Italic italic1 = new Italic();
            ItalicComplexScript italicComplexScript1 = new ItalicComplexScript();
            FontSizeComplexScript fontSizeComplexScript3 = new FontSizeComplexScript() { Val = "16" };

            paragraphMarkRunProperties2.Append(runFonts3);
            paragraphMarkRunProperties2.Append(italic1);
            paragraphMarkRunProperties2.Append(italicComplexScript1);
            paragraphMarkRunProperties2.Append(fontSizeComplexScript3);

            paragraphProperties2.Append(tabs2);
            paragraphProperties2.Append(justification2);
            paragraphProperties2.Append(paragraphMarkRunProperties2);

            Run run2 = new Run() { RsidRunProperties = "006F4615" };

            RunProperties runProperties2 = new RunProperties();
            RunFonts runFonts4 = new RunFonts() { Ascii = "Univers LT Std 57 Condensed", HighAnsi = "Univers LT Std 57 Condensed" };
            Italic italic2 = new Italic();
            ItalicComplexScript italicComplexScript2 = new ItalicComplexScript();
            FontSizeComplexScript fontSizeComplexScript4 = new FontSizeComplexScript() { Val = "16" };

            runProperties2.Append(runFonts4);
            runProperties2.Append(italic2);
            runProperties2.Append(italicComplexScript2);
            runProperties2.Append(fontSizeComplexScript4);
            Text text2 = new Text();
            // TODO: second line?
            text2.Text = "";

            run2.Append(runProperties2);
            run2.Append(text2);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);
            tableCell1.Append(paragraph2);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2683", Type = TableWidthUnitValues.Dxa };
            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00FA14B1", RsidParagraphAddition = "008B6C0B", RsidParagraphProperties = "008B6C0B", RsidRunAdditionDefault = "008B6C0B", ParagraphId = "15B51B7C", TextId = "77777777" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();

            Tabs tabs3 = new Tabs();
            TabStop tabStop3 = new TabStop() { Val = TabStopValues.Right, Position = 9360 };

            tabs3.Append(tabStop3);
            Justification justification3 = new Justification() { Val = secondJustification };

            ParagraphMarkRunProperties paragraphMarkRunProperties3 = new ParagraphMarkRunProperties();
            RunFonts runFonts5 = new RunFonts() { Ascii = "Univers LT Std 55 Roman", HighAnsi = "Univers LT Std 55 Roman" };
            FontSize fontSize3 = new FontSize() { Val = "40" };
            FontSizeComplexScript fontSizeComplexScript5 = new FontSizeComplexScript() { Val = "40" };

            paragraphMarkRunProperties3.Append(runFonts5);
            paragraphMarkRunProperties3.Append(fontSize3);
            paragraphMarkRunProperties3.Append(fontSizeComplexScript5);

            paragraphProperties3.Append(tabs3);
            paragraphProperties3.Append(justification3);
            paragraphProperties3.Append(paragraphMarkRunProperties3);

            Run run3 = new Run() { RsidRunProperties = "00FA14B1" };

            RunProperties runProperties3 = new RunProperties();
            RunFonts runFonts6 = new RunFonts() { Ascii = "Univers LT Std 55 Roman", HighAnsi = "Univers LT Std 55 Roman" };
            NoProof noProof1 = new NoProof();

            runProperties3.Append(runFonts6);
            runProperties3.Append(noProof1);

            #endregion
                        
            var dimensions = AddImageToHeaderPart(headerContent, doc, headerPart);

            Drawing drawing1 = new Drawing();

            DW.Inline inline1 = new DW.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "1169EC09", EditId = "2790D988" };
            DW.Extent extent1 = new DW.Extent() { Cx = dimensions.Item1, Cy = dimensions.Item2 };
            DW.EffectExtent effectExtent1 = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            DW.DocProperties docProperties1 = new DW.DocProperties() { Id = (UInt32Value)11U, Name = "Picture 11", Description = "A close up of a sign\n\nDescription automatically generated" };

            DW.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new DW.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)16U, Name = "Logo" };
            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId999" };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 1828800L, Cy = 348614L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run3.Append(runProperties3);
            run3.Append(drawing1);

            paragraph3.Append(paragraphProperties3);
            paragraph3.Append(run3);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph3);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            table.Append(tableProperties1);
            table.Append(tableGrid1);
            table.Append(tableRow1);

            header.Append(table);

        }

        private Tuple<long, long> AddImageToHeaderPart(HeaderContent headerContent, WordprocessingDocument doc, HeaderPart headerPart)
        {

            long width = 0, height = 0;
            var mainDocumentPart = doc.MainDocumentPart;

            var imagePartId = string.Empty;
            if (!string.IsNullOrEmpty(headerContent.Image))
            {

                var imgPart = headerPart.AddImagePart(ImagePartType.Jpeg, "rId999");
                imagePartId = headerPart.GetIdOfPart(imgPart);
                WebClient wc = new WebClient();
                using (Stream fs = wc.OpenRead(headerContent.Image))
                {
                    using (Image sourceImage = Image.FromStream(fs, true, true))
                    {
                        width = sourceImage.Width;
                        height = sourceImage.Height;
                    }
                }

                using (Stream fs = wc.OpenRead(headerContent.Image))
                {
                    imgPart.FeedData(fs);
                }

                log.Info("SetHeaderContent : iWidth:" + width);
                log.Info("SetHeaderContent : iHeight:" + height);

            }

            width = (long)Math.Round((decimal)width * 4127);
            height = (long)Math.Round((decimal)height * 4127);

            return new Tuple<long, long>(width, height);

        }

        public void AddImageAndPageNumber(HeaderContent headerContent, Header header, WordprocessingDocument doc, HeaderPart headerPart, JustificationValues firstJustification, JustificationValues secondJustification)
        {

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };
            TableJustification tableJustification1 = new TableJustification() { Val = TableRowAlignmentValues.Right };
            TableIndentation tableIndentation1 = new TableIndentation() { Width = -5, Type = TableWidthUnitValues.Dxa };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableIndentation1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);
            tableProperties1.Append(tableJustification1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "6298" };
            GridColumn gridColumn2 = new GridColumn() { Width = "2545" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "002F26B7", RsidTableRowAddition = "008B6C0B", RsidTableRowProperties = "008B6C0B", ParagraphId = "5FEC0FF4", TextId = "77777777" };

            TableRowProperties tableRowProperties1 = new TableRowProperties();
            TableRowHeight tableRowHeight1 = new TableRowHeight() { Val = (UInt32Value)990U };

            tableRowProperties1.Append(tableRowHeight1);

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "6667", Type = TableWidthUnitValues.Dxa };
            Shading shading1 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(shading1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "002F26B7", RsidParagraphAddition = "009A0695", RsidParagraphProperties = "002F26B7", RsidRunAdditionDefault = "009D198E", ParagraphId = "1A24ED6D", TextId = "7948A913" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Left };

            paragraphProperties1.Append(justification1);

            Run run1 = new Run() { RsidRunProperties = "002F26B7" };

            var dimensions = AddImageToHeaderPart(headerContent, doc, headerPart);

            Drawing drawing1 = new Drawing();

            DW.Inline inline1 = new DW.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "1169EC09", EditId = "2790D988" };
            DW.Extent extent1 = new DW.Extent() { Cx = dimensions.Item1, Cy = dimensions.Item2 };
            DW.EffectExtent effectExtent1 = new DW.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            DW.DocProperties docProperties1 = new DW.DocProperties() { Id = (UInt32Value)11U, Name = "Picture 11", Description = "A close up of a sign\n\nDescription automatically generated" };

            DW.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new DW.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)16U, Name = "Logo" };
            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId999" };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 1828800L, Cy = 348614L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run1.Append(drawing1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2683", Type = TableWidthUnitValues.Dxa };
            TableJustification tableJustification2 = new TableJustification() { Val = TableRowAlignmentValues.Right };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            Shading shading2 = new Shading() { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "auto" };

            tableCellProperties2.Append(tableCellWidth2);
            tableCellProperties2.Append(shading2);
            tableCellProperties2.Append(tableCellVerticalAlignment1);
            tableCellProperties2.Append(tableJustification2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "002F26B7", RsidParagraphAddition = "008B6C0B", RsidParagraphProperties = "002F26B7", RsidRunAdditionDefault = "009D198E", ParagraphId = "15B51B7C", TextId = "5066E76E" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties2.Append(justification2);

            Run run2 = new Run() { RsidRunProperties = "002F26B7" };
                        
            Run run_2 = new Run();
            FieldChar fieldChar_1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };
            run_2.Append(fieldChar_1);
            paragraph2.Append(run_2);

            Run run_3 = new Run();
            FieldCode fieldCode_1 = new FieldCode();
            fieldCode_1.Text = "PAGE";
            run_3.Append(fieldCode_1);
            paragraph2.Append(run_3);

            Run run_4 = new Run();
            FieldChar fieldChar_3 = new FieldChar() { FieldCharType = FieldCharValues.End };
            run_4.Append(fieldChar_3);
            paragraph2.Append(run_4);

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(run2);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            tableRow1.Append(tableRowProperties1);
            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);

            header.Append(table1);

        }

        public void AddTextAndPageNumber(HeaderContent headerContent, Header header)
        {
            
            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "2942" };
            GridColumn gridColumn2 = new GridColumn() { Width = "2943" };
            GridColumn gridColumn3 = new GridColumn() { Width = "2943" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00030AD9", RsidTableRowAddition = "00030AD9", RsidTableRowProperties = "00030AD9", ParagraphId = "5C30337A", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2942", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "00030AD9", ParagraphId = "0B418038", TextId = "10DDCC8B" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Left };

            paragraphProperties1.Append(justification1);

            Run run1 = new Run() { RsidRunProperties = "00030AD9" };
            Text text1 = new Text();
            text1.Text = headerContent.Content;

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);
            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "00030AD9", ParagraphId = "6AE09A11", TextId = "3F58E314" };

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };

            tableCellProperties3.Append(tableCellWidth3);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "0061456D", ParagraphId = "45836115", TextId = "39548499" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties2.Append(justification2);

            Run run2 = new Run() { RsidRunProperties = "0061456D" };
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run2.Append(fieldChar1);

            Run run3 = new Run() { RsidRunProperties = "0061456D" };
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

            run3.Append(fieldCode1);

            Run run4 = new Run() { RsidRunProperties = "0061456D" };
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run4.Append(fieldChar2);

            Run run5 = new Run() { RsidRunProperties = "0061456D" };

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);
            Text text2 = new Text();
            text2.Text = "1";

            run5.Append(runProperties1);
            run5.Append(text2);

            Run run6 = new Run() { RsidRunProperties = "0061456D" };

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run6.Append(runProperties2);
            run6.Append(fieldChar3);

            paragraph3.Append(paragraphProperties2);
            paragraph3.Append(run2);
            paragraph3.Append(run3);
            paragraph3.Append(run4);
            paragraph3.Append(run5);
            paragraph3.Append(run6);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "008B6C0B", RsidParagraphAddition = "008B6C0B", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "008B6C0B", ParagraphId = "1248EBF5", TextId = "20FD8BD3" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties3.Append(paragraphStyleId1);

            paragraph4.Append(paragraphProperties3);

            header.Append(table1);
            header.Append(paragraph4);

        }
        
        public void AddImageAndPageNumber(HeaderContent headerContent, Header header, WordprocessingDocument doc, HeaderPart headerPart)
        {

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "4836" };
            GridColumn gridColumn2 = new GridColumn() { Width = "1986" };
            GridColumn gridColumn3 = new GridColumn() { Width = "2016" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00030AD9", RsidTableRowAddition = "00030AD9", RsidTableRowProperties = "00030AD9", ParagraphId = "5C30337A", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2942", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "00D64038", ParagraphId = "0B418038", TextId = "298C2602" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Left };

            paragraphProperties1.Append(justification1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);
            
            ImagePart imagePart1 = headerPart.AddNewPart<ImagePart>("image/png", "rId1");
            var sizes = GenerateImagePart1Content(imagePart1, headerContent);

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "009AF219", EditId = "10D8591E" };
            Wp.Extent extent1 = new Wp.Extent() { Cx = sizes.Width, Cy = sizes.Height };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)2U, Name = "Picture 2", Description = "A picture containing leaf, outdoor object, grate\n\nDescription automatically generated" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 2", Description = "A picture containing leaf, outdoor object, grate\n\nDescription automatically generated" };
            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId1" };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = sizes.Width, Cy = sizes.Height };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run1.Append(runProperties1);
            run1.Append(drawing1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);
            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "00030AD9", ParagraphId = "6AE09A11", TextId = "3F58E314" };

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellVerticalAlignment1);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "0061456D", ParagraphId = "45836115", TextId = "39548499" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties2.Append(justification2);

            Run run2 = new Run() { RsidRunProperties = "0061456D" };
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run2.Append(fieldChar1);

            Run run3 = new Run() { RsidRunProperties = "0061456D" };
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

            run3.Append(fieldCode1);

            Run run4 = new Run() { RsidRunProperties = "0061456D" };
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run4.Append(fieldChar2);

            Run run5 = new Run() { RsidRunProperties = "0061456D" };

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);
            Text text1 = new Text();
            text1.Text = "1";

            run5.Append(runProperties2);
            run5.Append(text1);

            Run run6 = new Run() { RsidRunProperties = "0061456D" };

            RunProperties runProperties3 = new RunProperties();
            NoProof noProof3 = new NoProof();

            runProperties3.Append(noProof3);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run6.Append(runProperties3);
            run6.Append(fieldChar3);

            paragraph3.Append(paragraphProperties2);
            paragraph3.Append(run2);
            paragraph3.Append(run3);
            paragraph3.Append(run4);
            paragraph3.Append(run5);
            paragraph3.Append(run6);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);

            Paragraph paragraph4 = new Paragraph() { RsidParagraphMarkRevision = "008B6C0B", RsidParagraphAddition = "008B6C0B", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "008B6C0B", ParagraphId = "1248EBF5", TextId = "20FD8BD3" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties3.Append(paragraphStyleId1);

            paragraph4.Append(paragraphProperties3);


            header.Append(table1);
            header.Append(paragraph4);

        }

        public void AddImageAndText(HeaderContent headerContent, Header header, WordprocessingDocument doc, HeaderPart headerPart)
        {

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "4836" };
            GridColumn gridColumn2 = new GridColumn() { Width = "1907" };
            GridColumn gridColumn3 = new GridColumn() { Width = "2095" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00030AD9", RsidTableRowAddition = "00030AD9", RsidTableRowProperties = "00030AD9", ParagraphId = "5C30337A", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2942", Type = TableWidthUnitValues.Dxa };

            tableCellProperties1.Append(tableCellWidth1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "00D64038", ParagraphId = "0B418038", TextId = "298C2602" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Left };

            paragraphProperties1.Append(justification1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            #region image

            ImagePart imagePart1 = headerPart.AddNewPart<ImagePart>("image/png", "rId1");
            var sizes = GenerateImagePart1Content(imagePart1, headerContent);

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "009AF219", EditId = "10D8591E" };
            Wp.Extent extent1 = new Wp.Extent() { Cx = sizes.Width, Cy = sizes.Height };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)2U, Name = "Picture 2", Description = "A picture containing leaf, outdoor object, grate\n\nDescription automatically generated" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 2", Description = "A picture containing leaf, outdoor object, grate\n\nDescription automatically generated" };
            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId1" };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = sizes.Width, Cy = sizes.Height };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run1.Append(runProperties1);
            run1.Append(drawing1);

            #endregion image

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);
            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "00030AD9", ParagraphId = "6AE09A11", TextId = "3F58E314" };

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
            
            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellVerticalAlignment1);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "009322A8", RsidParagraphProperties = "00197CB7", RsidRunAdditionDefault = "009322A8", ParagraphId = "45836115", TextId = "5EE1CAA5" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification2 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties2.Append(justification2);

            Run run2 = new Run();
            Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text1.Text = headerContent.Content;

            run2.Append(text1);
            
            paragraph3.Append(paragraphProperties2);
            paragraph3.Append(run2);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);

            header.Append(table1);

        }

        public void AddTextAndImage(HeaderContent headerContent, Header header, WordprocessingDocument doc, HeaderPart headerPart)
        {

            Table table1 = new Table();

            TableProperties tableProperties1 = new TableProperties();
            TableStyle tableStyle1 = new TableStyle() { Val = "TableGrid" };
            TableWidth tableWidth1 = new TableWidth() { Width = "0", Type = TableWidthUnitValues.Auto };

            TableBorders tableBorders1 = new TableBorders();
            TopBorder topBorder1 = new TopBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            LeftBorder leftBorder1 = new LeftBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            BottomBorder bottomBorder1 = new BottomBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            RightBorder rightBorder1 = new RightBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideHorizontalBorder insideHorizontalBorder1 = new InsideHorizontalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };
            InsideVerticalBorder insideVerticalBorder1 = new InsideVerticalBorder() { Val = BorderValues.None, Color = "auto", Size = (UInt32Value)0U, Space = (UInt32Value)0U };

            tableBorders1.Append(topBorder1);
            tableBorders1.Append(leftBorder1);
            tableBorders1.Append(bottomBorder1);
            tableBorders1.Append(rightBorder1);
            tableBorders1.Append(insideHorizontalBorder1);
            tableBorders1.Append(insideVerticalBorder1);
            TableLook tableLook1 = new TableLook() { Val = "04A0" };

            tableProperties1.Append(tableStyle1);
            tableProperties1.Append(tableWidth1);
            tableProperties1.Append(tableBorders1);
            tableProperties1.Append(tableLook1);

            TableGrid tableGrid1 = new TableGrid();
            GridColumn gridColumn1 = new GridColumn() { Width = "2098" };
            GridColumn gridColumn2 = new GridColumn() { Width = "1904" };
            GridColumn gridColumn3 = new GridColumn() { Width = "4836" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            TableRow tableRow1 = new TableRow() { RsidTableRowMarkRevision = "00030AD9", RsidTableRowAddition = "00030AD9", RsidTableRowProperties = "00030AD9", ParagraphId = "5C30337A", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2942", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphMarkRevision = "00405C13", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00405C13", RsidRunAdditionDefault = "00405C13", ParagraphId = "0B418038", TextId = "195DF203" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();

            Run run1 = new Run();
            Text text1 = new Text() { Space = SpaceProcessingModeValues.Preserve };
            text1.Text = headerContent.Content;

            run1.Append(text1);
            ProofError proofError1 = new ProofError() { Type = ProofingErrorValues.SpellStart };
            
            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);
            paragraph1.Append(proofError1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);
            Paragraph paragraph2 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "00030AD9", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "00030AD9", ParagraphId = "6AE09A11", TextId = "3F58E314" };

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };

            tableCellProperties3.Append(tableCellWidth3);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphMarkRevision = "00030AD9", RsidParagraphAddition = "009322A8", RsidParagraphProperties = "00030AD9", RsidRunAdditionDefault = "00405C13", ParagraphId = "45836115", TextId = "032A12AC" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            Justification justification1 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties2.Append(justification1);

            Run run3 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            #region image 
            ImagePart imagePart1 = headerPart.AddNewPart<ImagePart>("image/png", "rId1");
            var sizes = GenerateImagePart1Content(imagePart1, headerContent);

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "001B0191", EditId = "2717CA2F" };
            Wp.Extent extent1 = new Wp.Extent() { Cx = sizes.Width, Cy = sizes.Height };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)2U, Name = "Picture 2", Description = "A picture containing leaf, outdoor object, grate\n\nDescription automatically generated" };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" };

            Pic.Picture picture1 = new Pic.Picture();
            picture1.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");

            Pic.NonVisualPictureProperties nonVisualPictureProperties1 = new Pic.NonVisualPictureProperties();
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)2U, Name = "Picture 2", Description = "A picture containing leaf, outdoor object, grate\n\nDescription automatically generated" };
            Pic.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Pic.NonVisualPictureDrawingProperties();

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Pic.BlipFill blipFill1 = new Pic.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "rId1" };

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Pic.ShapeProperties shapeProperties1 = new Pic.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = sizes.Width, Cy = sizes.Height};

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);

            graphicData1.Append(picture1);

            graphic1.Append(graphicData1);

            inline1.Append(extent1);
            inline1.Append(effectExtent1);
            inline1.Append(docProperties1);
            inline1.Append(nonVisualGraphicFrameDrawingProperties1);
            inline1.Append(graphic1);

            drawing1.Append(inline1);

            run3.Append(runProperties1);
            run3.Append(drawing1);

            #endregion image

            paragraph3.Append(paragraphProperties2);
            paragraph3.Append(run3);

            tableCell3.Append(tableCellProperties3);
            tableCell3.Append(paragraph3);

            tableRow1.Append(tableCell1);
            tableRow1.Append(tableCell2);
            tableRow1.Append(tableCell3);

            table1.Append(tableProperties1);
            table1.Append(tableGrid1);
            table1.Append(tableRow1);
            
            header.Append(table1);

        }

        private ImageSize GenerateImagePart1Content(ImagePart imagePart1, HeaderContent headerContent)
        {

            ImageSize sizes = new ImageSize();
            var imagePart1Data = string.Empty;
            if (!string.IsNullOrEmpty(headerContent.Image))
            {

                WebClient wc = new WebClient();
                using (Stream fs = wc.OpenRead(headerContent.Image))
                {
                    using (Image sourceImage = Image.FromStream(fs, true, true))
                    {

                        using (MemoryStream m = new MemoryStream())
                        {
                            sourceImage.Save(m, sourceImage.RawFormat);
                            byte[] imageBytes = m.ToArray();

                            sizes.Height = (int)Math.Round((decimal)sourceImage.Height * 9525); 
                            sizes.Width = (int)Math.Round((decimal)sourceImage.Width * 9525);

                            // Convert byte[] to Base64 String
                            imagePart1Data = Convert.ToBase64String(imageBytes);
                        }

                    }
                }

            }
            
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();

            return sizes;

        }

        private System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }

    }

}