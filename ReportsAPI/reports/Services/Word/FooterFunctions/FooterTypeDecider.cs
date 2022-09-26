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
using FooterContent = reports.Models.FooterContent;
using System.Net;
using System.IO;
using System.Drawing;
using log4net;
using System.Linq;
using HtmlToOpenXml.Extensions;
using reports.Services.Word.Models;

namespace reports.Services.Word.FooterFunctions
{

    public enum FooterTypes
    {
        none = 0,
        pagenumber_center,
        pagenumber_left,
        pagenumber_right,
        image_and_pagenumber,
        text_and_pagenumber
    }

    public abstract class FooterTypeDecider
    {

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        protected abstract void Run(WordprocessingDocument doc, FooterContent headerContent, FooterPart footerPart, Footer footer);
        
        public void TemplateMethod(WordprocessingDocument doc, FooterContent footerContent)
        {

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

            if(!string.IsNullOrEmpty(footerContent.Divider))
            {
                AddDivider(footerContent, footer);
            }

            Run(doc, footerContent, footerPart, footer);

            footerPart.Footer = footer;
            footerPart.Footer.Save();

        }

        protected void AddPageNumber(FooterContent footerContent, Footer footer, JustificationValues position)
        {
            //set page number in footer
            if (!string.IsNullOrEmpty(footerContent.Position))
            {

                Paragraph paragraph2 = new Paragraph();
                ParagraphProperties paragraphProperties = new ParagraphProperties();
                ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "Footer" };
                paragraphProperties.Append(paragraphStyleId);
                Justification justification = new Justification() { Val = position };
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
        }

        protected void AddText(FooterContent footerContent, Footer footer)
        {
            if (!string.IsNullOrEmpty(footerContent.Content))
            {
                Paragraph paragraph1 = new Paragraph() { };
                Run run1 = new Run();
                Text text1 = new Text();
                text1.Text = footerContent.Content;
                run1.Append(text1);
                paragraph1.Append(run1);
                footer.Append(paragraph1);
                log.Info("SetFooterContent : footer added content =" + footerContent.Content);
            }
        }

        public void AddDivider(FooterContent footerContent, Footer footer)
        {

            if (string.IsNullOrEmpty(footerContent.Divider))
            {
                return;
            }

            //AddEmptyParagraphToHeader(header);

            var fillColor = footerContent.Divider.FillColorOrDefault();
            var thickness = footerContent.Divider.ThicknessOrDefault();

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

            footer.Append(paragraph1);

        }

        public void GeneratePicture(FooterContent footerContent, FooterPart footerPart, WordprocessingDocument doc, Footer footer, JustificationValues justification)
        {

            long iWidth = 0, iHeight = 0;
            var mainDocumentPart = doc.MainDocumentPart;

            var imagePartId = string.Empty;
            if (!string.IsNullOrEmpty(footerContent.Image))
            {

                var imgPart = footerPart.AddImagePart(ImagePartType.Jpeg, "rId999");
                imagePartId = footerPart.GetIdOfPart(imgPart);
                WebClient wc = new WebClient();
                using (Stream fs = wc.OpenRead(footerContent.Image))
                {
                    using (Image sourceImage = Image.FromStream(fs, true, true))
                    {
                        iWidth = sourceImage.Width;
                        iHeight = sourceImage.Height;
                    }
                }

                using (Stream fs = wc.OpenRead(footerContent.Image))
                {
                    imgPart.FeedData(fs);
                }

                log.Info("SetHeaderContent : iWidth:" + iWidth);
                log.Info("SetHeaderContent : iHeight:" + iHeight);

            }

            iWidth = (long)Math.Round((decimal)iWidth * 4127);
            iHeight = (long)Math.Round((decimal)iHeight * 4127);

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
            footer.Append(paragraph);

        }

        public void AddImageAndPageNumber(WordprocessingDocument doc, FooterContent footerContent, FooterPart footerPart, Footer footer)
        {

            ImagePart imagePart1 = footerPart.AddNewPart<ImagePart>("image/png", "rId1");
            var sizes = GenerateImagePart1Content(imagePart1, footerContent);

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
            GridColumn gridColumn1 = new GridColumn() { Width = "3457" };
            GridColumn gridColumn2 = new GridColumn() { Width = "2685" };
            GridColumn gridColumn3 = new GridColumn() { Width = "2696" };

            tableGrid1.Append(gridColumn1);
            tableGrid1.Append(gridColumn2);
            tableGrid1.Append(gridColumn3);

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "00DA7385", RsidTableRowProperties = "00DA7385", ParagraphId = "4AD080B7", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2942", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00DA7385", RsidParagraphProperties = "00DA7385", RsidRunAdditionDefault = "00DA7385", ParagraphId = "3241FE9D", TextId = "67F379B3" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Drawing drawing1 = new Drawing();

            Wp.Inline inline1 = new Wp.Inline() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)0U, DistanceFromRight = (UInt32Value)0U, AnchorId = "0290BF18", EditId = "3ACBDB8A" };
            Wp.Extent extent1 = new Wp.Extent() { Cx = sizes.Width, Cy = sizes.Height };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L };
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = (UInt32Value)1U, Name = "Picture 1", Description = "A picture containing leaf, grate\n\nDescription automatically generated" };

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
            Pic.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Pic.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "Picture 1", Description = "A picture containing leaf, grate\n\nDescription automatically generated" };
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

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00DA7385", RsidRunAdditionDefault = "00DA7385", ParagraphId = "28D1A587", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties2.Append(paragraphStyleId2);

            paragraph2.Append(paragraphProperties2);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellVerticalAlignment2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00DA7385", RsidParagraphProperties = "00DA7385", RsidRunAdditionDefault = "00DA7385", ParagraphId = "0DC7F361", TextId = "114E9C72" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "Footer" };
            Justification justification1 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties3.Append(paragraphStyleId3);
            paragraphProperties3.Append(justification1);

            Run run2 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run2.Append(fieldChar1);

            Run run3 = new Run();
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

            run3.Append(fieldCode1);

            Run run4 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run4.Append(fieldChar2);

            Run run5 = new Run();

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);
            Text text1 = new Text();
            text1.Text = "1";

            run5.Append(runProperties2);
            run5.Append(text1);

            Run run6 = new Run();

            RunProperties runProperties3 = new RunProperties();
            NoProof noProof3 = new NoProof();

            runProperties3.Append(noProof3);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run6.Append(runProperties3);
            run6.Append(fieldChar3);

            paragraph3.Append(paragraphProperties3);
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

            footer.Append(table1);

        }

        public void AddTextAndPageNumber(WordprocessingDocument doc, FooterContent footerContent, FooterPart footerPart, Footer footer)
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

            TableRow tableRow1 = new TableRow() { RsidTableRowAddition = "00DA7385", RsidTableRowProperties = "00DA7385", ParagraphId = "4AD080B7", TextId = "77777777" };

            TableCell tableCell1 = new TableCell();

            TableCellProperties tableCellProperties1 = new TableCellProperties();
            TableCellWidth tableCellWidth1 = new TableCellWidth() { Width = "2942", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment1 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties1.Append(tableCellWidth1);
            tableCellProperties1.Append(tableCellVerticalAlignment1);

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00DA7385", RsidParagraphProperties = "00DA7385", RsidRunAdditionDefault = "00EE1BC0", ParagraphId = "3241FE9D", TextId = "42C94CA9" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = footerContent.Content;

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            tableCell1.Append(tableCellProperties1);
            tableCell1.Append(paragraph1);

            TableCell tableCell2 = new TableCell();

            TableCellProperties tableCellProperties2 = new TableCellProperties();
            TableCellWidth tableCellWidth2 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };

            tableCellProperties2.Append(tableCellWidth2);

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00DA7385", RsidRunAdditionDefault = "00DA7385", ParagraphId = "28D1A587", TextId = "77777777" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties2.Append(paragraphStyleId2);

            paragraph2.Append(paragraphProperties2);

            tableCell2.Append(tableCellProperties2);
            tableCell2.Append(paragraph2);

            TableCell tableCell3 = new TableCell();

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            TableCellWidth tableCellWidth3 = new TableCellWidth() { Width = "2943", Type = TableWidthUnitValues.Dxa };
            TableCellVerticalAlignment tableCellVerticalAlignment2 = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

            tableCellProperties3.Append(tableCellWidth3);
            tableCellProperties3.Append(tableCellVerticalAlignment2);

            Paragraph paragraph3 = new Paragraph() { RsidParagraphAddition = "00DA7385", RsidParagraphProperties = "00DA7385", RsidRunAdditionDefault = "00DA7385", ParagraphId = "0DC7F361", TextId = "114E9C72" };

            ParagraphProperties paragraphProperties3 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId3 = new ParagraphStyleId() { Val = "Footer" };
            Justification justification1 = new Justification() { Val = JustificationValues.Right };

            paragraphProperties3.Append(paragraphStyleId3);
            paragraphProperties3.Append(justification1);

            Run run2 = new Run();
            FieldChar fieldChar1 = new FieldChar() { FieldCharType = FieldCharValues.Begin };

            run2.Append(fieldChar1);

            Run run3 = new Run();
            FieldCode fieldCode1 = new FieldCode() { Space = SpaceProcessingModeValues.Preserve };
            fieldCode1.Text = " PAGE   \\* MERGEFORMAT ";

            run3.Append(fieldCode1);

            Run run4 = new Run();
            FieldChar fieldChar2 = new FieldChar() { FieldCharType = FieldCharValues.Separate };

            run4.Append(fieldChar2);

            Run run5 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);
            Text text2 = new Text();
            text2.Text = "1";

            run5.Append(runProperties1);
            run5.Append(text2);

            Run run6 = new Run();

            RunProperties runProperties2 = new RunProperties();
            NoProof noProof2 = new NoProof();

            runProperties2.Append(noProof2);
            FieldChar fieldChar3 = new FieldChar() { FieldCharType = FieldCharValues.End };

            run6.Append(runProperties2);
            run6.Append(fieldChar3);

            paragraph3.Append(paragraphProperties3);
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

            footer.Append(table1);

        }

        private ImageSize GenerateImagePart1Content(ImagePart imagePart1, FooterContent footerContent)
        {

            ImageSize sizes = new ImageSize();
            var imagePart1Data = string.Empty;
            if (!string.IsNullOrEmpty(footerContent.Image))
            {

                WebClient wc = new WebClient();
                using (Stream fs = wc.OpenRead(footerContent.Image))
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
