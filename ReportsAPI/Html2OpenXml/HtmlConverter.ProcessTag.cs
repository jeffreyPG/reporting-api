/* Copyright (C) Olivier Nizet https://github.com/onizet/html2openxml - All Rights Reserved
 * 
 * This source is subject to the Microsoft Permissive License.
 * Please see the License.txt file for more information.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY 
 * KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
 * PARTICULAR PURPOSE.
 */
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using A = DocumentFormat.OpenXml.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Wp14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;

namespace HtmlToOpenXml
{
	using a = DocumentFormat.OpenXml.Drawing;
	using pic = DocumentFormat.OpenXml.Drawing.Pictures;
	using wBorder = DocumentFormat.OpenXml.Wordprocessing.Border;

	partial class HtmlConverter
	{
        //____________________________________________________________________
        //
        // Processing known tags
        UInt32 uintDigit = 1;

        #region ProcessAcronym

        private void ProcessAcronym(HtmlEnumerator en)
		{
			// Transform the inline acronym/abbreviation to a reference to a foot note.

			string title = en.Attributes["title"];
			if (title == null) return;

            AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);

			if (elements.Count > 0 && elements[0] is Run)
			{
				string runStyle;
				FootnoteEndnoteReferenceType reference;

				if (this.AcronymPosition == AcronymPosition.PageEnd)
				{
					reference = new FootnoteReference() { Id = AddFootnoteReference(title) };
					runStyle = "FootnoteReference";
				}
				else
				{
					reference = new EndnoteReference() { Id = AddEndnoteReference(title) };
					runStyle = "EndnoteReference";
				}

				Run run;
				elements.Add(
					run = new Run(
						new RunProperties {
							RunStyle = new RunStyle() { Val = htmlStyles.GetStyle(runStyle, StyleValues.Character) }
						},
						reference));
			}
		}

		#endregion

		#region ProcessBlockQuote

		private void ProcessBlockQuote(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);

			string tagName = en.CurrentTag;
			string cite = en.Attributes["cite"];

			htmlStyles.Paragraph.BeginTag(en.CurrentTag, new ParagraphStyleId() { Val = htmlStyles.GetStyle("IntenseQuote") });

			AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);

			if (cite != null)
			{
				string runStyle;
				FootnoteEndnoteReferenceType reference;

				if (this.AcronymPosition == AcronymPosition.PageEnd)
				{
					reference = new FootnoteReference() { Id = AddFootnoteReference(cite) };
					runStyle = "FootnoteReference";
				}
				else
				{
					reference = new EndnoteReference() { Id = AddEndnoteReference(cite) };
					runStyle = "EndnoteReference";
				}

				Run run;
				elements.Add(
					run = new Run(
						new RunProperties {
							RunStyle = new RunStyle() { Val = htmlStyles.GetStyle(runStyle, StyleValues.Character) }
						},
						reference));
			}

			CompleteCurrentParagraph(true);
			htmlStyles.Paragraph.EndTag(tagName);
		}

		#endregion

		#region ProcessBody

		private void ProcessBody(HtmlEnumerator en)
		{

			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

            if (styleAttributes.Count > 0)
            {
                htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());
            }

			// Unsupported W3C attribute but claimed by users. Specified at <body> level, the page
			// orientation is applied on the whole document
			string attr = en.StyleAttributes["page-orientation"];
			if (attr != null)
			{
				PageOrientationValues orientation = Converter.ToPageOrientation(attr);

                SectionProperties sectionProperties = mainPart.Document.Body.GetFirstChild<SectionProperties>();
                if (sectionProperties == null || sectionProperties.GetFirstChild<PageSize>() == null)
                {
                    mainPart.Document.Body.Append(HtmlConverter.ChangePageOrientation(orientation));
                }
                else
                {
                    PageSize pageSize = sectionProperties.GetFirstChild<PageSize>();
                    if (!pageSize.Compare(orientation))
                    {
                        SectionProperties validSectionProp = ChangePageOrientation(orientation);
                        if (pageSize != null) pageSize.Remove();
                        sectionProperties.PrependChild(validSectionProp.GetFirstChild<PageSize>().CloneNode(true));
                    }
                }
            }
		}

		#endregion

		#region ProcessBr

		private void ProcessBr(HtmlEnumerator en)
		{
			elements.Add(new Run(new Break()));
		}

		#endregion

		#region ProcessCite

		private void ProcessCite(HtmlEnumerator en)
		{
			ProcessHtmlElement<RunStyle>(en, new RunStyle() { Val = htmlStyles.GetStyle("Quote", StyleValues.Character) });
		}

		#endregion

		#region ProcessDefinitionList

		private void ProcessDefinitionList(HtmlEnumerator en)
		{
			ProcessParagraph(en);
			currentParagraph.InsertInProperties(prop => prop.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" });
		}

		#endregion

		#region ProcessDefinitionListItem

		private void ProcessDefinitionListItem(HtmlEnumerator en)
		{
			AlternateProcessHtmlChunks(en, "</dd>");

			currentParagraph = htmlStyles.Paragraph.NewParagraph();
			currentParagraph.Append(elements);
			currentParagraph.InsertInProperties(prop => {
				prop.Indentation = new Indentation() { FirstLine = "708" };
				prop.SpacingBetweenLines = new SpacingBetweenLines() { After = "0" };
			});

			// Restore the original elements list
			AddParagraph(currentParagraph);
			this.elements.Clear();
		}

		#endregion

		#region ProcessDiv

		private void ProcessDiv(HtmlEnumerator en)
		{
			// The way the browser consider <div> is like a simple Break. But in case of any attributes that targets
			// the paragraph, we don't want to apply the style on the old paragraph but on a new one.
			if (en.Attributes.Count == 0 || (en.StyleAttributes["text-align"] == null && en.Attributes["align"] == null && en.StyleAttributes.GetAsBorder("border").IsEmpty))
			{
				List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
				bool newParagraph = ProcessContainerAttributes(en, runStyleAttributes);
				CompleteCurrentParagraph(newParagraph);

				if (runStyleAttributes.Count > 0)
                { 
					htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes);
                }

                // Any changes that requires a new paragraph?
                if (newParagraph)
				{
					// Insert before the break, complete this paragraph and start a new one
					this.paragraphs.Insert(this.paragraphs.Count - 1, currentParagraph);
					AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
					CompleteCurrentParagraph();
				}
			}
			else
			{
				// treat div as a paragraph
				ProcessParagraph(en);
			}
		}

		#endregion

		#region ProcessFont

		private void ProcessFont(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			ProcessContainerAttributes(en, styleAttributes);

			string attrValue = en.Attributes["size"];
			if (attrValue != null)
			{
				Unit fontSize = Converter.ToFontSize(attrValue);
                if (fontSize.IsFixed)
                { 
					styleAttributes.Add(new FontSize
                    {
                        Val = (fontSize.ValueInPoint * 2).ToString(CultureInfo.InvariantCulture)
                    });
                }
            }

			attrValue = en.Attributes["face"];
			if (attrValue != null)
			{
				// Set HightAnsi. Bug fixed by xjpmauricio on github.com/onizet/html2openxml/discussions/285439
				// where characters with accents were always using fallback font
				styleAttributes.Add(new RunFonts { Ascii = attrValue, HighAnsi = attrValue });
			}

			if (styleAttributes.Count > 0)
            { 
				htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);
            }
        }

		#endregion

		#region ProcessHeading

		private void ProcessHeading(HtmlEnumerator en)
		{
			char level = en.Current[2];

			// support also style attributes for heading (in case of css override)
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

			AlternateProcessHtmlChunks(en, "</h" + level + ">");
			Paragraph p = new Paragraph(elements);
			p.InsertInProperties(prop =>
				prop.ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle("Heading" + level, StyleValues.Paragraph) });

			htmlStyles.Paragraph.ApplyTags(p);
			htmlStyles.Paragraph.EndTag("<h" + level + ">");

			this.elements.Clear();
			AddParagraph(p);
			AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
		}

		#endregion

		#region ProcessHorizontalLine

		private void ProcessHorizontalLine(HtmlEnumerator en)
		{
			// Insert an horizontal line as it stands in many emails.
            CompleteCurrentParagraph(true);

			// If the previous paragraph contains a bottom border or is a Table, we add some spacing between the <hr>
			// and the previous element or Word will display only the last border.
			// (see Remarks: http://msdn.microsoft.com/en-us/library/documentformat.openxml.wordprocessing.bottomborder%28office.14%29.aspx)
            if (paragraphs.Count >= 2)
            {
                OpenXmlCompositeElement previousElement = paragraphs[paragraphs.Count - 2];
                bool addSpacing = false;
                ParagraphProperties prop = previousElement.GetFirstChild<ParagraphProperties>();
                if (prop != null)
                {
                    if (prop.ParagraphBorders != null 
                        && prop.ParagraphBorders.BottomBorder != null
                        && prop.ParagraphBorders.BottomBorder.Size > 0U)
                    { 
                            addSpacing = true;
                    }
                }
                else
                {
                    if (previousElement is Table)
                    { 
                        addSpacing = true;
                    }
                }

                if (addSpacing)
                {
                    currentParagraph.InsertInProperties(p => p.SpacingBetweenLines = new SpacingBetweenLines() { Before = "240" });
                }
			}

			// if this paragraph has no children, it will be deleted in RemoveEmptyParagraphs()
			// in order to kept the <hr>, we force an empty run
            currentParagraph.Append(new Run());

            currentParagraph.InsertInProperties(prop => 
				prop.ParagraphBorders = new ParagraphBorders {
					TopBorder = new TopBorder() { Val = BorderValues.Single, Size = 4U }
				});
		}

		#endregion

		#region ProcessHtml

		private void ProcessHtml(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			htmlStyles.Paragraph.ProcessCommonAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
            { 
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());
            }
        }

		#endregion

		#region ProcessHtmlElement

		private void ProcessHtmlElement<T>(HtmlEnumerator en) where T: OpenXmlLeafElement, new()
		{
			ProcessHtmlElement<T>(en, new T());
		}

		/// <summary>
		/// Generic handler for processing style on any Html element.
		/// </summary>
		private void ProcessHtmlElement<T>(HtmlEnumerator en, OpenXmlLeafElement style) where T: OpenXmlLeafElement
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>() { style };
			ProcessContainerAttributes(en, styleAttributes);
			htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);
		}

		#endregion

		#region ProcessFigureCaption

		private void ProcessFigureCaption(HtmlEnumerator en)
		{

			this.CompleteCurrentParagraph(true);

			currentParagraph.Append(
					new ParagraphProperties {
						ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle("Caption", StyleValues.Paragraph) },
						KeepNext = new KeepNext()
					},
					new Run(
						new Text("Figure ") { Space = SpaceProcessingModeValues.Preserve }
					),
					new SimpleField(
						new Run(
							new Text(AddFigureCaption().ToString(CultureInfo.InvariantCulture)))
					) { Instruction = " SEQ Figure \\* ARABIC " }
				);

			ProcessHtmlChunks(en, "</figcaption>");

			if (elements.Count > 0) // any caption?
			{
				Text t = (elements[0] as Run).GetFirstChild<Text>();
				t.Text = " " + t.InnerText; // append a space after the numero of the picture
			}

			this.CompleteCurrentParagraph(true);
		}

		#endregion

		#region ProcessImage

		private void ProcessImage(HtmlEnumerator en)
		{

            if (this.ImageProcessing == ImageProcessing.Ignore)
            {
                return;
            }

			Drawing drawing = null;
			wBorder border = new wBorder() { Val = BorderValues.None };
			string src = en.Attributes["src"];
            string timeStamp = string.IsNullOrEmpty(en.Attributes["timestamp"]) ? "" : en.Attributes["timestamp"];
            Uri uri = null;

			// Bug reported by Erik2014. Inline 64 bit images can be too big and Uri.TryCreate will fail silently with a SizeLimit error.
			// To circumvent this buffer size, we will work either on the Uri, either on the original src.
			if (src != null && (DataUri.IsWellFormed(src) || Uri.TryCreate(src, UriKind.RelativeOrAbsolute, out uri)))
			{

				string alt = (en.Attributes["title"] ?? en.Attributes["alt"]) ?? String.Empty;
				bool process = true;

				if (uri != null && !uri.IsAbsoluteUri && this.BaseImageUrl != null)
                { 
					uri = new Uri(this.BaseImageUrl, uri);
                }

                Size preferredSize = Size.Empty;
				Unit wu = en.Attributes.GetAsUnit("width");
                if (!wu.IsValid)
                {
                    wu = en.StyleAttributes.GetAsUnit("width");
                }

				Unit hu = en.Attributes.GetAsUnit("height");
                if (!hu.IsValid)
                {
                    hu = en.StyleAttributes.GetAsUnit("height");
                }

				// % is not supported
				if (wu.IsFixed && wu.Value > 0)
				{
					preferredSize.Width = wu.ValueInPx;
				}
                if (hu.IsFixed && hu.Value > 0)
				{
					// Image perspective skewed. Bug fixed by ddeforge on github.com/onizet/html2openxml/discussions/350500
					preferredSize.Height = hu.ValueInPx;
				}

				SideBorder attrBorder = en.StyleAttributes.GetAsSideBorder("border");
				if (attrBorder.IsValid)
				{
					border.Val = attrBorder.Style;
					border.Color = attrBorder.Color.ToHexString();
					border.Size = (uint) attrBorder.Width.ValueInPx * 4;
				}
				else
				{
					var attrBorderWidth = en.Attributes.GetAsUnit("border");
					if (attrBorderWidth.IsValid)
					{
						border.Val = BorderValues.Single;
						border.Size = (uint) attrBorderWidth.ValueInPx * 4;
					}
				}
                try
                {
                    if (process)
                    { 
                        drawing = AddImagePart(uri, src, alt, preferredSize, timeStamp);
                    }
                }
                catch
                {
                    // TODO: figure this out and throw/log accordingly

                }
			}

			if (drawing != null)
			{
				Run run = new Run(drawing);
				if (border.Val != BorderValues.None) run.InsertInProperties(prop => prop.Border = border);
				elements.Add(run);
			}
		}

		#endregion

		#region ProcessLi

		private void ProcessLi(HtmlEnumerator en)
		{

			CompleteCurrentParagraph(false);
			currentParagraph = htmlStyles.Paragraph.NewParagraph();

			int numberingId = htmlStyles.NumberingList.ProcessItem(en);
			int level = htmlStyles.NumberingList.LevelIndex;

			// Save the new paragraph reference to support nested numbering list.
			Paragraph p = currentParagraph;
			currentParagraph.InsertInProperties(prop => {
				prop.ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle("ListParagraph", StyleValues.Paragraph) };
				prop.Indentation = level < 2 ? null : new Indentation() { Left = (level * 780).ToString(CultureInfo.InvariantCulture) };
				prop.NumberingProperties = new NumberingProperties {
					NumberingLevelReference = new NumberingLevelReference() { Val = level - 1 },
					NumberingId = new NumberingId() { Val = numberingId }
				};
			});

			// Restore the original elements list
			AddParagraph(currentParagraph);

			// Continue to process the html until we found </li>
			HtmlStyles.Paragraph.ApplyTags(currentParagraph);
			AlternateProcessHtmlChunks(en, "</li>");
			p.Append(elements);
			this.elements.Clear();
		}

		#endregion

		#region ProcessLink

		private void ProcessLink(HtmlEnumerator en)
		{

			string att = en.Attributes["href"];
			Hyperlink h = null;
			Uri uri = null;
            
			if (!String.IsNullOrEmpty(att))
			{
				// handle link where the http:// is missing and that starts directly with www
				if(att.StartsWith("www.", StringComparison.OrdinalIgnoreCase))
                { 
					att = "http://" + att;
                }

                // is it an anchor?
                if (att[0] == '#' && att.Length > 1)
				{
					// Always accept _top anchor
					if (!this.ExcludeLinkAnchor || att == "#_top")
					{
						h = new Hyperlink(
							) { History = true, Anchor = att.Substring(1) };
					}
				}
				// ensure the links does not start with javascript:
				else if (Uri.TryCreate(att, UriKind.Absolute, out uri) && uri.Scheme != "javascript")
				{
					HyperlinkRelationship extLink = mainPart.AddHyperlinkRelationship(uri, true);

					h = new Hyperlink() { History = true, Id = extLink.Id };
				}
			}

			if (h == null)
			{
				// link to a broken url, simply process the content of the tag
				ProcessHtmlChunks(en, "</a>");
				return;
			}

			att = en.Attributes["title"];
            if (!String.IsNullOrEmpty(att))
            {
                h.Tooltip = att;
            }

			AlternateProcessHtmlChunks(en, "</a>", true);

			if (linkElements.Count == 0) return;

			// Let's see whether the link tag include an image inside its body.
			// If so, the Hyperlink OpenXmlElement is lost and we'll keep only the images
			// and applied a HyperlinkOnClick attribute.
			List<OpenXmlElement> imageInLink = linkElements.FindAll(e => { return e.HasChild<Drawing>(); });
			if (imageInLink.Count != 0)
			{
				for (int i = 0; i < imageInLink.Count; i++)
				{
					// Retrieves the "alt" attribute of the image and apply it as the link's tooltip
					Drawing d = imageInLink[i].GetFirstChild<Drawing>();
					var enDp = d.Descendants<pic.NonVisualDrawingProperties>().GetEnumerator();
					String alt;
                    if (enDp.MoveNext())
                    {
                        alt = enDp.Current.Description;
                    }
                    else
                    {
                        alt = null;
                    }

					d.InsertInDocProperties(
							new a.HyperlinkOnClick() { Id = h.Id ?? h.Anchor, Tooltip = alt });
				}
			}

			// Append the processed elements and put them to the Run of the Hyperlink
			h.Append(linkElements);

			// can't use GetFirstChild<Run> or we may find the one containing the image
			foreach (var el in h.ChildElements)
			{
				Run run = el as Run;
				if (run != null && !run.HasChild<Drawing>())
				{
					run.InsertInProperties(prop =>
						prop.RunStyle = new RunStyle() { Val = htmlStyles.GetStyle("Hyperlink", StyleValues.Character) });
					break;
				}
			}

			this.linkElements.Clear();

			// Append the hyperlink
			elements.Add(h);

            if (imageInLink.Count > 0)
            {
                CompleteCurrentParagraph(true);
            }
		}

		#endregion

		#region ProcessNumberingList

		private void ProcessNumberingList(HtmlEnumerator en)
		{
			htmlStyles.NumberingList.BeginList(en);
		}

		#endregion

		#region ProcessParagraph

		private void ProcessParagraph(HtmlEnumerator en)
		{

			CompleteCurrentParagraph(true);

			// Respect this order: this is the way the browsers apply them
			string attrValue = en.StyleAttributes["text-align"] ?? en.Attributes["align"];
			
			if (attrValue != null)
			{
				JustificationValues? align = Converter.ToParagraphAlign(attrValue);
				if (align.HasValue)
				{
					currentParagraph.InsertInProperties(prop => prop.Justification = new Justification { Val = align });
				}
			}

            var hasId = en.Attributes["id"];
            if (!string.IsNullOrEmpty(hasId))
            {
                // do something
                currentParagraph.SetAttribute(new OpenXmlAttribute("id", "", hasId.ToLower()));
            }

			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			bool newParagraph = ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
            { 
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());
            }

            if (newParagraph)
			{
				AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
				ProcessClosingParagraph(en);
			}
		}

		#endregion

		#region ProcessPre

		private void ProcessPre(HtmlEnumerator en)
		{
			CompleteCurrentParagraph();
			currentParagraph = htmlStyles.Paragraph.NewParagraph();

			// Oftenly, <pre> tag are used to renders some code examples. They look better inside a table
            if (this.RenderPreAsTable)
            {
                Table currentTable = new Table(
                    new TableProperties (
                        new TableStyle() { Val = htmlStyles.GetStyle("TableGrid", StyleValues.Table) },
                        new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" } // 100% * 50
					),
                    new TableGrid(
                        new GridColumn() { Width = "5610" }),
                    new TableRow(
                        new TableCell(
                    // Ensure the border lines are visible (regardless of the style used)
                            new TableCellProperties
                            {
                                TableCellBorders = new TableCellBorders(
                                   new TopBorder() { Val = BorderValues.Single },
                                   new LeftBorder() { Val = BorderValues.Single },
                                   new BottomBorder() { Val = BorderValues.Single },
                                   new RightBorder() { Val = BorderValues.Single })
                            },
                            currentParagraph))
                );

                AddParagraph(currentTable);
                tables.NewContext(currentTable);
            }
            else
            {
                AddParagraph(currentParagraph);
            }

			// Process the entire <pre> tag and append it to the document
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
            { 
				htmlStyles.Runs.BeginTag(en.CurrentTag, styleAttributes.ToArray());
            }

            AlternateProcessHtmlChunks(en, "</pre>");

			if (styleAttributes.Count > 0)
            { 
				htmlStyles.Runs.EndTag(en.CurrentTag);
            }

            if (RenderPreAsTable)
            { 
				tables.CloseContext();
            }
            CompleteCurrentParagraph();
		}

		#endregion

		#region ProcessQuote

		private void ProcessQuote(HtmlEnumerator en)
		{
			// The browsers render the quote tag between a kind of separators.
			// We add the Quote style to the nested runs to match more Word.

			Run run = new Run(
				new Text(" " + HtmlStyles.QuoteCharacters.chars[0]) { Space = SpaceProcessingModeValues.Preserve }
			);

			htmlStyles.Runs.ApplyTags(run);
			elements.Add(run);

			ProcessHtmlElement<RunStyle>(en, new RunStyle() { Val = htmlStyles.GetStyle("Quote", StyleValues.Character) });
		}

		#endregion

		#region ProcessSpan

		private void ProcessSpan(HtmlEnumerator en)
		{
			// A span style attribute can contains many information: font color, background color, font size,
			// font family, ...
			// We'll check for each of these and add apply them to the next build runs.

			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			bool newParagraph = ProcessContainerAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
            { 
				htmlStyles.Runs.MergeTag(en.CurrentTag, styleAttributes);
            }

            if (newParagraph)
			{
				AlternateProcessHtmlChunks(en, en.ClosingCurrentTag);
				CompleteCurrentParagraph(true);
			}
		}

		#endregion

		#region ProcessSubscript

		private void ProcessSubscript(HtmlEnumerator en)
		{
			ProcessHtmlElement<VerticalTextAlignment>(en, new VerticalTextAlignment() { Val = VerticalPositionValues.Subscript });
		}

		#endregion

		#region ProcessSuperscript

		private void ProcessSuperscript(HtmlEnumerator en)
		{
			ProcessHtmlElement<VerticalTextAlignment>(en, new VerticalTextAlignment() { Val = VerticalPositionValues.Superscript });
		}

		#endregion

		#region ProcessUnderline

		private void ProcessUnderline(HtmlEnumerator en)
		{
			ProcessHtmlElement<Underline>(en, new Underline() { Val = UnderlineValues.Single });
		}

		#endregion

		#region ProcessTable

		private void ProcessTable(HtmlEnumerator en)
		{
			TableProperties properties = new TableProperties(
				new TableStyle() { Val = htmlStyles.GetStyle("TableGrid", StyleValues.Table) }
			);

            if(en.Attributes["id"] == "twoColumn")
            {
                properties.TableCaption = new TableCaption() { Val = "TwoColumn" };
            }

            if (en.Attributes["id"] == "xCelHeader")
            {
                properties.TableCaption = new TableCaption() { Val = "XCelHeader" };
            }

            Table currentTable = new Table(properties);

			string classValue = en.Attributes["class"];
			if (classValue != null)
			{
				classValue = htmlStyles.GetStyle(classValue, StyleValues.Table, ignoreCase: true);
				if (classValue != null)
                { 
					properties.TableStyle.Val = classValue;
                }
            }

			int? border = en.Attributes.GetAsInt("border");
			if (border.HasValue && border.Value > 0)
			{
				bool handleBorders = true;
				if (classValue != null)
				{
					// check whether the style in use have borders
                    string styleId = this.htmlStyles.GetStyle(classValue, StyleValues.Table, true);
					if (styleId != null)
                    {
                        var s = mainPart.StyleDefinitionsPart.Styles.Elements<Style>().First(e => e.StyleId == styleId);
                        if (s.StyleTableProperties.TableBorders != null)
                        {
                            handleBorders = false;
                        }
                    }
				}

				// If the border has been specified, we display the Table Grid style which display
				// its grid lines. Otherwise the default table style hides the grid lines.
				if (handleBorders && properties.TableStyle.Val != "TableGrid")
				{
					uint borderSize = border.Value > 1 ? (uint) new Unit(UnitMetric.Pixel, border.Value).ValueInDxa : 1;
					properties.TableBorders = new TableBorders() {
						TopBorder = new TopBorder { Val = BorderValues.None },
						LeftBorder = new LeftBorder { Val = BorderValues.None },
						RightBorder = new RightBorder { Val = BorderValues.None },
						BottomBorder = new BottomBorder { Val = BorderValues.None },
						InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.Single, Size = borderSize },
						InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.Single, Size = borderSize }
					};
				}
			}
			// is the border=0? If so, we remove the border regardless the style in use
			else if (border == 0)
			{
				properties.TableBorders = new TableBorders() {
					TopBorder = new TopBorder { Val = BorderValues.None },
					LeftBorder = new LeftBorder { Val = BorderValues.None },
					RightBorder = new RightBorder { Val = BorderValues.None },
					BottomBorder = new BottomBorder { Val = BorderValues.None },
					InsideHorizontalBorder = new InsideHorizontalBorder { Val = BorderValues.None },
					InsideVerticalBorder = new InsideVerticalBorder { Val = BorderValues.None }
				};
			}

			Unit unit = en.StyleAttributes.GetAsUnit("width");
            if (!unit.IsValid)
            {
                unit = en.Attributes.GetAsUnit("width");
            }

			if (unit.IsValid)
			{
				switch (unit.Type)
				{
					case UnitMetric.Percent:
						properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) }; break;
					case UnitMetric.Point:
						properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.ValueInDxa * 20).ToString(CultureInfo.InvariantCulture) }; break;
					case UnitMetric.Pixel:
						properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.ValueInDxa).ToString(CultureInfo.InvariantCulture) }; break;
					//case UnitMetric.Percent:
					//	properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" }; break;
					//case UnitMetric.Point:
					//	properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = "5000" }; break;
					//case UnitMetric.Pixel:
					//	properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Dxa, Width = "5000" }; break;
				}
			}
			else
			{
				// Use Auto=0 instead of Pct=auto
				// bug reported by scarhand (https://html2openxml.codeplex.com/workitem/12494)
				properties.TableWidth = new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" };//case 1823 Table Width respect to page width
                
                // for printing table content wisw use below code
                // Type = TableWidthUnitValues.Auto, Width = "0"
			}

			string align = en.Attributes["align"];
			if (align != null)
			{
				JustificationValues? halign = Converter.ToParagraphAlign(align);
				if (halign.HasValue)
                { 
					properties.TableJustification = new TableJustification() { Val = halign.Value.ToTableRowAlignment() };
                }
            }

			// only if the table is left aligned, we can handle some left margin indentation
			// Right margin + Right align has no equivalent in OpenXml
			if (align == null || align == "left")
			{
				Margin margin = en.StyleAttributes.GetAsMargin("margin");

				// OpenXml doesn't support table margin in Percent, but Html does
				// the margin part has been implemented by Olek (patch #8457)

				TableCellMarginDefault cellMargin = new TableCellMarginDefault();
                if (margin.Left.IsFixed)
                { 
					cellMargin.TableCellLeftMargin = new TableCellLeftMargin() { Type = TableWidthValues.Dxa, Width = (short) margin.Left.ValueInDxa };
                }
                if (margin.Right.IsFixed)
                { 
					cellMargin.TableCellRightMargin = new TableCellRightMargin() { Type = TableWidthValues.Dxa, Width = (short) margin.Right.ValueInDxa };
                }
                if (margin.Top.IsFixed)
                { 
					cellMargin.TopMargin = new TopMargin() { Type = TableWidthUnitValues.Dxa, Width = margin.Top.ValueInDxa.ToString(CultureInfo.InvariantCulture) };
                }
                if (margin.Bottom.IsFixed)
                { 
					cellMargin.BottomMargin = new BottomMargin() { Type = TableWidthUnitValues.Dxa, Width = margin.Bottom.ValueInDxa.ToString(CultureInfo.InvariantCulture) };
                }
                // Align table according to the margin 'auto' as it stands in Html
                if (margin.Left.Type == UnitMetric.Auto || margin.Right.Type == UnitMetric.Auto)
                {
                    TableRowAlignmentValues justification;

                    if (margin.Left.Type == UnitMetric.Auto && margin.Right.Type == UnitMetric.Auto)
                    { 
                        justification = TableRowAlignmentValues.Center;
                    }
                    else if (margin.Left.Type == UnitMetric.Auto)
                    { 
                        justification = TableRowAlignmentValues.Right;
                    }
                    else
                    {
                        justification = TableRowAlignmentValues.Left;
                    }                       

                    properties.TableJustification = new TableJustification() { Val = justification };
                }

				if (cellMargin.HasChildren)
                { 
					properties.TableCellMarginDefault = cellMargin;
                }
            }

			int? spacing = en.Attributes.GetAsInt("cellspacing");
			if (spacing.HasValue)
            { 
                properties.TableCellSpacing = new TableCellSpacing { Type = TableWidthUnitValues.Dxa, Width = new Unit(UnitMetric.Pixel, spacing.Value).ValueInDxa.ToString(CultureInfo.InvariantCulture) };
            }

            int? padding = en.Attributes.GetAsInt("cellpadding");
            if (padding.HasValue)
            {
                int paddingDxa = (int) new Unit(UnitMetric.Pixel, padding.Value).ValueInDxa;

                TableCellMarginDefault cellMargin = new TableCellMarginDefault();
                cellMargin.TableCellLeftMargin = new TableCellLeftMargin() { Type = TableWidthValues.Dxa, Width = (short) paddingDxa };
                cellMargin.TableCellRightMargin = new TableCellRightMargin() { Type = TableWidthValues.Dxa, Width = (short) paddingDxa };
                cellMargin.TopMargin = new TopMargin() { Type = TableWidthUnitValues.Dxa, Width = paddingDxa.ToString(CultureInfo.InvariantCulture) };
                cellMargin.BottomMargin = new BottomMargin() { Type = TableWidthUnitValues.Dxa, Width = paddingDxa.ToString(CultureInfo.InvariantCulture) };
                properties.TableCellMarginDefault = cellMargin;
            }

			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();
			htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);

            if (runStyleAttributes.Count > 0)
            { 
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());
            }

            // are we currently inside another table?
            if (tables.HasContext)
			{
				// Okay we will insert nested table but beware the paragraph inside TableCell should contains at least 1 run.

				TableCell currentCell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();
                // don't add an empty paragraph if not required (bug #13608 by zanjo)
                if (elements.Count == 0)
                {
                    currentCell.Append(currentTable);
                }
                else
                {
                    currentCell.Append(new Paragraph(elements), currentTable);
                    elements.Clear();
                }
			}
			else
			{
				CompleteCurrentParagraph();
				this.paragraphs.Add(currentTable);
			}

			tables.NewContext(currentTable);
		}

		#endregion

		#region ProcessTableCaption

		private void ProcessTableCaption(HtmlEnumerator en)
		{
            if (!tables.HasContext)
            {
                return;
            }

			string att = en.StyleAttributes["text-align"];
            if (att == null)
            {
                att = en.Attributes["align"];
            }

			ProcessHtmlChunks(en, "</caption>");

			var legend = new Paragraph(
					new ParagraphProperties {
						ParagraphStyleId = new ParagraphStyleId() { Val = htmlStyles.GetStyle("Caption", StyleValues.Paragraph) }
					},
					new Run(
						new FieldChar() { FieldCharType = FieldCharValues.Begin }),
					new Run(
						new FieldCode(" SEQ TABLE \\* ARABIC ") { Space = SpaceProcessingModeValues.Preserve }),
					new Run(
						new FieldChar() { FieldCharType = FieldCharValues.End })
				);

			legend.Append(elements);
			elements.Clear();

			if (att != null)
			{
				JustificationValues? align = Converter.ToParagraphAlign(att);
				if (align.HasValue)
                { 
					legend.InsertInProperties(prop => prop.Justification = new Justification { Val = align } );
                }
            }
			else
			{
				// If no particular alignement has been specified for the legend, we will align the legend
				// relative to the owning table
				TableProperties props = tables.CurrentTable.GetFirstChild<TableProperties>();
				if (props != null)
				{
					TableJustification justif = props.GetFirstChild<TableJustification>();
                    if (justif != null)
                    {
                        legend.InsertInProperties(prop =>
                          prop.Justification = new Justification { Val = justif.Val.Value.ToJustification() });
                    }
				}
			}

			if (this.TableCaptionPosition == CaptionPositionValues.Above)
            { 
				this.paragraphs.Insert(this.paragraphs.Count - 1, legend);
            }
            else
            { 
				this.paragraphs.Add(legend);
            }
        }

		#endregion

		#region ProcessTableRow

		private void ProcessTableRow(HtmlEnumerator en)
		{
			// in case the html is bad-formed and use <tr> outside a <table> tag, we will ensure
			// a table context exists.
			if (!tables.HasContext) return;

			TableRowProperties properties = new TableRowProperties();
			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();

			htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);
            
			Unit unit = en.StyleAttributes.GetAsUnit("height");
            if (!unit.IsValid)
            {
                unit = en.Attributes.GetAsUnit("height");
            }

			switch (unit.Type)
			{
				case UnitMetric.Point:
					properties.Append(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) (unit.Value * 20) });
					break;
				case UnitMetric.Pixel:
					properties.Append(new TableRowHeight() { HeightType = HeightRuleValues.AtLeast, Val = (uint) unit.ValueInDxa });
					break;
			}

			TableRow row = new TableRow();
            if (properties.HasChildren)
            {
                row.AppendChild(properties);
            }

			htmlStyles.Runs.ProcessCommonAttributes(en, runStyleAttributes);
			if (runStyleAttributes.Count > 0)
            { 
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());
            }

            tables.CurrentTable.Append(row);
			tables.CellPosition = new CellPosition(tables.CellPosition.Row + 1, 0);
		}

		#endregion

		#region ProcessTableColumn

		private void ProcessTableColumn(HtmlEnumerator en)
		{
            if (!tables.HasContext)
            {
                return;
            }

			TableCellProperties properties = new TableCellProperties();
			// in Html, table cell are vertically centered by default
			properties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };

			HtmlBorder attrBorder = en.StyleAttributes.GetAsBorder("border");
			if (!attrBorder.IsEmpty)
			{
				properties.TableCellBorders = new TableCellBorders()
				{
					TopBorder = new TopBorder { Val = attrBorder.Top.Width.Value > 0 ? BorderValues.Single : BorderValues.None, Size = (uint)attrBorder.Top.Width.ValueInDxa, Color = attrBorder.Top.Color.ToHexString() },
					LeftBorder = new LeftBorder { Val = attrBorder.Left.Width.Value > 0 ? BorderValues.Single : BorderValues.None, Size = (uint)attrBorder.Left.Width.ValueInDxa, Color = attrBorder.Left.Color.ToHexString() },
					RightBorder = new RightBorder { Val = attrBorder.Right.Width.Value > 0 ? BorderValues.Single : BorderValues.None, Size = (uint)attrBorder.Right.Width.ValueInDxa, Color = attrBorder.Right.Color.ToHexString() },
					BottomBorder = new BottomBorder { Val = attrBorder.Bottom.Width.Value > 0 ? BorderValues.Single : BorderValues.None, Size = (uint)attrBorder.Bottom.Width.ValueInDxa, Color = attrBorder.Bottom.Color.ToHexString() },
				};
			}

			string attrColorValue = en.StyleAttributes["background-color"];
			if(attrColorValue != null)
			{
				HexBinaryValue hexBinaryColor = (HexBinaryValue)attrColorValue;
				properties.Shading = new Shading()
				{
					Fill = hexBinaryColor.ToString()
				};
			}
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();
			List<OpenXmlElement> runStyleAttributes = new List<OpenXmlElement>();

			Unit unit = en.StyleAttributes.GetAsUnit("width");
            if (!unit.IsValid)
            {
                unit = en.Attributes.GetAsUnit("width");
            }

			switch (unit.Type)
			{
				case UnitMetric.Percent:
                    properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Pct, Width = (unit.Value * 50).ToString(CultureInfo.InvariantCulture) };
					break;
				case UnitMetric.Point:
                    properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.ValueInDxa * 20).ToString(CultureInfo.InvariantCulture) };
					break;
				case UnitMetric.Pixel:
					properties.TableCellWidth = new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = (unit.ValueInDxa).ToString(CultureInfo.InvariantCulture) };
					break;
			}

			// fix an issue when specifying the RowSpan or ColSpan=1 (reported by imagremlin)
			int? colspan = en.Attributes.GetAsInt("colspan");
			if (colspan.HasValue && colspan.Value > 1)
			{
				properties.GridSpan = new GridSpan() { Val = colspan };
			}

			int? rowspan = en.Attributes.GetAsInt("rowspan");
			if (rowspan.HasValue && rowspan.Value > 1)
			{
				properties.VerticalMerge = new VerticalMerge() { Val = MergedCellValues.Restart };

				var p = tables.CellPosition;
                int shift = 0;
                // if there is already a running rowSpan on a left-sided column, we have to shift this position
                foreach (var rs in tables.RowSpan)
                {
                    if (rs.CellOrigin.Row < p.Row && rs.CellOrigin.Column <= p.Column + shift)
                    {
                        shift++;
                    }
                }

                p.Offset(0, shift);
                tables.RowSpan.Add(new HtmlTableSpan(p) {
                    RowSpan = rowspan.Value - 1,
                    ColSpan = colspan.HasValue && rowspan.Value > 1 ? colspan.Value : 0
                });
			}

			// Manage vertical text (only for table cell)
			string direction = en.StyleAttributes["writing-mode"];
			if (direction != null)
			{
				switch (direction)
				{
					case "tb-lr":
						properties.TextDirection = new TextDirection() { Val = TextDirectionValues.BottomToTopLeftToRight };
						properties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
						htmlStyles.Tables.BeginTagForParagraph(en.CurrentTag, new Justification() { Val = JustificationValues.Center });
						break;
					case "tb-rl":
						properties.TextDirection = new TextDirection() { Val = TextDirectionValues.TopToBottomRightToLeft };
						properties.TableCellVerticalAlignment = new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center };
						htmlStyles.Tables.BeginTagForParagraph(en.CurrentTag, new Justification() { Val = JustificationValues.Center });
						break;
				}
			}

			var padding = en.StyleAttributes.GetAsMargin("padding");
			if (!padding.IsEmpty)
			{
				TableCellMargin cellMargin = new TableCellMargin();
				var cellMarginSide = new List<KeyValuePair<Unit, TableWidthType>>();
				cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Top, new TopMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Right, new RightMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Bottom, new BottomMargin()));
				cellMarginSide.Add(new KeyValuePair<Unit, TableWidthType>(padding.Left, new LeftMargin()));

				foreach (var pair in cellMarginSide)
				{
					if (!pair.Key.IsValid || pair.Key.Value == 0) continue;
					if (pair.Key.Type == UnitMetric.Percent)
					{
						pair.Value.Width = (pair.Key.Value * 50).ToString(CultureInfo.InvariantCulture);
						pair.Value.Type = TableWidthUnitValues.Pct;
					}
					else
					{
						pair.Value.Width = pair.Key.ValueInDxa.ToString(CultureInfo.InvariantCulture);
						pair.Value.Type = TableWidthUnitValues.Dxa;
					}

					cellMargin.Append(pair.Value);
				}

				properties.TableCellMargin = cellMargin;
			}

			htmlStyles.Tables.ProcessCommonAttributes(en, runStyleAttributes);
			if (styleAttributes.Count > 0)
            { 
				htmlStyles.Tables.BeginTag(en.CurrentTag, styleAttributes);
            }

            if (runStyleAttributes.Count > 0)
            { 
				htmlStyles.Runs.BeginTag(en.CurrentTag, runStyleAttributes.ToArray());
            }

            TableCell cell = new TableCell();
            if (properties.HasChildren)
            {
                cell.TableCellProperties = properties;
            }

			tables.CurrentTable.GetLastChild<TableRow>().Append(cell);

			if (en.IsSelfClosedTag) // Force a call to ProcessClosingTableColumn
            { 
				ProcessClosingTableColumn(en);
            }
            else
			{
				// we create a new currentParagraph to add new runs inside the TableCell
				cell.Append(currentParagraph = new Paragraph());
			}
		}

		#endregion

		#region ProcessTablePart

		private void ProcessTablePart(HtmlEnumerator en)
		{
			List<OpenXmlElement> styleAttributes = new List<OpenXmlElement>();

			htmlStyles.Tables.ProcessCommonAttributes(en, styleAttributes);

			if (styleAttributes.Count > 0)
            { 
				htmlStyles.Tables.BeginTag(en.CurrentTag, styleAttributes.ToArray());
            }
        }

		#endregion

		#region ProcessXmlDataIsland

		private void ProcessXmlDataIsland(HtmlEnumerator en)
		{
			// Process inner Xml data island and do nothing.
			// The Xml has this format:
			/* <?xml:namespace prefix=o ns="urn:schemas-microsoft-com:office:office">
			   <globalGuideLine>
				   <employee>
					  <FirstName>Austin</FirstName>
					  <LastName>Hennery</LastName>
				   </employee>
			   </globalGuideLine>
			 */

			// Move to the first root element of the Xml then process until the end of the xml chunks.
			while (en.MoveNext() && !en.IsCurrentHtmlTag) ;

			if (en.Current != null)
			{
				string xmlRootElement = en.ClosingCurrentTag;
				while (en.MoveUntilMatch(xmlRootElement)) ;
			}
		}

        #endregion

        #region ProcessLine

        private void ProcessLine(HtmlEnumerator en)
        {

            string attrWeightValue = en.Attributes["thickness"];
            string attrColorValue = en.Attributes["color"];
            decimal decimalWeight = Convert.ToDecimal(attrWeightValue) * 12700;  //1 pt = 12,700
            int attrWeight = Convert.ToInt32(decimalWeight);
            HexBinaryValue hexBinaryColor = (HexBinaryValue)attrColorValue;

            Random random = new Random();
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            string randomValue = string.Empty;
            randomValue = new string(Enumerable.Repeat(chars, 8)
                                    .Select(s => s[random.Next(s.Length)]).ToArray());

            int rndDigit = random.Next(1, 99999);
            if (uintDigit == 1)
            { 
                uintDigit = Convert.ToUInt32(rndDigit);
            }
            else
            {
                uintDigit++;
            }
                
            Run run = new Run();
            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            AlternateContent alternateContent1 = new AlternateContent();

            AlternateContentChoice alternateContentChoice1 = new AlternateContentChoice() { Requires = "wps" };

            Drawing drawing1 = new Drawing();

            Wp.Anchor anchor1 = new Wp.Anchor() { DistanceFromTop = (UInt32Value)0U, DistanceFromBottom = (UInt32Value)0U, DistanceFromLeft = (UInt32Value)114300U, DistanceFromRight = (UInt32Value)114300U, SimplePos = false, RelativeHeight = (UInt32Value)251657216U, BehindDoc = false, Locked = false, LayoutInCell = true, AllowOverlap = true }; // EditId = "589E06C1", AnchorId = "3D6A5C96" };
            Wp.SimplePosition simplePosition1 = new Wp.SimplePosition() { X = 0L, Y = 0L };

            Wp.HorizontalPosition horizontalPosition1 = new Wp.HorizontalPosition() { RelativeFrom = Wp.HorizontalRelativePositionValues.Column };
            Wp.PositionOffset positionOffset1 = new Wp.PositionOffset();
            positionOffset1.Text = "13335";

            horizontalPosition1.Append(positionOffset1);

            Wp.VerticalPosition verticalPosition1 = new Wp.VerticalPosition() { RelativeFrom = Wp.VerticalRelativePositionValues.Paragraph };
            Wp.PositionOffset positionOffset2 = new Wp.PositionOffset();
            positionOffset2.Text = "88265";

            verticalPosition1.Append(positionOffset2);
            Wp.Extent extent1 = new Wp.Extent() { Cx = 5410200L, Cy = 19050L };
            Wp.EffectExtent effectExtent1 = new Wp.EffectExtent() { LeftEdge = 19050L, TopEdge = 19050L, RightEdge = 0L, BottomEdge = 0L };
            Wp.WrapNone wrapNone1 = new Wp.WrapNone();
            Wp.DocProperties docProperties1 = new Wp.DocProperties() { Id = UInt32Value.FromUInt32(uintDigit), Name = randomValue };

            Wp.NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties1 = new Wp.NonVisualGraphicFrameDrawingProperties();

            A.GraphicFrameLocks graphicFrameLocks1 = new A.GraphicFrameLocks();
            graphicFrameLocks1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            nonVisualGraphicFrameDrawingProperties1.Append(graphicFrameLocks1);

            A.Graphic graphic1 = new A.Graphic();
            graphic1.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.GraphicData graphicData1 = new A.GraphicData() { Uri = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape" };

            Wps.WordprocessingShape wordprocessingShape1 = new Wps.WordprocessingShape();

            Wps.NonVisualConnectorProperties nonVisualConnectorProperties1 = new Wps.NonVisualConnectorProperties();
            A.ConnectionShapeLocks connectionShapeLocks1 = new A.ConnectionShapeLocks();

            nonVisualConnectorProperties1.Append(connectionShapeLocks1);

            Wps.ShapeProperties shapeProperties1 = new Wps.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D() { VerticalFlip = true };
            A.Offset offset1 = new A.Offset() { X = 0L, Y = 0L };
            A.Extents extents1 = new A.Extents() { Cx = 5410200L, Cy = 19050L };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Line };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);
            A.NoFill noFill1 = new A.NoFill();

            A.Outline outline1 = new A.Outline() { Width = attrWeight, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };
            //outline1.Width = ;

            A.SolidFill solidFill1 = new A.SolidFill();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = hexBinaryColor };

            solidFill1.Append(rgbColorModelHex1);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill1);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);
            shapeProperties1.Append(noFill1);
            shapeProperties1.Append(outline1);
            Wps.TextBodyProperties textBodyProperties1 = new Wps.TextBodyProperties();

            wordprocessingShape1.Append(nonVisualConnectorProperties1);
            wordprocessingShape1.Append(shapeProperties1);
            wordprocessingShape1.Append(textBodyProperties1);

            graphicData1.Append(wordprocessingShape1);

            graphic1.Append(graphicData1);

            Wp14.RelativeWidth relativeWidth1 = new Wp14.RelativeWidth() { ObjectId = Wp14.SizeRelativeHorizontallyValues.Page };
            Wp14.PercentageWidth percentageWidth1 = new Wp14.PercentageWidth();
            percentageWidth1.Text = "0";

            relativeWidth1.Append(percentageWidth1);

            Wp14.RelativeHeight relativeHeight1 = new Wp14.RelativeHeight() { RelativeFrom = Wp14.SizeRelativeVerticallyValues.Page };
            Wp14.PercentageHeight percentageHeight1 = new Wp14.PercentageHeight();
            percentageHeight1.Text = "0";

            relativeHeight1.Append(percentageHeight1);

            anchor1.Append(simplePosition1);
            anchor1.Append(horizontalPosition1);
            anchor1.Append(verticalPosition1);
            anchor1.Append(extent1);
            anchor1.Append(effectExtent1);
            anchor1.Append(wrapNone1);
            anchor1.Append(docProperties1);
            anchor1.Append(nonVisualGraphicFrameDrawingProperties1);
            anchor1.Append(graphic1);
            anchor1.Append(relativeWidth1);
            anchor1.Append(relativeHeight1);

            drawing1.Append(anchor1);

            alternateContentChoice1.Append(drawing1);

            AlternateContentFallback alternateContentFallback1 = new AlternateContentFallback();

            Picture picture1 = new Picture();

            V.Line line1 = new V.Line()
            {
                Id = randomValue,
                Style = "position:absolute;flip:y;z-index:251657216;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:0;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:0;mso-position-horizontal:absolute;mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:page;mso-height-relative:page",
                OptionalString = "_x0000_s1026",
                StrokeColor = attrColorValue,
                StrokeWeight = attrWeightValue,
                From = "1.05pt,6.95pt",
                To = "427.05pt,8.45pt"
            };
            //line1.SetAttribute(new OpenXmlAttribute("w14", "anchorId", "http://schemas.microsoft.com/office/word/2010/wordml", "72D0737C"));
            line1.SetAttribute(new OpenXmlAttribute("o", "gfxdata", "urn:schemas-microsoft-com:office:office", "UEsDBBQABgAIAAAAIQC2gziS/gAAAOEBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRQU7DMBBF\n90jcwfIWJU67QAgl6YK0S0CoHGBkTxKLZGx5TGhvj5O2G0SRWNoz/78nu9wcxkFMGNg6quQqL6RA\n0s5Y6ir5vt9lD1JwBDIwOMJKHpHlpr69KfdHjyxSmriSfYz+USnWPY7AufNIadK6MEJMx9ApD/oD\nOlTrorhX2lFEilmcO2RdNtjC5xDF9pCuTyYBB5bi6bQ4syoJ3g9WQ0ymaiLzg5KdCXlKLjvcW893\nSUOqXwnz5DrgnHtJTxOsQfEKIT7DmDSUCaxw7Rqn8787ZsmRM9e2VmPeBN4uqYvTtW7jvijg9N/y\nJsXecLq0q+WD6m8AAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAALAAAAX3JlbHMvLnJl\nbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrbUb/Q94l/f/hMi1qR\nJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG5lrLq9biZkxWOiqY\n22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nTNEV3j6o9feQzro1i\nOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMAUEsDBBQABgAIAAAA\nIQDb/Ic92wEAAJEDAAAOAAAAZHJzL2Uyb0RvYy54bWysU02P0zAQvSPxHyzfaZJdFkrUdA8Ny2UF\nlQrcp46dWPhLHtO0/56x2+2ycEP0YNUz4zfz3rys7o/WsIOMqL3reLOoOZNO+EG7sePfvj68WXKG\nCdwAxjvZ8ZNEfr9+/Wo1h1be+MmbQUZGIA7bOXR8Sim0VYVikhZw4YN0lFQ+Wkh0jWM1RJgJ3Zrq\npq7fVbOPQ4heSESK9uckXxd8paRIX5RCmZjpOM2WyhnLuc9ntV5BO0YIkxaXMeAfprCgHTW9QvWQ\ngP2M+i8oq0X06FVaCG8rr5QWsnAgNk39B5vdBEEWLiQOhqtM+P9gxefDNjI9dPyWMweWVrRLEfQ4\nJbbxzpGAPrIm6zQHbKl847YxMxVHtwuPXvxAylUvkvmC4Vx2VNEyZXT4TvYoEhFpdiwbOF03II+J\nCQrevW1qWitngnLNh/qubKiCNsPkriFi+iS9ZflPx412WSBo4fCIKQ/yXJLDzj9oY8qSjWMzsVw2\nBR7Ia8pAok42EHt0I2dgRjKxSLFAojd6yM8zEMZxvzGRHYCM9LF/398WTajdi7I8VQ84netK6mwx\nqxP53Gjb8WWdfzlMr427iHfWKyu398NpG59Epb2XwotHs7F+v5fXz1/S+hcAAAD//wMAUEsDBBQA\nBgAIAAAAIQDrclHM2wAAAAcBAAAPAAAAZHJzL2Rvd25yZXYueG1sTI7BTsMwEETvSPyDtUjcqNMW\n0iTEqQBR0RMShQs3116SCHsdxW6b/j3LCY77ZjT76vXknTjiGPtACuazDASSCbanVsHH++amABGT\nJqtdIFRwxgjr5vKi1pUNJ3rD4y61gkcoVlpBl9JQSRlNh17HWRiQOPsKo9eJz7GVdtQnHvdOLrIs\nl173xB86PeBTh+Z7d/AKPovhvMJ8tXwMptz6/HVjnl+cUtdX08M9iIRT+ivDrz6rQ8NO+3AgG4VT\nsJhzkfGyBMFxcXfLYM8gL0E2tfzv3/wAAAD//wMAUEsBAi0AFAAGAAgAAAAhALaDOJL+AAAA4QEA\nABMAAAAAAAAAAAAAAAAAAAAAAFtDb250ZW50X1R5cGVzXS54bWxQSwECLQAUAAYACAAAACEAOP0h\n/9YAAACUAQAACwAAAAAAAAAAAAAAAAAvAQAAX3JlbHMvLnJlbHNQSwECLQAUAAYACAAAACEA2/yH\nPdsBAACRAwAADgAAAAAAAAAAAAAAAAAuAgAAZHJzL2Uyb0RvYy54bWxQSwECLQAUAAYACAAAACEA\n63JRzNsAAAAHAQAADwAAAAAAAAAAAAAAAAA1BAAAZHJzL2Rvd25yZXYueG1sUEsFBgAAAAAEAAQA\n8wAAAD0FAAAAAA==\n"));
            V.Stroke stroke1 = new V.Stroke() { JoinStyle = V.StrokeJoinStyleValues.Miter };
            Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, ShapeType = false };

            line1.Append(stroke1);
            line1.Append(lock1);

            picture1.Append(line1);

            alternateContentFallback1.Append(picture1);

            alternateContent1.Append(alternateContentChoice1);
            alternateContent1.Append(alternateContentFallback1);

            run.Append(runProperties1);
            run.Append(alternateContent1);

            elements.Add(run);
        }

        #endregion
        // Closing tags

        #region ProcessClosingBlockQuote

        private void ProcessClosingBlockQuote(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);
			htmlStyles.Paragraph.EndTag("<blockquote>");
		}

		#endregion

		#region ProcessClosingDiv

		private void ProcessClosingDiv(HtmlEnumerator en)
		{
			// Mimic the rendering of the browser:
            // TODO: BUIL-4851: we don't want break-space after a div element is closed (ie, a data widget)
			// ProcessBr(en);
			ProcessClosingTag(en);
		}

		#endregion

		#region ProcessClosingTag

		private void ProcessClosingTag(HtmlEnumerator en)
		{
			string openingTag = en.CurrentTag.Replace("/", "");
			htmlStyles.Runs.EndTag(openingTag);
			htmlStyles.Paragraph.EndTag(openingTag);
		}

		#endregion

		#region ProcessClosingNumberingList

		private void ProcessClosingNumberingList(HtmlEnumerator en)
		{
			htmlStyles.NumberingList.EndList();

			// If we are no more inside a list, we move to another paragraph (as we created
			// one for containing all the <li>. This will ensure the next run will not be added to the <li>.
			if (htmlStyles.NumberingList.LevelIndex == 0)
            { 
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
            }
        }

		#endregion

		#region ProcessClosingParagraph

		private void ProcessClosingParagraph(HtmlEnumerator en)
		{
			CompleteCurrentParagraph(true);

			string tag = en.CurrentTag.Replace("/", "");
			htmlStyles.Runs.EndTag(tag);
			htmlStyles.Paragraph.EndTag(tag);
		}

		#endregion

		#region ProcessClosingQuote

		private void ProcessClosingQuote(HtmlEnumerator en)
		{
			Run run = new Run(
				new Text(HtmlStyles.QuoteCharacters.chars[1]) { Space = SpaceProcessingModeValues.Preserve }
			);
			htmlStyles.Runs.ApplyTags(run);
			elements.Add(run);

			htmlStyles.Runs.EndTag("<q>");
		}

		#endregion

		#region ProcessClosingTable

		private void ProcessClosingTable(HtmlEnumerator en)
		{

			htmlStyles.Tables.EndTag("<table>");
			htmlStyles.Runs.EndTag("<table>");

			TableRow row = tables.CurrentTable.GetFirstChild<TableRow>();
			// Is this a misformed or empty table?
			if (row != null)
			{
				// Count the number of tableCell and add as much GridColumn as we need.
				TableGrid grid = new TableGrid();
				foreach (TableCell cell in row.Elements<TableCell>())
				{
					// If that column contains some span, we need to count them also
					int count = cell.TableCellProperties?.GridSpan?.Val ?? 1;
					for (int i=0; i<count; i++)
                    {
						grid.Append(new GridColumn());
					}
				}

				tables.CurrentTable.InsertAt<TableGrid>(grid, 1);
			}

			tables.CloseContext();

			if (!tables.HasContext)
            { 
				AddParagraph(currentParagraph = htmlStyles.Paragraph.NewParagraph());
            }
        }

		#endregion

		#region ProcessClosingTablePart

		private void ProcessClosingTablePart(HtmlEnumerator en)
		{
			string closingTag = en.CurrentTag.Replace("/", "");
			htmlStyles.Tables.EndTag(closingTag);
		}

		#endregion

		#region ProcessClosingTableRow

		private void ProcessClosingTableRow(HtmlEnumerator en)
		{
            if (!tables.HasContext)
            {
                return;
            }

			TableRow row = tables.CurrentTable.GetLastChild<TableRow>();
            if (row == null)
            {
                return;
            }

			// Word will not open documents with empty rows (reported by scwebgroup)
			if (row.GetFirstChild<TableCell>() == null)
			{
				row.Remove();
				return;
			}

			// Add empty columns to fill rowspan
			if (tables.RowSpan.Count > 0)
			{
				int rowIndex = tables.CellPosition.Row;

				for (int i = 0; i < tables.RowSpan.Count; i++)
				{
					HtmlTableSpan tspan = tables.RowSpan[i];
                    if (tspan.CellOrigin.Row == rowIndex)
                    {
                        continue;
                    }

                    TableCell emptyCell = new TableCell(new TableCellProperties {
								            TableCellWidth = new TableCellWidth() { Width = "0" },
								            VerticalMerge = new VerticalMerge() },
							            new Paragraph());

                    tspan.RowSpan--;

                    if (tspan.RowSpan == 0)
                    {
                        tables.RowSpan.RemoveAt(i);
                        i--;
                    }

                    // in case of both colSpan + rowSpan on the same cell, we have to reverberate the rowSpan on the next columns too
                    if (tspan.ColSpan > 0)
                    {
                        emptyCell.TableCellProperties.GridSpan = new GridSpan() { Val = tspan.ColSpan };
                    }

                    TableCell cell = row.GetFirstChild<TableCell>();
                    if (tspan.CellOrigin.Column == 0 || cell == null)
                    {
                        row.InsertAt(emptyCell, 0);
                        continue;
                    }

                    // find the good column position, taking care of eventual colSpan
                    int columnIndex = 0;
                    while (columnIndex < tspan.CellOrigin.Column)
                    {
                        columnIndex += cell.TableCellProperties?.GridSpan?.Val ?? 1;
                    }
                    //while ((cell = cell.NextSibling<TableCell>()) != null);

                    if (cell == null)
                    {
                        row.AppendChild(emptyCell);
                    }
                    else
                    {
                        row.InsertAfter<TableCell>(emptyCell, cell);
                    }
                }
			}

			htmlStyles.Tables.EndTag("<tr>");
			htmlStyles.Runs.EndTag("<tr>");
		}

		#endregion

		#region ProcessClosingTableColumn

		private void ProcessClosingTableColumn(HtmlEnumerator en)
		{

			if (!tables.HasContext)
			{
				// When the Html is bad-formed and doesn't contain <table>, the browser renders the column separated by a space.
				// So we do the same here
				Run run = new Run(new Text(" ") { Space = SpaceProcessingModeValues.Preserve });
				htmlStyles.Runs.ApplyTags(run);
				elements.Add(run);
				return;
			}
			TableCell cell = tables.CurrentTable.GetLastChild<TableRow>().GetLastChild<TableCell>();

			// As we add automatically a paragraph to the cell once we create it, we'll remove it if finally, it was not used.
			// For all the other children, we will ensure there is no more empty paragraphs (similarly to what we do at the end
			// of the convert processing).
			// use a basic loop instead of foreach to allow removal (bug reported by antgraf)
			for (int i = 0; i < cell.ChildElements.Count; )
			{
				Paragraph p = cell.ChildElements[i] as Paragraph;
                // care of hyperlinks as they are not inside Run (bug reported by mdeclercq github.com/onizet/html2openxml/workitem/11162)
                if (p != null && !p.HasChild<Run>() && !p.HasChild<Hyperlink>())
                {
                    p.Remove();
                }
                else
                {
                    i++;
                }
			}

            // We add this paragraph regardless it has elements or not. A TableCell requires at least a Paragraph, as the last child of
            // of a table cell.
            // additional check for a proper cleaning (reported by antgraf github.com/onizet/html2openxml/discussions/272744)
            if (!(cell.LastChild is Paragraph) || elements.Count > 0)
            {
                cell.Append(new Paragraph(elements));
            }

			htmlStyles.Tables.ApplyTags(cell);

			// Reset all our variables and move to next cell
			this.elements.Clear();
			String openingTag = en.CurrentTag.Replace("/", "");
			htmlStyles.Tables.EndTag(openingTag);
			htmlStyles.Runs.EndTag(openingTag);

			var pos = tables.CellPosition;
			pos.Column++;
			tables.CellPosition = pos;
		}

		#endregion
	}
}
