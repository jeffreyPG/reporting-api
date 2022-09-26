using HtmlAgilityPack;
using HtmlToOpenXml.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace HtmlToOpenXml.Services.HtmlCleaning
{
    public class HtmlCleanerBuilder : IHtmlCleaner
    {

        private string Html { get; set; }
        
        public HtmlCleanerBuilder(string _html)
        {
            Html = _html;
        }

        // Remove Script tags, doctype, comments, css style, controls and html head part
        public IHtmlCleaner RemoveHeaderContent()
        {
            var pattern = @"<xml.+?</xml>|<!--.+?-->|<script.+?</script>|<style.+?</style>|<head.+</head>|<!.+?>|<input.+?/>|<select.+?</select>|<textarea.+?</textarea>|<button.+?</button>";
            Html = Regex.Replace(Html, pattern, String.Empty,
                                 RegexOptions.IgnoreCase | RegexOptions.Singleline);
            return this;
        }

        // Removes tabs and whitespace inside and before|next the line-breaking tags (p, div, br and body) to preserve first whitespaces on the beginning of a 'pre' tag, we use '\bp\b' tag to exclude matching <pre> (by giorand, bug #13800)
        public IHtmlCleaner RemoveTabsAndWhiteSpace()
        {
            var pattern = @"(\s*)(</?(\bp\b|div|br|body)[^>]*/?>)(\s*)";
            Html = Regex.Replace(Html, pattern, "$2", RegexOptions.Multiline | RegexOptions.IgnoreCase);
            return this;
        }

        // Preserves whitespaces inside Pre tags.
        public IHtmlCleaner PreserveWhitespaceInsidePreTags()
        {
            var pattern = "(<pre.*?>)(.+?)</pre>";
            Html = Regex.Replace(Html, pattern, PreserveWhitespacesInPre, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            return this;
        }

        // Remove tabs and whitespace at the beginning of the lines
        public IHtmlCleaner RemoveTabsAndWhitespaceAtTheBeginning()
        {
            var pattern = @"^\s+";
            Html = Regex.Replace(Html, pattern, String.Empty, RegexOptions.Multiline);
            return this;
        }

        // and now at the end of the lines
        public IHtmlCleaner RemoveTabsAndWhitespaceAtTheEnd()
        {
            var pattern = @"\s+$";
            Html = Regex.Replace(Html, pattern, String.Empty, RegexOptions.Multiline);
            return this;
        }

        // Replace xml header by xml tag for further processing
        public IHtmlCleaner ReplaceXmlHeaderByXmlTag()
        {
            var pattern = @"<\?xml:namespace.+?>";
            Html = Regex.Replace(Html, pattern, "<xml>", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            return this;
        }

        // Ensure order of table elements are respected: thead, tbody and tfooter 
        // we select only the table that contains at least a tfoot or thead tag
        public IHtmlCleaner EnsureOrderOfTableElements()
        {
            var pattern = @"<table.*?>(\s+</?(?=(thead|tbody|tfoot))).+?</\2>\s+</table>";
            Html = Regex.Replace(Html, pattern, PreserveTablePartOrder, RegexOptions.Singleline);

            pattern = "(<table.*?>)(.*?)(</table>)";
            Html = Regex.Replace(Html, pattern, PreserveTablePartOrder, RegexOptions.Singleline | RegexOptions.IgnoreCase);
            return this;
        }

        public IHtmlCleaner RemoveCarriageReturns( )
        {
            var pattern = @"&#13;";
            Html = Regex.Replace(Html, pattern, string.Empty, RegexOptions.None);
            return this;
        }

        public IHtmlCleaner ReplaceAmpersandHtmlCodeSet()
        {
            var pattern = @"&#38;";
            Html = Regex.Replace(Html, pattern, "&", RegexOptions.Multiline);

            pattern = "&amp;";
            Html = Regex.Replace(Html, pattern, "&", RegexOptions.Multiline);

            pattern = "&nbsp;";
            Html = Regex.Replace(Html, pattern, " ", RegexOptions.Multiline);

            return this;
        }

        public IHtmlCleaner RemoveExceedingTableHeaders()
        {
            var pattern = "<thead><thead>";
            Html = Regex.Replace(Html, pattern, "<thead>", RegexOptions.Multiline);
            return this;
        }
        
        public IHtmlCleaner RemoveLineBreakAfterTableClosingElement()
        {
            var pattern = "</table><br/>";
            Html = Regex.Replace(Html, pattern, "</table>", RegexOptions.Multiline);
            return this;
        }

        public IHtmlCleaner RemoveLineBreakInsideParagraphTag()
        {
            var pattern = $"<p>(<br>|<br />|<br/>)</p>";
            Html = Regex.Replace(Html, pattern, "<p> </p>", RegexOptions.Multiline);
            for (var i = 2; i < 10; i++)
            {
                pattern = $"<p>(<br>|<br />|<br/>){{{i}}}</p>";
                Html = Regex.Replace(Html, pattern, "<p>$1</p>", RegexOptions.Multiline);
            }

            return this;
        }

        // TODO: we need to make this recursive
        public IHtmlCleaner EnsureOrderedListsAreIndented()
        {

            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(Html);
            HtmlNode.ElementsFlags["p"] = HtmlElementFlag.Closed;

            var orderedList = htmlDocument.DocumentNode.SelectNodes("//ol");

            if(orderedList == null)
            {
                return this;
            }

            var listCount = 0;
            var finalNodes = new List<HtmlNode>();

            foreach (var orderedListItem in orderedList)
            {

                var currentLevel = 0;
                var finalNode = HtmlNode.CreateNode($"<ol></ol>");
                finalNode.Id = $"list-{listCount}-level0";

                var attribute = htmlDocument.CreateAttribute("style", $"list-style-type: {currentLevel.GetOrderedListType()}");
                finalNode.Attributes.Add(attribute);

                foreach (var listItem in orderedListItem.ChildNodes)
                {

                    var level = listItem.GetClasses().Where(foo => foo.Contains("ql-indent-")).SingleOrDefault();
                    if (level != null)
                    {

                        var nodeCurrentLevel = level.GetQuillJSIndentLevel();             
                        var levelId = $"list-{listCount}-level{nodeCurrentLevel}";

                        if (currentLevel < nodeCurrentLevel)
                        {

                            var ol = HtmlNode.CreateNode($"<ol></ol>");
                            ol.Id = levelId;

                            attribute = htmlDocument.CreateAttribute("style", $"list-style-type: {nodeCurrentLevel.GetOrderedListType()}");
                            attribute.QuoteType = AttributeValueQuote.SingleQuote;
                            ol.Attributes.Add(attribute);

                            ol.AppendChild(listItem);

                            if (currentLevel == 0)
                            {
                                finalNode.AppendChild(ol);
                            }
                            else
                            {
                                levelId = $"list-{listCount}-level{currentLevel}";
                                finalNode.SelectSingleNode($"//ol[@id='{levelId}']").AppendChild(ol);
                            }                            

                        }
                        else if (currentLevel > nodeCurrentLevel)
                        {                            
                            finalNode.SelectSingleNode($"//ol[@id='{levelId}']").AppendChild(listItem);
                        }
                        else
                        {
                            finalNode.SelectSingleNode($"//ol[@id='{levelId}']").AppendChild(listItem);
                        }

                        currentLevel = nodeCurrentLevel;

                    }
                    else
                    {
                        finalNode.AppendChild(listItem);
                    }
                }

                orderedListItem.ParentNode.ReplaceChild(finalNode, orderedListItem);
                listCount++;

            }

            Html = htmlDocument.DocumentNode.InnerHtml;
            return this;

        }

        // TODO: we need to make this recursive
        public IHtmlCleaner EnsureUnorderedListsAreIndented()
        {

            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(Html);
            HtmlNode.ElementsFlags["p"] = HtmlElementFlag.Closed;

            var unorderedList = htmlDocument.DocumentNode.SelectNodes("//ul");
            if (unorderedList == null)
            {
                return this;
            }

            var listCount = 0;

            foreach (var unorderedListItem in unorderedList)
            {

                var currentLevel = 0;
                var finalNode = HtmlNode.CreateNode($"<ul></ul>");
                finalNode.Id = $"list-{listCount}-level{currentLevel}";

                var listItems = unorderedListItem.ChildNodes;

                foreach (var listItem in listItems)
                {

                    var level = listItem.GetClasses().Where(foo => foo.Contains("ql-indent-")).SingleOrDefault();
                    if (level != null)
                    {
                        
                        var nodeCurrentLevel = int.Parse(level.Substring(level.Length - 1, 1));
                        var levelId = $"list-{listCount}-level{nodeCurrentLevel}";

                        var ul = HtmlNode.CreateNode("<ul></ul>");
                        ul.Id = levelId;
                        ul.AppendChild(listItem);

                        if (currentLevel < nodeCurrentLevel)
                        {

                            if (currentLevel == 0)
                            {
                                finalNode.AppendChild(ul);
                            }
                            else
                            {
                                levelId = $"list-{listCount}-level{currentLevel}";
                                finalNode.SelectSingleNode($"//ul[@id='{levelId}']").AppendChild(ul);
                            }

                        }
                        else if (currentLevel > nodeCurrentLevel)
                        {
                            finalNode.SelectSingleNode($"//ul[@id='{levelId}']").AppendChild(listItem);
                        }
                        else
                        {
                            finalNode.SelectSingleNode($"//ul[@id='{levelId}']").AppendChild(listItem);
                        }

                        currentLevel = nodeCurrentLevel;

                    }
                    else
                    {
                        finalNode.AppendChild(listItem);
                    }
                }

                unorderedListItem.ParentNode.ReplaceChild(finalNode, unorderedListItem);
                listCount++;

            }
            
            Html = htmlDocument.DocumentNode.InnerHtml;
            return this;
        }

        public IHtmlCleaner ReplaceCSSStylesWithHtmlCorrespondents()
        {
            var pattern = $"class=\"ql-align-left\"";
            Html = Regex.Replace(Html, pattern, "align=\"left\"", RegexOptions.Multiline);

            //pattern = @"class=""ql-align-left""";
            //Html = Regex.Replace(Html, pattern, "align=\"left\"", RegexOptions.Multiline);

            pattern = $"class=\"ql-align-right\"";
            Html = Regex.Replace(Html, pattern, "align=\"right\"", RegexOptions.Multiline);

            //pattern = @"class=""ql-align-right""";
            //Html = Regex.Replace(Html, pattern, "align=\"right\"", RegexOptions.Multiline);

            pattern = $"class=\"ql-align-center\"";
            Html = Regex.Replace(Html, pattern, "align=\"center\"", RegexOptions.Multiline);

            //pattern = @"class=""ql-align-center""";
            //Html = Regex.Replace(Html, pattern, "align=\"center\"", RegexOptions.Multiline);

            pattern = $"class=\"ql-align-justify\"";
            Html = Regex.Replace(Html, pattern, "align=\"justify\"", RegexOptions.Multiline);

            //pattern = @"class=""ql-align-justify""";
            //Html = Regex.Replace(Html, pattern, "align=\"justify\"", RegexOptions.Multiline);

            return this;
        }

        public IHtmlCleaner ReplaceQuillJsIndentStyle()
        {
            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(Html);
            HtmlNode.ElementsFlags["p"] = HtmlElementFlag.Closed;

            var allPTagsWithIndentedStyles = htmlDocument.DocumentNode.SelectNodes("//p");

            if(allPTagsWithIndentedStyles == null)
            {
                return this;
            }

            foreach (var paragraph in allPTagsWithIndentedStyles)
            {
                var hasStyle = paragraph.GetClasses().Where(foo => foo.Contains("ql-indent")).FirstOrDefault();
                if(hasStyle != null)
                {
                    var numberOfIndents = hasStyle.GetQuillJSIndentLevel();
                    paragraph.InnerHtml = paragraph.InnerHtml.Repeat(numberOfIndents);
                }
            }

            Html = htmlDocument.DocumentNode.InnerHtml;
            return this;
        }

        public string Build()
        {
            return Html;
        }

        private string PreserveWhitespacesInPre(Match match)
        {
            // Convert new lines in <pre> to <br> tags for easier processing
            string innerHtml = Regex.Replace(match.Groups[2].Value, "\r?\n", "<br>");
            // Remove any whitespace at the end of the pre
            innerHtml = Regex.Replace(innerHtml, @"(<br>|\s+)$", String.Empty);
            return match.Groups[1].Value + innerHtml + "</pre>";
        }
        
        private string PreserveTablePartOrder(Match match)
        {
            // ensure the order of the table elements are set in the correct order.
            // bug #11016 reported by pauldbentley

            var sb = new System.Text.StringBuilder();
            sb.Append(match.Groups[1].Value);

            Regex tableSplitReg = new Regex(@"(<(?=(caption|colgroup|thead|tbody|tfoot|tr)).*?>.+?</\2>)", RegexOptions.Singleline | RegexOptions.IgnoreCase);
            MatchCollection matches = tableSplitReg.Matches(match.Groups[2].Value);

            foreach (String tagOrder in new[] { "caption", "colgroup", "thead", "tbody", "tfoot", "tr" })
                foreach (Match m in matches)
                {
                    if (m.Groups[2].Value.Equals(tagOrder, StringComparison.OrdinalIgnoreCase))
                        sb.Append(m.Groups[1].Value);
                }

            sb.Append(match.Groups[3].Value);
            return sb.ToString();
        }

    }

}