namespace HtmlToOpenXml.Services.HtmlCleaning
{
    public interface IHtmlCleaner
    {

        IHtmlCleaner RemoveHeaderContent();
        IHtmlCleaner RemoveTabsAndWhiteSpace();
        IHtmlCleaner PreserveWhitespaceInsidePreTags();
        IHtmlCleaner RemoveTabsAndWhitespaceAtTheBeginning();
        IHtmlCleaner RemoveTabsAndWhitespaceAtTheEnd();
        IHtmlCleaner ReplaceXmlHeaderByXmlTag();
        IHtmlCleaner EnsureOrderOfTableElements();
        IHtmlCleaner RemoveCarriageReturns();
        IHtmlCleaner ReplaceAmpersandHtmlCodeSet();
        IHtmlCleaner RemoveExceedingTableHeaders();
        IHtmlCleaner RemoveLineBreakAfterTableClosingElement();
        IHtmlCleaner EnsureUnorderedListsAreIndented();
        IHtmlCleaner EnsureOrderedListsAreIndented();
        IHtmlCleaner RemoveLineBreakInsideParagraphTag();
        IHtmlCleaner ReplaceCSSStylesWithHtmlCorrespondents();
        IHtmlCleaner ReplaceQuillJsIndentStyle();
        string Build();

    }
}