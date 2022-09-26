namespace reports.Models
{
    public class HtmlData
    {

        public string HtmlString { get; set; } 
        public bool isIncludeTOC { get; set; }
        public bool isShowPageNumber { get; set; }
        public string TOCDept { get; set; }
        public bool isPageNumberDisplayOnHeader { get; set; }
        public string pageNumberPosition { get; set; }
        public HeaderContent HeaderContent { get; set; }
        public FooterContent FooterContent { get; set; }
        public bool ReportStyles { get; set; }

    }
    public class HeaderContent
    {
        public string Image { get; set; }
        public string Content { get; set; }
        public string Position { get; set; }
        public string Divider { get; set; }

        public override string ToString()
        {
            return $"Image: {Image}; Content: {Content}; Position: {Position}; Divider: {Divider}";
        }
    }

    public class FooterContent
    {
        public string Image { get; set; }
        public string Content { get; set; }
        public string Position { get; set; }
        public string Divider { get; set; }

        public override string ToString()
        {
            return $"Image: {Image}; Content: {Content}; Position: {Position}; Divider: {Divider}";
        }
    }

}