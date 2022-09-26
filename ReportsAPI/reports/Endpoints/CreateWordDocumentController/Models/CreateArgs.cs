namespace reports.Endpoints.CreateWordDocumentController.Models
{
    public class CreateArgs
    {

        public string HtmlString { get; set; }
        public bool IncludeToC { get; set; }
        public bool ShowPageNumber { get; set; }
        public string ToCDepth { get; set; }
        public bool PageNumberDisplayOnHeader { get; set; }
        public string PageNumberPosition { get; set; }

        public string HeaderImage { get; set; }
        public string HeaderContent { get; set; }
        public string FooterContent { get; set; }

        // if null or empty, dont apply
        public string ReportStyleS3Id { get; set; }

        public string OrganizationId { get; set; }
        public string BuildingId { get; set; }
        public string ReportId { get; set; }

        public string ToString()
        {
            // TODO: implement
            
            return HtmlString;
        }

        // TODO: implement
        public void Validate()
        {

        }
    }

    public class HeaderContent
    {
        public string Image { get; set; }
        public string Content { get; set; }

        public string ToString()
        {
            // TODO: implement
            return Content;
        }

        // TODO: implement
        public void Validate()
        {

        }

    }

    public class FooterContent
    {
        public string Content { get; set; }

        public string ToString()
        {
            // TODO: implement
            return Content;
        }

        // TODO: implement
        public void Validate()
        {

        }

    }
}