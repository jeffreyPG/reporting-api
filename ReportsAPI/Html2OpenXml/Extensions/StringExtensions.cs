using HtmlAgilityPack;

namespace HtmlToOpenXml.Extensions
{
    public static class StringExtensions
    {

        public static int GetQuillJSIndentLevel(this string source)
        {
            // eg.: ql-indent-85
            var lastIndexOf = source.LastIndexOf('-');
            return int.Parse(source.Substring(lastIndexOf + 1));
        }

        public static string Repeat(this string source, int numberOfTabs)
        {
            return new string('\t', numberOfTabs) + source;
        }

        public static string ThicknessOrDefault(this string source)
        {
            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(source);            
            var thickness = htmlDocument.DocumentNode.SelectSingleNode("//line").GetAttributeValue("thickness", "1.5");
            return $"{thickness}pt";
        }

        public static string FillColorOrDefault(this string source)
        {
            HtmlDocument htmlDocument = new HtmlDocument();
            htmlDocument.LoadHtml(source);
            var color = htmlDocument.DocumentNode.SelectSingleNode("//line").GetAttributeValue("color", "000000");
            return $"#{color}";
        }

    }
}
