using System.Collections.Generic;

namespace HtmlToOpenXml.Extensions
{
    public static class IntegerExtensions
    {

        private static List<string> listsTypes = new List<string> { "decimal", "lower-alpha", "lower-roman" };

        public static string GetOrderedListType(this int listLevel)
        {
            return listsTypes[listLevel % listsTypes.Count];
        }

    }
}
