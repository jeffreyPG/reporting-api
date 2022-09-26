using System.Web;

namespace reports.Extensions
{
    public static class StringExtensions
    {

        public static string Decode(this string stringToDecode)
        {
            var result = HttpUtility.UrlDecode(stringToDecode);
            return result;
        }
    }
}