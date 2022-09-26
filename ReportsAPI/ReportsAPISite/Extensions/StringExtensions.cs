using System.Collections.Generic;
using System.Linq;

namespace ReportsAPISite.Extensions
{
    public static class StringExtensions
    {

        public static bool Empty(this IEnumerable<object> candidate)
        {
            return candidate.Any() == false;
        }

        public static bool IsNullOrEmpty(this string candidate)
        {
            return string.IsNullOrEmpty(candidate);
        }

        public static bool NotNullOrEmpty(this string candidate)
        {
            return !candidate.IsNullOrEmpty();
        }

        public static string UppercaseFirstCharacter(this string source)
        {
            if (string.IsNullOrEmpty(source))
            {
                return string.Empty;
            }

            var result = char.ToUpper(source[0]) + source.Substring(1);

            return result;
        }
    }
}