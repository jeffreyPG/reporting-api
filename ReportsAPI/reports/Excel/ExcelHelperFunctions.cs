using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace reports.Excel
{
    public class ExcelHelperFunctions
    {
        public static string IsStringIntDouble(string content)
        {
            if (int.TryParse(content, out int n) || double.TryParse(content, out double d))
            {
                if (int.TryParse(content, out int num))
                {
                    return "int";
                }
                else if (double.TryParse(content, out double doub)) 
                {
                    return "double";
                }
                else
                {
                    return "unknown";
                }
            }
            else if (content is string)
            {
                return "string";
            }
            else
            {
                return "unknown";
            }
        }
    }
}