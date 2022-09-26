using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace reports.Models.Excel
{
    /// <summary>
    /// The output type of Excel report API
    /// </summary>
    public class ExcelReportResult
    {
        /// <summary>
        /// Stores the file content as bytes
        /// </summary>
        public byte[] Content { get; set; }

        /// <summary>
        /// Determines whether the request is a valid or not
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// Stores the error message
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// Stores the invalid column names
        /// </summary>
        public List<string> InvalidColumnNames { get; set; }
    }
}