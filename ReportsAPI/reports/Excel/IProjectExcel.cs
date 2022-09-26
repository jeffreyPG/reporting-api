using reports.Models;
using reports.Models.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace reports.Excel
{
    /// <summary>
    /// Contains all methods for creating excel file
    /// </summary>
    public interface IProjectExcel
    {
        /// <summary>
        /// Returns the spreadsheet report
        /// </summary>
        /// <param name="model"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        ExcelReportResult GetSpreadsheetReport(SpreadSheetReport model, string type);
    }
}
