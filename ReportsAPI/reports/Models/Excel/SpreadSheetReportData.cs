using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace reports.Models
{

    /// <summary>
    /// 
    /// </summary>
    public class SpreadSheetReportData
    {
        /// <summary>
        /// For Storing the Datasource type
        /// </summary>
        public string DataSource { get; set; }

        /// <summary>
        /// Stores the sheet name
        /// </summary>
        [Required]
        public string SheetName { get; set; }

        /// <summary>
        /// Stores all the column names
        /// </summary>
        [Required]
        public List<ColumnNameType> ColumnNames { get; set; }

        /// <summary>
        /// Data that is used for generating the report
        /// </summary>
        [Required]
        public List<Dictionary<string, string>> Data { get; set; }
    }

    /// <summary>
    /// 
    /// </summary>
    public class ColumnNameType
    {
        /// <summary>
        /// The Column name 
        /// </summary>
        public string ColumnName { get; set; }
    }
}