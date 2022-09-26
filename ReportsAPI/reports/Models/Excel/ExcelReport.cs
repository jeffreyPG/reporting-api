using reports.Models.Excel;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace reports.Models
{
    /// <summary>
    /// This model is used for generating the report data
    /// </summary>
    public class SpreadSheetReport
    {
        /// <summary>
        /// Building ID of the project/building report
        /// </summary>
        public string BuildingId { get; set; }

        /// <summary>
        /// This property stores the Building Report Data
        /// </summary>
        public BuildingReport BuildingReportData { get; set; }

        /// <summary>
        /// This property stores the Project report data
        /// </summary>
        public ProjectReport ProjectReportData { get; set; }
    }

    /// <summary>
    /// This model is used for generating the building report data
    /// </summary>
    public class BuildingReport
    {
        /// <summary>
        /// This property stores the data for the source type 'Overview & Property'
        /// </summary>
        [ReportDataValidator]
        public List<SpreadSheetReportData> ReportData { get; set; }

    }

    /// <summary>
    /// This model is used for generating the project data
    /// </summary>
    public class ProjectReport
    {
        /// <summary>
        /// Defines the orientation whether its horizontal/vertical
        /// </summary>
        [Required]
        public string Layout { get; set; }

        /// <summary>
        /// Stores the project data of the report
        /// </summary>
        [ReportDataValidator]
        public List<SpreadSheetReportData> ProjectData { get; set; }
    }
}