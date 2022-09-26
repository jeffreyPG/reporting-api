using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;

namespace ReportsAPISite.Models.Excel
{
    public class ReportDataValidator : ValidationAttribute
    {
        protected override ValidationResult IsValid(object data, ValidationContext validationContext)
        {
            List<SpreadSheetReportData> rdata = (List<SpreadSheetReportData>)data;
            if (rdata?.Count > 0)
            {
                return ValidationResult.Success;
            }
            else
            {
                return new ValidationResult("Report data is not present");
            }
        }
    }
}