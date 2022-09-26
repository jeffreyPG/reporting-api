using ReportsAPISite.Models.Excel;

namespace ReportsAPISite.Services.Excel
{
    public interface ICreateExcelDocument
    {
        ExcelReportResult GetSpreadsheetReport(SpreadSheetReport model, string type);
    }
}