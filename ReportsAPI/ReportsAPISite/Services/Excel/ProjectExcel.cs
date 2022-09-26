using OfficeOpenXml;
using ReportsAPISite.Models.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;

namespace ReportsAPISite.Services.Excel
{
    public class ProjectExcel : ICreateExcelDocument
    {

        ExcelRange cell;

        public byte[] GenerateVerticalProject(ProjectData projectData)
        {
            using (var excelPackage = new ExcelPackage())
            {
                excelPackage.Workbook.Properties.Author = projectData.username;
                excelPackage.Workbook.Properties.Title = projectData.title;

                foreach (Report report in projectData.report)
                {
                    int rowIndex = 1;
                    int columnIndex = 1;
                    var sheet = excelPackage.Workbook.Worksheets.Add("Excel Report");
                    sheet.Name = report.sheetName;

                    #region Report Sheet title
                    sheet.Row(1).Height = 25;
                    cell = sheet.Cells[1, 1];
                    cell.Value = report.sheetName;
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.Size = 18;
                    rowIndex = rowIndex + 1;
                    #endregion

                    #region Report group titles
                    // fill merged columns that determine each group
                    foreach (Group group in report.data)
                    {
                        cell = sheet.Cells[rowIndex, 1];
                        cell.Value = group.group;
                        cell.Style.Font.Bold = true;
                        rowIndex = rowIndex + 1;

                        #region Report data titles
                        foreach (Section section in group.sections)
                        {
                            cell = sheet.Cells[rowIndex, 1];
                            cell.Value = section.title;

                            #region Report sections
                            foreach (string content in section.content)
                            {
                                columnIndex = columnIndex + 1;
                                cell = sheet.Cells[rowIndex, columnIndex];
                                string cellName = Regex.Replace(section.title, "[^\\w\\._]", "");
                                sheet.Names.Add(cellName + "_" + columnIndex, cell);
                                string contentType = ExcelHelperFunctions.IsStringIntDouble(content);
                                if (contentType == "int")
                                {
                                    cell.Value = int.Parse(content);
                                }
                                else if (contentType == "double")
                                {
                                    cell.Value = double.Parse(content);
                                }
                                else
                                {
                                    cell.Value = content;
                                }
                            }
                            #endregion

                            // move down 1 row to continue filling out
                            rowIndex = rowIndex + 1;
                            // move back to column 1to continue writing group section titles
                            columnIndex = 1;
                        }

                        rowIndex = rowIndex + 1;
                        #endregion
                    }
                    #endregion

                    // auto adjust cell width, with minimum and maximum size
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns(10, 40);
                    // text wrap on all columns
                    var start = sheet.Dimension.Start;
                    var end = sheet.Dimension.End;
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        sheet.Column(col).Style.WrapText = true;
                    }
                }

                return excelPackage.GetAsByteArray();
            }
        }

        public byte[] GenerateHorizontalProject(ProjectData projectData)
        {
            using (var excelPackage = new ExcelPackage())
            {
                excelPackage.Workbook.Properties.Author = "Corinne";
                excelPackage.Workbook.Properties.Title = " excel download";

                foreach (Report report in projectData.report)
                {
                    int rowIndex = 1;
                    int columnIndex = 1;
                    var sheet = excelPackage.Workbook.Worksheets.Add("Excel Report");
                    sheet.Name = report.sheetName;

                    #region Report Sheet title
                    sheet.Row(1).Height = 25;
                    cell = sheet.Cells[1, 1];
                    cell.Value = report.sheetName;
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.Size = 18;
                    rowIndex = rowIndex + 1;
                    #endregion

                    #region Report group titles
                    // fill merged columns that determine each group
                    foreach (Group group in report.data)
                    {
                        int sectionCount = group.sections.Length;
                        int mergedLength = columnIndex + sectionCount - 1;
                        sheet.Cells[rowIndex, columnIndex, rowIndex, mergedLength].Merge = true;
                        cell = sheet.Cells[rowIndex, columnIndex];
                        cell.Style.Font.Bold = true;
                        cell.Value = group.group;

                        #region Report data titles

                        // move down one row to start filling data titles
                        rowIndex = rowIndex + 1;
                        foreach (Section section in group.sections)
                        {
                            cell = sheet.Cells[rowIndex, columnIndex];
                            cell.Style.Font.Bold = true;
                            cell.Value = section.title;

                            #region Report sections
                            foreach (string content in section.content)
                            {
                                // move down one row
                                rowIndex = rowIndex + 1;
                                cell = sheet.Cells[rowIndex, columnIndex];
                                string cellName = Regex.Replace(section.title, "[^\\w\\._]", "");
                                sheet.Names.Add(cellName + "_" + rowIndex, cell);
                                string contentType = ExcelHelperFunctions.IsStringIntDouble(content);
                                if (contentType == "int")
                                {
                                    cell.Value = int.Parse(content);
                                }
                                else if (contentType == "double")
                                {
                                    cell.Value = double.Parse(content);
                                }
                                else
                                {
                                    cell.Value = content;
                                }
                            }
                            #endregion

                            // move forward 1 column to continue filling out
                            columnIndex = columnIndex + 1;
                            // move back up to row 3 to continue writing group section titles
                            rowIndex = 3;
                        }
                        // move back up to row 2 to continue writing group column titles
                        rowIndex = 2;

                        #endregion
                    }
                    #endregion

                    // merge top cell with sheet title all the way across the spreadsheet
                    sheet.Cells[1, 1, 1, columnIndex - 1].Merge = true;
                    // fix top three rows
                    sheet.View.FreezePanes(4, 1);
                    // auto adjust cell width, with minimum and maximum size
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns(10, 50);
                    // text wrap on all columns
                    var start = sheet.Dimension.Start;
                    var end = sheet.Dimension.End;
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        sheet.Column(col).Style.WrapText = true;
                    }

                    // testing cell formulas
                    sheet.Cells[7, 1].Value = "Sum of named cells WattageofExistingFixture_4 and WattageofExistingFixture_5";
                    sheet.Cells[7, 2].Formula = "=SUM(WattageofExistingFixture_4:WattageofExistingFixture_5)";

                }

                return excelPackage.GetAsByteArray();
            }
        }

        /// <summary>
        /// Generates the building/project report
        /// </summary>
        /// <param name="model"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public ExcelReportResult GetSpreadsheetReport(SpreadSheetReport model, string type)
        {
            var result = new ExcelReportResult();
            if (type.ToLower() == Utils.Constants.Building)
            {
                result = this.GetBuildingReport(model);
            }
            else if (type.ToLower() == Utils.Constants.Project)
            {
                result = this.GetProjectReport(model);
            }

            return result;
        }

        /// <summary>
        /// Returns the building report as byte array
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private ExcelReportResult GetBuildingReport(SpreadSheetReport model)
        {
            ExcelReportResult result = new ExcelReportResult();
            if (model?.BuildingReportData?.ReportData?.Count > 0)
            {
                using (var excelPackage = new ExcelPackage())
                {
                    foreach (var reportData in model.BuildingReportData.ReportData)
                    {
                        var columnNames = reportData?.ColumnNames;

                        // Filling the row headers/column names
                        var data = reportData?.Data;
                        int rowIndex = 2;
                        int colIndex = 1;
                        var sheetName = string.IsNullOrWhiteSpace(reportData?.SheetName) ? "untitled_" + colIndex : reportData?.SheetName;
                        var sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                        if (columnNames != null)
                        {
                            foreach (var colName in columnNames)
                            {
                                var cell = sheet.Cells[(rowIndex - 1), colIndex];
                                cell.Value = colName.ColumnName;
                                cell.Style.Font.Bold = true;
                                cell.Style.Font.Size = 18;
                                colIndex++;
                            }

                            // Filling the data
                            result.InvalidColumnNames = FillData(data, columnNames, sheet, rowIndex);

                            // auto adjust cell width, with minimum and maximum size
                            this.AutoAdjustCells(sheet);
                        }
                    }

                    result.Content = excelPackage.GetAsByteArray();
                }
            }

            return result;
        }

        /// <summary>
        /// Generates the project report in an excel file and returns the bytes
        /// </summary>
        /// <param name="model"></param>
        /// <returns></returns>
        private ExcelReportResult GetProjectReport(SpreadSheetReport model)
        {
            ExcelReportResult result = new ExcelReportResult();
            using (var excelPackage = new ExcelPackage())
            {
                foreach (var reportData in model.ProjectReportData.ProjectData)
                {
                    var data = reportData.Data;
                    var columnNames = reportData.ColumnNames;

                    if (model.ProjectReportData.Layout?.ToLower() == Utils.Constants.Vertical)
                    {
                        // Filling the row headers/column names
                        int rowIndex = 2;
                        int colIndex = 1;
                        var sheetName = string.IsNullOrWhiteSpace(reportData.SheetName) ? "untitled_" + colIndex : reportData.SheetName;
                        var sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                        if (columnNames != null)
                        {
                            foreach (var colName in columnNames)
                            {
                                var cell = sheet.Cells[(rowIndex - 1), colIndex];
                                cell.Value = colName.ColumnName;
                                cell.Style.Font.Bold = true;
                                cell.Style.Font.Size = 18;
                                colIndex++;
                            }

                            // Filling the data
                            result.InvalidColumnNames = FillData(data, columnNames, sheet, rowIndex);

                            // auto adjust cell width, with minimum and maximum size
                            this.AutoAdjustCells(sheet);
                        }
                    }
                    else if (model.ProjectReportData.Layout?.ToLower() == Utils.Constants.Horizontal)
                    {
                        int rowIndex = 1;
                        var sheetName = string.IsNullOrWhiteSpace(reportData.SheetName) ? "untitled_" + rowIndex : reportData.SheetName;
                        var sheet = excelPackage.Workbook.Worksheets.Add(sheetName);
                        int colIndex = 1;
                        if (columnNames != null)
                        {
                            foreach (var colName in columnNames)
                            {
                                var cell = sheet.Cells[(rowIndex), colIndex];
                                cell.Value = colName.ColumnName;
                                cell.Style.Font.Bold = true;
                                cell.Style.Font.Size = 18;
                                rowIndex++;
                            }

                            // Filling the data
                            FillDataHorizontally(data, columnNames, sheet, rowIndex);

                            // auto adjust cell width, with minimum and maximum size
                            this.AutoAdjustCells(sheet);
                        }
                    }
                }

                result.Content = excelPackage.GetAsByteArray();
            }

            return result;
        }

        /// <summary>
        /// Auto adjusts cell width, min & max size
        /// </summary>
        /// <param name="sheet"></param>
        private void AutoAdjustCells(ExcelWorksheet sheet)
        {
            if (sheet != null && sheet.Dimension != null && sheet.Cells != null)
            {
                // auto adjust cell width, with minimum and maximum size
                sheet.Cells[sheet.Dimension.Address].AutoFitColumns(10, 40);
                // text wrap on all columns
                var start = sheet.Dimension.Start;
                var end = sheet.Dimension.End;
                if (start != null && end != null)
                {
                    for (int col = start.Column; col <= end?.Column; col++)
                    {
                        if (sheet.Column(col) != null && sheet.Column(col).Style != null)
                            sheet.Column(col).Style.WrapText = true;
                    }
                }
            }
        }

        private List<string> FillDataHorizontally(List<Dictionary<string, string>> data, List<ColumnNameType> columnNames, ExcelWorksheet sheet, int rowIndex)
        {
            rowIndex = 1;
            List<string> unIdentifiedColumns = new List<string>();

            foreach (var columnName in columnNames)
            {
                var columnIndex = 2;
                foreach (var item in data)
                {
                    try
                    {
                        var cellData = "";

                        if (item.ContainsKey(columnName.ColumnName))
                            cellData = item[columnName.ColumnName];
                        else
                            unIdentifiedColumns.Add(columnName.ColumnName);

                        var cell = sheet.Cells[rowIndex, columnIndex];
                        this.RenderCell(cell, cellData, columnName.ColumnName);
                    }
                    catch (System.Exception)
                    {
                    }
                    columnIndex++;
                }
                rowIndex++;
            }

            return unIdentifiedColumns;
        }

        private List<string> FillData(List<Dictionary<string, string>> data, List<ColumnNameType> columnNames, ExcelWorksheet sheet, int rowIndex)
        {
            List<string> unIdentifiedColumns = new List<string>();
            foreach (var item in data)
            {
                /* Adding the try catch here to make sure that if any exception occurs in some particular row, 
                it wont stop there and the process continues for rest of the rows */
                try
                {
                    int columnIndex = 1;
                    foreach (var columnName in columnNames)
                    {
                        var cellData = "";
                        if (item.ContainsKey(columnName.ColumnName))
                            cellData = item[columnName.ColumnName];
                        else
                            unIdentifiedColumns.Add(columnName.ColumnName);
                        var cell = sheet.Cells[rowIndex, columnIndex];

                        this.RenderCell(cell, cellData, columnName.ColumnName);
                        columnIndex++;
                    }
                }
                catch (System.Exception)
                {
                }
                rowIndex++;
            }

            return unIdentifiedColumns;
        }

        private void RenderCell(ExcelRange cell, string content, string columnName)
        {
            string contentType = ExcelHelperFunctions.IsStringIntDouble(content);
            var matchCollection = Regex.Matches(content, "((http|ftp|https|www))");

            if (!string.IsNullOrWhiteSpace(content) && matchCollection?.Count > 0)
            {
                if (matchCollection[0].Index == 0)
                {
                    cell.Formula = "=HYPERLINK(\"" + content + "\", \"" + "Link" + "\")";
                    cell.Style.Font.UnderLine = true;
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                }
                else if (matchCollection[0].Index > 0)
                {
                    var displayText = content.Substring(0, matchCollection[0].Index).Trim();
                    var hyperLink = content.Substring(matchCollection[0].Index);

                    cell.Formula = "=HYPERLINK(\"" + hyperLink + "\", \"" + displayText + "\")";
                    cell.Style.Font.UnderLine = true;
                    cell.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                }
            }
            else
            {
                if (contentType == "int")
                {
                    cell.Value = int.Parse(content);
                }
                else if (contentType == "double")
                {
                    cell.Value = double.Parse(content);
                }
                else
                {
                    cell.Value = content;
                }
            }
        }
    }
}