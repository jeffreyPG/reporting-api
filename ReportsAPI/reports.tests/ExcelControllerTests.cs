using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using reports.Controllers;
using reports.Excel;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Net;

namespace reports.tests
{
    [TestClass]
    public class ExcelControllerTests
    {
        IProjectExcel projectExcel = new ProjectExcel();
        ExcelController excelController;
        private string ReportType_Building = "building";
        private string ReportType_Project = "project";

        public ExcelControllerTests()
        {
            excelController = new ExcelController(projectExcel);
        }

        /// <summary>
        /// This Testcase compares the response length and 
        /// </summary>
        /// <returns></returns>
        [TestMethod]
        public async Task GetBuildingReport_HappyPath()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var reportData = new List<Models.SpreadSheetReportData>();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "Name" });
            var excelData = new List<Dictionary<string, string>>();
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "1" },
                { "Name", "Rajeev" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "2" },
                { "Name", "John" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "3" },
                { "Name", "Alex" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "4" },
                { "Name", "Scott" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "5" },
                { "Name", "Matt" }
            });
            reportData.Add(new Models.SpreadSheetReportData()
            {
                ColumnNames = columnNames,
                DataSource = "Overview",
                Data = excelData,
                SheetName = "Sheet1"
            });
            data.BuildingReportData = new Models.BuildingReport()
            {
                ReportData = reportData,
            };

            #endregion

            //Get the Filedata in Bytes
            var currentPath = Directory.GetCurrentDirectory();
            var fileData = File.ReadAllBytes(currentPath + "\\..\\..\\Files\\SampleResponse.xlsx");

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Building);
            var result = await response.Content.ReadAsByteArrayAsync();

            //Assert - Compare the byte lenght of response and file
            Assert.AreEqual(result.Length, fileData.Length);
        }

        /// <summary>
        /// Invalid column name is specified in the Columnheader
        /// </summary>
        [TestMethod]
        public async Task GetBuildingReport_InvalidColumns()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var reportData = new List<Models.SpreadSheetReportData>();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpName" });
            var excelData = new List<Dictionary<string, string>>();
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "1" },
                { "Name", "Rajeev" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "2" },
                { "Name", "John" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "3" },
                { "Name", "Alex" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "4" },
                { "Name", "Scott" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "5" },
                { "Name", "Matt" }
            });
            reportData.Add(new Models.SpreadSheetReportData()
            {
                ColumnNames = columnNames,
                DataSource = "Overview",
                Data = excelData,
                SheetName = "Sheet1"
            });
            data.BuildingReportData = new Models.BuildingReport()
            {
                ReportData = reportData,
            };

            #endregion

            //Get the Filedata in Bytes
            var currentPath = Directory.GetCurrentDirectory();
            var fileData = File.ReadAllBytes(currentPath + "\\..\\..\\Files\\InvalidColumns.xlsx");

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Building);
            var result = await response.Content.ReadAsByteArrayAsync();

            //Assert - Compare the byte lenght of response and file
            Assert.AreEqual(result.Length, fileData.Length);
        }

        /// <summary>
        /// Invalid column name is specified in the data
        /// </summary>
        [TestMethod]
        public async Task GetBuildingReport_InvalidData()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var reportData = new List<Models.SpreadSheetReportData>();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "Name" });
            var excelData = new List<Dictionary<string, string>>();
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "1" },
                { "EmpName", "Rajeev" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "2" },
                { "EmpName", "John" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "3" },
                { "EmpName", "Alex" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "4" },
                { "EmpName", "Scott" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "5" },
                { "EmpName", "Matt" }
            });
            reportData.Add(new Models.SpreadSheetReportData()
            {
                ColumnNames = columnNames,
                DataSource = "Overview",
                Data = excelData,
                SheetName = "Sheet1"
            });
            data.BuildingReportData = new Models.BuildingReport()
            {
                ReportData = reportData,
            };

            #endregion

            //Get the Filedata in Bytes
            var currentPath = Directory.GetCurrentDirectory();
            var fileData = File.ReadAllBytes(currentPath + "\\..\\..\\Files\\InvalidData.xlsx");

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Building);
            var result = await response.Content.ReadAsByteArrayAsync();

            //Assert - Compare the byte lenght of response and file
            Assert.AreEqual(result.Length, fileData.Length);
        }

        /// <summary>
        /// Testcase with out any data
        /// </summary>
        [TestMethod]
        public async Task GetBuildingReport_NoData()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var reportData = new List<Models.SpreadSheetReportData>();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "Name" });
            var excelData = new List<Dictionary<string, string>>();
            reportData.Add(new Models.SpreadSheetReportData()
            {
                ColumnNames = columnNames,
                DataSource = "Overview",
                Data = excelData,
                SheetName = "Sheet1"
            });
            data.BuildingReportData = new Models.BuildingReport()
            {
                ReportData = reportData,
            };

            #endregion

            //Get the Filedata in Bytes
            var currentPath = Directory.GetCurrentDirectory();
            var fileData = File.ReadAllBytes(currentPath + "\\..\\..\\Files\\NoData.xlsx");

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Building);
            var result = await response.Content.ReadAsByteArrayAsync();

            //Assert - Compare the byte lenght of response and file
            Assert.AreEqual(result.Length, fileData.Length);
        }

        /// <summary>
        /// This method is for testing the project report in vertical alignment
        /// </summary>
        /// <returns></returns>
        [TestMethod]
        public async Task GetProjectReport_Vertical_HappyPath()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "Name" });
            var excelData = new List<Dictionary<string, string>>();
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "1" },
                { "Name", "Rajeev" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "2" },
                { "Name", "John" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "3" },
                { "Name", "Alex" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "4" },
                { "Name", "Scott" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "5" },
                { "Name", "Matt" }
            });
            var sheetData = new Models.SpreadSheetReportData()
            {
                ColumnNames = columnNames,
                Data = excelData,
                DataSource = "Overview",
                SheetName = "Sheet1"
            };
            data.ProjectReportData = new Models.ProjectReport()
            {
                ProjectData = new List<Models.SpreadSheetReportData>() { sheetData },
                Layout = "vertical"
            };

            #endregion

            //Get the Filedata in Bytes
            var currentPath = Directory.GetCurrentDirectory();
            var fileData = File.ReadAllBytes(currentPath + "\\..\\..\\Files\\ProjectReport_Vertical.xlsx");

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Project);
            var result = await response.Content.ReadAsByteArrayAsync();

            //Assert - Compare the byte lenght of response and file
            Assert.AreEqual(result.Length, fileData.Length);
        }

        /// <summary>
        /// This method is for testing the project report in horizontal alignment
        /// </summary>
        /// <returns></returns>
        [TestMethod]
        public async Task GetProjectReport_Horizontal_HappyPath()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "Name" });
            var excelData = new List<Dictionary<string, string>>();
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "1" },
                { "Name", "Rajeev" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "2" },
                { "Name", "John" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "3" },
                { "Name", "Alex" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "4" },
                { "Name", "Scott" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "5" },
                { "Name", "Matt" }
            });
            var sheetData = new Models.SpreadSheetReportData()
            {
                ColumnNames = columnNames,
                Data = excelData,
                DataSource = "Overview",
                SheetName = "Sheet1"
            };
            data.ProjectReportData = new Models.ProjectReport()
            {
                ProjectData = new List<Models.SpreadSheetReportData>() { sheetData },
                Layout = "horizontal"
            };

            #endregion

            //Get the Filedata in Bytes
            var currentPath = Directory.GetCurrentDirectory();
            var fileData = File.ReadAllBytes(currentPath + "\\..\\..\\Files\\ProjectReport_Horizontal.xlsx");

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Project);
            var result = await response.Content.ReadAsByteArrayAsync();

            //Assert - Compare the byte lenght of response and file
            Assert.AreEqual(result.Length, fileData.Length);
        }

        /// <summary>
        /// This method is for testing the project report in horizontal alignment
        /// </summary>
        /// <returns></returns>
        [TestMethod]
        public async Task GetProjectReport_Horizontal_InvalidColumns()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "Name1" });
            var excelData = new List<Dictionary<string, string>>();
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "1" },
                { "Name", "Rajeev" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "2" },
                { "Name", "John" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "3" },
                { "Name", "Alex" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "4" },
                { "Name", "Scott" }
            });
            excelData.Add(new Dictionary<string, string>
            {
                { "EmpId", "5" },
                { "Name", "Matt" }
            });
            var sheetData = new Models.SpreadSheetReportData()
            {
                ColumnNames = columnNames,
                Data = excelData,
                DataSource = "Overview",
                SheetName = "Sheet1"
            };
            data.ProjectReportData = new Models.ProjectReport()
            {
                ProjectData = new List<Models.SpreadSheetReportData>() { sheetData },
                Layout = "horizontal"
            };

            #endregion

            //Get the Filedata in Bytes
            var currentPath = Directory.GetCurrentDirectory();
            var fileData = File.ReadAllBytes(currentPath + "\\..\\..\\Files\\ProjectReport_InValidColumns.xlsx");

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Project);
            var result = await response.Content.ReadAsByteArrayAsync();

            //Assert - Compare the byte lenght of response and file
            Assert.AreEqual(result.Length, fileData.Length);
        }

        /// <summary>
        /// This method is for testing the project report in horizontal alignment
        /// </summary>
        /// <returns></returns>
        [TestMethod]
        public async Task GetProjectReport_Horizontal_NoData()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "Name" });
            var excelData = new List<Dictionary<string, string>>();
            var sheetData = new Models.SpreadSheetReportData()
            {
                ColumnNames = columnNames,
                Data = excelData,
                DataSource = "Overview",
                SheetName = "Sheet1"
            };
            data.ProjectReportData = new Models.ProjectReport()
            {
                ProjectData = new List<Models.SpreadSheetReportData>() { sheetData },
                Layout = "horizontal"
            };

            #endregion

            //Get the Filedata in Bytes
            var currentPath = Directory.GetCurrentDirectory();
            var fileData = File.ReadAllBytes(currentPath + "\\..\\..\\Files\\ProjectReport_NoData.xlsx");

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Project);
            var result = await response.Content.ReadAsByteArrayAsync();

            //Assert - Compare the byte lenght of response and file
            Assert.AreEqual(result.Length, fileData.Length);
        }

        /// <summary>
        /// This method is for testing the project report in horizontal alignment
        /// </summary>
        /// <returns></returns>
        [TestMethod]
        public void GetProjectReport_BadModel_Exception()
        {
            #region Arrange Data
            Models.SpreadSheetReport data = new Models.SpreadSheetReport();
            var columnNames = new List<Models.ColumnNameType>();
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "EmpId" });
            columnNames.Add(new Models.ColumnNameType() { ColumnName = "Name" });
            data.ProjectReportData = new Models.ProjectReport()
            {
                ProjectData = new List<Models.SpreadSheetReportData>()
                {
                },
                Layout = ""
            };

            #endregion

            //Call the API
            var response = excelController.GetReport(data, this.ReportType_Project);

            //Assert - Compare the http status code
            Assert.AreEqual(response.StatusCode, HttpStatusCode.InternalServerError);
        }
    }
}
