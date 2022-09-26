using OfficeOpenXml;
using reports.Models;
using System;
using System.IO;
using System.Web;

namespace reports.Excel
{
    public class NYCExcel
    {

        public byte[] PopulateNYCData(NYCData nycData)
        {
            string nycDoc = HttpContext.Current.Server.MapPath("~/NYC-reports/NYC-retro-Cx-reporting.xlsx");
            var nycTemplate = new FileInfo(nycDoc);

            using (ExcelPackage nycReport = new ExcelPackage(nycTemplate))
            {
                DateTime today = DateTime.Today;

                #region Submittal Information
                var submittalInfo = nycReport.Workbook.Worksheets["Submittal Information"];
                submittalInfo.Cells["C8"].Value = nycData.SubmittalInfo.SubmittedBy;
                submittalInfo.Cells["C9"].Value = nycData.SubmittalInfo.Company;
                submittalInfo.Cells["C10"].Value = nycData.SubmittalInfo.Phone;
                submittalInfo.Cells["C12"].Value = nycData.SubmittalInfo.Email;
                submittalInfo.Cells["C13"].Value = today.ToString("d");
                submittalInfo.Cells["C18"].Value = nycData.SubmittalInfo.Borough;
                submittalInfo.Cells["C19"].Value = nycData.SubmittalInfo.Block;
                submittalInfo.Cells["C20"].Value = nycData.SubmittalInfo.Lot;
                submittalInfo.Cells["C21"].Value = 1;
                submittalInfo.Cells["B25"].Value = nycData.SubmittalInfo.BinNumber;
                submittalInfo.Cells["C25"].Value = nycData.SubmittalInfo.Address;
                submittalInfo.Cells["D25"].Value = nycData.SubmittalInfo.Zip;
                #endregion

                #region Team Info
                var teamInfo = nycReport.Workbook.Worksheets["Team Info"];
                teamInfo.Cells["C9"].Value = nycData.TeamInfo.ProfessionalName;
                teamInfo.Cells["C10"].Value = nycData.TeamInfo.License;
                teamInfo.Cells["C11"].Value = nycData.TeamInfo.LicenseNo;
                teamInfo.Cells["C13"].Value = nycData.TeamInfo.Company;
                teamInfo.Cells["C14"].Value = nycData.TeamInfo.Address;
                teamInfo.Cells["C15"].Value = nycData.TeamInfo.Phone;
                teamInfo.Cells["C17"].Value = nycData.TeamInfo.CommissioningAgent;
                teamInfo.Cells["C18"].Value = nycData.TeamInfo.YearsExperience;
                teamInfo.Cells["C19"].Value = nycData.TeamInfo.CertType;
                teamInfo.Cells["C20"].Value = nycData.TeamInfo.CertExpirationDate;
                teamInfo.Cells["C21"].Value = today.ToString("d");
                teamInfo.Cells["F8"].Value = nycData.SubmittalInfo.BinNumber;
                #endregion

                #region Building Info
                var buildingInfo = nycReport.Workbook.Worksheets["Building Info"];
                buildingInfo.Cells["C9"].Value = nycData.BuildingInfo.Owner;
                buildingInfo.Cells["C10"].Value = nycData.BuildingInfo.OwnerRepresentative;
                buildingInfo.Cells["C11"].Value = nycData.BuildingInfo.ManagementCompany;
                buildingInfo.Cells["C12"].Value = nycData.BuildingInfo.ManagementContact;
                buildingInfo.Cells["C13"].Value = nycData.BuildingInfo.Phone;
                buildingInfo.Cells["C15"].Value = nycData.BuildingInfo.OperatorName;
                buildingInfo.Cells["C16"].Value = nycData.BuildingInfo.OperatorCert;
                buildingInfo.Cells["C18"].Value = nycData.BuildingInfo.OperatorLicenseNo;
                buildingInfo.Cells["C19"].Value = nycData.BuildingInfo.State;
                buildingInfo.Cells["F8"].Value = nycData.SubmittalInfo.BinNumber;
                #endregion

                #region RCMs
                var rcms = nycReport.Workbook.Worksheets["RCMs"];

                for (int row = 13; row <= 39; row++)
                {
                    foreach (Project project in nycData.Projects)
                    {
                        String cellValue = rcms.Cells[row, 2].Text;
                        if (cellValue.Contains(project.Name))
                        {
                            rcms.Cells[row, 3].Value = project.Compliant;
                            rcms.Cells[row, 4].Value = project.Notes;
                            rcms.Cells[row, 5].Value = project.DeficiencyCorrected;
                            rcms.Cells[row, 6].Value = project.ApproachToCompliance;
                            rcms.Cells[row, 7].Value = project.ImplementationCost;
                            rcms.Cells[row, 8].Value = project.Electricity;
                            rcms.Cells[row, 9].Value = project.Gas;
                            rcms.Cells[row, 10].Value = project.Oil;
                            rcms.Cells[row, 11].Value = project.Steam;
                            rcms.Cells[row, 12].Value = project.Other;
                            rcms.Cells[row, 13].Value = project.AnnualEnergySavings;
                            rcms.Cells[row, 14].Value = project.AnnualCostSavings;
                        }
                    }
                }
                rcms.Cells["Q15"].Value = nycData.SubmittalInfo.BinNumber;
                #endregion

                // return it as byte array
                return nycReport.GetAsByteArray();
            }
        }
    }
}