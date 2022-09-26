using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ReportsAPISite.Services.Excel
{
    public class SubmittalInfo
    {
        public string SubmittedBy { get; set; }
        public string Company { get; set; }
        public string Phone { get; set; }
        public string Email { get; set; }
        public string Borough { get; set; }
        public string Block { get; set; }
        public string Lot { get; set; }
        public string BinNumber { get; set; }
        public string Address { get; set; }
        public string Zip { get; set; }
    }

    public class TeamInfo
    {
        public string ProfessionalName { get; set; }
        public string License { get; set; }
        public string LicenseNo { get; set; }
        public string Company { get; set; }
        public string Address { get; set; }
        public string Phone { get; set; }
        public string CommissioningAgent { get; set; }
        public int YearsExperience { get; set; }
        public string CertType { get; set; }
        public string CertExpirationDate { get; set; }
    }

    public class BuildingInfo
    {
        public string Owner { get; set; }
        public string OwnerRepresentative { get; set; }
        public string ManagementCompany { get; set; }
        public string ManagementContact { get; set; }
        public string Phone { get; set; }
        public string OperatorName { get; set; }
        public string OperatorCert { get; set; }
        public string OperatorLicenseNo { get; set; }
        public string State { get; set; }
    }

    public class Project
    {
        public int Row { get; set; }
        public string Name { get; set; }
        public string Compliant { get; set; }
        public string Notes { get; set; }
        public string DeficiencyCorrected { get; set; }
        public string ApproachToCompliance { get; set; }
        public string ImplementationCost { get; set; }
        public string Electricity { get; set; }
        public string Gas { get; set; }
        public string Oil { get; set; }
        public string Steam { get; set; }
        public string Other { get; set; }
        public string AnnualEnergySavings { get; set; }
        public string AnnualCostSavings { get; set; }
    }

    public class NYCData
    {
        public SubmittalInfo SubmittalInfo { get; set; }
        public TeamInfo TeamInfo { get; set; }
        public BuildingInfo BuildingInfo { get; set; }
        public List<Project> Projects { get; set; }
    }
}