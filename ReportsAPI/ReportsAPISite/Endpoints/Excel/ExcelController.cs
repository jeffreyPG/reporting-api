using log4net;
using ReportsAPISite.Services.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;

namespace ReportsAPISite.Endpoints.Excel
{
    public partial class ExcelController : ApiController
    {

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private readonly ICreateExcelDocument createExcelDocument;

        public ExcelController(ICreateExcelDocument _createExcelDocument)
        {
            // IoC
            createExcelDocument = _createExcelDocument;
        }
    }
}