using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;

namespace ReportsAPISite.Endpoints.Pdf
{
    public partial class PdfController : ApiController
    {

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public PdfController()
        {
            // TODO: ioc
        }
    }
}