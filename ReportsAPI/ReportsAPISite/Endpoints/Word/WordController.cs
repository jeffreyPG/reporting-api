using log4net;
using ReportsAPISite.Services.ConfigProvider;
using ReportsAPISite.Services.DocumentStorage;
using ReportsAPISite.Services.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;

namespace ReportsAPISite.Endpoints.Word
{
    public partial class WordController : ApiController
    {

        private static readonly ILog log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private readonly IDocumentStorage s3DocumentStorage;
        private readonly IDocumentStorage fileDocumentStorage;
        private readonly IConfigProvider configProvider;
        private readonly ICreateWordDocument createWordDocument;

        public WordController(IDocumentStorage _s3DocumentStorage, IDocumentStorage _fileDocumentStorage, IConfigProvider _configProvider, ICreateWordDocument _createWordDocument)
        {
            s3DocumentStorage = _s3DocumentStorage;
            fileDocumentStorage = _fileDocumentStorage;
            configProvider = _configProvider;
            createWordDocument = _createWordDocument;
        }

    }
}