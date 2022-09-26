using reports.Services.ConfigurationProviders;
using reports.Services.DocumentStorage;
using reports.Services.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Http;

namespace reports.Endpoints.CreateWordDocumentController
{
    public partial class CreateWordDocumentController : ApiController
    {

        protected IDocumentStorage s3DocumentStorage;
        protected IDocumentStorage fileDocumentStorage;
        protected IConfigProvider configProvider;
        protected ICreateWordDocument createDocumentService;

        public CreateWordDocumentController()
        {
            // TODO: IoC
            configProvider = new ConfigProvider();
            s3DocumentStorage = new S3DocumentStorage(configProvider.AWSAccessKey, configProvider.AWSAccessSecretKey, "buildee-test"); // TODO: add bucket as config driven
            fileDocumentStorage = new FileDocumentStorage(HttpContext.Current.Server.MapPath("~/GeneratedDocs"));
            createDocumentService = new CreateDocumentService();
        }

    }
}