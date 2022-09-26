using System.Configuration;

namespace ReportsAPISite.Services.ConfigProvider
{
    public class WebConfigProvider : IConfigProvider
    {

        public string AWSAccessKey => ConfigurationManager.AppSettings["AWSAccessKey"];
        public string AWSAccessSecretKey => ConfigurationManager.AppSettings["AWSAccessSecretKey"];
        public string AWSS3BucketName => ConfigurationManager.AppSettings["AWSS3BucketName"];

        public string ChartUrl => ConfigurationManager.AppSettings["ChartUrl"];
                
        public string DbConnectionString => ConfigurationManager.AppSettings["DbConnectionString"];
        public string EnvironmentName => ConfigurationManager.AppSettings["EnvironmentName"];
        public string LoggingMinimumLevel => ConfigurationManager.AppSettings["LoggingMinimumLevel"];
        public string LogglyToken => ConfigurationManager.AppSettings["LogglyToken"];
        public string ReportsAPISiteKey => ConfigurationManager.AppSettings["ReportsAPISiteKey"];
        public string SiteName => ConfigurationManager.AppSettings["SiteName"];
                
        public string TableauUserName => ConfigurationManager.AppSettings["TableauUserName"];
        public string TableauPassword => ConfigurationManager.AppSettings["TableauPassword"];

        public string FileDocumentStorageDirectory => ConfigurationManager.AppSettings["FileDocumentStorageDirectory"];

    }
}