using System.Configuration;

namespace reports.Services.ConfigurationProviders
{
    public class ConfigProvider : IConfigProvider
    {
        public string AWSAccessKey => ConfigurationManager.AppSettings["AWSAccessKey"];
        public string AWSAccessSecretKey => ConfigurationManager.AppSettings["AWSAccessSecretKey"];

        public string ChartUrl => ConfigurationManager.AppSettings["ChartUrl"];
        public string TableauUserName => ConfigurationManager.AppSettings["TableauUserName"];
        public string TableauPassword => ConfigurationManager.AppSettings["TableauPassword"];
    }
}