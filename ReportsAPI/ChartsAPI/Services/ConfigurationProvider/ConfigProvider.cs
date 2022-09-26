using System.Configuration;

namespace ChartsAPI.Services.ConfigurationProvider
{
    public class ConfigProvider : IConfigProvider
    {
        public string ChartUrl => ConfigurationManager.AppSettings["ChartUrl"];
        public int ChartsMaxAge => int.Parse(ConfigurationManager.AppSettings["ChartsMaxAge"]);
        public string TableauUserName => ConfigurationManager.AppSettings["TableauUserName"];
        public string TableauPassword => ConfigurationManager.AppSettings["TableauPassword"];
    }
}