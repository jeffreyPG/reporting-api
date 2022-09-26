
namespace reports.Services.ConfigurationProviders
{
    public interface IConfigProvider
    {
        string AWSAccessKey { get; }
        string AWSAccessSecretKey { get; }

        string ChartUrl { get; }
        string TableauUserName { get; }
        string TableauPassword { get; }

    }
}