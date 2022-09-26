namespace ReportsAPISite.Services.ConfigProvider
{
    public interface IConfigProvider
    {

        string AWSAccessKey { get; }
        string AWSAccessSecretKey { get; }
        string AWSS3BucketName { get; }

        string ChartUrl { get; }

        string DbConnectionString { get; }
        string EnvironmentName { get; }

        string FileDocumentStorageDirectory { get; }

        string LoggingMinimumLevel { get; }
        string LogglyToken { get; }
        string ReportsAPISiteKey { get; }
        string SiteName { get; }

        string TableauUserName { get; }
        string TableauPassword { get; }

    }
}