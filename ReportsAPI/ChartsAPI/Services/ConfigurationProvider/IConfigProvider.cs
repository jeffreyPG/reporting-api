namespace ChartsAPI.Services.ConfigurationProvider
{
    public interface IConfigProvider
    {
        string ChartUrl { get; }
        int ChartsMaxAge { get; }
        string TableauUserName { get; }
        string TableauPassword { get; }
    }
}