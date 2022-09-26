using System;
using ReportsAPISite.Services.ConfigProvider;
using ReportsAPISite.Services.ResourceProvider;
using DbUp;
using DbUp.Engine;
using Serilog;

namespace ReportsAPISite.Services
{
    public class DbUpProcessor
    {
        private readonly IConfigProvider _configProvider;
        private readonly IResourceProvider _resourceProvider;
        private readonly bool _shouldCreateDatabaseIfNeeded;

        public DbUpProcessor(IConfigProvider configProvider, IResourceProvider resourceProvider, bool shouldCreateDatabaseIfNeeded)
        {
            _configProvider = configProvider;
            _resourceProvider = resourceProvider;
            _shouldCreateDatabaseIfNeeded = shouldCreateDatabaseIfNeeded;
        }

        public void EnsureDbVersion()
        {
            var connectionString = _configProvider.DbConnectionString;
            var databaseProvided = !string.IsNullOrEmpty(connectionString);
            if (databaseProvided)
            {
                try
                {
                    if (_shouldCreateDatabaseIfNeeded)
                    {
                        EnsureDatabase.For.PostgresqlDatabase(_configProvider.DbConnectionString);
                    }
                    UpdateDatabase(connectionString);
                }
                catch (Exception ex)
                {
                    Log.Error($"DbUp failure: {ex.Message}");
                    throw;
                }
            }
        }

        private void UpdateDatabase(string connectionString)
        {
            var scriptsFound = _resourceProvider.GetAllFilesOfTypeInFolder("ReportsAPISite.Services.Repository.schema", ".sql");
            var dbUpScripts = new SqlScript[scriptsFound.Count];
            for (int i = 0; i < scriptsFound.Count; i++)
            {
                dbUpScripts[i] = new SqlScript(scriptsFound[i].Item1, scriptsFound[i].Item2);
            }

            var upgrader = DeployChanges.To
                .PostgresqlDatabase(connectionString)
                .WithScripts(dbUpScripts)
                .LogToConsole()
                .Build();

            var result = upgrader.PerformUpgrade();

            if (!result.Successful)
            {
                throw result.Error;
            }
        }

    }
}