using Infrastructure.Common;

using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

namespace Infrastructure.Extensions
{
    public static class ConfigurationExtension
    {
        public static void ConfigureConstants(
            this WebApplicationBuilder builder)
        {
            builder.Services.Configure<AzureBlobStorageSettings>(builder.Configuration.GetSection("AzureStorage"));
            builder.Services.Configure<AppSettings>(builder.Configuration.GetSection("App"));
            builder.Services.Configure<AzureAppSettings>(builder.Configuration.GetSection("AzureApp"));
        }

        public static void ConfigureHealthCheck(
            this WebApplicationBuilder builder)
        {
            var settings = builder.Configuration.GetSection("AzureStorage").Get<AzureBlobStorageSettings>();

            builder.Services.AddHealthChecks()
                .AddAzureBlobStorage(
                    connectionString: settings!.ConnectionString,
                    containerName: settings.ContainerName,
                    failureStatus: Microsoft.Extensions.Diagnostics.HealthChecks.HealthStatus.Unhealthy,
                    tags: ["azure", "storage", "blob"]);

        }

        public static void ConfigureGraphAuthProvider(
            this WebApplicationBuilder builder)
        {
            var settings = builder.Configuration.GetSection("AzureApp").Get<AzureAppSettings>();

            builder.Services.AddSingleton(sp =>
            {
                var provider = new Infrastructure.Services.GraphAuthProvider(
                    settings!.ClientId,
                    settings!.TenantId,
                    settings!.ClientSecret); return provider.GetGraphClient();
            });
        }

        public static void ConfigureLogging(
            this WebApplicationBuilder builder)
        {
            builder.Services.AddLogging(loggingBuilder =>
                {
                    loggingBuilder.ClearProviders();
                    loggingBuilder.AddConsole();
                    loggingBuilder.AddDebug();
                    loggingBuilder.SetMinimumLevel(LogLevel.Information);
                }
            ); 
        }
    }
}
