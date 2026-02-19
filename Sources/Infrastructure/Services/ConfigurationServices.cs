using Infrastructure.Common;
using Infrastructure.Extensions.Interfaces;

namespace Infrastructure.Services
{
    public class ConfigurationServices : IServiceCollectionExtensions
    {
        private readonly AzureBlobStorageSettings _settings;

        public ConfigurationServices(AzureBlobStorageSettings settings)
        {
            _settings = settings;
        }

        public string AzureBlobStorageConnectionString()
        {
            return _settings.ConnectionString;
        }
    }
}
