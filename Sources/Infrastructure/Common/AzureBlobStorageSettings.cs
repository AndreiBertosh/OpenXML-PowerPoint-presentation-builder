namespace Infrastructure.Common
{
    public class AzureBlobStorageSettings
    {
        public required string ConnectionString { get; set; }

        public required string ContainerName { get; set; }

        public required string ResultContainerName { get; set; }

        public required int MaxRetries { get; set; } = 5;

        public required int NetworkTimeout { get; set; } = 5;
    }
}
