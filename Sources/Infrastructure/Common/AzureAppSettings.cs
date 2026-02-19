namespace Infrastructure.Common
{
    public class AzureAppSettings
    {
        public required string ClientId { get; set; }

        public required string TenantId { get; set; }

        public required string ClientSecret { get; set; }

        public required string SiteHost { get; set; }

        public required string SitePath { get; set; }

        public required string TemplatesFolderName { get; set; }

        public required string ResultFolderName { get; set; }

        public required string WebProxyAddress { get; set; }

        public required bool UseWebProxy { get; set; }
    }
}
