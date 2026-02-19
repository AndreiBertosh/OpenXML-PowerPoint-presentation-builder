using Asp.Versioning.Builder;
using Asp.Versioning.Conventions;

namespace API.Endpoints.Common
{
    public static class ApiVersions
    {
        public const string ApiPrefix = "/v{version:apiVersion}";

        public const double Version1_0 = 1.0;

        public const double Version2_0 = 2.0;

        public const string Name = "PresentationBuilder";

        public static ApiVersionSet VersionSet { get; private set; } = null!;

        public static WebApplication RegisterVersionSet(this WebApplication app)
        {
            VersionSet = app.NewApiVersionSet(Name)
            .ReportApiVersions()
            .HasApiVersion(Version1_0)
            .HasApiVersion(Version2_0)
            .Build();

            return app;
        }
    }
}
