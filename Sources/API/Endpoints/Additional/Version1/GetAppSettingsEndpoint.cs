using Infrastructure.Common;
using Microsoft.Extensions.Options;

namespace API.Endpoints.Additional.Version1
{
    public class GetAppSettingsEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            IOptions<AzureAppSettings> settings,
            CancellationToken cancellationToken = default)
        {
            return Results.Ok(settings);
        }
    }
}
