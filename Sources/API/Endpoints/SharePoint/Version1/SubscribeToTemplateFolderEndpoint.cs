using Application.Services.SharePointServices.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace API.Endpoints.SharePoint.Version1
{
    public class SubscribeToTemplateFolderEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromBody] string azureFunctionUrl,
            ISharePointServices sharePointServices,
            ILogger<UploadFileFromBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (!string.IsNullOrWhiteSpace(azureFunctionUrl))
            {
                var result = sharePointServices.SubscribeWebHookForFolderAsync(azureFunctionUrl, cancellationToken);
                return Results.Ok(result);
            }
            return Results.BadRequest("incorrect url");
        }
    }
}
