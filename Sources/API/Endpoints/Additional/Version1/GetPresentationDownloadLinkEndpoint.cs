using Application.Services.AzureServices.Interfaces;

using Azure;

using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.Additional.Version1
{
    public class GetpresentationDownloadLinkEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string blobName,
            IAzureServices azureServices,
            IOptions<AzureBlobStorageSettings> settings,
            ILogger<GetpresentationDownloadLinkEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            try
            {
                string link = await azureServices.GetDownloadLinkAsync(settings.Value.ResultContainerName, blobName, cancellationToken);

                logger.LogInformation("The Link for download presentation have been successfully generated.");
                return Results.Ok(link);
            }
            catch (RequestFailedException ex) when(ex.Status >= 400 && ex.Status< 500)
            {
                logger.LogWarning(ex, "Client error occurred while processing the request: {blobName} ", blobName);
                return Results.Problem("A client-side error occurred while processing the request. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error occurred while processing the request: {blobName} ", blobName);
                return Results.Problem("A server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An unexpected error occurred while processing the request: {blobName} ", blobName);
                return Results.Problem("An unexpected error occurred. Please try again later.", statusCode: 500);
            }

        }
    }
}
