using Application.Services.AzureServices.Interfaces;

using Azure;

using Microsoft.AspNetCore.Mvc;

namespace API.Endpoints.Additional.Version1
{
    public class DownloadBlobFromStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string container,
            [FromRoute] string blobName,
            IAzureServices azureServices,
            ILogger<DownloadBlobFromStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(blobName))
            {
                return Results.BadRequest("Blob name is empty!");
            }

            try
            {
                // Retrieve the presentation stream from Azure Blob Storage
                var stream = await azureServices.GetBlobStreamAsync(container, blobName, cancellationToken);

                // Set MIME type for PowerPoint (.pptx)
                const string contentType = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

                // Return the stream as a downloadable file
                return Results.File(stream, contentType, blobName);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex, "Client error while downloading blob '{BlobName}': {Message}", blobName, ex.Message);
                return Results.Problem($"Blob '{blobName}' could not be retrieved. Error: {ex.Message}", statusCode: 400);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Azure error while downloading blob '{BlobName}': {Message}", blobName, ex.Message);
                return Results.Problem($"Azure error: {ex.Message}", statusCode: 500);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Unexpected error while downloading blob '{BlobName}': {Message}", blobName, ex.Message);
                return Results.Problem($"Unexpected error: {ex.Message}", statusCode: 500);
            }
        }
    }
}
