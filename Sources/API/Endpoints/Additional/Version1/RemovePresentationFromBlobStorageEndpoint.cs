using API.QueryModels.AzureBlobStorage;
using Application.Services.AzureServices.Interfaces;

using Azure;

using Microsoft.AspNetCore.Mvc;

namespace API.Endpoints.Additional.Version1
{
    public class RemovePresentationFromBlobStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromBody] DeleteBlobsInStorageQueryModel query,
            IAzureServices azureServices,
            ILogger<RemovePresentationFromBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrEmpty(query.ContainerName))
            {
                return Results.BadRequest("Container name is empty!");
            }

            if (query.Blobs.Length == 0)
            {
                return Results.BadRequest("Blob names is empty!");
            }

            string blobName = string.Empty;

            try
            {
                List<string> results = [];

                foreach (var blob in query.Blobs)
                {
                    blobName = blob;
                    bool result = await azureServices.DeleteBlobAsync(query.ContainerName, blobName, cancellationToken);

                    if (result)
                    {
                        results.Add($"Blob '{blobName}' deleted successfully.");
                    }
                    else
                    {
                        results.Add($"Blob '{blobName}' does not exist or was already deleted.");
                    }
                }

                return Results.Ok(results);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex, "Client error while removing blob '{BlobName}': {Message}", blobName, ex.Message);
                return Results.Problem($"Blob '{blobName}' could not be retrieved. Error: {ex.Message}", statusCode: 400);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Azure error while removing blob '{BlobName}': {Message}", blobName, ex.Message);
                return Results.Problem($"Azure error: {ex.Message}", statusCode: 500);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Unexpected error while removing blob '{BlobName}': {Message}", blobName, ex.Message);
                return Results.Problem($"Unexpected error: {ex.Message}", statusCode: 500);
            }
        }

    }
}
