using Application.Services.AzureServices.Interfaces;

using Azure;
using Infrastructure.Common;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.Additional.Version1
{
    public class GetListBlobsInStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromQuery] string storageType,
            IAzureServices azureServices,
            IOptions<AzureBlobStorageSettings> settings,
            ILogger<GetListBlobsInStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            string storage = settings.Value.ContainerName;

            if (storageType == "results")
            {
                storage = settings.Value.ResultContainerName;
            }

            try
            {
                // Logging before making the request
                logger.LogInformation("Starting request to get list of blobs from storage: {Storage}", storage);

                var result = await azureServices.GetBlobsAsync(storage, cancellationToken);

                logger.LogInformation("Successfully retrieved {Count} blobs from storage '{Storage}'.", result?.Count() ?? 0, storage);

                return Results.Ok(result);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                // Client-side error (HTTP 4xx)
                logger.LogWarning(ex,
                    "Client request failed while getting blob list for storage '{Storage}'. Status: {Status}, Message: {Message}",
                    storage, ex.Status, ex.Message);
                return Results.Problem(
                    $"Could not load blob list for storage '{storage}'. Client error {ex.Status}: {ex.Message}",
                    statusCode: 400);
            }
            catch (RequestFailedException ex)
            {
                // Azure/server-side error
                logger.LogError(ex,
                    "Azure service error while getting blob list for storage '{Storage}'. Status: {Status}, Message: {Message}",
                    storage, ex.Status, ex.Message);
                return Results.Problem(
                    $"Azure error while loading blob list for storage '{storage}'. Server error {ex.Status}: {ex.Message}",
                    statusCode: 500);
            }
            catch (Exception ex)
            {
                // Unexpected or unhandled errors
                logger.LogError(ex,
                    "Unexpected error while getting blob list for storage '{Storage}'. Message: {Message}",
                    storage, ex.Message);
                return Results.Problem(
                    $"An unexpected error occurred when loading blob list for storage '{storage}'. Error: {ex.Message}",
                    statusCode: 500);
            }
        }
    }
}