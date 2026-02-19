using Application.Services.AzureServices.Interfaces;
using Application.Services.SharePointServices.Interfaces;
using Azure;
using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.SharePoint.Version1
{
    public class UploadFileFromBlobStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string fileName,
            ISharePointServices sharePointServices,
            IAzureServices azureServices,
            IOptions<AzureAppSettings> settings,
            IOptions<AzureBlobStorageSettings> settingsBlob,
            ILogger<UploadFileFromBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (fileName == null)
            {
                logger.LogWarning("CopyFile request received with null query model.");
                return Results.BadRequest("Query model must not be null.");
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                logger.LogWarning("CopyFile request received with empty file name.");
                return Results.BadRequest("File name must not be empty.");
            }

            try
            {
                logger.LogInformation("Attempting to retrieve blob stream for file '{FileName}' from Azure Storage.", fileName);

                var stream = await azureServices.GetBlobStreamAsync(settingsBlob.Value.ResultContainerName, fileName, cancellationToken);

                if (stream != null)
                {
                    stream.Position = 0;
                    logger.LogInformation("Successfully retrieved blob stream for '{FileName}'. Uploading to SharePoint...", fileName);

                    var result = await sharePointServices.UploadFileAsync(settings.Value.ResultFolderName, fileName, stream, cancellationToken);

                    logger.LogInformation("File '{FileName}' uploaded to SharePoint successfully.", fileName);

                    return Results.Ok(result);
                }

                logger.LogWarning("File '{FileName}' was not found in Azure Storage during copy request.", fileName);
                return Results.BadRequest($"The file '{fileName}' was not found!");
            }
            catch (FileNotFoundException ex)
            {
                logger.LogWarning(ex, "FileNotFoundException while copying file '{FileName}': {Message}", fileName, ex.Message);
                return Results.NotFound($"The file '{fileName}' was not found in Azure Storage. {ex.Message}");
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex,
                    "Client error occurred while copying file '{FileName}'. Status: {Status}, Message: {Message}",
                    fileName, ex.Status, ex.Message);
                return Results.Problem(
                    $"Client error ({ex.Status}) while copying file '{fileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex,
                    "Server error occurred while copying file '{FileName}'. Status: {Status}, Message: {Message}",
                    fileName, ex.Status, ex.Message);
                return Results.Problem(
                    $"Server error ({ex.Status}) while copying file '{fileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex,
                    "Unexpected error occurred while copying file '{FileName}': {Message}",
                    fileName, ex.Message);
                return Results.Problem(
                    $"An unexpected error occurred while copying file '{fileName}': {ex.Message}",
                    statusCode: 500);
            }
        }
    }
}
