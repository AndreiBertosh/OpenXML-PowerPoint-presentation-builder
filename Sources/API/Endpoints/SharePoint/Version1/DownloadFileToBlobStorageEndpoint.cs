using API.QueryModels;

using Application.Services.AzureServices.Interfaces;
using Application.Services.SharePointServices.Interfaces;
using Azure;
using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.SharePoint.Version1
{
    public class DownloadFileToBlobStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromBody] CopyFileQueryModel query,
            ISharePointServices sharePointServices,
            IAzureServices azureServices,
            IOptions<AzureBlobStorageSettings> settings,
            ILogger<DownloadFileToBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (query == null)
            {
                logger.LogWarning("Received null query object in CopyFile request.");
                return Results.BadRequest("Query object must not be null.");
            }

            if (string.IsNullOrWhiteSpace(query.FileName))
            {
                logger.LogWarning("Received CopyFile request with empty file name.");
                return Results.BadRequest("File name must not be empty.");
            }

            try
            {
                string folderName = query.FolderName;

                logger.LogInformation("Attempting to download file '{FileName}' from SharePoint folder '{FolderName}'.", query.FileName, folderName);

                var stream = await sharePointServices.DownloadFileByNameAsync(folderName, query.FileName, CancellationToken.None);

                if (stream != null)
                {
                    stream.Position = 0;

                    logger.LogInformation("Uploading file '{FileName}' from SharePoint to Azure Blob container '{ContainerName}'.", query.FileName, settings.Value.ContainerName);

                    await azureServices.UploadFileAsync(stream, settings.Value.ContainerName, query.FileName, cancellationToken);

                    logger.LogInformation("File '{FileName}' successfully downloaded from SharePoint and uploaded to Azure Blob Storage.", query.FileName);

                    return Results.Ok($"File '{query.FileName}' was successfully copied.");
                }

                logger.LogWarning("File '{FileName}' was not found in SharePoint folder '{FolderName}'.", query.FileName, folderName);
                return Results.BadRequest($"The file '{query.FileName}' was not found!");
            }
            catch (FileNotFoundException ex)
            {
                logger.LogWarning(ex, "FileNotFoundException while copying file '{FileName}': {Message}", query?.FileName, ex.Message);
                return Results.NotFound($"The file '{query?.FileName}' was not found in SharePoint. {ex.Message}");
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex,
                    "Client error occurred while copying file '{FileName}'. Status: {Status}, Message: {Message}",
                    query?.FileName, ex.Status, ex.Message);
                return Results.Problem(
                    $"Client error ({ex.Status}) while copying file '{query?.FileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex,
                    "Server error occurred while copying file '{FileName}'. Status: {Status}, Message: {Message}",
                    query?.FileName, ex.Status, ex.Message);
                return Results.Problem(
                    $"Server error ({ex.Status}) while copying file '{query?.FileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex,
                    "Unexpected error occurred while copying file '{FileName}': {Message}",
                    query?.FileName, ex.Message);
                return Results.Problem(
                    $"An unexpected error occurred while copying file '{query?.FileName}': {ex.Message} \n {ex.StackTrace}",
                    statusCode: 500);
            }
        }
    }
}
