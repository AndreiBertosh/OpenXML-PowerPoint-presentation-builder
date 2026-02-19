using Application.Services.SharePointServices.Interfaces;
using Azure;
using Infrastructure.Common;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.SharePoint.Version1
{
    public class DeleteFileEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string folderName,
            [FromRoute] string fileName,
            ISharePointServices services,
            IOptions<AzureAppSettings> settings,
            ILogger<DeleteFileEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(folderName))
            {
                logger.LogWarning("Delete request received with empty folder name.");
                return Results.Problem("Folder name must not be empty.", statusCode: 400);
            }

            if (folderName != settings.Value.ResultFolderName && folderName != settings.Value.TemplatesFolderName)
            {
                logger.LogWarning("Delete request received with incorrect folder name.");
                return Results.Problem("Folder name is incorrect.", statusCode: 400);
            }

            if (string.IsNullOrWhiteSpace(fileName))
            {
                logger.LogWarning("Delete request received with empty file name.");
                return Results.Problem("File name must not be empty.", statusCode: 400);
            }

            try
            {
                logger.LogInformation("Attempting to delete file '{FileName}' from folder '{FolderName}'.", fileName, folderName);

                await services.DeleteFileAsync(folderName, fileName, cancellationToken);

                logger.LogInformation("File '{FileName}' from folder '{FolderName}' deleted successfully.", fileName, folderName);
                return Results.Ok($"File '{folderName}/{fileName}' was deleted.");
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex,
                    "Client error occurred while deleting file '{FileName}' from folder '{FolderName}'. Status: {Status}, Message: {Message}",
                    fileName, folderName, ex.Status, ex.Message);

                return Results.Problem(
                    $"Client error ({ex.Status}) while deleting file '{folderName}/{fileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex,
                    "Server error occurred while deleting file '{FileName}' from folder '{FolderName}'. Status: {Status}, Message: {Message}",
                    fileName, folderName, ex.Status, ex.Message);

                return Results.Problem(
                    $"Server error ({ex.Status}) while deleting file '{folderName}/{fileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (FileNotFoundException ex)
            {
                logger.LogWarning(ex, "File '{FileName}' not found in folder '{FolderName}' during delete request.", fileName, folderName);
                return Results.NotFound($"File '{folderName}/{fileName}' not found.");
            }
            catch (Exception ex)
            {
                logger.LogError(ex,
                    "An unexpected error occurred while deleting file '{FileName}' from folder '{FolderName}': {Message}",
                    fileName, folderName, ex.Message);

                return Results.Problem(
                    $"Unexpected error while deleting file '{folderName}/{fileName}': {ex.Message}",
                    statusCode: 500);
            }
        }
    }
}
