using Application.Services.SharePointServices.Interfaces;
using Azure;
using Infrastructure.Common;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.SharePoint.Version1
{
    public class DoesFileExistEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string folderName,
            [FromRoute] string fileName,
            ISharePointServices services,
            IOptions<AzureAppSettings> settings,
            ILogger<DoesFileExistEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(folderName))
            {
                logger.LogWarning("DoesFileExist request received with empty folder name.");
                return Results.Problem("Folder name must not be empty.", statusCode: 400);
            }
            if (folderName != settings.Value.ResultFolderName && folderName != settings.Value.TemplatesFolderName)
            {
                logger.LogWarning("DoesFileExist request received with invalid folder name '{FolderName}'. Allowed folders: '{Allowed1}', '{Allowed2}'.",
                    folderName, settings.Value.ResultFolderName, settings.Value.TemplatesFolderName);
                return Results.Problem(
                    $"Folder name is incorrect. Allowed: '{settings.Value.ResultFolderName}', '{settings.Value.TemplatesFolderName}'.",
                    statusCode: 400);
            }
            if (string.IsNullOrWhiteSpace(fileName))
            {
                logger.LogWarning("DoesFileExist request received with empty file name.");
                return Results.Problem("File name must not be empty.", statusCode: 400);
            }

            try
            {
                logger.LogInformation("Checking if file '{FileName}' exists in folder '{FolderName}'.", fileName, folderName);

                var doesFileExist = await services.DoesFileExistAsync(folderName, fileName, cancellationToken);

                logger.LogInformation("Existence check for file '{FileName}' in folder '{FolderName}': {Exists}.", fileName, folderName, doesFileExist);
                return Results.Ok(doesFileExist);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex,
                    "Client error occurred while checking existence of file '{FileName}' in folder '{FolderName}'. Status: {Status}, Message: {Message}",
                    fileName, folderName, ex.Status, ex.Message);
                return Results.Problem(
                    $"Client error ({ex.Status}) while checking file '{folderName}/{fileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex,
                    "Server error occurred while checking existence of file '{FileName}' in folder '{FolderName}'. Status: {Status}, Message: {Message}",
                    fileName, folderName, ex.Status, ex.Message);
                return Results.Problem(
                    $"Server error ({ex.Status}) while checking file '{folderName}/{fileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex,
                    "Unexpected error occurred while checking existence of file '{FileName}' in folder '{FolderName}': {Message}",
                    fileName, folderName, ex.Message);
                return Results.Problem(
                    $"Unexpected error while checking file '{folderName}/{fileName}': {ex.Message}",
                    statusCode: 500);
            }
        }
    }
}