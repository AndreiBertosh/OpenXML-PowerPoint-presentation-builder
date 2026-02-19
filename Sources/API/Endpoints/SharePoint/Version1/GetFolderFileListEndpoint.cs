using API.Endpoints.Common;
using API.Endpoints.SharePoint.Models;
using Application.Services.SharePointServices.Interfaces;
using Azure;
using Microsoft.AspNetCore.Mvc;

namespace API.Endpoints.SharePoint.Version1
{
    public class GetFolderFileListEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string folderName,
            ISharePointServices services,
            ILogger<GetFolderFileListEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(folderName))
            {
                logger.LogWarning("GetFilesInFolder request received with empty folder name.");
                return Results.Problem("Folder name must not be empty.", statusCode: 400);
            }

            try
            {
                logger.LogInformation("Attempting to retrieve files from folder '{FolderName}' in SharePoint.", folderName);

                var folderInfo = await services.GetFilesInFolderAsync(folderName, cancellationToken);

                logger.LogInformation("Successfully retrieved files from folder '{FolderName}' in SharePoint.", folderName);

                Dictionary<string, List<FileInfoViewModel>> result = folderInfo
                    .ToDictionary(
                        pair => pair.Key,
                        pair => pair.Value.Select(l => l.ToViewModel()).ToList()
                    );

                return Results.Ok(result);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex,
                    "Client error occurred while retrieving files from folder '{FolderName}'. Status: {Status}, Message: {Message}",
                    folderName, ex.Status, ex.Message);

                return Results.Problem(
                    $"Client error ({ex.Status}) while retrieving files from folder '{folderName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex,
                    "Server error occurred while retrieving files from folder '{FolderName}'. Status: {Status}, Message: {Message}",
                    folderName, ex.Status, ex.Message);

                return Results.Problem(
                    $"Server error ({ex.Status}) while retrieving files from folder '{folderName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (FileNotFoundException ex)
            {
                logger.LogWarning(ex,
                    "Folder '{FolderName}' was not found in SharePoint while retrieving files.", folderName);
                return Results.NotFound($"Folder '{folderName}' was not found in SharePoint.");
            }
            catch (Exception ex)
            {
                logger.LogError(ex,
                    "Unexpected error occurred while retrieving files from folder '{FolderName}': {Message}",
                    folderName, ex.Message);

                return Results.Problem(
                    $"An unexpected error occurred while retrieving files from folder '{folderName}': {ex.Message}",
                    statusCode: 500);
            }
        }
    }
}
