using Application.Services.SharePointServices.Interfaces;
using Azure;
using Microsoft.AspNetCore.Mvc;

namespace API.Endpoints.SharePoint.Version1
{
    public class GetPresentationDownloadLinkEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string fileName,
            ISharePointServices services,
            ILogger<GetPresentationDownloadLinkEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                logger.LogWarning("GetFileDownloadLink request received with empty file name.");
                return Results.Problem("File name must not be empty.", statusCode: 400);
            }

            try
            {
                logger.LogInformation("Attempting to retrieve download link for file '{FileName}' from SharePoint.", fileName);

                var result = await services.GetFileDownloadLinkAsync(fileName, cancellationToken);

                logger.LogInformation("Successfully retrieved download link for file '{FileName}' from SharePoint.", fileName);

                return Results.Ok(result);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex,
                    "Client error occurred while retrieving download link for file '{FileName}'. Status: {Status}, Message: {Message}",
                    fileName, ex.Status, ex.Message);

                return Results.Problem(
                    $"Client error ({ex.Status}) while retrieving download link for file '{fileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex,
                    "Server error occurred while retrieving download link for file '{FileName}'. Status: {Status}, Message: {Message}",
                    fileName, ex.Status, ex.Message);

                return Results.Problem(
                    $"Server error ({ex.Status}) while retrieving download link for file '{fileName}': {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (FileNotFoundException ex)
            {
                logger.LogWarning(ex,
                    "File '{FileName}' was not found in SharePoint while retrieving download link.", fileName);
                return Results.NotFound($"File '{fileName}' was not found in SharePoint.");
            }
            catch (Exception ex)
            {
                logger.LogError(ex,
                    "Unexpected error occurred while retrieving download link for file '{FileName}': {Message}",
                    fileName, ex.Message);

                return Results.Problem(
                    $"An unexpected error occurred while retrieving download link for file '{fileName}': {ex.Message}",
                    statusCode: 500);
            }
        }
    }
}
