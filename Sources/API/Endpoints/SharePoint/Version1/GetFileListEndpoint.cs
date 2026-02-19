using Application.Services.SharePointServices.Interfaces;
using Azure;

namespace API.Endpoints.SharePoint.Version1
{
    public class GetFileListEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            ISharePointServices services,
            ILogger<GetFileListEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            try
            {
                logger.LogInformation("Starting request to get list of root folders and their files from SharePoint.");

                var result = await services.GetRootFilesAsync(cancellationToken);

                logger.LogInformation("Successfully retrieved root folders and files from SharePoint.");
                return Results.Ok(result);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex,
                    "Client error occurred while retrieving root files from SharePoint. Status: {Status}, Message: {Message}",
                    ex.Status, ex.Message);

                return Results.Problem(
                    $"Client error ({ex.Status}) while retrieving root files: {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex,
                    "Server error occurred while retrieving root files from SharePoint. Status: {Status}, Message: {Message}",
                    ex.Status, ex.Message);

                return Results.Problem(
                    $"Server error ({ex.Status}) while retrieving root files: {ex.Message}",
                    statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex,
                    "Unexpected error occurred while retrieving root files from SharePoint: {Message}", ex.Message);

                return Results.Problem(
                    $"An unexpected error occurred while retrieving root files: {ex.Message}",
                    statusCode: 500);
            }
        }
    }
}
