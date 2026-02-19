using API.Endpoints.PresentationActions.Models;
using API.QueryModels.AzureBlobStorage;

using Application.Services.AzureServices.Interfaces;
using Application.Services.PresentationActions.Interfaces;
using Application.Services.PresentationData.Interfaces;

using Azure;

using DocumentFormat.OpenXml.Packaging;

using Infrastructure.Common;

using Microsoft.Extensions.Options;

namespace API.Endpoints.PresentationActions.Version1
{
    public class CreateNewPresentationBySlideIndexesAzureBlobStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [AsParameters] CreatePresentationBySlideIndexesQueryModel query,
            IAzureServices azureServices,
            IPresentationActionsServices presentationActionsServices,
            IPresentationDataServices presentationServices,
            IOptions<AzureBlobStorageSettings> settings,
            ILogger<CreateNewPresentationBySlideIndexesAzureBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            IEnumerable<string>? errors;
            string className = typeof(CreateNewPresentationBySlideIndexesAzureBlobStorageEndpoint).Name;

            logger.LogInformation("Start processing request for blob: {BlobName}. Class is {ClassName}", query.SourceBlobName, className);

            if (string.IsNullOrWhiteSpace(query.SourceBlobName))
            {
                logger.LogError("Blob name cannot be empty!");
                return Results.BadRequest("Blob name cannot be empty!");
            }

            try
            { 
                string presentationBlobName = presentationServices.GetPresentationName(query.SourceBlobName);

                var stream = await azureServices.GetWritableBlobStreamAsync(settings.Value.ContainerName, presentationBlobName, cancellationToken);

                using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, true))
                {
                    logger.LogInformation("Presentation document opened successfully for writing: {PresentationName}", presentationBlobName);

                    // Removing slides from the presentation for later saving
                    logger.LogInformation("Removing slides from the presentation for later saving.");
                    errors = presentationActionsServices.NewPresentationBySlideIndexes(presentationDocument, query.SlideIndexes.ToList(), query.CommentMessage, logger).Result;
                }

                stream.Position = 0;

                string newBlobName = query.DestinationBlobName;

                // Upload the updated presentation to Azure Blob Storage
                await azureServices.UploadFileAsync(stream, settings.Value.ResultContainerName, newBlobName, cancellationToken);
                logger.LogInformation("Updated presentation successfully uploaded: {NewBlobName}", newBlobName);

                // Return result based on the presence of errors
                var result = new ResponseViewModel(
                    newBlobName,
                    "",
                    errors);

                // Return result based on the presence of errors
                if (errors != null && errors.Any())
                {
                    logger.LogWarning("Slide insertion completed with warnings/errors: {Errors}", string.Join(", ", errors));
                }

                logger.LogInformation("Slide(s) have been successfully added to the presentation.");
                return Results.Ok(result);
            }
            catch (RequestFailedException ex) when(ex.Status >= 400 && ex.Status< 500)
            {
                logger.LogWarning(ex, "Client error occurred while processing the request: {BlobName}", query.SourceBlobName);
                return Results.Problem("A client-side error occurred while processing the request. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error occurred while processing the request: {BlobName}", query.SourceBlobName);
                return Results.Problem("A server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An unexpected error occurred while processing the request: {BlobName}", query.SourceBlobName);
                return Results.Problem("An unexpected error occurred. Please try again later.", statusCode: 500);
            }
        }
    }
}
