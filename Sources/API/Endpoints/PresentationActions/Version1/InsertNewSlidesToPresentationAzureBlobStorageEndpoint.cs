using API.Endpoints.PresentationActions.Models;
using API.QueryModels.AzureBlobStorage;

using Application.Services.AzureServices.Interfaces;
using Application.Services.PresentationActions.Interfaces;
using Application.Services.PresentationActions.Models;

using Azure;

using DocumentFormat.OpenXml.Packaging;

using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.PresentationActions.Version1
{
    public class InsertNewSlidesToPresentationAzureBlobStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromBody] AddNewSlidesDataQueryModel query,
            IAzureServices azureServices,
            IAddNewSlideToPresentationServices insertNewSlideToPresentationServices,
            IOptions<AzureBlobStorageSettings> settings,
            ILogger<InsertNewSlideToPresentationAzureBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            IEnumerable<string>? errors;
            string className = typeof(InsertNewSlideToPresentationAzureBlobStorageEndpoint).Name;

            logger.LogInformation("Start processing request for blob: {DestinationBlobName}. Class is {ClassName}", query.DestinationBlobName, className);

            try
            {
                if (string.IsNullOrWhiteSpace(query.DestinationBlobName))
                {
                    logger.LogError("Blob name cannot be empty!");
                    return Results.BadRequest("Blob name cannot be empty!");
                }

                Stream? destinationStream = null;

                if (await azureServices.DoesBlobExistAsync(settings.Value.ResultContainerName, query.DestinationBlobName, cancellationToken))
                {
                    destinationStream = await azureServices.GetWritableBlobStreamAsync(settings.Value.ResultContainerName, query.DestinationBlobName, cancellationToken);
                }
                else
                {
                    destinationStream = await azureServices.GetWritableBlobStreamAsync(settings.Value.ContainerName, query.TemplateBlobName, cancellationToken);
                }

                using (PresentationDocument presentationDocument = PresentationDocument.Open(destinationStream, true))
                {
                    logger.LogInformation("Presentation document opened successfully for writing: {PresentationName}", query.DestinationBlobName);

                    var slidesData = query.SlidesData.Select(s =>
                        new NewSlideData(
                            s.ThemeName,
                            s.LayoutName,
                            s.Title,
                            s.SubTitle,
                            s.BodyText,
                            string.Empty,
                            s.SlideComment))
                        .ToArray();

                    // Insert the new slide
                    logger.LogInformation("Attempting to insert a new slide into the presentation.");
                    errors = insertNewSlideToPresentationServices.AddNewSlidesByLayout(presentationDocument, slidesData, logger);
                }

                destinationStream.Position = 0;

                // Upload the updated presentation to Azure Blob Storage
                await azureServices.UploadFileAsync(destinationStream, settings.Value.ResultContainerName, query.DestinationBlobName,cancellationToken);

                logger.LogInformation("Updated presentation successfully uploaded: {NewBlobName}", query.DestinationBlobName);

                var result = new ResponseViewModel(
                    query.DestinationBlobName,
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
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex, "Client error occurred while processing the request: {BlobName}", query.DestinationBlobName);
                return Results.Problem("A client-side error occurred while processing the request. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error occurred while processing the request: {BlobName}", query.DestinationBlobName);
                return Results.Problem("A server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An unexpected error occurred while processing the request: {BlobName}", query.DestinationBlobName);
                return Results.Problem("An unexpected error occurred. Please try again later.", statusCode: 500);
            }
        }
    }
}
