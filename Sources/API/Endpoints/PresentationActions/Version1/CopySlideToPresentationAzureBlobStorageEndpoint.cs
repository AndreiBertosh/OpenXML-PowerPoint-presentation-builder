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
    public class CopySlideToPresentationAzureBlobStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [AsParameters] CopySlideQueryModel query,
            IAzureServices azureServices,
            ICopySlideServices copySlideServices,
            IPresentationDataServices presentationServices,
            IOptions<AzureBlobStorageSettings> settings,
            ILogger<CopySlideToPresentationAzureBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            string className = typeof(CopySlideToPresentationAzureBlobStorageEndpoint).Name;

            logger.LogInformation("Start processing request for copy data from {SourceBlobName} to {DestinationBlobName}. Class is {ClassName}",
                query.SourceBlobName, query.DestinationBlobName, className);

            if (string.IsNullOrWhiteSpace(query.SourceBlobName))
            {
                return Results.BadRequest("Source Blob name cannot be empty!");
            }

            if (string.IsNullOrWhiteSpace(query.TemplateBlobName))
            {
                return Results.BadRequest("Template Blob name cannot be empty!");
            }

            string storage = settings.Value.ContainerName;

            if (query.SourceStorageType?.ToLower() == "results")
            {
                storage = settings.Value.ResultContainerName;
            }

            try
            {
                await using var sourceStream = await azureServices.GetBlobStreamAsync(storage, query.SourceBlobName, cancellationToken);

                Stream? destinationStream = null;

                if (await azureServices.DoesBlobExistAsync(settings.Value.ResultContainerName, query.DestinationBlobName, cancellationToken))
                {
                    destinationStream = await azureServices.GetWritableBlobStreamAsync(settings.Value.ResultContainerName, query.DestinationBlobName, cancellationToken);
                }
                else
                {
                    destinationStream = await azureServices.GetWritableBlobStreamAsync(settings.Value.ContainerName, query.TemplateBlobName, cancellationToken);
                }

                string commentMessage = query.CommentMessage;
                IEnumerable<string>? errors;

                using var destinationPresentationDocument = PresentationDocument.Open(destinationStream, true);
                using var sourcePresentationDocument = PresentationDocument.Open(sourceStream, false);

                logger.LogInformation("Presentation documents opened successfully.");

                errors = await copySlideServices.CopySlides(sourcePresentationDocument, destinationPresentationDocument, query.SlideIndexes, commentMessage, logger);

                destinationPresentationDocument.Save();
                destinationStream.Position = 0;

                string newBlobName = string.IsNullOrWhiteSpace(query.DestinationBlobName)
                    ? presentationServices.GetNewRandomPresentationName()
                    : query.DestinationBlobName;

                await azureServices.UploadFileAsync(destinationStream, settings.Value.ResultContainerName, newBlobName, cancellationToken);
                logger.LogInformation("Updated presentation uploaded to blob: {NewBlobName}", newBlobName);

                var result = new ResponseViewModel(newBlobName, "", errors);

                if (errors != null && errors.Any())
                {
                    logger.LogWarning("Slide insertion completed with warnings/errors: {Errors}", string.Join(", ", errors));
                }

                logger.LogInformation("Slide(s) successfully added to presentation.");

                // 🔄 Run Validation and logging
                //_ = Task.Run(() =>
                //{
                //    try
                //    {
                //        var validationErrors = PresentationCommonServices.ValidateDocument(destinationPresentationDocument, logger);
                //        if (validationErrors.Any())
                //        {
                //            logger.LogWarning("Background validation found {Count} issues.", validationErrors.Count());
                //        }
                //        else
                //        {
                //            logger.LogInformation("Background validation completed with no issues.");
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //        logger.LogError(ex, "Background validation failed unexpectedly.");
                //    }
                //}, cancellationToken);

                return Results.Ok(result);
            }
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex, "Client error while processing request: {SourceBlobName} → {DestinationBlobName}",
                    query.SourceBlobName, query.DestinationBlobName);
                return Results.Problem("Client-side error occurred. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error while processing request: {SourceBlobName} → {DestinationBlobName}",
                    query.SourceBlobName, query.DestinationBlobName);
                return Results.Problem("Server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Unexpected error while processing request: {SourceBlobName} → {DestinationBlobName}",
                    query.SourceBlobName, query.DestinationBlobName);
                return Results.Problem("Unexpected error occurred. Please try again later.", statusCode: 500);
            }
        }
    }
}
