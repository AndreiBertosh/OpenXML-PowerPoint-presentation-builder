using API.Endpoints.PresentationActions.Models;
using API.QueryModels;

using Application.Services.PresentationActions.Interfaces;
using Application.Services.PresentationData.Interfaces;
using Application.Services.SharePointServices.Interfaces;

using Azure;

using DocumentFormat.OpenXml.Packaging;

using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.PresentationActions.Version1
{
    public class CopySlideToPresentationEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromBody] CopySlideQueryModel query,
            ISharePointServices sharePointServices,
            ICopySlideServices copySlideServices,
            IOptions<AzureAppSettings> settings,
            IPresentationDataServices presentationServices,
            ILogger<CopySlideToPresentationAzureBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            string className = typeof(CopySlideToPresentationAzureBlobStorageEndpoint).Name;

            logger.LogInformation("Start processing request for copy data from {SourcePresentationName} to {DestinationPresentationName}. Class is {ClassName}",
                query.SourcePresentationName, query.DestinationPresentationName, className);

            if (string.IsNullOrWhiteSpace(query.SourcePresentationName))
                return Results.BadRequest("Source Presentation name cannot be empty!");

            if (string.IsNullOrWhiteSpace(query.TemplatePresentationName))
                return Results.BadRequest("Template Presentation name cannot be empty!");

            try
            {
                await using var sourceStream = await sharePointServices.DownloadFileByNameAsync(settings.Value.TemplatesFolderName, query.SourcePresentationName, cancellationToken);

                Stream? destinationStream = null;

                if (await sharePointServices.DoesFileExistAsync(settings.Value.ResultFolderName, query.DestinationPresentationName, cancellationToken))
                {
                    destinationStream = await sharePointServices.DownloadFileByNameAsync(settings.Value.ResultFolderName, query.DestinationPresentationName, cancellationToken);
                }
                else
                {
                    destinationStream = await sharePointServices.DownloadFileByNameAsync(settings.Value.TemplatesFolderName, query.TemplatePresentationName, cancellationToken);
                }

                string commentMessage = query.CommentMessage;
                IEnumerable<string>? errors;

                using var destinationPresentationDocument = PresentationDocument.Open(destinationStream, true);
                using var sourcePresentationDocument = PresentationDocument.Open(sourceStream, true);

                logger.LogInformation("Presentation documents opened successfully.");

                errors = await copySlideServices.CopySlides(sourcePresentationDocument, destinationPresentationDocument, query.SlideIndexes, commentMessage, logger);

                destinationPresentationDocument.Save();
                destinationStream.Position = 0;

                string newPresentationName = string.IsNullOrWhiteSpace(query.DestinationPresentationName)
                    ? presentationServices.GetNewRandomPresentationName()
                    : query.DestinationPresentationName;

                string link = await sharePointServices.UploadFileAsync("", newPresentationName, destinationStream, cancellationToken);
                logger.LogInformation("Updated presentation uploaded to blob: {NewPresentationName}", newPresentationName);

                var result = new ResponseViewModel(newPresentationName, link, errors);

                if (errors != null && errors.Any())
                {
                    logger.LogWarning("Slide insertion completed with warnings/errors: {Errors}", string.Join(", ", errors));
                }

                logger.LogInformation("Slide(s) successfully added to presentation.");

                // Run Validation and logging
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
                logger.LogWarning(ex, "Client error while processing request: {SourcePresentationName} → {DestinationPresentationName}",
                    query.SourcePresentationName, query.DestinationPresentationName);
                return Results.Problem("Client-side error occurred. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error while processing request: {SourcePresentationName} → {DestinationPresentationName}",
                    query.SourcePresentationName, query.DestinationPresentationName);
                return Results.Problem("Server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Unexpected error while processing request: {SourcePresentationName} → {DestinationPresentationName}",
                    query.SourcePresentationName, query.DestinationPresentationName);
                return Results.Problem("Unexpected error occurred. Please try again later.", statusCode: 500);
            }
        }
    }
}
