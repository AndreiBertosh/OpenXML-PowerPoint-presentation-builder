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
    public class CreateNewPresentationBySlideIndexesEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromBody] CreatePresentationBySlideIndexesQueryModel query,
            ISharePointServices sharePointServices,
            IPresentationActionsServices presentationActionsServices,
            IPresentationDataServices presentationServices,
            IOptions<AzureAppSettings> settings,
            ILogger<CreateNewPresentationBySlideIndexesAzureBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            IEnumerable<string>? errors;
            string className = typeof(CreateNewPresentationBySlideIndexesAzureBlobStorageEndpoint).Name;

            logger.LogInformation("Start processing request for SharePoint: {PresentationName}. Class is {ClassName}", query.SourcePresentationName, className);

            if (string.IsNullOrWhiteSpace(query.SourcePresentationName))
            {
                logger.LogError("PresentationName name cannot be empty!");
                return Results.BadRequest("PresentationName name cannot be empty!");
            }

            try
            { 
                string presentationBlobName = presentationServices.GetPresentationName(query.SourcePresentationName);

                var stream = await sharePointServices.DownloadFileByNameAsync(settings.Value.TemplatesFolderName, presentationBlobName, cancellationToken);

                using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, true))
                {
                    logger.LogInformation("Presentation document opened successfully for writing: {PresentationName}", presentationBlobName);

                    // Removing slides from the presentation for later saving
                    logger.LogInformation("Removing slides from the presentation for later saving.");
                    errors = presentationActionsServices.NewPresentationBySlideIndexes(presentationDocument, query.SlideIndexes.ToList(), query.CommentMessage, logger).Result;
                }

                stream.Position = 0;

                //string newBlobName = presentationServices.GetNewRandomPresentationName();
                string newPresentationName = query.DestinationPresentationName;

                // Upload the updated presentation to Azure Blob Storage
                string link = await sharePointServices.UploadFileAsync(settings.Value.ResultFolderName, newPresentationName, stream, cancellationToken);
                logger.LogInformation("Updated presentation successfully uploaded: {NewPresentationName}", newPresentationName);

                // Return result based on the presence of errors
                var result = new ResponseViewModel(
                    newPresentationName,
                    link,
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
                logger.LogWarning(ex, "Client error occurred while processing the request: {PresentationName}", query.SourcePresentationName);
                return Results.Problem("A client-side error occurred while processing the request. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error occurred while processing the request: {PresentationName}", query.SourcePresentationName);
                return Results.Problem("A server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An unexpected error occurred while processing the request: {PresentationName}", query.SourcePresentationName);
                return Results.Problem("An unexpected error occurred. Please try again later.", statusCode: 500);
            }
        }
    }
}
