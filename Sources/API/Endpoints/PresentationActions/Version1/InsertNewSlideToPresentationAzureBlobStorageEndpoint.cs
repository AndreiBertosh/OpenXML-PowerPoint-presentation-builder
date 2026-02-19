using API.Endpoints.PresentationActions.Models;
using API.QueryModels.AzureBlobStorage;

using Application.Services.AzureServices.Interfaces;
using Application.Services.PresentationActions.Interfaces;
using Application.Services.PresentationActions.Models;

using Azure;

using DocumentFormat.OpenXml.Packaging;

using Infrastructure.Common;

using Microsoft.Extensions.Options;

namespace API.Endpoints.PresentationActions.Version1
{
    public class InsertNewSlideToPresentationAzureBlobStorageEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [AsParameters] AddNewSlideDataQueryModel query,
            IAzureServices azureServices,
            IAddNewSlideToPresentationServices insertNewSlideToPresentationServices,
            IOptions<AzureBlobStorageSettings> settings,
            ILogger<InsertNewSlideToPresentationAzureBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            IEnumerable<string>? errors;
            string className = typeof(InsertNewSlideToPresentationAzureBlobStorageEndpoint).Name;

            logger.LogInformation("Start processing request for blob: {BlobName}. Class is {ClassName}", query.BlobName, className);

            try
            {
                if (string.IsNullOrWhiteSpace(query.BlobName) && string.IsNullOrWhiteSpace(query.TemplateBlobName))
                {
                    logger.LogError("Blob name and Template Blob Name cannot be empty!");
                    return Results.BadRequest("Blob name and Template Blob Name cannot be empty!");
                }

                Stream? stream = null;
                string presentation = string.Empty;

                if (await azureServices.DoesBlobExistAsync(settings.Value.ResultContainerName, query.BlobName, cancellationToken))
                {
                    stream = await azureServices.GetWritableBlobStreamAsync(settings.Value.ResultContainerName, query.BlobName, cancellationToken);
                    presentation = query.BlobName;
                }
                else
                {
                    stream = await azureServices.GetWritableBlobStreamAsync(settings.Value.ContainerName, query.TemplateBlobName, cancellationToken);
                    presentation = query.TemplateBlobName;
                }

                using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, true))
                {
                    logger.LogInformation("Presentation document opened successfully for writing: {PresentationName}", presentation);

                    NewSlideData slideData;

                    slideData = new(
                        query.ThemeName,
                        query.LayoutName,
                        query.Title.ToUpper(),
                        query.SubTitle.ToUpper(),
                        query.BodyText,
                        string.Empty,
                        query.SlideComment);

                    int position = (query.Position != null && query.Position > 0) ? position = (int)query.Position : 0;

                    // Insert the new slide
                    logger.LogInformation("Attempting to insert a new slide into the presentation.");
                    errors = insertNewSlideToPresentationServices.AddNewSlideByLayout(presentationDocument, slideData, logger);
                }

                stream.Position = 0;
                string newBlobName = query.BlobName;

                // Upload the updated presentation to Azure Blob Storage
                await azureServices.UploadFileAsync(stream, settings.Value.ResultContainerName, newBlobName, cancellationToken);

                logger.LogInformation("Updated presentation successfully uploaded: {NewBlobName}", newBlobName);

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
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex, "Client error occurred while processing the request: {BlobName}", query.BlobName);
                return Results.Problem("A client-side error occurred while processing the request. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error occurred while processing the request: {BlobName}", query.BlobName);
                return Results.Problem("A server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An unexpected error occurred while processing the request: {BlobName}", query.BlobName);
                return Results.Problem("An unexpected error occurred. Please try again later.", statusCode: 500);
            }
        }
    }
}
