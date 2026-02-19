using API.Endpoints.PresentationActions.Models;
using API.QueryModels;

using Application.Services.PresentationActions.Interfaces;
using Application.Services.PresentationActions.Models;
using Application.Services.SharePointServices.Interfaces;

using Azure;

using DocumentFormat.OpenXml.Packaging;

using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.PresentationActions.Version1
{
    public class InsertNewSlidesToPresentationEndpoint
    {

        public static async Task<IResult> ExecuteAsync(
            [FromBody] AddNewSlidesDataSharePointQueryModel query,
            ISharePointServices sharePointServices,
            IAddNewSlideToPresentationServices insertNewSlideToPresentationServices,
            IOptions<AzureAppSettings> settings,
            ILogger<InsertNewSlidesToPresentationEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            IEnumerable<string>? errors;
            string className = typeof(InsertNewSlidesToPresentationEndpoint).Name;

            logger.LogInformation("Start processing request for Presentation: {DestinationPresentationName}. Class is {ClassName}", query.DestinationPresentationName, className);

            try
            {
                if (string.IsNullOrWhiteSpace(query.DestinationPresentationName))
                {
                    logger.LogError("Presentation name cannot be empty!");
                    return Results.BadRequest("Presentation name cannot be empty!");
                }

                Stream destinationStream = null;

                if (await sharePointServices.DoesFileExistAsync(settings.Value.ResultFolderName, query.DestinationPresentationName, cancellationToken))
                {
                    destinationStream = await sharePointServices.DownloadFileByNameAsync(settings.Value.ResultFolderName, query.DestinationPresentationName, cancellationToken);
                }
                else
                {
                    destinationStream = await sharePointServices.DownloadFileByNameAsync(settings.Value.TemplatesFolderName, query.TemplatePresentationName, cancellationToken);
                }

                using (PresentationDocument presentationDocument = PresentationDocument.Open(destinationStream, true))
                {
                    logger.LogInformation("Presentation document opened successfully for writing: {PresentationName}", query.DestinationPresentationName);

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
                var link = await sharePointServices.UploadFileAsync(settings.Value.ResultFolderName, query.DestinationPresentationName, destinationStream, cancellationToken);

                logger.LogInformation("Updated presentation successfully uploaded: {NewPresentationName}", query.DestinationPresentationName);

                var result = new ResponseViewModel(
                    query.DestinationPresentationName,
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
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex, "Client error occurred while processing the request: {PresentationName}", query.DestinationPresentationName);
                return Results.Problem("A client-side error occurred while processing the request. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error occurred while processing the request: {PresentationName}", query.DestinationPresentationName);
                return Results.Problem("A server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An unexpected error occurred while processing the request: {PresentationName}", query.DestinationPresentationName);
                return Results.Problem("An unexpected error occurred. Please try again later.", statusCode: 500);
            }
        }
    }
}
