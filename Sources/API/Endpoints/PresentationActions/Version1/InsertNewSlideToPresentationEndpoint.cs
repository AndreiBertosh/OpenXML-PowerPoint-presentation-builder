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
    public class InsertNewSlideToPresentationEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromBody] AddNewSlideDataQueryModel query,
            ISharePointServices sharePointServices,
            IOptions<AzureAppSettings> settings,
            IAddNewSlideToPresentationServices insertNewSlideToPresentationServices,
            ILogger<InsertNewSlideToPresentationAzureBlobStorageEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            IEnumerable<string>? errors;
            string className = typeof(InsertNewSlideToPresentationAzureBlobStorageEndpoint).Name;

            logger.LogInformation("Start processing request for Presentation: {PresentationName}. Class is {ClassName}", query.PresentationName, className);

            try
            {
                if (string.IsNullOrWhiteSpace(query.PresentationName) && string.IsNullOrWhiteSpace(query.TemplatePresentationName))
                {
                    logger.LogError("Presentation name and Template Presentation Name cannot be empty!");
                    return Results.BadRequest("Presentation name and Template Presentation Name cannot be empty!");
                }

                Stream? stream = null;
                string presentation = string.Empty;

                if (await sharePointServices.DoesFileExistAsync(settings.Value.ResultFolderName, query.PresentationName, cancellationToken))
                {
                    stream = await sharePointServices.DownloadFileByNameAsync(settings.Value.ResultFolderName, query.PresentationName, cancellationToken);
                    presentation = query.PresentationName;
                }
                else
                {
                    stream = await sharePointServices.DownloadFileByNameAsync(settings.Value.TemplatesFolderName, query.TemplatePresentationName, cancellationToken);
                    presentation = query.TemplatePresentationName;
                }

                using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, true))
                {
                    logger.LogInformation("Presentation document opened successfully for writing: {PresentationName}", presentation);

                    NewSlideData slideData;
                    slideData = new(
                        query.ThemeName,
                        query.LayoutName,
                        query.Title,
                        query.SubTitle,
                        query.BodyText,
                        string.Empty,
                        query.SlideComment);

                    int position = (query.Position != null && query.Position > 0) ? position = (int)query.Position : 0;

                    // Insert the new slide
                    logger.LogInformation("Attempting to insert a new slide into the presentation.");
                    errors = insertNewSlideToPresentationServices.AddNewSlideByLayout(presentationDocument, slideData, logger);
                }

                stream.Position = 0;
                string newPresentationName = query.PresentationName;

                // Upload the updated presentation to SharePoint
                string link = await sharePointServices.UploadFileAsync("", newPresentationName, stream, cancellationToken);

                logger.LogInformation("Updated presentation successfully uploaded: {NewPresentationName}", newPresentationName);

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
            catch (RequestFailedException ex) when (ex.Status >= 400 && ex.Status < 500)
            {
                logger.LogWarning(ex, "Client error occurred while processing the request: {PresentationName}", query.PresentationName);
                return Results.Problem("A client-side error occurred while processing the request. Please check the input and try again.", statusCode: ex.Status);
            }
            catch (RequestFailedException ex)
            {
                logger.LogError(ex, "Server error occurred while processing the request: {PresentationName}", query.PresentationName);
                return Results.Problem("A server error occurred. Please try again later.", statusCode: ex.Status);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "An unexpected error occurred while processing the request: {BlobName}", query.PresentationName);
                return Results.Problem("An unexpected error occurred. Please try again later.", statusCode: 500);
            }
        }
    }
}
