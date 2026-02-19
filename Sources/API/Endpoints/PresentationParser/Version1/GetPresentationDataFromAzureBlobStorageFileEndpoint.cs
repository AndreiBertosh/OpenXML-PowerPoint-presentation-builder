using API.Endpoints.PresentationParser.Models;

using Application.Services.AzureServices.Interfaces;
using Application.Services.PresentationParser.Interfaces;

using AutoMapper;

using DocumentFormat.OpenXml.Packaging;

using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.PresentationParser.Version1
{
    public class GetPresentationDataFromAzureBlobStorageFileEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string storageType,
            [FromRoute] string presentationName,
            IPresentationParserServices presentationParserServices,
            IAzureServices azureServices,
            IOptions<AzureBlobStorageSettings> settings,
            IMapper mapper,
            ILogger<GetPresentationDataFromAzureBlobStorageFileEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            try
            {
                string containerName = settings.Value.ContainerName;

                if (storageType == "results")
                {
                    containerName = settings.Value.ResultContainerName;
                }

                string className = typeof(GetPresentationDataFromAzureBlobStorageFileEndpoint).Name;

                logger.LogInformation("Start processing request for blob: {BlobName}. Class is {ClassName}", presentationName, className);

                // Attempt to get the presentation name
                logger.LogInformation("Extracted presentation name: {PresentationName}", presentationName);

                // Attempt to fetch the blob stream from Azure Blob Storage
                var stream = await azureServices.GetBlobStreamAsync(containerName, presentationName, cancellationToken);

                // Parse the presentation file into a list of slides
                List<SlideDataViewModel> slides = [];
                using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, false))
                {
                    logger.LogInformation("Successfully opened the presentation document for: {PresentationName}", presentationName);

                    slides = [.. presentationParserServices
                        .GetAllPresentationData(presentationDocument)
                        .Select(slide => mapper.Map<SlideDataViewModel>(slide))];

                    logger.LogInformation("Parsed {SlideCount} slides from the presentation: {PresentationName}", slides.Count, presentationName);
                }

                stream.Close();

                string link = await azureServices.GetDownloadLinkAsync(containerName, presentationName, cancellationToken);

                logger.LogInformation("Successfully processed request for blob: {BlobName}", presentationName);
                return Results.Ok(new PresentationDataViewModel(presentationName, link, slides));
            }
            catch (Exception ex)
            {
                // Log the error
                logger.LogError(ex, "An error occurred while processing blob: {BlobName}", presentationName);

                // Return a problem result for the client
                return Results.Problem(
                    detail: ex.Message,
                    statusCode: 500
                );
            }
        }
    }
}
