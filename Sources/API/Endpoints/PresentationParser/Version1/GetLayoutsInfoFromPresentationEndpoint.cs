using API.Endpoints.PresentationParser.Models;

using Application.Services.AzureServices.Interfaces;
using Application.Services.PresentationData.Interfaces;
using Application.Services.PresentationParser.Interfaces;

using AutoMapper;

using DocumentFormat.OpenXml.Packaging;

using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.PresentationParser.Version1
{
    public class GetLayoutsInfoFromPresentationEndpoint
    {
        public static async Task<IResult> ExecuteAsync(
            [FromRoute] string blobName,
            IPresentationParserServices presentationParserServices,
            IAzureServices azureServices,
            IPresentationDataServices presentationServices,
            IOptions<AzureBlobStorageSettings> settings,
            IMapper mapper,
            ILogger<GetLayoutsInfoFromPresentationEndpoint> logger,
            CancellationToken cancellationToken = default)
        {
            try
            {
                string className = typeof(GetLayoutsInfoFromPresentationEndpoint).Name;

                logger.LogInformation("Start processing request for blob: {BlobName}. Class is {ClassName}", blobName, className);

                // Attempt to get the presentation name
                var presentationName = presentationServices.GetPresentationName(blobName);
                logger.LogInformation("Extracted presentation name: {PresentationName}", presentationName);

                // Attempt to fetch the blob stream from Azure Blob Storage
                var stream = await azureServices.GetBlobStreamAsync(settings.Value.ContainerName, presentationName, cancellationToken);

                // Parse the presentation file into a list of slides
                List<SlideMasterInfoViewModel> slideMasterStructure = [];

                using (PresentationDocument presentationDocument = PresentationDocument.Open(stream, false))
                {
                    logger.LogInformation("Successfully opened the presentation document for: {PresentationName}", presentationName);

                    slideMasterStructure = [.. presentationParserServices
                        .AnalyzePresentationLayouts(presentationDocument)
                        .Select(slideMasterData => mapper.Map<SlideMasterInfoViewModel>(slideMasterData))];

                    logger.LogInformation("Parsed {SlideCount} slides from the presentation: {PresentationName}", slideMasterStructure.Count, presentationName);
                }

                stream.Close();

                string link = await azureServices.GetDownloadLinkAsync(settings.Value.ContainerName, blobName, cancellationToken);

                logger.LogInformation("Successfully processed request for blob: {BlobName}", blobName);
                return Results.Ok(new PresentationSlideMasterInfoViewModel(blobName, link, slideMasterStructure));
            }
            catch (Exception ex)
            {
                // Log the error
                logger.LogError(ex, "An error occurred while processing blob: {BlobName}", blobName);

                // Return a problem result for the client
                return Results.Problem(
                    //detail: $"An error occurred while processing blob: {blobName}. Please check the logs for more details.",
                    detail: ex.Message,
                    statusCode: 500
                );
            }
        }
    }
}
