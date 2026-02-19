using System.Net.Mime;

using API.Endpoints.Common;
using API.Endpoints.PresentationParser.Models;
using API.Endpoints.PresentationParser.Version1;
using Infrastructure.Common;

namespace API.Endpoints.PresentationParser
{
    public static class ParserUrls
    {
        public const string ParserFromAzureBlobStorage = $"{ApiVersions.ApiPrefix}/presentation/parser/azureBlobStorage/{{storageType}}/{{presentationName}}";
        public const string ParserLayoutsInfoFromAzureBlobStorage = $"{ApiVersions.ApiPrefix}/presentation/parser/layoutsInfo/{{blobName}}";
    }

    public static partial class EndpointRegistration
    {
        private const string _endpointParserGroupName = "Parser";

        public static void RegistrationParserEndpoints(this WebApplication app)
        {
            app
                .MapGet(ParserUrls.ParserFromAzureBlobStorage, GetPresentationDataFromAzureBlobStorageFileEndpoint.ExecuteAsync)
                .WithTags(_endpointParserGroupName)
                .Produces<PresentationDataViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapGet(ParserUrls.ParserLayoutsInfoFromAzureBlobStorage, GetLayoutsInfoFromPresentationEndpoint.ExecuteAsync)
                .WithTags(_endpointParserGroupName)
                .Produces<PresentationSlideMasterInfoViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

        }
    }
}
