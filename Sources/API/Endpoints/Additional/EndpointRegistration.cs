using System.ComponentModel;
using System.Net.Mime;
using API.Endpoints.Additional.Version1;
using API.Endpoints.Common;

using Infrastructure.Common;

namespace API.Endpoints.Additional
{
    public static class AdditionalUrls
    {
        public const string RemovePresentation = $"{ApiVersions.ApiPrefix}/presentation/additional/remove";
        public const string BlobList = $"{ApiVersions.ApiPrefix}/presentation/additional/blobList/{{storageType}}";
        public const string DownloadPresentation = $"{ApiVersions.ApiPrefix}/presentation/additional/download/blobs/{{container}}/{{blobName}}";
        public const string GetLinkToDownloadNewPresentation = $"{ApiVersions.ApiPrefix}/additional/GetLinkToDownloadNewPresentation/{{blobName}}";
        public const string GetAppSettings = $"{ApiVersions.ApiPrefix}/presentation/additional/settings";
    }

    public static partial class EndpointRegistration
    {
        private const string _endpointAdditionalGroupName = "Additional";

        public static void RegistrationAdditionalEndpoints(this WebApplication app)
        {
            app
                .MapGet(AdditionalUrls.BlobList, GetListBlobsInStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<IEnumerable<string>>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapGet(AdditionalUrls.DownloadPresentation, DownloadBlobFromStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<string>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapGet(AdditionalUrls.GetLinkToDownloadNewPresentation, GetpresentationDownloadLinkEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<string>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            //app
            //    .MapGet(AdditionalUrls.GetAppSettings, GetAppSettingsEndpoint.ExecuteAsync)
            //    .WithTags(_endpointAdditionalGroupName)
            //    .Produces<string>(contentType: MediaTypeNames.Application.Json)
            //    .Produces(StatusCodes.Status401Unauthorized)
            //    .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
            //    .Produces(StatusCodes.Status500InternalServerError)
            //    .WithApiVersionSet(ApiVersions.VersionSet)
            //    .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPost(AdditionalUrls.RemovePresentation, RemovePresentationFromBlobStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<List<string>>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);
        }
    }
}
