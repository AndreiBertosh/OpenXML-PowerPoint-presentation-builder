using System.Net.Mime;

using API.Endpoints.Common;
using API.Endpoints.SharePoint.Version1;
using Infrastructure.Common;

namespace API.Endpoints.SharePoint
{
    public static class SharePointUrls
    {
        public const string SharePointList = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/folderList";
        public const string SharePointFileList = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/foldersData";
        public const string SharePointFolderFileList = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/folderData/{{folderName}}";
        public const string SharePointDownloadList = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/downloadLink/{{fileName}}";
        public const string SharePointDoesFileExist = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/doesFileExist/{{folderName}}/{{fileName}}";

        public const string SharePointCopyFileToStorage = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/CopyFileToStorage";
        public const string SharePointCopyFileToSharePoint = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/CopyFileToSharePoint/{{fileName}}";

        public const string SharePointDeleteFile = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/delete/{{folderName}}/{{fileName}}";
        public const string SharePointSubscription = $"{ApiVersions.ApiPrefix}/presentation/SharePoint/subscription";
    }

    public static partial class EndpointRegistration
    {
        private const string _endpointAdditionalGroupName = "SharePoint";

        public static void RegistrationSharePointEndpoints(this WebApplication app)
        {
            app
                .MapGet(SharePointUrls.SharePointFileList, GetFileListEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<Dictionary<string, List<string>>>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapGet(SharePointUrls.SharePointFolderFileList, GetFolderFileListEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<Dictionary<string, List<string>>>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapGet(SharePointUrls.SharePointDownloadList, GetPresentationDownloadLinkEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<string>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapGet(SharePointUrls.SharePointDoesFileExist, DoesFileExistEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<bool>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPut(SharePointUrls.SharePointCopyFileToStorage, DownloadFileToBlobStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<string>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPut(SharePointUrls.SharePointCopyFileToSharePoint, UploadFileFromBlobStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<string>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapDelete(SharePointUrls.SharePointDeleteFile, DeleteFileEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<string>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPost(SharePointUrls.SharePointSubscription, SubscribeToTemplateFolderEndpoint.ExecuteAsync)
                .WithTags(_endpointAdditionalGroupName)
                .Produces<string>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);
        }
    }
}
