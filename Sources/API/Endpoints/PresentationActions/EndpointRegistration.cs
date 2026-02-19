using System.Net.Mime;

using API.Endpoints.Common;
using API.Endpoints.PresentationActions.Models;
using API.Endpoints.PresentationActions.Version1;

using Infrastructure.Common;

namespace API.Endpoints.PresentationActions
{
    public static class ActionsUrls
    {
        public const string CreateNewPresentationAzureBlobStorageBySlideIndexes = $"{ApiVersions.ApiPrefix}/action/AzureBlobStorage/createNewPresentationBySlideIndexes";
        public const string AddSlideToPresentationAzureBlobStorage = $"{ApiVersions.ApiPrefix}/action/AzureBlobStorage/addSlideToPresentation";
        public const string AddSlidesToPresentationAzureBlobStorage = $"{ApiVersions.ApiPrefix}/action/AzureBlobStorage/addSlidesToPresentation";
        public const string CopySlideToPresentationAzureBlobStorage = $"{ApiVersions.ApiPrefix}/action/AzureBlobStorage/copySlideToPresentation";

        public const string CopySlideToPresentation = $"{ApiVersions.ApiPrefix}/action/copySlideToPresentation";
        public const string CreateNewPresentationBySlideIndexes = $"{ApiVersions.ApiPrefix}/action/createNewPresentationBySlideIndexes";
        public const string AddSlideToPresentation = $"{ApiVersions.ApiPrefix}/action/addSlideToPresentation";
        public const string AddSlidesToPresentation = $"{ApiVersions.ApiPrefix}/action/addSlidesToPresentation";
    }

    public static partial class EndpointRegistration
    {
        private const string _endpointActionsGroupName = "Actions";
        private const string _endpointSharePointGroupName = "Actions SharePoint";

        public static void RegistrationActionsEndpoints(this WebApplication app)
        {
            #region version V1

            app
                .MapPatch(ActionsUrls.CopySlideToPresentationAzureBlobStorage, CopySlideToPresentationAzureBlobStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointActionsGroupName)
                .Produces<ResponseViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPatch(ActionsUrls.AddSlideToPresentationAzureBlobStorage, InsertNewSlideToPresentationAzureBlobStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointActionsGroupName)
                .Produces<ResponseViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPatch(ActionsUrls.AddSlidesToPresentationAzureBlobStorage, InsertNewSlidesToPresentationAzureBlobStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointActionsGroupName)
                .Produces<ResponseViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPost(ActionsUrls.CreateNewPresentationAzureBlobStorageBySlideIndexes, CreateNewPresentationBySlideIndexesAzureBlobStorageEndpoint.ExecuteAsync)
                .WithTags(_endpointActionsGroupName)
                .Produces<ResponseViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
            .MapToApiVersion(ApiVersions.Version1_0);

            #endregion

            #region version V1 SharePoint

            app
                .MapPost(ActionsUrls.CreateNewPresentationBySlideIndexes, CreateNewPresentationBySlideIndexesEndpoint.ExecuteAsync)
                .WithTags(_endpointSharePointGroupName)
                .Produces<ResponseViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPatch(ActionsUrls.CopySlideToPresentation, CopySlideToPresentationEndpoint.ExecuteAsync)
                .WithTags(_endpointSharePointGroupName)
                .Produces<ResponseViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPatch(ActionsUrls.AddSlidesToPresentation, InsertNewSlidesToPresentationEndpoint.ExecuteAsync)
                .WithTags(_endpointSharePointGroupName)
                .Produces<ResponseViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            app
                .MapPatch(ActionsUrls.AddSlideToPresentation, InsertNewSlideToPresentationEndpoint.ExecuteAsync)
                .WithTags(_endpointSharePointGroupName)
                .Produces<ResponseViewModel>(contentType: MediaTypeNames.Application.Json)
                .Produces(StatusCodes.Status401Unauthorized)
                .Produces<ProblemDetailsWithErrors>(StatusCodes.Status400BadRequest)
                .Produces(StatusCodes.Status500InternalServerError)
                .WithApiVersionSet(ApiVersions.VersionSet)
                .MapToApiVersion(ApiVersions.Version1_0);

            #endregion

            #region version V2

            #endregion
        }
    }
}
