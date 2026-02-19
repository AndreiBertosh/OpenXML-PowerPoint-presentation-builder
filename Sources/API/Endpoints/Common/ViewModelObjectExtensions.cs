using API.Endpoints.SharePoint.Models;

using Application.Services.SharePointServices.Models;

namespace API.Endpoints.Common
{
    public static class ViewModelObjectExtensions
    {
        public static FileInfoViewModel ToViewModel(this FileInfoDomainModel source) =>
            new(
                Name: source.Name,
                Size: source.Size,
                LastModifiedDateTime: source.LastModifiedDateTime
                );
    }
}
