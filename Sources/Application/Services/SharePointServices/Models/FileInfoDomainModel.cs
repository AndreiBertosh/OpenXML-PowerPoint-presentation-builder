namespace Application.Services.SharePointServices.Models
{
    public record FileInfoDomainModel(
        string Name,
        long? Size,
        DateTimeOffset? LastModifiedDateTime
    );
}
