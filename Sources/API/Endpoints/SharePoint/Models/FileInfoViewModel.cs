namespace API.Endpoints.SharePoint.Models
{
    public record FileInfoViewModel(
        string Name,
        long? Size,
        DateTimeOffset? LastModifiedDateTime
    );
}
