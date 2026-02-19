using Microsoft.AspNetCore.Mvc;

namespace API.QueryModels.AzureBlobStorage
{
    public record CopySlideQueryModel(
        [property: FromQuery] string SourceBlobName,
        [property: FromQuery] string? SourceStorageType,
        [property: FromQuery] string TemplateBlobName,
        [property: FromQuery] string DestinationBlobName,
        [property: FromQuery] int[] SlideIndexes,
        [property: FromQuery] string CommentMessage);
}
