using Microsoft.AspNetCore.Mvc;

namespace API.QueryModels.AzureBlobStorage
{
    public record CreatePresentationBySlideIndexesQueryModel(
        [property: FromQuery] string SourceBlobName,
        [property: FromQuery] string DestinationBlobName,
        [property: FromQuery] int[] SlideIndexes,
        [property: FromQuery] string CommentMessage);
}
