using Microsoft.AspNetCore.Mvc;

namespace API.QueryModels.AzureBlobStorage
{
    public record AddNewSlideDataQueryModel(
        [property: FromQuery] string BlobName,
        [property: FromQuery] string TemplateBlobName,
        [property: FromQuery] string? ThemeName,
        [property: FromQuery] string? LayoutName,
        [property: FromQuery] string? Title,
        [property: FromQuery] string? SubTitle,
        [property: FromQuery] string[]? BodyText,
        [property: FromQuery] string? SlideComment,
        [property: FromQuery] int? Position = -1
        );
}
