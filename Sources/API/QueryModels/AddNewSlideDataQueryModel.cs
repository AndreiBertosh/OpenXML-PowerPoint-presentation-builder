namespace API.QueryModels
{
    public record AddNewSlideDataQueryModel(
        string PresentationName,
        string TemplatePresentationName,
        string? ThemeName,
        string? LayoutName,
        string? Title,
        string? SubTitle,
        string[]? BodyText,
        string? SlideComment,
        bool IsPictureLeft = false,
        int? Position = -1
    );
}
