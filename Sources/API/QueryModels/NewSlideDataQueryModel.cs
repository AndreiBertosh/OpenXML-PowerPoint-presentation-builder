namespace API.QueryModels
{
    public record NewSlideDataQueryModel(
        string? ThemeName,
        string? LayoutName,
        string? Title,
        string? SubTitle,
        string[]? BodyText,
        string? SlideComment);
}
