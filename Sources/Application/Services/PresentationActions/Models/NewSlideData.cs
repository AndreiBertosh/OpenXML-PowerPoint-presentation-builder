namespace Application.Services.PresentationActions.Models
{
    public record NewSlideData(
        string ThemeName,
        string LayoutName,
        string TitleText,
        string SubTitleText,
        string[] BodyText,
        string Author,
        string CommentMessage,
        long SlideWidth = 0,
        long SlideHeight = 0
        );
}
