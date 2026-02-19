namespace API.Endpoints.PresentationParser.Models
{
    public record PresentationDataViewModel(
        string PresentationFileName,
        string PresentationLink,
        List<SlideDataViewModel> Slides
        );
}
