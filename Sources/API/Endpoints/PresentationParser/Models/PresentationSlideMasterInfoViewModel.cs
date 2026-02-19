namespace API.Endpoints.PresentationParser.Models
{
    public record PresentationSlideMasterInfoViewModel(
        string PresentationFileName,
        string PresentationLink,
        IEnumerable<SlideMasterInfoViewModel> SlideMasters
        );
}
