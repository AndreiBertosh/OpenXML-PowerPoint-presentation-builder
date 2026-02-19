namespace API.Endpoints.PresentationParser.Models
{
    public sealed record SlideMasterInfoViewModel(
        string MasterName,                                  // display name of the slide master
        uint? MasterId,                                     // unique identifier of the master (may be null)
        IReadOnlyList<SlideLayoutInfoViewModel> Layouts     // collection of layouts belonging to this master
        );
}
