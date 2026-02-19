namespace Application.Services.PresentationParser.Models
{
    /// <summary>
    /// Represents a slide master, including its display name, internal identifier,
    /// and the collection of all slide layouts associated with this master.
    /// </summary>
    public sealed record SlideMasterInfoDomainModel(
        string MasterName,                                  // display name of the slide master
        uint? MasterId,                                     // unique identifier of the master (may be null)
        IReadOnlyList<SlideLayoutInfoDomainModel> Layouts   // collection of layouts belonging to this master
    );
}

