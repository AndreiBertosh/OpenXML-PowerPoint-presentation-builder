namespace Application.Services.PresentationParser.Models
{ 
    /// <summary>
    /// Represents a slide layout, including its display name, unique layout identifier,
    /// and the collection of shapes defined within this layout.
    /// </summary>
    public sealed record SlideLayoutInfoDomainModel(
        string LayoutName,                              // display name of the slide layout
        uint? LayoutId,                                 // unique layout identifier (may be null)
        IReadOnlyList<ShapeInfoDomainModel> Shapes      // collection of shapes defined in this layout
    );
}
