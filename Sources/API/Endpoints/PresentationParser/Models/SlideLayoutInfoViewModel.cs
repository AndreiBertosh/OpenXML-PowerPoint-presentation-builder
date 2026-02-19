using Application.Services.PresentationParser.Models;

namespace API.Endpoints.PresentationParser.Models
{
    public sealed record SlideLayoutInfoViewModel(
        string LayoutName,                // display name of the slide layout
        uint? LayoutId,                   // unique layout identifier (may be null)
        IReadOnlyList<ShapeInfoViewModel> Shapes    // collection of shapes defined in this layout
        );
}
