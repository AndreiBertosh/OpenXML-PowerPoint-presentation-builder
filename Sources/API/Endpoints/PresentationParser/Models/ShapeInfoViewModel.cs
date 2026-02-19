using DocumentFormat.OpenXml.Presentation;

namespace API.Endpoints.PresentationParser.Models
{
    /// <summary>
    /// Represents a single text shape (placeholder/container) on a slide layout:
    /// identity, binding to master, placeholder semantics, and approximate capacity,
    /// with its full text content organized by paragraphs and runs.
    /// </summary>
    public sealed record ShapeInfoViewModel(
        // Identity / binding
        string ShapeName,           // shape display name (cNvPr/@name)
        uint? ShapeId,              // shape id (cNvPr/@id) on layout
        uint? PlaceholderIndex,     // p:ph/@idx used to bind to master
        uint? MasterShapeId,        // matched master shape id (cNvPr/@id)

        // Placeholder semantics
        string PlaceholderName,     // PlaceholderValues.ToString() or "None"
        PlaceholderValues? PlaceholderType, // resolved placeholder type (Title, Body, etc.)
        string PlaceholderRole,     // friendly role label derived from type

        // Shape-level capacity (computed once per shape)
        int? MaxLinesApprox,        // approximate number of lines that fit the text box
        int? MaxCharsPerLineApprox, // approximate number of characters per line that fit

        // Full text content
        List<ParagraphInfoViewModel> Paragraphs // ordered paragraphs with their runs
    );
}
