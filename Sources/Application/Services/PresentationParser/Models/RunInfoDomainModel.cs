namespace Application.Services.PresentationParser.Models
{
    /// <summary>
    /// Represents a single text node inside a paragraph: either a normal run (<a:r>) or a field (<a:fld>).
    /// Holds resolved typography and field metadata.
    /// </summary>
    public sealed record RunInfoDomainModel(
        int RunIndex,             // 0-based ordinal within paragraph
        string Text,              // textual content for this node
        string? FontName,         // resolved font name (null if not specified anywhere)
        double? FontSizePt,       // resolved font size in points (null if not specified anywhere)
        bool IsField = false,     // true if node is <a:fld>
        string? FieldType = null, // e.g., "slidenum", "datetime", "footer"
        string? FieldId = null    // fld/@id (GUID-like string)
    );
}
