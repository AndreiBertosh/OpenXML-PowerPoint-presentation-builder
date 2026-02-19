namespace API.Endpoints.PresentationParser.Models
{
    /// <summary>
    /// Represents a single paragraph (<a:p>) within a shape's TextBody, with its ordered list of runs/fields.
    /// </summary>
    public record ParagraphInfoViewModel(
        int ParagraphIndex,           // 0-based ordinal within TextBody
        List<RunInfoViewModel> Runs       // ordered runs/fields in this paragraph
        );
}
