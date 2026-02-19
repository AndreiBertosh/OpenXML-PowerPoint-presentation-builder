namespace Application.Services.PresentationParser.Models
{
    /// <summary>
    /// Represents a single paragraph (<a:p>) within a shape's TextBody, with its ordered list of runs/fields.
    /// </summary>
    public sealed record ParagraphInfoDomainModel(
        int ParagraphIndex,                 // 0-based ordinal within TextBody
        List<RunInfoDomainModel> Runs       // ordered runs/fields in this paragraph
    );
}
