namespace API.Endpoints.PresentationActions.Models
{
    public record ResponseViewModel(
        string PresentationName,
        string? presentationLink,
        IEnumerable<string> Errors
        );
}
