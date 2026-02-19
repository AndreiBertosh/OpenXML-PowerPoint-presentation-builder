namespace Infrastructure.Common
{
    public sealed record ProblemDetailsWithErrors
    {
        public int? StatusCode { get; set; }

        public string? Type { get; set; }

        public string Reason { get; set; } = "Generic";

        public string? Title { get; set; }

        public string? Detail { get; set; }

        public string? Instance { get; set; }

        public List<ErrorDetail>? Errors { get; set; }
    }
}
