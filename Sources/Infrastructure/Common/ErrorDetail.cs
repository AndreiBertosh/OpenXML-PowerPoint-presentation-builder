namespace Infrastructure.Common
{
    public record ErrorDetail
    {
        public string? Title { get; set; }

        public string Detail { get; set; } = string.Empty;

        public string? ErrorCode { get; set; }

        public string? ErrorPath { get; set; }
    }
}
