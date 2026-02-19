namespace Infrastructure.Common
{
    public class AppSettings
    {
        public required string AppName { get; set; }

        public int HttpPort { get; set; } = 8080;

        public int HttpsPort { get; set; } = 8081;
    }
}
