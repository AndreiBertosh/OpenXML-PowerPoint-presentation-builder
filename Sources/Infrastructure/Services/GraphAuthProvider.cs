using Azure.Identity;

using Microsoft.Graph;

namespace Infrastructure.Services
{
    public class GraphAuthProvider
    {
        private readonly string _clientId;
        private readonly string _tenantId;
        private readonly string _clientSecret;

        public GraphAuthProvider(string clientId, string tenantId, string clientSecret)
        {
            _clientId = clientId;
            _tenantId = tenantId;
            _clientSecret = clientSecret;
        }

        /// <summary>
        /// Creates GraphServiceClient with ClientSecretCredential.
        /// </summary>
        public GraphServiceClient GetGraphClient()
        {
            var credential = new ClientSecretCredential(_tenantId, _clientId, _clientSecret);

            // Scope must always be https://graph.microsoft.com/.default for client credentials flow
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            return new GraphServiceClient(credential, scopes);
        }
    }
}
