namespace Infrastructure.Common
{
    /// <summary>
    /// Business logic exception for SharePoint service errors.
    /// </summary>
    public class SharePointServiceException : Exception
    {
        public SharePointServiceException(string message) : base(message) { }
        public SharePointServiceException(string message, Exception inner) : base(message, inner) { }
    }
}
