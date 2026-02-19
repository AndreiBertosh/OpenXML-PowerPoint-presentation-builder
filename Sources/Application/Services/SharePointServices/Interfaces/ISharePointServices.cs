using Application.Services.SharePointServices.Models;

namespace Application.Services.SharePointServices.Interfaces
{
    public interface ISharePointServices
    {
        Task<Dictionary<string, List<string>>> GetRootFilesAsync(CancellationToken cancellationToken);

        Task<MemoryStream> DownloadFileByNameAsync(string folderName, string fileName, CancellationToken cancellationToken);

        Task<string> UploadFileAsync(string folderPath, string fileName, Stream contentStream, CancellationToken cancellationToken);

        Task<Dictionary<string, List<FileInfoDomainModel>>> GetFilesInFolderAsync(string folderName, CancellationToken cancellationToken);

        Task<string> GetFileDownloadLinkAsync(string fileName, CancellationToken cancellationToken);

        Task<bool> DoesFileExistAsync(string folderName, string fileName, CancellationToken cancellationToken);

        Task DeleteFileAsync(string folderName, string fileName, CancellationToken cancellationToken);

        Task<string> SubscribeWebHookForFolderAsync(string azureFunctionUrl, CancellationToken cancellationToken);
    }
}
