namespace Application.Services.PresentationData.Interfaces
{
    public interface IPresentationDataServices
    {
        string GetPresentationPath();

        string GetPresentationPath(string directoryPatch, string fileName);

        string GetNewPresentationFile(string directoryPatch, string fileName, string outputPatch);

        string GetPresentationName(string fileName);

        string GetNewRandomPresentationName();

        string GetBlobNameWithoutType(string fileName);
    }
}
