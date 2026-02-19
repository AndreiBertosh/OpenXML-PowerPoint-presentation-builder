using Application.Services.PresentationData.Interfaces;

namespace Application.Services.PresentationData
{
    public class PresentationDataServices : IPresentationDataServices
    {
        public string GetPresentationPath()
        {
            return @"C:\Projects\ADNO-CPLT\Files\Corporate Presentation_December 10 2024_External Use - Notes.pptx";
        }

        public string GetNewPresentationFile(string directoryPatch, string fileName, string outputPatch)
        {
            string presentationFile = GetPresentationPath(directoryPatch, fileName);
            string newPresentationFile = @$"{outputPatch}\{GetNewRandomPresentationName()}";

            File.Copy(presentationFile, newPresentationFile, overwrite: true);

            return newPresentationFile;
        }

        public string GetPresentationPath(string directoryPatch, string fileName)
        {
            if (!fileName.EndsWith(".pptx"))
            {
                fileName = $"{fileName}.pptx";
            }

            return @$"{directoryPatch}\{fileName}";
        }

        public string GetPresentationName(string fileName)
        {
            var fileTypes = new List<string>() { "pptx", "potx" };
            var fileType = fileName.Substring(fileName.LastIndexOf(".") + 1);

            if (!fileTypes.Contains(fileType))
            {
                fileName = $"{fileName}.pptx";
            }

            return fileName;
        }

        public string GetNewRandomPresentationName()
        {
            return $"{Guid.NewGuid()}.pptx";
        }

        public string GetBlobNameWithoutType(string fileName)
        {
            return Path.GetFileNameWithoutExtension(fileName) ;
        }
    }
}
