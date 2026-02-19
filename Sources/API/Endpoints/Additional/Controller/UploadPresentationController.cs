using Application.Services.AzureServices.Interfaces;

using Infrastructure.Common;

using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace API.Endpoints.Additional.Controller
{
    [ApiController]
    [Route("v{version:apiVersion}/presentation/additional")]
    public class UploadPresentationController : ControllerBase
    {
        [HttpPost("upload")]
        [RequestSizeLimit(2L * 1024 * 1024 * 1024)] // 2 GB
        [IgnoreAntiforgeryToken] 
        public async Task<IActionResult> Upload(
            IFormFile file,
            [FromServices] IAzureServices azureServices,
            IOptions<AzureBlobStorageSettings> settings,
            ILogger<UploadPresentationController> logger,
            CancellationToken cancellationToken = default)
        {
            {
                if (file == null || file.Length == 0)
                    return BadRequest("No file uploaded.");

                logger.LogInformation("Uploading presentation: {fileName}", file.FileName);

                using var stream = file.OpenReadStream();
                using var memoryStream = new MemoryStream();
                await stream.CopyToAsync(memoryStream); 
                memoryStream.Position = 0;

                await azureServices.UploadFileAsync(memoryStream, settings.Value.ResultContainerName, file.FileName, cancellationToken);
                return Ok("Presentation was uploaded.");
            }
        }
    }
}
