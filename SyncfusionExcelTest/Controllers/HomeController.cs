using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using SyncfusionExcelTest.Infra;

namespace SyncfusionExcelTest.Controllers
{
    [Route("api/v1")]
    [ApiController]
    public class HomeController : ControllerBase
    {
        private readonly IFileService _fileService;

        public HomeController(IFileService fileService)
        {
            _fileService = fileService;
        }

        [HttpGet("files/download")]
        public async Task<IActionResult> CreateExcel()
        {
            await _fileService.GenerateExcel();
            return Ok();
        }

    }
}
