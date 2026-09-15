using EditarPDFCuentasPorPagar.Models;
using EditarPDFCuentasPorPagar.Services;
using Microsoft.AspNetCore.Mvc;

namespace EditarPDFCuentasPorPagar.Controllers
{
    [ApiController]
    [Route("api/facturas")]
    public class FacturasController : ControllerBase
    {
        private readonly IPdfFolderProcessorService _pdfFolderProcessorService;

        public FacturasController(IPdfFolderProcessorService pdfFolderProcessorService)
        {
            _pdfFolderProcessorService = pdfFolderProcessorService;
        }

        [HttpPost("process-folder")]
        public async Task<IActionResult> ProcessFolder([FromBody] ProcessPdfFolderRequest request)
        {
            try
            {
                var result = await _pdfFolderProcessorService.ProcessAsync(request);
                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { error = ex.Message, stackTrace = ex.StackTrace });
            }
        }
    }
}
