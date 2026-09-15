using EditarPDFCuentasPorPagar.Models;
using EditarPDFCuentasPorPagar.Services;
using Microsoft.AspNetCore.Mvc;

namespace EditarPDFCuentasPorPagar.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class EstadoCuentaController : ControllerBase
    {
        private readonly IEstadoCuentaService _estadoCuentaService;

        public EstadoCuentaController(IEstadoCuentaService estadoCuentaService)
        {
            _estadoCuentaService = estadoCuentaService;
        }

        [HttpPost("procesarCuentas")]
        public IActionResult Procesar([FromBody] ProcesarEstadoCuentaDto request)
        {
            try
            {
                _estadoCuentaService.ProcesarYFolearPdf(request);
                return Ok(new { Mensaje = "PDF procesado y foleado correctamente.", Ruta = request.OutputPath });
            }
            catch (Exception ex)
            {
                return BadRequest(new { Error = ex.Message });
            }
        }
    }
}
