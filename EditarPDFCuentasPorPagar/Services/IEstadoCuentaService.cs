using EditarPDFCuentasPorPagar.Models;

namespace EditarPDFCuentasPorPagar.Services
{
    public interface IEstadoCuentaService
    {
        void ProcesarYFolearPdf(ProcesarEstadoCuentaDto request);
    }
}
