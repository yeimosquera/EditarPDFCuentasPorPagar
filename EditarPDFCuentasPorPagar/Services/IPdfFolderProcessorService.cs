using EditarPDFCuentasPorPagar.Models;

namespace EditarPDFCuentasPorPagar.Services
{
    public interface IPdfFolderProcessorService
    {
        Task<ProcessPdfResponse> ProcessAsync(ProcessPdfFolderRequest request);
    }
}
