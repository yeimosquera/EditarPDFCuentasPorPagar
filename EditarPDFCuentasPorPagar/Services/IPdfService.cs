using EditarPDFCuentasPorPagar.Models;
using System.Collections.Generic;

namespace EditarPDFCuentasPorPagar.Services
{
    public interface IPdfService
    {
        bool ProcessPdf(EditPdfRequest request);
        bool ProcessPdfWithCoordinates(EditPdfCoordinatesRequest request);
        bool ProcessPdfWithCoordinatessabadell(EditPdfCoordinatesRequest request);
        bool ModificarPdfPorCoordenadassantander(EditPdfCoordinatesRequest request);
        bool ModificarPdfPorCoordenadasbanorte(EditPdfCoordinatesBnorteRequest request);
        List<int> ModificarPdfPorCoordenadasmultiva(EditPdfCoordinatesRequest request);
    }
}

