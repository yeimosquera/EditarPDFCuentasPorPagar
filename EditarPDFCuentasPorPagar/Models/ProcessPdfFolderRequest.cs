namespace EditarPDFCuentasPorPagar.Models
{
    public class ProcessPdfFolderRequest
    {
        public string FilePath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public List<ReglaBusquedaDto> Reglas { get; set; } = new();
    }

    public class ReglaBusquedaDto
    {
        public string PalabraClave { get; set; } = string.Empty;
        public string TextoAInsertar { get; set; } = string.Empty;
    }

    public class ProcessPdfResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; } = string.Empty;
        public string LogFilePath { get; set; } = string.Empty;
        public List<string> ArchivosProcesados { get; set; } = new();
        public List<string> PalabrasNoEncontradas { get; set; } = new();
    }
}
