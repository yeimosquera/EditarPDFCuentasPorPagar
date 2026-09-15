namespace EditarPDFCuentasPorPagar.Models
{
    public class ProcesarEstadoCuentaDto
    {
        public string FilePath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public List<ReglaFoleoDto> Reglas { get; set; } = new();
    }

    public class ReglaFoleoDto
    {
        public string PalabraClave { get; set; } = string.Empty;
        public string GuiID { get; set; } = string.Empty;
        public string RContable { get; set; } = string.Empty;
        public string Fecha { get; set; } = string.Empty;
        public string TextoAInsertar { get; set; } = string.Empty;
    }
}
