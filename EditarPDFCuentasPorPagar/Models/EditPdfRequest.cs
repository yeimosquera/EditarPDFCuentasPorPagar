namespace EditarPDFCuentasPorPagar.Models
{
    public class EditPdfRequest
    {
        public string FilePath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public string TextToAdd { get; set; } = string.Empty;
        // Identificador del 1 al 7 para saber qué reglas aplicar
        public int DocumentType { get; set; }
    }

    public class EditPdfCoordinatesRequest
    {
        public string FilePath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public List<ReglaInsercion> Reglas { get; set; } = new();
    }

    public class EditPdfCoordinatesBnorteRequest
    {
        public string FilePath { get; set; } = string.Empty;
        public string OutputPath { get; set; } = string.Empty;
        public List<ReglaInsercionBanorte> Reglas { get; set; } = new();
    }

    public class TextItem
    {
        public string Word { get; set; } = string.Empty;
        public float Width { get; set; }
        public float Height { get; set; }
        public float Left { get; set; }
        public float Top { get; set; }
        public string FontName { get; set; } = string.Empty;
        public int Page { get; set; }
        public int Index { get; set; }
    }

    public class ReglaInsercion
    {
        public string PalabraClave { get; set; } = string.Empty;
        public string Fecha { get; set; } = string.Empty;       
        public string TextoAInsertar { get; set; } = string.Empty;
    }

    public class ReglaInsercionBanorte
    {
        public string PalabraClave { get; set; } = string.Empty;
        public string Fecha { get; set; } = string.Empty;
        public string Descripcion { get; set; } = string.Empty;
        public string TextoAInsertar { get; set; } = string.Empty;
    }

    public class WordLocation
    {
        public string Text { get; set; } = string.Empty;
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
    }

}
