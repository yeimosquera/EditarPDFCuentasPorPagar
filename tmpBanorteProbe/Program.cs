using System;
using System.Collections.Generic;
using EditarPDF.Models;
using EditarPDF.Services;

var file = @"C:\Proyectos\EditarPDF\EDO DE CUENTA MAYO 25 ELOR 533.pdf";
var output = @"C:\Proyectos\EditarPDF\EDO_DE_CUENTA_MAYO_25_ELOR_533_PROCESADO_API_HTTPS.pdf";
var rules = new List<ReglaInsercionBanorte>
{
new ReglaInsercionBanorte { PalabraClave = "0.11", Fecha = "02-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-02", TextoAInsertar = "T03" },
new ReglaInsercionBanorte { PalabraClave = "0.16", Fecha = "05-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-05", TextoAInsertar = "T08" },
new ReglaInsercionBanorte { PalabraClave = "0.05", Fecha = "06-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-06", TextoAInsertar = "T10" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "07-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-07", TextoAInsertar = "T11" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "08-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-08", TextoAInsertar = "T12" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "09-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-09", TextoAInsertar = "T15" },
new ReglaInsercionBanorte { PalabraClave = "0.18", Fecha = "12-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-12", TextoAInsertar = "T16" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "13-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-13", TextoAInsertar = "T25" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "15-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-15", TextoAInsertar = "T27" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "16-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-16", TextoAInsertar = "T28" },
new ReglaInsercionBanorte { PalabraClave = "0.17", Fecha = "19-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-19", TextoAInsertar = "T40" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "20-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-20", TextoAInsertar = "T41" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "21-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-21", TextoAInsertar = "T42" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "22-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-22", TextoAInsertar = "T43" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "23-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-23", TextoAInsertar = "T52" },
new ReglaInsercionBanorte { PalabraClave = "0.19", Fecha = "26-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-26", TextoAInsertar = "T56" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "27-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-27", TextoAInsertar = "T57" },
new ReglaInsercionBanorte { PalabraClave = "0.06", Fecha = "28-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-28", TextoAInsertar = "T59" },
new ReglaInsercionBanorte { PalabraClave = "221.62", Fecha = "29-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-29", TextoAInsertar = "T61" },
new ReglaInsercionBanorte { PalabraClave = "447.26", Fecha = "30-MAY-25", Descripcion = "LIQ.INT.BRUTOS LIQ 2025-05-31", TextoAInsertar = "T72" }
};
var service = new PdfService();
var result = service.ModificarPdfPorCoordenadasbanorte(new EditPdfCoordinatesBnorteRequest { FilePath = file, OutputPath = output, Reglas = rules });
Console.WriteLine($"RESULT={result}");
Console.WriteLine(System.IO.File.Exists(output) ? "FILE_EXISTS" : "FILE_MISSING");
