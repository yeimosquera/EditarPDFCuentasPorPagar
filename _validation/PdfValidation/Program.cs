using System;
using System.Collections.Generic;
using System.IO;
using EditarPDF.Models;
using EditarPDF.Services;

static string ResolveProjectRoot()
{
    var roots = new[]
    {
        Environment.CurrentDirectory,
        AppContext.BaseDirectory
    };

    foreach (var root in roots)
    {
        var current = root;
        for (var i = 0; i < 10; i++)
        {
            if (File.Exists(Path.Combine(current, "EDO DE CUENTA MAYO 25 ELOR 423.pdf")))
                return current;

            if (File.Exists(Path.Combine(current, "EditarPDF.sln")))
                return current;

            var parent = Directory.GetParent(current);
            if (parent == null)
                break;

            current = parent.FullName;
        }
    }

    return Directory.GetCurrentDirectory();
}

static void EjecutarPruebaIntegracionSantander()
{
    var projectRoot = ResolveProjectRoot();
    var inputPath = Path.Combine(projectRoot, "EDO DE CUENTA MAYO 25 ELOR 423.pdf");
    var outputPath = Path.Combine(projectRoot, "EDO_DE_CUENTA_MAYO_25_ELOR_423_PROCESADO.pdf");
    var logPath = Path.Combine(projectRoot, Path.GetFileNameWithoutExtension(outputPath) + "_log.xlsx");

    if (!File.Exists(inputPath))
        throw new FileNotFoundException($"No se encontró el PDF de prueba en la ruta: {inputPath}");

    var request = new EditPdfCoordinatesRequest
    {
        FilePath = inputPath,
        OutputPath = outputPath,
        Reglas = new List<ReglaInsercion>
        {
            new() { PalabraClave = "16,477,000.00", Fecha = "02-MAY-2025", TextoAInsertar = "T51" },
            new() { PalabraClave = "12,316.56", Fecha = "05-MAY-2025", TextoAInsertar = "T52" },
            new() { PalabraClave = "12,746.12", Fecha = "06-MAY-2025", TextoAInsertar = "T53" },
            new() { PalabraClave = "290.78", Fecha = "07-MAY-2025", TextoAInsertar = "T54" },
            new() { PalabraClave = "290.78", Fecha = "08-MAY-2025", TextoAInsertar = "T55" },
            new() { PalabraClave = "290.78", Fecha = "09-MAY-2025", TextoAInsertar = "T56" },
            new() { PalabraClave = "873.09", Fecha = "12-MAY-2025", TextoAInsertar = "T57" },
            new() { PalabraClave = "1,268.25", Fecha = "13-MAY-2025", TextoAInsertar = "T58" },
            new() { PalabraClave = "1,220.92", Fecha = "14-MAY-2025", TextoAInsertar = "T59" },
            new() { PalabraClave = "1,752.89", Fecha = "15-MAY-2025", TextoAInsertar = "T60" },
            new() { PalabraClave = "67,909.15", Fecha = "16-MAY-2025", TextoAInsertar = "T61" },
            new() { PalabraClave = "195,639.36", Fecha = "19-MAY-2025", TextoAInsertar = "T62" },
            new() { PalabraClave = "19,057.50", Fecha = "20-MAY-2025", TextoAInsertar = "T63" },
            new() { PalabraClave = "4,197.45", Fecha = "21-MAY-2025", TextoAInsertar = "T64" },
            new() { PalabraClave = "5,002.01", Fecha = "22-MAY-2025", TextoAInsertar = "T65" },
            new() { PalabraClave = "1,976.33", Fecha = "23-MAY-2025", TextoAInsertar = "T66" },
            new() { PalabraClave = "8,102.25", Fecha = "26-MAY-2025", TextoAInsertar = "T67" },
            new() { PalabraClave = "4,108.01", Fecha = "02-MAY-2025", TextoAInsertar = "T68" },
            new() { PalabraClave = "9,261.15", Fecha = "27-MAY-2025", TextoAInsertar = "T69" },
            new() { PalabraClave = "9,646.39", Fecha = "28-MAY-2025", TextoAInsertar = "T70" },
            new() { PalabraClave = "10,209.64", Fecha = "29-MAY-2025", TextoAInsertar = "T71" },
            new() { PalabraClave = "11,196.63", Fecha = "30-MAY-2025", TextoAInsertar = "T72" }
        }
    };

    var service = new PdfService();
    var resultado = service.ModificarPdfPorCoordenadassantander(request);

    Console.WriteLine($"Resultado: {resultado}");
    Console.WriteLine($"Archivo procesado existe: {File.Exists(outputPath)}");
    Console.WriteLine($"Ruta del PDF procesado: {outputPath}");
    Console.WriteLine($"Ruta del log Excel: {logPath}");
    Console.WriteLine($"Cantidad de reglas: {request.Reglas.Count}");
}

EjecutarPruebaIntegracionSantander();
