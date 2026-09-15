using System;
using System.IO;
using System.Collections.Generic;
using EditarPDF.Models;
using EditarPDF.Services;

class Program
{
    static void Main()
    {
        var request = new EditPdfCoordinatesRequest
        {
            FilePath = @"C:\Rutas\BANORTE\EDO DE CUENTA ELOR BANORTE 25.pdf",
            OutputPath = @"C:\Rutas\Modificado\EDO DE CUENTA ELOR BANORTE 25.pdf",
            Reglas = new List<ReglaInsercion>
            {
                new ReglaInsercion
                {
                    PalabraClave = "122,868.50",
                    Fecha = "01-OCT-25",
                    TextoAInsertar = "T1"
                }
            }
        };

        var svc = new PdfService();
        try
        {
            Console.WriteLine("Llamando a ModificarPdfPorCoordenadasbanorte...");
            bool ok = svc.ModificarPdfPorCoordenadasbanorte(request);
            Console.WriteLine($"Resultado método: {ok}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Excepción al ejecutar el servicio:");
            Console.WriteLine(ex.ToString());
        }

        var logPath = Path.ChangeExtension(request.OutputPath, ".banorte.debug.log");
        if (File.Exists(logPath))
        {
            Console.WriteLine("---- Inicio log ----");
            var text = File.ReadAllText(logPath);
            Console.WriteLine(text.Length > 10000 ? text.Substring(0, 10000) + "\n... (truncado)" : text);
            Console.WriteLine("---- Fin log ----");
        }
        else
        {
            Console.WriteLine($"No existe el log esperado: {logPath}");
        }

        Test-Path "C:\Rutas\Modificado\EDO DE CUENTA ELOR BANORTE 25.banorte.debug.log";
        Get-ChildItem "C:\Rutas\Modificado" | Sort-Object LastWriteTime -Descending | Select-Object -First 10;
    }
}