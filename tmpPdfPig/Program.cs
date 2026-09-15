using System;
using UglyToad.PdfPig;

var file = @"C:\Proyectos\EditarPDF\EDO DE CUENTA MAYO 25 ELOR 533.pdf";
using var document = PdfDocument.Open(file);
var found = false;
foreach (var page in document.GetPages())
{
    var text = page.Text;
    if (text.Contains("LIQ.INT.BRUTOS", StringComparison.OrdinalIgnoreCase))
    {
        found = true;
        var idx = text.IndexOf("LIQ.INT.BRUTOS", StringComparison.OrdinalIgnoreCase);
        Console.WriteLine(text.Substring(Math.Max(0, idx - 200), Math.Min(700, text.Length - Math.Max(0, idx - 200))));
        break;
    }
}
Console.WriteLine($"FOUND={found}");
