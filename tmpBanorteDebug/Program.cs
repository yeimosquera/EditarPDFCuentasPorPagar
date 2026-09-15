using System;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;

var file = @"C:\Proyectos\EditarPDF\EDO DE CUENTA MAYO 25 ELOR 533.pdf";
using var reader = new PdfReader(file);
using var doc = new PdfDocument(reader);
var page = doc.GetPage(1);
var text = PdfTextExtractor.GetTextFromPage(page);
Console.WriteLine(text.Contains("LIQ.INT.BRUTOS", StringComparison.OrdinalIgnoreCase));
var idx = text.IndexOf("LIQ.INT.BRUTOS", StringComparison.OrdinalIgnoreCase);
Console.WriteLine(idx);
if (idx >= 0)
{
    Console.WriteLine(text.Substring(Math.Max(0, idx - 200), Math.Min(800, text.Length - Math.Max(0, idx - 200))));
}
