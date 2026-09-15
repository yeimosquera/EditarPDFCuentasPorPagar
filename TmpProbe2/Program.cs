using System;
using System.Linq;
using System.Text.RegularExpressions;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas.Parser;
using EditarPDF.Services;

static string NormalizarTextoBusqueda(string? valor)
{
    if (string.IsNullOrWhiteSpace(valor)) return string.Empty;
    return Regex.Replace(valor.Trim(), @"\s+", " ").Replace(" ", " ").ToUpperInvariant();
}
static string NormalizarFechaParaBusqueda(string? fecha)
{
    if (string.IsNullOrWhiteSpace(fecha)) return string.Empty;
    var fechaTrim = fecha.Trim();
    if (DateTime.TryParse(fechaTrim, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out var fechaDate))
        return fechaDate.ToString("dd-MMM-yyyy", System.Globalization.CultureInfo.InvariantCulture).ToUpperInvariant();
    var fechaLimpia = Regex.Replace(fechaTrim, @"[^0-9A-Z]", string.Empty).ToUpperInvariant();
    if (fechaLimpia.Length >= 8) { var dia = fechaLimpia.Substring(0,2); var mes = fechaLimpia.Substring(2,2); var anno = fechaLimpia.Substring(4); return $"{dia}-{mes}-{anno}"; }
    return fechaTrim.ToUpperInvariant();
}
static string NormalizarMontoParaBusqueda(string? valor)
{
    if (string.IsNullOrWhiteSpace(valor)) return string.Empty;
    var numero = valor.Trim().Replace(",", ".");
    var soloNumero = Regex.Replace(numero, @"[^0-9\.\-]", string.Empty);
    return soloNumero;
}
static bool TieneFechaEnTexto(string texto, string fechaBuscada)
{
    var textoNormalizado = NormalizarTextoBusqueda(texto);
    var fechaNormalizada = NormalizarFechaParaBusqueda(fechaBuscada);
    var patronesFecha = new[]
    {
        Regex.Escape(fechaNormalizada),
        Regex.Escape(fechaNormalizada.Replace("-", string.Empty)),
        Regex.Escape(fechaBuscada.Trim().ToUpperInvariant()),
        @"\b\d{1,2}[-/ ]*(?:ENE|FEB|MAR|ABR|MAY|JUN|JUL|AGO|SEP|OCT|NOV|DIC)\b",
        @"\b\d{4}[-/]\d{2}[-/]\d{2}\b",
        @"\b\d{1,2}[-/ ]*\d{1,2}[-/ ]*\d{2,4}\b"
    };
    return patronesFecha.Any(p => Regex.IsMatch(textoNormalizado, p, RegexOptions.IgnoreCase));
}
static bool TieneMontoEnTexto(string texto, string montoBuscado)
{
    var montoNormalizado = NormalizarMontoParaBusqueda(montoBuscado);
    var textoNormalizado = NormalizarTextoBusqueda(texto);
    var textoSoloDigitos = Regex.Replace(textoNormalizado, @"[^0-9\.\-]", string.Empty);
    return textoNormalizado.Contains(montoNormalizado, StringComparison.OrdinalIgnoreCase)
        || textoSoloDigitos.Contains(montoNormalizado, StringComparison.OrdinalIgnoreCase);
}
static bool TieneDescripcionEnTexto(string texto, string descripcionBuscada)
{
    var textoNormalizado = NormalizarTextoBusqueda(texto);
    var descripcionNormalizada = NormalizarTextoBusqueda(descripcionBuscada);
    var descripcionCompacta = Regex.Replace(descripcionNormalizada, @"[^A-Z0-9]", string.Empty);
    var textoCompacto = Regex.Replace(textoNormalizado, @"[^A-Z0-9]", string.Empty);
    return textoNormalizado.Contains(descripcionNormalizada, StringComparison.OrdinalIgnoreCase)
        || textoCompacto.Contains(descripcionCompacta, StringComparison.OrdinalIgnoreCase);
}

using var fs = new System.IO.FileStream(@"C:\Proyectos\EditarPDF\EDO DE CUENTA MAYO 25 ELOR 533.pdf", System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read);
using var reader = new PdfReader(fs);
using var doc = new PdfDocument(reader);
var page = doc.GetPage(2);
var strategy = new FilaExtractionStrategy(page.GetPageSize().GetHeight());
var processor = new PdfCanvasProcessor(strategy);
processor.ProcessPageContent(page);
var filas = strategy.ObtenerFilas();
var row = filas.FirstOrDefault(f => string.Join(" ", f.Select(w => w.Text ?? string.Empty)).Contains("2025-05-02", StringComparison.OrdinalIgnoreCase));
var text = row == null ? "" : string.Join(" ", row.Select(w => w.Text ?? string.Empty));
Console.WriteLine("ROW=" + text);
Console.WriteLine("FECHA=" + TieneFechaEnTexto(text, "02-MAY-25"));
Console.WriteLine("MONTO=" + TieneMontoEnTexto(text, "0.11"));
Console.WriteLine("DESC=" + TieneDescripcionEnTexto(text, "LIQ.INT.BRUTOS LIQ 2025-05-02"));
Console.WriteLine("COMPACT=" + Regex.Replace(NormalizarTextoBusqueda(text), @"[^A-Z0-9]", string.Empty));
Console.WriteLine("COMPACTDESC=" + Regex.Replace(NormalizarTextoBusqueda("LIQ.INT.BRUTOS LIQ 2025-05-02"), @"[^A-Z0-9]", string.Empty));
