using EditarPDFCuentasPorPagar.Models;
using iText.IO.Font.Constants;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using System.Text.RegularExpressions;
using iText.Kernel.Geom; // Asegúrate de tener este namespace para el Rectangle
using ClosedXML.Excel;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Globalization;

namespace EditarPDFCuentasPorPagar.Services
{
    public class PdfService : IPdfService
    {
        public bool ProcessPdf(EditPdfRequest request)
        {
            if (!File.Exists(request.FilePath))
                throw new FileNotFoundException("El PDF de origen no existe.");

            try
            {
                using (FileStream fsOriginal = new FileStream(request.FilePath, FileMode.Open, FileAccess.Read))
                using (FileStream fsDestino = new FileStream(request.OutputPath, FileMode.Create, FileAccess.Write))
                {
                    PdfReader reader = new PdfReader(fsOriginal);
                    PdfWriter writer = new PdfWriter(fsDestino);

                    using (PdfDocument pdfDoc = new PdfDocument(reader, writer))
                    {
                        PdfPage page = pdfDoc.GetFirstPage();

                        if (request.DocumentType == 1)
                        {
                            AddTextByCoordinates(page, request.TextToAdd, 450, 520);
                        }
                        else
                        {
                            AddTextByCoordinates(page, request.TextToAdd, 200, 500);
                        }

                        pdfDoc.Close();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error: {ex.Message}", ex);
            }
        }

        private void AddTextByCoordinates(PdfPage page, string text, float x, float y)
        {
            var pageSize = page.GetPageSize();
            if (x < 0 || x > pageSize.GetWidth() || y < 0 || y > pageSize.GetHeight())
            {
                x = 50; y = 50;
            }

            PdfCanvas canvas = new PdfCanvas(page);
            PdfFont font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);

            canvas.BeginText()
                  .SetFontAndSize(font, 12)
                  .MoveText(x, y)
                  .ShowText(text)
                  .EndText();

            canvas.Release();
        }

        public bool ProcessPdfWithCoordinates(EditPdfCoordinatesRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request), "La petición no puede ser nula.");

            if (!File.Exists(request.FilePath))
                throw new FileNotFoundException("El PDF de origen no existe.");

            try
            {
                var modificaciones = new List<(int PageIndex, WordLocation Anchor, string Texto)>();
                var paginasModificadas = new List<int>();

                using (var fsOriginal = new FileStream(request.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var reader = new PdfReader(fsOriginal))
                using (var srcDoc = new PdfDocument(reader))
                {
                    PdfFont font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    int totalPaginas = srcDoc.GetNumberOfPages();

                    // Primera pasada: detectar modificaciones sin escribir
                    for (int i = 1; i <= totalPaginas; i++)
                    {
                        var page = srcDoc.GetPage(i);
                        var pageSize = page.GetPageSize();

                        var strategy = new FilaExtractionStrategy(pageSize.GetHeight());
                        var processor = new PdfCanvasProcessor(strategy);
                        processor.ProcessPageContent(page);

                        var filasDeLaPagina = strategy.ObtenerFilas();

                        foreach (var regla in request.Reglas)
                        {
                            foreach (var fila in filasDeLaPagina)
                            {
                                var itemPalabraClave = fila.FirstOrDefault(w => w.Text.Contains(regla.PalabraClave, StringComparison.OrdinalIgnoreCase));
                                var itemFecha = fila.FirstOrDefault(w => w.Text.Contains(regla.Fecha, StringComparison.OrdinalIgnoreCase));

                                if (itemPalabraClave != null && itemFecha != null)
                                {
                                    modificaciones.Add((i, itemPalabraClave, regla.TextoAInsertar));
                                    if (!paginasModificadas.Contains(i))
                                        paginasModificadas.Add(i);

                                    // Se detectó modificación en esta línea para la regla actual; pasar a la siguiente regla
                                    break;
                                }
                            }
                        }
                    }

                    // Escritura: copiar solo página 1 y páginas modificadas
                    using (var fsDestino = new FileStream(request.OutputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    using (var writer = new PdfWriter(fsDestino))
                    using (var dstDoc = new PdfDocument(writer))
                    {
                        void DibujarTextoEnPagina(PdfPage dstPage, WordLocation anchor, string texto)
                        {
                            var pageSize = dstPage.GetPageSize();
                            PdfCanvas canvas = new PdfCanvas(dstPage.NewContentStreamAfter(), dstPage.GetResources(), dstDoc);

                            float offsetHorizontal = 2f;
                            float x_iText = anchor.Left + anchor.Width + offsetHorizontal;

                            float y_iText = pageSize.GetHeight() - anchor.Top - anchor.Height;
                            y_iText += 1f;

                            float fontSizeBase = anchor.Height > 0 ? anchor.Height : 10f;
                            float fontSize = fontSizeBase + 8f;

                            canvas.BeginText()
                                  .SetFontAndSize(font, fontSize)
                                  .MoveText(x_iText, y_iText);

                            canvas.SetTextRenderingMode(PdfCanvasConstants.TextRenderingMode.FILL_STROKE);
                            canvas.SetLineWidth(0.5f);
                            canvas.SetStrokeColor(ColorConstants.BLACK);

                            canvas.ShowText(texto)
                                  .EndText();

                            canvas.Release();
                        }

                        // 1) Copiar página 1 si existe
                        if (srcDoc.GetNumberOfPages() >= 1)
                        {
                            srcDoc.CopyPagesTo(1, 1, dstDoc);
                            var modsPagina1 = modificaciones.Where(m => m.PageIndex == 1).ToList();
                            if (modsPagina1.Any())
                            {
                                var dstPage = dstDoc.GetPage(dstDoc.GetNumberOfPages());
                                foreach (var mod in modsPagina1)
                                    DibujarTextoEnPagina(dstPage, mod.Anchor, mod.Texto);
                            }
                        }

                        // 2) Copiar demás páginas modificadas (sin duplicar la 1)
                        var paginasParaCopiar = paginasModificadas
                            .Distinct()
                            .OrderBy(p => p)
                            .Where(p => p != 1)
                            .ToList();

                        foreach (var p in paginasParaCopiar)
                        {
                            srcDoc.CopyPagesTo(p, p, dstDoc);
                            var dstPage = dstDoc.GetPage(dstDoc.GetNumberOfPages());
                            var mods = modificaciones.Where(m => m.PageIndex == p).ToList();
                            foreach (var mod in mods)
                                DibujarTextoEnPagina(dstPage, mod.Anchor, mod.Texto);
                        }

                        // Asegurar que la lista interna tenga la página 1 al inicio (si la necesitas luego)
                        if (!paginasModificadas.Contains(1))
                            paginasModificadas.Insert(0, 1);
                        else
                        {
                            paginasModificadas = paginasModificadas.Distinct().OrderBy(p => paginasModificadas.IndexOf(p)).ToList();
                            paginasModificadas.Remove(1);
                            paginasModificadas.Insert(0, 1);
                        }
                    }

                    // --- CREAR LOG (BBVA) ---
                    try
                    {
                        var filasLog = (request.Reglas ?? Enumerable.Empty<ReglaInsercion>())
                            .Select(r =>
                            {
                                var pagina = modificaciones
                                    .Where(m => m.Texto == (r.TextoAInsertar ?? string.Empty))
                                    .Select(m => m.PageIndex)
                                    .FirstOrDefault();

                                return (Texto: r.TextoAInsertar ?? string.Empty,
                                        Exito: modificaciones.Any(m => m.Texto == r.TextoAInsertar),
                                        Pagina: pagina > 0 ? pagina.ToString() : string.Empty);
                            })
                            .ToList();

                        // Guardar Excel junto al outputPath con nombre fijo
                        string dir = System.IO.Path.GetDirectoryName(request.OutputPath) ?? "";
                        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                            Directory.CreateDirectory(dir);

                        string fileName = System.IO.Path.GetFileNameWithoutExtension(request.OutputPath) + "_log.xlsx";
                        string logPath = System.IO.Path.Combine(dir, fileName);

                        using (var wb = new ClosedXML.Excel.XLWorkbook())
                        {
                            var ws = wb.Worksheets.Add("Log");
                            ws.Cell(1, 1).Value = "TextoAInsertar";
                            ws.Cell(1, 2).Value = "Exito";
                            ws.Cell(1, 3).Value = "Página";

                            int r = 2;
                            foreach (var f in filasLog)
                            {
                                ws.Cell(r, 1).Value = f.Texto ?? "";
                                ws.Cell(r, 2).Value = f.Exito ? "TRUE" : "FALSE";
                                ws.Cell(r, 3).Value = f.Pagina ?? string.Empty;
                                r++;
                            }

                            ws.Columns().AdjustToContents();
                            if (File.Exists(logPath))
                                File.Delete(logPath);
                            wb.SaveAs(logPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"No se pudo crear log Excel para {request.OutputPath}: {ex.Message}");
                        // No interrumpir el flujo principal por fallo en log
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en motor de renderizado geométrico de PDF (BBVA): {ex.Message}", ex);
            }
        }

        public bool ProcessPdfWithCoordinatessabadell(EditPdfCoordinatesRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request), "La petición no puede ser nula.");

            if (!File.Exists(request.FilePath))
                throw new FileNotFoundException("El PDF de origen no existe.");

            try
            {
                // Estructuras para detección y marcado posterior
                var modificaciones = new List<(int PageIndex, WordLocation Anchor, string Texto)>();
                var paginasModificadas = new List<int>();

                using (var fsOriginal = new FileStream(request.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var reader = new PdfReader(fsOriginal))
                using (var srcDoc = new PdfDocument(reader))
                {
                    PdfFont font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    int totalPaginas = srcDoc.GetNumberOfPages();

                    // 1) Detección (no se modifica el documento origen)
                    for (int i = 1; i <= totalPaginas; i++)
                    {
                        PdfPage page = srcDoc.GetPage(i);
                        var pageSize = page.GetPageSize();
                        float anchoPagina = pageSize.GetWidth();

                        float limiteIzquierdoColumnas = anchoPagina * 0.55f;
                        float limiteDerechoColumnas = anchoPagina * 0.85f;

                        var strategy = new FilaExtractionStrategy(pageSize.GetHeight());
                        var processor = new PdfCanvasProcessor(strategy);
                        processor.ProcessPageContent(page);

                        var filasDeLaPagina = strategy.ObtenerFilas();
                        bool dentroDeDetalleMovimientos = false;

                        foreach (var fila in filasDeLaPagina)
                        {
                            string textoCompletoFila = string.Join(" ", fila.Select(w => w.Text));

                            if (textoCompletoFila.Contains("DETALLE MOVIMIENTOS", StringComparison.OrdinalIgnoreCase))
                            {
                                dentroDeDetalleMovimientos = true;
                                continue;
                            }

                            if (textoCompletoFila.Contains("CARGOS OBJETADOS", StringComparison.OrdinalIgnoreCase) ||
                                textoCompletoFila.Contains("INFORMACION FISCAL", StringComparison.OrdinalIgnoreCase) ||
                                textoCompletoFila.Contains("ESTE PRODUCTO SE ENCUENTRA ENCUENTRA GARANTIZADO", StringComparison.OrdinalIgnoreCase))
                            {
                                dentroDeDetalleMovimientos = false;
                            }

                            if (!dentroDeDetalleMovimientos)
                                continue;

                            foreach (var regla in request.Reglas)
                            {
                                var itemFecha = fila.FirstOrDefault(w => w.Text.Contains(regla.Fecha, StringComparison.OrdinalIgnoreCase));

                                var itemPalabraClave = fila.FirstOrDefault(w =>
                                    w.Text.Contains(regla.PalabraClave, StringComparison.OrdinalIgnoreCase) &&
                                    w.Left >= limiteIzquierdoColumnas &&
                                    w.Left <= limiteDerechoColumnas
                                );

                                if (itemPalabraClave != null && itemFecha != null)
                                {
                                    modificaciones.Add((i, itemPalabraClave, regla.TextoAInsertar));
                                    if (!paginasModificadas.Contains(i))
                                        paginasModificadas.Add(i);

                                    // una regla aplicada en esta línea -> pasar a la siguiente regla
                                }
                            }
                        }
                    }

                    // Asegurar carpeta destino
                    string dirDestino = System.IO.Path.GetDirectoryName(request.OutputPath);
                    if (!string.IsNullOrEmpty(dirDestino) && !Directory.Exists(dirDestino))
                        Directory.CreateDirectory(dirDestino);

                    // 2) Escritura: copiar solo página 1 y páginas modificadas, aplicar marcas sobre las copias
                    using (var fsDestino = new FileStream(request.OutputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    using (var writer = new PdfWriter(fsDestino))
                    using (var dstDoc = new PdfDocument(writer))
                    {
                        void DibujarTextoEnPagina(PdfPage dstPage, WordLocation anchor, string texto)
                        {
                            var pageSize = dstPage.GetPageSize();
                            PdfCanvas canvas = new PdfCanvas(dstPage.NewContentStreamAfter(), dstPage.GetResources(), dstDoc);

                            float offsetHorizontal = 2f;
                            float x_iText = anchor.Left + anchor.Width + offsetHorizontal;

                            float y_iText = pageSize.GetHeight() - anchor.Top - anchor.Height;
                            y_iText += 1f;

                            float fontSizeBase = anchor.Height > 0 ? anchor.Height : 10f;
                            float fontSize = fontSizeBase + 8f;

                            canvas.BeginText()
                                  .SetFontAndSize(font, fontSize)
                                  .MoveText(x_iText, y_iText);

                            canvas.SetTextRenderingMode(PdfCanvasConstants.TextRenderingMode.FILL_STROKE);
                            canvas.SetLineWidth(0.5f);
                            canvas.SetStrokeColor(ColorConstants.BLACK);

                            canvas.ShowText(texto)
                                  .EndText();

                            canvas.Release();
                        }

                        // Copiar página 1 si existe (siempre se copia la primera página, igual que Multiva)
                        if (srcDoc.GetNumberOfPages() >= 1)
                        {
                            srcDoc.CopyPagesTo(1, 1, dstDoc);
                            var modsPagina1 = modificaciones.Where(m => m.PageIndex == 1).ToList();
                            if (modsPagina1.Any())
                            {
                                var dstPage = dstDoc.GetPage(dstDoc.GetNumberOfPages());
                                foreach (var mod in modsPagina1)
                                    DibujarTextoEnPagina(dstPage, mod.Anchor, mod.Texto);
                            }
                        }

                        // Copiar demás páginas modificadas (sin duplicar la 1)
                        var paginasParaCopiar = paginasModificadas
                            .Distinct()
                            .OrderBy(p => p)
                            .Where(p => p != 1)
                            .ToList();

                        foreach (var p in paginasParaCopiar)
                        {
                            srcDoc.CopyPagesTo(p, p, dstDoc);
                            var dstPage = dstDoc.GetPage(dstDoc.GetNumberOfPages());
                            var mods = modificaciones.Where(m => m.PageIndex == p).ToList();
                            foreach (var mod in mods)
                                DibujarTextoEnPagina(dstPage, mod.Anchor, mod.Texto);
                        }

                        // Mantener consistencia si necesitas la lista de páginas (opcional)
                        if (!paginasModificadas.Contains(1))
                            paginasModificadas.Insert(0, 1);
                        else
                        {
                            paginasModificadas = paginasModificadas.Distinct().OrderBy(p => paginasModificadas.IndexOf(p)).ToList();
                            paginasModificadas.Remove(1);
                            paginasModificadas.Insert(0, 1);
                        }
                    }

                    // --- CREAR LOG (BBVA) ---
                    try
                    {
                        var filasLog = (request.Reglas ?? Enumerable.Empty<ReglaInsercion>())
                            .Select(r => (Texto: r.TextoAInsertar ?? string.Empty, Exito: modificaciones.Any(m => m.Texto == r.TextoAInsertar)))
                            .ToList();

                        // Guardar Excel junto al outputPath con nombre fijo
                        string dir = System.IO.Path.GetDirectoryName(request.OutputPath) ?? "";
                        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                            Directory.CreateDirectory(dir);

                        string fileName = System.IO.Path.GetFileNameWithoutExtension(request.OutputPath) + "_log.xlsx";
                        string logPath = System.IO.Path.Combine(dir, fileName);

                        using (var wb = new ClosedXML.Excel.XLWorkbook())
                        {
                            var ws = wb.Worksheets.Add("Log");
                            ws.Cell(1, 1).Value = "TextoAInsertar";
                            ws.Cell(1, 2).Value = "Exito";

                            int r = 2;
                            foreach (var f in filasLog)
                            {
                                ws.Cell(r, 1).Value = f.Texto ?? "";
                                ws.Cell(r, 2).Value = f.Exito ? "TRUE" : "FALSE";
                                r++;
                            }

                            ws.Columns().AdjustToContents();
                            if (File.Exists(logPath))
                                File.Delete(logPath);
                            wb.SaveAs(logPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"No se pudo crear log Excel para {request.OutputPath}: {ex.Message}");
                        // No interrumpir el flujo principal por fallo en log
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en motor de renderizado Sabadell: {ex.Message}", ex);
            }
        }

        public bool ModificarPdfPorCoordenadasbanorte(EditPdfCoordinatesBnorteRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request), "La petición no puede ser nula.");

            if (!File.Exists(request.FilePath))
                throw new FileNotFoundException("El PDF de origen no existe.");

            try
            {
                string dirDestino = System.IO.Path.GetDirectoryName(request.OutputPath);
                if (!string.IsNullOrEmpty(dirDestino) && !Directory.Exists(dirDestino))
                    Directory.CreateDirectory(dirDestino);

                var modificaciones = new List<(int PageIndex, WordLocation Anchor, string Texto)>();
                var paginasModificadas = new HashSet<int>();
                var exitos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var paginaPorTexto = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var filasUsadas = new HashSet<(int PageIndex, int RowIndex)>();
                var reglasAplicadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                string NormalizarTextoBusqueda(string? valor)
                {
                    if (string.IsNullOrWhiteSpace(valor))
                        return string.Empty;

                    return Regex.Replace(valor.Trim(), @"\s+", " ").Replace(" ", " ").ToUpperInvariant();
                }

                string NormalizarFechaParaBusqueda(string? fecha)
                {
                    if (string.IsNullOrWhiteSpace(fecha))
                        return string.Empty;

                    var fechaTrim = fecha.Trim();

                    if (DateTime.TryParse(fechaTrim, CultureInfo.InvariantCulture, DateTimeStyles.None, out var fechaDate))
                    {
                        return fechaDate.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture).ToUpperInvariant();
                    }

                    if (DateTime.TryParseExact(fechaTrim, new[] { "d/M/yyyy", "dd/M/yyyy", "d/MM/yyyy", "dd/MM/yyyy" }, CultureInfo.InvariantCulture, DateTimeStyles.None, out fechaDate))
                    {
                        return fechaDate.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture).ToUpperInvariant();
                    }

                    var fechaLimpia = Regex.Replace(fechaTrim, @"[^0-9A-Z]", string.Empty).ToUpperInvariant();
                    if (fechaLimpia.Length >= 8)
                    {
                        var dia = fechaLimpia.Substring(0, 2);
                        var mes = fechaLimpia.Substring(2, 2);
                        var anno = fechaLimpia.Substring(4);
                        return $"{dia}-{mes}-{anno}";
                    }

                    return fechaTrim.ToUpperInvariant();
                }

                string NormalizarMontoParaBusqueda(string? valor)
                {
                    if (string.IsNullOrWhiteSpace(valor))
                        return string.Empty;

                    var numero = valor.Trim().Replace(",", ".");
                    var soloNumero = Regex.Replace(numero, @"[^0-9\.\-]", string.Empty);
                    return soloNumero;
                }

                bool TieneFechaEnTexto(string texto, string fechaBuscada)
                {
                    if (string.IsNullOrWhiteSpace(texto) || string.IsNullOrWhiteSpace(fechaBuscada))
                        return false;

                    var textoNormalizado = NormalizarTextoBusqueda(texto);
                    if (string.IsNullOrWhiteSpace(textoNormalizado))
                        return false;

                    var fechaNormalizada = NormalizarFechaParaBusqueda(fechaBuscada);
                    if (string.IsNullOrWhiteSpace(fechaNormalizada))
                        return false;

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

                bool TieneMontoEnTexto(string texto, string montoBuscado)
                {
                    if (string.IsNullOrWhiteSpace(texto) || string.IsNullOrWhiteSpace(montoBuscado))
                        return false;

                    var montoNormalizado = NormalizarMontoParaBusqueda(montoBuscado);
                    if (string.IsNullOrWhiteSpace(montoNormalizado))
                        return false;

                    var textoNormalizado = NormalizarTextoBusqueda(texto);
                    var textoSoloDigitos = Regex.Replace(textoNormalizado, @"[^0-9\.\-]", string.Empty);
                    return textoNormalizado.Contains(montoNormalizado, StringComparison.OrdinalIgnoreCase)
                        || textoSoloDigitos.Contains(montoNormalizado, StringComparison.OrdinalIgnoreCase);
                }

                bool TieneDescripcionEnTexto(string texto, string descripcionBuscada)
                {
                    if (string.IsNullOrWhiteSpace(texto))
                        return false;

                    var textoNormalizado = NormalizarTextoBusqueda(texto);
                    var descripcionNormalizada = NormalizarTextoBusqueda(descripcionBuscada);
                    if (string.IsNullOrWhiteSpace(descripcionNormalizada))
                        return true;

                    var descripcionCompacta = Regex.Replace(descripcionNormalizada, @"[^A-Z0-9]", string.Empty);
                    var textoCompacto = Regex.Replace(textoNormalizado, @"[^A-Z0-9]", string.Empty);

                    return textoNormalizado.Contains(descripcionNormalizada, StringComparison.OrdinalIgnoreCase)
                        || textoCompacto.Contains(descripcionCompacta, StringComparison.OrdinalIgnoreCase);
                }

                using (var fsOriginal = new FileStream(request.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var reader = new PdfReader(fsOriginal))
                using (var srcDoc = new PdfDocument(reader))
                {
                    int totalPaginas = srcDoc.GetNumberOfPages();

                    for (int i = 1; i <= totalPaginas; i++)
                    {
                        var page = srcDoc.GetPage(i);
                        var pageSize = page.GetPageSize();

                        var strategy = new FilaExtractionStrategy(pageSize.GetHeight());
                        var processor = new PdfCanvasProcessor(strategy);
                        processor.ProcessPageContent(page);

                        var filasDeLaPagina = strategy.ObtenerFilas();

                        for (int filaIndex = 0; filaIndex < filasDeLaPagina.Count; filaIndex++)
                        {
                            var fila = filasDeLaPagina[filaIndex];
                            if (fila == null || fila.Count == 0)
                                continue;

                            if (!filasUsadas.Add((i, filaIndex)))
                                continue;

                            string textoCompletoFila = string.Join(" ", fila.Select(w => w.Text ?? string.Empty)).Trim();

                            if (textoCompletoFila.Contains("DETALLE DE MOVIMIENTOS CUENTA DE CHEQUES", StringComparison.OrdinalIgnoreCase) ||
                                textoCompletoFila.Contains("DETALLE MOVIMIENTOS CUENTA DE CHEQUES", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            if (textoCompletoFila.Contains("DETALLES DE MOVIMIENTOS DINERO CRECIENTE", StringComparison.OrdinalIgnoreCase) ||
                                textoCompletoFila.Contains("INFORMACION FISCAL", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            var textosParaEvaluar = new List<string>();
                            textosParaEvaluar.Add(string.Join(" ", fila.Select(w => w.Text ?? string.Empty)));

                            for (int offset = -1; offset <= 1; offset++)
                            {
                                int idx = filaIndex + offset;
                                if (idx < 0 || idx >= filasDeLaPagina.Count)
                                    continue;

                                var bloque = filasDeLaPagina[idx];
                                string bloqueText = string.Join(" ", bloque.Select(w => w.Text ?? string.Empty));
                                if (!string.IsNullOrWhiteSpace(bloqueText))
                                    textosParaEvaluar.Add(bloqueText);
                            }

                            foreach (var regla in request.Reglas ?? Enumerable.Empty<ReglaInsercionBanorte>())
                            {
                                if (string.IsNullOrWhiteSpace(regla.TextoAInsertar))
                                    continue;
 
                                string reglaKey = $"{regla.Fecha ?? string.Empty}|{regla.PalabraClave ?? string.Empty}|{regla.Descripcion ?? string.Empty}|{regla.TextoAInsertar ?? string.Empty}";
                                Console.WriteLine($"[Banorte][RULE] key={reglaKey} fila={textoCompletoFila}");
 
                                bool reglaCoincide = false;
                                foreach (var textoFila in textosParaEvaluar.Distinct(StringComparer.OrdinalIgnoreCase))
                                {
                                    string textoFilaNormalizado = NormalizarTextoBusqueda(textoFila);
 
                                    bool lineaTieneFecha = TieneFechaEnTexto(textoFilaNormalizado, regla.Fecha);
                                    bool lineaTieneMonto = TieneMontoEnTexto(textoFilaNormalizado, regla.PalabraClave);
                                    bool lineaTieneDescripcion = TieneDescripcionEnTexto(textoFilaNormalizado, regla.Descripcion);
 
                                    if (lineaTieneFecha && lineaTieneMonto && lineaTieneDescripcion)
                                    {
                                        Console.WriteLine($"[Banorte][MATCH] regla={regla.TextoAInsertar} text={textoFila}");
                                        reglaCoincide = true;
                                        break;
                                    }
                                }
 
                                if (!reglaCoincide)
                                    continue;
 
                                if (!reglasAplicadas.Add(reglaKey))
                                    continue;
 
                                var elementosLinea = fila.OrderBy(e => e.Top).ThenBy(e => e.Left).ToList();
                                float filaTop = elementosLinea.Min(e => e.Top);
                                float filaBottom = elementosLinea.Max(e => e.Top + e.Height);
                                float filaHeight = Math.Max(filaBottom - filaTop, 8f);

                                const float centimetrosAlaIzquierda = 3f;
                                const float puntosPorCentimetro = 28.346456f;
                                float columnaFoleado = pageSize.GetWidth() * 0.85f - (centimetrosAlaIzquierda * puntosPorCentimetro);

                                var anchor = new WordLocation
                                {
                                    Text = regla.TextoAInsertar,
                                    Left = columnaFoleado,
                                    Top = filaTop,
                                    Width = 0f,
                                    Height = filaHeight
                                };

                                modificaciones.Add((i, anchor, regla.TextoAInsertar));
                                paginasModificadas.Add(i);
                                exitos.Add(regla.TextoAInsertar);
                                paginaPorTexto[regla.TextoAInsertar] = i;

                                break;
                            }
                        }
                    }
                }

                using (var fsOriginal2 = new FileStream(request.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var reader2 = new PdfReader(fsOriginal2))
                using (var srcDoc2 = new PdfDocument(reader2))
                using (var fsDestino = new FileStream(request.OutputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                using (var writerDest = new PdfWriter(fsDestino))
                using (var dstDoc = new PdfDocument(writerDest))
                {
                    var paginasParaCopiar = new List<int>();
                    if (srcDoc2.GetNumberOfPages() >= 1)
                        paginasParaCopiar.Add(1);

                    var adicionales = paginasModificadas.Distinct().OrderBy(x => x).Where(p => p != 1).ToList();
                    paginasParaCopiar.AddRange(adicionales);

                    foreach (var p in paginasParaCopiar)
                    {
                        if (p >= 1 && p <= srcDoc2.GetNumberOfPages())
                            srcDoc2.CopyPagesTo(p, p, dstDoc);
                    }

                    var mapping = paginasParaCopiar
                        .Select((orig, idx) => new { Orig = orig, Dst = idx + 1 })
                        .ToDictionary(x => x.Orig, x => x.Dst);

                    var modsPorPagina = modificaciones.GroupBy(m => m.PageIndex)
                        .ToDictionary(g => g.Key, g => g.ToList());

                    void DibujarTextoEnPagina(PdfPage dstPage, WordLocation anchor, string texto)
                    {
                        var pageSize = dstPage.GetPageSize();
                        PdfCanvas canvas = new PdfCanvas(dstPage.NewContentStreamAfter(), dstPage.GetResources(), dstDoc);

                        const float desplazamientoVerticalPuntos = 14f;
                        const float centimetrosAlaIzquierda = 3.5f;
                        const float puntosPorCentimetro = 28.346456f;
                        float x = pageSize.GetWidth() * 0.85f - (centimetrosAlaIzquierda * puntosPorCentimetro);
                        float y = pageSize.GetHeight() - anchor.Top - anchor.Height + 2f - desplazamientoVerticalPuntos;
                        float fontSize = Math.Max(anchor.Height > 0 ? anchor.Height + 6f : 10f, 10f);

                        canvas.SetFillColor(ColorConstants.BLACK);
                        canvas.SetStrokeColor(ColorConstants.BLACK);
                        canvas.SetLineWidth(0.4f);

                        canvas.BeginText()
                              .SetFontAndSize(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD), fontSize)
                              .MoveText(x, y);

                        canvas.SetTextRenderingMode(PdfCanvasConstants.TextRenderingMode.FILL);
                        canvas.ShowText(texto ?? string.Empty)
                              .EndText();

                        canvas.Release();
                    }

                    foreach (var kv in modsPorPagina)
                    {
                        int originalIndex = kv.Key;
                        if (!mapping.TryGetValue(originalIndex, out int dstIndex))
                            continue;

                        if (dstIndex < 1 || dstIndex > dstDoc.GetNumberOfPages())
                            continue;

                        var dstPage = dstDoc.GetPage(dstIndex);
                        foreach (var mod in kv.Value)
                        {
                            DibujarTextoEnPagina(dstPage, mod.Anchor, mod.Texto);
                        }
                    }
                }

                try
                {
                    var filasLog = (request.Reglas ?? Enumerable.Empty<ReglaInsercionBanorte>())
                        .Select(r =>
                        {
                            paginaPorTexto.TryGetValue(r.TextoAInsertar ?? string.Empty, out var pagina);
                            return (Texto: r.TextoAInsertar ?? string.Empty,
                                    Exito: exitos.Contains(r.TextoAInsertar ?? string.Empty),
                                    Pagina: pagina > 0 ? pagina.ToString() : string.Empty);
                        })
                        .ToList();

                    string dir = System.IO.Path.GetDirectoryName(request.OutputPath) ?? string.Empty;
                    if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        Directory.CreateDirectory(dir);

                    string fileName = System.IO.Path.GetFileNameWithoutExtension(request.OutputPath) + "_log.xlsx";
                    string logPath = System.IO.Path.Combine(dir, fileName);

                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("Log");
                        ws.Cell(1, 1).Value = "TextoAInsertar";
                        ws.Cell(1, 2).Value = "Exito";
                        ws.Cell(1, 3).Value = "Página";

                        int r = 2;
                        foreach (var f in filasLog)
                        {
                            ws.Cell(r, 1).Value = f.Texto ?? string.Empty;
                            ws.Cell(r, 2).Value = f.Exito ? "TRUE" : "FALSE";
                            ws.Cell(r, 3).Value = f.Pagina ?? string.Empty;
                            r++;
                        }

                        if (paginasModificadas.Any())
                        {
                            var wsPages = wb.Worksheets.Add("PaginasModificadas");
                            wsPages.Cell(1, 1).Value = "Pagina";
                            int pr = 2;
                            foreach (var p in paginasModificadas.Distinct().OrderBy(x => x))
                            {
                                wsPages.Cell(pr, 1).Value = p;
                                pr++;
                            }
                        }

                        ws.Columns().AdjustToContents();
                        if (File.Exists(logPath))
                            File.Delete(logPath);
                        wb.SaveAs(logPath);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"No se pudo crear log Excel para {request.OutputPath}: {ex.Message}");
                }

                return modificaciones.Count > 0;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error Crítico en Renderizado Geométrico Banorte: {ex.Message}", ex);
            }
        }

        public bool ModificarPdfPorCoordenadassantander(EditPdfCoordinatesRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request), "La petición no puede ser nula.");

            if (!File.Exists(request.FilePath))
                throw new FileNotFoundException("El PDF de origen no existe.");

            try
            {
                static bool ContieneTituloDineroCreciente(string texto)
                {
                    return texto.Contains("Detalles de movimientos Dinero Creciente Santander", StringComparison.OrdinalIgnoreCase)
                        || texto.Contains("DETALLES DE MOVIMIENTOS DINERO CRECIENTE", StringComparison.OrdinalIgnoreCase);
                }

                static string NormalizarFechaParaComparacion(string input)
                {
                    if (string.IsNullOrWhiteSpace(input))
                        return string.Empty;

                    string normalized = input.Trim();
                    normalized = normalized.Replace("de", " ", StringComparison.OrdinalIgnoreCase)
                                           .Replace("del", " ", StringComparison.OrdinalIgnoreCase);
                    normalized = normalized.ToUpperInvariant();
                    normalized = Regex.Replace(normalized, @"\s+", " ");

                    normalized = normalized.Replace("ENERO", "JAN")
                                           .Replace("JANUARY", "JAN")
                                           .Replace("FEBRERO", "FEB")
                                           .Replace("FEBRUARY", "FEB")
                                           .Replace("MARZO", "MAR")
                                           .Replace("MARCH", "MAR")
                                           .Replace("ABRIL", "APR")
                                           .Replace("APRIL", "APR")
                                           .Replace("MAYO", "MAY")
                                           .Replace("MAY", "MAY")
                                           .Replace("JUNIO", "JUN")
                                           .Replace("JUNE", "JUN")
                                           .Replace("JULIO", "JUL")
                                           .Replace("JULY", "JUL")
                                           .Replace("AGOSTO", "AUG")
                                           .Replace("AUGUST", "AUG")
                                           .Replace("SEPTIEMBRE", "SEP")
                                           .Replace("SEPTEMBER", "SEP")
                                           .Replace("OCTUBRE", "OCT")
                                           .Replace("OCTOBER", "OCT")
                                           .Replace("NOVIEMBRE", "NOV")
                                           .Replace("NOVEMBER", "NOV")
                                           .Replace("DICIEMBRE", "DEC")
                                           .Replace("DECEMBER", "DEC");

                    normalized = normalized.Replace("-", " ")
                                           .Replace("/", " ")
                                           .Replace(".", " ")
                                           .Replace(",", " ");
                    normalized = Regex.Replace(normalized, @"[^A-Z0-9 ]", " ");
                    normalized = Regex.Replace(normalized, @"\s+", string.Empty);
                    return normalized;
                }

                static bool CoincideFechaEnFila(IEnumerable<WordLocation> fila, string fechaBuscada)
                {
                    if (string.IsNullOrWhiteSpace(fechaBuscada))
                        return false;

                    string textoFila = string.Join(" ", fila.Select(w => w.Text ?? string.Empty));
                    string textoNormalizado = NormalizarFechaParaComparacion(textoFila);
                    if (string.IsNullOrEmpty(textoNormalizado))
                        return false;

                    var variantesFecha = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                    {
                        fechaBuscada,
                        fechaBuscada.Replace('/', '-'),
                        fechaBuscada.Replace('-', '/'),
                        fechaBuscada.Replace("-", string.Empty),
                        fechaBuscada.Replace("/", string.Empty),
                        Regex.Replace(fechaBuscada, @"[^0-9A-Z]", string.Empty)
                    };

                    foreach (var variante in variantesFecha)
                    {
                        string normalizada = NormalizarFechaParaComparacion(variante);
                        if (string.IsNullOrEmpty(normalizada))
                            continue;

                        if (textoNormalizado.Contains(normalizada, StringComparison.OrdinalIgnoreCase))
                            return true;

                        foreach (var token in fila)
                        {
                            string tokenNormalizado = NormalizarFechaParaComparacion(token.Text ?? string.Empty);
                            if (tokenNormalizado == normalizada || tokenNormalizado.Contains(normalizada, StringComparison.OrdinalIgnoreCase))
                                return true;
                        }
                    }

                    return false;
                }

                static decimal? ParseMontoDecimal(string input)
                {
                    if (string.IsNullOrWhiteSpace(input))
                        return null;

                    string cleaned = input.Trim();
                    cleaned = cleaned.Replace("$", string.Empty).Replace("€", string.Empty);
                    cleaned = cleaned.Replace(" ", string.Empty);
                    cleaned = Regex.Replace(cleaned, @"[^0-9,.-]", string.Empty);

                    if (string.IsNullOrEmpty(cleaned) || cleaned == "-" || cleaned == "." || cleaned == ",")
                        return null;

                    if (cleaned.Contains(',') && cleaned.Contains('.'))
                    {
                        if (cleaned.LastIndexOf('.') > cleaned.LastIndexOf(','))
                            cleaned = cleaned.Replace(",", string.Empty);
                        else
                            cleaned = cleaned.Replace(".", string.Empty).Replace(",", ".");
                    }
                    else if (cleaned.Contains(','))
                    {
                        int lastCommaIndex = cleaned.LastIndexOf(',');
                        int decimals = cleaned.Length - lastCommaIndex - 1;
                        if (decimals > 2)
                            cleaned = cleaned.Replace(",", string.Empty);
                        else
                            cleaned = cleaned.Replace(",", ".");
                    }
                    else if (cleaned.Contains('.'))
                    {
                        int lastDotIndex = cleaned.LastIndexOf('.');
                        int decimals = cleaned.Length - lastDotIndex - 1;
                        if (decimals > 2)
                            cleaned = cleaned.Replace(".", string.Empty);
                    }

                    if (decimal.TryParse(cleaned, NumberStyles.Number, CultureInfo.InvariantCulture, out var value))
                        return value;

                    return null;
                }

                static List<WordLocation> ObtenerTokensMontoExactoEnFila(IEnumerable<WordLocation> fila, string montoBuscado)
                {
                    if (string.IsNullOrWhiteSpace(montoBuscado))
                        return new List<WordLocation>();

                    decimal? montoEsperado = ParseMontoDecimal(montoBuscado);
                    if (!montoEsperado.HasValue)
                        return new List<WordLocation>();

                    var tokensNumericos = fila
                        .Where(w => !string.IsNullOrWhiteSpace(w?.Text))
                        .Select(w => new { Token = w, Monto = ParseMontoDecimal(w.Text) })
                        .Where(x => x.Monto.HasValue)
                        .OrderBy(x => x.Token.Left)
                        .ToList();

                    if (tokensNumericos.Count == 0)
                        return new List<WordLocation>();

                    var exactos = tokensNumericos
                        .Where(x => x.Monto.Value == montoEsperado.Value)
                        .Select(x => x.Token)
                        .OrderBy(t => t.Left)
                        .ThenBy(t => t.Top)
                        .ToList();

                    if (exactos.Count == 0)
                        return new List<WordLocation>();

                    var ultimoTokenNumerico = tokensNumericos.Last();
                    var exactosSinSaldoFinal = exactos
                        .Where(t => !(Math.Abs(t.Left - ultimoTokenNumerico.Token.Left) < 0.1f && Math.Abs(t.Top - ultimoTokenNumerico.Token.Top) < 0.1f))
                        .ToList();

                    if (exactosSinSaldoFinal.Count > 0)
                        exactos = exactosSinSaldoFinal;

                    return exactos;
                }

                static bool CoincideMontoEnFila(IEnumerable<WordLocation> fila, string montoBuscado)
                {
                    if (string.IsNullOrWhiteSpace(montoBuscado))
                        return false;

                    decimal? montoEsperado = ParseMontoDecimal(montoBuscado);
                    if (!montoEsperado.HasValue)
                        return false;

                    return fila
                        .Where(w => !string.IsNullOrWhiteSpace(w?.Text))
                        .Select(w => ParseMontoDecimal(w.Text))
                        .Any(m => m.HasValue && m.Value == montoEsperado.Value);
                }

                static bool EsTokenMontoCoincidente(string textoToken, decimal? montoEsperado)
                {
                    if (string.IsNullOrWhiteSpace(textoToken) || !montoEsperado.HasValue)
                        return false;

                    decimal? montoToken = ParseMontoDecimal(textoToken);
                    if (!montoToken.HasValue)
                        return false;

                    return montoToken.Value == montoEsperado.Value;
                }

                static WordLocation ObtenerTokenMontoEnFila(IEnumerable<WordLocation> fila, string montoBuscado)
                {
                    if (string.IsNullOrWhiteSpace(montoBuscado))
                        return null;

                    decimal? montoEsperado = ParseMontoDecimal(montoBuscado);
                    if (!montoEsperado.HasValue)
                        return null;

                    var tokensNumericos = fila
                        .Where(w => !string.IsNullOrWhiteSpace(w?.Text))
                        .Select(w => new { Token = w, Monto = ParseMontoDecimal(w.Text) })
                        .Where(x => x.Monto.HasValue)
                        .OrderBy(x => x.Token.Left)
                        .ToList();

                    if (tokensNumericos.Count == 0)
                        return null;

                    var exactos = tokensNumericos
                        .Where(x => x.Monto.Value == montoEsperado.Value)
                        .Select(x => x.Token)
                        .OrderBy(t => t.Left)
                        .ThenBy(t => t.Top)
                        .ToList();

                    if (exactos.Count == 0)
                        return null;

                    var ultimoTokenNumerico = tokensNumericos.Last().Token;
                    var exactosSinSaldoFinal = exactos
                        .Where(t => !(Math.Abs(t.Left - ultimoTokenNumerico.Left) < 0.1f && Math.Abs(t.Top - ultimoTokenNumerico.Top) < 0.1f))
                        .ToList();

                    return (exactosSinSaldoFinal.Count > 0 ? exactosSinSaldoFinal.First() : exactos.First());
                }

                string dirDestino = System.IO.Path.GetDirectoryName(request.OutputPath);
                if (!string.IsNullOrEmpty(dirDestino) && !Directory.Exists(dirDestino))
                {
                    Directory.CreateDirectory(dirDestino);
                }

                var exitos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var paginaPorTexto = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var paginasModificadas = new HashSet<int>();
                var reglasAplicadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                using (FileStream fsOriginal = new FileStream(request.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (FileStream fsDestino = new FileStream(request.OutputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    PdfReader reader = new PdfReader(fsOriginal);
                    PdfWriter writer = new PdfWriter(fsDestino);

                    using (PdfDocument pdfDoc = new PdfDocument(reader, writer))
                    {
                        PdfFont font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                        int totalPaginas = pdfDoc.GetNumberOfPages();

                        bool dentroDeDineroCreciente = false;

                        for (int i = 1; i <= totalPaginas; i++)
                        {
                            PdfPage page = pdfDoc.GetPage(i);
                            var pageSize = page.GetPageSize();

                            var strategy = new FilaExtractionStrategy(pageSize.GetHeight());
                            var processor = new PdfCanvasProcessor(strategy);
                            processor.ProcessPageContent(page);

                            var filasDeLaPagina = strategy.ObtenerFilas();

                            for (int filaIndex = 0; filaIndex < filasDeLaPagina.Count; filaIndex++)
                            {
                                var fila = filasDeLaPagina[filaIndex];
                                var bloqueFila = new List<WordLocation>();
                                for (int offset = -1; offset <= 1; offset++)
                                {
                                    int idx = filaIndex + offset;
                                    if (idx >= 0 && idx < filasDeLaPagina.Count)
                                        bloqueFila.AddRange(filasDeLaPagina[idx]);
                                }

                                string textoCompletoFila = string.Join(" ", fila.Select(w => w.Text ?? string.Empty)).Trim();
                                string textoBloque = string.Join(" ", bloqueFila.Select(w => w.Text ?? string.Empty)).Trim();
                                if (string.IsNullOrEmpty(textoCompletoFila))
                                    continue;

                                if (ContieneTituloDineroCreciente(textoCompletoFila))
                                {
                                    dentroDeDineroCreciente = true;
                                    Console.WriteLine($"[Santander][Debug][Pagina {i}] TITULO DETECTADO: {textoCompletoFila}");
                                    continue;
                                }

                                if (!dentroDeDineroCreciente)
                                    continue;

                                Console.WriteLine($"[Santander][Debug][Pagina {i}] Fila leída: {textoCompletoFila}");

                                bool esCabeceraSeccion = textoCompletoFila.Contains("DEPOSITOS", StringComparison.OrdinalIgnoreCase)
                                    || textoCompletoFila.Contains("RETIROS", StringComparison.OrdinalIgnoreCase);
                                if (esCabeceraSeccion)
                                    continue;

                                foreach (var regla in request.Reglas ?? Enumerable.Empty<ReglaInsercion>())
                                {
                                    if (string.IsNullOrWhiteSpace(regla.TextoAInsertar))
                                        continue;

                                    string reglaKey = $"{regla.Fecha ?? string.Empty}|{regla.PalabraClave ?? string.Empty}|{regla.TextoAInsertar ?? string.Empty}";
                                    if (reglasAplicadas.Contains(reglaKey))
                                        continue;

                                    string fechaBuscada = regla.Fecha?.Trim() ?? string.Empty;
                                    if (string.IsNullOrEmpty(fechaBuscada))
                                        continue;

                                    bool tieneFecha = CoincideFechaEnFila(bloqueFila, fechaBuscada);
                                    bool tieneMonto = CoincideMontoEnFila(bloqueFila, regla.PalabraClave ?? string.Empty);

                                    Console.WriteLine($"[Santander][Eval] Fecha={fechaBuscada} Monto={regla.PalabraClave} FechaOK={tieneFecha} MontoOK={tieneMonto} Fila={textoBloque}");

                                    if (!tieneFecha || !tieneMonto)
                                        continue;

                                    var itemMonto = ObtenerTokenMontoEnFila(bloqueFila, regla.PalabraClave ?? string.Empty);

                                    if (itemMonto == null)
                                        continue;

                                    if (!reglasAplicadas.Add(reglaKey))
                                        continue;

                                    PdfCanvas canvas = new PdfCanvas(page);

                                    float offsetHorizontal = 2f;
                                    float x_iText = itemMonto.Left + itemMonto.Width + offsetHorizontal;
                                    float y_iText = pageSize.GetHeight() - itemMonto.Top - itemMonto.Height + 1f;

                                    float fontSizeBase = itemMonto.Height > 0 ? itemMonto.Height : 10f;
                                    float fontSize = fontSizeBase + 6f;

                                    canvas.BeginText()
                                          .SetFontAndSize(font, fontSize)
                                          .MoveText(x_iText, y_iText);

                                    canvas.SetTextRenderingMode(PdfCanvasConstants.TextRenderingMode.FILL_STROKE);
                                    canvas.SetLineWidth(0.5f);
                                    canvas.SetStrokeColor(ColorConstants.BLACK);

                                    canvas.ShowText(regla.TextoAInsertar)
                                          .EndText();

                                    canvas.Release();

                                    Console.WriteLine($"[Santander][MATCH][Pagina {i}] Fecha={fechaBuscada} Monto={regla.PalabraClave} Texto={regla.TextoAInsertar} Fila={textoBloque}");

                                    if (!string.IsNullOrEmpty(regla.TextoAInsertar))
                                    {
                                        exitos.Add(regla.TextoAInsertar);
                                        paginaPorTexto[regla.TextoAInsertar] = i;
                                    }

                                    paginasModificadas.Add(i);
                                    break;
                                }
                            }
                        }

                        pdfDoc.Close();
                    }
                }

                string dir = System.IO.Path.GetDirectoryName(request.OutputPath) ?? "";
                string fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(request.OutputPath);
                string ext = System.IO.Path.GetExtension(request.OutputPath);
                string resumenPath = System.IO.Path.Combine(dir, $"{fileNameWithoutExt}_resumen{ext}");

                bool resumenGeneradoConExito = false;

                try
                {
                    var paginasAExtraer = new List<int> { 1 };
                    foreach (int p in paginasModificadas.OrderBy(p => p))
                    {
                        if (p != 1)
                        {
                            paginasAExtraer.Add(p);
                        }
                    }

                    using (PdfReader resReader = new PdfReader(request.OutputPath))
                    using (PdfWriter resWriter = new PdfWriter(resumenPath))
                    using (PdfDocument srcDoc = new PdfDocument(resReader))
                    using (PdfDocument destDoc = new PdfDocument(resWriter))
                    {
                        srcDoc.CopyPagesTo(paginasAExtraer, destDoc);
                    }
                    resumenGeneradoConExito = true;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Error al generar el PDF recortado de resumen: {ex.Message}");
                }

                if (resumenGeneradoConExito)
                {
                    try
                    {
                        if (File.Exists(request.OutputPath))
                        {
                            File.Delete(request.OutputPath);
                        }

                        if (File.Exists(resumenPath))
                        {
                            File.Move(resumenPath, request.OutputPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error al suplantar el PDF completo por el resumen: {ex.Message}");
                    }
                }

                try
                {
                    var filasLog = (request.Reglas ?? Enumerable.Empty<ReglaInsercion>())
                        .Select(r =>
                        {
                            paginaPorTexto.TryGetValue(r.TextoAInsertar ?? string.Empty, out var pagina);
                            return (Texto: r.TextoAInsertar ?? string.Empty,
                                    Exito: exitos.Contains(r.TextoAInsertar ?? string.Empty),
                                    Pagina: pagina > 0 ? pagina.ToString() : string.Empty);
                        })
                        .ToList();

                    string fileName = System.IO.Path.GetFileNameWithoutExtension(request.OutputPath) + "_log.xlsx";
                    string logPath = System.IO.Path.Combine(dir, fileName);

                    using (var wb = new XLWorkbook())
                    {
                        var ws = wb.Worksheets.Add("Log");
                        ws.Cell(1, 1).Value = "TextoAInsertar";
                        ws.Cell(1, 2).Value = "Exito";
                        ws.Cell(1, 3).Value = "Página";

                        int r = 2;
                        foreach (var f in filasLog)
                        {
                            ws.Cell(r, 1).Value = f.Texto ?? "";
                            ws.Cell(r, 2).Value = f.Exito ? "TRUE" : "FALSE";
                            ws.Cell(r, 3).Value = f.Pagina ?? string.Empty;
                            r++;
                        }

                        ws.Columns().AdjustToContents();
                        if (File.Exists(logPath))
                            File.Delete(logPath);
                        wb.SaveAs(logPath);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"No se pudo crear log Excel para {request.OutputPath}: {ex.Message}");
                }

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en motor de renderizado Santander: {ex.Message}", ex);
            }
        }

        public List<int> ModificarPdfPorCoordenadasmultiva(EditPdfCoordinatesRequest request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request), "La petición no puede ser nula.");

            if (!File.Exists(request.FilePath))
                throw new FileNotFoundException("El PDF de origen no existe.");

            var paginasModificadas = new List<int>();

            // Aseguramos carpeta destino
            string dirDestino = System.IO.Path.GetDirectoryName(request.OutputPath);
            if (!string.IsNullOrEmpty(dirDestino) && !Directory.Exists(dirDestino))
                Directory.CreateDirectory(dirDestino);

            try
            {
                // Estructura para guardar la información necesaria para estampar tras copiar
                var modificaciones = new List<(int PageIndex, WordLocation Anchor, string Texto)>();
                var reglasDetectadas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var paginaPorTexto = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Abrimos origen
                using (var fsOriginal = new FileStream(request.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                using (var reader = new PdfReader(fsOriginal))
                using (var srcDoc = new PdfDocument(reader))
                {
                    PdfFont font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    int totalPaginas = srcDoc.GetNumberOfPages();

                    // Primera pasada: detectar páginas que deben modificarse (no copiamos todavía)
                    for (int i = 1; i <= totalPaginas; i++)
                    {
                        var page = srcDoc.GetPage(i);
                        var pageSize = page.GetPageSize();

                        var strategy = new FilaExtractionStrategy(pageSize.GetHeight());
                        var processor = new PdfCanvasProcessor(strategy);
                        processor.ProcessPageContent(page);

                        var lineasFisicas = strategy.ObtenerFilas()
                            .OrderByDescending(fila => fila.FirstOrDefault()?.Top ?? 0)
                            .Select(fila => fila.AsEnumerable())
                            .ToList();

                        foreach (var grupoLinea in lineasFisicas)
                        {
                            var elementosLinea = grupoLinea.OrderBy(e => e.Left).ToList();
                            string textoLineaContinuo = string.Join("", elementosLinea.Select(e => e.Text));
                            string textoLineaSinEspacios = textoLineaContinuo.Replace(" ", "").ToUpper();

                            if (string.IsNullOrEmpty(textoLineaSinEspacios))
                                continue;

                            foreach (var regla in request.Reglas)
                            {
                                string fechaJson = regla.Fecha?.Replace("-", "").Replace("/", "").Replace(" ", "").Trim().ToUpper() ?? "";
                                if (string.IsNullOrEmpty(fechaJson)) continue;

                                string diaBuscado = fechaJson.Substring(0, Math.Min(2, fechaJson.Length));
                                string mesTexto = fechaJson.Length > 2 ? fechaJson.Substring(2, Math.Min(3, fechaJson.Length - 2)) : "";

                                string mesNumerico = "10";
                                if (mesTexto.Contains("ENE")) mesNumerico = "01";
                                if (mesTexto.Contains("FEB")) mesNumerico = "02";
                                if (mesTexto.Contains("MAR")) mesNumerico = "03";
                                if (mesTexto.Contains("ABR")) mesNumerico = "04";
                                if (mesTexto.Contains("MAY")) mesNumerico = "05";
                                if (mesTexto.Contains("JUN")) mesNumerico = "06";
                                if (mesTexto.Contains("JUL")) mesNumerico = "07";
                                if (mesTexto.Contains("AGO")) mesNumerico = "08";
                                if (mesTexto.Contains("SEP")) mesNumerico = "09";
                                if (mesTexto.Contains("OCT")) mesNumerico = "10";
                                if (mesTexto.Contains("NOV")) mesNumerico = "11";
                                if (mesTexto.Contains("DIC")) mesNumerico = "12";

                                string patronTextoBuscado = $"{diaBuscado}{mesTexto}";
                                string patronSlashBuscado = $"{diaBuscado}{mesNumerico}";

                                bool lineaTieneFecha = elementosLinea.Any(e =>
                                    Regex.IsMatch(e.Text ?? string.Empty, $@"(?<!\d){Regex.Escape(diaBuscado)}(?!\d)") ||
                                    (!string.IsNullOrEmpty(mesTexto) &&
                                     (e.Text.Contains(patronTextoBuscado, StringComparison.OrdinalIgnoreCase) ||
                                      e.Text.Contains(patronSlashBuscado, StringComparison.OrdinalIgnoreCase))));

                                string montoReglaLimpio = Regex.Replace(regla.PalabraClave ?? "", @"[^\d]", "");
                                string montoFilaLimpio = Regex.Replace(textoLineaSinEspacios, @"[^\d]", "");

                                bool lineaTieneMonto = !string.IsNullOrEmpty(montoReglaLimpio) && montoFilaLimpio.Contains(montoReglaLimpio);

                                if (lineaTieneFecha && lineaTieneMonto)
                                {
                                    // localizamos el ancla: priorizar fragmento numérico que contenga los dígitos exactos del monto buscado
                                    var candidatosMonto = elementosLinea.Where(e =>
                                        e.Left > pageSize.GetWidth() * 0.38f &&
                                        e.Left < pageSize.GetWidth() * 0.82f &&
                                        Regex.IsMatch(e.Text, @"[\d]")).ToList();

                                    WordLocation itemAncla = null;

                                    // 1) buscar en la zona por fragmento que contenga los dígitos del monto
                                    if (!string.IsNullOrEmpty(montoReglaLimpio))
                                    {
                                        itemAncla = candidatosMonto
                                            .FirstOrDefault(e => Regex.Replace(e.Text, @"[^\d]", "").Contains(montoReglaLimpio));
                                    }

                                    // 2) si no se encontró, buscar en toda la línea por el fragmento que contenga los dígitos
                                    if (itemAncla == null && !string.IsNullOrEmpty(montoReglaLimpio))
                                    {
                                        itemAncla = elementosLinea
                                            .FirstOrDefault(e => Regex.Replace(e.Text, @"[^\d]", "").Contains(montoReglaLimpio));
                                    }

                                    // 3) fallback: último candidato en la zona o último elemento de la línea (comportamiento anterior)
                                    if (itemAncla == null)
                                    {
                                        itemAncla = candidatosMonto.LastOrDefault() ?? elementosLinea.LastOrDefault();
                                    }

                                    if (itemAncla != null)
                                    {
                                        if (!string.IsNullOrEmpty(regla.TextoAInsertar))
                                        {
                                            reglasDetectadas.Add(regla.TextoAInsertar);
                                            paginaPorTexto[regla.TextoAInsertar] = i;
                                        }

                                        modificaciones.Add((i, itemAncla, regla.TextoAInsertar));
                                        if (!paginasModificadas.Contains(i))
                                            paginasModificadas.Add(i);
                                    }
                                }

                            }
                        }
                    } // fin detección

                    // Ahora abrimos writer para crear archivo destino y copiar en el orden deseado
                    using (var fsDestino = new FileStream(request.OutputPath, FileMode.Create, FileAccess.Write, FileShare.None))
                    using (var writer = new PdfWriter(fsDestino))
                    using (var dstDoc = new PdfDocument(writer))
                    {
                        // Helper local para dibujar texto junto al anchor con ajuste vertical
                        void DibujarTextoEnPagina(PdfPage dstPage, WordLocation anchor, string texto)
                        {
                            var pageSize = dstPage.GetPageSize();
                            PdfCanvas canvas = new PdfCanvas(dstPage.NewContentStreamAfter(), dstPage.GetResources(), dstDoc);

                            // Offset reducido para que quede prácticamente al lado del anchor
                            float offsetHorizontal = 1f;
                            float x_iText = anchor.Left + anchor.Width + offsetHorizontal;

                            // Baseline original
                            float baselineY = pageSize.GetHeight() - anchor.Top - anchor.Height;

                            // Restaurar tamaño original (como antes): altura del anchor + 7
                            float fontSizeBase = anchor.Height > 0 ? anchor.Height : 10f;
                            float fontSize = fontSizeBase + 7f;

                            // Alineación vertical usando baseline (igual que en la versión original)
                            float y_iText = baselineY + 1f;

                            canvas.BeginText()
                                  .SetFontAndSize(font, fontSize)
                                  .MoveText(x_iText, y_iText);

                            canvas.SetTextRenderingMode(PdfCanvasConstants.TextRenderingMode.FILL_STROKE);
                            canvas.SetLineWidth(0.4f);
                            canvas.SetStrokeColor(ColorConstants.BLACK);

                            canvas.ShowText(texto)
                                  .EndText();

                            canvas.Release();
                        }

                        // 1) Siempre copiar página 1 primero si existe
                        if (srcDoc.GetNumberOfPages() >= 1)
                        {
                            srcDoc.CopyPagesTo(1, 1, dstDoc);
                            // Si hay modificaciones para la página 1, las aplicamos sobre la copia
                            var modsPagina1 = modificaciones.Where(m => m.PageIndex == 1).ToList();
                            if (modsPagina1.Any())
                            {
                                var dstPage = dstDoc.GetPage(dstDoc.GetNumberOfPages());
                                foreach (var mod in modsPagina1)
                                {
                                    DibujarTextoEnPagina(dstPage, mod.Anchor, mod.Texto);
                                }
                            }
                        }

                        // 2) Copiar demás páginas modificadas (sin duplicar la 1) y aplicar las marcas
                        var paginasParaCopiar = paginasModificadas
                            .Distinct()
                            .OrderBy(p => p)
                            .Where(p => p != 1)
                            .ToList();

                        foreach (var p in paginasParaCopiar)
                        {
                            srcDoc.CopyPagesTo(p, p, dstDoc);
                            var dstPage = dstDoc.GetPage(dstDoc.GetNumberOfPages());

                            var mods = modificaciones.Where(m => m.PageIndex == p).ToList();
                            foreach (var mod in mods)
                            {
                                DibujarTextoEnPagina(dstPage, mod.Anchor, mod.Texto);
                            }
                        }

                        // Aseguramos que la lista devuelta tenga la página 1 al inicio
                        if (!paginasModificadas.Contains(1))
                            paginasModificadas.Insert(0, 1);
                        else
                        {
                            paginasModificadas = paginasModificadas.Distinct().OrderBy(p => paginasModificadas.IndexOf(p)).ToList();
                            paginasModificadas.Remove(1);
                            paginasModificadas.Insert(0, 1);
                        }
                    }

                    // --- CREAR LOG (Multiva) ---
                    try
                    {
                        var filasLog = (request.Reglas ?? Enumerable.Empty<ReglaInsercion>())
                            .Select(r =>
                            {
                                paginaPorTexto.TryGetValue(r.TextoAInsertar ?? string.Empty, out var pagina);
                                return (Texto: r.TextoAInsertar ?? string.Empty,
                                        Exito: !string.IsNullOrEmpty(r.TextoAInsertar) && reglasDetectadas.Contains(r.TextoAInsertar),
                                        Pagina: pagina > 0 ? pagina.ToString() : string.Empty);
                            })
                            .ToList();

                        string dir = System.IO.Path.GetDirectoryName(request.OutputPath) ?? "";
                        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                            Directory.CreateDirectory(dir);

                        string fileName = System.IO.Path.GetFileNameWithoutExtension(request.OutputPath) + "_log.xlsx";
                        string logPath = System.IO.Path.Combine(dir, fileName);

                        using (var wb = new XLWorkbook())
                        {
                            var ws = wb.Worksheets.Add("Log");
                            ws.Cell(1, 1).Value = "TextoAInsertar";
                            ws.Cell(1, 2).Value = "Exito";
                            ws.Cell(1, 3).Value = "Página";

                            int r = 2;
                            foreach (var f in filasLog)
                            {
                                ws.Cell(r, 1).Value = f.Texto ?? "";
                                ws.Cell(r, 2).Value = f.Exito ? "TRUE" : "FALSE";
                                ws.Cell(r, 3).Value = f.Pagina ?? string.Empty;
                                r++;
                            }

                            ws.Columns().AdjustToContents();
                            if (File.Exists(logPath))
                                File.Delete(logPath);
                            wb.SaveAs(logPath);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"No se pudo crear log Excel para {request.OutputPath}: {ex.Message}");
                        // No interrumpir el flujo principal por fallo en log
                    }

                    return paginasModificadas;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en motor de renderizado geométrico de PDF (Multiva): {ex.Message}", ex);
            }
        }
    }

    public class FilaExtractionStrategy : LocationTextExtractionStrategy
    {
        private readonly List<WordLocation> _words = new List<WordLocation>();
        private readonly float _pageHeight;

        public FilaExtractionStrategy(float pageHeight)
        {
            _pageHeight = pageHeight;
        }

        public override void EventOccurred(iText.Kernel.Pdf.Canvas.Parser.Data.IEventData data, EventType type)
        {
            base.EventOccurred(data, type);

            if (type.Equals(EventType.RENDER_TEXT))
            {
                var renderInfo = (iText.Kernel.Pdf.Canvas.Parser.Data.TextRenderInfo)data;
                string text = renderInfo.GetText();

                if (string.IsNullOrWhiteSpace(text)) return;

                var baseline = renderInfo.GetBaseline().GetStartPoint();
                var topRight = renderInfo.GetAscentLine().GetEndPoint();

                float left = baseline.Get(0);
                float top = _pageHeight - topRight.Get(1);
                float width = topRight.Get(0) - left;
                float height = topRight.Get(1) - baseline.Get(1);

                _words.Add(new WordLocation
                {
                    Text = text,
                    Left = left,
                    Top = top,
                    Width = width,
                    Height = height
                });
            }
        }

        public List<List<WordLocation>> ObtenerFilas()
        {
            float toleranciaFila = 7f;

            var palabrasOrdenadas = _words
                .Where(w => !string.IsNullOrWhiteSpace(w.Text))
                .OrderBy(w => w.Top)
                .ThenBy(w => w.Left)
                .ToList();

            var tokensFusionados = new List<WordLocation>();
            foreach (var word in palabrasOrdenadas)
            {
                var texto = word.Text.Trim();
                if (string.IsNullOrEmpty(texto))
                    continue;

                var token = new WordLocation
                {
                    Text = texto,
                    Left = word.Left,
                    Top = word.Top,
                    Width = word.Width,
                    Height = word.Height
                };

                if (tokensFusionados.Count > 0)
                {
                    var anterior = tokensFusionados[^1];
                    float gap = token.Left - (anterior.Left + anterior.Width);
                    bool mismoToken = Math.Abs(token.Top - anterior.Top) <= 4f && gap <= 8f && gap >= -8f;

                    if (mismoToken)
                    {
                        anterior.Text += token.Text;
                        anterior.Width = Math.Max(anterior.Width, token.Left + token.Width - anterior.Left);
                        anterior.Height = Math.Max(anterior.Height, token.Height);
                        anterior.Top = Math.Min(anterior.Top, token.Top);
                        continue;
                    }
                }

                tokensFusionados.Add(token);
            }

            var filas = new List<List<WordLocation>>();
            foreach (var word in tokensFusionados)
            {
                var filaExistente = filas.LastOrDefault(f =>
                    f.Count > 0 &&
                    Math.Abs(f.Last().Top - word.Top) <= toleranciaFila);

                if (filaExistente != null)
                {
                    filaExistente.Add(word);
                }
                else
                {
                    filas.Add(new List<WordLocation> { word });
                }
            }

            foreach (var fila in filas)
            {
                fila.Sort((a, b) => a.Left.CompareTo(b.Left));
            }

            return filas;
        }
    }   
}