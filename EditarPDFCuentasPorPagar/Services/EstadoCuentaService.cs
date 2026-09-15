using System.Text.RegularExpressions;
using ClosedXML.Excel;
using EditarPDFCuentasPorPagar.Models;
using iText.IO.Font.Constants;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Data;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace EditarPDFCuentasPorPagar.Services
{
    public class EstadoCuentaService : IEstadoCuentaService
    {
        public void ProcesarYFolearPdf(ProcesarEstadoCuentaDto request)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            if (string.IsNullOrWhiteSpace(request.FilePath))
                throw new ArgumentException("La ruta del archivo PDF es obligatoria.", nameof(request));

            if (!File.Exists(request.FilePath))
                throw new FileNotFoundException("El PDF de origen no existe.", request.FilePath);

            if (string.IsNullOrWhiteSpace(request.OutputPath))
                throw new ArgumentException("La ruta de salida del PDF es obligatoria.", nameof(request));

            var outputDirectory = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(request.OutputPath));
            if (!string.IsNullOrEmpty(outputDirectory) && !Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            using var reader = new PdfReader(request.FilePath);
            using var sourceDocument = new PdfDocument(reader);

            var pageTextLocations = new Dictionary<int, List<PdfTextLocation>>();
            for (var pageNumber = 1; pageNumber <= sourceDocument.GetNumberOfPages(); pageNumber++)
            {
                var page = sourceDocument.GetPage(pageNumber);
                var strategy = new PdfWordExtractionStrategy(page.GetPageSize().GetHeight());
                var processor = new PdfCanvasProcessor(strategy);
                processor.ProcessPageContent(page);
                pageTextLocations[pageNumber] = strategy.GetLocations();
            }

            var paginasFoleadas = new HashSet<int>();
            var auditoria = new List<ReglaAuditRecord>();
            var coincidencias = new List<PdfMatch>();
            var reglas = request.Reglas ?? Enumerable.Empty<ReglaFoleoDto>();

            var indiceRegla = 0;
            foreach (var regla in reglas)
            {
                indiceRegla++;
                var match = BuscarPrimeraCoincidenciaJerarquica(pageTextLocations, regla);
                if (match == null)
                {
                    auditoria.Add(new ReglaAuditRecord
                    {
                        NumeroRegla = indiceRegla,
                        Regla = regla.TextoAInsertar,
                        Caso = "Sin coincidencia",
                        Pagina = 0,
                        X = 0,
                        Y = 0,
                        Resultado = "No encontrado",
                        Exito = false,
                        Fecha = regla.Fecha,
                        PalabraClave = regla.PalabraClave,
                        GuiId = regla.GuiID,
                        RContable = regla.RContable
                    });
                    continue;
                }

                paginasFoleadas.Add(match.PageNumber);
                coincidencias.Add(match);
                auditoria.Add(new ReglaAuditRecord
                {
                    NumeroRegla = indiceRegla,
                    Regla = regla.TextoAInsertar,
                    Caso = match.Caso,
                    Pagina = match.PageNumber,
                    X = match.Location.Left,
                    Y = match.Location.Top,
                    Resultado = "Procesado",
                    Exito = true,
                    Fecha = regla.Fecha,
                    PalabraClave = regla.PalabraClave,
                    GuiId = regla.GuiID,
                    RContable = regla.RContable
                });
            }

            GenerarLogAuditoria(request.OutputPath, auditoria);

            using var outputStream = new FileStream(request.OutputPath, FileMode.Create, FileAccess.Write, FileShare.None);
            using var writer = new PdfWriter(outputStream);
            using var outputDocument = new PdfDocument(writer);

            if (paginasFoleadas.Count > 0)
            {
                var paginasOrdenadas = paginasFoleadas.OrderBy(x => x).ToList();
                sourceDocument.CopyPagesTo(paginasOrdenadas, outputDocument);

                foreach (var match in coincidencias)
                {
                    var nuevaPagina = paginasOrdenadas.IndexOf(match.PageNumber) + 1;
                    if (nuevaPagina <= 0)
                        continue;

                    EstamparTexto(outputDocument, nuevaPagina, match.Location, match.TextoAInsertar);
                }
            }
        }

        private static PdfMatch? BuscarPrimeraCoincidenciaJerarquica(Dictionary<int, List<PdfTextLocation>> pageTextLocations, ReglaFoleoDto regla)
        {
            foreach (var kvp in pageTextLocations.OrderBy(k => k.Key))
            {
                var result = BuscarGuiId(kvp.Value, regla.GuiID);
                if (result != null)
                    return new PdfMatch { PageNumber = kvp.Key, Location = result, Caso = "Caso 1", TextoAInsertar = regla.TextoAInsertar };
            }

            foreach (var kvp in pageTextLocations.OrderBy(k => k.Key))
            {
                var result = BuscarRContableEnDescripcion(kvp.Value, regla.RContable);
                if (result != null)
                    return new PdfMatch { PageNumber = kvp.Key, Location = result, Caso = "Caso 2", TextoAInsertar = regla.TextoAInsertar };
            }

            foreach (var kvp in pageTextLocations.OrderBy(k => k.Key))
            {
                var result = BuscarFechaYPalabraClave(kvp.Value, regla.Fecha, regla.PalabraClave);
                if (result != null)
                    return new PdfMatch { PageNumber = kvp.Key, Location = result, Caso = "Caso 3", TextoAInsertar = regla.TextoAInsertar };
            }

            return null;
        }

        private static PdfTextLocation? BuscarGuiId(List<PdfTextLocation> words, string guiId)
        {
            if (string.IsNullOrWhiteSpace(guiId))
                return null;

            foreach (var word in words)
            {
                if (ContieneTexto(word.Text, guiId))
                    return word;
            }

            return null;
        }

        private static PdfTextLocation? BuscarRContableEnDescripcion(List<PdfTextLocation> words, string rContable)
        {
            if (string.IsNullOrWhiteSpace(rContable))
                return null;

            var filas = AgruparPorRenglones(words);
            foreach (var fila in filas)
            {
                var textoFila = string.Join(" ", fila.Select(w => w.Text));
                if (!ContieneTexto(textoFila, rContable))
                    continue;

                var columnasDescripcion = fila
                    .Where(w => w.Left >= 80f && w.Left <= 375f)
                    .OrderBy(w => w.Left)
                    .ToList();

                if (columnasDescripcion.Count == 0)
                    continue;

                var match = columnasDescripcion.FirstOrDefault(w => ContieneTexto(w.Text, rContable));
                return match ?? columnasDescripcion.First();
            }

            return null;
        }

        private static PdfTextLocation? BuscarFechaYPalabraClave(List<PdfTextLocation> words, string fecha, string palabraClave)
        {
            if (string.IsNullOrWhiteSpace(fecha) || string.IsNullOrWhiteSpace(palabraClave))
                return null;

            var filas = AgruparPorRenglones(words);
            foreach (var fila in filas)
            {
                var textoFila = string.Join(" ", fila.Select(w => w.Text));
                var tieneFecha = ContieneTexto(textoFila, fecha);
                if (!tieneFecha)
                    continue;

                var cargos = string.Join(" ", fila.Where(w => w.Left >= 315f && w.Left <= 445f).Select(w => w.Text));
                var abonos = string.Join(" ", fila.Where(w => w.Left >= 435f && w.Left <= 530f).Select(w => w.Text));

                if (!ContieneTexto(cargos, palabraClave) && !ContieneTexto(abonos, palabraClave))
                    continue;

                var candidatos = fila
                    .Where(w => w.Left >= 315f && w.Left <= 530f)
                    .OrderBy(w => w.Left)
                    .ToList();

                var match = candidatos.FirstOrDefault(w => ContieneTexto(w.Text, palabraClave));
                return match ?? candidatos.FirstOrDefault() ?? fila.First();
            }

            return null;
        }

        private static List<List<PdfTextLocation>> AgruparPorRenglones(List<PdfTextLocation> words)
        {
            var ordenado = words
                .Where(w => !string.IsNullOrWhiteSpace(w.Text))
                .OrderBy(w => w.Top)
                .ThenBy(w => w.Left)
                .ToList();

            var filas = new List<List<PdfTextLocation>>();
            foreach (var word in ordenado)
            {
                var filaActual = filas.LastOrDefault();
                if (filaActual != null && Math.Abs(filaActual[^1].Top - word.Top) <= 7f)
                {
                    filaActual.Add(word);
                }
                else
                {
                    filas.Add(new List<PdfTextLocation> { word });
                }
            }

            foreach (var fila in filas)
            {
                fila.Sort((a, b) => a.Left.CompareTo(b.Left));
            }

            return filas;
        }

        private static bool ContieneTexto(string? source, string? value)
        {
            if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(value))
                return false;

            var normalizedSource = Normalizar(NormalizarFormatoImporte(source));
            var normalizedValue = Normalizar(NormalizarFormatoImporte(value));
            return !string.IsNullOrEmpty(normalizedValue) && normalizedSource.Contains(normalizedValue, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizarFormatoImporte(string texto)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return texto;

            var valor = texto.Trim();
            if (!valor.Any(char.IsDigit))
                return valor;

            var tieneComa = valor.Contains(',');
            var tienePunto = valor.Contains('.');

            if (tieneComa && tienePunto)
            {
                var ultimoPunto = valor.LastIndexOf('.');
                var ultimaComa = valor.LastIndexOf(',');

                if (ultimaComa > ultimoPunto)
                {
                    return valor
                        .Replace(".", "#TEMP#")
                        .Replace(",", ".")
                        .Replace("#TEMP#", ",");
                }

                return valor;
            }

            if (tieneComa && !tienePunto)
            {
                return valor.Replace(",", ".");
            }

            return valor;
        }

        private static string Normalizar(string value)
        {
            var texto = value.Trim();
            texto = Regex.Replace(texto, @"\s+", " ");
            texto = Regex.Replace(texto, @"[^A-Za-z0-9,./-]", string.Empty);
            return texto.ToUpperInvariant();
        }

        private static void GenerarLogAuditoria(string outputPath, List<ReglaAuditRecord> audios)
        {
            var directory = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
            var logPath = System.IO.Path.Combine(directory, System.IO.Path.GetFileNameWithoutExtension(outputPath) + "_audit.xlsx");

            using var workbook = new XLWorkbook();
            var ws = workbook.Worksheets.Add("Auditoria");
            ws.Cell(1, 1).Value = "TextoAInsertar";
            ws.Cell(1, 2).Value = "Exito";
            ws.Cell(1, 3).Value = "Página";

            for (var index = 0; index < audios.Count; index++)
            {
                var row = index + 2;
                var record = audios[index];
                ws.Cell(row, 1).Value = record.Regla ?? string.Empty;
                ws.Cell(row, 2).Value = record.Exito ? "TRUE" : "FALSE";

                if (record.Exito)
                    ws.Cell(row, 3).Value = record.Pagina;
                else
                    ws.Cell(row, 3).Value = string.Empty;
            }

            ws.Columns().AdjustToContents();
            workbook.SaveAs(logPath);
        }

        private static void EstamparTexto(PdfDocument document, int pageNumber, PdfTextLocation location, string textoAInsertar)
        {
            if (string.IsNullOrWhiteSpace(textoAInsertar))
                return;

            var page = document.GetPage(pageNumber);
            var pageSize = page.GetPageSize();
            var font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
            var canvas = new PdfCanvas(page.NewContentStreamBefore(), page.GetResources(), document);

            var x = location.Left + location.Width + 2f;
            var y = pageSize.GetHeight() - location.Top - location.Height -2f;

            canvas.BeginText()
                .SetFontAndSize(font, 15f)
                .MoveText(x, y)
                .ShowText(textoAInsertar)
                .EndText();

            canvas.Release();
        }
    }

    internal sealed class PdfMatch
    {
        public int PageNumber { get; set; }
        public PdfTextLocation Location { get; set; } = new();
        public string Caso { get; set; } = string.Empty;
        public string TextoAInsertar { get; set; } = string.Empty;
    }

    internal sealed class PdfTextLocation
    {
        public string Text { get; set; } = string.Empty;
        public float Left { get; set; }
        public float Top { get; set; }
        public float Width { get; set; }
        public float Height { get; set; }
    }

    internal sealed class ReglaAuditRecord
    {
        public int NumeroRegla { get; set; }
        public string? Regla { get; set; }
        public string? Caso { get; set; }
        public int Pagina { get; set; }
        public float X { get; set; }
        public float Y { get; set; }
        public string? Resultado { get; set; }
        public bool Exito { get; set; }
        public string? Fecha { get; set; }
        public string? PalabraClave { get; set; }
        public string? GuiId { get; set; }
        public string? RContable { get; set; }
    }

    internal sealed class PdfWordExtractionStrategy : LocationTextExtractionStrategy
    {
        private readonly List<PdfTextLocation> _words = new();
        private readonly float _pageHeight;

        public PdfWordExtractionStrategy(float pageHeight)
        {
            _pageHeight = pageHeight;
        }

        public List<PdfTextLocation> GetLocations() => _words;

        public override void EventOccurred(IEventData data, EventType type)
        {
            base.EventOccurred(data, type);

            if (type != EventType.RENDER_TEXT)
                return;

            var renderInfo = (TextRenderInfo)data;
            var text = renderInfo.GetText();
            if (string.IsNullOrWhiteSpace(text))
                return;

            var start = renderInfo.GetBaseline().GetStartPoint();
            var end = renderInfo.GetAscentLine().GetEndPoint();

            var left = start.Get(0);
            var top = _pageHeight - end.Get(1);
            var width = end.Get(0) - left;
            var height = end.Get(1) - start.Get(1);

            _words.Add(new PdfTextLocation
            {
                Text = text.Trim(),
                Left = left,
                Top = top,
                Width = Math.Max(width, 1f),
                Height = Math.Max(height, 1f)
            });
        }
    }
}
