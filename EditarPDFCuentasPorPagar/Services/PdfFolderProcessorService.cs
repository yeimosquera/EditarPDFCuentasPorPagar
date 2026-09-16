using ClosedXML.Excel;
using EditarPDFCuentasPorPagar.Models;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;

namespace EditarPDFCuentasPorPagar.Services
{
    public class PdfFolderProcessorService : IPdfFolderProcessorService
    {
        private readonly ILogger<PdfFolderProcessorService> _logger;

        public PdfFolderProcessorService(ILogger<PdfFolderProcessorService> logger)
        {
            _logger = logger;
        }

        public Task<ProcessPdfResponse> ProcessAsync(ProcessPdfFolderRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.FilePath))
            {
                throw new ArgumentException("La ruta de origen no puede estar vacía.", nameof(request));
            }

            if (string.IsNullOrWhiteSpace(request.OutputPath))
            {
                throw new ArgumentException("La ruta de salida no puede estar vacía.", nameof(request));
            }

            var response = new ProcessPdfResponse();
            var reglas = request.Reglas ?? new List<ReglaBusquedaDto>();
            var resultadosPorRegla = new List<(string PalabraClave, bool Encontrado)>();

            if (reglas.Count == 0)
            {
                var message = "No se recibieron reglas de búsqueda para procesar.";
                _logger.LogWarning(message);
                response.Success = false;
                response.Message = message;
                return Task.FromResult(response);
            }

            _logger.LogInformation("Iniciando procesamiento de carpeta PDF en {FilePath}.", request.FilePath);

            if (!Directory.Exists(request.FilePath))
            {
                throw new DirectoryNotFoundException($"La carpeta de origen no existe: {request.FilePath}");
            }

            var pdfFiles = Directory
                .GetFiles(request.FilePath, "*.pdf", SearchOption.TopDirectoryOnly)
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();

            if (pdfFiles.Length == 0)
            {
                var message = $"No se encontraron archivos PDF en la carpeta {request.FilePath}.";
                _logger.LogWarning(message);
                response.Success = false;
                response.Message = message;
                return Task.FromResult(response);
            }

            Directory.CreateDirectory(request.OutputPath);

            foreach (var regla in reglas)
            {
                if (string.IsNullOrWhiteSpace(regla.PalabraClave))
                {
                    _logger.LogWarning("Se omite una regla sin palabra clave.");
                    continue;
                }

                var palabraClave = regla.PalabraClave.Trim();
                var matchedFile = pdfFiles.FirstOrDefault(file =>
                    Path.GetFileName(file).Contains(palabraClave, StringComparison.OrdinalIgnoreCase));

                if (matchedFile == null)
                {
                    response.PalabrasNoEncontradas.Add(palabraClave);
                    resultadosPorRegla.Add((palabraClave, false));
                    _logger.LogWarning(
                        "No se encontró la palabra clave {PalabraClave} en ningún PDF de la carpeta {FilePath}.",
                        palabraClave,
                        request.FilePath);
                    continue;
                }

                try
                {
                    var outputFile = Path.Combine(request.OutputPath, Path.GetFileName(matchedFile));
                    StampTextOnPdf(matchedFile, outputFile, regla.TextoAInsertar);

                    response.ArchivosProcesados.Add(outputFile);
                    resultadosPorRegla.Add((palabraClave, true));

                    _logger.LogInformation(
                        "Se encontró el GUID {PalabraClave} en {PdfFile}. Se guardó el archivo en {OutputFile}.",
                        palabraClave,
                        matchedFile,
                        outputFile);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex,
                        "Error al procesar la palabra clave {PalabraClave} en el archivo {PdfFile}.",
                        palabraClave,
                        matchedFile);
                    response.PalabrasNoEncontradas.Add(palabraClave);
                    resultadosPorRegla.Add((palabraClave, false));
                }
            }

            response.Success = response.PalabrasNoEncontradas.Count == 0;
            response.Message = response.Success
                ? "Procesamiento completado correctamente."
                : "Procesamiento completado con advertencias: algunas palabras clave no se encontraron.";

            var logFilePath = GenerateExcelLog(request.OutputPath, resultadosPorRegla);
            response.LogFilePath = logFilePath;

            _logger.LogInformation(
                "Se generó el archivo de trazabilidad Excel en {LogFilePath}.",
                logFilePath);

            _logger.LogInformation(
                "Finaliza el procesamiento de la carpeta PDF. Archivos generados: {GeneratedFiles}.",
                response.ArchivosProcesados.Count);

            return Task.FromResult(response);
        }

        private static void StampTextOnPdf(string sourcePdfPath, string destinationPdfPath, string textToInsert)
        {
            if (string.IsNullOrWhiteSpace(textToInsert))
            {
                return;
            }

            var outputDirectory = Path.GetDirectoryName(destinationPdfPath);
            if (!string.IsNullOrEmpty(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            using var reader = new PdfReader(sourcePdfPath);
            using var writer = new PdfWriter(destinationPdfPath);
            using var pdfDocument = new PdfDocument(reader, writer);

            var font = PdfFontFactory.CreateFont(iText.IO.Font.Constants.StandardFonts.HELVETICA_BOLD);
            var fontSize = 20f;
            var marginSuperior = 30f;
            const float cmToPoints = 28.3464567f;
            var offsetDerecha = 2.5f * cmToPoints;
            var offsetAbajo = 0.8f * cmToPoints;

            for (var pageNumber = 1; pageNumber <= pdfDocument.GetNumberOfPages(); pageNumber++)
            {
                var page = pdfDocument.GetPage(pageNumber);
                var pageSize = page.GetPageSize();
                var textWidth = font.GetWidth(textToInsert, fontSize);
                var xText = (pageSize.GetWidth() / 2) - (textWidth / 2f) + offsetDerecha;
                var yText = pageSize.GetHeight() - marginSuperior - offsetAbajo;

                var canvas = new PdfCanvas(page);
                canvas.BeginText()
                      .SetFontAndSize(font, fontSize)
                      .SetFillColor(iText.Kernel.Colors.ColorConstants.BLACK)
                      .MoveText(xText, yText)
                      .ShowText(textToInsert)
                      .EndText();
                canvas.Release();
            }
        }

        private static string BuildOutputPath(string outputPath, string keyword)
        {
            var normalizedPath = outputPath.Trim();

            if (normalizedPath.Contains("{palabraClave}", StringComparison.OrdinalIgnoreCase))
            {
                return normalizedPath.Replace("{palabraClave}", keyword, StringComparison.OrdinalIgnoreCase);
            }

            var endsWithDirectorySeparator = normalizedPath.EndsWith("\\", StringComparison.Ordinal) || normalizedPath.EndsWith("/", StringComparison.Ordinal);
            var hasExtension = Path.HasExtension(normalizedPath);

            if (endsWithDirectorySeparator || (!hasExtension && !string.IsNullOrEmpty(normalizedPath)))
            {
                var folder = normalizedPath;
                Directory.CreateDirectory(folder);
                return Path.Combine(folder, $"{keyword}.pdf");
            }

            var directory = Path.GetDirectoryName(normalizedPath);
            if (!string.IsNullOrEmpty(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var fileName = $"{Path.GetFileNameWithoutExtension(normalizedPath)}_{keyword}{Path.GetExtension(normalizedPath)}";
            return Path.Combine(directory ?? string.Empty, fileName);
        }

        private static string GenerateExcelLog(string outputPath, IEnumerable<(string PalabraClave, bool Encontrado)> resultados)
        {
            var directory = ResolveOutputDirectory(outputPath);
            Directory.CreateDirectory(directory);

            var fileName = "Log_Procesamiento.xlsx";
            var logFilePath = Path.Combine(directory, fileName);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Log");
                worksheet.Cell(1, 1).Value = "PalabraClave";
                worksheet.Cell(1, 2).Value = "Resultado";

                var rowIndex = 2;
                foreach (var item in resultados)
                {
                    worksheet.Cell(rowIndex, 1).Value = item.PalabraClave;
                    worksheet.Cell(rowIndex, 2).Value = item.Encontrado ? "TRUE" : "FALSE";
                    rowIndex++;
                }

                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(logFilePath);
            }

            return logFilePath;
        }

        private static string ResolveOutputDirectory(string outputPath)
        {
            if (string.IsNullOrWhiteSpace(outputPath))
            {
                return Directory.GetCurrentDirectory();
            }

            var trimmed = outputPath.Trim();
            if (trimmed.EndsWith("\\", StringComparison.Ordinal) || trimmed.EndsWith("/", StringComparison.Ordinal))
            {
                return trimmed;
            }

            if (string.IsNullOrEmpty(Path.GetExtension(trimmed)))
            {
                return trimmed;
            }

            var directory = Path.GetDirectoryName(trimmed);
            return string.IsNullOrEmpty(directory) ? Directory.GetCurrentDirectory() : directory;
        }
    }
}
