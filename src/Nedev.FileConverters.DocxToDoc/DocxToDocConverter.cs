using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Nedev.FileConverters.DocxToDoc
{
    /// <summary>
    /// Provides high-performance conversion from OpenXML (.docx) into MS-DOC legacy binary (.doc) format.
    /// Does not rely on any third-party libraries.
    /// </summary>
    public class DocxToDocConverter
    {
        private readonly ILogger _logger;

        static DocxToDocConverter()
        {
            // attempt to register code pages encoding provider if available (some targets may not include the
            // type by default, e.g. netstandard2.1). Use reflection so compilation succeeds across all TFM.
            var providerType = Type.GetType("System.Text.CodePagesEncodingProvider, System.Text.Encoding.CodePages");
            if (providerType != null)
            {
                var instanceProp = providerType.GetProperty("Instance", System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
                var instance = instanceProp?.GetValue(null);
                if (instance is EncodingProvider ep)
                {
                    Encoding.RegisterProvider(ep);
                }
            }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocxToDocConverter"/> class.
        /// </summary>
        public DocxToDocConverter() : this(NullLogger.Instance)
        {
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DocxToDocConverter"/> class with a logger.
        /// </summary>
        /// <param name="logger">The logger to use for recording operations and errors.</param>
        public DocxToDocConverter(ILogger logger)
        {
            _logger = logger ?? NullLogger.Instance;
        }

        /// <summary>
        /// Converts a .docx file to a .doc file.
        /// </summary>
        /// <param name="docxPath">The path to the source .docx file.</param>
        /// <param name="docPath">The path to the destination .doc file.</param>
        /// <exception cref="ArgumentNullException">Thrown when docxPath or docPath is null or empty.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the source file does not exist.</exception>
        /// <exception cref="ConversionException">Thrown when an error occurs during conversion.</exception>
        public void Convert(string docxPath, string docPath)
        {
            if (string.IsNullOrWhiteSpace(docxPath))
                throw new ArgumentNullException(nameof(docxPath));
            if (string.IsNullOrWhiteSpace(docPath))
                throw new ArgumentNullException(nameof(docPath));

            _logger.LogInfo($"Starting conversion: '{docxPath}' -> '{docPath}'");

            try
            {
                ValidateInputFile(docxPath);

                using var inputStream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var outputStream = new FileStream(docPath, FileMode.Create, FileAccess.Write, FileShare.None);
                
                Convert(inputStream, outputStream);
                
                _logger.LogInfo($"Conversion completed successfully: '{docPath}'");
            }
            catch (Exception ex) when (!(ex is ArgumentNullException || ex is FileNotFoundException || ex is ConversionException))
            {
                _logger.LogError($"Conversion failed: '{docxPath}' -> '{docPath}'", ex);
                throw new ConversionException(
                    $"Failed to convert '{docxPath}' to '{docPath}'.",
                    docxPath,
                    docPath,
                    ConversionStage.Unknown,
                    ex);
            }
        }

        /// <summary>
        /// Converts a .docx stream to a .doc stream.
        /// </summary>
        /// <param name="docxStream">A stream containing the OpenXML document. Must support Read.</param>
        /// <param name="docStream">A stream where the .doc binary will be written. Must support Write.</param>
        /// <exception cref="ArgumentNullException">Thrown when docxStream or docStream is null.</exception>
        /// <exception cref="ArgumentException">Thrown when streams do not support required operations.</exception>
        /// <exception cref="ConversionException">Thrown when an error occurs during conversion.</exception>
        public void Convert(Stream docxStream, Stream docStream)
        {
            if (docxStream == null) throw new ArgumentNullException(nameof(docxStream));
            if (docStream == null) throw new ArgumentNullException(nameof(docStream));

            if (!docxStream.CanRead)
                throw new ArgumentException("Input stream must support reading.", nameof(docxStream));
            if (!docStream.CanWrite)
                throw new ArgumentException("Output stream must support writing.", nameof(docStream));

            _logger.LogDebug("Starting stream conversion");

            using var monitor = new PerformanceMonitor(_logger);
            monitor.Start();

            try
            {
                // Log initial memory usage
                MemoryMonitor.LogMemoryUsage(_logger, "Before conversion");

                // Create a DocxReader to parse the document contents
                _logger.LogDebug("Initializing DOCX reader");
                using var reader = new Format.DocxReader(docxStream);

                // Extract necessary layout/styles/content out of OpenXML and map to MS-DOC
                _logger.LogDebug("Reading document content");
                var documentModel = reader.ReadDocument();
                    int imageCount = CountImages(documentModel);

                monitor.RecordParagraphs(documentModel.Paragraphs.Count);
                monitor.RecordTables(documentModel.Content.Count(c => c is Model.TableModel));
                    monitor.RecordImages(imageCount);

                _logger.LogInfo($"Document parsed: {documentModel.Paragraphs.Count} paragraphs, {documentModel.Styles.Count} styles");

                // Provide data blocks for the MS-DOC writer
                _logger.LogDebug("Writing DOC format");
                var writer = new Format.DocWriter();
                writer.WriteDocBlocks(documentModel, docStream);

                // Record bytes written
                if (docStream.CanSeek)
                {
                    monitor.RecordBytesWritten(docStream.Position);
                }

                // Log final memory usage
                MemoryMonitor.LogMemoryUsage(_logger, "After conversion");

                monitor.Stop();
                monitor.LogSummary();

                _logger.LogDebug("Stream conversion completed");
            }
            catch (FileNotFoundException ex)
            {
                _logger.LogError("Required file not found in DOCX archive", ex);
                throw new ConversionException(
                    "The DOCX file is missing required components.",
                    ConversionStage.Reading,
                    ex);
            }
            catch (InvalidDataException ex)
            {
                _logger.LogError("Invalid DOCX format", ex);
                throw new ConversionException(
                    "The input file is not a valid DOCX file or is corrupted.",
                    ConversionStage.Parsing,
                    ex);
            }
            catch (Exception ex) when (!(ex is ArgumentNullException || ex is ArgumentException || ex is ConversionException))
            {
                _logger.LogError("Unexpected error during conversion", ex);
                throw new ConversionException(
                    "An unexpected error occurred during conversion.",
                    ConversionStage.Unknown,
                    ex);
            }
        }

        /// <summary>
        /// Asynchronously converts a .docx file to a .doc file.
        /// </summary>
        /// <param name="docxPath">The path to the source .docx file.</param>
        /// <param name="docPath">The path to the destination .doc file.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <returns>A task representing the asynchronous conversion operation.</returns>
        /// <exception cref="ArgumentNullException">Thrown when docxPath or docPath is null or empty.</exception>
        /// <exception cref="FileNotFoundException">Thrown when the source file does not exist.</exception>
        /// <exception cref="ConversionException">Thrown when an error occurs during conversion.</exception>
        /// <exception cref="OperationCanceledException">Thrown when the operation is cancelled.</exception>
        public async Task ConvertAsync(string docxPath, string docPath, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(docxPath))
                throw new ArgumentNullException(nameof(docxPath));
            if (string.IsNullOrWhiteSpace(docPath))
                throw new ArgumentNullException(nameof(docPath));

            _logger.LogInfo($"Starting async conversion: '{docxPath}' -> '{docPath}'");

            try
            {
                ValidateInputFile(docxPath);

                using var inputStream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                using var outputStream = new FileStream(docPath, FileMode.Create, FileAccess.Write, FileShare.None);
                
                await ConvertAsync(inputStream, outputStream, cancellationToken).ConfigureAwait(false);
                
                _logger.LogInfo($"Async conversion completed successfully: '{docPath}'");
            }
            catch (OperationCanceledException)
            {
                _logger.LogWarning($"Conversion cancelled: '{docxPath}' -> '{docPath}'");
                throw;
            }
            catch (Exception ex) when (!(ex is ArgumentNullException || ex is FileNotFoundException || ex is ConversionException || ex is OperationCanceledException))
            {
                _logger.LogError($"Async conversion failed: '{docxPath}' -> '{docPath}'", ex);
                throw new ConversionException(
                    $"Failed to convert '{docxPath}' to '{docPath}'.",
                    docxPath,
                    docPath,
                    ConversionStage.Unknown,
                    ex);
            }
        }

        /// <summary>
        /// Asynchronously converts a .docx stream to a .doc stream.
        /// </summary>
        /// <param name="docxStream">A stream containing the OpenXML document. Must support Read.</param>
        /// <param name="docStream">A stream where the .doc binary will be written. Must support Write.</param>
        /// <param name="cancellationToken">Cancellation token to cancel the operation.</param>
        /// <returns>A task representing the asynchronous conversion operation.</returns>
        /// <exception cref="ArgumentNullException">Thrown when docxStream or docStream is null.</exception>
        /// <exception cref="ArgumentException">Thrown when streams do not support required operations.</exception>
        /// <exception cref="ConversionException">Thrown when an error occurs during conversion.</exception>
        /// <exception cref="OperationCanceledException">Thrown when the operation is cancelled.</exception>
        public async Task ConvertAsync(Stream docxStream, Stream docStream, CancellationToken cancellationToken = default)
        {
            if (docxStream == null) throw new ArgumentNullException(nameof(docxStream));
            if (docStream == null) throw new ArgumentNullException(nameof(docStream));

            if (!docxStream.CanRead)
                throw new ArgumentException("Input stream must support reading.", nameof(docxStream));
            if (!docStream.CanWrite)
                throw new ArgumentException("Output stream must support writing.", nameof(docStream));

            cancellationToken.ThrowIfCancellationRequested();

            _logger.LogDebug("Starting async stream conversion");

            try
            {
                // Create a DocxReader to parse the document contents
                _logger.LogDebug("Initializing DOCX reader");
                using var reader = new Format.DocxReader(docxStream);
                
                cancellationToken.ThrowIfCancellationRequested();

                // Extract necessary layout/styles/content out of OpenXML and map to MS-DOC
                _logger.LogDebug("Reading document content");
                var documentModel = await Task.Run(() => reader.ReadDocument(), cancellationToken).ConfigureAwait(false);
                
                cancellationToken.ThrowIfCancellationRequested();

                _logger.LogInfo($"Document parsed: {documentModel.Paragraphs.Count} paragraphs, {documentModel.Styles.Count} styles");
                
                // Provide data blocks for the MS-DOC writer
                _logger.LogDebug("Writing DOC format");
                var writer = new Format.DocWriter();
                await Task.Run(() => writer.WriteDocBlocks(documentModel, docStream), cancellationToken).ConfigureAwait(false);
                
                _logger.LogDebug("Async stream conversion completed");
            }
            catch (OperationCanceledException)
            {
                _logger.LogWarning("Async conversion cancelled");
                throw;
            }
            catch (FileNotFoundException ex)
            {
                _logger.LogError("Required file not found in DOCX archive", ex);
                throw new ConversionException(
                    "The DOCX file is missing required components.",
                    ConversionStage.Reading,
                    ex);
            }
            catch (InvalidDataException ex)
            {
                _logger.LogError("Invalid DOCX format", ex);
                throw new ConversionException(
                    "The input file is not a valid DOCX file or is corrupted.",
                    ConversionStage.Parsing,
                    ex);
            }
            catch (Exception ex) when (!(ex is ArgumentNullException || ex is ArgumentException || ex is ConversionException || ex is OperationCanceledException))
            {
                _logger.LogError("Unexpected error during async conversion", ex);
                throw new ConversionException(
                    "An unexpected error occurred during conversion.",
                    ConversionStage.Unknown,
                    ex);
            }
        }

        /// <summary>
        /// Validates that the input file exists and is accessible.
        /// </summary>
        private void ValidateInputFile(string docxPath)
        {
            _logger.LogDebug($"Validating input file: '{docxPath}'");

            if (!File.Exists(docxPath))
            {
                _logger.LogError($"Input file not found: '{docxPath}'");
                throw new FileNotFoundException($"The input file '{docxPath}' does not exist.", docxPath);
            }

            var fileInfo = new FileInfo(docxPath);
            if (fileInfo.Length == 0)
            {
                _logger.LogError($"Input file is empty: '{docxPath}'");
                throw new ConversionException(
                    $"The input file '{docxPath}' is empty.",
                    docxPath,
                    null,
                    ConversionStage.Validation);
            }

            // Check file extension (optional validation)
            if (!string.Equals(Path.GetExtension(docxPath), ".docx", StringComparison.OrdinalIgnoreCase))
            {
                _logger.LogWarning($"Input file does not have .docx extension: '{docxPath}'");
            }

            _logger.LogDebug($"Input file validated: {fileInfo.Length} bytes");
        }

        private static int CountImages(Model.DocumentModel documentModel)
        {
            int imageCount = 0;

            foreach (var item in documentModel.Content)
            {
                if (item is Model.ParagraphModel paragraph)
                {
                    imageCount += CountImages(paragraph);
                }
                else if (item is Model.TableModel table)
                {
                    foreach (var row in table.Rows)
                    {
                        foreach (var cell in row.Cells)
                        {
                            foreach (var cellParagraph in cell.Paragraphs)
                            {
                                imageCount += CountImages(cellParagraph);
                            }
                        }
                    }
                }
            }

            return imageCount;

            static int CountImages(Model.ParagraphModel paragraph)
            {
                return paragraph.Runs.Count(run => run.Image != null);
            }
        }
    }
}
