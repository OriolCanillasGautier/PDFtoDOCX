using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using UglyToad.PdfPig;
using PDFtoDOCX.Models;
using PDFtoDOCX.Extraction;
using PDFtoDOCX.Layout;
using PDFtoDOCX.Tables;
using PDFtoDOCX.Docx;

namespace PDFtoDOCX
{
    /// <summary>
    /// Main entry point for converting PDF files to DOCX format.
    /// Orchestrates the full conversion pipeline:
    ///   1. Content extraction (PdfPig)
    ///   2. Table detection and reconstruction
    ///   3. Layout analysis (line grouping, paragraphs, columns)
    ///   4. DOCX packaging (manual OOXML generation)
    /// </summary>
    public class PdfToDocxConverter : IPdfToDocxConverter, IDisposable
    {
        private readonly ConversionOptions _options;
        private readonly PdfContentExtractor _extractor;
        private readonly LayoutAnalyzer _layoutAnalyzer;
        private readonly TableDetector _tableDetector;
        private readonly DocxPackager _packager;

        /// <summary>
        /// Creates a new converter with default options.
        /// </summary>
        public PdfToDocxConverter() : this(new ConversionOptions()) { }

        /// <summary>
        /// Creates a new converter with the specified options.
        /// </summary>
        public PdfToDocxConverter(ConversionOptions options)
        {
            _options = options ?? new ConversionOptions();
            _extractor = new PdfContentExtractor(_options);
            _layoutAnalyzer = new LayoutAnalyzer(_options);
            _tableDetector = new TableDetector(_options);
            _packager = new DocxPackager(_options);
        }

        /// <summary>
        /// Converts a PDF file to DOCX format.
        /// </summary>
        /// <param name="pdfPath">Path to the input PDF file.</param>
        /// <param name="docxPath">Path for the output DOCX file.</param>
        public void Convert(string pdfPath, string docxPath)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath));
            if (string.IsNullOrEmpty(docxPath))
                throw new ArgumentNullException(nameof(docxPath));
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF file not found.", pdfPath);

            using var pdfDoc = PdfDocument.Open(pdfPath);
            var docStructure = AnalyzeDocument(pdfDoc);
            _packager.Save(docStructure, docxPath);
        }

        /// <summary>
        /// Asynchronously converts a PDF file to DOCX format.
        /// </summary>
        /// <param name="pdfPath">Path to the input PDF file.</param>
        /// <param name="docxPath">Path for the output DOCX file.</param>
        /// <param name="cancellationToken">Token to cancel the operation.</param>
        /// <param name="progress">Optional progress reporter (0–100).</param>
        public Task ConvertAsync(string pdfPath, string docxPath,
            CancellationToken cancellationToken = default,
            IProgress<int>? progress = null)
        {
            return Task.Run(() =>
            {
                if (string.IsNullOrEmpty(pdfPath))
                    throw new ArgumentNullException(nameof(pdfPath));
                if (string.IsNullOrEmpty(docxPath))
                    throw new ArgumentNullException(nameof(docxPath));
                if (!File.Exists(pdfPath))
                    throw new FileNotFoundException("PDF file not found.", pdfPath);

                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(0);

                using var pdfDoc = PdfDocument.Open(pdfPath);
                cancellationToken.ThrowIfCancellationRequested();
                progress?.Report(20);

                var docStructure = AnalyzeDocument(pdfDoc, cancellationToken, progress, 20, 90);
                cancellationToken.ThrowIfCancellationRequested();

                _packager.Save(docStructure, docxPath);
                progress?.Report(100);
            }, cancellationToken);
        }

        /// <summary>
        /// Converts a PDF file to DOCX and returns the DOCX as a byte array.
        /// </summary>
        /// <param name="pdfPath">Path to the input PDF file.</param>
        /// <returns>DOCX file contents as bytes.</returns>
        public byte[] ConvertToBytes(string pdfPath)
        {
            if (string.IsNullOrEmpty(pdfPath))
                throw new ArgumentNullException(nameof(pdfPath));
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF file not found.", pdfPath);

            using var pdfDoc = PdfDocument.Open(pdfPath);
            var docStructure = AnalyzeDocument(pdfDoc);
            return _packager.Generate(docStructure);
        }

        /// <summary>
        /// Asynchronously converts a PDF file to DOCX and returns the DOCX as a byte array.
        /// </summary>
        public Task<byte[]> ConvertToBytesAsync(string pdfPath,
            CancellationToken cancellationToken = default,
            IProgress<int>? progress = null)
        {
            return Task.Run(() =>
            {
                if (string.IsNullOrEmpty(pdfPath))
                    throw new ArgumentNullException(nameof(pdfPath));
                if (!File.Exists(pdfPath))
                    throw new FileNotFoundException("PDF file not found.", pdfPath);

                cancellationToken.ThrowIfCancellationRequested();
                using var pdfDoc = PdfDocument.Open(pdfPath);
                progress?.Report(20);
                var docStructure = AnalyzeDocument(pdfDoc, cancellationToken, progress, 20, 90);
                cancellationToken.ThrowIfCancellationRequested();
                var result = _packager.Generate(docStructure);
                progress?.Report(100);
                return result;
            }, cancellationToken);
        }

        /// <summary>
        /// Converts a PDF from a stream to DOCX and returns the DOCX as a byte array.
        /// </summary>
        /// <param name="pdfStream">Stream containing the PDF data.</param>
        /// <returns>DOCX file contents as bytes.</returns>
        public byte[] ConvertToBytes(Stream pdfStream)
        {
            if (pdfStream == null)
                throw new ArgumentNullException(nameof(pdfStream));

            using var pdfDoc = PdfDocument.Open(pdfStream);
            var docStructure = AnalyzeDocument(pdfDoc);
            return _packager.Generate(docStructure);
        }

        /// <summary>
        /// Asynchronously converts a PDF from a stream to DOCX.
        /// </summary>
        public Task<byte[]> ConvertToBytesAsync(Stream pdfStream,
            CancellationToken cancellationToken = default,
            IProgress<int>? progress = null)
        {
            return Task.Run(() =>
            {
                if (pdfStream == null)
                    throw new ArgumentNullException(nameof(pdfStream));

                cancellationToken.ThrowIfCancellationRequested();
                using var pdfDoc = PdfDocument.Open(pdfStream);
                progress?.Report(20);
                var docStructure = AnalyzeDocument(pdfDoc, cancellationToken, progress, 20, 90);
                cancellationToken.ThrowIfCancellationRequested();
                var result = _packager.Generate(docStructure);
                progress?.Report(100);
                return result;
            }, cancellationToken);
        }

        /// <summary>
        /// Converts a PDF from a byte array to DOCX and returns the DOCX as a byte array.
        /// </summary>
        /// <param name="pdfBytes">Byte array containing the PDF data.</param>
        /// <returns>DOCX file contents as bytes.</returns>
        public byte[] ConvertToBytes(byte[] pdfBytes)
        {
            if (pdfBytes == null || pdfBytes.Length == 0)
                throw new ArgumentNullException(nameof(pdfBytes));

            using var stream = new MemoryStream(pdfBytes);
            return ConvertToBytes(stream);
        }

        /// <summary>
        /// Asynchronously converts a PDF byte array to DOCX.
        /// </summary>
        public Task<byte[]> ConvertToBytesAsync(byte[] pdfBytes,
            CancellationToken cancellationToken = default,
            IProgress<int>? progress = null)
        {
            using var stream = new MemoryStream(pdfBytes);
            return ConvertToBytesAsync(stream, cancellationToken, progress);
        }

        /// <summary>
        /// Analyzes the full PDF document and produces a DocumentStructure
        /// ready for DOCX generation.
        /// </summary>
        private DocumentStructure AnalyzeDocument(PdfDocument pdfDoc,
            CancellationToken cancellationToken = default,
            IProgress<int>? progress = null,
            int progressStart = 0, int progressEnd = 100)
        {
            var document = new DocumentStructure();

            // Phase 1: Extract raw content from all pages
            var pageContents = _extractor.ExtractAll(pdfDoc);

            int totalPages = pageContents.Count;
            if (_options.EnableDiagnostics)
                System.Console.WriteLine($"[Converter] Extracted {totalPages} page(s)");

            int totalTablesDetected = 0;

            // Phase 2: Analyze each page
            for (int i = 0; i < pageContents.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var pageContent = pageContents[i];
                var pageStructure = AnalyzePage(pageContent, ref totalTablesDetected);
                document.Pages.Add(pageStructure);

                if (progress != null && totalPages > 0)
                {
                    int pct = progressStart + (int)((i + 1.0) / totalPages * (progressEnd - progressStart));
                    progress.Report(pct);
                }
            }

            if (_options.EnableDiagnostics)
                System.Console.WriteLine(
                    $"[Converter] Total tables detected across all pages: {totalTablesDetected}");

            return document;
        }

        /// <summary>
        /// Analyzes a single page's content and produces a PageStructure
        /// with content blocks in reading order.
        /// </summary>
        private PageStructure AnalyzePage(PageContent pageContent, ref int totalTablesDetected)
        {
            var page = new PageStructure
            {
                PageNumber = pageContent.PageNumber,
                Width = pageContent.Width,
                Height = pageContent.Height
            };

            // Step 1: Detect tables
            var tables = new List<DetectedTable>();
            if (_options.DetectTables)
            {
                tables = _tableDetector.DetectTables(pageContent);

                // Phase 4.1: Diagnostic table count logging
                totalTablesDetected += tables.Count;
                if (_options.EnableDiagnostics)
                    System.Console.WriteLine(
                        $"[Converter] Page {pageContent.PageNumber}: {tables.Count} table(s) detected " +
                        $"({pageContent.Lines.Count} line segments, {pageContent.TextElements.Count} text elements)");
            }

            // Step 2: Get text elements NOT inside tables
            var nonTableText = LayoutAnalyzer.ExcludeTableRegions(
                pageContent.TextElements, tables);

            // Step 3: Analyze layout of non-table text
            var paragraphs = _layoutAnalyzer.Analyze(
                nonTableText, pageContent.Width, pageContent.Height);

            // Step 4: Apply hyperlinks to text runs
            if (pageContent.Hyperlinks.Count > 0)
            {
                ApplyHyperlinks(paragraphs, pageContent.Hyperlinks);
            }

            // Step 5: Collect all content blocks with their vertical positions
            var blocks = new List<ContentBlock>();

            // Add paragraphs as blocks
            foreach (var para in paragraphs)
            {
                blocks.Add(new ContentBlock
                {
                    Type = ContentBlockType.Paragraph,
                    Paragraph = para,
                    Bounds = para.Bounds
                });
            }

            // Add tables as blocks
            foreach (var table in tables)
            {
                blocks.Add(new ContentBlock
                {
                    Type = ContentBlockType.Table,
                    Table = table,
                    Bounds = table.Bounds
                });
            }

            // Add images as blocks (only images not inside tables)
            if (_options.ExtractImages)
            {
                foreach (var image in pageContent.Images)
                {
                    // Skip images that are inside table bounds
                    bool insideTable = tables.Any(t => t.Bounds.Contains(image.Bounds));
                    if (!insideTable)
                    {
                        blocks.Add(new ContentBlock
                        {
                            Type = ContentBlockType.Image,
                            Image = image,
                            Bounds = image.Bounds
                        });
                    }
                }
            }

            // Step 6: Sort blocks into reading order (top-to-bottom)
            blocks.Sort((a, b) =>
            {
                int cmp = a.Bounds.Top.CompareTo(b.Bounds.Top);
                if (cmp != 0) return cmp;
                return a.Bounds.Left.CompareTo(b.Bounds.Left);
            });

            page.Blocks = blocks;
            return page;
        }

        /// <summary>
        /// Applies hyperlink annotations to the appropriate text runs.
        /// If a text run's position overlaps with a hyperlink annotation,
        /// the hyperlink URI is applied to that run.
        /// </summary>
        private void ApplyHyperlinks(List<TextParagraph> paragraphs, List<HyperlinkInfo> hyperlinks)
        {
            foreach (var para in paragraphs)
            {
                foreach (var line in para.Lines)
                {
                    foreach (var run in line.Runs)
                    {
                        // Check if any hyperlink covers this line's position
                        foreach (var link in hyperlinks)
                        {
                            if (line.Bounds.Intersects(link.Bounds))
                            {
                                run.HyperlinkUri = link.Uri;
                                break;
                            }
                        }
                    }
                }
            }
        }

        public void Dispose()
        {
            // No unmanaged resources, but implementing IDisposable for good practice
        }
    }
}
