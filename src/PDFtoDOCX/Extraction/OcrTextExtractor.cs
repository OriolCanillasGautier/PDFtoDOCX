using System.Collections.Generic;
using UglyToad.PdfPig.Content;
using PDFtoDOCX.Models;

namespace PDFtoDOCX.Extraction
{
    /// <summary>
    /// Interface for text extraction back-ends, allowing the primary word-based
    /// extractor and OCR fallback to be swapped or combined.
    /// </summary>
    public interface ITextExtractor
    {
        /// <summary>
        /// Extracts text elements from the given PDF page.
        /// </summary>
        /// <param name="page">The PdfPig page to process.</param>
        /// <param name="pageHeight">Page height in PDF points (used for Y-axis inversion).</param>
        /// <returns>List of extracted text elements in top-left-origin coordinates.</returns>
        List<TextElement> ExtractText(Page page, double pageHeight);
    }

    /// <summary>
    /// Stub OCR text extractor that can be subclassed or replaced when a real
    /// OCR library (e.g. Tesseract.NET) is available.
    /// <para>
    /// To activate OCR:
    /// <list type="number">
    ///   <item>Add the <c>Tesseract</c> NuGet package to the project.</item>
    ///   <item>Subclass <see cref="OcrTextExtractor"/> and override <see cref="ExtractText"/>.</item>
    ///   <item>Set <see cref="ConversionOptions.EnableOcr"/> to <c>true</c>.</item>
    ///   <item>Override <see cref="PdfContentExtractor.OcrExtractText"/> to call your extractor.</item>
    /// </list>
    /// </para>
    /// </summary>
    public class OcrTextExtractor : ITextExtractor
    {
        private readonly ConversionOptions _options;

        /// <summary>
        /// Creates an OCR extractor with the specified options.
        /// </summary>
        public OcrTextExtractor(ConversionOptions options)
        {
            _options = options ?? new ConversionOptions();
        }

        /// <summary>
        /// Rasterizes the page and runs OCR to obtain text elements.
        /// <para>
        /// This base implementation returns an empty list.
        /// Override in a subclass to provide a real OCR back-end.
        /// </para>
        /// Example override sketch (requires Tesseract NuGet + SkiaSharp for rasterization):
        /// <code>
        /// public override List&lt;TextElement&gt; ExtractText(Page page, double pageHeight)
        /// {
        ///     using var bitmap = RasterizePage(page, dpi: 150);
        ///     using var engine = new TesseractEngine(@"./tessdata", "eng");
        ///     using var img = Pix.LoadFromMemory(bitmap);
        ///     using var result = engine.Process(img);
        ///     return ParseHOCR(result.GetHOCRText(0), pageHeight);
        /// }
        /// </code>
        /// </summary>
        public virtual List<TextElement> ExtractText(Page page, double pageHeight)
        {
            // Stub: no OCR back-end wired up.
            // Return empty list so the pipeline continues gracefully.
            return new List<TextElement>();
        }
    }
}
