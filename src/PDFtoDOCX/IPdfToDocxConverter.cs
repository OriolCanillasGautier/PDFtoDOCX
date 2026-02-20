using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace PDFtoDOCX
{
    /// <summary>
    /// Contract for converting PDF documents to DOCX format.
    /// Implement this interface to swap in alternative back-ends via dependency injection.
    /// </summary>
    public interface IPdfToDocxConverter
    {
        /// <summary>
        /// Converts a PDF file at <paramref name="pdfPath"/> and writes the resulting
        /// DOCX to <paramref name="docxPath"/>.
        /// </summary>
        void Convert(string pdfPath, string docxPath);

        /// <summary>
        /// Converts a PDF file and returns the DOCX content as a byte array.
        /// </summary>
        byte[] ConvertToBytes(string pdfPath);

        /// <summary>
        /// Converts a PDF stream and returns the DOCX content as a byte array.
        /// </summary>
        byte[] ConvertToBytes(Stream pdfStream);

        /// <summary>
        /// Converts a PDF byte array and returns the DOCX content as a byte array.
        /// </summary>
        byte[] ConvertToBytes(byte[] pdfBytes);

        /// <summary>
        /// Asynchronously converts a PDF file and writes the DOCX to <paramref name="docxPath"/>.
        /// </summary>
        Task ConvertAsync(string pdfPath, string docxPath,
            CancellationToken cancellationToken = default,
            IProgress<int>? progress = null);

        /// <summary>
        /// Asynchronously converts a PDF file and returns the DOCX as a byte array.
        /// </summary>
        Task<byte[]> ConvertToBytesAsync(string pdfPath,
            CancellationToken cancellationToken = default,
            IProgress<int>? progress = null);

        /// <summary>
        /// Asynchronously converts a PDF stream and returns the DOCX as a byte array.
        /// </summary>
        Task<byte[]> ConvertToBytesAsync(Stream pdfStream,
            CancellationToken cancellationToken = default,
            IProgress<int>? progress = null);
    }
}
