using System;
using System.Diagnostics;
using System.IO;
using PDFtoDOCX;

namespace PDFtoDOCX.Console
{
    /// <summary>
    /// Command-line interface for the PDF-to-DOCX converter.
    /// Usage: PDFtoDOCX.Console [options] &lt;input.pdf&gt; [output.docx]
    /// </summary>
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0 || args[0] == "--help" || args[0] == "-h")
            {
                PrintUsage();
                return 0;
            }

            string inputPath = null;
            string outputPath = null;
            var options = new ConversionOptions();

            // Parse arguments
            for (int i = 0; i < args.Length; i++)
            {
                switch (args[i])
                {
                    case "--no-images":
                        options.ExtractImages = false;
                        break;
                    case "--no-tables":
                        options.DetectTables = false;
                        break;
                    case "--no-hyperlinks":
                        options.DetectHyperlinks = false;
                        break;
                    case "--start-page":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int sp))
                            options.StartPage = sp;
                        break;
                    case "--end-page":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int ep))
                            options.EndPage = ep;
                        break;
                    case "--max-pages":
                        if (i + 1 < args.Length && int.TryParse(args[++i], out int mp))
                            options.MaxPages = mp;
                        break;
                    case "--line-tolerance":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out double lt))
                            options.LineGroupingTolerance = lt;
                        break;
                    case "--para-gap":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out double pg))
                            options.ParagraphGapMultiplier = pg;
                        break;
                    case "--column-gap":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out double cg))
                            options.MinColumnGap = cg;
                        break;
                    case "--diagnostics":
                        options.EnableDiagnostics = true;
                        break;
                    case "--ocr":
                        options.EnableOcr = true;
                        break;
                    case "--line-spacing":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out double ls))
                            options.LineSpacingMultiplier = ls;
                        break;
                    case "--para-spacing-after":
                        if (i + 1 < args.Length && double.TryParse(args[++i], out double psa))
                            options.ParagraphSpacingAfter = psa;
                        break;
                    default:
                        if (args[i].StartsWith("-"))
                        {
                            System.Console.Error.WriteLine($"Unknown option: {args[i]}");
                            return 1;
                        }
                        if (inputPath == null)
                            inputPath = args[i];
                        else if (outputPath == null)
                            outputPath = args[i];
                        break;
                }
            }

            if (string.IsNullOrEmpty(inputPath))
            {
                System.Console.Error.WriteLine("Error: No input PDF file specified.");
                PrintUsage();
                return 1;
            }

            if (!File.Exists(inputPath))
            {
                System.Console.Error.WriteLine($"Error: File not found: {inputPath}");
                return 1;
            }

            // Default output path: same as input with .docx extension
            if (string.IsNullOrEmpty(outputPath))
            {
                outputPath = Path.ChangeExtension(inputPath, ".docx");
            }

            try
            {
                System.Console.WriteLine($"Converting: {inputPath}");
                System.Console.WriteLine($"Output:     {outputPath}");
                System.Console.WriteLine();

                var sw = Stopwatch.StartNew();

                using var converter = new PdfToDocxConverter(options);
                converter.Convert(inputPath, outputPath);

                sw.Stop();

                var fileInfo = new FileInfo(outputPath);
                System.Console.WriteLine($"Conversion completed successfully.");
                System.Console.WriteLine($"  Time:   {sw.ElapsedMilliseconds} ms");
                System.Console.WriteLine($"  Output: {fileInfo.Length:N0} bytes");

                return 0;
            }
            catch (Exception ex)
            {
                System.Console.Error.WriteLine($"Error during conversion: {ex.Message}");
                System.Console.Error.WriteLine(ex.StackTrace);
                return 2;
            }
        }

        static void PrintUsage()
        {
            System.Console.WriteLine(@"
PDFtoDOCX - High-Fidelity PDF to DOCX Converter
=================================================

Usage:
  PDFtoDOCX.Console [options] <input.pdf> [output.docx]

Arguments:
  <input.pdf>          Path to the input PDF file (required)
  [output.docx]        Path for the output DOCX file (default: same name as input)

Options:
  --no-images          Do not extract or embed images
  --no-tables          Do not detect or reconstruct tables
  --no-hyperlinks      Do not detect hyperlinks
  --start-page N       Start converting from page N (1-based, default: 1)
  --end-page N         Stop converting at page N (inclusive, default: last page)
  --max-pages N        Maximum number of pages to convert (default: all)
  --line-tolerance N   Y-coordinate tolerance for line grouping in points (default: 3.0)
  --para-gap N         Paragraph gap multiplier (default: 1.3)
  --column-gap N       Minimum column gap in points (default: 20.0)
  --line-spacing N     Line spacing multiplier relative to font size (default: 1.15)
  --para-spacing-after N  Space after each paragraph in points (default: 6.0)
  --diagnostics        Print detailed extraction and detection info during conversion
  --ocr                Enable OCR fallback for pages with no extractable text
  -h, --help           Show this help message

Examples:
  PDFtoDOCX.Console document.pdf
  PDFtoDOCX.Console document.pdf output.docx
  PDFtoDOCX.Console --start-page 2 --end-page 5 document.pdf
  PDFtoDOCX.Console --no-tables --no-images document.pdf simple.docx
");
        }
    }
}
