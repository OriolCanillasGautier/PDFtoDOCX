using System;

namespace PDFtoDOCX.Utils
{
    /// <summary>
    /// Converts between PDF coordinate units and DOCX/OOXML units.
    /// </summary>
    public static class UnitConverter
    {
        // PDF uses points (1 pt = 1/72 inch)
        // DOCX uses multiple unit systems:

        /// <summary>1 inch = 72 PDF points.</summary>
        public const double PointsPerInch = 72.0;

        /// <summary>1 inch = 914400 EMU (English Metric Units). Used for DrawingML (images, shapes).</summary>
        public const long EmuPerInch = 914400;

        /// <summary>1 inch = 1440 twips. Used for page layout (margins, page size, table widths).</summary>
        public const double TwipsPerInch = 1440.0;

        /// <summary>1 pt = 20 twips.</summary>
        public const double TwipsPerPoint = 20.0;

        /// <summary>DOCX font sizes use half-points: w:sz value of 24 = 12pt font.</summary>
        public const double HalfPointsPerPoint = 2.0;

        /// <summary>1 PDF point = 12700 EMU.</summary>
        public const long EmuPerPoint = 12700;

        /// <summary>Standard US Letter width in points (8.5 inches).</summary>
        public const double LetterWidthPt = 612.0;

        /// <summary>Standard US Letter height in points (11 inches).</summary>
        public const double LetterHeightPt = 792.0;

        /// <summary>Standard A4 width in points (210mm).</summary>
        public const double A4WidthPt = 595.28;

        /// <summary>Standard A4 height in points (297mm).</summary>
        public const double A4HeightPt = 841.89;

        // --- Conversion methods ---

        /// <summary>Convert PDF points to OOXML twips.</summary>
        public static int PointsToTwips(double points) => (int)Math.Round(points * TwipsPerPoint);

        /// <summary>Convert PDF points to EMU (for DrawingML images/shapes).</summary>
        public static long PointsToEmu(double points) => (long)Math.Round(points * EmuPerPoint);

        /// <summary>Convert PDF font size (in points) to OOXML half-points (w:sz value).</summary>
        public static int FontSizeToHalfPoints(double pointSize) => (int)Math.Round(pointSize * HalfPointsPerPoint);

        /// <summary>Convert twips to PDF points.</summary>
        public static double TwipsToPoints(int twips) => twips / TwipsPerPoint;

        /// <summary>Convert EMU to PDF points.</summary>
        public static double EmuToPoints(long emu) => (double)emu / EmuPerPoint;

        /// <summary>Convert pixels to EMU at a given DPI (default 96 DPI).</summary>
        public static long PixelsToEmu(double pixels, double dpi = 96.0) => (long)Math.Round(pixels / dpi * EmuPerInch);

        /// <summary>Convert PDF points to OOXML border width (in 1/8 of a point).</summary>
        public static int PointsToBorderWidth(double points) => Math.Max(1, (int)Math.Round(points * 8));

        /// <summary>
        /// Convert a PDF RGB color (0.0-1.0 per channel) to a hex string like "FF0000".
        /// </summary>
        public static string RgbToHex(double r, double g, double b)
        {
            int ri = Math.Clamp((int)Math.Round(r * 255), 0, 255);
            int gi = Math.Clamp((int)Math.Round(g * 255), 0, 255);
            int bi = Math.Clamp((int)Math.Round(b * 255), 0, 255);
            return $"{ri:X2}{gi:X2}{bi:X2}";
        }

        /// <summary>
        /// Convert a gray value (0.0 = black, 1.0 = white) to hex string.
        /// </summary>
        public static string GrayToHex(double gray)
        {
            return RgbToHex(gray, gray, gray);
        }
    }
}
