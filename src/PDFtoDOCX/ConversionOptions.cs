namespace PDFtoDOCX
{
    /// <summary>
    /// Configuration options for the PDF-to-DOCX conversion.
    /// </summary>
    public class ConversionOptions
    {
        /// <summary>
        /// Tolerance (in PDF points) for grouping text elements into the same line.
        /// Elements whose vertical midpoints differ by less than this are considered on the same line.
        /// Default: 3.0 points.
        /// </summary>
        public double LineGroupingTolerance { get; set; } = 3.0;

        /// <summary>
        /// Multiplier of average line height used to detect paragraph breaks.
        /// If the vertical gap between consecutive lines exceeds this multiple
        /// of the average line height, a new paragraph starts.
        /// Default: 1.3.
        /// </summary>
        public double ParagraphGapMultiplier { get; set; } = 1.3;

        /// <summary>
        /// Minimum horizontal gap (in PDF points) between text columns.
        /// If adjacent text clusters are separated by more than this gap, they
        /// are treated as separate columns.
        /// Default: 20.0 points.
        /// </summary>
        public double MinColumnGap { get; set; } = 20.0;

        /// <summary>
        /// Minimum length (in PDF points) for a line segment to be considered
        /// as a potential table border.
        /// Default: 10.0 points.
        /// </summary>
        public double MinTableLineLength { get; set; } = 10.0;

        /// <summary>
        /// Tolerance (in PDF points) for snapping line endpoints to grid intersections
        /// during table detection.
        /// Default: 5.0 points.
        /// </summary>
        public double TableGridSnapTolerance { get; set; } = 5.0;

        /// <summary>
        /// Minimum number of rows for a grid to be recognized as a table.
        /// Default: 2.
        /// </summary>
        public int MinTableRows { get; set; } = 2;

        /// <summary>
        /// Minimum number of columns for a grid to be recognized as a table.
        /// Default: 2.
        /// </summary>
        public int MinTableCols { get; set; } = 2;

        /// <summary>
        /// Default page margins in PDF points if the source PDF doesn't specify them.
        /// Default: 72 (1 inch).
        /// </summary>
        public double DefaultMargin { get; set; } = 72.0;

        /// <summary>
        /// Whether to extract and embed images from the PDF into the DOCX.
        /// Default: true.
        /// </summary>
        public bool ExtractImages { get; set; } = true;

        /// <summary>
        /// Whether to detect and reconstruct tables.
        /// Default: true.
        /// </summary>
        public bool DetectTables { get; set; } = true;

        /// <summary>
        /// Whether to detect hyperlinks from PDF annotations.
        /// Default: true.
        /// </summary>
        public bool DetectHyperlinks { get; set; } = true;

        /// <summary>
        /// Maximum number of pages to convert. 0 = all pages.
        /// Default: 0.
        /// </summary>
        public int MaxPages { get; set; } = 0;

        /// <summary>
        /// Starting page number (1-based). Default: 1.
        /// </summary>
        public int StartPage { get; set; } = 1;

        /// <summary>
        /// Ending page number (1-based, inclusive). 0 = last page.
        /// Default: 0.
        /// </summary>
        public int EndPage { get; set; } = 0;

        // ── Phase 2: Inline spacing ──────────────────────────────────────────

        /// <summary>
        /// Multiplier applied to font size to calculate line height in the DOCX output.
        /// A value of 1.15 produces single-spaced output with a slight openness.
        /// Default: 1.15.
        /// </summary>
        public double LineSpacingMultiplier { get; set; } = 1.15;

        /// <summary>
        /// Space added after each paragraph, in PDF points.
        /// Maps to OOXML &lt;w:spacing w:after="…"/&gt;.
        /// Default: 6.0 points.
        /// </summary>
        public double ParagraphSpacingAfter { get; set; } = 6.0;

        // ── Phase 3: Text extraction ─────────────────────────────────────────

        /// <summary>
        /// Gap multiplier used when grouping individual letters into words during
        /// fallback letter-grouping extraction.
        /// A gap larger than <c>avgCharWidth × GapMultiplier</c> starts a new word.
        /// Default: 0.5.
        /// </summary>
        public double LetterGroupingGapMultiplier { get; set; } = 0.5;

        /// <summary>
        /// When true, enables the OCR fallback for pages where no text operators
        /// can be detected. Requires the optional Tesseract integration to be present.
        /// Default: false.
        /// </summary>
        public bool EnableOcr { get; set; } = false;

        // ── Phase 7: Diagnostics ─────────────────────────────────────────────

        /// <summary>
        /// When true, writes diagnostic information (extraction counts, table
        /// confidence scores, image failures) to the console during conversion.
        /// Default: false.
        /// </summary>
        public bool EnableDiagnostics { get; set; } = false;
    }
}
