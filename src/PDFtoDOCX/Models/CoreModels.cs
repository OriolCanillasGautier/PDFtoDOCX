using System;
using System.Collections.Generic;

namespace PDFtoDOCX.Models
{
    /// <summary>
    /// A rectangle with top-left origin (Y increases downward).
    /// All coordinates are in PDF points (1 pt = 1/72 inch).
    /// </summary>
    public class Rect
    {
        public double Left { get; set; }
        public double Top { get; set; }
        public double Right { get; set; }
        public double Bottom { get; set; }

        public Rect() { }

        public Rect(double left, double top, double right, double bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public double Width => Right - Left;
        public double Height => Bottom - Top;
        public double MidX => (Left + Right) / 2.0;
        public double MidY => (Top + Bottom) / 2.0;

        /// <summary>
        /// Returns true if this rectangle intersects with another rectangle.
        /// </summary>
        public bool Intersects(Rect other)
        {
            return Left < other.Right && Right > other.Left &&
                   Top < other.Bottom && Bottom > other.Top;
        }

        /// <summary>
        /// Returns true if this rectangle fully contains another rectangle.
        /// </summary>
        public bool Contains(Rect other)
        {
            return Left <= other.Left && Right >= other.Right &&
                   Top <= other.Top && Bottom >= other.Bottom;
        }

        /// <summary>
        /// Returns true if a point (x, y) is inside this rectangle.
        /// </summary>
        public bool ContainsPoint(double x, double y)
        {
            return x >= Left && x <= Right && y >= Top && y <= Bottom;
        }

        public override string ToString() => $"Rect({Left:F1}, {Top:F1}, {Right:F1}, {Bottom:F1})";
    }

    /// <summary>
    /// Represents a single text element (word-level) extracted from a PDF page.
    /// Coordinates use top-left origin.
    /// </summary>
    public class TextElement
    {
        public string Text { get; set; } = string.Empty;
        public Rect Bounds { get; set; } = new Rect();
        public string FontName { get; set; } = string.Empty;
        public double FontSize { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        /// <summary>Hex RGB color string, e.g. "000000" for black.</summary>
        public string Color { get; set; } = "000000";

        public override string ToString() => $"\"{Text}\" at {Bounds}";
    }

    /// <summary>
    /// Represents an image extracted from a PDF page.
    /// </summary>
    public class ImageElement
    {
        public Rect Bounds { get; set; } = new Rect();
        public byte[] ImageData { get; set; } = Array.Empty<byte>();
        /// <summary>File format: "png", "jpeg", "gif", etc.</summary>
        public string Format { get; set; } = "png";
        public double WidthPx { get; set; }
        public double HeightPx { get; set; }
    }

    /// <summary>
    /// Represents a straight line segment on the PDF page (for table detection).
    /// Coordinates use top-left origin.
    /// </summary>
    public class LineSegment
    {
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
        public double Thickness { get; set; } = 1.0;
        /// <summary>Hex RGB color string.</summary>
        public string Color { get; set; } = "000000";

        /// <summary>True if the line is roughly horizontal (within 5 degrees).</summary>
        public bool IsHorizontal => Math.Abs(Y2 - Y1) < Math.Max(Math.Abs(X2 - X1) * 0.1, 0.5);
        /// <summary>True if the line is roughly vertical (within 5 degrees).</summary>
        public bool IsVertical => Math.Abs(X2 - X1) < Math.Max(Math.Abs(Y2 - Y1) * 0.1, 0.5);

        public double Length => Math.Sqrt(Math.Pow(X2 - X1, 2) + Math.Pow(Y2 - Y1, 2));

        /// <summary>Normalize so X1 &lt;= X2 for horizontal; Y1 &lt;= Y2 for vertical.</summary>
        public void Normalize()
        {
            if (IsHorizontal && X1 > X2)
            {
                (X1, X2) = (X2, X1);
                (Y1, Y2) = (Y2, Y1);
            }
            else if (IsVertical && Y1 > Y2)
            {
                (X1, X2) = (X2, X1);
                (Y1, Y2) = (Y2, Y1);
            }
        }
    }

    /// <summary>
    /// Represents a filled or stroked rectangle on the PDF page.
    /// Used for detecting cell shading and table boundaries.
    /// </summary>
    public class RectangleElement
    {
        public Rect Bounds { get; set; } = new Rect();
        public string FillColor { get; set; } = string.Empty;
        public string StrokeColor { get; set; } = string.Empty;
        public double StrokeWidth { get; set; }
        public bool IsFilled { get; set; }
        public bool IsStroked { get; set; }
    }

    /// <summary>
    /// Hyperlink annotation from the PDF.
    /// </summary>
    public class HyperlinkInfo
    {
        public Rect Bounds { get; set; } = new Rect();
        public string Uri { get; set; } = string.Empty;
    }

    /// <summary>
    /// A contiguous run of text with uniform formatting within a line.
    /// </summary>
    public class TextRun
    {
        public string Text { get; set; } = string.Empty;
        public string FontName { get; set; } = string.Empty;
        public double FontSize { get; set; }
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        /// <summary>Hex RGB color string.</summary>
        public string Color { get; set; } = "000000";
        /// <summary>Optional hyperlink URI.</summary>
        public string HyperlinkUri { get; set; } = string.Empty;

        /// <summary>
        /// Returns true if this run has the same formatting as another.
        /// </summary>
        public bool HasSameFormatting(TextRun other)
        {
            return FontName == other.FontName &&
                   Math.Abs(FontSize - other.FontSize) < 0.5 &&
                   IsBold == other.IsBold &&
                   IsItalic == other.IsItalic &&
                   Color == other.Color &&
                   HyperlinkUri == other.HyperlinkUri;
        }
    }

    /// <summary>
    /// A single line of text, composed of one or more runs.
    /// </summary>
    public class TextLine
    {
        public List<TextRun> Runs { get; set; } = new List<TextRun>();
        public Rect Bounds { get; set; } = new Rect();

        /// <summary>
        /// Font-based line height in PDF points, set by <see cref="Layout.LayoutAnalyzer"/>
        /// using <c>DominantFontSize × LineSpacingMultiplier</c>.
        /// 0 means "use bounding box height" (legacy fallback).
        /// </summary>
        public double LineHeight { get; set; } = 0.0;

        /// <summary>
        /// The full text of the line (concatenation of all runs).
        /// </summary>
        public string FullText
        {
            get
            {
                var sb = new System.Text.StringBuilder();
                foreach (var run in Runs) sb.Append(run.Text);
                return sb.ToString();
            }
        }

        /// <summary>The dominant font size in this line.</summary>
        public double DominantFontSize
        {
            get
            {
                if (Runs.Count == 0) return 12;
                double maxLen = 0;
                double fontSize = 12;
                foreach (var run in Runs)
                {
                    if (run.Text.Length > maxLen)
                    {
                        maxLen = run.Text.Length;
                        fontSize = run.FontSize;
                    }
                }
                return fontSize;
            }
        }
    }

    /// <summary>
    /// A paragraph composed of one or more text lines.
    /// </summary>
    public class TextParagraph
    {
        public List<TextLine> Lines { get; set; } = new List<TextLine>();
        public Rect Bounds { get; set; } = new Rect();
        /// <summary>"left", "center", "right", or "justify".</summary>
        public string Alignment { get; set; } = "left";

        /// <summary>All runs across all lines, flattened.</summary>
        public List<TextRun> AllRuns
        {
            get
            {
                var runs = new List<TextRun>();
                foreach (var line in Lines)
                    runs.AddRange(line.Runs);
                return runs;
            }
        }
    }

    /// <summary>
    /// Border style for a table cell edge.
    /// </summary>
    public class BorderStyle
    {
        public double Width { get; set; }
        /// <summary>Hex RGB color string.</summary>
        public string Color { get; set; } = "000000";
        /// <summary>OOXML border style: "single", "dashed", "dotted", "none".</summary>
        public string Style { get; set; } = "single";

        public static BorderStyle None => new BorderStyle { Width = 0, Style = "none", Color = "000000" };
        public static BorderStyle Default => new BorderStyle { Width = 0.5, Style = "single", Color = "000000" };
    }

    /// <summary>
    /// A single cell in a detected table.
    /// </summary>
    public class TableCell
    {
        public int Row { get; set; }
        public int Col { get; set; }
        public int RowSpan { get; set; } = 1;
        public int ColSpan { get; set; } = 1;
        public Rect Bounds { get; set; } = new Rect();
        public List<TextParagraph> Paragraphs { get; set; } = new List<TextParagraph>();
        /// <summary>Hex RGB background color, or empty for no shading.</summary>
        public string BackgroundColor { get; set; } = string.Empty;
        public BorderStyle TopBorder { get; set; } = BorderStyle.Default;
        public BorderStyle BottomBorder { get; set; } = BorderStyle.Default;
        public BorderStyle LeftBorder { get; set; } = BorderStyle.Default;
        public BorderStyle RightBorder { get; set; } = BorderStyle.Default;
        /// <summary>True if this cell is part of a merge but not the origin cell.</summary>
        public bool IsMergedContinuation { get; set; }
    }

    /// <summary>
    /// A complete detected table with cells, dimensions, and styles.
    /// </summary>
    public class DetectedTable
    {
        public Rect Bounds { get; set; } = new Rect();
        public int RowCount { get; set; }
        public int ColCount { get; set; }
        /// <summary>Cell grid [row, col]. Some cells may be merge continuations.</summary>
        public TableCell[,] Cells { get; set; } = new TableCell[0, 0];
        /// <summary>Column widths in PDF points.</summary>
        public double[] ColumnWidths { get; set; } = Array.Empty<double>();
        /// <summary>Row heights in PDF points.</summary>
        public double[] RowHeights { get; set; } = Array.Empty<double>();
    }

    /// <summary>
    /// Types of content blocks in the page layout.
    /// </summary>
    public enum ContentBlockType
    {
        Paragraph,
        Table,
        Image
    }

    /// <summary>
    /// A logical block of content on a page (paragraph, table, or image).
    /// Blocks are ordered in reading sequence.
    /// </summary>
    public class ContentBlock
    {
        public ContentBlockType Type { get; set; }
        public TextParagraph? Paragraph { get; set; }
        public DetectedTable? Table { get; set; }
        public ImageElement? Image { get; set; }
        public Rect Bounds { get; set; } = new Rect();
    }
}
