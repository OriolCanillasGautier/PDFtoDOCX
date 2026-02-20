using System;
using System.Collections.Generic;
using System.Linq;
using PDFtoDOCX.Models;

namespace PDFtoDOCX.Layout
{
    /// <summary>
    /// Analyzes extracted PDF page content to reconstruct the logical document layout.
    /// Groups text elements into lines, detects columns, and assembles paragraphs.
    /// </summary>
    public class LayoutAnalyzer
    {
        private readonly ConversionOptions _options;

        public LayoutAnalyzer(ConversionOptions options)
        {
            _options = options ?? new ConversionOptions();
        }

        /// <summary>
        /// Analyzes a page's text elements and produces a list of text paragraphs
        /// in reading order. Does NOT handle table regions — those text elements
        /// should be excluded before calling this method.
        /// </summary>
        public List<TextParagraph> Analyze(List<TextElement> textElements, double pageWidth, double pageHeight)
        {
            if (textElements == null || textElements.Count == 0)
                return new List<TextParagraph>();

            // Step 1: Group text elements into lines
            var lines = GroupIntoLines(textElements);

            // Step 2: Detect columns
            var columns = DetectColumns(lines, pageWidth);

            // Step 3: Order lines within each column top-to-bottom
            foreach (var column in columns)
            {
                column.Sort((a, b) => a.Bounds.Top.CompareTo(b.Bounds.Top));
            }

            // Step 4: Merge columns into reading order (left-to-right columns)
            var orderedLines = new List<TextLine>();
            foreach (var column in columns)
            {
                orderedLines.AddRange(column);
            }

            // Step 5: Group lines into paragraphs
            var paragraphs = GroupIntoParagraphs(orderedLines);

            // Step 6: Detect alignment for each paragraph
            foreach (var para in paragraphs)
            {
                para.Alignment = DetectAlignment(para, pageWidth);
            }

            return paragraphs;
        }

        /// <summary>
        /// Groups individual text elements into horizontal lines by Y-coordinate clustering.
        /// </summary>
        public List<TextLine> GroupIntoLines(List<TextElement> elements)
        {
            if (elements.Count == 0) return new List<TextLine>();

            // Sort by Y (top), then by X (left)
            var sorted = elements.OrderBy(e => e.Bounds.Top).ThenBy(e => e.Bounds.Left).ToList();

            var lines = new List<TextLine>();
            var currentLineElements = new List<TextElement> { sorted[0] };
            double currentLineY = sorted[0].Bounds.MidY;

            for (int i = 1; i < sorted.Count; i++)
            {
                var elem = sorted[i];
                double elemMidY = elem.Bounds.MidY;

                // Dynamic tolerance based on font size
                double tolerance = Math.Max(_options.LineGroupingTolerance,
                    elem.Bounds.Height * 0.5);

                if (Math.Abs(elemMidY - currentLineY) <= tolerance)
                {
                    // Same line
                    currentLineElements.Add(elem);
                    // Update running average Y
                    currentLineY = currentLineElements.Average(e => e.Bounds.MidY);
                }
                else
                {
                    // New line
                    lines.Add(BuildTextLine(currentLineElements));
                    currentLineElements = new List<TextElement> { elem };
                    currentLineY = elemMidY;
                }
            }

            // Don't forget the last line
            if (currentLineElements.Count > 0)
                lines.Add(BuildTextLine(currentLineElements));

            return lines;
        }

        /// <summary>
        /// Builds a TextLine from a collection of TextElements on the same horizontal band.
        /// Elements are sorted left-to-right, and adjacent elements with the same
        /// formatting are merged into single TextRuns.
        /// </summary>
        private TextLine BuildTextLine(List<TextElement> elements)
        {
            // Sort elements left-to-right
            elements.Sort((a, b) => a.Bounds.Left.CompareTo(b.Bounds.Left));

            var runs = new List<TextRun>();
            TextRun currentRun = null;
            TextElement prevElement = null;

            foreach (var elem in elements)
            {
                bool needsSpace = false;
                if (prevElement != null)
                {
                    // Calculate horizontal gap between previous element's right edge
                    // and this element's left edge
                    double gap = elem.Bounds.Left - prevElement.Bounds.Right;
                    // Average character width approximation
                    double avgCharWidth = prevElement.Bounds.Width /
                        Math.Max(1, prevElement.Text.Length);

                    // If gap is more than ~30% of average char width, insert a space
                    if (gap > avgCharWidth * 0.3)
                        needsSpace = true;
                }

                var newRun = new TextRun
                {
                    Text = elem.Text,
                    FontName = elem.FontName,
                    FontSize = elem.FontSize,
                    IsBold = elem.IsBold,
                    IsItalic = elem.IsItalic,
                    Color = elem.Color
                };

                if (currentRun != null && currentRun.HasSameFormatting(newRun))
                {
                    // Merge into current run
                    currentRun.Text += (needsSpace ? " " : "") + elem.Text;
                }
                else
                {
                    // Start a new run
                    if (currentRun != null)
                        runs.Add(currentRun);

                    if (needsSpace && currentRun != null)
                    {
                        // Append trailing space to previous run or prepend to new
                        newRun.Text = " " + newRun.Text;
                    }

                    currentRun = newRun;
                }

                prevElement = elem;
            }

            if (currentRun != null)
                runs.Add(currentRun);

            // Calculate bounding box
            double left = elements.Min(e => e.Bounds.Left);
            double top = elements.Min(e => e.Bounds.Top);
            double right = elements.Max(e => e.Bounds.Right);
            double bottom = elements.Max(e => e.Bounds.Bottom);

            return new TextLine
            {
                Runs = runs,
                Bounds = new Rect(left, top, right, bottom)
            };
        }

        /// <summary>
        /// Detects columns by clustering lines by their horizontal (X) position.
        /// Returns a list of columns (each column is a list of lines), ordered left-to-right.
        /// </summary>
        public List<List<TextLine>> DetectColumns(List<TextLine> lines, double pageWidth)
        {
            if (lines.Count == 0)
                return new List<List<TextLine>>();

            // Cluster lines by their left X position
            // Use a simple approach: sort lines by left edge and detect gaps
            var sortedByX = lines.OrderBy(l => l.Bounds.Left).ToList();

            // Find the distribution of left-edge positions
            var leftEdges = sortedByX.Select(l => l.Bounds.Left).OrderBy(x => x).ToList();
            var rightEdges = sortedByX.Select(l => l.Bounds.Right).OrderBy(x => x).ToList();

            // Try to detect column boundaries by looking for vertical gaps in the text
            // A "gap" is a horizontal region with no text
            var columnRanges = FindColumnRanges(lines, pageWidth);

            if (columnRanges.Count <= 1)
            {
                // Single column - return all lines as one column
                return new List<List<TextLine>> { lines.ToList() };
            }

            // Assign each line to a column based on maximum overlap
            var columns = new List<List<TextLine>>();
            for (int i = 0; i < columnRanges.Count; i++)
                columns.Add(new List<TextLine>());

            foreach (var line in lines)
            {
                int bestCol = 0;
                double bestOverlap = 0;

                for (int i = 0; i < columnRanges.Count; i++)
                {
                    double overlapLeft = Math.Max(line.Bounds.Left, columnRanges[i].Item1);
                    double overlapRight = Math.Min(line.Bounds.Right, columnRanges[i].Item2);
                    double overlap = Math.Max(0, overlapRight - overlapLeft);

                    if (overlap > bestOverlap)
                    {
                        bestOverlap = overlap;
                        bestCol = i;
                    }
                }

                columns[bestCol].Add(line);
            }

            // Remove empty columns
            columns.RemoveAll(c => c.Count == 0);

            return columns;
        }

        /// <summary>
        /// Finds column ranges by locating X positions where no text line crosses through.
        /// A true multi-column layout has a clear vertical gap that no line spans across.
        /// Varying-length lines in a single-column document do NOT create column splits.
        /// Returns a list of (leftX, rightX) tuples for each detected column.
        /// </summary>
        private List<(double, double)> FindColumnRanges(List<TextLine> lines, double pageWidth)
        {
            if (lines.Count == 0)
                return new List<(double, double)> { (0, pageWidth) };

            // Resolution: 1-point steps give precise gap detection without float noise
            int resolution = (int)Math.Ceiling(pageWidth) + 1;
            // For each X bucket, count how many lines span across it (left <= x <= right)
            int[] crossCount = new int[resolution];

            foreach (var line in lines)
            {
                int left  = Math.Max(0, (int)line.Bounds.Left);
                int right = Math.Min(resolution - 1, (int)Math.Ceiling(line.Bounds.Right));
                for (int x = left; x <= right; x++)
                    crossCount[x]++;
            }

            // A column gap is a contiguous X range where NO line crosses through.
            // We do not count the page margins (first/last 10% of width) as column gaps.
            double marginX = pageWidth * 0.10;
            int marginBucket = (int)marginX;

            var gapRegions = new List<(int start, int end)>();
            int? gapStart = null;

            for (int x = marginBucket; x < resolution - marginBucket; x++)
            {
                if (crossCount[x] == 0)
                {
                    if (gapStart == null) gapStart = x;
                }
                else
                {
                    if (gapStart != null)
                    {
                        gapRegions.Add((gapStart.Value, x - 1));
                        gapStart = null;
                    }
                }
            }
            if (gapStart != null)
                gapRegions.Add((gapStart.Value, resolution - marginBucket - 1));

            // Keep only gaps that are at least MinColumnGap wide
            var significantGaps = gapRegions
                .Where(g => (g.end - g.start) >= _options.MinColumnGap)
                .ToList();

            if (significantGaps.Count == 0)
                return new List<(double, double)> { (0, pageWidth) };

            // Build column ranges from the gaps
            var columns = new List<(double, double)>();
            double colStart = 0;

            foreach (var gap in significantGaps)
            {
                double colEnd = gap.start;
                if (colEnd > colStart)
                    columns.Add((colStart, colEnd));
                colStart = gap.end + 1;
            }

            // Last column after the final gap
            if (colStart < pageWidth)
                columns.Add((colStart, pageWidth));

            return columns.Count > 0 ? columns : new List<(double, double)> { (0, pageWidth) };
        }

        /// <summary>
        /// Groups consecutive lines into paragraphs based on vertical spacing.
        /// A new paragraph starts when the gap between lines exceeds the threshold.
        /// </summary>
        public List<TextParagraph> GroupIntoParagraphs(List<TextLine> lines)
        {
            if (lines.Count == 0)
                return new List<TextParagraph>();

            // Calculate average line spacing
            var lineGaps = new List<double>();
            for (int i = 1; i < lines.Count; i++)
            {
                double gap = lines[i].Bounds.Top - lines[i - 1].Bounds.Bottom;
                if (gap > 0) lineGaps.Add(gap);
            }

            double avgGap = lineGaps.Count > 0 ? lineGaps.Average() : 5.0;
            double avgLineHeight = lines.Average(l => l.Bounds.Height);
            double paraThreshold = avgLineHeight * _options.ParagraphGapMultiplier;

            var paragraphs = new List<TextParagraph>();
            var currentLines = new List<TextLine> { lines[0] };

            for (int i = 1; i < lines.Count; i++)
            {
                double gap = lines[i].Bounds.Top - lines[i - 1].Bounds.Bottom;

                // Also check for significant font size changes (possible heading boundary)
                bool fontSizeChanged = Math.Abs(lines[i].DominantFontSize -
                    lines[i - 1].DominantFontSize) > 2.0;

                // Check for significant left-indent change (possible new paragraph or list item)
                bool indentChanged = Math.Abs(lines[i].Bounds.Left -
                    lines[i - 1].Bounds.Left) > avgLineHeight;

                if (gap > paraThreshold || fontSizeChanged || indentChanged)
                {
                    // Start new paragraph
                    paragraphs.Add(BuildParagraph(currentLines));
                    currentLines = new List<TextLine> { lines[i] };
                }
                else
                {
                    currentLines.Add(lines[i]);
                }
            }

            if (currentLines.Count > 0)
                paragraphs.Add(BuildParagraph(currentLines));

            return paragraphs;
        }

        /// <summary>
        /// Builds a TextParagraph from a group of consecutive lines.
        /// Handles joining lines within a paragraph (e.g., re-inserting spaces).
        /// Stores a font-size-based <see cref="TextLine.LineHeight"/> on each line
        /// so the packager can use it for accurate spacing.
        /// </summary>
        private TextParagraph BuildParagraph(List<TextLine> lines)
        {
            double left   = lines.Min(l => l.Bounds.Left);
            double top    = lines.Min(l => l.Bounds.Top);
            double right  = lines.Max(l => l.Bounds.Right);
            double bottom = lines.Max(l => l.Bounds.Bottom);

            // Phase 2.1: store font-based line height on each line
            foreach (var line in lines)
            {
                double fontSize = line.DominantFontSize > 0 ? line.DominantFontSize : 12.0;
                line.LineHeight = fontSize * _options.LineSpacingMultiplier;
            }

            return new TextParagraph
            {
                Lines  = lines,
                Bounds = new Rect(left, top, right, bottom)
            };
        }

        /// <summary>
        /// Detects paragraph alignment by comparing text bounds to the page width.
        /// </summary>
        private string DetectAlignment(TextParagraph para, double pageWidth)
        {
            if (para.Lines.Count == 0) return "left";

            // Margins
            double leftMargin = _options.DefaultMargin;
            double rightMargin = pageWidth - _options.DefaultMargin;
            double textAreaWidth = rightMargin - leftMargin;
            double pageCenter = pageWidth / 2.0;

            // Check line positions relative to the text area
            double avgLeft = para.Lines.Average(l => l.Bounds.Left);
            double avgRight = para.Lines.Average(l => l.Bounds.Right);

            // Justified: all lines except last span nearly full text width
            bool allLinesFullWidth = para.Lines.Count > 1 &&
                para.Lines.Take(para.Lines.Count - 1)
                    .All(l => (l.Bounds.Right - l.Bounds.Left) > textAreaWidth * 0.9);

            if (allLinesFullWidth)
                return "justify";

            // Center-aligned: requires 2+ lines, all must be tightly centered,
            // and the block must not start at or near the left margin.
            if (para.Lines.Count >= 2)
            {
                // Tight threshold: within 5% of text area width OR 15pt — whichever is smaller
                double centerThreshold = Math.Min(textAreaWidth * 0.05, 15.0);
                bool allCentered = para.Lines.All(l =>
                    Math.Abs(l.Bounds.MidX - pageCenter) < centerThreshold);
                // Also ensure the paragraph isn't simply left-aligned full-width text
                bool notLeftOrigin = avgLeft > leftMargin + 20.0;

                if (allCentered && notLeftOrigin)
                    return "center";
            }

            // Right-aligned: requires 2+ lines with consistent right edge near the right margin
            if (para.Lines.Count >= 2)
            {
                double maxRightDiff = 0;
                for (int i = 1; i < para.Lines.Count; i++)
                    maxRightDiff = Math.Max(maxRightDiff,
                        Math.Abs(para.Lines[i].Bounds.Right - para.Lines[0].Bounds.Right));

                bool rightAligned = maxRightDiff < 5.0 &&
                    Math.Abs(avgRight - rightMargin) < 10.0 &&
                    (avgLeft - leftMargin) > 20.0;

                if (rightAligned)
                    return "right";
            }

            return "left";
        }

        /// <summary>
        /// Filters out text elements that fall within any of the given table boundaries.
        /// Returns elements outside all tables.
        /// Uses the same midpoint criterion as ElementsInRegion to avoid double-counting.
        /// </summary>
        public static List<TextElement> ExcludeTableRegions(
            List<TextElement> elements, List<DetectedTable> tables)
        {
            if (tables == null || tables.Count == 0)
                return elements;

            return elements.Where(e =>
                !tables.Any(t =>
                    e.Bounds.MidY >= t.Bounds.Top   &&
                    e.Bounds.MidY <= t.Bounds.Bottom &&
                    e.Bounds.MidX >= t.Bounds.Left   &&
                    e.Bounds.MidX <= t.Bounds.Right)).ToList();
        }

        /// <summary>
        /// Filters text elements that overlap significantly with a specific rectangle.
        /// Uses horizontal overlap ≥ 50% of the element width and vertical center containment,
        /// which is more robust than pure midpoint containment when text sits at cell edges.
        /// </summary>
        public static List<TextElement> ElementsInRegion(
            List<TextElement> elements, Rect region)
        {
            // Expand vertically by a small tolerance to handle border-touching elements
            const double tol = 2.0;
            double rTop    = region.Top    - tol;
            double rBottom = region.Bottom + tol;
            double rLeft   = region.Left   - tol;
            double rRight  = region.Right  + tol;

            return elements.Where(e =>
            {
                // Vertical: element's vertical midpoint must be inside the region
                if (e.Bounds.MidY < rTop || e.Bounds.MidY > rBottom)
                    return false;

                // Horizontal: at least 50% of the element must overlap the region
                double overlapLeft  = Math.Max(e.Bounds.Left,  rLeft);
                double overlapRight = Math.Min(e.Bounds.Right, rRight);
                double overlap = overlapRight - overlapLeft;
                double elemWidth = Math.Max(1.0, e.Bounds.Width);
                return overlap / elemWidth >= 0.5;
            }).ToList();
        }

        /// <summary>
        /// Calculates what fraction of 'inner' is overlapped by 'outer'.
        /// </summary>
        private static double OverlapRatio(Rect inner, Rect outer)
        {
            double overlapLeft = Math.Max(inner.Left, outer.Left);
            double overlapTop = Math.Max(inner.Top, outer.Top);
            double overlapRight = Math.Min(inner.Right, outer.Right);
            double overlapBottom = Math.Min(inner.Bottom, outer.Bottom);

            double overlapWidth = Math.Max(0, overlapRight - overlapLeft);
            double overlapHeight = Math.Max(0, overlapBottom - overlapTop);
            double overlapArea = overlapWidth * overlapHeight;
            double innerArea = inner.Width * inner.Height;

            return innerArea > 0 ? overlapArea / innerArea : 0;
        }
    }
}
