using System;
using System.Collections.Generic;
using System.Linq;
using PDFtoDOCX.Models;

namespace PDFtoDOCX.Tables
{
    /// <summary>
    /// Detects tables in PDF page content by analyzing line segments (grid detection),
    /// reconstructs cell structure including merged cells, populates cells with text,
    /// and maps border styles via line properties.
    /// </summary>
    public class TableDetector
    {
        private readonly ConversionOptions _options;
        private readonly Layout.LayoutAnalyzer _layoutAnalyzer;

        public TableDetector(ConversionOptions options)
        {
            _options = options ?? new ConversionOptions();
            _layoutAnalyzer = new Layout.LayoutAnalyzer(options);
        }

        /// <summary>
        /// Detects all tables in the given page content.
        /// Returns a list of fully populated DetectedTable structures.
        /// </summary>
        public List<DetectedTable> DetectTables(PageContent content)
        {
            if (content.Lines == null || content.Lines.Count == 0)
                return new List<DetectedTable>();

            // Step 1: Separate lines into horizontal and vertical
            var hLines = content.Lines.Where(l => l.IsHorizontal && l.Length >= _options.MinTableLineLength).ToList();
            var vLines = content.Lines.Where(l => l.IsVertical && l.Length >= _options.MinTableLineLength).ToList();

            if (hLines.Count < 2 || vLines.Count < 2)
                return new List<DetectedTable>();

            // Step 2: Find clusters of intersecting horizontal and vertical lines (grids)
            var grids = FindGrids(hLines, vLines, content.Width, content.Height);

            // Step 3: Build table structures from grids
            var tables = new List<DetectedTable>();
            foreach (var grid in grids)
            {
                var table = BuildTableFromGrid(grid, content);
                if (table != null &&
                    table.RowCount >= _options.MinTableRows &&
                    table.ColCount >= _options.MinTableCols)
                {
                    double confidence = CalculateConfidence(table, grid);
                    if (_options.EnableDiagnostics)
                        System.Console.WriteLine(
                            $"[TableDetector] Page {content.PageNumber}: {table.RowCount}×{table.ColCount} table " +
                            $"at ({table.Bounds.Left:F0},{table.Bounds.Top:F0}) confidence={confidence:F2}");

                    if (confidence >= 0.4)
                        tables.Add(table);
                    else if (_options.EnableDiagnostics)
                        System.Console.WriteLine(
                            $"[TableDetector] Rejected low-confidence table (confidence={confidence:F2})");
                }
            }

            return tables;
        }

        /// <summary>
        /// Calculates a confidence score [0..1] for a detected table.
        /// Factors: interior line density, cell text coverage, border completeness.
        /// </summary>
        private double CalculateConfidence(DetectedTable table, GridCandidate grid)
        {
            double snap = _options.TableGridSnapTolerance;
            double score = 0.0;

            // ── Factor 1: Interior line density (0–0.4) ─────────────────────
            int expectedInteriorH = table.RowCount - 1;
            int expectedInteriorV = table.ColCount - 1;
            int foundH = 0, foundV = 0;

            for (int r = 1; r < grid.YPositions.Count - 1; r++)
                if (CountLinesAtPosition(grid.HorizontalLines, grid.YPositions[r], snap, true) > 0)
                    foundH++;

            for (int c = 1; c < grid.XPositions.Count - 1; c++)
                if (CountLinesAtPosition(grid.VerticalLines, grid.XPositions[c], snap, false) > 0)
                    foundV++;

            double densityH = expectedInteriorH > 0 ? (double)foundH / expectedInteriorH : 1.0;
            double densityV = expectedInteriorV > 0 ? (double)foundV / expectedInteriorV : 1.0;
            score += (densityH + densityV) / 2.0 * 0.4;

            // ── Factor 2: Cell text coverage (0–0.4) ────────────────────────
            int totalCells = 0, cellsWithText = 0;
            for (int r = 0; r < table.RowCount; r++)
            {
                for (int c = 0; c < table.ColCount; c++)
                {
                    var cell = table.Cells[r, c];
                    if (cell.IsMergedContinuation) continue;
                    totalCells++;

                    bool hasText = cell.Paragraphs.Count > 0 &&
                        cell.Paragraphs.Any(p =>
                            p.Lines.Any(l =>
                                l.Runs.Any(run => !string.IsNullOrWhiteSpace(run.Text))));
                    if (hasText) cellsWithText++;
                }
            }
            // Even empty tables are valid; use 0.5 as a base when all cells are empty
            double textCoverage = totalCells > 0 ? (double)cellsWithText / totalCells : 0.0;
            score += textCoverage * 0.4;

            // ── Factor 3: Outer border completeness (0–0.2) ─────────────────
            bool hasTop    = CountLinesAtPosition(grid.HorizontalLines, grid.YPositions.First(), snap, true) > 0;
            bool hasBottom = CountLinesAtPosition(grid.HorizontalLines, grid.YPositions.Last(),  snap, true) > 0;
            bool hasLeft   = CountLinesAtPosition(grid.VerticalLines,   grid.XPositions.First(), snap, false) > 0;
            bool hasRight  = CountLinesAtPosition(grid.VerticalLines,   grid.XPositions.Last(),  snap, false) > 0;
            int borderCount = (hasTop ? 1 : 0) + (hasBottom ? 1 : 0) + (hasLeft ? 1 : 0) + (hasRight ? 1 : 0);
            score += (borderCount / 4.0) * 0.2;

            return score;
        }

        /// <summary>
        /// Finds grids by clustering horizontal and vertical lines that intersect.
        /// A grid is a set of horizontal and vertical lines that form a closed structure.
        /// </summary>
        private List<GridCandidate> FindGrids(List<LineSegment> hLines, List<LineSegment> vLines,
            double pageWidth, double pageHeight)
        {
            var grids = new List<GridCandidate>();
            double snap = _options.TableGridSnapTolerance;

            // Get unique Y positions from horizontal lines (snapped)
            var yPositions = SnapAndDedupe(hLines.SelectMany(l => new[] { l.Y1, l.Y2 }), snap);
            // Get unique X positions from vertical lines (snapped)
            var xPositions = SnapAndDedupe(vLines.SelectMany(l => new[] { l.X1, l.X2 }), snap);

            if (yPositions.Count < 2 || xPositions.Count < 2)
                return grids;

            // Find connected groups of X/Y positions that form grids
            // For simplicity, try to build a single grid from all positions first
            // Then validate that sufficient lines exist between adjacent positions

            // Build intersection matrix
            var gridLines = new GridCandidate
            {
                XPositions = xPositions,
                YPositions = yPositions,
                HorizontalLines = hLines,
                VerticalLines = vLines
            };

            // Validate that this forms a coherent grid
            if (ValidateGrid(gridLines, pageWidth, pageHeight))
            {
                grids.Add(gridLines);
            }
            else
            {
                // Try to find sub-grids by clustering lines spatially
                var subGrids = FindSubGrids(hLines, vLines, snap, pageWidth, pageHeight);
                grids.AddRange(subGrids);
            }

            return grids;
        }

        /// <summary>
        /// Snaps values to a grid defined by the tolerance, deduplicates, and sorts.
        /// </summary>
        private List<double> SnapAndDedupe(IEnumerable<double> values, double tolerance)
        {
            var sorted = values.OrderBy(v => v).ToList();
            var result = new List<double>();

            foreach (var val in sorted)
            {
                bool merged = false;
                for (int i = 0; i < result.Count; i++)
                {
                    if (Math.Abs(result[i] - val) <= tolerance)
                    {
                        // Average the two values (snap to midpoint)
                        result[i] = (result[i] + val) / 2.0;
                        merged = true;
                        break;
                    }
                }
                if (!merged)
                    result.Add(val);
            }

            result.Sort();
            return result;
        }

        /// <summary>
        /// Validates that the candidate grid has sufficient line coverage
        /// to be considered a real table.
        /// Rejects 1×N or M×1 "tables", grids that span &gt;80% of the page
        /// (likely page borders), and grids with insufficient interior lines.
        /// </summary>
        private bool ValidateGrid(GridCandidate grid, double pageWidth = 0, double pageHeight = 0)
        {
            double snap = _options.TableGridSnapTolerance;

            int rows = grid.YPositions.Count - 1;
            int cols = grid.XPositions.Count - 1;

            // Must have at least 2 rows and 2 columns to be a real table
            if (rows < 2 || cols < 2) return false;

            // Reject grids that span >80% of the page in BOTH dimensions (likely page frames)
            if (pageWidth > 0 && pageHeight > 0)
            {
                double gridWidth  = grid.XPositions.Last() - grid.XPositions.First();
                double gridHeight = grid.YPositions.Last() - grid.YPositions.First();
                if (gridWidth  > pageWidth  * 0.80 &&
                    gridHeight > pageHeight * 0.80)
                    return false;
            }

            // Check outer boundary lines exist
            if (CountLinesAtPosition(grid.HorizontalLines, grid.YPositions.First(), snap, true) == 0) return false;
            if (CountLinesAtPosition(grid.HorizontalLines, grid.YPositions.Last(), snap, true) == 0) return false;
            if (CountLinesAtPosition(grid.VerticalLines, grid.XPositions.First(), snap, false) == 0) return false;
            if (CountLinesAtPosition(grid.VerticalLines, grid.XPositions.Last(), snap, false) == 0) return false;

            // Check that interior horizontal lines exist (at least half the expected ones)
            int expectedInteriorH = rows - 1;
            int foundInteriorH = 0;
            for (int r = 1; r < grid.YPositions.Count - 1; r++)
                if (CountLinesAtPosition(grid.HorizontalLines, grid.YPositions[r], snap, true) > 0)
                    foundInteriorH++;

            // Check that interior vertical lines exist (at least half the expected ones)
            int expectedInteriorV = cols - 1;
            int foundInteriorV = 0;
            for (int c = 1; c < grid.XPositions.Count - 1; c++)
                if (CountLinesAtPosition(grid.VerticalLines, grid.XPositions[c], snap, false) > 0)
                    foundInteriorV++;

            // Require interior lines: at least 50% for H and V
            if (foundInteriorH < Math.Max(1, expectedInteriorH * 0.5)) return false;
            if (foundInteriorV < Math.Max(1, expectedInteriorV * 0.5)) return false;

            return true;
        }

        /// <summary>
        /// Counts how many lines exist at a given Y (horizontal) or X (vertical) position.
        /// </summary>
        private int CountLinesAtPosition(List<LineSegment> lines, double position, double tolerance, bool horizontal)
        {
            return lines.Count(l =>
            {
                if (horizontal)
                    return Math.Abs(l.Y1 - position) <= tolerance || Math.Abs(l.Y2 - position) <= tolerance;
                else
                    return Math.Abs(l.X1 - position) <= tolerance || Math.Abs(l.X2 - position) <= tolerance;
            });
        }

        /// <summary>
        /// Finds sub-grids within the full set of lines by spatial clustering.
        /// </summary>
        private List<GridCandidate> FindSubGrids(List<LineSegment> hLines, List<LineSegment> vLines,
            double snap, double pageWidth = 0, double pageHeight = 0)
        {
            var results = new List<GridCandidate>();

            // Cluster horizontal lines by Y position
            var hClusters = ClusterLinesByPosition(hLines, true, snap * 3);
            var vClusters = ClusterLinesByPosition(vLines, false, snap * 3);

            foreach (var hCluster in hClusters)
            {
                foreach (var vCluster in vClusters)
                {
                    // Check if these clusters overlap spatially
                    double hMinX = hCluster.Min(l => Math.Min(l.X1, l.X2));
                    double hMaxX = hCluster.Max(l => Math.Max(l.X1, l.X2));
                    double hMinY = hCluster.Min(l => Math.Min(l.Y1, l.Y2));
                    double hMaxY = hCluster.Max(l => Math.Max(l.Y1, l.Y2));

                    double vMinX = vCluster.Min(l => Math.Min(l.X1, l.X2));
                    double vMaxX = vCluster.Max(l => Math.Max(l.X1, l.X2));
                    double vMinY = vCluster.Min(l => Math.Min(l.Y1, l.Y2));
                    double vMaxY = vCluster.Max(l => Math.Max(l.Y1, l.Y2));

                    // Overlap check
                    if (hMinX < vMaxX + snap && hMaxX > vMinX - snap &&
                        hMinY < vMaxY + snap && hMaxY > vMinY - snap)
                    {
                        var yPositions = SnapAndDedupe(hCluster.SelectMany(l => new[] { l.Y1, l.Y2 }), snap);
                        var xPositions = SnapAndDedupe(vCluster.SelectMany(l => new[] { l.X1, l.X2 }), snap);

                        if (yPositions.Count >= 2 && xPositions.Count >= 2)
                        {
                            var candidate = new GridCandidate
                            {
                                XPositions = xPositions,
                                YPositions = yPositions,
                                HorizontalLines = hCluster,
                                VerticalLines = vCluster
                            };

                            if (ValidateGrid(candidate, pageWidth, pageHeight))
                                results.Add(candidate);
                        }
                    }
                }
            }

            // Remove duplicates / overlapping grids (keep largest)
            return RemoveOverlapping(results);
        }

        /// <summary>
        /// Clusters lines by their primary position (Y for horizontal, X for vertical).
        /// Returns groups of lines that are spatially close.
        /// </summary>
        private List<List<LineSegment>> ClusterLinesByPosition(List<LineSegment> lines, bool horizontal, double maxGap)
        {
            // Get the range positions
            var positions = lines.Select(l => horizontal ?
                (Math.Min(l.Y1, l.Y2) + Math.Max(l.Y1, l.Y2)) / 2 :
                (Math.Min(l.X1, l.X2) + Math.Max(l.X1, l.X2)) / 2)
                .Zip(lines, (pos, line) => (pos, line))
                .OrderBy(x => x.pos)
                .ToList();

            var clusters = new List<List<LineSegment>>();
            if (positions.Count == 0) return clusters;

            var currentCluster = new List<LineSegment> { positions[0].line };
            double lastPos = positions[0].pos;

            for (int i = 1; i < positions.Count; i++)
            {
                if (positions[i].pos - lastPos > maxGap)
                {
                    if (currentCluster.Count >= 2)
                        clusters.Add(currentCluster);
                    currentCluster = new List<LineSegment>();
                }
                currentCluster.Add(positions[i].line);
                lastPos = positions[i].pos;
            }

            if (currentCluster.Count >= 2)
                clusters.Add(currentCluster);

            return clusters;
        }

        /// <summary>
        /// Removes overlapping grid candidates, keeping the ones with more cells.
        /// </summary>
        private List<GridCandidate> RemoveOverlapping(List<GridCandidate> grids)
        {
            if (grids.Count <= 1) return grids;

            var result = new List<GridCandidate>();

            // Sort by cell count (descending) to prefer larger grids
            grids.Sort((a, b) =>
                (b.XPositions.Count * b.YPositions.Count)
                .CompareTo(a.XPositions.Count * a.YPositions.Count));

            var usedBounds = new List<Rect>();
            foreach (var grid in grids)
            {
                var bounds = new Rect(
                    grid.XPositions.First(), grid.YPositions.First(),
                    grid.XPositions.Last(), grid.YPositions.Last());

                bool overlaps = usedBounds.Any(b => b.Intersects(bounds));
                if (!overlaps)
                {
                    result.Add(grid);
                    usedBounds.Add(bounds);
                }
            }

            return result;
        }

        /// <summary>
        /// Builds a complete DetectedTable from a validated grid candidate.
        /// Handles cell creation, merged cell detection, content population, and border styling.
        /// </summary>
        private DetectedTable BuildTableFromGrid(GridCandidate grid, PageContent content)
        {
            int rows = grid.YPositions.Count - 1;
            int cols = grid.XPositions.Count - 1;

            if (rows < 1 || cols < 1) return null;

            var table = new DetectedTable
            {
                RowCount = rows,
                ColCount = cols,
                Cells = new TableCell[rows, cols],
                ColumnWidths = new double[cols],
                RowHeights = new double[rows],
                Bounds = new Rect(
                    grid.XPositions.First(), grid.YPositions.First(),
                    grid.XPositions.Last(), grid.YPositions.Last())
            };

            // Calculate column widths and row heights
            for (int c = 0; c < cols; c++)
                table.ColumnWidths[c] = grid.XPositions[c + 1] - grid.XPositions[c];
            for (int r = 0; r < rows; r++)
                table.RowHeights[r] = grid.YPositions[r + 1] - grid.YPositions[r];

            // Initialize cells
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    table.Cells[r, c] = new TableCell
                    {
                        Row = r,
                        Col = c,
                        Bounds = new Rect(
                            grid.XPositions[c], grid.YPositions[r],
                            grid.XPositions[c + 1], grid.YPositions[r + 1])
                    };
                }
            }

            // Detect merged cells
            DetectMergedCells(table, grid);

            // Populate cells with text content
            PopulateCells(table, content.TextElements);

            // Apply border styles from the original line data
            ApplyBorderStyles(table, grid);

            // Apply cell shading from rectangles
            ApplyCellShading(table, content.Rectangles);

            return table;
        }

        /// <summary>
        /// Detects merged cells by checking for missing internal grid lines.
        /// A horizontal merge occurs when a vertical line is missing between adjacent cells.
        /// A vertical merge occurs when a horizontal line is missing between adjacent cells.
        /// </summary>
        private void DetectMergedCells(DetectedTable table, GridCandidate grid)
        {
            double snap = _options.TableGridSnapTolerance;

            // Build a boolean grid of internal lines
            // hLineExists[r, c] = true if there's a horizontal line between row r-1 and r at column c
            // vLineExists[r, c] = true if there's a vertical line between col c-1 and c at row r

            // Check horizontal merges: look for missing vertical lines between cells
            for (int r = 0; r < table.RowCount; r++)
            {
                for (int c = 0; c < table.ColCount - 1; c++)
                {
                    double x = grid.XPositions[c + 1]; // Vertical line between col c and c+1
                    double yTop = grid.YPositions[r];
                    double yBottom = grid.YPositions[r + 1];

                    bool hasVLine = HasLineSegment(grid.VerticalLines, x, yTop, x, yBottom, snap);

                    if (!hasVLine && !table.Cells[r, c].IsMergedContinuation)
                    {
                        // Merge cell [r, c] with [r, c+1]
                        table.Cells[r, c].ColSpan++;
                        table.Cells[r, c + 1].IsMergedContinuation = true;

                        // Update bounds
                        table.Cells[r, c].Bounds = new Rect(
                            table.Cells[r, c].Bounds.Left,
                            table.Cells[r, c].Bounds.Top,
                            table.Cells[r, c + 1].Bounds.Right,
                            table.Cells[r, c].Bounds.Bottom);
                    }
                }
            }

            // Check vertical merges: look for missing horizontal lines between cells
            for (int c = 0; c < table.ColCount; c++)
            {
                for (int r = 0; r < table.RowCount - 1; r++)
                {
                    if (table.Cells[r, c].IsMergedContinuation) continue;

                    double y = grid.YPositions[r + 1]; // Horizontal line between row r and r+1
                    double xLeft = grid.XPositions[c];
                    double xRight = grid.XPositions[c + table.Cells[r, c].ColSpan];

                    bool hasHLine = HasLineSegment(grid.HorizontalLines, xLeft, y, xRight, y, snap);

                    if (!hasHLine && !table.Cells[r + 1, c].IsMergedContinuation)
                    {
                        // Merge cell [r, c] with [r+1, c]
                        table.Cells[r, c].RowSpan++;
                        table.Cells[r + 1, c].IsMergedContinuation = true;

                        // Update bounds
                        table.Cells[r, c].Bounds = new Rect(
                            table.Cells[r, c].Bounds.Left,
                            table.Cells[r, c].Bounds.Top,
                            table.Cells[r, c].Bounds.Right,
                            table.Cells[r + 1, c].Bounds.Bottom);
                    }
                }
            }
        }

        /// <summary>
        /// Checks if there's a line segment at the specified position covering the given range.
        /// </summary>
        private bool HasLineSegment(List<LineSegment> lines, double x1, double y1, double x2, double y2, double tolerance)
        {
            bool isHorizontal = Math.Abs(y2 - y1) < tolerance;

            foreach (var line in lines)
            {
                if (isHorizontal && line.IsHorizontal)
                {
                    if (Math.Abs(line.Y1 - y1) <= tolerance)
                    {
                        double lineLeft = Math.Min(line.X1, line.X2);
                        double lineRight = Math.Max(line.X1, line.X2);
                        double targetLeft = Math.Min(x1, x2);
                        double targetRight = Math.Max(x1, x2);

                        // Line must cover at least 80% of the target span
                        double overlap = Math.Min(lineRight, targetRight) - Math.Max(lineLeft, targetLeft);
                        double targetSpan = targetRight - targetLeft;
                        if (targetSpan > 0 && overlap / targetSpan > 0.8)
                            return true;
                    }
                }
                else if (!isHorizontal && line.IsVertical)
                {
                    if (Math.Abs(line.X1 - x1) <= tolerance)
                    {
                        double lineTop = Math.Min(line.Y1, line.Y2);
                        double lineBottom = Math.Max(line.Y1, line.Y2);
                        double targetTop = Math.Min(y1, y2);
                        double targetBottom = Math.Max(y1, y2);

                        double overlap = Math.Min(lineBottom, targetBottom) - Math.Max(lineTop, targetTop);
                        double targetSpan = targetBottom - targetTop;
                        if (targetSpan > 0 && overlap / targetSpan > 0.8)
                            return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Populates each table cell with text content by finding text elements
        /// that fall within each cell's bounds.
        /// </summary>
        private void PopulateCells(DetectedTable table, List<TextElement> allTextElements)
        {
            int cellsWithText = 0;
            for (int r = 0; r < table.RowCount; r++)
            {
                for (int c = 0; c < table.ColCount; c++)
                {
                    var cell = table.Cells[r, c];
                    if (cell.IsMergedContinuation) continue;

                    // Find text elements within this cell
                    var cellElements = Layout.LayoutAnalyzer.ElementsInRegion(
                        allTextElements, cell.Bounds);

                    if (cellElements.Count > 0)
                    {
                        // Use the layout analyzer to group into paragraphs
                        cell.Paragraphs = _layoutAnalyzer.Analyze(
                            cellElements, cell.Bounds.Width, cell.Bounds.Height);
                        cellsWithText++;
                    }

                    if (cell.Paragraphs.Count == 0)
                    {
                        cell.Paragraphs.Add(new TextParagraph
                        {
                            Lines = new List<TextLine>
                            {
                                new TextLine { Runs = new List<TextRun> { new TextRun { Text = "" } } }
                            }
                        });
                    }
                }
            }

            if (_options.EnableDiagnostics)
                System.Console.WriteLine(
                    $"[TableDetector] Table populated: {cellsWithText}/{table.RowCount * table.ColCount} cells contain text");
        }

        /// <summary>
        /// Applies border styles to each cell by examining the original line segments
        /// at each cell boundary.
        /// </summary>
        private void ApplyBorderStyles(DetectedTable table, GridCandidate grid)
        {
            double snap = _options.TableGridSnapTolerance;

            for (int r = 0; r < table.RowCount; r++)
            {
                for (int c = 0; c < table.ColCount; c++)
                {
                    var cell = table.Cells[r, c];
                    if (cell.IsMergedContinuation) continue;

                    // Top border
                    cell.TopBorder = FindBorderStyle(grid.HorizontalLines,
                        cell.Bounds.Left, cell.Bounds.Top, cell.Bounds.Right, cell.Bounds.Top, snap);

                    // Bottom border
                    cell.BottomBorder = FindBorderStyle(grid.HorizontalLines,
                        cell.Bounds.Left, cell.Bounds.Bottom, cell.Bounds.Right, cell.Bounds.Bottom, snap);

                    // Left border
                    cell.LeftBorder = FindBorderStyle(grid.VerticalLines,
                        cell.Bounds.Left, cell.Bounds.Top, cell.Bounds.Left, cell.Bounds.Bottom, snap);

                    // Right border
                    cell.RightBorder = FindBorderStyle(grid.VerticalLines,
                        cell.Bounds.Right, cell.Bounds.Top, cell.Bounds.Right, cell.Bounds.Bottom, snap);
                }
            }
        }

        /// <summary>
        /// Finds the border style at a specific cell edge by matching line segments.
        /// </summary>
        private BorderStyle FindBorderStyle(List<LineSegment> lines,
            double x1, double y1, double x2, double y2, double tolerance)
        {
            bool isHorizontal = Math.Abs(y2 - y1) < tolerance;

            LineSegment bestMatch = null;
            double bestOverlap = 0;

            foreach (var line in lines)
            {
                if (isHorizontal && line.IsHorizontal && Math.Abs(line.Y1 - y1) <= tolerance)
                {
                    double overlap = Math.Min(Math.Max(line.X1, line.X2), x2) -
                                     Math.Max(Math.Min(line.X1, line.X2), x1);
                    if (overlap > bestOverlap)
                    {
                        bestOverlap = overlap;
                        bestMatch = line;
                    }
                }
                else if (!isHorizontal && line.IsVertical && Math.Abs(line.X1 - x1) <= tolerance)
                {
                    double overlap = Math.Min(Math.Max(line.Y1, line.Y2), y2) -
                                     Math.Max(Math.Min(line.Y1, line.Y2), y1);
                    if (overlap > bestOverlap)
                    {
                        bestOverlap = overlap;
                        bestMatch = line;
                    }
                }
            }

            if (bestMatch != null)
            {
                return new BorderStyle
                {
                    Width = bestMatch.Thickness,
                    Color = bestMatch.Color ?? "000000",
                    Style = bestMatch.Thickness > 0 ? "single" : "none"
                };
            }

            return BorderStyle.None;
        }

        /// <summary>
        /// Applies cell background shading from filled rectangles.
        /// </summary>
        private void ApplyCellShading(DetectedTable table, List<RectangleElement> rectangles)
        {
            if (rectangles == null || rectangles.Count == 0) return;

            for (int r = 0; r < table.RowCount; r++)
            {
                for (int c = 0; c < table.ColCount; c++)
                {
                    var cell = table.Cells[r, c];
                    if (cell.IsMergedContinuation) continue;

                    // Find filled rectangles that closely match this cell's bounds
                    foreach (var rect in rectangles)
                    {
                        if (!rect.IsFilled || string.IsNullOrEmpty(rect.FillColor))
                            continue;

                        // Check if rectangle covers most of the cell
                        if (rect.Bounds.Intersects(cell.Bounds))
                        {
                            double overlapLeft = Math.Max(rect.Bounds.Left, cell.Bounds.Left);
                            double overlapTop = Math.Max(rect.Bounds.Top, cell.Bounds.Top);
                            double overlapRight = Math.Min(rect.Bounds.Right, cell.Bounds.Right);
                            double overlapBottom = Math.Min(rect.Bounds.Bottom, cell.Bounds.Bottom);

                            double overlapArea = Math.Max(0, overlapRight - overlapLeft) *
                                                 Math.Max(0, overlapBottom - overlapTop);
                            double cellArea = cell.Bounds.Width * cell.Bounds.Height;

                            if (cellArea > 0 && overlapArea / cellArea > 0.7)
                            {
                                cell.BackgroundColor = rect.FillColor;
                                break;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Internal data structure for a grid candidate (set of aligned H and V lines).
        /// </summary>
        private class GridCandidate
        {
            public List<double> XPositions { get; set; } = new List<double>();
            public List<double> YPositions { get; set; } = new List<double>();
            public List<LineSegment> HorizontalLines { get; set; } = new List<LineSegment>();
            public List<LineSegment> VerticalLines { get; set; } = new List<LineSegment>();
        }
    }
}
