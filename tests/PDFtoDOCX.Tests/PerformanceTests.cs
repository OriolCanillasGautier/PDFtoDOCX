using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using PDFtoDOCX.Docx;
using PDFtoDOCX.Layout;
using PDFtoDOCX.Models;
using PDFtoDOCX.Tables;
using Xunit;
using Xunit.Abstractions;

namespace PDFtoDOCX.Tests
{
    /// <summary>
    /// Performance benchmarks for the core conversion pipeline components.
    /// These tests measure elapsed time for in-memory operations and assert
    /// that processing stays within acceptable bounds on a standard developer machine.
    ///
    /// Target: &lt;2 000 ms total per simulated 10-page document for each subsystem.
    /// </summary>
    public class PerformanceTests
    {
        private readonly ITestOutputHelper _output;

        public PerformanceTests(ITestOutputHelper output)
        {
            _output = output;
        }

        // ── Layout analyzer ───────────────────────────────────────────────────

        [Fact]
        public void LayoutAnalyzer_1000Elements_CompletesIn500ms()
        {
            var options  = new ConversionOptions();
            var analyzer = new LayoutAnalyzer(options);

            // Build 1 000 text elements spread across 40 rows, 25 columns
            var elements = new List<TextElement>();
            for (int row = 0; row < 40; row++)
            {
                for (int col = 0; col < 25; col++)
                {
                    double x = 72 + col * 18.0;
                    double y = 72 + row * 14.0;
                    elements.Add(new TextElement
                    {
                        Text     = "Word",
                        Bounds   = new Rect(x, y, x + 16, y + 12),
                        FontName = "Arial",
                        FontSize = 10
                    });
                }
            }

            var sw = Stopwatch.StartNew();
            var paragraphs = analyzer.Analyze(elements, 612, 792);
            sw.Stop();

            _output.WriteLine($"LayoutAnalyzer — 1000 elements: {sw.ElapsedMilliseconds} ms, " +
                              $"{paragraphs.Count} paragraph(s)");

            Assert.True(sw.ElapsedMilliseconds < 500,
                $"LayoutAnalyzer took {sw.ElapsedMilliseconds} ms (limit: 500 ms)");
        }

        // ── Table detector ────────────────────────────────────────────────────

        [Fact]
        public void TableDetector_5x10Grid_CompletesIn200ms()
        {
            var detector = new TableDetector(new ConversionOptions());
            var lines    = BuildGridLines(100, 100, colCount: 5, rowCount: 10,
                                          cellWidth: 80, cellHeight: 30);

            var content = new PageContent
            {
                PageNumber   = 1, Width = 612, Height = 792,
                Lines        = lines,
                TextElements = new List<TextElement>(),
                Rectangles   = new List<RectangleElement>()
            };

            var sw = Stopwatch.StartNew();
            var tables = detector.DetectTables(content);
            sw.Stop();

            _output.WriteLine($"TableDetector — 5×10 grid: {sw.ElapsedMilliseconds} ms, " +
                              $"{tables.Count} table(s)");

            Assert.True(sw.ElapsedMilliseconds < 200,
                $"TableDetector took {sw.ElapsedMilliseconds} ms (limit: 200 ms)");
        }

        // ── DocxPackager ─────────────────────────────────────────────────────

        [Fact]
        public void DocxPackager_100Paragraphs_CompletesIn1000ms()
        {
            var packager = new DocxPackager();
            var blocks   = Enumerable.Range(0, 100).Select(i =>
            {
                double y = 72 + i * 14.0;
                var line = new TextLine
                {
                    Bounds = new Rect(72, y, 540, y + 12),
                    Runs   = new List<TextRun>
                    {
                        new TextRun { Text = $"Paragraph {i}: Lorem ipsum dolor sit amet.", FontName = "Calibri", FontSize = 11 }
                    }
                };
                var para = new TextParagraph
                {
                    Lines  = new List<TextLine> { line },
                    Bounds = line.Bounds
                };
                return new ContentBlock
                {
                    Type      = ContentBlockType.Paragraph,
                    Paragraph = para,
                    Bounds    = para.Bounds
                };
            }).ToList();

            var doc = new DocumentStructure
            {
                Pages = new List<PageStructure>
                {
                    new PageStructure { PageNumber = 1, Width = 612, Height = 792, Blocks = blocks }
                }
            };

            var sw = Stopwatch.StartNew();
            var bytes = packager.Generate(doc);
            sw.Stop();

            _output.WriteLine($"DocxPackager — 100 paragraphs: {sw.ElapsedMilliseconds} ms, " +
                              $"{bytes.Length:N0} bytes");

            Assert.True(sw.ElapsedMilliseconds < 1000,
                $"DocxPackager took {sw.ElapsedMilliseconds} ms (limit: 1000 ms)");
            Assert.True(bytes.Length > 0);
        }

        // ── Memory: large table ───────────────────────────────────────────────

        [Fact]
        public void DocxPackager_LargeTable_CompletesIn2000ms()
        {
            const int rows = 50, cols = 10;
            var table = new DetectedTable
            {
                Bounds       = new Rect(72, 72, 540, 72 + rows * 18.0),
                RowCount     = rows,
                ColCount     = cols,
                ColumnWidths = Enumerable.Repeat(46.8, cols).ToArray(),
                RowHeights   = Enumerable.Repeat(18.0, rows).ToArray(),
                Cells        = new TableCell[rows, cols]
            };

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    table.Cells[r, c] = new TableCell
                    {
                        Row = r, Col = c,
                        Bounds = new Rect(72 + c * 46.8, 72 + r * 18.0,
                                          72 + (c + 1) * 46.8, 72 + (r + 1) * 18.0),
                        Paragraphs = new List<TextParagraph>
                        {
                            new TextParagraph
                            {
                                Lines = new List<TextLine>
                                {
                                    new TextLine
                                    {
                                        Runs = new List<TextRun>
                                        {
                                            new TextRun { Text = $"{r},{c}", FontSize = 9 }
                                        }
                                    }
                                }
                            }
                        }
                    };
                }
            }

            var packager = new DocxPackager();
            var doc = new DocumentStructure
            {
                Pages = new List<PageStructure>
                {
                    new PageStructure
                    {
                        PageNumber = 1, Width = 612, Height = 792,
                        Blocks = new List<ContentBlock>
                        {
                            new ContentBlock
                            {
                                Type   = ContentBlockType.Table,
                                Table  = table,
                                Bounds = table.Bounds
                            }
                        }
                    }
                }
            };

            var sw = Stopwatch.StartNew();
            var bytes = packager.Generate(doc);
            sw.Stop();

            _output.WriteLine($"DocxPackager — {rows}×{cols} table: {sw.ElapsedMilliseconds} ms, " +
                              $"{bytes.Length:N0} bytes");

            Assert.True(sw.ElapsedMilliseconds < 2000,
                $"DocxPackager (large table) took {sw.ElapsedMilliseconds} ms (limit: 2000 ms)");
        }

        // ── Helper ────────────────────────────────────────────────────────────

        private static List<LineSegment> BuildGridLines(
            double startX, double startY,
            int colCount, int rowCount,
            double cellWidth, double cellHeight)
        {
            var lines = new List<LineSegment>();
            double endX = startX + colCount * cellWidth;
            double endY = startY + rowCount * cellHeight;

            for (int r = 0; r <= rowCount; r++)
            {
                double y = startY + r * cellHeight;
                lines.Add(new LineSegment { X1 = startX, Y1 = y, X2 = endX, Y2 = y });
            }

            for (int c = 0; c <= colCount; c++)
            {
                double x = startX + c * cellWidth;
                lines.Add(new LineSegment { X1 = x, Y1 = startY, X2 = x, Y2 = endY });
            }

            return lines;
        }
    }
}
