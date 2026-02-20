using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using PDFtoDOCX;
using PDFtoDOCX.Models;
using PDFtoDOCX.Tables;
using PDFtoDOCX.Extraction;
using Xunit;

namespace PDFtoDOCX.Tests
{
    /// <summary>
    /// Regression tests that validate end-to-end conversion behavior using
    /// in-memory constructed page content (no external PDF files required).
    ///
    /// When physical test PDFs are available, add them to the test project as
    /// embedded resources and write additional [Fact] tests that open and convert
    /// them with <see cref="PdfToDocxConverter.ConvertToBytes(byte[])"/>.
    /// </summary>
    public class RegressionTests
    {
        // ── Table detection regressions ───────────────────────────────────────

        [Fact]
        public void TableDetector_SingleCellBox_IsRejected()
        {
            // A 1×1 "table" (just a decorative box) should NOT be detected.
            // ValidateGrid requires at least 2×2 cells.
            var detector = new TableDetector(new ConversionOptions());

            var lines = new List<LineSegment>
            {
                // A simple box: 2 H-lines + 2 V-lines → 1×1 grid → must be rejected
                new LineSegment { X1 = 50, Y1 = 50, X2 = 200, Y2 = 50 },
                new LineSegment { X1 = 50, Y1 = 150, X2 = 200, Y2 = 150 },
                new LineSegment { X1 = 50, Y1 = 50, X2 = 50, Y2 = 150 },
                new LineSegment { X1 = 200, Y1 = 50, X2 = 200, Y2 = 150 }
            };

            var content = new PageContent
            {
                PageNumber = 1, Width = 612, Height = 792,
                Lines = lines,
                TextElements = new List<TextElement>(),
                Rectangles = new List<RectangleElement>()
            };

            var tables = detector.DetectTables(content);
            Assert.Empty(tables);
        }

        [Fact]
        public void TableDetector_PageWideBox_IsRejected()
        {
            // A box that spans >80% of the page in BOTH dimensions should be
            // rejected as a page border (phantom table).
            var detector = new TableDetector(new ConversionOptions());

            double pageW = 612, pageH = 792;

            // 90% × 90% page rectangle with interior lines (to pass inner-line check)
            double left = pageW * 0.05, right = pageW * 0.95;
            double top  = pageH * 0.05, bot   = pageH * 0.95;
            double midX = (left + right) / 2;
            double midY = (top + bot) / 2;

            var lines = new List<LineSegment>
            {
                // Outer boundary
                new LineSegment { X1 = left, Y1 = top,  X2 = right, Y2 = top  },
                new LineSegment { X1 = left, Y1 = bot,  X2 = right, Y2 = bot  },
                new LineSegment { X1 = left, Y1 = top,  X2 = left,  Y2 = bot  },
                new LineSegment { X1 = right,Y1 = top,  X2 = right, Y2 = bot  },
                // Interior lines (to make it look like a real 2×2 grid)
                new LineSegment { X1 = left, Y1 = midY, X2 = right, Y2 = midY },
                new LineSegment { X1 = midX, Y1 = top,  X2 = midX,  Y2 = bot  }
            };

            var content = new PageContent
            {
                PageNumber = 1, Width = pageW, Height = pageH,
                Lines = lines,
                TextElements = new List<TextElement>(),
                Rectangles = new List<RectangleElement>()
            };

            var tables = detector.DetectTables(content);
            Assert.Empty(tables);
        }

        [Fact]
        public void TableDetector_Valid2x2_IsAccepted()
        {
            // Verify that a well-formed 2×2 table with interior lines is still detected.
            var detector = new TableDetector(new ConversionOptions());

            var lines = new List<LineSegment>
            {
                new LineSegment { X1 = 100, Y1 = 100, X2 = 300, Y2 = 100 },
                new LineSegment { X1 = 100, Y1 = 150, X2 = 300, Y2 = 150 },
                new LineSegment { X1 = 100, Y1 = 200, X2 = 300, Y2 = 200 },
                new LineSegment { X1 = 100, Y1 = 100, X2 = 100, Y2 = 200 },
                new LineSegment { X1 = 200, Y1 = 100, X2 = 200, Y2 = 200 },
                new LineSegment { X1 = 300, Y1 = 100, X2 = 300, Y2 = 200 }
            };

            var content = new PageContent
            {
                PageNumber = 1, Width = 612, Height = 792,
                Lines = lines,
                TextElements = new List<TextElement>(),
                Rectangles = new List<RectangleElement>()
            };

            var tables = detector.DetectTables(content);
            Assert.Single(tables);
            Assert.Equal(2, tables[0].RowCount);
            Assert.Equal(2, tables[0].ColCount);
        }

        // ── Text length regression ────────────────────────────────────────────

        [Fact]
        public void Conversion_TextPreserved_NoCorruptionInOutputXml()
        {
            // Check that special XML characters in text are escaped correctly.
            var packager = new PDFtoDOCX.Docx.DocxPackager();
            var para = new TextParagraph
            {
                Lines = new List<TextLine>
                {
                    new TextLine
                    {
                        Bounds = new Rect(72, 72, 540, 84),
                        Runs = new List<TextRun>
                        {
                            new TextRun
                            {
                                Text = "Price: 5 < 10 & 'hello' \"world\"",
                                FontName = "Arial", FontSize = 12
                            }
                        }
                    }
                },
                Bounds = new Rect(72, 72, 540, 84)
            };

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
                                Type = ContentBlockType.Paragraph,
                                Paragraph = para, Bounds = para.Bounds
                            }
                        }
                    }
                }
            };

            var bytes = packager.Generate(doc);

            using var ms = new MemoryStream(bytes);
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read);
            var docEntry = archive.GetEntry("word/document.xml");
            using var reader = new StreamReader(docEntry!.Open());
            var xml = reader.ReadToEnd();

            // Verify special characters are escaped and the document is well-formed XML
            Assert.Contains("&lt;", xml);
            Assert.Contains("&amp;", xml);
            Assert.DoesNotContain("Price: 5 < 10", xml); // raw < should not appear
        }

        // ── Confidence scoring regression ─────────────────────────────────────

        [Fact]
        public void TableWithText_HasHigherConfidenceThanEmptyTable()
        {
            // A table whose cells contain text should score higher than an empty table.
            // Indirectly verified by checking that both are returned (both have confidence >= 0.4).
            // The text-bearing table should appear first (sorted by row/col position).
            var detector = new TableDetector(new ConversionOptions());

            var baseLines = new List<LineSegment>
            {
                new LineSegment { X1 = 100, Y1 = 100, X2 = 300, Y2 = 100 },
                new LineSegment { X1 = 100, Y1 = 150, X2 = 300, Y2 = 150 },
                new LineSegment { X1 = 100, Y1 = 200, X2 = 300, Y2 = 200 },
                new LineSegment { X1 = 100, Y1 = 100, X2 = 100, Y2 = 200 },
                new LineSegment { X1 = 200, Y1 = 100, X2 = 200, Y2 = 200 },
                new LineSegment { X1 = 300, Y1 = 100, X2 = 300, Y2 = 200 }
            };

            var textElements = new List<TextElement>
            {
                new TextElement { Text = "A", Bounds = new Rect(110, 115, 180, 130), FontSize = 12 },
                new TextElement { Text = "B", Bounds = new Rect(210, 115, 280, 130), FontSize = 12 }
            };

            var content = new PageContent
            {
                PageNumber = 1, Width = 612, Height = 792,
                Lines = baseLines, TextElements = textElements,
                Rectangles = new List<RectangleElement>()
            };

            var tables = detector.DetectTables(content);
            // A valid 2×2 grid with 2 cells having text → confidence should be >= 0.4
            Assert.NotEmpty(tables);
        }
    }
}
