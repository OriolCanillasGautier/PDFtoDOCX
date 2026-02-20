using System.Collections.Generic;
using PDFtoDOCX.Models;
using PDFtoDOCX.Tables;
using Xunit;

namespace PDFtoDOCX.Tests
{
    public class TableDetectorTests
    {
        private readonly TableDetector _detector;

        public TableDetectorTests()
        {
            _detector = new TableDetector(new ConversionOptions());
        }

        [Fact]
        public void DetectTables_NoLines_ReturnsEmpty()
        {
            var content = new PageContent
            {
                PageNumber = 1,
                Width = 612,
                Height = 792,
                Lines = new List<LineSegment>()
            };

            var tables = _detector.DetectTables(content);
            Assert.Empty(tables);
        }

        [Fact]
        public void DetectTables_SimpleGrid_DetectsTable()
        {
            // Create a simple 2x2 table grid
            var lines = new List<LineSegment>
            {
                // Horizontal lines (3 rows: top, middle, bottom)
                new LineSegment { X1 = 100, Y1 = 100, X2 = 300, Y2 = 100, Thickness = 1 },
                new LineSegment { X1 = 100, Y1 = 150, X2 = 300, Y2 = 150, Thickness = 1 },
                new LineSegment { X1 = 100, Y1 = 200, X2 = 300, Y2 = 200, Thickness = 1 },
                
                // Vertical lines (3 columns: left, middle, right)
                new LineSegment { X1 = 100, Y1 = 100, X2 = 100, Y2 = 200, Thickness = 1 },
                new LineSegment { X1 = 200, Y1 = 100, X2 = 200, Y2 = 200, Thickness = 1 },
                new LineSegment { X1 = 300, Y1 = 100, X2 = 300, Y2 = 200, Thickness = 1 }
            };

            var content = new PageContent
            {
                PageNumber = 1,
                Width = 612,
                Height = 792,
                Lines = lines,
                TextElements = new List<TextElement>(),
                Rectangles = new List<RectangleElement>()
            };

            var tables = _detector.DetectTables(content);

            Assert.NotEmpty(tables);
            var table = tables[0];
            Assert.Equal(2, table.RowCount);
            Assert.Equal(2, table.ColCount);
        }

        [Fact]
        public void DetectTables_WithText_PopulatesCells()
        {
            // Create a 2x2 table with text inside cells
            var lines = new List<LineSegment>
            {
                // Horizontal lines (top, middle, bottom)
                new LineSegment { X1 = 100, Y1 = 100, X2 = 300, Y2 = 100, Thickness = 1 },
                new LineSegment { X1 = 100, Y1 = 150, X2 = 300, Y2 = 150, Thickness = 1 },
                new LineSegment { X1 = 100, Y1 = 200, X2 = 300, Y2 = 200, Thickness = 1 },
                // Vertical lines (left, center, right)
                new LineSegment { X1 = 100, Y1 = 100, X2 = 100, Y2 = 200, Thickness = 1 },
                new LineSegment { X1 = 200, Y1 = 100, X2 = 200, Y2 = 200, Thickness = 1 },
                new LineSegment { X1 = 300, Y1 = 100, X2 = 300, Y2 = 200, Thickness = 1 }
            };

            var textElements = new List<TextElement>
            {
                new TextElement { Text = "CellA", Bounds = new Rect(110, 115, 180, 130), FontSize = 12 },
                new TextElement { Text = "CellB", Bounds = new Rect(210, 115, 280, 130), FontSize = 12 },
                new TextElement { Text = "CellC", Bounds = new Rect(110, 165, 180, 180), FontSize = 12 },
                new TextElement { Text = "CellD", Bounds = new Rect(210, 165, 280, 180), FontSize = 12 }
            };

            var content = new PageContent
            {
                PageNumber = 1,
                Width = 612,
                Height = 792,
                Lines = lines,
                TextElements = textElements,
                Rectangles = new List<RectangleElement>()
            };

            var tables = _detector.DetectTables(content);

            Assert.NotEmpty(tables);
            var table = tables[0];
            Assert.Equal(2, table.RowCount);
            Assert.Equal(2, table.ColCount);

            // Check that cells have content
            Assert.NotEmpty(table.Cells[0, 0].Paragraphs);
            Assert.NotEmpty(table.Cells[0, 1].Paragraphs);
        }

        [Fact]
        public void DetectTables_InsufficientLines_ReturnsEmpty()
        {
            // Only one horizontal and one vertical line
            var lines = new List<LineSegment>
            {
                new LineSegment { X1 = 100, Y1 = 100, X2 = 300, Y2 = 100, Thickness = 1 },
                new LineSegment { X1 = 100, Y1 = 100, X2 = 100, Y2 = 200, Thickness = 1 }
            };

            var content = new PageContent
            {
                PageNumber = 1,
                Width = 612,
                Height = 792,
                Lines = lines
            };

            var tables = _detector.DetectTables(content);
            Assert.Empty(tables);
        }
    }
}
