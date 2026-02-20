using System.Collections.Generic;
using PDFtoDOCX.Layout;
using PDFtoDOCX.Models;
using Xunit;

namespace PDFtoDOCX.Tests
{
    public class LayoutAnalyzerTests
    {
        private readonly LayoutAnalyzer _analyzer;

        public LayoutAnalyzerTests()
        {
            _analyzer = new LayoutAnalyzer(new ConversionOptions());
        }

        [Fact]
        public void GroupIntoLines_SingleElement_ReturnsOneLine()
        {
            var elements = new List<TextElement>
            {
                new TextElement { Text = "Hello", Bounds = new Rect(10, 10, 50, 22), FontSize = 12 }
            };

            var lines = _analyzer.GroupIntoLines(elements);

            Assert.Single(lines);
            Assert.Equal("Hello", lines[0].FullText);
        }

        [Fact]
        public void GroupIntoLines_TwoElementsSameLine_ReturnsSingleLine()
        {
            var elements = new List<TextElement>
            {
                new TextElement { Text = "Hello", Bounds = new Rect(10, 10, 50, 22), FontSize = 12.0 },
                new TextElement { Text = "World", Bounds = new Rect(55, 10, 95, 22), FontSize = 12.0 }
            };

            var lines = _analyzer.GroupIntoLines(elements);

            Assert.Single(lines);
            Assert.Contains("Hello", lines[0].FullText);
            Assert.Contains("World", lines[0].FullText);
        }

        [Fact]
        public void GroupIntoLines_TwoElementsDifferentLines_ReturnsTwoLines()
        {
            var elements = new List<TextElement>
            {
                new TextElement { Text = "Line1", Bounds = new Rect(10, 10, 50, 22), FontSize = 12.0 },
                new TextElement { Text = "Line2", Bounds = new Rect(10, 30, 50, 42), FontSize = 12.0 }
            };

            var lines = _analyzer.GroupIntoLines(elements);

            Assert.Equal(2, lines.Count);
            Assert.Equal("Line1", lines[0].FullText);
            Assert.Equal("Line2", lines[1].FullText);
        }

        [Fact]
        public void GroupIntoLines_ElementsOutOfOrder_SortsCorrectly()
        {
            var elements = new List<TextElement>
            {
                new TextElement { Text = "Second", Bounds = new Rect(60, 10, 110, 22), FontSize = 12.0 },
                new TextElement { Text = "First", Bounds = new Rect(10, 10, 50, 22), FontSize = 12.0 }
            };

            var lines = _analyzer.GroupIntoLines(elements);

            Assert.Single(lines);
            // First should come before Second in the line
            Assert.StartsWith("First", lines[0].FullText);
        }

        [Fact]
        public void GroupIntoParagraphs_SmallGap_SingleParagraph()
        {
            var lines = new List<TextLine>
            {
                new TextLine
                {
                    Bounds = new Rect(10, 10, 200, 22),
                    Runs = new List<TextRun> { new TextRun { Text = "Line 1", FontSize = 12 } }
                },
                new TextLine
                {
                    Bounds = new Rect(10, 24, 200, 36),
                    Runs = new List<TextRun> { new TextRun { Text = "Line 2", FontSize = 12 } }
                }
            };

            var paragraphs = _analyzer.GroupIntoParagraphs(lines);

            Assert.Single(paragraphs);
            Assert.Equal(2, paragraphs[0].Lines.Count);
        }

        [Fact]
        public void GroupIntoParagraphs_LargeGap_TwoParagraphs()
        {
            var lines = new List<TextLine>
            {
                new TextLine
                {
                    Bounds = new Rect(10, 10, 200, 22),
                    Runs = new List<TextRun> { new TextRun { Text = "Para 1", FontSize = 12 } }
                },
                new TextLine
                {
                    Bounds = new Rect(10, 50, 200, 62),
                    Runs = new List<TextRun> { new TextRun { Text = "Para 2", FontSize = 12 } }
                }
            };

            var paragraphs = _analyzer.GroupIntoParagraphs(lines);

            Assert.Equal(2, paragraphs.Count);
        }

        [Fact]
        public void DetectColumns_SingleColumn_ReturnsOneColumn()
        {
            var lines = new List<TextLine>
            {
                new TextLine { Bounds = new Rect(72, 10, 540, 22) },
                new TextLine { Bounds = new Rect(72, 30, 540, 42) },
                new TextLine { Bounds = new Rect(72, 50, 540, 62) }
            };

            var columns = _analyzer.DetectColumns(lines, 612);

            Assert.Single(columns);
        }

        [Fact]
        public void DetectColumns_TwoColumns_ReturnsTwoColumns()
        {
            var lines = new List<TextLine>
            {
                // Left column
                new TextLine { Bounds = new Rect(50, 10, 260, 22) },
                new TextLine { Bounds = new Rect(50, 30, 260, 42) },
                new TextLine { Bounds = new Rect(50, 50, 260, 62) },
                // Right column (with a clear gap)
                new TextLine { Bounds = new Rect(320, 10, 550, 22) },
                new TextLine { Bounds = new Rect(320, 30, 550, 42) },
                new TextLine { Bounds = new Rect(320, 50, 550, 62) }
            };

            var columns = _analyzer.DetectColumns(lines, 612);

            Assert.Equal(2, columns.Count);
        }

        [Fact]
        public void Analyze_EmptyInput_ReturnsEmptyList()
        {
            var result = _analyzer.Analyze(new List<TextElement>(), 612, 792);
            Assert.Empty(result);
        }

        [Fact]
        public void ExcludeTableRegions_FiltersCorrectly()
        {
            var elements = new List<TextElement>
            {
                new TextElement { Text = "Outside", Bounds = new Rect(10, 10, 50, 22) },
                new TextElement { Text = "Inside", Bounds = new Rect(110, 110, 150, 122) }
            };

            var tables = new List<DetectedTable>
            {
                new DetectedTable { Bounds = new Rect(100, 100, 300, 300) }
            };

            var filtered = LayoutAnalyzer.ExcludeTableRegions(elements, tables);

            Assert.Single(filtered);
            Assert.Equal("Outside", filtered[0].Text);
        }

        [Fact]
        public void ElementsInRegion_FindsCorrectElements()
        {
            var elements = new List<TextElement>
            {
                new TextElement { Text = "Outside", Bounds = new Rect(10, 10, 50, 22) },
                new TextElement { Text = "Inside", Bounds = new Rect(110, 110, 150, 122) }
            };

            var region = new Rect(100, 100, 300, 300);
            var found = LayoutAnalyzer.ElementsInRegion(elements, region);

            Assert.Single(found);
            Assert.Equal("Inside", found[0].Text);
        }
    }
}
