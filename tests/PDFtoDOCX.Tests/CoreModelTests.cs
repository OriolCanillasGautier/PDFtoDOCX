using PDFtoDOCX.Models;
using PDFtoDOCX.Utils;
using Xunit;

namespace PDFtoDOCX.Tests
{
    public class CoreModelTests
    {
        [Fact]
        public void Rect_Intersects_OverlappingRects_ReturnsTrue()
        {
            var a = new Rect(0, 0, 100, 100);
            var b = new Rect(50, 50, 150, 150);

            Assert.True(a.Intersects(b));
            Assert.True(b.Intersects(a));
        }

        [Fact]
        public void Rect_Intersects_NonOverlapping_ReturnsFalse()
        {
            var a = new Rect(0, 0, 100, 100);
            var b = new Rect(200, 200, 300, 300);

            Assert.False(a.Intersects(b));
        }

        [Fact]
        public void Rect_Contains_InnerRect_ReturnsTrue()
        {
            var outer = new Rect(0, 0, 200, 200);
            var inner = new Rect(50, 50, 150, 150);

            Assert.True(outer.Contains(inner));
            Assert.False(inner.Contains(outer));
        }

        [Fact]
        public void Rect_ContainsPoint_InsidePoint_ReturnsTrue()
        {
            var rect = new Rect(10, 10, 100, 100);

            Assert.True(rect.ContainsPoint(50, 50));
            Assert.True(rect.ContainsPoint(10, 10)); // Edge
            Assert.False(rect.ContainsPoint(5, 50)); // Outside
        }

        [Fact]
        public void Rect_Dimensions_CorrectValues()
        {
            var rect = new Rect(10, 20, 110, 70);

            Assert.Equal(100, rect.Width);
            Assert.Equal(50, rect.Height);
            Assert.Equal(60, rect.MidX);
            Assert.Equal(45, rect.MidY);
        }

        [Fact]
        public void TextRun_HasSameFormatting_IdenticalRuns_ReturnsTrue()
        {
            var a = new TextRun { FontName = "Arial", FontSize = 12, IsBold = true, IsItalic = false, Color = "000000" };
            var b = new TextRun { FontName = "Arial", FontSize = 12, IsBold = true, IsItalic = false, Color = "000000" };

            Assert.True(a.HasSameFormatting(b));
        }

        [Fact]
        public void TextRun_HasSameFormatting_DifferentBold_ReturnsFalse()
        {
            var a = new TextRun { FontName = "Arial", FontSize = 12, IsBold = true, Color = "000000" };
            var b = new TextRun { FontName = "Arial", FontSize = 12, IsBold = false, Color = "000000" };

            Assert.False(a.HasSameFormatting(b));
        }

        [Fact]
        public void LineSegment_IsHorizontal_HorizontalLine_ReturnsTrue()
        {
            var line = new LineSegment { X1 = 0, Y1 = 100, X2 = 200, Y2 = 100 };
            Assert.True(line.IsHorizontal);
            Assert.False(line.IsVertical);
        }

        [Fact]
        public void LineSegment_IsVertical_VerticalLine_ReturnsTrue()
        {
            var line = new LineSegment { X1 = 100, Y1 = 0, X2 = 100, Y2 = 200 };
            Assert.True(line.IsVertical);
            Assert.False(line.IsHorizontal);
        }

        [Fact]
        public void LineSegment_Normalize_ReversedHorizontal_FixesOrder()
        {
            var line = new LineSegment { X1 = 200, Y1 = 100, X2 = 100, Y2 = 100 };
            line.Normalize();
            Assert.True(line.X1 <= line.X2);
        }

        [Fact]
        public void UnitConverter_PointsToTwips_ConvertCorrectly()
        {
            Assert.Equal(1440, UnitConverter.PointsToTwips(72)); // 72pt = 1 inch = 1440 twips
        }

        [Fact]
        public void UnitConverter_FontSizeToHalfPoints_ConvertCorrectly()
        {
            Assert.Equal(24, UnitConverter.FontSizeToHalfPoints(12)); // 12pt = 24 half-points
        }

        [Fact]
        public void UnitConverter_PointsToEmu_ConvertCorrectly()
        {
            Assert.Equal(914400, UnitConverter.PointsToEmu(72)); // 72pt = 1 inch = 914400 EMU
        }

        [Fact]
        public void UnitConverter_RgbToHex_ConvertCorrectly()
        {
            Assert.Equal("FF0000", UnitConverter.RgbToHex(1.0, 0.0, 0.0));
            Assert.Equal("00FF00", UnitConverter.RgbToHex(0.0, 1.0, 0.0));
            Assert.Equal("0000FF", UnitConverter.RgbToHex(0.0, 0.0, 1.0));
            Assert.Equal("000000", UnitConverter.RgbToHex(0.0, 0.0, 0.0));
            Assert.Equal("FFFFFF", UnitConverter.RgbToHex(1.0, 1.0, 1.0));
        }

        [Fact]
        public void UnitConverter_GrayToHex_ConvertCorrectly()
        {
            Assert.Equal("000000", UnitConverter.GrayToHex(0.0));
            Assert.Equal("FFFFFF", UnitConverter.GrayToHex(1.0));
            Assert.Equal("808080", UnitConverter.GrayToHex(0.5));
        }
    }
}
