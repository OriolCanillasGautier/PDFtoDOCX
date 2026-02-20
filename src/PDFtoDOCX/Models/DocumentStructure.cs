using System.Collections.Generic;

namespace PDFtoDOCX.Models
{
    /// <summary>
    /// All raw content extracted from a single PDF page.
    /// Coordinates use top-left origin (Y inverted from PDF standard).
    /// </summary>
    public class PageContent
    {
        public int PageNumber { get; set; }
        /// <summary>Page width in PDF points.</summary>
        public double Width { get; set; }
        /// <summary>Page height in PDF points.</summary>
        public double Height { get; set; }
        public List<TextElement> TextElements { get; set; } = new List<TextElement>();
        public List<ImageElement> Images { get; set; } = new List<ImageElement>();
        public List<LineSegment> Lines { get; set; } = new List<LineSegment>();
        public List<RectangleElement> Rectangles { get; set; } = new List<RectangleElement>();
        public List<HyperlinkInfo> Hyperlinks { get; set; } = new List<HyperlinkInfo>();
    }

    /// <summary>
    /// The analyzed logical structure of a single page.
    /// Contains content blocks in reading order.
    /// </summary>
    public class PageStructure
    {
        public int PageNumber { get; set; }
        /// <summary>Page width in PDF points.</summary>
        public double Width { get; set; }
        /// <summary>Page height in PDF points.</summary>
        public double Height { get; set; }
        /// <summary>Content blocks in reading order (paragraphs, tables, images).</summary>
        public List<ContentBlock> Blocks { get; set; } = new List<ContentBlock>();
    }

    /// <summary>
    /// The complete logical structure of a multi-page document.
    /// </summary>
    public class DocumentStructure
    {
        public List<PageStructure> Pages { get; set; } = new List<PageStructure>();
    }
}
