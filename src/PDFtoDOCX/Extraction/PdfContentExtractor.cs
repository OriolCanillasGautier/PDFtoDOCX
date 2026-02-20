using System;
using System.Collections.Generic;
using System.Linq;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;
using UglyToad.PdfPig.Core;
using UglyToad.PdfPig.Graphics.Colors;
using PDFtoDOCX.Models;
using PDFtoDOCX.Utils;

namespace PDFtoDOCX.Extraction
{
    /// <summary>
    /// Extracts all content elements from PDF pages using PdfPig.
    /// Converts PDF bottom-left origin coordinates to top-left origin.
    /// </summary>
    public class PdfContentExtractor
    {
        private readonly ConversionOptions _options;

        public PdfContentExtractor(ConversionOptions options)
        {
            _options = options ?? new ConversionOptions();
        }

        /// <summary>
        /// Extracts content from all pages (or a range) of the given PDF document.
        /// </summary>
        public List<PageContent> ExtractAll(PdfDocument document)
        {
            var pages = new List<PageContent>();
            int totalPages = document.NumberOfPages;
            int startPage = Math.Max(1, _options.StartPage);
            int endPage = _options.EndPage > 0 ? Math.Min(_options.EndPage, totalPages) : totalPages;

            if (_options.MaxPages > 0)
                endPage = Math.Min(endPage, startPage + _options.MaxPages - 1);

            for (int i = startPage; i <= endPage; i++)
            {
                var page = document.GetPage(i);
                var content = ExtractPage(page, i);
                pages.Add(content);
            }

            return pages;
        }

        /// <summary>
        /// Extracts all content from a single PDF page.
        /// </summary>
        public PageContent ExtractPage(Page page, int pageNumber)
        {
            double pageHeight = page.Height;
            double pageWidth = page.Width;

            var content = new PageContent
            {
                PageNumber = pageNumber,
                Width = pageWidth,
                Height = pageHeight
            };

            // Extract text elements (word-level)
            content.TextElements = ExtractTextElements(page, pageHeight);

            // Extract images
            if (_options.ExtractImages)
                content.Images = ExtractImages(page, pageHeight);

            // Extract lines and rectangles from paths (for table detection)
            ExtractPathElements(page, pageHeight, content);

            // Extract hyperlinks from annotations
            if (_options.DetectHyperlinks)
                content.Hyperlinks = ExtractHyperlinks(page, pageHeight);

            return content;
        }

        /// <summary>
        /// Extracts text elements at the word level with formatting info.
        /// Falls back to letter-based grouping when GetWords() returns nothing,
        /// and then optionally to OCR when letters are also unavailable.
        /// </summary>
        private List<TextElement> ExtractTextElements(Page page, double pageHeight)
        {
            var elements = new List<TextElement>();

            // Primary: PdfPig word extraction
            var words = page.GetWords();
            foreach (var word in words)
            {
                if (string.IsNullOrWhiteSpace(word.Text))
                    continue;

                var bounds = ConvertBounds(word.BoundingBox, pageHeight);
                var firstLetter = word.Letters.FirstOrDefault();
                string fontName = firstLetter?.FontName ?? "Unknown";
                double fontSize = firstLetter?.PointSize ?? 12.0;
                string color = ExtractLetterColor(firstLetter);

                elements.Add(new TextElement
                {
                    Text = word.Text,
                    Bounds = bounds,
                    FontName = CleanFontName(fontName),
                    FontSize = fontSize,
                    IsBold = DetectBold(fontName),
                    IsItalic = DetectItalic(fontName),
                    Color = color
                });
            }

            // Fallback 1: if GetWords() returned nothing, group individual Letters into pseudo-words
            if (elements.Count == 0)
            {
                var letters = page.Letters;
                if (letters != null && letters.Count > 0)
                {
                    if (_options.EnableDiagnostics)
                        System.Console.WriteLine(
                            $"[Extractor] Page words=0, using letter-grouping fallback ({letters.Count} letters)");
                    elements = GroupLettersIntoWords(letters, pageHeight);
                }
                else if (_options.EnableOcr)
                {
                    // Fallback 2: OCR (requires optional Tesseract integration)
                    if (_options.EnableDiagnostics)
                        System.Console.WriteLine("[Extractor] No words or letters — attempting OCR fallback");
                    elements = OcrExtractText(page, pageHeight);
                }
                else
                {
                    if (_options.EnableDiagnostics)
                        System.Console.WriteLine("[Extractor] No text operators found and OCR is disabled — page will be empty");
                }
            }

            return elements;
        }

        /// <summary>
        /// OCR fallback stub — returns an empty list unless a real OCR back-end is provided.
        /// Implement a subclass or inject a custom extractor to activate OCR.
        /// </summary>
        protected virtual List<TextElement> OcrExtractText(Page page, double pageHeight)
        {
            // Stub: real implementation requires Tesseract or similar OCR library.
            return new List<TextElement>();
        }

        /// <summary>
        /// Groups individual PdfPig Letters into word-level TextElements by proximity.
        /// Sorts by line (Y) then position (X), splitting words on gaps > GapMultiplier × char width.
        /// </summary>
        private List<TextElement> GroupLettersIntoWords(IReadOnlyList<Letter> letters, double pageHeight)
        {
            var elements = new List<TextElement>();

            // Sort letters: primary Y (baseline, PDF coords = high Y = top), secondary X
            var sorted = letters
                .Where(l => !string.IsNullOrWhiteSpace(l.Value))
                .OrderByDescending(l => l.GlyphRectangle.Bottom) // bottom-left origin: larger Y = higher on page
                .ThenBy(l => l.GlyphRectangle.Left)
                .ToList();

            if (sorted.Count == 0) return elements;

            // Group into lines by baseline proximity
            double lineTolerancePt = 3.0;
            var lines = new List<List<Letter>>();
            var currentLine = new List<Letter> { sorted[0] };
            double currentBaselineY = sorted[0].GlyphRectangle.Bottom;

            for (int i = 1; i < sorted.Count; i++)
            {
                double y = sorted[i].GlyphRectangle.Bottom;
                if (Math.Abs(y - currentBaselineY) <= lineTolerancePt)
                {
                    currentLine.Add(sorted[i]);
                }
                else
                {
                    lines.Add(currentLine);
                    currentLine = new List<Letter> { sorted[i] };
                    currentBaselineY = y;
                }
            }
            lines.Add(currentLine);

            // Within each line, split into words by horizontal gap
            foreach (var line in lines)
            {
                var byX = line.OrderBy(l => l.GlyphRectangle.Left).ToList();
                var wordLetters = new List<Letter> { byX[0] };

                for (int i = 1; i < byX.Count; i++)
                {
                    var prev = byX[i - 1];
                    var curr = byX[i];
                    double gap = curr.GlyphRectangle.Left - prev.GlyphRectangle.Right;
                    double avgWidth = prev.GlyphRectangle.Width;
                    // Split on gap > GapMultiplier × char width (inter-word gap, configurable)
                    if (gap > avgWidth * _options.LetterGroupingGapMultiplier)
                    {
                        if (wordLetters.Count > 0)
                            elements.Add(MakeWordElement(wordLetters, pageHeight));
                        wordLetters = new List<Letter>();
                    }
                    wordLetters.Add(curr);
                }
                if (wordLetters.Count > 0)
                    elements.Add(MakeWordElement(wordLetters, pageHeight));
            }

            return elements;
        }

        private TextElement MakeWordElement(List<Letter> letters, double pageHeight)
        {
            double minX = letters.Min(l => l.GlyphRectangle.Left);
            double maxX = letters.Max(l => l.GlyphRectangle.Right);
            double minY = letters.Min(l => l.GlyphRectangle.Bottom);
            double maxY = letters.Max(l => l.GlyphRectangle.Top);

            var text = string.Concat(letters.Select(l => l.Value));
            var first = letters[0];
            string fontName = first.FontName ?? "Unknown";

            return new TextElement
            {
                Text = text,
                Bounds = new Rect(minX, pageHeight - maxY, maxX, pageHeight - minY),
                FontName = CleanFontName(fontName),
                FontSize = first.PointSize,
                IsBold = DetectBold(fontName),
                IsItalic = DetectItalic(fontName),
                Color = ExtractLetterColor(first)
            };
        }

        /// <summary>
        /// Extracts images from the page, including nested images inside Form XObjects.
        /// </summary>
        private List<ImageElement> ExtractImages(Page page, double pageHeight)
        {
            var images = new List<ImageElement>();

            // Extract top-level images
            ExtractImagesFromSource(page.GetImages(), pageHeight, images);

            if (_options.EnableDiagnostics)
                System.Console.WriteLine(
                    $"[Extractor] Page {page.Number}: {images.Count} top-level image(s) extracted");

            // Phase 5.1: Form XObject image extraction is reserved for a future PdfPig API version.
            // The current build does not expose GetFormXObjects() on ExperimentalAccess.
            // When the API becomes available, recurse into each XObject and call ExtractImagesFromSource.

            return images;
        }

        /// <summary>
        /// Processes a sequence of PdfPig IPdfImage objects and populates the images list.
        /// </summary>
        private void ExtractImagesFromSource(
            IEnumerable<UglyToad.PdfPig.Content.IPdfImage> source,
            double pageHeight, List<ImageElement> images)
        {
            if (source == null) return;

            foreach (var image in source)
            {
                try
                {
                    var bounds = ConvertBounds(image.Bounds, pageHeight);
                    byte[] imageData = null;
                    string format = "png";

                    // Skip degenerate images with no visible area
                    if (bounds.Width < 1.0 || bounds.Height < 1.0)
                        continue;

                    // Phase 5.3: enforce minimum pixel size
                    if (image.Bounds.Width < 10 || image.Bounds.Height < 10)
                        continue;

                    // Try to get the image as PNG first
                    if (image.TryGetPng(out var pngBytes) && pngBytes != null && pngBytes.Length > 0)
                    {
                        imageData = pngBytes;
                        format = "png";
                    }
                    else
                    {
                        // Fall back to raw bytes
                        imageData = image.RawBytes.ToArray();
                        format = DetermineImageFormat(imageData);
                    }

                    // Normalize "jpeg" to "jpg" for broader Word compatibility
                    if (format == "jpeg") format = "jpg";

                    if (imageData != null && imageData.Length > 100)
                    {
                        images.Add(new ImageElement
                        {
                            Bounds = bounds,
                            ImageData = imageData,
                            Format = format,
                            WidthPx = image.Bounds.Width,
                            HeightPx = image.Bounds.Height
                        });
                    }
                }
                catch (Exception ex)
                {
                    if (_options.EnableDiagnostics)
                        System.Console.WriteLine($"[Extractor] Image extraction failed: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Extracts line segments and rectangles from PDF vector paths.
        /// These are used for table detection.
        /// </summary>
        private void ExtractPathElements(Page page, double pageHeight, PageContent content)
        {
            double pageWidth = page.Width;
            try
            {
                var paths = page.ExperimentalAccess.Paths;
                if (paths == null) return;

                foreach (var path in paths)
                {
                    bool isStroked = path.IsStroked;
                    bool isFilled = path.IsFilled;
                    string strokeColor = ColorToHex(path.StrokeColor);
                    string fillColor = ColorToHex(path.FillColor);
                    double lineWidth = (double)path.LineWidth;

                    // Process each subpath
                    foreach (var subpath in path)
                    {
                        var commands = subpath.Commands;
                        if (commands == null || commands.Count == 0) continue;

                        var points = new List<PdfPoint>();

                        foreach (var cmd in commands)
                        {
                            if (cmd is PdfSubpath.Move move)
                            {
                                points.Add(move.Location);
                            }
                            else if (cmd is PdfSubpath.Line line)
                            {
                                points.Add(line.From);
                                points.Add(line.To);

                                if (isStroked)
                                {
                                    var seg = new LineSegment
                                    {
                                        X1 = line.From.X,
                                        Y1 = pageHeight - line.From.Y,
                                        X2 = line.To.X,
                                        Y2 = pageHeight - line.To.Y,
                                        Thickness = lineWidth,
                                        Color = strokeColor
                                    };
                                    seg.Normalize();

                                    if (seg.Length >= _options.MinTableLineLength &&
                                        (seg.IsHorizontal || seg.IsVertical))
                                    {
                                        content.Lines.Add(seg);
                                    }
                                }
                            }
                            else if (cmd is PdfSubpath.Close)
                            {
                                // Closed path - may be a rectangle
                            }
                        }

                        // Only treat as a rectangle if the path has exactly 2 unique X values
                        // and 2 unique Y values — the fingerprint of an axis-aligned rectangle.
                        // Text glyph outlines have many unique X/Y values and are excluded.
                        if (points.Count >= 4)
                        {
                            var uniqueX = points.Select(p => Math.Round(p.X, 1)).Distinct().ToList();
                            var uniqueY = points.Select(p => Math.Round(p.Y, 1)).Distinct().ToList();
                            if (uniqueX.Count == 2 && uniqueY.Count == 2)
                            {
                                TryExtractRectangle(points, pageHeight, pageWidth, isFilled, isStroked,
                                    fillColor, strokeColor, lineWidth, content);
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Path extraction is experimental; gracefully handle failures
            }
        }

        /// <summary>
        /// Tries to interpret a set of points as a rectangle and adds it to content.
        /// Phase 1.2: Skips section banners (&gt;60% page width), background panels (&gt;40% page height),
        /// and rectangles with extreme aspect ratios.
        /// </summary>
        private void TryExtractRectangle(List<PdfPoint> points, double pageHeight, double pageWidth,
            bool isFilled, bool isStroked, string fillColor, string strokeColor,
            double lineWidth, PageContent content)
        {
            // Get bounding box of all points
            double minX = points.Min(p => p.X);
            double maxX = points.Max(p => p.X);
            double minY = points.Min(p => p.Y);
            double maxY = points.Max(p => p.Y);

            double width  = maxX - minX;
            double height = maxY - minY;

            // Only consider as a rectangle if it has meaningful area
            if (width < 2 || height < 2) return;

            // Phase 1.2: Filter decorative / page-spanning rectangles
            if (pageWidth > 0 && width > pageWidth * 0.60 &&
                pageHeight > 0 && height > pageHeight * 0.40)
                return;  // full-page banners / background fills

            // Aspect-ratio guard: keep only shapes between 0.2 and 5.0
            double aspectRatio = height > 0 ? width / height : double.MaxValue;
            if (aspectRatio > 5.0 || aspectRatio < 0.2)
            {
                // Allow only thin stroked lines (they become line segments below)
                bool isThin = height <= 5.0 || width <= 5.0;
                if (!isThin) return;
            }

            var bounds = new Rect(minX, pageHeight - maxY, maxX, pageHeight - minY);

            if (isFilled || isStroked)
            {
                content.Rectangles.Add(new RectangleElement
                {
                    Bounds = bounds,
                    FillColor = isFilled ? fillColor : string.Empty,
                    StrokeColor = isStroked ? strokeColor : string.Empty,
                    StrokeWidth = lineWidth,
                    IsFilled = isFilled,
                    IsStroked = isStroked
                });
            }

            // Derive line segments for table border detection.
            // Many PDFs draw borders as thin FILLED rectangles rather than stroked lines.
            string lineColor = isStroked ? strokeColor : fillColor;
            double lineThickness = lineWidth > 0 ? lineWidth : 1.0;

            // Case 1: Thin filled rect → treat as a single line segment
            const double ThinThreshold = 5.0;
            if (isFilled && !isStroked)
            {
                if (height <= ThinThreshold && width >= _options.MinTableLineLength)
                {
                    // Horizontal line: use the vertical midpoint
                    content.Lines.Add(new LineSegment
                    {
                        X1 = bounds.Left, Y1 = bounds.MidY,
                        X2 = bounds.Right, Y2 = bounds.MidY,
                        Thickness = height, Color = lineColor
                    });
                    return;
                }
                if (width <= ThinThreshold && height >= _options.MinTableLineLength)
                {
                    // Vertical line: use the horizontal midpoint
                    content.Lines.Add(new LineSegment
                    {
                        X1 = bounds.MidX, Y1 = bounds.Top,
                        X2 = bounds.MidX, Y2 = bounds.Bottom,
                        Thickness = width, Color = lineColor
                    });
                    return;
                }
            }

            // Case 2: Normal stroked rectangle → add its 4 edges as line segments
            // Case 3: Normal filled rectangle whose size suggests a table cell area → add edges
            // The aspect-ratio and page-size filters above already rejected decorative shapes.
            bool addEdges = isStroked ||
                (isFilled && width < 400 && height < 400);
            if (addEdges)
            {
                // Top edge
                content.Lines.Add(new LineSegment
                {
                    X1 = bounds.Left, Y1 = bounds.Top,
                    X2 = bounds.Right, Y2 = bounds.Top,
                    Thickness = lineThickness, Color = lineColor
                });
                // Bottom edge
                content.Lines.Add(new LineSegment
                {
                    X1 = bounds.Left, Y1 = bounds.Bottom,
                    X2 = bounds.Right, Y2 = bounds.Bottom,
                    Thickness = lineThickness, Color = lineColor
                });
                // Left edge
                content.Lines.Add(new LineSegment
                {
                    X1 = bounds.Left, Y1 = bounds.Top,
                    X2 = bounds.Left, Y2 = bounds.Bottom,
                    Thickness = lineThickness, Color = lineColor
                });
                // Right edge
                content.Lines.Add(new LineSegment
                {
                    X1 = bounds.Right, Y1 = bounds.Top,
                    X2 = bounds.Right, Y2 = bounds.Bottom,
                    Thickness = lineThickness, Color = lineColor
                });
            }
        }

        /// <summary>
        /// Extracts hyperlink annotations from the page.
        /// </summary>
        private List<HyperlinkInfo> ExtractHyperlinks(Page page, double pageHeight)
        {
            var links = new List<HyperlinkInfo>();

            try
            {
                var annotations = page.ExperimentalAccess.GetAnnotations();
                if (annotations == null) return links;

                foreach (var annotation in annotations)
                {
                    // Check if it's a link annotation
                    if (annotation.Type == UglyToad.PdfPig.Annotations.AnnotationType.Link)
                    {
                        var dict = annotation.AnnotationDictionary;
                        // Try to get the URI from the action dictionary
                        string uri = null;

                        if (dict.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create("A"), out var actionToken))
                        {
                            if (actionToken is UglyToad.PdfPig.Tokens.DictionaryToken actionDict)
                            {
                                if (actionDict.TryGet(UglyToad.PdfPig.Tokens.NameToken.Create("URI"), out var uriToken))
                                {
                                    if (uriToken is UglyToad.PdfPig.Tokens.StringToken strToken)
                                    {
                                        uri = strToken.Data;
                                    }
                                }
                            }
                        }

                        if (!string.IsNullOrEmpty(uri))
                        {
                            var rect = annotation.Rectangle;
                            links.Add(new HyperlinkInfo
                            {
                                Bounds = ConvertBounds(rect, pageHeight),
                                Uri = uri
                            });
                        }
                    }
                }
            }
            catch (Exception)
            {
                // Annotation extraction may fail on some PDFs
            }

            return links;
        }

        // --- Helper Methods ---

        /// <summary>
        /// Converts PdfPig's bottom-left origin rectangle to top-left origin.
        /// </summary>
        private Rect ConvertBounds(PdfRectangle pdfRect, double pageHeight)
        {
            return new Rect(
                left: pdfRect.Left,
                top: pageHeight - pdfRect.Top,
                right: pdfRect.Right,
                bottom: pageHeight - pdfRect.Bottom
            );
        }

        /// <summary>
        /// Extracts color from a PdfPig Letter as a hex string.
        /// </summary>
        private string ExtractLetterColor(Letter letter)
        {
            if (letter?.Color == null) return "000000";
            return ColorToHex(letter.Color);
        }

        /// <summary>
        /// Converts a PdfPig IColor to a hex RGB string.
        /// </summary>
        private string ColorToHex(IColor color)
        {
            if (color == null) return "000000";

            if (color is RGBColor rgb)
            {
                return UnitConverter.RgbToHex((double)rgb.R, (double)rgb.G, (double)rgb.B);
            }
            else if (color is GrayColor gray)
            {
                return UnitConverter.GrayToHex((double)gray.Gray);
            }
            else if (color is CMYKColor cmyk)
            {
                // Approximate CMYK to RGB conversion
                double c = (double)cmyk.C, m = (double)cmyk.M, y = (double)cmyk.Y, k = (double)cmyk.K;
                double r = (1 - c) * (1 - k);
                double g = (1 - m) * (1 - k);
                double b = (1 - y) * (1 - k);
                return UnitConverter.RgbToHex(r, g, b);
            }

            return "000000";
        }

        /// <summary>
        /// Detects bold from font name patterns (e.g., "Arial-Bold", "TimesNewRoman,Bold").
        /// </summary>
        private bool DetectBold(string fontName)
        {
            if (string.IsNullOrEmpty(fontName)) return false;
            string fn = fontName.ToLowerInvariant();
            return fn.Contains("bold") || fn.Contains("heavy") || fn.Contains("black");
        }

        /// <summary>
        /// Detects italic from font name patterns (e.g., "Arial-Italic", "TimesNewRoman-BoldItalic").
        /// </summary>
        private bool DetectItalic(string fontName)
        {
            if (string.IsNullOrEmpty(fontName)) return false;
            string fn = fontName.ToLowerInvariant();
            return fn.Contains("italic") || fn.Contains("oblique");
        }

        /// <summary>
        /// Cleans font name by removing subset prefix (e.g., "BCDFGH+Arial" -> "Arial").
        /// </summary>
        private string CleanFontName(string fontName)
        {
            if (string.IsNullOrEmpty(fontName)) return "Arial";

            // Remove subset prefix (6 uppercase letters + '+')
            int plusIndex = fontName.IndexOf('+');
            if (plusIndex >= 0 && plusIndex <= 7)
            {
                fontName = fontName.Substring(plusIndex + 1);
            }

            // Remove style suffixes for base font name
            string baseName = fontName
                .Replace("-Bold", "")
                .Replace("-Italic", "")
                .Replace("-BoldItalic", "")
                .Replace(",Bold", "")
                .Replace(",Italic", "")
                .Replace(",BoldItalic", "")
                .Replace("-Regular", "")
                .Replace(",Regular", "");

            // Map common PDF font names to Windows font names
            return MapToWindowsFont(baseName);
        }

        /// <summary>
        /// Maps common PDF font names to standard Windows/DOCX fonts.
        /// </summary>
        private string MapToWindowsFont(string fontName)
        {
            string lower = fontName.ToLowerInvariant().Replace(" ", "");

            if (lower.Contains("timesnewroman") || lower.Contains("times"))
                return "Times New Roman";
            if (lower.Contains("arial") || lower.Contains("helvetica"))
                return "Arial";
            if (lower.Contains("couriernew") || lower.Contains("courier"))
                return "Courier New";
            if (lower.Contains("calibri"))
                return "Calibri";
            if (lower.Contains("cambria"))
                return "Cambria";
            if (lower.Contains("georgia"))
                return "Georgia";
            if (lower.Contains("verdana"))
                return "Verdana";
            if (lower.Contains("tahoma"))
                return "Tahoma";
            if (lower.Contains("trebuchet"))
                return "Trebuchet MS";
            if (lower.Contains("symbol"))
                return "Symbol";
            if (lower.Contains("wingdings"))
                return "Wingdings";

            // Return the original name if no mapping found
            return fontName;
        }

        /// <summary>
        /// Determines image format from raw bytes by checking magic numbers.
        /// </summary>
        private string DetermineImageFormat(byte[] data)
        {
            if (data == null || data.Length < 4) return "png";

            // PNG: 0x89 0x50 0x4E 0x47
            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                return "png";

            // JPEG: 0xFF 0xD8
            if (data[0] == 0xFF && data[1] == 0xD8)
                return "jpg";

            // GIF: "GIF8"
            if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x38)
                return "gif";

            // BMP: "BM"
            if (data[0] == 0x42 && data[1] == 0x4D)
                return "bmp";

            // TIFF: "II" or "MM"
            if ((data[0] == 0x49 && data[1] == 0x49) || (data[0] == 0x4D && data[1] == 0x4D))
                return "tiff";

            return "png"; // Default to PNG
        }
    }
}
