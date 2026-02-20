using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using PDFtoDOCX.Models;
using PDFtoDOCX.Utils;

namespace PDFtoDOCX.Docx
{
    /// <summary>
    /// Packages the analyzed document structure into a valid DOCX (OOXML) file.
    /// Constructs all required XML parts and assembles them into a ZIP archive.
    /// </summary>
    public class DocxPackager
    {
        // OOXML namespace URIs
        private const string NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        private const string NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        private const string NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships";

        private const string REL_TYPE_DOCUMENT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        private const string REL_TYPE_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
        private const string REL_TYPE_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings";
        private const string REL_TYPE_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
        private const string REL_TYPE_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

        private readonly ConversionOptions _options;

        private int _relationshipIdCounter = 3; // rId1=styles, rId2=settings
        private readonly List<(string id, string type, string target, bool external)> _docRelationships = new();
        private readonly List<(byte[] data, string filename, string format)> _mediaFiles = new();
        private readonly List<string> _contentTypeOverrides = new();

        /// <summary>
        /// Creates a packager with default conversion options.
        /// </summary>
        public DocxPackager() : this(new ConversionOptions()) { }

        /// <summary>
        /// Creates a packager that uses the provided conversion options for spacing etc.
        /// </summary>
        public DocxPackager(ConversionOptions options)
        {
            _options = options ?? new ConversionOptions();
        }

        /// <summary>
        /// Generates a complete DOCX file from the document structure and saves it to the specified path.
        /// </summary>
        public void Save(DocumentStructure document, string outputPath)
        {
            var bytes = Generate(document);
            File.WriteAllBytes(outputPath, bytes);
        }

        /// <summary>
        /// Generates a complete DOCX file as a byte array.
        /// </summary>
        public byte[] Generate(DocumentStructure document)
        {
            // Reset state
            _relationshipIdCounter = 3;
            _docRelationships.Clear();
            _mediaFiles.Clear();
            _contentTypeOverrides.Clear();

            // Add standard relationships
            _docRelationships.Add(("rId1", REL_TYPE_STYLES, "styles.xml", false));
            _docRelationships.Add(("rId2", REL_TYPE_SETTINGS, "settings.xml", false));

            // Determine page size from the first page (or use defaults)
            double pageWidthPt = UnitConverter.LetterWidthPt;
            double pageHeightPt = UnitConverter.LetterHeightPt;
            if (document.Pages.Count > 0)
            {
                pageWidthPt = document.Pages[0].Width;
                pageHeightPt = document.Pages[0].Height;
            }

            // Build document.xml content
            string documentXml = BuildDocumentXml(document, pageWidthPt, pageHeightPt);
            string stylesXml = BuildStylesXml();
            string settingsXml = BuildSettingsXml();
            string contentTypesXml = BuildContentTypesXml();
            string rootRelsXml = BuildRootRelsXml();
            string docRelsXml = BuildDocRelsXml();

            // Create ZIP archive
            using var memStream = new MemoryStream();
            using (var archive = new ZipArchive(memStream, ZipArchiveMode.Create, true))
            {
                AddZipEntry(archive, "[Content_Types].xml", contentTypesXml);
                AddZipEntry(archive, "_rels/.rels", rootRelsXml);
                AddZipEntry(archive, "word/document.xml", documentXml);
                AddZipEntry(archive, "word/_rels/document.xml.rels", docRelsXml);
                AddZipEntry(archive, "word/styles.xml", stylesXml);
                AddZipEntry(archive, "word/settings.xml", settingsXml);

                // Add media files
                foreach (var media in _mediaFiles)
                {
                    var entry = archive.CreateEntry($"word/media/{media.filename}", CompressionLevel.Optimal);
                    using var entryStream = entry.Open();
                    entryStream.Write(media.data, 0, media.data.Length);
                }
            }

            return memStream.ToArray();
        }

        /// <summary>
        /// Builds the main document.xml containing all content.
        /// </summary>
        private string BuildDocumentXml(DocumentStructure document, double pageWidthPt, double pageHeightPt)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append($"<w:document xmlns:w=\"{NS_W}\" xmlns:r=\"{NS_R}\" ");
            sb.Append($"xmlns:wp=\"{NS_WP}\" xmlns:a=\"{NS_A}\" ");
            sb.AppendLine($"xmlns:pic=\"{NS_PIC}\">");
            sb.AppendLine("  <w:body>");

            foreach (var page in document.Pages)
            {
                foreach (var block in page.Blocks)
                {
                    switch (block.Type)
                    {
                        case ContentBlockType.Paragraph:
                            WriteParagraph(sb, block.Paragraph, "    ");
                            break;

                        case ContentBlockType.Table:
                            WriteTable(sb, block.Table, "    ");
                            break;

                        case ContentBlockType.Image:
                            WriteImage(sb, block.Image, "    ");
                            break;
                    }
                }

                // Add page break between pages (except after the last one)
                if (page != document.Pages.Last())
                {
                    sb.AppendLine("    <w:p>");
                    sb.AppendLine("      <w:r>");
                    sb.AppendLine("        <w:br w:type=\"page\"/>");
                    sb.AppendLine("      </w:r>");
                    sb.AppendLine("    </w:p>");
                }
            }

            // Section properties (page size and margins)
            int pgW = UnitConverter.PointsToTwips(pageWidthPt);
            int pgH = UnitConverter.PointsToTwips(pageHeightPt);
            int margin = UnitConverter.PointsToTwips(72); // 1-inch margins

            sb.AppendLine("    <w:sectPr>");
            sb.AppendLine($"      <w:pgSz w:w=\"{pgW}\" w:h=\"{pgH}\"/>");
            sb.AppendLine($"      <w:pgMar w:top=\"{margin}\" w:right=\"{margin}\" w:bottom=\"{margin}\" w:left=\"{margin}\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>");
            sb.AppendLine("    </w:sectPr>");
            sb.AppendLine("  </w:body>");
            sb.AppendLine("</w:document>");

            return sb.ToString();
        }

        /// <summary>
        /// Writes a paragraph element to the document XML.
        /// </summary>
        private void WriteParagraph(StringBuilder sb, TextParagraph para, string indent)
        {
            if (para == null) return;

            sb.AppendLine($"{indent}<w:p>");

            // Paragraph properties
            sb.AppendLine($"{indent}  <w:pPr>");
            if (!string.IsNullOrEmpty(para.Alignment) && para.Alignment != "left")
            {
                string jcVal = para.Alignment switch
                {
                    "center" => "center",
                    "right" => "right",
                    "justify" => "both",
                    _ => "left"
                };
                sb.AppendLine($"{indent}    <w:jc w:val=\"{jcVal}\"/>");
            }

            // Phase 2.3: font-based line spacing + configurable after-spacing
            {
                // Choose the highest line height across lines (most representative font)
                double lineHeightPt = para.Lines.Count > 0
                    ? para.Lines.Max(l => l.LineHeight > 0 ? l.LineHeight : l.Bounds.Height)
                    : 12.0 * _options.LineSpacingMultiplier;

                // OOXML line value in 240ths of a line (240 = single-spaced baseline)
                // lineRule="auto" lets Word scale properly
                int lineVal = (int)Math.Round(UnitConverter.PointsToTwips(lineHeightPt) * 240.0 / 240.0);
                // Simpler: express as twips directly with lineRule="exact"
                int lineTwips = UnitConverter.PointsToTwips(lineHeightPt);
                int afterTwips = UnitConverter.PointsToTwips(_options.ParagraphSpacingAfter);
                sb.AppendLine($"{indent}    <w:spacing w:after=\"{afterTwips}\" w:line=\"{lineTwips}\" w:lineRule=\"atLeast\"/>");
            }

            sb.AppendLine($"{indent}  </w:pPr>");

            // Write runs line by line; add a space-run between consecutive lines
            // to prevent words from different lines being concatenated in Word.
            for (int li = 0; li < para.Lines.Count; li++)
            {
                var line = para.Lines[li];
                if (li > 0 && line.Runs.Count > 0)
                {
                    var lastRun = para.Lines[li - 1].Runs.LastOrDefault();
                    WriteRun(sb, new TextRun
                    {
                        Text = " ",
                        FontName = lastRun?.FontName ?? string.Empty,
                        FontSize = lastRun?.FontSize ?? 12
                    }, indent + "  ");
                }
                foreach (var run in line.Runs)
                    WriteRun(sb, run, indent + "  ");
            }

            sb.AppendLine($"{indent}</w:p>");
        }

        /// <summary>
        /// Writes a text run element with formatting properties.
        /// </summary>
        private void WriteRun(StringBuilder sb, TextRun run, string indent)
        {
            if (run == null) return;

            bool hasHyperlink = !string.IsNullOrEmpty(run.HyperlinkUri);

            if (hasHyperlink)
            {
                string hyperlinkRId = RegisterHyperlink(run.HyperlinkUri);
                sb.AppendLine($"{indent}<w:hyperlink r:id=\"{hyperlinkRId}\">");
                indent += "  ";
            }

            sb.AppendLine($"{indent}<w:r>");

            // Run properties
            sb.AppendLine($"{indent}  <w:rPr>");
            if (!string.IsNullOrEmpty(run.FontName))
            {
                string escaped = EscapeXml(run.FontName);
                sb.AppendLine($"{indent}    <w:rFonts w:ascii=\"{escaped}\" w:hAnsi=\"{escaped}\" w:cs=\"{escaped}\"/>");
            }
            if (run.FontSize > 0)
            {
                int halfPts = UnitConverter.FontSizeToHalfPoints(run.FontSize);
                sb.AppendLine($"{indent}    <w:sz w:val=\"{halfPts}\"/>");
                sb.AppendLine($"{indent}    <w:szCs w:val=\"{halfPts}\"/>");
            }
            if (run.IsBold)
            {
                sb.AppendLine($"{indent}    <w:b/>");
                sb.AppendLine($"{indent}    <w:bCs/>");
            }
            if (run.IsItalic)
            {
                sb.AppendLine($"{indent}    <w:i/>");
                sb.AppendLine($"{indent}    <w:iCs/>");
            }
            if (!string.IsNullOrEmpty(run.Color) && run.Color != "000000")
            {
                sb.AppendLine($"{indent}    <w:color w:val=\"{run.Color}\"/>");
            }
            if (hasHyperlink)
            {
                sb.AppendLine($"{indent}    <w:color w:val=\"0563C1\"/>");
                sb.AppendLine($"{indent}    <w:u w:val=\"single\"/>");
            }
            sb.AppendLine($"{indent}  </w:rPr>");

            // Text content (preserve spaces)
            string escapedText = EscapeXml(run.Text);
            sb.AppendLine($"{indent}  <w:t xml:space=\"preserve\">{escapedText}</w:t>");

            sb.AppendLine($"{indent}</w:r>");

            if (hasHyperlink)
            {
                indent = indent.Substring(0, indent.Length - 2);
                sb.AppendLine($"{indent}</w:hyperlink>");
            }
        }

        /// <summary>
        /// Writes a table element to the document XML.
        /// </summary>
        private void WriteTable(StringBuilder sb, DetectedTable table, string indent)
        {
            if (table == null) return;

            sb.AppendLine($"{indent}<w:tbl>");

            // Table properties
            sb.AppendLine($"{indent}  <w:tblPr>");
            int totalWidth = UnitConverter.PointsToTwips(table.Bounds.Width);
            sb.AppendLine($"{indent}    <w:tblW w:w=\"{totalWidth}\" w:type=\"dxa\"/>");
            sb.AppendLine($"{indent}    <w:tblLayout w:type=\"fixed\"/>");
            sb.AppendLine($"{indent}    <w:tblBorders>");
            sb.AppendLine($"{indent}      <w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
            sb.AppendLine($"{indent}      <w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
            sb.AppendLine($"{indent}      <w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
            sb.AppendLine($"{indent}      <w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
            sb.AppendLine($"{indent}      <w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
            sb.AppendLine($"{indent}      <w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
            sb.AppendLine($"{indent}    </w:tblBorders>");
            sb.AppendLine($"{indent}  </w:tblPr>");

            // Grid columns
            sb.AppendLine($"{indent}  <w:tblGrid>");
            for (int c = 0; c < table.ColCount; c++)
            {
                int colW = UnitConverter.PointsToTwips(table.ColumnWidths[c]);
                sb.AppendLine($"{indent}    <w:gridCol w:w=\"{colW}\"/>");
            }
            sb.AppendLine($"{indent}  </w:tblGrid>");

            // Table rows
            for (int r = 0; r < table.RowCount; r++)
            {
                sb.AppendLine($"{indent}  <w:tr>");

                // Row properties
                int rowH = UnitConverter.PointsToTwips(table.RowHeights[r]);
                sb.AppendLine($"{indent}    <w:trPr>");
                sb.AppendLine($"{indent}      <w:trHeight w:val=\"{rowH}\" w:hRule=\"atLeast\"/>");
                sb.AppendLine($"{indent}    </w:trPr>");

                for (int c = 0; c < table.ColCount; c++)
                {
                    var cell = table.Cells[r, c];

                    // Skip cells that are continuations of a merge
                    if (cell.IsMergedContinuation)
                    {
                        // We still need to emit a <w:tc> for vertical merge continuations
                        // Check if this is a vertical merge continuation
                        bool isVMergeContinuation = r > 0 && !table.Cells[r - 1, c].IsMergedContinuation
                            && table.Cells[r - 1, c].RowSpan > 1;

                        // Check ancestors for vMerge
                        bool partOfVMerge = false;
                        for (int rr = r - 1; rr >= 0; rr--)
                        {
                            if (!table.Cells[rr, c].IsMergedContinuation &&
                                table.Cells[rr, c].RowSpan > 1 &&
                                rr + table.Cells[rr, c].RowSpan > r)
                            {
                                partOfVMerge = true;
                                break;
                            }
                        }

                        if (partOfVMerge)
                        {
                            // Emit a vMerge continuation cell
                            sb.AppendLine($"{indent}    <w:tc>");
                            sb.AppendLine($"{indent}      <w:tcPr>");
                            int cellW = UnitConverter.PointsToTwips(table.ColumnWidths[c]);
                            sb.AppendLine($"{indent}        <w:tcW w:w=\"{cellW}\" w:type=\"dxa\"/>");
                            sb.AppendLine($"{indent}        <w:vMerge/>");
                            sb.AppendLine($"{indent}      </w:tcPr>");
                            sb.AppendLine($"{indent}      <w:p/>");
                            sb.AppendLine($"{indent}    </w:tc>");
                        }
                        // For horizontal merge continuations, skip entirely
                        // (the gridSpan on the origin cell handles it)
                        continue;
                    }

                    WriteTableCell(sb, cell, table, indent + "    ");
                }

                sb.AppendLine($"{indent}  </w:tr>");
            }

            sb.AppendLine($"{indent}</w:tbl>");

            // Empty paragraph after table (required by OOXML spec)
            sb.AppendLine($"{indent}<w:p/>");
        }

        /// <summary>
        /// Writes a single table cell element.
        /// </summary>
        private void WriteTableCell(StringBuilder sb, TableCell cell, DetectedTable table, string indent)
        {
            sb.AppendLine($"{indent}<w:tc>");

            // Cell properties
            sb.AppendLine($"{indent}  <w:tcPr>");

            // Calculate cell width (sum of spanned columns)
            double cellWidthPt = 0;
            for (int c = cell.Col; c < cell.Col + cell.ColSpan && c < table.ColCount; c++)
                cellWidthPt += table.ColumnWidths[c];
            int cellW = UnitConverter.PointsToTwips(cellWidthPt);
            sb.AppendLine($"{indent}    <w:tcW w:w=\"{cellW}\" w:type=\"dxa\"/>");

            // Horizontal merge (gridSpan)
            if (cell.ColSpan > 1)
                sb.AppendLine($"{indent}    <w:gridSpan w:val=\"{cell.ColSpan}\"/>");

            // Vertical merge (restart)
            if (cell.RowSpan > 1)
                sb.AppendLine($"{indent}    <w:vMerge w:val=\"restart\"/>");

            // Cell borders
            WriteCellBorders(sb, cell, indent + "    ");

            // Cell shading
            if (!string.IsNullOrEmpty(cell.BackgroundColor))
            {
                sb.AppendLine($"{indent}    <w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{cell.BackgroundColor}\"/>");
            }

            sb.AppendLine($"{indent}  </w:tcPr>");

            // Cell content (paragraphs)
            if (cell.Paragraphs.Count > 0)
            {
                foreach (var para in cell.Paragraphs)
                {
                    WriteParagraph(sb, para, indent + "  ");
                }
            }
            else
            {
                // Empty cell must have at least one paragraph
                sb.AppendLine($"{indent}  <w:p/>");
            }

            sb.AppendLine($"{indent}</w:tc>");
        }

        /// <summary>
        /// Writes cell border properties.
        /// </summary>
        private void WriteCellBorders(StringBuilder sb, TableCell cell, string indent)
        {
            sb.AppendLine($"{indent}<w:tcBorders>");
            WriteSingleBorder(sb, "top", cell.TopBorder, indent + "  ");
            WriteSingleBorder(sb, "left", cell.LeftBorder, indent + "  ");
            WriteSingleBorder(sb, "bottom", cell.BottomBorder, indent + "  ");
            WriteSingleBorder(sb, "right", cell.RightBorder, indent + "  ");
            sb.AppendLine($"{indent}</w:tcBorders>");
        }

        private void WriteSingleBorder(StringBuilder sb, string edge, BorderStyle border, string indent)
        {
            if (border == null || border.Style == "none")
            {
                sb.AppendLine($"{indent}<w:{edge} w:val=\"none\" w:sz=\"0\" w:space=\"0\" w:color=\"auto\"/>");
            }
            else
            {
                int sz = UnitConverter.PointsToBorderWidth(border.Width);
                sb.AppendLine($"{indent}<w:{edge} w:val=\"{border.Style}\" w:sz=\"{sz}\" w:space=\"0\" w:color=\"{border.Color}\"/>");
            }
        }

        /// <summary>
        /// Writes an inline image to the document XML.
        /// </summary>
        private void WriteImage(StringBuilder sb, ImageElement image, string indent)
        {
            if (image == null || image.ImageData == null || image.ImageData.Length == 0)
                return;

            // Register the image as a media file
            string rId = RegisterImage(image);

            // Calculate image dimensions in EMU
            long cxEmu = UnitConverter.PointsToEmu(image.Bounds.Width);
            long cyEmu = UnitConverter.PointsToEmu(image.Bounds.Height);

            // Ensure reasonable size (max 6 inches wide)
            long maxWidth = 6 * UnitConverter.EmuPerInch;
            if (cxEmu > maxWidth)
            {
                double scale = (double)maxWidth / cxEmu;
                cxEmu = maxWidth;
                cyEmu = (long)(cyEmu * scale);
            }

            int docPrId = _mediaFiles.Count;

            sb.AppendLine($"{indent}<w:p>");
            sb.AppendLine($"{indent}  <w:r>");
            sb.AppendLine($"{indent}    <w:drawing>");
            sb.AppendLine($"{indent}      <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">");
            sb.AppendLine($"{indent}        <wp:extent cx=\"{cxEmu}\" cy=\"{cyEmu}\"/>");
            sb.AppendLine($"{indent}        <wp:docPr id=\"{docPrId}\" name=\"Image {docPrId}\"/>");
            sb.AppendLine($"{indent}        <a:graphic>");
            sb.AppendLine($"{indent}          <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">");
            sb.AppendLine($"{indent}            <pic:pic>");
            sb.AppendLine($"{indent}              <pic:nvPicPr>");
            sb.AppendLine($"{indent}                <pic:cNvPr id=\"0\" name=\"Image {docPrId}\"/>");
            sb.AppendLine($"{indent}                <pic:cNvPicPr/>");
            sb.AppendLine($"{indent}              </pic:nvPicPr>");
            sb.AppendLine($"{indent}              <pic:blipFill>");
            sb.AppendLine($"{indent}                <a:blip r:embed=\"{rId}\"/>");
            sb.AppendLine($"{indent}                <a:stretch><a:fillRect/></a:stretch>");
            sb.AppendLine($"{indent}              </pic:blipFill>");
            sb.AppendLine($"{indent}              <pic:spPr>");
            sb.AppendLine($"{indent}                <a:xfrm>");
            sb.AppendLine($"{indent}                  <a:off x=\"0\" y=\"0\"/>");
            sb.AppendLine($"{indent}                  <a:ext cx=\"{cxEmu}\" cy=\"{cyEmu}\"/>");
            sb.AppendLine($"{indent}                </a:xfrm>");
            sb.AppendLine($"{indent}                <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>");
            sb.AppendLine($"{indent}              </pic:spPr>");
            sb.AppendLine($"{indent}            </pic:pic>");
            sb.AppendLine($"{indent}          </a:graphicData>");
            sb.AppendLine($"{indent}        </a:graphic>");
            sb.AppendLine($"{indent}      </wp:inline>");
            sb.AppendLine($"{indent}    </w:drawing>");
            sb.AppendLine($"{indent}  </w:r>");
            sb.AppendLine($"{indent}</w:p>");
        }

        /// <summary>
        /// Registers an image as a media file and returns its relationship ID.
        /// </summary>
        private string RegisterImage(ImageElement image)
        {
            string rId = $"rId{_relationshipIdCounter++}";
            int imageIndex = _mediaFiles.Count + 1;
            string filename = $"image{imageIndex}.{image.Format}";

            _mediaFiles.Add((image.ImageData, filename, image.Format));
            _docRelationships.Add((rId, REL_TYPE_IMAGE, $"media/{filename}", false));

            return rId;
        }

        /// <summary>
        /// Registers a hyperlink and returns its relationship ID.
        /// </summary>
        private string RegisterHyperlink(string uri)
        {
            // Check if this URI is already registered
            var existing = _docRelationships.FirstOrDefault(r => r.type == REL_TYPE_HYPERLINK && r.target == uri);
            if (existing.id != null)
                return existing.id;

            string rId = $"rId{_relationshipIdCounter++}";
            _docRelationships.Add((rId, REL_TYPE_HYPERLINK, uri, true));
            return rId;
        }

        // --- XML for supporting parts ---

        /// <summary>
        /// Builds the [Content_Types].xml manifest.
        /// </summary>
        private string BuildContentTypesXml()
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.AppendLine("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            sb.AppendLine("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            sb.AppendLine("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>");

            // Add media type defaults
            var formats = _mediaFiles.Select(m => m.format).Distinct();
            foreach (var format in formats)
            {
                string contentType = format switch
                {
                    "png" => "image/png",
                    "jpg" or "jpeg" => "image/jpeg",
                    "gif" => "image/gif",
                    "bmp" => "image/bmp",
                    "tiff" => "image/tiff",
                    _ => "application/octet-stream"
                };
                sb.AppendLine($"  <Default Extension=\"{format}\" ContentType=\"{contentType}\"/>");
            }

            sb.AppendLine("  <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>");
            sb.AppendLine("  <Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>");
            sb.AppendLine("  <Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/>");
            sb.AppendLine("</Types>");

            return sb.ToString();
        }

        /// <summary>
        /// Builds the root _rels/.rels file.
        /// </summary>
        private string BuildRootRelsXml()
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.AppendLine($"<Relationships xmlns=\"{NS_RELS}\">");
            sb.AppendLine($"  <Relationship Id=\"rId1\" Type=\"{REL_TYPE_DOCUMENT}\" Target=\"word/document.xml\"/>");
            sb.AppendLine("</Relationships>");
            return sb.ToString();
        }

        /// <summary>
        /// Builds the word/_rels/document.xml.rels file.
        /// </summary>
        private string BuildDocRelsXml()
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.AppendLine($"<Relationships xmlns=\"{NS_RELS}\">");

            foreach (var rel in _docRelationships)
            {
                string targetAttr = EscapeXml(rel.target);
                if (rel.external)
                {
                    sb.AppendLine($"  <Relationship Id=\"{rel.id}\" Type=\"{rel.type}\" Target=\"{targetAttr}\" TargetMode=\"External\"/>");
                }
                else
                {
                    sb.AppendLine($"  <Relationship Id=\"{rel.id}\" Type=\"{rel.type}\" Target=\"{targetAttr}\"/>");
                }
            }

            sb.AppendLine("</Relationships>");
            return sb.ToString();
        }

        /// <summary>
        /// Builds the word/styles.xml file with basic document styles.
        /// </summary>
        private string BuildStylesXml()
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.AppendLine($"<w:styles xmlns:w=\"{NS_W}\">");

            // Default styles
            sb.AppendLine("  <w:docDefaults>");
            sb.AppendLine("    <w:rPrDefault>");
            sb.AppendLine("      <w:rPr>");
            sb.AppendLine("        <w:rFonts w:ascii=\"Calibri\" w:hAnsi=\"Calibri\" w:cs=\"Calibri\"/>");
            sb.AppendLine("        <w:sz w:val=\"22\"/>");
            sb.AppendLine("        <w:szCs w:val=\"22\"/>");
            sb.AppendLine("      </w:rPr>");
            sb.AppendLine("    </w:rPrDefault>");
            sb.AppendLine("    <w:pPrDefault>");
            sb.AppendLine("      <w:pPr>");
            sb.AppendLine("        <w:spacing w:after=\"0\" w:line=\"240\" w:lineRule=\"auto\"/>");
            sb.AppendLine("      </w:pPr>");
            sb.AppendLine("    </w:pPrDefault>");
            sb.AppendLine("  </w:docDefaults>");

            // Normal style
            sb.AppendLine("  <w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\">");
            sb.AppendLine("    <w:name w:val=\"Normal\"/>");
            sb.AppendLine("    <w:qFormat/>");
            sb.AppendLine("  </w:style>");

            // Heading 1
            sb.AppendLine("  <w:style w:type=\"paragraph\" w:styleId=\"Heading1\">");
            sb.AppendLine("    <w:name w:val=\"heading 1\"/>");
            sb.AppendLine("    <w:basedOn w:val=\"Normal\"/>");
            sb.AppendLine("    <w:qFormat/>");
            sb.AppendLine("    <w:pPr>");
            sb.AppendLine("      <w:spacing w:before=\"240\" w:after=\"60\"/>");
            sb.AppendLine("    </w:pPr>");
            sb.AppendLine("    <w:rPr>");
            sb.AppendLine("      <w:b/>");
            sb.AppendLine("      <w:sz w:val=\"32\"/>");
            sb.AppendLine("    </w:rPr>");
            sb.AppendLine("  </w:style>");

            // Heading 2
            sb.AppendLine("  <w:style w:type=\"paragraph\" w:styleId=\"Heading2\">");
            sb.AppendLine("    <w:name w:val=\"heading 2\"/>");
            sb.AppendLine("    <w:basedOn w:val=\"Normal\"/>");
            sb.AppendLine("    <w:qFormat/>");
            sb.AppendLine("    <w:pPr>");
            sb.AppendLine("      <w:spacing w:before=\"240\" w:after=\"60\"/>");
            sb.AppendLine("    </w:pPr>");
            sb.AppendLine("    <w:rPr>");
            sb.AppendLine("      <w:b/>");
            sb.AppendLine("      <w:sz w:val=\"28\"/>");
            sb.AppendLine("    </w:rPr>");
            sb.AppendLine("  </w:style>");

            // Hyperlink character style
            sb.AppendLine("  <w:style w:type=\"character\" w:styleId=\"Hyperlink\">");
            sb.AppendLine("    <w:name w:val=\"Hyperlink\"/>");
            sb.AppendLine("    <w:rPr>");
            sb.AppendLine("      <w:color w:val=\"0563C1\"/>");
            sb.AppendLine("      <w:u w:val=\"single\"/>");
            sb.AppendLine("    </w:rPr>");
            sb.AppendLine("  </w:style>");

            // Table Normal style
            sb.AppendLine("  <w:style w:type=\"table\" w:default=\"1\" w:styleId=\"TableNormal\">");
            sb.AppendLine("    <w:name w:val=\"Normal Table\"/>");
            sb.AppendLine("    <w:tblPr>");
            sb.AppendLine("      <w:tblInd w:w=\"0\" w:type=\"dxa\"/>");
            sb.AppendLine("      <w:tblCellMar>");
            sb.AppendLine("        <w:top w:w=\"0\" w:type=\"dxa\"/>");
            sb.AppendLine("        <w:left w:w=\"108\" w:type=\"dxa\"/>");
            sb.AppendLine("        <w:bottom w:w=\"0\" w:type=\"dxa\"/>");
            sb.AppendLine("        <w:right w:w=\"108\" w:type=\"dxa\"/>");
            sb.AppendLine("      </w:tblCellMar>");
            sb.AppendLine("    </w:tblPr>");
            sb.AppendLine("  </w:style>");

            sb.AppendLine("</w:styles>");
            return sb.ToString();
        }

        /// <summary>
        /// Builds the word/settings.xml file.
        /// </summary>
        private string BuildSettingsXml()
        {
            var sb = new StringBuilder();
            sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.AppendLine($"<w:settings xmlns:w=\"{NS_W}\">");
            sb.AppendLine("  <w:defaultTabStop w:val=\"720\"/>");
            sb.AppendLine("  <w:characterSpacingControl w:val=\"doNotCompress\"/>");
            sb.AppendLine("  <w:compat>");
            sb.AppendLine("    <w:compatSetting w:name=\"compatibilityMode\" w:uri=\"http://schemas.microsoft.com/office/word\" w:val=\"15\"/>");
            sb.AppendLine("  </w:compat>");
            sb.AppendLine("</w:settings>");
            return sb.ToString();
        }

        // --- Utility Methods ---

        /// <summary>
        /// Adds a text entry to the ZIP archive.
        /// </summary>
        private void AddZipEntry(ZipArchive archive, string entryPath, string content)
        {
            var entry = archive.CreateEntry(entryPath, CompressionLevel.Optimal);
            using var stream = entry.Open();
            var bytes = Encoding.UTF8.GetBytes(content);
            stream.Write(bytes, 0, bytes.Length);
        }

        /// <summary>
        /// Escapes special XML characters in a string.
        /// </summary>
        private string EscapeXml(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }
    }
}
