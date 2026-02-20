using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using PDFtoDOCX.Docx;
using PDFtoDOCX.Models;
using Xunit;

namespace PDFtoDOCX.Tests
{
    /// <summary>
    /// Validates that generated DOCX files conform to the OOXML specification.
    /// Uses only built-in .NET ZIP and XML libraries — no external SDK required.
    /// </summary>
    public class DocxValidationTests
    {
        // ── Helpers ───────────────────────────────────────────────────────────

        private static byte[] GenerateMinimalDocx(IEnumerable<ContentBlock>? blocks = null)
        {
            var packager = new DocxPackager();
            var doc = new DocumentStructure
            {
                Pages = new List<PageStructure>
                {
                    new PageStructure
                    {
                        PageNumber = 1, Width = 612, Height = 792,
                        Blocks = (blocks ?? Enumerable.Empty<ContentBlock>()).ToList()
                    }
                }
            };
            return packager.Generate(doc);
        }

        private static ZipArchive OpenDocx(byte[] bytes)
            => new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);

        private static string ReadEntry(ZipArchive archive, string entryPath)
        {
            var entry = archive.GetEntry(entryPath)
                ?? throw new InvalidOperationException($"Missing DOCX part: {entryPath}");
            using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
            return reader.ReadToEnd();
        }

        // ── [Content_Types].xml ────────────────────────────────────────────────

        [Fact]
        public void ContentTypes_ContainsDocumentOverride()
        {
            var bytes = GenerateMinimalDocx();
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "[Content_Types].xml");

            Assert.Contains("word/document.xml", xml);
            Assert.Contains("wordprocessingml.document.main+xml", xml);
        }

        [Fact]
        public void ContentTypes_ContainsStylesAndSettings()
        {
            var bytes = GenerateMinimalDocx();
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "[Content_Types].xml");

            Assert.Contains("word/styles.xml", xml);
            Assert.Contains("word/settings.xml", xml);
        }

        [Fact]
        public void ContentTypes_ImagesPresentWhenDocxHasImages()
        {
            var imageData = new byte[200]; // dummy PNG-like data
            imageData[0] = 0x89; imageData[1] = 0x50; imageData[2] = 0x4E; imageData[3] = 0x47; // PNG magic

            var block = new ContentBlock
            {
                Type = ContentBlockType.Image,
                Bounds = new Rect(72, 72, 200, 150),
                Image = new ImageElement
                {
                    Bounds = new Rect(72, 72, 200, 150),
                    ImageData = imageData,
                    Format = "png",
                    WidthPx = 100, HeightPx = 60
                }
            };

            var bytes = GenerateMinimalDocx(new[] { block });
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "[Content_Types].xml");

            Assert.Contains("image/png", xml);
        }

        // ── _rels/.rels ────────────────────────────────────────────────────────

        [Fact]
        public void RootRels_PointsToDocument()
        {
            var bytes = GenerateMinimalDocx();
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "_rels/.rels");

            Assert.Contains("word/document.xml", xml);
            Assert.Contains("officeDocument", xml);
        }

        // ── word/document.xml ─────────────────────────────────────────────────

        [Fact]
        public void DocumentXml_IsWellFormedXml()
        {
            var bytes = GenerateMinimalDocx();
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "word/document.xml");

            // Parse as XML — throws XmlException on malformed content
            var doc = new XmlDocument();
            doc.LoadXml(xml);   // must not throw
            Assert.NotNull(doc.DocumentElement);
        }

        [Fact]
        public void DocumentXml_ContainsSectionProperties()
        {
            var bytes = GenerateMinimalDocx();
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "word/document.xml");

            Assert.Contains("w:sectPr", xml);
            Assert.Contains("w:pgSz",   xml);
            Assert.Contains("w:pgMar",  xml);
        }

        [Fact]
        public void DocumentXml_TableStructureIsValid()
        {
            var table = new DetectedTable
            {
                Bounds = new Rect(72, 100, 540, 300),
                RowCount = 2, ColCount = 2,
                ColumnWidths = new double[] { 234, 234 },
                RowHeights   = new double[] { 100, 100 },
                Cells        = new TableCell[2, 2]
            };

            for (int r = 0; r < 2; r++)
            {
                for (int c = 0; c < 2; c++)
                {
                    table.Cells[r, c] = new TableCell
                    {
                        Row = r, Col = c,
                        Bounds = new Rect(72 + c * 234, 100 + r * 100,
                                          72 + (c + 1) * 234, 100 + (r + 1) * 100),
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
                                            new TextRun { Text = $"R{r}C{c}", FontSize = 12 }
                                        }
                                    }
                                }
                            }
                        }
                    };
                }
            }

            var block = new ContentBlock
            {
                Type = ContentBlockType.Table,
                Table = table,
                Bounds = table.Bounds
            };

            var bytes = GenerateMinimalDocx(new[] { block });
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "word/document.xml");

            // Structural checks
            Assert.Contains("<w:tbl>",   xml);
            Assert.Contains("<w:tr>",    xml);
            Assert.Contains("<w:tc>",    xml);
            Assert.Contains("<w:tblW",   xml);
            Assert.Contains("<w:tblGrid", xml);
        }

        // ── word/_rels/document.xml.rels ──────────────────────────────────────

        [Fact]
        public void DocRels_ContainsStylesAndSettingsRelationships()
        {
            var bytes = GenerateMinimalDocx();
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "word/_rels/document.xml.rels");

            Assert.Contains("styles.xml", xml);
            Assert.Contains("settings.xml", xml);
        }

        [Fact]
        public void DocRels_AllImageRelationshipsResolveToMediaFiles()
        {
            var imageData = new byte[200];
            imageData[0] = 0x89; imageData[1] = 0x50; imageData[2] = 0x4E; imageData[3] = 0x47;

            var block = new ContentBlock
            {
                Type = ContentBlockType.Image,
                Bounds = new Rect(72, 72, 200, 150),
                Image = new ImageElement
                {
                    Bounds = new Rect(72, 72, 200, 150),
                    ImageData = imageData, Format = "png",
                    WidthPx = 100, HeightPx = 60
                }
            };

            var bytes = GenerateMinimalDocx(new[] { block });
            using var archive = OpenDocx(bytes);

            var relsXml = ReadEntry(archive, "word/_rels/document.xml.rels");
            var entries = archive.Entries.Select(e => e.FullName).ToHashSet();

            // Parse relationship targets
            var xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(relsXml);
            var ns = new XmlNamespaceManager(xmlDoc.NameTable);
            ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

            foreach (XmlElement rel in xmlDoc.SelectNodes("//r:Relationship[@Type[contains(., 'image')]]", ns)!)
            {
                string target = rel.GetAttribute("Target");
                string fullPath = $"word/{target}";
                Assert.Contains(fullPath, entries);
            }
        }

        // ── word/styles.xml ───────────────────────────────────────────────────

        [Fact]
        public void StylesXml_ContainsNormalAndTableNormalStyles()
        {
            var bytes = GenerateMinimalDocx();
            using var archive = OpenDocx(bytes);
            var xml = ReadEntry(archive, "word/styles.xml");

            Assert.Contains("Normal", xml);
            Assert.Contains("TableNormal", xml);
            Assert.Contains("Hyperlink", xml);
        }
    }
}
