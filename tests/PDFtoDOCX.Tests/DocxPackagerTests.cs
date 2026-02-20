using System.IO;
using System.Linq;
using System.IO.Compression;
using PDFtoDOCX.Docx;
using PDFtoDOCX.Models;
using System.Collections.Generic;
using Xunit;

namespace PDFtoDOCX.Tests
{
    public class DocxPackagerTests
    {
        [Fact]
        public void Generate_EmptyDocument_ProducesValidZip()
        {
            var packager = new DocxPackager();
            var doc = new DocumentStructure
            {
                Pages = new List<PageStructure>
                {
                    new PageStructure
                    {
                        PageNumber = 1,
                        Width = 612,
                        Height = 792,
                        Blocks = new List<ContentBlock>()
                    }
                }
            };

            var bytes = packager.Generate(doc);

            Assert.NotNull(bytes);
            Assert.NotEmpty(bytes);

            // Verify it's a valid ZIP archive
            using var stream = new MemoryStream(bytes);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

            var entryNames = archive.Entries.Select(e => e.FullName).ToList();

            // Verify required DOCX parts exist
            Assert.Contains("[Content_Types].xml", entryNames);
            Assert.Contains("_rels/.rels", entryNames);
            Assert.Contains("word/document.xml", entryNames);
            Assert.Contains("word/_rels/document.xml.rels", entryNames);
            Assert.Contains("word/styles.xml", entryNames);
            Assert.Contains("word/settings.xml", entryNames);
        }

        [Fact]
        public void Generate_WithParagraph_ContainsText()
        {
            var packager = new DocxPackager();
            var doc = new DocumentStructure
            {
                Pages = new List<PageStructure>
                {
                    new PageStructure
                    {
                        PageNumber = 1,
                        Width = 612,
                        Height = 792,
                        Blocks = new List<ContentBlock>
                        {
                            new ContentBlock
                            {
                                Type = ContentBlockType.Paragraph,
                                Bounds = new Rect(72, 72, 540, 84),
                                Paragraph = new TextParagraph
                                {
                                    Bounds = new Rect(72, 72, 540, 84),
                                    Lines = new List<TextLine>
                                    {
                                        new TextLine
                                        {
                                            Bounds = new Rect(72, 72, 540, 84),
                                            Runs = new List<TextRun>
                                            {
                                                new TextRun
                                                {
                                                    Text = "Hello World",
                                                    FontName = "Arial",
                                                    FontSize = 12,
                                                    Color = "000000"
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            var bytes = packager.Generate(doc);

            // Verify document.xml contains our text
            using var stream = new MemoryStream(bytes);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var docEntry = archive.GetEntry("word/document.xml");
            using var reader = new StreamReader(docEntry.Open());
            var xml = reader.ReadToEnd();

            Assert.Contains("Hello World", xml);
            Assert.Contains("Arial", xml);
        }

        [Fact]
        public void Generate_WithTable_ContainsTableXml()
        {
            var packager = new DocxPackager();

            var table = new DetectedTable
            {
                Bounds = new Rect(100, 100, 500, 200),
                RowCount = 1,
                ColCount = 2,
                ColumnWidths = new double[] { 200, 200 },
                RowHeights = new double[] { 100 },
                Cells = new TableCell[1, 2]
            };

            table.Cells[0, 0] = new TableCell
            {
                Row = 0, Col = 0,
                Bounds = new Rect(100, 100, 300, 200),
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
                                    new TextRun { Text = "CellA" }
                                }
                            }
                        }
                    }
                }
            };

            table.Cells[0, 1] = new TableCell
            {
                Row = 0, Col = 1,
                Bounds = new Rect(300, 100, 500, 200),
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
                                    new TextRun { Text = "CellB" }
                                }
                            }
                        }
                    }
                }
            };

            var doc = new DocumentStructure
            {
                Pages = new List<PageStructure>
                {
                    new PageStructure
                    {
                        PageNumber = 1,
                        Width = 612,
                        Height = 792,
                        Blocks = new List<ContentBlock>
                        {
                            new ContentBlock
                            {
                                Type = ContentBlockType.Table,
                                Table = table,
                                Bounds = table.Bounds
                            }
                        }
                    }
                }
            };

            var bytes = packager.Generate(doc);

            using var stream = new MemoryStream(bytes);
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
            var docEntry = archive.GetEntry("word/document.xml");
            using var reader = new StreamReader(docEntry.Open());
            var xml = reader.ReadToEnd();

            Assert.Contains("w:tbl", xml);
            Assert.Contains("w:tr", xml);
            Assert.Contains("w:tc", xml);
            Assert.Contains("CellA", xml);
            Assert.Contains("CellB", xml);
        }

        [Fact]
        public void Save_CreatesFile()
        {
            var packager = new DocxPackager();
            var doc = new DocumentStructure
            {
                Pages = new List<PageStructure>
                {
                    new PageStructure
                    {
                        PageNumber = 1,
                        Width = 612,
                        Height = 792,
                        Blocks = new List<ContentBlock>()
                    }
                }
            };

            var tempPath = Path.GetTempFileName() + ".docx";
            try
            {
                packager.Save(doc, tempPath);
                Assert.True(File.Exists(tempPath));
                Assert.True(new FileInfo(tempPath).Length > 0);
            }
            finally
            {
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }
    }
}
