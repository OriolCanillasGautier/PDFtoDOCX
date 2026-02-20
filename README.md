# PDFtoDOCX

A high-fidelity PDF-to-DOCX conversion library written in C# (.NET 6+).  
Replicates the layout intelligence of Python's `pdf2docx` — entirely in .NET, with no Python dependencies.

## Features

- **Text extraction with formatting** — font name, size, bold, italic, color
- **Layout reconstruction** — line grouping, paragraph detection, multi-column support
- **Table detection** — grid detection via line intersections, merged cells, border styles, cell shading
- **Image embedding** — extracts and embeds images (PNG, JPEG, GIF, BMP, TIFF)
- **Hyperlink preservation** — detects link annotations and maps them to DOCX hyperlinks
- **Manual OOXML generation** — builds the DOCX ZIP package from scratch for maximum control and spec conformance
- **Page range support** — convert specific pages or ranges
- **Configurable heuristics** — tune line grouping tolerance, paragraph gap, column gap, table detection sensitivity

## Architecture

```
PDF Input
  │
  ▼
┌──────────────────────────┐
│  Phase 1: Extraction     │  PdfPig (Apache 2.0)
│  Text, Images, Paths,    │  Letters → Words → TextElements
│  Annotations             │  Paths → LineSegments / Rectangles
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│  Phase 2: Table Detection│  Custom C# algorithms
│  Grid detection          │  Line intersection analysis
│  Cell population         │  Merged cell detection
│  Border style mapping    │  Cell shading from filled rectangles
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│  Phase 3: Layout Analysis│  Custom C# algorithms
│  Y-coordinate clustering │  Line grouping
│  X-band detection        │  Column detection
│  Gap-based segmentation  │  Paragraph assembly
└──────────┬───────────────┘
           │
           ▼
┌──────────────────────────┐
│  Phase 4: DOCX Packaging │  System.IO.Compression
│  Manual XML generation   │  document.xml, styles.xml, etc.
│  ZIP archive assembly    │  [Content_Types].xml, .rels files
│  Image embedding         │  word/media/ folder
└──────────────────────────┘
           │
           ▼
      DOCX Output
```

## Project Structure

```
src/
  PDFtoDOCX/                    # Core library (class library)
    Models/CoreModels.cs        # All data model classes
    Models/DocumentStructure.cs # Page/document structure models
    Extraction/                 # PdfPig-based PDF content extraction
    Layout/                     # Layout analysis (lines, paragraphs, columns)
    Tables/                     # Table detection and reconstruction
    Docx/                       # DOCX XML generation and ZIP packaging
    Utils/                      # Unit conversion utilities
    PdfToDocxConverter.cs       # Main entry point / orchestrator
    ConversionOptions.cs        # Configuration options
  PDFtoDOCX.Console/           # Command-line application
tests/
  PDFtoDOCX.Tests/             # Unit tests (xUnit)
```

## Quick Start

### As a Library

```csharp
using PDFtoDOCX;

// Simple conversion
using var converter = new PdfToDocxConverter();
converter.Convert("input.pdf", "output.docx");

// With options
var options = new ConversionOptions
{
    StartPage = 1,
    EndPage = 5,
    DetectTables = true,
    ExtractImages = true,
    ParagraphGapMultiplier = 1.5
};

using var converter2 = new PdfToDocxConverter(options);
converter2.Convert("input.pdf", "output.docx");

// From byte array / stream
byte[] pdfBytes = File.ReadAllBytes("input.pdf");
byte[] docxBytes = converter.ConvertToBytes(pdfBytes);
File.WriteAllBytes("output.docx", docxBytes);
```

### Command Line

```bash
# Basic conversion (output: input.docx)
dotnet run --project src/PDFtoDOCX.Console -- input.pdf

# Specify output path
dotnet run --project src/PDFtoDOCX.Console -- input.pdf output.docx

# Convert specific pages
dotnet run --project src/PDFtoDOCX.Console -- --start-page 2 --end-page 10 input.pdf

# Skip tables and images for faster conversion
dotnet run --project src/PDFtoDOCX.Console -- --no-tables --no-images input.pdf
```

### CLI Options

| Option | Description | Default |
|--------|-------------|---------|
| `--no-images` | Skip image extraction | `false` |
| `--no-tables` | Skip table detection | `false` |
| `--no-hyperlinks` | Skip hyperlink detection | `false` |
| `--start-page N` | First page to convert (1-based) | `1` |
| `--end-page N` | Last page to convert (inclusive) | last page |
| `--max-pages N` | Maximum pages to convert | all |
| `--line-tolerance N` | Y-tolerance for line grouping (points) | `3.0` |
| `--para-gap N` | Paragraph gap multiplier | `1.3` |
| `--column-gap N` | Minimum column gap (points) | `20.0` |

## Building

```bash
# Build entire solution
dotnet build

# Run tests
dotnet test

# Publish console app
dotnet publish src/PDFtoDOCX.Console -c Release -o publish
```

## Dependencies

| Package | License | Purpose |
|---------|---------|---------|
| [UglyToad.PdfPig](https://github.com/UglyToad/PdfPig) | Apache 2.0 | PDF content extraction (text, images, paths, annotations) |
| System.IO.Compression | MIT (.NET BCL) | ZIP archive creation for DOCX packaging |

## How It Works

### 1. Content Extraction
Uses PdfPig to scan each PDF page and extract:
- **Text** at word level with bounding boxes, font name, size, color
- **Images** with position data, converted to PNG/JPEG
- **Vector paths** decomposed into horizontal/vertical line segments for table detection
- **Annotations** (link annotations with URIs)

All coordinates are converted from PDF bottom-left origin to top-left origin for easier processing.

### 2. Table Detection
Analyzes extracted line segments to find tabular structures:
- Separates horizontal and vertical lines
- Snaps endpoints to a grid and finds intersection clusters
- Validates that intersections form a complete grid (top, bottom, left, right edges present)
- Detects **merged cells** by checking for missing internal grid lines
- Populates cells with text content that falls within their bounds
- Maps **border styles** (width, color, solid/dashed) from the original PDF lines
- Applies **cell shading** from filled rectangles overlapping cell areas

### 3. Layout Analysis
Transforms the remaining (non-table) text elements into paragraphs:
- **Line grouping**: clusters text by Y-coordinate with dynamic tolerance based on font size
- **Column detection**: builds a horizontal histogram of text coverage, finds gaps exceeding the minimum column gap threshold
- **Paragraph detection**: measures vertical gaps between consecutive lines; gaps exceeding the threshold (line height × multiplier) start a new paragraph
- **Alignment detection**: analyzes line positions relative to page margins to determine left/center/right/justify

### 4. DOCX Generation
Constructs a valid Office Open XML document from scratch:
- Generates all required XML parts: `document.xml`, `styles.xml`, `settings.xml`
- Builds relationship files (`_rels/.rels`, `document.xml.rels`)
- Creates the `[Content_Types].xml` manifest
- Embeds images in `word/media/` with proper relationship references
- Handles table markup including `gridSpan` for merged columns and `vMerge` for merged rows
- Assembles everything into a ZIP archive with correct structure

## Configuration Guide

| Option | Effect | When to Adjust |
|--------|--------|----------------|
| `LineGroupingTolerance` | How close (in points) text must be vertically to be on the same line | Increase for documents with uneven baselines |
| `ParagraphGapMultiplier` | How much larger than average the gap must be to start a new paragraph | Increase for tightly-spaced documents, decrease for loosely-spaced |
| `MinColumnGap` | Minimum horizontal gap (points) to split text into separate columns | Decrease for narrow-column layouts |
| `MinTableLineLength` | Minimum line length (points) to consider for table borders | Increase to ignore small decorative lines |
| `TableGridSnapTolerance` | How close line endpoints must be to snap to the same grid point | Increase for slightly misaligned table borders |

## License

See [LICENSE](LICENSE) for details.