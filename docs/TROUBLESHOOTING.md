# PDFtoDOCX Conversion Troubleshooting Guide

## Table of Contents
1. [Phantom Tables Detected](#phantom-tables-detected)
2. [Missing or Corrupted Text](#missing-or-corrupted-text)
3. [Images Not Appearing](#images-not-appearing)
4. [Incorrect Text Spacing](#incorrect-text-spacing)
5. [Output Opens with Repair Dialog](#output-opens-with-repair-dialog)
6. [Performance Issues](#performance-issues)
7. [Enabling Diagnostic Mode](#enabling-diagnostic-mode)
8. [Known Limitations](#known-limitations)

---

## Phantom Tables Detected

**Symptom:** Lines of text rendered inside a table when they should be free-flowing paragraphs. Often looks like a page-wide border became a table.

**Causes:**
- Page-border or decorative rectangle lines being treated as table borders.
- Section dividers (full-width horizontal rules) interpreted as table rows.

**Diagnostic steps:**
1. Run with `--diagnostics` to see confidence scores per detected table.
2. Look for tables with very large bounds (>80% of page in both dimensions) — these should be rejected automatically.
3. Check line thickness: decorative lines are often ≤1 pt; table borders are typically ≥0.5 pt.

**Fixes:**
- Increase `MinTableLineLength` (default 10 pt) to ignore short decorative strokes.
- Increase `TableGridSnapTolerance` (default 5 pt) if grid lines are slightly misaligned.
- Set `DetectTables = false` to disable table detection entirely for affected documents.
- The validator already rejects:
  - 1×1 grids (decorative boxes)
  - Grids spanning >80% of page width AND height simultaneously
  - Grids with confidence score <0.4

---

## Missing or Corrupted Text

**Symptom A:** Large sections of the page are blank in the DOCX even though text is visible in the PDF.

**Cause:** The PDF stores text as vector paths (glyph outlines) with no Unicode text operators. PdfPig cannot extract text from pure vector glyphs.

**Diagnostic steps:**
1. Run with `--diagnostics`. Look for: `No text operators found and OCR is disabled`.
2. Open the PDF in Acrobat and try **Edit → Copy All Text**: if nothing is selected, the text is rasterized or path-based.

**Fixes:**
- Enable OCR: `--ocr` flag (requires an OCR back-end — see [OCR Integration](#known-limitations)).
- If using the API: set `options.EnableOcr = true` and subclass `OcrTextExtractor`.

**Symptom B:** Words run together without spaces ("HelloWorld" instead of "Hello World").

**Cause:** Letter grouping fallback uses a gap threshold that is too wide for the font metrics.

**Fix:** Lower `LetterGroupingGapMultiplier` (default 0.5). Try 0.3 for tightly-spaced fonts:
```csharp
options.LetterGroupingGapMultiplier = 0.3;
```
Or via CLI: this option is not yet exposed as a command-line flag.

---

## Images Not Appearing

**Symptom:** Images visible in the PDF are absent from the DOCX.

**Cause A:** Images are part of Form XObjects (nested resource dictionaries). The current PdfPig version does not expose `GetFormXObjects()`.

**Cause B:** Images are smaller than 10×10 pixels (filtered by the minimum-size check).

**Cause C:** The image data could not be decoded (corrupted stream or unsupported compression).

**Diagnostic steps:**
1. Run with `--diagnostics`. Look for: `X top-level image(s) extracted`.
2. If 0 images are extracted, the images are likely stored as Form XObjects or rasterized paths.

**Fixes:**
- Upgrade to a newer PdfPig build that exposes Form XObject iterators.
- For logos rasterized as vector paths, enable vector-to-raster conversion (Phase 5.2 — not yet implemented; planned via SkiaSharp).
- Disable the minimum-size filter by editing `PdfContentExtractor.ExtractImagesFromSource` if small icons are important.

---

## Incorrect Text Spacing

**Symptom:** Lines of text are too close together or too far apart in the DOCX output.

**Cause:** The default line height multiplier (1.15×) may not suit all font sizes.

**Fixes:**
- Adjust `LineSpacingMultiplier` (default 1.15):
  ```csharp
  options.LineSpacingMultiplier = 1.0;  // single-spaced
  options.LineSpacingMultiplier = 1.5;  // more open
  ```
  CLI: `--line-spacing 1.5`
- Adjust `ParagraphSpacingAfter` (default 6 pt):
  ```csharp
  options.ParagraphSpacingAfter = 0.0;  // no space after paragraphs
  options.ParagraphSpacingAfter = 12.0; // double space
  ```
  CLI: `--para-spacing-after 12`
- Adjust `ParagraphGapMultiplier` (default 1.3) if paragraph breaks are not detected correctly.

---

## Output Opens with Repair Dialog

**Symptom:** Word displays "We found a problem with some content…" when opening the DOCX.

**Causes:**
- Malformed XML (unescaped special characters, unclosed tags).
- Relationship IDs referenced in `document.xml` that don't exist in `document.xml.rels`.
- Missing media files referenced by image relationships.

**Diagnostic steps:**
1. Rename the `.docx` to `.zip` and extract it.
2. Open `word/document.xml` in a text editor or XML validator.
3. Look for raw `<`, `>`, `&` characters not inside CDATA.
4. Check that every `r:embed="rIdN"` in `word/document.xml.rels` resolves to a file in `word/media/`.

**Known safe characters:** The packager escapes `<`, `>`, `&`, `"`, and `'` in all text content. If you see unescaped characters, please file an issue with the source PDF.

---

## Performance Issues

**Target:** <2 seconds per page for documents with standard content.

**Symptom:** Conversion of a 50-page document takes >2 minutes.

**Likely causes:**
- Complex path structures (many thousands of line segments per page).
- A large number of rectangles triggering table detection on every page.

**Fixes:**
- Set `MinTableLineLength = 20.0` to reduce the set of lines processed by the table detector.
- Set `DetectTables = false` if tables are not needed.
- Use `StartPage` and `EndPage` or `MaxPages` to limit the pages processed.
- Run with `--diagnostics` to identify which page has the largest line/element counts.

---

## Enabling Diagnostic Mode

Add `--diagnostics` to the CLI command or set `EnableDiagnostics = true` in code:

```bash
# CLI
PDFtoDOCX.Console --diagnostics document.pdf output.docx

# Code
var options = new ConversionOptions { EnableDiagnostics = true };
using var converter = new PdfToDocxConverter(options);
converter.Convert("input.pdf", "output.docx");
```

Diagnostic output includes:
- Number of pages extracted.
- Per-page: text element count, line segment count, table count.
- Per-table: dimensions, confidence score, and rejection reason (if any).
- Per-page: number of images extracted at the top level.
- When OCR is enabled: fallback activation messages.

---

## Known Limitations

| Limitation | Status | Workaround |
|---|---|---|
| Text stored as vector glyph paths | No native extraction | Enable OCR (requires Tesseract back-end) |
| Form XObject images | Not extracted (PdfPig API pending) | Upgrade PdfPig when API stabilizes |
| Vector graphics / logos | Not rasterized | Planned: SkiaSharp path renderer (Phase 5.2) |
| Right-to-left / bidirectional text | Not supported | No workaround |
| PDF forms (AcroForms) | Fields not extracted | No workaround |
| Layered PDFs (Optional Content Groups) | All layers merged | Set desired layer visibility before export in Acrobat |
| Password-protected PDFs | Not supported | Decrypt before conversion |
| OCR back-end | Stub only — returns empty list | Subclass `OcrTextExtractor` with Tesseract.NET |
