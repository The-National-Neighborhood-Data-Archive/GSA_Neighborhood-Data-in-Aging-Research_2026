# Tools

These scripts made the public copies of the slide decks, the PDFs, and the HTML transcripts in the slides folder. They are kept here so the same package can be rebuilt for the next workshop, or rebuilt now if a deck changes. Like everything else in this repository, they are released under the Creative Commons Attribution 4.0 International license in the repository root.

## What each script does

- **make_public_decks.ps1** (PowerShell; needs PowerPoint installed). For each original deck it copies the file into a scratch folder, opens the copy in PowerPoint with no visible window, deletes every hidden slide, removes comments and personal information, clears the speaker notes on every slide, saves the cleaned deck as `slides\<stem>.pptx`, exports `slides\<stem>.pdf`, and strips any leftover comment-author part from the saved package. The originals are never opened for writing. The mapping from original filename to public stem is the `$Decks` table at the top of the script.
- **pptx_to_html.py** (Python; needs python-pptx). Reads the public `.pptx` files, never the originals, and writes one self-contained HTML transcript per deck: a skip link, a header, a navigation list with one link per slide, and one section per slide with the slide's text in reading order, nested lists, real tables, hyperlinks, and image descriptions written out as text. The deck titles and presenters are the `DECKS` table at the top of the script.
- **check_public_decks.py** (Python; needs python-pptx and pymupdf). Proves the public copies are clean and the transcripts are complete, without PowerPoint. The expected public slide counts are the `EXPECTED_SLIDES` table at the top of the script.

## Rerunning them

1. Install the Python dependencies:

   ```
   pip install python-pptx pymupdf
   ```

2. If the decks or their slide counts changed, update the three tables named above.

3. From the repository root, in PowerShell:

   ```
   powershell -ExecutionPolicy Bypass -File tools\make_public_decks.ps1 -SourceDir "C:\path\to\the\original\decks"
   python tools\pptx_to_html.py
   python tools\check_public_decks.py
   ```

4. Open each public `.pptx` in PowerPoint once and confirm that no repair prompt appears.

The public copies for the 2026 workshop were made from the originals in `C:\Users\gypin\ClaudeDesktop\projects\gsa-workshop`.

## Outcome checks

`check_public_decks.py` exits with status 0 only when every one of these passes:

- Each public `.pptx` has no comment parts, comment relationships, or comment content types; no speaker-notes text (every text run in the notes parts is empty, a slide number, or the slide-number field marker); no hidden slides; and the expected slide count, read both from the ZIP and by opening the file with python-pptx.
- Each PDF has a page count equal to the public slide count, selectable text (the script prints the first line of text on page 4), and document structure tags (`/MarkInfo` with `/Marked true` and a `/StructTreeRoot`).
- Each HTML transcript has one slide heading per public slide and no empty headings; `lang="en"`, a title, a skip link to `#main`, and a `<nav>`; every relative link resolves to a file in the repository; and every hyperlink found in the public `.pptx` appears in the HTML.

For the 2026 workshop the expected public slide counts are 28, 16, 19, 34, and 9. The third deck (NaNDA and Immune Aging) has one fewer slide than the original because its hidden slide 13 was removed.

## Things to know

- **PDF export.** The script exports the PDF with `SaveAs` and the PDF file type. That uses PowerPoint's default PDF options, which include document structure tags; the checker verifies the tags are there. The `ExportAsFixedFormat` method, which exposes those options explicitly, cannot be called from PowerShell because the COM interop treats it as a property.
- **PDF titles.** A PDF's title metadata comes from the deck's Title document property. The script does not change document properties, so a deck with no Title, or a placeholder such as "PptxGenJS Presentation", carries that over. Set the Title property in the original deck before rerunning if a better PDF title is wanted.
- **Reading order.** The transcript reads the title first, then the remaining shapes in rows from top to bottom, and left to right within a row. Shapes whose top edges are within a quarter inch are treated as one row. On a slide where visual layout carries meaning, read the transcript once after generating it.
- **Alt text.** Pictures marked decorative, or with no alt text, are skipped. Alt text that is only a filename is treated as missing. The generator prints every missing description with its slide number so it can be fixed in the original deck.
