<#
.SYNOPSIS
Makes the public copies of the workshop slide decks: a cleaned .pptx and a
tagged PDF for each deck, with speaker notes, comments, and hidden slides
removed.

.DESCRIPTION
For each original deck listed in $Decks, the script copies the file into a
scratch folder (_work\ under the repository root), opens the copy in PowerPoint
with no visible window, and then:

  1. deletes every hidden slide;
  2. removes comments and personal information (Presentation.RemoveDocumentInformation);
  3. clears the speaker notes on every slide, because PpRemoveDocInfoType has no
     member for notes;
  4. saves the cleaned deck as slides\<stem>.pptx;
  5. exports slides\<stem>.pdf with document structure tags on and hidden slides off.

The originals are never opened for writing. The scratch copy is deleted when the
deck is done, and the scratch folder is removed at the end.

PowerPoint must be installed (the script drives it through COM). After it runs,
use tools\check_public_decks.py to prove the public copies are clean without
relying on PowerPoint, and tools\pptx_to_html.py to write the HTML transcripts.

.PARAMETER SourceDir
Folder that holds the original decks named in $Decks.

.PARAMETER RepoRoot
Repository root. Defaults to the parent of the folder this script lives in.

.EXAMPLE
powershell -ExecutionPolicy Bypass -File tools\make_public_decks.ps1 -SourceDir "C:\Users\gypin\ClaudeDesktop\projects\gsa-workshop"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceDir,

    [string]$RepoRoot
)

$ErrorActionPreference = 'Stop'

# Default the repository root to the parent of this script's folder. This is
# resolved here rather than in the param block because $PSScriptRoot is empty
# while PowerShell 5.1 binds parameters.
if (-not $RepoRoot) {
    $RepoRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

# Original filename -> public filename stem. Edit this table for a new workshop.
$Decks = [ordered]@{
    '1_Welcome_and_Overview_NaNDA-GSA2026.pptx'  = '01_Welcome_and_Overview'
    '2_Clarke.pptx'                              = '02_Neighborhood_Environment_and_Aging_with_Disability'
    '3_Noppert.pptx'                             = '03_NaNDA_and_Immune_Aging'
    '4_Choosing_Your_First_NaNDA_Dataset.pptx'   = '04_Choosing_Your_First_NaNDA_Dataset'
    '5_Reconvene_and_Wrap-Up_NaNDA-GSA2026.pptx' = '05_Reconvene_and_Wrap-Up'
}

# Office object model constants. Values checked against the Microsoft VBA reference.
$msoTrue  = -1
$msoFalse = 0
$ppAlertsNone = 1                      # PpAlertLevel
$ppSaveAsOpenXMLPresentation = 24      # PpSaveAsFileType
$ppSaveAsPDF = 32                      # PpSaveAsFileType
$ppRDIComments = 1                     # PpRemoveDocInfoType
$ppRDIRemovePersonalInformation = 4    # PpRemoveDocInfoType
$msoPlaceholder = 14                   # MsoShapeType
$ppPlaceholderSlideNumber = 13         # PpPlaceholderType

$SlidesDir = Join-Path $RepoRoot 'slides'
$WorkDir   = Join-Path $RepoRoot '_work'
New-Item -ItemType Directory -Force -Path $SlidesDir | Out-Null
New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null

Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Read-ZipEntryText {
    param($Entry)
    $stream = $Entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        return $reader.ReadToEnd()
    }
    finally { $stream.Dispose() }
}

function Write-ZipEntryText {
    param($Entry, [string]$Text)
    $stream = $Entry.Open()
    try {
        $stream.SetLength(0)
        $writer = New-Object System.IO.StreamWriter($stream, (New-Object System.Text.UTF8Encoding($false)))
        $writer.Write($Text)
        $writer.Flush()
    }
    finally { $stream.Dispose() }
}

function Remove-CommentParts {
    <#
    Deletes any comment or comment-author part left in a saved .pptx, along with
    the relationships and content-type overrides that point at them. Returns the
    names of the parts removed. A .pptx is a ZIP package, so this edits the
    package directly and touches nothing else.
    #>
    param([Parameter(Mandatory = $true)][string]$PptxPath)

    $partPattern = '^ppt/(authors\.xml|commentAuthors\.xml|comments/.+)$'
    $relPattern  = '<Relationship\b[^>]*Type="[^"]*/(comments|commentAuthors|authors)"[^>]*/>'
    $typePattern = '<Override\b[^>]*PartName="/ppt/(authors\.xml|commentAuthors\.xml|comments/[^"]+)"[^>]*/>'

    $zip = [System.IO.Compression.ZipFile]::Open($PptxPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $removed = @()
        foreach ($entry in @($zip.Entries)) {
            if ($entry.FullName -match $partPattern) {
                $removed += $entry.FullName
                $entry.Delete()
            }
        }
        if ($removed.Count -eq 0) { return @() }

        foreach ($entry in @($zip.Entries)) {
            if ($entry.FullName -like '*.rels') {
                $xml = Read-ZipEntryText $entry
                $new = [regex]::Replace($xml, $relPattern, '')
                if ($new -ne $xml) { Write-ZipEntryText $entry $new }
            }
        }

        $contentTypes = $zip.GetEntry('[Content_Types].xml')
        $xml = Read-ZipEntryText $contentTypes
        $new = [regex]::Replace($xml, $typePattern, '')
        if ($new -ne $xml) { Write-ZipEntryText $contentTypes $new }

        return $removed
    }
    finally { $zip.Dispose() }
}

$ppt = $null

try {
    $ppt = New-Object -ComObject PowerPoint.Application
    $ppt.DisplayAlerts = $ppAlertsNone

    foreach ($entry in $Decks.GetEnumerator()) {
        $originalName = $entry.Key
        $stem = $entry.Value
        $original = Join-Path $SourceDir $originalName
        if (-not (Test-Path -LiteralPath $original)) { throw "Missing original: $original" }

        # Work on a copy so the original is never opened for writing.
        $workCopy = Join-Path $WorkDir $originalName
        Copy-Item -LiteralPath $original -Destination $workCopy -Force
        Set-ItemProperty -LiteralPath $workCopy -Name IsReadOnly -Value $false

        Write-Host "== $originalName -> $stem"
        # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
        $pres = $ppt.Presentations.Open($workCopy, $msoFalse, $msoFalse, $msoFalse)
        try {
            $before = $pres.Slides.Count

            # 1. Delete hidden slides, last to first so the indexes stay valid.
            $hiddenDeleted = 0
            for ($i = $pres.Slides.Count; $i -ge 1; $i--) {
                $slide = $pres.Slides.Item($i)
                if ($slide.SlideShowTransition.Hidden -eq $msoTrue) {
                    Write-Host "   deleting hidden slide $i"
                    $slide.Delete()
                    $hiddenDeleted++
                }
            }

            # 2. Remove comments and personal information.
            $pres.RemoveDocumentInformation($ppRDIComments)
            $pres.RemoveDocumentInformation($ppRDIRemovePersonalInformation)

            # 3. Clear speaker notes: every text-bearing shape on each notes page
            #    except the slide number placeholder.
            $notesCleared = 0
            foreach ($slide in $pres.Slides) {
                foreach ($shape in $slide.NotesPage.Shapes) {
                    if (-not $shape.HasTextFrame) { continue }
                    if ($shape.Type -eq $msoPlaceholder -and $shape.PlaceholderFormat.Type -eq $ppPlaceholderSlideNumber) { continue }
                    if ($shape.TextFrame.HasText -eq $msoTrue) {
                        $shape.TextFrame.TextRange.Text = ''
                        $notesCleared++
                    }
                }
            }

            # 4. Save the public copy.
            $publicPptx = Join-Path $SlidesDir "$stem.pptx"
            if (Test-Path -LiteralPath $publicPptx) { Remove-Item -LiteralPath $publicPptx -Force }
            $pres.SaveAs($publicPptx, $ppSaveAsOpenXMLPresentation)

            # 5. Export the PDF. SaveAs with the PDF file type uses PowerPoint's
            #    default PDF options, which include document structure tags, and
            #    the hidden slides are already gone. (ExportAsFixedFormat, which
            #    exposes those options explicitly, cannot be called from
            #    PowerShell: the COM interop treats it as a property.)
            #    tools\check_public_decks.py verifies that the output is tagged.
            $publicPdf = Join-Path $SlidesDir "$stem.pdf"
            if (Test-Path -LiteralPath $publicPdf) { Remove-Item -LiteralPath $publicPdf -Force }
            $pres.SaveAs($publicPdf, $ppSaveAsPDF)

            Write-Host ("   slides {0} -> {1}; hidden deleted {2}; notes shapes cleared {3}" -f `
                $before, $pres.Slides.Count, $hiddenDeleted, $notesCleared)
        }
        finally {
            $pres.Close()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
            $pres = $null
        }
        Remove-Item -LiteralPath $workCopy -Force

        # 6. RemoveDocumentInformation deletes the comments but can leave the
        #    (anonymized) comment author list behind. Strip any such part from
        #    the saved package so the public copy carries no comment machinery.
        $stripped = Remove-CommentParts -PptxPath $publicPptx
        if ($stripped.Count -gt 0) {
            Write-Host "   removed leftover comment parts: $($stripped -join ', ')"
        }
    }
}
finally {
    if ($ppt) {
        $ppt.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
        $ppt = $null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $WorkDir) {
        Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "Done. Public copies are in $SlidesDir. Next: python tools\check_public_decks.py"
