# Exports each DOCX of a directory to a PDF with Microsoft Word for Windows, as the references of the corpus.
# Usage: powershell -ExecutionPolicy Bypass -File export.ps1 -Source corpus\docx -Target corpus\reference\windows
param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Target
)
$ErrorActionPreference = "Stop"
$wdExportFormatPDF = 17
$wdDoNotSaveChanges = 0
$source = (Resolve-Path $Source).Path
New-Item -ItemType Directory -Force -Path $Target | Out-Null
$target = (Resolve-Path $Target).Path

$word = New-Object -ComObject Word.Application
$word.Visible = $false
# No dialog: the fields are updated below, as Word for Mac does when it asks
$word.DisplayAlerts = 0
try {
    foreach ($file in Get-ChildItem -Path $source -Filter *.docx | Sort-Object Name) {
        $out = Join-Path $target ($file.BaseName + ".pdf")
        if (Test-Path $out) { continue }
        # ConfirmConversions false, ReadOnly true, AddToRecentFiles false
        $document = $word.Documents.Open($file.FullName, $false, $true, $false)
        try {
            # Word for Mac answers Yes to updating the fields and the whole table of contents when it opens a document of pandoc-service
            foreach ($toc in $document.TablesOfContents) { $toc.Update() }
            foreach ($tof in $document.TablesOfFigures) { $tof.Update() }
            $document.Fields.Update() | Out-Null
            $document.ExportAsFixedFormat($out, $wdExportFormatPDF)
            Write-Output "$($file.BaseName): ok"
        } catch {
            Write-Output "$($file.BaseName): $($_.Exception.Message)"
        } finally {
            $document.Close($wdDoNotSaveChanges)
        }
    }
} finally {
    $word.Quit()
}
