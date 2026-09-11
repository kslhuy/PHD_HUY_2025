param([string]$InputDocx,[string]$OutputPdf,[switch]$Save)
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$originalPrinter = $word.ActivePrinter
$word.ActivePrinter = 'Microsoft Print to PDF'
$doc = $null
try {
    Write-Output 'Opening document'
    $doc = $word.Documents.Open($InputDocx, $false, (-not $Save.IsPresent))
    Write-Output 'Document opened'
    if ($Save) {
        Write-Output 'Updating fields'
        $doc.Fields.Update() | Out-Null
        Write-Output 'Updating contents'
        foreach ($toc in $doc.TablesOfContents) { $toc.Update() }
        Write-Output 'Repagination'
        $doc.Repaginate()
        Write-Output 'Saving document'
        $doc.Save()
    }
    Write-Output 'Exporting PDF'
    $doc.ExportAsFixedFormat($OutputPdf,17)
    Write-Output ('Pages: ' + $doc.ComputeStatistics(2))
} finally {
    if ($null -ne $doc) { $doc.Close(0); [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null }
    $word.ActivePrinter = $originalPrinter
    $word.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}
