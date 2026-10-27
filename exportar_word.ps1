param([string]$Documento,[string]$Pdf)
$ErrorActionPreference = 'Stop'
$simplexWord = $null
$simplexDoc = $null
try {
    $simplexWord = New-Object -ComObject Word.Application
    $simplexWord.Visible = $false
    $simplexWord.ScreenUpdating = $false
    $simplexWord.DisplayAlerts = 0
    $simplexDoc = $simplexWord.Documents.Open($Documento, $false, $true)
    $simplexDoc.Repaginate()
    $simplexDoc.ExportAsFixedFormat($Pdf, 17)
    Write-Output ('Paginas: ' + $simplexDoc.ComputeStatistics(2))
} finally {
    if ($null -ne $simplexDoc) { $simplexDoc.Close(0); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($simplexDoc) }
    if ($null -ne $simplexWord) { $simplexWord.Quit(); [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($simplexWord) }
}
