#requires -version 5.1
[CmdletBinding()]
param([string]$ReportPath = (Join-Path $PSScriptRoot 'Word-kabul-sonucu.txt'))
$ErrorActionPreference = 'Stop'
if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Once tum Word pencerelerini kaydedip kapatin.' }
if (Test-Path -LiteralPath $ReportPath) { throw 'Rapor dosyasi var. Yeni bir -ReportPath secin.' }
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    # Executes the installed add-in only. The macro creates/discards its own
    # synthetic document. It never imports catalog data or edits a thesis.
    $result = $word.Run('Vefeyat0341UiNativeChecks')
    [string]$result | Set-Content -LiteralPath $ReportPath -Encoding UTF8
    Write-Output $result
    if (-not ([string]$result).StartsWith('PASS ')) { throw 'Word kabul testi basarisiz; raporu inceleyin.' }
} finally {
    if ($word) {
        if ($word.Documents.Count -eq 0) { $word.Quit() }
        # If an unrelated document appeared, leave Word open for the author.
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
}
