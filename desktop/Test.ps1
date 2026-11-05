#requires -Version 5.1
param(
    [string] $CsvFile
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$temporary = $null

try {
    # Test fixtures use temporary directories. This script never starts Word,
    # registers COM classes, or opens the user's installed catalog.
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw 'Bu test komutu Windows ve yerleşik .NET Framework 4.8 derleyicisi gerektirir.'
    }
    $sourceNames = @('Core.cs', 'CatalogForm.cs', 'Connect.cs', 'SelectionRules.cs')
    $testNames = @('CoreTests', 'SelectionTests', 'ComContractTests')
    $sources = @()
    foreach ($source in $sourceNames) {
        $path = Join-Path $PSScriptRoot $source
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Kaynak dosya bulunamadı: $source" }
        $sources += $path
    }
    foreach ($test in $testNames) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot ('tests\' + $test + '.cs')) -PathType Leaf)) { throw "Test dosyası bulunamadı: $test.cs" }
    }
    $resolvedCsv = $null
    if (-not [string]::IsNullOrWhiteSpace($CsvFile)) {
        $resolvedCsv = (Resolve-Path -LiteralPath $CsvFile).ProviderPath
        if (-not (Test-Path -LiteralPath $resolvedCsv -PathType Leaf)) { throw 'CsvFile bir dosyayı göstermelidir.' }
    }
    $frameworkDirectories = @((Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319'))
    if ([Environment]::Is64BitOperatingSystem) { $frameworkDirectories += Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319' }
    $compiler = $null
    foreach ($directory in $frameworkDirectories) {
        $candidate = Join-Path $directory 'csc.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $compiler = $candidate; break }
    }
    if (-not $compiler) { throw '.NET Framework 4.8 C# derleyicisi bulunamadı.' }
    $temporary = Join-Path ([IO.Path]::GetTempPath()) ('Vefeyat-tests-' + [Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $temporary)
    $commonArguments = @('/nologo', '/platform:anycpu', '/optimize+', '/langversion:5', '/codepage:65001', '/utf8output')
    foreach ($reference in @('System.dll', 'System.Core.dll', 'System.Xml.dll', 'System.Windows.Forms.dll', 'System.Drawing.dll', 'Microsoft.CSharp.dll')) { $commonArguments += '/reference:' + $reference }

    $libraryPath = Join-Path $temporary 'Vefeyat.WordAddin.dll'
    $arguments = $commonArguments + @('/target:library', ('/out:' + $libraryPath)) + $sources
    $compilerOutput = & $compiler @arguments 2>&1
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $libraryPath -PathType Leaf)) {
        throw ('Eklenti C# 5 biçiminde derlenemedi.' + [Environment]::NewLine + ($compilerOutput -join [Environment]::NewLine))
    }
    Write-Host 'Eklenti: .NET Framework ve C# 5 derlemesi başarılı.'
    foreach ($test in $testNames) {
        $testPath = Join-Path $temporary ($test + '.exe')
        $testSource = Join-Path $PSScriptRoot ('tests\' + $test + '.cs')
        $arguments = $commonArguments + @('/target:exe', ('/main:' + $test), ('/out:' + $testPath)) + $sources + @($testSource)
        $compilerOutput = & $compiler @arguments 2>&1
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $testPath -PathType Leaf)) {
            throw ($test + ' derlenemedi.' + [Environment]::NewLine + ($compilerOutput -join [Environment]::NewLine))
        }
        Write-Host ('Çalıştırılıyor: ' + $test)
        $testArguments = @()
        if ($test -eq 'CoreTests' -and $null -ne $resolvedCsv) { $testArguments += $resolvedCsv }
        & $testPath @testArguments
        if ($LASTEXITCODE -ne 0) { throw ($test + ' başarısız; çıkış kodu: ' + $LASTEXITCODE) }
    }
    $installerTests = Join-Path $temporary 'InstallerPlanTests.exe'
    $installerArguments = $commonArguments + @('/target:exe', '/main:InstallerPlanTests', ('/out:' + $installerTests),
        (Join-Path $PSScriptRoot 'Installer.cs'), (Join-Path $PSScriptRoot 'tests\InstallerPlanTests.cs'))
    $compilerOutput = & $compiler @installerArguments 2>&1
    if ($LASTEXITCODE -ne 0) { throw ('Kurucu kayıt planı testleri derlenemedi.' + [Environment]::NewLine + ($compilerOutput -join [Environment]::NewLine)) }
    & $installerTests
    if ($LASTEXITCODE -ne 0) { throw 'Kurucu kayıt planı testleri başarısız.' }
    Write-Host 'Dört test grubu başarılı. Word uygulaması başlatılmadı; kurulu katalog değiştirilmedi.'
    if ($null -ne $resolvedCsv) { Write-Host 'Ek CSV denetimi, bu projedeki 7.934 kayıtlı özgün veri dosyasını bekler.' }
} catch {
    Write-Host ('Testler tamamlanamadı: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
} finally {
    if ($null -ne $temporary -and (Test-Path -LiteralPath $temporary)) { Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue }
}
