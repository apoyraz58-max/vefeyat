#requires -Version 5.1
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$progId = 'Vefeyat.Connect'
$classId = '{2B0AF7B7-BD0D-4C70-9136-BED2DBD8F8F3}'
$bases = @()
try {
    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
        throw 'Önce bütün Word pencerelerini kapatın; sonra Kaldir.cmd dosyasını yeniden çalıştırın.'
    }
    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { throw 'Windows yerel uygulama veri klasörü bulunamadı.' }
    $installDirectory = Join-Path $env:LOCALAPPDATA 'Vefeyat\Addin'
    $views = @([Microsoft.Win32.RegistryView]::Registry32)
    if ([Environment]::Is64BitOperatingSystem) { $views += [Microsoft.Win32.RegistryView]::Registry64 }
    # Validate both views before deleting anything.
    foreach ($view in $views) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        $bases += $base
        $key = $base.OpenSubKey("Software\Classes\CLSID\$classId\InprocServer32")
        if ($null -ne $key) {
            try {
                if ($key.GetValue('Class', '') -cne 'Vefeyat.Connect') { throw 'COM kaydı başka bir uygulamaya ait; otomatik kaldırma durduruldu.' }
            } finally { $key.Dispose() }
        }
        $key = $base.OpenSubKey("Software\Classes\$progId\CLSID")
        if ($null -ne $key) {
            try {
                if ($key.GetValue('', '') -ine $classId) { throw 'Program kimliği başka bir COM sınıfına bağlı; otomatik kaldırma durduruldu.' }
            } finally { $key.Dispose() }
        }
    }
    foreach ($base in $bases) {
        foreach ($path in @("Software\Microsoft\Office\Word\Addins\$progId", "Software\Classes\$progId", "Software\Classes\CLSID\$classId")) { $base.DeleteSubKeyTree($path, $false) }
    }
    # Remove only the two installed files. Leave unknown files and all user data.
    foreach ($relative in @('Vefeyat.WordAddin.dll', 'data\people.csv')) {
        $path = Join-Path $installDirectory $relative
        if (Test-Path -LiteralPath $path -PathType Leaf) { Remove-Item -LiteralPath $path -Force }
    }
    foreach ($path in @((Join-Path $installDirectory 'data'), $installDirectory)) {
        if ((Test-Path -LiteralPath $path -PathType Container) -and @(Get-ChildItem -LiteralPath $path -Force).Count -eq 0) { Remove-Item -LiteralPath $path -Force }
    }
    Write-Host 'Vefeyât eklentisi kaldırıldı.'
    Write-Host ('Kişisel verileriniz korundu: ' + (Join-Path $env:LOCALAPPDATA 'Vefeyat\Data'))
} catch {
    Write-Host ('Kaldırma tamamlanamadı: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
} finally {
    foreach ($base in $bases) { $base.Dispose() }
}
