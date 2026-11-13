#requires -version 5.1
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [string]$DotmPath,
    [string]$StartupPath = (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'),
    [string]$ExpectedSha256,
    [switch]$GeriAl
)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Bu arac Windows Word icindir.' }
if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    throw 'Calismanizi kaydedip Word un BUTUN pencerelerini kapatin. Arac Word u zorla kapatmaz.'
}
function Assert-LocalPath([string]$Path) {
    if ($Path -notmatch '^[a-zA-Z]:\\' -or $Path.StartsWith('\\')) { throw 'Yerel bir Windows yolu gerekli.' }
    $drive = New-Object System.IO.DriveInfo ($Path.Substring(0,3))
    if ($drive.DriveType -eq [System.IO.DriveType]::Network) { throw 'Ag surucusu kullanilamaz.' }
}
Assert-LocalPath $StartupPath
$StartupPath = [System.IO.Path]::GetFullPath($StartupPath)
$target = Join-Path $StartupPath 'VefeyatEklenti.dotm'
$backupRoot = Join-Path $env:LOCALAPPDATA 'Vefeyat\WordVBA\InstallerBackups'
$statePath = Join-Path $backupRoot 'last-install.json'
if ($GeriAl) {
    if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { throw 'Geri donus kaydi yok.' }
    $state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($state.Target -ne $target) { throw 'Startup yolu onceki kurulumla ayni olmali.' }
    $backup = [System.IO.Path]::GetFullPath([string]$state.Backup)
    if (-not $backup.StartsWith($backupRoot + '\',[StringComparison]::OrdinalIgnoreCase)) { throw 'Gecersiz yedek yolu.' }
    if (-not (Test-Path -LiteralPath $backup -PathType Leaf)) { throw 'Onceki surum yedegi bulunamadi.' }
    if ((Get-FileHash -LiteralPath $backup -Algorithm SHA256).Hash -ne $state.BackupHash) { throw 'Yedek dogrulanamadi.' }
    if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $state.InstalledHash) { throw 'Kurulumdan sonra hedef degismis; ustune yazilmadi.' }
    $DotmPath = $backup
} else {
    if (-not $DotmPath) {
        $DotmPath = Join-Path $env:LOCALAPPDATA 'Vefeyat\WordVBA\Guncelleme\VefeyatEklenti.dotm'
    }
}
Assert-LocalPath $DotmPath
$DotmPath = [System.IO.Path]::GetFullPath($DotmPath)
if ($DotmPath -eq $target) { throw 'Kaynak ve Startup hedefi farkli olmali.' }
if (-not (Test-Path -LiteralPath $DotmPath -PathType Leaf)) { throw 'Once hazirlik sablonunda Vefeyat0341Kur ile guncelleme DOTM dosyasini olusturun; gerekirse -DotmPath verin.' }
$sourceHash = (Get-FileHash -LiteralPath $DotmPath -Algorithm SHA256).Hash
if ($ExpectedSha256 -and $sourceHash -ne $ExpectedSha256) { throw 'Beklenen SHA256 ile uyusmuyor.' }
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($DotmPath)
try {
    if (-not $zip.GetEntry('word/vbaProject.bin')) { throw 'Makro projesi bulunamadi.' }
    $ribbon = $zip.GetEntry('customUI/customUI14.xml')
    if (-not $ribbon) { throw 'Vefeyat seridi bulunamadi.' }
    $reader = New-Object System.IO.StreamReader ($ribbon.Open())
    try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if (-not $GeriAl -and $xml.DocumentElement.onLoad -ne 'VefeyatEklenti0341.Vefeyat0341UiOnLoad') {
        throw 'Hazir DOTM 0.3.41 seridini icermiyor. Yanlis surum kopyalanmadi.'
    }
} finally { $zip.Dispose() }
$duplicates = @()
if (Test-Path -LiteralPath $StartupPath -PathType Container) {
    $duplicates = @(Get-ChildItem -LiteralPath $StartupPath -Filter '*Vefeyat*.dotm' -File | Where-Object { $_.FullName -ne $target })
}
if ($duplicates.Count) { throw 'Startup ta birden fazla Vefeyat kopyasi var. Kopyalari inceleyip fazlaligi Startup disina tasiyin; arac hicbirini silmedi.' }
if (-not $PSCmdlet.ShouldProcess($target, 'Dogrulanmis Vefeyat kopyasini kur')) { return }
New-Item -ItemType Directory -Path $StartupPath -Force | Out-Null
New-Item -ItemType Directory -Path $backupRoot -Force | Out-Null
$backupPath = ''
$backupHash = ''
if (Test-Path -LiteralPath $target -PathType Leaf) {
    $backupPath = Join-Path $backupRoot ((Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N') + '.dotm')
    Copy-Item -LiteralPath $target -Destination $backupPath
    $backupHash = (Get-FileHash -LiteralPath $backupPath -Algorithm SHA256).Hash
    if ($backupHash -ne (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash) { throw 'Onceki surum yedegi dogrulanamadi.' }
}
$stage = Join-Path $StartupPath ([guid]::NewGuid().ToString('N') + '.tmp')
try {
    Copy-Item -LiteralPath $DotmPath -Destination $stage
    if ((Get-FileHash -LiteralPath $stage -Algorithm SHA256).Hash -ne $sourceHash) { throw 'Kopya dogrulanamadi.' }
    if (Get-Process WINWORD -ErrorAction SilentlyContinue) { throw 'Islem sirasinda Word acildi; kurulum durduruldu.' }
    if (Test-Path -LiteralPath $target -PathType Leaf) {
        if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $backupHash) { throw 'Hedef degisti; ustune yazilmadi.' }
        [System.IO.File]::Replace($stage,$target,$null)
    } else {
        [System.IO.File]::Move($stage,$target)
    }
    if (-not $GeriAl) {
        @{ Target=$target; Backup=$backupPath; BackupHash=$backupHash; InstalledHash=$sourceHash; Version='0.3.41' } |
            ConvertTo-Json | Set-Content -LiteralPath $statePath -Encoding UTF8
    }
    Write-Output 'Kurulum tamamlandi. Word u acip Veri bilgisi surumunu kontrol edin.'
    Write-Output ('SHA256: ' + $sourceHash)
} finally {
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage }
}
# No execution-policy changes, registry edits, VBProject access, downloads,
# process termination, CSV writes or personal-document operations occur here.
