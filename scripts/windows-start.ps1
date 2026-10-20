[CmdletBinding()]
param(
    [string]$DataFile = ""
)

$ErrorActionPreference = "Stop"
$repoDir = Split-Path -Parent $PSScriptRoot
Set-Location -LiteralPath $repoDir

if (-not (Get-Command docker -ErrorAction SilentlyContinue)) {
    throw "Docker Desktop kurulu değil veya docker komutu bulunamıyor. Docker Desktop'ı Linux containers modunda başlatın."
}
& docker info --format '{{.OSType}}' | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Docker Desktop çalışmıyor. Linux containers motorunu başlatıp yeniden deneyin."
}
$dockerOS = & docker info --format '{{.OSType}}'
if ($dockerOS -ne "linux") {
    throw "Docker Desktop Linux containers modunda çalışmalıdır."
}
& docker compose version | Out-Null
if ($LASTEXITCODE -ne 0) {
    throw "Docker Compose v2 gerekli. Docker Desktop kurulumunu kontrol edin."
}

$localDir = Join-Path $repoDir ".local"
$sourceDir = Join-Path $localDir "source"
$sourceFile = Join-Path $sourceDir "people.xlsx"
New-Item -ItemType Directory -Path $sourceDir -Force | Out-Null
if ($DataFile) {
    $incoming = (Resolve-Path -LiteralPath $DataFile).Path
    if (-not (Test-Path -LiteralPath $incoming -PathType Leaf)) {
        throw "-DataFile geçerli bir Excel dosyasını göstermelidir."
    }
    if (Test-Path -LiteralPath $sourceFile) {
        if ((Get-FileHash -LiteralPath $incoming -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $sourceFile -Algorithm SHA256).Hash) {
            throw ".local/source/people.xlsx zaten var ve farklı içerikte. Mevcut dosya korundu; değiştirmeden önce yedeğini alın."
        }
    } else {
        Copy-Item -LiteralPath $incoming -Destination $sourceFile
    }
}
if (-not (Test-Path -LiteralPath $sourceFile -PathType Leaf)) {
    throw "Excel dosyasını .local/source/people.xlsx olarak koyun veya -DataFile 'C:\dosyalar\Sahislar.xlsx' parametresiyle çalıştırın."
}

function New-RandomHex([int]$ByteCount) {
    $bytes = New-Object byte[] $ByteCount
    $random = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    try { $random.GetBytes($bytes) } finally { $random.Dispose() }
    return [BitConverter]::ToString($bytes).Replace("-", "").ToLowerInvariant()
}

$envFile = Join-Path $localDir "windows.env"
if (-not (Test-Path -LiteralPath $envFile)) {
    $content = "POSTGRES_PASSWORD=$(New-RandomHex 32)`nDJANGO_SECRET_KEY=$(New-RandomHex 64)`n"
    $stream = [System.IO.File]::Open($envFile, [System.IO.FileMode]::CreateNew, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $bytes = [System.Text.UTF8Encoding]::new($false).GetBytes($content)
        $stream.Write($bytes, 0, $bytes.Length)
    } finally {
        $stream.Dispose()
        $content = $null
    }
}
$envLines = Get-Content -LiteralPath $envFile
if (-not ($envLines -match '^POSTGRES_PASSWORD=[0-9a-f]{64}$') -or -not ($envLines -match '^DJANGO_SECRET_KEY=[0-9a-f]{128}$')) {
    throw ".local/windows.env beklenen yerel kurulum biçiminde değil; parolalar yazdırılmadan işlem durduruldu."
}
$envLines = $null

# Limit the credential file to this Windows user and the operating system.
$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent().User
$acl = [System.Security.AccessControl.FileSecurity]::new()
$acl.SetOwner($identity)
$acl.SetAccessRuleProtection($true, $false)
$acl.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new($identity, "FullControl", "Allow"))
$systemSid = [System.Security.Principal.SecurityIdentifier]::new("S-1-5-18")
$acl.AddAccessRule([System.Security.AccessControl.FileSystemAccessRule]::new($systemSid, "FullControl", "Allow"))
Set-Acl -LiteralPath $envFile -AclObject $acl

& docker compose --env-file $envFile -f compose.windows.yml config --quiet
if ($LASTEXITCODE -ne 0) { throw "Docker Compose yapılandırması doğrulanamadı." }
& docker compose --env-file $envFile -f compose.windows.yml up --detach --build --wait --wait-timeout 240
if ($LASTEXITCODE -ne 0) {
    throw "Yerel servisler başlatılamadı. docs/windows-local.md içindeki durum ve günlük komutlarını inceleyin."
}

Write-Host "Yerel servisler başlatıldı. Önce docs/windows-local.md içindeki sertifika inceleme ve güven adımlarını tamamlayın."
Write-Host "Ardından https://localhost:8443 adresini açın. Tarayıcı otomatik açılmadı; güven ayarı otomatik değiştirilmedi."
