#requires -Version 5.1
[CmdletBinding()]
param(
    [string] $SigningCertificateThumbprint,
    [string] $TimestampUrl,
    [string] $OutDirectory,
    [string] $CsvFile,
    [string] $SigntoolPath
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$releaseStage = $null
$pendingOutput = $null

function Invoke-CheckedTool([string] $ToolPath, [string[]] $Arguments, [string] $FailureMessage) {
    $toolOutput = & $ToolPath @Arguments 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw ($FailureMessage + [Environment]::NewLine + ($toolOutput -join [Environment]::NewLine))
    }
}

function Test-WithinDirectory([string] $Path, [string] $Directory) {
    $directoryRoot = [IO.Path]::GetFullPath($Directory).TrimEnd([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
    return $Path.Equals($directoryRoot, [StringComparison]::OrdinalIgnoreCase) -or
        $Path.StartsWith($directoryRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)
}

function Get-CsvRecordCount([string] $Path) {
    # Decode strictly as UTF-8 before using the built-in quoted-field parser.
    $decoder = New-Object Text.UTF8Encoding($false, $true)
    $text = $decoder.GetString([IO.File]::ReadAllBytes($Path))
    if ($text.Length -gt 0 -and $text[0] -eq [char]0xFEFF) { $text = $text.Substring(1) }
    if ($text.IndexOf([char]0) -ge 0) { throw 'CSV NUL karakteri içeremez.' }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $reader = New-Object IO.StringReader($text)
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($reader)
    try {
        $parser.SetDelimiters([string[]]@(';'))
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        $expected = @('Kısa Ad', 'Tam Ad', 'Orijinal Ad', 'Vefat Tarihi', 'Tanıtım', 'Kaynak Link')
        $headers = $parser.ReadFields()
        if ($null -eq $headers -or ($headers -join '|') -cne ($expected -join '|')) {
            throw 'CSV, beklenen altı başlığı aynı sırada içermelidir; ayraç noktalı virgül olmalıdır.'
        }
        $count = 0
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields()
            if ($null -eq $fields) { continue }
            $hasContent = $false
            foreach ($field in $fields) { if (-not [string]::IsNullOrWhiteSpace($field)) { $hasContent = $true } }
            if (-not $hasContent) { continue }
            if ($fields.Length -ne 6) { throw 'CSV kaydı altı alan içermelidir.' }
            if ([string]::IsNullOrWhiteSpace($fields[0]) -or $fields[0].Trim() -eq '---') { throw 'CSV kaydında Kısa Ad boş olamaz.' }
            foreach ($field in $fields) { [void][Xml.XmlConvert]::VerifyXmlChars($field) }
            $count++
        }
        if ($count -eq 0) { throw 'CSV en az bir şahıs kaydı içermelidir.' }
        return $count
    } finally {
        $parser.Dispose()
        $reader.Dispose()
    }
}

function Confirm-Authenticode([string] $Path, [string] $ExpectedThumbprint, [string] $SignTool) {
    Invoke-CheckedTool $SignTool @('verify', '/pa', '/all', $Path) 'İmzalanan dosyanın Windows güven doğrulaması başarısız.'
    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    if ($signature.Status -ne [Management.Automation.SignatureStatus]::Valid -or
        $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -ine $ExpectedThumbprint -or
        $null -eq $signature.TimeStamperCertificate) {
        throw 'Dosyada beklenen yayıncıya ait geçerli ve zaman damgalı Authenticode imzası doğrulanamadı.'
    }
}

try {
    # Missing signing resources must fail, rather than produce an unsigned release.
    if ([string]::IsNullOrWhiteSpace($SigningCertificateThumbprint)) {
        throw 'Güvenilir sağlayıcıdan alınmış mevcut kod imzalama sertifikasının parmak izini -SigningCertificateThumbprint ile belirtin. İmzasız dağıtım paketi oluşturulmadı.'
    }
    $thumbprint = $SigningCertificateThumbprint -replace '\s', ''
    if ($thumbprint -notmatch '^[0-9a-fA-F]{40}$') { throw 'Kod imzalama sertifikası parmak izi 40 onaltılık karakter içermelidir.' }
    [Uri] $timestampUri = $null
    if ([string]::IsNullOrWhiteSpace($TimestampUrl) -or
        -not [Uri]::TryCreate($TimestampUrl, [UriKind]::Absolute, [ref]$timestampUri) -or
        $timestampUri.Scheme -ine 'https' -or [string]::IsNullOrWhiteSpace($timestampUri.Host) -or
        -not [string]::IsNullOrEmpty($timestampUri.UserInfo)) {
        throw 'İmzalama sağlayıcınızın HTTPS zaman damgası adresini -TimestampUrl ile belirtin.'
    }
    if ([string]::IsNullOrWhiteSpace($OutDirectory)) { throw 'Tamamlanmış imzalı ZIP için -OutDirectory belirtin.' }
    if ([Environment]::OSVersion.Platform -ne [PlatformID]::Win32NT) {
        throw 'İmzalı dağıtım derlemesi Windows, .NET Framework 4.8 ve Windows SDK SignTool gerektirir.'
    }
    $outputDirectory = [IO.Path]::GetFullPath($OutDirectory)
    if ((Test-WithinDirectory $outputDirectory $PSScriptRoot) -or
        (Test-WithinDirectory $outputDirectory (Join-Path $PSScriptRoot '..\.git'))) {
        throw 'Çıktı klasörü desktop kaynaklarının veya Git iç dosyalarının içinde olamaz.'
    }
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA) -and
        (Test-WithinDirectory $outputDirectory (Join-Path $env:LOCALAPPDATA 'Vefeyat\Data'))) {
        throw 'Çıktı klasörü kişisel Vefeyât verilerinizin içinde olamaz.'
    }
    if (Test-Path -LiteralPath $outputDirectory -PathType Leaf) { throw 'OutDirectory bir dosya değil klasör olmalıdır.' }
    $finalZip = Join-Path $outputDirectory 'Vefeyat-Windows-Imzali.zip'
    $seedPath = Join-Path $PSScriptRoot 'data\people.csv'
    $sampleData = [string]::IsNullOrWhiteSpace($CsvFile)
    if (-not $sampleData) {
        $seedPath = (Resolve-Path -LiteralPath $CsvFile).ProviderPath
        if (-not (Test-Path -LiteralPath $seedPath -PathType Leaf)) { throw 'CsvFile bir dosyayı göstermelidir.' }
    }
    if ([IO.Path]::GetFullPath($seedPath).Equals($finalZip, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Çıktı ZIP dosyası başlangıç CSV dosyasının üzerine yazılamaz.'
    }
    $sourceNames = @('Core.cs', 'CatalogForm.cs', 'Connect.cs', 'SelectionRules.cs', 'Installer.cs', 'KULLANIM-IMZALI.md')
    foreach ($source in $sourceNames) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $source) -PathType Leaf)) { throw "Kaynak dosya bulunamadı: $source" }
    }
    $recordCount = Get-CsvRecordCount $seedPath
    $certificate = Get-Item -LiteralPath ('Cert:\CurrentUser\My\' + $thumbprint) -ErrorAction Stop
    if (-not $certificate.HasPrivateKey) { throw 'Seçilen sertifikanın kullanılabilir özel anahtarı yok. Sertifika/anahtar dosyalarını sohbete yüklemeyin.' }
    if ($certificate.NotBefore -gt (Get-Date) -or $certificate.NotAfter -le (Get-Date)) { throw 'Seçilen kod imzalama sertifikası geçerlilik süresi dışında.' }
    if ($certificate.Subject -eq $certificate.Issuer) { throw 'Kendinden imzalı sertifika bu dağıtım için kabul edilmez; güvenilir sağlayıcının kod imzalama sertifikası gerekir.' }
    $codeSigningEku = $false
    foreach ($extension in $certificate.Extensions) {
        if ($extension -is [Security.Cryptography.X509Certificates.X509EnhancedKeyUsageExtension]) {
            foreach ($oid in $extension.EnhancedKeyUsages) { if ($oid.Value -eq '1.3.6.1.5.5.7.3.3') { $codeSigningEku = $true } }
        }
    }
    if (-not $codeSigningEku) { throw 'Seçilen sertifikada Kod İmzalama kullanım amacı bulunmuyor.' }

    $frameworkDirectories = @((Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319'))
    if ([Environment]::Is64BitOperatingSystem) { $frameworkDirectories += Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319' }
    $compiler = $null
    foreach ($directory in $frameworkDirectories) {
        $candidate = Join-Path $directory 'csc.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $compiler = $candidate; break }
    }
    if (-not $compiler) { throw '.NET Framework C# derleyicisi bulunamadı.' }
    $signTool = $null
    if (-not [string]::IsNullOrWhiteSpace($SigntoolPath)) {
        $signTool = (Resolve-Path -LiteralPath $SigntoolPath).ProviderPath
        if (-not (Test-Path -LiteralPath $signTool -PathType Leaf)) { throw 'SigntoolPath bir dosyayı göstermelidir.' }
    } else {
        $command = Get-Command signtool.exe -CommandType Application -ErrorAction SilentlyContinue
        if ($null -ne $command) { $signTool = $command.Source }
        if ($null -eq $signTool -and -not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
            $sdkBin = Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin'
            $architecture = 'x86'
            if ([Environment]::Is64BitOperatingSystem) { $architecture = 'x64' }
            if (Test-Path -LiteralPath $sdkBin -PathType Container) {
                foreach ($version in @(Get-ChildItem -LiteralPath $sdkBin -Directory | Where-Object { $_.Name -match '^10\.\d+\.\d+\.\d+$' } | Sort-Object { [Version]$_.Name } -Descending)) {
                    $candidate = Join-Path $version.FullName ($architecture + '\signtool.exe')
                    if (Test-Path -LiteralPath $candidate -PathType Leaf) { $signTool = $candidate; break }
                }
            }
        }
    }
    if ($null -eq $signTool) { throw 'Windows SDK SignTool bulunamadı. Kurulu signtool.exe yolunu -SigntoolPath ile belirtin.' }

    $releaseStage = Join-Path ([IO.Path]::GetTempPath()) ('Vefeyat-signed-build-' + [Guid]::NewGuid().ToString('N'))
    $buildDirectory = Join-Path $releaseStage 'build'
    $releaseRoot = Join-Path $releaseStage 'Vefeyat-Windows'
    [void](New-Item -ItemType Directory -Path $buildDirectory)
    [void](New-Item -ItemType Directory -Path (Join-Path $releaseRoot 'data'))
    $libraryPath = Join-Path $buildDirectory 'Vefeyat.WordAddin.dll'
    $installerPath = Join-Path $buildDirectory 'Vefeyat.Kurulum.exe'
    $installerManifest = Join-Path $buildDirectory 'installer.manifest'
    $manifestText = @'
<?xml version="1.0" encoding="utf-8"?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
  <assemblyIdentity version="1.0.0.0" name="Vefeyat.Kurulum" />
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
    <security><requestedPrivileges><requestedExecutionLevel level="asInvoker" uiAccess="false" /></requestedPrivileges></security>
  </trustInfo>
</assembly>
'@
    [IO.File]::WriteAllText($installerManifest, $manifestText, (New-Object Text.UTF8Encoding($false)))
    $common = @('/nologo', '/platform:anycpu', '/optimize+', '/langversion:5', '/codepage:65001', '/utf8output')
    foreach ($reference in @('System.dll', 'System.Core.dll', 'System.Xml.dll', 'System.Windows.Forms.dll', 'System.Drawing.dll', 'Microsoft.CSharp.dll')) { $common += '/reference:' + $reference }
    $librarySources = @()
    foreach ($source in @('Core.cs', 'CatalogForm.cs', 'Connect.cs', 'SelectionRules.cs')) { $librarySources += Join-Path $PSScriptRoot $source }
    Invoke-CheckedTool $compiler ($common + @('/target:library', ('/out:' + $libraryPath)) + $librarySources) 'Word eklentisi derlenemedi.'
    Invoke-CheckedTool $compiler ($common + @('/target:winexe', ('/out:' + $installerPath), ('/win32manifest:' + $installerManifest), (Join-Path $PSScriptRoot 'Installer.cs'))) 'Windows kurulum programı derlenemedi.'

    foreach ($file in @($libraryPath, $installerPath)) {
        Invoke-CheckedTool $signTool @('sign', '/s', 'My', '/sha1', $thumbprint, '/fd', 'SHA256', '/tr', $timestampUri.AbsoluteUri, '/td', 'SHA256', $file) 'Kod imzalama veya zaman damgası işlemi tamamlanamadı.'
        Confirm-Authenticode $file $thumbprint $signTool
    }
    Copy-Item -LiteralPath $libraryPath -Destination (Join-Path $releaseRoot 'Vefeyat.WordAddin.dll')
    Copy-Item -LiteralPath $installerPath -Destination (Join-Path $releaseRoot 'Vefeyat.Kurulum.exe')
    Copy-Item -LiteralPath $seedPath -Destination (Join-Path $releaseRoot 'data\people.csv')
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'KULLANIM-IMZALI.md') -Destination (Join-Path $releaseRoot 'KULLANIM-IMZALI.md')
    foreach ($file in @('Vefeyat.WordAddin.dll', 'Vefeyat.Kurulum.exe')) {
        Confirm-Authenticode (Join-Path $releaseRoot $file) $thumbprint $signTool
    }
    $metadata = [ordered]@{
        format = 1; product = 'Vefeyât Windows Word'; distribution = 'authenticode-signed';
        assembly_version = [Reflection.AssemblyName]::GetAssemblyName($libraryPath).Version.ToString();
        publisher = $certificate.Subject; signer_thumbprint = $thumbprint;
        seed_records = $recordCount; sample_data = $sampleData;
        seed_sha256 = (Get-FileHash -LiteralPath $seedPath -Algorithm SHA256).Hash.ToLowerInvariant();
        requires_network = $false; requires_server = $false;
        requires_runtime_compilation = $false; windows_word_acceptance = 'pending'
    }
    $utf8 = New-Object Text.UTF8Encoding($false)
    [IO.File]::WriteAllText((Join-Path $releaseRoot 'paket-bilgisi.json'), (($metadata | ConvertTo-Json -Depth 4) + [Environment]::NewLine), $utf8)
    $packageFiles = @('Vefeyat.Kurulum.exe', 'Vefeyat.WordAddin.dll', 'data\people.csv', 'KULLANIM-IMZALI.md', 'paket-bilgisi.json')
    $manifest = @()
    foreach ($relative in $packageFiles) {
        $manifest += ((Get-FileHash -LiteralPath (Join-Path $releaseRoot $relative) -Algorithm SHA256).Hash.ToLowerInvariant() + '  ' + $relative.Replace('\', '/'))
    }
    [IO.File]::WriteAllText((Join-Path $releaseRoot 'SHA256SUMS.txt'), (($manifest -join "`n") + "`n"), $utf8)
    $stagedZip = Join-Path $releaseStage 'Vefeyat-Windows-Imzali.zip'
    Compress-Archive -LiteralPath $releaseRoot -DestinationPath $stagedZip -CompressionLevel Optimal

    # Publish only after both binaries have valid, timestamped signatures.
    [void](New-Item -ItemType Directory -Path $outputDirectory -Force)
    $pendingOutput = Join-Path $outputDirectory ('.Vefeyat-signed-' + [Guid]::NewGuid().ToString('N') + '.tmp')
    Copy-Item -LiteralPath $stagedZip -Destination $pendingOutput
    if (Test-Path -LiteralPath $finalZip -PathType Leaf) { [IO.File]::Replace($pendingOutput, $finalZip, $null) }
    else { [IO.File]::Move($pendingOutput, $finalZip) }
    $pendingOutput = $null
    Write-Host ('İmzalı dağıtım paketi oluşturuldu: ' + $finalZip)
    Write-Host ('Kayıt sayısı: ' + $recordCount)
    Write-Host ('SHA-256: ' + (Get-FileHash -LiteralPath $finalZip -Algorithm SHA256).Hash.ToLowerInvariant())
    Write-Host 'Derleme bilgisayarında Authenticode doğrulandı. Akıllı Uygulama Denetimi açık gerçek Windows Word kabul testi ayrıca gerekir.'
} catch {
    Write-Host ('İmzalı dağıtım tamamlanamadı: ' + $_.Exception.Message) -ForegroundColor Red
    exit 1
} finally {
    if ($null -ne $pendingOutput -and (Test-Path -LiteralPath $pendingOutput)) { Remove-Item -LiteralPath $pendingOutput -Force -ErrorAction SilentlyContinue }
    if ($null -ne $releaseStage -and (Test-Path -LiteralPath $releaseStage)) { Remove-Item -LiteralPath $releaseStage -Recurse -Force -ErrorAction SilentlyContinue }
}
