#requires -Version 5.1
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# This script does not change execution policy, macro security, or machine-wide keys.
$progId = 'Vefeyat.Connect'
$classId = '{2B0AF7B7-BD0D-4C70-9136-BED2DBD8F8F3}'
$className = 'Vefeyat.Connect'
$registryPaths = @("Software\Classes\CLSID\$classId", "Software\Classes\$progId", "Software\Microsoft\Office\Word\Addins\$progId")

function Get-RegistrySnapshot([Microsoft.Win32.RegistryKey] $Key) {
    $values = @()
    foreach ($name in $Key.GetValueNames()) {
        $values += [pscustomobject]@{ Name = $name; Kind = $Key.GetValueKind($name); Data = $Key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames) }
    }
    $children = @()
    foreach ($name in $Key.GetSubKeyNames()) {
        $child = $Key.OpenSubKey($name)
        try { $children += [pscustomobject]@{ Name = $name; Tree = Get-RegistrySnapshot $child } }
        finally { $child.Dispose() }
    }
    return [pscustomobject]@{ Values = $values; Children = $children }
}

function Set-RegistrySnapshot([Microsoft.Win32.RegistryKey] $Key, $Tree) {
    foreach ($value in $Tree.Values) { $Key.SetValue($value.Name, $value.Data, $value.Kind) }
    foreach ($child in $Tree.Children) {
        $subkey = $Key.CreateSubKey($child.Name)
        try { Set-RegistrySnapshot $subkey $child.Tree } finally { $subkey.Dispose() }
    }
}

function Set-RegistryValues([Microsoft.Win32.RegistryKey] $Base, [string] $Path, [hashtable] $Values) {
    $key = $Base.CreateSubKey($Path)
    try {
        foreach ($name in $Values.Keys) {
            $kind = [Microsoft.Win32.RegistryValueKind]::String
            if ($Values[$name] -is [int]) { $kind = [Microsoft.Win32.RegistryValueKind]::DWord }
            $key.SetValue($name, $Values[$name], $kind)
        }
    } finally { $key.Dispose() }
}

$temporary = $null
$bases = @()
$snapshots = @()
$fileBackups = @()
$changedFiles = @()
$registryChanged = $false
try {
    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) {
        throw 'Önce bütün Word pencerelerini kapatın; sonra Kur.cmd dosyasını yeniden çalıştırın.'
    }
    if ($null -eq [Type]::GetTypeFromProgID('Word.Application')) { throw 'Windows masaüstü Microsoft Word kurulu ve bu kullanıcı için kayıtlı olmalıdır.' }
    if ([string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { throw 'Windows yerel uygulama veri klasörü bulunamadı.' }
    $installDirectory = Join-Path $env:LOCALAPPDATA 'Vefeyat\Addin'
    $assemblyPath = Join-Path $installDirectory 'Vefeyat.WordAddin.dll'
    $seedPath = Join-Path $PSScriptRoot 'data\people.csv'
    $sourceNames = @('Core.cs', 'CatalogForm.cs', 'Connect.cs', 'SelectionRules.cs')
    foreach ($source in $sourceNames) {
        if (-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $source) -PathType Leaf)) { throw "Paket eksik: $source" }
    }
    if (-not (Test-Path -LiteralPath $seedPath -PathType Leaf)) { throw 'Paket eksik: data\people.csv' }
    $seed = @(Import-Csv -LiteralPath $seedPath -Delimiter ';' -Encoding UTF8)
    $expectedHeaders = @('Kısa Ad', 'Tam Ad', 'Orijinal Ad', 'Vefat Tarihi', 'Tanıtım', 'Kaynak Link')
    if ($seed.Count -eq 0 -or (@($seed[0].PSObject.Properties.Name) -join '|') -cne ($expectedHeaders -join '|')) {
        throw 'Başlangıç CSV dosyasının altı sütunu beklenen biçimde değil veya dosya boş.'
    }
    $frameworkDirectories = @((Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319'))
    if ([Environment]::Is64BitOperatingSystem) { $frameworkDirectories += Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319' }
    $compiler = $null
    foreach ($directory in $frameworkDirectories) {
        $candidate = Join-Path $directory 'csc.exe'
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { $compiler = $candidate; break }
    }
    if (-not $compiler) { throw 'Windows .NET Framework 4.x derleyicisi bulunamadı. .NET Framework 4.8 kurulu Windows gerekir.' }
    $temporary = Join-Path ([IO.Path]::GetTempPath()) ('Vefeyat-install-' + [Guid]::NewGuid().ToString('N'))
    [void](New-Item -ItemType Directory -Path $temporary)
    $compiledPath = Join-Path $temporary 'Vefeyat.WordAddin.dll'
    $arguments = @('/nologo', '/target:library', '/platform:anycpu', '/optimize+', '/langversion:5', '/codepage:65001', '/utf8output', ('/out:' + $compiledPath))
    foreach ($reference in @('System.dll', 'System.Core.dll', 'System.Xml.dll', 'System.Windows.Forms.dll', 'System.Drawing.dll', 'Microsoft.CSharp.dll')) { $arguments += '/reference:' + $reference }
    foreach ($source in $sourceNames) { $arguments += Join-Path $PSScriptRoot $source }
    $compilerOutput = & $compiler @arguments 2>&1
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $compiledPath -PathType Leaf)) {
        throw ('Eklenti derlenemedi. Hiçbir kurulum kaydı değiştirilmedi.' + [Environment]::NewLine + ($compilerOutput -join [Environment]::NewLine))
    }
    $assemblyName = [Reflection.AssemblyName]::GetAssemblyName($compiledPath)
    if ($assemblyName.Name -cne 'Vefeyat.WordAddin') { throw 'Derlenen eklentinin kimliği beklenen biçimde değil.' }
    $version = $assemblyName.Version.ToString()
    $codeBase = ([Uri]$assemblyPath).AbsoluteUri
    $views = @([Microsoft.Win32.RegistryView]::Registry32)
    if ([Environment]::Is64BitOperatingSystem) { $views += [Microsoft.Win32.RegistryView]::Registry64 }
    foreach ($view in $views) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        $bases += $base
        $existingClass = $base.OpenSubKey("Software\Classes\CLSID\$classId\InprocServer32")
        if ($null -ne $existingClass) {
            try {
                if ($existingClass.GetValue('Class', '') -cne $className) { throw 'Aynı COM kimliği başka bir uygulamaya ait. Kurulum durduruldu.' }
            } finally { $existingClass.Dispose() }
        }
        $existingProgId = $base.OpenSubKey("Software\Classes\$progId\CLSID")
        if ($null -ne $existingProgId) {
            try {
                if ($existingProgId.GetValue('', '') -ine $classId) { throw 'Vefeyât program kimliği başka bir COM sınıfına bağlı. Kurulum durduruldu.' }
            } finally { $existingProgId.Dispose() }
        }
        foreach ($path in $registryPaths) {
            $key = $base.OpenSubKey($path)
            $tree = $null
            if ($null -ne $key) { try { $tree = Get-RegistrySnapshot $key } finally { $key.Dispose() } }
            $snapshots += [pscustomobject]@{ Base = $base; Path = $path; Tree = $tree }
        }
    }
    if (Get-Process -Name WINWORD -ErrorAction SilentlyContinue) { throw 'Kurulum sırasında Word açıldı. Word uygulamasını kapatıp yeniden deneyin.' }
    # Back up only files owned by this installer. Never touch Vefeyat\Data.
    $files = @(
        [pscustomobject]@{ Source = $compiledPath; Destination = $assemblyPath },
        [pscustomobject]@{ Source = $seedPath; Destination = (Join-Path $installDirectory 'data\people.csv') }
    )
    foreach ($file in $files) {
        $backup = $null
        if (Test-Path -LiteralPath $file.Destination -PathType Leaf) {
            $backup = Join-Path $temporary ('backup-' + $fileBackups.Count)
            Copy-Item -LiteralPath $file.Destination -Destination $backup
        }
        $fileBackups += [pscustomobject]@{ Destination = $file.Destination; Backup = $backup }
    }
    foreach ($file in $files) {
        [void](New-Item -ItemType Directory -Force -Path ([IO.Path]::GetDirectoryName($file.Destination)))
        $changedFiles += $file.Destination
        Copy-Item -LiteralPath $file.Source -Destination $file.Destination -Force
    }
    $registration = @{ '' = 'mscoree.dll'; ThreadingModel = 'Both'; Class = $className; Assembly = $assemblyName.FullName; RuntimeVersion = 'v4.0.30319'; CodeBase = $codeBase }
    $registryChanged = $true
    foreach ($base in $bases) {
        $classPath = "Software\Classes\CLSID\$classId"
        # Clear obsolete version subkeys during repair; the full tree is backed up.
        $base.DeleteSubKeyTree($classPath, $false)
        Set-RegistryValues $base $classPath @{ '' = $className }
        Set-RegistryValues $base "$classPath\InprocServer32" $registration
        Set-RegistryValues $base "$classPath\InprocServer32\$version" $registration
        Set-RegistryValues $base "$classPath\ProgId" @{ '' = $progId }
        Set-RegistryValues $base "Software\Classes\$progId" @{ '' = 'Vefeyât Word Eklentisi' }
        Set-RegistryValues $base "Software\Classes\$progId\CLSID" @{ '' = $classId }
        Set-RegistryValues $base "Software\Microsoft\Office\Word\Addins\$progId" @{ FriendlyName = 'Vefeyât'; Description = 'Bilgisayarınızdaki şahıs kataloğu'; LoadBehavior = [int]3; CommandLineSafe = [int]0 }
    }
    Write-Host 'Vefeyât kuruldu. Word uygulamasını açın; şeritte Vefeyât sekmesi görünecektir.'
    Write-Host ('Eklenti: ' + $installDirectory)
    Write-Host ('Kişisel veriler: ' + (Join-Path $env:LOCALAPPDATA 'Vefeyat\Data'))
    Write-Host 'Kurulum ve günlük kullanım internet bağlantısı gerektirmez.'
} catch {
    $originalError = $_.Exception.Message
    $rollbackErrors = @()
    if ($registryChanged) {
        foreach ($snapshot in $snapshots) {
            try {
                $snapshot.Base.DeleteSubKeyTree($snapshot.Path, $false)
                if ($null -ne $snapshot.Tree) {
                    $key = $snapshot.Base.CreateSubKey($snapshot.Path)
                    try { Set-RegistrySnapshot $key $snapshot.Tree } finally { $key.Dispose() }
                }
            } catch { $rollbackErrors += $_.Exception.Message }
        }
    }
    foreach ($file in $fileBackups) {
        if ($changedFiles -contains $file.Destination) {
            try {
                if ($null -ne $file.Backup) { Copy-Item -LiteralPath $file.Backup -Destination $file.Destination -Force }
                elseif (Test-Path -LiteralPath $file.Destination -PathType Leaf) { Remove-Item -LiteralPath $file.Destination -Force }
            } catch { $rollbackErrors += $_.Exception.Message }
        }
    }
    Write-Host ('Kurulum tamamlanamadı: ' + $originalError) -ForegroundColor Red
    foreach ($failure in $rollbackErrors) { Write-Host ('Önceki kurulum geri yüklenirken hata: ' + $failure) -ForegroundColor Red }
    exit 1
} finally {
    foreach ($base in $bases) { $base.Dispose() }
    if ($null -ne $temporary -and (Test-Path -LiteralPath $temporary)) { Remove-Item -LiteralPath $temporary -Recurse -Force -ErrorAction SilentlyContinue }
}
