@echo off
rem {{NAME}} kurulumu: cift tiklayin. Word'e eklentiyi kaydeder (yonetici izni gerekmez).
rem Bu dosyanin geri kalani PowerShell betigidir; asagidaki satir onu calistirir.
setlocal
set "VEFEYAT_SELF=%~f0"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:VEFEYAT_SELF,[Text.Encoding]::UTF8); $i=$t.IndexOf('#'+'POWERSHELL'); Invoke-Expression $t.Substring($i)"
echo.
pause
exit /b
#POWERSHELL
$ErrorActionPreference = 'Stop'
$Base = '{{BASE_URL}}'
$Id = '{{ID}}'
$Name = '{{NAME}}'
$Folder = Join-Path $env:LOCALAPPDATA 'Vefeyat'
$Key = 'HKCU:\Software\Microsoft\Office\16.0\WEF\Developer'
Write-Host ''
Write-Host "  $Name kurulumu" -ForegroundColor Cyan
Write-Host '  ------------------------'
try {
    while (Get-Process WINWORD -ErrorAction SilentlyContinue) {
        Write-Host ''
        Write-Host '  Word acik. Calismanizi kaydedip Word''un BUTUN pencerelerini kapatin,' -ForegroundColor Yellow
        Read-Host '  sonra Enter''a basin (vazgecmek icin pencereyi kapatin)' | Out-Null
    }

    # Word 2408 (build 17928) and later: keyboard shortcuts (manifest.xml); older: manifest-uyumlu.xml.
    $file = 'manifest.xml'
    $word = $null
    foreach ($root in 'HKLM:', 'HKCU:') {
        $p = Get-ItemProperty "$root\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Winword.exe" -ErrorAction SilentlyContinue
        if ($p -and $p.'(default)' -and (Test-Path $p.'(default)')) { $word = $p.'(default)'; break }
    }
    if ($word) {
        $build = (Get-Item $word).VersionInfo.FileBuildPart
        Write-Host "  Word bulundu (yapi $build)."
        if ($build -lt 17928) {
            $file = 'manifest-uyumlu.xml'
            Write-Host '  Bu Word surumu klavye kisayollarini desteklemiyor; uyumlu surum kurulacak.'
        }
    } else {
        Write-Host '  Word''un yeri bulunamadi; guncel surum kurulacak.' -ForegroundColor Yellow
    }

    New-Item -ItemType Directory -Force -Path $Folder | Out-Null
    $manifest = Join-Path $Folder 'manifest.xml'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Write-Host "  Indiriliyor: $Base$file"
    Invoke-WebRequest -UseBasicParsing -Uri ($Base + $file) -OutFile $manifest
    $xml = [IO.File]::ReadAllText($manifest)
    if ($xml -notmatch [Regex]::Escape("<Id>$Id</Id>")) { throw 'Indirilen dosya beklenen eklenti degil.' }

    if (-not (Test-Path $Key)) { New-Item -Path $Key -Force | Out-Null }
    New-ItemProperty -Path $Key -Name $Id -Value $manifest -PropertyType String -Force | Out-Null

    Write-Host ''
    Write-Host "  Tamam, eklenti Word'e kaydedildi." -ForegroundColor Green
    Write-Host ''
    Write-Host "  Simdi Word'u acin. Seritte '$Name' sekmesi gorunmezse (yalniz ilk seferde):" -ForegroundColor Yellow
    Write-Host '    Giris -> Eklentiler (ya da Ekle -> Eklentilerim) -> Diger Eklentiler'
    Write-Host "    -> 'Gelistirici Eklentileri' bolumunde '$Name' -> Ekle"
    Write-Host ''
    Write-Host '  Guncelleme gerektiginde bu dosyayi yeniden calistirmaniz yeterli.'
    Write-Host '  Kaldirmak icin: Vefeyat-Kaldir.cmd'
} catch {
    Write-Host ''
    Write-Host ('  Kurulum tamamlanamadi: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host '  Bu pencerenin ekran goruntusunu apoyraz@cumhuriyet.edu.tr adresine gonderin.'
}
