#requires -Version 5.1
# CekRingkas QC Hub: 3 pilihan dalam satu kali menjalankan PowerShell.
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $MyInvocation.MyCommand.Path
$modules=Join-Path $root 'Modules'
$reports=Join-Path $root 'Reports'
New-Item -ItemType Directory -Force -Path $modules,$reports | Out-Null
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
function Get-OriginalGist {
    $local=Join-Path $modules 'CekRingkas_Original_Gist.ps1'
    $api='https://api.github.com/gists/44b2bbf12700b5ec9e18760c3e1630ce'
    try {
        Write-Host 'Memeriksa versi terbaru script Gist...' -ForegroundColor Cyan
        $g=Invoke-RestMethod -Uri $api -Headers @{'User-Agent'='CekRingkas-QC-Hub'} -TimeoutSec 12
        $file=$g.files.PSObject.Properties | Where-Object {$_.Name -match '\.ps1$'} | Select-Object -First 1
        if($file -and $file.Value.raw_url){
            $content=Invoke-RestMethod -Uri $file.Value.raw_url -Headers @{'User-Agent'='CekRingkas-QC-Hub'} -TimeoutSec 20
            if($content -match '(?i)SKRIP INSPEKSI|CekRingkas|Get-CimInstance|Win32_ComputerSystem'){
                [IO.File]::WriteAllText($local,[string]$content,(New-Object System.Text.UTF8Encoding($true)))
                Write-Host "Gist terbaru tersimpan: $($file.Name)" -ForegroundColor Green
                return $local
            }
        }
    } catch { Write-Host 'Gist tidak dapat diunduh saat ini; menggunakan salinan fallback yang disertakan.' -ForegroundColor Yellow }
    if(Test-Path $local){return $local}
    $fallback=Join-Path $modules 'CekRingkas_Fallback_V11.1.ps1'
    if(Test-Path $fallback){return $fallback}
    throw 'Script Gist dan fallback tidak ditemukan.'
}
function Invoke-Script([string]$path,[string[]]$arguments=@()) {
    if(!(Test-Path $path)){throw "File tidak ditemukan: $path"}
    Write-Host "`nMenjalankan: $([IO.Path]::GetFileName($path))" -ForegroundColor Cyan
    & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $path @arguments
    Write-Host "`nProses selesai. Tekan Enter untuk kembali ke menu." -ForegroundColor DarkGray
    [void](Read-Host)
}
while($true){
    Clear-Host
    Write-Host '==========================================================' -ForegroundColor DarkCyan
    Write-Host '       CEKRINGKAS LAPTOP — INSPEKSI & QC HUB' -ForegroundColor Cyan
    Write-Host '==========================================================' -ForegroundColor DarkCyan
    Write-Host 'Pilih mode yang ingin dijalankan:'
    Write-Host ''
    Write-Host '  [1] SUMMARY REPORT AWAL (script Gist asli)' -ForegroundColor Green
    Write-Host '      Laporan spesifikasi lengkap seperti versi awal.'
    Write-Host '      Mode non-interaktif dipakai jika didukung oleh Gist.'
    Write-Host ''
    Write-Host '  [2] SUMMARY QC + TES INTERAKTIF' -ForegroundColor Yellow
    Write-Host '      Tes keyboard, LCD warna, kamera, audio, port, baterai,'
    Write-Host '      Device Manager, RAM, disk, dan hasil LULUS/GAGAL.'
    Write-Host ''
    Write-Host '  [3] COMPONENT / MOTHERBOARD SCANNER' -ForegroundColor Magenta
    Write-Host '      Detail motherboard, CPU, GPU, RAM per modul, storage,'
    Write-Host '      baterai, Wi-Fi, audio, webcam, error driver + diagram.'
    Write-Host ''
    Write-Host '  [0] Keluar' -ForegroundColor Gray
    Write-Host ''
    $choice=Read-Host 'Masukkan pilihan (0-3)'
    try {
        switch($choice){
            '1' {
                $original=Get-OriginalGist
                $source=Get-Content -LiteralPath $original -Raw
                if($source -match '(?is)param\s*\([^)]*\bSkipInteractive\b'){
                    Invoke-Script $original @('-SkipInteractive')
                } elseif($source -match '(?i)\[switch\]\s*\$SkipInteractive') {
                    Invoke-Script $original @('-SkipInteractive')
                } else {
                    Write-Host 'Gist ini tidak mendeklarasikan -SkipInteractive; menjalankan mode aslinya agar tidak mengubah alur bawaan.' -ForegroundColor Yellow
                    Invoke-Script $original @()
                }
            }
            '2' { Invoke-Script (Join-Path $modules 'QC_Guided_V13.ps1') @() }
            '3' { Invoke-Script (Join-Path $modules 'ComponentScanner.ps1') @() }
            '0' { break }
            default { Write-Host 'Pilihan tidak dikenal.' -ForegroundColor Red; Start-Sleep -Seconds 1 }
        }
    } catch { Write-Host "Terjadi kendala: $($_.Exception.Message)" -ForegroundColor Red; [void](Read-Host 'Tekan Enter untuk kembali') }
    if($choice -eq '0'){break}
}
