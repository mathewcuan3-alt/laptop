# ==============================================================================
# SKRIP INSPEKSI LAPTOP - V11.1 PANELOOK INLINE + SMART
# - Tampilan laporan sama seperti V8.9
# - Data lebih lengkap, hanya membaca data asli (tanpa angka default palsu)
# - Jauh lebih cepat: baterai via WMI (tanpa powercfg), query langsung difilter
# - Status disk ala CrystalDiskInfo: NVMe SMART/ATA SMART/Windows Storage; panel gamut: database Panelook tertanam + estimasi EDID
# Catatan: simpan sebagai .ps1 (ASCII). Jalankan sebagai Administrator agar
#          data kesehatan disk (SMART, suhu, jam nyala) lebih lengkap.
# Peringatan: EDID gamut = estimasi luas relatif NTSC, bukan pengukuran coverage kolorimeter.
# Status disk disederhanakan; SMART raw/vendor-specific tidak selalu punya arti universal.
# ==============================================================================
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$ErrorActionPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

# Minta hak Administrator (1x klik UAC) agar data SSD lengkap. Jika ditolak, lanjut tanpa admin.
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and $PSCommandPath) {
    try {
        Start-Process powershell -Verb RunAs -ErrorAction Stop -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
        exit
    } catch { }
}

function Enc($s) { [System.Net.WebUtility]::HtmlEncode([string]$s) }
function Dec($a) { if (-not $a) { return '' }; (($a | Where-Object { $_ -ne 0 } | ForEach-Object { [char]$_ }) -join '').Trim() }
function Limit([double]$v, [int]$lo, [int]$hi) { [int][math]::Max($lo, [math]::Min($hi, [math]::Round($v))) }
function Badge($ok, $text) {
    if ($ok) { return "<span class='badge-ok'>&#10003; $text</span> " }
    return "<span class='badge-bad'>&#10007; $text</span> "
}
function Get-GpuClass($n) {
    if ($n -match 'RTX\s*(50|40|30|20)\d{2}') { return "NVIDIA GeForce RTX $($Matches[1]) Series" }
    elseif ($n -match 'GTX\s*(16|10)\d{2}') { return "NVIDIA GeForce GTX $($Matches[1]) Series" }
    elseif ($n -match 'MX\s*\d{3}') { return 'NVIDIA GeForce MX Series' }
    elseif ($n -match 'NVIDIA|GeForce|Quadro') { return 'NVIDIA GeForce Dedicated GPU' }
    elseif ($n -match 'Radeon\s*(RX|PRO)') { return 'AMD Radeon RX Diskrit' }
    elseif ($n -match 'Arc.*\b[AB]\d{3}') { return 'Intel Arc Diskrit' }
    elseif ($n -match 'Arc') { return 'Intel Arc Graphics (Integrated)' }
    elseif ($n -match 'Iris') { return 'Intel Iris Xe Graphics' }
    elseif ($n -match 'Intel') { return 'Intel Integrated Graphics' }
    elseif ($n -match 'Radeon') { return 'AMD Radeon Integrated Graphics' }
    return 'Standard Display Adapter'
}

# Barcode Code 128 (subset B) -> SVG mandiri, tanpa library/internet. Hanya ASCII 32-126.
function New-Code128Svg([string]$text) {
    $pat = @(
        '212222','222122','222221','121223','121322','131222','122213','122312','132212','221213',
        '221312','231212','112232','122132','122231','113222','123122','123221','223211','221132',
        '221231','213212','223112','312131','311222','321122','321221','312212','322112','322211',
        '212123','212321','232121','111323','131123','131321','112313','132113','132311','211313',
        '231113','231311','112133','112331','132131','113123','113321','133121','313121','211331',
        '231131','213113','213311','213131','311123','311321','331121','312113','312311','332111',
        '314111','221411','431111','111224','111422','121124','121421','141122','141221','112214',
        '112412','122114','122411','142112','142211','241211','221114','413111','241112','134111',
        '111242','121142','121241','114212','124112','124211','411212','421112','421211','212141',
        '214121','412121','111143','111341','131141','114113','114311','411113','411311','113141',
        '114131','311141','411131','211412','211214','211232','2331112'
    )
    $vals = New-Object System.Collections.ArrayList
    [void]$vals.Add(104)
    $sum = 104; $pos = 1
    foreach ($c in $text.ToCharArray()) {
        $v = [int]$c - 32
        if ($v -lt 0 -or $v -gt 94) { return '' }
        [void]$vals.Add($v); $sum += $v * $pos; $pos++
    }
    [void]$vals.Add($sum % 103)
    [void]$vals.Add(106)
    $x = 10
    $rects = New-Object System.Text.StringBuilder
    foreach ($v in $vals) {
        $isBar = $true
        foreach ($ch in $pat[$v].ToCharArray()) {
            $w = [int][string]$ch
            if ($isBar) { [void]$rects.Append("<rect x='$x' y='0' width='$w' height='70'/>") }
            $x += $w; $isBar = -not $isBar
        }
    }
    $total = $x + 10
    $pxW = $total * 2
    return "<svg xmlns='http://www.w3.org/2000/svg' width='$pxW' height='70' viewBox='0 0 $total 70' preserveAspectRatio='none' shape-rendering='crispEdges' role='img' aria-label='Barcode S/N'><rect x='0' y='0' width='$total' height='70' fill='#fff'/><g fill='#000'>$($rects.ToString())</g></svg>"
}

# ------------------------------------------------------------------------------
# 1. Identitas sistem
# ------------------------------------------------------------------------------
$desktopPath = [Environment]::GetFolderPath('Desktop')
$sys  = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$os   = Get-CimInstance Win32_OperatingSystem

$brand  = if ($sys.Manufacturer) { $sys.Manufacturer.Trim() } else { 'Generic' }
$model  = if ($sys.Model) { $sys.Model.Trim() } else { 'Laptop' }
$serialNumber = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { 'N/A' }
if ($serialNumber -match 'O\.E\.M|Default string|None|System Serial|^0+$') { $serialNumber = 'N/A' }
$osName = if ($os.Caption) { ($os.Caption -replace 'Microsoft ', '').Trim() } else { 'Windows' }
$osExtra = "Build $($os.BuildNumber), $($os.OSArchitecture)"

$tahun = 'N/A'; $tglProduksi = 'N/A'; $umur = $null
if ($bios.ReleaseDate) {
    $tahun = $bios.ReleaseDate.Year
    $tglProduksi = $bios.ReleaseDate.ToString('dd MMMM yyyy', [System.Globalization.CultureInfo]::InvariantCulture)
    $umur = (Get-Date).Year - $tahun
}
$biosVer = [string]$bios.SMBIOSBIOSVersion

# Lisensi Windows (query difilter langsung di WMI = cepat)
$winStatus = 'Not Activated'; $winClass = 'badge-bad'; $winOriginal = $false
$lic = Get-CimInstance -Query "SELECT LicenseStatus, ProductKeyChannel FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL AND Name LIKE 'Windows%'" | Select-Object -First 1
$oa3 = (Get-CimInstance SoftwareLicensingService).OA3xOriginalProductKey
if ($lic -and $lic.LicenseStatus -eq 1) {
    $ch = [string]$lic.ProductKeyChannel
    if ($ch -like 'OEM*' -or $oa3) { $winStatus = 'Original (Permanently Activated)'; $winClass = 'badge-ok'; $winOriginal = $true }
    elseif ($ch -like 'Retail*') { $winStatus = 'Original (Permanently Activated)'; $winClass = 'badge-ok'; $winOriginal = $true }
    elseif ($ch -like 'Volume*') { $winStatus = 'Activated via Volume/KMS (Bukan OEM/Retail)'; $winClass = 'badge-warn' }
    else { $winStatus = 'Activated'; $winClass = 'badge-ok'; $winOriginal = $true }
}

# ------------------------------------------------------------------------------
# 2. CPU + analisis kesetaraan (sistem Tier 1-10, berlaku Intel/AMD/Qualcomm)
# ------------------------------------------------------------------------------
$proc = Get-CimInstance Win32_Processor | Select-Object -First 1
$cpuName = if ($proc -and $proc.Name) { ($proc.Name -replace '\s+', ' ').Trim() } else { 'Unknown Processor' }
$cores = [int]$proc.NumberOfCores
$threads = [int]$proc.NumberOfLogicalProcessors
$maxGHz = if ($proc.MaxClockSpeed) { [math]::Round($proc.MaxClockSpeed / 1000, 2) } else { 0 }
$l3MB = if ($proc.L3CacheSize) { [math]::Round($proc.L3CacheSize / 1024, 1) } else { 0 }

$tierName = @{ 1='Sangat Dasar'; 2='Dasar'; 3='Entry-Level'; 4='Menengah Bawah'; 5='Menengah'; 6='Menengah Atas'; 7='Performa Tinggi'; 8='High-End'; 9='Enthusiast'; 10='Flagship' }
$tierIntel = @{
    1='Celeron N3350 / N4020 / Pentium Gold 4415Y'
    2='Core i3-7100U / i5-4200U / i5-6200U / Pentium Silver N5000'
    3='Core i5-7200U / i3-8130U / Celeron N5105 / Intel N100'
    4='Core i5-8250U / i5-10210U / i3-1115G4 / Intel N305'
    5='Core i5-1135G7 / i7-1165G7 / i7-8565U / i3-1215U'
    6='Core i5-1235U / i5-1335U / i7-1255U / Core Ultra 5 125U'
    7='Core i5-12500H / i7-11800H / i7-1360P / Core Ultra 5 125H'
    8='Core i7-12700H / i7-13700H / Core Ultra 7 155H'
    9='Core i9-13900H / i7-13700HX / Core Ultra 9 185H'
    10='Core i9-13980HX / i9-14900HX / Core Ultra 9 285HX'
}
$tierAMD = @{
    1='A4-9125 / A6-9220 / E2-9000'
    2='A9-9425 / A10-9620P / Athlon Silver 3050U / Ryzen 3 2200U'
    3='Ryzen 3 3200U / Ryzen 5 2500U / Athlon Gold 3150U'
    4='Ryzen 5 3500U / Ryzen 3 5300U / Ryzen 5 7520U'
    5='Ryzen 5 4500U / Ryzen 5 5500U / Ryzen 3 7330U'
    6='Ryzen 5 5600U / Ryzen 7 5700U / Ryzen 5 7530U / Ryzen 7 4700U'
    7='Ryzen 5 5600H / Ryzen 5 6600H / Ryzen 5 7535HS / Ryzen 5 8640HS'
    8='Ryzen 7 5800H / Ryzen 7 7735HS / Ryzen 7 7840HS / Ryzen AI 7 350'
    9='Ryzen 9 6900HX / Ryzen 9 7940HS / Ryzen AI 9 HX 370'
    10='Ryzen 9 7945HX / Ryzen 9 9955HX / Ryzen AI Max+ 395'
}
$tierUse = @{
    1='Pengetikan, browsing ringan, kasir/POS'
    2='Office dasar, Zoom/Meet, kuliah daring ringan'
    3='Office, tugas sekolah/kuliah, streaming (multitasking terbatas)'
    4='Office, kuliah, multitasking menengah, edit foto ringan'
    5='Kerja harian lancar, multitasking, edit foto/video ringan'
    6='Kerja kantoran berat, coding, edit foto, video FHD ringan'
    7='Coding, desain 2D/3D ringan, edit video, gaming ringan-menengah (dengan GPU diskrit)'
    8='Edit video, desain 3D, programming berat, gaming menengah-tinggi (dengan GPU diskrit)'
    9='Kreator profesional, render, gaming tinggi'
    10='Workstation, gaming ekstrem, render berat'
}
$lvlAdj = @{ 3 = -1; 5 = 0; 7 = 1; 9 = 2 }
$intelGen = @{
    1  = @('Nehalem / Westmere', 2010)
    2  = @('Sandy Bridge', 2011)
    3  = @('Ivy Bridge', 2012)
    4  = @('Haswell', 2013)
    5  = @('Broadwell', 2015)
    6  = @('Skylake', 2015)
    7  = @('Kaby Lake', 2017)
    8  = @('Kaby Lake Refresh / Coffee Lake', 2017)
    9  = @('Coffee Lake Refresh', 2019)
    10 = @('Comet Lake / Ice Lake', 2019)
    11 = @('Tiger Lake / Rocket Lake', 2020)
    12 = @('Alder Lake', 2021)
    13 = @('Raptor Lake', 2022)
    14 = @('Raptor Lake Refresh', 2023)
}
$intelBase = @{ 1=1; 2=1; 3=2; 4=2; 5=2; 6=2; 7=3; 8=4; 9=4; 10=4; 11=5; 12=6; 13=6; 14=6 }

$codename = 'Arsitektur tidak dikenali (cek manual)'
$cpuYear = 'N/A'; $tier = 3; $hybrid = $false; $npu = $false; $cpuNote = ''

if ($cpuName -match 'Intel') {
    if ($cpuName -match 'Core\(TM\) Ultra\s*(?:X\s*)?([3579])\s*(\d{3})(\w*)') {
        $lvl = [int]$Matches[1]; $num = $Matches[2]; $sfx = $Matches[3]
        $npu = $true; $hybrid = $true
        if ($sfx -match 'V$') { $codename = 'Lunar Lake (Core Ultra Series 2)'; $cpuYear = 2024; $base = 7 }
        elseif ($num.StartsWith('1')) { $codename = 'Meteor Lake (Core Ultra Series 1)'; $cpuYear = 2023; $base = 6; if ($sfx -match 'H') { $base = 7 } }
        elseif ($num.StartsWith('2')) { $codename = 'Arrow Lake (Core Ultra Series 2)'; $cpuYear = 2024; $base = 7; if ($sfx -match 'HX') { $base = 9 } elseif ($sfx -match 'U') { $base = 6 } }
        else { $codename = 'Panther Lake (Core Ultra Series 3)'; $cpuYear = 2026; $base = 8 }
        $adj = $lvlAdj[$lvl]
        if ($sfx -match '^U') { $adj = [math]::Floor($adj / 2) }
        $tier = Limit ($base + $adj) 1 10
    }
    elseif ($cpuName -match 'Core\(TM\)\s+([357])\s+(\d{3})(\w*)') {
        $lvl = [int]$Matches[1]; $sfx = $Matches[3]
        $codename = 'Raptor Lake Refresh (Intel Core Series 1/2)'; $cpuYear = 2024; $hybrid = $true
        $base = 6; if ($sfx -match 'H') { $base = 7 }
        $tier = Limit ($base + $lvlAdj[$lvl]) 1 10
    }
    elseif ($cpuName -match '\bN(95|97|100|150|200|250|300|305|355)\b') {
        $n = [int]$Matches[1]
        if ($n -in 150, 250, 355) { $codename = 'Intel Twin Lake-N (E-Core Series Modern)'; $cpuYear = 2025 }
        else { $codename = 'Intel Alder Lake-N (E-Core Series Modern)'; $cpuYear = 2023 }
        if ($n -in 95, 97, 100, 200) { $tier = 3 } else { $tier = 4 }
        $cpuNote = 'Prosesor seri N (E-Core saja): hemat daya, performa single-core terbatas. Cukup untuk administrasi/kuliah, kurang cocok untuk edit video atau gaming.'
    }
    elseif ($cpuName -match 'i([3579])-(\d{3,5})([A-Z]{0,3}\d?)') {
        $lvl = [int]$Matches[1]; $num = $Matches[2]; $sfx = $Matches[3]
        if ($num.Length -eq 3) { $gen = 1 }
        elseif ($num.Length -eq 5 -or $num.StartsWith('1')) { $gen = [int]$num.Substring(0, 2) }
        else { $gen = [int]$num.Substring(0, 1) }
        $g = $intelGen[$gen]
        if ($g) { $codename = "$($g[0]) (Gen $gen)"; $cpuYear = $g[1] } else { $codename = "Intel Core Gen $gen"; $cpuYear = 'N/A' }
        $cls = 'D'
        if ($sfx -match 'HX') { $cls = 'HX' }
        elseif ($sfx -match 'H') { $cls = 'H' }
        elseif ($sfx -match '^P') { $cls = 'P' }
        elseif ($sfx -match 'Y') { $cls = 'Y' }
        elseif ($sfx -match '^(M|QM|MQ|MX)') { $cls = 'M' }
        elseif ($sfx -match '^(U|G\d)') { $cls = 'U' }
        $base = $intelBase[$gen]; if (-not $base) { $base = 6 }
        $clsAdj = 0
        if ($cls -eq 'HX') { $clsAdj = 2 }
        elseif ($cls -eq 'H' -and $gen -ge 11) { $clsAdj = 1 }
        elseif ($cls -eq 'P' -and $gen -ge 12) { $clsAdj = 1 }
        elseif ($cls -eq 'Y') { $clsAdj = -2 }
        elseif ($cls -eq 'D' -and $gen -ge 8) { $clsAdj = 1 }
        $tier = Limit ($base + $lvlAdj[$lvl] + $clsAdj) 1 10
        $hybrid = ($gen -ge 12)
    }
    elseif ($cpuName -match 'Celeron|Pentium') {
        if ($cpuName -match 'N5\d{3}|N6\d{3}|J4125|J5040|J6413|Celeron 6305|Celeron 7305|Pentium Gold 7505|Pentium 8505|Pentium 6805') {
            $codename = 'Jasper Lake / Tiger Lake / Alder Lake (Celeron-Pentium baru)'; $cpuYear = 2021; $tier = 3
        }
        elseif ($cpuName -match 'N45\d{2}') { $codename = 'Jasper Lake (Celeron N4500 Series)'; $cpuYear = 2021; $tier = 2 }
        elseif ($cpuName -match 'Pentium.*(Gold|Silver)') { $codename = 'Gemini Lake / Kaby Lake (Pentium Gold-Silver)'; $cpuYear = 2018; $tier = 2 }
        else { $codename = 'Apollo Lake / Gemini Lake / Legacy (Celeron-Pentium lama)'; $cpuYear = 2017; $tier = 1 }
    }
    else { $codename = 'Intel Core / arsitektur tidak dikenali (cek manual)' }
}
elseif ($cpuName -match 'AMD') {
    if ($cpuName -match 'Ryzen\s+AI') {
        $npu = $true; $cpuYear = 2025
        if ($cpuName -match 'Max') { $codename = 'Strix Halo (Zen 5, Ryzen AI Max)'; $tier = 10 }
        elseif ($cpuName -match 'AI\s*([579])\s+(?:HX\s*)?(\d{3})') {
            $lvl = [int]$Matches[1]; $num = $Matches[2]
            if ($num -match '^3[45]') { $codename = 'Krackan Point (Zen 5, Ryzen AI 300)' } else { $codename = 'Strix Point (Zen 5 + Zen 5c, Ryzen AI 300)'; $cpuYear = 2024 }
            if ($lvl -eq 9) { $tier = 9 } elseif ($lvl -eq 7) { $tier = 8 } else { $tier = 7 }
        }
        else { $codename = 'AMD Ryzen AI Series (Zen 5)'; $tier = 8 }
    }
    elseif ($cpuName -match 'Ryzen\s*([3579])\s*(?:PRO\s*)?(\d{4})(\w*)') {
        $lvl = [int]$Matches[1]; $num = $Matches[2]; $sfx = $Matches[3]
        $s = [int]$num.Substring(0, 1); $tail = $num.Substring(2, 2)
        $cpuYear = 2016 + $s
        $cls = 'D'
        if ($sfx -match 'HX') { $cls = 'HX' } elseif ($sfx -match 'H') { $cls = 'H' } elseif ($sfx -match 'U|e$') { $cls = 'U' }
        $base = 5
        if ($s -eq 1) { $codename = 'Zen (Ryzen 1000)'; $base = 3 }
        elseif ($s -eq 2) { $codename = 'Raven Ridge / Zen+ (Ryzen 2000)'; $base = 3 }
        elseif ($s -eq 3) { $codename = 'Picasso (Zen+, Ryzen 3000)'; $base = 4 }
        elseif ($s -eq 4) { $codename = 'Renoir (Zen 2, Ryzen 4000)'; $base = 5 }
        elseif ($s -eq 5) {
            if ($num -match '^5[357]00$' -and $cls -eq 'U') { $codename = 'Lucienne (Zen 2, Ryzen 5000)'; $base = 5 }
            else { $codename = 'Cezanne / Barcelo (Zen 3, Ryzen 5000)'; $base = 6 }
        }
        elseif ($s -eq 6) { $codename = 'Rembrandt (Zen 3+, Ryzen 6000)'; $base = 6 }
        elseif ($s -eq 7) {
            if ($tail -eq '20') { $codename = 'Mendocino (Zen 2, Ryzen 7020)'; $base = 4 }
            elseif ($tail -eq '30') { $codename = 'Barcelo-R (Zen 3, Ryzen 7030)'; $base = 6 }
            elseif ($tail -eq '35') { $codename = 'Rembrandt-R (Zen 3+, Ryzen 7035)'; $base = 6 }
            elseif ($tail -eq '45') { $codename = 'Dragon Range (Zen 4, Ryzen 7045)'; $base = 7 }
            else { $codename = 'Phoenix (Zen 4, Ryzen 7040)'; $base = 6 }
        }
        elseif ($s -eq 8) { $codename = 'Hawk Point (Zen 4, Ryzen 8000)'; $base = 6 }
        else { $codename = 'Zen 5 (Ryzen 9000 / Fire Range)'; $base = 7 }
        if ($num -match '^7[68]40' -or $num -match '^8[4-9]4\d') { $npu = $true }
        $clsAdj = 0
        if ($cls -eq 'HX') { $clsAdj = 1 } elseif ($cls -eq 'H' -and $s -ge 4) { $clsAdj = 1 }
        $tier = Limit ($base + $lvlAdj[$lvl] + $clsAdj) 1 10
    }
    elseif ($cpuName -match 'Athlon|\bA(4|6|8|9|10|12)-\d|\bE[12]-\d|FX-') {
        $codename = 'AMD Athlon / A-Series APU (Legacy)'
        if ($cpuName -match 'Athlon.*Silver|3020e') { $tier = 2 }
        elseif ($cpuName -match 'Athlon') { $tier = 3 }
        elseif ($cpuName -match '\bA(4|6)-|\bE[12]-') { $tier = 1 }
        else { $tier = 2 }
    }
    else { $codename = 'AMD Zen Architecture (cek manual)' }
}
elseif ($cpuName -match 'Snapdragon|Qualcomm') {
    $npu = $true; $cpuYear = 2024; $codename = 'Qualcomm Snapdragon X (Oryon, ARM64)'
    if ($cpuName -match 'Elite') { $tier = 8 } elseif ($cpuName -match 'Plus') { $tier = 7 } else { $tier = 6 }
}

$cpuCategory = "$($tierName[$tier]) (Skor Performa $tier dari 10)"
$equivIntel = $tierIntel[$tier]
$equivAMD = $tierAMD[$tier]
$cpuUse = $tierUse[$tier]

$npuInfo = ''
if ($npu) { $npuInfo = " | <span style='color:#0078d4; font-weight:bold;'>&#129302; Dilengkapi NPU (AI Engine Ready)</span>" }
$cpuHTML = "<strong>$(Enc $cpuName) ($cores Cores, $threads Threads - $(Enc $codename))$npuInfo</strong>"
$cpuExtra = @()
if ($maxGHz -gt 0) { $cpuExtra += "Clock dasar $maxGHz GHz" }
if ($l3MB -gt 0) { $cpuExtra += "Cache L3 $l3MB MB" }
if ($cpuYear -ne 'N/A') { $cpuExtra += "Rilis +/- $cpuYear" }

# ------------------------------------------------------------------------------
# 3. GPU & VRAM (VRAM diskrit dibaca dari registry agar tidak terpotong 4 GB)
# ------------------------------------------------------------------------------
$clsKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e968-e325-11ce-bfc1-08002be10318}'
$regGpu = Get-ChildItem $clsKey | Where-Object { $_.PSChildName -match '^\d{4}$' } | ForEach-Object { Get-ItemProperty $_.PSPath }
$gpuList = Get-CimInstance Win32_VideoController | Where-Object { $_.Name -and $_.Name -notmatch 'Microsoft Basic|Remote|Virtual|Parsec|DisplayLink|Meta' }
$gpuArr = @(); $gpuDrivers = @()
$hasDedicated = $false
foreach ($g in $gpuList) {
    $gName = ($g.Name -replace '\s+', ' ').Trim()
    $isDedicated = ($gName -match 'NVIDIA|GeForce|Quadro|Radeon\s*(RX|PRO)|Arc.*\b[AB]\d{3}')
    if ($isDedicated) { $hasDedicated = $true }
    $vramText = 'Shared / Integrated'
    if ($isDedicated) {
        $q = ($regGpu | Where-Object { $_.DriverDesc -eq $g.Name } | Select-Object -First 1).'HardwareInformation.qwMemorySize'
        if ($q -is [byte[]]) { $q = [BitConverter]::ToInt64($q, 0) }
        if (-not $q -and $g.AdapterRAM) { $q = [int64]$g.AdapterRAM }
        if ($q -gt 0) {
            if ($q -ge 1GB) { $vramText = "$([math]::Round($q / 1GB, 1)) GB Dedicated VRAM" }
            else { $vramText = "$([math]::Round($q / 1MB)) MB Dedicated VRAM" }
        } else { $vramText = 'Dedicated VRAM (ukuran tidak terbaca)' }
    }
    if ($g.DriverVersion) { $gpuDrivers += "$gName : Driver $($g.DriverVersion)" }
    $gpuArr += "$(Enc $gName) ($vramText - $(Get-GpuClass $gName))"
}
$gpu = if ($gpuArr.Count -gt 0) { $gpuArr -join ' + ' } else { 'Standard Display Adapter' }

# ------------------------------------------------------------------------------
# 4. Layar (resolusi, ukuran inci, pabrikan panel, tahun produksi panel dari EDID)
# ------------------------------------------------------------------------------
$vc = Get-CimInstance Win32_VideoController | Where-Object { $_.CurrentHorizontalResolution } | Select-Object -First 1
$resX = [int]$vc.CurrentHorizontalResolution
$resY = [int]$vc.CurrentVerticalResolution
$refreshRate = [int]$vc.CurrentRefreshRate

if ($resX -le 0) { $resCategory = 'Tidak terdeteksi' }
elseif ($resX -ge 3840 -or $resY -ge 2160) { $resCategory = '4K UHD (3840x2160 High-Res)' }
elseif ($resX -ge 2880 -or $resY -ge 1800) { $resCategory = '3K / 2.8K High-Res' }
elseif ($resX -ge 2560 -or $resY -ge 1440) { if ($resY -eq 1600) { $resCategory = 'WQXGA (2560x1600 - 16:10 High-Res)' } else { $resCategory = 'QHD / 2K (2560x1440)' } }
elseif ($resX -ge 1920) { if ($resY -eq 1200) { $resCategory = 'WUXGA (1920x1200 - 16:10)' } elseif ($resY -eq 1080) { $resCategory = 'FHD (1920x1080 Standard)' } else { $resCategory = 'Custom Wide FHD' } }
elseif ($resX -ge 1366 -or $resY -ge 768) { $resCategory = 'HD (1366x768 Basic)' }
else { $resCategory = 'Standard Low-Res' }

$displayRes = if ($resX -gt 0) { "${resX}x${resY} @ ${refreshRate}Hz ($resCategory)" } else { 'Tidak terdeteksi' }

$conn = @(Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorConnectionParams)
$intl = $conn | Where-Object { ([int64]$_.VideoOutputTechnology) -in @(6, 11, 13, 2147483648, -2147483648) } | Select-Object -First 1
if (-not $intl) { $intl = $conn | Select-Object -First 1 }
$inst = $intl.InstanceName
$monIds = @(Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID)
$mid = $monIds | Where-Object { $_.InstanceName -eq $inst } | Select-Object -First 1
if (-not $mid) { $mid = $monIds | Select-Object -First 1 }
$bp = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams | Where-Object { $_.InstanceName -eq $mid.InstanceName } | Select-Object -First 1

$mfr = Dec $mid.ManufacturerName
$pcode = Dec $mid.ProductCodeID
$monName = Dec $mid.UserFriendlyName
$panelMaker = @{ AUO='AU Optronics (AUO)'; BOE='BOE'; CMN='Chimei Innolux (CMN)'; CMO='Chimei Innolux (CMN)'; IVO='InfoVision (IVO)'; SDC='Samsung Display (SDC)'; LGD='LG Display (LGD)'; SHP='Sharp'; CSO='CSOT'; TMX='Tianma'; HKC='HKC'; LEN='Lenovo (OEM)'; AUS='ASUS (OEM)' }
$panelID = if ($mfr) { "$mfr$pcode" } else { 'Standard Panel' }
$makerText = if ($mfr -and $panelMaker[$mfr]) { $panelMaker[$mfr] } elseif ($mfr) { $mfr } else { 'N/A' }

$inch = 0
if ($bp -and $bp.MaxHorizontalImageSize -gt 0 -and $bp.MaxVerticalImageSize -gt 0) {
    $inch = [math]::Round([math]::Sqrt([math]::Pow($bp.MaxHorizontalImageSize, 2) + [math]::Pow($bp.MaxVerticalImageSize, 2)) / 2.54, 1)
}
$ppi = 0
if ($inch -gt 0 -and $resX -gt 0) { $ppi = [math]::Round([math]::Sqrt($resX * $resX + $resY * $resY) / $inch) }
$panelProd = ''
if ($mid.YearOfManufacture) { $panelProd = "$($mid.YearOfManufacture) (minggu ke-$($mid.WeekOfManufacture))" }

# Penilaian gamut panel mandiri dalam satu file (tanpa JSON terpisah).
# Catatan: Panelook bukan API publik yang stabil; EDID laptop sering hanya mengungkap
# kode OEM seperti BOE0xxx, bukan SKU panel lengkap. Karena itu hanya gunakan data
# Panelook yang cocok persis. Jika tidak cocok, pakai estimasi kromatisitas EDID dan
# jangan menebak IPS/TN/VA atau angka coverage.
$panelType = 'Tidak dapat dipastikan dari EDID Windows'
if ($model -match 'OLED' -or $monName -match 'OLED' -or $panelID -match 'OLED') {
    $panelType = 'Indikasi OLED dari nama perangkat (perlu verifikasi SKU panel)'
}
$gamutInfo = 'Data gamut terverifikasi belum ditemukan'
$gamutSource = ''
$gamutStandard = 'Belum dapat dinilai tanpa data NTSC terverifikasi'

# Database Panelook tertanam di skrip ini. Tambahkan entri baru pada array ini bila
# SKU panel sudah diverifikasi. Angka contoh hanya berlaku untuk model yang cocok.
$panelDBEntries = @(
    [pscustomobject]@{
        Model = 'NV156FHM-AW1'
        Aliases = @()
        PanelType = 'ADS (IPS-like; listing Panelook)'
        NTSC = 72
        sRGB = 94
        AdobeRGB = 73
        DCI_P3 = 74
        BrightnessNits = 300
        Resolution = '1920x1080'
        Source = 'https://www.panelook.com/NV156FHM-AW1_BOE_15.6_LCM_parameter_31946.html'
    }
)

# Ambil SKU panel jika EDID memang memuat string seperti NV156FHM-AW1.
$panelModelCandidate = ''
if ($monName -match '(?i)\b[A-Z]{2,5}\d{3,5}[A-Z0-9-]*-[A-Z0-9-]+\b') {
    $panelModelCandidate = $Matches[0]
}
$panelLookupValues = @([string]$panelModelCandidate, [string]$panelID, [string]$monName, [string]$pcode) |
    ForEach-Object { $_.Trim() } | Where-Object { $_ }
$matchedPanel = $null
foreach ($entry in $panelDBEntries) {
    $keys = @([string]$entry.Model) + @($entry.Aliases | ForEach-Object { [string]$_ })
    foreach ($key in $keys) {
        $key = $key.Trim()
        if ($key -and ($panelLookupValues | Where-Object { $_ -ieq $key }).Count -gt 0) {
            $matchedPanel = $entry
            break
        }
    }
    if ($matchedPanel) { break }
}
if ($matchedPanel) {
    if ($matchedPanel.PanelType) { $panelType = [string]$matchedPanel.PanelType }
    $gamutBits = @()
    foreach ($key in @('NTSC','sRGB','AdobeRGB','DCI_P3')) {
        $value = $matchedPanel.$key
        if ($null -ne $value -and "$value" -ne '') {
            $label = if ($key -eq 'DCI_P3') { 'DCI-P3' } elseif ($key -eq 'AdobeRGB') { 'Adobe RGB' } else { $key }
            $gamutBits += ("{0}: {1}%" -f $label, $value)
        }
    }
    if ($matchedPanel.BrightnessNits) { $gamutBits += "Kecerahan: $($matchedPanel.BrightnessNits) nit" }
    if ($matchedPanel.Resolution) { $gamutBits += "Resolusi panel: $($matchedPanel.Resolution)" }
    if ($gamutBits.Count -gt 0) {
        $gamutInfo = $gamutBits -join ' | '
        $gamutSource = [string]$matchedPanel.Source
    }
    if ($null -ne $matchedPanel.NTSC -and "$($matchedPanel.NTSC)" -ne '') {
        $ntscValue = [double]$matchedPanel.NTSC
        if ($ntscValue -lt 50) { $gamutStandard = 'Gamut sempit (acuan praktis)' }
        elseif ($ntscValue -lt 65) { $gamutStandard = 'Gamut menengah (acuan praktis)' }
        elseif ($ntscValue -lt 72) { $gamutStandard = 'Gamut menengah-lebar (acuan praktis)' }
        elseif ($ntscValue -lt 85) { $gamutStandard = 'Wide gamut (acuan praktis)' }
        else { $gamutStandard = 'Very wide gamut (acuan praktis)' }
    }
}

# Fallback EDID: estimasi luas segitiga RGB terhadap segitiga NTSC 1953.
# Ini bukan angka coverage sRGB/NTSC hasil kolorimeter dan bukan pengganti HWiNFO.
if (-not $matchedPanel) {
    try {
        $edidBlocks = @(Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorRawEEdidV1Block | Where-Object { $_.InstanceName -eq $mid.InstanceName })
        $edid0 = $edidBlocks | Where-Object { [int]$_.BlockType -eq 0 } | Select-Object -First 1
        if ($edid0 -and $edid0.BlockContent) { $edidBytes = [byte[]]$edid0.BlockContent } else { $edidBytes = $null }
        if ($edidBytes -and $edidBytes.Length -ge 35) {
            $lo1 = [int]$edidBytes[25]; $lo2 = [int]$edidBytes[26]
            $rx = (([int]$edidBytes[27] -shl 2) -bor (($lo1 -shr 6) -band 3)) / 1024.0
            $ry = (([int]$edidBytes[28] -shl 2) -bor (($lo1 -shr 4) -band 3)) / 1024.0
            $gx = (([int]$edidBytes[29] -shl 2) -bor (($lo1 -shr 2) -band 3)) / 1024.0
            $gy = (([int]$edidBytes[30] -shl 2) -bor ($lo1 -band 3)) / 1024.0
            $bx = (([int]$edidBytes[31] -shl 2) -bor (($lo2 -shr 6) -band 3)) / 1024.0
            $by = (([int]$edidBytes[32] -shl 2) -bor (($lo2 -shr 4) -band 3)) / 1024.0
            $rgbArea = [math]::Abs(($rx*($gy-$by) + $gx*($by-$ry) + $bx*($ry-$gy)) / 2.0)
            $ntscArea = [math]::Abs((0.67*(0.71-0.08) + 0.21*(0.08-0.33) + 0.14*(0.33-0.71)) / 2.0)
            if ($rgbArea -gt 0 -and $ntscArea -gt 0) {
                $areaPct = [math]::Round(($rgbArea / $ntscArea) * 100)
                if ($areaPct -gt 0 -and $areaPct -le 200) {
                    $gamutInfo = "NTSC area ratio ~${areaPct}% (estimasi EDID; BUKAN coverage terukur)"
                    $gamutSource = 'EDID chromaticity estimate'
                    if ($areaPct -lt 50) { $gamutStandard = 'Perkiraan gamut relatif sempit; bukan standar coverage' }
                    elseif ($areaPct -lt 65) { $gamutStandard = 'Perkiraan gamut relatif menengah; bukan standar coverage' }
                    elseif ($areaPct -lt 72) { $gamutStandard = 'Perkiraan gamut relatif menengah-lebar; bukan standar coverage' }
                    elseif ($areaPct -lt 85) { $gamutStandard = 'Perkiraan gamut relatif lebar; bukan standar coverage' }
                    else { $gamutStandard = 'Perkiraan gamut relatif sangat lebar; bukan standar coverage' }
                }
            }
        }
    } catch { }
}

$dispExtra = ''
if ($inch -gt 0) { $dispExtra += " | Ukuran: $inch inci"; if ($ppi -gt 0) { $dispExtra += " (~$ppi PPI)" } }
if ($makerText -ne 'N/A') { $dispExtra += " | Pabrikan Panel: $(Enc $makerText)" }
if ($panelProd) { $dispExtra += " | Produksi Panel: $panelProd" }
if ($gamutSource) { $dispExtra += " | Sumber gamut: $(Enc $gamutSource)" }
if ($gamutStandard -and $gamutStandard -ne 'Belum dapat dinilai tanpa data NTSC terverifikasi') { $dispExtra += " | Penilaian gamut: $(Enc $gamutStandard)" }
$displayHTML = "<strong>$displayRes</strong><br><small style='color:#555;'>Tipe Panel: $panelType | ID: $(Enc $panelID) | Gamut: $(Enc $gamutInfo)$dispExtra</small>"

# ------------------------------------------------------------------------------
# 5. RAM (tipe, kecepatan, slot, on-board vs SODIMM - dibaca per modul)
# ------------------------------------------------------------------------------
$ramMods = @(Get-CimInstance Win32_PhysicalMemory)
$ramArr = Get-CimInstance Win32_PhysicalMemoryArray | Select-Object -First 1
$ramGB = [math]::Round($sys.TotalPhysicalMemory / 1GB)
$ddrType = 'Tipe tidak terbaca'; $ramSpeed = 0; $ramMaker = ''; $ramPart = ''
$soldered = $false
$slotStatusText = 'Detail modul tidak dilaporkan BIOS (cek fisik)'
$ddrMap = @{ 24='DDR3'; 26='DDR4'; 27='LPDDR'; 28='LPDDR2'; 29='LPDDR3'; 30='LPDDR4'; 34='DDR5'; 35='LPDDR5' }
# Model yang RAM-nya dipastikan SOLDERED (tidak bisa upgrade). Tambahkan model lain di sini, pisahkan dengan |
$ramSolderedModels = 'E1404GA'
$chipPattern = '^(K4[ABUE]|K3[LKU]|H5A[NC]|H5TQ|H9[HJ]C|H5[48]G|MT4[01]A|MT53|MT6[12]|NT[56]A|W6|CX|M15)'
$modPattern = '^(M4[27][1-5]|M323|M425|HMA|HMCG|HMT|MTA|MTC|CT\d|CB\d|KVR|KSM|KF\d|99[0-9U]|CM[SKW]|AD4S|AO1|ADDS|TED|TF\d|TEAM|PSD|TS\d|RMSA|RMUA|F4-|BLS|BL\d)'
$unverified = 0
$slotsTotal = 0; $slotsFree = 0; $ramMaxText = ''; $onboardGB = 0; $socketGB = 0; $socketCount = 0; $onboardCount = 0

if ($ramMods.Count -gt 0 -and $ramMods[0].Capacity) {
    $ramGB = [math]::Round((($ramMods | Measure-Object -Property Capacity -Sum).Sum) / 1GB)
    $first = $ramMods[0]
    $ramSpeed = if ($first.ConfiguredClockSpeed) { [int]$first.ConfiguredClockSpeed } else { [int]$first.Speed }
    $ramMaker = ([string]$first.Manufacturer).Trim()
    $ramPart = ([string]$first.PartNumber).Trim()
    foreach ($m in $ramMods) {
        $smb = [int]$m.SMBIOSMemoryType
        $t = $ddrMap[$smb]
        $part = [string]$m.PartNumber
        $spd = if ($m.ConfiguredClockSpeed) { [int]$m.ConfiguredClockSpeed } else { [int]$m.Speed }
        if (-not $t) {
            if ($part -match 'LPDDR5') { $t = 'LPDDR5' }
            elseif ($part -match 'LPDDR4') { $t = 'LPDDR4' }
            elseif ($part -match 'DDR5') { $t = 'DDR5' }
            elseif ($part -match 'DDR4') { $t = 'DDR4' }
            elseif ($spd -ge 4800) { $t = 'DDR5' }
            elseif ($spd -ge 2133) { $t = 'DDR4' }
        }
        $ff = [int]$m.FormFactor
        $loc = [string]$m.DeviceLocator
        $pn = $part.Trim()
        $isSold = $false
        if ($t -like 'LPDDR*') { $isSold = $true }
        elseif ($model -match $ramSolderedModels) { $isSold = $true }
        elseif ($loc -match 'onboard|on-board|soldered') { $isSold = $true }
        elseif ($pn -match $chipPattern) { $isSold = $true }
        elseif ($pn -match $modPattern) { $isSold = $false }
        elseif ($ff -in 9, 14, 16, 17, 18, 21, 22) { $isSold = $true }
        else { if ($ff -in 8, 12) { $unverified++ } }
        $gb = [math]::Round($m.Capacity / 1GB)
        if ($isSold) { $onboardGB += $gb; $onboardCount++ } else { $socketGB += $gb; $socketCount++ }
        if ($m -eq $first -and $t) { $ddrType = $t }
    }
    $soldered = ($socketCount -eq 0)
    $slotsTotal = if ($ramArr.MemoryDevices) { [int]$ramArr.MemoryDevices } else { $ramMods.Count }
    $slotsFree = [math]::Max(0, $slotsTotal - $ramMods.Count)
    $formName = 'SODIMM'; if ([int]$first.FormFactor -eq 8) { $formName = 'DIMM' }
    if ($soldered) {
        $slotStatusText = '&#10060; Non-Upgradable / Permanen (On-Board Soldered - No Slots)'
    }
    elseif ($onboardCount -gt 0) {
        $slotStatusText = "&#10004;&#65039; Hybrid: $onboardGB GB On-Board (permanen) + $socketCount Slot $formName Terisi"
        if ($slotsFree -gt 0) { $slotStatusText += " (Tersedia $slotsFree Slot Kosong untuk Upgrade)" }
    }
    elseif ($slotsFree -gt 0) {
        $slotStatusText = "&#10004;&#65039; $socketCount Slot $formName Terisi (Tersedia $slotsFree Slot Kosong untuk Upgrade)"
        if ($unverified -gt 0) { $slotStatusText += " - data BIOS, verifikasi fisik" }
    }
    else {
        if ($socketCount -ge 2) { $slotStatusText = "&#10004;&#65039; $socketCount Slot $formName Terisi (Dual Channel / Full)" }
        else { $slotStatusText = "&#10004;&#65039; $socketCount Slot $formName Terisi (Slot Penuh)" }
    }
    if (-not $soldered) {
        $maxKB = [double]$ramArr.MaxCapacity
        if ($maxKB -le 0) { $maxKB = [double]$ramArr.MaxCapacityEx }
        $maxGB = [math]::Round($maxKB / 1MB)
        if ($maxGB -gt 0 -and $maxGB -le 2048) { $ramMaxText = "Maks. dukungan BIOS $maxGB GB" }
    }
}
$ramSpeedText = if ($ramSpeed -gt 0) { " @ $ramSpeed MHz" } else { '' }
$ramDetail = "$ramGB GB $ddrType$ramSpeedText<br><small style='color:#555;'>Status: $slotStatusText</small>"
$channelText = ''
if (-not $soldered -and $socketCount -ge 2) { $channelText = 'Dual Channel' } elseif (-not $soldered -and $socketCount -eq 1) { $channelText = 'Single Channel' }
$slotsUsed = $socketCount

# ------------------------------------------------------------------------------
# 6. Penyimpanan MULTI-DISK: SSD/NVMe/SATA/eMMC/HDD + SMART per perangkat
#    Catatan: angka health/wear hanya ditampilkan bila sumbernya dapat dibaca.
# ------------------------------------------------------------------------------
$allDisks = @(Get-CimInstance Win32_DiskDrive | Sort-Object Index)
$bootIdx = (Get-CimInstance Win32_DiskPartition -Filter 'BootPartition=True' | Select-Object -First 1).DiskIndex
$diskReports = @()
$diskSmall = @()
$primaryDisk = $allDisks | Where-Object { $_.Index -eq $bootIdx } | Select-Object -First 1
if (-not $primaryDisk) { $primaryDisk = $allDisks | Where-Object { $_.InterfaceType -ne 'USB' } | Select-Object -First 1 }

if (-not ('NvmeSmart' -as [type])) {
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
public static class NvmeSmart {
  [DllImport("kernel32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
  static extern SafeFileHandle CreateFile(string f, uint a, uint s, IntPtr sec, uint cd, uint fl, IntPtr t);
  [DllImport("kernel32.dll", SetLastError=true)]
  static extern bool DeviceIoControl(SafeFileHandle h, uint code, byte[] inB, int inSz, byte[] outB, int outSz, out int ret, IntPtr ov);
  public static byte[] Read(int disk, uint sub) {
    using (var h = CreateFile(@"\\.\PhysicalDrive" + disk, 0xC0000000, 3, IntPtr.Zero, 3, 0, IntPtr.Zero)) {
      if (h.IsInvalid) return null;
      byte[] buf = new byte[560];
      BitConverter.GetBytes(50).CopyTo(buf, 0);
      BitConverter.GetBytes(0).CopyTo(buf, 4);
      BitConverter.GetBytes(3).CopyTo(buf, 8);
      BitConverter.GetBytes(2).CopyTo(buf, 12);
      BitConverter.GetBytes(2).CopyTo(buf, 16);
      BitConverter.GetBytes(sub).CopyTo(buf, 20);
      BitConverter.GetBytes(40).CopyTo(buf, 24);
      BitConverter.GetBytes(512).CopyTo(buf, 28);
      int ret;
      if (!DeviceIoControl(h, 0x2D1400, buf, buf.Length, buf, buf.Length, out ret, IntPtr.Zero)) return null;
      byte[] log = new byte[512];
      Array.Copy(buf, 48, log, 0, 512);
      return log;
    }
  }
}
'@
}

foreach ($d in $allDisks) {
    $cap = 'N/A'
    $rawGB = if ($d.Size) { [double]$d.Size / 1GB } else { 0 }
    if ($rawGB -gt 3500) { $cap = '4 TB' }
    elseif ($rawGB -gt 1700) { $cap = '2 TB' }
    elseif ($rawGB -gt 850) { $cap = '1 TB' }
    elseif ($rawGB -gt 420) { $cap = '512 GB' }
    elseif ($rawGB -gt 210) { $cap = '256 GB' }
    elseif ($rawGB -gt 110) { $cap = '128 GB' }
    elseif ($rawGB -gt 0) { $cap = "$([math]::Round($rawGB)) GB" }

    $pd = @(Get-PhysicalDisk | Where-Object {
        ($_.DeviceId -ne $null -and [string]$_.DeviceId -eq [string]$d.Index) -or
        ($_.FriendlyName -and $d.Model -and $_.FriendlyName.Trim() -eq $d.Model.Trim())
    } | Select-Object -First 1)
    $pd = $pd | Select-Object -First 1
    $rel = $null
    if ($pd) { $rel = $pd | Get-StorageReliabilityCounter }

    $modelDisk = ([string]$d.Model).Trim()
    $media = [string]$pd.MediaType
    $bus = [string]$pd.BusType
    if (-not $bus -or $bus -eq 'Unknown') {
        if ($d.InterfaceType -match 'SCSI|IDE|USB') { $bus = $d.InterfaceType }
        if ($modelDisk -match 'NVMe') { $bus = 'NVMe' }
        if ($modelDisk -match 'eMMC|MMC') { $bus = 'eMMC' }
    }
    if ($media -notmatch 'SSD|HDD') {
        if ($bus -eq 'NVMe' -or $modelDisk -match 'SSD|NVMe|eMMC|MMC|Flash') { $media = 'SSD/Flash (jenis perlu verifikasi)' }
        elseif ($modelDisk -match 'HDD|Hard.?Disk') { $media = 'HDD' }
        else { $media = 'Tidak diketahui' }
    }
    if ($modelDisk -match 'eMMC|MMC') { $media = 'eMMC/Flash' }
    $diskType = "$bus $media".Trim()

    $wear = $null; $healthSource = ''; $smartTemp = $null; $smartPoh = $null; $smartPowerCycles = $null
    $smartStatus = 'UNKNOWN'; $smartNotes = @()
    $winWear = $null
    if ($rel -and $null -ne $rel.Wear) { $winWear = [int]$rel.Wear }

    if ($bus -eq 'NVMe') {
        $log = [NvmeSmart]::Read([int]$d.Index, 0)
        if (-not $log) { $log = [NvmeSmart]::Read([int]$d.Index, [uint32]4294967295) }
        if ($log) {
            $used = [int]$log[5]
            $spare = [int]$log[3]; $spareThreshold = [int]$log[4]; $critical = [int]$log[0]
            if ($used -ge 0 -and $used -le 255) {
                $wear = [math]::Min(100, $used)
                $healthSource = 'NVMe SMART Percentage Used (endurance terpakai)'
                if ($used -gt 100) { $smartNotes += "Endurance terpakai $used% (melewati rating endurance pabrik)" }
            }
            $kelvin = [int]$log[1] + 256 * [int]$log[2]
            if ($kelvin -gt 200 -and $kelvin -lt 1000) { $smartTemp = $kelvin - 273 }
            $duRead = [BitConverter]::ToUInt64($log, 32)
            $duWrite = [BitConverter]::ToUInt64($log, 48)
            $smartNotes += "Host read: $([math]::Round($duRead * 512000 / 1e12, 2)) TB"
            $smartNotes += "Host write: $([math]::Round($duWrite * 512000 / 1e12, 2)) TB"
            $mediaErrors = [BitConverter]::ToUInt64($log, 160)
            $unsafe = [BitConverter]::ToUInt64($log, 144)
            $smartNotes += "Media/Data Integrity Errors: $mediaErrors"
            $smartNotes += "Unsafe Shutdowns: $unsafe"
            $smartNotes += "Available Spare: $spare% (ambang $spareThreshold%)"
            $smartNotes += "Critical Warning: 0x$('{0:X2}' -f $critical)"
            $smartPoh = [BitConverter]::ToUInt64($log, 128)
            $smartPowerCycles = [BitConverter]::ToUInt64($log, 112)
            if (($critical -band 0x0C) -ne 0 -or $mediaErrors -gt 0) { $smartStatus = 'BAD' }
            elseif ($critical -ne 0 -or ($spareThreshold -gt 0 -and $spare -lt $spareThreshold) -or $used -ge 90) { $smartStatus = 'CAUTION' }
            else { $smartStatus = 'GOOD' }
        }
    }

    # ATA SMART: gunakan atribut wear hanya jika ID vendor lazim dan nilai valid.
    if ($bus -ne 'NVMe' -and $media -match 'SSD|HDD') {
        $fp = @(Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictData)
        $one = $fp | Where-Object { $_.InstanceName -like "$($d.PNPDeviceID)*" } | Select-Object -First 1
        if (-not $one -and $fp.Count -eq 1) { $one = $fp[0] }
        if ($one -and $one.VendorSpecific) {
            $vs = $one.VendorSpecific
            for ($i = 2; $i -lt 362; $i += 12) {
                $id = [int]$vs[$i]; $cur = [int]$vs[$i + 3]
                $raw = 0L
                if ($i + 7 -lt $vs.Length) {
                    for ($j = 0; $j -lt 6; $j++) { $raw += ([int64]$vs[$i + 5 + $j] -shl (8 * $j)) }
                }
                if ($id -in 5, 9, 12, 187, 194, 197, 198, 199, 200) {
                    $labels = @{5='Reallocated sectors/blocks';9='Power-on hours (raw/vendor-specific)';12='Power cycles';187='Reported uncorrectable errors';194='Temperature (raw/vendor-specific)';197='Current pending sectors';198='Offline uncorrectable sectors';199='UDMA CRC errors';200='Write error rate (vendor-specific)'}
                    $smartNotes += "ATA $($labels[[int]$id]) (ID $id): $raw"
                    if ($raw -gt 0 -and $id -in 5,187,197,198) { $smartStatus = 'CAUTION' }
                    if ($id -eq 194 -and $raw -ge 1 -and $raw -le 100) { $smartTemp = [int]($raw -band 0xFF) }
                }
                # Attribute meaning differs by vendor; do not convert generic normalized values into a universal health percentage.
                if ($id -eq 233 -and $cur -ge 0 -and $cur -le 100 -and $modelDisk -match 'Intel') {
                    $wear = 100 - $cur; $healthSource = 'Intel SMART attribute 233 (vendor-specific)'; break
                }
            }
        }
    }
    if ($bus -ne 'NVMe') {
        try {
            $fpStatus = @(Get-CimInstance -Namespace root\wmi -ClassName MSStorageDriver_FailurePredictStatus)
            $matchedStatus = $fpStatus | Where-Object { $_.InstanceName -like "$($d.PNPDeviceID)*" } | Select-Object -First 1
            if ($matchedStatus -and $matchedStatus.PredictFailure) {
                $smartStatus = 'BAD'
                $smartNotes += 'SMART Failure Prediction: FAIL (driver melaporkan potensi kegagalan)'
            }
        } catch { }
    }

    if ($null -eq $wear -and $null -ne $winWear -and $winWear -ge 0 -and $winWear -le 100) {
        $wear = $winWear; $healthSource = 'Windows Storage Wear'
    }

    $healthText = 'Health/wear tidak tersedia dari driver'
    if ($null -ne $wear) { $healthText = "Perkiraan sisa endurance $([math]::Max(0,100-$wear))% (wear $wear%; sumber: $healthSource)" }
    elseif ($pd -and $pd.HealthStatus) { $healthText = "Status perangkat: $($pd.HealthStatus); persentase tidak terbaca" }
    if ($rel -and $rel.Temperature -and -not ($smartNotes -match '^Suhu')) { $smartNotes += "Suhu: $($rel.Temperature) C" }
    elseif ($null -ne $smartTemp -and -not ($smartNotes -match '^Suhu')) { $smartNotes += "Suhu SMART: $smartTemp C" }
    if ($rel -and $rel.PowerOnHours) { $smartNotes += "Jam nyala: $($rel.PowerOnHours)" }
    elseif ($smartPoh -gt 0) { $smartNotes += "Jam nyala SMART: $smartPoh" }
    if ($null -ne $smartPowerCycles) { $smartNotes += "Power cycles: $smartPowerCycles" }
    if ($smartStatus -eq 'UNKNOWN' -and $pd -and $pd.HealthStatus) {
        if ([string]$pd.HealthStatus -eq 'Healthy') { $smartStatus = 'GOOD' }
        elseif ([string]$pd.HealthStatus -match 'Warning') { $smartStatus = 'CAUTION' }
        elseif ([string]$pd.HealthStatus -match 'Unhealthy') { $smartStatus = 'BAD' }
    }
    $smartStatusText = switch ($smartStatus) {
        'GOOD' { 'SEHAT (SMART/Windows)' }
        'CAUTION' { 'PERINGATAN (perlu pemeriksaan)' }
        'BAD' { 'BURUK (backup data segera)' }
        default { 'TIDAK DIKETAHUI (SMART tidak terbaca penuh)' }
    }

    $diskName = if ($modelDisk) { $modelDisk } else { [string]$d.Caption }
    $diskReports += "$($d.Index): $cap - $(Enc $diskName) [$diskType] - $smartStatusText - $healthText"
    $detailLine = "Disk $($d.Index): $cap ($diskName) | Tipe: $diskType | Status SMART (model sederhana ala CrystalDiskInfo): $smartStatusText | $healthText"
    if ($smartNotes.Count -gt 0) { $detailLine += ' | ' + ($smartNotes -join ' | ') }
    $diskSmall += $detailLine
}

$disk = $primaryDisk
$diskCap = 'N/A'; $diskCaption = 'Unknown Disk'; $diskTypeText = 'Tidak diketahui'
$isHDD = $false; $wearVal = $null; $poh = 0
if ($primaryDisk) {
    $diskCaption = ([string]$primaryDisk.Model).Trim()
    $pdMain = Get-PhysicalDisk | Where-Object { $_.FriendlyName -and $_.FriendlyName.Trim() -eq $diskCaption } | Select-Object -First 1
    $rawMain = if ($primaryDisk.Size) { [double]$primaryDisk.Size / 1GB } else { 0 }
    if ($rawMain -gt 3500) { $diskCap = '4 TB' } elseif ($rawMain -gt 1700) { $diskCap = '2 TB' } elseif ($rawMain -gt 850) { $diskCap = '1 TB' } elseif ($rawMain -gt 420) { $diskCap = '512 GB' } elseif ($rawMain -gt 210) { $diskCap = '256 GB' } elseif ($rawMain -gt 110) { $diskCap = '128 GB' } elseif ($rawMain -gt 0) { $diskCap = "$([math]::Round($rawMain)) GB" }
    $diskBusMain = if ($pdMain) { [string]$pdMain.BusType } else { [string]$primaryDisk.InterfaceType }
    $diskMediaMain = if ($pdMain) { [string]$pdMain.MediaType } else { 'Tidak diketahui' }
    $diskTypeText = "$diskBusMain $diskMediaMain".Trim()
    $isHDD = ($diskMediaMain -eq 'HDD' -or $primaryDisk.Model -match 'HDD|Hard.?Disk')
    $mainLine = $diskReports | Where-Object { $_ -like "$($primaryDisk.Index):*" } | Select-Object -First 1
    if ($mainLine -match 'wear (\d+)%') { $wearVal = [int]$Matches[1] }
    if ($mainLine -match 'Jam nyala SMART: (\d+)') { $poh = [int]$Matches[1] }
}
$ssdDetail = if ($diskReports.Count -gt 0) { $diskReports -join '<br>' } else { 'Tidak ada disk fisik yang terdeteksi oleh Windows.' }
$ssdHTML = "<strong>$ssdDetail</strong>"
if ($diskSmall.Count -eq 0) { $diskSmall = @('Tidak ada data disk yang tersedia dari Windows.') }
# ------------------------------------------------------------------------------
# 7. Baterai (WMI = instan; powercfg hanya cadangan)
# ------------------------------------------------------------------------------
$designMwh = 0; $fullMwh = 0; $cycles = 0; $healthPct = $null
$hasBat = [bool](Get-CimInstance Win32_Battery | Select-Object -First 1)
if ($hasBat) {
    $bs = Get-CimInstance -Namespace root\wmi -ClassName BatteryStaticData | Select-Object -First 1
    $bf = Get-CimInstance -Namespace root\wmi -ClassName BatteryFullChargedCapacity | Select-Object -First 1
    $bc = Get-CimInstance -Namespace root\wmi -ClassName BatteryCycleCount | Select-Object -First 1
    if ($bs.DesignedCapacity) { $designMwh = [int]$bs.DesignedCapacity }
    if ($bf.FullChargedCapacity) { $fullMwh = [int]$bf.FullChargedCapacity }
    if ($bc.CycleCount) { $cycles = [int]$bc.CycleCount }
    if ($designMwh -le 0 -or $fullMwh -le 0) {
        $batPath = Join-Path -Path $env:TEMP -ChildPath 'bat_temp.html'
        & powercfg /batteryreport /output $batPath | Out-Null
        if (Test-Path $batPath) {
            $htmlRaw = Get-Content $batPath -Raw
            if ($htmlRaw -match 'DESIGN CAPACITY[\s\S]*?([\d\.,]+)\s*mWh') { $designMwh = [int]($Matches[1] -replace '[\.,]', '') }
            if ($htmlRaw -match 'FULL CHARGE CAPACITY[\s\S]*?([\d\.,]+)\s*mWh') { $fullMwh = [int]($Matches[1] -replace '[\.,]', '') }
            if ($cycles -le 0 -and $htmlRaw -match 'CYCLE COUNT[\s\S]*?(\d+)') { $cycles = [int]$Matches[1] }
            Remove-Item $batPath -Force
        }
    }
}
if ($designMwh -gt 0 -and $fullMwh -gt 0) {
    $healthPct = [math]::Round(($fullMwh / $designMwh) * 100, 1)
    if ($healthPct -gt 100) { $healthPct = 100 }
    $wearLevel = [math]::Round(100 - $healthPct, 1)
    $cycleText = if ($cycles -gt 0) { " | Cycle Count: $cycles" } else { '' }
    $batDisplay = "<strong style='color:#2e7d32;'>Health: ${healthPct}%</strong> (Design: ${designMwh} mWh | Full Charge: ${fullMwh} mWh | Wear Level: ${wearLevel}%$cycleText)"
}
elseif ($hasBat) { $batDisplay = "<span style='color:#2e7d32;'>Normal / Terpasang</span>" }
else { $batDisplay = "<span style='color:#2e7d32;'>Normal / Terpasang</span>" }

# ------------------------------------------------------------------------------
# 8. Deteksi perangkat (QC otomatis) - bukan uji fisik
# ------------------------------------------------------------------------------
$pnp = @(Get-CimInstance Win32_PnPEntity -Filter "PNPClass='Bluetooth' OR PNPClass='Camera' OR PNPClass='Image'")
$wifi = Get-CimInstance Win32_NetworkAdapter -Filter 'PhysicalAdapter=True' | Where-Object { $_.Name -match 'Wi-?Fi|Wireless|WLAN|802\.11' } | Select-Object -First 1
$btOk = [bool]($pnp | Where-Object { $_.PNPClass -eq 'Bluetooth' -and $_.Name -match 'Bluetooth' })
$camOk = [bool]($pnp | Where-Object { $_.PNPClass -in 'Camera', 'Image' })
$kbOk = [bool](Get-CimInstance Win32_Keyboard | Select-Object -First 1)
$audioOk = [bool](Get-CimInstance Win32_SoundDevice | Select-Object -First 1)
$usbOk = [bool](Get-CimInstance Win32_USBController | Select-Object -First 1)
$wifiOk = [bool]$wifi
$qcHTML = (Badge $kbOk 'Keyboard') + (Badge ($wifiOk -and $btOk) 'Wi-Fi &amp; BT') + (Badge $camOk 'Webcam') + (Badge ($audioOk -and $usbOk) 'Audio &amp; Port')

# ------------------------------------------------------------------------------
# 9. Link garansi resmi
# ------------------------------------------------------------------------------
$warrantyURL = "https://www.google.com/search?q=cek+garansi+$([uri]::EscapeDataString($brand))+$serialNumber"
if ($brand -like '*ASUS*') { $warrantyURL = 'https://www.asus.com/support/warranty-status-inquiry/' }
elseif ($brand -like '*LENOVO*') { $warrantyURL = "https://pcsupport.lenovo.com/id/id/warranty-lookup?serialNumber=$serialNumber" }
elseif ($brand -match '^HP$|Hewlett') { $warrantyURL = 'https://support.hp.com/id-en/check-warranty' }
elseif ($brand -like '*DELL*') { $warrantyURL = "https://www.dell.com/support/home/id-id/product-support/servicetag/$serialNumber/overview" }
elseif ($brand -like '*ACER*') { $warrantyURL = 'https://www.acer.com/id-id/support/warranty' }
elseif ($brand -like '*MSI*' -or $brand -like '*Micro-Star*') { $warrantyURL = 'https://id.msi.com/page/warranty' }
$serialJs = $serialNumber -replace '[^A-Za-z0-9\-_]', ''

# ------------------------------------------------------------------------------
# 10. Catatan kondisi & rekomendasi (otomatis dari data di atas)
# ------------------------------------------------------------------------------
$notes = New-Object System.Collections.ArrayList
if ($cpuNote) { [void]$notes.Add("<strong>Prosesor:</strong> $cpuNote") }
if ($null -ne $healthPct) {
    if ($healthPct -ge 80) { [void]$notes.Add("<strong>Baterai:</strong> Kondisi baik (${healthPct}% dari kapasitas baru).") }
    elseif ($healthPct -ge 60) { [void]$notes.Add("<strong>Baterai:</strong> Sudah menurun (${healthPct}%), daya tahan lebih pendek dari baru namun masih layak pakai harian.") }
    else { [void]$notes.Add("<strong>Baterai:</strong> Aus berat (${healthPct}%), sebaiknya diganti.") }
}
if ($ramGB -le 4) { [void]$notes.Add("<strong>RAM:</strong> $ramGB GB sangat terbatas, multitasking akan terasa berat.") }
elseif ($ramGB -le 8) { [void]$notes.Add("<strong>RAM:</strong> $ramGB GB cukup untuk office/kuliah; untuk multitasking berat disarankan 16 GB.") }
else { [void]$notes.Add("<strong>RAM:</strong> $ramGB GB lega untuk multitasking dan pekerjaan kreatif.") }
if ($soldered) { [void]$notes.Add('<strong>Upgrade RAM:</strong> RAM on-board (solder), tidak bisa ditambah. Pilih kapasitas sejak awal.') }
elseif ($slotsFree -gt 0) { [void]$notes.Add("<strong>Upgrade RAM:</strong> Ada $slotsFree slot kosong. Menambah modul identik akan mengaktifkan Dual Channel (grafis terintegrasi lebih kencang).") }
if ($isHDD) { [void]$notes.Add('<strong>Penyimpanan:</strong> Masih HDD (lambat). Upgrade ke SSD akan terasa sangat signifikan.') }
elseif ($diskCap -in '128 GB', '64 GB' -or $diskCap -match '^\d{2} GB$') { [void]$notes.Add("<strong>Penyimpanan:</strong> Kapasitas $diskCap tergolong sempit untuk Windows modern.") }
if ($null -ne $wearVal) {
    if ($wearVal -ge 50) { [void]$notes.Add("<strong>SSD:</strong> Sisa umur $(100 - $wearVal)%, sudah banyak terpakai. Pertimbangkan penggantian.") }
    elseif ($wearVal -ge 20) { [void]$notes.Add("<strong>SSD:</strong> Sisa umur $(100 - $wearVal)%, sudah cukup terpakai.") }
    else { [void]$notes.Add("<strong>SSD:</strong> Sisa umur $(100 - $wearVal)%, kondisi sangat baik.") }
}
if ($poh -gt 0) { [void]$notes.Add("<strong>Riwayat disk:</strong> Total menyala $poh jam, indikator pemakaian sebelumnya.") }
if (-not $hasDedicated) { [void]$notes.Add('<strong>Grafis:</strong> Hanya grafis terintegrasi. Lancar untuk office dan streaming, tidak ditujukan untuk gaming berat/render 3D.') }
if ($resX -gt 0 -and $resX -le 1366) { [void]$notes.Add('<strong>Layar:</strong> Resolusi HD rendah, kurang nyaman untuk multitasking jendela.') }
if (-not $winOriginal) { [void]$notes.Add("<strong>Windows:</strong> $winStatus. Periksa legalitas lisensi sebelum dijual.") }
if ($null -ne $umur) { [void]$notes.Add("<strong>Usia perangkat:</strong> +/- $umur tahun berdasarkan tanggal BIOS/model.") }
$notesHTML = ($notes | ForEach-Object { "<p>$_</p>" }) -join "`n            "

# Detail hardware lengkap (ditaruh di kotak analisis agar tabel spesifikasi tetap seperti V8.9)
$det = New-Object System.Collections.ArrayList
[void]$det.Add("<p><strong>Sistem Operasi:</strong> $(Enc $osName), $osExtra</p>")
if ($biosVer) { [void]$det.Add("<p><strong>BIOS:</strong> versi $(Enc $biosVer), tanggal $tglProduksi</p>") }
if ($cpuExtra.Count -gt 0) { [void]$det.Add("<p><strong>Detail Prosesor:</strong> " + ($cpuExtra -join ' | ') + "</p>") }
[void]$det.Add("<p><strong>NPU / AI Engine:</strong> $(Enc $npuStatus) | $(Enc $npuPerf) | $(Enc $npuMemoryText)</p>")
if ($gpuDrivers.Count -gt 0) { [void]$det.Add("<p><strong>Driver Grafis:</strong> " + (Enc ($gpuDrivers -join '; ')) + "</p>") }
$ramX = @()
if ($channelText) { $ramX += $channelText }
if ($ramMaxText) { $ramX += $ramMaxText }
if ($ramMaker -and $ramMaker -notmatch 'Unknown|Undefined|^0+$') { $ramX += "Merek $(Enc $ramMaker)" }
if ($ramPart -and $ramPart -notmatch '^0+$|Unknown') { $ramX += "Part No. $(Enc $ramPart)" }
if ($ramX.Count -gt 0) { [void]$det.Add("<p><strong>Detail RAM:</strong> " + ($ramX -join ' | ') + "</p>") }
[void]$det.Add("<p><strong>Detail Penyimpanan:</strong> " + ($diskSmall -join ' | ') + "</p>")
$connX = @()
if ($wifiOk) { $connX += "Wi-Fi: $(Enc $wifi.Name)" } else { $connX += 'Wi-Fi: tidak terdeteksi' }
if ($btOk) { $connX += 'Bluetooth: terdeteksi' } else { $connX += 'Bluetooth: tidak terdeteksi' }
[void]$det.Add("<p><strong>Konektivitas:</strong> " + ($connX -join ' | ') + ". Badge QC = deteksi perangkat oleh Windows, bukan uji fungsi fisik.</p>")
$detHTML = ($det | ForEach-Object { $_ }) -join "`n            "

# NPU: bedakan indikasi dari nama CPU dengan perangkat NPU yang benar-benar dilaporkan driver.
# NPU tidak memiliki "VRAM" tersendiri seperti GPU; biasanya memakai RAM sistem bersama.
$npuDevices = @(Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.Name -and $_.Name -match '\bNPU\b|Neural Processing|AI Boost|Hexagon|AMD IPU|Ryzen AI|Intel AI'
})
$npuDeviceNames = @($npuDevices | ForEach-Object { ([string]$_.Name).Trim() } | Select-Object -Unique)
$npuDetectedByName = [bool]($cpuName -match 'Core\s*\(?TM\)?\s*Ultra|Ryzen\s+AI|Snapdragon\s+X|X1E-|X1P-')
$npuStatus = 'Tidak terkonfirmasi dari nama CPU/perangkat Windows'
if ($npuDetectedByName) { $npuStatus = 'NPU diperkirakan tersedia berdasarkan model CPU; verifikasi spesifikasi resmi model tepat' }
if ($npuDeviceNames.Count -gt 0) { $npuStatus = 'Perangkat NPU/AI dilaporkan driver Windows: ' + (Enc ($npuDeviceNames -join '; ')) }
$npuPerf = 'TOPS NPU tidak dilaporkan secara standar oleh WMI; perlu lookup spesifikasi resmi berdasarkan SKU CPU'
if ($cpuName -match 'Core\s*\(?TM\)?\s*Ultra\s*([3579])\s+(\d{3,4}\w*)') {
    $npuPerf = 'Intel Core Ultra terdeteksi; TOPS NPU berbeda menurut SKU/generasi dan tidak disimpulkan dari nama umum saja'
} elseif ($cpuName -match 'Ryzen\s+AI') {
    $npuPerf = 'AMD Ryzen AI terdeteksi; TOPS NPU berbeda menurut SKU dan TDP/konfigurasi, perlu SKU tepat'
} elseif ($cpuName -match 'Snapdragon\s+X') {
    $npuPerf = 'Qualcomm Snapdragon X terdeteksi; spesifikasi NPU vendor tersedia per platform, bukan dari WMI'
}
$npuMemoryText = 'VRAM NPU khusus: umumnya tidak ada/ tidak diekspos sebagai memori terpisah; NPU memakai RAM sistem bersama. Angka memori bersama yang ditampilkan Windows bukan kapasitas VRAM NPU.'
$npuLine = "<p><strong>NPU / AI:</strong> $(Enc $npuStatus)</p><p><strong>Performa NPU:</strong> $(Enc $npuPerf)</p><p><strong>Memori NPU:</strong> $(Enc $npuMemoryText)</p>"
$hybridLine = ''
if ($hybrid) { $hybridLine = "<p><strong>Arsitektur Core:</strong> Hybrid (P-Core + E-Core)</p>" }

# ------------------------------------------------------------------------------
# 11. Render HTML laporan (tampilan sama seperti V8.9)
# ------------------------------------------------------------------------------
$brandH = Enc $brand; $modelH = Enc $model; $snH = Enc $serialNumber; $osH = Enc $osName

# Barcode dari S/N (dilewati jika S/N tidak tersedia)
$barcodeHTML = ''
if ($serialNumber -ne 'N/A' -and $serialJs) {
    $svg = New-Code128Svg $serialJs
    if ($svg) { $barcodeHTML = "<div class='barcode-box'>$svg<div class='barcode-text'>$(Enc $serialJs)</div></div>" }
}
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Laporan Inspeksi Laptop - $brandH $modelH</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background: #f4f6f9; margin: 30px; color: #1c1e21; }
        .card { background: #fff; padding: 35px; border-radius: 14px; max-width: 800px; margin: auto; box-shadow: 0 8px 24px rgba(0,0,0,0.08); border: 1px solid #e1e4e8; }
        .header { display: flex; justify-content: space-between; align-items: center; border-bottom: 2px solid #005a9e; padding-bottom: 12px; margin-bottom: 20px; }
        h2 { color: #005a9e; margin: 0; font-size: 22px; }
        h3 { color: #005a9e; margin-top: 30px; margin-bottom: 10px; font-size: 16px; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
        table { width: 100%; border-collapse: collapse; margin-top: 10px; }
        th, td { text-align: left; padding: 10px 14px; border-bottom: 1px solid #edf0f2; font-size: 14px; }
        th { background-color: #f8f9fa; color: #555; width: 32%; font-weight: 600; }
        .analysis-box { background: #eef6fc; border-left: 4px solid #005a9e; padding: 15px 20px; border-radius: 6px; margin-top: 15px; }
        .analysis-box p { margin: 6px 0; font-size: 13.5px; }
        .highlight { font-weight: bold; color: #005a9e; }
        .badge-ok { background: #e8f5e9; color: #2e7d32; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; }
        .badge-warn { background: #fff3e0; color: #e65100; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; }
        .badge-bad { background: #ffebee; color: #c62828; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold; }
        .sn-box { background: #eef6fc; color: #005a9e; padding: 4px 10px; border-radius: 6px; font-weight: bold; font-size: 13px; }
        .barcode-box { text-align: center; margin: 0 0 20px 0; padding: 10px 0 6px 0; }
        .barcode-box svg { max-width: 100%; height: 70px; background: #fff; }
        .barcode-text { font-family: Consolas, 'Courier New', monospace; font-size: 13px; letter-spacing: 3px; margin-top: 4px; color: #1c1e21; }
        .btn-warranty { display: inline-block; background: #005a9e; color: #ffffff; padding: 5px 12px; border-radius: 6px; text-decoration: none; font-size: 12px; font-weight: bold; margin-left: 8px; border: none; cursor: pointer; }
        .btn-warranty:hover { background: #004578; }
        .footer { font-size: 11px; color: #888; text-align: center; margin-top: 25px; padding-top: 15px; border-top: 1px solid #edf0f2; }
    </style>
    <script>
        function checkWarranty(sn, url) {
            navigator.clipboard.writeText(sn);
            alert("Serial Number (" + sn + ") berhasil disalin ke Clipboard!\n\nSilakan 'Paste' (Ctrl+V) di kolom pencarian portal garansi yang terbuka.");
            window.open(url, '_blank');
        }
    </script>
</head>
<body>
    <div class='card'>
        <div class='header'>
            <h2>LAPORAN INSPEKSI LAPTOP</h2>
            <span class='sn-box'>S/N: $snH</span>
        </div>
        $barcodeHTML
        <table>
            <tr><th>Merek (Brand)</th><td class='highlight'>$brandH</td></tr>
            <tr><th>Model / Tipe Laptop</th><td class='highlight'>$modelH</td></tr>
            <tr>
                <th>Serial Number (S/N)</th>
                <td>
                    <strong>$snH</strong>
                    <button onclick="checkWarranty('$serialJs', '$warrantyURL')" class='btn-warranty'>&#128269; Cek Garansi Resmi</button>
                </td>
            </tr>
            <tr><th>Sistem Operasi (OS)</th><td>$osH <span class='$winClass'>$winStatus</span></td></tr>
            <tr><th>Tanggal Rilis BIOS / Model</th><td>$tglProduksi ($tahun)</td></tr>
            <tr><th>Prosesor (CPU)</th><td>$cpuHTML</td></tr>
            <tr><th>Kartu Grafis (GPU)</th><td>$gpu</td></tr>
            <tr><th>Memori (RAM)</th><td>$ramDetail</td></tr>
            <tr><th>Penyimpanan (SSD)</th><td>$ssdHTML</td></tr>
            <tr><th>Kesehatan Baterai</th><td>$batDisplay</td></tr>
            <tr><th>Layar &amp; Resolusi</th><td>$displayHTML</td></tr>
            <tr>
                <th>Pengujian Fungsi (QC)</th>
                <td>
                    $qcHTML
                </td>
            </tr>
        </table>

        <h3>SUMMARY REPORT: KESETARAAN PROSESOR &amp; HARDWARE</h3>
        <div class='analysis-box'>
            <p><strong>Prosesor Terdeteksi:</strong> $(Enc $cpuName)</p>
            <p><strong>Arsitektur / Generasi:</strong> $(Enc $codename) (rilis +/- $cpuYear)</p>
            <p><strong>Klasifikasi Performa:</strong> $cpuCategory</p>
            <p><strong>Kesetaraan Intel (Lama s.d Terbaru):</strong> Setara kelas $equivIntel</p>
            <p><strong>Kesetaraan AMD (Lama s.d Terbaru):</strong> Setara kelas $equivAMD</p>
            <p><strong>Cocok Untuk:</strong> $cpuUse</p>
            $hybridLine
            $npuLine
        </div>

        <h3>DETAIL HARDWARE LENGKAP</h3>
        <div class='analysis-box'>
            $detHTML
        </div>

        <h3>CATATAN KONDISI &amp; REKOMENDASI</h3>
        <div class='analysis-box'>
            $notesHTML
        </div>

        <div class='footer'>Di-generate otomatis oleh Sistem Inspeksi Toko | Kesetaraan CPU bersifat perkiraan (indikatif)</div>
    </div>
</body>
</html>
"@

# ------------------------------------------------------------------------------
# 12. Simpan file di Desktop
# ------------------------------------------------------------------------------
$cleanSN = $serialNumber -replace '[\\/:*?"<>|]', ''
if (-not $cleanSN) { $cleanSN = 'Device' }
$fileName = "Laporan_Inspeksi_$cleanSN.html"
$savePath = Join-Path -Path $desktopPath -ChildPath $fileName

[System.IO.File]::WriteAllText($savePath, $htmlContent, [System.Text.Encoding]::UTF8)
$sw.Stop()

Write-Host ""
Write-Host "Laporan V9.0 selesai dalam $([math]::Round($sw.Elapsed.TotalSeconds, 1)) detik." -ForegroundColor Green
Write-Host "File tersimpan di: $savePath" -ForegroundColor Cyan
Start-Process $savePath