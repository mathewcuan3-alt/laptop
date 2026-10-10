#requires -Version 5.1
# QC LAPTOP GUIDED V13.0 - GUI wizard for Windows PowerShell 5.1
# Uses WinForms. Hardware detection is advisory; guided physical tests require technician confirmation.
$ErrorActionPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Results = New-Object System.Collections.Generic.List[object]
$script:Started = Get-Date
$script:Desktop = [Environment]::GetFolderPath('Desktop')
if (-not (Test-Path $script:Desktop)) { $script:Desktop = $env:TEMP }

function Add-QCResult([string]$Area,[string]$Status,[string]$Finding,[string]$NextStep='') {
    $script:Results.Add([pscustomobject]@{
        Waktu=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        Area=$Area; Status=$Status; Temuan=$Finding; TindakLanjut=$NextStep
    }) | Out-Null
}
function Show-Info([string]$title,[string]$text) {
    [System.Windows.Forms.MessageBox]::Show($text,$title,[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}
function Ask-QC([string]$area,[string]$instructions) {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "QC Laptop - $area"
    $form.Size = New-Object System.Drawing.Size(620,300)
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $instructions
    $lbl.Location = New-Object System.Drawing.Point(22,22)
    $lbl.Size = New-Object System.Drawing.Size(560,125)
    $lbl.Font = New-Object System.Drawing.Font('Segoe UI',11)
    $form.Controls.Add($lbl)
    $detail = New-Object System.Windows.Forms.TextBox
    $detail.Location = New-Object System.Drawing.Point(22,155)
    $detail.Size = New-Object System.Drawing.Size(560,28)
    $detail.PlaceholderText = 'Catatan/gejala (opsional)'
    $form.Controls.Add($detail)
    $result = $null
    $buttons = @(
        @{Text='LULUS';X=22;Status='LULUS'},
        @{Text='GAGAL';X=205;Status='GAGAL'},
        @{Text='BELUM DIUJI';X=388;Status='BELUM DIUJI'}
    )
    foreach ($b in $buttons) {
        $btn = New-Object System.Windows.Forms.Button
        $btn.Text = $b.Text
        $btn.Tag = $b.Status
        $btn.Location = New-Object System.Drawing.Point($b.X,205)
        $btn.Size = New-Object System.Drawing.Size(165,38)
        $btn.Font = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold)
        $btn.Add_Click({
            $script:tempResult = [pscustomobject]@{Status=$this.Tag;Detail=$detail.Text}
            $form.Close()
        }.GetNewClosure())
        $form.Controls.Add($btn)
    }
    $form.Add_Shown({$form.Activate()})
    $form.ShowDialog() | Out-Null
    if ($script:tempResult) {
        $result = $script:tempResult
        $script:tempResult = $null
    } else { $result = [pscustomobject]@{Status='BELUM DIUJI';Detail='Jendela ditutup sebelum memilih hasil.'} }
    $next = if ($result.Status -eq 'GAGAL') { 'Ulangi dengan perangkat/kabel pembanding yang diketahui normal; catat gejala dan lakukan pemeriksaan lanjutan.' } elseif ($result.Status -eq 'LULUS') { 'Lulus untuk tes yang dilakukan saja; bukan jaminan seluruh bagian sehat.' } else { 'Lakukan pemeriksaan ketika alat/perangkat uji tersedia.' }
    Add-QCResult $area $result.Status $result.Detail $next
}
function Show-ColorTest {
    $colors = @(
        @{Name='HITAM';Color=[System.Drawing.Color]::Black},
        @{Name='PUTIH';Color=[System.Drawing.Color]::White},
        @{Name='MERAH';Color=[System.Drawing.Color]::Red},
        @{Name='HIJAU';Color=[System.Drawing.Color]::Lime},
        @{Name='BIRU';Color=[System.Drawing.Color]::Blue},
        @{Name='ABU-ABU';Color=[System.Drawing.Color]::Gray}
    )
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'QC LCD - Tes Warna'
    $form.WindowState = 'Maximized'
    $form.FormBorderStyle = 'None'
    $form.TopMost = $true
    $form.KeyPreview = $true
    $index = 0
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Dock = 'Fill'
    $form.Controls.Add($panel)
    $label = New-Object System.Windows.Forms.Label
    $label.Dock = 'Top'; $label.Height = 55
    $label.TextAlign = 'MiddleCenter'
    $label.Font = New-Object System.Drawing.Font('Segoe UI',18,[System.Drawing.FontStyle]::Bold)
    $panel.Controls.Add($label)
    $hint = New-Object System.Windows.Forms.Label
    $hint.Dock='Bottom'; $hint.Height=70
    $hint.Text='Periksa dead pixel, garis, bercak, flicker dan warna tidak merata. Klik LANJUT atau tekan Enter. Esc untuk keluar.'
    $hint.TextAlign='MiddleCenter'
    $hint.Font=New-Object System.Drawing.Font('Segoe UI',12)
    $panel.Controls.Add($hint)
    $next = New-Object System.Windows.Forms.Button
    $next.Text='LANJUT  →'; $next.Size=New-Object System.Drawing.Size(150,42)
    $next.Anchor='Bottom'
    $next.Location=New-Object System.Drawing.Point(([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Width-160),([System.Windows.Forms.Screen]::PrimaryScreen.Bounds.Height-120))
    $panel.Controls.Add($next)
    $setColor = {
        $panel.BackColor = $colors[$index].Color
        $label.Text = "TES LCD: $($colors[$index].Name)  ($($index+1)/$($colors.Count))"
        $label.ForeColor = if ($colors[$index].Name -in @('PUTIH','ABU-ABU','HIJAU')) {[System.Drawing.Color]::Black} else {[System.Drawing.Color]::White}
        $hint.ForeColor = $label.ForeColor
        $next.ForeColor = [System.Drawing.Color]::Black
        $form.Refresh()
    }
    $next.Add_Click({
        $index++
        if ($index -ge $colors.Count) { $form.Close() } else { & $setColor }
    }.GetNewClosure())
    $form.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) { $form.Close() }
        elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter -or $_.KeyCode -eq [System.Windows.Forms.Keys]::Space) { $next.PerformClick() }
    }.GetNewClosure())
    & $setColor
    $form.ShowDialog() | Out-Null
    Ask-QC 'LCD / Panel' 'Setelah melihat semua warna, apakah layar bebas dari dead pixel, garis, bercak, flicker, dan warna tidak normal? Catatan: pemeriksaan visual oleh teknisi.'
}
function Show-KeyboardTest {
    $form=New-Object System.Windows.Forms.Form
    $form.Text='QC Keyboard - Tekan Semua Tombol'
    $form.Size=New-Object System.Drawing.Size(780,430)
    $form.StartPosition='CenterScreen'; $form.TopMost=$true; $form.KeyPreview=$true
    $lbl=New-Object System.Windows.Forms.Label
    $lbl.Text="Klik kotak di bawah, lalu tekan tombol keyboard satu per satu.`r`nTombol yang terbaca akan tampil. Tes juga Shift, Ctrl, Alt, Fn, panah, Enter, Backspace, tombol angka dan fungsi."
    $lbl.Location=New-Object System.Drawing.Point(20,20); $lbl.Size=New-Object System.Drawing.Size(720,70)
    $lbl.Font=New-Object System.Drawing.Font('Segoe UI',11); $form.Controls.Add($lbl)
    $display=New-Object System.Windows.Forms.TextBox
    $display.Location=New-Object System.Drawing.Point(20,110); $display.Size=New-Object System.Drawing.Size(720,160)
    $display.Multiline=$true; $display.ReadOnly=$true; $display.Font=New-Object System.Drawing.Font('Consolas',12)
    $display.ScrollBars='Vertical'; $display.Text='Belum ada tombol yang dicatat.'; $form.Controls.Add($display)
    $capture=New-Object System.Windows.Forms.TextBox
    $capture.Location=New-Object System.Drawing.Point(20,285); $capture.Size=New-Object System.Drawing.Size(720,32)
    $capture.Text='Klik di sini, lalu tekan tombol untuk mulai tes'; $form.Controls.Add($capture)
    $seen=New-Object System.Collections.Generic.HashSet[string]
    $capture.Add_KeyDown({
        $key=$_.KeyCode.ToString()
        [void]$seen.Add($key)
        $display.Text=(($seen | Sort-Object) -join ', ')
        $_.SuppressKeyPress=$true
    }.GetNewClosure())
    $capture.Add_KeyPress({ $_.Handled=$true })
    $btn=New-Object System.Windows.Forms.Button
    $btn.Text='SELESAI TES KEYBOARD'; $btn.Location=New-Object System.Drawing.Point(20,335); $btn.Size=New-Object System.Drawing.Size(220,38)
    $btn.Add_Click({$form.Close()}); $form.Controls.Add($btn)
    $form.ShowDialog() | Out-Null
    $keys=($seen | Sort-Object) -join ', '
    Ask-QC 'Keyboard' ("Tombol yang terdeteksi oleh jendela: $keys`r`nApakah semua tombol fisik yang diuji berfungsi? Catatan: Fn dan tombol khusus tertentu tidak selalu dapat dideteksi oleh Windows.")
}
function Start-Loading {
    $f=New-Object System.Windows.Forms.Form
    $f.Text='QC Laptop Checker'; $f.Size=New-Object System.Drawing.Size(520,220); $f.StartPosition='CenterScreen'
    $f.FormBorderStyle='FixedDialog'; $f.ControlBox=$false; $f.TopMost=$true
    $title=New-Object System.Windows.Forms.Label
    $title.Text='QC LAPTOP CHECKER'; $title.Font=New-Object System.Drawing.Font('Segoe UI',20,[System.Drawing.FontStyle]::Bold)
    $title.Location=New-Object System.Drawing.Point(25,25); $title.Size=New-Object System.Drawing.Size(450,45); $title.TextAlign='MiddleCenter'; $f.Controls.Add($title)
    $msg=New-Object System.Windows.Forms.Label
    $msg.Text='Menyiapkan pemeriksaan hardware...'; $msg.Location=New-Object System.Drawing.Point(25,80); $msg.Size=New-Object System.Drawing.Size(450,28); $msg.TextAlign='MiddleCenter'; $f.Controls.Add($msg)
    $bar=New-Object System.Windows.Forms.ProgressBar
    $bar.Location=New-Object System.Drawing.Point(30,125); $bar.Size=New-Object System.Drawing.Size(440,25); $bar.Style='Marquee'; $f.Controls.Add($bar)
    $f.Show(); $f.Refresh()
    return @{Form=$f;Message=$msg}
}
# Splash screen
$loading=Start-Loading
try {
    $loading.Message.Text='Membaca identitas dan spesifikasi laptop...'; $loading.Form.Refresh()
    $sys=Get-CimInstance Win32_ComputerSystem
    $bios=Get-CimInstance Win32_BIOS
    $os=Get-CimInstance Win32_OperatingSystem
    $cpu=Get-CimInstance Win32_Processor | Select-Object -First 1
    $ram=@(Get-CimInstance Win32_PhysicalMemory)
    $disks=@(Get-CimInstance Win32_DiskDrive)
    $pnp=@(Get-CimInstance Win32_PnPEntity)
    $bad=@($pnp | Where-Object {$null -ne $_.ConfigManagerErrorCode -and $_.ConfigManagerErrorCode -ne 0})
    Add-QCResult 'Identitas' 'INFO' "$($sys.Manufacturer) $($sys.Model); Serial BIOS=$($bios.SerialNumber); OS=$($os.Caption) build $($os.BuildNumber)" 'Cocokkan serial dengan label fisik.'
    Add-QCResult 'CPU' 'TERDETEKSI' "$($cpu.Name); core=$($cpu.NumberOfCores); thread=$($cpu.NumberOfLogicalProcessors)" 'Deteksi nama tidak menguji kestabilan CPU.'
    $ramGB=[math]::Round((($ram | Measure-Object Capacity -Sum).Sum/1GB),1)
    Add-QCResult 'RAM' 'TERDETEKSI' "$ramGB GB; modul=$($ram.Count)" 'Untuk tes mendalam gunakan Windows Memory Diagnostic/MemTest86.'
    foreach($d in $disks) {
        Add-QCResult "Disk $($d.Model)" $(if($d.Status -eq 'OK'){'TERDETEKSI'}else{'PERINGATAN'}) "Status Windows=$($d.Status); kapasitas=$([math]::Round($d.Size/1GB,1)) GB; serial=$($d.SerialNumber)" 'Cek SMART vendor bila ada gejala.'
    }
    Add-QCResult 'Device Manager' $(if($bad.Count){'PERINGATAN'}else{'TIDAK ADA ERROR YANG DILAPORKAN'}) "Perangkat PnP diperiksa=$($pnp.Count); error code non-zero=$($bad.Count)" 'Buka Device Manager untuk detail kode; tidak adanya error bukan bukti hardware sehat.'
    foreach($dev in $bad) {
        Add-QCResult "PnP: $($dev.Name)" 'PERINGATAN' "ErrorCode=$($dev.ConfigManagerErrorCode); ID=$($dev.PNPDeviceID)" 'Periksa driver, koneksi perangkat, dan kemungkinan hardware.'
    }
    $groups=@(
        @{Name='Wi-Fi';Pattern='Wi-?Fi|Wireless|802\.11|WLAN';Next='Tes koneksi dengan access point yang diketahui normal.'},
        @{Name='Bluetooth';Pattern='Bluetooth';Next='Pair dengan perangkat Bluetooth pembanding.'},
        @{Name='Webcam';Pattern='Camera|Webcam|Integrated Camera|USB Video';Next='Buka Camera dan pastikan gambar hidup.'},
        @{Name='Audio';Pattern='Audio|Realtek.*(Sound|Audio)|High Definition Audio';Next='Tes speaker dan mikrofon.'},
        @{Name='USB controller';Pattern='USB.*(Host Controller|Root Hub)|xHCI';Next='Controller terdeteksi tidak menjamin setiap port fisik berfungsi.'},
        @{Name='LAN Ethernet';Pattern='Ethernet|Gigabit|PCIe.*GbE|LAN';Next='Tes kabel LAN dan jaringan pembanding.'},
        @{Name='Display adapter';Pattern='NVIDIA|AMD Radeon|Intel.*Graphics|Display adapter';Next='Tes visual LCD dan monitor eksternal jika tersedia.'},
        @{Name='Touchpad';Pattern='Touchpad|I2C HID|ELAN.*Input|Synaptics';Next='Tes pointer, klik, scroll dan gesture.'}
    )
    foreach($g in $groups) {
        $m=@($pnp | Where-Object {$_.Name -match $g.Pattern})
        Add-QCResult $g.Name $(if($m.Count){'TERDETEKSI'}else{'TIDAK TERDETEKSI / VERIFIKASI'}) $(if($m.Count){($m.Name -join '; ')}else{'Tidak ditemukan nama perangkat yang cocok; bisa karena nama/driver/model berbeda.'}) $g.Next
    }
    $loading.Form.Close()
} catch {
    $loading.Form.Close()
    Show-Info 'Peringatan scan' "Sebagian scan otomatis tidak selesai: $($_.Exception.Message)`r`nWizard akan tetap dilanjutkan."
}

Show-Info 'Selamat datang' "QC Laptop Checker V13.0`r`n`r`nProgram akan membuka tes satu per satu. Siapkan flashdisk, headset, jaringan Wi-Fi, dan monitor eksternal bila ingin menguji port terkait.`r`n`r`nHasil fisik dicatat berdasarkan konfirmasi teknisi."
Show-KeyboardTest
Show-ColorTest
# Camera app: external app access is needed; no assumption of success just because device is detected.
$cam = [System.Windows.Forms.MessageBox]::Show("Program akan mencoba membuka aplikasi Windows Camera.`r`nPeriksa apakah gambar muncul, jelas, dan stabil. Tutup aplikasi Camera setelah tes.`r`n`r`nLanjut membuka Camera?","Tes Webcam",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Question)
if($cam -eq [System.Windows.Forms.DialogResult]::Yes) {
    try { Start-Process 'microsoft.windows.camera:'; Start-Sleep -Seconds 2 } catch {}
    Ask-QC 'Webcam' 'Apakah gambar kamera muncul dan stabil? Pastikan izin kamera tidak memblokir aplikasi.'
} else { Add-QCResult 'Webcam' 'BELUM DIUJI' 'Teknisi melewati tes kamera.' 'Buka aplikasi Camera dan verifikasi gambar.' }
Ask-QC 'Wi-Fi / Internet' 'Sambungkan ke Wi-Fi yang diketahui normal. Buka halaman web atau lakukan tes jaringan. Apakah koneksi berhasil dan stabil?'
Ask-QC 'Bluetooth' 'Jika tersedia, aktifkan Bluetooth dan coba pairing dengan perangkat pembanding. Apakah pairing dan koneksi berhasil?'
Ask-QC 'Speaker dan mikrofon' 'Putar suara, tes kanal kiri/kanan jika tersedia, lalu rekam suara dan putar kembali. Apakah semua yang diuji berfungsi?'
Ask-QC 'USB - setiap port' 'Uji SETIAP port USB dengan flashdisk yang diketahui normal. Bila memungkinkan, salin file uji. Apakah semua port yang diuji berfungsi dan tidak putus?'
Ask-QC 'HDMI / USB-C DisplayPort' 'Jika port tersedia, sambungkan monitor dan kabel yang diketahui normal. Apakah tampilan keluar stabil? Jika tidak tersedia, pilih BELUM DIUJI.'
Ask-QC 'LAN Ethernet' 'Jika port tersedia, uji dengan kabel dan router/switch yang diketahui normal. Apakah link dan koneksi jaringan berfungsi?'
Ask-QC 'Audio jack' 'Jika tersedia, uji headset. Periksa suara putus-putus saat konektor bergerak perlahan tanpa dipaksa. Apakah berfungsi?'
Ask-QC 'Touchpad dan klik' 'Uji gerakan pointer, klik kiri/kanan, scroll dan gesture jika didukung. Apakah semua fungsi yang diuji berjalan?'
Ask-QC 'Baterai dan charger' 'Periksa indikator pengisian dan persentase baterai. Bila aman, cabut charger sebentar. Apakah pengisian dan perpindahan daya normal?'
Ask-QC 'Kipas dan suhu' 'Dengarkan kipas dan amati suhu saat pemakaian normal. Jangan menutup ventilasi atau memaksa tes beban bila laptop sudah panas. Apakah ada suara abnormal atau panas berlebihan?'
Ask-QC 'Port lain / pembaca kartu' 'Uji card reader, fingerprint, SIM, Thunderbolt atau port lain jika memang tersedia. Catat komponen yang tidak diuji.'
Ask-QC 'Indikasi konektor / motherboard' 'Apakah ditemukan gejala seperti perangkat internal hilang, koneksi terputus-putus, port mati atau restart abnormal? Ini hanya indikasi; konfirmasi fisik perlu teknisi.'

# Save report
$finish=Get-Date
$safeSerial=[string]$bios.SerialNumber -replace '[\\/:*?"<>|]',''
if([string]::IsNullOrWhiteSpace($safeSerial) -or $safeSerial -match 'Default string|System Serial|To be filled|^0+$'){$safeSerial='Laptop'}
$stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$base=Join-Path $script:Desktop ("QC_Laptop_{0}_{1}" -f $safeSerial,$stamp)
$jsonPath="$base.json"; $htmlPath="$base.html"
$script:Results | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $jsonPath -Encoding UTF8
function Enc([object]$x){[System.Net.WebUtility]::HtmlEncode([string]$x)}
$rows=foreach($r in $script:Results) {
    $cls=if($r.Status -eq 'LULUS' -or $r.Status -eq 'TERDETEKSI' -or $r.Status -eq 'INFO'){'ok'}elseif($r.Status -eq 'GAGAL'){'bad'}else{'warn'}
    "<tr><td>$(Enc $r.Waktu)</td><td>$(Enc $r.Area)</td><td><b class='$cls'>$(Enc $r.Status)</b></td><td>$(Enc $r.Temuan)</td><td>$(Enc $r.TindakLanjut)</td></tr>"
}
$failed=@($script:Results | Where-Object {$_.Status -eq 'GAGAL'}).Count
$pending=@($script:Results | Where-Object {$_.Status -eq 'BELUM DIUJI' -or $_.Status -match 'PERINGATAN|VERIFIKASI'}).Count
$html=@"
<!doctype html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>QC Laptop Checker V13</title>
<style>body{font-family:Segoe UI,Arial;margin:24px;background:#f3f5f7;color:#202733}.wrap{max-width:1400px;margin:auto;background:white;padding:24px;border-radius:12px}h1{color:#174f78}table{width:100%;border-collapse:collapse;margin-top:18px}td,th{border:1px solid #dce2e8;padding:8px;text-align:left;vertical-align:top;font-size:12px;overflow-wrap:anywhere}th{background:#eaf1f7}.ok{color:#16733a}.bad{color:#b42318}.warn{color:#8a5a00}.note{padding:12px;background:#fff4d6;border-left:4px solid #e6a700;margin:16px 0}.stats{display:flex;gap:12px;flex-wrap:wrap}.stats div{background:#f0f4f8;padding:12px;border-radius:8px}button{padding:10px 16px}@media print{body{margin:0;background:white}.wrap{padding:0}button{display:none}td,th{font-size:9px;padding:5px}}</style></head><body><div class="wrap">
<h1>QC LAPTOP CHECKER V13.0</h1><p>Wizard pemeriksaan komponen dengan tes visual/interaktif dan scan Windows.</p>
<p><b>Perangkat:</b> $(Enc "$($sys.Manufacturer) $($sys.Model)") &nbsp; <b>Serial:</b> $(Enc $bios.SerialNumber)<br><b>Windows:</b> $(Enc $os.Caption) &nbsp; <b>Mulai:</b> $(Enc $script:Started.ToString('yyyy-MM-dd HH:mm:ss')) &nbsp; <b>Selesai:</b> $(Enc $finish.ToString('yyyy-MM-dd HH:mm:ss'))</p>
<div class="stats"><div><b>$($script:Results.Count)</b><br>Temuan</div><div><b>$failed</b><br>Tes gagal</div><div><b>$pending</b><br>Perlu tindak lanjut/belum diuji</div></div>
<div class="note"><b>Catatan penting:</b> hasil LULUS hanya berlaku untuk tes yang dilakukan. Program tidak dapat memastikan konektor longgar, jalur putus, korosi, short circuit atau kerusakan IC. Untuk diagnosis motherboard, gunakan pemeriksaan fisik dan alat ukur yang sesuai.</div>
<button onclick="window.print()">Cetak / Simpan PDF</button><table><thead><tr><th>Waktu</th><th>Komponen</th><th>Status</th><th>Temuan</th><th>Tindak lanjut</th></tr></thead><tbody>$($rows -join "`n")</tbody></table></div></body></html>
"@
[System.IO.File]::WriteAllText($htmlPath,$html,[System.Text.Encoding]::UTF8)
$open= [System.Windows.Forms.MessageBox]::Show("Pemeriksaan selesai.`r`n`r`nLaporan HTML: $htmlPath`r`nData JSON: $jsonPath`r`n`r`nBuka laporan sekarang?","QC Laptop Checker - Selesai",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Information)
if($open -eq [System.Windows.Forms.DialogResult]::Yes){Start-Process $htmlPath}
