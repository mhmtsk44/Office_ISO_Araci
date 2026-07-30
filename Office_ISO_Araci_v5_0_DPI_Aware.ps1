<#
.SYNOPSIS
    Office Taşınabilir ISO Aracı - v5.0.9 (Temp Dizin + Kesin UI Çözümü)
.DESCRIPTION
    v5.0.9: ODT indirme ve çıkartma işlemleri C:\ODT yerine %TEMP% dizininde 
            benzersiz bir klasöre taşındı. İzin (Permission) sorunları önlendi.
            Adım 4'teki yazıların kesilmesini önlemek için AutoSize kapatılıp 
            yükseklikler (Height) kalıcı olarak devasa boyutlara sabitlendi.
.NOTES
    Hazırlayan : Mehmet IŞIK
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------- YÜKSEK ÇÖZÜNÜRLÜK (DPI) VE KONSOL GİZLEME YARDIMCISI ----------
$csharpCode = @"
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public class Win32Native {
    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
    
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}

public class IsoHelper {
    public static void SaveIStreamToFile(object comStream, string filePath, Action<long> onProgress) {
        IStream stream = comStream as IStream;
        if (stream == null) throw new ArgumentException("IStream COM nesnesine dönüştürülemedi.");
        
        using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)) {
            byte[] buffer = new byte[1048576]; 
            IntPtr ptr = Marshal.AllocHGlobal(4);
            long totalWritten = 0;
            int updateCounter = 0;
            try {
                while (true) {
                    stream.Read(buffer, buffer.Length, ptr);
                    int bytesRead = Marshal.ReadInt32(ptr);
                    if (bytesRead <= 0) break;
                    
                    fs.Write(buffer, 0, bytesRead);
                    totalWritten += bytesRead;
                    updateCounter++;
                    
                    if (updateCounter % 40 == 0 && onProgress != null) {
                        onProgress(totalWritten);
                    }
                }
            } finally {
                Marshal.FreeHGlobal(ptr);
            }
        }
    }
}
"@
Add-Type -TypeDefinition $csharpCode -ErrorAction Ignore
try { [Win32Native]::SetProcessDPIAware() | Out-Null } catch {}

# Eğer kod buraya kadar çökmeden (örneğin karakter hatası vermeden) geldiyse,
# arkadaki siyah PowerShell ekranını artık güvenle gizleyebiliriz (0 = Hide).
try {
    $konsol = [Win32Native]::GetConsoleWindow()
    if ($konsol -ne [IntPtr]::Zero) { [Win32Native]::ShowWindow($konsol, 0) | Out-Null }
} catch {}

[System.Windows.Forms.Application]::EnableVisualStyles()
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch {}

# ---------- AUTO-ELEVATE VE UZAKTAN ÇALIŞTIRMA (IRM) YEDEĞİ ----------
$idn = [Security.Principal.WindowsIdentity]::GetCurrent()
$prc = New-Object Security.Principal.WindowsPrincipal($idn)
if (-not $prc.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    $sp = $MyInvocation.MyCommand.Path
    try {
        if (-not [string]::IsNullOrWhiteSpace($sp) -and (Test-Path $sp)) {
            # Yerel dosya çalıştırılıyorsa (-NoExit eklendi ki hata olursa kapanmasın)
            $argList = "-NoProfile -ExecutionPolicy Bypass -NoExit -File `"$sp`""
            Start-Process powershell -ArgumentList $argList -Verb RunAs -ErrorAction Stop
        } else {
            # GitHub üzerinden (irm) çalıştırılıyorsa
            $ScriptUrl = "https://raw.githubusercontent.com/mhmtsk44/Office_ISO_Araci/refs/heads/main/Office_ISO_Araci_v5_0_DPI_Aware.ps1"
            $argList = "-NoProfile -ExecutionPolicy Bypass -NoExit -Command `"irm '$ScriptUrl' | iex`""
            Start-Process powershell -ArgumentList $argList -Verb RunAs -ErrorAction Stop
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Yönetici izni alınırken hata oluştu veya işlem reddedildi.`n`nAyrıntı: $($_.Exception.Message)", "Yönetici İzni Gerekli", 'OK', 'Error') | Out-Null
    }
    exit
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$script:LogYolu = ""
try { $script:LogYolu = Join-Path $env:TEMP "office_iso_v50.log" } catch {}
if ([string]::IsNullOrWhiteSpace($script:LogYolu)) { $script:LogYolu = "C:\Windows\Temp\office_iso_v50.log" }

# ---------- PALET ----------
$C_Bg    = [System.Drawing.Color]::White
$C_Head  = [System.Drawing.Color]::FromArgb(0, 90, 158)
$C_Text  = [System.Drawing.Color]::FromArgb(30, 35, 45)
$C_Mute  = [System.Drawing.Color]::FromArgb(120, 128, 140)
$C_Line  = [System.Drawing.Color]::FromArgb(226, 230, 236)
$C_Sel   = [System.Drawing.Color]::FromArgb(232, 242, 252)
$C_Ok    = [System.Drawing.Color]::FromArgb(16, 137, 62)
$C_Gray  = [System.Drawing.Color]::FromArgb(245, 246, 248)

$F     = New-Object System.Drawing.Font("Segoe UI", 9.5)
$F_S   = New-Object System.Drawing.Font("Segoe UI", 8.5)
$F_Q   = New-Object System.Drawing.Font("Segoe UI", 15)
$F_O   = New-Object System.Drawing.Font("Segoe UI Semibold", 10.5)
$F_B   = New-Object System.Drawing.Font("Segoe UI Semibold", 10)
$F_Big = New-Object System.Drawing.Font("Segoe UI Light", 30)
$F_Mn  = New-Object System.Drawing.Font("Consolas", 8.5)

# =====================================================================
# PENCERE
# =====================================================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Office ISO Aracı"
$form.ClientSize = New-Object System.Drawing.Size(640, 560)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.Font = $F
$form.BackColor = $C_Bg

# --- adım göstergesi (üst ince şerit) ---
$stepBar = New-Object System.Windows.Forms.Panel
$stepBar.Location='0,0'; $stepBar.Size='640,4'; $stepBar.BackColor=$C_Line
$form.Controls.Add($stepBar)
$stepFill = New-Object System.Windows.Forms.Panel
$stepFill.Location='0,0'; $stepFill.Size='160,4'; $stepFill.BackColor=$C_Head
$stepBar.Controls.Add($stepFill)

$lblAdim = New-Object System.Windows.Forms.Label
$lblAdim.Location='28,20'; $lblAdim.AutoSize=$true; $lblAdim.Font=$F_S; $lblAdim.ForeColor=$C_Mute
$form.Controls.Add($lblAdim)

$lblSoru = New-Object System.Windows.Forms.Label
$lblSoru.Location='26,45'; $lblSoru.AutoSize=$true; $lblSoru.Font=$F_Q; $lblSoru.ForeColor=$C_Text
$lblSoru.UseCompatibleTextRendering=$true
$form.Controls.Add($lblSoru)

# --- içerik alanı ---
$pnl = New-Object System.Windows.Forms.Panel
$pnl.Location='26,85'; $pnl.Size='588,398'; $pnl.BackColor=$C_Bg
$form.Controls.Add($pnl)

# --- alt buton çubuğu ---
$footer = New-Object System.Windows.Forms.Panel
$footer.Location='0,492'; $footer.Size='640,68'; $footer.BackColor=$C_Gray
$form.Controls.Add($footer)
$fLine = New-Object System.Windows.Forms.Panel
$fLine.Location='0,0'; $fLine.Size='640,1'; $fLine.BackColor=$C_Line
$footer.Controls.Add($fLine)

function New-Btn {
    param($Text,$X,$Y,$W,$H,$Bg,$Fg,$Font,$Parent)
    $b = New-Object System.Windows.Forms.Button
    $b.Text=$Text; $b.Location=New-Object System.Drawing.Point($X,$Y)
    $b.Size=New-Object System.Drawing.Size($W,$H)
    $b.FlatStyle='Flat'; $b.FlatAppearance.BorderSize=0
    $b.BackColor=$Bg; $b.ForeColor=$Fg; $b.Font=$Font; $b.Cursor='Hand'
    $Parent.Controls.Add($b); return $b
}
$btnGeri  = New-Btn "‹  Geri"     26  20 100 32 $C_Gray $C_Text  $F   $footer
$btnGeri.FlatAppearance.BorderSize=1; $btnGeri.FlatAppearance.BorderColor=$C_Line
$btnIptal = New-Btn "İptal"      386  20  98 32 $C_Gray $C_Mute  $F   $footer
$btnIptal.FlatAppearance.BorderSize=1; $btnIptal.FlatAppearance.BorderColor=$C_Line
$btnIleri = New-Btn "İleri  ›"   494  20 120 32 $C_Head ([System.Drawing.Color]::White) $F_B $footer

# =====================================================================
# SEÇİM DURUMU
# =====================================================================
$script:Sec = @{
    Paket   = "temel"
    Surum   = "365"
    Mimari  = "x64"
    Dil     = "tr-tr"
    Klasor  = [Environment]::GetFolderPath('Desktop')
    Ozel    = @("Word","Excel","PowerPoint","Outlook","OneNote")
    OdtDosya = ""
}
$script:Adim = 1
$script:Calisiyor = $false

# --- büyük seçim kartı üretici ---
function New-Secenek {
    param($Baslik,$Aciklama,$Y,$Deger,$Grup,$Parent)
    $p = New-Object System.Windows.Forms.Panel
    $p.Location = New-Object System.Drawing.Point(0,$Y)
    $p.Size = New-Object System.Drawing.Size(588,80)
    $p.BackColor=$C_Bg; $p.BorderStyle='FixedSingle'; $p.Cursor='Hand'
    $p.Tag = @{ Deger=$Deger; Grup=$Grup }

    $t = New-Object System.Windows.Forms.Label
    $t.Text=$Baslik; $t.Location='16,12'; $t.AutoSize=$true; $t.Font=$F_O; $t.ForeColor=$C_Text
    $t.Cursor='Hand'; $p.Controls.Add($t)

    $a = New-Object System.Windows.Forms.Label
    $a.Text=$Aciklama; $a.Location='17,38'; $a.AutoSize=$true; $a.MaximumSize=New-Object System.Drawing.Size(530,0)
    $a.Font=$F_S; $a.ForeColor=$C_Mute; $a.Cursor='Hand'; $p.Controls.Add($a)

    $tik = New-Object System.Windows.Forms.Label
    $tik.Text="✓"; $tik.Location='550,26'; $tik.Size='26,26'
    $tik.Font=$F_O; $tik.ForeColor=$C_Head; $tik.Visible=$false
    $p.Controls.Add($tik); $p.Tag.Tik = $tik

    $tikla = {
        $kart = $this; while ($kart -isnot [System.Windows.Forms.Panel]) { $kart = $kart.Parent }
        foreach ($k in $kart.Parent.Controls) {
            if ($k -is [System.Windows.Forms.Panel] -and $k.Tag -and $k.Tag.Grup -eq $kart.Tag.Grup) {
                $k.BackColor = $C_Bg; $k.Tag.Tik.Visible = $false
            }
        }
        $kart.BackColor = $C_Sel; $kart.Tag.Tik.Visible = $true
        $script:Sec[$kart.Tag.Grup] = $kart.Tag.Deger
        if ($script:GuncelleOzet) { & $script:GuncelleOzet }
    }
    $p.Add_Click($tikla); $t.Add_Click($tikla); $a.Add_Click($tikla)
    $Parent.Controls.Add($p)
    return $p
}

# =====================================================================
# ADIM 1 — NE KURULACAK
# =====================================================================
$s1 = New-Object System.Windows.Forms.Panel
$s1.Location='0,0'; $s1.Size='588,398'; $s1.Visible=$false; $pnl.Controls.Add($s1)

$o1a = New-Secenek "Temel Paket"  "Word, Excel, PowerPoint, Outlook, OneNote  —  çoğu kullanıcı için yeterli"  0 "temel" "Paket" $s1
$o1b = New-Secenek "Tam Paket"    "Temel paket + Access, Publisher, Teams, OneDrive"                         90 "tam"   "Paket" $s1
$o1c = New-Secenek "Özel Seçim"   "Bileşenleri tek tek ben seçeyim"                                          180 "ozel"  "Paket" $s1

$clbOzel = New-Object System.Windows.Forms.CheckedListBox
$clbOzel.Location='0,270'; $clbOzel.Size='588,120'; $clbOzel.CheckOnClick=$true
$clbOzel.MultiColumn=$true; $clbOzel.ColumnWidth=186; $clbOzel.BorderStyle='FixedSingle'
$clbOzel.Font=$F; $clbOzel.Visible=$false
$tumBilesen = @("Word","Excel","PowerPoint","Outlook","OneNote","Access",
                "Publisher","Microsoft Teams","OneDrive","Skype for Business",
                "Visio Pro","Project Pro")
foreach ($b in $tumBilesen) {
    [void]$clbOzel.Items.Add($b)
    $clbOzel.SetItemChecked($clbOzel.Items.Count-1, ($script:Sec.Ozel -contains $b))
}
$s1.Controls.Add($clbOzel)

# =====================================================================
# ADIM 2 — SÜRÜM
# =====================================================================
$s2 = New-Object System.Windows.Forms.Panel
$s2.Location='0,0'; $s2.Size='588,398'; $s2.Visible=$false; $pnl.Controls.Add($s2)

$o2a = New-Secenek "Microsoft 365 Apps"  "Abonelik  —  sürekli güncellenir, en yeni özellikler"  0 "365"  "Surum" $s2
$o2b = New-Secenek "Office LTSC 2024"    "Kalıcı lisans  —  KMS/MAK, kurumsal ortamlar için"     90 "2024" "Surum" $s2

$lblMim = New-Object System.Windows.Forms.Label
$lblMim.Text="Mimari ve dil"; $lblMim.Location='0,185'; $lblMim.AutoSize=$true
$lblMim.Font=$F_B; $lblMim.ForeColor=$C_Mute; $s2.Controls.Add($lblMim)

$cmbMim = New-Object System.Windows.Forms.ComboBox
$cmbMim.Location='0,210'; $cmbMim.Size='284,24'; $cmbMim.DropDownStyle='DropDownList'; $cmbMim.FlatStyle='Flat'
[void]$cmbMim.Items.AddRange(@("64-bit  (önerilen)","32-bit","Her ikisi  (x86 + x64)"))
$cmbMim.SelectedIndex=0
$cmbMim.Add_SelectedIndexChanged({
    if     ($cmbMim.SelectedIndex -eq 0) { $script:Sec.Mimari="x64" }
    elseif ($cmbMim.SelectedIndex -eq 1) { $script:Sec.Mimari="x86" }
    else                                 { $script:Sec.Mimari="x86x64" }
})
$s2.Controls.Add($cmbMim)

$cmbDil = New-Object System.Windows.Forms.ComboBox
$cmbDil.Location='304,210'; $cmbDil.Size='284,24'; $cmbDil.DropDownStyle='DropDownList'; $cmbDil.FlatStyle='Flat'
[void]$cmbDil.Items.AddRange(@("Türkçe","İngilizce"))
$cmbDil.SelectedIndex=0
$cmbDil.Add_SelectedIndexChanged({
    if ($cmbDil.SelectedIndex -eq 0) { $script:Sec.Dil="tr-tr" } else { $script:Sec.Dil="en-us" }
})
$s2.Controls.Add($cmbDil)

$lblLtscNot = New-Object System.Windows.Forms.Label
$lblLtscNot.Location='0,250'; $lblLtscNot.AutoSize=$true; $lblLtscNot.Font=$F_S; $lblLtscNot.ForeColor=$C_Mute
$lblLtscNot.Text="ISO yalnızca kurulum kaynağıdır, lisans içermez."
$s2.Controls.Add($lblLtscNot)

# =====================================================================
# ADIM 3 — KONUM + ÖZET
# =====================================================================
$s3 = New-Object System.Windows.Forms.Panel
$s3.Location='0,0'; $s3.Size='588,398'; $s3.Visible=$false; $pnl.Controls.Add($s3)

$txtKlasor = New-Object System.Windows.Forms.TextBox
$txtKlasor.Location='0,10'; $txtKlasor.Size='476,24'; $txtKlasor.BorderStyle='FixedSingle'
$txtKlasor.Text=$script:Sec.Klasor
$s3.Controls.Add($txtKlasor)
$btnGz = New-Btn "Gözat" 486 9 102 26 $C_Gray $C_Text $F $s3
$btnGz.FlatAppearance.BorderSize=1; $btnGz.FlatAppearance.BorderColor=$C_Line
$btnGz.Add_Click({
    $d = New-Object System.Windows.Forms.FolderBrowserDialog
    $d.Description="ISO dosyasının kaydedileceği klasör"
    if ($d.ShowDialog() -eq 'OK') { $txtKlasor.Text=$d.SelectedPath; $script:Sec.Klasor=$d.SelectedPath }
})

$ozetKutu = New-Object System.Windows.Forms.Panel
$ozetKutu.Location='0,50'; $ozetKutu.Size='588,190'; $ozetKutu.BackColor=$C_Gray; $ozetKutu.BorderStyle='FixedSingle'
$s3.Controls.Add($ozetKutu)

$lblOzet = New-Object System.Windows.Forms.Label
$lblOzet.Location='16,12'; $lblOzet.AutoSize=$true; $lblOzet.Font=$F; $lblOzet.ForeColor=$C_Text
$ozetKutu.Controls.Add($lblOzet)

# --- sistem durumu + onarım ---
$sysKutu = New-Object System.Windows.Forms.Panel
$sysKutu.Location='0,250'; $sysKutu.Size='588,70'; $sysKutu.BorderStyle='FixedSingle'
$s3.Controls.Add($sysKutu)

$lblSys = New-Object System.Windows.Forms.Label
$lblSys.Location='14,12'; $lblSys.AutoSize=$true; $lblSys.Font=$F_S
$sysKutu.Controls.Add($lblSys)

$btnOnar = New-Btn "Onar" 472 20 100 26 $C_Gray $C_Text $F $sysKutu
$btnOnar.FlatAppearance.BorderSize=1; $btnOnar.FlatAppearance.BorderColor=$C_Line
$btnOnar.Visible=$false

$chkOdt = New-Object System.Windows.Forms.CheckBox
$chkOdt.Text="ODT indirilemezse hazır dosyayı kullan"
$chkOdt.Location='0,335'; $chkOdt.AutoSize=$true
$chkOdt.ForeColor=$C_Mute; $chkOdt.Font=$F_S
$s3.Controls.Add($chkOdt)

$btnOdtSec = New-Btn "Dosya seç" 486 330 102 24 $C_Gray $C_Text $F_S $s3
$btnOdtSec.FlatAppearance.BorderSize=1; $btnOdtSec.FlatAppearance.BorderColor=$C_Line
$btnOdtSec.Visible=$false
$btnOdtSec.Add_Click({
    $o = New-Object System.Windows.Forms.OpenFileDialog
    $o.Filter="ODT kurulum dosyası (*.exe)|*.exe"
    if ($o.ShowDialog() -eq 'OK') { $script:Sec.OdtDosya=$o.FileName; $chkOdt.Text="ODT: " + [IO.Path]::GetFileName($o.FileName) }
})
$chkOdt.Add_CheckedChanged({
    $btnOdtSec.Visible = $chkOdt.Checked
    if (-not $chkOdt.Checked) { $script:Sec.OdtDosya=""; $chkOdt.Text="ODT indirilemezse hazır dosyayı kullan" }
})

# =====================================================================
# ADIM 4 — İLERLEME (KESİLMELER ÖNLENDİ, KUTULAR BÜYÜTÜLDÜ)
# =====================================================================
$s4 = New-Object System.Windows.Forms.Panel
$s4.Location='0,0'; $s4.Size='588,398'; $s4.Visible=$false; $pnl.Controls.Add($s4)

$lblYuzde = New-Object System.Windows.Forms.Label
$lblYuzde.Text="0%"
$lblYuzde.Location='0,0'; $lblYuzde.Size='588,90'
$lblYuzde.Font=$F_Big; $lblYuzde.ForeColor=$C_Head; $lblYuzde.TextAlign='BottomCenter'
$lblYuzde.AutoSize=$false
$lblYuzde.UseCompatibleTextRendering=$true
$s4.Controls.Add($lblYuzde)

$pbArka = New-Object System.Windows.Forms.Panel
$pbArka.Location='0,100'; $pbArka.Size='588,8'; $pbArka.BackColor=$C_Line
$s4.Controls.Add($pbArka)
$pbDolu = New-Object System.Windows.Forms.Panel
$pbDolu.Location='0,0'; $pbDolu.Size='0,8'; $pbDolu.BackColor=$C_Head
$pbArka.Controls.Add($pbDolu)

# AutoSize KAPALI. Sabit ve devasa yükseklik (40px) verildi. Asla kesilemez.
$lblAsama = New-Object System.Windows.Forms.Label
$lblAsama.Location='0,115'; $lblAsama.Size='588,40'
$lblAsama.Font=$F_B; $lblAsama.ForeColor=$C_Text; $lblAsama.TextAlign='MiddleCenter'
$lblAsama.AutoSize=$false 
$lblAsama.UseCompatibleTextRendering=$true
$s4.Controls.Add($lblAsama)

# AutoSize KAPALI. 40px bol yükseklik.
$lblDetay = New-Object System.Windows.Forms.Label
$lblDetay.Location='0,155'; $lblDetay.Size='588,40'
$lblDetay.Font=$F_S; $lblDetay.ForeColor=$C_Mute; $lblDetay.TextAlign='MiddleCenter'
$lblDetay.AutoSize=$false
$lblDetay.UseCompatibleTextRendering=$true
$s4.Controls.Add($lblDetay)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline=$true; $txtLog.ScrollBars="Vertical"
$txtLog.Location='0,200'; $txtLog.Size='588,195'
$txtLog.ReadOnly=$true; $txtLog.BackColor=[System.Drawing.Color]::FromArgb(28,30,36)
$txtLog.ForeColor=[System.Drawing.Color]::FromArgb(130,225,140)
$txtLog.BorderStyle='FixedSingle'; $txtLog.Font=$F_Mn
$s4.Controls.Add($txtLog)

# =====================================================================
# IMAPI2 KONTROL + ONARIM
# =====================================================================
function Test-Imapi {
    try {
        $t = New-Object -ComObject IMAPI2FS.MsftFileSystemImage
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($t) | Out-Null
        return $true
    } catch { return $false }
}
function Get-Oscdimg {
    $yollar = @(
        "${env:ProgramFiles(x86)}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe",
        "${env:ProgramFiles}\Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\amd64\Oscdimg\oscdimg.exe"
    )
    foreach ($y in $yollar) { if (Test-Path $y) { return $y } }
    $c = Get-Command oscdimg.exe -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    return ""
}

$script:ImapiVar = Test-Imapi
$script:Oscdimg  = Get-Oscdimg

function Guncelle-Sys {
    if ($script:ImapiVar) {
        $lblSys.Text = "✓  ISO motoru hazır (IMAPI2)`r`n     Sistem ISO oluşturmaya uygun."
        $lblSys.ForeColor = $C_Ok
        $btnOnar.Visible = $false
    } elseif ($script:Oscdimg) {
        $lblSys.Text = "!  IMAPI2 çalışmıyor — yedek motor (oscdimg) kullanılacak.`r`n     İşlem yine de tamamlanabilir."
        $lblSys.ForeColor = [System.Drawing.Color]::FromArgb(190,120,0)
        $btnOnar.Visible = $true
    } else {
        $lblSys.Text = "✕  ISO motoru bulunamadı. Onar'a basın.`r`n     Başarısız olursa dosyalar klasör olarak bırakılır."
        $lblSys.ForeColor = [System.Drawing.Color]::FromArgb(200,50,50)
        $btnOnar.Visible = $true
    }
}

$btnOnar.Add_Click({
    $btnOnar.Enabled=$false; $btnOnar.Text="Onarılıyor..."
    $rapor = @()
    $dll = "$env:SystemRoot\System32\imapi2fs.dll"
    if (Test-Path $dll) {
        $rapor += "imapi2fs.dll mevcut."
        try {
            $r = Start-Process regsvr32.exe -ArgumentList "/s `"$dll`"" -Wait -PassThru -NoNewWindow
            $rapor += "regsvr32 kaydı yenilendi (kod $($r.ExitCode))."
        } catch { $rapor += "regsvr32 çalıştırılamadı: $($_.Exception.Message)" }
    } else {
        $rapor += "imapi2fs.dll BULUNAMADI — sistem dosyası eksik."
        $rapor += "Yönetici komut isteminde 'sfc /scannow' çalıştırın."
    }
    $script:ImapiVar = Test-Imapi
    $script:Oscdimg  = Get-Oscdimg
    Guncelle-Sys
    $sonucMetni = "IMAPI2 hâlâ çalışmıyor."
    if ($script:ImapiVar) { $sonucMetni = "IMAPI2 artık ÇALIŞIYOR." }
    [System.Windows.Forms.MessageBox]::Show(
        ($rapor -join "`r`n") + "`r`n`r`nSonuç: $sonucMetni",
        "Onarım Raporu",'OK','Information')|Out-Null
    $btnOnar.Enabled=$true; $btnOnar.Text="Onar"
})

# =====================================================================
# ÖZET + GEZİNME
# =====================================================================
$script:GuncelleOzet = {
    $bilesenler = @()
    if     ($script:Sec.Paket -eq "temel") { $bilesenler = @("Word","Excel","PowerPoint","Outlook","OneNote") }
    elseif ($script:Sec.Paket -eq "tam")   { $bilesenler = @("Word","Excel","PowerPoint","Outlook","OneNote","Access","Publisher","Microsoft Teams","OneDrive") }
    else {
        for($i=0;$i -lt $clbOzel.Items.Count;$i++){ if($clbOzel.GetItemChecked($i)){ $bilesenler += [string]$clbOzel.Items[$i] } }
    }
    $script:Sec.Ozel = $bilesenler
    $clbOzel.Visible = ($script:Sec.Paket -eq "ozel")

    $surumAd = "Microsoft 365 Apps"
    if ($script:Sec.Surum -eq "2024") { $surumAd = "Office LTSC 2024" }
    $mimAd = $script:Sec.Mimari.Replace("x86x64","32 + 64 bit").Replace("x64","64-bit").Replace("x86","32-bit")
    $dilAd = "Türkçe"; if ($script:Sec.Dil -eq "en-us") { $dilAd = "İngilizce" }
    $carp = 1; if ($script:Sec.Mimari -eq "x86x64") { $carp = 2 }
    $tahmin = [math]::Round((1.6 + ($bilesenler.Count * 0.35)) * $carp, 1)

    $lblOzet.Text = "Sürüm`t`t: $surumAd`r`n" +
                    "Mimari / Dil`t: $mimAd  ·  $dilAd`r`n" +
                    "Bileşenler`t: $($bilesenler.Count) adet`r`n" +
                    "  $($bilesenler -join ', ')`r`n`r`n" +
                    "Tahmini boyut`t: ~$tahmin GB"
    if ($script:Sec.Surum -eq "2024") { $lblLtscNot.Text = "LTSC 2024 için kanal otomatik PerpetualVL2024 olarak ayarlanır.`r`nISO lisans içermez (KMS/MAK gerekir)." }
    else { $lblLtscNot.Text = "ISO yalnızca kurulum kaynağıdır, lisans içermez (abonelik girişi gerekir)." }
}
$clbOzel.Add_ItemCheck({ $form.BeginInvoke([Action]{ & $script:GuncelleOzet }) | Out-Null })

$basliklar = @(
    @{ A="Adım 1 / 4"; S="Hangi uygulamalar kurulsun?" },
    @{ A="Adım 2 / 4"; S="Hangi Office sürümü?" },
    @{ A="Adım 3 / 4"; S="Nereye kaydedilsin?" },
    @{ A="Adım 4 / 4"; S="ISO oluşturuluyor" }
)

function Goster-Adim($n) {
    $script:Adim = $n
    $s1.Visible=($n -eq 1); $s2.Visible=($n -eq 2); $s3.Visible=($n -eq 3); $s4.Visible=($n -eq 4)
    $lblAdim.Text = $basliklar[$n-1].A
    $lblSoru.Text = $basliklar[$n-1].S
    $stepFill.Width = [int](640 * $n / 4)
    $btnGeri.Visible  = ($n -gt 1 -and $n -lt 4)
    $btnIptal.Visible = ($n -lt 4)
    if ($n -eq 3) { $btnIleri.Text = "ISO Oluştur"; $btnIleri.BackColor = $C_Ok; & $script:GuncelleOzet; Guncelle-Sys }
    elseif ($n -eq 4) { $btnIleri.Text = "Kapat"; $btnIleri.BackColor = $C_Head; $btnIleri.Enabled = $false }
    else { $btnIleri.Text = "İleri  ›"; $btnIleri.BackColor = $C_Head }
}

$btnGeri.Add_Click({ if ($script:Adim -gt 1) { Goster-Adim ($script:Adim - 1) } })
$btnIptal.Add_Click({ $form.Close() })

# ilk seçimleri işaretle
$o1a.BackColor=$C_Sel; $o1a.Tag.Tik.Visible=$true
$o2a.BackColor=$C_Sel; $o2a.Tag.Tik.Visible=$true
Goster-Adim 1

# =====================================================================
# İLERİ / BAŞLAT
# =====================================================================
$btnIleri.Add_Click({

    if ($script:Adim -eq 1) {
        & $script:GuncelleOzet
        if ($script:Sec.Ozel.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("En az bir uygulama seçin.","Eksik Seçim",'OK','Warning')|Out-Null; return
        }
        Goster-Adim 2; return
    }
    if ($script:Adim -eq 2) { Goster-Adim 3; return }
    if ($script:Adim -eq 4) { $form.Close(); return }

    # ---- ADIM 3: BAŞLAT ----
    $script:Sec.Klasor = $txtKlasor.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($script:Sec.Klasor) -or -not (Test-Path $script:Sec.Klasor -PathType Container)) {
        [System.Windows.Forms.MessageBox]::Show("Geçerli bir klasör seçin.","Eksik Bilgi",'OK','Warning')|Out-Null; return
    }
    if ($chkOdt.Checked -and -not (Test-Path $script:Sec.OdtDosya -PathType Leaf)) {
        [System.Windows.Forms.MessageBox]::Show("ODT dosyasını seçin veya kutunun işaretini kaldırın.","Eksik Bilgi",'OK','Warning')|Out-Null; return
    }

    $gerekli = 10; if ($script:Sec.Mimari -eq "x86x64") { $gerekli = 18 }
    try { $cBos = [math]::Round((Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace/1GB,1) } catch { $cBos = 999 }
    if ($cBos -lt $gerekli) {
        $c = [System.Windows.Forms.MessageBox]::Show(
            "C: sürücüsünde $cBos GB boş alan var, ~$gerekli GB gerekiyor.`r`n`r`nDevam edilsin mi?",
            "Disk Alanı",'YesNo','Warning')
        if ($c -ne 'Yes') { return }
    }

    # ---- config (açık atamalar) ----
    $kanal = "Current"; if ($script:Sec.Surum -eq "2024") { $kanal = "PerpetualVL2024" }
    $urun  = "Office365"; if ($script:Sec.Surum -eq "2024") { $urun = "Office2024" }

    $cfg = @{}
    $cfg['UrunAdi']    = $urun
    $cfg['Kanal']      = $kanal
    $cfg['Mimari']     = $script:Sec.Mimari
    $cfg['Dil']        = $script:Sec.Dil
    $cfg['Hedef']      = $script:Sec.Klasor
    $cfg['LogYolu']    = [string]$script:LogYolu
    $cfg['ManuelOdt']  = [string]$script:Sec.OdtDosya
    $cfg['Secimler']   = @($script:Sec.Ozel)
    $cfg['Oscdimg']    = [string]$script:Oscdimg
    $cfg['ImapiVar']   = [bool]$script:ImapiVar

    foreach ($k in @('UrunAdi','Kanal','Mimari','Dil','Hedef','LogYolu')) {
        if ([string]::IsNullOrWhiteSpace([string]$cfg[$k])) {
            [System.Windows.Forms.MessageBox]::Show("İç hata: '$k' boş.","Yapılandırma Hatası",'OK','Error')|Out-Null; return
        }
    }

    # ---- İLERLEME PAYLAŞIMI (gerçek bayt) ----
    $script:Ilerleme = [hashtable]::Synchronized(@{
        Yuzde=0; Asama="Hazırlanıyor"; Detay=""; Bitti=$false
    })

    $script:Calisiyor = $true
    Goster-Adim 4
    $txtLog.Clear()

    $scriptBlock = {
        param($cfg, $ilr)

        $LogYolu = [string]$cfg.LogYolu
        if ([string]::IsNullOrWhiteSpace($LogYolu)) { $LogYolu = "C:\Windows\Temp\office_iso_fb.log" }
        function LogYaz($m) {
            $s = "[$(Get-Date -Format 'HH:mm:ss')] $m"
            try { $s | Out-File -FilePath $LogYolu -Append -Encoding UTF8 -ErrorAction Stop } catch {}
            try { Write-Information $s -Tags "Log" } catch {}
        }
        function Ilerle($y,$a,$d) { $ilr.Yuzde=[int]$y; $ilr.Asama=$a; $ilr.Detay=$d }
        function Sonuc($ok,$m,$iso){ [pscustomobject]@{ Basarili=$ok; Mesaj=$m; IsoYolu=$iso } }
        function Test-Exe([string]$y){
            if ([string]::IsNullOrWhiteSpace($y) -or !(Test-Path $y)) { return $false }
            try {
                $fs=[System.IO.File]::OpenRead($y)
                try { $b=New-Object byte[] 2; $n=$fs.Read($b,0,2) } finally { $fs.Close() }
                if ($n -lt 2) { return $false }
                return ([System.Text.Encoding]::ASCII.GetString($b) -eq "MZ" -and (Get-Item $y).Length -ge 1MB)
            } catch { return $false }
        }
        function KlasorBoyut($p) {
            try { return (Get-ChildItem $p -Recurse -File -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum } catch { return 0 }
        }

        LogYaz "===== v5.0.9 işlem başladı ====="
        LogYaz "Ürün=$($cfg.UrunAdi) Kanal=$($cfg.Kanal) Mimari=$($cfg.Mimari) Dil=$($cfg.Dil)"
        LogYaz "Bileşenler: $($cfg.Secimler -join ', ')"

        # ODT'Yİ TEMP KLASÖRÜNE AÇMA (İzin sorunlarını engeller)
        $ODT = Join-Path $env:TEMP "ODT_Kurulum_$([guid]::NewGuid().ToString().Substring(0,8))"
        New-Item -ItemType Directory -Force -Path $ODT | Out-Null
        
        $X86="$ODT\x86"; $X64="$ODT\x64"
        if ($cfg.Mimari -in @("x86","x86x64")) { New-Item -ItemType Directory -Force -Path $X86|Out-Null }
        if ($cfg.Mimari -in @("x64","x86x64")) { New-Item -ItemType Directory -Force -Path $X64|Out-Null }

        # ---------- 1) ODT (0-5%) ----------
        Ilerle 1 "Dağıtım aracı hazırlanıyor" ""
        $odtExe="$ODT\odt.exe"
        if (-not [string]::IsNullOrWhiteSpace($cfg.ManuelOdt)) {
            LogYaz "Manuel ODT: $($cfg.ManuelOdt)"
            if (-not (Test-Exe $cfg.ManuelOdt)) { return Sonuc $false "Seçilen ODT dosyası geçersiz." $null }
            Copy-Item $cfg.ManuelOdt $odtExe -Force
        } else {
            try { [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12 } catch {}
            $hdr=@{ "User-Agent"="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36" }
            $adaylar=@("https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_20131-20090.exe")
            LogYaz "Microsoft İndirme Merkezi taranıyor..."
            foreach ($u in @("https://www.microsoft.com/en-us/download/details.aspx?id=49117",
                             "https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117")) {
                try {
                    $r=Invoke-WebRequest -Uri $u -Headers $hdr -UseBasicParsing -TimeoutSec 40 -ErrorAction Stop
                    foreach ($x in [regex]::Matches($r.Content,'https://download\.microsoft\.com/[^"'' \\]*officedeploymenttool[^"'' \\]*\.exe')) {
                        if ($adaylar -notcontains $x.Value) { $adaylar += $x.Value }
                    }
                } catch {}
            }
            $adaylar += "https://go.microsoft.com/fwlink/p/?LinkID=626065"
            $indi=$false
            foreach ($link in $adaylar) {
                try {
                    Remove-Item $odtExe -Force -ErrorAction SilentlyContinue
                    Invoke-WebRequest -Uri $link -OutFile $odtExe -Headers $hdr -MaximumRedirection 10 -UseBasicParsing -TimeoutSec 180 -ErrorAction Stop
                    if (Test-Exe $odtExe) { LogYaz "ODT indirildi ($([math]::Round((Get-Item $odtExe).Length/1MB,1)) MB)."; $indi=$true; break }
                    else { LogYaz "Geçersiz içerik, sıradaki link." }
                } catch { LogYaz "Link başarısız: $($_.Exception.Message)" }
            }
            if (-not $indi) {
                return Sonuc $false ("Dağıtım aracı indirilemedi.`r`n`r`nÇÖZÜM: microsoft.com/download/details.aspx?id=49117 " +
                    "adresinden dosyayı elle indirip 3. adımdaki `"hazır dosyayı kullan`" seçeneğiyle gösterin.") $null
            }
        }
        Unblock-File -Path $odtExe -ErrorAction SilentlyContinue
        $p=Start-Process $odtExe -ArgumentList "/extract:`"$ODT`" /quiet" -Wait -PassThru -NoNewWindow
        Remove-Item $odtExe -Force -ErrorAction SilentlyContinue
        
        # KAFA KARIŞTIRAN ÖRNEK XML DOSYALARINI KÖKÜNDEN SİL
        Remove-Item "$ODT\*.xml" -Force -ErrorAction SilentlyContinue
        
        if ($p.ExitCode -ne 0 -or !(Test-Path "$ODT\setup.exe")) { return Sonuc $false "Dağıtım aracı açılamadı (kod $($p.ExitCode))." $null }
        Ilerle 5 "Yapılandırma oluşturuluyor" ""
        LogYaz "setup.exe hazır."

        # ---------- 2) XML ----------
        $ltsc = ($cfg.UrunAdi -eq "Office2024")
        if ($ltsc) { $uid="ProPlus2024Volume"; $vid="VisioPro2024Volume"; $pid2="ProjectPro2024Volume" }
        else       { $uid="O365ProPlusRetail"; $vid="VisioProRetail";     $pid2="ProjectProRetail" }

        $tumApp=@("Word","Excel","PowerPoint","Outlook","OneNote","Access","Publisher","Teams","Groove","Lync")
        $haric=@()
        foreach ($a in $tumApp) {
            if ($ltsc -and ($a -in @("Publisher","Groove","Lync"))) { continue }
            
            $g=$a
            if ($a -eq "Teams")  { $g="Microsoft Teams" }
            if ($a -eq "Lync")   { $g="Skype for Business" }
            if ($a -eq "Groove") { $g="OneDrive" }
            if ($cfg.Secimler -notcontains $g) { $haric += $a }
        }
        
        $excXML=($haric|ForEach-Object{"      <ExcludeApp ID=`"$_`" />"}) -join "`r`n"
        $ekExc=($haric|Where-Object{$_ -in @("Groove","Teams","Lync")}|ForEach-Object{"<ExcludeApp ID=`"$_`" />"}) -join ""
        
        $ekXML=""
        if ($cfg.Secimler -contains "Visio Pro")   { $ekXML += "`r`n    <Product ID=`"$vid`"><Language ID=`"$($cfg.Dil)`" />$ekExc</Product>" }
        if ($cfg.Secimler -contains "Project Pro") { $ekXML += "`r`n    <Product ID=`"$pid2`"><Language ID=`"$($cfg.Dil)`" />$ekExc</Product>" }

        $sablon=@"
<Configuration>
  <Add OfficeClientEdition="{0}" Channel="{1}">
    <Product ID="{2}">
      <Language ID="{3}" />
{4}
    </Product>{5}
  </Add>
  <Display Level="None" AcceptEULA="TRUE" />
</Configuration>
"@
        $utf8=New-Object System.Text.UTF8Encoding($false)
        if ($cfg.Mimari -in @("x86","x86x64")) {
            [System.IO.File]::WriteAllText("$X86\c86.xml",($sablon -f "32",$cfg.Kanal,$uid,$cfg.Dil,$excXML,$ekXML),$utf8)
            Copy-Item "$ODT\setup.exe" "$X86\setup.exe"
        }
        if ($cfg.Mimari -in @("x64","x86x64")) {
            [System.IO.File]::WriteAllText("$X64\c64.xml",($sablon -f "64",$cfg.Kanal,$uid,$cfg.Dil,$excXML,$ekXML),$utf8)
            Copy-Item "$ODT\setup.exe" "$X64\setup.exe"
        }
        LogYaz "XML hazır. Hariç tutulanlar: $($haric -join ', ')"

        # ---------- 3) İNDİRME (5-70%) — GERÇEK BAYT ----------
        $adimlar=@()
        if ($cfg.Mimari -in @("x86","x86x64")) { $adimlar += ,@($X86,"c86.xml","32-bit") }
        if ($cfg.Mimari -in @("x64","x86x64")) { $adimlar += ,@($X64,"c64.xml","64-bit") }

        $tahminGB = 1.6 + ($cfg.Secimler.Count * 0.35)
        $hedefBayt = [int64]($tahminGB * 1GB)
        $dilim = 65.0 / $adimlar.Count
        $sira = 0

        foreach ($ad in $adimlar) {
            $taban = 5 + ($sira * $dilim)
            LogYaz "$($ad[2]) paketleri indiriliyor..."
            Ilerle $taban "$($ad[2]) paketleri indiriliyor" "başlatılıyor..."
            $proc = Start-Process -FilePath "$($ad[0])\setup.exe" -ArgumentList "/download `"$($ad[1])`"" `
                        -WorkingDirectory $ad[0] -PassThru -NoNewWindow
            $baslangic = Get-Date
            while (-not $proc.HasExited) {
                Start-Sleep -Milliseconds 1500
                $b = KlasorBoyut $ad[0]
                $oran = 0.0
                if ($hedefBayt -gt 0) { $oran = [math]::Min(1.0, $b / $hedefBayt) }
                $gecen = ((Get-Date) - $baslangic).TotalSeconds
                $kalanMetin = ""
                if ($oran -gt 0.02 -and $gecen -gt 10) {
                    $toplamTah = $gecen / $oran
                    $kalanDk = [math]::Ceiling(($toplamTah - $gecen)/60)
                    if ($kalanDk -gt 0 -and $kalanDk -lt 600) { $kalanMetin = "  ·  ~$kalanDk dk kaldı" }
                }
                $mbStr = "{0:N2} / ~{1:N1} GB" -f ($b/1GB), $tahminGB
                Ilerle ($taban + ($dilim * $oran)) "$($ad[2]) paketleri indiriliyor" "$mbStr$kalanMetin"
            }
            
            $cikisKodu = -1
            try { $proc.WaitForExit(); $cikisKodu = [int]$proc.ExitCode } catch {}
            
            if ($cikisKodu -ne 0) { 
                $hataMesaji = if ($cikisKodu -eq -1) { "Bilinmeyen ODT Çökmesi" } else { $cikisKodu }
                return Sonuc $false "$($ad[2]) indirme başarısız (kod: $hataMesaji). Kanal/ürün uyumsuz olabilir." $null 
            }
            
            if (-not (Test-Path "$($ad[0])\Office\Data")) { return Sonuc $false "$($ad[2]): paket klasörü oluşmadı." $null }
            $sira++
            LogYaz "$($ad[2]) tamamlandı ($([math]::Round((KlasorBoyut $ad[0])/1GB,2)) GB)."
        }
        Remove-Item "$ODT\setup.exe" -Force -ErrorAction SilentlyContinue

        # ---------- 4) ISO (70-92%) — GERÇEK BAYT VE C# BYPASS ----------
        $kaynakBayt = KlasorBoyut $ODT
        $isoAdi="$($cfg.UrunAdi)_$($cfg.Mimari)_$(Get-Date -Format 'yyyyMMdd').iso"
        $isoYolu=Join-Path $cfg.Hedef $isoAdi
        Ilerle 70 "ISO kalıbı yazılıyor" ""
        LogYaz "ISO yazılıyor: $isoYolu"

        $isoTamam = $false
        if ($cfg.ImapiVar) {
            try {
                $fsi=New-Object -ComObject IMAPI2FS.MsftFileSystemImage
                $fsi.FreeMediaBlocks=-1; $fsi.FileSystemsToCreate=4; $fsi.VolumeName="OFFICE_ISO"
                $fsi.Root.AddTree($ODT,$false)
                $img=$fsi.CreateResultImage()
                
                $yaziciKallback = [System.Action[long]] {
                    param($yazilan)
                    if ($kaynakBayt -gt 0) {
                        $o = [math]::Min(1.0, $yazilan / $kaynakBayt)
                        Ilerle (70 + 22*$o) "ISO kalıbı yazılıyor" ("{0:N2} / {1:N2} GB" -f ($yazilan/1GB),($kaynakBayt/1GB))
                    }
                }
                [IsoHelper]::SaveIStreamToFile($img.ImageStream, $isoYolu, $yaziciKallback)
                
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($img)|Out-Null
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($fsi)|Out-Null
                [GC]::Collect()
                
                $isoTamam = $true
                LogYaz "IMAPI2 (C# bypass) ile ISO yazıldı."
            } catch { LogYaz "IMAPI2 başarısız: $($_.Exception.Message)" }
        }

        if (-not $isoTamam -and -not [string]::IsNullOrWhiteSpace($cfg.Oscdimg)) {
            LogYaz "Yedek motor (oscdimg) deneniyor..."
            Ilerle 78 "ISO yazılıyor (yedek motor)" ""
            $op = Start-Process -FilePath $cfg.Oscdimg -ArgumentList "-u2 -udfver102 -lOFFICE_ISO `"$ODT`" `"$isoYolu`"" -Wait -PassThru -NoNewWindow
            if ($op.ExitCode -eq 0) { $isoTamam = $true; LogYaz "oscdimg ile ISO yazıldı." }
            else { LogYaz "oscdimg başarısız (kod $($op.ExitCode))." }
        }

        if (-not $isoTamam) {
            $yedekKlasor = Join-Path $cfg.Hedef ("Office_Kurulum_" + (Get-Date -Format 'yyyyMMdd'))
            LogYaz "ISO oluşturulamadı, dosyalar klasöre taşınıyor: $yedekKlasor"
            Ilerle 85 "Dosyalar klasöre kopyalanıyor" ""
            Copy-Item $ODT $yedekKlasor -Recurse -Force -ErrorAction SilentlyContinue
            Ilerle 100 "Tamamlandı (klasör)" ""
            return Sonuc $true "ISO motoru çalışmadı; kurulum dosyaları klasör olarak hazırlandı. USB'ye bu klasörü kopyalayabilirsiniz." $yedekKlasor
        }

        if (!(Test-Path $isoYolu) -or (Get-Item $isoYolu).Length -lt 100MB) {
            return Sonuc $false "ISO boyutu beklenenden küçük, doğrulanamadı." $isoYolu
        }
        $gb=[math]::Round((Get-Item $isoYolu).Length/1GB,2)
        LogYaz "ISO tamam: $gb GB"

        # ---------- 5) SHA256 (92-100%) ----------
        Ilerle 93 "Bütünlük doğrulanıyor (SHA256)" "$gb GB"
        $h=(Get-FileHash -Path $isoYolu -Algorithm SHA256).Hash
        LogYaz "SHA256: $h"
        [System.IO.File]::WriteAllText("$isoYolu.sha256","$h  $isoAdi",[System.Text.Encoding]::ASCII)

        Ilerle 98 "Geçici dosyalar temizleniyor" ""
        Remove-Item $ODT -Recurse -Force -ErrorAction SilentlyContinue
        Ilerle 100 "Tamamlandı" "$gb GB"
        LogYaz "BAŞARILI."
        return Sonuc $true "ISO başarıyla oluşturuldu ($gb GB)." $isoYolu
    }

    $script:rs=[runspacefactory]::CreateRunspace()
    $script:rs.ApartmentState=[System.Threading.ApartmentState]::STA
    $script:rs.ThreadOptions=[System.Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $script:rs.Open()
    $script:ps=[powershell]::Create().AddScript($scriptBlock).AddArgument($cfg).AddArgument($script:Ilerleme)
    $script:ps.Runspace=$script:rs
    $script:li=0
    $script:ar=$script:ps.BeginInvoke()

    $script:timer=New-Object System.Windows.Forms.Timer
    $script:timer.Interval=300
    $script:timer.Add_Tick({
        # gerçek ilerleme
        $y = [int]$script:Ilerleme.Yuzde
        $lblYuzde.Text = "$y%"
        $pbDolu.Width  = [int](588 * $y / 100)
        $lblAsama.Text = [string]$script:Ilerleme.Asama
        $lblDetay.Text = [string]$script:Ilerleme.Detay

        while ($script:li -lt $script:ps.Streams.Information.Count) {
            $txtLog.AppendText("$([string]$script:ps.Streams.Information[$script:li].MessageData)`r`n"); $script:li++
        }
        if (-not $script:ar.IsCompleted) { return }
        $script:timer.Stop(); $script:Calisiyor=$false

        $son=$null
        try { $son=$script:ps.EndInvoke($script:ar)|Select-Object -Last 1 }
        catch { $son=[pscustomobject]@{Basarili=$false;Mesaj="Beklenmeyen hata: $($_.Exception.Message)";IsoYolu=$null} }
        foreach ($e in $script:ps.Streams.Error) { $txtLog.AppendText("[HATA] $e`r`n") }
        try { $script:ps.Dispose(); $script:rs.Close() } catch {}

        $btnIleri.Enabled=$true; $btnIleri.Text="Kapat"

        if ($son -and $son.Basarili) {
            $lblYuzde.Text="100%"; $pbDolu.Width=588; $pbDolu.BackColor=$C_Ok
            $lblYuzde.ForeColor=$C_Ok; $lblAsama.Text="Tamamlandı"
            $lblDetay.Text=[string]$son.IsoYolu
            [System.Windows.Forms.MessageBox]::Show("$($son.Mesaj)`r`n`r`nKonum:`r`n$($son.IsoYolu)","Tamamlandı",'OK','Information')|Out-Null
        } else {
            $pbDolu.BackColor=[System.Drawing.Color]::FromArgb(200,50,50)
            $lblYuzde.ForeColor=[System.Drawing.Color]::FromArgb(200,50,50)
            $lblAsama.Text="Başarısız"
            $m="Sonuç alınamadı."; if ($son) { $m=$son.Mesaj }
            $lblDetay.Text=""
            $txtLog.AppendText("[BAŞARISIZ] $m`r`n")
            [System.Windows.Forms.MessageBox]::Show("İşlem başarısız.`r`n`r`n$m`r`n`r`nGünlük: $script:LogYolu","Hata",'OK','Error')|Out-Null
        }
    })
    $script:timer.Start()
})

$form.Add_FormClosing({
    if ($script:Calisiyor) {
        if ([System.Windows.Forms.MessageBox]::Show("İşlem sürüyor. Kapatılsın mı?","Uyarı",'YesNo','Warning') -ne 'Yes') { $_.Cancel=$true }
        else { try { $script:timer.Stop(); $script:ps.Stop(); $script:rs.Close() } catch {} }
    }
})
$form.ShowDialog()|Out-Null