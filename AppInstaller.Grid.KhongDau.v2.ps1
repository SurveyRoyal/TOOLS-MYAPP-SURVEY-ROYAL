# AppInstaller.Grid.KhongDau.v2.ps1
# UI: PowerShell + WPF (XAML) — Tabs: Install / CSVV / FONT / AutoCAD
# PowerShell 5.1 compatible, Light theme
# Tùy chọn: "Chay ngoai Console" để đẩy lệnh ra cửa sổ PowerShell riêng (UI mượt)
# ĐÃ XÓA HOÀN TOÀN: Office ODT/Offline

# TLS cho máy cũ
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
try { Add-Type -AssemblyName System.Web } catch {}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
[void][Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

# ---- XAML (Light) ----
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="App Installer - Khong Dau" Width="1100" Height="720"
        Background="#FFFFFF" Foreground="#1C1C1C"
        FontFamily="Segoe UI" FontSize="13"
        WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <SolidColorBrush x:Key="Accent"       Color="#2563EB"/>
    <SolidColorBrush x:Key="TileBg"       Color="#F2F4F7"/>
    <SolidColorBrush x:Key="TileBgHover"  Color="#E6EAF0"/>
    <SolidColorBrush x:Key="TileBorder"   Color="#D0D5DD"/>
    <SolidColorBrush x:Key="TextFg"       Color="#1C1C1C"/>

    <Style x:Key="TileCheckBox" TargetType="CheckBox">
      <Setter Property="Margin" Value="6"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="Foreground" Value="{StaticResource TextFg}"/>
      <Setter Property="Background" Value="{StaticResource TileBg}"/>
      <Setter Property="BorderBrush" Value="{StaticResource TileBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="CheckBox">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="8">
              <Grid Margin="2">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Grid>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="{StaticResource TileBgHover}"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter Property="Background"  Value="{StaticResource Accent}"/>
                <Setter Property="BorderBrush" Value="{StaticResource Accent}"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.6"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="GroupHeader" TargetType="TextBlock">
      <Setter Property="FontSize" Value="16"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="Margin" Value="0,6,0,8"/>
    </Style>
  </Window.Resources>

  <DockPanel Margin="10">
    <!-- Top bar -->
    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,10">
      <Button Name="BtnInstallSelected" Content="Install Selected" Width="140" Height="32" Margin="0,0,8,0"/>
      <Button Name="BtnClear" Content="Clear Selection" Width="140" Height="32" Margin="0,0,8,0"/>
      <Button Name="BtnGetInstalled" Content="Get Installed" Width="120" Height="32" Margin="0,0,8,0"/>
      <CheckBox Name="ChkSilent" IsChecked="True" Content="Silent" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <CheckBox Name="ChkAccept" IsChecked="True" Content="Accept EULA" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <!-- NEW: chay ngoai console -->
      <CheckBox Name="ChkConsole" IsChecked="True" Content="Chay ngoai Console" VerticalAlignment="Center" Margin="0,0,8,0"/>
    </StackPanel>

    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="200"/>
      </Grid.RowDefinitions>

      <TabControl Grid.Row="0">
        <TabItem Header="Install">
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel Name="PanelGroups" Margin="6"/>
          </ScrollViewer>
        </TabItem>

        <TabItem Header="CSVV">
          <Grid><TextBlock Margin="10" Text="Tab CSVV (de trong de sua sau)" /></Grid>
        </TabItem>

        <TabItem Header="FONT">
          <Grid><TextBlock Margin="10" Text="Tab FONT (de trong de sua sau)"/></Grid>
        </TabItem>

        <TabItem Header="AutoCAD">
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel Name="PanelAutoCAD" Margin="6"/>
          </ScrollViewer>
        </TabItem>
      </TabControl>

      <!-- Log -->
      <Grid Grid.Row="1">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Log" FontWeight="Bold" Margin="0,0,0,4"/>
        <TextBox Grid.Row="1" Name="TxtLog" Background="#FFFFFF" Foreground="#1C1C1C" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
      </Grid>
    </Grid>
  </DockPanel>
</Window>
"@

# ---- Load XAML ----
$ErrorActionPreference = 'Stop'
try {
  if ([string]::IsNullOrWhiteSpace($xaml)) { throw "XAML string rong." }
  $window = [Windows.Markup.XamlReader]::Parse($xaml)
} catch {
  Write-Host "XAML ERROR: $($_.Exception.Message)" -ForegroundColor Red
  if ($_.Exception.InnerException) { Write-Host "INNER: $($_.Exception.InnerException.Message)" -ForegroundColor Yellow }
  throw
}

# Controls
$PanelGroups        = $window.FindName("PanelGroups")
$PanelAutoCAD       = $window.FindName("PanelAutoCAD")
$BtnInstallSelected = $window.FindName("BtnInstallSelected")
$BtnClear           = $window.FindName("BtnClear")
$BtnGetInstalled    = $window.FindName("BtnGetInstalled")
$TxtLog             = $window.FindName("TxtLog")
$ChkSilent          = $window.FindName("ChkSilent")
$ChkAccept          = $window.FindName("ChkAccept")
$ChkConsole         = $window.FindName("ChkConsole")

# ==== Helpers chung ====
function Log-Msg([string]$msg){
  $TxtLog.AppendText(("{0}  {1}`r`n" -f (Get-Date).ToString("HH:mm:ss"), $msg))
  $TxtLog.ScrollToEnd()
}
function Resolve-Id([string[]]$candidates){
  foreach($id in $candidates){
    $p = Start-Process -FilePath "winget" -ArgumentList @("show","-e","--id",$id) -PassThru -WindowStyle Hidden
    $p.WaitForExit()
    if($p.ExitCode -eq 0){ return $id }
  }
  return $null
}
function Start-ExternalConsole([string]$title, [string]$scriptText){
  try{
    $tmp = Join-Path $env:TEMP ("run_" + [Guid]::NewGuid().ToString("N") + ".ps1")
    Set-Content -Path $tmp -Value $scriptText -Encoding UTF8
    $exe = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
    $args = "-NoProfile -ExecutionPolicy Bypass -NoExit -File `"$tmp`""
    Start-Process -FilePath $exe -ArgumentList $args -WindowStyle Normal | Out-Null
    Log-Msg ("[LAUNCH] {0}" -f $title)
  } catch {
    Log-Msg ("[ERR] Start-ExternalConsole: {0}" -f $_.Exception.Message)
  }
}

# ---- Install bằng winget ----
function Install-ById([string]$id, [string[]]$ExtraArgs=$null){
  if(-not $id){ return $false }

  $cmd = "winget install -e --id `"$id`""
  if($ChkSilent.IsChecked){ $cmd += " --silent" }
  if($ChkAccept.IsChecked){ $cmd += " --accept-package-agreements --accept-source-agreements" }
  if($ExtraArgs){ $cmd += " " + ($ExtraArgs -join ' ') }

  if($ChkConsole.IsChecked){
    $ps = @"
`$ErrorActionPreference='Continue'
Write-Host "=== $id ==="
Write-Host "$cmd" -ForegroundColor Cyan
& $ExecutionContext.InvokeCommand.ExpandString("$cmd")
Write-Host "`nDone. Nhan Enter de dong cua so..."
[Console]::ReadLine() | Out-Null
"@
    Start-ExternalConsole "winget $id" $ps
    return $true
  }

  Log-Msg ("Install: {0}" -f $id)
  $p = Start-Process -FilePath "winget" -ArgumentList ($cmd -replace '^winget\s+','') -PassThru -WindowStyle Hidden
  $p.WaitForExit()
  $code = $p.ExitCode
  if(($code -eq 0) -or ($code -eq -1978335189)){
    if($code -eq -1978335189){ Log-Msg ("[OK] already installed / not applicable: {0}" -f $id) }
    else { Log-Msg ("[OK] installed: {0}" -f $id) }
    return $true
  } else { Log-Msg ("[WARN] install failed (ExitCode={0})" -f $code); return $false }
}

# ---- EXE/MSI ----
function Install-Exe([hashtable]$exe){
  try{
    $url = [string]$exe.Url
    $args = if([string]::IsNullOrWhiteSpace($exe.Args)) { "/S" } else { [string]$exe.Args }
    $sha  = [string]$exe.Sha256

    if($ChkConsole.IsChecked){
      $ps = @"
`$ErrorActionPreference='Continue'
`$u = "$url"
`$f = Join-Path `$env:TEMP ([IO.Path]::GetFileName((`$u -split '\?')[0]))
Write-Host "Download: `$u" -ForegroundColor Cyan
iwr -useb "`$u" -OutFile "`$f"
if("$sha"){
  `$h = (Get-FileHash -Algorithm SHA256 -Path "`$f").Hash.ToLower()
  if(`$h -ne "$sha".ToLower()){ Write-Host "[ERR] SHA256 mismatch" -ForegroundColor Red; Write-Host "Nhan Enter de dong..."; [Console]::ReadLine() | Out-Null; return }
}
Write-Host "Run: `"`$f`" $args" -ForegroundColor Yellow
Start-Process -FilePath "`$f" -ArgumentList "$args" -Wait
Write-Host "`nDone. Nhan Enter de dong cua so..."
[Console]::ReadLine() | Out-Null
"@
      Start-ExternalConsole "exe $(Split-Path -Leaf $url)" $ps
      return $true
    }

    if([string]::IsNullOrWhiteSpace($url)){ Log-Msg "[ERR] Exe.Url rong"; return $false }
    $file = Join-Path $env:TEMP ([IO.Path]::GetFileName(($url -split '\?')[0]))
    Log-Msg ("Download: {0}" -f $url); iwr -useb $url -OutFile $file
    if($sha){
      $hash = (Get-FileHash -Algorithm SHA256 -Path $file).Hash.ToLower()
      if($hash -ne $sha.ToLower()){ Log-Msg "[ERR] SHA256 mismatch"; return $false }
    }
    if($file.ToLower().EndsWith(".msi")){
      $msiArgs = "/i `"$file`" /qn /norestart"; Log-Msg ("MSI: msiexec {0}" -f $msiArgs)
      $p = Start-Process msiexec -ArgumentList $msiArgs -PassThru -WindowStyle Hidden
    } else {
      Log-Msg ("EXE: {0} {1}" -f $file,$args)
      $p = Start-Process -FilePath $file -ArgumentList $args -PassThru -WindowStyle Hidden
    }
    $p.WaitForExit()
    if($p.ExitCode -eq 0){ Log-Msg "[OK] installed"; return $true } else { Log-Msg ("[WARN] exit {0}" -f $p.ExitCode); return $false }
  } catch { Log-Msg ("[ERR] Install-Exe: {0}" -f $_.Exception.Message); return $false }
}

# ---- ZIP (giải + tạo shortcut/startup nếu cấu hình) ----
function Install-ZipPackage([hashtable]$zip){
  try{ Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null } catch {}
  $url=[string]$zip.Url; $dest=[Environment]::ExpandEnvironmentVariables([string]$zip.DestDir)
  $exe=[string]$zip.Exe; $args=[string]$zip.RunArgs; $mkDesk=[bool]$zip.CreateShortcut; $startup=[bool]$zip.AddStartup
  if([string]::IsNullOrWhiteSpace($url) -or [string]::IsNullOrWhiteSpace($dest)){ Log-Msg "[ERR] Zip.Url/DestDir rong"; return $false }

  if($ChkConsole.IsChecked){
    $ps = @"
`$ErrorActionPreference='Continue'
`$u = "$url"
`$zipPath = Join-Path `$env:TEMP ([IO.Path]::GetFileName((`$u -split '\?')[0]))
`$dest = "$dest"
Write-Host "Download: `$u" -ForegroundColor Cyan
iwr -useb "`$u" -OutFile "`$zipPath"
if(-not (Test-Path "`$dest")){ New-Item -ItemType Directory -Path "`$dest" -Force | Out-Null }
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::ExtractToDirectory("`$zipPath","`$dest",`$true)
if("$mkDesk" -eq "True" -and "$exe"){
  `$lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) "UniKey.lnk"
  `$target = Join-Path "`$dest" "$exe"
  `$ws = New-Object -ComObject WScript.Shell
  `$sc = `$ws.CreateShortcut(`$lnk); `$sc.TargetPath = `$target; if("$args"){ `$sc.Arguments="$args" }; `$sc.WorkingDirectory="`$dest"; `$sc.Save()
}
if("$startup" -eq "True" -and "$exe"){
  `$target = Join-Path "`$dest" "$exe"
  New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "UniKey" -Value "`"`$target`"" -PropertyType String -Force | Out-Null
}
Write-Host "`nDone. Nhan Enter de dong cua so..."
[Console]::ReadLine() | Out-Null
"@
    Start-ExternalConsole "zip $(Split-Path -Leaf $url)" $ps
    return $true
  }

  $zipPath = Join-Path $env:TEMP ([IO.Path]::GetFileName(($url -split '\?')[0]))
  Log-Msg ("Download: {0}" -f $url); iwr -useb $url -OutFile $zipPath
  if(-not (Test-Path $dest)){ New-Item -ItemType Directory -Path $dest -Force | Out-Null }
  [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $dest, $true)
  if($mkDesk -and $exe){
    $lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) "UniKey.lnk"
    $target = Join-Path $dest $exe
    $ws = New-Object -ComObject WScript.Shell
    $sc = $ws.CreateShortcut($lnk); $sc.TargetPath = $target; if($args){ $sc.Arguments=$args }; $sc.WorkingDirectory=$dest; $sc.Save()
  }
  if($startup -and $exe){
    $target = Join-Path $dest $exe
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" -Name "UniKey" -Value "`"$target`"" -PropertyType String -Force | Out-Null
  }
  Log-Msg "[OK] zip extracted"; return $true
}

# ---- GitHub latest (EVKey ...) ----
function Install-GitHubLatest([hashtable]$gh){
  try{
    $repo=[string]$gh.Repo; if([string]::IsNullOrWhiteSpace($repo)){ Log-Msg "[ERR] GitHub.Repo rong"; return $false }
    $api="https://api.github.com/repos/$repo/releases/latest"; Log-Msg ("GitHub API: {0}" -f $api)
    $rel = Invoke-RestMethod -UseBasicParsing -Headers @{ 'User-Agent'='PowerShell' } -Uri $api -ErrorAction Stop
    $assets=@($rel.assets)
    $cand = $assets | Where-Object { $_.name -match '(?i)\.(msi|exe)$' } | Select-Object -First 1
    if($cand){ return Install-Exe @{ Url=$cand.browser_download_url; Args="/S"; Sha256="" } }
    $zip = $assets | Where-Object { $_.name -match '(?i)\.zip$' } | Select-Object -First 1
    if($zip){ return Install-ZipPackage @{ Url=$zip.browser_download_url; DestDir="$Env:ProgramFiles\EVKey"; Exe="EVKey.exe"; RunArgs=""; CreateShortcut=$true; AddStartup=$true } }
    Log-Msg "[ERR] Khong tim thay asset phu hop"; return $false
  } catch { Log-Msg ("[ERR] Install-GitHubLatest: {0}" -f $_.Exception.Message); return $false }
}

# ==== 7-Zip & Archive helpers (dùng cho AutoCAD/MEGA nếu cần) ====
function Find-7z(){
  $c = Get-Command 7z -ErrorAction SilentlyContinue
  if($c){ return $c.Source }
  foreach($p in @("C:\Program Files\7-Zip\7z.exe","C:\Program Files (x86)\7-Zip\7z.exe")){ if(Test-Path $p){ return $p } }
  return $null
}
function Ensure-7Zip(){
  $exe = Find-7z
  if($exe){ return $exe }
  try{ Start-Process winget -ArgumentList @("install","-e","--id","7zip.7zip","--silent","--accept-package-agreements","--accept-source-agreements") -Wait -WindowStyle Hidden | Out-Null } catch {}
  return (Find-7z)
}
function Extract-7z([string]$archive,[string]$dest,[string]$password=""){
  $seven = Ensure-7Zip
  if(-not $seven){ Log-Msg "[ERR] Khong tim thay 7-Zip (7z.exe)."; return $false }
  if(-not (Test-Path $dest)){ New-Item -ItemType Directory -Path $dest -Force | Out-Null }
  $args = @("x","-y","-aoa","-o$dest",$archive)
  if($password){ $args = @("x","-y","-aoa","-p$password","-o$dest",$archive) }
  Log-Msg ("7z: {0} {1}" -f $seven, ($args -join ' '))
  $p = Start-Process -FilePath $seven -ArgumentList $args -PassThru -WindowStyle Hidden
  $p.WaitForExit()
  if($p.ExitCode -eq 0){ return $true } else { Log-Msg ("[WARN] 7z exit {0}" -f $p.ExitCode); return $false }
}
function Extract-ArchiveAny([string]$file,[string]$dest,[string]$password=""){
  $ext = [IO.Path]::GetExtension($file).ToLower()
  if($ext -eq ".zip"){
    try{ Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null } catch {}
    [System.IO.Compression.ZipFile]::ExtractToDirectory($file,$dest,$true)
    return $true
  } elseif($ext -in @(".7z",".rar",".001",".cab")){
    return (Extract-7z -archive $file -dest $dest -password $password)
  } else { return $false }
}

# ---- MEGA (tuỳ chọn) ----
function Ensure-MegaCmd(){
  $cmd = Get-Command mega-get -ErrorAction SilentlyContinue
  if($cmd){ return $true }
  foreach($id in @("MEGA.MEGAcmd","MegaLimited.MEGAcmd","MEGA.nz.MEGAcmd")){
    try{
      $p = Start-Process -FilePath "winget" -ArgumentList @("show","-e","--id",$id) -PassThru -WindowStyle Hidden
      $p.WaitForExit()
      if($p.ExitCode -eq 0){
        Start-Process -FilePath "winget" -ArgumentList @("install","-e","--id",$id,"--silent","--accept-package-agreements","--accept-source-agreements") -Wait -WindowStyle Hidden | Out-Null
        if(Get-Command mega-get -ErrorAction SilentlyContinue){ return $true }
      }
    } catch {}
  }
  return $false
}
function Mega-DownloadToTemp([string]$megaUrl){
  if(-not (Ensure-MegaCmd)){ Log-Msg "[ERR] Khong tim thay MEGAcmd (mega-get)."; return $null }
  $outDir = Join-Path $env:TEMP ("mega_" + (Get-Random))
  New-Item -ItemType Directory -Path $outDir -Force | Out-Null
  Log-Msg ("MEGA: mega-get -> {0}" -f $outDir)
  $p = Start-Process -FilePath "mega-get" -ArgumentList @($megaUrl,$outDir) -PassThru -WindowStyle Hidden
  $p.WaitForExit()
  if($p.ExitCode -ne 0){ Log-Msg ("[ERR] mega-get exit {0}" -f $p.ExitCode); return $null }
  $items = Get-ChildItem -Path $outDir -Force -ErrorAction SilentlyContinue
  if(-not $items){ return $outDir }
  $latest = $items | Sort-Object LastWriteTime -Descending | Select-Object -First 1
  return $latest.FullName
}

# ---- Direct hóa URL (OneDrive/SharePoint/Dropbox) ----
function Transform-UrlForDownload([string]$url){
  try{
    if($url -match 'onedrive\.live\.com'){
      $u = [System.Uri]$url
      $q = [System.Web.HttpUtility]::ParseQueryString($u.Query)
      if($q["cid"] -and $q["resid"]){
        $auth = if($q["authkey"]){ "&authkey=$($q['authkey'])" } else { "" }
        return "https://onedrive.live.com/download?cid=$($q['cid'])&resid=$($q['resid'])$auth"
      }
    }
    if($url -match 'sharepoint\.com'){
      if($url -match '\?'){ if($url -notmatch 'download=1'){ return "$url&download=1" } }
      else { return "$url?download=1" }
    }
    if($url -match 'dropbox\.com'){
      if($url -match 'dl=0'){ return $url -replace 'dl=0','dl=1' } else { return "$url`?dl=1" }
    }
  } catch {}
  return $url
}

# ---- URL/local -> local temp (file/folder) ----
function Get-LocalFromSource([string]$src,[string]$password=""){
  if([string]::IsNullOrWhiteSpace($src)){ return $null }

  if($src -match '^https?://'){
    if($src -match 'mega(\.co)?\.nz'){
      $p = Mega-DownloadToTemp $src
      if(-not $p){ return $null }
      if(Test-Path $p -PathType Leaf){ return @{ Kind="file"; Path=$p } }
      if(Test-Path $p -PathType Container){ return @{ Kind="folder"; Path=$p } }
      return $null
    }
    $src = Transform-UrlForDownload $src
    $tmp = Join-Path $env:TEMP ("pkg_" + [IO.Path]::GetFileName(($src -split '\?')[0]))
    Log-Msg ("Download: {0}" -f $src)
    try{ iwr -useb $src -OutFile $tmp } catch { Log-Msg ("[ERR] Download failed: {0}" -f $_.Exception.Message); return $null }

    $low = $tmp.ToLower()
    if($low.EndsWith(".zip") -or $low.EndsWith(".7z") -or $low.EndsWith(".rar") -or $low.EndsWith(".001") -or $low.EndsWith(".cab")){
      $ext = Join-Path $env:TEMP ("pkg_" + (Get-Random))
      if(-not (Extract-ArchiveAny -file $tmp -dest $ext -password $password)){ Log-Msg "[ERR] Extract failed."; return $null }
      return @{ Kind="folder"; Path=$ext }
    } else {
      return @{ Kind="file"; Path=$tmp }
    }
  }

  if(Test-Path $src -PathType Leaf){ return @{ Kind="file"; Path=(Resolve-Path $src).Path } }
  if(Test-Path $src -PathType Container){ return @{ Kind="folder"; Path=(Resolve-Path $src).Path } }

  Log-Msg ("[ERR] Nguon khong ton tai: {0}" -f $src)
  return $null
}

# ---- AutoCAD installers (có hỗ trợ console) ----
function Invoke-Proc($file,$args,$wd){
  $p = Start-Process -FilePath $file -ArgumentList $args -WorkingDirectory $wd -PassThru -WindowStyle Hidden
  $p.WaitForExit(); return $p.ExitCode
}
function Install-AutoCADFromExe([string]$file){
  if(-not (Test-Path $file)){ Log-Msg "[ERR] EXE/MSI khong ton tai."; return $false }
  $wd = Split-Path $file -Parent

  if($ChkConsole.IsChecked){
    $ps = @'
$ErrorActionPreference='Continue'
$f = "{FILE}"
$wd = "{WD}"
if($f.ToLower().EndsWith(".msi")){
  $msi = "/i `"$f`" /qn /norestart"
  Write-Host "msiexec $msi" -ForegroundColor Cyan
  Start-Process msiexec -ArgumentList "$msi" -Wait
} else {
  $cands = @("/quiet","/silent","--quiet","--silent","/S","/VERYSILENT","/s /v`"/qn REBOOT=ReallySuppress`"","-q","")
  foreach($a in $cands){
    Write-Host "Try: `"$f`" $a" -ForegroundColor Yellow
    $p = Start-Process -FilePath "$f" -ArgumentList "$a" -PassThru -WorkingDirectory "$wd"
    $p.WaitForExit()
    if($p.ExitCode -eq 0){ break }
  }
}
Write-Host "`nDone. Nhan Enter de dong cua so..."
[Console]::ReadLine() | Out-Null
'@
    $ps = $ps.Replace('{FILE}',$file).Replace('{WD}',$wd)
    Start-ExternalConsole "AutoCAD $(Split-Path -Leaf $file)" $ps
    return $true
  }

  if($file.ToLower().EndsWith(".msi")){
    $msiArgs = "/i `"$file`" /qn /norestart"; Log-Msg ("MSI: msiexec {0}" -f $msiArgs)
    $ec = Invoke-Proc "msiexec.exe" $msiArgs $wd
    if($ec -eq 0){ Log-Msg "[OK] MSI installed"; return $true } else { Log-Msg ("[WARN] MSI exit {0}" -f $ec); return $false }
  }
  $cands = @("/quiet","/silent","--quiet","--silent","/S","/VERYSILENT","/s /v`"/qn REBOOT=ReallySuppress`"","-q","")
  foreach($a in $cands){
    Log-Msg ("Try: `"{0}`" {1}" -f $file,$a)
    $ec = Invoke-Proc $file $a $wd
    if($ec -eq 0){ Log-Msg "[OK] installed"; return $true }
  }
  Log-Msg "[WARN] Tat ca tham so silent thu nghiem deu that bai."; return $false
}
function Install-AutoCADFromFolder([string]$dir){
  if(-not (Test-Path $dir)){ Log-Msg "[ERR] Thu muc khong ton tai."; return $false }
  $msi = Get-ChildItem -Path $dir -Recurse -Filter *.msi -ErrorAction SilentlyContinue | Select-Object -First 1
  if($msi){ return Install-AutoCADFromExe $msi.FullName }
  $exe = Get-ChildItem -Path $dir -Recurse -Include AutodeskInstaller.exe,Install.exe,install.exe,Setup.exe,setup.exe,*.exe -ErrorAction SilentlyContinue |
         Sort-Object Length | Select-Object -First 1
  if($exe){ return Install-AutoCADFromExe $exe.FullName }
  Log-Msg "[ERR] Khong tim thay MSI/EXE phu hop trong thu muc."; return $false
}
function Install-AutoCADAuto([string]$version){
  if(-not $AutoCADSources.ContainsKey($version) -or [string]::IsNullOrWhiteSpace($AutoCADSources[$version])){
    Log-Msg ("[NOTE] {0}: chua chon nguon (EXE/Folder/URL). Chuot phai -> chon nguon truoc." -f $version); return $false
  }
  $src = $AutoCADSources[$version]
  $pwd = if($AutoCADPwds.ContainsKey($version)){$AutoCADPwds[$version]} else {""}
  $loc = Get-LocalFromSource -src $src -password $pwd
  if(-not $loc){ return $false }

  if($loc.Kind -eq "file"){
    Log-Msg ("{0}: cai tu EXE/MSI {1}" -f $version,$loc.Path)
    return Install-AutoCADFromExe $loc.Path
  } else {
    Log-Msg ("{0}: cai tu thu muc {1}" -f $version,$loc.Path)
    return Install-AutoCADFromFolder $loc.Path
  }
}
function Add-AutoCADContextMenu($cb,[string]$ver){
  $cm = New-Object System.Windows.Controls.ContextMenu

  $miUrl = New-Object System.Windows.Controls.MenuItem
  $miUrl.Header = "Nhap URL (OneDrive/SharePoint/Dropbox/MEGA)..."
  $miUrl.Add_Click({
    $def = if($AutoCADSources.ContainsKey($ver)){$AutoCADSources[$ver]} else {""}
    $v = [Microsoft.VisualBasic.Interaction]::InputBox("Nhap URL bo cai:","AutoCAD Source URL",$def)
    if($v){ $AutoCADSources[$ver] = $v; Log-Msg ("[{0}] Source = {1}" -f $ver,$v) }
  })
  $cm.Items.Add($miUrl) | Out-Null

  $mi1 = New-Object System.Windows.Controls.MenuItem
  $mi1.Header = "Chon EXE/MSI/ZIP/7Z/RAR..."
  $mi1.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Executable/Archive (*.exe;*.msi;*.zip;*.7z;*.rar;*.001)|*.exe;*.msi;*.zip;*.7z;*.rar;*.001|All files (*.*)|*.*"
    if($dlg.ShowDialog() -eq 'OK'){ $AutoCADSources[$ver] = $dlg.FileName; Log-Msg ("[{0}] Source = {1}" -f $ver,$dlg.FileName) }
  })
  $cm.Items.Add($mi1) | Out-Null

  $mi2 = New-Object System.Windows.Controls.MenuItem
  $mi2.Header = "Chon thu muc bo cai..."
  $mi2.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if($dlg.ShowDialog() -eq 'OK'){ $AutoCADSources[$ver] = $dlg.SelectedPath; Log-Msg ("[{0}] Source = {1}" -f $ver,$dlg.SelectedPath) }
  })
  $cm.Items.Add($mi2) | Out-Null

  $mi3 = New-Object System.Windows.Controls.MenuItem
  $mi3.Header = "Cai ngay (Auto)"
  $mi3.Add_Click({ [void](Install-AutoCADAuto -version $ver) })
  $cm.Items.Add($mi3) | Out-Null

  $cm.Items.Add((New-Object System.Windows.Controls.Separator)) | Out-Null

  $miPwd = New-Object System.Windows.Controls.MenuItem
  $miPwd.Header = "Dat mat khau giai nen..."
  $miPwd.Add_Click({
    $def = if($AutoCADPwds.ContainsKey($ver)){$AutoCADPwds[$ver]} else {""}
    $v = [Microsoft.VisualBasic.Interaction]::InputBox("Nhap mat khau (neu co):","Password",$def)
    if($v -ne $null){
      if($v){ $AutoCADPwds[$ver] = $v; Log-Msg ("[{0}] Password set ({1} ky tu)" -f $ver,$v.Length) }
      else { $AutoCADPwds.Remove($ver) | Out-Null; Log-Msg ("[{0}] Password cleared" -f $ver) }
    }
  })
  $cm.Items.Add($miPwd) | Out-Null

  $mi4 = New-Object System.Windows.Controls.MenuItem
  $mi4.Header = "Xem nguon dang dung"
  $mi4.Add_Click({
    if($AutoCADSources.ContainsKey($ver)){ Log-Msg ("[{0}] Source = {1}" -f $ver,$AutoCADSources[$ver]) }
    else { Log-Msg ("[{0}] Chua chon nguon." -f $ver) }
  })
  $cm.Items.Add($mi4) | Out-Null

  $mi5 = New-Object System.Windows.Controls.MenuItem
  $mi5.Header = "Xoa nguon"
  $mi5.Add_Click({ $AutoCADSources.Remove($ver) | Out-Null; Log-Msg ("[{0}] Source cleared" -f $ver) })
  $cm.Items.Add($mi5) | Out-Null

  $cb.ContextMenu = $cm
}

# ==== DỮ LIỆU APP ====
$AppCatalog = @{
  "7zip"          = @{ Name="7zip";            Ids=@("7zip.7zip") }
  "Chrome"        = @{ Name="Chrome";          Ids=@("Google.Chrome") }
  "Notepad++"     = @{ Name="Notepad++";       Ids=@("Notepad++.Notepad++") }
  "VS Code"       = @{ Name="VS Code";         Ids=@("Microsoft.VisualStudioCode") }
  "PowerToys"     = @{ Name="PowerToys";       Ids=@("Microsoft.PowerToys") }
  "PC Manager"    = @{ Name="PC Manager";      Ids=@("Microsoft.PCManager") }
  "Rainmeter"     = @{ Name="Rainmeter";       Ids=@("Rainmeter.Rainmeter") }

  "Zalo"          = @{
    Name="Zalo";
    Exe = @{ Url="https://res-download-pc-te-vnno-cm-1.zadn.vn/win/ZaloSetup-25.8.2.exe"; Args="/S"; Sha256="" }
    Ids = @("VNG.ZaloPC","Zalo.Zalo","VNG.Zalo","VNGCorp.Zalo")
  }
  "EVKey"         = @{ Name="EVKey"; GitHub=@{ Repo="lamquangminh/EVKey" }; Ids=@("tranxuanthang.EVKey","EVKey.EVKey","EVKey") }
  "UniKey"        = @{ Name="UniKey"; Zip=@{ Url="https://www.unikey.org/assets/release/unikey46RC2-230919-win64.zip"; DestDir="$Env:ProgramFiles\UniKey"; Exe="UniKeyNT.exe"; RunArgs=""; CreateShortcut=$true; AddStartup=$true } }

  "AutoHotkey"    = @{ Name="AutoHotkey";      Ids=@("AutoHotkey.AutoHotkey","AutoHotkey.AutoHotkey.Portable") }
  "AHK Sample"    = @{ Name="AHK Sample (Startup)"; ScriptAction="AHK_SAMPLE" }
}

$Groups = @(
  @{ Title="Essentials";       Keys=@("7zip","Chrome","Notepad++","VS Code","PowerToys","PC Manager","Rainmeter") },
  @{ Title="VN Chat & Input";  Keys=@("Zalo","EVKey","UniKey","AutoHotkey","AHK Sample") }
)

# ==== AUTO CAD ====
$AutoCADVersions = @("AutoCAD 2007","AutoCAD 2010","AutoCAD 2019","AutoCAD 2020","AutoCAD 2021","AutoCAD 2022","AutoCAD 2023","AutoCAD 2024","AutoCAD 2025","AutoCAD 2026")
$AutoCADSources  = @{}  # version -> source (EXE/Folder/URL)
$AutoCADPwds     = @{}  # version -> password (for archives)

# ---- Input helpers ----
function Ask([string]$title,[string]$label,[string]$def=""){
  [Microsoft.VisualBasic.Interaction]::InputBox($label,$title,$def)
}
function Pick-File([string]$filter="Executable/Archive (*.exe;*.msi;*.zip;*.7z;*.rar;*.001)|*.exe;*.msi;*.zip;*.7z;*.rar;*.001|All files (*.*)|*.*"){
  $dlg = New-Object System.Windows.Forms.OpenFileDialog
  $dlg.Filter = $filter; $dlg.Multiselect=$false
  if($dlg.ShowDialog() -eq 'OK'){ return $dlg.FileName } else { return $null }
}
function Pick-Folder(){
  $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
  if($dlg.ShowDialog() -eq 'OK'){ return $dlg.SelectedPath } else { return $null }
}

# ==== UI: Install tab ====
$CheckBoxes = @{}  # key -> CheckBox (kể cả AutoCAD để Install Selected)
foreach($g in $Groups){
  $gb = New-Object System.Windows.Controls.GroupBox
  $gb.Header = $g.Title
  $gb.Margin = "0,0,0,10"

  $panel = New-Object System.Windows.Controls.WrapPanel
  foreach($k in $g.Keys){
    $info = $AppCatalog[$k]; if(-not $info){ continue }
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Style = $window.Resources["TileCheckBox"]
    $cb.Content = $info.Name
    $cb.Tag = $k
    $cb.Width = 180; $cb.Height = 38
    $panel.Children.Add($cb) | Out-Null
    $CheckBoxes[$k] = $cb

    # Double-click = cài nhanh (Auto)
    $cb.AddHandler([System.Windows.Controls.Control]::MouseDoubleClickEvent,
      [System.Windows.Input.MouseButtonEventHandler]{ param($s,$e)
        $key = $s.Tag; $s.IsEnabled = $false
        try {
          $info2 = $AppCatalog[$key]
          if($info2.ScriptAction -eq "AHK_SAMPLE"){ [void](Install-AHKSample); return }
          if($info2.Exe){ [void](Install-Exe -exe $info2.Exe); return }
          if($info2.Zip){ [void](Install-ZipPackage -zip $info2.Zip); return }
          if($info2.GitHub){ [void](Install-GitHubLatest -gh $info2.GitHub); return }
          if($info2.Ids){ $id = Resolve-Id -candidates $info2.Ids; if($id){ [void](Install-ById -id $id) } else { Log-Msg ("[ERR] not found on winget: {0}" -f ($info2.Ids -join " | ")) } }
        } finally { $s.IsEnabled = $true }
      })
  }
  $gb.Content = $panel
  $PanelGroups.Children.Add($gb) | Out-Null
}

# ==== UI: AutoCAD tab ====
$gbAC = New-Object System.Windows.Controls.GroupBox
$gbAC.Header = "AutoCAD Versions"
$gbAC.Margin = "0,0,0,10"
$panelAC = New-Object System.Windows.Controls.WrapPanel
foreach($ver in $AutoCADVersions){
  $cb = New-Object System.Windows.Controls.CheckBox
  $cb.Style = $window.Resources["TileCheckBox"]
  $cb.Content = $ver
  $cb.Tag = $ver
  $cb.Width = 180; $cb.Height = 38
  $panelAC.Children.Add($cb) | Out-Null
  $CheckBoxes[$ver] = $cb

  Add-AutoCADContextMenu -cb $cb -ver $ver

  $cb.AddHandler([System.Windows.Controls.Control]::MouseDoubleClickEvent,
    [System.Windows.Input.MouseButtonEventHandler]{ param($s,$e)
      $v = $s.Tag; $s.IsEnabled = $false
      try { [void](Install-AutoCADAuto -version $v) } finally { $s.IsEnabled = $true }
    })
}
$gbAC.Content = $panelAC
$PanelAutoCAD.Children.Add($gbAC) | Out-Null

# ==== Buttons ====
$BtnClear.Add_Click({
  foreach($cb in $CheckBoxes.Values){ $cb.IsChecked = $false }
  Log-Msg "Selection cleared."
})
$BtnGetInstalled.Add_Click({
  Log-Msg "winget list ..."
  $tmp = [System.IO.Path]::GetTempFileName()
  $p = Start-Process -FilePath "winget" -ArgumentList @("list") -PassThru -WindowStyle Hidden -RedirectStandardOutput $tmp
  $p.WaitForExit()
  try { Log-Msg (Get-Content -Raw $tmp) } catch { Log-Msg "[WARN] cannot read output." }
  Remove-Item -ErrorAction SilentlyContinue $tmp
})
$BtnInstallSelected.Add_Click({
  $selected = @(); foreach($kv in $CheckBoxes.GetEnumerator()){ if($kv.Value.IsChecked){ $selected += $kv.Key } }
  if($selected.Count -eq 0){ Log-Msg "Chua chon ung dung nao."; return }
  Log-Msg ("Installing {0} item(s)..." -f $selected.Count)
  foreach($k in $selected){
    $cb = $CheckBoxes[$k]; $cb.IsEnabled = $false
    try {
      if($k -like "AutoCAD *"){
        [void](Install-AutoCADAuto -version $k)
      } else {
        $info = $AppCatalog[$k]
        if($info.ScriptAction -eq "AHK_SAMPLE"){ [void](Install-AHKSample); continue }
        if($info.Exe){ [void](Install-Exe -exe $info.Exe); continue }
        if($info.Zip){ [void](Install-ZipPackage -zip $info.Zip); continue }
        if($info.GitHub){ [void](Install-GitHubLatest -gh $info.GitHub); continue }
        if($info.Ids){ $id = Resolve-Id -candidates $info.Ids; if($id){ [void](Install-ById -id $id) } else { Log-Msg ("[ERR] not found on winget: {0}" -f ($info.Ids -join " | ")) } }
      }
    } finally { $cb.IsEnabled = $true }
  }
  Log-Msg "Done."
})

# Show UI
$window.ShowDialog() | Out-Null
