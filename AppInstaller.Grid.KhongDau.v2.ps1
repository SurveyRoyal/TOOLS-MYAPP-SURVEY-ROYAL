# AppInstaller.Grid.KhongDau.v2.ps1
# UI: PowerShell + WPF (XAML) — Tabs: Install / CSVV / FONT
# PowerShell 5.1 compatible — Có nút chuyển Light/Dark

Add-Type -AssemblyName PresentationFramework

# ---- XAML ----
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="App Installer - Khong Dau" Width="1100" Height="700"
        Background="#14161A" Foreground="#F2F2F2"
        FontFamily="Segoe UI" FontSize="13"
        WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <!-- Palette resources (sẽ bị đổi màu động khi bấm nút Theme) -->
    <SolidColorBrush x:Key="Accent"       Color="#2D7DFF"/>
    <SolidColorBrush x:Key="TileBg"       Color="#30343B"/>
    <SolidColorBrush x:Key="TileBgHover"  Color="#3B4048"/>
    <SolidColorBrush x:Key="TileBorder"   Color="#60646D"/>
    <SolidColorBrush x:Key="TextFg"       Color="#F2F2F2"/>

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
      <Button Name="BtnTheme" Content="Theme: Dark" Width="120" Height="32" Margin="10,0,0,0"/>
    </StackPanel>

    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="160"/>
      </Grid.RowDefinitions>

      <TabControl Grid.Row="0">
        <TabItem Header="Install">
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel Name="PanelGroups" Margin="6"/>
          </ScrollViewer>
        </TabItem>

        <TabItem Header="CSVV">
          <Grid>
            <TextBlock Margin="10" Text="Tab CSVV (de trong de sua sau)" />
          </Grid>
        </TabItem>

        <TabItem Header="FONT">
          <Grid>
            <TextBlock Margin="10" Text="Tab FONT (de trong de sua sau)"/>
          </Grid>
        </TabItem>
      </TabControl>

      <!-- Log -->
      <Grid Grid.Row="1">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Text="Log" FontWeight="Bold" Margin="0,0,0,4"/>
        <TextBox Grid.Row="1" Name="TxtLog" Background="#0F1115" Foreground="#F2F2F2" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
      </Grid>
    </Grid>
  </DockPanel>
</Window>
"@

# ---- Load XAML ----
$reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find controls
$PanelGroups        = $window.FindName("PanelGroups")
$BtnInstallSelected = $window.FindName("BtnInstallSelected")
$BtnClear           = $window.FindName("BtnClear")
$BtnGetInstalled    = $window.FindName("BtnGetInstalled")
$BtnTheme           = $window.FindName("BtnTheme")
$TxtLog             = $window.FindName("TxtLog")
$ChkSilent          = $window.FindName("ChkSilent")
$ChkAccept          = $window.FindName("ChkAccept")

# ---- Theme helpers ----
function Set-BrushColor([string]$key, [string]$hex){
  $b = $window.Resources[$key]
  if(-not $b){
    $b = New-Object Windows.Media.SolidColorBrush
    $window.Resources.Add($key, $b)
  }
  $b.Color = [Windows.Media.ColorConverter]::ConvertFromString($hex)
}

$ThemeDark = @{
  WindowBg = "#14161A"; WindowFg = "#F2F2F2"; LogBg = "#0F1115"; LogFg = "#F2F2F2";
  Accent   = "#2D7DFF"; TileBg = "#30343B"; TileBgHover = "#3B4048"; TileBorder = "#60646D"; TextFg="#F2F2F2"
}
$ThemeLight = @{
  WindowBg = "#FFFFFF"; WindowFg = "#1C1C1C"; LogBg = "#FFFFFF"; LogFg = "#1C1C1C";
  Accent   = "#2563EB"; TileBg = "#F2F4F7"; TileBgHover = "#E6EAF0"; TileBorder = "#D0D5DD"; TextFg="#1C1C1C"
}

function Apply-Theme([string]$mode){
  $t = if($mode -eq "Light"){ $ThemeLight } else { $ThemeDark }
  Set-BrushColor 'Accent'      $t.Accent
  Set-BrushColor 'TileBg'      $t.TileBg
  Set-BrushColor 'TileBgHover' $t.TileBgHover
  Set-BrushColor 'TileBorder'  $t.TileBorder
  Set-BrushColor 'TextFg'      $t.TextFg

  $window.Background = New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($t.WindowBg))
  $window.Foreground = New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($t.WindowFg))
  $TxtLog.Background = New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($t.LogBg))
  $TxtLog.Foreground = New-Object Windows.Media.SolidColorBrush ([Windows.Media.ColorConverter]::ConvertFromString($t.LogFg))
}

# default theme
Apply-Theme -mode "Dark"
$BtnTheme.Tag = "Dark"
$BtnTheme.Add_Click({
  if($BtnTheme.Tag -eq "Dark"){
    Apply-Theme -mode "Light"; $BtnTheme.Tag="Light"; $BtnTheme.Content="Theme: Light"
  } else {
    Apply-Theme -mode "Dark";  $BtnTheme.Tag="Dark";  $BtnTheme.Content="Theme: Dark"
  }
})

# ---- Common helpers ----
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
function Install-ById([string]$id){
  if(-not $id){ return $false }
  $args = @("install","-e","--id",$id)
  if($ChkSilent.IsChecked){ $args += "--silent" }
  if($ChkAccept.IsChecked){ $args += @("--accept-package-agreements","--accept-source-agreements") }
  Log-Msg ("Install: {0}" -f $id)
  $p = Start-Process -FilePath "winget" -ArgumentList $args -PassThru -WindowStyle Hidden
  $p.WaitForExit()
  $code = $p.ExitCode
  if(($code -eq 0) -or ($code -eq -1978335189)){
    if($code -eq -1978335189){ Log-Msg ("[OK] already installed / not applicable: {0}" -f $id) }
    else { Log-Msg ("[OK] installed: {0}" -f $id) }
    return $true
  } else { Log-Msg ("[WARN] install failed (ExitCode={0})" -f $code); return $false }
}

# --- EXE/MSI/ZIP, GitHub latest, Office ODT ---
function Install-Exe([hashtable]$exe){
  try{
    $url = [string]$exe.Url; if([string]::IsNullOrWhiteSpace($url)){ Log-Msg "[ERR] Exe.Url rong"; return $false }
    $file = Join-Path $env:TEMP ([IO.Path]::GetFileName(($url -split '\?')[0]))
    Log-Msg ("Download: {0}" -f $url); iwr -useb $url -OutFile $file
    $sha = $exe.Sha256
    if($sha){
      $hash = (Get-FileHash -Algorithm SHA256 -Path $file).Hash.ToLower()
      if($hash -ne $sha.ToLower()){ Log-Msg "[ERR] SHA256 mismatch"; return $false }
    }
    if($file.ToLower().EndsWith(".msi")){
      $msiArgs = "/i `"$file`" /qn /norestart"; Log-Msg ("MSI: msiexec {0}" -f $msiArgs)
      $p = Start-Process msiexec -ArgumentList $msiArgs -PassThru -WindowStyle Hidden
    } else {
      $args = if([string]::IsNullOrWhiteSpace($exe.Args)) { "/S" } else { [string]$exe.Args }
      Log-Msg ("EXE: {0} {1}" -f $file,$args)
      $p = Start-Process -FilePath $file -ArgumentList $args -PassThru -WindowStyle Hidden
    }
    $p.WaitForExit()
    if($p.ExitCode -eq 0){ Log-Msg "[OK] installed"; return $true } else { Log-Msg ("[WARN] exit {0}" -f $p.ExitCode); return $false }
  } catch { Log-Msg ("[ERR] Install-Exe: {0}" -f $_.Exception.Message); return $false }
}
function Install-ZipPackage([hashtable]$zip){
  try{ Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null } catch {}
  $url=[string]$zip.Url; $dest=[Environment]::ExpandEnvironmentVariables([string]$zip.DestDir)
  $exe=[string]$zip.Exe; $args=[string]$zip.RunArgs; $mkDesk=[bool]$zip.CreateShortcut; $startup=[bool]$zip.AddStartup
  if([string]::IsNullOrWhiteSpace($url) -or [string]::IsNullOrWhiteSpace($dest)){ Log-Msg "[ERR] Zip.Url/DestDir rong"; return $false }
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
function Install-OfficeODT([hashtable]$opt){
  try{
    $channel = if($opt.Channel){$opt.Channel}else{"Current"}
    $product = if($opt.Product){$opt.Product}else{"O365ProPlusRetail"}
    $lang    = if($opt.Language){$opt.Language}else{"vi-vn"}
    $srcEnv  = if($opt.SourceEnvVar){$opt.SourceEnvVar}else{"OFFICE_SRC"}
    $work = Join-Path $env:TEMP "ODT_$(Get-Random)"; New-Item -ItemType Directory -Path $work -Force | Out-Null
    $odtExe = Join-Path $work "officedeploymenttool.exe"; $odtUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Log-Msg ("Download ODT: {0}" -f $odtUrl); iwr -useb $odtUrl -OutFile $odtExe
    Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:`"$work`"" -Wait
    $setup = Join-Path $work "setup.exe"; if(-not (Test-Path $setup)){ Log-Msg "[ERR] Khong tim thay setup.exe"; return $false }
    $cfg = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel">
    <Product ID="$product"><Language ID="$lang" /></Product>
  </Add>
  <RemoveMSI /><Updates Enabled="TRUE" Channel="$channel"/>
  <Display Level="None" AcceptEULA="TRUE"/><Property Name="AUTOACTIVATE" Value="1"/>
</Configuration>
"@
    $xml = Join-Path $work "config.xml"; Set-Content -Path $xml -Value $cfg -Encoding UTF8
    $src = [Environment]::GetEnvironmentVariable($srcEnv,"Process"); if(-not $src){ $src=[Environment]::GetEnvironmentVariable($srcEnv,"Machine") }; if(-not $src){ $src=[Environment]::GetEnvironmentVariable($srcEnv,"User") }
    if($src -and (Test-Path $src)){
      Log-Msg ("Office Offline: SourcePath = {0}" -f $src)
      $cfg2 = $cfg -replace "<Add ","<Add SourcePath=`"$([IO.Path]::GetFullPath($src))`" "
      Set-Content -Path $xml -Value $cfg2 -Encoding UTF8
      Start-Process -FilePath $setup -ArgumentList "/configure `"$xml`"" -Wait
      Log-Msg "[OK] Office offline configured."; return $true
    } else {
      $dlCfg = $cfg -replace "<Add ","<Add DownloadPath=`"$work\Office`" "
      $xmlDl = Join-Path $work "download.xml"; Set-Content -Path $xmlDl -Value $dlCfg -Encoding UTF8
      Log-Msg "Downloading Office content (online)..."; Start-Process -FilePath $setup -ArgumentList "/download `"$xmlDl`"" -Wait
      Log-Msg "Installing Office from downloaded cache..."; Start-Process -FilePath $setup -ArgumentList "/configure `"$xml`"" -Wait
      Log-Msg "[OK] Office installed."; return $true
    }
  } catch { Log-Msg ("[ERR] Install-OfficeODT: {0}" -f $_.Exception.Message); return $false }
}

# ---- Data: Apps & Groups ----
$AppCatalog = @{
  "7zip"          = @{ Name="7zip";            Ids=@("7zip.7zip") }
  "Chrome"        = @{ Name="Chrome";          Ids=@("Google.Chrome") }
  "Notepad++"     = @{ Name="Notepad++";       Ids=@("Notepad++.Notepad++") }
  "VS Code"       = @{ Name="VS Code";         Ids=@("Microsoft.VisualStudioCode") }
  "PowerToys"     = @{ Name="PowerToys";       Ids=@("Microsoft.PowerToys") }
  "PC Manager"    = @{ Name="PC Manager";      Ids=@("Microsoft.PCManager") }
  "Rainmeter"     = @{ Name="Rainmeter";       Ids=@("Rainmeter.Rainmeter") }

  # Zalo: ưu tiên EXE bạn cung cấp + winget dự phòng
  "Zalo"          = @{
    Name="Zalo";
    Exe = @{ Url="https://res-download-pc-te-vnno-cm-1.zadn.vn/win/ZaloSetup-25.8.2.exe"; Args="/S"; Sha256="" };
    Ids = @("VNG.ZaloPC","Zalo.Zalo","VNG.Zalo","VNGCorp.Zalo")
  }

  # EVKey: release mới nhất trên GitHub
  "EVKey"         = @{ Name="EVKey"; GitHub=@{ Repo="lamquangminh/EVKey" }; Ids=@("tranxuanthang.EVKey","EVKey.EVKey","EVKey") }

  # UniKey: ZIP portable
  "UniKey"        = @{ Name="UniKey"; Zip=@{ Url="https://www.unikey.org/assets/release/unikey46RC2-230919-win64.zip"; DestDir="$Env:ProgramFiles\UniKey"; Exe="UniKeyNT.exe"; RunArgs=""; CreateShortcut=$true; AddStartup=$true } }

  # Office
  "Office ODT"    = @{ Name="Office ODT"; Ids=@("Microsoft.OfficeDeploymentTool") }
  "Office Offline"= @{ Name="Office Offline"; OfficeODT=@{ Channel="Current"; Product="O365ProPlusRetail"; Language="vi-vn"; SourceEnvVar="OFFICE_SRC" } }

  # Design & CAD
  "Creative Cloud"= @{ Name="Creative Cloud"; Ids=@("Adobe.CreativeCloud","Adobe.Photoshop") }
  "AutoCAD"       = @{ Name="AutoCAD";      Ids=@("Autodesk.AutoCAD","Autodesk.AutoCADLT") }
}

$Groups = @(
  @{ Title="Essentials";       Keys=@("7zip","Chrome","Notepad++","VS Code","PowerToys","PC Manager","Rainmeter") },
  @{ Title="VN Chat & Input";  Keys=@("Zalo","EVKey","UniKey") },
  @{ Title="Office";           Keys=@("Office ODT","Office Offline") },
  @{ Title="Design & CAD";     Keys=@("Creative Cloud","AutoCAD") }
)

# UI build
$CheckBoxes = @{}
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

    # Double-click to install
    $cb.AddHandler([System.Windows.Controls.Control]::MouseDoubleClickEvent,
      [System.Windows.Input.MouseButtonEventHandler]{ param($s,$e)
        $key = $s.Tag; $info2 = $AppCatalog[$key]; if($null -eq $info2){ return }
        $s.IsEnabled = $false
        try {
          if($info2.ContainsKey("GitHub"))       { [void](Install-GitHubLatest -gh $info2.GitHub) }
          elseif($info2.ContainsKey("Zip"))      { [void](Install-ZipPackage -zip $info2.Zip) }
          elseif($info2.ContainsKey("Exe"))      { [void](Install-Exe -exe $info2.Exe) }
          elseif($info2.ContainsKey("OfficeODT")){ [void](Install-OfficeODT -opt $info2.OfficeODT) }
          else {
            $id = Resolve-Id -candidates $info2.Ids
            if($id){ [void](Install-ById -id $id) } else {
              if($info2.Ids){ Log-Msg ("[ERR] not found on winget: {0}" -f ($info2.Ids -join " | ")) } else { Log-Msg "[ERR] no Ids defined" }
            }
          }
        } finally { $s.IsEnabled = $true }
      })
  }
  $gb.Content = $panel
  $PanelGroups.Children.Add($gb) | Out-Null
}

# Buttons
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
    $info = $AppCatalog[$k]; $cb = $CheckBoxes[$k]; $cb.IsEnabled = $false
    try {
      if($info.ContainsKey("GitHub"))       { [void](Install-GitHubLatest -gh $info.GitHub); continue }
      if($info.ContainsKey("Zip"))          { [void](Install-ZipPackage -zip $info.Zip);   continue }
      if($info.ContainsKey("Exe"))          { [void](Install-Exe -exe $info.Exe);          continue }
      if($info.ContainsKey("OfficeODT"))    { [void](Install-OfficeODT -opt $info.OfficeODT); continue }
      $id = Resolve-Id -candidates $info.Ids
      if($id){ [void](Install-ById -id $id) } else { Log-Msg ("[ERR] not found on winget: {0}" -f ($info.Ids -join " | ")) }
    } finally { $cb.IsEnabled = $true }
  }
  Log-Msg "Done."
})

# Show
$window.ShowDialog() | Out-Null
