# AppInstaller.Grid.KhongDau.ps1
# UI: PowerShell + WPF (XAML)
# Tabs: Install / CSVV / FONT (CSVV & FONT as placeholders to edit later)
# Labels: ASCII (khong dau)
# Tested on PowerShell 5.1+ (no ternary operator)

Add-Type -AssemblyName PresentationFramework

# ---- XAML ----
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="App Installer - Khong Dau" Width="1100" Height="700"
        Background="#1e1e1e" Foreground="White"
        FontFamily="Segoe UI" FontSize="13"
        WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <!-- Accent + tile brushes (de doi mau dong bo) -->
    <SolidColorBrush x:Key="Accent"       Color="#0f6cbd"/>
    <SolidColorBrush x:Key="TileBg"       Color="#2a2a2a"/>
    <SolidColorBrush x:Key="TileBgHover"  Color="#3a3a3a"/>
    <SolidColorBrush x:Key="TileBorder"   Color="#3f3f3f"/>

    <Style x:Key="TileCheckBox" TargetType="CheckBox">
      <Setter Property="Margin" Value="6"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Background" Value="{StaticResource TileBg}"/>
      <Setter Property="BorderBrush" Value="{StaticResource TileBorder}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="HorizontalContentAlignment" Value="Center"/>
      <Setter Property="VerticalContentAlignment" Value="Center"/>
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
        <TextBox Grid.Row="1" Name="TxtLog" Background="#181818" Foreground="White" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
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
$TxtLog             = $window.FindName("TxtLog")
$ChkSilent          = $window.FindName("ChkSilent")
$ChkAccept          = $window.FindName("ChkAccept")

# ---- Helpers ----
function Log-Msg([string]$msg){
  $TxtLog.AppendText(("{0}  {1}`r`n" -f (Get-Date).ToString("HH:mm:ss"), $msg))
  $TxtLog.ScrollToEnd()
}

function Resolve-Id([string[]]$candidates){
  foreach($id in $candidates){
    # Use 'winget show -e --id' to verify existence
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
  if($ChkAccept.IsChecked){
    $args += @("--accept-package-agreements","--accept-source-agreements")
  }
  Log-Msg ("Install: {0}" -f $id)
  $p = Start-Process -FilePath "winget" -ArgumentList $args -PassThru -WindowStyle Hidden
  $p.WaitForExit()
  $code = $p.ExitCode
  # Treat -1978335189 (0x8A15002B, APPINSTALLER_CLI_ERROR_UPDATE_NOT_APPLICABLE) as success
  if(($code -eq 0) -or ($code -eq -1978335189)){
    if($code -eq -1978335189){
      Log-Msg ("[OK] already installed / no applicable update: {0}" -f $id)
    } else {
      Log-Msg ("[OK] installed: {0}" -f $id)
    }
    return $true
  } else {
    Log-Msg ("[WARN] install failed (ExitCode={0})" -f $code)
    return $false
  }
}

# --- New: installer helpers for EXE/MSI/ZIP and GitHub + Office ODT (offline/online) ---

function Install-Exe([hashtable]$exe){
  try{
    $url = [string]$exe.Url
    if([string]::IsNullOrWhiteSpace($url)){ Log-Msg "[ERR] Exe.Url rong"; return $false }

    $file = Join-Path $env:TEMP ([IO.Path]::GetFileName(($url -split '\?')[0]))
    Log-Msg ("Download: {0}" -f $url)
    iwr -useb $url -OutFile $file

    $sha = $exe.Sha256
    if($sha){
      $hash = (Get-FileHash -Algorithm SHA256 -Path $file).Hash.ToLower()
      if($hash -ne $sha.ToLower()){
        Log-Msg ("[ERR] SHA256 mismatch. expected={0} actual={1}" -f $sha,$hash)
        return $false
      }
    }

    if($file.ToLower().EndsWith(".msi")){
      $msiArgs = "/i `"$file`" /qn /norestart"
      Log-Msg ("MSI: msiexec {0}" -f $msiArgs)
      $p = Start-Process msiexec -ArgumentList $msiArgs -PassThru -WindowStyle Hidden
      $p.WaitForExit()
      if($p.ExitCode -eq 0){ Log-Msg "[OK] msi installed"; return $true }
      Log-Msg ("[WARN] msi exit {0}" -f $p.ExitCode); return $false
    } else {
      $args = [string]$exe.Args
      if([string]::IsNullOrWhiteSpace($args)){ $args = "/S" } # default guess NSIS
      Log-Msg ("EXE: {0} {1}" -f $file,$args)
      $p = Start-Process -FilePath $file -ArgumentList $args -PassThru -WindowStyle Hidden
      $p.WaitForExit()
      if($p.ExitCode -eq 0){ Log-Msg "[OK] exe installed"; return $true }
      Log-Msg ("[WARN] exe exit {0}" -f $p.ExitCode); return $false
    }
  } catch {
    Log-Msg ("[ERR] Install-Exe: {0}" -f $_.Exception.Message); return $false
  }
}

function Install-ZipPackage([hashtable]$zip){
  try{ Add-Type -AssemblyName System.IO.Compression.FileSystem | Out-Null } catch {}
  $url     = [string]$zip.Url
  $destDir = [Environment]::ExpandEnvironmentVariables([string]$zip.DestDir)
  $exeName = [string]$zip.Exe
  $args    = [string]$zip.RunArgs
  $mkDesk  = [bool]  $zip.CreateShortcut
  $startup = [bool]  $zip.AddStartup

  if([string]::IsNullOrWhiteSpace($url) -or [string]::IsNullOrWhiteSpace($destDir)){
    Log-Msg "[ERR] Zip.Url/DestDir rong"; return $false
  }

  $zipPath = Join-Path $env:TEMP ([IO.Path]::GetFileName(($url -split '\?')[0]))
  Log-Msg ("Download: {0}" -f $url)
  iwr -useb $url -OutFile $zipPath

  if(-not (Test-Path $destDir)){ New-Item -ItemType Directory -Path $destDir -Force | Out-Null }
  [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $destDir, $true)

  if($mkDesk -and $exeName){
    $lnkPath = Join-Path ([Environment]::GetFolderPath('Desktop')) "UniKey.lnk"
    $target  = Join-Path $destDir $exeName
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($lnkPath)
    $Shortcut.TargetPath = $target
    if($args){ $Shortcut.Arguments = $args }
    $Shortcut.WorkingDirectory = $destDir
    $Shortcut.Save()
  }

  if($startup -and $exeName){
    $target = Join-Path $destDir $exeName
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run" `
      -Name "UniKey" -Value "`"$target`"" -PropertyType String -Force | Out-Null
  }

  Log-Msg "[OK] zip extracted"
  return $true
}

function Install-GitHubLatest([hashtable]$gh){
  try{
    $repo = [string]$gh.Repo   # "owner/name"
    if([string]::IsNullOrWhiteSpace($repo)){ Log-Msg "[ERR] GitHub.Repo rong"; return $false }
    $api  = "https://api.github.com/repos/$repo/releases/latest"
    Log-Msg ("GitHub API: {0}" -f $api)
    $rel = Invoke-RestMethod -UseBasicParsing -Headers @{ 'User-Agent'='PowerShell' } -Uri $api -ErrorAction Stop
    $assets = @($rel.assets)

    # uu tien EXE/MSI
    $cand = $assets | Where-Object { $_.name -match '(?i)\.(msi|exe)$' } | Select-Object -First 1
    if($null -ne $cand){
      $url = $cand.browser_download_url
      Log-Msg ("EVKey asset: {0}" -f $cand.name)
      return Install-Exe @{ Url=$url; Args="/S"; Sha256="" }
    }

    # fallback ZIP (portable)
    $zip = $assets | Where-Object { $_.name -match '(?i)\.zip$' } | Select-Object -First 1
    if($null -ne $zip){
      $url = $zip.browser_download_url
      Log-Msg ("EVKey asset zip: {0}" -f $zip.name)
      $dest = "$Env:ProgramFiles\EVKey"
      return Install-ZipPackage @{ Url=$url; DestDir=$dest; Exe="EVKey.exe"; RunArgs=""; CreateShortcut=$true; AddStartup=$true }
    }

    Log-Msg "[ERR] Khong tim thay asset phu hop tren GitHub"
    return $false
  } catch {
    Log-Msg ("[ERR] Install-GitHubLatest: {0}" -f $_.Exception.Message)
    return $false
  }
}

function Install-OfficeODT([hashtable]$opt){
  try{
    # Option defaults
    $channel  = if($opt.Channel){ $opt.Channel } else { "Current" }
    $product  = if($opt.Product){ $opt.Product } else { "O365ProPlusRetail" }
    $lang     = if($opt.Language){ $opt.Language } else { "vi-vn" }
    $srcEnv   = if($opt.SourceEnvVar){ $opt.SourceEnvVar } else { "OFFICE_SRC" }

    $work = Join-Path $env:TEMP "ODT_$(Get-Random)"
    New-Item -ItemType Directory -Path $work -Force | Out-Null

    # Download ODT setup (microsoft) and extract
    $odtExe = Join-Path $work "officedeploymenttool.exe"
    $odtUrl = "https://officecdn.microsoft.com/pr/wsus/setup.exe"
    Log-Msg ("Download ODT: {0}" -f $odtUrl)
    iwr -useb $odtUrl -OutFile $odtExe
    Start-Process -FilePath $odtExe -ArgumentList "/quiet /extract:`"$work`"" -Wait

    $setup = Join-Path $work "setup.exe"
    if(-not (Test-Path $setup)){ Log-Msg "[ERR] Khong tim thay setup.exe sau khi extract"; return $false }

    $cfg = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="$channel">
    <Product ID="$product">
      <Language ID="$lang" />
    </Product>
  </Add>
  <RemoveMSI />
  <Updates Enabled="TRUE" Channel="$channel"/>
  <Display Level="None" AcceptEULA="TRUE"/>
  <Property Name="AUTOACTIVATE" Value="1"/>
</Configuration>
"@
    $xml = Join-Path $work "config.xml"
    Set-Content -Path $xml -Value $cfg -Encoding UTF8

    # If offline source provided, use it; else download, then configure
    $src = [Environment]::GetEnvironmentVariable($srcEnv, "Process")
    if([string]::IsNullOrWhiteSpace($src)){ $src = [Environment]::GetEnvironmentVariable($srcEnv, "Machine") }
    if([string]::IsNullOrWhiteSpace($src)){ $src = [Environment]::GetEnvironmentVariable($srcEnv, "User") }

    if(-not [string]::IsNullOrWhiteSpace($src) -and (Test-Path $src)){
      Log-Msg ("Office Offline: SourcePath = {0}" -f $src)
      # create config with SourcePath
      $cfg2 = $cfg -replace "<Add ","<Add SourcePath=`"$([IO.Path]::GetFullPath($src))`" "
      Set-Content -Path $xml -Value $cfg2 -Encoding UTF8
      Start-Process -FilePath $setup -ArgumentList "/configure `"$xml`"" -Wait
      Log-Msg "[OK] Office offline configured."
      return $true
    } else {
      # online download -> local cache then configure
      $dlCfg = $cfg -replace "<Add ","<Add DownloadPath=`"$work\Office`" "
      $xmlDl = Join-Path $work "download.xml"
      Set-Content -Path $xmlDl -Value $dlCfg -Encoding UTF8
      Log-Msg "Downloading Office content (online)..."
      Start-Process -FilePath $setup -ArgumentList "/download `"$xmlDl`"" -Wait
      Log-Msg "Installing Office from downloaded cache..."
      Start-Process -FilePath $setup -ArgumentList "/configure `"$xml`"" -Wait
      Log-Msg "[OK] Office installed."
      return $true
    }
  } catch {
    Log-Msg ("[ERR] Install-OfficeODT: {0}" -f $_.Exception.Message); return $false
  }
}

# ---- Data: Apps & Groups ----
# Use "Keys" to map to $AppCatalog entries
$AppCatalog = @{
  "7zip"          = @{ Name = "7zip";            Ids = @("7zip.7zip") }
  "Chrome"        = @{ Name = "Chrome";          Ids = @("Google.Chrome") }
  "Notepad++"     = @{ Name = "Notepad++";       Ids = @("Notepad++.Notepad++") }
  "VS Code"       = @{ Name = "VS Code";         Ids = @("Microsoft.VisualStudioCode") }
  "PowerToys"     = @{ Name = "PowerToys";       Ids = @("Microsoft.PowerToys") }
  "PC Manager"    = @{ Name = "PC Manager";      Ids = @("Microsoft.PCManager") }
  "Rainmeter"     = @{ Name = "Rainmeter";       Ids = @("Rainmeter.Rainmeter") }

  # Zalo giu winget (link zalo.me chi la trang dieu huong)
  "Zalo"          = @{ Name = "Zalo";            Ids = @("VNG.ZaloPC","Zalo.Zalo","VNG.Zalo","VNGCorp.Zalo") }

  # EVKey lay release moi nhat tren GitHub (uu tien MSI/EXE, fallback ZIP)
  "EVKey"         = @{ Name = "EVKey";           GitHub = @{ Repo = "lamquangminh/EVKey" }; Ids=@("tranxuanthang.EVKey","EVKey.EVKey","EVKey") }

  # UniKey portable (ZIP)
  "UniKey"        = @{ Name = "UniKey";          Zip = @{ Url="https://www.unikey.org/assets/release/unikey46RC2-230919-win64.zip"; DestDir="$Env:ProgramFiles\UniKey"; Exe="UniKeyNT.exe"; RunArgs=""; CreateShortcut=$true; AddStartup=$true } }

  # Office ODT (winget)
  "Office ODT"    = @{ Name = "Office ODT";      Ids = @("Microsoft.OfficeDeploymentTool") }

  # Office Offline (ODT) - dung $env:OFFICE_SRC neu co (thu muc cache offline)
  "Office Offline"= @{ Name = "Office Offline";  OfficeODT = @{ Channel="Current"; Product="O365ProPlusRetail"; Language="vi-vn"; SourceEnvVar="OFFICE_SRC" } }

  # Adobe Creative Cloud (Photoshop)
  "Creative Cloud"= @{ Name = "Creative Cloud";  Ids = @("Adobe.CreativeCloud","Adobe.Photoshop") }

  # AutoCAD
  "AutoCAD"       = @{ Name = "AutoCAD";         Ids = @("Autodesk.AutoCAD","Autodesk.AutoCADLT") }
}

$Groups = @(
  @{ Title = "Essentials";       Keys = @("7zip","Chrome","Notepad++","VS Code","PowerToys","PC Manager","Rainmeter") },
  @{ Title = "VN Chat & Input";  Keys = @("Zalo","EVKey","UniKey") },
  @{ Title = "Office";           Keys = @("Office ODT","Office Offline") },
  @{ Title = "Design & CAD";     Keys = @("Creative Cloud","AutoCAD") }
)

# Dictionary to keep generated CheckBoxes
$CheckBoxes = @{}  # key -> CheckBox

# Build UI groups dynamically
foreach($g in $Groups){
  $gb = New-Object System.Windows.Controls.GroupBox
  $gb.Header = $g.Title
  $gb.Margin = "0,0,0,10"

  $panel = New-Object System.Windows.Controls.WrapPanel
  $panel.Margin = "0,0,0,0"

  foreach($k in $g.Keys){
    $info = $AppCatalog[$k]
    if(-not $info){ continue }
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Style = $window.Resources["TileCheckBox"]
    $cb.Content = $info.Name
    $cb.Tag = $k
    $cb.Width = 180
    $cb.Height = 38
    $panel.Children.Add($cb) | Out-Null
    $CheckBoxes[$k] = $cb

    # Double-click to install immediately (Ids/GitHub/Zip/Exe/OfficeODT)
    $cb.AddHandler([System.Windows.Controls.Control]::MouseDoubleClickEvent, [System.Windows.Input.MouseButtonEventHandler]{ param($s,$e)
      $key = $s.Tag
      $info2 = $AppCatalog[$key]
      if($null -eq $info2){ return }
      $s.IsEnabled = $false
      try {
        if($info2.ContainsKey("GitHub")){
          [void](Install-GitHubLatest -gh $info2.GitHub)
        } elseif($info2.ContainsKey("Zip")){
          [void](Install-ZipPackage -zip $info2.Zip)
        } elseif($info2.ContainsKey("Exe")){
          [void](Install-Exe -exe $info2.Exe)
        } elseif($info2.ContainsKey("OfficeODT")){
          [void](Install-OfficeODT -opt $info2.OfficeODT)
        } else {
          $id = Resolve-Id -candidates $info2.Ids
          if($null -eq $id){
            if($info2.Ids){ Log-Msg ("[ERR] not found on winget: {0}" -f ($info2.Ids -join " | ")) } else { Log-Msg "[ERR] no Ids defined" }
          } else {
            [void](Install-ById -id $id)
          }
        }
      } finally { $s.IsEnabled = $true }
    })
  }

  $gb.Content = $panel
  $PanelGroups.Children.Add($gb) | Out-Null
}

# Button handlers
$BtnClear.Add_Click({
  foreach($cb in $CheckBoxes.Values){ $cb.IsChecked = $false }
  Log-Msg "Selection cleared."
})

$BtnGetInstalled.Add_Click({
  Log-Msg "winget list ..."
  $p = Start-Process -FilePath "winget" -ArgumentList @("list") -PassThru -WindowStyle Hidden -RedirectStandardOutput ([System.IO.Path]::GetTempFileName())
  $p.WaitForExit()
  try {
    $out = Get-Content -Raw $p.RedirectStandardOutput
    Log-Msg $out
  } catch {
    Log-Msg "[WARN] cannot read output."
  }
})

$BtnInstallSelected.Add_Click({
  $selectedKeys = @()
  foreach($kv in $CheckBoxes.GetEnumerator()){
    if($kv.Value.IsChecked){ $selectedKeys += $kv.Key }
  }
  if($selectedKeys.Count -eq 0){ Log-Msg "Chua chon ung dung nao."; return }
  Log-Msg ("Installing {0} item(s)..." -f $selectedKeys.Count)
  foreach($k in $selectedKeys){
    $info = $AppCatalog[$k]
    $cb = $CheckBoxes[$k]
    $cb.IsEnabled = $false
    try {
      if($info.ContainsKey("GitHub")){
        [void](Install-GitHubLatest -gh $info.GitHub)
        continue
      }
      if($info.ContainsKey("Zip")){
        [void](Install-ZipPackage -zip $info.Zip)
        continue
      }
      if($info.ContainsKey("Exe")){
        [void](Install-Exe -exe $info.Exe)
        continue
      }
      if($info.ContainsKey("OfficeODT")){
        [void](Install-OfficeODT -opt $info.OfficeODT)
        continue
      }

      $id = Resolve-Id -candidates $info.Ids
      if($null -eq $id){
        Log-Msg ("[ERR] not found on winget: {0}" -f ($info.Ids -join " | "))
        continue
      }
      [void](Install-ById -id $id)
    } finally { $cb.IsEnabled = $true }
  }
  Log-Msg "Done."
})

# Show
$window.ShowDialog() | Out-Null
