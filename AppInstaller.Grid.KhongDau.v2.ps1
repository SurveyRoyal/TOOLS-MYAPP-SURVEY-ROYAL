# AppInstaller.Grid.KhongDau.v2.ps1 (enhanced)
# UI: PowerShell + WPF (XAML) - PS 5.1 compatible
# Tabs: Install / CSVV / FONT (2 tab sau de trong)
# Tinh nang moi:
#  - Search loc tile
#  - Force cai (winget --force)
#  - Upgrade Selected
#  - Select All/None/Install Group tren tung nhom
#  - Exit code -1978335189 coi la OK (already installed / not applicable)

Add-Type -AssemblyName PresentationFramework

# ---- XAML ----
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="App Installer - Khong Dau" Width="1180" Height="720"
        Background="#1e1e1e" Foreground="White" WindowStartupLocation="CenterScreen">
  <Window.Resources>
    <Style x:Key="TileCheckBox" TargetType="CheckBox">
      <Setter Property="Margin" Value="6"/>
      <Setter Property="Padding" Value="10,6"/>
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Background" Value="#2a2a2a"/>
      <Setter Property="BorderBrush" Value="#3f3f3f"/>
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
                <Setter Property="Background" Value="#3a3a3a"/>
              </Trigger>
              <Trigger Property="IsChecked" Value="True">
                <Setter Property="Background" Value="#0f6cbd"/>
                <Setter Property="BorderBrush" Value="#0f6cbd"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.6"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style x:Key="GroupHeaderText" TargetType="TextBlock">
      <Setter Property="FontSize" Value="16"/>
      <Setter Property="FontWeight" Value="Bold"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
      <Setter Property="Margin" Value="0,0,8,0"/>
    </Style>
  </Window.Resources>

  <DockPanel Margin="10">
    <!-- Top bar -->
    <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,10">
      <Button Name="BtnInstallSelected" Content="Install Selected" Width="140" Height="32" Margin="0,0,8,0"/>
      <Button Name="BtnUpgradeSelected" Content="Upgrade Selected" Width="150" Height="32" Margin="0,0,8,0"/>
      <Button Name="BtnClear" Content="Clear Selection" Width="140" Height="32" Margin="0,0,8,0"/>
      <Button Name="BtnGetInstalled" Content="Get Installed" Width="120" Height="32" Margin="0,0,8,0"/>
      <CheckBox Name="ChkForce" Content="Force" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <CheckBox Name="ChkSilent" IsChecked="True" Content="Silent" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <CheckBox Name="ChkAccept" IsChecked="True" Content="Accept EULA" VerticalAlignment="Center" Margin="0,0,8,0"/>
      <TextBox Name="TxtSearch" Width="240" Height="32" Margin="10,0,6,0" ToolTip="Search"/>
      <Button Name="BtnSearchClear" Content="X" Width="30" Height="30"/>
    </StackPanel>

    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="*"/>
        <RowDefinition Height="180"/>
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
        <TextBox Grid.Row="1" Name="TxtLog" Background="#181818" Foreground="White" IsReadOnly="True"
                 TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
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
$BtnUpgradeSelected = $window.FindName("BtnUpgradeSelected")
$BtnClear           = $window.FindName("BtnClear")
$BtnGetInstalled    = $window.FindName("BtnGetInstalled")
$TxtLog             = $window.FindName("TxtLog")
$ChkSilent          = $window.FindName("ChkSilent")
$ChkAccept          = $window.FindName("ChkAccept")
$ChkForce           = $window.FindName("ChkForce")
$TxtSearch          = $window.FindName("TxtSearch")
$BtnSearchClear     = $window.FindName("BtnSearchClear")

# ---- Helpers ----
function Is-Admin {
  try {
    $wi = [Security.Principal.WindowsIdentity]::GetCurrent()
    $wp = New-Object Security.Principal.WindowsPrincipal($wi)
    return $wp.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  } catch { return $false }
}

function Log-Msg([string]$msg){
  $TxtLog.AppendText(("{0}  {1}`r`n" -f (Get-Date).ToString("HH:mm:ss"), $msg))
  $TxtLog.ScrollToEnd()
}

if(-not (Is-Admin)){
  Log-Msg "[WARN] Nen chay PowerShell 'Run as Administrator' de winget cai dat khong bi chan."
}

# ---- Data: Apps & Groups ----
# Keys map -> $AppCatalog entries (ASCII/khong dau)
$AppCatalog = @(
  @{ Key="7zip";           Name="7zip";            Ids=@("7zip.7zip") },
  @{ Key="Chrome";         Name="Chrome";          Ids=@("Google.Chrome") },
  @{ Key="Notepad++";      Name="Notepad++";       Ids=@("Notepad++.Notepad++") },
  @{ Key="VS Code";        Name="VS Code";         Ids=@("Microsoft.VisualStudioCode") },
  @{ Key="PowerToys";      Name="PowerToys";       Ids=@("Microsoft.PowerToys") },
  @{ Key="PC Manager";     Name="PC Manager";      Ids=@("Microsoft.PCManager") },
  @{ Key="Rainmeter";      Name="Rainmeter";       Ids=@("Rainmeter.Rainmeter") },
  @{ Key="Zalo";           Name="Zalo";            Ids=@("VNG.ZaloPC","Zalo.Zalo","VNG.Zalo","VNGCorp.Zalo") },
  @{ Key="EVKey";          Name="EVKey";           Ids=@("tranxuanthang.EVKey","EVKey.EVKey","EVKey") },
  @{ Key="Office ODT";     Name="Office ODT";      Ids=@("Microsoft.OfficeDeploymentTool") },
  @{ Key="Creative Cloud"; Name="Creative Cloud";  Ids=@("Adobe.CreativeCloud","Adobe.Photoshop") },
  @{ Key="AutoCAD";        Name="AutoCAD";         Ids=@("Autodesk.AutoCAD","Autodesk.AutoCADLT") }
)

# Index for quick lookup by key
$AppByKey = @{}
foreach($a in $AppCatalog){ $AppByKey[$a.Key] = $a }

$Groups = @(
  @{ Title="Essentials";      Keys=@("7zip","Chrome","Notepad++","VS Code","PowerToys","PC Manager","Rainmeter") },
  @{ Title="VN Chat & Input"; Keys=@("Zalo","EVKey") },
  @{ Title="Office";          Keys=@("Office ODT") },
  @{ Title="Design & CAD";    Keys=@("Creative Cloud","AutoCAD") }
)

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
  if($ChkForce.IsChecked){ $args += "--force" }
  if($ChkSilent.IsChecked){ $args += "--silent" }
  if($ChkAccept.IsChecked){ $args += @("--accept-package-agreements","--accept-source-agreements") }

  Log-Msg ("Install: {0}" -f $id)
  $p = Start-Process -FilePath "winget" -ArgumentList $args -PassThru -WindowStyle Hidden
  $p.WaitForExit()
  $code = $p.ExitCode

  # 0: success ; -1978335189: already installed / update not applicable
  if($code -eq 0 -or $code -eq -1978335189){
    if($code -eq -1978335189){
      Log-Msg ("[OK] already installed / not applicable: {0}" -f $id)
    } else {
      Log-Msg ("[OK] installed: {0}" -f $id)
    }
    return $true
  } else {
    Log-Msg ("[WARN] install failed (ExitCode={0})" -f $code)
    return $false
  }
}

function Upgrade-ById([string]$id){
  if(-not $id){ return $false }
  $args = @("upgrade","-e","--id",$id)
  if($ChkSilent.IsChecked){ $args += "--silent" }
  if($ChkAccept.IsChecked){ $args += @("--accept-package-agreements","--accept-source-agreements") }

  Log-Msg ("Upgrade: {0}" -f $id)
  $p = Start-Process -FilePath "winget" -ArgumentList $args -PassThru -WindowStyle Hidden
  $p.WaitForExit()
  if($p.ExitCode -eq 0){
    Log-Msg ("[OK] upgraded or no update: {0}" -f $id)
    return $true
  } else {
    Log-Msg ("[WARN] upgrade failed (ExitCode={0})" -f $p.ExitCode)
    return $false
  }
}

# Dictionary for CheckBoxes and group containers
$CheckBoxes = @{}     # key -> CheckBox
$GroupPanels = @{}    # groupTitle -> WrapPanel

# Build group UI
foreach($g in $Groups){
  $gb = New-Object System.Windows.Controls.GroupBox
  $gb.Margin = "0,0,0,10"

  # Header: title + actions
  $header = New-Object System.Windows.Controls.DockPanel
  $t = New-Object System.Windows.Controls.TextBlock
  $t.Text = $g.Title
  $t.Style = $window.Resources["GroupHeaderText"]
  [System.Windows.Controls.DockPanel]::SetDock($t, "Left")
  $header.Children.Add($t) | Out-Null

  $hstack = New-Object System.Windows.Controls.StackPanel
  $hstack.Orientation = "Horizontal"
  $hstack.HorizontalAlignment = "Right"

  $btnAll = New-Object System.Windows.Controls.Button
  $btnAll.Content = "Select All"
  $btnAll.Margin = "4,0,0,0"
  $btnNone = New-Object System.Windows.Controls.Button
  $btnNone.Content = "None"
  $btnNone.Margin = "4,0,0,0"
  $btnInstallGroup = New-Object System.Windows.Controls.Button
  $btnInstallGroup.Content = "Install Group"
  $btnInstallGroup.Margin = "4,0,0,0"

  $hstack.Children.Add($btnAll) | Out-Null
  $hstack.Children.Add($btnNone) | Out-Null
  $hstack.Children.Add($btnInstallGroup) | Out-Null
  [System.Windows.Controls.DockPanel]::SetDock($hstack, "Right")
  $header.Children.Add($hstack) | Out-Null
  $gb.Header = $header

  $panel = New-Object System.Windows.Controls.WrapPanel
  $panel.Margin = "0,6,0,0"

  foreach($k in $g.Keys){
    $info = $AppByKey[$k]
    if(-not $info){ continue }
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.Style = $window.Resources["TileCheckBox"]
    $cb.Content = $info.Name
    $cb.Tag = $k
    $cb.Width = 180
    $cb.Height = 38
    $panel.Children.Add($cb) | Out-Null
    $CheckBoxes[$k] = $cb

    # Double-click: install single
    $cb.AddHandler([System.Windows.Controls.Control]::MouseDoubleClickEvent,
      [System.Windows.Input.MouseButtonEventHandler]{ param($s,$e)
        $key = $s.Tag
        $i2 = $AppByKey[$key]
        if($null -eq $i2){ return }
        $id = Resolve-Id -candidates $i2.Ids
        if($null -eq $id){ Log-Msg ("[ERR] not found on winget: {0}" -f ($i2.Ids -join " | ")); return }
        $s.IsEnabled = $false
        try { [void](Install-ById -id $id) } finally { $s.IsEnabled = $true }
      })
  }

  # Header buttons actions
  $btnAll.Add_Click({ foreach($k in $g.Keys){ if($CheckBoxes.ContainsKey($k)){ $CheckBoxes[$k].IsChecked = $true } } })
  $btnNone.Add_Click({ foreach($k in $g.Keys){ if($CheckBoxes.ContainsKey($k)){ $CheckBoxes[$k].IsChecked = $false } } })
  $btnInstallGroup.Add_Click({
    foreach($k in $g.Keys){
      if(-not $CheckBoxes.ContainsKey($k)){ continue }
      $info = $AppByKey[$k]
      $id = Resolve-Id -candidates $info.Ids
      if($null -eq $id){ Log-Msg ("[ERR] not found on winget: {0}" -f ($info.Ids -join " | ")); continue }
      $cb = $CheckBoxes[$k]; $cb.IsEnabled = $false
      try { [void](Install-ById -id $id) } finally { $cb.IsEnabled = $true }
    }
  })

  $gb.Content = $panel
  $PanelGroups.Children.Add($gb) | Out-Null
  $GroupPanels[$g.Title] = $panel
}

# Search filter
function Apply-Search {
  $q = ($TxtSearch.Text).ToString()
  $q = if([string]::IsNullOrWhiteSpace($q)) { "" } else { $q.Trim().ToLower() }
  foreach($kv in $CheckBoxes.GetEnumerator()){
    $cb = $kv.Value
    $name = ($cb.Content).ToString().ToLower()
    $cb.Visibility = if($q -eq "" -or $name -like "*$q*"){ "Visible" } else { "Collapsed" }
  }
}
$TxtSearch.Add_TextChanged({ Apply-Search })
$BtnSearchClear.Add_Click({ $TxtSearch.Text=""; Apply-Search })

# Top buttons
$BtnClear.Add_Click({
  foreach($cb in $CheckBoxes.Values){ $cb.IsChecked = $false }
  Log-Msg "Selection cleared."
})

$BtnGetInstalled.Add_Click({
  Log-Msg "winget list ..."
  $tmpOut = [System.IO.Path]::GetTempFileName()
  $p = Start-Process -FilePath "winget" -ArgumentList @("list") -PassThru -WindowStyle Hidden -RedirectStandardOutput $tmpOut
  $p.WaitForExit()
  try { Log-Msg (Get-Content -Raw $tmpOut) } catch { Log-Msg "[WARN] cannot read output." }
  Remove-Item -ErrorAction SilentlyContinue $tmpOut
})

$BtnInstallSelected.Add_Click({
  $sel = @()
  foreach($kv in $CheckBoxes.GetEnumerator()){ if($kv.Value.IsChecked){ $sel += $kv.Key } }
  if($sel.Count -eq 0){ Log-Msg "Chua chon ung dung nao."; return }
  Log-Msg ("Installing {0} item(s)..." -f $sel.Count)
  foreach($k in $sel){
    $info = $AppByKey[$k]
    $id = Resolve-Id -candidates $info.Ids
    if($null -eq $id){ Log-Msg ("[ERR] not found on winget: {0}" -f ($info.Ids -join " | ")); continue }
    $cb = $CheckBoxes[$k]; $cb.IsEnabled = $false
    try { [void](Install-ById -id $id) } finally { $cb.IsEnabled = $true }
  }
  Log-Msg "Done."
})

$BtnUpgradeSelected.Add_Click({
  $sel = @()
  foreach($kv in $CheckBoxes.GetEnumerator()){ if($kv.Value.IsChecked){ $sel += $kv.Key } }
  if($sel.Count -eq 0){ Log-Msg "Chua chon ung dung nao."; return }
  Log-Msg ("Upgrading {0} item(s)..." -f $sel.Count)
  foreach($k in $sel){
    $info = $AppByKey[$k]
    $id = Resolve-Id -candidates $info.Ids
    if($null -eq $id){ Log-Msg ("[ERR] not found on winget: {0}" -f ($info.Ids -join " | ")); continue }
    $cb = $CheckBoxes[$k]; $cb.IsEnabled = $false
    try { [void](Upgrade-ById -id $id) } finally { $cb.IsEnabled = $true }
  }
  Log-Msg "Done."
})

# Show UI
$window.ShowDialog() | Out-Null
