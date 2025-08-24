# AppPack.Index.PS5.ps1 - PowerShell 5.1 compatible (no ternary)
[CmdletBinding()]
param(
  [int]$Index,
  [switch]$Upgrade,
  [switch]$Preview
)

function W($m){ Write-Host $m }
function Info($m){ Write-Host "[*] $m" }
function Ok($m){ Write-Host "[OK] $m" -ForegroundColor Green }
function Err($m){ Write-Host "[ERR] $m" -ForegroundColor Red }

# Check winget
if(-not (Get-Command winget -ErrorAction SilentlyContinue)){
  Err "winget not found. Please install/update 'App Installer' from Microsoft Store."
  exit 1
}

function Resolve-Id([string[]]$Candidates){
  foreach($c in $Candidates){
    $p = Start-Process -FilePath "winget" -ArgumentList @("show","-e","--id",$c) -PassThru -WindowStyle Hidden
    $p.WaitForExit()
    if($p.ExitCode -eq 0){ return $c }
  }
  return $null
}

function Install-Any($entry){
  $id = $null
  if($entry -is [System.Array]){
    $id = Resolve-Id -Candidates $entry
    if(-not $id){
      Err ("Package not found on winget: {0}" -f (($entry -join " | ")))
      return $false
    }
  } else {
    $id = [string]$entry
  }

  Info ("Install: {0}" -f $id)
  $p = Start-Process -FilePath "winget" -ArgumentList @(
    "install","-e","--id",$id,"--silent",
    "--accept-package-agreements","--accept-source-agreements"
  ) -PassThru -WindowStyle Hidden
  $p.WaitForExit()

  # Treat -1978335189 (already installed / not applicable) as success
  if($p.ExitCode -eq 0 -or $p.ExitCode -eq -1978335189){
    if($p.ExitCode -eq -1978335189){
      Ok ("Already installed / not applicable: {0}" -f $id)
    } else {
      Ok ("Installed: {0}" -f $id)
    }
    return $true
  } else {
    W "[WARN] install failed (ExitCode=$($p.ExitCode))"
    return $false
  }
}

$Packs = @(
  @{
    Index = 1; Name = "Essentials";
    Note  = "7zip, Chrome, Notepad++, VS Code, PowerToys, PC Manager, Rainmeter";
    Ids   = @(
      "7zip.7zip",
      "Google.Chrome",
      "Notepad++.Notepad++",
      "Microsoft.VisualStudioCode",
      "Microsoft.PowerToys",
      "Microsoft.PCManager",
      "Rainmeter.Rainmeter"
    )
  },
  @{
    Index = 2; Name = "VN Chat & Input";
    Note  = "Zalo, EVKey";
    Ids   = @(
      @("VNG.ZaloPC","Zalo.Zalo","VNG.Zalo","VNGCorp.Zalo"),
      @("tranxuanthang.EVKey","EVKey.EVKey","EVKey")
    )
  },
  @{
    Index = 3; Name = "Office";
    Note  = "Office Deployment Tool (ODT) + ONLYOFFICE (free)";
    Ids   = @(
      "Microsoft.OfficeDeploymentTool",
      "ONLYOFFICE.DesktopEditors"
    )
  },
  @{
    Index = 4; Name = "Design & CAD";
    Note  = "Adobe Creative Cloud (Photoshop) + Autodesk AutoCAD";
    Ids   = @(
      @("Adobe.CreativeCloud","Adobe.Photoshop"),
      @("Autodesk.AutoCAD","Autodesk.AutoCADLT")
    )
  },
  @{
    Index = 5; Name = "ALL (1+2+3+4)";
    Note  = "Combined packs";
    Ids   = $null
  }
)

# Compose ALL pack
$all = $Packs | Where-Object { $_.Index -eq 5 }
if($all){
  $ids = @()
  foreach($i in 1..4){
    $ids += ($Packs | Where-Object { $_.Index -eq $i }).Ids
  }
  $all.Ids = $ids
}

function Show-Menu {
  W ""
  W "=== AppPack by Index (PS5 compatible) ==="
  foreach($p in $Packs){
    W ("[{0}] {1}  - {2}" -f $p.Index, $p.Name, $p.Note)
  }
  W "[0] Exit"
  W ""
}

$choice = $Index
if(-not $choice){
  Show-Menu
  $choice = Read-Host "Enter index to install"
}

if([string]::IsNullOrWhiteSpace("$choice") -or $choice -eq 0){
  Info "Cancelled."
  exit 0
}

$target = $Packs | Where-Object { $_.Index -eq [int]$choice }
if(-not $target){
  Err "Invalid index: $choice"
  exit 1
}

Info ("Selected: [{0}] {1} - {2}" -f $target.Index, $target.Name, $target.Note)
if($Preview){
  W "Packages:"
  foreach($e in $target.Ids){
    if($e -is [System.Array]){ W ("  - " + ($e -join " | ")) } else { W ("  - " + $e) }
  }
  exit 0
}

if($Upgrade){
  Info "Upgrading available packages first..."
  & winget upgrade --all --silent --accept-package-agreements --accept-source-agreements | Out-Host
}

$fail = @()
foreach($e in $target.Ids){
  $ok = Install-Any -entry $e
  if(-not $ok){
    if($e -is [System.Array]){
      $fail += ($e -join " | ")
    } else {
      $fail += $e
    }
  }
}

if($fail.Count -gt 0){
  Err ("Some packages failed: {0}" -f ($fail -join ", "))
  W "Notes:"
  W " - Photoshop: install via Adobe Creative Cloud app if direct ID not available."
  W " - AutoCAD: may require Autodesk sign-in/subscription; installer may prompt."
  W " - EVKey: if not found on winget, download from project homepage."
  exit 2
} else {
  Ok "All done."
}
