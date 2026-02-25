# Based heavily on the approach used in https://github.com/denoland/deno_install
# Copyright 2018 the Deno authors. All rights reserved. MIT license.

# Installer for vbapm on Windows

$ErrorActionPreference = 'Stop'

if ($args.Length -gt 0) {
  $Version = $args.Get(0)
}

$LibDir = "$env:APPDATA\vbapm"
$BinDir ="$LibDir\bin"
$ZipFile = "$LibDir\vbapm.zip"
$AddinsDir = "$LibDir\addins\build"

# Create the lib directory if it doesn't exist
if (!(Test-Path $LibDir)) {
  New-Item $LibDir -ItemType Directory | Out-Null
}

# GitHub requires TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ReleaseUri = if (!$Version) {
	# Use the GitHub API to find the latest release asset URL.
	# The releases HTML page requires JavaScript to render asset links,
	# so we query the API which returns JSON instead.
	$Response = Invoke-RestMethod "https://api.github.com/repos/vbapm/core/releases/latest"
	$Response.assets |
		Where-Object { $_.name -eq "vbapm-win.zip" } |
		ForEach-Object { $_.browser_download_url } |
		Select-Object -First 1
} else {
	"https://github.com/vbapm/core/releases/download/$Version/vbapm-win.zip"
}

Write-Output "[1/5] Downloading vbapm..."
Write-Output "($ReleaseUri)"
Invoke-WebRequest $ReleaseUri -Out $ZipFile

Write-Output "[2/5] Extracting vbapm..."
Expand-Archive $ZipFile -Destination $LibDir -Force
Remove-Item $ZipFile

Write-Output "[3/5] Adding vbapm to PATH..."
$User = [EnvironmentVariableTarget]::User
$Path = [Environment]::GetEnvironmentVariable('Path', $User)
if (!(";$Path;".ToLower() -like "*;$BinDir;*".ToLower())) {
  [Environment]::SetEnvironmentVariable('Path', "$Path;$BinDir", $User)
}
$env:Path += ";$BinDir"
# We add the bin directory to GITHUB_PATH in case this script is running inside a GitHub Actions runner
if ($env:GITHUB_PATH) {
	$BinDir | Out-File -FilePath $env:GITHUB_PATH -Encoding utf8 -Append
}

function New-Shortcut ($Src, $Dest) {
  try {
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($Dest)
    $Shortcut.TargetPath = $Src
    $Shortcut.Save()
  } catch {
    Write-Output "Failed to link add-ins, they can instead be found in $AddinsDir."
  }
}

# The add-ins directory is where the browse button used to add a new addin will start from, so we create a shortcut to it on the desktop for easy access.
Write-Output "[4/5] Creating shortcut to add-ins..."
New-Shortcut "$AddinsDir" "$env:AppData\Microsoft\Addins\vbapm Add-ins.lnk"

function Enable-VBOM ($App) {
  try {
    $CurVer = Get-ItemProperty -Path Registry::HKEY_CLASSES_ROOT\$App.Application\CurVer -ErrorAction Stop
    $OfficeVersion = $CurVer.'(default)'.replace("$App.Application.", "") + ".0"

    Set-ItemProperty -Path HKCU:\Software\Microsoft\Office\$OfficeVersion\$App\Security -Name AccessVBOM -Value 1 -ErrorAction Stop
  } catch {
    Write-Output "Failed to enable access to VBA project object model for $App."
  }
}

Write-Output "[5/5] Enabling access to VBA project object model..."
Enable-VBOM "Excel"
# TODO Enable-VBOM "Word"
# TODO Enable-VBOM "PowerPoint"
# TODO Enable-VBOM "Access"

if (!(Test-Path (Join-Path $BinDir "vba.cmd"))) {
	throw "Expected vba.cmd in $BinDir, but it was not found."
}

Write-Output ""
Write-Output "Success! vbapm was installed successfully."
Write-Output ""
Write-Output "Command: \"$BinDir\vba\""
Write-Output "Add-ins: \"$AddinsDir\""
Write-Output ""
Write-Output "Run 'vba --help' to get started"
