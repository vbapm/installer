# Uninstaller for vbapm on Windows
# Reverses the steps performed by install.ps1

$ErrorActionPreference = 'Stop'

$LibDir = "$env:APPDATA\vbapm"
$BinDir = "$LibDir\bin"
$Shortcut = "$env:AppData\Microsoft\Addins\vbapm Add-ins.lnk"

Write-Output "[1/4] Removing vbapm from PATH..."
$User = [EnvironmentVariableTarget]::User
$Path = [Environment]::GetEnvironmentVariable('Path', $User)
$NewPath = ($Path -split ';' | Where-Object { $_ -ne $BinDir }) -join ';'
if ($NewPath -ne $Path) {
  [Environment]::SetEnvironmentVariable('Path', $NewPath, $User)
  $Env:Path = ($Env:Path -split ';' | Where-Object { $_ -ne $BinDir }) -join ';'
  Write-Output "Removed $BinDir from PATH."
} else {
  Write-Output "$BinDir was not in PATH, skipping."
}

Write-Output "[2/3] Removing add-ins shortcut..."
if (Test-Path $Shortcut) {
  Remove-Item $Shortcut -Force
  Write-Output "Removed $Shortcut."
} else {
  Write-Output "Shortcut not found, skipping."
}

Write-Output "[3/3] Removing vbapm installation directory..."
if (Test-Path $LibDir) {
  Remove-Item $LibDir -Recurse -Force
  Write-Output "Removed $LibDir."
} else {
  Write-Output "$LibDir not found, skipping."
}

Write-Output ""
Write-Output "vbapm has been uninstalled."
