$ErrorActionPreference = "Stop"

$RootPath = $PSScriptRoot
$BinPath = Join-Path $RootPath "bin"

Write-Host ""
Write-Host "Installing DevUtils-PS..."
Write-Host "Location: $RootPath"
Write-Host ""

if (!(Test-Path $BinPath)) {
    New-Item -ItemType Directory -Path $BinPath | Out-Null
}

$CmdPath = Join-Path $BinPath "devutils.cmd"

$CmdContent = @'
@echo off
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0..\devutils.ps1" %*
'@

Set-Content -Path $CmdPath -Value $CmdContent

Write-Host "Created launcher:"
Write-Host "  $CmdPath"

$CurrentPath = [Environment]::GetEnvironmentVariable(
    "Path",
    "User"
)

$PathEntries = $CurrentPath -split ";" |
    Where-Object { $_ -ne "" }

if ($PathEntries -notcontains $BinPath) {

    $NewPath = ($PathEntries + $BinPath) -join ";"

    [Environment]::SetEnvironmentVariable(
        "Path",
        $NewPath,
        "User"
    )

    Write-Host ""
    Write-Host "Added to user PATH:"
    Write-Host "  $BinPath"
}
else {
    Write-Host ""
    Write-Host "PATH already configured."
}

Write-Host ""
Write-Host "Installation complete."
Write-Host ""
Write-Host "Open a NEW terminal and run:"
Write-Host ""
Write-Host "  devutils help"
Write-Host ""