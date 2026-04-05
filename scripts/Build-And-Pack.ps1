#Requires -Version 5.1
<#
.SYNOPSIS
  Единый скрипт: сборка Win32 и Win64 (Release) и упаковка dist\visualRecognitionAddIn.zip.

.DESCRIPTION
  1) cmake --preset win32-release + сборка Release  
  2) cmake --preset x64-release + сборка Release  
  3) ZIP: MANIFEST.xml, visualAddInWin32.dll, visualAddInWin64.dll (плоская структура).

  Требуется: CMake в PATH, Visual Studio 2022 (CMakePresets.json).

.EXAMPLE
  .\Build-And-Pack.ps1
.EXAMPLE
  .\Build-And-Pack.ps1 -OutDir D:\artifacts
.EXAMPLE
  .\Build-And-Pack.ps1 -NoFresh
#>
[CmdletBinding()]
param(
    [string] $OutDir,
    [switch] $NoFresh
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$ResolvedOutDir = if ($OutDir) { $OutDir } else { (Join-Path $ProjectRoot 'dist') }
$ZipPath = Join-Path $ResolvedOutDir 'visualRecognitionAddIn.zip'
$ManifestPath = Join-Path $ProjectRoot 'MANIFEST.xml'
$Win32BuildDir = Join-Path $ProjectRoot 'build\win32-release'
$X64BuildDir = Join-Path $ProjectRoot 'build\x64-release'

function Invoke-CMakeConfigure {
    param(
        [Parameter(Mandatory)][string] $PresetName
    )
    $cmakeArgs = @('--preset', $PresetName)
    if (-not $NoFresh) {
        $cmakeArgs += '--fresh'
    }
    & cmake @cmakeArgs
}

function Find-ReleaseDll {
    param(
        [Parameter(Mandatory)][string]$BuildDir,
        [Parameter(Mandatory)][string]$FileName
    )
    if (-not (Test-Path -LiteralPath $BuildDir)) {
        return $null
    }
    $direct = Join-Path $BuildDir "Release\$FileName"
    if (Test-Path -LiteralPath $direct) {
        return (Get-Item -LiteralPath $direct)
    }
    $binRelease = Join-Path $BuildDir "bin\Release\$FileName"
    if (Test-Path -LiteralPath $binRelease) {
        return (Get-Item -LiteralPath $binRelease)
    }
    Get-ChildItem -LiteralPath $BuildDir -Recurse -Filter $FileName -File -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match '[\\/]Release[\\/]' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

if (-not (Test-Path -LiteralPath $ManifestPath)) {
    throw "MANIFEST.xml не найден: $ManifestPath"
}

Push-Location $ProjectRoot
try {
    Invoke-CMakeConfigure -PresetName 'win32-release'
    if ($LASTEXITCODE -ne 0) { throw "cmake --preset win32-release failed: exit $LASTEXITCODE" }

    cmake --build --preset win32-release --config Release
    if ($LASTEXITCODE -ne 0) { throw "cmake --build win32-release Release failed: exit $LASTEXITCODE" }

    Invoke-CMakeConfigure -PresetName 'x64-release'
    if ($LASTEXITCODE -ne 0) { throw "cmake --preset x64-release failed: exit $LASTEXITCODE" }

    cmake --build --preset x64-release --config Release
    if ($LASTEXITCODE -ne 0) { throw "cmake --build x64-release Release failed: exit $LASTEXITCODE" }
}
finally {
    Pop-Location
}

$dll32 = Find-ReleaseDll -BuildDir $Win32BuildDir -FileName 'visualAddInWin32.dll'
$dll64 = Find-ReleaseDll -BuildDir $X64BuildDir -FileName 'visualAddInWin64.dll'

if (-not $dll32) {
    throw "visualAddInWin32.dll не найден под $Win32BuildDir после сборки Release."
}
if (-not $dll64) {
    throw "visualAddInWin64.dll не найден под $X64BuildDir после сборки Release."
}

$staging = Join-Path ([System.IO.Path]::GetTempPath()) ("visualRecognitionPack_" + [Guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $staging | Out-Null
try {
    if (-not (Test-Path -LiteralPath $ResolvedOutDir)) {
        New-Item -ItemType Directory -Path $ResolvedOutDir -Force | Out-Null
    }

    Copy-Item -LiteralPath $ManifestPath -Destination (Join-Path $staging 'MANIFEST.xml')
    Copy-Item -LiteralPath $dll32.FullName -Destination (Join-Path $staging 'visualAddInWin32.dll')
    Copy-Item -LiteralPath $dll64.FullName -Destination (Join-Path $staging 'visualAddInWin64.dll')

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    if (Test-Path -LiteralPath $ZipPath) {
        Remove-Item -LiteralPath $ZipPath -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $ZipPath, [System.IO.Compression.CompressionLevel]::Optimal, $false)
}
finally {
    Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host "Создан архив: $ZipPath"
Write-Host ("  Win32: {0} ({1:u})" -f $dll32.FullName, $dll32.LastWriteTimeUtc)
Write-Host ("  x64:   {0} ({1:u})" -f $dll64.FullName, $dll64.LastWriteTimeUtc)
