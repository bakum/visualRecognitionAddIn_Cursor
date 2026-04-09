#Requires -Version 5.1
<#
.SYNOPSIS
  Единый скрипт: сборка Win32 и Win64 (Release) и упаковка dist\visualRecognitionAddIn.zip.

.DESCRIPTION
  1) cmake --preset win32-release + сборка Release  
  2) cmake --preset x64-release + сборка Release  
  3) ZIP: MANIFEST.xml + Windows DLL; Linux .so добавляются автоматически, если найдены.

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
    [switch] $NoFresh,
    [string] $LinuxLin64So,
    [string] $LinuxLin32So
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Ensure Unicode (Cyrillic) output renders correctly in Windows terminals.
[Console]::InputEncoding = [System.Text.UTF8Encoding]::new()
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
$OutputEncoding = [Console]::OutputEncoding

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$ResolvedOutDir = if ($OutDir) { $OutDir } else { (Join-Path $ProjectRoot 'dist') }
$ZipPath = Join-Path $ResolvedOutDir 'visualRecognitionAddIn.zip'
$ManifestPath = Join-Path $ProjectRoot 'MANIFEST.xml'
$Win32BuildDir = Join-Path $ProjectRoot 'build\win32-release'
$X64BuildDir = Join-Path $ProjectRoot 'build\x64-release'

function Resolve-BoostRoot {
    $candidates = @()
    if ($env:BOOST_ROOT) { $candidates += $env:BOOST_ROOT }
    $candidates += @(
        'E:\boost\boost_1_84_0',
        'E:\boost\boost_1_85_0',
        'C:\boost\boost_1_84_0',
        'C:\local\boost_1_84_0'
    )
    foreach ($c in $candidates) {
        if (-not $c) { continue }
        $hdr = Join-Path $c 'boost\json.hpp'
        if (Test-Path -LiteralPath $hdr) {
            return (Resolve-Path -LiteralPath $c).Path
        }
    }
    throw "Boost не найден. Установите BOOST_ROOT (папка, содержащая boost\json.hpp), например E:\boost\boost_1_84_0."
}

$BoostRoot = Resolve-BoostRoot

function Invoke-CMakeConfigure {
    param(
        [Parameter(Mandatory)][string] $PresetName
    )
    $cmakeArgs = @('--preset', $PresetName)
    if (-not $NoFresh) {
        $cmakeArgs += '--fresh'
    }
    $cmakeArgs += @('-DBOOST_ROOT=' + $BoostRoot)
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

function Write-ManifestCopy {
    param(
        [Parameter(Mandatory)][string] $DestinationPath,
        [Parameter(Mandatory)][bool] $IncludeLinux32,
        [Parameter(Mandatory)][bool] $IncludeLinux64
    )
    [xml]$doc = Get-Content -LiteralPath $ManifestPath -Encoding UTF8
    $ns = 'http://v8.1c.ru/8.2/addin/bundle'
    $mgr = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $mgr.AddNamespace('b', $ns)
    $bundle = $doc.SelectSingleNode('/b:bundle', $mgr)
    if (-not $bundle) {
        throw "Некорректный MANIFEST.xml: отсутствует корневой узел bundle."
    }

    $lin64 = $doc.SelectSingleNode("//b:component[@os='Linux' and @arch='x86_64']", $mgr)
    if ($IncludeLinux64) {
        if (-not $lin64) {
            $node = $doc.CreateElement('component', $ns)
            [void]$node.SetAttribute('os', 'Linux')
            [void]$node.SetAttribute('path', 'visualAddInLin64.so')
            [void]$node.SetAttribute('type', 'native')
            [void]$node.SetAttribute('arch', 'x86_64')
            [void]$bundle.AppendChild($node)
        } else {
            [void]$lin64.SetAttribute('path', 'visualAddInLin64.so')
            [void]$lin64.SetAttribute('type', 'native')
        }
    } elseif ($lin64) {
        [void]$lin64.ParentNode.RemoveChild($lin64)
    }

    $lin32 = $doc.SelectSingleNode("//b:component[@os='Linux' and @arch='i386']", $mgr)
    if ($IncludeLinux32) {
        if (-not $lin32) {
            $node = $doc.CreateElement('component', $ns)
            [void]$node.SetAttribute('os', 'Linux')
            [void]$node.SetAttribute('path', 'visualAddInLin32.so')
            [void]$node.SetAttribute('type', 'native')
            [void]$node.SetAttribute('arch', 'i386')
            [void]$bundle.AppendChild($node)
        } else {
            [void]$lin32.SetAttribute('path', 'visualAddInLin32.so')
            [void]$lin32.SetAttribute('type', 'native')
        }
    } elseif ($lin32) {
        [void]$lin32.ParentNode.RemoveChild($lin32)
    }

    $doc.Save($DestinationPath)
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

$autoLin64 = Join-Path $ProjectRoot 'build\linux-release\visualAddInLin64.so'
$autoLin32 = Join-Path $ProjectRoot 'build\linux32-release\visualAddInLin32.so'
$so64Path = if ($LinuxLin64So) { $LinuxLin64So } elseif (Test-Path -LiteralPath $autoLin64) { $autoLin64 } else { $null }
$so32Path = if ($LinuxLin32So) { $LinuxLin32So } elseif (Test-Path -LiteralPath $autoLin32) { $autoLin32 } else { $null }

$haveLin64 = $null -ne $so64Path -and (Test-Path -LiteralPath $so64Path)
$haveLin32 = $null -ne $so32Path -and (Test-Path -LiteralPath $so32Path)
if ($LinuxLin64So -and -not (Test-Path -LiteralPath $LinuxLin64So)) {
    throw "-LinuxLin64So не найден: $LinuxLin64So"
}
if ($LinuxLin32So -and -not (Test-Path -LiteralPath $LinuxLin32So)) {
    throw "-LinuxLin32So не найден: $LinuxLin32So"
}
if (-not $haveLin64 -and -not $haveLin32) {
    Write-Warning "Linux .so не найдены; в MANIFEST внутри zip останутся только Windows-компоненты."
}

$staging = Join-Path ([System.IO.Path]::GetTempPath()) ("visualRecognitionPack_" + [Guid]::NewGuid().ToString('n'))
New-Item -ItemType Directory -Path $staging | Out-Null
try {
    if (-not (Test-Path -LiteralPath $ResolvedOutDir)) {
        New-Item -ItemType Directory -Path $ResolvedOutDir -Force | Out-Null
    }

    $stagingManifest = Join-Path $staging 'MANIFEST.xml'
    Write-ManifestCopy -DestinationPath $stagingManifest -IncludeLinux32:$haveLin32 -IncludeLinux64:$haveLin64
    Copy-Item -LiteralPath $dll32.FullName -Destination (Join-Path $staging 'visualAddInWin32.dll')
    Copy-Item -LiteralPath $dll64.FullName -Destination (Join-Path $staging 'visualAddInWin64.dll')
    if ($haveLin64) {
        Copy-Item -LiteralPath $so64Path -Destination (Join-Path $staging 'visualAddInLin64.so')
    }
    if ($haveLin32) {
        Copy-Item -LiteralPath $so32Path -Destination (Join-Path $staging 'visualAddInLin32.so')
    }

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
if ($haveLin64) { Write-Host ("  Lin64: {0}" -f $so64Path) }
if ($haveLin32) { Write-Host ("  Lin32: {0}" -f $so32Path) }
