#Requires -Version 5.1
<#
.SYNOPSIS
  Build VisualRecognitionAddIn .so in Docker (x64 or i386) and copy artifact under build/.

.DESCRIPTION
  Runs Linux build inside container (CMake + Ninja), then copies resulting shared library to:

    x64:  <repo>\build\linux-release\visualAddInLin64.so
    i386: <repo>\build\linux32-release\visualAddInLin32.so

  Requires Docker Desktop (Linux containers).

.PARAMETER Arch
  x64 (default) or i386.

.PARAMETER UbuntuImage
  Override base image. Defaults:
    x64  -> ubuntu:24.04
    i386 -> i386/debian:bookworm-slim

.PARAMETER NoClean
  Do not delete docker build directory before configure.

.PARAMETER OutDir
  Destination directory for copied .so (defaults: build\linux-release or build\linux32-release).
#>
[CmdletBinding()]
param(
    [ValidateSet('x64', 'i386')]
    [string] $Arch = 'x64',
    [string] $UbuntuImage,
    [switch] $NoClean,
    [string] $OutDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$is32 = ($Arch -eq 'i386')

if (-not $UbuntuImage) {
    $UbuntuImage = if ($is32) { 'i386/debian:bookworm-slim' } else { 'ubuntu:24.04' }
}

if (-not $OutDir) {
    $OutDir = if ($is32) {
        Join-Path $ProjectRoot 'build\linux32-release'
    } else {
        Join-Path $ProjectRoot 'build\linux-release'
    }
}

$dockerRel = if ($is32) { 'build\docker-linux32' } else { 'build\docker-linux' }
$dockerPosix = if ($is32) { 'build/docker-linux32' } else { 'build/docker-linux' }
$copiedSoName = if ($is32) { 'visualAddInLin32.so' } else { 'visualAddInLin64.so' }
$linuxPreset = if ($is32) { 'linux32-release' } else { 'linux-release' }

# Expected target names from CMake (Linux output name set in CMakeLists.txt).
$candidateSos = @(
    "$dockerPosix/bin/Release/$copiedSoName",
    "$dockerPosix/bin/$copiedSoName",
    "$dockerPosix/$copiedSoName",
    "$dockerPosix/bin/Release/libVisualRecognitionAddIn.so",
    "$dockerPosix/bin/libVisualRecognitionAddIn.so",
    "$dockerPosix/libVisualRecognitionAddIn.so"
)

$steps = [System.Collections.Generic.List[string]]::new()
if (-not $NoClean) {
    $steps.Add("rm -rf $dockerPosix")
}
$steps.Add('export DEBIAN_FRONTEND=noninteractive')
$steps.Add('apt-get update -qq')
$steps.Add('apt-get install -y -qq build-essential cmake ninja-build libuuid1 uuid-dev libssl-dev libboost-dev ca-certificates >/dev/null')
$steps.Add('BOOST_ROOT_PATH=/usr/include; if [ ! -f /usr/include/boost/json.hpp ]; then apt-get install -y -qq wget >/dev/null; test -f /tmp/boost_1_84_0.tar.gz || wget -q -O /tmp/boost_1_84_0.tar.gz https://archives.boost.io/release/1.84.0/source/boost_1_84_0.tar.gz; mkdir -p /opt; rm -rf /opt/boost_1_84_0; tar -xzf /tmp/boost_1_84_0.tar.gz -C /opt; BOOST_ROOT_PATH=/opt/boost_1_84_0; fi; export BOOST_ROOT=$BOOST_ROOT_PATH')
if ($is32) {
    $steps.Add("cmake -S . -B $dockerPosix -G Ninja -DCMAKE_BUILD_TYPE=Release -DBOOST_ROOT=`$BOOST_ROOT -DCMAKE_CXX_FLAGS='-m32'")
} else {
    $steps.Add("cmake -S . -B $dockerPosix -G Ninja -DCMAKE_BUILD_TYPE=Release -DBOOST_ROOT=`$BOOST_ROOT")
}
$steps.Add("cmake --build $dockerPosix --parallel `$(nproc)")
$steps.Add("test -f $($candidateSos[0]) || test -f $($candidateSos[1]) || test -f $($candidateSos[2]) || test -f $($candidateSos[3]) || test -f $($candidateSos[4]) || test -f $($candidateSos[5])")

$bash = $steps -join ' && '

Write-Host "Docker image: $UbuntuImage"
Write-Host "Arch:         $Arch"
Write-Host "Preset note:  $linuxPreset (used for output naming only)"
Write-Host "Project:      $ProjectRoot"
Write-Host "Build dir:    $(Join-Path $ProjectRoot $dockerRel)"

& docker run --rm `
    -v "${ProjectRoot}:/work:rw" `
    -w /work `
    $UbuntuImage `
    bash -lc $bash

if ($LASTEXITCODE -ne 0) {
    throw "docker build failed with exit code $LASTEXITCODE"
}

$builtCandidates = foreach ($candidate in $candidateSos) {
    [System.IO.Path]::Combine($ProjectRoot, ($candidate -replace '/', '\'))
}
$builtSo = $builtCandidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
if (-not $builtSo) {
    throw "Expected artifact not found in docker build output. Checked: $($builtCandidates -join ', ')"
}

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$dest = Join-Path $OutDir $copiedSoName
Copy-Item -LiteralPath $builtSo -Destination $dest -Force

Write-Host "Copied: $dest"
