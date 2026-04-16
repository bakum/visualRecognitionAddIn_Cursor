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

.PARAMETER NoDeps
  Skip apt-get update/install dependency steps. Use for repeat builds with a pre-warmed image.
#>
[CmdletBinding()]
param(
    [ValidateSet('x64', 'i386')]
    [string] $Arch = 'x64',
    [string] $UbuntuImage,
    [switch] $NoClean,
    [string] $OutDir,
    [switch] $NoDeps
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
$steps.Add('set -euo pipefail')
$steps.Add('echo "[docker-build] Step 1/6: Preparing build directory"')
if (-not $NoClean) {
    $steps.Add("rm -rf $dockerPosix")
}
if ($NoDeps) {
    $steps.Add('echo "[docker-build] Step 2/6: Skipping dependency install (-NoDeps)"')
    $steps.Add('echo "[docker-build] Verify toolchain in image: cmake and ninja"')
    $steps.Add('command -v cmake >/dev/null || { echo "[docker-build] ERROR: cmake not found in image. Re-run without -NoDeps or use pre-warmed image."; exit 1; }')
    $steps.Add('command -v ninja >/dev/null || command -v ninja-build >/dev/null || { echo "[docker-build] ERROR: ninja not found in image. Re-run without -NoDeps or use pre-warmed image."; exit 1; }')
    $steps.Add('echo "[docker-build] Step 4/6: Preparing Boost root (NoDeps mode)"')
    $steps.Add('test -f /usr/include/boost/json.hpp || { echo "[docker-build] ERROR: boost/json.hpp not found in image. Re-run without -NoDeps or use pre-warmed image."; exit 1; }; export BOOST_ROOT=/usr/include')
} else {
    $steps.Add('echo "[docker-build] Step 2/6: apt-get update (with retries)"')
    $steps.Add('export DEBIAN_FRONTEND=noninteractive')
    $steps.Add('APT_UPDATE_OK=0; for i in 1 2 3; do rm -rf /var/lib/apt/lists/*; apt-get clean; if apt-get update; then APT_UPDATE_OK=1; break; fi; if [ "$i" -eq 2 ] && [ -f /etc/apt/sources.list ]; then echo "[docker-build] Switching Ubuntu mirror to mirrors.edge.kernel.org/ubuntu"; sed -i "s|http://archive.ubuntu.com/ubuntu|http://mirrors.edge.kernel.org/ubuntu|g; s|http://security.ubuntu.com/ubuntu|http://mirrors.edge.kernel.org/ubuntu|g" /etc/apt/sources.list; fi; echo "[docker-build] apt-get update failed (attempt $i/3), retrying..."; sleep 5; done; [ "$APT_UPDATE_OK" -eq 1 ] || { echo "[docker-build] ERROR: apt-get update failed after retries"; exit 1; }')
    $steps.Add('echo "[docker-build] Step 3/6: apt-get install toolchain (with retries)"')
    $steps.Add('APT_INSTALL_OK=0; for i in 1 2 3; do if apt-get install -y --fix-missing build-essential cmake ninja-build libuuid1 uuid-dev libssl-dev libboost-dev ca-certificates; then APT_INSTALL_OK=1; break; fi; echo "[docker-build] apt-get install failed (attempt $i/3), retrying..."; sleep 5; done; [ "$APT_INSTALL_OK" -eq 1 ] || { echo "[docker-build] ERROR: apt-get install failed after retries"; exit 1; }')
    $steps.Add('echo "[docker-build] Verify toolchain: cmake and ninja"')
    $steps.Add('command -v cmake >/dev/null || { echo "[docker-build] ERROR: cmake not found after apt-get install"; exit 1; }')
    $steps.Add('command -v ninja >/dev/null || command -v ninja-build >/dev/null || { echo "[docker-build] ERROR: ninja not found after apt-get install"; exit 1; }')
    $steps.Add('echo "[docker-build] Step 4/6: Preparing Boost root"')
    $steps.Add('BOOST_ROOT_PATH=/usr/include; if [ ! -f /usr/include/boost/json.hpp ]; then apt-get install -y wget; test -f /tmp/boost_1_84_0.tar.gz || wget -O /tmp/boost_1_84_0.tar.gz https://archives.boost.io/release/1.84.0/source/boost_1_84_0.tar.gz; mkdir -p /opt; rm -rf /opt/boost_1_84_0; tar -xzf /tmp/boost_1_84_0.tar.gz -C /opt; BOOST_ROOT_PATH=/opt/boost_1_84_0; fi; export BOOST_ROOT=$BOOST_ROOT_PATH')
}
if ($is32) {
    $steps.Add('echo "[docker-build] Step 5/6: Configuring CMake"')
    $steps.Add("cmake -S . -B $dockerPosix -G Ninja -DCMAKE_BUILD_TYPE=Release -DBOOST_ROOT=`$BOOST_ROOT -DCMAKE_CXX_FLAGS='-m32'")
} else {
    $steps.Add('echo "[docker-build] Step 5/6: Configuring CMake"')
    $steps.Add("cmake -S . -B $dockerPosix -G Ninja -DCMAKE_BUILD_TYPE=Release -DBOOST_ROOT=`$BOOST_ROOT")
}
$steps.Add('echo "[docker-build] Step 6/6: Building with CMake"')
$steps.Add("cmake --build $dockerPosix --parallel `$(nproc)")
$steps.Add("test -f $($candidateSos[0]) || test -f $($candidateSos[1]) || test -f $($candidateSos[2]) || test -f $($candidateSos[3]) || test -f $($candidateSos[4]) || test -f $($candidateSos[5])")

$bash = $steps -join ' && '
$bashBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($bash))

Write-Host "Docker image: $UbuntuImage"
Write-Host "Arch:         $Arch"
Write-Host "Preset note:  $linuxPreset (used for output naming only)"
Write-Host "Project:      $ProjectRoot"
Write-Host "Build dir:    $(Join-Path $ProjectRoot $dockerRel)"

& docker run --rm `
    -v "${ProjectRoot}:/work:rw" `
    -w /work `
    $UbuntuImage `
    bash -lc "echo '$bashBase64' | base64 -d | bash"

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
