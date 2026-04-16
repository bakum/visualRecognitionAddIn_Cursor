#Requires -Version 5.1
<#
.SYNOPSIS
  Build pre-warmed Docker images for fast Linux add-in builds.

.DESCRIPTION
  Pre-installs build dependencies used by docker-build-linux.ps1 so repeated
  builds can run with -NoDeps for both x64 and i386.

.PARAMETER Arch
  x64, i386, or all (default).

.PARAMETER X64BaseImage
  Base image for x64 warm image (default: ubuntu:24.04).

.PARAMETER I386BaseImage
  Base image for i386 warm image (default: i386/debian:bookworm-slim).

.PARAMETER X64Tag
  Output tag for warmed x64 image (default: visualaddin-build:ubuntu24-x64).

.PARAMETER I386Tag
  Output tag for warmed i386 image (default: visualaddin-build:debian12-i386).
#>
[CmdletBinding()]
param(
    [ValidateSet('all', 'x64', 'i386')]
    [string] $Arch = 'all',
    [string] $X64BaseImage = 'ubuntu:24.04',
    [string] $I386BaseImage = 'i386/debian:bookworm-slim',
    [string] $X64Tag = 'visualaddin-build:ubuntu24-x64',
    [string] $I386Tag = 'visualaddin-build:debian12-i386'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Build-WarmImage {
    param(
        [Parameter(Mandatory = $true)][string] $BaseImage,
        [Parameter(Mandatory = $true)][string] $OutTag,
        [Parameter(Mandatory = $true)][string] $ArchLabel
    )

    Write-Host "[warm] Building $ArchLabel image: $OutTag (base: $BaseImage)"

    $cmd = @'
set -euo pipefail
export DEBIAN_FRONTEND=noninteractive
rm -rf /var/lib/apt/lists/*
apt-get clean
apt-get update
apt-get install -y --fix-missing build-essential cmake ninja-build libuuid1 uuid-dev libssl-dev libboost-dev ca-certificates wget
if [ ! -f /usr/include/boost/json.hpp ]; then
  test -f /tmp/boost_1_84_0.tar.gz || wget -O /tmp/boost_1_84_0.tar.gz https://archives.boost.io/release/1.84.0/source/boost_1_84_0.tar.gz
  mkdir -p /opt
  rm -rf /opt/boost_1_84_0
  tar -xzf /tmp/boost_1_84_0.tar.gz -C /opt
fi
cmake --version
(ninja --version || ninja-build --version)
'@

    $cmdBase64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($cmd))

    & docker rm -f visualaddin-warm-tmp 2>$null | Out-Null
    & docker run --name visualaddin-warm-tmp `
        $BaseImage `
        bash -lc "echo '$cmdBase64' | base64 -d | bash"
    if ($LASTEXITCODE -ne 0) {
        & docker rm -f visualaddin-warm-tmp | Out-Null
        throw "[warm] Failed while preparing commit container for $ArchLabel."
    }

    & docker commit visualaddin-warm-tmp $OutTag | Out-Null
    if ($LASTEXITCODE -ne 0) {
        & docker rm -f visualaddin-warm-tmp | Out-Null
        throw "[warm] docker commit failed for $ArchLabel."
    }

    & docker rm -f visualaddin-warm-tmp | Out-Null
    Write-Host "[warm] Ready: $OutTag"
}

if ($Arch -in @('all', 'x64')) {
    Build-WarmImage -BaseImage $X64BaseImage -OutTag $X64Tag -ArchLabel 'x64'
}

if ($Arch -in @('all', 'i386')) {
    Build-WarmImage -BaseImage $I386BaseImage -OutTag $I386Tag -ArchLabel 'i386'
}

Write-Host '[warm] Done.'
