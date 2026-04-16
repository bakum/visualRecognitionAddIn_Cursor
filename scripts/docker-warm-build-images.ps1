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

apt_retry_update() {
  local ok=0
  for i in 1 2 3; do
    rm -rf /var/lib/apt/lists/*
    apt-get clean

    if apt-get update; then ok=1; break; fi
    if [ "$i" -eq 2 ] && [ -f /etc/apt/sources.list ]; then
      echo "[warm] Switching Ubuntu mirror to mirrors.edge.kernel.org/ubuntu"
      for f in /etc/apt/sources.list /etc/apt/sources.list.d/*.list /etc/apt/sources.list.d/*.sources; do
        [ -f "$f" ] || continue
        sed -i 's|http://archive.ubuntu.com/ubuntu|http://mirrors.edge.kernel.org/ubuntu|g; s|http://security.ubuntu.com/ubuntu|http://mirrors.edge.kernel.org/ubuntu|g; s|https://archive.ubuntu.com/ubuntu|http://mirrors.edge.kernel.org/ubuntu|g; s|https://security.ubuntu.com/ubuntu|http://mirrors.edge.kernel.org/ubuntu|g' "$f" || true
      done
    fi
    echo "[warm] apt-get update failed (attempt $i/3), retrying..."
    sleep 5
  done
  [ "$ok" -eq 1 ] || { echo "[warm] ERROR: apt-get update failed after retries"; exit 1; }
}

apt_retry_install() {
  local ok=0
  for i in 1 2 3; do
    if apt-get install -y --fix-missing build-essential cmake ninja-build libuuid1 uuid-dev libssl-dev libboost-dev ca-certificates wget; then ok=1; break; fi
    echo "[warm] apt-get install failed (attempt $i/3), retrying..."
    sleep 5
  done
  [ "$ok" -eq 1 ] || { echo "[warm] ERROR: apt-get install failed after retries"; exit 1; }
}

apt_retry_update
apt_retry_install

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

    try {
        & docker rm -f visualaddin-warm-tmp *> $null
    } catch {
        # ignore: container may not exist yet
    }
    & docker run --name visualaddin-warm-tmp `
        $BaseImage `
        bash -lc "echo '$cmdBase64' | base64 -d | tr -d '\r' | bash"
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
