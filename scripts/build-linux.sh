#!/usr/bin/env bash
# Build visualRecognition add-in .so on Linux (or WSL). Requires: cmake, ninja, g++, OpenSSL dev package, Boost headers.
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
BUILD_DIR="${1:-$ROOT/build/linux-release}"
cmake -S "$ROOT" -B "$BUILD_DIR" -G Ninja -DCMAKE_BUILD_TYPE=Release "${@:2}"
cmake --build "$BUILD_DIR" --parallel "$(nproc 2>/dev/null || sysctl -n hw.ncpu 2>/dev/null || echo 4)"
echo "Build finished. Artifact is under: $BUILD_DIR/bin (shared library for Linux target)."
