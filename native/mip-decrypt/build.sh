#!/usr/bin/env bash
# Builds the mip-decrypt sidecar against an extracted MIP SDK.
#
# Usage: MIP_SDK_ROOT=/path/to/mipsdk ./build.sh
#
# MIP_SDK_ROOT must hold the SDK's `include/` (mip/*.h headers) and `bin/`
# (release shared libraries) trees, as extracted from the bundles at
# https://aka.ms/mipsdkbins (see README.md). The binary lands in ./build/.
set -euo pipefail

cd "$(dirname "$0")"

if [[ -z "${MIP_SDK_ROOT:-}" ]]; then
  echo "error: set MIP_SDK_ROOT to the extracted MIP SDK directory" >&2
  exit 2
fi
if [[ ! -d "$MIP_SDK_ROOT/include/mip" ]]; then
  echo "error: $MIP_SDK_ROOT/include/mip not found; see README.md" >&2
  exit 2
fi

CXX="${CXX:-g++}"
BUILD_DIR="${BUILD_DIR:-build}"
mkdir -p "$BUILD_DIR"

OS="$(uname -s)"
ARCH="$(uname -m)"

# Library layout differs per platform but the link set is the same File SDK
# trio; rpath keeps the binary loadable from its build dir without installs.
case "$OS" in
  Darwin)
    LIB_DIR="$MIP_SDK_ROOT/bin/release/macos"
    RPATH_ARGS=(-Wl,-rpath,"@executable_path" -Wl,-rpath,"$LIB_DIR")
    ;;
  Linux)
    LIB_DIR="$MIP_SDK_ROOT/bin/release/generic"
    RPATH_ARGS=(-Wl,-rpath,'$ORIGIN' -Wl,-rpath,"$LIB_DIR")
    ;;
  *)
    echo "error: unsupported OS $OS (macOS and Linux only; see README.md)" >&2
    exit 2
    ;;
esac

if [[ ! -d "$LIB_DIR" ]]; then
  # Older bundles flatten the per-OS dirs; fall back to the top level.
  LIB_DIR="$MIP_SDK_ROOT/bin/release"
fi

# The ucs companion library ships on some platforms only (absent from the
# macOS bundle); link it only when the file is there, or the link fails.
LINK_LIBS=(-lmip_file_sdk -lmip_protection_sdk)
if [[ -f "$LIB_DIR/libmip_ucs_sdk.dylib" || -f "$LIB_DIR/libmip_ucs_sdk.so" ]]; then
  LINK_LIBS+=(-lmip_ucs_sdk)
fi

"$CXX" -std=c++17 -O2 -o "$BUILD_DIR/mip-decrypt" decrypt.cpp \
  -I"$MIP_SDK_ROOT/include" \
  -L"$LIB_DIR" \
  "${LINK_LIBS[@]}" \
  "${RPATH_ARGS[@]}"

echo "built $BUILD_DIR/mip-decrypt ($OS/$ARCH)"
