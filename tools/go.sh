#!/usr/bin/env sh
# Explicit toolchain, isolated caches, offline by default, like go.ps1.
set -eu
repo=$(CDPATH= cd -- "$(dirname -- "$0")/.." && pwd)
tools=$(dirname "$repo")/.tools
toolchain=${WORDUP_GO:-$tools/go/bin/go}
[ -x "$toolchain" ] || { echo "Pinned Go toolchain not found: $toolchain" >&2; exit 1; }
if [ "${1:-}" = -Online ]; then shift; else export GOPROXY=off; fi
export GOCACHE="$tools/gocache" GOMODCACHE="$tools/gomodcache" GOTMPDIR="$tools/gotmp"
export GOTOOLCHAIN=local GOFLAGS="${GOFLAGS:-} -buildvcs=false"
mkdir -p "$GOCACHE" "$GOMODCACHE" "$GOTMPDIR"
cd "$repo"
exec "$toolchain" "$@"
