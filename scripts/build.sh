#!/usr/bin/env sh
set -eu
cd "$(dirname "$0")/.."
mkdir -p dist
for target in windows/amd64 windows/arm64 darwin/amd64 darwin/arm64 linux/amd64; do
  os=${target%/*}; arch=${target#*/}; name=$os; a=$arch; ext=
  [ "$name" != darwin ] || name=macos
  [ "$a" != amd64 ] || a=x64
  [ "$os" != windows ] || ext=.exe
  GOOS="$os" GOARCH="$arch" CGO_ENABLED=0 go build -trimpath -ldflags='-s -w' -o "dist/wordwright-$name-$a$ext" ./cmd/wordwright
done
