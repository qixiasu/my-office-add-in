#!/bin/bash
set -e

APP_NAME="office-addin-server"
BUILD_DIR="dist"

mkdir -p "$BUILD_DIR"

GOOS=windows GOARCH=amd64 go build -o "${BUILD_DIR}/${APP_NAME}.exe" .

echo "Build complete: ${BUILD_DIR}/${APP_NAME}.exe"