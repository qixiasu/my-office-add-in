@echo off
setlocal

set APP_NAME=office-addin-server
set BUILD_DIR=dist

if not exist "%BUILD_DIR%" mkdir "%BUILD_DIR%"

GOOS=windows GOARCH=amd64 go build -o "%BUILD_DIR%\%APP_NAME%.exe" .

echo Build complete: %BUILD_DIR%\%APP_NAME%.exe
