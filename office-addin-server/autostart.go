package main

import (
	"fmt"
	"syscall"
	"unsafe"
)

const (
	regKeyPath = `Software\Microsoft\Windows\CurrentVersion\Run`
)

var (
	advapi32       = syscall.NewLazyDLL("advapi32.dll")
	regOpenKeyExW  = advapi32.NewProc("RegOpenKeyExW")
	regSetValueExW = advapi32.NewProc("RegSetValueExW")
	regCloseKey    = advapi32.NewProc("RegCloseKey")
)

func SetAutostart(name string, exePath string) error {
	var key syscall.Handle
	ret, _, _ := regOpenKeyExW.Call(
		syscall.HKEY_CURRENT_USER,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(regKeyPath))),
		0,
		syscall.KEY_WRITE,
		uintptr(unsafe.Pointer(&key)),
	)
	if ret != 0 {
		return fmt.Errorf("failed to open registry key: %d", ret)
	}
	defer regCloseKey.Call(uintptr(key))

	exePathPtr := syscall.StringToUTF16Ptr(exePath)
	ret, _, _ = regSetValueExW.Call(
		uintptr(key),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(name))),
		0,
		syscall.REG_SZ,
		uintptr(unsafe.Pointer(exePathPtr)),
		uintptr(len(exePath)*2),
	)
	if ret != 0 {
		return fmt.Errorf("failed to set registry value: %d", ret)
	}

	return nil
}

func RemoveAutostart(name string) error {
	var key syscall.Handle
	ret, _, _ := regOpenKeyExW.Call(
		syscall.HKEY_CURRENT_USER,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(regKeyPath))),
		0,
		syscall.KEY_WRITE,
		uintptr(unsafe.Pointer(&key)),
	)
	if ret != 0 {
		return fmt.Errorf("failed to open registry key: %d", ret)
	}
	defer regCloseKey.Call(uintptr(key))

	ret, _, _ = syscall.NewLazyDLL("advapi32.dll").NewProc("RegDeleteValueW").Call(
		uintptr(key),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(name))),
	)
	if ret != 0 && ret != 2 {
		return fmt.Errorf("failed to delete registry value: %d", ret)
	}

	return nil
}