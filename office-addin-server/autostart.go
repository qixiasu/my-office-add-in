package main

import (
	"fmt"
	"syscall"
	"unsafe"
)

var (
	advapi32 = syscall.MustLoadDLL("advapi32.dll")
	regOpenKeyExW  = advapi32.MustFindProc("RegOpenKeyExW")
	regSetValueExW = advapi32.MustFindProc("RegSetValueExW")
	regCloseKey    = advapi32.MustFindProc("RegCloseKey")
	regDeleteValueW = advapi32.MustFindProc("RegDeleteValueW")
)

func SetAutostart(name string, exePath string) error {
	var key syscall.Handle
	r0, _, err := regOpenKeyExW.Call(
		syscall.HKEY_CURRENT_USER,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(`Software\Microsoft\Windows\CurrentVersion\Run`))),
		0,
		syscall.KEY_WRITE,
		uintptr(unsafe.Pointer(&key)),
	)
	if r0 != 0 {
		return fmt.Errorf("failed to open registry key: %w", err)
	}
	defer regCloseKey.Call(uintptr(key))

	r0, _, err = regSetValueExW.Call(
		uintptr(key),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(name))),
		0,
		syscall.REG_SZ,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(exePath))),
		uintptr(len(exePath)*2),
	)
	if r0 != 0 {
		return fmt.Errorf("failed to set registry value: %w", err)
	}

	return nil
}

func RemoveAutostart(name string) error {
	var key syscall.Handle
	r0, _, err := regOpenKeyExW.Call(
		syscall.HKEY_CURRENT_USER,
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(`Software\Microsoft\Windows\CurrentVersion\Run`))),
		0,
		syscall.KEY_WRITE,
		uintptr(unsafe.Pointer(&key)),
	)
	if r0 != 0 {
		return fmt.Errorf("failed to open registry key: %w", err)
	}
	defer regCloseKey.Call(uintptr(key))

	r0, _, err = regDeleteValueW.Call(
		uintptr(key),
		uintptr(unsafe.Pointer(syscall.StringToUTF16Ptr(name))),
	)
	if r0 != 0 && r0 != 2 {
		return fmt.Errorf("failed to delete registry value: %w", err)
	}

	return nil
}