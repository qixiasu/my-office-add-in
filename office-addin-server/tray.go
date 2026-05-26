package main

import (
	"fmt"
	"os"
	"os/exec"
	"path/filepath"

	"github.com/getlantern/systray"
)

type TrayManager struct {
	port      int
	autostart bool
	onToggle  func(bool)
}

func NewTrayManager(port int, autostart bool, onToggle func(bool)) *TrayManager {
	return &TrayManager{
		port:      port,
		autostart: autostart,
		onToggle:  onToggle,
	}
}

func (t *TrayManager) Setup() {
	systray.SetTitle("Office Add-in Server")
	systray.SetTooltip(fmt.Sprintf("Serving on port %d", t.port))

	mOpen := systray.AddMenuItem("Open", "Open in browser")
	mToggleAuto := systray.AddMenuItemCheckbox("Enable Autostart", "Toggle autostart", t.autostart)
	mQuit := systray.AddMenuItem("Quit", "Exit")

	go func() {
		for {
			select {
			case <-mOpen.ClickedCh:
				t.openBrowser()
			case <-mToggleAuto.ClickedCh:
				if mToggleAuto.Checked() {
					mToggleAuto.Uncheck()
					t.autostart = false
				} else {
					mToggleAuto.Check()
					t.autostart = true
				}
				if t.onToggle != nil {
					t.onToggle(t.autostart)
				}
			case <-mQuit.ClickedCh:
				systray.Quit()
				os.Exit(0)
			}
		}
	}()
}

func (t *TrayManager) openBrowser() {
	url := fmt.Sprintf("http://localhost:%d", t.port)
	exec.Command("cmd", "/c", "start", url).Start()
}

func GetExecutablePath() (string, error) {
	exe, err := os.Executable()
	if err != nil {
		return "", err
	}
	return filepath.Abs(exe)
}