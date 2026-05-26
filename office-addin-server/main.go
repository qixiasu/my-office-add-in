package main

import (
	"flag"
	"fmt"
	"log"
	"os"
	"os/signal"
	"syscall"

	"github.com/getlantern/systray"
)

var (
	flagPort = flag.Int("port", 3000, "HTTP server port")
	flagDir  = flag.String("dir", "./dist", "Static files directory")
	flagHelp = flag.Bool("h", false, "Show help")
)

func main() {
	flag.Parse()

	if *flagHelp {
		flag.Usage()
		os.Exit(0)
	}

	cfg, err := LoadConfig()
	if err != nil {
		log.Printf("Warning: failed to load config: %v, using defaults", err)
		cfg = DefaultConfig()
	}

	port := *flagPort
	dir := *flagDir
	if dir == "./dist" && cfg.Dir != "" {
		dir = cfg.Dir
	}
	if port == 3000 && cfg.Port != 0 {
		port = cfg.Port
	}

	exePath, err := GetExecutablePath()
	if err != nil {
		log.Fatalf("Failed to get executable path: %v", err)
	}

	httpServer := NewHTTPServer(port, dir)

	sigCh := make(chan os.Signal, 1)
	signal.Notify(sigCh, syscall.SIGINT, syscall.SIGTERM)

	go func() {
		<-sigCh
		httpServer.Stop()
		os.Exit(0)
	}()

	systray.Run(func() {
		tray := NewTrayManager(port, cfg.Autostart, func(enabled bool) {
			cfg.Autostart = enabled
			if err := SaveConfig(cfg); err != nil {
				log.Printf("Failed to save config: %v", err)
			}

			if enabled {
				if err := SetAutostart("OfficeAddinServer", exePath); err != nil {
					log.Printf("Failed to enable autostart: %v", err)
				}
			} else {
				if err := RemoveAutostart("OfficeAddinServer"); err != nil {
					log.Printf("Failed to disable autostart: %v", err)
				}
			}
		})
		tray.Setup()

		go func() {
			fmt.Printf("Server starting on http://localhost:%d\n", port)
			if err := httpServer.Start(); err != nil {
				log.Printf("HTTP server error: %v", err)
			}
		}()
	}, func() {})
}