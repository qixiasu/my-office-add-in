package main

import (
	"encoding/json"
	"os"
	"path/filepath"
)

type Config struct {
	Port      int    `json:"port"`
	Dir       string `json:"dir"`
	Autostart bool   `json:"autostart"`
}

func DefaultConfig() *Config {
	return &Config{
		Port:      3000,
		Dir:       "./dist",
		Autostart: false,
	}
}

func LoadConfig() (*Config, error) {
	cfg := DefaultConfig()

	exePath, err := os.Executable()
	if err != nil {
		return cfg, err
	}
	cfgDir := filepath.Dir(exePath)
	cfgFile := filepath.Join(cfgDir, "config.json")

	data, err := os.ReadFile(cfgFile)
	if err != nil {
		if os.IsNotExist(err) {
			return cfg, nil
		}
		return nil, err
	}

	if err := json.Unmarshal(data, cfg); err != nil {
		cfg = DefaultConfig()
		return cfg, nil
	}

	return cfg, nil
}

func SaveConfig(cfg *Config) error {
	exePath, err := os.Executable()
	if err != nil {
		return err
	}
	cfgDir := filepath.Dir(exePath)
	cfgFile := filepath.Join(cfgDir, "config.json")

	data, err := json.MarshalIndent(cfg, "", "  ")
	if err != nil {
		return err
	}

	return os.WriteFile(cfgFile, data, 0644)
}