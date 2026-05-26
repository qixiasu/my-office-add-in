package main

import (
	"fmt"
	"net/http"
	"os"
	"path/filepath"
)

type HTTPServer struct {
	port  int
	dir   string
	server *http.Server
}

func NewHTTPServer(port int, dir string) *HTTPServer {
	return &HTTPServer{
		port: port,
		dir:  dir,
	}
}

func (h *HTTPServer) Start() error {
	absDir, err := filepath.Abs(h.dir)
	if err != nil {
		return fmt.Errorf("failed to get absolute path: %w", err)
	}

	if _, err := os.Stat(absDir); os.IsNotExist(err) {
		if err := os.MkdirAll(absDir, 0755); err != nil {
			return fmt.Errorf("failed to create directory %s: %w", absDir, err)
		}
	}

	mux := http.NewServeMux()
	mux.Handle("/", http.FileServer(http.Dir(absDir)))

	h.server = &http.Server{
		Addr:    fmt.Sprintf(":%d", h.port),
		Handler: mux,
	}

	return h.server.ListenAndServe()
}

func (h *HTTPServer) Stop() error {
	if h.server != nil {
		return h.server.Close()
	}
	return nil
}