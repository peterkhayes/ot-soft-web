.PHONY: build test serve dev clean help

# Default target
.DEFAULT_GOAL := help

# Ensure cargo is in PATH
CARGO := $(shell if [ -f ~/.cargo/bin/cargo ]; then echo "PATH=$$HOME/.cargo/bin:$$PATH"; fi)
WASM_PACK := $(shell if [ -f ~/.cargo/bin/wasm-pack ]; then echo "PATH=$$HOME/.cargo/bin:$$PATH wasm-pack"; else echo "wasm-pack"; fi)

help: ## Show this help message
	@echo "Available commands:"
	@echo ""
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'

build: ## Build Rust to WebAssembly
	@echo "Building Rust to WebAssembly..."
	@cd rust && $(WASM_PACK) build --target web --out-dir ../web/pkg
	@echo "✓ Build complete"

test: ## Run Rust tests
	@echo "Running Rust tests..."
	@cd rust && $(CARGO) cargo test
	@echo "✓ Tests complete"

serve: ## Start local web server on port 8000
	@echo "Starting web server at http://localhost:8000"
	@echo "Press Ctrl+C to stop"
	@cd web && python3 -m http.server 8000

dev: build serve ## Build and start development server

clean: ## Clean build artifacts
	@echo "Cleaning build artifacts..."
	@cd rust && $(CARGO) cargo clean
	@rm -rf web/pkg
	@echo "✓ Clean complete"

check: ## Check Rust code without building
	@echo "Checking Rust code..."
	@cd rust && $(CARGO) cargo check
	@echo "✓ Check complete"

fmt: ## Format Rust code
	@echo "Formatting Rust code..."
	@cd rust && $(CARGO) cargo fmt
	@echo "✓ Format complete"

watch: ## Watch for changes and rebuild (requires cargo-watch)
	@echo "Watching for changes..."
	@cd rust && $(CARGO) cargo watch -s 'wasm-pack build --target web --out-dir ../web/pkg'
