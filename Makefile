.PHONY: build test serve dev clean help web-build web-test lint

# Default target
.DEFAULT_GOAL := help

# Ensure cargo and wasm-pack are in PATH
export PATH := $(HOME)/.cargo/bin:$(PATH)

help: ## Show this help message
	@echo "Available commands:"
	@echo ""
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'

build: ## Build Rust to WebAssembly
	@echo "Building Rust to WebAssembly..."
	@cd rust && wasm-pack build --target web --out-dir ../web/pkg
	@echo "✓ Build complete"

test: ## Run Rust tests
	@echo "Running Rust tests..."
	@cd rust && cargo test
	@echo "✓ Tests complete"

serve: ## Start Vite dev server
	@cd web && npm run dev

dev: build serve ## Build WASM and start Vite dev server

web-build: ## Build web frontend for production
	@cd web && npm run build

web-test: ## Run web frontend tests (Vitest + Playwright)
	@cd web && npm test

web-check: ## Type-check web frontend without building
	@echo "Type-checking web frontend..."
	@cd web && npx tsc -b
	@echo "✓ Web check complete"

clean: ## Clean build artifacts
	@echo "Cleaning build artifacts..."
	@cd rust && cargo clean
	@rm -rf web/pkg
	@echo "✓ Clean complete"

check: ## Check Rust code without building
	@echo "Checking Rust code..."
	@cd rust && cargo check
	@echo "✓ Check complete"

lint: ## Lint Rust code with Clippy
	@echo "Linting Rust code..."
	@cd rust && cargo clippy -- -D warnings
	@echo "✓ Lint complete"

fmt: ## Format Rust code
	@echo "Formatting Rust code..."
	@cd rust && cargo fmt
	@echo "✓ Format complete"

watch: ## Watch for changes and rebuild (requires cargo-watch)
	@echo "Watching for changes..."
	@cd rust && cargo watch -s 'wasm-pack build --target web --out-dir ../web/pkg'
