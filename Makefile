.PHONY: build test serve dev clean help web-install web-build web-test web-lint web-fmt web-check precommit lint fmt conformance-test watch

# Default target
.DEFAULT_GOAL := help

# Ensure cargo and wasm-pack are in PATH
export PATH := $(HOME)/.cargo/bin:$(PATH)

help: ## Show this help message
	@echo "Available commands:"
	@echo ""
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | awk 'BEGIN {FS = ":.*?## "}; {printf "  \033[36m%-15s\033[0m %s\n", $$1, $$2}'
	@echo ""
	@echo "Flags (pass as VAR=1):"
	@echo "  \033[33mVERBOSE\033[0m    Show full output (test, web-test)"
	@echo "  \033[33mFIX\033[0m        Auto-fix issues (lint, web-lint)"
	@echo "  \033[33mUPDATE\033[0m     Update snapshots (web-test)"
	@echo "  \033[33mCHECK\033[0m      Check only, no writes (web-fmt)"

build: ## Build Rust to WebAssembly
	@echo "Building Rust to WebAssembly..."
	@cd rust && wasm-pack build --target web --out-dir ../web/pkg
	@echo "✓ Build complete"

test: ## Run Rust tests
	@echo "Running Rust tests..."
ifdef VERBOSE
	@cd rust && cargo test
else
	@cd rust && cargo test --quiet
endif
	@echo "✓ Tests complete"

serve: ## Start Vite dev server
	@cd web && npm run dev

dev: build serve ## Build WASM and start Vite dev server

web-install: ## Install web dependencies and Playwright browser (run once after checkout)
	@cd web && npm install && npx playwright install chromium
	@echo "✓ Web install complete"

web-build: ## Build web frontend for production
	@cd web && npm run build

web-test: ## Run web frontend tests
ifdef UPDATE
	@cd web && npx vitest run -u
else ifdef VERBOSE
	@cd web && VERBOSE_TESTS=1 npm test
else
	@cd web && npm test
endif

web-lint: ## Lint web frontend with ESLint
	@echo "Linting web frontend..."
ifdef FIX
	@cd web && npm run lint -- --fix
	@echo "✓ Web lint fix complete"
else
	@cd web && npm run lint
	@echo "✓ Web lint complete"
endif

web-fmt: ## Format web frontend with Prettier
ifdef CHECK
	@echo "Checking web frontend formatting..."
	@cd web && npm run fmt:check
	@echo "✓ Web format check complete"
else
	@echo "Formatting web frontend..."
	@cd web && npm run fmt
	@echo "✓ Web format complete"
endif

web-check: ## Type-check web frontend without building
	@echo "Type-checking web frontend..."
	@cd web && npx tsc -b
	@echo "✓ Web check complete"

precommit: ## Run all checks (Rust + web lint, tests, build)
	@$(MAKE) lint
	@$(MAKE) test
	@$(MAKE) build
	@$(MAKE) web-lint
	@$(MAKE) web-fmt CHECK=1
	@$(MAKE) web-check
	@$(MAKE) web-test

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
ifdef FIX
	@cd rust && cargo clippy --fix --allow-dirty -- -D warnings
	@echo "✓ Lint fix complete"
else
	@cd rust && cargo clippy -- -D warnings
	@echo "✓ Lint complete"
endif

fmt: ## Format Rust code
	@echo "Formatting Rust code..."
	@cd rust && cargo fmt
	@echo "✓ Format complete"

conformance-test: ## Run conformance tests against VB6 golden files
	@echo "Running conformance tests..."
	@cd rust && cargo test conformance_tests -- --nocapture
	@echo "✓ Conformance tests complete"

watch: ## Watch for changes and rebuild (requires cargo-watch)
	@echo "Watching for changes..."
	@cd rust && cargo watch -s 'wasm-pack build --target web --out-dir ../web/pkg'
