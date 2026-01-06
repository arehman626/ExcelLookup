# ExcelLookup Development Makefile

.PHONY: help install install-dev test test-cov lint format clean build upload docs

# Default target
help:
	@echo "Available targets:"
	@echo "  install     - Install package for production"
	@echo "  install-dev - Install package for development"
	@echo "  test        - Run tests"
	@echo "  test-cov    - Run tests with coverage"
	@echo "  lint        - Run linting"
	@echo "  format      - Format code"
	@echo "  clean       - Clean build artifacts"
	@echo "  build       - Build package"
	@echo "  docs        - Generate documentation"

# Installation
install:
	pip install .

install-dev:
	pip install -e ".[dev]"

# Testing
test:
	pytest

test-cov:
	pytest --cov=src --cov-report=html --cov-report=term

# Code quality
lint:
	flake8 src tests
	mypy src

format:
	black src tests

# Cleanup
clean:
	rm -rf build/
	rm -rf dist/
	rm -rf *.egg-info/
	rm -rf .pytest_cache/
	rm -rf htmlcov/
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -type f -name "*.pyc" -delete

# Build
build: clean
	python -m build

# Development server (if needed)
dev:
	python main.py --help