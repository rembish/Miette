VENV   := .venv
PYTHON := $(VENV)/bin/python
PIP    := $(VENV)/bin/pip

.PHONY: help install lint format typecheck test pre-commit clean

help:
	@echo "Usage: make <target>"
	@echo ""
	@echo "  install     set up the virtual environment"
	@echo "  lint        run ruff"
	@echo "  format      run black"
	@echo "  typecheck   run mypy"
	@echo "  test        run pytest with coverage"
	@echo "  pre-commit  install pre-commit hooks"
	@echo "  clean       remove build artifacts and caches"

$(VENV)/bin/activate: pyproject.toml
	python3 -m venv $(VENV)
	$(PIP) install --upgrade pip
	$(PIP) install -e "../cfb"
	$(PIP) install -e ".[dev]"
	touch $(VENV)/bin/activate

install: $(VENV)/bin/activate
	@echo "Virtual environment ready at $(VENV)/"

lint: $(VENV)/bin/activate
	$(VENV)/bin/ruff check miette tests

format: $(VENV)/bin/activate
	$(VENV)/bin/black miette tests

typecheck: $(VENV)/bin/activate
	$(VENV)/bin/mypy miette

test: $(VENV)/bin/activate
	$(VENV)/bin/pytest --cov=miette --cov-report=term-missing

pre-commit: $(VENV)/bin/activate
	$(VENV)/bin/pre-commit install

clean:
	rm -rf $(VENV) build dist *.egg-info .mypy_cache .ruff_cache .pytest_cache __pycache__
	find . -type d -name __pycache__ -exec rm -rf {} +
	find . -name "*.pyc" -delete
