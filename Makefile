## Makefile — docker-first dev workflow for docxsphinx.
##
## Usage:
##   make build                 build the dev image (python:3.12 by default)
##   make shell                 interactive shell inside the container
##   make test                  run pytest in the container
##   make lint                  run ruff in the container
##   make cov                   run pytest with coverage
##   make example N=2           build examples/sample_2 in the container
##   make profile N=2           profile writer.py against examples/sample_2
##   make release-check         python -m build + twine check dist/*
##   make clean                 remove example builds and python caches
##
## Host-native escape hatch (skip docker entirely):
##   make test LOCAL=1
##
## Python version override:
##   PYTHON_VERSION=3.11 make build
##
## UID/GID of the container dev user default to the host's so that files
## written via the bind-mount stay owned by you on the host.

export UID              ?= $(shell id -u)
export GID              ?= $(shell id -g)
export PYTHON_VERSION   ?= 3.12

COMPOSE ?= docker compose
RUN     ?= $(COMPOSE) run --rm dev
ifdef LOCAL
RUN =
endif

.DEFAULT_GOAL := help

.PHONY: help build shell test test-smoke test-unit test-integration test-e2e test-fast test-all lint cov example profile roundtrip release-check clean distclean

# roundtrip corpus defaults: `make roundtrip` analyses the committed
# fixture set; `make roundtrip CORPUS=corpus` analyses the user's private
# corpus (./examples/corpus/, gitignored).  Reports always land in
# examples/corpus/reports/ which is gitignored.
CORPUS  ?= corpus_samples
REPORTS ?= examples/corpus/reports

help:
	@awk 'BEGIN{FS=":.*##"} /^[a-zA-Z0-9_-]+:.*##/ {printf "  %-18s %s\n", $$1, $$2}' $(MAKEFILE_LIST)

build: ## Build the dev image
	$(COMPOSE) build

shell: ## Interactive shell in the container
	$(COMPOSE) run --rm dev bash

test: ## Run the full pytest suite (all tiers)
	$(RUN) pytest -v

test-smoke: ## Smoke tier only (seconds)
	$(RUN) pytest -v -m smoke

test-unit: ## Unit tier only (no subprocess, no disk I/O)
	$(RUN) pytest -v -m unit

test-integration: ## Integration tier only (visitor + python-docx, in-memory)
	$(RUN) pytest -v -m integration

test-e2e: ## End-to-end tier only (sphinx-build subprocess per example)
	$(RUN) pytest -v -m e2e

test-fast: ## Smoke + unit + integration (skip slow e2e)
	$(RUN) pytest -v -m "smoke or unit or integration"

test-all: test ## Alias for the full suite

lint: ## Run ruff
	$(RUN) ruff check src tests

cov: ## Pytest with coverage
	$(RUN) pytest --cov=docxsphinx --cov-report=term-missing

EXAMPLE_DIR = $(if $(DIR),$(DIR),$(if $(N),sample_$(N),))

example: ## Build an example  (usage: make example N=2  |  make example DIR=md_basic)
ifeq ($(EXAMPLE_DIR),)
	$(error set N=<sample-number> or DIR=<example-dir>, e.g. make example N=2 or make example DIR=md_basic)
endif
	$(RUN) bash -c "cd examples/$(EXAMPLE_DIR) && rm -rf build && PYTHONPATH=/workspace/src sphinx-build -b docx source build"

profile: ## cProfile the writer against an example (usage: make profile N=2  |  make profile DIR=md_basic)
ifeq ($(EXAMPLE_DIR),)
	$(error set N=<sample-number> or DIR=<example-dir>, e.g. make profile N=2 or make profile DIR=md_basic)
endif
	$(RUN) bash -c "cd examples/$(EXAMPLE_DIR) && rm -rf build && PYTHONPATH=/workspace/src python -m cProfile -s calls \$$(which sphinx-build) -M docx source build/docx/ 2>&1 | grep writer.py | awk '{print \$$6}' | sort | uniq -c | sort -rn | head -30"

roundtrip: ## Roundtrip .docx files through pandoc→MD→docxsphinx; diff vs original (CORPUS=corpus_samples|corpus)
	$(RUN) python tools/roundtrip.py examples/$(CORPUS) $(REPORTS)

release-check: ## Build sdist+wheel and run twine check
	$(RUN) python -m build
	$(RUN) twine check dist/*

clean: ## Remove example build artifacts and python caches
	@rm -rf examples/*/build dist build .pytest_cache .ruff_cache .coverage coverage.xml htmlcov
	@rm -f docx.log examples/*/docx.log
	@find . -type d -name __pycache__ -prune -exec rm -rf {} +
	@find . -type d -name '*.egg-info' -prune -exec rm -rf {} +

distclean: clean ## Remove images and caches
	-$(COMPOSE) down --rmi local --volumes
