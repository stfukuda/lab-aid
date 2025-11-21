set dotenv-load := true
uv := env_var_or_default("UV", "uv")

# Default recipe shows available commands
default: help

# Display available recipes
help:
    @just --list --unsorted

# Initialize git, create dev branch, install deps/hooks
setup:
    #!/usr/bin/env sh
    if command -v git >/dev/null 2>&1; then
        if [ ! -d .git ]; then
            git init
        fi
        if ! git rev-parse --verify HEAD >/dev/null 2>&1; then
            git -c commit.gpgSign=false commit --allow-empty -m "Initial commit"
        fi
        if git rev-parse --verify dev >/dev/null 2>&1; then
            git switch dev
        else
            git switch -c dev
        fi
        git add -A
        if ! git diff --cached --quiet; then
            git -c commit.gpgSign=false commit -m "chore: scaffold project"
        fi
    else
        echo "git not found; skipping repository initialization."
    fi
    uv sync --all-groups
    if command -v git >/dev/null 2>&1; then
        git add -A
        if ! git diff --cached --quiet; then
            git -c commit.gpgSign=false commit -m "chore: sync dependencies"
        fi
        prek install
    fi
    echo "Setup complete."

# Re-sync dependencies and hooks
sync:
    #!/usr/bin/env sh
    uv sync --all-groups
    if command -v git >/dev/null 2>&1; then
        prek install
    else
        echo "git not found; skipping pre-commit installation."
    fi
    echo "Environment sync complete."

# Upgrade locked deps/hooks and commit changes
update:
    #!/usr/bin/env sh
    uv lock --upgrade --refresh
    uv sync --all-groups
    prek autoupdate
    if command -v git >/dev/null 2>&1 && [ -d .git ]; then
        for file in pyproject.toml uv.lock .pre-commit-config.yaml; do
            if [ -e "$file" ]; then
                git add "$file"
            fi
        done
        if ! git diff --cached --quiet; then
            git -c commit.gpgSign=false commit -m "chore: refresh tooling"
        else
            echo "No update-related changes to commit."
        fi
        prek install
    else
        echo "git not found; skipping commit and hook installation."
    fi
    echo "Tooling update complete."

# Format Python sources/tests
format:
    uv run ruff format ./src ./tests

# Run Ruff lint, mypy, and Bandit
lint:
    uv run ruff check --fix ./src ./tests
    uv run mypy ./src ./tests
    uv run bandit -c pyproject.toml -r ./src ./tests

check: format lint
    @echo "Formatting and linting complete."

# Run pytest with coverage
test:
    uv run pytest --cov=src --cov-report=term-missing -n auto ./tests

# Build sdist/wheel artifacts (after tests succeed)
build: test
    uv build

# Build Sphinx HTML docs
docs:
    uv run sphinx-build -b html docs docs/_build/html

# Live-reload docs with sphinx-autobuild
docs-serve:
    uv run sphinx-autobuild docs docs/_build/html

# Clean caches and build artifacts
clean:
    find ./ -type f -name "*.py[co]" -delete
    find ./ -type d -name "__pycache__" -prune -exec rm -rf {} +
    rm -rf dist htmlcov .coverage .pytest_cache .ruff_cache .mypy_cache docs/_build
    echo "Workspace cleaned."
