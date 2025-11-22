# Lab-Aid

Tools and assets for Lab-Aid.

## Installation

Install from PyPI:

```shell
pip install -U lab_aid
```

## Usage

Describe how to use the library in this section.

## Development

### Prerequisites

- Docker Desktop (or Docker Engine with the Dev Container CLI)
- One of the following:
  - [Visual Studio Code](https://code.visualstudio.com/) with the Dev Containers
    extension
  - The [Dev Container CLI](https://github.com/devcontainers/cli)
    (`npm install -g @devcontainers/cli`)
- [Just](https://github.com/casey/just) task runner (optional but recommended)
- [PreK](https://github.com/chriskuehl/prek) (`uv tool install prek`) if you run hooks
  outside the Dev Container

> The Dev Container image installs `uv`, `just`, and `prek`, so host installation is
> only needed if you plan to run the recipes outside the container.

### Getting started (VS Code)

1. Open the repository in VS Code.
2. When prompted, select **“Reopen in Container”** (or run the Dev Containers command).
3. Wait for the container build to complete, then run `just setup` inside the container
   to sync dependencies and install PreK hooks.

### Getting started (Dev Container CLI)

```shell
devcontainer up --workspace-folder .
devcontainer exec --workspace-folder . just setup
```

Run additional commands as needed, for example
`devcontainer exec --workspace-folder . just check`.

### Common workflows (inside the container)

- `just setup` – initialize Git (if available), create the `dev` branch, install deps,
  and bootstrap PreK hooks
- `just sync` – refresh dependencies via `uv sync` and re-install PreK hooks when Git
  is available
- `just update` – upgrade locked dependencies + PreK hooks and commit the changes
- `just check` – format & lint/type-check the codebase (Ruff + mypy + Bandit)
- `just test` – execute the test suite with pytest
- `just build` – build distributable artifacts into dist/
- `just docs` – build the Sphinx docs into `docs/_build/html`
- `just docs-serve` – run `sphinx-autobuild` for live preview while editing docs
- `just --list` – show all available recipes

### Cleaning up

Remove cached artifacts and build caches with:

```shell
just clean
```

### Release workflow

- Formal release: create an annotated git tag like `v1.2.3` and push it.
  `hatch-vcs` derives the package version from the tag during `uv build`, so you never
  edit `pyproject.toml` manually.
- Test release: create a pre-release tag such as `v1.2.3-dev.1`. The release workflow
  publishes the build to TestPyPI using the `TEST_PYPI_API_TOKEN` secret.
- Each release reuses the existing artifacts; ensure the secrets are configured before
  tagging.

## Documentation

- Docstrings drive the API reference via Sphinx + autodoc.
- Author additional guides under `docs/` (Markdown or reStructuredText). Update the
  toctree in `docs/index.md` as needed.
- Build locally with `just docs`; preview with `just docs-serve`.

## License

[MIT License](LICENSE)
