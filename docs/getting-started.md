# Getting Started

This guide helps you build and preview the documentation locally.

## Build the docs

Run the provided Just recipe (recommended):

```bash
just docs
```

Or call Sphinx directly:

```bash
uv run sphinx-build -b html docs docs/_build/html
```

## Live preview while editing

Use the Just live preview recipe (recommended):

```bash
just docs-serve
```

Or run Sphinx autobuild directly:

```bash
uv run sphinx-autobuild docs docs/_build/html
```

## Writing content

- Write prose in Markdown or reStructuredText.
- Reference API members directly; docstrings are pulled in via `autodoc`.
- Keep long code examples in `examples/` (if any) and include them with code blocks.
