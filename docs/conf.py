from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path
from importlib import metadata

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

project = "Lab-Aid"
author = "stfukuda"
current_year = datetime.utcnow().year
copyright = f"{current_year}, {author}"
try:
    release = metadata.version("lab-aid")
except metadata.PackageNotFoundError:
    try:
        from lab_aid import __version__ as release
    except Exception:  # pragma: no cover - docs fallback
        release = "0.0.0"

extensions = [
    "myst_parser",
    "sphinx.ext.autodoc",
    "sphinx.ext.autosummary",
    "sphinx.ext.napoleon",
    "sphinx.ext.viewcode",
    "sphinx.ext.intersphinx",
    "sphinx.ext.todo",
    "sphinx_autodoc_typehints",
]

templates_path = ["_templates"]
exclude_patterns: list[str] = [
    "guidelines",
    "design",
    "official-docs",
]

html_theme = "furo"
html_static_path = ["_static"]

source_suffix = {
    ".rst": "restructuredtext",
    ".md": "markdown",
}

autodoc_member_order = "bysource"
autodoc_typehints = "description"
autosummary_generate = True

napoleon_google_docstring = False
napoleon_numpy_docstring = True

intersphinx_mapping = {
    "python": ("https://docs.python.org/3", None),
}

todo_include_todos = True
