"""Lab-Aid public API."""

from __future__ import annotations

import subprocess  # nosec B404
from importlib import metadata

__all__ = ["__version__", "hello"]

_FALLBACK_VERSION = "0.0.0"


def _detect_version() -> str:
    try:
        return metadata.version("lab-aid")
    except metadata.PackageNotFoundError:
        git_version = _detect_version_from_git()
        if git_version:
            return git_version
        return _FALLBACK_VERSION


def _detect_version_from_git() -> str | None:
    try:
        completed = subprocess.run(  # nosec B603 B607
            ["git", "describe", "--tags", "--abbrev=0"],
            check=True,
            capture_output=True,
            text=True,
        )
    except (FileNotFoundError, subprocess.CalledProcessError):
        return None

    tag = completed.stdout.strip()
    if tag.startswith("v"):
        return tag[1:] or None
    return tag or None


__version__ = _detect_version()


def hello(name: str) -> str:
    """Return a friendly greeting string.

    Parameters
    ----------
    name:
        The person or thing to greet.
    """

    return f"Hello, {name}!"
