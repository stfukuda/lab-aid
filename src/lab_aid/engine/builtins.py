"""互換性維持のために残している旧ビルトイン API。"""

from __future__ import annotations

from .runtime.functions import (
    NUMERIC_FUNCTIONS,
    PACKAGE_FUNCTIONS,
    STRING_FUNCTIONS,
    BuiltinNumericResult,
    format_roundjisb_output,
    str_comp,
)
from .runtime.functions.package import execute_print, execute_print2

__all__ = [
    "NUMERIC_FUNCTIONS",
    "STRING_FUNCTIONS",
    "PACKAGE_FUNCTIONS",
    "BuiltinNumericResult",
    "execute_print",
    "execute_print2",
    "str_comp",
    "format_roundjisb_output",
]
