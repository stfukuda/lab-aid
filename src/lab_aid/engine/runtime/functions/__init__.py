"""ビルトイン関数のレジストリをまとめたモジュール。"""

from __future__ import annotations

from .base import (
    BuiltinNumericResult,
    collect_numeric_values,
    ensure_int,
    ensure_number,
    force_int_if_integral,
    normalize_number,
    select_value,
    to_decimal,
)
from .numeric import NUMERIC_FUNCTIONS, format_roundjisb_output
from .package import PACKAGE_FUNCTIONS
from .string import STRING_FUNCTIONS, str_comp

__all__ = [
    "NUMERIC_FUNCTIONS",
    "STRING_FUNCTIONS",
    "PACKAGE_FUNCTIONS",
    "BuiltinNumericResult",
    "collect_numeric_values",
    "ensure_int",
    "ensure_number",
    "force_int_if_integral",
    "normalize_number",
    "to_decimal",
    "select_value",
    "format_roundjisb_output",
    "str_comp",
]
