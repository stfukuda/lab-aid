"""Lab-Aid の印字系ビルトイン関数（print 系）をまとめたモジュール。"""

from __future__ import annotations

from collections.abc import Callable
from typing import Any


def execute_print(
    arg_expr: str,
    evaluator: Callable[[str], Any],
    to_text_fn: Callable[[Any], str | None],
) -> str | None:
    """`print` ビルトインを評価し、Lab-Aid 互換の文字列を返す。

    Args:
        arg_expr: 引数式を表す文字列。
        evaluator: 式を評価するコールバック。
        to_text_fn: 評価結果を文字列化するコールバック。

    Returns:
        文字列化された評価結果。値が ``None`` の場合は ``None``。
    """
    value = evaluator(arg_expr)
    return to_text_fn(value)


def execute_print2(
    arg_expr: str,
    evaluator: Callable[[str], Any],
    to_text_fn: Callable[[Any], str | None],
) -> str | None:
    """`print2` ビルトインを評価し、Lab-Aid 互換の文字列を返す。

    Args:
        arg_expr: 引数式を表す文字列。
        evaluator: 式を評価するコールバック。
        to_text_fn: 評価結果を文字列化するコールバック。

    Returns:
        文字列化された評価結果。値が ``None`` の場合は ``None``。
    """
    value = evaluator(arg_expr)
    return to_text_fn(value)


PACKAGE_FUNCTIONS = {
    "print": execute_print,
    "print2": execute_print2,
}


__all__ = ["PACKAGE_FUNCTIONS", "execute_print", "execute_print2"]
