"""Lab-Aid スクリプト向けの入力解析ヘルパー。"""

from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any

from .text import (
    parse_number_like,
    replace_word_ci_outside_quotes,
    strip_comment_quote_aware,
    unquote_single,
    validate_hash_name,
)

RE_ITEM_ANY = re.compile(
    r"#([A-Za-z0-9_]+(?:\.[A-Za-z0-9_]+)?)(?:\[([A-Za-z0-9_]+)\])?"
)
RE_ITEM_STRICT = re.compile(r"#([A-Z0-9_]+(?:\.[A-Z0-9_]+)?)(?:\[([A-Z0-9_]+)\])?")
RE_INPUT_LINE = re.compile(
    r"^\s*(#?[A-Za-z0-9_]+(?:\.[A-Za-z0-9_]+)?(?:\[[A-Za-z0-9_]+\])?)\s*=\s*(.+?)\s*$"
)
THIS_ASSIGN_RE = re.compile(r"^\s*this\s*=", re.IGNORECASE)


@dataclass
class VarRef:
    """通常変数（`this` 以外）への参照を表す。

    Attributes:
        name: 参照対象の変数名。
    """

    name: str


def _split_multi_values(value: str) -> list[str]:
    """カンマ区切りの値文字列を単一値リストへ分割する。

    単一引用符で囲まれたカンマは文字列として扱い、エスケープされた二重引用符も
    適切に復元する。

    Args:
        value: 分割対象となる値文字列。

    Returns:
        先頭・末尾の空白を除いた値文字列のリスト。
    """
    tokens: list[str] = []
    current: list[str] = []
    in_sq = False
    index = 0
    length = len(value)
    while index < length:
        char = value[index]
        if char == "'":
            current.append(char)
            if in_sq and index + 1 < length and value[index + 1] == "'":
                current.append("'")
                index += 2
                continue
            in_sq = not in_sq
            index += 1
            continue
        if char == "," and not in_sq:
            tokens.append("".join(current).strip())
            current = []
            index += 1
            continue
        current.append(char)
        index += 1
    if current or value.endswith(","):
        tokens.append("".join(current).strip())
    return [token for token in tokens if token]


def _parse_value(token: str, line_no: int) -> Any:
    """E タイプ入力の値トークンを解析する。

    Args:
        token: 単一値を表す文字列。
        line_no: エラーメッセージ用の行番号。

    Returns:
        数値、もしくは単一引用符を除去した文字列。

    Raises:
        ValueError: 数値と文字列以外の形式が指定された場合。
    """
    if token.startswith("'") and token.endswith("'"):
        return unquote_single(token)
    number = parse_number_like(token)
    if number is not None:
        return number
    raise ValueError(
        f"E入力の値は数値または '文字' を指定してください（{line_no}行目）: {token!r}"
    )


def parse_inputs_E(inputs: str) -> dict[Any, Any]:
    """E タイプ計算のために複数行の `NAME=VALUE` を解析する。

    Args:
        inputs: 複数行で構成される Lab-Aid 入力文字列。

    Returns:
        試験項目コードや単位をキーにした辞書。複数値はリストで保持する。

    Raises:
        ValueError: 行の書式が `NAME=VALUE` に一致しない場合。
    """
    items: dict[Any, Any] = {}
    if not inputs or not inputs.strip():
        return items
    for index, raw in enumerate(inputs.splitlines(), 1):
        line = raw.strip()
        if not line or line.lower().startswith("rem"):
            continue
        match = RE_INPUT_LINE.match(line)
        if not match:
            raise ValueError(f"E入力の書式エラー（{index}行目）: {raw!r}")
        name, value_str = match.group(1), match.group(2).strip()

        name_no_hash = name[1:] if name.startswith("#") else name
        unit = None
        if "[" in name_no_hash and name_no_hash.endswith("]"):
            code, unit = name_no_hash[:-1].split("[", 1)
        else:
            code = name_no_hash

        validate_hash_name(code, unit)

        values_raw = _split_multi_values(value_str)
        if values_raw:
            parsed_values = [_parse_value(token, index) for token in values_raw]
            value: Any = parsed_values if len(parsed_values) > 1 else parsed_values[0]
        else:
            value = _parse_value(value_str.strip(), index)

        if unit is None:
            items[code] = value
        else:
            items[(code, unit)] = value
    return items


def parse_input_R(value: str) -> tuple[Any, str]:
    """R タイプ計算で用いる単一値を解析する。

    Args:
        value: Lab-Aid R タイプの第 3 引数となる文字列。

    Returns:
        `(値, リテラル文字列)` のタプル。値は数値、文字列、もしくは `VarRef`。

    Raises:
        ValueError: 空文字列や裸の文字列で変数名と認識できない場合。
    """
    token = (value or "").strip()
    if not token:
        raise ValueError(
            "Rタイプの第3引数は数値または '文字'、もしくは通常変数名を1つ指定してください。"
        )
    if token.startswith("'") and token.endswith("'"):
        return unquote_single(token), unquote_single(token)
    number = parse_number_like(token)
    if number is not None:
        return number, str(token)
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", token):
        raise ValueError(
            f"R入力の値が裸文字列で、かつ通常変数名として解釈できません: {token!r}"
        )
    raise ValueError(
        "Rタイプの第3引数は数値または '文字' を指定してください（通常変数は不可）。"
    )


def ensure_has_this_assignment_E(script: str) -> None:
    """E タイプのスクリプトに `this = ...` が含まれるか検証する。

    Args:
        script: Lab-Aid E タイプのスクリプト。

    Raises:
        ValueError: `this = ...` の行が 1 つも見つからない場合。
    """
    for raw in script.splitlines():
        line = strip_comment_quote_aware(raw)
        if not line or line.lower().startswith("rem"):
            continue
        if THIS_ASSIGN_RE.match(line):
            return
    raise ValueError("Eタイプでは式中に少なくとも1行の 'this = ...' が必要です。")


def replace_rhs_this_for_R(script: str) -> str:
    """R タイプ実行用に右辺の `this` を `__THIS_IN__` へ置換する。

    Args:
        script: Lab-Aid R タイプのスクリプト。

    Returns:
        右辺のみ `__THIS_IN__` に置き換えられたスクリプト文字列。
    """
    lines = []
    for raw in script.splitlines():
        if "=" not in raw:
            lines.append(raw)
            continue
        lhs, rhs = raw.split("=", 1)
        rhs2 = replace_word_ci_outside_quotes(rhs, {"this": "__THIS_IN__"})
        lines.append(f"{lhs}={rhs2}")
    return "\n".join(lines)


__all__ = [
    "RE_ITEM_ANY",
    "RE_ITEM_STRICT",
    "VarRef",
    "parse_inputs_E",
    "parse_input_R",
    "ensure_has_this_assignment_E",
    "replace_rhs_this_for_R",
]
