"""Lab-Aid のビルトイン関数で共有するヘルパー群。"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from typing import Any


@dataclass
class BuiltinNumericResult:
    """ビルトイン関数の評価結果を保持するコンテナ。

    Attributes:
        value: ビルトイン関数が返す数値または文字列。
        format_hint: 表示形式を指示するヒント。`None` の場合は指定なし。
    """

    value: Any
    format_hint: str | None = None


def to_decimal(value: Any) -> Decimal:
    """値を `Decimal` に変換する。

    Args:
        value: 数値化対象の値。浮動小数点は桁落ちを防ぐため文字列化してから変換する。

    Returns:
        変換後の `Decimal` 値。
    """
    if isinstance(value, float):
        return Decimal(str(value))
    return Decimal(value)


def force_int_if_integral(dec_value: Decimal, rounding: str | None) -> Any:
    """端数がない場合は整数へ変換し、そうでなければ浮動小数点のまま返す。

    Args:
        dec_value: 判定対象の `Decimal` 値。
        rounding: `Decimal.to_integral_value` に渡す丸めモード。

    Returns:
        整数または浮動小数点。
    """
    if dec_value == dec_value.to_integral_value(rounding=rounding):
        return int(dec_value)
    return float(dec_value)


def normalize_number(value: float | int) -> Any:
    """浮動小数点で表現された整数値を int に縮約する。

    Args:
        value: 正規化対象の数値。

    Returns:
        整数として表現可能な場合は int、それ以外は元の値。
    """
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def ensure_number(value: Any, func: str) -> float:
    """ビルトイン関数向けに数値を保証する。

    Args:
        value: 検証対象の値。
        func: エラーメッセージに利用する関数名。

    Returns:
        float 型に変換した数値。

    Raises:
        TypeError: 真偽値や非数値が指定された場合。
    """
    if isinstance(value, bool):
        raise TypeError(f"{func}: 真偽値は指定できません。")
    if isinstance(value, (int, float)):
        return float(value)
    raise TypeError(f"{func}: 数値を指定してください。")


def ensure_int(value: Any, func: str, name: str) -> int:
    """ビルトイン関数向けに整数を保証する。

    Args:
        value: 検証対象の値。
        func: エラーメッセージに利用する関数名。
        name: エラーメッセージに利用する引数名。

    Returns:
        int 型に変換した整数。

    Raises:
        TypeError: 整数へ変換できない場合。
    """
    number = ensure_number(value, func)
    if not float(number).is_integer():
        raise TypeError(f"{func}: {name} には整数を指定してください。")
    return int(number)


def collect_numeric_values(source: Any, func: str) -> list[float]:
    """入力値をリスト化し、すべて数値であることを保証する。

    Args:
        source: 単一値または値リスト。
        func: エラーメッセージに利用する関数名。

    Returns:
        数値のみから成るリスト。

    Raises:
        ValueError: 空リストが指定された場合。
        TypeError: リスト内に非数値が含まれる場合。
    """
    if isinstance(source, list):
        values = [ensure_number(item, func) for item in source]
        if not values:
            raise ValueError(f"{func}: 空のデータは指定できません。")
        return values
    return [ensure_number(source, func)]


def select_value(source: Any, index: int | None) -> Any:
    """リストから指定位置の値を取り出すか、単一値をそのまま返す。

    Args:
        source: 単一値または値リスト。
        index: 1 始まりのインデックス。省略時は先頭を選択。

    Returns:
        選択された値。範囲外の場合は空文字列を返す。
    """
    if isinstance(source, list):
        if not source:
            return ""
        if index is None:
            return source[0]
        position = index - 1
        if 0 <= position < len(source):
            return source[position]
        return ""
    return source
