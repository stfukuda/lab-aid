"""Lab-Aid の数値系ビルトイン関数群。"""

from __future__ import annotations

import math
import statistics
from collections.abc import Sequence
from decimal import (
    ROUND_DOWN,
    ROUND_FLOOR,
    ROUND_HALF_EVEN,
    ROUND_HALF_UP,
    Decimal,
    InvalidOperation,
)
from typing import Any

from .base import (
    BuiltinNumericResult,
    collect_numeric_values,
    ensure_int,
    ensure_number,
    force_int_if_integral,
    normalize_number,
    to_decimal,
)


def sqrt_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """平方根を計算する。

    Args:
        args: 引数リスト。1 つの数値を想定。

    Returns:
        `BuiltinNumericResult`: 平方根の結果。

    Raises:
        TypeError: 引数数が 1 つでない場合、または数値以外が指定された場合。
        ValueError: 負の数を指定した場合。
    """
    if len(args) != 1:
        raise TypeError("sqrt: 引数は1個 (d) を指定してください。")
    value = ensure_number(args[0], "sqrt")
    if value < 0:
        raise ValueError("sqrt: 負の数には適用できません。")
    return BuiltinNumericResult(normalize_number(math.sqrt(value)))


def log_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """自然対数を計算する。

    Args:
        args: 引数リスト。1 つの数値を想定。

    Returns:
        `BuiltinNumericResult`: 自然対数の結果。

    Raises:
        TypeError: 引数数が 1 つでない場合、または数値以外が指定された場合。
        ValueError: 0 以下の値を指定した場合。
    """
    if len(args) != 1:
        raise TypeError("log: 引数は1個 (d) を指定してください。")
    value = ensure_number(args[0], "log")
    if value <= 0:
        raise ValueError("log: 0 以下の値には適用できません。")
    return BuiltinNumericResult(normalize_number(math.log(value)))


def log10_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """常用対数を計算する。

    Args:
        args: 引数リスト。1 つの数値を想定。

    Returns:
        `BuiltinNumericResult`: 常用対数の結果。

    Raises:
        TypeError: 引数数が 1 つでない場合、または数値以外が指定された場合。
        ValueError: 0 以下の値を指定した場合。
    """
    if len(args) != 1:
        raise TypeError("log10: 引数は1個 (d) を指定してください。")
    value = ensure_number(args[0], "log10")
    if value <= 0:
        raise ValueError("log10: 0 以下の値には適用できません。")
    return BuiltinNumericResult(normalize_number(math.log10(value)))


def exp_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """自然対数の底 `e` による指数関数を計算する。

    Args:
        args: 引数リスト。1 つの数値を想定。

    Returns:
        `BuiltinNumericResult`: 指数関数の計算結果。

    Raises:
        TypeError: 引数数が 1 つでない場合、または数値以外が指定された場合。
    """
    if len(args) != 1:
        raise TypeError("exp: 引数は1個 (d) を指定してください。")
    value = ensure_number(args[0], "exp")
    return BuiltinNumericResult(normalize_number(math.exp(value)))


def pow_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """冪乗を計算する。

    Args:
        args: 引数リスト。底と指数の 2 つの数値を想定。

    Returns:
        `BuiltinNumericResult`: 冪乗の計算結果。

    Raises:
        TypeError: 引数数が 2 つでない場合、または数値以外が指定された場合。
    """
    if len(args) != 2:
        raise TypeError("pow: 引数は2個 (x, y) を指定してください。")
    x = ensure_number(args[0], "pow")
    y = ensure_number(args[1], "pow")
    return BuiltinNumericResult(normalize_number(math.pow(x, y)))


def modi_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """整数部を切り出す。

    Args:
        args: 引数リスト。1 つの数値を想定。

    Returns:
        `BuiltinNumericResult`: 整数部のみを返す。

    Raises:
        TypeError: 引数数が 1 つでない場合、または数値以外が指定された場合。
    """
    if len(args) != 1:
        raise TypeError("modi: 引数は1個 (d) を指定してください。")
    value = ensure_number(args[0], "modi")
    return BuiltinNumericResult(math.trunc(value))


def modd_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """小数部を正値として切り出す。

    Args:
        args: 引数リスト。1 つの数値を想定。

    Returns:
        `BuiltinNumericResult`: 小数部とフォーマットヒント。

    Raises:
        TypeError: 引数数が 1 つでない場合、または数値以外が指定された場合。
    """
    if len(args) != 1:
        raise TypeError("modd: 引数は1個 (d) を指定してください。")
    numeric = ensure_number(args[0], "modd")
    dec_value = to_decimal(numeric)
    integer_part = Decimal(math.trunc(numeric))
    fraction_dec = (dec_value - integer_part).copy_abs()
    exponent = int(fraction_dec.as_tuple().exponent)
    digits = max(0, -exponent)
    fraction = float(fraction_dec)
    format_hint = f"fixed:{digits}" if digits > 0 else None
    return BuiltinNumericResult(normalize_number(fraction), format_hint)


def round_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """Lab-Aid 互換の四捨五入処理を行う。

    Args:
        args: `(d, x, y [, z])` 形式の引数リスト。

    Returns:
        `BuiltinNumericResult`: 丸め結果とフォーマットヒント。

    Raises:
        TypeError: 引数数や型が仕様に合致しない場合。
        ValueError: 丸め単位や桁数が不正な場合。
    """
    if len(args) not in (3, 4):
        raise TypeError(
            "round: 引数は3個または4個 (d, x, y [, z]) を指定してください。"
        )
    dec = to_decimal(ensure_number(args[0], "round"))
    x = ensure_int(args[1], "round", "x")
    y = ensure_int(args[2], "round", "y")
    if y not in (0, 1):
        raise ValueError("round: 第3引数 y には 0 または 1 を指定してください。")
    z = 1 if len(args) == 3 else ensure_number(args[3], "round")
    if z <= 0:
        raise ValueError("round: 丸め単位 z は正の数を指定してください。")

    if y == 0:
        if dec.is_zero():
            quantized = Decimal(0)
        else:
            exponent = dec.adjusted() - x + 1
            quantum = Decimal(1).scaleb(exponent)
            quantized = dec.quantize(quantum, rounding=ROUND_HALF_UP)
        value = force_int_if_integral(quantized, ROUND_HALF_UP)
        format_hint = f"sig:{x}" if x > 0 else None
        return BuiltinNumericResult(value, format_hint)

    step = to_decimal(z) * Decimal(1).scaleb(-x)
    if step.is_zero():
        raise ValueError("round: 丸め単位の指定が不正です。")
    scaled = dec / step
    rounded = scaled.quantize(Decimal(1), rounding=ROUND_HALF_UP)
    quantized = (rounded * step).quantize(step, rounding=ROUND_HALF_UP)
    value = force_int_if_integral(quantized, ROUND_HALF_UP)
    format_hint = f"fixed:{x}" if x >= 0 else None
    return BuiltinNumericResult(value, format_hint)


def roundjisb_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """JIS B 互換の四捨五入処理を行う。

    Args:
        args: `(d, p, f)` 形式の引数リスト。

    Returns:
        `BuiltinNumericResult`: 丸め結果とフォーマットヒント。

    Raises:
        TypeError: 引数数や型が仕様に合致しない場合。
    """
    if len(args) != 3:
        raise TypeError("roundjisb: 引数は3個 (d, p, f) を指定してください。")
    d, p, f = args
    numeric = ensure_number(d, "roundjisb")
    dec_value = to_decimal(numeric)
    if isinstance(p, float) and p.is_integer():
        p = int(p)
    if not isinstance(p, int):
        raise TypeError("roundjisb: 第2引数 p は整数である必要があります。")
    if not isinstance(f, (int, float)) or not float(f).is_integer():
        raise TypeError("roundjisb: 第3引数 f は 0/1 の整数である必要があります。")
    f = int(f)
    if f == 1:
        quantized = dec_value.quantize(Decimal(f"1e-{p}"), rounding=ROUND_HALF_UP)
        value = force_int_if_integral(quantized, ROUND_HALF_UP)
        return BuiltinNumericResult(normalize_number(value), f"fixed:{p}")
    if dec_value.is_zero():
        quantized = Decimal(0)
    else:
        exp10 = dec_value.adjusted()
        quant_exp = exp10 - p + 1
        quantized = dec_value.quantize(
            Decimal(f"1e{quant_exp}"), rounding=ROUND_HALF_UP
        )
    value = force_int_if_integral(quantized, ROUND_HALF_UP)
    return BuiltinNumericResult(normalize_number(value), f"sig:{p}")


def floor_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """指定桁での切り捨てを行う。

    Args:
        args: `(d, p)` 形式の引数リスト。

    Returns:
        `BuiltinNumericResult`: 切り捨て後の値。

    Raises:
        TypeError: 引数数が 2 つでない場合、または型が不正な場合。
        TypeError: 丸め処理に失敗した場合。
    """
    if len(args) != 2:
        raise TypeError("floor: 引数は2個 (d, p) を指定してください。")
    d, p = args
    if not isinstance(d, (int, float)):
        raise TypeError("floor: 第1引数 d は数値である必要があります。")
    if isinstance(p, float) and p.is_integer():
        p = int(p)
    if not isinstance(p, int):
        raise TypeError("floor: 第2引数 p は整数である必要があります。")
    dx = to_decimal(d)
    try:
        if p >= 0:
            quantized = dx.quantize(Decimal(f"1e-{p}"), rounding=ROUND_FLOOR)
        else:
            quantized = dx.quantize(Decimal(f"1e{-p}"), rounding=ROUND_FLOOR)
    except InvalidOperation as exc:
        raise TypeError("floor: 丸めに失敗しました。") from exc
    return BuiltinNumericResult(force_int_if_integral(quantized, ROUND_HALF_EVEN))


def trunc_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """指定桁での絶対値切り捨てを行う。

    Args:
        args: `(d, p)` 形式の引数リスト。

    Returns:
        `BuiltinNumericResult`: 切り捨て結果。

    Raises:
        TypeError: 引数数が 2 つでない場合、または型が不正な場合。
    """
    if len(args) != 2:
        raise TypeError("trunc: 引数は2個 (d, p) を指定してください。")
    d, p = args
    if not isinstance(d, (int, float)):
        raise TypeError("trunc: 第1引数 d は数値である必要があります。")
    if isinstance(p, float) and p.is_integer():
        p = int(p)
    if not isinstance(p, int):
        raise TypeError("trunc: 第2引数 p は整数である必要があります。")
    sign = -1 if d < 0 else 1
    dx = to_decimal(abs(d))
    if p >= 0:
        quantized = dx.quantize(Decimal(f"1e-{p}"), rounding=ROUND_DOWN)
    else:
        quantized = dx.quantize(Decimal(f"1e{-p}"), rounding=ROUND_DOWN)
    if sign < 0:
        quantized = quantized.copy_negate()
    return BuiltinNumericResult(force_int_if_integral(quantized, ROUND_HALF_EVEN))


def max_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """最大値を返す。

    Args:
        args: 数値または数値リストを含む引数リスト。

    Returns:
        `BuiltinNumericResult`: 最大値。

    Raises:
        TypeError: 引数が空、または非数値を含む場合。
    """
    if not args:
        raise TypeError("max: 引数を1つ以上指定してください。")
    if len(args) == 1:
        numbers = collect_numeric_values(args[0], "max")
        return BuiltinNumericResult(normalize_number(max(numbers)))
    best: float | None = None
    for arg in args:
        values = collect_numeric_values(arg, "max")
        candidate = sum(values) / len(values)
        if best is None or candidate > best:
            best = candidate
    if best is None:
        raise TypeError("max: 比較可能な数値がありません。")
    return BuiltinNumericResult(normalize_number(best))


def min_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """最小値を返す。

    Args:
        args: 数値または数値リストを含む引数リスト。

    Returns:
        `BuiltinNumericResult`: 最小値。

    Raises:
        TypeError: 引数が空、または非数値を含む場合。
    """
    if not args:
        raise TypeError("min: 引数を1つ以上指定してください。")
    if len(args) == 1:
        numbers = collect_numeric_values(args[0], "min")
        return BuiltinNumericResult(normalize_number(min(numbers)))
    best: float | None = None
    for arg in args:
        values = collect_numeric_values(arg, "min")
        candidate = sum(values) / len(values)
        if best is None or candidate < best:
            best = candidate
    if best is None:
        raise TypeError("min: 比較可能な数値がありません。")
    return BuiltinNumericResult(normalize_number(best))


def ave_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """平均値を返す。

    Args:
        args: 数値または数値リストを含む引数リスト。

    Returns:
        `BuiltinNumericResult`: 平均値。

    Raises:
        TypeError: 引数が空、または非数値を含む場合。
    """
    if not args:
        raise TypeError("ave: 引数を1つ以上指定してください。")
    values: list[float] = []
    for arg in args:
        values.extend(collect_numeric_values(arg, "ave"))
    result = sum(values) / len(values)
    return BuiltinNumericResult(normalize_number(result))


def sum_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """合計値を返す。

    Args:
        args: 数値または数値リストを含む引数リスト。

    Returns:
        `BuiltinNumericResult`: 合計値。

    Raises:
        TypeError: 引数が空、または非数値を含む場合。
    """
    if not args:
        raise TypeError("sum: 引数を1つ以上指定してください。")
    total = 0.0
    for arg in args:
        total += sum(collect_numeric_values(arg, "sum"))
    return BuiltinNumericResult(normalize_number(total))


def stdev_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """標本標準偏差を計算する。

    Args:
        args: 数値または数値リストを含む引数リスト。

    Returns:
        `BuiltinNumericResult`: 標本標準偏差。

    Raises:
        TypeError: 引数が不足している、または非数値を含む場合。
        ValueError: 値の総数が 2 未満の場合。
    """
    if not args:
        raise TypeError("stdev: 引数を2つ以上指定してください。")
    values: list[float] = []
    for arg in args:
        values.extend(collect_numeric_values(arg, "stdev"))
    if len(values) < 2:
        raise ValueError("stdev: 少なくとも2つの値が必要です。")
    return BuiltinNumericResult(normalize_number(statistics.stdev(values)))


def stdeva_func(args: Sequence[Any]) -> BuiltinNumericResult:
    """母標準偏差（Lab-Aid 仕様では `pstdev`）を計算する。

    Args:
        args: 数値または数値リストを含む引数リスト。

    Returns:
        `BuiltinNumericResult`: 母標準偏差。

    Raises:
        TypeError: 引数が不足している、または非数値を含む場合。
        ValueError: 値の総数が 2 未満の場合。
    """
    if not args:
        raise TypeError("stdeva: 引数を2つ以上指定してください。")
    values: list[float] = []
    for arg in args:
        values.extend(collect_numeric_values(arg, "stdeva"))
    if len(values) < 2:
        raise ValueError("stdeva: 少なくとも2つの値が必要です。")
    return BuiltinNumericResult(normalize_number(statistics.pstdev(values)))


NUMERIC_FUNCTIONS = {
    "sqrt": sqrt_func,
    "log": log_func,
    "log10": log10_func,
    "exp": exp_func,
    "pow": pow_func,
    "modi": modi_func,
    "modd": modd_func,
    "round": round_func,
    "roundjisb": roundjisb_func,
    "floor": floor_func,
    "trunc": trunc_func,
    "max": max_func,
    "min": min_func,
    "ave": ave_func,
    "sum": sum_func,
    "stdev": stdev_func,
    "stdeva": stdeva_func,
}


def _format_fixed_frac(dec: Decimal, digits: int) -> str:
    """固定小数点形式で文字列整形する。

    Args:
        dec: 対象となる `Decimal` 値。
        digits: 少数部の桁数。

    Returns:
        少数部が `digits` 桁になるよう整形した文字列。
    """
    quantized = dec.quantize(Decimal(f"1e-{digits}"), rounding=ROUND_HALF_UP)
    return f"{quantized:.{digits}f}"


def _format_sig_digits(dec: Decimal, sig: int) -> str:
    """有効数字形式で文字列整形する。

    Args:
        dec: 対象となる `Decimal` 値。
        sig: 有効数字の桁数。

    Returns:
        有効数字 `sig` 桁で表現された文字列。
    """
    if dec.is_zero():
        if sig <= 1:
            return "0"
        return "0." + ("0" * (sig - 1))
    exp10 = dec.adjusted()
    quant_exp = exp10 - sig + 1
    quantized = dec.quantize(Decimal(f"1e{quant_exp}"), rounding=ROUND_HALF_UP)
    if quant_exp >= 0:
        return f"{quantized:f}"
    frac = -quant_exp
    return f"{quantized:.{frac}f}"


def format_roundjisb_output(value: Any, format_hint: str | None) -> str | None:
    """`roundjisb` 系のフォーマットヒントに従い文字列化する。

    Args:
        value: 整形対象の値。
        format_hint: `fixed:n` または `sig:n` 形式のヒント。`None` なら変換しない。

    Returns:
        整形済み文字列。ヒントが無効な場合は `None`。
    """
    if not format_hint:
        return None
    try:
        dec_value = to_decimal(value)
    except (InvalidOperation, TypeError, ValueError):
        return None
    if format_hint.startswith("fixed:"):
        digits = int(format_hint.split(":", 1)[1])
        return _format_fixed_frac(dec_value, digits)
    if format_hint.startswith("sig:"):
        digits = int(format_hint.split(":", 1)[1])
        return _format_sig_digits(dec_value, digits)
    return None


__all__ = ["NUMERIC_FUNCTIONS", "format_roundjisb_output"]
