"""Lab-Aid の文字列系ビルトイン関数群。"""

from __future__ import annotations

from collections.abc import Sequence
from typing import Any

from .base import BuiltinNumericResult, ensure_int, select_value


def _shift_jis_length(char: str, func: str) -> int:
    """Shift_JIS でのバイト長を算出する。

    Args:
        char: 対象となる 1 文字。
        func: エラーメッセージに使用する関数名。

    Returns:
        Shift_JIS でのバイト長。

    Raises:
        ValueError: Shift_JIS でエンコードできない文字の場合。
    """
    try:
        return len(char.encode("shift_jis"))
    except UnicodeEncodeError as exc:
        raise ValueError(
            f"{func}: Shift_JIS で表現できない文字が含まれています。"
        ) from exc


def _slice_shift_jis(src: str, start: int, length: int | None, func: str) -> str:
    """Shift_JIS バイト長基準で文字列を切り出す。

    Args:
        src: 元となる文字列。
        start: 1 始まりの開始位置。
        length: 取得するバイト長。``None`` の場合は末尾まで取得。
        func: エラーメッセージに使用する関数名。

    Returns:
        指定範囲を切り出した文字列。

    Raises:
        ValueError: 開始位置が 1 未満、または切り出し長が負の場合。
    """
    if start < 1:
        raise ValueError(f"{func}: start には 1 以上の整数を指定してください。")
    if not src:
        return ""
    char_index = start - 1
    if char_index >= len(src):
        return ""
    remainder = src[char_index:]
    if length is None:
        return remainder
    if length == 0:
        return ""
    copied: list[str] = []
    used = 0
    for char in remainder:
        byte_len = _shift_jis_length(char, func)
        if used + byte_len > length:
            break
        copied.append(char)
        used += byte_len
    return "".join(copied)


def _select_text(value: Any, index: int | None) -> str:
    """文字列またはリストから表示用テキストを取得する。

    Args:
        value: 単一値または値リスト。
        index: 1 始まりのインデックス。``None`` の場合は先頭要素を使用。

    Returns:
        取得した文字列。値が ``None`` の場合は空文字列。
    """
    selected = select_value(value, index)
    if selected is None:
        return ""
    return selected if isinstance(selected, str) else f"{selected}"


def str_comp(args: Sequence[Any]) -> BuiltinNumericResult:
    """Lab-Aid 仕様の文字列比較を行う。

    Args:
        args: 2〜4 個の引数。第 3、第 4 引数は試験回指定。

    Returns:
        `BuiltinNumericResult`: 比較結果（左<右:-1、等しい:0、左>右:1）。

    Raises:
        TypeError: 引数構成が仕様に合致しない場合。
    """
    if len(args) not in (2, 3, 4):
        raise TypeError("str_comp: 引数数が不正です。2 / 3 / 4 個を指定してください。")

    index_left: int | None = None
    index_right: int | None = None

    if len(args) == 2:
        left_raw, right_raw = args
    elif len(args) == 3:
        left_raw, index_raw, right_raw = args
        index_left = ensure_int(index_raw, "str_comp", "n")
    else:  # len(args) == 4
        left_raw, right_raw, index1_raw, index2_raw = args
        index_left = ensure_int(index1_raw, "str_comp", "n1")
        index_right = ensure_int(index2_raw, "str_comp", "n2")

    left = _select_text(left_raw, index_left)
    right = _select_text(right_raw, index_right)

    if left == right:
        result = 0
    elif left < right:
        result = -1
    else:
        result = 1

    return BuiltinNumericResult(result)


def is_char(args: Sequence[Any]) -> BuiltinNumericResult:
    """値が文字列かどうかを判定する。

    Args:
        args: 1 または 2 個の引数。第 2 引数は試験回指定。

    Returns:
        `BuiltinNumericResult`: 文字列なら 1、それ以外は 0。

    Raises:
        TypeError: 引数構成が仕様に合致しない場合。
    """
    if len(args) not in (1, 2):
        raise TypeError("is_char: 引数は1個または2個 (#T [, n]) を指定してください。")
    index = ensure_int(args[1], "is_char", "n") if len(args) == 2 else None
    value = select_value(args[0], index)
    if value is None:
        return BuiltinNumericResult(1)
    return BuiltinNumericResult(1 if isinstance(value, str) else 0)


def strlen(args: Sequence[Any]) -> BuiltinNumericResult:
    """文字列長を返す。

    Args:
        args: 1 または 2 個の引数。第 2 引数は試験回指定。

    Returns:
        `BuiltinNumericResult`: 文字列長。

    Raises:
        TypeError: 引数構成が仕様に合致しない場合。
    """
    if len(args) not in (1, 2):
        raise TypeError("strlen: 引数は1個または2個 (d [, n]) を指定してください。")
    index = ensure_int(args[1], "strlen", "n") if len(args) == 2 else None
    text = _select_text(args[0], index)
    return BuiltinNumericResult(len(text))


def strcat(args: Sequence[Any]) -> BuiltinNumericResult:
    """文字列を連結する。

    Args:
        args: 2 個の引数。左辺・右辺のテキスト。

    Returns:
        `BuiltinNumericResult`: 結合後の文字列。

    Raises:
        TypeError: 引数数が 2 個でない場合。
    """
    if len(args) != 2:
        raise TypeError("strcat: 引数は2個 (a, b) を指定してください。")
    left = _select_text(args[0], None)
    right = _select_text(args[1], None)
    return BuiltinNumericResult(f"{left}{right}")


def strncpy(args: Sequence[Any]) -> BuiltinNumericResult:
    """Shift_JIS バイト長で部分文字列を取得する。

    Args:
        args: 2〜4 個の引数 `(dest, src [, start [, len]])`。

    Returns:
        `BuiltinNumericResult`: 切り出した文字列。

    Raises:
        TypeError: 引数数が 2〜4 個でない場合、または型が不正な場合。
        ValueError: `len` に負の値が指定された場合。
    """
    if len(args) < 2 or len(args) > 4:
        raise TypeError(
            "strncpy: 引数は2〜4個 (dest, src [, start [, len]]) を指定してください。"
        )
    src = _select_text(args[1], None)
    start = 1
    length: int | None = None
    if len(args) >= 3:
        start = ensure_int(args[2], "strncpy", "start")
    if len(args) == 4:
        length = ensure_int(args[3], "strncpy", "len")
        if length < 0:
            raise ValueError("strncpy: len には 0 以上の整数を指定してください。")
    return BuiltinNumericResult(_slice_shift_jis(src, start, length, "strncpy"))


def isempty(args: Sequence[Any]) -> BuiltinNumericResult:
    """値が空かどうかを判定する。

    Args:
        args: 1 または 2 個の引数。第 2 引数は試験回指定。

    Returns:
        `BuiltinNumericResult`: 空文字列または `None` の場合は 1、それ以外は 0。

    Raises:
        TypeError: 引数構成が仕様に合致しない場合。
    """
    if len(args) not in (1, 2):
        raise TypeError("isempty: 引数は1個または2個 (d [, n]) を指定してください。")
    index = ensure_int(args[1], "isempty", "n") if len(args) == 2 else None
    value = select_value(args[0], index)
    if value is None:
        return BuiltinNumericResult(1)
    if isinstance(value, str):
        return BuiltinNumericResult(1 if value == "" else 0)
    return BuiltinNumericResult(0)


def isspace(args: Sequence[Any]) -> BuiltinNumericResult:
    """値が空白文字のみで構成されているか判定する。

    Args:
        args: 1 または 2 個の引数。第 2 引数は試験回指定。

    Returns:
        `BuiltinNumericResult`: 空白のみなら 1、それ以外は 0。

    Raises:
        TypeError: 引数構成が仕様に合致しない場合。
    """
    if len(args) not in (1, 2):
        raise TypeError("isspace: 引数は1個または2個 (d [, n]) を指定してください。")
    index = ensure_int(args[1], "isspace", "n") if len(args) == 2 else None
    value = select_value(args[0], index)
    if not isinstance(value, str):
        return BuiltinNumericResult(0)
    return BuiltinNumericResult(1 if value != "" and value.isspace() else 0)


STRING_FUNCTIONS = {
    "str_comp": str_comp,
    "is_char": is_char,
    "strlen": strlen,
    "strcat": strcat,
    "strncpy": strncpy,
    "isempty": isempty,
    "isspace": isspace,
}


__all__ = [
    "STRING_FUNCTIONS",
    "str_comp",
    "is_char",
    "strlen",
    "strcat",
    "strncpy",
    "isempty",
    "isspace",
]
