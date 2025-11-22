"""Lab-Aid 実行エンジンで利用する文字列・式操作ユーティリティ。"""

from __future__ import annotations

import re
import string
from typing import Any

_INT_RE = re.compile(r"^[+-]?\d+$")
_FLOAT_RE = re.compile(
    r"^[+-]?(?:\d+\.\d*|\.\d+|\d+\.)(?:[eE][+-]?\d+)?$|^[+-]?\d+(?:[eE][+-]?\d+)$"
)
_WORD_CHARS = string.ascii_letters + string.digits + "_"


def parse_number_like(value: str) -> Any | None:
    """文字列を整数または浮動小数点数として解釈できれば数値に変換する。

    Args:
        value: 数値解釈を試みる文字列。

    Returns:
        整数または浮動小数点数として解釈できた場合はその数値。解釈できなければ ``None``。
    """
    token = value.strip()
    if _INT_RE.match(token):
        try:
            return int(token)
        except Exception:
            return None
    if _FLOAT_RE.match(token):
        try:
            return float(token)
        except Exception:
            return None
    return None


def unquote_single(value: str) -> str:
    """単一引用符で囲まれた文字列から外側の引用符を除去し、内部の二重引用符を復元する。

    Args:
        value: 単一引用符を含む可能性がある文字列。

    Returns:
        外側の単一引用符を外し、連続した ``''`` を単一の `'` に戻した文字列。
    """
    token = value.strip()
    if len(token) >= 2 and token[0] == "'" and token[-1] == "'":
        return token[1:-1].replace("''", "'")
    return token


def to_text(value: Any) -> str | None:
    """値を Lab-Aid が理解できる文字列表現へ変換する。

    Args:
        value: 文字列化対象の値。

    Returns:
        文字列をそのまま、数値等は `str()` 相当で文字列化した結果。
        値が ``None`` の場合は ``None`` を返す。
    """
    if value is None:
        return None
    return value if isinstance(value, str) else f"{value}"


def strip_comment_quote_aware(line: str) -> str:
    """単一引用符内を除き、`;` 以降のコメントを取り除く。

    Args:
        line: Lab-Aid スクリプトの 1 行。

    Returns:
        コメントを除去し、前後の空白を取り除いた文字列。
    """
    line = line.rstrip("\n")
    out: list[str] = []
    in_sq = False
    index = 0
    while index < len(line):
        char = line[index]
        if char == "'":
            if in_sq and index + 1 < len(line) and line[index + 1] == "'":
                out.append("''")
                index += 2
                continue
            in_sq = not in_sq
            out.append(char)
            index += 1
            continue
        if char == ";" and not in_sq:
            break
        out.append(char)
        index += 1
    return "".join(out).strip()


def replace_word_ci_outside_quotes(text: str, mapping: dict[str, str]) -> str:
    """単一引用符の外側で、単語単位の大小文字非区別置換を行う。

    Args:
        text: 置換対象となる文字列。
        mapping: 小文字比較で一致させる置換ルール。

    Returns:
        置換後の文字列。
    """
    result: list[str] = []
    index = 0
    length = len(text)
    in_sq = False
    keys = sorted(mapping.keys(), key=len, reverse=True)
    while index < length:
        char = text[index]
        if char == "'":
            if in_sq and index + 1 < length and text[index + 1] == "'":
                result.append("''")
                index += 2
                continue
            in_sq = not in_sq
            result.append(char)
            index += 1
            continue
        if not in_sq:
            matched = False
            for key in keys:
                span = len(key)
                if text[index : index + span].lower() == key:
                    prev_ok = index == 0 or text[index - 1] not in _WORD_CHARS
                    next_ok = (
                        index + span == length or text[index + span] not in _WORD_CHARS
                    )
                    if prev_ok and next_ok:
                        result.append(mapping[key])
                        index += span
                        matched = True
                        break
            if matched:
                continue
        result.append(char)
        index += 1
    return "".join(result)


def validate_hash_name(code: str, unit: str | None) -> None:
    """Lab-Aid の `#` 変数表記が仕様に合致しているかを検証する。

    Args:
        code: 試験項目コード。
        unit: 単位を表す文字列。指定しない場合は ``None``。

    Raises:
        ValueError: コードまたは単位が仕様外の文字で構成されていた場合。
    """
    if not re.fullmatch(r"[A-Z0-9_]+(?:\.[A-Z0-9_]+)?", code or ""):
        raise ValueError(
            f"#変数名は大文字英字・数字・'_'（必要ならドット1個で区切り）のみ: #{code}"
        )
    if unit is not None and not re.fullmatch(r"[A-Z0-9_]+", unit):
        raise ValueError(f"#変数の単位名も大文字英字・数字・'_' のみ: #{code}[{unit}]")
