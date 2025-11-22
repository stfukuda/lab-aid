"""Lab-Aid の計算スクリプトを実行する公開 API。"""

from __future__ import annotations

from .engine_core import Engine
from .inputs import (
    VarRef,
    ensure_has_this_assignment_E,
    parse_input_R,
    parse_inputs_E,
    replace_rhs_this_for_R,
    RE_ITEM_ANY,
)
from .text import strip_comment_quote_aware, to_text


def assert_no_hash_usage(text: str, where: str) -> None:
    """禁止箇所に `#CODE` 形式の項目参照が含まれていないことを検証する。

    コメントや単一引用符で囲まれたリテラル内の `#` は許容し、実際の
    `#CODE` 記法のみを検出する。

    Args:
        text: 検査対象となる文字列。
        where: エラーメッセージに表示する対象箇所の説明。

    Raises:
        ValueError: コメント／リテラル外で `#CODE` 記法が見つかった場合。
    """

    for raw_line in text.splitlines():
        line = strip_comment_quote_aware(raw_line)
        if not line:
            continue
        if line.lstrip().lower().startswith("rem"):
            continue
        if _contains_hash_reference(line):
            raise ValueError(f"Rタイプでは {where} に #変数は使用できません。")


def _contains_hash_reference(line: str) -> bool:
    """単一引用符外で `#CODE` 記法が存在するか判定する。"""

    index = 0
    length = len(line)
    in_sq = False
    while index < length:
        char = line[index]
        if char == "'":
            if in_sq and index + 1 < length and line[index + 1] == "'":
                index += 2
                continue
            in_sq = not in_sq
            index += 1
            continue
        if not in_sq and char == "#":
            match = RE_ITEM_ANY.match(line, index)
            if match:
                return True
        index += 1
    return False


def evaluate(
    calc_type: str,
    script: str,
    inputs: str,
) -> tuple[str | None, str | None, str | None]:
    """Lab-Aid の推定計算（E）または丸め計算（R）を評価する。

    E タイプは試験項目の値を `this` へ代入するスクリプト、R タイプは丸め条件を
    評価するスクリプトを入力として受け取る。戻り値はいずれも Lab-Aid 互換の
    文字列表現を想定しており、例外発生時は仕様に従ったエラー文字列を返却する。

    Args:
        calc_type: "E" または "R" を示す計算種別。
        script: Lab-Aid 形式で記述された計算スクリプト。複数行を許可。
        inputs: E タイプでは `NAME=VALUE` 形式、R タイプでは単一値の入力文字列。

    Returns:
        E タイプの場合は `(raw_text, edited_text, reported_text)` のタプル。
        R タイプの場合は `(None, edited_text, reported_text)` のタプル。
        いずれもエラー発生時は仕様どおり `"エラー"` を含むタプルを返す。

    Raises:
        なし。入力不備は Lab-Aid 互換のエラー文字列として呼び出し元へ返却される。
    """

    try:
        ctype = (calc_type or "").strip().upper()
        if ctype not in ("E", "R"):
            raise ValueError("calc_type は 'E' または 'R' を指定してください。")

        if ctype == "E":
            ensure_has_this_assignment_E(script)
            items = parse_inputs_E(inputs)
            engine = Engine(items=items, vars={"this": 0})
            vars_after = engine.run_lines(script.splitlines())

            if engine.this_assigned_count == 0:
                raise ValueError("Eタイプでは this= が必須です。")

            this_value = vars_after.get("this")
            raw_text = (
                engine.this_formatted
                if engine.this_formatted is not None
                else to_text(this_value)
            )
            edited_text = engine.last_print
            reported_text = engine.last_print2
            return raw_text, edited_text, reported_text

        assert_no_hash_usage(script, "第2引数（計算式）")
        assert_no_hash_usage(inputs, "第3引数（入力値）")

        this_in_raw, literal = parse_input_R(inputs)
        if isinstance(this_in_raw, VarRef):
            this_initial = 0
            placeholder_value = 0
        else:
            this_initial = this_in_raw
            placeholder_value = this_in_raw

        runnable_script = replace_rhs_this_for_R(script)
        engine = Engine(
            items={},
            vars={"this": this_initial, "__THIS_IN__": placeholder_value},
        )
        vars_after = engine.run_lines(runnable_script.splitlines())

        edited_text = None
        reported_text = None

        if engine.last_print is not None:
            edited_text = engine.last_print
            reported_text = engine.last_print
        elif engine.this_assigned_count > 0:
            this_value = vars_after.get("this")
            base_text = (
                engine.this_formatted
                if engine.this_formatted is not None
                else to_text(this_value)
            )
            edited_text = base_text
            reported_text = base_text
        else:
            edited_text = literal
            reported_text = literal

        if engine.last_print2 is not None:
            reported_text = engine.last_print2

        return None, edited_text, reported_text

    except Exception:
        if (calc_type or "").strip().upper() == "R":
            return None, "エラー", "エラー"
        return "エラー", None, None


__all__ = ["evaluate", "assert_no_hash_usage"]
