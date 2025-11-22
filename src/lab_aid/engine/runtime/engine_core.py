"""Lab-Aid スクリプトを実行する中核エンジン。"""

from __future__ import annotations

import ast
import re
from collections.abc import Iterable
from dataclasses import dataclass, field
from typing import Any

from .constants import MAX_FOR_ITERS, MAX_NEST_DEPTH
from .functions import (
    NUMERIC_FUNCTIONS,
    STRING_FUNCTIONS,
    BuiltinNumericResult,
    format_roundjisb_output,
)
from .functions.package import execute_print, execute_print2
from .inputs import RE_ITEM_ANY, VarRef
from .text import (
    parse_number_like,
    replace_word_ci_outside_quotes,
    strip_comment_quote_aware,
    to_text,
    validate_hash_name,
)

PRINT_RE = re.compile(r"^\s*print\s*\(\s*this\s*,\s*(.+)\)\s*$", re.IGNORECASE)
PRINT2_RE = re.compile(r"^\s*print2\s*\(\s*this\s*,\s*(.+)\)\s*$", re.IGNORECASE)

FUNCTION_DISPATCH = {**NUMERIC_FUNCTIONS, **STRING_FUNCTIONS}
STATEMENT_FUNCTIONS = {"strcat", "strncpy"}

OP_MAPPING = {
    "eq": "==",
    "ne": "!=",
    "gt": ">",
    "ge": ">=",
    "lt": "<",
    "le": "<=",
    "and": "and",
    "or": "or",
}


class FormatAwareNumber(float):
    """フォーマット済み文字列表現を保持する数値ラッパー。"""

    def __new__(cls, value: float, formatted: str | None):
        obj = float.__new__(cls, value)
        obj.formatted = formatted
        return obj

    def __str__(self) -> str:
        formatted = getattr(self, "formatted", None)
        return formatted if formatted is not None else super().__str__()


@dataclass
class Engine:
    """Lab-Aid のスクリプトを評価するエンジン。

    Attributes:
        items: `#CODE[UNIT]` 参照に対応する試験項目の値マップ。
        vars: 通常変数と `this` を保持する可変辞書。
        var_formats: 変数毎の表示ヒントとフォーマット済み文字列。
        last_format_hint: 直近の式評価で得られたフォーマットヒント。
        this_formatted: `this` に対して適用されたフォーマット済み文字列。
        last_print: `print` によって最後に出力された文字列。
        last_print2: `print2` によって最後に出力された文字列。
        this_assigned_count: `this` への代入回数。E タイプでは 1 以上が要求される。
    """

    items: dict[Any, Any] = field(default_factory=dict)
    vars: dict[str, Any] = field(default_factory=lambda: {"this": 0})
    var_formats: dict[str, tuple[str | None, str | None]] = field(default_factory=dict)

    last_format_hint: str | None = None
    this_formatted: str | None = None
    last_print: str | None = None
    last_print2: str | None = None
    this_assigned_count: int = 0

    @staticmethod
    def _coerce_numeric(value: Any) -> int | float | None:
        """大小比較用に値を数値へ変換する。"""
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, str):
            parsed = parse_number_like(value)
            if isinstance(parsed, (int, float)):
                return parsed
        return None

    def _format_to_text(self, value: Any) -> str | None:
        """直近のフォーマットヒントを考慮して文字列化する。"""
        formatted = format_roundjisb_output(value, self.last_format_hint)
        if formatted is not None:
            self.last_format_hint = None
            return formatted
        return to_text(value)

    def eval_expr(self, expr: str) -> Any:
        """Lab-Aid 互換の式文字列を評価する。

        Lab-Aid 特有の `#CODE[UNIT]` 表記や大小比較演算子を Python AST に変換し、
        安全に評価した結果を返す。評価途中で取得したフォーマットヒントは
        `last_format_hint` に保持する。

        Args:
            expr: Lab-Aid 互換の式文字列。

        Returns:
            式の評価結果。数値・文字列・真偽値などを返却する。

        Raises:
            SyntaxError: 構文解析に失敗した場合。
            TypeError: サポート外の値型や演算子が含まれていた場合。
            ValueError: `#` 変数名が Lab-Aid 仕様に違反していた場合。
            KeyError: 未定義の試験項目を参照した場合。
        """
        self.last_format_hint = None

        def build_quote_mask(text: str) -> list[bool]:
            mask = [False] * len(text)
            in_sq = False
            index = 0
            while index < len(text):
                char = text[index]
                if char == "'":
                    if in_sq and index + 1 < len(text) and text[index + 1] == "'":
                        mask[index] = True
                        mask[index + 1] = True
                        index += 2
                        continue
                    in_sq = not in_sq
                    mask[index] = True
                else:
                    mask[index] = in_sq
                index += 1
            return mask

        quote_mask = build_quote_mask(expr)
        matches: list[tuple[int, int, re.Match]] = []
        for match in RE_ITEM_ANY.finditer(expr):
            start, end = match.span()
            if any(quote_mask[start:end]):
                continue
            code, unit = match.group(1), match.group(2)
            validate_hash_name(code, unit)
            matches.append((start, end, match))

        subst_map: dict[str, Any] = {}
        rewritten_parts: list[str] = []
        last_index = 0

        for start, end, match in matches:
            code = match.group(1)
            unit = match.group(2)
            value = self.resolve_item(code, unit)
            placeholder = f"__item_{code}__{unit}" if unit else f"__item_{code}"
            placeholder = re.sub(r"[^A-Za-z0-9_]", "_", placeholder).lower()
            subst_map[placeholder] = value
            rewritten_parts.append(expr[last_index:start])
            rewritten_parts.append(placeholder)
            last_index = end

        rewritten_parts.append(expr[last_index:])
        rewritten = "".join(rewritten_parts)
        rewritten = replace_word_ci_outside_quotes(rewritten, OP_MAPPING)

        try:
            node = ast.parse(rewritten, mode="eval")
        except SyntaxError as exc:
            raise SyntaxError(f"式の構文エラー: {expr}") from exc

        names: dict[str, Any] = {}
        for key, value in self.vars.items():
            canonical = "this" if key.lower() == "this" else key
            names[canonical] = value
        names.update(subst_map)

        return self.eval_ast(node.body, names)

    def eval_ast(self, node: ast.AST, names: dict[str, Any]) -> Any:
        """AST ノードを再帰的に評価する。

        Args:
            node: 評価対象の AST ノード。
            names: 変数名と値を紐付けた辞書。

        Returns:
            ノードに対応する評価結果。

        Raises:
            TypeError: サポート外のノード種別や型が検出された場合。
            NameError: 未定義の通常変数を参照した場合。
            RuntimeError: IF 条件式等で例外をラップする際に利用。
        """
        if isinstance(node, ast.BinOp):
            self.last_format_hint = None
            left_val = self.eval_ast(node.left, names)
            right_val = self.eval_ast(node.right, names)
            if not isinstance(left_val, (int, float)) or not isinstance(
                right_val, (int, float)
            ):
                raise TypeError(
                    f"数値でない値に四則演算は適用不可: {left_val!r}, {right_val!r}"
                )
            if isinstance(node.op, ast.Add):
                return left_val + right_val
            if isinstance(node.op, ast.Sub):
                return left_val - right_val
            if isinstance(node.op, ast.Mult):
                return left_val * right_val
            if isinstance(node.op, ast.Div):
                return left_val / right_val
            raise TypeError(f"未対応の演算子: {type(node.op).__name__}")

        if isinstance(node, ast.UnaryOp):
            self.last_format_hint = None
            operand = self.eval_ast(node.operand, names)
            if not isinstance(operand, (int, float)):
                raise TypeError(f"数値でない値に単項演算は適用不可: {operand!r}")
            if isinstance(node.op, ast.UAdd):
                return +operand
            if isinstance(node.op, ast.USub):
                return -operand
            raise TypeError(f"未対応の単項演算子: {type(node.op).__name__}")

        if isinstance(node, ast.BoolOp):
            self.last_format_hint = None
            if isinstance(node.op, ast.And):
                return all(bool(self.eval_ast(value, names)) for value in node.values)
            if isinstance(node.op, ast.Or):
                return any(bool(self.eval_ast(value, names)) for value in node.values)
            raise TypeError("未対応の論理演算子")

        if isinstance(node, ast.Compare):
            self.last_format_hint = None
            left = self.eval_ast(node.left, names)
            for op, comparator in zip(node.ops, node.comparators, strict=True):
                right = self.eval_ast(comparator, names)
                if isinstance(op, (ast.Eq, ast.NotEq)):
                    ok = (left == right) if isinstance(op, ast.Eq) else (left != right)
                else:
                    left_num = self._coerce_numeric(left)
                    right_num = self._coerce_numeric(right)
                    if left_num is None or right_num is None:
                        raise TypeError(f"大小比較は数値のみ: {left!r}, {right!r}")
                    if isinstance(op, ast.Gt):
                        ok = left_num > right_num
                    elif isinstance(op, ast.GtE):
                        ok = left_num >= right_num
                    elif isinstance(op, ast.Lt):
                        ok = left_num < right_num
                    elif isinstance(op, ast.LtE):
                        ok = left_num <= right_num
                    else:
                        raise TypeError("未対応の比較子")
                if not ok:
                    return False
                left = right
            return True

        if isinstance(node, ast.Call):
            if not isinstance(node.func, ast.Name):
                raise TypeError("未対応の関数呼び出しです。")
            name = node.func.id.lower()
            func = FUNCTION_DISPATCH.get(name)
            if func is None:
                raise TypeError(f"未対応の関数: {name}")
            self._validate_call(node, name)
            args = [self.eval_ast(arg, names) for arg in node.args]
            result = func(args)
            if not isinstance(result, BuiltinNumericResult):
                raise TypeError(f"{name}: 無効なビルトイン関数の戻り値です。")
            self.last_format_hint = result.format_hint
            return result.value

        if isinstance(node, ast.Constant):
            if isinstance(node.value, (int, float, str, bool)):
                return node.value
            raise TypeError(f"未対応のリテラル: {node.value!r}")

        if isinstance(node, ast.Name):
            key = "this" if node.id.lower() == "this" else node.id
            if key in names:
                fmt = self.var_formats.get(key)
                if fmt is not None:
                    hint, _formatted = fmt
                    self.last_format_hint = hint
                return names[key]
            if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", node.id):
                return 0
            raise NameError(f"未定義名: {node.id}")

        if isinstance(node, ast.Expr):
            return self.eval_ast(node.value, names)

        raise TypeError(f"未対応の式: {ast.dump(node)}")

    def resolve_item(self, code: str, unit: str | None) -> Any:
        """`#CODE[UNIT]` 形式の試験項目値を取得する。

        Args:
            code: 試験項目コード。
            unit: 単位を指定する場合の文字列。未指定なら ``None``。

        Returns:
            対応する値、もしくは `VarRef` を解決した通常変数の値。

        Raises:
            KeyError: 指定した試験項目が `items` に存在しない場合。
        """
        key = (code, unit) if unit is not None else code
        if key not in self.items:
            raise KeyError(f"試験項目が見つかりません: #{code}[{unit}]")
        value = self.items[key]
        if isinstance(value, VarRef):
            return self.vars.get(value.name, 0)
        return value

    def run_lines(self, lines: Iterable[str]) -> dict[str, Any]:
        """与えられたスクリプト行を Lab-Aid 構文として逐次実行する。

        Args:
            lines: 実行対象のスクリプト行イテラブル。

        Returns:
            実行完了時点の通常変数および `this` の辞書。

        Raises:
            SyntaxError: IF/ELSE/END・FOR/NEXT の対応が崩れた場合。
            RuntimeError: FOR ループの反復上限超過など、実行時の制約違反が起きた場合。
            TypeError: 構文上許可されないステートメントが実行された場合。
        """
        # コメントを除去し、解析しやすいプログラム（行リスト）を構築する。
        program = [strip_comment_quote_aware(line) for line in lines]
        pc = 0
        count = len(program)
        stack: list[dict[str, Any]] = []

        def is_active() -> bool:
            return all(frame["active"] for frame in stack)

        while pc < count:
            raw = program[pc].strip()
            pc += 1
            if not raw or raw.lower().startswith("rem"):
                continue

            match = PRINT_RE.match(raw)
            if match:
                if is_active():
                    arg_expr = match.group(1).strip()
                    if RE_ITEM_ANY.search(arg_expr):
                        raise SyntaxError(
                            "print: 引数には #項目を直接指定できません（通常変数か '文字' を使用）"
                        )
                    self.last_print = execute_print(
                        arg_expr, self.eval_expr, self._format_to_text
                    )
                continue

            match = PRINT2_RE.match(raw)
            if match:
                if is_active():
                    arg_expr = match.group(1).strip()
                    if RE_ITEM_ANY.search(arg_expr):
                        raise SyntaxError(
                            "print2: 引数には #項目を直接指定できません（通常変数か '文字' を使用）"
                        )
                    self.last_print2 = execute_print2(
                        arg_expr, self.eval_expr, self._format_to_text
                    )
                continue

            match = re.match(r"^\s*if\s+(.+?)\s*$", raw, re.IGNORECASE)
            if match:
                if len(stack) >= MAX_NEST_DEPTH:
                    raise SyntaxError(
                        f"制御構文のネストが上限を超えました（最大{MAX_NEST_DEPTH}）"
                    )
                condition = match.group(1)
                parent_active = is_active()
                cond_value = False
                if parent_active:
                    try:
                        cond_value = bool(self.eval_expr(condition))
                    except Exception as exc:
                        raise RuntimeError(
                            f"IF 条件評価エラー: {condition} : {exc}"
                        ) from exc
                stack.append(
                    {
                        "type": "IF",
                        "active": parent_active and cond_value,
                        "parent": parent_active,
                        "in_else": False,
                        "cond": cond_value,
                    }
                )
                continue

            if re.match(r"^\s*else\s*$", raw, re.IGNORECASE):
                if not stack or stack[-1]["type"] != "IF":
                    raise SyntaxError("ELSE に対応する IF がありません。")
                frame = stack[-1]
                if frame["in_else"]:
                    raise SyntaxError("同一 IF ブロック内で複数の ELSE は使えません。")
                frame["in_else"] = True
                frame["active"] = (not frame["cond"]) and frame["parent"]
                continue

            if re.match(r"^\s*end\s*$", raw, re.IGNORECASE):
                if not stack or stack[-1]["type"] != "IF":
                    raise SyntaxError("END に対応する IF がありません。")
                stack.pop()
                continue

            match = re.match(
                r"^\s*for\s+([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.+?)\s+to\s+(.+?)(?:\s+step\s+(.+?))?\s*$",
                raw,
                re.IGNORECASE,
            )
            if match:
                if len(stack) >= MAX_NEST_DEPTH:
                    raise SyntaxError(
                        f"制御構文のネストが上限を超えました（最大{MAX_NEST_DEPTH}）"
                    )
                var_name, from_expr, to_expr, step_expr = match.groups()
                parent_active = is_active()
                if not parent_active:
                    stack.append(
                        {
                            "type": "FOR",
                            "active": False,
                            "parent": False,
                            "skipping": True,
                        }
                    )
                    continue
                from_val = self.eval_expr(from_expr)
                to_val = self.eval_expr(to_expr)
                step_val = self.eval_expr(step_expr) if step_expr is not None else 1
                if not all(
                    isinstance(x, (int, float)) for x in (from_val, to_val, step_val)
                ):
                    raise TypeError("FOR の範囲/ステップは数値である必要があります。")
                if step_val == 0:
                    raise ValueError("FOR の STEP に 0 は指定できません。")
                self.vars[var_name] = from_val
                stack.append(
                    {
                        "type": "FOR",
                        "active": True,
                        "parent": True,
                        "var": var_name,
                        "to": to_val,
                        "step": step_val,
                        "start_pc": pc,
                        "iters": 0,
                        "skipping": False,
                    }
                )
                continue

            match = re.match(
                r"^\s*next(?:\s+([A-Za-z_][A-Za-z0-9_]*))?\s*$", raw, re.IGNORECASE
            )
            if match:
                if not stack or stack[-1]["type"] != "FOR":
                    raise SyntaxError("NEXT に対応する FOR がありません。")
                frame = stack[-1]
                var_name = match.group(1)
                if var_name and frame.get("var") != var_name:
                    raise SyntaxError("NEXT の変数名が対応する FOR と一致しません。")
                if frame.get("skipping"):
                    stack.pop()
                    continue
                variable = frame["var"]
                current = self.vars.get(variable, 0)
                step = frame["step"]
                limit = frame["to"]
                frame["iters"] += 1
                if frame["iters"] > MAX_FOR_ITERS:
                    raise RuntimeError("FOR 反復回数が上限を超えました。")
                next_val = current + step
                condition = (next_val <= limit) if step > 0 else (next_val >= limit)
                if condition:
                    self.vars[variable] = next_val
                    pc = frame["start_pc"]
                else:
                    stack.pop()
                continue

            if is_active() and self.exec_function_statement(raw):
                continue

            if is_active():
                self.exec_assign(raw)

        if stack:
            raise SyntaxError("ブロックの閉じ忘れがあります（END/NEXT の不足）。")

        return dict(self.vars)

    def exec_function_statement(self, line: str) -> bool:
        """`strcat(B, A)` のような関数ステートメントを実行する。

        Args:
            line: ステートメント形式の関数呼び出し文字列。

        Returns:
            ステートメントとして解釈できた場合は ``True``、それ以外は ``False``。

        Raises:
            TypeError: 引数数や引数型がステートメント仕様に反した場合。
        """
        try:
            expr = ast.parse(line, mode="eval")
        except SyntaxError:
            return False
        node = expr.body
        if not isinstance(node, ast.Call) or not isinstance(node.func, ast.Name):
            return False
        name = node.func.id.lower()
        if name not in STATEMENT_FUNCTIONS:
            return False
        if not node.args:
            raise TypeError(f"{name}: 引数が不足しています。")
        dest_node = node.args[0]
        if not isinstance(dest_node, ast.Name):
            raise TypeError(f"{name}: 第1引数には変数を指定してください。")
        dest = "this" if dest_node.id.lower() == "this" else dest_node.id
        if dest == "this":
            raise TypeError(f"{name}: 第1引数に this は指定できません。")
        value = self.eval_expr(line)
        formatted_value = format_roundjisb_output(value, self.last_format_hint)
        if formatted_value is not None and isinstance(value, (int, float)):
            value = FormatAwareNumber(float(value), formatted_value)
        self.vars[dest] = value
        if self.last_format_hint is not None or formatted_value is not None:
            self.var_formats[dest] = (self.last_format_hint, formatted_value)
        else:
            self.var_formats.pop(dest, None)
        return True

    def _validate_call(self, call: ast.Call, name: str) -> None:
        """`str_comp` 呼び出し専用の構文検証を行う。

        Args:
            call: 解析済みの関数呼び出しノード。
            name: 小文字化された関数名。

        Raises:
            TypeError: Lab-Aid の `str_comp` 仕様に違反する引数構成だった場合。
        """
        if name != "str_comp":
            return
        argc = len(call.args)
        if argc < 2:
            raise TypeError("str_comp: 引数数が不正です。")
        first = call.args[0]
        if isinstance(first, ast.Constant) and isinstance(first.value, str):
            raise TypeError("str_comp: 第1引数には文字列リテラルを指定できません。")
        if argc == 3:
            index_node = call.args[1]
            target = call.args[2]
            if not isinstance(index_node, ast.Name):
                raise TypeError("str_comp: 第2引数には変数を指定してください。")
            if not isinstance(target, ast.Constant) or not isinstance(
                target.value, str
            ):
                raise TypeError(
                    "str_comp: 第3引数には文字列リテラルを指定してください。"
                )
        if argc == 4:
            for idx_node in call.args[2:]:
                if not isinstance(idx_node, (ast.Constant, ast.Name)):
                    raise TypeError("str_comp: 試験回指定が不正です。")
                if isinstance(idx_node, ast.Constant) and not isinstance(
                    idx_node.value, int
                ):
                    raise TypeError("str_comp: 試験回には整数を指定してください。")

    def exec_assign(self, line: str) -> None:
        """代入文を解析して右辺式を評価し、変数へ格納する。

        Args:
            line: `=` を含む Lab-Aid の代入文。

        Raises:
            SyntaxError: `=` を含まない、または REM 行以外の不正な代入文だった場合。
            NameError: 左辺の変数名が Lab-Aid 仕様に適合しない場合。
            TypeError: 右辺評価でサポート外の型が発生した場合。
        """
        if line.strip().lower().startswith("rem"):
            return
        if "=" not in line:
            raise SyntaxError(f"代入行のみ対応: {line}")
        lhs, rhs = line.split("=", 1)
        lhs = lhs.strip()
        rhs = rhs.strip()

        if lhs.lower() == "this":
            lhs = "this"

        if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", lhs) and lhs != "this":
            raise NameError(f"左辺変数名が不正: {lhs}")

        value = self.eval_expr(rhs)
        formatted_value = format_roundjisb_output(value, self.last_format_hint)
        if formatted_value is not None and isinstance(value, (int, float)):
            value = FormatAwareNumber(float(value), formatted_value)
        self.vars[lhs] = value

        if self.last_format_hint is not None or formatted_value is not None:
            self.var_formats[lhs] = (self.last_format_hint, formatted_value)
        else:
            self.var_formats.pop(lhs, None)

        if lhs == "this":
            self.this_assigned_count += 1
            self.this_formatted = formatted_value
