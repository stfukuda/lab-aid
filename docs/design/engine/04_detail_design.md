# 4. 詳細設計（How）

## 4.1 ディレクトリ構成

```text
./  
 ├ src/
 │  └ lab_aid/
 │     ├ __init__.py
 │     ├ engine/
 │     │  ├ __init__.py            # 外部向け API を re-export
 │     │  ├ builtins.py            # 互換レイヤー（旧 API）
 │     │  └ runtime/
 │     │     ├ __init__.py
 │     │     ├ api.py              # evaluate(calc_type, script, inputs)
 │     │     ├ constants.py        # MAX_NEST_DEPTH などの安全装置
 │     │     ├ engine_core.py      # Engine クラスと実行ループ
 │     │     ├ inputs.py           # #CODE 形式の入力パーサ
 │     │     ├ text.py             # コメント除去・置換ユーティリティ
 │     │     └ functions/
 │     │        ├ __init__.py      # FUNCTION_DISPATCH の定義
 │     │        ├ base.py          # BuiltinNumericResult 等の基盤
 │     │        ├ numeric.py       # roundjisb 等の数値ビルトイン
 │     │        ├ string.py        # str_comp 等の文字列ビルトイン
 │     │        └ package.py       # print / print2 ラッパー
 │     ├ builtins.py               # 古い import 互換用
 │     └ excel_cli.py              # Excel 一括実行 CLI
 ├ tests/
 │  ├ conftest.py
 │  ├ engine_test.py              # E/R 正常系・制御構文・ビルトイン
 │  └ test_basic.py               # CLI や簡易ケース（拡張余地）
 └ docs/
    └ design/engine/*.md         # 本設計書
```

## 4.2 実装コンポーネント

- `runtime.api.evaluate`
  - `calc_type` を判別し、E では `ensure_has_this_assignment_E`、R では `assert_no_hash_usage` を適用。
  - `parse_inputs_E` / `parse_input_R` / `replace_rhs_this_for_R` で入力正規化。
  - `Engine` を初期化して `run_lines` し、戻り値を `(raw, edited, reported)` へ整形。
- `runtime.engine_core.Engine`
  - 状態: `items`, `vars`, `var_formats`, `this_formatted`, `last_print`, `last_print2`, `control_stack`。
  - `run_lines`: コメント除去済みスクリプトを走査し、`exec_assign`、`exec_function_statement`、`exec_print[_2]` を呼び出す。
  - `eval_expr` → `eval_ast`: Lab-Aid 記法を Python AST に変換し、安全なノードのみ評価。
  - 制御構文: IF/ELSE/END、FOR/NEXT のスタックを保持し、`constants` の上限で例外化。
- `runtime.inputs`
  - `parse_inputs_E`: `RE_INPUT_LINE` で行を抽出し、`validate_hash_name` で `#CODE`/`#HOLDER.CODE` を検証。
  - `_split_multi_values` と `_parse_value` でカンマ区切り・単一引用符を正しく処理。
  - `parse_input_R`: 数値・文字列・通常変数名のいずれかを返し、検証メッセージを統一。
  - `replace_rhs_this_for_R`: RHS の `this` を `__THIS_IN__` に置換し、入力値を保持。
- `runtime.functions`
  - `base.py`: `BuiltinNumericResult(value, format_hint, formatted)`、`coerce_number` 等。
  - `numeric.py`: 四則系、統計系、丸め (`roundjisb`/`round`)。フォーマットヒント (`fixed:n`, `sig:n`) を返す。
  - `string.py`: `str_comp`, `substr`, `strcat` など Lab-Aid 仕様に沿った検証を実装。
  - `package.py`: `print` / `print2` の引数評価と結果フォーマット。
- `excel_cli.py`
  - `click` ベースの CLI。テンプレート生成と Excel ワークブック一括処理を行い、`lab_aid.engine.evaluate` を呼び出す。
 - `windows/run_lab_aid.cmd` / `windows/setup_lab_aid.ps1`
  - 埋め込み Python を展開し、CLI を実行する Windows 利用者向けラッパー。

## 4.3 制御ロジック概要

- **IF/ELSE/END**
  - スタックに `(condition_met, in_else)` を push。
  - ELSE 時は直前 IF の `condition_met` を参照し、`in_else` フラグを立てる。
  - END で pop。ネスト超過は `MAX_NEST_DEPTH` を超えた時点で `ValueError`。
- **FOR/NEXT**
  - `for_stack` に `(var, start, end, step, iterations, body_index)` を保持。
  - `step` 省略時は `1` / `-1` を start/end から推測。
  - `iterations` が `MAX_FOR_ITERS` を超えるとエラー。
- **式評価**
  - `#CODE` や `#HOLDER.CODE[UNIT]` を `items` から取得。未定義は `ValueError`。
  - `replace_word_ci_outside_quotes` で `gt`, `eq` 等を Python オペレータへ置換。
  - `ast.parse` の結果を `eval_ast` で再帰処理。許可ノード以外は `TypeError`。
  - 関数呼び出しは `FUNCTION_DISPATCH[name_lower]` を参照し、`BuiltinNumericResult` で値＋フォーマットを返す。

## 4.4 エラー処理・安全装置

- `constants.MAX_NEST_DEPTH = 10`、`MAX_FOR_ITERS = 1_000_000` で暴走を抑制。
- `assert_no_hash_usage` が R モードでの `#` 利用を事前に遮断。
- `ensure_has_this_assignment_E/R` で `this = ...` が無いスクリプトを拒否。
- `runtime.api.evaluate` が全例外を捕捉し、E では `("エラー", None, None)`、R では `(None, "エラー", "エラー")` を返す。
- `excel_cli` は行ごとの失敗をログに記録し、最終的な exit code で異常を表現する。

## 4.5 使用ライブラリ

- ランタイム本体: 標準ライブラリのみ（`ast`, `math`, `re`, `decimal` 等）。
- CLI: `click`, `openpyxl`。
- 開発ツール: `pytest`, `ruff`, `pre-commit`, `just`（任意）。

## 4.6 例外ハンドリングとリトライ

- エンジン内部の例外は `RuntimeError` などへ変換せず、捕捉後にエラー文字列を返すのが基本方針。
- CLI では 1 行失敗しても他行の処理を継続し、失敗行の情報を Excel シートへ書き戻す（例: `status="エラー"`）。
- Windows バッチは PowerShell スクリプトの戻り値で成功/失敗を判定し、必要に応じて再実行案内を表示。
