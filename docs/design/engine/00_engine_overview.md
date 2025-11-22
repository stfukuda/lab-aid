# Lab-Aid Engine Overview

## 1. Scope

この資料では Lab-Aid 計算式エンジンの全体像を整理し、主要モジュール・データフロー・責務分担を共有する。詳細な制御構文やビルトイン仕様などは後続ドキュメントに委ねる。

## 2. Top-Level Flow

```text
lab_aid.engine.evaluate
 └─ runtime.api.evaluate(calc_type, script, inputs)
     ├─ runtime.inputs.parse_inputs_E / parse_input_R
     ├─ Engine(items, vars, ...)
     │   └─ run_lines(lines)
     │       ├─ exec_assign / exec_function_statement
     │       ├─ eval_expr → eval_ast → FUNCTION_DISPATCH
     │       └─ state tracking (IF / FOR / var_formats / this)
     └─ runtime.text / runtime.functions ヘルパー
```

- `lab_aid.engine.evaluate`: 外部公開 API。`calc_type` に応じて E/R を切り替える。
- `runtime.api.evaluate`: 入力の正規化・検証を行い、`Engine` を初期化して実行結果を整形する中核。
- `runtime.inputs`: `#CODE` 形式の試験項目や R モードの単一入力を解析し、`Engine` が扱える形へ変換。
- `runtime.engine_core.Engine`: Lab-Aid スクリプトを 1 行ずつ解釈し、式評価や制御構文のステート管理を担う。
- `runtime.functions`: 数値／文字列／印字系ビルトインをまとめたディスパッチレイヤー。
- `runtime.text`: コメント除去や単語置換など、Lab-Aid 固有のテキスト操作を提供。

## 3. Module Responsibilities

| Module | Key Responsibilities | Notes |
| --- | --- | --- |
| `lab_aid/engine/__init__.py` | `evaluate`, `Engine`, `VarRef` を re-export し、外部 API を単純化。 | バックコンパチ維持のため `lab_aid.builtins` も残す。 |
| `runtime/api.py` | 例外→互換エラー文字列への変換、E/R 判別、入力パース、`Engine` 起動。 | `assert_no_hash_usage` で R モードの制限も enforce。 |
| `runtime/engine_core.py` | 実行ループ、式評価 (`eval_expr`/`eval_ast`)、IF/ELSE/END と FOR/NEXT のステート管理。 | `var_formats` や `last_print*` を保持し結果整形に活用。 |
| `runtime/inputs.py` | `parse_inputs_E`, `parse_input_R`, `replace_rhs_this_for_R` など入力正規化。 | `VarRef` を通じた遅延解決を担当。 |
| `runtime/text.py` | コメント除去、キーワード置換、数値解析、`#CODE` 書式検証。 | エンジン全体で共有するテキストユーティリティ。 |
| `runtime/functions/` | ビルトイン実装 (`numeric`, `string`, `package`) とヘルパー (`base`). | `FUNCTION_DISPATCH` を構築し `eval_ast` から参照。 |

## 4. Data & State

| State | Description |
| --- | --- |
| `items` | `#CODE` / `#CODE[UNIT]` をキーとする試験項目値（`VarRef` を含む）。 |
| `vars` | `this` を含む通常変数テーブル。Lab-Aid は大文字小文字を区別。 |
| `var_formats` | `roundjisb` 系で得たフォーマットヒントと整形済み文字列を保存。 |
| `this_formatted` | `this` の表示文字列。raw 出力に利用。 |
| `last_print` / `last_print2` | 印字系ビルトインの直近結果。edited / reported に反映。 |
| `control_stack` | IF/ELSE/END および FOR/NEXT のネスト状態。`MAX_NEST_DEPTH`, `MAX_FOR_ITERS` を監視。 |

## 5. Error Contracts

- 解析・実行中に Python 例外が発生した場合、`runtime.api.evaluate` で捕捉し Lab-Aid 互換の `"エラー"` を返す。
- R モードで `#` 変数が使われた場合は早期に拒否。
- IF/ELSE/END や FOR/NEXT の不整合は `SyntaxError` として検出し、呼び出し元へエラー文字列を返す。

## 6. Extension Notes

- 新ビルトインは `runtime/functions` にカテゴリ別に追加し、戻り値を `BuiltinNumericResult` へ統一する。
- 新しい入力形式や制御構文を導入する場合は `runtime/constants.py` の制限値やエラー仕様への影響を明示する。
- CLI や Excel 連携は `excel_cli.py` を入口としてエンジン API を利用するため、API 互換性を最優先とする。
