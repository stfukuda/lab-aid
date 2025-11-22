# 8. 制約事項・注意点

## 8.1 技術的制約

- **制御構文の上限**  
  - IF/ELSE/END と FOR/NEXT のネスト深度は `MAX_NEST_DEPTH = 10` に制限。超過すると `ValueError`。
  - FOR ループの反復回数は `MAX_FOR_ITERS = 1_000_000`。超過時は `RuntimeError`。
- **R モードの `#` 禁止**  
  - `assert_no_hash_usage` により、スクリプト／入力値の両方で `#CODE` を使用するとエラー。
- **R モードの `this` 置換**  
  - 右辺に現れる `this` は `__THIS_IN__` に置換され、入力値を参照する。複数回の代入でも入力値が元になる。
- **フォーマットヒント**  
  - `roundjisb`/`round` 実行時に `fixed:n`/`sig:n` といったヒントを `Engine.last_format_hint` に保持し、直後の `this` 代入や `print` 系で `format_roundjisb_output` を通じて桁数固定表示を行う。ヒントは1回の評価で消費され、未対応の型では無視される。
- **print 系の評価**  
  - `print`/`print2` は任意の式を `Engine.eval_expr` に渡しており、`#CODE` を含めた複合式でも評価できる。制約は Lab-Aid 式構文の範囲内に限られる。
- **`str_comp` 仕様**  
  - 引数は 2〜4 個に限定され、試験回インデックス（第3、第4引数）は整数である必要がある。範囲外の引数数や非整数を渡すと `TypeError`。
- **文字列系関数**  
  - `strcat`/`strncpy` などの文字列関数はすべて式評価として動作し、副作用は持たない。引数には `this` や `#CODE` を含めた任意の式を指定できる。
- **入力パーサ**  
  - E モードは `#CODE[UNIT]` や `#HOLDER.CODE` をサポートする。名前が仕様外の場合は `ValueError`。
  - R モードの第3引数は数値または `'文字列'` のみに制限。空文字や通常変数名は `ValueError`。

## 8.2 運用上の注意

- **Excel ファイルのロック**  
  - CLI 実行前に Excel を閉じる。開いたままだと `status="エラー"` になる。
- **Devcontainer/Git Bash**  
  - 開発は Devcontainer または Git Bash + `uv` で行い、`just sync` → `just test` の手順を守る。
- **Windows 配布**  
  - 利用者は `windows/setup_lab_aid.cmd`・`windows/run_lab_aid.cmd` をダブルクリックするだけで実行できるが、`lab_aid_input.xlsx` と同ディレクトリ配置が必要。
- **ログ監視**  
  - CLI 実行ログにエラー行・フォーマット不正が出力されるので、再実行時にはターミナルのログを確認する。
- **CLI の入力フォーマット**  
  - `calc_type` は `E` か `R`、`script` は Lab-Aid 記法、`inputs` は `#CODE=value` の複数行文字列に限定する。

## 8.3 既知の制限

- **`lab_aid/excel_cli.py`**  
  - CLI モジュール自体はユニットテストカバレッジ外。手動確認での保証となる。
- **ヘッダーなし CSV 等のバッチ入力**  
  - Excel 以外の入力フォーマットはサポート対象外。必要に応じて別途拡張が必要。
- **型チェックの厳密差**  
  - ビルトイン関数の引数チェックは Python の `TypeError` をそのまま返す箇所がある（例: `round` に真偽値を渡した場合）。Lab-Aid 仕様との差異がある場合は `tests/engine_test.py` で明示的に補う。
