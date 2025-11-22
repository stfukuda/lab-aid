# 9. 保守メモ

## 9.1 よくある問い合わせ

- **Excel を閉じ忘れてエラーになる**  
  - `lab_aid.excel_cli` は `openpyxl` ベースのため、対象ブックが開いていると `status="エラー"` で失敗する。操作手順の際に必ず閉じる旨を案内する。
- **R タイプ第 3 引数の形式**  
  - 実装上は数値または `'文字列'` しか受け付けない。通常変数名を渡すと `ValueError` になり、Excel CLI では「入力値が不正」と表示される。
- **`this=` が無いスクリプト**  
  - E タイプは `this=` を最低 1 行含む必要がある。`ensure_has_this_assignment_E` で弾かれるので、質問を受けた際はスクリプト内の REM 以外を確認してもらう。
- **print の表示桁数が期待どおりにならない**  
  - `roundjisb` / `round` が直前に実行されているか確認。ヒントが消費されるため、間に別の式評価を挟むと桁数指定が失われる。

## 9.2 改修時の注意

- **埋め込み Python と wheel 配布**  
  - Windows フォルダには `lab_aid-*.whl` を必ず含める。`setup_lab_aid.ps1` は wheel 未検出で即終了するため、ビルド後に `just ship` でコピーする手順を守る。
- **`Engine` のフォーマット処理**  
  - `FormatAwareNumber` と `var_formats` は `print` 系だけでなく `this` 代入でも参照される。新たに数値計算を追加する際は `format_roundjisb_output` の呼び出し箇所を忘れない。
- **R タイプの `#` 禁止**  
  - `assert_no_hash_usage` がスクリプトと入力値を検査している。仕様を緩和する場合は `tests/engine_test.py` の R 系テストを更新し、`replace_rhs_this_for_R` との整合を保つ。
- **Excel CLI の生成物パス**  
  - `lab_aid_input.xlsx`、埋め込み Python、一時ログはすべて `windows/` 配下を前提にしている。出力先を変更する場合は PowerShell スクリプトとドキュメントの両方を更新する。

## 9.3 今後の検討事項

- **Excel ブックを開いたまま編集できるライブラリ**  
  - 現状は `openpyxl` 固定。将来的に `xlwings` 等の採用を検討する場合、依存ライブラリや配布サイズ、Windows での配布容易性を比較した上で設計を改訂する。
- **CLI エラーハンドリングの自動テスト**  
  - Excel CLI の挙動は手動検証に依存している。`pytest` で CLI 層を直接たたく E2E テストを整備し、テンプレート生成やログ出力の回帰検知を行いたい。
- **Windows 配布物の自動パッケージング**  
  - 埋め込み Python・wheel・テンプレート・バッチを 1 つのアーカイブにまとめるスクリプトを追加し、配布手順を簡略化する余地がある。
