# 7. テスト仕様

## 7.1 テスト方針

- E/R 各モードで代表的な制御構文・ビルトイン・エラー系を網羅する。
- CLI（Excel 入力）と API（直接呼び出し）の両方を検証する。
- 正常系・異常系ともに `tests/engine_test.py` でカバーする（必要に応じてサブテストを追加する）。

### E タイプ

- **NE-01 `this` 代入のみ**  
  `test_e_assign_literal_number`, `test_e_assign_literal_string`
- **NE-02 通常変数 + `#CODE` 参照**  
  `test_e_assign_variable_number`, `test_e_hash_lookup_with_unit`
- **NE-03 文字列代入・比較**  
  `test_e_assign_literal_string`, `test_e_string_functions_is_char_strlen`
- **NE-04 IF/ELSE 分岐**  
  `test_e_if_else_branching`, `test_e_comparison_operator`
- **NE-05 複合条件・論理演算**  
  `test_e_comparison_and_logical`, `test_e_comparison_logical_with_parentheses`
- **NE-06 FOR/NEXT ループ**  
  `test_e_for_loop_accumulates_values`, `test_e_for_with_negative_step`
- **NE-07 文字列ビルトイン／ステートメント**  
  `test_e_strcat_statement_updates_target`, `test_e_strncpy_statement_uses_shift_jis_bytes`,  
  `test_e_string_functions_empty_and_space`
- **NE-08 数値ビルトイン／丸め**  
  `test_e_roundjisb_preserves_trailing_zero`, `test_e_round_significant_digits_and_errors`
- **NE-09 数値ビルトイン（floor/trunc/max/min/ave/sum/stdev）**  
  `test_e_floor_positive_value`, `test_e_trunc_negative_toward_zero`,  
  `test_e_max_min_with_multi_measurements`, `test_e_aggregate_functions`, `test_e_stdev_and_stdeva`
- **NE-10 print/print2 の組み合わせ**  
  `test_e_print_outputs`, `test_e_print2_outputs`
- **NE-11 多段ネスト**  
  `test_e_deep_nesting_respects_limits`, `test_e_nested_for_if_structure`
- **NE-12 複数入力フォーマット**  
  `test_e_multi_unit_and_holder_inputs`

### R タイプ

- **NR-01 R モード丸め**  
  `test_r_roundjisb_uses_input_as_this`
- **NR-02 固有挙動（入力エコー・print/print2）**  
  `test_r_literal_echo_when_no_this_or_print`, `test_r_print_overrides_this_value`,  
  `test_r_print2_overrides_reported_only`, `test_r_print2_condition_false_keeps_literal`
- **NR-03 RHS の `this` 置換**  
  `test_r_rhs_this_always_uses_input_value`
- **NR-04 コメント／文字列での `#` / `this` 無視**  
  `test_r_allows_hash_in_literals_and_comments`, `test_r_comments_and_literals_ignore_hash_and_this`
- **NR-05 print2 による roundjisb フォーマット反映**  
  `test_r_print2_preserves_roundjisb_formatting`

### 共通

- **NC-01 CLI 実行**  
  `windows/lab_aid_input.xlsx` 複数行 → D〜G列が更新され `status="OK"`

## 7.3 異常系テストケース

- **EE-01 `this =` が無い**  
  `test_e_requires_this_assignment`
- **EE-02 IF/END 不整合**  
  `test_e_missing_end_reports_error`
- **EE-03 FOR 反復超過・STEP 0**  
  `test_e_for_loop_iteration_limits`, `test_e_for_with_zero_step_reports_error`
- **EE-04 数値演算の不正値**  
  `test_e_sqrt_and_errors`, `test_e_logarithms`
- **EE-05 ビルトイン引数不正（roundjisb/round/floor/trunc）**  
  `test_e_roundjisb_invalid_argument_count`, `test_e_round_significant_digits_and_errors`,  
  `test_e_floor_invalid_argument_type`, `test_e_trunc_invalid_argument_type`
- **EE-06 統計系の入力不足**  
  `test_e_stdev_and_stdeva`
- **EE-07 `#` 参照ミスや VarRef 誤用**  
  `test_e_hash_item_missing_raises_error`, `test_e_hash_argument_rejects_varref`
- **EE-08 文字列ビルトインの制約違反**  
  `test_e_str_comp_rejects_literal_left_operand`,  
  `test_e_str_comp_requires_variable_index_in_three_arg_form`,  
  `test_e_str_comp_requires_string_literal_in_three_arg_form`
- **EE-09 ネスト上限超過**  
  `test_e_deep_nesting_respects_limits`

### R タイプ

- **ER-01 `#CODE` 使用違反** (`test_r_rejects_hash_variables`)
- **ER-02 入力欠如/VarRef** (`test_r_varref_defaults_to_zero_before_execution`)

### 共通

- **EC-01 CLI で Excel が開かれている**  
  ファイルロック → `status="エラー"` とログ警告
- **EC-02 CLI で `#` 指定ミス**  
  `inputs` がフォーマット不正 → `status="エラー"` とログ警告

## 7.4 手動確認項目

- Excel CLI: `windows/run_lab_aid.cmd` を実行し、エラー行が正しくステータス表示されるか確認する。
- Windows 配布: `windows/setup_lab_aid.cmd` → `windows/run_lab_aid.cmd` を順に実行し、環境を再現できるかチェックする。
- ログ出力: 失敗行がターミナルに表示されるか、`uv run pytest -k "error"` で異常系がカバーされているかを監視する。

## 7.5 カバレッジと保守

- `just test` を定期的に実行し、E/R の代表ケースを網羅する。
- 新しいビルトインや制御構文を追加する場合は、正常・異常テストを最低 1 件ずつ追加する。
- CLI の挙動変更（列追加など）があれば `tests/engine_test.py` に対応するリグレッションテストを追加する。
