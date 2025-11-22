# 6. 動作手順・ユースケース

## 6.1 CLI 実行手順（Excel ブック）

1. `windows/lab_aid_input.xlsx` を用意し、次の列を入力する。  
   - `calc_type`（A列）: `E` または `R`  
   - `script`（B列）: Lab-Aid 記法のスクリプト。セル内改行で複数行を記述する。  
   - `inputs`（E）または `input`（R）（C列）: `#CODE=value` 形式や単一値。
2. `windows\run_lab_aid.cmd` をダブルクリックして実行する。
3. 実行後、Excel の結果列（D〜G列の `raw`/`edited`/`reported`/`status`）が更新される。エラー行は `status="エラー"` となる。

## 6.2 API 実行例（Python から）

```python
from lab_aid.engine import evaluate

script = """
this = #A + #B
"""
inputs = """
A=5
B=3
"""
result = evaluate("E", script, inputs)
print(result)  # ("8", None, None)
```

- R タイプの例: `evaluate("R", "this = roundjisb(this, 2, 1)", "25.346")`
- 例外発生時は Lab-Aid 互換のエラー文字列が戻る。

## 6.3 代表ユースケース

- **検査測定の自動判定**: Excel に蓄積された測定値を一括評価し、同ファイルの `edited`/`reported` 列をレポートへ転記する。人手による再計算を削減。
- **教育・認定制度**: 認定試験の模擬入力を Excel で作成し、CLI で一括採点する。既存エンジンと同じ結果を保証。
- **シミュレーション / 開発検証**: API を Python スクリプトから直接呼び出し、新しい Lab-Aid 式やビルトイン追加の挙動をテストする。

## 6.4 操作時の注意

- Excel ファイルは書き込み中にロックされるため、CLI 実行前に Excel を閉じる。
- `calc_type` によらず、式中には `this = ...` を必ず 1 回は含める。`R` では `#` 変数を禁止する。
- CLI 実行ログはターミナルに表示される。失敗行を確認して再処理する。
