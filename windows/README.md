# Lab-Aid ユーザーガイドライン（Windows 配布版）

Lab-Aid ユーザー向けの Windows 配布手順です。
公式の埋め込み版 Python を自動取得する PowerShell スクリプトを利用し、
利用者が Python を個別にインストールする必要がない構成を提供します。
ネットワークに制限がある場合でも、事前ダウンロードしたアーカイブを指定することで
セットアップできます（「オフラインまたは制限付きネットワークの場合」を参照）。

## 1. 対象読者

- Windows 11 (64bit) で Lab-Aid ランタイムを利用したい運用担当者
- ネットワーク経由で `python.org` / `bootstrap.pypa.io` / PyPI へアクセス可能な環境
- PowerShell が利用できること（標準環境を想定）

## 2. 同梱物とフォルダ構成

- `windows/setup_lab_aid.cmd`（セットアップ実行バッチ） … セットアップ用ラッパーバッチ
- `windows/setup_lab_aid.ps1`（セットアップスクリプト） … セットアップスクリプト本体
- `windows/run_lab_aid.cmd`（ランタイム実行バッチ） … Windows フォルダに配置された実行用ラッパーバッチ
- `lab_aid/`（ソースコード一式）
- `docs/`（設計書・引き継ぎ資料）
- `LICENSE` … Lab-Aid プロジェクトのライセンス

セットアップ完了後に自動生成される主なフォルダ:

| フォルダ / ファイル | 内容 |
| --- | --- |
| `runtime/python/` | 埋め込み版 Python 本体と site-packages |
| `licenses/python/` | Python 本体に同梱されるライセンスファイル群 |
| `licenses/third_party/` | 主要依存ライブラリ（openpyxl 等）のライセンス写し |
| `windows/lab_aid_input.xlsx` | Lab-Aid 実行用の入力テンプレート（自動生成） |
| `windows/run_lab_aid.cmd` / `lab_aid.bat` | 実行用ラッパーバッチ |
| `windows/lab_aid-*.whl` | Lab-Aid エンジンの wheel（オフラインインストール用） |

`windows/lab_aid-*.whl` が存在する場合、セットアップスクリプトはこの wheel を優先的に
インストールし、ネットワークに接続できない環境でも構築できるようになっています。

## 3. セットアップ手順

1. **`windows/setup_lab_aid.cmd`（セットアップ実行バッチ）を実行する**  
   エクスプローラーで `windows` フォルダを開き、`setup_lab_aid.cmd` を右クリックして
   「管理者として実行」を選びます（ダブルクリックでも OK）。
   画面の指示に従えば、埋め込み版 Python の展開・依存インストール・ライセンス収集まで
   自動で完了します。

   - SmartScreen などで警告が表示された場合は「詳細情報」→「実行」を選択してください。
   - オフライン利用時はこのフォルダに配置した `python-*.zip` および `get-pip.py` が
     自動で使用されます（次節参照）。

2. **`windows/lab_aid_input.xlsx` に計算条件を入力する**  
   `windows` フォルダに自動生成された `lab_aid_input.xlsx` を開き、次の形式でデータを
   入力します。1 行目は大分類（入力 / 出力）、2 行目は項目名、3 行目以降にテストパターンを
   記入します。

   - A 列: **計算式タイプ** … `E`（推定計算）または `R`（丸め計算）
   - B 列: **計算式** … Lab-Aid のスクリプト。セル内で改行（Alt + Enter）すると複数行のスクリプトを記述できます。
   - C 列: **変数** … `A=5` のような入力値。複数の値を指定する場合はセル内で改行します。
   - D〜F 列: **生データ / 編集後 / 報告値** … 実行結果が文字列として書き戻されます。
   - G 列: **status** … 行ごとの実行ステータス（OK / ERROR など）

3. **`windows/run_lab_aid.cmd`（ランタイム実行バッチ）を実行して評価する**  
   `windows` フォルダ内の `run_lab_aid.cmd`（または `lab_aid.bat`）をダブルクリックすると、
   Excel の各行を読み込んで Lab-Aid 評価を実行し、結果を同じファイルに書き戻します。
   `--help` オプションで利用方法を表示できます。

   ```cmd
   C:\path\to\lab-aid> windows\run_lab_aid.cmd
   C:\path\to\lab-aid> windows\run_lab_aid.cmd --help
   ```

> 補足: PowerShell コマンドラインで直接実行したい場合は、`setup_lab_aid.ps1`
> （セットアップスクリプト）を呼び出す従来の手順も利用できます（付録 A 参照）。

### オフラインまたは制限付きネットワークの場合

事前に以下のファイルをダウンロードし、`windows/setup_lab_aid.ps1` と同じフォルダ
（`windows/`）に保存してください。

| 必要ファイル | 取得元の例 | 備考 |
| --- | --- | --- |
| `python-<version>-embed-amd64.zip` | <https://www.python.org/ftp/python/> | 例: `python-3.13.0-embed-amd64.zip` |
| `get-pip.py` | <https://bootstrap.pypa.io/get-pip.py> | pip 導入用スクリプト |

保存したファイル名をそのまま利用できるようになっているため、
`setup_lab_aid.cmd` を実行するだけで自動的に検出されます。

特殊な配置を行いたい場合のみ、`setup_lab_aid.ps1`（セットアップスクリプト）を手動で
実行し、`-PythonZipPath` や `-PipScriptPath` 引数で明示的にパスを指定します
（付録 A 参照）。

## 4. ライセンスについて

- セットアップ完了後、`licenses/` フォルダに以下が自動生成されます。
  - `licenses/python/` … 埋め込み版 Python に含まれるライセンス (`python*.txt` 等)
  - `licenses/third_party/` … `openpyxl`、`et_xmlfile` など主要依存のライセンス写し
  - `licenses/LICENSE_lab_aid.txt` … プロジェクトの `LICENSE` のコピー
- 再配布する場合は `licenses/` フォルダと `LICENSE` を必ず同梱してください。

## 5. アップデート / 再セットアップ

- 新しいバージョンを導入する場合や環境を初期化したい場合は、再度
  `setup_lab_aid.cmd`（セットアップ実行バッチ）を実行します。
  既存の `runtime\python` を削除するため、必要に応じて `-Force` を付けて実行してください
  （付録 A 参照）。
- Python バージョンを固定したい場合は、`setup_lab_aid.ps1`
  （セットアップスクリプト）を直接呼び出し `-PythonVersion` 引数で明示的に指定します
  （付録 A）。

## 6. トラブルシューティング

- **ファイアウォールやプロキシによりダウンロードできない**  
  - `python.org` / `bootstrap.pypa.io` / `pypi.org` へのアクセスを許可するか、
    事前にダウンロードしたアーカイブを共有ディレクトリに置き、
    スクリプトのダウンロード部分を書き換えて利用してください。
- **PowerShell の実行ポリシーでブロックされる**  
  - `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` を再度実行してください。
- **ライセンスファイルが生成されない**  
  - ネットワーク断や権限不足が原因でライブラリのインストールに失敗していないか
    確認してください。`licenses/third_party` にファイルが無い場合は
    `pip show openpyxl` などでインストール状況を確認してください。

## 7. 参考情報

- セットアップスクリプトの詳細: `windows/setup_lab_aid.ps1`（セットアップスクリプト）
- 設計情報と運用ガイド:
  - `docs/guidelines/handover_guide.md`
  - `docs/design/engine/00_engine_overview.md` ほか `docs/design/engine/01-09*.md`

---

### 付録 A: コマンドラインでの実行例（上級者向け）

PowerShell から直接セットアップしたい場合は次のコマンドを利用できます。

```powershell
PS C:\> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
PS C:\path\to\lab-aid> powershell -ExecutionPolicy Bypass -File windows\setup_lab_aid.ps1 -Force
```

特殊な配置を行う場合の例:

```powershell
PS C:\path\to\lab-aid> powershell -ExecutionPolicy Bypass -File windows\setup_lab_aid.ps1 `
    -PythonZipPath D:\offline\python-3.13.0-embed-amd64.zip `
    -PipScriptPath D:\offline\get-pip.py
```

実行のみを行いたい場合は、`windows\run_lab_aid.cmd`（ランタイム実行バッチ）を
右クリックし「管理者として実行」するか、以下のようにコマンドから呼び出してください。

```cmd
C:\path\to\lab-aid> windows\run_lab_aid.cmd
```

不明点があれば開発チームまでお問い合わせください。
