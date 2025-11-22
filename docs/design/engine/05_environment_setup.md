# 5. 環境構築手順

## 5.1 前提ソフトウェア

- OS: Windows 11
- Python 3.13 以上
- uv（https://github.com/astral-sh/uv）
- git（https://git-scm.com/）
- just（https://github.com/casey/just）
- prek（https://github.com/chriskuehl/prek）
- PowerShell 5.1+
- Microsoft Excel 2019 以降。

## 5.2 ローカルでの開発

1. Devcontainer 環境、もしくは Git Bash + `uv` の仮想環境で作業する。
2. 依存同期は `just sync` を実行する。
3. 動作確認は `just test`（必要に応じて `uv run python -m lab_aid.excel_cli --help`）で行う。

## 5.3 Windows 利用者向け配布

1. `windows/setup_lab_aid.cmd` をダブルクリックして埋め込み Python と依存を展開する。
2. `windows\run_lab_aid.cmd` をダブルクリックし、`windows\lab_aid_input.xlsx` と同じフォルダで CLI を起動する。

## 5.4 トラブルシュート

- `uv sync` が証明書エラーとなる場合は社内 CA を `REQUESTS_CA_BUNDLE` に設定。
- PowerShell 実行制限は `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` で緩和。
- Excel CLI がファイルロックに遭遇した際は Excel を閉じて再実行する。
