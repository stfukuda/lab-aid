<#
  Lab-Aid ランタイムを Windows で動作させるための埋め込み版 Python セットアップスクリプト。
  - 公式 Python Embed パッケージのダウンロードと展開
  - pip 導入および `lab_aid` パッケージのインストール
  - ライセンスファイルの収集とランチャーバッチの生成
#>

param(
    [string]$PythonVersion = "3.13.0",
    [string]$InstallDir = "windows\runtime\python",
    [switch]$Force = $true,
    [string]$PythonZipPath = $null,
    [string]$PipScriptPath = $null
)

$ErrorActionPreference = "Stop"

# 進捗や警告を日本語で出力するヘルパー
function Write-Info($Message) {
    Write-Host "[INFO] $Message"
}

function Write-Warn($Message) {
    Write-Host "[WARNING] $Message" -ForegroundColor Yellow
}

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$targetPath = Join-Path $repoRoot $InstallDir

if (Test-Path $targetPath) {
    # 既存環境がある場合は Force 指定なしならそのまま終了し、誤削除を防ぐ
    if (-not $Force) {
        Write-Warn "既に '$InstallDir' が存在します。上書きする場合は -Force を付けて再実行してください。"
        return
    }
    Write-Info "'$InstallDir' を削除しています。"
    Remove-Item -Path $targetPath -Recurse -Force
}

New-Item -ItemType Directory -Path $targetPath | Out-Null

if ($PythonZipPath) {
    # 事前にダウンロードしたアーカイブを利用
    if (-not (Test-Path $PythonZipPath)) {
        throw "指定された Python アーカイブが見つかりません: $PythonZipPath"
    }
    Write-Info "ローカルの Python アーカイブ '$PythonZipPath' を展開します。"
    Expand-Archive -LiteralPath $PythonZipPath -DestinationPath $targetPath -Force
} else {
    $arch = "amd64"
    $pythonZipName = "python-$PythonVersion-embed-$arch.zip"
    $localZipCandidate = Join-Path $PSScriptRoot $pythonZipName
    if (Test-Path $localZipCandidate) {
        Write-Info "スクリプトと同じフォルダにある '$pythonZipName' を使用します。"
        Expand-Archive -LiteralPath $localZipCandidate -DestinationPath $targetPath -Force
    } else {
        $pythonDownloadUrl = "https://www.python.org/ftp/python/$PythonVersion/$pythonZipName"
        $tempZipPath = Join-Path ([System.IO.Path]::GetTempPath()) $pythonZipName

        Write-Info "Python $PythonVersion 埋め込み版をダウンロードしています: $pythonDownloadUrl"
        Invoke-WebRequest -Uri $pythonDownloadUrl -OutFile $tempZipPath

        Write-Info "埋め込み版 Python を '$InstallDir' に展開しています。"
        Expand-Archive -LiteralPath $tempZipPath -DestinationPath $targetPath -Force
        Remove-Item $tempZipPath
    }
}

$majorMinor = ($PythonVersion.Split(".")[0..1] -join "")
# Embed 版では site-packages が無効化されているため、_pth の import site を有効化する
$pthFile = Join-Path $targetPath ("python{0}._pth" -f $majorMinor)
if (-not (Test-Path $pthFile)) {
    throw "Could not find expected path configuration file: $pthFile"
}

$pthContent = Get-Content $pthFile
if ($pthContent -match "#import site") {
    Write-Info "埋め込み版 Python の site-packages を有効化します。"
    $pthContent = $pthContent -replace "#import site", "import site"
    Set-Content -Path $pthFile -Value $pthContent -Encoding ASCII
} elseif ($pthContent -notmatch "import site") {
    Write-Info "'import site' を追記して site-packages を有効化します。"
    Add-Content -Path $pthFile -Value "import site"
}

$pythonExe = Join-Path $targetPath "python.exe"
if (-not (Test-Path $pythonExe)) {
    throw "Python 実行ファイルが見つかりません: $pythonExe"
}

# Embed 環境内に pip を導入
Push-Location $targetPath
try {
    if ($PipScriptPath) {
        if (-not (Test-Path $PipScriptPath)) {
            throw "指定された get-pip.py が見つかりません: $PipScriptPath"
        }
        Write-Info "ローカルに用意した get-pip.py を使用します。"
        $getPipPath = $PipScriptPath
    } else {
        $localPipCandidate = Join-Path $PSScriptRoot "get-pip.py"
        if (Test-Path $localPipCandidate) {
            Write-Info "スクリプトと同じフォルダの get-pip.py を使用します。"
            $getPipPath = $localPipCandidate
        } else {
            $getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
            $getPipPath = Join-Path $targetPath "get-pip.py"
            Write-Info "pip インストーラ (get-pip.py) をダウンロードしています。"
            Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipPath
        }
    }

    Write-Info "pip を埋め込み環境にインストールしています。"
    & $pythonExe $getPipPath
    if (-not $PipScriptPath -and -not (Test-Path (Join-Path $PSScriptRoot "get-pip.py"))) {
        Remove-Item $getPipPath
    }
} finally {
    Pop-Location
}

Write-Info "pip / setuptools / wheel を最新化しています。"
& $pythonExe -m pip install --upgrade pip setuptools wheel

Write-Info "Lab-Aid パッケージのインストール方法を確認しています。"
$localWheel = Get-ChildItem -Path $PSScriptRoot -Filter "lab_aid-*.whl" -ErrorAction SilentlyContinue | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1
if ($null -eq $localWheel) {
    throw "windows フォルダに 'lab_aid-*.whl' が見つかりません。配布パッケージに wheel を同梱してください。"
}
Write-Info "ローカルの wheel '$($localWheel.Name)' をインストールします。"
& $pythonExe -m pip install --upgrade $localWheel.FullName

# ---- ライセンス収集 -------------------------------------------------------
Write-Info "Collecting license files."
$licensesRoot = Join-Path $repoRoot "licenses"
New-Item -ItemType Directory -Path $licensesRoot -Force | Out-Null

$pythonLicenseDir = Join-Path $licensesRoot "python"
New-Item -ItemType Directory -Path $pythonLicenseDir -Force | Out-Null
Get-ChildItem -Path $targetPath -Filter "*.txt" | ForEach-Object {
    Copy-Item -Path $_.FullName -Destination (Join-Path $pythonLicenseDir $_.Name) -Force
}

$projectLicense = Join-Path $repoRoot "LICENSE"
if (Test-Path $projectLicense) {
    Copy-Item -Path $projectLicense -Destination (Join-Path $licensesRoot "LICENSE_lab_aid.txt") -Force
}

$thirdPartyLicenseDir = Join-Path $licensesRoot "third_party"
# 依存ライブラリで明示的にライセンスを確認したいパッケージ
$packages = @("openpyxl", "et_xmlfile")

$collectorScriptPath = Join-Path $targetPath "collect_licenses.py"
# site-packages を走査してライセンスファイルらしきものを抽出する簡易スクリプト
$collectorScript = @'
import importlib
import json
import os
import sys

names = sys.argv[1:]
result = {}
for name in names:
    try:
        module = importlib.import_module(name)
    except Exception:
        continue
    base = os.path.dirname(getattr(module, "__file__", ""))
    if not base or not os.path.isdir(base):
        continue
    entries = []
    for entry in os.listdir(base):
        lower = entry.lower()
        if lower.startswith(("license", "copying", "notice")):
            entries.append(os.path.join(base, entry))
    if entries:
        result[name] = entries

print(json.dumps(result))
'@
Set-Content -Path $collectorScriptPath -Value $collectorScript -Encoding UTF8

# 収集結果（JSON）を取得し、PowerShell 側で複製
$licenseJson = & $pythonExe $collectorScriptPath @packages
Remove-Item $collectorScriptPath -Force

if ($licenseJson) {
    $licenseMap = $licenseJson | ConvertFrom-Json
    if ($licenseMap.PSObject.Properties.Name.Count -gt 0) {
        New-Item -ItemType Directory -Path $thirdPartyLicenseDir -Force | Out-Null
        foreach ($property in $licenseMap.PSObject.Properties) {
            $packageName = $property.Name
            foreach ($licensePath in $property.Value) {
                if (Test-Path $licensePath) {
                    $destinationName = "{0}_{1}" -f $packageName, (Split-Path $licensePath -Leaf)
                    Copy-Item -Path $licensePath -Destination (Join-Path $thirdPartyLicenseDir $destinationName) -Force
                }
            }
        }
    } else {
        Write-Warn "依存パッケージのライセンスファイルを検出できませんでした。"
    }
} else {
    Write-Warn "ライセンス収集スクリプトの結果が空でした。"
}

# ---- ランチャースクリプト生成 ---------------------------------------------
$launcherPath = Join-Path $repoRoot "lab_aid.bat"
$launcherContent = "@echo off`r`n""%~dp0$InstallDir\python.exe"" -m lab_aid.excel_cli %*`r`n"
Set-Content -Path $launcherPath -Value $launcherContent -Encoding ASCII
Write-Info "ランチャー 'lab_aid.bat' を生成しました。"

Write-Info "Excel 入力テンプレートを確認しています。"
Push-Location $repoRoot
try {
    & $pythonExe -m lab_aid.excel_cli --create-template
} finally {
    Pop-Location
}

Write-Info "セットアップが完了しました。"
Write-Host ""
Write-Host "Lab-Aid を実行するには次の手順を実行してください。"
Write-Host "  1. 'windows\\lab_aid_input.xlsx' に計算条件を入力する"
Write-Host "  2. windows\\run_lab_aid.cmd (または lab_aid.bat) を実行する"
Write-Host "  3. Excel に出力された結果を確認する"
Write-Host ""
