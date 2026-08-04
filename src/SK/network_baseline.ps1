# =============================================================================
# SK サーバー移行 ネットワーク基礎性能測定スクリプト
#
# Access のベンチマーク (modBenchmark.bas) とは別に、ネットワーク自体の
# 「遅延(RTT)」「大容量スループット」「小サイズ多数操作」を素の状態で測る。
# Access の結果が悪いとき、原因がネットワークか Access かを切り分けるために使う。
#
# 使い方 (PowerShell を通常権限で):
#   .\network_baseline.ps1 -Label "旧サーバー" -TargetFolder "<<旧サーバーの検証用フォルダUNCパス>>"
#   .\network_baseline.ps1 -Label "新サーバー" -TargetFolder "<<新サーバーの検証用フォルダUNCパス>>"
#
# 結果はスクリプトと同じフォルダの network_baseline_results.csv に追記される。
# 対象フォルダに一時ファイルを作成するが、終了時にすべて削除する。
# =============================================================================
param(
    [Parameter(Mandatory = $true)][string]$Label,
    [Parameter(Mandatory = $true)][string]$TargetFolder,
    [int]$LargeMB = 100,
    [int]$SmallCount = 200
)

$ErrorActionPreference = "Stop"
$resultCsv = Join-Path $PSScriptRoot "network_baseline_results.csv"

function Write-Result([string]$test, [double]$value, [string]$unit, [string]$detail) {
    if (-not (Test-Path $resultCsv)) {
        "日時,サーバーラベル,テスト項目,値,単位,詳細" | Out-File $resultCsv -Encoding utf8
    }
    $line = "{0},{1},{2},{3},{4},{5}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Label, $test, [math]::Round($value, 2), $unit, $detail
    $line | Out-File $resultCsv -Append -Encoding utf8
    Write-Host ("  {0}: {1} {2}  {3}" -f $test, [math]::Round($value, 2), $unit, $detail)
}

if (-not (Test-Path $TargetFolder)) {
    Write-Host "フォルダーにアクセスできません: $TargetFolder" -ForegroundColor Red
    exit 1
}

Write-Host "=== ネットワーク基礎性能測定: $Label ($TargetFolder) ==="

# ---- 1. RTT (ping) --------------------------------------------------------
# UNC パスからサーバー名を取り出す
$server = ($TargetFolder -replace '^\\\\', '') -split '\\' | Select-Object -First 1
try {
    $pings = Test-Connection -ComputerName $server -Count 10
    $rtts = $pings | ForEach-Object { $_.ResponseTime }
    $avg = ($rtts | Measure-Object -Average).Average
    $max = ($rtts | Measure-Object -Maximum).Maximum
    Write-Result "1_RTT(ping平均)" $avg "ms" ("最大 {0}ms / 10回" -f $max)
} catch {
    Write-Result "1_RTT(ping平均)" -1 "ms" "ping不可(ICMP遮断の可能性)"
}

# ---- 準備: ローカル一時ファイル --------------------------------------------
$workLocal = Join-Path $env:TEMP ("netbench_" + [guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $workLocal | Out-Null
$workRemote = Join-Path $TargetFolder ("netbench_" + $env:COMPUTERNAME)
New-Item -ItemType Directory -Force -Path $workRemote | Out-Null

$largeLocal = Join-Path $workLocal "large.bin"
$buf = New-Object byte[] (1MB)
(New-Object Random).NextBytes($buf)
$fs = [IO.File]::OpenWrite($largeLocal)
for ($i = 0; $i -lt $LargeMB; $i++) { $fs.Write($buf, 0, $buf.Length) }
$fs.Close()

try {
    # ---- 2. 大容量書き込み (ローカル → サーバー) ---------------------------
    $largeRemote = Join-Path $workRemote "large.bin"
    $t = Measure-Command { Copy-Item $largeLocal $largeRemote }
    Write-Result "2_大容量書込" ($LargeMB / $t.TotalSeconds) "MB/秒" ("{0}MB を {1}秒" -f $LargeMB, [math]::Round($t.TotalSeconds, 1))

    # ---- 3. 大容量読み取り (サーバー → ローカル) ---------------------------
    $largeBack = Join-Path $workLocal "large_back.bin"
    $t = Measure-Command { Copy-Item $largeRemote $largeBack }
    Write-Result "3_大容量読取" ($LargeMB / $t.TotalSeconds) "MB/秒" ("{0}MB を {1}秒" -f $LargeMB, [math]::Round($t.TotalSeconds, 1))

    # ---- 4. 小ファイル多数作成 (遅延に敏感 = Access的なアクセスパターン) ----
    $smallLocal = Join-Path $workLocal "small.bin"
    [IO.File]::WriteAllBytes($smallLocal, $buf[0..4095])
    $t = Measure-Command {
        for ($i = 0; $i -lt $SmallCount; $i++) {
            Copy-Item $smallLocal (Join-Path $workRemote ("s{0:0000}.bin" -f $i))
        }
    }
    Write-Result "4_小ファイル書込" ($t.TotalMilliseconds / $SmallCount) "ms/件" ("4KB × {0}件 合計{1}秒" -f $SmallCount, [math]::Round($t.TotalSeconds, 1))

    # ---- 5. 小ファイル列挙+読み取り ----------------------------------------
    $t = Measure-Command {
        Get-ChildItem $workRemote -Filter "s*.bin" | ForEach-Object {
            [IO.File]::ReadAllBytes($_.FullName) | Out-Null
        }
    }
    Write-Result "5_小ファイル読取" ($t.TotalMilliseconds / $SmallCount) "ms/件" ("4KB × {0}件" -f $SmallCount)
}
finally {
    # ---- 後片付け ----------------------------------------------------------
    try { Remove-Item -Recurse -Force $workRemote -ErrorAction Stop } catch {}
    try { Remove-Item -Recurse -Force $workLocal -ErrorAction Stop } catch {}
}

Write-Host "=== 完了。結果: $resultCsv ==="
