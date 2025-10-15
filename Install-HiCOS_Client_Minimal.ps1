<#
.SYNOPSIS
  HiCOS Client 自動安裝與更新（簡潔 版）
.DESCRIPTION
  - 檢查 HiCOS 是否已安裝
  - 比對版本號，若未安裝或版本較舊則自動安裝
  - 自動複製安裝檔到本機 C:\TEMP 執行
  - 支援安裝重試機制
.NOTES
  作者: IT 部門 Wang Wang
  版本: 2025.10.15.簡潔
  適用於 GPO Startup Script 或本機執行測試
#>

param(
    [string]$NASPath = "\\NAS\LogFiles",
    [string]$GPO_NAME = "",
    [int]$Days = 30,  # 保留參數（未使用）
    [string[]]$InstallFiles = @(
        "\\FileSrv\SoftwareDeploy\HiCOS_Client.exe",
        "\\ADSrv\scripts\"
    ), # === 安裝程式網路路徑清單 ===
    [string]$TargetVersion = "3.2.2.00901" # === 目標 HiCOS 版本 ===
)

# === 偏好設定 ===
$WarningPreference = "SilentlyContinue"
$ErrorActionPreference = "Stop"

# === 本機暫存路徑 ===
$LocalTempPath  = "$env:SystemDrive\TEMP"
$LocalInstaller = Join-Path $LocalTempPath "HiCOS_Client.exe"

# === 安裝重試次數 ===
$MaxRetries = 2

# === 版本比對函數 ===
function Compare-Version($v1, $v2) {
    try {
        return [version]($v1 -replace '[^\d\.]','') -lt [version]($v2 -replace '[^\d\.]','')
    } catch {
        return $true
    }
}

# === 取得已安裝版本 (32/64-bit) ===
function Get-HiCOS {
    $apps = @()
    $apps += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
    $apps += Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
    return $apps | Where-Object { $_.DisplayName -like "*HiCOS*" } | Select-Object -First 1 DisplayName, DisplayVersion
}

# === 主程序 ===
$HiCOS = Get-HiCOS
$NeedInstall = -not $HiCOS -or (Compare-Version $HiCOS.DisplayVersion $TargetVersion)

if ($NeedInstall) {
    if (-not (Test-Path $LocalTempPath)) { New-Item -Path $LocalTempPath -ItemType Directory -Force | Out-Null }

    # 複製安裝檔到本機
    foreach ($InstallFile in $InstallFiles) {
        if (Test-Path $InstallFile -PathType Leaf) {
            Copy-Item -Path $InstallFile -Destination $LocalInstaller -Force
            break
        }
    }

    if (Test-Path $LocalInstaller) {
        # 執行安裝（含重試）
        for ($i = 1; $i -le $MaxRetries; $i++) {
            Start-Process -FilePath $LocalInstaller -ArgumentList "/S" -Wait
            $HiCOSNew = Get-HiCOS
            if ($HiCOSNew -and -not (Compare-Version $HiCOSNew.DisplayVersion $TargetVersion)) { break }
            Start-Sleep -Seconds 10
        }

        # 可選：刪除安裝檔
        # Remove-Item $LocalInstaller -Force -ErrorAction SilentlyContinue
    }
}
