<#
.SYNOPSIS
  收集本機 HiCOS 憑證元件版本，並寫入 NAS 或本地暫存 CSV。
.DESCRIPTION
  - 支援 32/64-bit Windows 安裝檢查
  - 優先寫入 NAS，若無法存取則寫入 C:\Temp
  - 每台電腦生成獨立紀錄，可 Append 到同一 CSV
.NOTES
  作者: IT 部門 Wang Wang
  版本: 2025.10.15.006
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$NASPath = "\\DEFAULT_NAS\path",  # 作為備用/測試用的預設路徑
    [string]$GPO_NAME = "Get-HiCOSVersion.ps1"
)

# === 版本號 ===
$ScriptVersion = "2025.10.15.006"

# === 控制輸出偏好 ===
$WarningPreference = "SilentlyContinue"
$ErrorActionPreference = "Stop"

# === 路徑設定 ===
$LocalTempPath = "$env:SystemDrive\TEMP"
$FileName = "HiCOS_Versions_$(Get-Date -Format yyyyMMdd).csv"

# 預設日誌檔案路徑 (優先使用 NAS)
$LogFile = Join-Path $NASPath $FileName

# 建立資料夾並設定日誌路徑（容錯處理）
try {
    if (-not (Test-Path $NASPath)) {
        New-Item -Path $NASPath -ItemType Directory -Force | Out-Null
    }
} catch {
    Write-Warning "無法存取或建立 NAS 資料夾 $NASPath，切換至 $LocalTempPath。"
    try {
        if (-not (Test-Path $LocalTempPath)) {
            New-Item -Path $LocalTempPath -ItemType Directory -Force | Out-Null
        }
        $LogFile = Join-Path $LocalTempPath $FileName
    } catch {
        Write-Error "無法建立本地暫存資料夾 $LocalTempPath，腳本終止。"
        exit 2
    }
}

# === 取得 HiCOS 安裝資訊（32/64-bit）===
$HiCOS = @()
$HiCOS += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
$HiCOS += Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
$HiCOS = $HiCOS | Where-Object { $_.DisplayName -like "*HiCOS*" } | Select-Object -First 1 DisplayName, DisplayVersion

# === 建立核心輸出物件 ===
$Result = [PSCustomObject]@{
    ComputerName  = $env:COMPUTERNAME
    DisplayName   = if ($HiCOS) { $HiCOS.DisplayName } else { "未安裝" }
    Version       = if ($HiCOS) { $HiCOS.DisplayVersion } else { "" }
    CheckedTime   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogPath       = $LogFile
    ScriptVersion = $ScriptVersion
    GPOName       = $GPO_NAME
}

# === 將結果寫入 CSV ===
$PropertyOrder = @(
    "ComputerName",
    "DisplayName",
    "Version",
    "CheckedTime",
    "ScriptVersion",
    "GPOName"
)

# 如果檔案已存在就 Append，否則建立新檔
$Append = Test-Path $LogFile
$Result | Select-Object -Property $PropertyOrder | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8 -Append:$Append

Write-Host "結果已寫入：$LogFile"
