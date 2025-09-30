# Get-HiCOSVersion.ps1
# 腳本功能：收集本機 HiCOS 憑證元件版本，並將結果記錄到 NAS 或 C:\Temp。
param(
    [Parameter(Mandatory=$false)]
    [string]$NASPath = "\\NAS\Path"  # 設置一個預設值，如果 BAT 檔沒有傳入，則使用此值
)

# === 版本號 ===
$ScriptVersion = "2025.09.30.005"

# === 控制腳本輸出 (新加入) ===
# 設置偏好變數，確保所有警告和非終止錯誤不會輸出到控制台。
$WarningPreference = "SilentlyContinue"  # 靜默忽略所有警告 (例如 Write-Warning)
$ErrorActionPreference = "Stop"         # 僅在發生終止錯誤時才停止腳本

# === 路徑設定 (請根據您的環境修改 $NASPath) ===
$TempPath = "C:\Temp"
$FileName = "HiCOS_Versions_$(Get-Date -Format yyyyMMdd).csv"

# 預設日誌檔案路徑 (優先使用 NAS)
$LogFile = Join-Path $NASPath $FileName

# 建立資料夾並設定日誌路徑（容錯處理）
try {
    # 嘗試建立 NAS 資料夾
    if (-not (Test-Path $NASPath)) { 
        Write-Host "嘗試建立 NAS 資料夾：$NASPath"
        New-Item -Path $NASPath -ItemType Directory -Force | Out-Null 
    }
} catch {
    # 如果 NAS 失敗，切換到本地 C:\Temp
    Write-Warning "無法存取或建立 NAS 資料夾 $NASPath。切換至 $TempPath。"
    
    # 檢查並建立 C:\Temp
    try {
        if (-not (Test-Path $TempPath)) {
            Write-Host "嘗試建立本地暫存資料夾：$TempPath"
            New-Item -Path $TempPath -ItemType Directory -Force | Out-Null
        }
        # 更新日誌檔案路徑到本地
        $LogFile = Join-Path $TempPath $FileName
    } catch {
        # 如果連 C:\Temp 都無法建立，則發出致命錯誤並停止
        Write-Error "無法建立本地暫存資料夾 $TempPath。無法寫入日誌。腳本終止。"
        exit 1
    }
}

# --- 取得 HiCOS 安裝資訊（32/64-bit）---
$HiCOS = @()
# 搜尋 64-bit 登錄檔
$HiCOS += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
# 搜尋 32-bit 登錄檔（在 64-bit 系統上）
$HiCOS += Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue

# 篩選並只取第一個 HiCOS 項目
$HiCOS = $HiCOS | 
    Where-Object { $_.DisplayName -like "*HiCOS*" } | 
    Select-Object -First 1 DisplayName, DisplayVersion

# --- 建立核心輸出物件 ---
$Result = New-Object PSObject -Property @{
    ComputerName  = $env:COMPUTERNAME
    DisplayName   = if ($HiCOS) { $HiCOS.DisplayName } else { "未安裝" }
    Version       = if ($HiCOS) { $HiCOS.DisplayVersion } else { "" }
    CheckedTime   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    LogPath       = $LogFile  # 記錄實際寫入的路徑，方便除錯
    ScriptVersion = $ScriptVersion
}

# --- 將結果寫入 CSV ---

# 1. 定義您想要的欄位順序
$PropertyOrder = @(
    "ComputerName",
    "DisplayName",
    "Version",
    "CheckedTime",
    "ScriptVersion"
)

# 2. 使用 Select-Object 確保欄位順序後，再導出為 CSV
$Result | Select-Object -Property $PropertyOrder | Export-Csv -Path $LogFile -NoTypeInformation -Encoding UTF8 -Append:$Append

Write-Host "結果已寫入：$LogFile"
