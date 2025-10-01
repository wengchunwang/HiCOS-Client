<#
.SYNOPSIS
    HiCOS Client 自動安裝與更新 (本機安裝優化)
.DESCRIPTION
    - 檢查 HiCOS Client 是否已安裝
    - 比對版本號 $TargetVersion
    - 未安裝或版本較舊則自動安裝
    - 支援安裝失敗重試機制
    - 安裝程式先複製到本機 C:\Temp，減少網路耗時
    - 安裝完成後將 HiCOS Client Debug log 複製到 NAS
    - LOG 記錄完整流程，包括版本比對、安裝結果、Debug log 備份
.NOTES
    作者: IT 部門 Wang Wang
    版本: 2025.10.01.010
    建立日期: 2025-09-26
    更新日期: 2025-10-01
    使用於 GPO Startup Script
#>
param(
    [string]$NASPath = "\\NAS\LogFiles",
    [string]$GPO_NAME = "",
    [int]$Days = 30, # RetentionDays
    [string[]]$InstallFiles
)

# === 偏好設定 ===
$WarningPreference = "SilentlyContinue"
$ErrorActionPreference = "Stop"

# === 版本號 ===
$ScriptVersion = "2025.10.01.010"

# === 目標 HiCOS 版本 ===
$TargetVersion = "3.2.2.00901"

# === 本機暫存路徑 ===
$LocalTempPath = "C:\Temp"
$LocalInstaller = Join-Path $LocalTempPath "HiCOS_Client.exe"
$LocalDebugFile = Join-Path $LocalTempPath "HiCOS_Client_Debug.log"
$LocalLog = Join-Path $LocalTempPath "HiCOS_Update_Local.log" # NAS 寫入失敗時，轉為寫入本機 $LocalLog

# === 建立本機暫存資料夾 (C:\Temp 通常存在，但保留檢查) ===
if (-not (Test-Path $LocalTempPath)) { New-Item -Path $LocalTempPath -ItemType Directory -Force | Out-Null }

# === LOG 設定 ===
$LogFile = Join-Path $NASPath "Update_$(Get-Date -Format yyyyMMdd).log"

# === 安裝重試次數 ===
$MaxRetries = 2

# === LOG 函數 (已加入 Unicode 編碼修正亂碼問題) ===
function Write-Log($msg) {
    # 建立 LOG 檔頭 (若不存在)
    if (-not (Test-Path $LogFile)) {
        Add-Content -Path $LogFile -Value @"
# =========================================================
# HiCOS Client 安裝/更新 LOG
# 版本: $ScriptVersion
# 目標 HiCOS 版本: $TargetVersion
# 建立日期: $(Get-Date -Format "yyyy-MM-dd")
# 執行來源 GPO: $GPO_NAME
# 記錄內容: 每次啟動 GPO 或手動執行
# =========================================================
"@ -Encoding Unicode # 使用 Unicode 編碼確保 LOG 檔頭正確
    }
    try {
        Add-Content -Path $LogFile -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $env:COMPUTERNAME - $msg" -Encoding Unicode
    } catch {
        # NAS 寫入失敗時，轉為寫入本機 $LocalLog
        Add-Content -Path $LocalLog -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $env:COMPUTERNAME - $msg" -Encoding Unicode
    }
}

# === 參數驗證 (新增) ===
if (-not $InstallFiles -or $InstallFiles.Count -eq 0) {
    Write-Log "001: 錯誤：GPO 參數 InstallFiles 未提供或為空，無法開始安裝。"
    exit 1
}

# === 版本比對函數 ===
function Compare-Version($v1, $v2) {
    try {
        return [version]($v1 -replace '[^\d\.]','') -lt [version]($v2 -replace '[^\d\.]','')
    } catch {
        Write-Log "000: 警告：版本比對失敗，強制進行安裝。"
        return $true
    }
}

# === 檢查 HiCOS 安裝狀態 (32/64-bit) ===
$HiCOS = @()
$HiCOS += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
$HiCOS += Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
$HiCOS = $HiCOS | Where-Object { $_.DisplayName -like "*HiCOS*" } | Select-Object -First 1 DisplayName, DisplayVersion

if (-not $HiCOS) {
    Write-Log "102: 未安裝 HiCOS，需要安裝"
    $NeedInstall = $true
} elseif (Compare-Version $HiCOS.DisplayVersion $TargetVersion) {
    Write-Log "103: 已安裝版本 $($HiCOS.DisplayVersion) < 目標版本 $TargetVersion，需要更新"
    $NeedInstall = $true
} else {
    Write-Log "104: 已安裝版本 $($HiCOS.DisplayVersion) >= 目標版本 $TargetVersion，不需更新"
    $NeedInstall = $false
}

# === 執行安裝與重試機制 ===
if ($NeedInstall) {
    $Installed = $false
    foreach ($InstallFile in $InstallFiles) {
        if (Test-Path $InstallFile) {
            try {
                Copy-Item -Path $InstallFile -Destination $LocalInstaller -Force
                Write-Log "205: 安裝程式已複製到本機：$LocalInstaller"
            } catch {
                Write-Log "206: 複製安裝程式失敗：$($_.Exception.Message)"
                return
            }
            break
        }
    }
    if (-not (Test-Path $LocalInstaller)) {
        Write-Log "207: 本機安裝檔不存在，無法安裝，結束"
        return
    }

    for ($i=1; $i -le $MaxRetries+1; $i++) {
        Write-Log "208: 安裝嘗試 #$i，本機來源：$LocalInstaller"
        try {
            Start-Process -FilePath $LocalInstaller -ArgumentList "/S /L `"$LocalDebugFile`"" -Wait

            # 重新檢查版本
            $HiCOSNew = @()
            $HiCOSNew += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
            $HiCOSNew += Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
            $HiCOSNew = $HiCOSNew | Where-Object { $_.DisplayName -like "*HiCOS*" } | Select-Object -First 1 DisplayName, DisplayVersion

            if ($HiCOSNew -and -not (Compare-Version $HiCOSNew.DisplayVersion $TargetVersion)) {
                Write-Log "209: 安裝成功，版本：$($HiCOSNew.DisplayVersion)"
                $Installed = $true
                
                # === 優化：安裝成功後立即清理本機安裝檔 ===
                Write-Log "210: 安裝成功，立即清理本機安裝檔：$LocalInstaller"
                Remove-Item $LocalInstaller -Force -ErrorAction SilentlyContinue
                # ==========================================
                
                break
            } else {
                Write-Log "211: 安裝後版本仍未達目標 $TargetVersion"
            }
        } catch {
            Write-Log "212: 安裝或版本檢查失敗：$($_.Exception.Message)"
        }
        if ($i -le $MaxRetries) {
            Write-Log "213: 等待 10 秒後重試..."
            Start-Sleep -Seconds 10
        }
    }

    if (-not $Installed) { Write-Log "214: 安裝失敗，已達最大重試次數 $MaxRetries" }

    # === 複製 Debug log 到 NAS ===
    try {
        if (Test-Path $LocalDebugFile) {
            if (Test-Path $NASPath -PathType Container) {
                $DestPath = Join-Path $NASPath "$env:COMPUTERNAME`_HiCOS_Debug.log"
                Copy-Item -Path $LocalDebugFile -Destination $DestPath -Force -ErrorAction SilentlyContinue
                Write-Log "315: Debug log 複製到 NAS：$DestPath"
            } else {
                Write-Log "316: NAS 路徑不可寫，跳過 Debug log 備份"
            }
            Remove-Item $LocalDebugFile -Force -ErrorAction SilentlyContinue
        } else {
            Write-Log "317: 本機 Debug log 不存在，無法複製到 NAS"
        }
    } catch {
        Write-Log "318: Debug log 處理失敗：$($_.Exception.Message)"
    }
} else {
    Write-Log "219: 跳過安裝"
}

# 可選：刪除本機安裝檔 (若安裝失敗，則在此處執行刪除)
if (-not $Installed) {
    try {
        Remove-Item $LocalInstaller -Force -ErrorAction SilentlyContinue
        Write-Log "420: 本機安裝檔已刪除：$LocalInstaller (後續清理)"
    } catch {
        Write-Log "421: 刪除本機安裝檔失敗：$($_.Exception.Message)"
    }
}
	
# === 清理過期 LOG 檔案 ===
try {
    Write-Log "522: 檢查並清理超過 $Days 天的歷史 LOG..."
    $CutoffDate = (Get-Date).AddDays(-$Days)
    Get-ChildItem -Path $NASPath -Filter "Update_*.log" |
        Where-Object { $_.CreationTime -lt $CutoffDate } |
        Remove-Item -Force -ErrorAction SilentlyContinue | Out-Null
} catch {
    Write-Log "523: 清理歷史 LOG 檔案失敗：$($_.Exception.Message)"
}
