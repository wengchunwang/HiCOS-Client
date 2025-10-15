<#
.SYNOPSIS
  HiCOS Client 自動安裝與更新 (本機安裝優化最終版)
.DESCRIPTION
  - 檢查 HiCOS Client 是否已安裝
  - 比對版本號 $TargetVersion
  - 未安裝或版本較舊則自動安裝
  - 支援安裝失敗重試機制
  - 安裝程式先複製到本機 C:\Temp，減少網路耗時
  - 安裝完成後將 HiCOS Client Debug log 複製到 NAS
  - LOG 記錄完整流程，統一編號格式 [001: Message]
.NOTES
  作者: IT 部門 Wang Wang
  版本: 2025.10.15.002 (最終正式版)
  建立日期: 2025-09-26
  更新日期: 2025-10-15
  使用於 GPO Startup Script
#>

param(
    [string]$NASPath = "\\NAS\LogFiles",
    [string]$GPO_NAME = "Install-HiCOS_Client.ps1",
    [int]$Days = 30,
    [string[]]$InstallFiles
)

# === 偏好設定 ===
$WarningPreference = "SilentlyContinue"
$ErrorActionPreference = "Stop"

# === 版本號 ===
$ScriptVersion = "2025.10.15.002"

# === 目標 HiCOS 版本 ===
$TargetVersion = "3.2.2.00901"

# === 本機暫存路徑 ===
$LocalTempPath = "$env:SystemDrive\TEMP"
$LocalInstaller = Join-Path $LocalTempPath "HiCOS_Client.exe"
$LocalDebugFile = Join-Path $LocalTempPath "HiCOS_Client_Debug.log"
$LocalLog = Join-Path $LocalTempPath "HiCOS_Update_Local.log"

# 建立本機暫存資料夾
if (-not (Test-Path $LocalTempPath)) { New-Item -Path $LocalTempPath -ItemType Directory -Force | Out-Null }

# === LOG 檔案設定 ===
$LogFile = Join-Path $NASPath ("Update_" + (Get-Date -Format yyyyMMdd) + ".log")

# === 狀態變數 ===
$Installed = $false
$NeedInstall = $false
$MaxRetries = 2

# === 日誌函數 ===
$script:LogIndex = 0
function Write-Log($msg) {
    $script:LogIndex++
    $IndexFormatted = "{0:D3}" -f $script:LogIndex
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Line = "[${TimeStamp}_${env:COMPUTERNAME}_${GPO_NAME}] ${IndexFormatted}: ${msg}"

    try {
        Add-Content -Path $LogFile -Value $Line -Encoding Unicode -ErrorAction Stop
    } catch {
        Add-Content -Path $LocalLog -Value $Line -Encoding Unicode
    }
}

# === 初始化 LOG 檔頭 ===
$LogHeader = @"
# =========================================================
# HiCOS Client 安裝/更新 LOG 啟動
# 版本: $ScriptVersion
# 目標 HiCOS 版本: $TargetVersion
# 建立日期: $(Get-Date -Format "yyyy-MM-dd")
# 執行來源 GPO: $GPO_NAME
# =========================================================
"@
try {
    Add-Content -Path $LogFile -Value $LogHeader -Encoding Unicode -ErrorAction SilentlyContinue
} catch {
    Add-Content -Path $LocalLog -Value $LogHeader -Encoding Unicode -ErrorAction SilentlyContinue
}

# === 參數驗證 ===
if (-not $InstallFiles -or $InstallFiles.Count -eq 0) {
    Write-Log "錯誤：GPO 參數 InstallFiles 未提供或為空，無法開始安裝。"
    exit 1
}

# === 版本比對函數 ===
function Compare-Version($v1, $v2) {
    try {
        $CleanV1 = $v1 -replace '[^\d\.]',''
        $CleanV2 = $v2 -replace '[^\d\.]',''
        return [version]($CleanV1) -lt [version]($CleanV2)
    } catch {
        Write-Log "警告：版本比對失敗，強制進行安裝。（版本字串格式錯誤：$v1 或 $v2）"
        return $true
    }
}

# === 檢查 HiCOS 安裝狀態 ===
$HiCOS = @()
$HiCOS += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
$HiCOS += Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
$HiCOS = $HiCOS | Where-Object { $_.DisplayName -like "*HiCOS*" } | Select-Object -First 1 DisplayName, DisplayVersion

if (-not $HiCOS) {
    Write-Log "未安裝 HiCOS，需要安裝"
    $NeedInstall = $true
} elseif (Compare-Version $HiCOS.DisplayVersion $TargetVersion) {
    Write-Log "已安裝版本 $($HiCOS.DisplayVersion) < 目標版本 $TargetVersion，需要更新"
    $NeedInstall = $true
} else {
    Write-Log "已安裝版本 $($HiCOS.DisplayVersion) >= 目標版本 $TargetVersion，不需更新"
}

# === 執行安裝與重試 ===
if ($NeedInstall) {
    foreach ($InstallFile in $InstallFiles) {
        if (Test-Path $InstallFile) {
            try {
                Copy-Item -Path $InstallFile -Destination $LocalInstaller -Force
                Write-Log "安裝程式已複製到本機：$LocalInstaller"
            } catch {
                Write-Log "複製安裝程式失敗：$($_.Exception.Message)"
                return
            }
            break
        }
    }

    if (-not (Test-Path $LocalInstaller)) {
        Write-Log "本機安裝檔不存在，無法安裝，結束"
        return
    }

    for ($i=1; $i -le $MaxRetries+1; $i++) {
        Write-Log "安裝嘗試 #$i，本機來源：$LocalInstaller"
        try {
            Start-Process -FilePath $LocalInstaller -ArgumentList "/S /L `"$LocalDebugFile`"" -Wait

            # 重新檢查版本
            $HiCOSNew = @()
            $HiCOSNew += Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
            $HiCOSNew += Get-ItemProperty "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue
            $HiCOSNew = $HiCOSNew | Where-Object { $_.DisplayName -like "*HiCOS*" } | Select-Object -First 1 DisplayName, DisplayVersion

            if ($HiCOSNew -and -not (Compare-Version $HiCOSNew.DisplayVersion $TargetVersion)) {
                Write-Log "安裝成功，版本：$($HiCOSNew.DisplayVersion)"
                $Installed = $true
                if (Test-Path $LocalInstaller) {
                    Remove-Item $LocalInstaller -Force -ErrorAction SilentlyContinue
                    Write-Log "安裝成功，已刪除本機安裝檔：$LocalInstaller"
                }
                break
            } else {
                Write-Log "安裝後版本仍未達目標 $TargetVersion"
            }
        } catch {
            Write-Log "安裝或版本檢查失敗：$($_.Exception.Message)"
        }

        if ($i -le $MaxRetries) {
            Write-Log "等待 10 秒後重試..."
            Start-Sleep -Seconds 10
        }
    }

    if (-not $Installed) { Write-Log "安裝失敗，已達最大重試次數 $MaxRetries" }

    # 複製 Debug log 到 NAS
    try {
        if (Test-Path $LocalDebugFile) {
            if (Test-Path $NASPath -PathType Container) {
                $DestPath = Join-Path $NASPath "$env:COMPUTERNAME`_HiCOS_Debug.log"
                Copy-Item -Path $LocalDebugFile -Destination "$DestPath" -Force -ErrorAction Stop
                Write-Log "Debug log 複製到 NAS：$DestPath"
                Remove-Item $LocalDebugFile -Force -ErrorAction SilentlyContinue
            } else {
                Write-Log "NAS 不可寫，保留本機 Debug log：$LocalDebugFile"
            }
        } else {
            Write-Log "本機 Debug log 不存在，無法複製到 NAS"
        }
    } catch {
        Write-Log "Debug log 處理失敗：$($_.Exception.Message)"
    }
} else {
    Write-Log "跳過安裝"
}

# 安裝失敗仍刪除本機安裝檔
if (-not $Installed -and (Test-Path $LocalInstaller)) {
    try {
        Remove-Item $LocalInstaller -Force -ErrorAction SilentlyContinue
        Write-Log "本機安裝檔已刪除：$LocalInstaller (後續清理)"
    } catch {
        Write-Log "刪除本機安裝檔失敗：$($_.Exception.Message)"
    }
}

# 清理過期 LOG 檔案
try {
    Write-Log "檢查並清理超過 $Days 天的歷史 LOG..."
    $CutoffDate = (Get-Date).AddDays(-$Days)
    Get-ChildItem -Path $NASPath -Filter "Update_*.log" |
        Where-Object { $_.CreationTime -lt $CutoffDate } |
        Remove-Item -Force -ErrorAction SilentlyContinue | Out-Null
} catch {
    Write-Log "清理歷史 LOG 檔案失敗：$($_.Exception.Message)"
}

Write-Log "=== 執行結束 (版本 $ScriptVersion) ==="
exit 0
