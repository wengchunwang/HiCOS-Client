# === 目標 HiCOS 版本 ===
$TargetVersion = "3.2.2.00901"

# === 安裝程式網路路徑清單 ===
$InstallFiles  = @(
    "\\NAS\Path\HiCOS_Client.exe",
    "\\AD\SYSVOL\scripts\HiCOS_Client.exe"
)

# === 本機暫存路徑 ===
$LocalTempPath = "C:\TEMP"
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

# === 檢查是否已安裝 HiCOS ===
$HiCOS = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" `
        -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -like "*HiCOS*" } |
        Select-Object -First 1 DisplayName, DisplayVersion

$NeedInstall = -not $HiCOS -or (Compare-Version $HiCOS.DisplayVersion $TargetVersion)

if ($NeedInstall) {
    # 建立本機暫存資料夾
    if (-not (Test-Path $LocalTempPath)) { New-Item -Path $LocalTempPath -ItemType Directory -Force }

    # 複製安裝程式到本機
    foreach ($path in $InstallFiles) {
        if (Test-Path $path) {
            Copy-Item -Path $path -Destination $LocalInstaller -Force
            break
        }
    }

    if (Test-Path $LocalInstaller) {
        # 安裝
        for ($i=1; $i -le $MaxRetries+1; $i++) {
            Start-Process -FilePath $LocalInstaller -ArgumentList "/S" -Wait

            # 檢查版本
            $HiCOSNew = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" `
                       -ErrorAction SilentlyContinue |
                       Where-Object { $_.DisplayName -like "*HiCOS*" } |
                       Select-Object -First 1 DisplayName, DisplayVersion

            if ($HiCOSNew -and -not (Compare-Version $HiCOSNew.DisplayVersion $TargetVersion)) { break }
            Start-Sleep -Seconds 10
        }

        # 刪除本機安裝檔
        #Remove-Item $LocalInstaller -Force
    }
}
