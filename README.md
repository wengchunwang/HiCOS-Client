# HiCOS Client 安裝與版本收集 Script

## 專案說明
本專案提供兩個主要功能：

1. **HiCOS_Client Minimal 安裝Script (`Install-HiCOS_Client_Minimal.ps1`)**
   - 檢查目標電腦是否已安裝 HiCOS Client。
   - 若未安裝或版本過舊，自動安裝指定版本。
   - 支援將安裝程式複製到本機暫存資料夾（C:\TEMP）後安裝。
   - 設計為 minimal，不產生 LOG 或 Debug 檔。
   - 適合透過 **GPO Startup Script** 部署。
   - **安裝檔來源說明**：HiCOS_Client.exe 可自 [https://moica.nat.gov.tw/download_1.html](https://moica.nat.gov.tw/download_1.html) 下載 ZIP 後，自行解壓縮。

2. **HiCOS 版本收集Script (`Get-HiCOSVersion_GPO.ps1`)**
   - 取得本機已安裝 HiCOS Client 的版本資訊（支援 32/64-bit 登錄檔）。
   - 將結果寫入 **CSV**，可存放於 NAS 或本地暫存。
   - 每筆紀錄包含：
     - `ComputerName`：電腦名稱
     - `DisplayName`：已安裝程式名稱
     - `Version`：已安裝版本號
     - `CheckedTime`：檢查時間
     - `LogPath`：實際寫入路徑（方便除錯）
     - `ScriptVersion`：Script版本號
   - 設計為核心類型物件，可在受限語言模式（GPO Startup Script）下安全執行。

## Script版本管理
- `Install-HiCOS_Client_Minimal.ps1` 版本：`2025.09.27.003`
- `Get-HiCOSVersion_GPO.ps1` 版本：`2025.09.27.003`

## 使用方式

### 1. Minimal 安裝Script
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "\\AD\SYSVOL\scripts\Install-HiCOS_Client_Minimal.ps1"
```

### 2. 版本收集Script
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "\\AD\SYSVOL\scripts\Get-HiCOSVersion_GPO.ps1"
```

## 注意事項
1. **NAS 權限**  
   - 確保電腦有寫入 NAS 的權限，否則版本收集Script會切換到本地暫存。
2. **GPO 受限語言模式**  
   - 版本收集Script已改用核心類型 `[PSObject]`，可在 GPO Startup Script 下執行。
3. **多臺電腦同時寫入 NAS**  
   - 若多臺電腦同時寫入同一份 CSV，可能造成檔案鎖定衝突。建議每臺電腦寫入單獨檔，再集中合併。

## 目錄結構建議
```
/HiCOS-Scripts
├─ Install-HiCOS_Client_Minimal.ps1
├─ Get-HiCOSVersion_GPO.ps1
└─ README.md
```

## 開發者資訊
- 作者：Weng, Chun-Wang
- 聯絡方式：wengchunwang@hotmail.com
- GitHub：https://github.com/wengchunwang/HiCOS-Client

## 版本紀錄
| 日期       | 版本         | 說明 |
|------------|--------------|------|
| 2025-09-27 | 2025.09.27.003 | Minimal 安裝 & 版本收集Script更新，支援 NAS 容錯與核心類型物件 |
| 2025-09-26 | 2025.09.26.001 | 初始版本，支援 GPO 安裝 & CSV 收集 |
