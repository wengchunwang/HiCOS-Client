# 🧩 HiCOS Client 自動安裝與更新Script (PowerShell)

## 🎯 專案簡介
本 PowerShell Script用於 **自動化企業環境中 HiCOS Client 的安裝、更新與版本管理**。  
可透過 **GPO（Group Policy Object）** 或 **排程任務** 執行，確保所有用戶端電腦上的 HiCOS Client 均維持在指定目標版本。  
Script具備高容錯性，並提供詳細的日誌記錄與遠端備份機制。

### 官方來源與說明
- **安裝檔來源**：[MOICA 內政部憑證管理中心 - 檔案下載](https://moica.nat.gov.tw/download_1.html)  
- **最新公告**：[HiCOS卡片管理工具與跨平臺網頁元件更新（2025-09-15）](https://moica.nat.gov.tw/news_in_17e9501c4f4000006dc2.html)
- **HiCOS 功能說明**：  
  HiCOS 卡片管理工具為 CSP（Cryptography Service Provider），提供 IC 卡憑證註冊至作業系統之功能，支援電子郵件簽章與加密應用。下載後包含：
  - HiCOS 卡片管理工具  
  - 用戶端環境檢測工具  
  - UP2Date Agent  
  - 相關使用手冊  

---

## ✨ 功能總覽

1. **自動版本比對**：比對登錄檔版本與 `$TargetVersion`，僅在版本過舊或未安裝時才執行。
2. **智慧安裝 / 更新機制**：避免重複安裝，並支援重試以提升成功率。
3. **本地暫存加速**：安裝檔先複製至 `C:\TEMP`，減少網路延遲。
4. **高容錯日誌機制**：  
   - 優先寫入 NAS。  
   - NAS 不可用時自動切換至本地路徑 (`C:\TEMP\HiCOS_Update_Local.log`)。  
5. **Debug Log 備份**：自動將 HiCOS 產生的 Debug Log 備份至 NAS。
6. **日誌保留與清理機制**：自動清除 NAS 上超過 `$Days` 天的舊檔案。
7. **靜默安裝支援**：支援無人值守安裝，適用於 GPO 環境。

---

## 🛠️ 環境需求與參數設定

### 系統需求
- **OS**：Windows 10/11 或 Windows Server  
- **PowerShell**：版本 5.1 以上  
- **執行權限**：需具備系統或本機管理者權限（GPO Startup Script 預設為 SYSTEM 帳號）  
- **網路需求**：
  - 能存取安裝檔來源 (`$InstallFiles`)  
  - 能存取並寫入日誌備份 NAS (`$NASPath`)

---

### Script參數 (`param`)

| 參數名稱 | 類型 | 預設值 | 說明 |
|-----------|------|---------|------|
| `NASPath` | `string` | `\\NAS\LogFiles` | NAS 備份與日誌存放路徑。需具備寫入權限。 |
| `GPO_NAME` | `string` | `Install-HiCOS_Client.ps1` | Script識別名稱，顯示於日誌。 |
| `Days` | `int` | `30` | 日誌檔在 NAS 上保留天數。 |
| `InstallFiles` | `string[]` | (必填) | HiCOS 安裝檔網路路徑（陣列形式，依序尋找可用來源）。 |

**範例 GPO 啟動Script參數：**
```powershell
-NASPath "\\YourDomainNAS\IT_Logs\HiCOS" -InstallFiles "\\FileServer\Software\HiCOS\HiCOS_Client.exe"
```

---

## 🚀 部署教學

### 1️⃣ 準備安裝檔案
將 `HiCOS_Client.exe`（或等效靜默安裝檔）放置於可供所有用戶端存取的網路共享資料夾。

### 2️⃣ 配置 GPO 啟動Script
1. 於網域控制站開啟 Group Policy Management。  
2. 編輯目標 OU 的 GPO。  
3. 前往 **Computer Configuration → Policies → Windows Settings → Scripts (Startup/Shutdown)**。  
4. 選取 **Startup → PowerShell Scripts → Add**。  
5. 指定 `Install-HiCOS_Client.ps1` 為執行檔。  
6. 在參數欄輸入上述設定。

### 3️⃣ 驗證日誌輸出
- **主要路徑 (NAS)**：`\\NAS\LogFiles\HiCOS_Update_YYYYMMDD.log`  
- **備援路徑 (本機)**：`C:\TEMP\HiCOS_Update_Local.log`

---

## 📂 專案結構建議

```
/HiCOS-Scripts
├─ Install-HiCOS_Client.ps1
├─ Install-HiCOS_Client_Minimal.ps1
├─ Get-HiCOSVersion_GPO.ps1
└─ README.md
```

---

## 🔧 相關Script說明

| Script名稱 | 功能 | 備註 |
|-----------|------|------|
| `Install-HiCOS_Client.ps1` | 自動偵測版本並執行安裝或更新。 | 主Script（含日誌、容錯機制） |
| `Install-HiCOS_Client_Minimal.ps1` | 簡化版安裝Script，不產生日誌。 | 適用無需追蹤的環境 |
| `Get-HiCOSVersion_GPO.ps1` | 收集 HiCOS 安裝版本資訊並輸出 CSV。 | 支援 32/64-bit 登錄檔查詢與 NAS 備援 |

### `Get-HiCOSVersion_GPO.ps1` 產出欄位：
| 欄位 | 說明 |
|------|------|
| `ComputerName` | 電腦名稱 |
| `DisplayName` | 程式名稱 |
| `Version` | 已安裝版本號 |
| `CheckedTime` | 檢查時間 |
| `LogPath` | 實際輸出路徑 |
| `ScriptVersion` | Script版本號 |

> 💡 建議：若多臺電腦同時寫入同一 CSV，請採用獨立檔案命名方式（例如 `hostname.csv`），再集中彙整。

---

## ⚙️ 主要變數一覽

| 變數 | 說明 | 範例值 |
|------|------|--------|
| `$ScriptVersion` | Script版本 | `2025.10.03.016` |
| `$TargetVersion` | HiCOS 目標版本 | `3.2.2.00901` |
| `$LocalTempPath` | 暫存路徑 | `C:\TEMP` |
| `$MaxRetries` | 安裝重試次數 | `2` |
| `Write-Log` | 日誌函數，具 NAS 容錯 | — |

---

## 🧾 版本紀錄

| 日期 | 版本 | 說明 |
|------|------|------|
| **2025-10-15** | **2025.10.03.016** | 修正：安裝失敗時未刪除暫存檔（`Remove-Item`）。<br>優化：分離 NAS/Local Log 檔頭，確保完整記錄。 |
| 2025-10-01 | 2025.10.01.010 | 加入 `InstallFiles` 驗證；安裝成功後清理暫存檔；日誌訊息分類重整。 |
| 2025-09-27 | 2025.09.27.003 | Minimal 安裝版與版本收集 Script 更新，支援核心物件與 NAS 容錯。 |
| 2025-09-26 | 2025.09.26.001 | 初始版本：支援 GPO 安裝與 CSV 收集。 |

---

## 👨‍💻 開發者資訊

| 項目 | 資訊 |
|------|------|
| 作者 | **Weng, Chun-Wang** |
| 聯絡方式 | [wengchunwang@hotmail.com](mailto:wengchunwang@hotmail.com) |
| GitHub | [github.com/wengchunwang/HiCOS-Client](https://github.com/wengchunwang/HiCOS-Client) |
