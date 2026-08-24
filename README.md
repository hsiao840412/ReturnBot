# ReturnBot v3.0

ReturnBot 是一款專為 Apple 維修退料流程設計的 macOS 工具。匯入 ePacking List CSV 後，即可自動產生 KBB／Mail-in Excel 文件；特定類型也會同步建立 DHL 上傳用 CSV。

![macOS](https://img.shields.io/badge/macOS-26%20Tahoe-111111?logo=apple)
![Version](https://img.shields.io/badge/version-3.0-0A84FF)
![Architecture](https://img.shields.io/badge/architecture-Apple%20Silicon-555555)

## 功能

- 支援四種退料類型：
  - Mail-in KBB
  - Mail-in 電池膨脹
  - 一般 KBB
  - 單獨鋰電池 KBB
- 原生 SwiftUI 介面，採用 macOS Tahoe Liquid Glass 風格。
- 匯入 ePacking List CSV 後，一鍵產生完整 Excel 退料文件。
- 自動建立發票編號、填入工作表並依資料筆數調整列數。
- Mail-in KBB 與一般 KBB 可同步產生 DHL 貨物上傳 CSV。
- 單獨鋰電池 KBB 支援條碼工作表與公式填入。
- 自動辨識 CSV 編碼與檢查必要欄位、空白資料及退料類型。
- 遇到未辨識國家時仍會輸出，並在檔案內加入原始國家名稱備註。
- 啟動時預先要求 Excel 與「下載項目」存取權，減少產生途中被權限視窗中斷。

## 系統需求

- Apple Silicon Mac（arm64）
- macOS 26 Tahoe 或更新版本
- Microsoft Excel

## 安裝

1. 前往 [ReturnBot v3.0 Release](https://github.com/hsiao840412/ReturnBot/releases/tag/v3.0) 下載 `ReturnBot-v3.0-arm64.dmg`。
2. 開啟 DMG，將 ReturnBot 拖入「應用程式」。
3. 第一次啟動若 macOS 阻擋 App，請在 Finder 對 ReturnBot 按右鍵選擇「打開」，或至「系統設定 → 隱私權與安全性」允許開啟。
4. 依畫面提示授予 ReturnBot／Microsoft Excel 自動化及「下載項目」存取權。
5. 若使用單獨鋰電池條碼功能，請先安裝作業所需的條碼字體。

> 此版本採用 ad-hoc 簽署，未經 Apple Developer ID 公證，因此首次開啟可能出現安全提示。

## 使用方式

1. 選擇退料類型。
2. 選擇 ePacking List CSV。
3. 按下「生成 Excel 退料文件」。
4. 產生的檔案會儲存在「下載項目」資料夾。

## 從原始碼建置

此專案的 macOS 介面使用 SwiftUI，資料處理與 Excel 自動化則由 Python helper 執行。建置腳本會將兩者及必要資源封裝為 App 與 DMG。

```bash
./scripts/package_macos.sh
```

建置結果位於：

- `build/macos/ReturnBot.app`
- `build/macos/ReturnBot-v3.0-arm64.dmg`

Excel 範本包含作業格式與資料，因此不存放於公開 GitHub 倉庫；自行建置前需在專案根目錄準備對應範本。

## 隱私

ReturnBot 在本機讀取 CSV、操作 Microsoft Excel 並將結果存入「下載項目」。退料資料不會由 App 上傳至 GitHub。
