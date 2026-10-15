# 專案維護

## 目錄用途

| 位置 | 用途 |
| --- | --- |
| `Sources/ReturnBotMac/` | macOS SwiftUI 介面、截圖辨識與更新功能 |
| `returnbot_cli.py` | App 與 Python 文件產生流程之間的入口 |
| `ReturnBot.py` | 一般退料的資料驗證、文件產生與 DHL CSV 處理，仍由 CLI 使用 |
| `normal_invoice.py`、`kgb_template.py`、`invoice_common.py`、`recall.py` | Excel 文件與寄銷召回處理 |
| `packaging/` | App 設定、權限、更新歷史與正式圖示 |
| `scripts/` | 打包、更新簽署與更新說明產生工具 |
| `tests/` | Python 與 Swift 測試 |
| `build/macos-版本號/` | 該版本的 App、DMG、更新 ZIP 與 appcast |
| `dist/` | 自動產生的當次更新說明與完整更新歷史 |

正式圖示只維護 `packaging/AppIcon.icon/`，打包時由 Xcode 編譯。

## 必須保留的本機資源

目前 Excel 由程式直接寫入，不需啟動 Microsoft Excel；文件版面、圖片與寄件資料仍會讀取根目錄的五份 `.xlsx` 資源。這些檔案未納入 Git，重新建立開發環境時需要另行備妥：

- `mail-in template.xlsx`
- `mail-in swollen template.xlsx`
- `kbb template.xlsx`
- `battery kbb template.xlsx`
- `kgb template.xlsx`

`.venv/` 是 Python 執行與 PyInstaller 打包環境；`.build/` 是 Swift 建置與 Sparkle 工具所在位置，日常維護保留以加快建置。

## 清理原則

- 新版確認完成後，可刪除舊版的 `build/macos-版本號/`、試用 App 與驗證輸出。
- 最新版保留 `ReturnBot.app`、DMG、ZIP 與 `appcast.xml`；其餘打包中間檔可刪除，下次打包會重新產生。
- `__pycache__/`、`.pyc` 與 `.DS_Store` 可隨時清除。
- 實際案件、價格表與人工修改的文件需另行保存，不當作建置暫存清理。

發佈流程見 [macOS 更新發布](updates.md)。
