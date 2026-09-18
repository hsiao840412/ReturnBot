# macOS 更新發布

App 使用 Sparkle 2.10.0。啟動時自動檢查，選單提供「檢查更新…」與自動檢查開關。下載與安裝需要使用者點選；處理中或未匯出的召回案件會阻止直接退出。

## 建置與簽署

1. 使用 Xcode，執行 `DEVELOPER_DIR=/Applications/Xcode.app/Contents/Developer zsh scripts/package_macos.sh 版本號`。
2. 執行 `python3 scripts/create_appcast.py 版本號`，以本機鑰匙圈 account `com.returnbot.app` 簽署更新 ZIP 與 appcast。
3. GitHub Release `v版本號` 同時上傳 DMG、ZIP、appcast.xml，確認檔案完整後才發布為 latest。

來源：`https://github.com/hsiao840412/ReturnBot/releases/latest/download/appcast.xml`。每次發布都必須保留同一簽章金鑰、bundle id `com.returnbot.app` 與 `ReturnBot.app` 名稱。私密金鑰在發布者的 macOS 鑰匙圈中，不在原始碼或 App 內；更换電腦須安全移轉金鑰，不可直接產生另一把。

3.2.1 是本機試用版本，3.2.2 是第一個公開更新來源。舊版 3.2 沒有 Sparkle，需先手動安裝一次含更新功能的版本。App 須放在可寫入位置（通常為「應用程式」），不能直接在唯讀 DMG 中更新。仍依 macOS 原本的安全性與安裝權限處理，不繞過 Gatekeeper。

## 驗證

- `swift test --disable-sandbox --filter UpdateStateTests`：多視窗處理狀態、召回匯出後修改保護。
- `create_appcast.py` 自動驗證產生的 ZIP 與 feed 簽章。
- 發布後以 3.2.1 驗證能讀取正式 feed 並看到 3.2.2；安裝與重新啟動仍需實際操作測試。
