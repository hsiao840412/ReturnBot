# ReturnBot 退料機器人

用於製作退料與寄銷召回文件的 macOS App。匯入清單、核對零件資料，即可產生 Excel 與 DHL CSV，減少重複填寫。

[下載最新版本](https://github.com/hsiao840412/ReturnBot/releases/latest) · [更新歷史](packaging/RELEASE_NOTES.md)

## 系統需求

- Apple Silicon Mac（arm64）
- macOS 26 Tahoe 或更新版本
- 生成文件不需安裝或啟動 Excel；檢視、編輯及列印成品可使用 Excel。

## 安裝

1. 前往 [最新版本](https://github.com/hsiao840412/ReturnBot/releases/latest) 下載 DMG。
2. 開啟 DMG，將 `ReturnBot.app` 拖入「應用程式」，已有同名 App 時選擇取代。
3. 第一次啟動若 macOS 阻擋 App，請依「系統設定 → 隱私權與安全性」提示允許開啟。
4. 若 macOS 詢問輸出資料夾存取權，允許 ReturnBot 儲存文件即可，不需 Excel 自動化授權。
5. 使用鋰電池條碼時，請先下載並安裝 [ConnectCode39 字體](https://github.com/hsiao840412/ReturnBot/releases/latest/download/ConnectCode39.ttf)。

安裝後可使用 App 的「檢查更新…」取得新版並查看更新歷史。

## 一般退料

支援「Mail-in KBB」「Mail-in 電池膨脹」「一般 KBB」與「單獨鋰電池 KBB」。

1. 依本次作業選擇退料方式；不確定分類時，請先向資深夥伴確認。
2. 拖入或選擇從 ePacking List 下載的 CSV 清單，保留原始欄位與內容。
3. 按下「製作 Excel 文件」。
4. 完成後按「開啟 Excel，核對內容」，檢查料號、數量與退料資料。

文件儲存在 Mac 的「下載項目」，包含 Invoice、ePacking List、標籤及適用類型的 DHL CSV。更換清單或退料方式後，請重新製作文件。

## 寄銷召回 Beta

1. **準備資料**：匯入最新的零件價格表（Excel 或 CSV），輸入召回單號與本次匯率，再貼上截圖、選擇截圖或手動輸入零件資料。
2. **核對零件明細**：核對料號、數量、單價、重量、原產地及商品描述。
3. **確認分單與儲存**：查看分單結果、勾選核對項目，再選擇資料夾儲存。

系統依每單 USD 7,500 上限分單，每單包含 KGB Invoice、外箱標籤與 DHL CSV。預設每單一箱，毛重與尺寸請裝箱後填寫。

價格表只保存在本機，不隨 App 發布。

文件製作完成表示檔案已建立，內容仍需人工核對。

## 專案維護

[目錄與本機資源說明](docs/development.md) · [打包與更新發佈](docs/updates.md)
