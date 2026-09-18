# 退料機器人 3.2

## 系統需求

- Apple Silicon Mac（arm64）
- macOS 26 Tahoe 或更新版本
- 生成文件不需安裝或啟動 Excel；檢視、編輯及列印成品可使用 Excel。

## 安裝

1. 前往 [3.2 Release](https://github.com/hsiao840412/ReturnBot/releases/tag/v3.2) 下載 `ReturnBot-3.2-arm64.dmg`。
2. 開啟 DMG，將 `RetuenBot.app` 拖入「應用程式」，已有同名 App 時選擇取代。
3. 第一次啟動若 macOS 阻擋 App，請依「系統設定 → 隱私權與安全性」提示允許開啟。
4. 若 macOS 詢問輸出資料夾存取權，允許 ReturnBot 儲存文件即可，不需 Excel 自動化授權。
5. 鋰電池條碼列印仍需安裝原作業使用的條碼字體。

App 名稱為 **RetuenBot**，開啟後的視窗標題為「退料機器人」，版本顯示 **3.2**。

> 此版本採用 ad-hoc 簽署，未經 Apple Developer ID 公證，因此首次開啟可能出現安全提示。

## 一般退料

1. 選擇退料類型。
2. 選擇 ePacking List CSV。
3. 按下「生成 Excel 退料文件」。
4. 成品儲存在「下載項目」，包含 Invoice、ePacking List、標籤及適用類型的 DHL CSV。

3.2 改由 App 直接產生 Excel 文件，不再啟動 Excel 或執行模板授權預檢。

## 寄銷召回 Beta

1. 匯入自己的台幣價格表（Excel 或 CSV），輸入召回單號與本次匯率。
2. 貼上截圖、選擇圖片或手動輸入料號與數量，再加入明細。OCR 預設快速模式，也可選精確模式。
3. 核對料號、數量、交換價格、商品描述、重量與原產地。
4. 預覽分單，核對後選擇資料夾匯出。

美金單價四捨五入至整數，每單上限 USD 5,000。每單輸出 Invoice、外箱標籤與 DHL CSV，不含 ePacking List。商品描述保留原文，0 美金品項需逐筆確認。請於裝箱後補上實際毛重及尺寸。

價格表只保存在本機，不隨 App 發布。尚未提供未匯出草稿自動保存；關閉 App 前請先匯出。
