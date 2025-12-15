# 🤖 ReturnBot - 退料機器人 v2.0

一個基於 Python 和 `tkinter` 的桌面應用程式，用於自動化生成和填寫 Apple 維修退料流程所需的 Excel (KBB/Mail-in) 文件，大幅提高處理效率。

## ✨ 核心功能

* **多類型支援：** 支援四種主要的退料類型：
    * Mail-in KBB
    * Mail-in 電池膨脹
    * 一般 KBB
    * 單獨鋰電池 KBB
* **一鍵生成：** 匯入 ePacking List CSV 檔案後，一鍵生成包含所有必要數據的 Excel 模板檔案。
* **自動化數據處理：**
    * **發票單號自動命名：** 根據退料類型自動產生日期和類型化的發票編號 (e.g., `800935_YYYYMMDD`, `SRR#YYYY-MMT935(KBB)` )。
    * **Excel 模板填充：** 自動將 CSV 數據寫入 Excel 模板中的「KBB&KGB invoice」與「ePacking List」工作表。
    * **動態調整：** 根據 CSV 數據筆數，自動在 Excel 發票工作表中插入/刪除列，確保格式正確。
* **DHL 上傳檔生成：** 針對 Mail-in KBB 和一般 KBB 類型，自動生成 DHL 貨物上傳所需的 CSV 檔案，包含國家代碼和預估重量。
* **條碼支援：** 針對「單獨鋰電池 KBB」，自動將數據填入「條碼」工作表，並複製公式以產生所需條碼。

##  安裝與使用 

1.  前往 [Releases](https://github.com/hsiao840412/ReturnBot/releases) 頁面下載最新版本的 `ReturnBot.dmg`。
2.  開啟後將 App 拖入「應用程式」資料夾。
3.  打開終端機複製 “xattr -cr `"App拉進去"` 然後 Return。
4.  視需求安裝條碼字體
