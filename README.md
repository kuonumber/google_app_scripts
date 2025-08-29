# Google Apps Script - PM2.5 空氣品質資料匯入工具

## 📋 **專案簡介**

這是一個 Google Apps Script 專案，用於從環境部環境資料開放平台自動獲取 PM2.5 等空氣品質監測資料，並匯入到 Google Sheets 中進行分析。

## 🚀 **主要功能**

- **自動匯入空氣品質資料**：從環境部 API 獲取即時 PM2.5 資料
- **智慧欄位選擇**：自動偵測有值的欄位，避免空白欄位
- **資料預覽功能**：可建立 AQI_Preview 工作表進行資料預覽
- **縣市篩選**：支援按縣市篩選空氣品質資料
- **API Key 管理**：完整的 API Key 設定、測試、清除功能

## 🔧 **安裝與設定**

### **步驟 1：建立 Google Apps Script 專案**
1. 前往 [script.google.com](https://script.google.com)
2. 建立新專案或選擇現有專案
3. 將 `pm2.5.gs` 的內容複製到編輯器中

### **步驟 2：設定 Script Properties（重要！）**
1. 在 Apps Script 編輯器中點擊左側的「專案設定」（齒輪圖示）
2. 在「腳本屬性」區段點擊「新增腳本屬性」
3. **屬性名稱**：輸入 `ENV_API_KEY`
4. **屬性值**：輸入您的環境部 API Key
   ```
   範例：af1e421d-ceb9-48a3-ab82-35bdbb1429d4
   ```
5. 點擊「儲存腳本屬性」

### **步驟 3：設定 appsscript.json**
1. 將 `appsscript.json` 的內容複製到專案設定中
2. 儲存專案

## 🔑 **API Key 取得**

1. 前往 [環境部環境資料開放平台](https://data.moenv.gov.tw/)
2. 註冊/登入帳戶
3. 申請 API Key 存取權限
4. 取得 API Key 後設定到 Script Properties

## 📊 **資料欄位說明**

目前依據 `aqx_p_02` 回傳並實際輸出的欄位（固定順序）：
1. **縣市** (`county`)
2. **測站** (`site`)
3. **PM2.5** (`pm25`，同時相容舊鍵 `pm2.5`)
4. **單位** (`itemunit`)
5. **資料建立日期** (`datacreationdate`)

> 說明：程式會固定輸出上述欄位，若欄位值缺漏則為空字串。

## 🎯 **使用方法**

### **基本操作**
1. **匯入資料**：執行 `importAirQualityData()` 或從選單選擇「匯入空氣品質資料」
2. **預覽資料**：執行 `writeApiPreviewToSheet(20)` 建立 AQI_Preview 工作表
3. **重新整理**：執行 `refreshData()` 清除舊資料並重新匯入

### **進階功能**
- **縣市篩選**：`getCountyAirQuality('臺北市')` 篩選特定縣市資料
- **API Key 測試**：`testApiKey()` 確認 API Key 是否有效
- **資料列印**：`printApiRecordsFull(100)` 將資料列印到執行記錄

### **選單操作**
在 Google Sheets 中：
1. 點擊「擴充功能」→「巨集」
2. 選擇要執行的功能
3. 授權後即可執行

## 📝 **函數說明**

### **核心函數**
- `importAirQualityData()` - 主要資料匯入函數
- `writeApiPreviewToSheet(limit)` - 建立預覽工作表
- `refreshData()` - 重新整理資料
- `getCountyAirQuality(county)` - 縣市資料篩選

### **API Key 管理**
- `setApiKey(apiKey)` - 設定 API Key
- `getApiKey()` - 取得已設定的 API Key
- `testApiKey()` - 測試 API Key 有效性
- `clearApiKey()` - 清除 API Key

### **資料列印**
- `printApiResponse(limit)` - 列印 API 回應摘要
- `printSheetRows(limit)` - 列印工作表資料
- `printApiRecordsFull(limit)` - 列印完整 API 紀錄

## ⚠️ **注意事項**

1. **API Key 安全性**：API Key 儲存在 Script Properties 中，不會暴露在程式碼裡
2. **資料量限制**：預設獲取 1000 筆資料，可透過參數調整
3. **執行權限**：首次執行需要授權 Google Apps Script 存取試算表
4. **網路連線**：需要網路連線來存取環境部 API

## 🔄 **更新與維護**

### **更新程式碼**
1. 在 Apps Script 編輯器中更新程式碼
2. 儲存專案
3. 重新載入試算表頁面

### **更新選單**
1. 修改 `appsscript.json` 中的選單項目
2. 儲存後重新載入試算表

## 📞 **技術支援**

- **執行錯誤**：檢查執行記錄中的錯誤訊息
- **API 問題**：確認 API Key 是否有效，網路連線是否正常
- **權限問題**：確認 Google 帳戶對試算表的編輯權限

## 📄 **授權**

本專案使用環境部環境資料開放平台的公開 API，請遵守相關使用條款。

---

**版本**：1.0.0  
**最後更新**：2024年12月  
**維護者**：Jimmy Kuo