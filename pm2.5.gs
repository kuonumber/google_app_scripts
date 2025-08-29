/**
 * 從環境部環境資料開放平台獲取空氣品質監測資料
 * 資料來源：https://data.moenv.gov.tw/swagger/#/%E5%A4%A7%E6%B0%A3/get_aqx_p_02
 * 使用環境部 v2 API 以 JSON 方式取得 PM2.5 等空氣品質資料
 */
function importAirQualityData() {
  try {
    var apiKey = getApiKey();
    if (!apiKey) {
      throw new Error('尚未設定 ENV_API_KEY（腳本屬性）。請先執行 setApiKey() 或在專案設定中新增。');
    }
    importAirQualityDataWithApiKey(apiKey);
  } catch (error) {
    Logger.log('匯入資料時發生錯誤: ' + error.toString());
    throw error;
  }
}

/**
 * 使用 API Key 從環境部 API 獲取資料
 * @param {string} apiKey - 環境部 API Key（存於 Script Properties 的 ENV_API_KEY）
 */
function importAirQualityDataWithApiKey(apiKey) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();

    // 以查詢參數傳遞 api_key，回傳 JSON 並限制筆數
    var apiUrl = 'https://data.moenv.gov.tw/api/v2/aqx_p_02?api_key=' + encodeURIComponent(apiKey) + '&format=json&limit=1000';
    var options = {
      method: 'GET',
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(apiUrl, options);
    var text = response.getContentText();

    var data;
    try {
      data = JSON.parse(text);
    } catch (parseErr) {
      Logger.log('API 回應非 JSON：' + text);
      throw parseErr;
    }

    if (data && data.records && Array.isArray(data.records)) {
      var records = data.records;

      // 候選欄位與顯示名稱（僅輸出 aqx_p_02 record 會回傳且有值的欄位；PM2.5 一定輸出）
      var candidates = [
        { key: 'county', label: '縣市' },
        { key: 'site', label: '測站' },
        { key: 'pm25', label: 'PM2.5' },
        { key: 'itemunit', label: '單位' },
        { key: 'datacreationdate', label: '資料建立日期' }
      ];

      var included = candidates.filter(function(c) {
        if (c.key === 'pm25') return true; // 一定包含 PM2.5
        return records.some(function(r) {
          var v = r[c.key];
          if (v == null || v === '') return false;
          return true;
        });
      });

      // 若資料為舊鍵名 'pm2.5'，也視為有值以保留 PM2.5 欄位
      if (!included.some(function(c) { return c.key === 'pm25'; })) {
        var hasOldPm = records.some(function(r) { return r['pm2.5'] != null && r['pm2.5'] !== ''; });
        if (hasOldPm) included.unshift({ key: 'pm25', label: 'PM2.5' });
      }

      var headers = included.map(function(c) { return c.label; });
      var rows = records.map(function(r) {
        return included.map(function(c) {
          if (c.key === 'pm25') return r.pm25 || r['pm2.5'] || '';
          return r[c.key] || '';
        });
      });

      sheet.clear();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
      }
      sheet.autoResizeColumns(1, headers.length);
      Logger.log('使用 API Key 成功獲取空氣品質資料，共 ' + rows.length + ' 筆資料；欄位：' + headers.join(','));
    } else {
      throw new Error('API 回應格式不正確或無資料: ' + JSON.stringify(data));
    }
  } catch (error) {
    Logger.log('使用 API Key 獲取資料失敗: ' + error.toString());
    throw error;
  }
}

/**
 * 設定 API Key
 * @param {string} apiKey - 環境部 API Key
 */
function setApiKey(apiKey) {
  try {
    PropertiesService.getScriptProperties().setProperty('ENV_API_KEY', apiKey);
    Logger.log('API Key 設定成功');
    return true;
  } catch (error) {
    Logger.log('設定 API Key 失敗: ' + error.toString());
    return false;
  }
}

/**
 * 取得已設定的 API Key
 * @returns {string|null} API Key 或 null
 */
function getApiKey() {
  try {
    return PropertiesService.getScriptProperties().getProperty('ENV_API_KEY');
  } catch (error) {
    Logger.log('獲取 API Key 失敗: ' + error.toString());
    return null;
  }
}

/**
 * 清除已設定的 API Key
 */
function clearApiKey() {
  try {
    PropertiesService.getScriptProperties().deleteProperty('ENV_API_KEY');
    Logger.log('API Key 已清除');
    return true;
  } catch (error) {
    Logger.log('清除 API Key 失敗: ' + error.toString());
    return false;
  }
}

/**
 * 檢查 API Key 是否有效（以 query 參數方式呼叫）
 * @returns {boolean} API Key 是否有效
 */
function testApiKey() {
  try {
    var apiKey = getApiKey();
    if (!apiKey) {
      Logger.log('未設定 API Key');
      return false;
    }

    var apiUrl = 'https://data.moenv.gov.tw/api/v2/aqx_p_02?api_key=' + encodeURIComponent(apiKey) + '&format=json&limit=1';
    var response = UrlFetchApp.fetch(apiUrl, { method: 'GET', muteHttpExceptions: true });
    var text = response.getContentText();
    var data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      Logger.log('測試回應非 JSON：' + text);
      return false;
    }

    if (data && data.records && Array.isArray(data.records)) {
      Logger.log('API Key 有效，可以獲取 ' + data.records.length + ' 筆（測試）');
      return true;
    }
    Logger.log('API Key 無效或 API 回應異常：' + JSON.stringify(data));
    return false;
  } catch (error) {
    Logger.log('測試 API Key 失敗: ' + error.toString());
    return false;
  }
}

/**
 * 重新整理試算表中的資料
 * 清除所有資料並重新匯入最新資料
 */
function refreshData() {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    sheet.clear();
    importAirQualityData();
    Logger.log('資料重新整理完成');
  } catch (error) {
    Logger.log('重新整理資料時發生錯誤: ' + error.toString());
    throw error;
  }
}

/**
 * 獲取特定縣市的空氣品質資料
 * @param {string} county - 縣市名稱，例如：'臺北市', '新北市'
 */
function getCountyAirQuality(county) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var countyIndex = headers.indexOf('縣市');
    if (countyIndex === -1) {
      throw new Error('找不到縣市欄位');
    }

    var filteredData = data.filter(function(row, index) {
      return index === 0 || row[countyIndex] === county; // 保留表頭
    });

    sheet.clear();
    sheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
    Logger.log('已篩選 ' + county + ' 的空氣品質資料，共 ' + (filteredData.length - 1) + ' 筆');
  } catch (error) {
    Logger.log('篩選縣市資料時發生錯誤: ' + error.toString());
    throw error;
  }
}

/**
 * 直接從 API 列印前幾筆紀錄到執行記錄（Logger）
 * @param {number} limit 要列印的筆數，預設 5
 */
function printApiResponse(limit) {
  try {
    var apiKey = getApiKey();
    if (!apiKey) {
      Logger.log('未設定 API Key');
      return;
    }
    var count = (typeof limit === 'number' && limit > 0) ? limit : 5;
    var apiUrl = 'https://data.moenv.gov.tw/api/v2/aqx_p_02?api_key=' + encodeURIComponent(apiKey) + '&format=json&limit=' + count;
    var res = UrlFetchApp.fetch(apiUrl, { method: 'GET', muteHttpExceptions: true });
    var text = res.getContentText();
    var data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      Logger.log('API 回應非 JSON：' + text);
      return;
    }
    if (!data || !Array.isArray(data.records)) {
      Logger.log('API 回應無 records：' + JSON.stringify(data));
      return;
    }
    Logger.log('將列印前 ' + data.records.length + ' 筆：');
    data.records.forEach(function(r, i) {
      var brief = {
        index: i + 1,
        sitename: r.sitename || '',
        county: r.county || '',
        pm25: r.pm25 || r['pm2.5'] || '',
        publishtime: r.publishtime || ''
      };
      Logger.log(JSON.stringify(brief));
    });
  } catch (err) {
    Logger.log('printApiResponse 發生錯誤：' + err);
  }
}

/**
 * 從目前工作表列印前幾筆資料到執行記錄（Logger）
 * @param {number} limit 要列印的筆數，預設 5（不含表頭）
 */
function printSheetRows(limit) {
  try {
    var sheet = SpreadsheetApp.getActiveSheet();
    var values = sheet.getDataRange().getValues();
    if (!values || values.length <= 1) {
      Logger.log('工作表尚無資料可列印');
      return;
    }
    var headers = values[0];
    var max = Math.min((typeof limit === 'number' && limit > 0) ? limit : 5, values.length - 1);
    Logger.log('將列印表頭與前 ' + max + ' 筆資料：');
    Logger.log('Headers: ' + JSON.stringify(headers));
    for (var i = 1; i <= max; i++) {
      Logger.log('Row ' + i + ': ' + JSON.stringify(values[i]));
    }
  } catch (err) {
    Logger.log('printSheetRows 發生錯誤：' + err);
  }
}

/**
 * 將 API 回應的前幾筆資料寫入名為 AQI_Preview 的工作表，並顯示提示
 * @param {number} limit 要寫入的筆數，預設 10
 */
function writeApiPreviewToSheet(limit) {
  try {
    var apiKey = getApiKey();
    if (!apiKey) {
      SpreadsheetApp.getActive().toast('未設定 API Key', 'AQI_Preview', 5);
      return;
    }
    var count = (typeof limit === 'number' && limit > 0) ? limit : 10;
    var apiUrl = 'https://data.moenv.gov.tw/api/v2/aqx_p_02?api_key=' + encodeURIComponent(apiKey) + '&format=json&limit=' + count;
    var res = UrlFetchApp.fetch(apiUrl, { method: 'GET', muteHttpExceptions: true });
    var text = res.getContentText();
    var data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      SpreadsheetApp.getActive().toast('API 回應非 JSON，請查看執行記錄', 'AQI_Preview', 5);
      Logger.log('API 回應非 JSON：' + text);
      return;
    }
    if (!data || !Array.isArray(data.records)) {
      SpreadsheetApp.getActive().toast('API 無 records，請查看執行記錄', 'AQI_Preview', 5);
      Logger.log('API 回應無 records：' + JSON.stringify(data));
      return;
    }

    var records = data.records;
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('AQI_Preview') || ss.insertSheet('AQI_Preview');
    sheet.clear();

    // 候選欄位與顯示名稱（僅輸出 aqx_p_02 record 會回傳且有值的欄位；PM2.5 一定輸出）
    var candidates = [
      { key: 'county', label: '縣市' },
      { key: 'site', label: '測站' },
      { key: 'pm25', label: 'PM2.5' },
      { key: 'itemunit', label: '單位' },
      { key: 'datacreationdate', label: '資料建立日期' }
    ];

    var included = candidates.filter(function(c) {
      if (c.key === 'pm25') return true;
      return records.some(function(r) {
        var v = r[c.key];
        return !(v == null || v === '');
      });
    });
    if (!included.some(function(c) { return c.key === 'pm25'; })) {
      var hasOldPm = records.some(function(r) { return r['pm2.5'] != null && r['pm2.5'] !== ''; });
      if (hasOldPm) included.unshift({ key: 'pm25', label: 'PM2.5' });
    }

    var headers = included.map(function(c) { return c.label; });
    var rows = records.map(function(r) {
      return included.map(function(c) {
        if (c.key === 'pm25') return r.pm25 || r['pm2.5'] || '';
        return r[c.key] || '';
      });
    });

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    sheet.autoResizeColumns(1, headers.length);

    SpreadsheetApp.getActive().toast('已寫入 AQI_Preview 前 ' + rows.length + ' 筆資料；欄位：' + headers.join(','), 'AQI_Preview', 5);
  } catch (err) {
    Logger.log('writeApiPreviewToSheet 發生錯誤：' + err);
    SpreadsheetApp.getActive().toast('寫入預覽時發生錯誤，請查看記錄', 'AQI_Preview', 5);
  }
}

/**
 * 將 API 回應的紀錄完整以 JSON 逐筆列印到執行記錄（Logger）
 * 注意：列印 1000 筆會產生大量日誌，請在需要時使用
 * @param {number} limit 要列印的筆數，預設 1000
 */
function printApiRecordsFull(limit) {
  try {
    var apiKey = getApiKey();
    if (!apiKey) {
      Logger.log('未設定 API Key');
      return;
    }
    var count = (typeof limit === 'number' && limit > 0) ? limit : 1000;
    var apiUrl = 'https://data.moenv.gov.tw/api/v2/aqx_p_02?api_key=' + encodeURIComponent(apiKey) + '&format=json&limit=' + count;
    var res = UrlFetchApp.fetch(apiUrl, { method: 'GET', muteHttpExceptions: true });
    var text = res.getContentText();
    var data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      Logger.log('API 回應非 JSON：' + text);
      return;
    }
    if (!data || !Array.isArray(data.records)) {
      Logger.log('API 回應無 records：' + JSON.stringify(data));
      return;
    }
    Logger.log('即將列印 ' + data.records.length + ' 筆完整紀錄');
    data.records.forEach(function(rec, idx) {
      Logger.log('#' + (idx + 1) + ' ' + JSON.stringify(rec));
    });
  } catch (err) {
    Logger.log('printApiRecordsFull 發生錯誤：' + err);
  }
}

// 保留舊函數名稱以維持向後相容性
function myFunction() {
  importAirQualityData();
}

function refreshmacro() {
  refreshData();
}
