/**
 * =================================================================
 * 【ファイル 1/2: 日次集計ツール (F_DailyTool.gs)】- キャッシュ上限突破FIX版
 * =================================================================
 */

function getAvailableMonths() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const urlListSheet = ss.getSheetByName("URLリスト");
    if (!urlListSheet) throw new Error("「URLリスト」シートが見つかりません。");
    
    const urlListData = urlListSheet.getDataRange().getValues();
    const shiftEntries = urlListData.filter(row => row[0] && row[0].toString().includes("確定シフト"));
    
    const monthRegex = /^\d{4}\/(0[1-9]|1[0-2])$/;
    const availableMonths = new Set();

    shiftEntries.forEach(entry => {
      if (!entry[1]) return;
      try {
        const shiftSpreadsheet = SpreadsheetApp.openByUrl(entry[1]);
        const sheets = shiftSpreadsheet.getSheets();
        
        sheets.forEach(sheet => {
          const sheetName = sheet.getName();
          
          if (monthRegex.test(sheetName) && sheet.getLastRow() > 2) {
            availableMonths.add(sheetName);
          }
          
          if (sheetName === "確定シフト") {
            const data = sheet.getDataRange().getValues();
            if (data.length < 2) return;
            
            const header = data[0];
            const dateColIdx = header.indexOf("勤務日");
            if (dateColIdx === -1) return;
            
            for (let i = 1; i < data.length; i++) {
              const dateVal = data[i][dateColIdx];
              if (dateVal instanceof Date) {
                const yyyy = dateVal.getFullYear();
                const mm = String(dateVal.getMonth() + 1).padStart(2, '0');
                availableMonths.add(`${yyyy}/${mm}`);
              }
            }
          }
        });
      } catch (e) {
        Logger.log(`スプレッドシート読込エラー(${entry[0]}): ${e.message}`);
      }
    });
    
    return Array.from(availableMonths).sort((a, b) => b.localeCompare(a));

  } catch (e) {
    Logger.log(`getAvailableMonths Error: ${e.message}`);
    return [];
  }
}

function startCalculation(monthStr) {
  const year = parseInt(monthStr.split('/')[0], 10);
  const month = parseInt(monthStr.split('/')[1], 10) - 1;
  const startDate = new Date(year, month, 1);
  const endDate = new Date(year, month + 1, 0);
  const totalDays = (endDate.getTime() - startDate.getTime()) / (1000 * 60 * 60 * 24) + 1;
  const settings = _loadInitialData(SpreadsheetApp.getActiveSpreadsheet());
  if (!settings) throw new Error("マスターデータの読み込みに失敗しました。");

  const lastYearMonthStr = `${year - 1}/${String(month + 1).padStart(2, '0')}`;
  let allDailyData_LastYear = {};
  try {
    allDailyData_LastYear = _calculateLastYearActuals_FOR_DAILY_TOOL(lastYearMonthStr, settings);
  } catch (e) {
    Logger.log(`昨年度 (${lastYearMonthStr}) の実績データの取得に失敗: ${e.message}`);
  }

  let monthlyCache = null;
  if (typeof NewDailyCostCalculator !== 'undefined' && NewDailyCostCalculator.buildMonthlyCache) {
    monthlyCache = NewDailyCostCalculator.buildMonthlyCache(startDate, endDate);
  }

  const lightSettings = {};
  const heavySettings = {};
  for (const key in settings) {
    const value = settings[key];
    if (key === 'salesUrlMaps' && value) {
      const salesUrlMapsAsArrays = {};
      for (const yearKey in value) {
        if (value[yearKey] instanceof Map) {
          salesUrlMapsAsArrays[yearKey] = Array.from(value[yearKey].entries());
        }
      }
      lightSettings[key] = salesUrlMapsAsArrays;
    } else if (value instanceof Map) {
      heavySettings[key] = Array.from(value.entries());
    } else if (value instanceof Set) {
      lightSettings[key] = Array.from(value);
    } else {
      lightSettings[key] = value;
    }
  }

  const cache = CacheService.getScriptCache();
  cache.put('calculation_settings_light', JSON.stringify(lightSettings), 3600);
  cache.put('calculation_settings_heavy', JSON.stringify(heavySettings), 3600);
  
  const sessionInfo = { startDate: startDate.toISOString(), totalDays: totalDays };
  cache.put('calculation_session_info', JSON.stringify(sessionInfo), 3600);
  
  // ★ 容量オーバー対策: チャンク分割して保存
  _putLargeCache(cache, 'last_year_data_cache', JSON.stringify(allDailyData_LastYear), 3600);
  
  if (monthlyCache) {
    const serializableCache = {
      shifts: monthlyCache.shifts,
      ftDoctors: Array.from(monthlyCache.ftDoctors),
      ptDoctors: Array.from(monthlyCache.ptDoctors)
    };
    _putLargeCache(cache, 'new_monthly_cache', JSON.stringify(serializableCache), 3600);
  }
  
  return { totalDays: totalDays };
}

function processOneDay(dayIndex) {
  try {
    const cache = CacheService.getScriptCache();
    const sessionInfo = JSON.parse(cache.get('calculation_session_info'));
    const targetDate = new Date(sessionInfo.startDate);
    targetDate.setDate(targetDate.getDate() + dayIndex);
    
    const today = new Date();
    today.setHours(0,0,0,0);
    if (targetDate >= today) {
      return `${Utilities.formatDate(targetDate, "JST", "yyyy/M/d")} は未来の日付のため、後で予測値を計算します。`;
    }
    
    const settings = restoreSettingsFromCache('calculation_');

    // ★ 容量オーバー対策: チャンクから復元
    let monthlyCache = null;
    const monthlyCacheJson = _getLargeCache(cache, 'new_monthly_cache');
    if (monthlyCacheJson) {
      const parsedCache = JSON.parse(monthlyCacheJson);
      monthlyCache = {
        shifts: parsedCache.shifts,
        ftDoctors: new Set(parsedCache.ftDoctors),
        ptDoctors: new Set(parsedCache.ptDoctors)
      };
    }

    const dateStr = `${targetDate.getFullYear()}/${targetDate.getMonth() + 1}/${targetDate.getDate()}`;
    const dailyDataForOneDay = {};
    dailyDataForOneDay[dateStr] = {};
    
    if (typeof NewDailyCostCalculator !== 'undefined' && NewDailyCostCalculator.routeAndCalculateDailyCost) {
      NewDailyCostCalculator.routeAndCalculateDailyCost(targetDate, dailyDataForOneDay, settings, monthlyCache);
    } else {
      _calculateDailyPersonnelCosts(targetDate, dailyDataForOneDay, settings);
    }
    
    cache.put(`day_result_${dayIndex}`, JSON.stringify(dailyDataForOneDay), 3600);
    
    return `${Utilities.formatDate(targetDate, "JST", "yyyy/M/d")}\n - 人件費 計算完了`;

  } catch (e) {
    const date = new Date();
    const errorMessage = `エラー: ${e.message || '不明なエラー'}`;
    Logger.log(`processOneDayでエラー発生: ${errorMessage}\n${e.stack || ''}`);
    return `${Utilities.formatDate(date, "JST", "yyyy/M/d")}\n ${errorMessage}`;
  }
}

function finishCalculation() {
  const cache = CacheService.getScriptCache();
  const sessionInfoJson = cache.get('calculation_session_info');
  if (!sessionInfoJson) throw new Error("セッション情報が見つかりません。");
  const sessionInfo = JSON.parse(sessionInfoJson);

  const allDailyData = {};
  const cacheKeysToRemove = ['calculation_settings_light', 'calculation_settings_heavy', 'calculation_session_info'];
  for (let i = 0; i < sessionInfo.totalDays; i++) {
    const dayKey = `day_result_${i}`;
    const dayResultJson = cache.get(dayKey);
    if (dayResultJson) {
      Object.assign(allDailyData, JSON.parse(dayResultJson));
    }
    cacheKeysToRemove.push(dayKey);
  }

  const fullSettings = restoreSettingsFromCache('calculation_');
  
  // ★ 容量オーバー対策: チャンクから復元
  const lastYearCacheJson = _getLargeCache(cache, 'last_year_data_cache');
  if (lastYearCacheJson) {
    fullSettings.lastYearDailyData = JSON.parse(lastYearCacheJson);
  }
    
  const startDate = new Date(sessionInfo.startDate);
  const endDate = new Date(startDate);
  endDate.setDate(startDate.getDate() + sessionInfo.totalDays - 1);

  const futureCosts = _calculateFutureCostsForDailyTool(startDate, endDate, fullSettings);
  
  for (const dateKey in futureCosts) {
      if (!allDailyData[dateKey]) allDailyData[dateKey] = {};
      for (const clinicKey in futureCosts[dateKey]) {
          allDailyData[dateKey][clinicKey] = futureCosts[dateKey][clinicKey];
      }
  }

  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  const OUTPUT_SHEET = SPREADSHEET.getSheetByName("拠点別コスト");
  if (!OUTPUT_SHEET) throw new Error("「拠点別コスト」シートが見つかりません。");
  
  _writeResultsToSheet(allDailyData, startDate, endDate, OUTPUT_SHEET, fullSettings, fullSettings.holidays, sessionInfo.totalDays);
  
  cache.removeAll(cacheKeysToRemove);
  _removeLargeCache(cache, 'last_year_data_cache');
  _removeLargeCache(cache, 'new_monthly_cache');
  
  return "シートへの書き込みが完了しました。";
}

function cancelCalculation(processedDays) {
  const cache = CacheService.getScriptCache();
  const cacheKeysToRemove = ['calculation_settings_light', 'calculation_settings_heavy', 'calculation_session_info'];
  for (let i = 0; i < processedDays; i++) {
    cacheKeysToRemove.push(`day_result_${i}`);
  }
  cache.removeAll(cacheKeysToRemove);
  _removeLargeCache(cache, 'last_year_data_cache');
  _removeLargeCache(cache, 'new_monthly_cache');
}

// ====================================================================
// ▼ 100KB制限回避のためのチャンク処理ヘルパー関数 ▼
// ====================================================================

function _putLargeCache(cache, key, valueStr, time) {
  const MAX_SIZE = 90000; // 90KB安全マージン
  if (valueStr.length > MAX_SIZE) {
    const chunks = Math.ceil(valueStr.length / MAX_SIZE);
    cache.put(key + '_chunks', chunks.toString(), time);
    for (let i = 0; i < chunks; i++) {
      cache.put(key + '_' + i, valueStr.substring(i * MAX_SIZE, (i + 1) * MAX_SIZE), time);
    }
  } else {
    cache.put(key, valueStr, time);
  }
}

function _getLargeCache(cache, key) {
  const chunkCountStr = cache.get(key + '_chunks');
  if (chunkCountStr) {
    const chunks = parseInt(chunkCountStr, 10);
    let result = '';
    for (let i = 0; i < chunks; i++) {
      const chunk = cache.get(key + '_' + i);
      if (!chunk) return null; 
      result += chunk;
    }
    return result;
  } else {
    return cache.get(key);
  }
}

function _removeLargeCache(cache, key) {
  const chunkCountStr = cache.get(key + '_chunks');
  if (chunkCountStr) {
    const chunks = parseInt(chunkCountStr, 10);
    for (let i = 0; i < chunks; i++) {
      cache.remove(key + '_' + i);
    }
    cache.remove(key + '_chunks');
  }
  cache.remove(key);
}