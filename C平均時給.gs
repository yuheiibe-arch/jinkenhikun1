/**
 * =================================================================
 * 【ファイル X/X: 平均時給計算エンジン (G_AvgWageTool.gs)】
 * =================================================================
 */
function calculateAverageWages_v2() {
  const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
  const outputSheet = SPREADSHEET.getSheetByName("平均時給");
  if (!outputSheet) {
    SpreadsheetApp.getUi().alert("「平均時給」シートが見つかりません。");
    return;
  }
  const layoutData = outputSheet.getRange(1, 1, outputSheet.getLastRow(), 4).getValues();
  if (layoutData.length < 2) {
    SpreadsheetApp.getUi().alert("「平均時給」シートにレイアウト行がありません。");
    return;
  }
  Logger.log("共有の設定ファイルをロードします...");
  SpreadsheetApp.flush(); 
  let settings;
  try {
    settings = _loadInitialData(SPREADSHEET); 
  } catch (e) {
    SpreadsheetApp.getUi().alert(`設定の読み込み中にエラー:\n${e.message}`);
    return;
  }
  Logger.log("全期間のデータを事前計算します...");
  const monthlyDataCache = new Map(); 
  const today = new Date();
  const endMonthStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}`; 
  const startMonthStr_2025 = `2025/04`;
  const uniqueMonths = new Set([
    ...gt_getMonthsInRange("2024/04", "2025/03"), 
    ...gt_getMonthsInRange(startMonthStr_2025, endMonthStr) 
  ]);
  const monthsToCalc = Array.from(uniqueMonths);
  for (const monthStr of monthsToCalc) {
    if (monthStr > endMonthStr && !(new Date(monthStr + '/01') < today) ) continue;
    try {
      const { clinicDataMap } = _calculateActualsForAllClinics_OneMonth(monthStr, settings); 
      monthlyDataCache.set(monthStr, clinicDataMap);
      Logger.log(`  - ${monthStr} の計算完了`);
    } catch (e) {
      Logger.log(`  - ${monthStr} の計算中にエラー: ${e.message}`);
    }
  }
  Logger.log("データ事前計算完了。");
  const outputMatrix = [];
  const monthOrder = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3]; 
  for (let i = 1; i < layoutData.length; i++) { 
    const row = layoutData[i];
    const lookupKey = row[0]; 
    const yearLabel = row[3]; 
    if (!lookupKey || !yearLabel) {
      outputMatrix.push(Array(15).fill(null)); 
      continue;
    }
    const yearPrefix = (yearLabel.includes("昨年度")) ? 2024 : 2025;
    const monthlyAveragesMap = new Map(); 
    for (const monthNum of monthOrder) {
      const year = (monthNum <= 3) ? (yearPrefix + 1) : yearPrefix; 
      const monthStr = `${year}/${String(monthNum).padStart(2, '0')}`;
      const monthDataMap = monthlyDataCache.get(monthStr); 
      let avgWage = null;
      if (monthDataMap && monthDataMap.size > 0) {
        if (lookupKey === "関東" || lookupKey === "関西") {
          const regionalAverages = _calculateAveragesFromActualsMap(monthDataMap, settings.clinicGroupMap);
          avgWage = regionalAverages[lookupKey] || null; 
        } else {
          const normalizedClinic = gt_normalizeClinicName(lookupKey, settings.clinicAliasMap);
          const clinicData = monthDataMap.get(normalizedClinic);
          avgWage = (clinicData && clinicData.hours > 0) ? (clinicData.cost / clinicData.hours) : null;
        }
      }
      monthlyAveragesMap.set(monthNum, avgWage);
    }
    const generalValues = [monthlyAveragesMap.get(4), monthlyAveragesMap.get(5), monthlyAveragesMap.get(6), monthlyAveragesMap.get(7), monthlyAveragesMap.get(8), monthlyAveragesMap.get(9)].filter(v => v !== null && v > 0);
    const regularValues = [monthlyAveragesMap.get(10), monthlyAveragesMap.get(11), monthlyAveragesMap.get(12), monthlyAveragesMap.get(1), monthlyAveragesMap.get(2), monthlyAveragesMap.get(3)].filter(v => v !== null && v > 0);
    const allValues = [...generalValues, ...regularValues];
    const avgQ_Annual = allValues.length > 0 ? allValues.reduce((a, b) => a + b, 0) / allValues.length : null; 
    const avgR_General = generalValues.length > 0 ? generalValues.reduce((a, b) => a + b, 0) / generalValues.length : null; 
    const avgS_Regular = regularValues.length > 0 ? regularValues.reduce((a, b) => a + b, 0) / regularValues.length : null; 
    const monthlyValues = monthOrder.map(m => monthlyAveragesMap.get(m) || null);
    outputMatrix.push([...monthlyValues, avgQ_Annual, avgR_General, avgS_Regular]); 
  }
  if (outputMatrix.length > 0) {
    const targetRange = outputSheet.getRange(2, 5, outputMatrix.length, 15); 
    targetRange.clearContent(); 
    targetRange.setValues(outputMatrix);
    targetRange.setNumberFormat('"¥"#,##0'); 
    Logger.log("シートへの書き込みが完了しました。");
    SpreadsheetApp.getUi().alert('平均時給シートの更新が完了しました。');
  } else {
    Logger.log("書き込むデータがありませんでした。");
  }
}