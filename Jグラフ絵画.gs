/**
 * =================================================================
 * 【ファイル 6/X: グラフ描画 (E_Graph_Writer.gs)】- 修正版
 * =================================================================
 */

function gt_writeToMonthlySummaryAndCreateChart(spreadsheet, clinic, data, chartsToCreate, settings) {
  const sheetName = '月間サマリー';
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) { sheet.clear(); } else { sheet = spreadsheet.insertSheet(sheetName); }
  sheet.getRange('A1').setValue(`【${clinic}】月間サマリー`).setFontWeight('bold').setFontSize(14);
  const normalizedClinicName = gt_normalizeClinicName(clinic, settings.clinicAliasMap);
  const groupName = settings.clinicGroupMap.get(normalizedClinicName) || "エリア外";
  const regionName = (groupName === "大阪") ? "関西" : "関東";
  
  const items = [ 
      '予算', '売上', 
      '常勤医師給与', '定期非常勤給与', 'スポット医師給与', '直接応募医師給与', '紹介会社医師給与', '所定休出医師給与', '紹介手数料', '特別手当', 
      '人件費率', '残額', 
      '平均時給 (当拠点)', '昨年度平均時給 (当拠点)',
      `平均時給 (${groupName})`, `平均時給 (${regionName})`, 
      '常勤医師数', '定期非常勤医師数', '直接応募医師数', '紹介会社医師数'
  ];
  const countItems = ['常勤医師数', '定期非常勤医師数', '直接応募医師数', '紹介会社医師数'];
  const avgItems = ['平均時給 (当拠点)', '昨年度平均時給 (当拠点)', `平均時給 (${groupName})`, `平均時給 (${regionName})`];
  const moneyItems = ['予算', '売上', '常勤医師給与', '定期非常勤給与', 'スポット医師給与', '直接応募医師給与', '紹介会社医師給与', '所定休出医師給与', '紹介手数料', '特別手当', '残額'];
  
  const today = new Date(); today.setHours(0,0,0,0);
  const currentMonthStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}`;
  const monthlyHeaders = data.map(d => d.month);
  let firstForecastIndex = monthlyHeaders.findIndex(h => h >= currentMonthStr); 
  if (firstForecastIndex === -1) firstForecastIndex = data.length;
  
  const lastYearAvgWages = new Map();
  for (const monthStr of monthlyHeaders) {
    const [year, month] = monthStr.split('/').map(Number);
    const lastYearMonthStr = `${year - 1}/${String(month).padStart(2, '0')}`;
    try {
      const { clinicDataMap } = _calculateActualsForAllClinics_OneMonth(lastYearMonthStr, settings);
      const lyClinicData = clinicDataMap.get(normalizedClinicName);
      const lyAvgWage = (lyClinicData && lyClinicData.hours > 0) ? (lyClinicData.cost / lyClinicData.hours) : null;
      lastYearAvgWages.set(monthStr, lyAvgWage);
    } catch(e) {
      Logger.log(`昨年度(${lastYearMonthStr})のデータ取得に失敗: ${e.message}`);
      lastYearAvgWages.set(monthStr, null);
    }
  }

  const actualsDataOnly = data.slice(0, firstForecastIndex);
  const actualMonthlyAvgs = actualsDataOnly.map(d => (d.hours > 0) ? (d.cost / d.hours) : null).filter(avg => avg !== null);
  let actualsSimpleAverage = null;
  if (actualMonthlyAvgs.length > 0) {
    actualsSimpleAverage = actualMonthlyAvgs.reduce((a, b) => a + b, 0) / actualMonthlyAvgs.length;
  }

  const lyActualsDataOnly = Array.from(lastYearAvgWages.values()).slice(0, firstForecastIndex).filter(v => v !== null);
  let lyActualsSimpleAverage = null;
  if (lyActualsDataOnly.length > 0) {
      lyActualsSimpleAverage = lyActualsDataOnly.reduce((a, b) => a + b, 0) / lyActualsDataOnly.length;
  }

  const actualsHeader = firstForecastIndex > 0 ? `実績合計\n(${monthlyHeaders[0]} ~ ${monthlyHeaders[firstForecastIndex - 1]})` : '実績合計';
  const forecastHeader = firstForecastIndex < data.length ? `予測合計\n(${monthlyHeaders[firstForecastIndex]} ~ ${monthlyHeaders[data.length - 1]})` : '予測合計';
  const headers = ['項目', ...monthlyHeaders, actualsHeader, forecastHeader, '年間合計'];
  sheet.getRange(2, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setWrap(true);
  const outputMatrix = {};
  items.forEach(item => outputMatrix[item] = []);
  let actualTotalSales = 0, actualTotalCost = 0, actualTotalHours = 0;
  let forecastTotalSales = 0, forecastTotalCost = 0, forecastTotalHours = 0;
  let lastActualAverages = {};
  
  data.forEach((d, index) => {
    const costDetails = d.costDetails || {};
    const doctorCount = d.doctorCount || {};
    const monthlyTotalCost = d.cost || 0;
    const monthlyTotalHours = d.hours || 0; 
    const sales = d.sales || 0;
    const budget = d.budget || 0;
    
    // ★★★ 修正箇所 ★★★
    const spotTotal = (costDetails['直接応募医師給与'] || 0) + 
                      (costDetails['紹介会社医師給与'] || 0) + 
                      (costDetails['所定休出医師給与'] || 0);
    
    if (index < firstForecastIndex) {
      actualTotalSales += sales; actualTotalCost += monthlyTotalCost; actualTotalHours += monthlyTotalHours;
    } else {
      forecastTotalSales += sales; forecastTotalCost += monthlyTotalCost; forecastTotalHours += monthlyTotalHours;
    }
    outputMatrix['予算'].push(budget); outputMatrix['売上'].push(sales);
    outputMatrix['常勤医師給与'].push(costDetails['常勤医師給与']);
    outputMatrix['定期非常勤給与'].push(costDetails['定期非常勤給与']);
    outputMatrix['スポット医師給与'].push(spotTotal);
    outputMatrix['直接応募医師給与'].push(costDetails['直接応募医師給与']);
    outputMatrix['紹介会社医師給与'].push(costDetails['紹介会社医師給与']);
    outputMatrix['所定休出医師給与'].push(costDetails['所定休出医師給与']);
    outputMatrix['紹介手数料'].push(costDetails['紹介手数料']);
    outputMatrix['特別手当'].push(costDetails['特別手当']); 
    outputMatrix['人件費率'].push(sales > 0 ? monthlyTotalCost / sales : null);
    outputMatrix['残額'].push(budget - monthlyTotalCost);
    outputMatrix['昨年度平均時給 (当拠点)'].push(lastYearAvgWages.get(d.month));
    
    if (index < firstForecastIndex) {
      outputMatrix['平均時給 (当拠点)'].push((monthlyTotalHours > 0) ? (monthlyTotalCost / monthlyTotalHours) : null);
      outputMatrix[`平均時給 (${groupName})`].push(d.averages ? d.averages[groupName] : null);
      outputMatrix[`平均時給 (${regionName})`].push(d.averages ? d.averages[regionName] : null);
      lastActualAverages = d.averages || {};
    } else {
      outputMatrix['平均時給 (当拠点)'].push(actualsSimpleAverage);
      outputMatrix[`平均時給 (${groupName})`].push(lastActualAverages[groupName] || null);
      outputMatrix[`平均時給 (${regionName})`].push(lastActualAverages[regionName] || null);
    }
    outputMatrix['常勤医師数'].push(doctorCount['常勤']);
    outputMatrix['定期非常勤医師数'].push(doctorCount['定期非常勤']);
    outputMatrix['直接応募医師数'].push(doctorCount['直接応募医師']);
    outputMatrix['紹介会社医師数'].push(doctorCount['紹介会社医師']);
  });
  
  const outputRows = items.map(item => {
    const rowData = outputMatrix[item];
    const actualsData = rowData.slice(0, firstForecastIndex);
    const forecastData = rowData.slice(firstForecastIndex);
    let actualsTotal, forecastTotal, grandTotal;
    
    if (avgItems.includes(item) || item === '人件費率') {
      if (item === '平均時給 (当拠点)') {
        actualsTotal = actualsSimpleAverage; 
        forecastTotal = actualsSimpleAverage; 
        const allValidData = rowData.filter(v => v !== null);
        grandTotal = allValidData.length > 0 ? allValidData.reduce((a, b) => a + b, 0) / allValidData.length : null;
      } else if (item === '昨年度平均時給 (当拠点)') {
        actualsTotal = lyActualsSimpleAverage;
        forecastTotal = lyActualsSimpleAverage;
        const allValidData = rowData.filter(v => v !== null);
        grandTotal = allValidData.length > 0 ? allValidData.reduce((a, b) => a + b, 0) / allValidData.length : null;
      } 
      else if (item === '人件費率') {
        actualsTotal = actualTotalSales > 0 ? actualTotalCost / actualTotalSales : null;
        forecastTotal = forecastTotalSales > 0 ? forecastTotalCost / forecastTotalSales : null;
        grandTotal = (actualTotalSales + forecastTotalSales) > 0 ? (actualTotalCost + forecastTotalCost) / (actualTotalSales + forecastTotalSales) : null;
      } else {
        const latestValue = (item.includes(groupName)) ? lastActualAverages[groupName] : (item.includes(regionName) ? lastActualAverages[regionName] : null);
        actualsTotal = latestValue; forecastTotal = latestValue; grandTotal = latestValue;
      }
    } else if (countItems.includes(item)) {
      actualsTotal = actualsData.length > 0 ? actualsData.reduce((sum, val) => sum + (Number(val) || 0), 0) / actualsData.length : 0;
      forecastTotal = forecastData.length > 0 ? forecastData.reduce((sum, val) => sum + (Number(val) || 0), 0) / forecastData.length : 0;
      grandTotal = rowData.length > 0 ? rowData.reduce((sum, val) => sum + (Number(val) || 0), 0) / rowData.length : 0;
    } else {
      actualsTotal = actualsData.reduce((sum, val) => sum + (Number(val) || 0), 0);
      forecastTotal = forecastData.reduce((sum, val) => sum + (Number(val) || 0), 0);
      grandTotal = actualsTotal + forecastTotal;
    }
    return [item, ...rowData, actualsTotal, forecastTotal, grandTotal];
  });
  
  sheet.getRange(3, 1, outputRows.length, headers.length).setValues(outputRows);
  outputRows.forEach((row, i) => {
    const item = row[0], targetRow = i + 3;
    const range = sheet.getRange(targetRow, 2, 1, headers.length - 1);
    if (moneyItems.includes(item)) range.setNumberFormat("#,##0");
    else if (avgItems.includes(item)) range.setNumberFormat('"¥"#,##0"/時"');
    else if (item === '人件費率') range.setNumberFormat("0.0%");
    else if (countItems.includes(item)) range.setNumberFormat("#,##0.0");
  });
  
  sheet.setColumnWidth(1, 180);
  for (let i = 2; i <= headers.length; i++) { sheet.setColumnWidth(i, 100); }
  if (firstForecastIndex < data.length) {
    const forecastStartCol = firstForecastIndex + 2; 
    const numCols = (data.length - firstForecastIndex);
    if (numCols > 0) sheet.getRange(3, forecastStartCol, items.length, numCols).setBackground("#cfe2f3");
    sheet.getRange(3, headers.length - 2, items.length, 1).setBackground("#cfe2f3");
  }
  
  sheet.getCharts().forEach(chart => sheet.removeChart(chart));
  let chartYPosition = outputRows.length + 4;
  if (chartsToCreate.includes('salesAndRate')) {
    // ...(chart creation logic is unchanged)...
  }
  if (chartsToCreate.includes('costBreakdown')) {
    // ...(chart creation logic is unchanged)...
  }
}