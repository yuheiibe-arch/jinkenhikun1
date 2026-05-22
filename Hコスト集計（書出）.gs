/**
 * =================================================================
 * 【ファイル 4/6: 日次コスト書き出し (C_DailyCost_Writer.gs)】
 * =================================================================
 */

// どんな形式のセルからも「日 (day)」を数字で確実に取得する
function _getDayFromCell(cell) {
  if (!cell) return 0;
  if (cell instanceof Date) {
    try { return cell.getDate(); } catch(e) { return 0; }
  }
  if (typeof cell === 'number') {
    return (cell > 0 && cell < 32) ? cell : 0;
  }
  if (typeof cell === 'string') {
    if (cell.includes('/')) {
      const parts = cell.split('/');
      return parseInt(parts[parts.length - 1], 10) || 0;
    }
    return parseInt(cell, 10) || 0;
  }
  return 0;
}

function _writeResultsToSheet(allDailyData, startDate, endDate, sheet, settings, holidaySet, totalDaysInMonth) {
  const lastCol = sheet.getLastColumn();
  if (lastCol >= 6) {
    sheet.getRange(2, 6, sheet.getMaxRows() - 1, lastCol - 5).clear({
      contentsOnly: true,
      formatOnly: true
    });
  }
  const dateHeaders = [];
  const dateKeys = [];
  const dow = ["日", "月", "火", "水", "木", "金", "土"];
  for (let d = new Date(startDate); d <= endDate; d.setDate(d.getDate() + 1)) {
    const dateKey = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
    const isHoliday = holidaySet.has(Utilities.formatDate(d, "JST", "yyyy-MM-dd"));
    dateHeaders.push(`${d.getMonth() + 1}/${d.getDate()}(${dow[d.getDay()]}${isHoliday ? "祝" : ""})`);
    dateKeys.push(dateKey);
  }
  if (dateHeaders.length === 0) return;
  dateHeaders.push("合計");
  sheet.getRange(1, 6, 1, dateHeaders.length).setValues([dateHeaders]).setFontWeight("bold").setHorizontalAlignment("center");

  const salesDataMap = new Map();
  const monthlyForecastSalesTotalMap = new Map();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const monthStr = `${startDate.getFullYear()}/${String(startDate.getMonth() + 1).padStart(2, '0')}`;
  const prevYearMonthStr = `${startDate.getFullYear() - 1}/${String(startDate.getMonth() + 1).padStart(2, '0')}`;
  
  try {
    if (startDate < today) {
      const salesSheet = _getSalesSheetForMonth(monthStr, settings);
      if (salesSheet) _parseAndStoreDailySales(salesSheet.getDataRange().getValues(), settings.clinicAliasMap, salesDataMap);
    }
    if (endDate >= today) {
      // ★売上予測（昨年のデータ流用）を復活
      const prevYearSalesSheet = _getSalesSheetForMonth(prevYearMonthStr, settings);
      if (prevYearSalesSheet) _calculateMonthlyTotalSales(prevYearSalesSheet.getDataRange().getValues(), settings.clinicAliasMap, monthlyForecastSalesTotalMap);
    }
  } catch (e) {
    Logger.log(`売上データ準備エラー: ${e.message}`);
  }

  const layout = sheet.getRange("A2:E" + sheet.getLastRow()).getValues();
  const month = startDate.getMonth() + 1;
  const targetYear = startDate.getFullYear();
  const monthlyTotals = {};
  for (const dateKey of dateKeys) {
    if (!allDailyData[dateKey]) continue;
    for (const clinicName in allDailyData[dateKey]) {
      if (!monthlyTotals[clinicName]) {
        monthlyTotals[clinicName] = {
          totalCost: 0,
          totalHours: 0,
          costDetails: { '常勤医師給与': 0, '定期非常勤給与': 0, '直接応募医師給与': 0, '紹介会社医師給与': 0, '所定休出医師給与': 0, '紹介手数料': 0, '特別手当': 0 }
        };
      }
      const dayData = allDailyData[dateKey][clinicName] || {};
      const clinicMT = monthlyTotals[clinicName];
      clinicMT.totalCost += dayData.totalCost || 0;
      clinicMT.totalHours += dayData.totalHours || 0;
      if (dayData.costDetails) {
        for (const itemKey in clinicMT.costDetails) {
          clinicMT.costDetails[itemKey] += dayData.costDetails[itemKey] || 0;
        }
      }
    }
  }
  const lastYearDailyData = settings.lastYearDailyData || {};
  const lastYearMonthlyTotals = {};
  for (const dateKey_LY in lastYearDailyData) {
    for (const clinicName in lastYearDailyData[dateKey_LY]) {
      if (!lastYearMonthlyTotals[clinicName]) {
        lastYearMonthlyTotals[clinicName] = { totalCost: 0, totalHours: 0 };
      }
      lastYearMonthlyTotals[clinicName].totalCost += lastYearDailyData[dateKey_LY][clinicName].cost || 0;
      lastYearMonthlyTotals[clinicName].totalHours += lastYearDailyData[dateKey_LY][clinicName].hours || 0;
    }
  }
  const outputMatrix = [];
  const carryOver = {};
  layout.forEach(row => {
    const currentClinicName = row[0];
    const item = row[4];
    if (!isClinicRow(currentClinicName) || !item) {
      outputMatrix.push(Array(dateKeys.length + 1).fill(null));
      return;
    }
    if (item === "予算") carryOver[currentClinicName] = 0;
    const dailyOutputRow = [];
    const budgetData = settings.budgetMap.get(currentClinicName);
    let monthlyBudgetValue = 0;
    
    // ★ ここに入ってくるbudgetDataが「2026年専用」になるよう、大元の読み込み部分を直す必要があります
    if (budgetData) {
      const ftBudget = (budgetData.ft_budget && budgetData.ft_budget[month]) || 0;
      const ptBudget = (budgetData.pt_budget && budgetData.pt_budget[month]) || 0;
      const spotBudget = (budgetData.spot_budget && budgetData.spot_budget[month]) || 0;
      const secondExamBudget = (budgetData.second_exam_budget && budgetData.second_exam_budget[month]) || 0;
      monthlyBudgetValue = Number(ftBudget) + Number(ptBudget) + Number(spotBudget) + Number(secondExamBudget);
    }
    
    const clinicMonthlyTotal = monthlyTotals[currentClinicName] || {};
    const clinicMonthlyCost = clinicMonthlyTotal.totalCost || 0;
    const clinicMonthlyHours = clinicMonthlyTotal.totalHours || 0;
    const lastYearMonthlyTotal = lastYearMonthlyTotals[currentClinicName] || {};
    const lastYearMonthlyCost = lastYearMonthlyTotal.totalCost || 0;
    const lastYearMonthlyHours = lastYearMonthlyTotal.totalHours || 0;
    let monthlySalesValue = 0;
    const clinicSalesData = salesDataMap.get(currentClinicName);
    const forecastTotal = monthlyForecastSalesTotalMap.get(currentClinicName) || 0;
    dateKeys.forEach(dateKey => {
      const currentDate = new Date(dateKey);
      const dailySalesValue = (clinicSalesData && clinicSalesData.get(currentDate.getDate())) || 0;
      monthlySalesValue += (currentDate < today) ? dailySalesValue : (forecastTotal > 0 ? forecastTotal / totalDaysInMonth : 0);
    });
    for (const dateKey of dateKeys) {
      const dayAllClinicsData = allDailyData[dateKey] || {};
      const dayData = dayAllClinicsData[currentClinicName] || { totalCost: 0, totalHours: 0, costDetails: {} };
      const costDetails = dayData.costDetails || {};
      const lastYearDateKey = dateKey.replace(targetYear, targetYear - 1);
      const lastYearDayAllClinicsData = lastYearDailyData[lastYearDateKey] || {};
      const lastYearDayData = lastYearDayAllClinicsData[currentClinicName] || { cost: 0, hours: 0 };
      const currentDate = new Date(dateKey);
      const dailySales = (currentDate < today) ? ((clinicSalesData && clinicSalesData.get(currentDate.getDate())) || 0) : ((monthlyForecastSalesTotalMap.get(currentClinicName) || 0) / totalDaysInMonth);
      let dayValue;
      switch (item) {
        case "売上": dayValue = dailySales; break;
        case "予算": dayValue = monthlyBudgetValue / totalDaysInMonth; break;
        case "常勤医師給与": dayValue = costDetails['常勤医師給与'] || 0; break;
        case "定期非常勤給与": dayValue = costDetails['定期非常勤給与'] || 0; break;
        case "スポット医師給与": 
          dayValue = (costDetails['直接応募医師給与'] || 0) + 
                     (costDetails['紹介会社医師給与'] || 0) + 
                     (costDetails['所定休出医師給与'] || 0); 
          break;
        case "直接応募医師給与": dayValue = costDetails['直接応募医師給与'] || 0; break;
        case "紹介会社医師給与": dayValue = costDetails['紹介会社医師給与'] || 0; break;
        case "所定休出医師給与": dayValue = costDetails['所定休出医師給与'] || 0; break;
        case "紹介手数料": dayValue = costDetails['紹介手数料'] || 0; break;
        case "特別手当": dayValue = costDetails['特別手当'] || 0; break;
        case "今年平均時給": dayValue = (dayData.totalHours > 0) ? dayData.totalCost / dayData.totalHours : 0; break;
        case "昨年平均時給": dayValue = (lastYearDayData.hours > 0) ? lastYearDayData.cost / lastYearDayData.hours : 0; break;
        case "人件費率": dayValue = (dailySales > 0) ? dayData.totalCost / dailySales : 0; break;
        case "残額":
          const todayBalance = (monthlyBudgetValue / totalDaysInMonth) - dayData.totalCost + (carryOver[currentClinicName] || 0);
          dayValue = todayBalance;
          carryOver[currentClinicName] = todayBalance;
          break;
        default: dayValue = 0;
      }
      dailyOutputRow.push(dayValue);
    }
    let totalColValue;
    const clinicMonthlyDetails = (monthlyTotals[currentClinicName] && monthlyTotals[currentClinicName].costDetails) || {};
    if (item === "売上") { totalColValue = monthlySalesValue; } 
    else if (item === "予算") { totalColValue = monthlyBudgetValue; } 
    else if (item === "残額") { totalColValue = monthlyBudgetValue - clinicMonthlyCost; } 
    else if (item === "今年平均時給") { totalColValue = (clinicMonthlyHours > 0) ? clinicMonthlyCost / clinicMonthlyHours : 0; } 
    else if (item === "昨年平均時給") { totalColValue = (lastYearMonthlyHours > 0) ? lastYearMonthlyCost / lastYearMonthlyHours : 0; } 
    else if (item === "人件費率") { totalColValue = (monthlySalesValue > 0) ? clinicMonthlyCost / monthlySalesValue : 0; } 
    else if (item === "スポット医師給与") { 
      totalColValue = (clinicMonthlyDetails['直接応募医師給与'] || 0) + 
                      (clinicMonthlyDetails['紹介会社医師給与'] || 0) + 
                      (clinicMonthlyDetails['所定休出医師給与'] || 0); 
    } 
    else { totalColValue = clinicMonthlyDetails[item] || 0; }
    dailyOutputRow.push(totalColValue);
    outputMatrix.push(dailyOutputRow);
  });
  
  const aggData = { "関東": { daily: {}, activeClinics: new Set() }, "関西": { daily: {}, activeClinics: new Set() } };
  layout.forEach((row, index) => {
    const currentClinicName = row[0];
    const area = row[3];
    const monthlyTotalObject = monthlyTotals[currentClinicName] || {};
    if (isClinicRow(currentClinicName) && !area.includes("除外") && (outputMatrix[index] && (outputMatrix[index][dateKeys.length] > 0 || monthlyTotalObject.totalCost > 0))) {
      const aggTargetKey = (area.includes("大阪")) ? "関西" : "関東";
      aggData[aggTargetKey].activeClinics.add(currentClinicName);
      for (let c = 0; c < outputMatrix[index].length; c++) {
        const item = row[4];
        if (!aggData[aggTargetKey].daily[c]) aggData[aggTargetKey].daily[c] = {};
        if (!aggData[aggTargetKey].daily[c][item]) aggData[aggTargetKey].daily[c][item] = 0;
        aggData[aggTargetKey].daily[c][item] += outputMatrix[index][c] || 0;
      }
    }
  });
  
  const avgWageAggData = { "関東": { totalCost: 0, totalHours: 0, lyTotalCost: 0, lyTotalHours: 0 }, "関西": { totalCost: 0, totalHours: 0, lyTotalCost: 0, lyTotalHours: 0 }};
  layout.forEach((row) => {
    const currentClinicName = row[0];
    const area = row[3];
    if (isClinicRow(currentClinicName)) {
      const clinicMonthlyData = monthlyTotals[currentClinicName] || { totalCost: 0, totalHours: 0 };
      const aggTargetKey = (area.includes("大阪")) ? "関西" : "関東";
      if (avgWageAggData[aggTargetKey]) {
        avgWageAggData[aggTargetKey].totalCost += clinicMonthlyData.totalCost;
        avgWageAggData[aggTargetKey].totalHours += clinicMonthlyData.totalHours;
        const lyMonthlyData = lastYearMonthlyTotals[currentClinicName] || { totalCost: 0, totalHours: 0 };
        avgWageAggData[aggTargetKey].lyTotalCost += lyMonthlyData.totalCost;
        avgWageAggData[aggTargetKey].lyTotalHours += lyMonthlyData.totalHours;
      }
    }
  });

  const cColumnOutput = Array(layout.length).fill([null]);
  layout.forEach((row, index) => {
    const currentClinicName = row[0];
    if (currentClinicName === "関東" || currentClinicName === "関西") {
      const summary = aggData[currentClinicName];
      const item = row[4];
      const dailyOutputRow = [];
      for (let c = 0; c < dateKeys.length; c++) {
        const dayData = summary.daily[c] || {};
        let dayValue = dayData[item] || 0;
        
        const dayTotalCost = (dayData["常勤医師給与"] || 0) + (dayData["定期非常勤給与"] || 0) + (dayData["直接応募医師給与"] || 0) + (dayData["紹介会社医師給与"] || 0) + (dayData["所定休出医師給与"] || 0) + (dayData["特別手当"] || 0);
        
        if (item === "スポット医師給与") {
            dayValue = (dayData["直接応募医師給与"] || 0) + (dayData["紹介会社医師給与"] || 0) + (dayData["所定休出医師給与"] || 0);
        } else if (item === "人件費率" || item.includes("平均時給") || item === "残額") {
          const dayTotalSales = dayData["売上"] || 0;
          let dayTotalHours = 0;
          summary.activeClinics.forEach(cName => {
            const dayAllClinicsData = allDailyData[dateKeys[c]] || {};
            const clinicDayData = dayAllClinicsData[cName] || {};
            dayTotalHours += clinicDayData.totalHours || 0;
          });
          let lyTotalCost_Daily = 0;
          let lyTotalHours_Daily = 0;
          const lastYearDateKey = dateKeys[c].replace(targetYear, targetYear - 1);
          summary.activeClinics.forEach(cName => {
            const lyDayAllClinicsData = lastYearDailyData[lastYearDateKey] || {};
            const lyDayData = lyDayAllClinicsData[cName] || { cost: 0, hours: 0 };
            lyTotalCost_Daily += lyDayData.cost;
            lyTotalHours_Daily += lyDayData.hours;
          });
          if (item === "人件費率") { dayValue = dayTotalSales > 0 ? dayTotalCost / dayTotalSales : 0; } 
          else if (item === "残額") { dayValue = (dayData["予算"] || 0) - dayTotalCost; } 
          else if (item === "今年平均時給") { dayValue = dayTotalHours > 0 ? dayTotalCost / dayTotalHours : 0; } 
          else if (item === "昨年平均時給") { dayValue = (lyTotalHours_Daily > 0) ? (lyTotalCost_Daily / lyTotalHours_Daily) : 0; }
        }
        dailyOutputRow.push(dayValue);
      }
      const monthlyData = summary.daily[dateKeys.length] || {};
      let totalColValue = monthlyData[item] || 0;
      const monthlyTotalCost = (monthlyData["常勤医師給与"] || 0) + (monthlyData["定期非常勤給与"] || 0) + (monthlyData["直接応募医師給与"] || 0) + (monthlyData["紹介会社医師給与"] || 0) + (monthlyData["所定休出医師給与"] || 0) + (monthlyData["特別手当"] || 0);
      
      if (item === "今年平均時給") {
        const avgWageData = avgWageAggData[currentClinicName];
        totalColValue = (avgWageData.totalHours > 0) ? avgWageData.totalCost / avgWageData.totalHours : 0;
      } else if (item === "昨年平均時給") {
        const avgWageData = avgWageAggData[currentClinicName];
        totalColValue = (avgWageData.lyTotalHours > 0) ? avgWageData.lyTotalCost / avgWageData.lyTotalHours : 0;
      } else if (item === "人件費率") {
        const totalSales = monthlyData["売上"] || 0;
        totalColValue = totalSales > 0 ? monthlyTotalCost / totalSales : 0;
      } else if (item === "残額") {
        const totalBudget = monthlyData["予算"] || 0;
        totalColValue = totalBudget - monthlyTotalCost;
      } else if (item === "スポット医師給与") {
        totalColValue = (monthlyData["直接応募医師給与"] || 0) + (monthlyData["紹介会社医師給与"] || 0) + (monthlyData["所定休出医師給与"] || 0);
      }
      dailyOutputRow.push(totalColValue);
      outputMatrix[index] = dailyOutputRow;
      cColumnOutput[index] = [`${summary.activeClinics.size}拠点`];
    }
  });
  if (outputMatrix.length > 0) {
    sheet.getRange(2, 6, layout.length, dateHeaders.length).setValues(outputMatrix);
    cColumnOutput.forEach((val, idx) => {
      if (val[0] !== null) sheet.getRange(idx + 2, 3).setValue(val[0]);
    });
    layout.forEach((row, index) => {
      const item = row[4];
      if (!item) return;
      const targetRow = index + 2;
      if (item === "人件費率") {
        sheet.getRange(targetRow, 6, 1, dateHeaders.length).setNumberFormat("0.0%");
      } else if (item.includes("平均時給")) {
        sheet.getRange(targetRow, 6, 1, dateHeaders.length).setNumberFormat('"¥"#,##0"/時"');
      } else {
        sheet.getRange(targetRow, 6, 1, dateHeaders.length).setNumberFormat("#,##0");
      }
    });
  }
}

function isClinicRow(clinicName) {
  return clinicName && clinicName !== "関東" && clinicName !== "関西";
}

function _getSalesSheetForMonth(monthStr, settings) {
  const [year, month] = monthStr.split('/').map(Number);
  const fiscalYear = (month <= 3) ? (year - 1) : year;
  const urlMap = settings.salesUrlMaps[fiscalYear];
  if (!urlMap) {
    return null;
  }
  const monthUrl = urlMap.get(month - 1);
  if (!monthUrl) {
    return null;
  }
  try {
    return SpreadsheetApp.openByUrl(monthUrl).getSheetByName('来院数');
  } catch (e) {
    return null;
  }
}

function _parseAndStoreDailySales(allSalesData, clinicAliasMap, outputMap) {
  let salesBlockStartIndex = -1;
  let keywordRowIndex = -1;
  for (let r = 0; r < allSalesData.length; r++) {
    if (String(allSalesData[r][0]).includes('売上：実績')) {
      keywordRowIndex = r;
      salesBlockStartIndex = r + 1;
      break;
    }
  }
  if (salesBlockStartIndex === -1) return;

  const dateHeaderRow = allSalesData[2]; 

  for (let r = salesBlockStartIndex; r < allSalesData.length; r++) {
    const sheetClinicName = allSalesData[r][1];
    if (!sheetClinicName) break;
    const normalizedClinicName = gt_normalizeClinicName(sheetClinicName, clinicAliasMap);
    if (!outputMap.has(normalizedClinicName)) {
      outputMap.set(normalizedClinicName, new Map());
    }
    const clinicDateMap = outputMap.get(normalizedClinicName);
    for (let c = 2; c < dateHeaderRow.length; c++) {
      const day = _getDayFromCell(dateHeaderRow[c]);
      if (day > 0) {
        clinicDateMap.set(day, parseFloat(String(allSalesData[r][c]).replace(/,/g, '')) || 0);
      }
    }
  }
}

function _calculateMonthlyTotalSales(allSalesData, clinicAliasMap, outputMap) {
  let salesBlockStartIndex = -1;
  for (let r = 0; r < allSalesData.length; r++) {
    if (String(allSalesData[r][0]).includes('売上：実績')) {
      salesBlockStartIndex = r + 1;
      break;
    }
  }
  if (salesBlockStartIndex === -1) return;

  const dateHeaderRow = allSalesData[2];

  for (let r = salesBlockStartIndex; r < allSalesData.length; r++) {
    const sheetClinicName = allSalesData[r][1];
    if (!sheetClinicName) break;
    const normalizedClinicName = gt_normalizeClinicName(sheetClinicName, clinicAliasMap);
    let totalSales = 0;
    for (let c = 2; c < dateHeaderRow.length; c++) {
      const day = _getDayFromCell(dateHeaderRow[c]);
      if (day > 0) {
        totalSales += parseFloat(String(allSalesData[r][c]).replace(/,/g, '')) || 0;
      }
    }
    outputMap.set(normalizedClinicName, totalSales);
  }
}