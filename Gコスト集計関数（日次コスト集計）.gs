/**
 * =================================================================
 * 【ファイル 3/6: 日次コスト計算エンジン (C_DailyCost_Calculator.gs)】- 修正版
 * =================================================================
 */

function _calculateDailyPersonnelCosts(targetDate, dailyDataForOneDay, settings) {
  const targetDateStr = `${targetDate.getFullYear()}/${targetDate.getMonth() + 1}/${targetDate.getDate()}`;
  const monthStr = `${targetDate.getFullYear()}/${String(targetDate.getMonth() + 1).padStart(2, '0')}`;
  if (!dailyDataForOneDay[targetDateStr]) {
    dailyDataForOneDay[targetDateStr] = {};
  }
  const dailyAggregates = dailyDataForOneDay[targetDateStr];
  let shiftSheet;
  try {
    const indexSpreadsheet = SpreadsheetApp.openByUrl(settings.actualShiftUrl);
    const today = new Date();
    const currentMonthStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}`;
    if (monthStr === currentMonthStr && indexSpreadsheet.getSheetByName("最新")) {
      const targetUrl = indexSpreadsheet.getSheetByName("最新").getRange("A1").getValue();
      shiftSheet = SpreadsheetApp.openByUrl(targetUrl).getSheetByName("貼付用");
    } else {
      shiftSheet = indexSpreadsheet.getSheetByName(monthStr);
    }
  } catch (e) {
    Logger.log(`確定シフトシート(${monthStr})取得エラー: ${e.message}`);
    return;
  }
  if (!shiftSheet || shiftSheet.getLastRow() < 3) return;

  const allData = shiftSheet.getDataRange().getValues();
  const header = allData[0].map(String);
  const dataRows = allData.slice(2);

  const col = {
    name: header.indexOf("名前"), medicalId: header.indexOf("医籍番号"), clinic: header.indexOf("クリニック名"),
    startTime: header.indexOf("勤務開始時間"), endTime: header.indexOf("勤務終了時間"), totalPay: header.indexOf("時給合計"),
    extraPay1: header.indexOf("追加支給額1"), extraPay2: header.indexOf("追加支給額2"), workDate: header.indexOf("勤務日"),
    department: header.indexOf("診療科"), comment1: header.indexOf("スタッフコメント1")
  };

  const targetDateStrCompare = targetDate.toDateString();

  for (let i = 0; i < dataRows.length; i++) {
    const shift = dataRows[i];
    const workDateVal = shift[col.workDate];
    if (!workDateVal || !(workDateVal instanceof Date) || workDateVal.toDateString() !== targetDateStrCompare) continue;

    const officialClinicName = gt_normalizeClinicName(shift[col.clinic], settings.clinicAliasMap);
    const normalizedName = gt_normalizePersonName(shift[col.name], settings.nameAliasMap);
    const medicalId = settings.nameToMedicalIdMap.get(normalizedName) || (shift[col.medicalId] ? String(shift[col.medicalId]).trim() : null);

    if (!officialClinicName || !normalizedName) continue;
    
    if (!dailyAggregates[officialClinicName]) {
      dailyAggregates[officialClinicName] = {
        sales: 0, totalCost: 0, totalHours: 0,
        costDetails: { '常勤医師給与': 0, '定期非常勤給与': 0, '直接応募医師給与': 0, '紹介会社医師給与': 0, '所定休出医師給与': 0, '紹介手数料': 0, '特別手当': 0 }
      };
    }

    const agg = dailyAggregates[officialClinicName];
    const doctorType = _getDoctorTypeForShift(medicalId, normalizedName, workDateVal, officialClinicName, settings);
    
    let hours = 0;
    const startTimeObj = shift[col.startTime];
    const endTimeObj = shift[col.endTime];
    if (startTimeObj instanceof Date && endTimeObj instanceof Date) {
      hours = (endTimeObj.getTime() - startTimeObj.getTime()) / 3600000;
      if (hours > 4 && startTimeObj.getHours() < 13 && endTimeObj.getHours() >= 15) { hours -= 2; }
    }
    if (hours <= 0) continue;
    agg.totalHours += hours;

    const specialAllowance = Number(shift[col.extraPay1]) || 0;
    if (specialAllowance > 0) agg.costDetails['特別手当'] += specialAllowance;

    let cost = 0;
    
    if (doctorType === "常勤") {
      const hourlyWage = settings.ftHourlyWageMap.get(normalizedName) || 10000;
      cost = hours * hourlyWage;
      const comment = (shift[col.comment1] || '').toString();
      const costType = comment.includes("所定休出") ? '所定休出医師給与' : '常勤医師給与';
      agg.costDetails[costType] += cost;

    } else if (doctorType === "定期非常勤") {
      cost = (Number(shift[col.totalPay]) || 0) + (Number(shift[col.extraPay2]) || 0);
      if (cost === 0 && hours > 0) cost = 10000 * hours;
      agg.costDetails['定期非常勤給与'] += cost;

    } else { // スポット
      cost = (Number(shift[col.totalPay]) || 0) + (Number(shift[col.extraPay2]) || 0);
      if (cost === 0 && hours > 0) cost = 10000 * hours;
      
      const REFERRAL_FEE_RATE = 0.20;
      if (settings.agencyShiftSet.has(`${normalizedName}|${targetDateStr}`)) {
        const doctorSalary = cost / (1 + REFERRAL_FEE_RATE);
        const referralFee = cost - doctorSalary;
        agg.costDetails['紹介会社医師給与'] += doctorSalary;
        agg.costDetails['紹介手数料'] += referralFee;
      } else {
        agg.costDetails['直接応募医師給与'] += cost;
      }
    }
  }

  for (const clinic in dailyAggregates) {
    dailyAggregates[clinic].totalCost = Object.values(dailyAggregates[clinic].costDetails).reduce((sum, val) => sum + val, 0);
  }
}

function _calculateLastYearActuals_FOR_DAILY_TOOL(monthStr, settings) {
  const allDailyData_LastYear = {};
  const [year, monthNum] = monthStr.split('/').map(Number);
  
  // ▼▼▼ 修正箇所：1〜3月の場合は、年度(fiscalYear)を前年に設定 ▼▼▼
  const fiscalYear = (monthNum <= 3) ? year - 1 : year;
  // ▲▲▲ 修正箇所終わり ▲▲▲

  let shiftSheet;
  try {
    const urlMap = settings.actualShiftUrls;
    // 修正：year ではなく fiscalYear をキーにする
    if (!urlMap || !urlMap.has(fiscalYear)) throw new Error(`確定シフトURL(${fiscalYear}年度) が見つかりません。`);
    shiftSheet = SpreadsheetApp.openByUrl(urlMap.get(fiscalYear)).getSheetByName(monthStr);
  } catch (e) {
    Logger.log(`昨年度確定シフト(${monthStr})取得エラー: ${e.message}`);
    return allDailyData_LastYear;
  }
  if (!shiftSheet || shiftSheet.getLastRow() < 3) return allDailyData_LastYear;

  const header = shiftSheet.getRange(1, 1, 1, shiftSheet.getLastColumn()).getValues()[0];
  const col = {
    name: header.indexOf("名前"), medicalId: header.indexOf("医籍番号"), clinic: header.indexOf("クリニック名"),
    startTime: header.indexOf("勤務開始時間"), endTime: header.indexOf("勤務終了時間"), totalPay: header.indexOf("時給合計"),
    extraPay1: header.indexOf("追加支給額1"), extraPay2: header.indexOf("追加支給額2"), workDate: header.indexOf("勤務日"),
    department: header.indexOf("診療科")
  };
  const dataRange = shiftSheet.getRange(3, 1, shiftSheet.getLastRow() - 2, header.length);
  if (dataRange.getNumRows() === 0) return allDailyData_LastYear;

  const shiftData_Values = dataRange.getValues();
  const startTime_DisplayStrings = dataRange.offset(0, col.startTime, dataRange.getNumRows(), 1).getDisplayValues();
  const endTime_DisplayStrings = dataRange.offset(0, col.endTime, dataRange.getNumRows(), 1).getDisplayValues();

  for (let i = 0; i < shiftData_Values.length; i++) {
    const shift = shiftData_Values[i];
    const workDateVal = shift[col.workDate];
    if (!workDateVal || !(workDateVal instanceof Date) || (workDateVal.getMonth() + 1) !== monthNum) continue;
    
    const dateKey = `${workDateVal.getFullYear()}/${workDateVal.getMonth() + 1}/${workDateVal.getDate()}`;
    if (!allDailyData_LastYear[dateKey]) allDailyData_LastYear[dateKey] = {};
    
    const officialClinicName = gt_normalizeClinicName(shift[col.clinic], settings.clinicAliasMap);
    const medicalId = shift[col.medicalId] ? String(shift[col.medicalId]).trim() : null;
    const normalizedName = gt_normalizePersonName(shift[col.name], settings.nameAliasMap);
    if (!officialClinicName || (!medicalId && !normalizedName)) continue;
    
    if (!allDailyData_LastYear[dateKey][officialClinicName]) {
      allDailyData_LastYear[dateKey][officialClinicName] = { cost: 0, hours: 0 };
    }
    const agg = allDailyData_LastYear[dateKey][officialClinicName];
    let hours = 0;
    const startTimeObj = shift[col.startTime];
    const endTimeObj = shift[col.endTime];
    if (startTimeObj && endTimeObj && !isNaN(new Date(startTimeObj).getTime()) && !isNaN(new Date(endTimeObj).getTime())) {
      hours = (new Date(endTimeObj).getTime() - new Date(startTimeObj).getTime()) / 3600000;
    }
    const startHour = parseInt(startTime_DisplayStrings[i][0].split(':')[0], 10) || 9;
    const endHour = parseInt(endTime_DisplayStrings[i][0].split(':')[0], 10) || 18;
    if (hours > 4 && startHour < 13 && endHour >= 15) hours -= 2;
    if (hours <= 0) continue;
    
    let shiftCost = 0;
    const doctorType = _getDoctorTypeForShift(medicalId, normalizedName, workDateVal, officialClinicName, settings);
    
    if (doctorType === "常勤") {
      let hourlyWage = 0;
      if (medicalId) hourlyWage = settings.ftWageMap_LastYear.get(medicalId) || 0;
      if (hourlyWage === 0) hourlyWage = settings.ftHourlyWageMap.get(normalizedName) || 10000;
      shiftCost = hourlyWage * hours;
    } else {
      shiftCost = (Number(shift[col.totalPay]) || 0) + (Number(shift[col.extraPay1]) || 0) + (Number(shift[col.extraPay2]) || 0);
      if (shiftCost === 0 && hours > 0) {
        shiftCost = 10000 * hours;
      }
    }

    if (hours > 0 && shiftCost > 0) {
      agg.hours += hours;
      agg.cost += shiftCost;
    }
  }
  return allDailyData_LastYear;
}