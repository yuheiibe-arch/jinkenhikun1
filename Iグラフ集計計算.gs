/**
 * =================================================================
 * 【ファイル 5/X: グラフ計算 (D_Graph_Calculators.gs)】- 最終FIX版
 * =================================================================
 */

function gt_calculateActualsForMonth(clinic, monthStr, settings) {
  const allClinicsActuals = _calculateActualsForAllClinics_OneMonth(monthStr, settings);
  const normalizedClinic = gt_normalizeClinicName(clinic, settings.clinicAliasMap);
  const clinicData = allClinicsActuals.clinicDataMap.get(normalizedClinic) || { cost: 0, hours: 0, costDetails: {}, doctorCount: {} };
  const averages = _calculateAveragesFromActualsMap(allClinicsActuals.clinicDataMap, settings.clinicGroupMap);
  const salesData = gt_calculateSalesForMonth(clinic, monthStr, settings, false); 
  const month = parseInt(monthStr.split('/')[1], 10);
  const budgetData = settings.budgetMap.get(normalizedClinic) || {};
  const totalBudget = (budgetData.ft_budget?.[month] || 0) + (budgetData.pt_budget?.[month] || 0) + (budgetData.spot_budget?.[month] || 0) + (budgetData.second_exam_budget?.[month] || 0);
  return { 
    sales: salesData, cost: clinicData.cost, hours: clinicData.hours, budget: totalBudget, 
    costDetails: clinicData.costDetails, doctorCount: clinicData.doctorCount, averages: averages 
  };
}

function gt_calculateFuturePersonnelCosts(clinic, monthStr, settings, targetDate = null) {
  let totalCost = 0, totalHours = 0; 
  const costDetails = { '常勤医師給与': 0, '定期非常勤給与': 0, '直接応募医師給与': 0, '紹介会社医師給与': 0, '所定休出医師給与': 0, '紹介手数料': 0, '特別手当': 0 };
  const uniqueDoctors = { '常勤': new Set(), '定期非常勤': new Set(), '直接応募医師': new Set(), '紹介会社医師': new Set() };
  const [targetYear, month] = monthStr.split('/').map(Number);
  const targetMonth = month - 1;
  const normalizedClinic = gt_normalizeClinicName(clinic, settings.clinicAliasMap); 
  const defaultReturn = { totalCost: 0, totalHours: 0, costDetails, doctorCount: { '常勤': 0, '定期非常勤': 0, '直接応募医師': 0, '紹介会社医師': 0 } };
  if (!settings.kantoUrl || !settings.kansaiUrl) return defaultReturn;
  const kantoSs = SpreadsheetApp.openByUrl(settings.kantoUrl);
  const kansaiSs = SpreadsheetApp.openByUrl(settings.kansaiUrl);
  const allSheets = [...kantoSs.getSheets(), ...kansaiSs.getSheets()];
  const targetSheet = allSheets.find(s => gt_normalizeClinicName(s.getName(), settings.clinicAliasMap) === normalizedClinic); 
  if (!targetSheet) return defaultReturn;
  const data = targetSheet.getDataRange().getValues();
  if (data.length < 29) return defaultReturn;
  const dateRow = data[24];
  let monthStartCol = -1;
  for (let c = 0; c < dateRow.length; c++) { 
    if (dateRow[c] instanceof Date && dateRow[c].getMonth() === targetMonth && dateRow[c].getDate() === 1) { 
      monthStartCol = c; 
      break; 
    } 
  }
  if (monthStartCol === -1) return defaultReturn;
  const hourMapping = [9, 10, 11, 12, null, null, 15, 16, 17, 18, 19, 20];
  for (let r = 28; r < data.length; r += 2) {
    const ruleRow1 = data[r];
    if (!ruleRow1) continue;
    const dayOfWeekStr = ruleRow1[monthStartCol - 3], weekNumStr = ruleRow1[monthStartCol - 2];
    if (!dayOfWeekStr || !weekNumStr) continue;
    let dates = gt_calculateDatesFromRule(targetYear, targetMonth, weekNumStr, dayOfWeekStr); 
    if (targetDate) {
      dates = dates.filter(d => d.getDate() === targetDate.getDate());
    }
    if (dates.length === 0) continue;
    const shiftsInRule = dates.length;
    for (let c_offset = 0; c_offset < hourMapping.length; c_offset++) {
      if (hourMapping[c_offset] === null) continue;
      const personsRaw = [ (ruleRow1[monthStartCol + c_offset] || ''), (data[r+1] ? data[r+1][monthStartCol + c_offset] : '') ];
      personsRaw.forEach(personRaw => {
        const person = personRaw.toString().trim();
        if (!person) return;
        
        totalHours += shiftsInRule;
        
        const normalizedName = gt_normalizePersonName(person, settings.nameAliasMap); 
        
        let doctorType = '';
        if (settings.partTimeDoctorNameSet && settings.partTimeDoctorNameSet.has(normalizedName)) {
            doctorType = '定期非常勤';
        } else {
            const ftContractInfo = settings.doctorContractMap.get(normalizedName);
            doctorType = (ftContractInfo && ftContractInfo.isFullTime) ? '常勤' : 'スポット';
        }
        
        let hourlyWage = 0;
        if (doctorType === '常勤') {
          uniqueDoctors['常勤'].add(normalizedName);
          hourlyWage = settings.ftHourlyWageMap.get(normalizedName) || 10000;
          costDetails['常勤医師給与'] += hourlyWage * shiftsInRule;
        } else if (doctorType === '定期非常勤') {
          uniqueDoctors['定期非常勤'].add(normalizedName);
          hourlyWage = settings.ptContractWageMap.get(normalizedName) || 15000;
          costDetails['定期非常勤給与'] += hourlyWage * shiftsInRule;
        } else {
          const isAgency = dates.some(d => settings.agencyShiftSet.has(`${normalizedName}|${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`));
          if (isAgency) { 
              uniqueDoctors['紹介会社医師'].add(normalizedName);
              costDetails['紹介会社医師給与'] += 10000 * shiftsInRule;
          } else { 
              uniqueDoctors['直接応募医師'].add(normalizedName);
              costDetails['直接応募医師給与'] += 10000 * shiftsInRule;
          }
        }
      });
    }
  }
  
  totalCost = Object.values(costDetails).reduce((a, b) => a + b, 0);

  const finalDoctorCount = {
    '常勤': uniqueDoctors['常勤'].size, '定期非常勤': uniqueDoctors['定期非常勤'].size,
    '直接応募医師': uniqueDoctors['直接応募医師'].size, '紹介会社医師': uniqueDoctors['紹介会社医師'].size
  };
  return { totalCost, totalHours, costDetails, doctorCount: finalDoctorCount };
}

function _calculateActualsForAllClinics_OneMonth(monthStr, settings, limitDate = null) {
  const clinicDataMap = new Map();
  const [year, monthNum] = monthStr.split('/').map(Number);
  const targetMonthIndex = monthNum - 1;
  let shiftSheet;
  const REFERRAL_FEE_RATE = 0.20; 

  try {
    const today = new Date();
    const currentMonthStr = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}`;
    const indexSpreadsheet = SpreadsheetApp.openByUrl(settings.actualShiftUrl);
    
    if (monthStr === currentMonthStr && indexSpreadsheet.getSheetByName("最新")) {
      const targetUrl = indexSpreadsheet.getSheetByName("最新").getRange("A1").getValue();
      shiftSheet = SpreadsheetApp.openByUrl(targetUrl).getSheetByName("貼付用");
    } else {
      const fiscalYear = (monthNum <= 3) ? year - 1 : year; 
      const urlMap = settings.actualShiftUrls; 
      if (!urlMap || !urlMap.has(fiscalYear)) { throw new Error(`確定シフトURL(${fiscalYear}年度) が見つかりません。`); }
      const ss = SpreadsheetApp.openByUrl(urlMap.get(fiscalYear));
      shiftSheet = ss.getSheetByName(monthStr);
    }
  } catch (e) { 
    return { clinicDataMap }; 
  }
  
  if (!shiftSheet || shiftSheet.getLastRow() < 3) { return { clinicDataMap }; }
  
  const allData = shiftSheet.getDataRange().getValues();
  const header = allData[0].map(String);
  const dataRows = allData.slice(2);

  const col = { name: header.indexOf("名前"), clinic: header.indexOf("クリニック名"), startTime: header.indexOf("勤務開始時間"), endTime: header.indexOf("勤務終了時間"), totalPay: header.indexOf("時給合計"), extraPay1: header.indexOf("追加支給額1"), extraPay2: header.indexOf("追加支給額2"), workDate: header.indexOf("勤務日"), department: header.indexOf("診療科"), comment1: header.indexOf("スタッフコメント1"), medicalId: header.indexOf("医籍番号") };
  
  if(dataRows.length === 0) return { clinicDataMap }; 
  
  const monthlyUniqueDoctors = {}; 
  
  for (const shift of dataRows) { 
    const workDateVal = new Date(shift[col.workDate]);
    if (!workDateVal || isNaN(workDateVal.getTime()) || workDateVal.getMonth() !== targetMonthIndex) continue;
    
    const officialClinicName = gt_normalizeClinicName(shift[col.clinic], settings.clinicAliasMap); 
    const doctorName = shift[col.name];
    if (!doctorName || !officialClinicName) continue;
    
    if (!clinicDataMap.has(officialClinicName)) {
      clinicDataMap.set(officialClinicName, { cost: 0, hours: 0, costDetails: { '常勤医師給与': 0, '定期非常勤給与': 0, '直接応募医師給与': 0, '紹介会社医師給与': 0, '所定休出医師給与': 0, '紹介手数料': 0, '特別手当': 0} });
      monthlyUniqueDoctors[officialClinicName] = { '常勤': new Set(), '定期非常勤': new Set(), '直接応募医師': new Set(), '紹介会社医師': new Set() };
    }
    const agg = clinicDataMap.get(officialClinicName);
    const uniqueSets = monthlyUniqueDoctors[officialClinicName];
    
    const medicalId = settings.nameToMedicalIdMap.get(gt_normalizePersonName(doctorName, settings.nameAliasMap)) || (shift[col.medicalId] ? String(shift[col.medicalId]).trim() : null);
    const normalizedName = gt_normalizePersonName(doctorName, settings.nameAliasMap); 
    const doctorType = _getDoctorTypeForShift(medicalId, normalizedName, workDateVal, officialClinicName, settings); 
    
    let hours = 0;
    const startTimeObj = shift[col.startTime], endTimeObj = shift[col.endTime];     
    if (startTimeObj instanceof Date && endTimeObj instanceof Date) {
      hours = (endTimeObj.getTime() - startTimeObj.getTime()) / 3600000;
      if (hours > 4 && startTimeObj.getHours() < 13 && endTimeObj.getHours() >= 15) { hours -= 2; }
    }
    if (hours <= 0) continue;
    agg.hours += hours;

    const specialAllowance = Number(shift[col.extraPay1]) || 0;
    if (specialAllowance > 0) { agg.costDetails['特別手当'] += specialAllowance; }
    
    let cost = 0; 
    
    if (doctorType === "常勤") {
      uniqueSets['常勤'].add(normalizedName);
      const hourlyWage = settings.ftHourlyWageMap.get(normalizedName) || 10000;
      cost = hours * hourlyWage;
      const comment = (shift[col.comment1] || '').toString();
      const costType = comment.includes("所定休出") ? '所定休出医師給与' : '常勤医師給与';
      agg.costDetails[costType] += cost;
    } else {
      cost = (Number(shift[col.totalPay]) || 0) + (Number(shift[col.extraPay2]) || 0);
      if (cost === 0 && hours > 0) cost = 10000 * hours;
      
      if (doctorType === "定期非常勤") {
        uniqueSets['定期非常勤'].add(normalizedName);
        agg.costDetails['定期非常勤給与'] += cost;
      } else {
        const dateStr = `${workDateVal.getFullYear()}/${workDateVal.getMonth() + 1}/${workDateVal.getDate()}`;
        if (settings.agencyShiftSet.has(`${normalizedName}|${dateStr}`)) {
          uniqueSets['紹介会社医師'].add(normalizedName);
          const doctorSalary = cost / (1 + REFERRAL_FEE_RATE);
          const referralFee = cost - doctorSalary;
          agg.costDetails['紹介会社医師給与'] += doctorSalary;
          agg.costDetails['紹介手数料'] += referralFee;
        } else { 
          uniqueSets['直接応募医師'].add(normalizedName);
          agg.costDetails['直接応募医師給与'] += cost; 
        }
      }
    }
  }
  
  clinicDataMap.forEach((data, clinicName) => {
    data.cost = Object.values(data.costDetails).reduce((sum, val) => sum + val, 0);
    data.doctorCount = { 
      '常勤': monthlyUniqueDoctors[clinicName]['常勤'].size, '定期非常勤': monthlyUniqueDoctors[clinicName]['定期非常勤'].size,
      '直接応募医師': monthlyUniqueDoctors[clinicName]['直接応募医師'].size, '紹介会社医師': monthlyUniqueDoctors[clinicName]['紹介会社医師'].size
    };
  });
  return { clinicDataMap };
}

function _calculateAveragesFromActualsMap(clinicDataMap, clinicGroupMap) {
  const groupTotals = {}; 
  const kantoGroups = ["東京第一", "東京第二", "埼玉", "神奈川", "千葉"];
  const kansaiGroups = ["大阪"]; 
  clinicDataMap.forEach((data, clinicName) => {
    if (!clinicGroupMap.has(clinicName)) return;
    const group = clinicGroupMap.get(clinicName); 
    if (!group) return; 
    
    if (!groupTotals[group]) groupTotals[group] = { cost: 0, hours: 0 };
    groupTotals[group].cost += data.cost;
    groupTotals[group].hours += data.hours;
    if (kantoGroups.includes(group)) {
      if (!groupTotals["関東"]) groupTotals["関東"] = { cost: 0, hours: 0 };
      groupTotals["関東"].cost += data.cost;
      groupTotals["関東"].hours += data.hours;
    }
    if (kansaiGroups.includes(group)) {
      if (!groupTotals["関西"]) groupTotals["関西"] = { cost: 0, hours: 0 };
      groupTotals["関西"].cost += data.cost;
      groupTotals["関西"].hours += data.hours;
    }
  });
  const averages = {};
  for (const groupName in groupTotals) {
    const total = groupTotals[groupName];
    averages[groupName] = (total.hours > 0) ? (total.cost / total.hours) : 0;
  }
  return averages;
}

function gt_getMonthsInRange(startMonth, endMonth) {
  const months = [];
  const [startY, startM] = startMonth.split('/').map(Number);
  let currentDate = new Date(startY, startM - 1, 1);
  const [endY, endM] = endMonth.split('/').map(Number);
  const lastDate = new Date(endY, endM - 1, 1);
  while (currentDate <= lastDate) {
    const y = currentDate.getFullYear();
    const m = String(currentDate.getMonth() + 1).padStart(2, '0');
    months.push(`${y}/${m}`);
    currentDate.setMonth(currentDate.getMonth() + 1);
  }
  return months;
}

function gt_calculateSalesForMonth(clinic, monthStr, settings, isForecast) {
  const [year, month] = monthStr.split('/').map(Number);
  const targetYear = isForecast ? year - 1 : year;
  const targetMonthStr = isForecast ? `${targetYear}/${String(month).padStart(2, '0')}` : monthStr;
  return _getActualSalesForMonth(clinic, targetMonthStr, settings);
}

function _getActualSalesForMonth(clinic, monthStr, settings) {
  const [year, month] = monthStr.split('/').map(Number);
  const fiscalYear = (month <= 3) ? (year - 1) : year;
  const urlMap = settings.salesUrlMaps[fiscalYear];
  if (!urlMap) { return 0; }
  const monthUrl = urlMap.get(month - 1);
  if (!monthUrl) { return 0; }
  try {
    const salesSheet = SpreadsheetApp.openByUrl(monthUrl).getSheetByName('来院数');
    const data = salesSheet.getDataRange().getValues(); 
    let salesBlockStartIndex = -1;
    for(let r = 0; r < data.length; r++) { if (String(data[r][0]).includes('売上：実績')) { salesBlockStartIndex = r + 1; break; } }
    if (salesBlockStartIndex === -1) return 0;
    
    const normalizedSearchName = gt_normalizeClinicName(clinic, settings.clinicAliasMap); 
    let targetRowIndex = -1;
    for(let r = salesBlockStartIndex -1; r < data.length; r++) {
      const sheetClinicName = data[r][1]; if (!sheetClinicName) continue;
      if (gt_normalizeClinicName(sheetClinicName, settings.clinicAliasMap) === normalizedSearchName) { targetRowIndex = r; break; }
    }
    if (targetRowIndex === -1) return 0;
    
    const salesRow = data[targetRowIndex];
    let totalSales = 0;
    for (let c = 2; c < salesRow.length; c++) {
      if (data[2] && data[2][c] instanceof Date) { totalSales += parseFloat(String(salesRow[c]).replace(/,/g, '')) || 0; }
    }
    return totalSales;
  } catch (e) { Logger.log(`売上シート読込エラー (${monthStr}): ${e.message}`); return 0; }
}

function gt_getAvailableYears() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const urlListSheet = ss.getSheetByName("URLリスト");
    if (!urlListSheet) throw new Error("「URLリスト」シートが見つかりません。");

    const urlListData = urlListSheet.getDataRange().getValues();
    const yearRegex = /\d{4}/;
    
    const years = urlListData
      .map(row => row[0].toString())
      .filter(name => name.includes('確定シフト'))
      .map(name => {
        const match = name.match(yearRegex);
        return match ? parseInt(match[0], 10) : null;
      })
      .filter(year => year !== null);

    const uniqueYears = Array.from(new Set(years));
    uniqueYears.sort((a, b) => b - a);

    return uniqueYears.length > 0 ? uniqueYears : [new Date().getFullYear()];

  } catch (e) {
    Logger.log(e);
    return [new Date().getFullYear()];
  }
}

function gt_calculateDatesFromRule(year, month, weekNumStr, dayOfWeekStr) {
  const weekRule = String(weekNumStr);
  const dayMap = { "日": 0, "月": 1, "火": 2, "水": 3, "木": 4, "金": 5, "土": 6 };
  const targetDayOfWeek = dayMap[dayOfWeekStr.charAt(0)];
  if (targetDayOfWeek === undefined) return [];

  const dates = [];
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const weeks = weekRule.includes('毎週') ? null : weekRule.match(/\d/g).map(Number);

  for (let day = 1; day <= daysInMonth; day++) {
    const d = new Date(year, month, day);
    if (d.getDay() === targetDayOfWeek) {
      if (weeks) {
        const nthDay = Math.floor((d.getDate() - 1) / 7) + 1;
        if (weeks.includes(nthDay)) {
          dates.push(d);
        }
      } else {
        dates.push(d);
      }
    }
  }
  return dates;
}

function _calculateFutureCostsForDailyTool(startDate, endDate, settings) {
  const futureDailyData = {};
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (endDate < today) {
    return futureDailyData;
  }
  
  Logger.log("未来のコスト予測計算を開始します...");

  try {
    const kakuninSheetUrl = "https://docs.google.com/spreadsheets/d/1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA/";
    const kakuninSs = SpreadsheetApp.openByUrl(kakuninSheetUrl);
    const kakuninSheet = kakuninSs.getSheetByName('確認用');
    if (!kakuninSheet) throw new Error("未来予測の参照先「確認用」シートが見つかりません。");
    
    const kakuninData = kakuninSheet.getDataRange().getValues();
    const allClinics = Array.from(new Set(settings.clinicAliasMap.values()));

    for (let d = new Date(today.getTime()); d <= endDate; d.setDate(d.getDate() + 1)) {
      const dateStr = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
      const targetDateStrForSheet = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      futureDailyData[dateStr] = {};

      for (const clinicName of allClinics) {
          const normalizedClinic = clinicName;
          let totalDailyCost = 0;
          let totalDailyHours = 0;
          const costDetails = { '常勤医師給与': 0, '定期非常勤給与': 0, '直接応募医師給与': 0, '紹介会社医師給与': 0, '所定休出医師給与': 0, '紹介手数料': 0, '特別手当': 0 };

          const targetRow = kakuninData.find(row => 
              gt_normalizeClinicName(row[0], settings.clinicAliasMap) === normalizedClinic && 
              (row[1] ? row[1].toString().split('（')[0].trim() : '') === targetDateStrForSheet
          );

          if (targetRow) {
            const amDoctors = targetRow[7] ? targetRow[7].toString().split(', ').filter(Boolean) : [];
            const pmDoctors = targetRow[8] ? targetRow[8].toString().split(', ').filter(Boolean) : [];
            const nightDoctors = targetRow[9] ? targetRow[9].toString().split(', ').filter(Boolean) : [];
            
            const dailyUniqueDoctors = new Map();
            amDoctors.forEach(name => { if (!dailyUniqueDoctors.has(name)) dailyUniqueDoctors.set(name, new Set()); dailyUniqueDoctors.get(name).add('am'); });
            pmDoctors.forEach(name => { if (!dailyUniqueDoctors.has(name)) dailyUniqueDoctors.set(name, new Set()); dailyUniqueDoctors.get(name).add('pm'); });
            nightDoctors.forEach(name => { if (!dailyUniqueDoctors.has(name)) dailyUniqueDoctors.set(name, new Set()); dailyUniqueDoctors.get(name).add('night'); });
            
            dailyUniqueDoctors.forEach((segments, name) => {
              const normalizedName = gt_normalizePersonName(name, settings.nameAliasMap);
              const medicalId = settings.nameToMedicalIdMap.get(normalizedName) || null;
              const doctorType = _getDoctorTypeForShift(medicalId, normalizedName, d, normalizedClinic, settings);
              
              let hours = 0;
              if (segments.has('am')) hours += 4;
              if (segments.has('pm')) hours += 3;
              if (segments.has('night')) hours += 3;
              
              let cost = 0;
              if (doctorType === '常勤') {
                cost = hours * (settings.ftHourlyWageMap.get(normalizedName) || 10000);
                costDetails['常勤医師給与'] += cost;
              } else if (doctorType === '定期非常勤') {
                cost = hours * (settings.ptContractWageMap.get(normalizedName) || 10000);
                 costDetails['定期非常勤給与'] += cost;
              } else {
                cost = hours * 10000;
                const dateForAgencyCheck = `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
                if (settings.agencyShiftSet.has(`${normalizedName}|${dateForAgencyCheck}`)){
                  costDetails['紹介会社医師給与'] += cost;
                } else {
                  costDetails['直接応募医師給与'] += cost;
                }
              }
              totalDailyCost += cost;
              totalDailyHours += hours;
            });
            
            const processVacantSlot = (segment, doctors) => {
               if (doctors.length > 0) return;
               const dayType = gt_getDayType(d, settings.holidays);
               let payRates = settings.payRateMaster.get(`${normalizedClinic}||小児科`) || settings.payRateMaster.get(`${normalizedClinic}||共通`);
               let wage = payRates ? getCorrectPay({name: segment}, dayType, payRates) : 0;
               let hours = (segment === 'am') ? 4 : 3;
               const cost = hours * (wage || 10000);
               totalDailyCost += cost;
               totalDailyHours += hours;
               costDetails['直接応募医師給与'] += cost;
            };
            processVacantSlot('am', amDoctors);
            processVacantSlot('pm', pmDoctors);
            processVacantSlot('night', nightDoctors);
          } else {
             const segments = ['am', 'pm', 'night'];
             segments.forEach(segment => {
                 const dayType = gt_getDayType(d, settings.holidays);
                 let payRates = settings.payRateMaster.get(`${normalizedClinic}||小児科`) || settings.payRateMaster.get(`${normalizedClinic}||共通`);
                 let wage = payRates ? getCorrectPay({name: segment}, dayType, payRates) : 0;
                 let hours = (segment === 'am') ? 4 : 3;
                 const cost = hours * (wage || 10000);
                 totalDailyCost += cost;
                 totalDailyHours += hours;
                 costDetails['直接応募医師給与'] += cost;
             });
          }

          futureDailyData[dateStr][clinicName] = {
              totalCost: totalDailyCost,
              totalHours: totalDailyHours,
              costDetails: costDetails
          };
      }
    }
    Logger.log("未来のコスト予測計算が完了しました。");
    
  } catch (e) {
    Logger.log(`未来予測コストの計算中にエラーが発生しました: ${e.message}\n${e.stack}`);
  }
  
  return futureDailyData;
}