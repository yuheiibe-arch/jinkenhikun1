/**
 * =================================================================
 * 【ファイル 2/3: お財布くんサーバー処理 (H_OsaifuKun.gs)】- 最終完成版
 * =================================================================
 */

function getClinicListForUI() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const urlListSheet = ss.getSheetByName("URLリスト");
    if (!urlListSheet) throw new Error("「URLリスト」シートが見つかりません。");
    
    const urlListData = urlListSheet.getDataRange().getValues();
    const normalizationEntry = urlListData.find(row => row[0] === "正規表現");
    if (!normalizationEntry || !normalizationEntry[1]) throw new Error("URLリストに「正規表現」の定義が見つかりません。");

    const normalizationSs = SpreadsheetApp.openByUrl(normalizationEntry[1]);
    const clinicSheet = normalizationSs.getSheetByName('拠点名');
    if (!clinicSheet) throw new Error("「正規表現」スプレッドシートに「拠点名」シートが見つかりません。");

    const clinicNames = clinicSheet.getRange(2, 1, clinicSheet.getLastRow() - 1, 1).getValues()
      .flat()
      .filter(name => name);

    return Array.from(new Set(clinicNames)).sort();

  } catch (e) {
    Logger.log(`getClinicListForUI Error: ${e.message}`);
    return [];
  }
}

function getDailyScheduleAndCost(clinic, targetDate, settings) {
  const normalizedClinic = gt_normalizeClinicName(clinic, settings.clinicAliasMap);
  const finalSchedule = { am: [], pm: [], night: [] };
  let totalDailyCost = 0;
  
  const openingDate = settings.clinicOpeningDates.get(normalizedClinic);
  if (openingDate && targetDate < openingDate) {
    finalSchedule.am.push("（開業前）");
    finalSchedule.pm.push("（開業前）");
    finalSchedule.night.push("（開業前）");
    return { schedule: finalSchedule, dailyCost: 0 };
  }
  
  try {
    const kakuninSheetUrl = "https://docs.google.com/spreadsheets/d/1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA/";
    const kakuninSs = SpreadsheetApp.openByUrl(kakuninSheetUrl);
    const kakuninSheet = kakuninSs.getSheetByName('確認用');
    if (!kakuninSheet) throw new Error("シート「確認用」が見つかりません。");
    
    const kakuninData = kakuninSheet.getDataRange().getValues();
    const targetDateStr = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    let targetRowFound = false;
    for (const row of kakuninData) {
      const clinicNameInSheet = gt_normalizeClinicName(row[0], settings.clinicAliasMap);
      const workDateStr = row[1] ? row[1].toString().split('（')[0].trim() : '';
      
      if (clinicNameInSheet === normalizedClinic && workDateStr === targetDateStr) {
        targetRowFound = true;
        
        const amDoctors = row[7] ? row[7].toString().split(', ').filter(Boolean) : [];
        const pmDoctors = row[8] ? row[8].toString().split(', ').filter(Boolean) : [];
        const nightDoctors = row[9] ? row[9].toString().split(', ').filter(Boolean) : [];
        
        const dailyUniqueDoctors = new Map();
        amDoctors.forEach(name => { if (!dailyUniqueDoctors.has(name)) dailyUniqueDoctors.set(name, new Set()); dailyUniqueDoctors.get(name).add('am'); });
        pmDoctors.forEach(name => { if (!dailyUniqueDoctors.has(name)) dailyUniqueDoctors.set(name, new Set()); dailyUniqueDoctors.get(name).add('pm'); });
        nightDoctors.forEach(name => { if (!dailyUniqueDoctors.has(name)) dailyUniqueDoctors.set(name, new Set()); dailyUniqueDoctors.get(name).add('night'); });
        
        dailyUniqueDoctors.forEach((segments, name) => {
          const normalizedName = gt_normalizePersonName(name, settings.nameAliasMap);
          const doctorType = _getDoctorTypeForShift(settings.nameToMedicalIdMap.get(normalizedName) || null, normalizedName, targetDate, normalizedClinic, settings);
          
          let totalHours = 0;
          if (segments.has('am')) totalHours += 4;
          if (segments.has('pm')) totalHours += 3;
          if (segments.has('night')) totalHours += 3;
          
          let cost = 0;
          let wage = 0;
          if (doctorType === '常勤') {
            wage = settings.ftHourlyWageMap.get(normalizedName) || 10000;
            cost = totalHours * wage;
          } else if (doctorType === '定期非常勤') {
             wage = settings.ptContractWageMap.get(normalizedName) || 15000;
             cost = totalHours * wage;
          } else { // スポット
            const dayType = gt_getDayType(targetDate, settings.holidays);
            let spotCost = 0;
            if (segments.has('am')) {
                wage = getCorrectPay({name: 'am'}, dayType, settings.payRateMaster.get(`${normalizedClinic}||小児科`) || settings.payRateMaster.get(`${normalizedClinic}||共通`)) || 10000;
                spotCost += 4 * wage;
            }
            if (segments.has('pm')) {
                wage = getCorrectPay({name: 'pm'}, dayType, settings.payRateMaster.get(`${normalizedClinic}||小児科`) || settings.payRateMaster.get(`${normalizedClinic}||共通`)) || 10000;
                spotCost += 3 * wage;
            }
            if (segments.has('night')) {
                wage = getCorrectPay({name: 'night'}, dayType, settings.payRateMaster.get(`${normalizedClinic}||小児科`) || settings.payRateMaster.get(`${normalizedClinic}||共通`)) || 10000;
                spotCost += 3 * wage;
            }
            cost = spotCost;
          }
          totalDailyCost += cost;

          const dateForAgencyCheck = `${targetDate.getFullYear()}/${targetDate.getMonth() + 1}/${targetDate.getDate()}`;
          const typeAbbr = { '常勤': '常', '定期非常勤': '定', 'スポット': settings.agencyShiftSet.has(`${normalizedName}|${dateForAgencyCheck}`) ? '紹' : '直'}[doctorType];
          const displayStr = `(${typeAbbr}) ${normalizedName}`;
          
          if(segments.has('am')) finalSchedule.am.push(displayStr);
          if(segments.has('pm')) finalSchedule.pm.push(displayStr);
          if(segments.has('night')) finalSchedule.night.push(displayStr);
        });

        const processVacantSlot = (segment, doctors) => {
           if (doctors.length > 0) return;
           const dayType = gt_getDayType(targetDate, settings.holidays);
           let payRates = settings.payRateMaster.get(`${normalizedClinic}||小児科`) || settings.payRateMaster.get(`${normalizedClinic}||共通`);
           let wage = payRates ? getCorrectPay({name: segment}, dayType, payRates) : 0;
           let hours = (segment === 'am') ? 4 : 3;
           const cost = hours * (wage || 10000);
           totalDailyCost += cost;
           const wageStr = typeof wage === 'number' ? wage.toLocaleString() + '円' : wage;
           finalSchedule[segment].push(`募集枠（時給 ${wageStr}）`);
        };
        processVacantSlot('am', amDoctors);
        processVacantSlot('pm', pmDoctors);
        processVacantSlot('night', nightDoctors);
        
        break; 
      }
    }
    
    if (!targetRowFound) {
      const segments = ['am', 'pm', 'night'];
      segments.forEach(segment => {
          const dayType = gt_getDayType(targetDate, settings.holidays);
          let payRates = settings.payRateMaster.get(`${normalizedClinic}||小児科`) || settings.payRateMaster.get(`${normalizedClinic}||共通`);
          let wage = payRates ? getCorrectPay({name: segment}, dayType, payRates) : 0;
          const wageStr = typeof wage === 'number' ? wage.toLocaleString() + '円' : wage;
          finalSchedule[segment].push(`募集枠（時給 ${wageStr}）`);
          let hours = (segment === 'am') ? 4 : 3;
          totalDailyCost += hours * (wage || 10000);
      });
    }

  } catch (e) {
    Logger.log(`getDailyScheduleAndCost Error: ${e.message}`);
    finalSchedule.am = [`エラー: ${e.message}`];
  }
  
  return { schedule: finalSchedule, dailyCost: totalDailyCost };
}

function getDailySchedule(clinic, dateStr) {
  const settings = _loadInitialData(SpreadsheetApp.getActiveSpreadsheet());
  const targetDate = new Date(dateStr);
  const { schedule } = getDailyScheduleAndCost(clinic, targetDate, settings);
  return schedule;
}

function startOsaifuKunCalculation(clinic, dateStr) {
  try {
    const settings = _loadInitialData(SpreadsheetApp.getActiveSpreadsheet());
    const targetDate = new Date(dateStr);
    const monthStartDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
    const monthEndDate = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0);

    const sessionData = {
      clinic: clinic,
      targetDate: targetDate.toISOString(),
      monthStartDate: monthStartDate.toISOString(),
      monthEndDate: monthEndDate.toISOString(),
    };
    
    const lightSettings = {};
    const heavySettings = {};
    for (const key in settings) {
      const value = settings[key];
      if (key === 'salesUrlMaps' && value) {
        const salesUrlMapsAsArrays = {};
        for (const yearKey in value) {
          if (value[yearKey] instanceof Map) salesUrlMapsAsArrays[yearKey] = Array.from(value[yearKey].entries());
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
    cache.put('osaifukun_session', JSON.stringify(sessionData), 1800);
    cache.put('osaifukun_settings_light', JSON.stringify(lightSettings), 1800);
    cache.put('osaifukun_settings_heavy', JSON.stringify(heavySettings), 1800);
    cache.put('osaifukun_daily_results', JSON.stringify({}), 1800);

    const totalDays = (monthEndDate.getTime() - monthStartDate.getTime()) / (1000 * 60 * 60 * 24) + 1;
    return { success: true, totalDays: totalDays };
  } catch (e) {
    Logger.log(`startOsaifuKunCalculation Error: ${e.message}`);
    return { success: false, message: e.message };
  }
}

function processOsaifuKunDay(dayIndex) {
  const cache = CacheService.getScriptCache();
  const sessionJson = cache.get('osaifukun_session');
  if (!sessionJson) return "エラー: セッションがタイムアウトしました。";
  
  const session = JSON.parse(sessionJson);
  const startDate = new Date(session.monthStartDate);
  const currentDate = new Date(new Date(session.monthStartDate).setDate(startDate.getDate() + dayIndex));
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  try {
    const settings = restoreSettingsFromCache('osaifukun_');
    const clinic = session.clinic;
    let dailyCost = 0;

    if (currentDate < today) {
      const dailyDataForOneDay = {};
      _calculateDailyPersonnelCosts(currentDate, dailyDataForOneDay, settings);
      const dateStr = `${currentDate.getFullYear()}/${currentDate.getMonth() + 1}/${currentDate.getDate()}`;
      const clinicData = dailyDataForOneDay[dateStr]?.[clinic];
      dailyCost = clinicData ? clinicData.totalCost : 0;
    } else {
      const { dailyCost: predictedCost } = getDailyScheduleAndCost(clinic, currentDate, settings);
      dailyCost = predictedCost;
    }
    
    const dailyResults = JSON.parse(cache.get('osaifukun_daily_results'));
    dailyResults[dayIndex] = dailyCost;
    cache.put('osaifukun_daily_results', JSON.stringify(dailyResults), 1800);

    return `${Utilities.formatDate(currentDate, "JST", "MM/dd")} のコスト(¥${Math.round(dailyCost).toLocaleString()})を計算しました。`;
  } catch (e) {
    Logger.log(`processOsaifuKunDay Error: ${e.message}\n${e.stack}`);
    return `エラー (${Utilities.formatDate(currentDate, "JST", "MM/dd")}): ${e.message}`;
  }
}

function finishOsaifuKunCalculation() {
  const cache = CacheService.getScriptCache();
  const sessionJson = cache.get('osaifukun_session');
  if (!sessionJson) return { success: false, message: "セッションがタイムアウトしました。" };
  
  try {
    const session = JSON.parse(sessionJson);
    const settings = restoreSettingsFromCache('osaifukun_');
    const monthStartDate = new Date(session.monthStartDate);
    const month = monthStartDate.getMonth() + 1;
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const budgetData = settings.budgetMap.get(session.clinic) || {};
    const totalBudget = (budgetData.ft_budget?.[month] || 0) + (budgetData.pt_budget?.[month] || 0) + (budgetData.spot_budget?.[month] || 0) + (budgetData.second_exam_budget?.[month] || 0);

    let totalActuals = 0;
    let totalProjections = 0;
    const dailyResults = JSON.parse(cache.get('osaifukun_daily_results'));

    for (const dayIndex in dailyResults) {
      const currentDate = new Date(new Date(session.monthStartDate).setDate(new Date(session.monthStartDate).getDate() + parseInt(dayIndex)));
      if (currentDate < today) {
        totalActuals += dailyResults[dayIndex] || 0;
      } else {
        totalProjections += dailyResults[dayIndex] || 0;
      }
    }

    const remainingAmount = totalBudget - totalActuals - totalProjections;
    cancelOsaifuKunCalculation();

    return { success: true, remainingAmount, budget: totalBudget, totalActuals, totalProjections };
  } catch (e) {
    Logger.log(`finishOsaifuKunCalculation Error: ${e.message}\n${e.stack}`);
    return { success: false, message: e.message };
  }
}

function cancelOsaifuKunCalculation() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(['osaifukun_session', 'osaifukun_settings_light', 'osaifukun_settings_heavy', 'osaifukun_daily_results']);
}

function restoreSettingsFromCache(prefix) {
    const cache = CacheService.getScriptCache();
    const lightSettingsJson = cache.get(prefix + 'settings_light');
    const heavySettingsJson = cache.get(prefix + 'settings_heavy');
    if (!lightSettingsJson || !heavySettingsJson) throw new Error("セッションが見つかりません。");
    const settings = JSON.parse(lightSettingsJson);
    const heavySettings = JSON.parse(heavySettingsJson);
    for(const key in heavySettings) { settings[key] = new Map(heavySettings[key]); }
    if (settings.salesUrlMaps) {
      const restoredSalesUrlMaps = {};
      for (const yearKey in settings.salesUrlMaps) { restoredSalesUrlMaps[yearKey] = new Map(settings.salesUrlMaps[yearKey]); }
      settings.salesUrlMaps = restoredSalesUrlMaps;
    }
    if (settings.holidays) settings.holidays = new Set(settings.holidays);
    if (settings.agencyShiftSet) settings.agencyShiftSet = new Set(settings.agencyShiftSet);
    if(settings.ptDoctorContractMapById) {
      for (const contractInfo of settings.ptDoctorContractMapById.values()) {
        if (contractInfo.startDate) { contractInfo.startDate = new Date(contractInfo.startDate); }
      }
    }
    return settings;
}

function _isValidNameForSchedule(name) {
  if (!name || name.length < 2) return false;
  if (name.includes('募集')) return true;
  if (/^[\d\s-:]+$/.test(name)) return false; 
  if (/^(月|火|水|木|金|土|日|曜日|第)$/.test(name)) return false;
  return true;
}