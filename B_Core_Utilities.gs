/**
 * =================================================================
 * 【ファイル 2/6: 共通コア (B_Core_Utilities.gs)】- 修正版
 * =================================================================
 */

// ★ 修正箇所1: 引数に targetDateObj を追加し、計算対象の日付を受け取れるようにしました
function _loadInitialData(spreadsheet, targetDateObj) {
  const urlListSheet = spreadsheet.getSheetByName("URLリスト");
  if (!urlListSheet) throw new Error("「URLリスト」シートが見つかりません。");
  const urlListData = urlListSheet.getDataRange().getValues();
  const getEntry = (name) => {
    const entry = urlListData.find(row => row[0] === name);
    if (!entry || !entry[1]) throw new Error(`URLリストに「${name}」の定義が見つかりません。`);
    return entry;
  };

  const normalizationSs = SpreadsheetApp.openByUrl(getEntry('正規表現')[1]);
  const clinicAliasMap = new Map();
  const clinicGroupMap = new Map();
  const clinicOpeningDates = new Map(); 

  normalizationSs.getSheetByName('拠点名').getDataRange().getValues().slice(1).forEach(row => {
    const correctName = row[0];
    const group = row[5];
    const openingDate = row[7];
    if (correctName && group) clinicGroupMap.set(correctName, group);
    if (correctName && openingDate instanceof Date) {
      openingDate.setHours(0, 0, 0, 0);
      clinicOpeningDates.set(correctName, openingDate);
    }
    if (correctName) row.slice(1, 5).forEach(alias => {
      if (alias) clinicAliasMap.set(alias.toString().trim(), correctName);
    });
  });
  
  const nameAliasMap = new Map();
  normalizationSs.getSheetByName('氏名').getDataRange().getValues().slice(1).forEach(row => {
    const correctName = row[0];
    if (correctName) row.slice(1).forEach(alias => {
      if (alias) nameAliasMap.set(alias.toString().trim(), correctName);
    });
  });

  let ftHourlyWageMap = new Map();
  const ftMasterSheet = spreadsheet.getSheetByName('常勤マスタ');
  if (ftMasterSheet) {
    ftHourlyWageMap = new Map(ftMasterSheet.getDataRange().getValues().slice(1).map(row => [gt_normalizePersonName(row[2], nameAliasMap), row[25]]));
  }
  
  let ptContractWageMap = new Map();
  let partTimeDoctorNameSet = new Set();
  const ptMasterSheet = spreadsheet.getSheetByName('定期非常勤マスタ');
  if (ptMasterSheet) {
    const ptMasterData = ptMasterSheet.getDataRange().getValues().slice(1);
    ptContractWageMap = new Map(ptMasterData.map(row => [gt_normalizePersonName(row[2], nameAliasMap), row[25]])); 
    partTimeDoctorNameSet = new Set(ptMasterData.map(row => gt_normalizePersonName(row[2], nameAliasMap)).filter(Boolean));
  }

  const doctorContractMap = new Map();
  const ptDoctorContractMapById = new Map();
  const nameToMedicalIdMap = new Map();

  try {
    const doctorListEntry = getEntry('定期勤務医師リスト_2025年度');
    const doctorListSheet = SpreadsheetApp.openByUrl(doctorListEntry[1]).getSheetByName(doctorListEntry[2]);
    const doctorListValues = doctorListSheet.getDataRange().getValues();
    if (doctorListValues.length > 4) {
      const header = doctorListValues[4].map(h => String(h).trim());
      let nameColIdx = header.indexOf("氏名\nスペース\nいれない！！");
      if (nameColIdx === -1) nameColIdx = header.indexOf("氏名");
      
      const medicalIdColIdx = header.indexOf("医籍番号");
      const typeColIdx = header.indexOf("医師区分");
      const contractColIdx = header.indexOf("2025年度契約内容");
      const startDateColIdx = header.indexOf("入職日");

      doctorListValues.slice(5).forEach(row => {
        const type = row[typeColIdx];
        const normalizedName = gt_normalizePersonName(row[nameColIdx], nameAliasMap);
        const medicalId = row[medicalIdColIdx] ? String(row[medicalIdColIdx]).trim() : null;
        if (normalizedName && medicalId) {
            nameToMedicalIdMap.set(normalizedName, medicalId);
        }
        if (type === "常勤") {
          if (normalizedName) doctorContractMap.set(normalizedName, { isFullTime: true });
        } else if (type === "定期非常勤") {
          if (!medicalId) return;
          const contractText = row[contractColIdx];
          const startDateValue = row[startDateColIdx];
          if (contractText && startDateValue instanceof Date) {
            const startDate = new Date(startDateValue);
            startDate.setHours(0,0,0,0);
            const rules = parseContractToRules(contractText);
            const contracts = rules.map(rule => ({
              location: gt_normalizeClinicName(rule.workplace, clinicAliasMap),
              dayOfWeek: rule.dayOfWeek
            })).filter(c => c.location && typeof c.dayOfWeek !== 'undefined');
            ptDoctorContractMapById.set(medicalId, { startDate, contracts, name: normalizedName });
          }
        }
      });
    }
  } catch (e) {
    Logger.log(`警告: '定期勤務医師リスト_2025年度' の読み込みに失敗: ${e.message}\n${e.stack}`);
  }

  const salesUrlMaps = {};
  urlListData.filter(r => r[0].toString().includes('売上')).forEach(e => {
    const m = e[0].match(/(\d{4})売上/);
    if (m) salesUrlMaps[parseInt(m[1])] = gt_createSalesUrlMap(e[1]);
  });
  
  const actualShiftUrls = new Map(urlListData.filter(r => r[0].includes('確定シフト')).map(r => [parseInt(r[0].match(/\d{4}/)?.[0]), r[1]]).filter(r => r[0]));
  
  // ★ 修正箇所2: ターゲットの日付（または本日）から、動的に対象の「今年度」と「昨年度」を判定する
  const today = targetDateObj || new Date();
  const currentFiscalYear = (today.getMonth() + 1 <= 3) ? today.getFullYear() - 1 : today.getFullYear();
  const lastFiscalYear = currentFiscalYear - 1;
  const urlForLastYearShifts = actualShiftUrls.get(lastFiscalYear);
  const ftWageMap_LastYear = _loadLastYearFtWageMap(urlForLastYearShifts);
  
  return {
    clinicAliasMap, clinicGroupMap, nameAliasMap,
    kantoUrl: getEntry('2025関東シフト表')[1],
    kansaiUrl: getEntry('2025関西シフト表')[1],
    payRateMaster: gt_getPayRateMaster(getEntry('2025時給表')[1]),
    holidays: gt_getHolidays(spreadsheet),
    ftHourlyWageMap, ptContractWageMap, doctorContractMap,
    ptDoctorContractMapById,
    agencyShiftSet: _getAgencyShiftSet(urlListData, nameAliasMap),
    
    // ★ 修正箇所3: 決め打ちを辞め、対象の「今年度（currentFiscalYear）」を渡して動的に取得する
    budgetMap: _getBudgetMap(urlListData, clinicAliasMap, currentFiscalYear),
    
    salesUrlMaps, actualShiftUrls,
    actualShiftUrl: actualShiftUrls.get(2025) || '',
    ftWageMap_LastYear: ftWageMap_LastYear,
    nameToMedicalIdMap: nameToMedicalIdMap,
    clinicOpeningDates: clinicOpeningDates,
    partTimeDoctorNameSet: partTimeDoctorNameSet
  };
}

function _getDoctorTypeForShift(medicalId, normalizedName, shiftDate, shiftLocation, settings) {
  if (medicalId && settings.ptDoctorContractMapById.has(medicalId)) {
    const contractInfo = settings.ptDoctorContractMapById.get(medicalId);
    if (!contractInfo.startDate || shiftDate < contractInfo.startDate) return "スポット";
    const shiftDayOfWeek = shiftDate.getDay();
    const isContractedShift = contractInfo.contracts.some(c => c.location === shiftLocation && c.dayOfWeek === shiftDayOfWeek);
    return isContractedShift ? "定期非常勤" : "スポット";
  }
  const ftContractInfo = settings.doctorContractMap.get(normalizedName);
  if (ftContractInfo && ftContractInfo.isFullTime) {
    return "常勤";
  }
  return "スポット";
}

function parseContractToRules(text) {
  if (!text || typeof text !== 'string') return [];
  const lines = text.split('\n').map(line => line.trim()).filter(line => line);
  const rules = [];
  const dayMap = { "日": 0, "月": 1, "火": 2, "水": 3, "木": 4, "金": 5, "土": 6 };

  const parseLine = (line) => {
    if (line.startsWith('(') && line.includes('時給')) return;
    const dayMatch = line.match(/(月|火|水|木|金|土|日)曜日?/);
    if (!dayMatch) return;
    const workplaceMatch = line.match(/【([^】]+)】/);
    if (!workplaceMatch) return;

    rules.push({
      dayOfWeek: dayMap[dayMatch[0].charAt(0)],
      workplace: workplaceMatch[1].trim()
    });
  };

  lines.forEach(line => parseLine(line));

  const uniqueRules = rules.filter((rule, index, self) =>
    index === self.findIndex(r => (
      r.dayOfWeek === rule.dayOfWeek && r.workplace === rule.workplace
    ))
  );
  return uniqueRules;
}


function gt_createSalesUrlMap(indexUrl) {
  const salesIndexSs = SpreadsheetApp.openByUrl(indexUrl);
  const monthlyUrlsData = salesIndexSs.getSheets()[0].getDataRange().getValues();
  const monthlyUrlMap = new Map();
  monthlyUrlsData.forEach(row => {
    if (row[0] instanceof Date) {
      monthlyUrlMap.set(row[0].getMonth(), row[1]);
    }
  });
  return monthlyUrlMap;
}

function gt_normalizeClinicName(name, aliasMap) {
  if (!name) return '';
  const trimmedName = name.toString().trim();
  if (aliasMap.has(trimmedName)) return aliasMap.get(trimmedName);
  const normalized = trimmedName.replace(/^[0-9.\s]+/, '').split('(')[0].trim();
  return aliasMap.get(normalized) || normalized;
}

function gt_normalizePersonName(name, aliasMap) {
  if (!name || typeof name !== 'string') return '';
  const rawTrimmed = name.toString().trim();
  if (aliasMap.has(rawTrimmed)) return aliasMap.get(rawTrimmed);
  const halfSpaceVer = rawTrimmed.replace(/　/g, ' ');
  if (aliasMap.has(halfSpaceVer)) return aliasMap.get(halfSpaceVer);
  const noSpaceVer = halfSpaceVer.replace(/ /g, '');
  if (aliasMap.has(noSpaceVer)) return aliasMap.get(noSpaceVer);
  return noSpaceVer;
}

function gt_getPayRateMaster(url) {
  const masterSheet = SpreadsheetApp.openByUrl(url).getSheetByName('2025年度（年間）');
  if (!masterSheet) throw new Error("時給マスタに「2025年度（年間）」シートが見つかりません。");
  const values = masterSheet.getRange('B2:O' + masterSheet.getLastRow()).getValues();
  const masterData = new Map();
  values.forEach(row => {
    const clinic = (row[0] || '').toString().trim();
    const department = (row[1] || '').toString().replace(/\s+/g, '');
    if (!clinic) return;
    const key = `${clinic}||${department}`;
    const payRates = {
      weekday: { am: row[3], pm: row[4], night: row[5] },
      saturday: { am: row[7], pm: row[8], night: row[9] },
      holiday: { am: row[11], pm: row[12], night: row[13] }
    };
    masterData.set(key, payRates);
  });
  return masterData;
}

function gt_getHolidays(spreadsheet) {
  const holidaySheet = spreadsheet.getSheetByName('祝日');
  if (!holidaySheet) {
    return new Set();
  }
  return new Set(holidaySheet.getRange('A:A').getValues().flat().filter(String).map(d => Utilities.formatDate(new Date(d), "JST", "yyyy-MM-dd")));
}

function gt_getDayType(date, holidays) {
  const workDate = (date instanceof Date) ? date : new Date(date);
  if (isNaN(workDate.getTime())) return 'weekday';
  const dateStr = Utilities.formatDate(workDate, "JST", "yyyy-MM-dd");
  if (holidays.has(dateStr) || workDate.getDay() === 0) return 'holiday';
  if (workDate.getDay() === 6) return 'saturday';
  return 'weekday';
}

function gt_getCorrectPay(segment, dayType, masterRow) {
  if (!segment || !masterRow) return null;
  const type = dayType.toLowerCase();
  const zone = segment.name;
  if (masterRow[type] && masterRow[type][zone]) {
    return masterRow[type][zone];
  }
  return null;
}

function _getAgencyShiftSet(urlListData, nameAliasMap) {
  const agencyShiftSet = new Set();
  const entry = urlListData.find(row => row[0] === "紹介会社");
  if (!entry || !entry[1] || !entry[2]) {
    return agencyShiftSet;
  }
  const sheetNames = entry[2].toString().split(',').map(name => name.trim()).filter(String);
  const spreadsheet = SpreadsheetApp.openByUrl(entry[1]);
  sheetNames.forEach(sheetName => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 3) return;
    const header = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const nameCol = header.findIndex(h => h.includes("氏名") || h.includes("医師名"));
    const dateCol = header.indexOf("勤務日");
    if (nameCol === -1 || dateCol === -1) return;
    const data = sheet.getRange(3, 1, sheet.getLastRow() - 2, sheet.getLastColumn()).getValues();
    data.forEach(row => {
      const name = row[nameCol];
      const date = row[dateCol];
      if (name && date instanceof Date) {
        const dateStr = `${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`;
        agencyShiftSet.add(`${gt_normalizePersonName(name, nameAliasMap)}|${dateStr}`);
      }
    });
  });
  return agencyShiftSet;
}

// ★ 修正箇所4: 引数に fiscalYear を追加し、動的にその年の予算シートを探す
function _getBudgetMap(urlListData, clinicAliasMap, fiscalYear) {
  const budgetMap = new Map();
  const targetBudgetSheetName = `${fiscalYear}予算`; // 例: 2026予算

  const entry = urlListData.find(row => row[0] === targetBudgetSheetName);
  if (!entry || !entry[1] || !entry[2]) {
    Logger.log(`⚠️ URLリストに「${targetBudgetSheetName}」が見つかりません。予算はすべて0円として計算します。`);
    return budgetMap;
  }

  try {
    const sheet = SpreadsheetApp.openByUrl(entry[1]).getSheetByName(entry[2]);
    const data = sheet.getDataRange().getValues();
    const monthHeaders = data[0];
    const monthCols = {};
    for (let i = 4; i < monthHeaders.length; i++) {
      const monthMatch = String(monthHeaders[i]).match(/(\d+)\/(\d+)/);
      if (monthMatch) {
        monthCols[monthMatch[2]] = i;
      }
    }
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawClinicName = row[0];
      if (!rawClinicName) continue;
      const officialClinicName = gt_normalizeClinicName(rawClinicName, clinicAliasMap);
      if (!officialClinicName) continue;
      if (!budgetMap.has(officialClinicName)) {
        budgetMap.set(officialClinicName, {});
      }
      const clinicBudget = budgetMap.get(officialClinicName);
      const itemColC = row[2];
      const itemColD = row[3];
      if (itemColD === "基本給") {
        let targetType;
        if (itemColC === "常勤医師") targetType = "ft_budget";
        else if (itemColC === "定期非常勤医師") targetType = "pt_budget";
        else if (itemColC === "スポット医師") targetType = "spot_budget";
        else continue;
        if (!clinicBudget[targetType]) clinicBudget[targetType] = {};
        for (const month in monthCols) {
          clinicBudget[targetType][month] = row[monthCols[month]] || 0;
        }
      } else if (String(itemColC).trim() === "2診目" && String(itemColD).startsWith("コスト")) {
        const targetType = "second_exam_budget";
        if (!clinicBudget[targetType]) clinicBudget[targetType] = {};
        for (const month in monthCols) {
          clinicBudget[targetType][month] = row[monthCols[month]] || 0;
        }
      }
    }
  } catch (e) {
    Logger.log(`予算マップの作成中にエラー: ${e.message}`);
  }
  return budgetMap;
}

function _loadLastYearFtWageMap(spreadsheetUrl) {
  const ftWageMap_LastYear = new Map();
  if (!spreadsheetUrl) {
    return ftWageMap_LastYear;
  }
  try {
    const sheet = SpreadsheetApp.openByUrl(spreadsheetUrl).getSheetByName("常勤医師給与");
    if (!sheet) {
      return ftWageMap_LastYear;
    }
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    data.forEach(row => {
      const averageWage = row[2];
      const medicalId = row[3];
      if (medicalId && typeof averageWage === 'number' && averageWage > 0) {
        ftWageMap_LastYear.set(String(medicalId).trim(), averageWage);
      }
    });
  } catch (e) {
    Logger.log(`エラー: 「昨年度確定シフト」(常勤医師給与)の読み込みに失敗: ${e.message}`);
  }
  return ftWageMap_LastYear;
}