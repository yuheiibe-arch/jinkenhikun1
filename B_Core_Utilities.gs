/**
 * =================================================================
 * 【ファイル 2/6: 共通コア (B_Core_Utilities.gs)】- 最終完成版 (特別時給解析対応)
 * =================================================================
 */

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

  // ====================================================================
  // 🌟 2026年度 外部医師マスタ（常勤・非常勤）
  // ====================================================================
  const MASTER_URL = "https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit";
  const masterSs = SpreadsheetApp.openByUrl(MASTER_URL);
  const baseDate = targetDateObj || new Date(); 

  const createHeaderMap = (headerRow) => {
    const map = {};
    headerRow.forEach((val, idx) => {
      if (val) map[String(val).replace(/\n/g, '').trim()] = idx;
    });
    return map;
  };

  const ftHourlyWageMap = new Map();
  
  const ptWageTypeMap = new Map(); // "BASE", "FLAT", "SPECIAL"
  const ptContractWageMap = new Map(); // FLATの場合の固定時給
  const ptSpecialWageRulesMap = new Map(); // SPECIALの場合のルール配列

  const partTimeDoctorNameSet = new Set();
  const doctorContractMap = new Map();
  const ptDoctorContractMapById = new Map();
  const nameToMedicalIdMap = new Map();

  // --- [A] 常勤2026年度 ---
  const ftSheet = masterSs.getSheetByName('常勤2026年度');
  if (ftSheet) {
    const ftData = ftSheet.getDataRange().getValues();
    const ftHMap = createHeaderMap(ftData[0]);

    for (let i = 1; i < ftData.length; i++) {
      const row = ftData[i];
      const rawName = row[ftHMap['医師名']];
      const rawMedId = row[ftHMap['医籍番号']];
      const retireDateRaw = row[ftHMap['退職日']];
      
      if (!rawName) continue;
      if (retireDateRaw instanceof Date && !isNaN(retireDateRaw.getTime()) && retireDateRaw < baseDate) continue;

      const normName = gt_normalizePersonName(rawName, nameAliasMap);
      const medId = rawMedId ? String(rawMedId).trim() : null;
      const wageVal = parseInt(row[ftHMap['時給']], 10);
      const wage = isNaN(wageVal) ? 10000 : wageVal;

      if (normName) {
        ftHourlyWageMap.set(normName, wage);
        doctorContractMap.set(normName, { isFullTime: true });
        if (medId) nameToMedicalIdMap.set(normName, medId);
      }
    }
  }

  // --- [B] 定期非常勤2026年度 ---
  const ptSheet = masterSs.getSheetByName('定期非常勤2026年度');
  if (ptSheet) {
    const ptData = ptSheet.getDataRange().getValues();
    const ptHMap = createHeaderMap(ptData[0]);

    for (let i = 1; i < ptData.length; i++) {
      const row = ptData[i];
      const rawName = row[ptHMap['医師名']];
      const rawMedId = row[ptHMap['医籍番号']];
      const retireDateRaw = row[ptHMap['退職日']];
      const contractText = row[ptHMap['勤務備考']];
      
      const contractWageType = row[ptHMap['契約時給']] ? String(row[ptHMap['契約時給']]).trim() : "";
      const specialWageText = row[ptHMap['特別時給の内訳']] ? String(row[ptHMap['特別時給の内訳']]).trim() : "";
      
      if (!rawName) continue;
      if (retireDateRaw instanceof Date && !isNaN(retireDateRaw.getTime()) && retireDateRaw < baseDate) continue;

      const normName = gt_normalizePersonName(rawName, nameAliasMap);
      const medId = rawMedId ? String(rawMedId).trim() : null;

      if (normName) {
        partTimeDoctorNameSet.add(normName);
        if (medId) nameToMedicalIdMap.set(normName, medId);

        // 🌟【新規】非常勤の時給タイプの仕分け
        if (contractWageType === "時給表どおり") {
          ptWageTypeMap.set(normName, "BASE");
        } else if (contractWageType.includes("特別時給") || specialWageText) {
          ptWageTypeMap.set(normName, "SPECIAL");
          const rules = _parseSpecialWageText(specialWageText, clinicAliasMap);
          ptSpecialWageRulesMap.set(normName, rules);
        } else {
          const wageVal = parseInt(contractWageType.replace(/,/g, ''), 10);
          if (!isNaN(wageVal)) {
            ptWageTypeMap.set(normName, "FLAT");
            ptContractWageMap.set(normName, wageVal);
          } else {
            ptWageTypeMap.set(normName, "BASE"); // フォールバック
          }
        }

        // 契約シフト情報のパース
        if (medId && contractText) {
          const rules = parseContractToRules(contractText);
          const startDateRaw = row[ptHMap['入職日']];
          let startDate = new Date('2000/01/01');
          if (startDateRaw instanceof Date && !isNaN(startDateRaw.getTime())) {
            startDate = new Date(startDateRaw);
            startDate.setHours(0, 0, 0, 0);
          }
          const contracts = rules.map(rule => ({
            location: gt_normalizeClinicName(rule.workplace, clinicAliasMap),
            dayOfWeek: rule.dayOfWeek
          })).filter(c => c.location && typeof c.dayOfWeek !== 'undefined');

          ptDoctorContractMapById.set(medId, { startDate, contracts, name: normName });
        }
      }
    }
  }

  // ====================================================================
  // 🌟 [C] 紹介会社医師履歴 (外部シート) 読込
  // ====================================================================
  const agencyDoctorMedIds = new Map();
  try {
    const AGENCY_URL = "https://docs.google.com/spreadsheets/d/1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE/edit";
    const agencySs = SpreadsheetApp.openByUrl(AGENCY_URL);
    const agencySheet = agencySs.getSheetByName("紹介会社医師履歴");
    
    if (agencySheet) {
      const agencyData = agencySheet.getDataRange().getValues();
      const aHMap = createHeaderMap(agencyData[0]);
      for (let i = 1; i < agencyData.length; i++) {
        const rawMedId = agencyData[i][aHMap['医籍番号']];
        if (rawMedId) {
          agencyDoctorMedIds.set(String(rawMedId).trim(), true);
        }
      }
    }
  } catch (e) {
    Logger.log(`⚠️ 紹介会社マスタの読み込みに失敗しました: ${e.message}`);
  }
  // ====================================================================

  const salesUrlMaps = {};
  urlListData.filter(r => r[0].toString().includes('売上')).forEach(e => {
    const m = e[0].match(/(\d{4})売上/);
    if (m) salesUrlMaps[parseInt(m[1])] = gt_createSalesUrlMap(e[1]);
  });
  
  const actualShiftUrls = new Map(urlListData.filter(r => r[0].includes('確定シフト')).map(r => [parseInt(r[0].match(/\d{4}/)?.[0]), r[1]]).filter(r => r[0]));
  
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
    ftHourlyWageMap, 
    ptWageTypeMap, ptContractWageMap, ptSpecialWageRulesMap, // ★新規追加: PTの時給分岐用
    doctorContractMap, ptDoctorContractMapById,
    agencyShiftSet: _getAgencyShiftSet(urlListData, nameAliasMap), 
    agencyDoctorMedIds: agencyDoctorMedIds,
    budgetMap: _getBudgetMap(urlListData, clinicAliasMap, currentFiscalYear),
    salesUrlMaps, actualShiftUrls,
    actualShiftUrl: actualShiftUrls.get(2025) || '',
    ftWageMap_LastYear: ftWageMap_LastYear,
    nameToMedicalIdMap: nameToMedicalIdMap,
    clinicOpeningDates: clinicOpeningDates,
    partTimeDoctorNameSet: partTimeDoctorNameSet
  };
}

/**
 * 特別時給のテキストを解析するパーサー
 * 例: "平日\n【錦糸町】19:00-21:00：時給：14,000円"
 */
function _parseSpecialWageText(text, clinicAliasMap) {
  const rules = [];
  const lines = text.split('\n');
  let currentDayType = "ALL";
  
  for (let line of lines) {
    line = line.trim();
    if (!line) continue;
    if (line === "平日") { currentDayType = "平日"; continue; }
    if (line === "土曜日" || line === "土曜") { currentDayType = "土曜"; continue; }
    if (line === "日曜日" || line === "日曜") { currentDayType = "日曜"; continue; }
    if (line === "祝日" || line === "日祝") { currentDayType = "祝日"; continue; }
    
    // 【拠点】00:00-00:00：時給：00,000円 の形式を抽出
    const match = line.match(/【([^】]+)】\s*(\d{1,2}:\d{2})\s*[-~〜]\s*(\d{1,2}:\d{2}).*?([\d,]+)円/);
    if (match) {
      const loc = gt_normalizeClinicName(match[1], clinicAliasMap);
      const start = match[2];
      const end = match[3];
      const wage = parseInt(match[4].replace(/,/g, ''), 10);
      rules.push({ dayType: currentDayType, loc: loc, start: start, end: end, wage: wage });
    }
  }
  return rules;
}

// 既存のヘルパー関数群 (省略せずそのまま維持)
function _getDoctorTypeForShift(medicalId, normalizedName, shiftDate, shiftLocation, settings) {
  if (medicalId && settings.ptDoctorContractMapById.has(medicalId)) {
    const contractInfo = settings.ptDoctorContractMapById.get(medicalId);
    if (!contractInfo.startDate || shiftDate < contractInfo.startDate) return "スポット";
    const shiftDayOfWeek = shiftDate.getDay();
    const isContractedShift = contractInfo.contracts.some(c => c.location === shiftLocation && c.dayOfWeek === shiftDayOfWeek);
    return isContractedShift ? "定期非常勤" : "スポット";
  }
  const ftContractInfo = settings.doctorContractMap.get(normalizedName);
  if (ftContractInfo && ftContractInfo.isFullTime) { return "常勤"; }
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
    rules.push({ dayOfWeek: dayMap[dayMatch[0].charAt(0)], workplace: workplaceMatch[1].trim() });
  };
  lines.forEach(line => parseLine(line));
  return rules.filter((rule, index, self) => index === self.findIndex(r => (r.dayOfWeek === rule.dayOfWeek && r.workplace === rule.workplace)));
}

function gt_createSalesUrlMap(indexUrl) {
  const salesIndexSs = SpreadsheetApp.openByUrl(indexUrl);
  const monthlyUrlsData = salesIndexSs.getSheets()[0].getDataRange().getValues();
  const monthlyUrlMap = new Map();
  monthlyUrlsData.forEach(row => { if (row[0] instanceof Date) monthlyUrlMap.set(row[0].getMonth(), row[1]); });
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
    masterData.set(`${clinic}||${department}`, {
      weekday: { am: row[3], pm: row[4], night: row[5] },
      saturday: { am: row[7], pm: row[8], night: row[9] },
      holiday: { am: row[11], pm: row[12], night: row[13] }
    });
  });
  return masterData;
}

function gt_getHolidays(spreadsheet) {
  const holidaySheet = spreadsheet.getSheetByName('祝日');
  if (!holidaySheet) return new Set();
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
  if (masterRow[type] && masterRow[type][zone]) return masterRow[type][zone];
  return null;
}

function _getAgencyShiftSet(urlListData, nameAliasMap) {
  const agencyShiftSet = new Set();
  const entry = urlListData.find(row => row[0] === "紹介会社");
  if (!entry || !entry[1] || !entry[2]) return agencyShiftSet;
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
        agencyShiftSet.add(`${gt_normalizePersonName(name, nameAliasMap)}|${date.getFullYear()}/${date.getMonth() + 1}/${date.getDate()}`);
      }
    });
  });
  return agencyShiftSet;
}

function _getBudgetMap(urlListData, clinicAliasMap, fiscalYear) {
  const budgetMap = new Map();
  const entry = urlListData.find(row => row[0] === `${fiscalYear}予算`);
  if (!entry || !entry[1] || !entry[2]) return budgetMap;
  try {
    const sheet = SpreadsheetApp.openByUrl(entry[1]).getSheetByName(entry[2]);
    const data = sheet.getDataRange().getValues();
    const monthHeaders = data[0];
    const monthCols = {};
    for (let i = 4; i < monthHeaders.length; i++) {
      const monthMatch = String(monthHeaders[i]).match(/(\d+)\/(\d+)/);
      if (monthMatch) monthCols[monthMatch[2]] = i;
    }
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rawClinicName = row[0];
      if (!rawClinicName) continue;
      const officialClinicName = gt_normalizeClinicName(rawClinicName, clinicAliasMap);
      if (!officialClinicName) continue;
      if (!budgetMap.has(officialClinicName)) budgetMap.set(officialClinicName, {});
      
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
        for (const month in monthCols) clinicBudget[targetType][month] = row[monthCols[month]] || 0;
      } else if (String(itemColC).trim() === "2診目" && String(itemColD).startsWith("コスト")) {
        if (!clinicBudget["second_exam_budget"]) clinicBudget["second_exam_budget"] = {};
        for (const month in monthCols) clinicBudget["second_exam_budget"][month] = row[monthCols[month]] || 0;
      }
    }
  } catch (e) { Logger.log(`予算マップの作成中にエラー: ${e.message}`); }
  return budgetMap;
}

function _loadLastYearFtWageMap(spreadsheetUrl) {
  const ftWageMap_LastYear = new Map();
  if (!spreadsheetUrl) return ftWageMap_LastYear;
  try {
    const sheet = SpreadsheetApp.openByUrl(spreadsheetUrl).getSheetByName("常勤医師給与");
    if (!sheet) return ftWageMap_LastYear;
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    data.forEach(row => {
      const averageWage = row[2];
      const medicalId = row[3];
      if (medicalId && typeof averageWage === 'number' && averageWage > 0) ftWageMap_LastYear.set(String(medicalId).trim(), averageWage);
    });
  } catch (e) {}
  return ftWageMap_LastYear;
}